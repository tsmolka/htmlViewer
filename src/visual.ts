/*
*  Power BI Visual CLI
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*  MIT License
*
*  Permission is hereby granted, free of charge, to any person obtaining a copy
*  of this software and associated documentation files (the ""Software""), to deal
*  in the Software without restriction, including without limitation the rights
*  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*  copies of the Software, and to permit persons to whom the Software is
*  furnished to do so, subject to the following conditions:
*
*  The above copyright notice and this permission notice shall be included in
*  all copies or substantial portions of the Software.
*
*  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
*  THE SOFTWARE.
*/
"use strict";

import "core-js/stable";
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;

import { dataViewObjectsParser } from "powerbi-visuals-utils-dataviewutils";
import DataViewObjectsParser = dataViewObjectsParser.DataViewObjectsParser;

export class VisualSettings extends DataViewObjectsParser {
    public dataPoint: DataPointSettings = new DataPointSettings();
}

export class DataPointSettings {
    public renderAsHTML: boolean = true;
    public textAlign: string = "left";
    public fontColor: string = "#000000";
    public fontSize: number = 11;
}

interface VisualViewModel {
    dataPoints: VisualDataPoint[];
    settings: VisualSettings;
}

interface VisualDataPoint {
    category: string;
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private settings: VisualSettings;
    private target: HTMLElement;
    private divElement: HTMLDivElement;

    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
        this.host = options.host;

        if (typeof document !== 'undefined') {
            document.addEventListener('click', this._onClick);
            this.divElement = document.createElement('div');
            this.target.appendChild(this.divElement);
        }
    }

    private _onClick = (e: MouseEvent) => {
        const el = e.target as HTMLElement;
        if (el.tagName === 'A') {
            e.preventDefault();
            this.host.launchUrl(el.getAttribute('href'));
        }
    }

    private stripHtml(html: string) {
        let el = document.createElement("div");
        el.innerHTML = html;
        return el.textContent || el.innerText || "";
    }

    public update(options: VisualUpdateOptions) {
        let viewModel: VisualViewModel = this.visualTransform(options, this.host);
        this.settings = viewModel.settings;

        this.divElement.style.setProperty('text-align', this.settings.dataPoint.textAlign);
        this.divElement.style.setProperty('color', this.settings.dataPoint.fontColor);
        this.divElement.style.setProperty('font-size', this.settings.dataPoint.fontSize.toString() + 'px', 'important');

        for (let i = 0; i < viewModel.dataPoints.length; i++) {
            if (i >= this.divElement.children.length) {
                this.divElement.appendChild(document.createElement('div'));
            }

            let el = this.divElement.children[i] as HTMLDivElement;
            if (this.settings.dataPoint.renderAsHTML) {
                el.innerHTML = viewModel.dataPoints[i].category;
            } else {
                el.innerText = this.stripHtml(viewModel.dataPoints[i].category);
            }
        }

        while (viewModel.dataPoints.length < this.divElement.children.length) {
            this.divElement.children[viewModel.dataPoints.length].remove();
        }
    }

    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
        return VisualSettings.enumerateObjectInstances(this.settings || VisualSettings.getDefault(), options);
    }

    public visualTransform(options: VisualUpdateOptions, host: IVisualHost): VisualViewModel {
        let viewModel: VisualViewModel = {
            dataPoints: [],
            settings: <VisualSettings>{}
        };
        viewModel.settings = VisualSettings.parse<VisualSettings>(options && options.dataViews && options.dataViews[0]);

        let dataViews = options.dataViews;
        if (!dataViews
            || !dataViews[0]
            || !dataViews[0].categorical
            || !dataViews[0].categorical.categories
            || !dataViews[0].categorical.categories[0].values)
            return viewModel;

        let category = dataViews[0].categorical.categories[0];
        for (let i = 0; i < category.values.length; i++) {
            viewModel.dataPoints.push({
                category: category.values[i].toString()
            });
        }

        return viewModel;
    }
}