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
import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import DataView = powerbi.DataView;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import * as d3 from 'd3';
import DataViewObjects = powerbi.DataViewObjects;

import { VisualSettings } from "./settings";

export interface Pipeline {
    Title: String;
    Phase: string;
    Category: string;
    Name: string;
}

export interface Pipelines {
    SalesForce: Pipeline[];
}

export function logExceptions(): MethodDecorator {
    return (target: Object, propertyKey: string, descriptor: TypedPropertyDescriptor<any>)
        : TypedPropertyDescriptor<any> => {

        return {
            value: function () {
                try {
                    return descriptor.value.apply(this, arguments);
                } catch (e) {
                    // this.svg.append('text').text(e).style("stroke","black")
                    // .attr("dy", "1em");
                    throw e;
                }
            }
        };
    };
}

export function getCategoricalObjectValue<T>(objects: DataViewObjects, index: number, objectName: string, propertyName: string, defaultValue: T): T {
    if (objects) {
        let object = objects[objectName];
        if (object) {
            let property: T = <T>object[propertyName];
            if (property !== undefined) {
                return property;
            }
        }
    }
    return defaultValue;
}


export class Visual implements IVisual {
    private target: d3.Selection<HTMLElement, any, any, any>;
    private header: d3.Selection<HTMLElement, any, any, any>;
    private footer: d3.Selection<HTMLElement, any, any, any>;
    private margin = { top: 50, right: 40, bottom: 50, left: 40 };
    private settings: VisualSettings;
    private host: IVisualHost;
    private events: IVisualEventService;

    constructor(options: VisualConstructorOptions) {
        console.log('Visual Constructor', options);
        this.header = d3.select(options.element).append('div');
        this.target = d3.select(options.element).append('div');
        this.footer = d3.select(options.element).append('div');
        this.host = options.host;
        this.events = options.host.eventService;
    }

    @logExceptions()
    public update(options: VisualUpdateOptions) {
        this.events.renderingStarted(options);
        console.log('Visual Update ', options);
        this.settings = Visual.parseSettings(options && options.dataViews && options.dataViews[0]);
        this.target.selectAll('*').remove();
        let _this = this;
        this.target.attr('class', 'pipeline-container');
        if (this.settings.pipeline.layout) {
            this.target.attr('style', 'height:' + (options.viewport.height - 110) + 'px;width:' + (options.viewport.width) + 'px');
        }
        else {
            this.target.attr('style', 'height:' + (options.viewport.height) + 'px;width:' + (options.viewport.width) + 'px');
        }
        let gHeight = options.viewport.height - this.margin.top - this.margin.bottom;
        let gWidth = options.viewport.width - this.margin.left - this.margin.right;

        let pipelineData = Visual.CONVERTER(options.dataViews[0], this.host);
        pipelineData = pipelineData.slice(0, 250);

        let phaseData = this.settings.pipeline.phases.split(',');

        let colors = ['#2ECC71', '#336EFF', '#641E16', '#FF5733', '#3498DB', '#4A235A', '#154360', '#0B5345', '#784212', '#424949',
            '#17202A', '#E74C3C', '#00ff00', '#0000ff', '#252D48'];

        let categoryData = [];
        if (this.settings.pipeline.categories) {
            categoryData = this.settings.pipeline.categories.split(',');
        }
        else {
            categoryData = pipelineData.map(d => d.Category).filter((v, i, self) => self.indexOf(v) === i);
        }
        categoryData = categoryData.sort();
        let categoryColorData = categoryData.map((d, i) => {
            return {
                category: d,
                color: colors[i] // this.getRandomColor()
            };
        });

        this.renderLayout();

        this.renderPipelineReport(phaseData, pipelineData, categoryColorData);

        this.renderLegend(categoryColorData)

        this.handleScrollEvent();

        this.events.renderingFinished(options);
    }

    private renderLayout() {
        if (this.settings.pipeline.layout.toLowerCase() === 'header') {
            this.header
                .attr('class', 'visual-header')
                .html(() => {
                    if (this.settings.pipeline.imgURL) {
                        return '<img src="' + this.settings.pipeline.imgURL + '"/>';
                    }
                    else {
                        return "";
                    }
                });
        }
        else if (this.settings.pipeline.layout.toLowerCase() === 'footer') {
            this.footer
                .attr('class', 'visual-footer')
                .html(() => {
                    if (this.settings.pipeline.imgURL) {
                        return '<img src="' + this.settings.pipeline.imgURL + '"/>';
                    }
                    else {
                        return "";
                    }
                });
        }
    }

    private renderPipelineReport(phaseData, pipelineData, categoryColorData) {
        let mainContent = this.target.append('div')
            .attr('class', 'main-content');

        mainContent.append('div')
            .attr('class', 'header')
            .append('p').text(this.settings.pipeline.title);

        let pipelineWrap = mainContent.append('div')
            .attr('class', 'pipeline-wrap');

        let pipelineBar = pipelineWrap.append('div')
            .attr('class', 'pipeline-bar');

        let phases = pipelineBar.selectAll('.phase')
            .data(phaseData)
            .enter()
            .append('div')
            .attr('class', 'phase');

        phases.append('p')
            .attr('class', 'phase-text')
            .text((d: string, i) => {
                return d;
            });

        phases.append('div')
            .attr('class', 'phase-arrow');

        phases.append('div')
            .attr('class', 'phase-rope');

        phases.append('div')
            .attr('class', 'phase-rope-circle');

        let companiesWrap = pipelineWrap.append('div')
            .attr('class', 'companies-wrap');

        let phaseCompanies = companiesWrap.selectAll('.phase-companies')
            .data(phaseData)
            .enter()
            .append('div')
            .attr('class', 'phase-companies');

        let companies = phaseCompanies.selectAll('.phase-companies')
            .data((pd) => {
                return pipelineData.filter(d => d.Phase === pd);
            })
            .enter()
            .append('div')
            .attr('class', 'company');

        companies.append('p')
            .attr('class', 'company-name')
            .attr('style', (d: Pipeline) => {
                let [moAcolor] = categoryColorData.filter(cd => cd.category === d.Category);
                return 'color:' + moAcolor.color + ';';
            })
            .text((d: Pipeline) => {
                return d.Title ? d.Title.toString() : '';
            });

        companies.append('p')
            .attr('class', 'product-name')
            .attr('style', (d: Pipeline) => {
                let [moAcolor] = categoryColorData.filter(cd => cd.category === d.Category);
                return 'color:' + moAcolor.color + ';';
            }).text((d: Pipeline) => {
                return d.Name ? d.Name.toString() : '';
            });
    }

    private renderLegend(categoryColorData) {
        let pipelineWrap = d3.select('.pipeline-wrap');
        let legendWrap = pipelineWrap.append('div')
            .attr('class', 'legend-wrap');

        legendWrap.selectAll('.legend')
            .data(categoryColorData)
            .enter()
            .append('div')
            .attr('class', 'legend')
            .append('p')
            .attr('style', (d: any) => {
                return 'color:' + d.color + ';';
            })
            .text((d: any) => {
                return d.category ? d.category.toString() : '';
            });

        let legendWrapHeight = legendWrap.node().getBoundingClientRect().height;
        let calcHeight = 260 + legendWrapHeight;
        let companiesWrap = d3.select('.companies-wrap');
        companiesWrap.attr('style', 'height:calc(100% - ' + calcHeight + 'px);');
    }

    private handleScrollEvent() {
        let pipelineBar = d3.select('.pipeline-bar');
        let companiesWrap = d3.select('.companies-wrap');
        pipelineBar.attr('style', 'margin-left:0px;');
        companiesWrap.on('scroll', (e: Event) => {
            e = e || window.event;
            let target: any = e.target || e.srcElement;
            console.log('scroll', target.scrollLeft);
            pipelineBar.attr('style', 'margin-left:' + (-target.scrollLeft) + 'px;');
        });
    }

    public static CONVERTER(dataView: DataView, host: IVisualHost): Pipeline[] {
        let resultData: Pipeline[] = [];
        let tableView = dataView.table;
        let _rows = tableView.rows;
        let _columns = tableView.columns;
        let _titleIndex = -1, _phaseIndex = -1, _categoryIndex = -1, _nameIndex;
        for (let ti = 0; ti < _columns.length; ti++) {
            if (_columns[ti].roles.hasOwnProperty("Title")) {
                _titleIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Phase")) {
                _phaseIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Category")) {
                _categoryIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Name")) {
                _nameIndex = ti;
            }
        }
        for (let i = 0; i < _rows.length; i++) {
            let row = _rows[i];
            let dp = {
                Title: row[_titleIndex] ? row[_titleIndex].toString() : null,
                Phase: row[_phaseIndex] ? row[_phaseIndex].toString() : null,
                Category: row[_categoryIndex] ? row[_categoryIndex].toString() : null,
                Name: row[_nameIndex] ? row[_nameIndex].toString() : null
            };
            resultData.push(dp);
        }
        return resultData;
    }

    private static parseSettings(dataView: DataView): VisualSettings {
        return <VisualSettings>VisualSettings.parse(dataView);
    }

    /**
     * This function gets called for each of the objects defined in the capabilities files and allows you to select which of the
     * objects and properties you want to expose to the users in the property pane.
     *
     */
    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
        return VisualSettings.enumerateObjectInstances(this.settings || VisualSettings.getDefault(), options);
    }
}