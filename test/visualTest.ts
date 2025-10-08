/*
 *  Power BI Visualizations
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

import powerbi from "powerbi-visuals-api";
import {BaseType, select as d3Select} from "d3-selection";

import lodashMinBy from "lodash.minby";
import lodashMaxBy from "lodash.maxby";
import lodashUniq from "lodash.uniq";
import lodashUniqBy from "lodash.uniqby";

import {VisualData} from "./visualData";
import {VisualBuilder} from "./visualBuilder";
import {isColorAppliedToElements} from "./helpers/helpers";
import {
    assertColorsMatch,
    clickElement,
    createVisualHost,
    getRandomNumber,
    MockISelectionId,
    MockISelectionIdBuilder
} from "powerbi-visuals-utils-testutils";

import {pixelConverter as PixelConverter} from "powerbi-visuals-utils-typeutils";
import {valueFormatter} from "powerbi-visuals-utils-formattingutils";

import {Milestone, Task} from "../src/interfaces";
import {Gantt as VisualClass} from "../src/gantt";
import {getRandomHexColor, isValidDate} from "../src/utils";

import {DefaultOpacity, DimmedOpacity} from "../src/behavior";
import {DateType, MilestoneShape, ResourceLabelPosition} from "../src/enums";
import DataView = powerbi.DataView;
import PrimitiveValue = powerbi.PrimitiveValue;

import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
import IValueFormatter = valueFormatter.IValueFormatter;


const datesAmountForScroll: number = 90;

describe("Gantt", () => {
    let visualBuilder: VisualBuilder;
    let defaultDataViewBuilder: VisualData;
    let dataView: DataView;

    beforeEach(() => {
        visualBuilder = new VisualBuilder(1000, 500);

        defaultDataViewBuilder = new VisualData();
        dataView = defaultDataViewBuilder.getDataView();
        fixDataViewDateValuesAggregation(dataView);

    });

    function fixDataViewDateValuesAggregation(dataView: DataView) {
        let values = dataView.categorical.values[0].values;

        for (let i = 0; i < values.length; ++i) {
            let stringValue: string = values[i].toString();
            let index: number = stringValue.indexOf(")");

            if (stringValue.length - 1 !== index) {
                values[i] = new Date(stringValue.substring(0, index + 1));
            }
        }
    }

    function getUniqueParentsCount(dataView: DataView, parentColumnIndex: number) {
        let uniqueParents: string[] = [];

        dataView.table.rows.forEach(row => {
            if (row[parentColumnIndex] && uniqueParents.indexOf(row[parentColumnIndex] as string)) {
                uniqueParents.push(row[parentColumnIndex] as string);
            }
        });

        return uniqueParents.length;
    }

    describe("DOM tests", () => {

        // function that uses grep to filter
        function grep(val: BaseType[]) {
            return val.filter((e: Element) => e.innerHTML === "" || e.textContent === "");
        }

        it("svg element created", () => {
            expect(visualBuilder.mainElement[0]).not.toBeNull();
        });

        it("update", (done) => {
            visualBuilder.updateRenderTimeout(dataView, () => {
                const taskResources: HTMLElement[] = [];
                visualBuilder.tasks.forEach((element: HTMLElement) => {
                    const taskResourceNode: NodeListOf<HTMLElement> = element.querySelectorAll(".task-resource");
                    taskResourceNode.forEach((taskResource: HTMLElement) => {
                        taskResources.push(taskResource);
                    });
                });

                let countOfTaskLabels = taskResources.length;

                let countOfTaskLines = visualBuilder.taskLabels.length;
                let countOfTasks = visualBuilder.tasks.length;

                let uniqueParents = getUniqueParentsCount(dataView, 5);

                expect(countOfTaskLabels).toEqual((dataView.table?.rows?.length ?? 0) + uniqueParents);
                expect(countOfTaskLines).toEqual((dataView.table?.rows?.length ?? 0) + uniqueParents);
                expect(countOfTasks).toEqual((dataView.table?.rows?.length ?? 0) + uniqueParents);

                done();
            });
        });

        it("Task Elements are presented in DOM if and only if task name is available (specified)", (done) => {
            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnTask]);

            visualBuilder.updateRenderTimeout(dataView, () => {
                expect(visualBuilder.tasks.length).not.toEqual(0);
                done();
            });
        });

        it("When Task Element is Missing, empty viewport should be created", (done) => {
            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate,
                VisualData.ColumnResource]);

            visualBuilder.updateRenderTimeout(dataView, () => {
                let body = d3Select(visualBuilder.element);

                expect(body.select(".axis").selectAll("*").nodes().length).toEqual(1);
                expect(body.select(".task-lines").selectAll("task-labels").nodes().length).toEqual(0);
                expect(body.select(".chart .tasks").selectAll("*").nodes().length).toEqual(0);
                done();
            });
        });

        it("When task start time is missing, it should be set to today date", (done) => {
            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnEndDate,
                VisualData.ColumnResource]);

            visualBuilder.updateRenderTimeout(dataView, () => {
                let tasks = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];

                for (let task of tasks) {
                    expect(task.start.toDateString()).toEqual(new Date(Date.now()).toDateString());
                }

                done();
            });
        });

        it("Task Resource is Missing, not shown on dom", (done) => {
            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate]);

            fixDataViewDateValuesAggregation(dataView);

            visualBuilder.updateRenderTimeout(dataView, () => {
                let resources = d3Select(visualBuilder.element).selectAll(".task-resource").nodes();
                let returnResource = grep(resources);

                expect(returnResource.length).toEqual(resources.length);
                done();
            });
        });
        it("Verify task labels have tooltips", (done) => {
            defaultDataViewBuilder.valuesTaskTypeResource.forEach(x => x[1] = (x[1] + " ").repeat(5).trim());
            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate,
            ]);

            fixDataViewDateValuesAggregation(dataView);

            visualBuilder.updateRenderTimeout(dataView, () => {
                let taskLabelsInDom = d3Select(visualBuilder.element).selectAll(".label title").nodes();
                let taskLabels = d3Select(visualBuilder.element).selectAll(".label").data() as Task[];
                let tasks: PrimitiveValue[] | undefined = dataView.categorical?.categories?.[0].values;

                if (tasks) {
                    for (let i = 0; i < tasks.length; i++) {
                        expect(taskLabels[i].name).toEqual((taskLabelsInDom[i] as HTMLElement).textContent);
                        expect(tasks[i]).toEqual((taskLabelsInDom[i] as HTMLElement).textContent);
                    }
                }

                done();
            });
        });

        it("Verify tooltips have extra information", (done) => {
            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate,
                VisualData.ColumnExtraInformation,
                VisualData.ColumnResource]);

            fixDataViewDateValuesAggregation(dataView);

            visualBuilder.updateRenderTimeout(dataView, () => {
                let tasks = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                let index = 0;
                for (let task of tasks) {
                    for (let tooltipInfo of task.tooltipInfo) {
                        if (tooltipInfo.displayName === VisualData.ColumnExtraInformation) {
                            let value: string = tooltipInfo.value;

                            expect(value).toEqual(defaultDataViewBuilder.valuesExtraInformation[index++]);
                        }
                    }
                }

                done();
            });
        });

        it("Verify tooltips have extra information (date type)", (done) => {
            let host: IVisualHost = createVisualHost({});
            host.locale = host.locale || (<any>window.navigator).userLanguage || window.navigator["language"];
            let dateFormatter: IValueFormatter = valueFormatter.create({ format: undefined, cultureSelector: host.locale });

            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate,
                VisualData.ColumnExtraInformationDates]);

            visualBuilder.updateRenderTimeout(dataView, () => {
                let tasks = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                for (let task of tasks) {
                    for (let tooltipInfo of task.tooltipInfo) {
                        if (tooltipInfo.displayName === VisualData.ColumnExtraInformation) {
                            let value: string = tooltipInfo.value;

                            expect(value).toEqual(dateFormatter.format(task.start));
                        }
                    }
                }

                done();
            });
        });

        it("Verify tooltips include contract awarded information", (done) => {
            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate,
                VisualData.ColumnVergabeAn]);

            fixDataViewDateValuesAggregation(dataView);

            visualBuilder.updateRenderTimeout(dataView, () => {
                const tasks = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];

                tasks.forEach((task, index) => {
                    expect(task.contractAwarded).toEqual(defaultDataViewBuilder.valuesVergabeAn[index]);

                    const tooltipEntry = task.tooltipInfo.find(info => info.displayName === VisualData.ColumnVergabeAn);
                    if (defaultDataViewBuilder.valuesVergabeAn[index]) {
                        expect(tooltipEntry).toBeDefined();
                        expect(tooltipEntry?.value).toEqual(defaultDataViewBuilder.valuesVergabeAn[index]);
                    }
                });

                done();
            });
        });

        it("renders today line when today color is configured", (done) => {
            const currentDate = new Date();
            const todayStart = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate());

            defaultDataViewBuilder.valuesStartDate[0] = todayStart;
            defaultDataViewBuilder.valuesEndDate[0] = new Date(todayStart.getTime() + 24 * 60 * 60 * 1000);

            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate]);

            fixDataViewDateValuesAggregation(dataView);

            visualBuilder.updateRenderTimeout(dataView, () => {
                const todayLines = visualBuilder.mainElement.querySelectorAll("line.today-line");
                expect(todayLines.length).toBeGreaterThan(0);
                done();
            });
        });

        it("hides today line when no color is selected", (done) => {
            const currentDate = new Date();
            const todayStart = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate());

            defaultDataViewBuilder.valuesStartDate[0] = todayStart;
            defaultDataViewBuilder.valuesEndDate[0] = new Date(todayStart.getTime() + 24 * 60 * 60 * 1000);

            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate]);

            dataView.metadata = dataView.metadata || {};
            dataView.metadata.objects = dataView.metadata.objects || {};
            dataView.metadata.objects.dateType = Object.assign({}, dataView.metadata.objects.dateType, {
                todayColor: { solid: { color: null } }
            });

            fixDataViewDateValuesAggregation(dataView);

            visualBuilder.updateRenderTimeout(dataView, () => {
                const todayLines = visualBuilder.mainElement.querySelectorAll("line.today-line");
                expect(todayLines.length).toBe(0);
                done();
            });
        });

        it("hides today line when toggle disabled", (done) => {
            const currentDate = new Date();
            const todayStart = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate());

            defaultDataViewBuilder.valuesStartDate[0] = todayStart;
            defaultDataViewBuilder.valuesEndDate[0] = new Date(todayStart.getTime() + 24 * 60 * 60 * 1000);

            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate]);

            dataView.metadata = dataView.metadata || {};
            dataView.metadata.objects = dataView.metadata.objects || {};
            dataView.metadata.objects.dateType = Object.assign({}, dataView.metadata.objects.dateType, {
                todayLineVisible: false,
                todayColor: VisualBuilder.getSolidColorStructuralObject("#000000")
            });

            fixDataViewDateValuesAggregation(dataView);

            visualBuilder.updateRenderTimeout(dataView, () => {
                const todayLines = visualBuilder.mainElement.querySelectorAll("line.today-line");
                expect(todayLines.length).toBe(0);
                done();
            });
        });

        it("Verify tooltips have only string values", (done) => {
            const randomNumber = 134223;

            const task: any = {
                taskType: randomNumber,
                name: randomNumber,
                start: new Date(),
                end: new Date(),
                extraInformation: []
            };

            const formatters = {
                startDateFormatter: jasmine.createSpyObj("startDateFormatter", ["format"])
            };
            const localizationManager = visualBuilder.visualHost.createLocalizationManager();

            const tooltips = VisualClass.getTooltipInfo(task, formatters, localizationManager);
            tooltips
                .filter(t => t.value !== null && t.value !== undefined)
                .forEach(t => {
                    expect(typeof t.value).toBe("string");
                });
            done();
        });

        it("Verify sub tasks", (done) => {
            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate,
                VisualData.ColumnParent,
                VisualData.ColumnResource]);

            fixDataViewDateValuesAggregation(dataView);

            visualBuilder.updateRenderTimeout(dataView, () => {

                let tasks = d3Select(visualBuilder.element).selectAll(".task").data(),
                    uniqueParentsCount: number = getUniqueParentsCount(dataView, 3);

                expect(tasks.length).toEqual(defaultDataViewBuilder.valuesTaskTypeResource.length + uniqueParentsCount);

                let parentIndex: number = 4;
                let parentTask = visualBuilder.taskLabelsText[parentIndex];
                clickElement(parentTask);

                let childTaskMarginLeft: number = +(visualBuilder.taskLabelsText[++parentIndex].getAttribute("x") ?? 0);
                expect(childTaskMarginLeft).toEqual(VisualClass["DefaultValues"]["TaskLineWidth"]);

                childTaskMarginLeft = +(visualBuilder.taskLabelsText[++parentIndex].getAttribute("x") ?? 0);
                expect(childTaskMarginLeft).toEqual(VisualClass["DefaultValues"]["TaskLineWidth"]);

                done();
            });
        });

        it("Show collapse all arrow if parent is added", (done) => {
            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate,
                VisualData.ColumnResource,
                VisualData.ColumnParent
            ]);

            fixDataViewDateValuesAggregation(dataView);

            visualBuilder.updateRenderTimeout(dataView, () => {
                let collapseArrow = visualBuilder.collapseAllArrow;
                expect(collapseArrow).toBeDefined();

                dataView = defaultDataViewBuilder.getDataView([
                    VisualData.ColumnType,
                    VisualData.ColumnTask,
                    VisualData.ColumnStartDate,
                    VisualData.ColumnEndDate,
                    VisualData.ColumnResource
                ]);

                visualBuilder.updateRenderTimeout(dataView, () => {

                    let collapseArrow = visualBuilder.collapseAllArrow;
                    expect(collapseArrow.length).toBe(0);
                    done();
                });
            });
        });

        describe("Verify tooltips have info according 'parent' data", () => {
            function checkTasksHaveTooltipInfo(done: () => void) {
                fixDataViewDateValuesAggregation(dataView);
                visualBuilder.updateRenderTimeout(dataView, () => {
                    let tasks = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                    for (let task of tasks) {
                        if (!task.children) {
                            expect(task.tooltipInfo.length).not.toEqual(0);
                        }
                    }

                    done();
                });
            }

            it("With parent data", (done) => {
                dataView = defaultDataViewBuilder.getDataView([
                    VisualData.ColumnTask,
                    VisualData.ColumnStartDate,
                    VisualData.ColumnEndDate,
                    VisualData.ColumnParent]);

                fixDataViewDateValuesAggregation(dataView);

                checkTasksHaveTooltipInfo(done);
            });

            it("Without parent data", (done) => {
                dataView = defaultDataViewBuilder.getDataView([
                    VisualData.ColumnTask,
                    VisualData.ColumnStartDate,
                    VisualData.ColumnEndDate]);

                checkTasksHaveTooltipInfo(done);
            });
        });

        it("Verify Font Size set to default", (done) => {
            visualBuilder.updateRenderTimeout(dataView, () => {
                let element = d3Select(visualBuilder.element);
                let resources = element.selectAll(".task-resource").node();
                let labels = (element.selectAll(".label > .clickableArea").node() as HTMLElement).firstChild;

                expect((resources as SVGTextElement).style["font-size"]).toEqual("12px");
                expect((labels as SVGTextElement).style["font-size"]).toEqual("12px");
                done();
            });
        });

        for (let dateType in DateType) {
            it(`Verify date format (${dateType})`, ((dateType) => (done) => {
                switch (dateType) {
                    case DateType.Second:
                    case DateType.Minute:
                    case DateType.Hour:
                        defaultDataViewBuilder.valuesStartDate = VisualData.getRandomUniqueDates(
                            defaultDataViewBuilder.valuesTaskTypeResource.length,
                            new Date(2017, 7, 0),
                            new Date(2017, 7, 1)
                        );
                        dataView = defaultDataViewBuilder.getDataView();
                        break;
                }

                dataView.metadata.objects = { dateType: { type: dateType } };

                let host: IVisualHost = createVisualHost({});
                host.locale = host.locale || (<any>window.navigator).userLanguage || window.navigator["language"];

                let xAxisDateFormatter: IValueFormatter = valueFormatter.create({
                    format: VisualClass.DefaultValues.DateFormatStrings[dateType],
                    cultureSelector: host.locale
                });

                visualBuilder.updateRenderTimeout(dataView, () => {
                    const texts: SVGElement[]  = [];
                    visualBuilder.axisTicks.forEach((axisTick: HTMLElement) => {
                        const textsNode: NodeListOf<SVGElement> = axisTick.querySelectorAll("text");
                        textsNode.forEach((element: SVGElement) => {
                            texts.push(element);
                        })
                    });

                    texts.forEach((e) => {
                        let date: Date = new Date((<any>e).__data__);
                        expect(e.textContent).toEqual(xAxisDateFormatter.format(date));
                    });

                    done();
                });
            })(dateType));
        }

        it(`Verify milestone line is present in dom`, (done) => {
            let startDate: Date = new Date();
            let endDate: Date = new Date();

            startDate.setDate(startDate.getDate() - 30);
            endDate.setDate(endDate.getDate() + 30);

            defaultDataViewBuilder.valuesStartDate = VisualData.getRandomUniqueDates(
                defaultDataViewBuilder.valuesTaskTypeResource.length,
                startDate,
                endDate
            );
            dataView = defaultDataViewBuilder.getDataView();

            visualBuilder.updateRenderTimeout(dataView, () => {
                expect(visualBuilder.chartLine).toBeTruthy();

                done();
            });
        });

        it("Verify date format for culture which user have chosen", (done) => {
            let host: IVisualHost = createVisualHost({});
            host.locale = host.locale || (<any>window.navigator).userLanguage || window.navigator["language"];
            let dateFormatter: IValueFormatter = valueFormatter.create({ format: undefined, cultureSelector: host.locale });

            let formattedDates: string[] = [];
            for (let date of defaultDataViewBuilder.valuesStartDate) {
                formattedDates.push(dateFormatter.format(date));
            }

            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate]);

            fixDataViewDateValuesAggregation(dataView);

            visualBuilder.updateRenderTimeout(dataView, () => {
                let tasks = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                for (let task of tasks) {
                    for (let tooltipInfo of task.tooltipInfo) {
                        if (tooltipInfo.displayName === "Start Date") {
                            let value: string = tooltipInfo.value;
                            let idx: number = formattedDates.indexOf(value);

                            expect(value).toEqual(formattedDates[idx]);
                            formattedDates.splice(idx, 1);
                        }
                    }
                }

                done();
            });
        });

        it("Verify custom date format inside tooltip", (done) => {
            dataView.metadata.objects = {
                tooltipConfig: {
                    dateFormat: "MMMM dd,yyyy"
                }
            };

            visualBuilder.updateRenderTimeout(dataView, () => {
                let tasks = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];

                for (let task of tasks.filter(x => x.tooltipInfo)) {
                    const tooltipInfoArray = task.tooltipInfo;
                    tooltipInfoArray.forEach((tooltipInfo: VisualTooltipDataItem) => {
                        if (tooltipInfo.displayName === "Start Date") {
                            let value: string = tooltipInfo.value;

                            expect(value).toMatch(/([a-z].)\s([0-9]{2}),([0-9]{0,4})/);
                        }
                    });
                }

                done();
            });
        });

        it("Verify end date in tooltip", (done) => {
            let host: IVisualHost = createVisualHost({});
            host.locale = host.locale || (<any>window.navigator).userLanguage || window.navigator["language"];
            let dateFormatter: IValueFormatter = valueFormatter.create({ format: undefined, cultureSelector: host.locale });

            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate]);

            fixDataViewDateValuesAggregation(dataView);

            visualBuilder.updateRenderTimeout(dataView, () => {
                let tasks = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                for (let task of tasks) {
                    for (let tooltipInfo of task.tooltipInfo) {
                        if (tooltipInfo.displayName === "End Date") {
                            expect(tooltipInfo.value).toBe(dateFormatter.format(task.end));
                        }
                    }
                }

                done();
            });
        });

        it("Verify group tasks enabled", (done) => {
            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate]);

            dataView.metadata.objects = { general: { groupTasks: true } };

            fixDataViewDateValuesAggregation(dataView);

            visualBuilder.updateRenderTimeout(dataView, () => {
                let taskLinesText: HTMLElement[] = visualBuilder.taskLabelsText;
                let values = dataView.categorical?.categories?.[1].values;
                let taskGroups: HTMLElement[] = visualBuilder.tasksGroups;
                let tasks: Task[] = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];

                expect(values?.length).toBeGreaterThan(lodashUniq(values).length);
                expect(taskLinesText.length).toEqual(lodashUniq(values).length);

                taskGroups.forEach((taskGroup: HTMLElement, index: number) => {
                    const taskName: string | null = taskLinesText[index].children[0].textContent;
                    const tasksWithSameName = tasks.filter((task) => task.name === taskName);
                    expect(taskGroup.children.length).toBe(tasksWithSameName.length);
                });
                done();
            });
        });

        it("Verify group tasks enabled and then disabled", (done) => {

            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate]);

            dataView.metadata.objects = { general: { groupTasks: true } };

            fixDataViewDateValuesAggregation(dataView);

            visualBuilder.updateRenderTimeout(dataView, () => {
                dataView.metadata.objects = { general: { groupTasks: false } };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    let countOfTaskLines = visualBuilder.taskLabelsText.length;
                    let values = dataView.categorical?.categories?.[1].values;
                    let taskGroups: HTMLElement[] = visualBuilder.tasksGroups;

                    expect(countOfTaskLines).toEqual(values?.length ?? 0);
                    // in each row only one task - all the task re-rendered right
                    taskGroups.forEach((taskGroup: HTMLElement) => {
                        expect(taskGroup.children.length).toBe(1);
                    });
                    done();
                });
            });
        });

        it("Common task bar test with Grouping = OFF", (done) => {
            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate,
                VisualData.ColumnResource,
                VisualData.ColumnParent
            ]);

            dataView.metadata.objects = { general: { groupTasks: false } };
            fixDataViewDateValuesAggregation(dataView);

            visualBuilder.updateRenderTimeout(dataView, () => {
                let tasks: Task[] = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                let parentTasks: Task[] = tasks.filter((task: Task) => task.children);

                let parentIndex: number = getRandomNumber(0, parentTasks.length - 1),
                    parentTask = parentTasks[parentIndex],
                    parentTaskLabel = visualBuilder.taskLabelsText[parentTask.index];


                const minChildStart = lodashMinBy(parentTask.children, (t: Task) => t.start).start;
                const maxChildEnd = lodashMaxBy(parentTask.children, (t: Task) => t.end).end;
                const color = parentTask.children[0].color;


                clickElement(parentTaskLabel.parentElement);
                let collapsedTasksList = visualBuilder.instance["collapsedTasks"];
                dataView.metadata.objects = {
                    collapsedTasks: {
                        list: JSON.stringify(collapsedTasksList)
                    },
                    general: { groupTasks: false }
                };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    let taskGroups: HTMLElement[] = visualBuilder.tasksGroups;
                    let updatedTasks = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                    const updatedParentTask = updatedTasks[parentTask.index];

                    expect(updatedTasks.length).toBe(tasks.length - parentTask.children.length);
                    expect(taskGroups.length).toBe(tasks.length - parentTask.children.length);
                    expect(taskGroups[parentTask.index].children.length).toBe(1);

                    expect(updatedParentTask.start).toEqual(minChildStart);
                    expect(updatedParentTask.end).toEqual(maxChildEnd);
                    expect(updatedParentTask.color).toBe(color);
                    done();
                });
            });
        });

        it("Common task bar test with Grouping = ON", (done) => {
            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate,
                VisualData.ColumnResource,
                VisualData.ColumnParent
            ]);

            dataView.metadata.objects = { general: { groupTasks: true } };
            fixDataViewDateValuesAggregation(dataView);
            visualBuilder.updateRenderTimeout(dataView, () => {
                let tasks: Task[] = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                let parentTasks: Task[] = tasks.filter((task: Task) => task.children);

                let parentIndex: number = getRandomNumber(0, parentTasks.length - 1),
                    parentTask = parentTasks[parentIndex],
                    parentTaskLabel = visualBuilder.taskLabelsText[parentTask.index];

                const minChildStart = lodashMinBy(parentTask.children, (t: Task) => t.start).start;
                const maxChildEnd = lodashMaxBy(parentTask.children, (t: Task) => t.end).end;

                // Collapse
                clickElement(parentTaskLabel.parentElement);
                let collapsedTasksList = visualBuilder.instance["collapsedTasks"];
                dataView.metadata.objects = {
                    collapsedTasks: {
                        list: JSON.stringify(collapsedTasksList)
                    },
                    general: { groupTasks: true }
                };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    let taskGroups: HTMLElement[] = visualBuilder.tasksGroups;
                    let updatedTasks = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                    const updatedParentTask = updatedTasks[parentTask.index];
                    const tasksWithSameName = updatedTasks.filter((task) => task.name === parentTask.name);

                    // all params are similar because common task is not used with Grouping
                    expect(updatedParentTask.start).not.toBe(minChildStart);
                    expect(updatedParentTask.end).not.toBe(maxChildEnd);
                    expect(updatedParentTask.children).toBeNull();
                    expect(taskGroups[parentTask.index].children.length).toBe(tasksWithSameName.length);
                    done();
                });
            });
        });

        it("Milestone test", (done) => {
            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate,
                VisualData.ColumnResource,
                VisualData.ColumnParent,
                VisualData.ColumnMilestones
            ], true);

            const milestoneColumnIndex = 5;
            const categoriesColumn = dataView?.categorical?.categories?.[milestoneColumnIndex];
            const uniqueMilestoneTypes = lodashUniq(categoriesColumn?.values).filter(x => !!x);

            const randomColors = uniqueMilestoneTypes.map(() => getRandomHexColor());
            const randomTypes = uniqueMilestoneTypes.map(() => {
                const types = [MilestoneShape.Rhombus, MilestoneShape.Circle, MilestoneShape.Square];
                return types[Math.floor(getRandomNumber(0, types.length - 1))];
            });

            if (dataView.categorical?.categories && !dataView.categorical.categories[milestoneColumnIndex].objects) {
                dataView.categorical.categories[milestoneColumnIndex].objects = [];
            }

            categoriesColumn?.values.forEach((value: PrimitiveValue) => {
                let milestoneSettingsObject: { milestones: { fill: any; shapeType: string } } | null = null;
                if (value) {
                    const index = uniqueMilestoneTypes.indexOf(value);
                    milestoneSettingsObject = {
                        milestones: {
                            fill: VisualBuilder.getSolidColorStructuralObject(randomColors[index]),
                            shapeType: randomTypes[index]
                        }
                    };
                }

                dataView?.categorical?.categories?.[milestoneColumnIndex]?.objects?.push(milestoneSettingsObject);
            });

            // check for color and figure
            visualBuilder.updateRenderTimeout(dataView, () => {
                let tasks: Task[] = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                const taskWithMilestones = tasks.filter((task: Task) => task.Milestones?.length);
                const milestones: SVGElement[] = visualBuilder.milestones;

                expect(milestones.length).toBe(taskWithMilestones.length);

                // for each unique milestone type must be its own color and shapeType
                taskWithMilestones.forEach((task: Task) => {
                    task.Milestones?.forEach((milestone: Milestone) => {
                        const index = uniqueMilestoneTypes.indexOf(milestone.type);
                        const expectedColor = randomColors[index];
                        const actualColor = milestones[index].getAttribute("fill");
                        expect(actualColor).toBe(expectedColor);
                    });
                });

                done();
            });
        });

        it("Milestone before task start is preserved", (done) => {
            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate,
                VisualData.ColumnResource,
                VisualData.ColumnParent,
                VisualData.ColumnMilestones
            ], true);

            const categories = dataView.categorical?.categories ?? [];
            const milestoneCategory = categories.find(category => category.source.displayName === VisualData.ColumnMilestones);
            const startDateCategory = categories.find(category => category.source.displayName === VisualData.ColumnStartDate);

            const firstTaskIndex = 0;
            const startDate = startDateCategory?.values[firstTaskIndex] as Date;
            const milestoneDate = new Date((startDate?.getTime() ?? Date.now()) - (24 * 60 * 60 * 1000));

            if (milestoneCategory) {
                milestoneCategory.values[firstTaskIndex] = milestoneDate;
            }

            fixDataViewDateValuesAggregation(dataView);

            visualBuilder.updateRenderTimeout(dataView, () => {
                const tasks: Task[] = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                const taskWithMilestone = tasks.find(task => task.Milestones?.length);

                expect(taskWithMilestone).toBeDefined();
                const milestone = taskWithMilestone?.Milestones?.[0];
                expect(milestone).toBeDefined();
                if (milestone && taskWithMilestone) {
                    expect(milestone.start.getTime()).toBeLessThan(taskWithMilestone.start.getTime());
                }

                done();
            });
        });

        it("Common milestone test", (done) => {
            dataView = defaultDataViewBuilder.getDataView([
                VisualData.ColumnType,
                VisualData.ColumnTask,
                VisualData.ColumnStartDate,
                VisualData.ColumnEndDate,
                VisualData.ColumnResource,
                VisualData.ColumnParent,
                VisualData.ColumnMilestones
            ], true);
            visualBuilder.updateRenderTimeout(dataView, () => {
                const tasks: Task[] = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                const parentTasks: Task[] = tasks.filter((task: Task) => task.children);
                const oldMilestones: SVGElement[] = visualBuilder.milestones;

                let parentIndex: number = getRandomNumber(0, parentTasks.length - 1),
                    parentTask = parentTasks[parentIndex],
                    parentTaskLabel = visualBuilder.taskLabelsText[parentTask.index];

                // get uniq by date child milestones for current parent - they should be rendered on parent task bar
                const childMilestones = parentTask.children.map((childTask: Task) => {
                    if (childTask.Milestones?.length) {
                        return childTask.Milestones;
                    }
                });
                let mergedMilestone: Milestone[] | undefined  = parentTask.Milestones;
                childMilestones.forEach((milestoneArr) => {
                    mergedMilestone = mergedMilestone?.concat(milestoneArr);
                });

                const uniqDates = lodashUniqBy(mergedMilestone, "start");

                // Collapse
                clickElement(parentTaskLabel.parentNode);
                let collapsedTasksList = visualBuilder.instance["collapsedTasks"];
                dataView.metadata.objects = {
                    collapsedTasks: {
                        list: JSON.stringify(collapsedTasksList)
                    }
                };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    const updatedTasks: Task[] = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                    const updatedParentTask = updatedTasks[parentTask.index];
                    const milestones: SVGElement[] = visualBuilder.milestones;
                    const updatedTasksWithMilestones = updatedTasks.filter((t: Task) => t.Milestones?.length && t.index !== parentTask.index);

                    expect(milestones.length).toBe(oldMilestones.length - ((updatedParentTask.Milestones?.length ?? 0) - uniqDates.length));
                    expect(updatedParentTask.Milestones?.length).toBe(mergedMilestone?.length);

                    updatedTasksWithMilestones.forEach((t: Task) => {
                        expect(t.Milestones?.length).toBe(1);
                    });

                    done();
                });
            });
        });
    });

    describe("Selection", () => {
        describe("Single selection", () => {
            it("should highlight a proper data point after external filtering", () => {
                const selectionIds: MockISelectionId[] = [];
                let selectionIndex: number = -1;

                const customMockISelectionIdBuilder = new MockISelectionIdBuilder();
                customMockISelectionIdBuilder.createSelectionId = () => {
                    selectionIndex++;

                    if (selectionIds[selectionIndex]) {
                        return selectionIds[selectionIndex];
                    }

                    const selectionId: MockISelectionId = new MockISelectionId(`${selectionIndex}`);

                    selectionIds.push(selectionId);

                    return selectionId;
                };

                visualBuilder.visualHost.createSelectionIdBuilder = () => customMockISelectionIdBuilder;
                visualBuilder.updateFlushAllD3Transitions(dataView);

                // can't use lodash.deepCopy because we need to keep identity references
                const filteredDataView: DataView = {
                    ...dataView,
                    categorical: {
                        ...dataView.categorical,
                        categories: dataView.categorical?.categories?.map((category) => {
                            return {
                                ...category,
                                values: category.values.slice(0, 2)
                            };
                        })
                    }
                };

                selectionIndex = -1;

                visualBuilder.updateFlushAllD3Transitions(filteredDataView);

                clickElement(visualBuilder.tasks[0]);

                const selectedDataPoints: Task[] = getSelectedTasks(visualBuilder);

                selectionIndex = -1;

                visualBuilder.updateFlushAllD3Transitions(dataView);

                const selectedDataPointsAfterUpdateCall: Task[] = getSelectedTasks(visualBuilder);

                expect(selectedDataPoints.length).toBe(selectedDataPointsAfterUpdateCall.length);

                selectedDataPoints.forEach((selectedDataPoint: Task, index: number) => {
                    const selectedDataPointAfterUpdateCall: Task = selectedDataPointsAfterUpdateCall[index];

                    expect(selectedDataPoint.name).toBe(selectedDataPointAfterUpdateCall.name);
                    expect(selectedDataPoint.resource).toBe(selectedDataPointAfterUpdateCall.resource);
                    expect(selectedDataPoint.identity).toBe(selectedDataPointAfterUpdateCall.identity);
                });
            });

            function getSelectedTasks(visualBuilder: VisualBuilder): Task[] {
                return (visualBuilder.instance["interactivityService"]["selectableDataPoints"] as Task[])
                    .filter((task: Task) => task && task.selected);
            }
        });

        describe("Multi selection", () => {
            it("two data points should be selected", () => {
                visualBuilder.updateFlushAllD3Transitions(dataView);

                let firstGroup = visualBuilder.tasks[0];
                let secondGroup = visualBuilder.tasks[1];
                let thirdGroup = visualBuilder.tasks[2];

                clickElement(firstGroup);
                clickElement(secondGroup, true);

                expect(parseFloat(firstGroup.style.opacity)).toBe(1);
                expect(parseFloat(secondGroup.style.opacity)).toBe(1);
                expect(parseFloat(thirdGroup.style.opacity)).toBeLessThan(1);
            });
        });
    });

    describe("Format settings test", () => {
        describe("General", () => {
            it("Scroll to current time", (done) => {
                let todayDate = new Date();
                let startDate = new Date();
                let endDate = new Date();

                startDate.setDate(todayDate.getDate() - datesAmountForScroll);
                endDate.setDate(todayDate.getDate() + datesAmountForScroll);

                defaultDataViewBuilder.valuesStartDate = VisualData.getRandomUniqueDates(
                    defaultDataViewBuilder.valuesTaskTypeResource.length,
                    startDate,
                    endDate
                );
                dataView = defaultDataViewBuilder.getDataView();
                dataView.metadata.objects = {
                    general: {
                        scrollToCurrentTime: true
                    }
                };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    expect(visualBuilder.body.scrollLeft).not.toEqual(0);
                    done();
                });
            });

        });

        describe("Sub tasks", () => {
            beforeEach(() => {
                dataView = defaultDataViewBuilder.getDataView([
                    VisualData.ColumnTask,
                    VisualData.ColumnStartDate,
                    VisualData.ColumnEndDate,
                    VisualData.ColumnParent]);

                fixDataViewDateValuesAggregation(dataView);
            });

            it("sorting both parents and subtasks (tasks asc)", (done) => {
                dataView.metadata.columns[1].sort = 1; // 1 - ascending order

                visualBuilder.updateRenderTimeout(dataView, () => {
                    let tasks = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                    assertSortingOrderAsc(tasks);
                    done();
                });
            });

            it("sorting both parents and subtasks (tasks desc)", (done) => {
                dataView.metadata.columns[1].sort = 2; // 2 - descending order

                visualBuilder.updateRenderTimeout(dataView, () => {
                    let tasks = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                    assertSortingOrderDesc(tasks);
                    done();
                });
            });

            it("sorting both parents and subtasks (parent asc)", (done) => {

                dataView.metadata.columns[2].sort = 1; // 1 - ascending order

                visualBuilder.updateRenderTimeout(dataView, () => {
                    let tasks = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                    assertSortingOrderAsc(tasks);
                    done();
                });
            });

            it("sorting both parents and subtasks (parent desc)", (done) => {
                dataView.metadata.columns[2].sort = 2; // 2 - descending order

                visualBuilder.updateRenderTimeout(dataView, () => {
                    let tasks = d3Select(visualBuilder.element).selectAll(".task").data() as Task[];
                    assertSortingOrderDesc(tasks);
                    done();
                });
            });

            function assertSortingOrderAsc(tasks: Task[]) {
                let prevIndex: number = 0;

                for (let i = 1; i < tasks.length; ++i) {
                    if (!tasks[i].parent) {
                        expect(tasks[i].name >= tasks[prevIndex].name);
                        prevIndex = i;
                    }
                }
            }

            function assertSortingOrderDesc(tasks: Task[]) {
                let prevIndex: number = 0;

                for (let i = 1; i < tasks.length; ++i) {
                    if (!tasks[i].parent) {
                        expect(tasks[i].name <= tasks[prevIndex].name);
                        prevIndex = i;
                    }
                }
            }

            function getChildrenAndParents(tasks: Task[]) {
                let children: { [key: string]: Task[] } = {};
                let parents: Task[] = [];
                tasks.forEach((task) => {
                    if (task.parent) {
                        const parentName = task.parent.substring(0, task.parent.length - task.name.length - 1);

                        const parentTask: Task | undefined = tasks.find(t => t.name == parentName);

                        if (parentTask) {
                            if (!parents.find(parent => parent.name = parentTask.name)) {
                                parents.push(parentTask);
                            }

                            if (!children[parentTask.name]) {
                                children[parentTask.name] = [];
                            }

                            children[parentTask.name].push(task);
                        }
                    }
                });

                return { parents, children };
            }
        });

        describe("Data labels", () => {
            beforeEach(() => {
                dataView.metadata.objects = {
                    taskResource: {
                        show: true
                    }
                };

            });

            it("show", (done) => {
                visualBuilder.updateRenderTimeout(dataView, () => {
                    expect(visualBuilder.taskResources).toBeTruthy();

                    done();
                });
            });

            it("hide", (done) => {
                dataView.metadata.objects = {
                    taskResource: {
                        show: false
                    }
                };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    expect(visualBuilder.taskResources.length).toEqual(0);

                    done();
                });
            });

            it("color", (done) => {
                let color: string = getRandomHexColor();
                dataView.metadata.objects = {
                    taskResource: {
                        fill: VisualBuilder.getSolidColorStructuralObject(color)
                    }
                };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    visualBuilder.taskResources.forEach(e =>
                        assertColorsMatch(e.style.fill, color));

                    done();
                });
            });

            it("fontSize", (done) => {
                const fontSize: number = 10;
                dataView.metadata.objects = {
                    taskResource: {
                        fontSize
                    }
                };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    visualBuilder.taskResources.forEach(e => {
                        let fontSizeEl: string = e.style.fontSize;
                        fontSizeEl = fontSizeEl.substring(0, fontSizeEl.length - 2);

                        let fontSizePoint: string = PixelConverter.fromPoint(fontSize);
                        fontSizePoint = (+(fontSizePoint.substring(0, fontSizePoint.length - 2))).toFixed(4);

                       expect(fontSizeEl).toEqual(fontSizePoint);
                    });

                    done();
                });
            });

            it("position", (done) => {
                dataView.metadata.objects = {
                    taskResource: {
                        position: ResourceLabelPosition.Top
                    }
                };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    let taskRects: any[] = visualBuilder.taskRect;
                    visualBuilder.taskResources.forEach((e, i) => {
                        const text: string | null = e.textContent;
                        const taskResourcesX = +(e.getAttribute("x") ?? 0);
                        const taskResourcesY = +(e.getAttribute("y") ?? 0);
                        const taskRectX = taskRects[i].getBBox().x + VisualClass.RectRound;
                        const taskRectY = taskRects[i].getBBox().y;

                        if (text) {
                            expect(taskResourcesX.toFixed(2)).toBeCloseTo(taskRectX.toFixed(2), 1);
                            expect(taskResourcesY.toFixed(2)).toBeLessThan(taskRectY.toFixed(2));
                        }
                    });

                    done();
                });
            });

            it("fullText", (done) => {
                dataView.metadata.objects = {
                    taskResource: {
                        fullText: true
                    }
                };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    visualBuilder.taskResources.forEach(e =>
                        expect(e.textContent.indexOf("...")).toEqual(-1));

                    done();
                });
            });

            it("widthByTask", (done) => {
                dataView.metadata.objects = {
                    taskResource: {
                        position: ResourceLabelPosition.Top,
                        fullText: false,
                        widthByTask: true
                    }
                };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    let taskRects: HTMLElement[] = visualBuilder.taskRect;
                    visualBuilder.taskResources.forEach((e, i) => {
                        let labelElRawWidth: string = e.style.width;
                        let labelElWidth: number = +labelElRawWidth.substring(0, labelElRawWidth.length - 2);

                        let taskElRawWidth: string = taskRects[i].style.width;
                        let taskElWidth: number = +taskElRawWidth.substring(0, taskElRawWidth.length - 2);

                        expect(labelElWidth <= taskElWidth).toBeTruthy();
                    });

                    done();
                });
            });
        });

        describe("Task Settings", () => {
            it("color", (done) => {
                dataView = defaultDataViewBuilder.getDataView([
                    VisualData.ColumnTask,
                    VisualData.ColumnStartDate,
                    VisualData.ColumnEndDate,
                    VisualData.ColumnResource]);

                fixDataViewDateValuesAggregation(dataView);

                let color: string = getRandomHexColor();
                dataView.metadata.objects = {
                    taskConfig: {
                        fill: VisualBuilder.getSolidColorStructuralObject(color)
                    }
                };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    visualBuilder.taskLine.forEach(e =>
                        assertColorsMatch(e.style.fill, color));

                    done();
                });
            });

            it("height", (done) => {
                let height: number = 50;
                dataView.metadata.objects = {
                    taskConfig: {
                        height
                    }
                };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    visualBuilder.taskLine.forEach(e =>
                        expect(+(e.getAttribute("height") ?? 0)).toEqual(height));

                    done();
                });
            });
        });

        describe("Category Labels", () => {
            beforeEach(() => {
                dataView.metadata.objects = {
                    taskLabels: {
                        show: true
                    }
                };

            });

            it("show", (done) => {
                dataView.metadata.objects = {
                    taskLabels: {
                        show: true
                    }
                };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    const taskLabelsWidth: number = 110;
                    expect(visualBuilder.taskLabels).toBeTruthy();
                    expect(visualBuilder.taskLineRect[0].getAttribute("width")).toEqual(taskLabelsWidth.toString());
                    done();
                });
            });

            it("hide", (done) => {
                dataView.metadata.objects = {
                    taskLabels: {
                        show: false
                    }
                };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    expect(visualBuilder.taskLabels.length).toEqual(0);
                    expect(visualBuilder.taskLineRect[0].getAttribute("width")).toEqual("0");
                    done();
                });
            });

            it("color", (done) => {
                let color: string = getRandomHexColor();
                dataView.metadata.objects = {
                    taskLabels: {
                        fill: VisualBuilder.getSolidColorStructuralObject(color)
                    }
                };

                visualBuilder.updateRenderTimeout(dataView, () => {
                    visualBuilder.taskLabelsText.forEach(e =>
                        assertColorsMatch(e.getAttribute("fill"), color));

                    done();
                });
            });
        });

        describe("Gantt date types", () => {
            it("Today color", (done) => {
                let color: string = getRandomHexColor();
                dataView.metadata.objects = {
                    dateType: {
                        todayColor: VisualBuilder.getSolidColorStructuralObject(color)
                    }
                };

                checkColor(visualBuilder.chartLine, color, "stroke", done);
            });

            it("Axis color", (done) => {
                let color: string = getRandomHexColor();
                dataView.metadata.objects = {
                    dateType: {
                        axisColor: VisualBuilder.getSolidColorStructuralObject(color)
                    }
                };

                checkColor(visualBuilder.axisTicksLine, color, "stroke", done);
            });

            it("Axis text color", (done) => {
                let color: string = getRandomHexColor();
                dataView.metadata.objects = {
                    dateType: {
                        axisTextColor: VisualBuilder.getSolidColorStructuralObject(color)
                    }
                };

                checkColor(visualBuilder.axisTicksText, color, "fill", done);
            });

            function checkColor(
                elements: Element[] | SVGElement[],
                color: string,
                cssStyle: string,
                done: () => void
            ): void {
                visualBuilder.updateRenderTimeout(dataView, () => {
                    elements.forEach((e: SVGElement | HTMLElement) =>
                        assertColorsMatch(e.style.getPropertyValue(cssStyle), color));

                    done();
                });
            }
        });
    });

    describe("View Model tests", () => {
        it("Test result from enumeration", (done) => {
            const fontSize: number = 14;

            dataView.metadata.objects = {
                taskResource: {
                    show: true,
                    fill: { solid: { color: "#A3A3A3" } },
                    fontSize
                }
            };

            visualBuilder.updateRenderTimeout(dataView, () => {
                const taskResources = visualBuilder.taskResources;
                taskResources.forEach(resource => {
                    expect(resource).toBeDefined();
                    expect(resource.getAttribute("x")).not.toBeNaN();
                    expect(resource.getAttribute("y")).not.toBeNaN();
                    expect(resource.style.fill).toBe("rgb(163, 163, 163)");

                    let fontSizeEl: string = resource.style.fontSize;
                    fontSizeEl = fontSizeEl.substring(0, fontSizeEl.length - 2);

                    let fontSizePoint: string = PixelConverter.fromPoint(fontSize);
                    fontSizePoint = (+(fontSizePoint.substring(0, fontSizePoint.length - 2))).toFixed(4);

                    expect(fontSizeEl).toEqual(fontSizePoint);
                });
                done();
            });
        });
    });

    describe("Capabilities tests", () => {
        it("all items having displayName should have displayNameKey property", () => {
            const jsonData = require("../capabilities.json");

            let objectsChecker: Function = (obj: any) => {
                for (let property in obj) {
                    let value: any = obj[property];

                    if (value.displayName) {
                        expect(value.displayNameKey).toBeDefined();
                    }

                    if (typeof value === "object") {
                        objectsChecker(value);
                    }
                }
            };

            objectsChecker(jsonData);
        });
    });

    describe("High contrast mode", () => {
        const backgroundColor: string = "#000000";
        const foregroundColor: string = "#ff00ff";

        let taskRect: HTMLElement[],
            taskLineRect: HTMLElement[],
            axisTicksText: SVGElement[],
            axisTicksLine: SVGElement[],
            taskLabels: HTMLElement[],
            chartLine: HTMLElement[];

        beforeEach(() => {
            visualBuilder.visualHost.colorPalette.isHighContrast = true;

            visualBuilder.visualHost.colorPalette.background = { value: backgroundColor };
            visualBuilder.visualHost.colorPalette.foreground = { value: foregroundColor };

            taskRect = visualBuilder.taskRect;
            taskLineRect = visualBuilder.taskLineRect;

            axisTicksLine = visualBuilder.axisTicksLine;
            axisTicksText = visualBuilder.axisTicksLine;
            taskLabels = visualBuilder.taskLabels;
            chartLine = visualBuilder.chartLine;
        });

        it("should use high contrast mode colors", (done) => {
            visualBuilder.updateRenderTimeout(dataView, () => {
                expect(isColorAppliedToElements(chartLine, foregroundColor, "fill"));
                expect(isColorAppliedToElements(axisTicksLine, foregroundColor, "stroke"));
                expect(isColorAppliedToElements(axisTicksText, foregroundColor, "fill"));
                expect(isColorAppliedToElements(taskLabels, foregroundColor, "fill"));
                done();
            });
        });

        it("axis color and categories background should be taken from theme color", (done) => {
            visualBuilder.updateRenderTimeout(dataView, () => {
                expect(isColorAppliedToElements(taskLineRect, backgroundColor, "fill"));
                done();
            });
        });

        it("should apply high contrast colors to task rects", (done) => {
            visualBuilder.updateRenderTimeout(dataView, () => {
                expect(isColorAppliedToElements(taskRect, foregroundColor, "fill"));
                expect(isColorAppliedToElements(taskRect, foregroundColor, "stroke"));
                done();
            });
        });
    });

    describe("IsDateValid test", () => {
        it("test for valid Date", () => {
            let validDate = new Date();
            expect(isValidDate(validDate)).toBeTruthy();

            validDate = new Date(13425);
            expect(isValidDate(validDate)).toBeTruthy();
        });

        it("test for invalid Date", () => {
            const validDate = new Date("Hello");
            expect(isValidDate(validDate)).toBeFalsy();
        });
    });

    describe("Highlight test", () => {
        const defaultOpacity: string = DefaultOpacity.toString();
        const dimmedOpacity: string = DimmedOpacity.toString();

        it("Highlights property should not be received", (done) => {
            visualBuilder.updateRenderTimeout(dataView, () => {
                expect(dataView.categorical?.values?.some(value => value.highlights != null && value.highlights.length > 0)).toBe(false);

                const tasks: HTMLElement[] = visualBuilder.tasks;

                tasks.forEach((task: HTMLElement) => {
                    expect(task.style.opacity).toBe(defaultOpacity);
                });

                done();
            });
        });

        it("Elements should be highlighted", (done) => {
            const dataViewWithHighLighted: DataView = defaultDataViewBuilder.getDataViewWithHighlights();
            visualBuilder.updateRenderTimeout(dataViewWithHighLighted, () => {
                expect(dataViewWithHighLighted.categorical?.values?.some(value => value.highlights != null && value.highlights.length > 0)).toBe(true);

                let highlightedCount: number = 0;
                let nonHighlightedCount: number = 0;
                const expectedHighlightedCount: number = 1;

                const tasks: HTMLElement[] = visualBuilder.tasks;

                tasks.forEach((task: HTMLElement) => {
                    const opacity: string = task?.style?.opacity;
                    if (opacity === defaultOpacity)
                        highlightedCount++;
                    if (opacity === dimmedOpacity)
                        nonHighlightedCount++;
                });

                const expectedNonHighlightedCount: number = tasks.length - expectedHighlightedCount;
                expect(highlightedCount).toBe(expectedHighlightedCount);
                expect(nonHighlightedCount).toBe(expectedNonHighlightedCount);

                done();
            });
        });
    });

    describe("PersistProperties test", () => {

        const collapsedTasksUpdateIDs: string = "collapsedTasksUpdateIDs";

        it("Synchronous one task", (done) => {
            const newId = crypto?.randomUUID() || Math.random().toString();

            visualBuilder.instance[collapsedTasksUpdateIDs] = [newId];

            dataView.metadata.objects = {
                collapsedTasksUpdateId: {
                    value: newId
                }
            };

            visualBuilder.updateRenderTimeout(dataView, () => {
                expect(visualBuilder.instance[collapsedTasksUpdateIDs].length).toBe(0);
                done();
            });
        });

        it("Synchronous multiple tasks", (done) => {
            const collapsedTasksUpdateIDsRandom : string[] = []

            for (let count = 0; count < 3; count++) {
                const newId = crypto?.randomUUID() || Math.random().toString();
                collapsedTasksUpdateIDsRandom.push(newId);
            }

            visualBuilder.instance[collapsedTasksUpdateIDs] = collapsedTasksUpdateIDsRandom;

            const objects1 = {
                collapsedTasksUpdateId: {
                    value: collapsedTasksUpdateIDsRandom[0]
                }
            };

            const objects2 = {
                collapsedTasksUpdateId: {
                    value: collapsedTasksUpdateIDsRandom[1]
                }
            };

            const objects3 = {
                collapsedTasksUpdateId: {
                    value: collapsedTasksUpdateIDsRandom[2]
                }
            };


            dataView.metadata.objects = objects1;
            visualBuilder.update(dataView);
            expect(visualBuilder.instance[collapsedTasksUpdateIDs].length).toBe(2);


            dataView.metadata.objects = objects2;
            visualBuilder.update(dataView);
            expect(visualBuilder.instance[collapsedTasksUpdateIDs].length).toBe(1);


            dataView.metadata.objects = objects3;
            visualBuilder.update(dataView);
            expect(visualBuilder.instance[collapsedTasksUpdateIDs].length).toBe(0);

            done();
        });

        it("Asynchronous multiple tasks", async () => {
            const collapsedTasksUpdateIDsRandom : string[] = []

            for (let count = 0; count < 3; count++) {
                const newId = crypto?.randomUUID() || Math.random().toString();
                collapsedTasksUpdateIDsRandom.push(newId);
            }

            visualBuilder.instance[collapsedTasksUpdateIDs] = collapsedTasksUpdateIDsRandom;

            const objects1 = {
                collapsedTasksUpdateId: {
                    value: collapsedTasksUpdateIDsRandom[0]
                }
            };

            const objects2 = {
                collapsedTasksUpdateId: {
                    value: collapsedTasksUpdateIDsRandom[1]
                }
            };

            const objects3 = {
                collapsedTasksUpdateId: {
                    value: collapsedTasksUpdateIDsRandom[2]
                }
            };


            const promise1 = new Promise((resolve)=> {
                setTimeout(() => {
                    dataView.metadata.objects = objects1;
                    visualBuilder.update(dataView);
                    resolve(visualBuilder.instance[collapsedTasksUpdateIDs].includes(collapsedTasksUpdateIDsRandom[0]));
                },
                1_000);
            });

            const promise2 = new Promise((resolve)=> {
                setTimeout(() => {
                    dataView.metadata.objects = objects2;
                    visualBuilder.update(dataView);
                    resolve(visualBuilder.instance[collapsedTasksUpdateIDs].includes(collapsedTasksUpdateIDsRandom[1]));
                },
                    2_000);
            });

            const promise3 = new Promise((resolve)=> {
                setTimeout(() => {
                    dataView.metadata.objects = objects3;
                    visualBuilder.update(dataView);
                    resolve(visualBuilder.instance[collapsedTasksUpdateIDs].includes(collapsedTasksUpdateIDsRandom[2]));
                },
                    3_000);
            });

            const result = await Promise.all([promise1, promise2, promise3]);

            expect(result.filter(result => result === true).length).toBe(0);
        });
    });
});
