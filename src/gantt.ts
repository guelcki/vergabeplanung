/*
 *  Power BI Visualizations
 *
 *  Original work:
 *    Copyright (c) 2018 Microsoft Corporation, power-bi-gantt
 *
 *  Extensions and modifications:
 *    Copyright (c) 2025 Timo Gülck
 *
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the "Software"), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

import "./../style/gantt.less";

import {select as d3Select, Selection as d3Selection} from "d3-selection";
import {ScaleTime as timeScale} from "d3-scale";
import {nest as d3Nest} from "d3-collection";
import "d3-transition";

//lodash
import lodashIsEmpty from "lodash.isempty";
import lodashMin from "lodash.min";
import lodashMinBy from "lodash.minby";
import lodashMax from "lodash.max";
import lodashMaxBy from "lodash.maxby";
import lodashGroupBy from "lodash.groupby";
import lodashClone from "lodash.clone";
import lodashUniqBy from "lodash.uniqby";
import {Dictionary as lodashDictionary} from "lodash";

import powerbi from "powerbi-visuals-api";

// powerbi.extensibility.utils.svg
import * as SVGUtil from "powerbi-visuals-utils-svgutils";

// powerbi.extensibility.utils.type
import {pixelConverter as PixelConverter, valueType} from "powerbi-visuals-utils-typeutils";

// powerbi.extensibility.utils.formatting
import {textMeasurementService, valueFormatter as ValueFormatter} from "powerbi-visuals-utils-formattingutils";

// powerbi.extensibility.utils.interactivity
import {
    interactivityBaseService as interactivityService,
    interactivitySelectionService
} from "powerbi-visuals-utils-interactivityutils";

// powerbi.extensibility.utils.tooltip
import {
    createTooltipServiceWrapper,
    ITooltipServiceWrapper,
    TooltipEnabledDataPoint
} from "powerbi-visuals-utils-tooltiputils";

// powerbi.extensibility.utils.color
import {ColorHelper} from "powerbi-visuals-utils-colorutils";

// powerbi.extensibility.utils.chart
import {
    axis as AxisHelper,
    axisInterfaces,
    axisScale
} from "powerbi-visuals-utils-chartutils";

// behavior
import {Behavior, BehaviorOptions} from "./behavior";
import {
    ExtraInformation,
    GanttCalculateScaleAndDomainOptions,
    GanttChartFormatters,
    GanttViewModel,
    GroupedTask,
    Line,
    Milestone,
    MilestoneData,
    MilestoneDataPoint,
    MilestonePath,
    Task,
    TaskTypeMetadata,
    TaskTypes,
    TimeCriticalSegment
} from "./interfaces";
import {GanttColumns} from "./columns";
import {
    drawCircle,
    drawDiamond,
    drawNotRoundedRectByPath,
    drawRectangle,
    isValidDate
} from "./utils";
import {drawCollapseButton, drawExpandButton, drawMinusButton, drawPlusButton} from "./drawButtons";
import {TextProperties} from "powerbi-visuals-utils-formattingutils/lib/src/interfaces";

import {FormattingSettingsService} from "powerbi-visuals-utils-formattingmodel";
import {DateTypeCardSettings, GanttChartSettingsModel} from "./ganttChartSettingsModels";
import {DateType, GanttRole, LabelForDate, MilestoneShape, ResourceLabelPosition} from "./enums";

// d3
type Selection<T1, T2 = T1> = d3Selection<any, T1, any, T2>;

// powerbi
import DataView = powerbi.DataView;
import IViewport = powerbi.IViewport;
import SortDirection = powerbi.SortDirection;
import DataViewValueColumn = powerbi.DataViewValueColumn;
import DataViewValueColumns = powerbi.DataViewValueColumns;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import PrimitiveValue = powerbi.PrimitiveValue;

import DataViewObjectPropertyIdentifier = powerbi.DataViewObjectPropertyIdentifier;


import VisualObjectInstancesToPersist = powerbi.VisualObjectInstancesToPersist;

import IColorPalette = powerbi.extensibility.IColorPalette;
import ILocalizationManager = powerbi.extensibility.ILocalizationManager;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
// powerbi.visuals
import ISelectionIdBuilder = powerbi.visuals.ISelectionIdBuilder;
// powerbi.extensibility.visual
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
// powerbi.extensibility.utils.svg
import SVGManipulations = SVGUtil.manipulation;
import ClassAndSelector = SVGUtil.CssConstants.ClassAndSelector;
import createClassAndSelector = SVGUtil.CssConstants.createClassAndSelector;
import IMargin = SVGUtil.IMargin;
// powerbi.extensibility.utils.type
import PrimitiveType = valueType.PrimitiveType;
import ValueType = valueType.ValueType;
// powerbi.extensibility.utils.formatting
import IValueFormatter = ValueFormatter.IValueFormatter;
// powerbi.extensibility.utils.interactivity
import appendClearCatcher = interactivityService.appendClearCatcher;
import IInteractivityService = interactivityService.IInteractivityService;
import createInteractivityService = interactivitySelectionService.createInteractivitySelectionService;
// powerbi.extensibility.utils.chart
import IAxisProperties = axisInterfaces.IAxisProperties;

const ScrollMargin: number = 100;
const MillisecondsInASecond: number = 1000;
const MillisecondsInAMinute: number = 60 * MillisecondsInASecond;
const MillisecondsInAHour: number = 60 * MillisecondsInAMinute;
const MillisecondsInADay: number = 24 * MillisecondsInAHour;
const MillisecondsInWeek: number = 7 * MillisecondsInADay;
const MillisecondsInAMonth: number = 30 * MillisecondsInADay;
const MillisecondsInAYear: number = 365 * MillisecondsInADay;
const MillisecondsInAQuarter: number = MillisecondsInAYear / 4;
const PaddingTasks: number = 5;
const DefaultChartLineHeight = 40;
const TaskColumnName: string = "Task";
const ParentColumnName: string = "Parent";
export class SortingOptions {
    isCustomSortingNeeded: boolean;
    sortingDirection: SortDirection;
}



export class Gantt implements IVisual {
    private static ClassName: ClassAndSelector = createClassAndSelector("gantt");
    private static Chart: ClassAndSelector = createClassAndSelector("chart");
    private static ChartLine: ClassAndSelector = createClassAndSelector("chart-line");
    private static Body: ClassAndSelector = createClassAndSelector("gantt-body");
    private static AxisGroup: ClassAndSelector = createClassAndSelector("axis");
    private static Domain: ClassAndSelector = createClassAndSelector("domain");
    private static AxisTick: ClassAndSelector = createClassAndSelector("tick");
    private static Tasks: ClassAndSelector = createClassAndSelector("tasks");
    private static TaskGroup: ClassAndSelector = createClassAndSelector("task-group");
    private static SingleTask: ClassAndSelector = createClassAndSelector("task");
    private static TaskRect: ClassAndSelector = createClassAndSelector("task-rect");
    private static TaskMilestone: ClassAndSelector = createClassAndSelector("task-milestone");
    private static TaskResource: ClassAndSelector = createClassAndSelector("task-resource");
    private static TaskLabels: ClassAndSelector = createClassAndSelector("task-labels");
    private static TaskLines: ClassAndSelector = createClassAndSelector("task-lines");
    private static TaskLinesRect: ClassAndSelector = createClassAndSelector("task-lines-rect");
    private static TaskTopLine: ClassAndSelector = createClassAndSelector("task-top-line");
    private static CollapseAll: ClassAndSelector = createClassAndSelector("collapse-all");
    private static CollapseAllArrow: ClassAndSelector = createClassAndSelector("collapse-all-arrow");
    private static Label: ClassAndSelector = createClassAndSelector("label");
    private static ClickableArea: ClassAndSelector = createClassAndSelector("clickableArea");

    private viewport: IViewport;
    private colors: IColorPalette;
    private colorHelper: ColorHelper;

    private textProperties: TextProperties = {
        fontFamily: "wf_segoe-ui_normal",
        fontSize: PixelConverter.toString(9),
    };

    private static MilestonesPropertyIdentifier: DataViewObjectPropertyIdentifier = {
        objectName: "milestones",
        propertyName: "fill"
    };

    private static TaskResourcePropertyIdentifier: DataViewObjectPropertyIdentifier = {
        objectName: "taskResource",
        propertyName: "show"
    };

    private static CollapsedTasksPropertyIdentifier: DataViewObjectPropertyIdentifier = {
        objectName: "collapsedTasks",
        propertyName: "list"
    };

    private static CollapsedTasksUpdateIdPropertyIdentifier: DataViewObjectPropertyIdentifier = {
        objectName: "collapsedTasksUpdateId",
        propertyName: "value"
    };

    private static TaskConfigPropertyIdentifier: DataViewObjectPropertyIdentifier = {
        objectName: "taskConfig",
        propertyName: "fill"
    };

    public static DefaultValues = {
        AxisTickSize: 6,
        BarMargin: 2,
        ResourceWidth: 100,
        TaskColor: "#838383",
        SignalColor: "#FF341D",
        MilestoneFillColor: "#FFFA94",
        TaskLineColor: "#ccc",
        CollapseAllColor: "#000",
        PlusMinusColor: "#5F6B6D",
        CollapseAllTextColor: "#aaa",
        MilestoneLineColor: "#ccc",
        TaskCategoryLabelsRectColor: "#fafafa",
        TaskLineWidth: 15,
        IconMargin: 12,
        IconHeight: 16,
        IconWidth: 15,
        ChildTaskLeftMargin: 25,
        ParentTaskLeftMargin: 0,
        DefaultDateType: DateType.Month,
        DateFormatStrings: {
            Second: "HH:mm:ss",
            Minute: "HH:mm",
            Hour: "HH:mm (dd)",
            Day: "MMM dd",
            Week: "MMM dd",
            Month: "MMM yyyy",
            Quarter: "MMM yyyy",
            Year: "yyyy"
        }
    };

    private static DefaultGraphicWidthPercentage: number = 0.78;
    private static ResourceLabelDefaultDivisionCoefficient: number = 1.5;
    private static DefaultTicksLength: number = 50;
    private static DefaultDuration: number = 250;
    private static TaskLineCoordinateX: number = 15;
    private static AxisLabelClip: number = 40;
    private static AxisLabelStrokeWidth: number = 1;
    private static AxisTopMargin: number = 6;
    private static CollapseAllLeftShift: number = 7.5;
    private static BarHeightMargin: number = 5;
    private static ChartLineHeightDivider: number = 4;
    private static ResourceWidthPadding: number = 10;
    private static TaskLabelsMarginTop: number = 15;
    private static MinTasks: number = 1;
    private static ChartLineProportion: number = 1.5;
    private static MilestoneTop: number = 0;
    private static DividerForCalculatingPadding: number = 4;
    private static LabelTopOffsetForPadding: number = 0.5;
    private static DividerForCalculatingCenter: number = 2;
    private static SubtasksLeftMargin: number = 10;
    public static RectRound: number = 7;

    private static TimeScale: timeScale<any, any>;
    private xAxisProperties: IAxisProperties;

    private static get DefaultMargin(): IMargin {
        return {
            top: 50,
            right: 40,
            bottom: 40,
            left: 10
        };
    }

    private formattingSettings: GanttChartSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    private hasHighlights: boolean;

    private margin: IMargin = Gantt.DefaultMargin;

    private body: Selection<any>;
    private ganttSvg: Selection<any>;
    private viewModel: GanttViewModel;
    private collapseAllGroup: Selection<any>;
    private axisGroup: Selection<any>;
    private chartGroup: Selection<any>;
    private taskGroup: Selection<any>;
    private lineGroup: Selection<any>;
    private lineGroupWrapper: Selection<any>;
    private clearCatcher: Selection<any>;
    private ganttDiv: Selection<any>;
    private behavior: Behavior;
    private interactivityService: IInteractivityService<Task>;
    private eventService: IVisualEventService;
    private tooltipServiceWrapper: ITooltipServiceWrapper;
    private host: IVisualHost;
    private localizationManager: ILocalizationManager;
    private isInteractiveChart: boolean = false;
    private groupTasksPrevValue: boolean = false;
    private collapsedTasks: string[] = [];
    private collapseAllFlag: "data-is-collapsed";
    private parentLabelOffset: number = 5;
    private groupLabelSize: number = 25;
    private secondExpandAllIconOffset: number = 7;
    private hasNotNullableDates: boolean = false;

    private collapsedTasksUpdateIDs: string[] = [];

    constructor(options: VisualConstructorOptions) {
        this.init(options);
    }

    private init(options: VisualConstructorOptions): void {
        this.host = options.host;
        this.localizationManager = this.host.createLocalizationManager();
        this.formattingSettingsService = new FormattingSettingsService(this.localizationManager);
        this.colors = options.host.colorPalette;
        this.colorHelper = new ColorHelper(this.colors);
        this.body = d3Select(options.element);
        this.tooltipServiceWrapper = createTooltipServiceWrapper(this.host.tooltipService, options.element);
        this.behavior = new Behavior();
        this.interactivityService = createInteractivityService(this.host);
        this.eventService = options.host.eventService;

        this.createViewport();
    }

    /**
     * Create the viewport area of the gantt chart
     */
    private createViewport(): void {
        const axisBackgroundColor: string = this.colorHelper.getThemeColor();
        // create div container to the whole viewport area
        this.ganttDiv = this.body.append("div")
            .classed(Gantt.Body.className, true);

        // create container to the svg area
        this.ganttSvg = this.ganttDiv
            .append("svg")
            .classed(Gantt.ClassName.className, true);

        // create clear catcher
        this.clearCatcher = appendClearCatcher(this.ganttSvg);

        // create chart container
        this.chartGroup = this.ganttSvg
            .append("g")
            .classed(Gantt.Chart.className, true);

        // create tasks container
        this.taskGroup = this.chartGroup
            .append("g")
            .classed(Gantt.Tasks.className, true);

        // create axis container
        this.axisGroup = this.ganttSvg
            .append("g")
            .classed(Gantt.AxisGroup.className, true);
        this.axisGroup
            .append("rect")
            .attr("width", "100%")
            .attr("y", "-20")
            .attr("height", "40px")
            .attr("fill", axisBackgroundColor);

        // create task lines container
        this.lineGroup = this.ganttSvg
            .append("g")
            .classed(Gantt.TaskLines.className, true);

        this.lineGroupWrapper = this.lineGroup
            .append("rect")
            .classed(Gantt.TaskLinesRect.className, true)
            .attr("height", "100%")
            .attr("width", "0")
            .attr("fill", axisBackgroundColor)
            .attr("y", this.margin.top);

        this.lineGroup
            .append("rect")
            .classed(Gantt.TaskTopLine.className, true)
            .attr("width", "100%")
            .attr("height", 1)
            .attr("y", this.margin.top)
            .attr("fill", this.colorHelper.getHighContrastColor("foreground", Gantt.DefaultValues.TaskLineColor));

        this.collapseAllGroup = this.lineGroup
            .append("g")
            .classed(Gantt.CollapseAll.className, true);

        this.ganttDiv.on("scroll", (event) => {
            if (this.viewModel) {
                const taskLabelsWidth: number = this.viewModel.settings.taskLabelsCardSettings.show.value
                    ? this.viewModel.settings.taskLabelsCardSettings.width.value
                    : 0;

                const scrollTop: number = <number>event.target.scrollTop;
                const scrollLeft: number = <number>event.target.scrollLeft;

                this.axisGroup
                    .attr("transform", SVGManipulations.translate(taskLabelsWidth + this.margin.left + Gantt.SubtasksLeftMargin, Gantt.TaskLabelsMarginTop + scrollTop));
                this.lineGroup
                    .attr("transform", SVGManipulations.translate(scrollLeft, 0))
                    .attr("height", 20);
            }
        }, false);
    }

    /**
     * Clear the viewport area
     */
    private clearViewport(): void {
        this.ganttDiv
            .style("height", 0)
            .style("width", 0);

        this.axisGroup
            .selectAll(Gantt.AxisTick.selectorName)
            .remove();

        this.axisGroup
            .selectAll(Gantt.Domain.selectorName)
            .remove();

        this.collapseAllGroup
            .selectAll(Gantt.CollapseAll.selectorName)
            .remove();

        this.collapseAllGroup
            .selectAll(Gantt.CollapseAllArrow.selectorName)
            .remove();

        this.lineGroup
            .selectAll(Gantt.TaskLabels.selectorName)
            .remove();

        this.lineGroup
            .selectAll(Gantt.Label.selectorName)
            .remove();

        this.chartGroup
            .selectAll(Gantt.ChartLine.selectorName)
            .remove();

        this.chartGroup
            .selectAll(Gantt.TaskGroup.selectorName)
            .remove();

        this.chartGroup
            .selectAll(Gantt.SingleTask.selectorName)
            .remove();
    }

    /**
     * Update div container size to the whole viewport area
     */
    private updateChartSize(): void {
        this.ganttDiv
            .style("height", PixelConverter.toString(this.viewport.height))
            .style("width", PixelConverter.toString(this.viewport.width));
    }

    /**
     * Check if dataView has a given role
     * @param column The dataView headers
     * @param name The role to find
     */
    private static hasRole(column: DataViewMetadataColumn, name: string) {
        return column.roles && column.roles[name];
    }

    /**
     * Get the tooltip info (data display names & formatted values)
     * @param task All task attributes.
     * @param formatters Formatting options for gantt attributes.
     * @param localizationManager powerbi localization manager
     */
    public static getTooltipInfo(
        task: Task,
        formatters: GanttChartFormatters,
        localizationManager: ILocalizationManager): VisualTooltipDataItem[] {

        const tooltipDataArray: VisualTooltipDataItem[] = [];

        tooltipDataArray.push({
            displayName: localizationManager.getDisplayName("Role_Task"),
            value: task.name
        });

        if (task.start && !isNaN(task.start.getDate())) {
            tooltipDataArray.push({
                displayName: localizationManager.getDisplayName("Role_StartDate"),
                value: formatters.startDateFormatter.format(task.start)
            });
        }

        if (task.end && !isNaN(task.end.getDate())) {
            tooltipDataArray.push({
                displayName: localizationManager.getDisplayName("Role_EndDate"),
                value: formatters.startDateFormatter.format(task.end)
            });
        }

        if (task.resource) {
            tooltipDataArray.push({
                displayName: localizationManager.getDisplayName("Role_Resource"),
                value: task.resource
            });
        }

        if (task.tooltipInfo && task.tooltipInfo.length) {
            tooltipDataArray.push(...task.tooltipInfo);
        }

        task.extraInformation
            .map(tooltip => {
                if (typeof tooltip.value === "string") {
                    return tooltip;
                }

                const value: any = tooltip.value;

                if (isNaN(Date.parse(value)) || typeof value === "number") {
                    tooltip.value = value.toString();
                } else {
                    tooltip.value = formatters.startDateFormatter.format(value);
                }

                return tooltip;
            })
            .forEach(tooltip => tooltipDataArray.push(tooltip));

        tooltipDataArray
            .filter(x => x.value && typeof x.value !== "string")
            .forEach(tooltip => tooltip.value = tooltip.value.toString());

        return tooltipDataArray;
    }

    /**
    * Check if task has data for task
    * @param dataView
    */
    private static isChartHasTask(dataView: DataView): boolean {
        if (dataView?.metadata?.columns) {
            for (const column of dataView.metadata.columns) {
                if (Gantt.hasRole(column, GanttRole.Task)) {
                    return true;
                }
            }
        }

        return false;
    }

    /**
     * Returns the chart formatters
     * @param dataView The data Model
     * @param settings visual settings
     * @param cultureSelector The current user culture
     */
    private static getFormatters(
        dataView: DataView,
        settings: GanttChartSettingsModel,
        cultureSelector: string): GanttChartFormatters {

        if (!dataView?.metadata?.columns) {
            return null;
        }

        let dateFormat: string = "d";
        for (const dvColumn of dataView.metadata.columns) {
            if (Gantt.hasRole(dvColumn, GanttRole.StartDate)) {
                dateFormat = dvColumn.format;
            }
        }

        // Priority of using date format: Format from dvColumn -> Format by culture selector -> Custom Format
        if (cultureSelector) {
            dateFormat = null;
        }

        if (!settings.tooltipConfigCardSettings.dateFormat) {
            settings.tooltipConfigCardSettings.dateFormat.value = dateFormat;
        }

        if (settings.tooltipConfigCardSettings.dateFormat &&
            settings.tooltipConfigCardSettings.dateFormat.value !== dateFormat) {

            dateFormat = settings.tooltipConfigCardSettings.dateFormat.value;
        }

        return <GanttChartFormatters>{
            startDateFormatter: ValueFormatter.create({ format: dateFormat, cultureSelector })
        };
    }

    private static getSortingOptions(dataView: DataView): SortingOptions {
        const sortingOption: SortingOptions = new SortingOptions();

        dataView.metadata.columns.forEach(column => {
            if (column.roles && column.sort && (column.roles[ParentColumnName] || column.roles[TaskColumnName])) {
                sortingOption.isCustomSortingNeeded = true;
                sortingOption.sortingDirection = column.sort;

                return sortingOption;
            }
        });

        return sortingOption;
    }

    private static getUniqueMilestones(milestonesDataPoints: MilestoneDataPoint[]) {
        const milestonesWithoutDuplicates: {
            [name: string]: MilestoneDataPoint
        } = {};
        milestonesDataPoints.forEach((milestone: MilestoneDataPoint) => {
            if (milestone.name) {
                milestonesWithoutDuplicates[milestone.name] = milestone;
            }
        });

        return milestonesWithoutDuplicates;
    }

    private static createMilestones(
        dataView: DataView,
        host: IVisualHost): MilestoneData {
        let milestonesIndex = -1;
        for (const index in dataView.categorical.categories) {
            const category = dataView.categorical.categories[index];
            if (category.source.roles.Milestones) {
                milestonesIndex = +index;
            }
        }

        const milestoneData: MilestoneData = {
            dataPoints: []
        };
        const milestonesCategory = dataView.categorical.categories[milestonesIndex];
        const milestones: { value: PrimitiveValue, index: number }[] = [];

        if (milestonesCategory && milestonesCategory.values) {
            milestonesCategory.values.forEach((value: PrimitiveValue, index: number) => milestones.push({ value, index }));
            milestones.forEach((milestone) => {
                const milestoneObjects = milestonesCategory.objects?.[milestone.index];
                const selectionBuilder: ISelectionIdBuilder = host
                    .createSelectionIdBuilder()
                    .withCategory(milestonesCategory, milestone.index);

                const milestoneName: string = milestone.value instanceof Date
                    ? milestone.value.toISOString()
                    : String(milestone.value);

                const milestoneDataPoint: MilestoneDataPoint = {
                    name: milestoneName,
                    identity: selectionBuilder.createSelectionId(),
                    shapeType: milestoneObjects?.milestones?.shapeType ?
                        milestoneObjects.milestones.shapeType as string : MilestoneShape.Rhombus,
                    color: milestoneObjects?.milestones?.fill ?
                        (milestoneObjects.milestones as any).fill.solid.color : Gantt.DefaultValues.MilestoneFillColor
                };
                milestoneData.dataPoints.push(milestoneDataPoint);
            });
        }

        return milestoneData;
    }

    private static parseDateValue(value: PrimitiveValue): Date {
        if (value instanceof Date) {
            return isValidDate(value) ? value : null;
        }

        if (typeof value === "number") {
            const dateFromNumber: Date = new Date(value);
            return isValidDate(dateFromNumber) ? dateFromNumber : null;
        }

        if (typeof value === "string") {
            const germanDateMatch = value.trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
            if (germanDateMatch) {
                const [, day, month, year, hour = "0", minute = "0", second = "0"] = germanDateMatch;
                const parsedDate: Date = new Date(
                    Number(year),
                    Number(month) - 1,
                    Number(day),
                    Number(hour),
                    Number(minute),
                    Number(second));
                if (isValidDate(parsedDate)) {
                    return parsedDate;
                }
            }

            const directParse: Date = new Date(value);
            if (isValidDate(directParse)) {
                return directParse;
            }
        }

        return null;
    }

    private static ensureValidEndDate(startDate: Date, endDate: Date): Date {
        const fallbackEndDate = new Date(startDate.getTime() + MillisecondsInADay);

        if (!endDate || !isValidDate(endDate)) {
            return fallbackEndDate;
        }

        if (endDate.getTime() <= startDate.getTime()) {
            return fallbackEndDate;
        }

        return endDate;
    }

    private static parseMilestoneValue(
        milestoneValue: PrimitiveValue,
        fallbackStart: Date,
        fallbackEnd: Date): { type: string; date: Date; label: string } | null {

        if (milestoneValue === null
            || typeof milestoneValue === "undefined"
            || milestoneValue === "") {
            return null;
        }

        const parsedDate: Date = Gantt.parseDateValue(milestoneValue);
        const milestoneType: string = parsedDate
            ? parsedDate.toISOString()
            : String(milestoneValue);
        const milestoneLabel: string = parsedDate
            ? parsedDate.toLocaleDateString()
            : String(milestoneValue);

        const fallbackDate: Date = isValidDate(fallbackStart)
            ? fallbackStart
            : isValidDate(fallbackEnd)
                ? fallbackEnd
                : new Date();

        let milestoneDate: Date = parsedDate || fallbackDate;

        if (fallbackEnd && isValidDate(fallbackEnd) && milestoneDate > fallbackEnd) {
            milestoneDate = fallbackEnd;
        }

        return {
            type: milestoneType,
            date: milestoneDate,
            label: milestoneLabel
        };
    }

    private static parseTimeCriticalValue(value: PrimitiveValue): { isCritical: boolean | null; invalidValue?: string } {
        if (value === null || typeof value === "undefined") {
            return { isCritical: null };
        }

        if (typeof value === "boolean") {
            return { isCritical: value };
        }

        const rawValue: string = String(value);
        const normalized: string = rawValue.trim().toLowerCase();

        if (!normalized) {
            return { isCritical: null };
        }

        if (normalized === "ja") {
            return { isCritical: true };
        }

        if (normalized === "nein") {
            return { isCritical: false };
        }

        const displayValue = rawValue.trim() || rawValue;

        return {
            isCritical: null,
            invalidValue: displayValue
        };
    }

    private static calculateTimeCriticalSegments(start: Date, end: Date, milestone: Date | null): TimeCriticalSegment[] {
        if (!isValidDate(start) || !isValidDate(end) || end.getTime() <= start.getTime()) {
            return [];
        }

        if (!milestone || !isValidDate(milestone)) {
            return [];
        }

        if (milestone.getTime() > end.getTime()) {
            return [];
        }

        const segmentStart = milestone.getTime() <= start.getTime() ? start : milestone;

        if (segmentStart.getTime() >= end.getTime()) {
            return [];
        }

        return [{
            start: segmentStart,
            end
        }];
    }

    /**
     * Create task objects dataView
     * @param dataView The data Model.
     * @param formatters task attributes represented format.
     * @param taskColor Color of task
     * @param settings settings of visual
     * @param colors colors of groped tasks
     * @param host Host object
     * @param taskTypes
     * @param localizationManager powerbi localization manager
     * @param hasHighlights if any of the tasks has highlights
     */
    private static createTasks(
        dataView: DataView,
        taskTypes: TaskTypes,
        host: IVisualHost,
        formatters: GanttChartFormatters,
        colors: IColorPalette,
        settings: GanttChartSettingsModel,
        taskColor: string,
        localizationManager: ILocalizationManager,
        hasHighlights: boolean): Task[] {
        const categoricalValues: DataViewValueColumns = dataView?.categorical?.values;

        let tasks: Task[] = [];
        const addedParents: string[] = [];
        taskColor = taskColor || Gantt.DefaultValues.TaskColor;

        const values: GanttColumns<any> = GanttColumns.getCategoricalValues(dataView);

        if (!values.Task) {
            return tasks;
        }

        const colorHelper: ColorHelper = new ColorHelper(colors, Gantt.TaskConfigPropertyIdentifier);
        const groupValues: GanttColumns<DataViewValueColumn>[] = GanttColumns.getGroupedValueColumns(dataView);
        const sortingOptions: SortingOptions = Gantt.getSortingOptions(dataView);

        const collapsedTasks: string[] = JSON.parse(settings.collapsedTasksCardSettings.list.value);
        const invalidTimeCriticalEntries: { value: string; taskName: string }[] = [];

        values.Task.forEach((categoryValue: PrimitiveValue, index: number) => {
            const selectionBuilder: ISelectionIdBuilder = host
                .createSelectionIdBuilder()
                .withCategory(dataView.categorical.categories[0], index);

            const taskGroupAttributes = this.computeTaskGroupAttributes(taskColor, groupValues, values, index, taskTypes, selectionBuilder, colorHelper);
            const { color, taskType, endDate } = taskGroupAttributes;

            const {
                taskParentName,
                milestone,
                startDate,
                extraInformation,
                highlight,
                task
            } = this.createTask(values, index, hasHighlights, categoricalValues, color, categoryValue, endDate, taskType, selectionBuilder, invalidTimeCriticalEntries);

            if (taskParentName) {
                Gantt.addTaskToParentTask(
                    categoryValue,
                    task,
                    tasks,
                    taskParentName,
                    addedParents,
                    collapsedTasks,
                    milestone,
                    startDate,
                    highlight,
                    extraInformation,
                    selectionBuilder,
                );
            }

            tasks.push(task);
        });

        if (invalidTimeCriticalEntries.length) {
            const maxExamples = 5;
            const examples = invalidTimeCriticalEntries.slice(0, maxExamples)
                .map(({ taskName, value }) => {
                    const resolvedTaskName = taskName && String(taskName).trim().length ? String(taskName) : "(ohne Name)";
                    const sanitizedValue = typeof value === "string" ? value.trim() : value;
                    return `${resolvedTaskName}: "${sanitizedValue}"`;
                })
                .join("; ");
            const extraCount = invalidTimeCriticalEntries.length > maxExamples
                ? `; … (${invalidTimeCriticalEntries.length - maxExamples} weitere)`
                : "";
            host.displayWarningIcon(
                "Ungültiger Wert in \"Terminkritisch?\"",
                `Erlaubt sind \"ja\" oder \"nein\" (Groß-/Kleinschreibung egal). Ungültige Eingaben: ${examples}${extraCount}.`);
        }

        if (values.Parent) {
            tasks = Gantt.sortTasksWithParents(tasks, sortingOptions);
        }

        this.updateTaskDetails(tasks, settings, collapsedTasks);

        this.addTooltipInfoForCollapsedTasks(tasks, collapsedTasks, formatters, localizationManager, settings);

        return tasks;
    }

    private static updateTaskDetails(tasks: Task[], settings: GanttChartSettingsModel, collapsedTasks: string[]) {
        tasks.forEach(task => {
            if (task.children && task.children.length) {
                return;
            }

            task.end = Gantt.ensureValidEndDate(task.start, task.end);

            if (task.parent) {
                task.visibility = collapsedTasks.indexOf(task.parent) === -1;
            }
        });
    }

    private static addTooltipInfoForCollapsedTasks(tasks: Task[], collapsedTasks: string[], formatters: GanttChartFormatters, localizationManager: powerbi.extensibility.ILocalizationManager, settings: GanttChartSettingsModel) {
        tasks.forEach((task: Task) => {
            if (!task.children || collapsedTasks.includes(task.name)) {
                task.tooltipInfo = Gantt.getTooltipInfo(task, formatters, localizationManager);
                if (task.Milestones) {
                    task.Milestones.forEach((milestone) => {
                        const milestoneDate: Date = milestone.start || task.start;
                        const dateFormatted = formatters.startDateFormatter.format(milestoneDate);
                        const dateTypesSettings = settings.dateTypeCardSettings;
                        const milestoneLabel: string = milestone.label || milestone.type;
                        milestone.tooltipInfo = Gantt.getTooltipForMilestoneLine(dateFormatted, localizationManager, dateTypesSettings, [milestoneLabel], [milestone.category]);
                    });
                }
            }
        });
    }

    private static createTask(
        values: GanttColumns<any>,
        index: number,
        hasHighlights: boolean,
        categoricalValues: powerbi.DataViewValueColumns,
        color: string,
        categoryValue: string | number | Date | boolean,
        endDate: Date,
        taskType: TaskTypeMetadata,
        selectionBuilder: powerbi.visuals.ISelectionIdBuilder,
        invalidTimeCriticalEntries: { value: string; taskName: string }[]) {
        const resource: string = (values.Resource && values.Resource[index] as string) || "";
        const taskParentName: string = (values.Parent && values.Parent[index] as string) || null;
        const milestoneRaw: PrimitiveValue = values.Milestones ? values.Milestones[index] : null;

        const startDateValue: PrimitiveValue = values.StartDate ? values.StartDate[index] : null;
        const startDate: Date = Gantt.parseDateValue(startDateValue) || new Date(Date.now());

        const valueEndDate: PrimitiveValue = values.EndDate ? values.EndDate[index] : null;
        let resolvedEndDate: Date = endDate || Gantt.parseDateValue(valueEndDate);
        resolvedEndDate = Gantt.ensureValidEndDate(startDate, resolvedEndDate);

        const milestone = Gantt.parseMilestoneValue(milestoneRaw, startDate, resolvedEndDate);
        const milestoneOriginalDate = Gantt.parseDateValue(milestoneRaw);

        const extraInformation: ExtraInformation[] = this.getExtraInformationFromValues(values, index);

        let highlight: number = null;
        if (hasHighlights && categoricalValues) {
            const notNullIndex = categoricalValues.findIndex(value => value.highlights && value.values[index] != null);
            if (notNullIndex != -1) highlight = <number>categoricalValues[notNullIndex].highlights[index];
        }

        const timeCriticalEvaluation = Gantt.parseTimeCriticalValue(values.IsCritical ? values.IsCritical[index] : null);
        if (timeCriticalEvaluation.invalidValue) {
            invalidTimeCriticalEntries.push({
                value: timeCriticalEvaluation.invalidValue,
                taskName: categoryValue !== undefined && categoryValue !== null ? String(categoryValue) : ""
            });
        }

        const isTimeCritical: boolean = timeCriticalEvaluation.isCritical === true;
        const timeCriticalSegments: TimeCriticalSegment[] = isTimeCritical
            ? Gantt.calculateTimeCriticalSegments(startDate, resolvedEndDate, milestoneOriginalDate)
            : [];

        const task: Task = {
            color,
            resource,
            isTimeCritical,
            timeCriticalSegments,
            index: null,
            name: categoryValue as string,
            start: startDate,
            end: resolvedEndDate,
            parent: taskParentName,
            children: null,
            visibility: true,
            taskType: taskType && taskType.name,
            description: categoryValue as string,
            tooltipInfo: [],
            selected: false,
            identity: selectionBuilder.createSelectionId(),
            extraInformation,
            Milestones: milestone ? [{
                type: milestone.type,
                start: milestone.date,
                label: milestone.label,
                tooltipInfo: null,
                category: categoryValue as string
            }] : [],
            highlight: highlight !== null
        };

        return {taskParentName, milestone, startDate, extraInformation, highlight, task};
    }

    private static computeTaskGroupAttributes(
        taskColor: string,
        groupValues: GanttColumns<powerbi.DataViewValueColumn>[],
        values: GanttColumns<any>,
        index: number,
        taskTypes: TaskTypes,
        selectionBuilder: powerbi.visuals.ISelectionIdBuilder,
        colorHelper: ColorHelper) {
        let color: string = taskColor;
        let taskType: TaskTypeMetadata = null;
        let endDate: Date;

        if (groupValues) {
            groupValues.forEach((group: GanttColumns<DataViewValueColumn>) => {
                if (group.EndDate && group.EndDate.values[index] !== null) {
                    taskType = taskTypes?.types?.find((typeMeta: TaskTypeMetadata) => typeMeta.name === group.EndDate.source.groupName);

                    if (taskType) {
                        selectionBuilder.withCategory(taskType.selectionColumn, 0);
                        color = colorHelper.getColorForMeasure(taskType.columnGroup.objects, taskType.name);
                    }

                    const endDateValue: PrimitiveValue = group.EndDate.values[index];
                    endDate = Gantt.parseDateValue(endDateValue);
                } else if (group.Resource && group.Resource.values[index] !== null) {
                    taskType = taskTypes?.types?.find((typeMeta: TaskTypeMetadata) => typeMeta.name === group.Resource.source.groupName);
                    if (taskType) {
                        selectionBuilder.withCategory(taskType.selectionColumn, 0);
                        color = colorHelper.getColorForMeasure(taskType.columnGroup.objects, taskType.name);
                    }
                }
            });
        }

        return {
            color,
            taskType,
            endDate
        };
    }

    private static addTaskToParentTask(
        categoryValue: PrimitiveValue,
        task: Task,
        tasks: Task[],
        taskParentName: string,
        addedParents: string[],
        collapsedTasks: string[],
        milestone: { type: string; date: Date; label: string } | null,
        startDate: Date,
        highlight: number,
        extraInformation: ExtraInformation[],
        selectionBuilder: ISelectionIdBuilder,
    ) {
        if (addedParents.includes(taskParentName)) {
            const parentTask: Task = tasks.find(x => x.index === 0 && x.name === taskParentName);
            parentTask.children.push(task);
        } else {
            addedParents.push(taskParentName);

            const parentTask: Task = {
                index: 0,
                name: taskParentName,
                start: null,
                resource: null,
                isTimeCritical: false,
                timeCriticalSegments: [],
                end: null,
                parent: null,
                children: [task],
                visibility: true,
                taskType: null,
                description: null,
                color: null,
                tooltipInfo: null,
                extraInformation: collapsedTasks.includes(taskParentName) ? extraInformation : null,
                selected: false,
                identity: selectionBuilder.createSelectionId(),
                Milestones: [],
                highlight: highlight !== null
            };

            tasks.push(parentTask);
        }
    }

    private static getExtraInformationFromValues(values: GanttColumns<any>, taskIndex: number): ExtraInformation[] {
        const extraInformation: ExtraInformation[] = [];

        if (values.ExtraInformation) {
            const extraInformationKeys: any[] = Object.keys(values.ExtraInformation);
            for (const key of extraInformationKeys) {
                const value: string = values.ExtraInformation[key][taskIndex];
                if (value) {
                    extraInformation.push({
                        displayName: key,
                        value: value
                    });
                }
            }
        }

        return extraInformation;
    }

    public static sortTasksWithParents(tasks: Task[], sortingOptions: SortingOptions): Task[] {
        const sortingFunction = ((a: Task, b: Task) => {
            if (a.name < b.name) {
                return sortingOptions.sortingDirection === SortDirection.Ascending ? -1 : 1;
            }

            if (a.name > b.name) {
                return sortingOptions.sortingDirection === SortDirection.Ascending ? 1 : -1;
            }

            return 0;
        });

        if (sortingOptions.isCustomSortingNeeded) {
            tasks.sort(sortingFunction);
        }

        let index: number = 0;
        tasks.forEach(task => {
            if (!task.index && !task.parent) {
                task.index = index++;

                if (task.children) {
                    if (sortingOptions.isCustomSortingNeeded) {
                        task.children.sort(sortingFunction);
                    }

                    task.children.forEach(subtask => {
                        subtask.index = subtask.index === null ? index++ : subtask.index;
                    });
                }
            }
        });

        const resultTasks: Task[] = [];

        tasks.forEach((task) => {
            resultTasks[task.index] = task;
        });

        return resultTasks;
    }

    /**
     * Convert the dataView to view model
     * @param dataView The data Model
     * @param host Host object
     * @param colors Color palette
     * @param colorHelper powerbi color helper
     * @param localizationManager localization manager returns localized strings
     */
    public converter(
        dataView: DataView,
        host: IVisualHost,
        colors: IColorPalette,
        colorHelper: ColorHelper,
        localizationManager: ILocalizationManager): GanttViewModel {

        if (dataView?.categorical?.categories?.length === 0 || !Gantt.isChartHasTask(dataView)) {
            return null;
        }

        const settings: GanttChartSettingsModel = this.parseSettings(dataView, colorHelper);

        const taskTypes: TaskTypes = Gantt.getAllTasksTypes();

        this.hasHighlights = Gantt.hasHighlights(dataView);

        const formatters: GanttChartFormatters = Gantt.getFormatters(dataView, settings, host.locale || null);

        const isEndDateFilled: boolean = dataView.metadata.columns.findIndex(col => Gantt.hasRole(col, GanttRole.EndDate)) !== -1,
            isParentFilled: boolean = dataView.metadata.columns.findIndex(col => Gantt.hasRole(col, GanttRole.Parent)) !== -1,
            isResourcesFilled: boolean = dataView.metadata.columns.findIndex(col => Gantt.hasRole(col, GanttRole.Resource)) !== -1;

        const milestonesData: MilestoneData = Gantt.createMilestones(dataView, host);

        const taskColor: string = settings.taskConfigCardSettings.fill.value.value;

        const tasks: Task[] = Gantt.createTasks(dataView, taskTypes, host, formatters, colors, settings, taskColor, localizationManager, this.hasHighlights);

        return {
            dataView,
            settings,
            taskTypes,
            tasks,
            milestonesData,
            isEndDateFilled: isEndDateFilled,
            isParentFilled,
            isResourcesFilled
        };
    }

    public parseSettings(dataView: DataView, colorHelper: ColorHelper): GanttChartSettingsModel {

        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(GanttChartSettingsModel, dataView);
        const settings: GanttChartSettingsModel = this.formattingSettings;

        if (!colorHelper) {
            return settings;
        }

        if (colorHelper.isHighContrast) {
            settings.dateTypeCardSettings.axisColor.value.value = colorHelper.getHighContrastColor("foreground", settings.dateTypeCardSettings.axisColor.value.value);
            settings.dateTypeCardSettings.axisTextColor.value.value = colorHelper.getHighContrastColor("foreground", settings.dateTypeCardSettings.axisColor.value.value);
            settings.dateTypeCardSettings.todayColor.value.value = colorHelper.getHighContrastColor("foreground", settings.dateTypeCardSettings.todayColor.value.value);

            settings.taskConfigCardSettings.criticalFill.value.value = colorHelper.getHighContrastColor("foreground", settings.taskConfigCardSettings.criticalFill.value.value);
            settings.taskConfigCardSettings.fill.value.value = colorHelper.getHighContrastColor("foreground", settings.taskConfigCardSettings.fill.value.value);
            settings.taskLabelsCardSettings.fill.value.value = colorHelper.getHighContrastColor("foreground", settings.taskLabelsCardSettings.fill.value.value);
            settings.taskResourceCardSettings.fill.value.value = colorHelper.getHighContrastColor("foreground", settings.taskResourceCardSettings.fill.value.value);
        }

        return settings;
    }

    /**
    * Gets all unique types from the tasks array
    * @param dataView The data model.
    */
    private static getAllTasksTypes(): TaskTypes {
        return null;
    }

    private static hasHighlights(dataView: DataView): boolean {
        const values = (dataView?.categorical?.values?.length && dataView.categorical.values) || <DataViewValueColumns>[];
        const highlightsExist = values.some(({ highlights }) => highlights?.some(Number.isInteger));
        return !!highlightsExist;
    }

    private scaleAxisLength(axisLength: number): number {
        const fullScreenAxisLength: number = Gantt.DefaultGraphicWidthPercentage * this.viewport.width;
        if (axisLength < fullScreenAxisLength) {
            axisLength = fullScreenAxisLength;
        }

        return axisLength;
    }

    /**
    * Called on data change or resizing
    * @param options The visual option that contains the dataView and the viewport
    */
    public update(options: VisualUpdateOptions): void {
        if (!options || !options.dataViews || !options.dataViews[0]) {
            this.clearViewport();
            return;
        }

        const collapsedTasksUpdateId: any = options.dataViews[0].metadata?.objects?.collapsedTasksUpdateId?.value;

        if (this.collapsedTasksUpdateIDs.includes(collapsedTasksUpdateId)) {
            this.collapsedTasksUpdateIDs = this.collapsedTasksUpdateIDs.filter(id => id !== collapsedTasksUpdateId);
            return;
        }

        this.updateInternal(options);
    }

    private updateInternal(options: VisualUpdateOptions) : void {
        this.viewModel = this.converter(options.dataViews[0], this.host, this.colors, this.colorHelper, this.localizationManager);

        // for duplicated milestone types
        if (this.viewModel && this.viewModel.milestonesData) {
            const newMilestoneData: MilestoneData = this.viewModel.milestonesData;
            const milestonesWithoutDuplicates = Gantt.getUniqueMilestones(newMilestoneData.dataPoints);

            newMilestoneData.dataPoints.forEach((dataPoint: MilestoneDataPoint) => {
                if (dataPoint.name) {
                    const theSameUniqDataPoint: MilestoneDataPoint = milestonesWithoutDuplicates[dataPoint.name];
                    dataPoint.color = theSameUniqDataPoint.color;
                    dataPoint.shapeType = theSameUniqDataPoint.shapeType;
                }
            });

            this.viewModel.milestonesData = newMilestoneData;
        }

        if (!this.viewModel || !this.viewModel.tasks || this.viewModel.tasks.length <= 0) {
            this.clearViewport();
            return;
        }

        this.viewport = lodashClone(options.viewport);
        this.margin = Gantt.DefaultMargin;

        this.eventService.renderingStarted(options);

        this.render();

        this.eventService.renderingFinished(options);
    }

    private render(): void {
        const settings = this.viewModel.settings;

        this.updateChartSize();

        const visibleTasks = this.viewModel.tasks
            .filter((task: Task) => task.visibility);
        const tasks: Task[] = visibleTasks
            .map((task: Task, i: number) => {
                task.index = i;
                return task;
            });

        if (this.interactivityService) {
            this.interactivityService.applySelectionStateToData(tasks);
        }

        if (tasks.length < Gantt.MinTasks) {
            return;
        }

        this.collapsedTasks = JSON.parse(settings.collapsedTasksCardSettings.list.value);
        const groupTasks = this.viewModel.settings.generalCardSettings.groupTasks.value;
        const groupedTasks: GroupedTask[] = Gantt.getGroupTasks(tasks, groupTasks, this.collapsedTasks);
        // do something with task ids
        this.updateCommonTasks(groupedTasks);
        this.updateCommonMilestones(groupedTasks);

        const tasksWithDates: Task[] = groupedTasks
            .flatMap(t => t.tasks)
            .filter(task => task
                && task.start
                && task.end
                && isValidDate(task.start)
                && isValidDate(task.end));

        const minDateTask: Task = lodashMinBy(tasksWithDates, (t) => t && t.start);
        const maxDateTask: Task = lodashMaxBy(tasksWithDates, (t) => t && t.end);
        this.hasNotNullableDates = tasksWithDates.length > 0;

        let axisLength: number = 0;
        if (this.hasNotNullableDates) {
            const startDate: Date = minDateTask.start;
            let endDate: Date = maxDateTask.end;

            if (startDate.toString() === endDate.toString()) {
                endDate = new Date(endDate.valueOf() + (24 * 60 * 60 * 1000));
            }

            const dateTypeMilliseconds: number = Gantt.getDateType(DateType[settings.dateTypeCardSettings.type.value.value]);
            let ticks: number = Math.ceil(Math.round(endDate.valueOf() - startDate.valueOf()) / dateTypeMilliseconds);
            ticks = ticks < 2 ? 2 : ticks;

            axisLength = ticks * Gantt.DefaultTicksLength;
            axisLength = this.scaleAxisLength(axisLength);

            const viewportIn: IViewport = {
                height: this.viewport.height,
                width: axisLength
            };

            const xAxisProperties: IAxisProperties = this.calculateAxes(viewportIn, this.textProperties, startDate, endDate, ticks, false);
            this.xAxisProperties = xAxisProperties;
            Gantt.TimeScale = <timeScale<Date, Date>>xAxisProperties.scale;

            this.renderAxis(xAxisProperties);
        }

        axisLength = this.scaleAxisLength(axisLength);

        this.setDimension(groupedTasks, axisLength, settings);

        this.renderTasks(groupedTasks);
        this.updateTaskLabels(groupedTasks, settings.taskLabelsCardSettings.width.value);
        this.updateElementsPositions(this.margin);

        this.bindInteractivityService(tasks);
    }

    private bindInteractivityService(tasks: Task[]): void {
        if (this.interactivityService) {
            const behaviorOptions: BehaviorOptions = {
                clearCatcher: this.body,
                taskSelection: this.taskGroup.selectAll(Gantt.SingleTask.selectorName),
                subTasksCollapse: {
                    selection: this.body.selectAll(Gantt.ClickableArea.selectorName),
                    callback: this.subTasksCollapseCb.bind(this)
                },
                allSubtasksCollapse: {
                    selection: this.body
                        .selectAll(Gantt.CollapseAll.selectorName),
                    callback: this.subTasksCollapseAll.bind(this)
                },
                interactivityService: this.interactivityService,
                behavior: this.behavior,
                dataPoints: tasks
            };

            this.interactivityService.bind(behaviorOptions);

            this.behavior.renderSelection(this.hasHighlights);
        }
    }

    private static getDateType(dateType: DateType): number {
        switch (dateType) {
            case DateType.Second:
                return MillisecondsInASecond;

            case DateType.Minute:
                return MillisecondsInAMinute;

            case DateType.Hour:
                return MillisecondsInAHour;

            case DateType.Day:
                return MillisecondsInADay;

            case DateType.Week:
                return MillisecondsInWeek;

            case DateType.Month:
                return MillisecondsInAMonth;

            case DateType.Quarter:
                return MillisecondsInAQuarter;

            case DateType.Year:
                return MillisecondsInAYear;

            default:
                return MillisecondsInWeek;
        }
    }

    private calculateAxes(
        viewportIn: IViewport,
        textProperties: TextProperties,
        startDate: Date,
        endDate: Date,
        ticksCount: number,
        scrollbarVisible: boolean): IAxisProperties {

        const dataTypeDatetime: ValueType = ValueType.fromPrimitiveTypeAndCategory(PrimitiveType.Date);
        const category: DataViewMetadataColumn = {
            displayName: this.localizationManager.getDisplayName("Role_StartDate"),
            queryName: GanttRole.StartDate,
            type: dataTypeDatetime,
            index: 0
        };

        const visualOptions: GanttCalculateScaleAndDomainOptions = {
            viewport: viewportIn,
            margin: this.margin,
            forcedXDomain: [startDate, endDate],
            forceMerge: false,
            showCategoryAxisLabel: false,
            showValueAxisLabel: false,
            categoryAxisScaleType: axisScale.linear,
            valueAxisScaleType: null,
            valueAxisDisplayUnits: 0,
            categoryAxisDisplayUnits: 0,
            trimOrdinalDataOnOverflow: false,
            forcedTickCount: ticksCount
        };

        const width: number = viewportIn.width;
        const axes: IAxisProperties = this.calculateAxesProperties(viewportIn, visualOptions, category);
        axes.willLabelsFit = AxisHelper.LabelLayoutStrategy.willLabelsFit(
            axes,
            width,
            textMeasurementService.measureSvgTextWidth,
            textProperties);

        // If labels do not fit, and we are not scrolling, try word breaking
        axes.willLabelsWordBreak = (!axes.willLabelsFit && !scrollbarVisible) && AxisHelper.LabelLayoutStrategy.willLabelsWordBreak(
            axes, this.margin, width, textMeasurementService.measureSvgTextWidth,
            textMeasurementService.estimateSvgTextHeight, textMeasurementService.getTailoredTextOrDefault,
            textProperties);

        return axes;
    }

    private calculateAxesProperties(
        viewportIn: IViewport,
        options: GanttCalculateScaleAndDomainOptions,
        metaDataColumn: DataViewMetadataColumn): IAxisProperties {

        const dateType: DateType = DateType[this.viewModel.settings.dateTypeCardSettings.type.value.value];
        const cultureSelector: string = this.host.locale;
        const xAxisDateFormatter: IValueFormatter = ValueFormatter.create({
            format: Gantt.DefaultValues.DateFormatStrings[dateType],
            cultureSelector
        });
        const xAxisProperties: IAxisProperties = AxisHelper.createAxis({
            pixelSpan: viewportIn.width,
            dataDomain: options.forcedXDomain,
            metaDataColumn: metaDataColumn,
            formatString: Gantt.DefaultValues.DateFormatStrings[dateType],
            outerPadding: 5,
            isScalar: true,
            isVertical: false,
            forcedTickCount: options.forcedTickCount,
            useTickIntervalForDisplayUnits: true,
            isCategoryAxis: true,
            getValueFn: (index) => {
                return xAxisDateFormatter.format(new Date(index));
            },
            scaleType: options.categoryAxisScaleType,
            axisDisplayUnits: options.categoryAxisDisplayUnits,
        });

        xAxisProperties.axisLabel = metaDataColumn.displayName;
        return xAxisProperties;
    }

    private setDimension(
        groupedTasks: GroupedTask[],
        axisLength: number,
        settings: GanttChartSettingsModel): void {

        const fullResourceLabelMargin = groupedTasks.length * this.getResourceLabelTopMargin();
        let widthBeforeConversion = this.margin.left + settings.taskLabelsCardSettings.width.value + axisLength;

        if (settings.taskResourceCardSettings.show.value && settings.taskResourceCardSettings.position.value.value === ResourceLabelPosition.Right) {
            widthBeforeConversion += Gantt.DefaultValues.ResourceWidth;
        } else {
            widthBeforeConversion += Gantt.DefaultValues.ResourceWidth / 2;
        }

        const height = PixelConverter.toString(groupedTasks.length * (settings.taskConfigCardSettings.height.value || DefaultChartLineHeight) + this.margin.top + fullResourceLabelMargin);
        const width = PixelConverter.toString(widthBeforeConversion);

        this.ganttSvg
            .attr("height", height)
            .attr("width", width);
    }

    private static getGroupTasks(tasks: Task[], groupTasks: boolean, collapsedTasks: string[]): GroupedTask[] {
        if (groupTasks) {
            const groupedTasks: lodashDictionary<Task[]> = lodashGroupBy(tasks,
                x => (x.parent ? `${x.parent}.${x.name}` : x.name));

            const result: GroupedTask[] = [];
            const taskKeys: string[] = Object.keys(groupedTasks);
            const alreadyReviewedKeys: string[] = [];

            taskKeys.forEach((key: string) => {
                const isKeyAlreadyReviewed = alreadyReviewedKeys.includes(key);
                if (!isKeyAlreadyReviewed) {
                    let name: string = key;
                    if (groupedTasks[key] && groupedTasks[key].length && groupedTasks[key][0].parent && key.indexOf(groupedTasks[key][0].parent) !== -1) {
                        name = key.substr(groupedTasks[key][0].parent.length + 1, key.length);
                    }

                    // add current task
                    const taskRecord = <GroupedTask>{
                        name,
                        tasks: groupedTasks[key]
                    };
                    result.push(taskRecord);
                    alreadyReviewedKeys.push(key);

                    // see all the children and add them
                    groupedTasks[key].forEach((task: Task) => {
                        if (task.children && !collapsedTasks.includes(task.name)) {
                            task.children.forEach((childrenTask: Task) => {
                                const childrenFullName = `${name}.${childrenTask.name}`;
                                const isChildrenKeyAlreadyReviewed = alreadyReviewedKeys.includes(childrenFullName);

                                if (!isChildrenKeyAlreadyReviewed) {
                                    const childrenRecord = <GroupedTask>{
                                        name: childrenTask.name,
                                        tasks: groupedTasks[childrenFullName]
                                    };
                                    result.push(childrenRecord);
                                    alreadyReviewedKeys.push(childrenFullName);
                                }
                            });
                        }
                    });
                }
            });

            result.forEach((x, i) => {
                x.tasks.forEach(t => t.index = i);
                x.index = i;
            });

            return result;
        }

        return tasks.map(x => <GroupedTask>{
            name: x.name,
            index: x.index,
            tasks: [x]
        });
    }

    private renderAxis(xAxisProperties: IAxisProperties, duration: number = Gantt.DefaultDuration): void {
        const axisColor: string = this.viewModel.settings.dateTypeCardSettings.axisColor.value.value;
        const axisTextColor: string = this.viewModel.settings.dateTypeCardSettings.axisTextColor.value.value;

        const xAxis = xAxisProperties.axis;
        this.axisGroup.call(xAxis.tickSizeOuter(xAxisProperties.outerPadding));

        this.axisGroup
            .transition()
            .duration(duration)
            .call(xAxis);

        this.axisGroup
            .selectAll("path")
            .style("stroke", axisColor);

        this.axisGroup
            .selectAll(".tick line")
            .style("stroke", (timestamp: number) => this.setTickColor(timestamp, axisColor));

        this.axisGroup
            .selectAll(".tick text")
            .style("fill", (timestamp: number) => this.setTickColor(timestamp, axisTextColor));
    }

    private setTickColor(
        timestamp: number,
        defaultColor: string): string {
        return defaultColor;
    }

    /**
    * Update task labels and add its tooltips
    * @param tasks All tasks array
    * @param width The task label width
    */
    // eslint-disable-next-line max-lines-per-function
    private updateTaskLabels(
        tasks: GroupedTask[],
        width: number): void {

        let axisLabel: Selection<any>;
        const taskLabelsShow: boolean = this.viewModel.settings.taskLabelsCardSettings.show.value;
        const displayGridLines: boolean = this.viewModel.settings.generalCardSettings.displayGridLines.value;
        const taskLabelsColor: string = this.viewModel.settings.taskLabelsCardSettings.fill.value.value;
        const taskLabelsFontSize: number = this.viewModel.settings.taskLabelsCardSettings.fontSize.value;
        const taskLabelsWidth: number = this.viewModel.settings.taskLabelsCardSettings.width.value;
        const taskConfigHeight: number = this.viewModel.settings.taskConfigCardSettings.height.value || DefaultChartLineHeight;
        const categoriesAreaBackgroundColor: string = this.colorHelper.getThemeColor();
        const isHighContrast: boolean = this.colorHelper.isHighContrast;

        this.updateCollapseAllGroup(categoriesAreaBackgroundColor, taskLabelsShow);

        if (taskLabelsShow) {
            this.lineGroupWrapper
                .attr("width", taskLabelsWidth)
                .attr("fill", isHighContrast ? categoriesAreaBackgroundColor : Gantt.DefaultValues.TaskCategoryLabelsRectColor)
                .attr("stroke", this.colorHelper.getHighContrastColor("foreground", Gantt.DefaultValues.TaskLineColor))
                .attr("stroke-width", 1);

            this.lineGroup
                .selectAll(Gantt.Label.selectorName)
                .remove();

            axisLabel = this.lineGroup
                .selectAll(Gantt.Label.selectorName)
                .data(tasks);

            const axisLabelGroup = axisLabel
                .enter()
                .append("g")
                .merge(axisLabel);

            axisLabelGroup.classed(Gantt.Label.className, true)
                .attr("transform", (task: GroupedTask) => SVGManipulations.translate(0, this.margin.top + this.getTaskLabelCoordinateY(task.index)));

            const clickableArea = axisLabelGroup
                .append("g")
                .classed(Gantt.ClickableArea.className, true)
                .merge(axisLabelGroup);

            clickableArea
                .append("text")
                .attr("x", (task: GroupedTask) => (Gantt.TaskLineCoordinateX +
                    (task.tasks.every((task: Task) => !!task.parent)
                        ? Gantt.SubtasksLeftMargin
                        : (task.tasks[0].children && !!task.tasks[0].children.length) ? this.parentLabelOffset : 0)))
                .attr("class", (task: GroupedTask) => task.tasks[0].children ? "parent" : task.tasks[0].parent ? "child" : "normal-node")
                .attr("y", (task: GroupedTask) => (task.index + 0.5) * this.getResourceLabelTopMargin())
                .attr("fill", taskLabelsColor)
                .attr("stroke-width", Gantt.AxisLabelStrokeWidth)
                .style("font-size", PixelConverter.fromPoint(taskLabelsFontSize))
                .text((task: GroupedTask) => task.name)
                .call(AxisHelper.LabelLayoutStrategy.clip, width - Gantt.AxisLabelClip, textMeasurementService.svgEllipsis)
                .append("title")
                .text((task: GroupedTask) => task.name);

            const buttonSelection = clickableArea
                .filter((task: GroupedTask) => task.tasks[0].children && !!task.tasks[0].children.length)
                .append("svg")
                .attr("viewBox", "0 0 32 32")
                .attr("width", Gantt.DefaultValues.IconWidth)
                .attr("height", Gantt.DefaultValues.IconHeight)
                .attr("y", (task: GroupedTask) => (task.index + 0.5) * this.getResourceLabelTopMargin() - Gantt.DefaultValues.IconMargin)
                .attr("x", Gantt.DefaultValues.BarMargin);

            clickableArea
                .append("rect")
                .attr("width", 2 * Gantt.DefaultValues.IconWidth)
                .attr("height", 2 * Gantt.DefaultValues.IconWidth)
                .attr("y", (task: GroupedTask) => (task.index + 0.5) * this.getResourceLabelTopMargin() - Gantt.DefaultValues.IconMargin)
                .attr("x", Gantt.DefaultValues.BarMargin)
                .attr("fill", "transparent");

            const buttonPlusMinusColor = this.colorHelper.getHighContrastColor("foreground", Gantt.DefaultValues.PlusMinusColor);
            buttonSelection
                .each(function (task: GroupedTask) {
                    const element = d3Select(this);
                    if (!task.tasks[0].children[0].visibility) {
                        drawPlusButton(element, buttonPlusMinusColor);
                    } else {
                        drawMinusButton(element, buttonPlusMinusColor);
                    }
                });

            let parentTask: string = "";
            let childrenCount = 0;
            let currentChildrenIndex = 0;
            axisLabelGroup
                .append("rect")
                .attr("x", (task: GroupedTask) => {
                    const isGrouped = this.viewModel.settings.generalCardSettings.groupTasks.value;
                    const drawStandardMargin: boolean = !task.tasks[0].parent || task.tasks[0].parent && task.tasks[0].parent !== parentTask;
                    parentTask = task.tasks[0].parent ? task.tasks[0].parent : task.tasks[0].name;
                    if (task.tasks[0].children) {
                        parentTask = task.tasks[0].name;
                        childrenCount = isGrouped ? lodashUniqBy(task.tasks[0].children, "name").length : task.tasks[0].children.length;
                        currentChildrenIndex = 0;
                    }

                    if (task.tasks[0].parent === parentTask) {
                        currentChildrenIndex++;
                    }
                    const isLastChild = childrenCount && childrenCount === currentChildrenIndex;
                    return drawStandardMargin || isLastChild ? Gantt.DefaultValues.ParentTaskLeftMargin : Gantt.DefaultValues.ChildTaskLeftMargin;
                })
                .attr("y", (task: GroupedTask) => (task.index + 1) * this.getResourceLabelTopMargin() + (taskConfigHeight - this.viewModel.settings.taskLabelsCardSettings.fontSize.value) / 2)
                .attr("width", () => displayGridLines ? this.viewport.width : 0)
                .attr("height", 1)
                .attr("fill", this.colorHelper.getHighContrastColor("foreground", Gantt.DefaultValues.TaskLineColor));

            axisLabel
                .exit()
                .remove();
        } else {
            this.lineGroupWrapper
                .attr("width", 0)
                .attr("fill", "transparent");

            this.lineGroup
                .selectAll(Gantt.Label.selectorName)
                .remove();
        }
    }

    private updateCollapseAllGroup(categoriesAreaBackgroundColor: string, taskLabelShow: boolean) {
        this.collapseAllGroup
            .selectAll("svg")
            .remove();

        this.collapseAllGroup
            .selectAll("rect")
            .remove();

        this.collapseAllGroup
            .selectAll("text")
            .remove();

        if (this.viewModel.isParentFilled) {
            const categoryLabelsWidth: number = this.viewModel.settings.taskLabelsCardSettings.show.value
                ? this.viewModel.settings.taskLabelsCardSettings.width.value
                : 0;

            this.collapseAllGroup
                .append("rect")
                .attr("width", categoryLabelsWidth)
                .attr("height", 2 * Gantt.TaskLabelsMarginTop)
                .attr("fill", categoriesAreaBackgroundColor);

            const expandCollapseButton = this.collapseAllGroup
                .append("svg")
                .classed(Gantt.CollapseAllArrow.className, true)
                .attr("viewBox", "0 0 48 48")
                .attr("width", this.groupLabelSize)
                .attr("height", this.groupLabelSize)
                .attr("x", Gantt.CollapseAllLeftShift + this.xAxisProperties.outerPadding || 0)
                .attr("y", this.secondExpandAllIconOffset)
                .attr(this.collapseAllFlag, (this.collapsedTasks.length ? "1" : "0"));

            expandCollapseButton
                .append("rect")
                .attr("width", this.groupLabelSize)
                .attr("height", this.groupLabelSize)
                .attr("x", 0)
                .attr("y", this.secondExpandAllIconOffset)
                .attr("fill", "transparent");

            const buttonExpandCollapseColor = this.colorHelper.getHighContrastColor("foreground", Gantt.DefaultValues.CollapseAllColor);
            if (this.collapsedTasks.length) {
                drawExpandButton(expandCollapseButton, buttonExpandCollapseColor);
            } else {
                drawCollapseButton(expandCollapseButton, buttonExpandCollapseColor);
            }

            if (taskLabelShow) {
                this.collapseAllGroup
                    .append("text")
                    .attr("x", this.secondExpandAllIconOffset + this.groupLabelSize)
                    .attr("y", this.groupLabelSize)
                    .attr("font-size", "12px")
                    .attr("fill", this.colorHelper.getHighContrastColor("foreground", Gantt.DefaultValues.CollapseAllTextColor))
                    .text(this.collapsedTasks.length ? this.localizationManager.getDisplayName("Visual_Expand_All") : this.localizationManager.getDisplayName("Visual_Collapse_All"));
            }
        }
    }

    /**
     * callback for subtasks click event
     * @param taskClicked Grouped clicked task
     */
    private subTasksCollapseCb(taskClicked: GroupedTask): void {
        const taskIsChild: boolean = taskClicked.tasks[0].parent && !taskClicked.tasks[0].children;
        const taskWithoutParentAndChildren: boolean = !taskClicked.tasks[0].parent && !taskClicked.tasks[0].children;
        if (taskIsChild || taskWithoutParentAndChildren) {
            return;
        }

        const taskClickedParent: string = taskClicked.tasks[0].parent || taskClicked.tasks[0].name;
        this.viewModel.tasks.forEach((task: Task) => {
            if (task.parent === taskClickedParent &&
                task.parent.length >= taskClickedParent.length) {
                const index: number = this.collapsedTasks.indexOf(task.parent);
                if (task.visibility) {
                    this.collapsedTasks.push(task.parent);
                } else {
                    if (taskClickedParent === task.parent) {
                        this.collapsedTasks.splice(index, 1);
                    }
                }
            }
        });

        // eslint-disable-next-line
        const newId = crypto?.randomUUID() || Math.random().toString();
        this.collapsedTasksUpdateIDs.push(newId);

        this.setJsonFiltersValues(this.collapsedTasks, newId);
    }

    /**
     * callback for subtasks collapse all click event
     */
    private subTasksCollapseAll(): void {
        const collapsedAllSelector = this.collapseAllGroup.select(Gantt.CollapseAllArrow.selectorName);
        const isCollapsed: string = collapsedAllSelector.attr(this.collapseAllFlag);
        const buttonExpandCollapseColor = this.colorHelper.getHighContrastColor("foreground", Gantt.DefaultValues.CollapseAllColor);

        collapsedAllSelector.selectAll("path").remove();
        if (isCollapsed === "1") {
            this.collapsedTasks = [];
            collapsedAllSelector.attr(this.collapseAllFlag, "0");
            drawCollapseButton(collapsedAllSelector, buttonExpandCollapseColor);

        } else {
            collapsedAllSelector.attr(this.collapseAllFlag, "1");
            drawExpandButton(collapsedAllSelector, buttonExpandCollapseColor);
            this.viewModel.tasks.forEach((task: Task) => {
                if (task.parent) {
                    if (task.visibility) {
                        this.collapsedTasks.push(task.parent);
                    }
                }
            });
        }

        // eslint-disable-next-line
        const newId = crypto?.randomUUID() || Math.random().toString();
        this.collapsedTasksUpdateIDs.push(newId);

        this.setJsonFiltersValues(this.collapsedTasks, newId);
    }

    private setJsonFiltersValues(collapsedValues: string[], collapsedTasksUpdateId: string) {
        this.host.persistProperties(<VisualObjectInstancesToPersist>{
            merge: [{
                objectName: "collapsedTasks",
                selector: null,
                properties: {
                    list: JSON.stringify(collapsedValues)
                }
            }, {
                objectName: "collapsedTasksUpdateId",
                selector: null,
                properties: {
                    value: JSON.stringify(collapsedTasksUpdateId)
                }
            }]
        });
    }

    /**
     * Render tasks
     * @param groupedTasks Grouped tasks
     */
    private renderTasks(groupedTasks: GroupedTask[]): void {
        const taskConfigHeight: number = this.viewModel.settings.taskConfigCardSettings.height.value || DefaultChartLineHeight;
        const taskGroupSelection: Selection<any> = this.taskGroup
            .selectAll(Gantt.TaskGroup.selectorName)
            .data(groupedTasks);

        taskGroupSelection
            .exit()
            .remove();

        // render task group container
        const taskGroupSelectionMerged = taskGroupSelection
            .enter()
            .append("g")
            .merge(taskGroupSelection);

        taskGroupSelectionMerged.classed(Gantt.TaskGroup.className, true);

        const taskSelection: Selection<Task> = this.taskSelectionRectRender(taskGroupSelectionMerged);
        this.taskMainRectRender(taskSelection, taskConfigHeight);
        this.MilestonesRender(taskSelection, taskConfigHeight);
        this.taskResourceRender(taskSelection, taskConfigHeight);

        this.renderTooltip(taskSelection);
    }


    /**
     * Change task structure to be able for
     * Rendering common tasks when all the children of current parent are collapsed
     * used only the Grouping mode is OFF
     * @param groupedTasks Grouped tasks
     */
    private updateCommonTasks(groupedTasks: GroupedTask[]): void {
        if (!this.viewModel.settings.generalCardSettings.groupTasks.value) {
            groupedTasks.forEach((groupedTask: GroupedTask) => {
                const currentTaskName: string = groupedTask.name;
                if (this.collapsedTasks.includes(currentTaskName)) {
                    const firstTask: Task = groupedTask.tasks && groupedTask.tasks[0];
                    const tasks = groupedTask.tasks;
                    tasks.forEach((task: Task) => {
                        if (task.children) {
                            const childrenColors = task.children.map((child: Task) => child.color).filter((color) => color);
                            const minChildDateStart = lodashMin(task.children.map((child: Task) => child.start).filter((dateStart) => dateStart));
                            const maxChildDateEnd = lodashMax(task.children.map((child: Task) => child.end).filter((dateStart) => dateStart));
                            firstTask.color = !firstTask.color && task.children ? childrenColors[0] : firstTask.color;
                            firstTask.start = lodashMin([firstTask.start, minChildDateStart]);
                            firstTask.end = <any>lodashMax([firstTask.end, maxChildDateEnd]);
                        }
                    });

                    if (firstTask) {
                        firstTask.isTimeCritical = false;
                        firstTask.timeCriticalSegments = [];
                    }

                    groupedTask.tasks = firstTask && [firstTask] || [];
                }
            });
        }
    }

    /**
     * Change task structure to be able for
     * Rendering common milestone when all the children of current parent are collapsed
     * used only the Grouping mode is OFF
     * @param groupedTasks Grouped tasks
     */
    private updateCommonMilestones(groupedTasks: GroupedTask[]): void {
        groupedTasks.forEach((groupedTask: GroupedTask) => {
            groupedTask.tasks?.forEach((task: Task) => {
                if (task.children && task.children.length) {
                    task.Milestones = [];
                }
            });
        });
    }

    /**
     * Render task progress rect
     * @param taskGroupSelection Task Group Selection
     */
    private taskSelectionRectRender(taskGroupSelection: Selection<any>) {
        const taskSelection: Selection<Task> = taskGroupSelection
            .selectAll(Gantt.SingleTask.selectorName)
            .data((d: GroupedTask) => d.tasks);

        taskSelection
            .exit()
            .remove();

        const taskSelectionMerged = taskSelection
            .enter()
            .append("g")
            .merge(taskSelection);

        taskSelectionMerged.classed(Gantt.SingleTask.className, true);

        return taskSelectionMerged;
    }

    /**
     * @param task
     */
    private getTaskRectWidth(task: Task): number {
        if (!this.hasNotNullableDates) {
            return 0;
        }

        if (!task.end || isNaN(task.end.getTime())) {
            return 0;
        }

        return Gantt.taskDurationToWidth(task.start, task.end);
    }


    /**
     *
     * @param task
     * @param taskConfigHeight
     * @param barsRoundedCorners are bars with rounded corners
     */
    private drawTaskRect(task: Task, taskConfigHeight: number): string {
        const x = this.hasNotNullableDates ? Gantt.TimeScale(task.start) : 0,
            y = Gantt.getBarYCoordinate(task.index, taskConfigHeight) + (task.index + 1) * this.getResourceLabelTopMargin(),
            width = this.getTaskRectWidth(task),
            height = Gantt.getBarHeight(taskConfigHeight);

        return drawNotRoundedRectByPath(x, y, width, height);
    }

    private drawCriticalSegment(task: Task, segment: TimeCriticalSegment, taskConfigHeight: number): string {
        if (!this.hasNotNullableDates || !segment || !isValidDate(segment.start) || !isValidDate(segment.end)) {
            return "";
        }

        const width = Gantt.taskDurationToWidth(segment.start, segment.end);

        if (width <= 0) {
            return "";
        }

        const x = Gantt.TimeScale(segment.start);
        const y = Gantt.getBarYCoordinate(task.index, taskConfigHeight) + (task.index + 1) * this.getResourceLabelTopMargin();
        const height = Gantt.getBarHeight(taskConfigHeight);

        return drawNotRoundedRectByPath(x, y, width, height);
    }

    /**
     * Render task progress rect
     * @param taskSelection Task Selection
     * @param taskConfigHeight Task heights from settings
     */
    private taskMainRectRender(
        taskSelection: Selection<Task>,
        taskConfigHeight: number): void {
        const highContrastModeTaskRectStroke: number = 1;
        const signalColorSetting = this.viewModel?.settings?.taskConfigCardSettings?.criticalFill?.value?.value
            || Gantt.DefaultValues.SignalColor;
        const resolvedSignalColor = this.colorHelper.getHighContrastColor("foreground", signalColorSetting);

        const taskRect: Selection<Task> = taskSelection
            .selectAll(Gantt.TaskRect.selectorName)
            .data((d: Task) => [d]);

        const taskRectMerged = taskRect
            .enter()
            .append("path")
            .merge(taskRect);

        taskRectMerged.classed(Gantt.TaskRect.className, true);

        taskRectMerged
            .attr("d", (task: Task) => this.drawTaskRect(task, taskConfigHeight))
            .attr("width", (task: Task) => this.getTaskRectWidth(task))
            .style("fill", (task: Task) => {
                const fallbackColor = this.viewModel?.settings?.taskConfigCardSettings?.fill?.value?.value
                    || Gantt.DefaultValues.TaskColor;
                const baseColor = task.color || fallbackColor;
                return this.colorHelper.getHighContrastColor("foreground", baseColor);
            });

        if (this.colorHelper.isHighContrast) {
            taskRectMerged
                .style("stroke", (task: Task) => this.colorHelper.getHighContrastColor("foreground", task.color))
                .style("stroke-width", highContrastModeTaskRectStroke);
        }

        const criticalRectSelection: Selection<{ segment: TimeCriticalSegment; task: Task }, Task> = taskSelection
            .selectAll(".task-rect-critical")
            .data((task: Task) => (task.timeCriticalSegments || []).map((segment: TimeCriticalSegment) => ({ segment, task })));

        const criticalRectMerged: Selection<{ segment: TimeCriticalSegment; task: Task }, Task> = criticalRectSelection
            .enter()
            .append("path")
            .merge(criticalRectSelection);

        criticalRectMerged
            .classed("task-rect-critical", true)
            .attr("d", (data: { segment: TimeCriticalSegment; task: Task }) =>
                this.drawCriticalSegment(data.task, data.segment, taskConfigHeight))
            .style("fill", resolvedSignalColor);

        if (this.colorHelper.isHighContrast) {
            criticalRectMerged
                .style("stroke", resolvedSignalColor)
                .style("stroke-width", highContrastModeTaskRectStroke);
        } else {
            criticalRectMerged
                .style("stroke", null)
                .style("stroke-width", null);
        }

        criticalRectSelection
            .exit()
            .remove();

        taskRect
            .exit()
            .remove();
    }

    /**
     *
     * @param milestoneType milestone type
     */
    private getMilestoneColor(): string {
        const colorSetting = this.viewModel?.settings?.milestonesCardSettings?.fill?.value?.value
            || Gantt.DefaultValues.MilestoneFillColor;
        return this.colorHelper.getHighContrastColor("foreground", colorSetting);
    }

    private getMilestonePath(taskConfigHeight: number): string {
        let shape: string;
        const convertedHeight: number = Gantt.getBarHeight(taskConfigHeight);
        const shapeType = (this.viewModel?.settings?.milestonesCardSettings?.shapeType?.value?.value as MilestoneShape)
            || MilestoneShape.Rhombus;
        switch (shapeType) {
            case MilestoneShape.Rhombus:
                shape = drawDiamond(convertedHeight);
                break;
            case MilestoneShape.Square:
                shape = drawRectangle(convertedHeight);
                break;
            case MilestoneShape.Circle:
                shape = drawCircle(convertedHeight);
        }

        return shape;
    }

    /**
     * Render milestones
     * @param taskSelection Task Selection
     * @param taskConfigHeight Task heights from settings
     */
    private MilestonesRender(
        taskSelection: Selection<Task>,
        taskConfigHeight: number): void {
            const taskMilestones: Selection<any> = taskSelection
            .selectAll(Gantt.TaskMilestone.selectorName)
            .data((d: Task) => {
                if ((d.children && d.children.length) || lodashIsEmpty(d.Milestones) || !this.hasNotNullableDates) {
                    return [];
                }
                const nestedByDate = d3Nest().key((d: Milestone) => d.start.toDateString()).entries(d.Milestones);
                const updatedMilestones: MilestonePath[] = nestedByDate.map((nestedObj) => {
                    const oneDateMilestones = nestedObj.values;
                    // if there is 2 or more milestones for concrete date => draw only one milestone for concrete date, but with tooltip for all of them
                    const currentMilestone = [...oneDateMilestones].pop();
                    const allTooltipInfo = oneDateMilestones.map((milestone: MilestonePath) => milestone.tooltipInfo);
                    currentMilestone.tooltipInfo = allTooltipInfo.reduce((a, b) => a.concat(b), []);

                    return {
                        type: currentMilestone.type,
                        start: currentMilestone.start,
                        taskID: d.index,
                        tooltipInfo: currentMilestone.tooltipInfo,
                        label: currentMilestone.label
                    };
                });

                return [{
                    key: d.index, values: <MilestonePath[]>updatedMilestones
                }];
            });


        taskMilestones
            .exit()
            .remove();

        const taskMilestonesAppend = taskMilestones
            .enter()
            .append("g");

        const taskMilestonesMerged = taskMilestonesAppend
            .merge(taskMilestones);

        taskMilestonesMerged.classed(Gantt.TaskMilestone.className, true);

        const transformForMilestone = (id: number, start: Date) => {
            return SVGManipulations.translate(
                Gantt.TimeScale(start) - Gantt.getBarHeight(taskConfigHeight) / 4,
                Gantt.getBarYCoordinate(id, taskConfigHeight) + (id + 1) * this.getResourceLabelTopMargin()
            );
        };

        const taskMilestonesSelection = taskMilestonesMerged.selectAll("path");
        const taskMilestonesSelectionData = taskMilestonesSelection.data(milestonesData => <MilestonePath[]>milestonesData.values);

        // add milestones: for collapsed task may be several milestones of its children, in usual case - just 1 milestone
        const taskMilestonesSelectionAppend = taskMilestonesSelectionData.enter()
            .append("path");

        taskMilestonesSelectionData
            .exit()
            .remove();

        const taskMilestonesSelectionMerged = taskMilestonesSelectionAppend
            .merge(<any>taskMilestonesSelection);

        taskMilestonesSelectionMerged
            .attr("d", () => this.getMilestonePath(taskConfigHeight))
            .attr("transform", (data: MilestonePath) => transformForMilestone(data.taskID, data.start))
            .attr("fill", () => this.getMilestoneColor())
            .attr("stroke", Gantt.DefaultValues.TaskColor)
            .attr("stroke-width", 1);

        const milestoneLabelFormatter = ValueFormatter.create({
            format: "dd.MM.",
            cultureSelector: this.host?.locale || null
        });

        const milestoneLabels = taskMilestonesMerged
            .selectAll("text")
            .data(milestonesData => <MilestonePath[]>milestonesData.values);

        milestoneLabels
            .exit()
            .remove();

        const milestoneLabelsMerged = milestoneLabels
            .enter()
            .append("text")
            .merge(<any>milestoneLabels);

        const milestoneLabelFontSize = this.viewModel.settings.milestonesCardSettings.labelFontSize.value || 16;

        milestoneLabelsMerged
            .classed("milestone-label", true)
            .text((data: MilestonePath) => milestoneLabelFormatter.format(data.start))
            .attr("x", (data: MilestonePath) => Gantt.TimeScale(data.start) + Gantt.getBarHeight(taskConfigHeight) + 4)
            .attr("y", (data: MilestonePath) => Gantt.getBarYCoordinate(data.taskID, taskConfigHeight)
                + (data.taskID + 1) * this.getResourceLabelTopMargin()
                + Gantt.getBarHeight(taskConfigHeight) / 2)
            .attr("dominant-baseline", "middle")
            .style("font-size", `${milestoneLabelFontSize}px`)
            .style("font-weight", "700")
            .style("fill", this.colorHelper.getHighContrastColor(
                "foreground",
                this.viewModel.settings.milestonesCardSettings.labelColor.value.value || Gantt.DefaultValues.TaskColor
            ));

        taskMilestonesMerged.each(function () {
            const element = this as SVGGElement;
            const parentNode = element.parentNode;
            if (parentNode && parentNode.lastChild !== element) {
                parentNode.appendChild(element);
            }
        });

        this.renderTooltip(taskMilestonesSelectionMerged);
    }

    /**
     * Render task resource labels
     * @param taskSelection Task Selection
     * @param taskConfigHeight Task heights from settings
     */
    private taskResourceRender(
        taskSelection: Selection<Task>,
        taskConfigHeight: number): void {

        const groupTasks: boolean = this.viewModel.settings.generalCardSettings.groupTasks.value;
        let newLabelPosition: ResourceLabelPosition | null = null;
        if (groupTasks && !this.groupTasksPrevValue) {
            newLabelPosition = ResourceLabelPosition.Inside;
        }

        if (!groupTasks && this.groupTasksPrevValue) {
            newLabelPosition = ResourceLabelPosition.Right;
        }

        if (newLabelPosition) {
            this.host.persistProperties(<VisualObjectInstancesToPersist>{
                merge: [{
                    objectName: "taskResource",
                    selector: null,
                    properties: { position: newLabelPosition }
                }]
            });

            this.viewModel.settings.taskResourceCardSettings.position.value.value = newLabelPosition;
            newLabelPosition = null;
        }

        this.groupTasksPrevValue = groupTasks;

        const isResourcesFilled: boolean = this.viewModel.isResourcesFilled;
        const taskResourceShow: boolean = this.viewModel.settings.taskResourceCardSettings.show.value;
        const taskResourceColor: string = this.viewModel.settings.taskResourceCardSettings.fill.value.value;
        const taskResourceFontSize: number = this.viewModel.settings.taskResourceCardSettings.fontSize.value;
        const taskResourcePosition: ResourceLabelPosition = ResourceLabelPosition[this.viewModel.settings.taskResourceCardSettings.position.value.value];
        const taskResourceFullText: boolean = this.viewModel.settings.taskResourceCardSettings.fullText.value;
        const taskResourceWidthByTask: boolean = this.viewModel.settings.taskResourceCardSettings.widthByTask.value;
        const isGroupedByTaskName: boolean = this.viewModel.settings.generalCardSettings.groupTasks.value;

        if (isResourcesFilled && taskResourceShow) {
            const taskResource: Selection<Task> = taskSelection
                .selectAll(Gantt.TaskResource.selectorName)
                .data((d: Task) => [d]);

            const taskResourceMerged = taskResource
                .enter()
                .append("text")
                .merge(taskResource);

            taskResourceMerged.classed(Gantt.TaskResource.className, true);

            taskResourceMerged
                .attr("x", (task: Task) => this.getResourceLabelXCoordinate(task, taskConfigHeight, taskResourceFontSize, taskResourcePosition))
                .attr("y", (task: Task) => Gantt.getBarYCoordinate(task.index, taskConfigHeight)
                    + Gantt.getResourceLabelYOffset(taskConfigHeight, taskResourceFontSize, taskResourcePosition)
                    + (task.index + 1) * this.getResourceLabelTopMargin())
                .text((task: Task) => this.getTaskRectWidth(task) > 0 ? (task.resource || "") : "")
                .style("fill", taskResourceColor)
                .style("font-size", PixelConverter.fromPoint(taskResourceFontSize))
                .style("alignment-baseline", taskResourcePosition === ResourceLabelPosition.Inside ? "central" : "auto");

            const hasNotNullableDates: boolean = this.hasNotNullableDates;
            const defaultWidth: number = Gantt.DefaultValues.ResourceWidth - Gantt.ResourceWidthPadding;

            if (taskResourceWidthByTask) {
                taskResourceMerged
                    .each(function (task: Task) {
                        const width: number = hasNotNullableDates ? Gantt.taskDurationToWidth(task.start, task.end) : 0;
                        AxisHelper.LabelLayoutStrategy.clip(d3Select(this), width, textMeasurementService.svgEllipsis);
                    });
            } else if (isGroupedByTaskName) {
                taskResourceMerged
                    .each(function (task: Task, outerIndex: number) {
                        const sameRowNextTaskStart: Date = Gantt.getSameRowNextTaskStartDate(task, outerIndex, taskResourceMerged);

                        if (sameRowNextTaskStart) {
                            let width: number = 0;
                            if (hasNotNullableDates) {
                                const startDate: Date = taskResourcePosition === ResourceLabelPosition.Top ? task.start : task.end;
                                width = Gantt.taskDurationToWidth(startDate, sameRowNextTaskStart);
                            }

                            AxisHelper.LabelLayoutStrategy.clip(d3Select(this), width, textMeasurementService.svgEllipsis);
                        } else {
                            if (!taskResourceFullText) {
                                AxisHelper.LabelLayoutStrategy.clip(d3Select(this), defaultWidth, textMeasurementService.svgEllipsis);
                            }
                        }
                    });
            } else if (!taskResourceFullText) {
                taskResourceMerged
                    .each(function () {
                        AxisHelper.LabelLayoutStrategy.clip(d3Select(this), defaultWidth, textMeasurementService.svgEllipsis);
                    });
            }

            taskResource
                .exit()
                .remove();
        } else {
            taskSelection
                .selectAll(Gantt.TaskResource.selectorName)
                .remove();
        }
    }

    private static getSameRowNextTaskStartDate(task: Task, index: number, selection: Selection<Task>) {
        let sameRowNextTaskStart: Date;

        selection
            .each(function (x: Task, i: number) {
                if (index !== i &&
                    x.index === task.index &&
                    x.start >= task.start &&
                    (!sameRowNextTaskStart || sameRowNextTaskStart < x.start)) {

                    sameRowNextTaskStart = x.start;
                }
            });

        return sameRowNextTaskStart;
    }

    private static getResourceLabelYOffset(
        taskConfigHeight: number,
        taskResourceFontSize: number,
        taskResourcePosition: ResourceLabelPosition): number {
        const barHeight: number = Gantt.getBarHeight(taskConfigHeight);
        switch (taskResourcePosition) {
            case ResourceLabelPosition.Right:
                return (barHeight / Gantt.DividerForCalculatingCenter) + (taskResourceFontSize / Gantt.DividerForCalculatingCenter);
            case ResourceLabelPosition.Top:
                return -(taskResourceFontSize / Gantt.DividerForCalculatingPadding) + Gantt.LabelTopOffsetForPadding;
            case ResourceLabelPosition.Inside:
                return -(taskResourceFontSize / Gantt.DividerForCalculatingPadding) + Gantt.LabelTopOffsetForPadding + barHeight / Gantt.ResourceLabelDefaultDivisionCoefficient;
        }
    }

    private getResourceLabelXCoordinate(
        task: Task,
        taskConfigHeight: number,
        taskResourceFontSize: number,
        taskResourcePosition: ResourceLabelPosition): number {
        if (!this.hasNotNullableDates) {
            return 0;
        }

        const barHeight: number = Gantt.getBarHeight(taskConfigHeight);
        switch (taskResourcePosition) {
            case ResourceLabelPosition.Right:
                return (Gantt.TimeScale(task.end) + (taskResourceFontSize / 2) + Gantt.RectRound) || 0;
            case ResourceLabelPosition.Top:
                return (Gantt.TimeScale(task.start) + Gantt.RectRound) || 0;
            case ResourceLabelPosition.Inside:
                return (Gantt.TimeScale(task.start) + barHeight / (2 * Gantt.ResourceLabelDefaultDivisionCoefficient) + Gantt.RectRound) || 0;
        }
    }

    /**
     * Returns the matching Y coordinate for a given task index
     * @param taskIndex Task Number
     */
    private getTaskLabelCoordinateY(taskIndex: number): number {
        const settings = this.viewModel.settings;
        const fontSize: number = + settings.taskLabelsCardSettings.fontSize.value;
        const taskConfigHeight = settings.taskConfigCardSettings.height.value || DefaultChartLineHeight;
        const taskYCoordinate = taskConfigHeight * taskIndex;
        const barHeight = Gantt.getBarHeight(taskConfigHeight);
        return taskYCoordinate + (barHeight + Gantt.BarHeightMargin - (taskConfigHeight - fontSize) / Gantt.ChartLineHeightDivider);
    }

    /**
    * Get bar y coordinate
    * @param lineNumber Line number that represents the task number
    * @param lineHeight Height of task line
    */
    private static getBarYCoordinate(
        lineNumber: number,
        lineHeight: number): number {
        return (lineHeight * lineNumber) + PaddingTasks;
    }

    /**
     * Get bar height
     * @param lineHeight The height of line
     */
    private static getBarHeight(lineHeight: number): number {
        return lineHeight / Gantt.ChartLineProportion;
    }

    /**
     * Get the margin that added to task rects and task category labels
     *
     * depends on resource label position and resource label font size
     */
    private getResourceLabelTopMargin(): number {
        const isResourcesFilled: boolean = this.viewModel.isResourcesFilled;
        const taskResourceShow: boolean = this.viewModel.settings.taskResourceCardSettings.show.value;
        const taskResourceFontSize: number = this.viewModel.settings.taskResourceCardSettings.fontSize.value;
        const taskResourcePosition: ResourceLabelPosition = ResourceLabelPosition[this.viewModel.settings.taskResourceCardSettings.position.value.value];

        let margin: number = 0;
        if (isResourcesFilled && taskResourceShow && taskResourcePosition === ResourceLabelPosition.Top) {
            margin = Number(taskResourceFontSize) + Gantt.LabelTopOffsetForPadding;
        }

        return margin;
    }

    /**
     * convert task duration to width in the timescale
     * @param start The start of task to convert
     * @param end The end of task to convert
     */
    private static taskDurationToWidth(
        start: Date,
        end: Date): number {
        return Gantt.TimeScale(end) - Gantt.TimeScale(start);
    }

    private static getTooltipForMilestoneLine(
        formattedDate: string,
        localizationManager: ILocalizationManager,
        dateTypeSettings: DateTypeCardSettings,
        milestoneTitle: string[] | LabelForDate[], milestoneCategoryName?: string[]): VisualTooltipDataItem[] {
        const result: VisualTooltipDataItem[] = [];

        for (let i = 0; i < milestoneTitle.length; i++) {
            if (!milestoneTitle[i]) {
                switch (dateTypeSettings.type.value.value) {
                    case DateType.Second:
                    case DateType.Minute:
                    case DateType.Hour:
                        milestoneTitle[i] = localizationManager.getDisplayName("Visual_Label_Now");
                        break;
                    default:
                        milestoneTitle[i] = localizationManager.getDisplayName("Visual_Label_Today");
                }
            }

            if (milestoneCategoryName) {
                result.push({
                    displayName: localizationManager.getDisplayName("Visual_Milestone_Name"),
                    value: milestoneCategoryName[i]
                });
            }

            result.push({
                displayName: <string>milestoneTitle[i],
                value: formattedDate
            });
        }

        return result;
    }

    /**
    * Create vertical dotted line that represent milestone in the time axis (by default it shows not time)
    * @param tasks All tasks array
    * @param milestoneTitle
    * @param timestamp the milestone to be shown in the time axis (default Date.now())
    */
    private createMilestoneLine(
        tasks: GroupedTask[],
        timestamp: number = Date.now(),
        milestoneTitle?: string): void {
        if (!this.hasNotNullableDates) {
            return;
        }

        const todayColor: string = this.viewModel.settings.dateTypeCardSettings.todayColor.value.value;
        // TODO: add not today milestones color
        const milestoneDates = [new Date(timestamp)];
        tasks.forEach((task: GroupedTask) => {
            const subtasks: Task[] = task.tasks;
            subtasks.forEach((task: Task) => {
                if (!lodashIsEmpty(task.Milestones)) {
                    task.Milestones.forEach((milestone) => {
                        if (milestone.start && !milestoneDates.some(existingDate => existingDate.getTime() === milestone.start.getTime())) {
                            milestoneDates.push(milestone.start);
                        }
                    });
                }
            });
        });

        const line: Line[] = [];
        const dateTypeSettings: DateTypeCardSettings = this.viewModel.settings.dateTypeCardSettings;
        milestoneDates.forEach((date: Date) => {
            const title = date === Gantt.TimeScale(timestamp) ? milestoneTitle : "Milestone";
            const lineOptions = {
                x1: Gantt.TimeScale(date),
                y1: Gantt.MilestoneTop,
                x2: Gantt.TimeScale(date),
                y2: this.getMilestoneLineLength(tasks.length),
                tooltipInfo: Gantt.getTooltipForMilestoneLine(date.toLocaleDateString(), this.localizationManager, dateTypeSettings, [title])
            };
            line.push(lineOptions);
        });

        const chartLineSelection: Selection<Line> = this.chartGroup
            .selectAll(Gantt.ChartLine.selectorName)
            .data(line);

        const chartLineSelectionMerged = chartLineSelection
            .enter()
            .append("line")
            .merge(chartLineSelection);

        chartLineSelectionMerged.classed(Gantt.ChartLine.className, true);

        chartLineSelectionMerged
            .attr("x1", (line: Line) => line.x1)
            .attr("y1", (line: Line) => line.y1)
            .attr("x2", (line: Line) => line.x2)
            .attr("y2", (line: Line) => line.y2)
            .style("stroke", (line: Line) => {
                const color: string = line.x1 === Gantt.TimeScale(timestamp) ? todayColor : Gantt.DefaultValues.MilestoneLineColor;
                return this.colorHelper.getHighContrastColor("foreground", color);
            });

        this.renderTooltip(chartLineSelectionMerged);

        chartLineSelection
            .exit()
            .remove();
    }

    private scrollToMilestoneLine(axisLength: number,
        timestamp: number = Date.now()): void {

        let scrollValue = Gantt.TimeScale(new Date(timestamp));
        scrollValue -= scrollValue > ScrollMargin
            ? ScrollMargin
            : 0;

        if (axisLength > scrollValue) {
            (this.body.node() as SVGSVGElement)
                .querySelector(Gantt.Body.selectorName).scrollLeft = scrollValue;
        }
    }

    private renderTooltip(selection: Selection<Line | Task | MilestonePath>): void {
        this.tooltipServiceWrapper.addTooltip(
            selection,
            (tooltipEvent: TooltipEnabledDataPoint) => tooltipEvent.tooltipInfo);
    }

    private updateElementsPositions(margin: IMargin): void {
        const settings: GanttChartSettingsModel = this.viewModel.settings;
        const taskLabelsWidth: number = settings.taskLabelsCardSettings.show.value
            ? settings.taskLabelsCardSettings.width.value
            : 0;

        let translateXValue: number = taskLabelsWidth + margin.left + Gantt.SubtasksLeftMargin;
        this.chartGroup
            .attr("transform", SVGManipulations.translate(translateXValue, margin.top));

        const translateYValue: number = Gantt.TaskLabelsMarginTop + (this.ganttDiv.node() as SVGSVGElement).scrollTop;
        this.axisGroup
            .attr("transform", SVGManipulations.translate(translateXValue, translateYValue));

        translateXValue = (this.ganttDiv.node() as SVGSVGElement).scrollLeft;
        this.lineGroup
            .attr("transform", SVGManipulations.translate(translateXValue, 0));
        this.collapseAllGroup
            .attr("transform", SVGManipulations.translate(0, margin.top / 4 + Gantt.AxisTopMargin));
    }

    private getMilestoneLineLength(numOfTasks: number): number {
        return numOfTasks * ((this.viewModel.settings.taskConfigCardSettings.height.value || DefaultChartLineHeight) + (1 + numOfTasks) * this.getResourceLabelTopMargin() / 2);
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        this.filterSettingsCards();
        this.formattingSettings.setLocalizedOptions(this.localizationManager);
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    public filterSettingsCards() {
        const settings: GanttChartSettingsModel = this.formattingSettings;

        settings.milestonesCardSettings.visible = !!(this.viewModel?.milestonesData?.dataPoints?.length);

        settings.cards.forEach(element => {
            switch(element.name) {
                case Gantt.TaskResourcePropertyIdentifier.objectName:
                    if (!this.viewModel.isResourcesFilled) {
                        settings.taskResourceCardSettings.visible = false;
                    }
                    break;
            }
        });
    }
}
