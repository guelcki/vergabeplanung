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

import powerbiVisualsApi from "powerbi-visuals-api";
import {formattingSettings} from "powerbi-visuals-utils-formattingmodel";
import {MilestoneShape} from "./enums";
import {DateType} from "./enums";
import {ResourceLabelPosition} from "./enums";

import ILocalizationManager = powerbi.extensibility.ILocalizationManager;

import Card = formattingSettings.SimpleCard;
import Model = formattingSettings.Model;

import IEnumMember = powerbi.IEnumMember;

export const dateTypeOptions : IEnumMember[] = [
    { displayName: "Visual_DateType_Second", value: DateType.Second },
    { displayName: "Visual_DateType_Minute", value: DateType.Minute },
    { displayName: "Visual_DateType_Hour", value: DateType.Hour },
    { displayName: "Visual_DateType_Day", value: DateType.Day },
    { displayName: "Visual_DateType_Week", value: DateType.Week },
    { displayName: "Visual_DateType_Month", value: DateType.Month },
    { displayName: "Visual_DateType_Quarter", value: DateType.Quarter },
    { displayName: "Visual_DateType_Year", value: DateType.Year }
]

const shapesOptions : IEnumMember[] = [
    { displayName: "Visual_Shape_Rhombus", value: MilestoneShape.Rhombus },
    { displayName: "Visual_Shape_Circle", value: MilestoneShape.Circle },
    { displayName: "Visual_Shape_Square", value: MilestoneShape.Square }
]

const resourcePositionOptions : IEnumMember[] = [
    { displayName: "Visual_Position_Top", value: ResourceLabelPosition.Top },
    { displayName: "Visual_Position_Right", value: ResourceLabelPosition.Right },
    { displayName: "Visual_Position_Inside", value: ResourceLabelPosition.Inside }
];

class FontSizeSettings {
    public static readonly DefaultFontSize: number = 9;
    public static readonly MinFontSize: number = 8;
}

class WidthSettings {
    public static readonly DefaultFontSize: number = 110;
    public static readonly MinFontSize: number = 0;
}

class HeightSettings {
    public static readonly DefaultFontSize: number = 25;
    public static readonly MinFontSize: number = 1;
}

export class GeneralCardSettings extends Card {

    groupTasks = new formattingSettings.ToggleSwitch({
        name: "groupTasks",
        displayNameKey: "Visual_GroupTasks",
        value: false
    });

    scrollToCurrentTime = new formattingSettings.ToggleSwitch({
        name: "scrollToCurrentTime",
        displayNameKey: "Visual_ScrollToCurrentTime",
        value: false
    });

    displayGridLines = new formattingSettings.ToggleSwitch({
        name: "displayGridLines",
        displayNameKey: "Visual_DisplayGridLines",
        value: true
    });

    name: string = "general";
    displayNameKey: string = "Visual_General";
    slices = [this.groupTasks, this.scrollToCurrentTime, this.displayGridLines];
}

export class SubTasksCardSettings extends Card {

    parentCompletionByChildren = new formattingSettings.ToggleSwitch({
        name: "parentCompletionByChildren",
        displayNameKey: "Visual_ParentCompletionByChildren",
        value: true
    });

    name: string = "subTasks";
    displayNameKey: string = "Visual_SubTasks";
    slices = [this.parentCompletionByChildren];
}

export class CollapsedTasksCardSettings extends Card {

    list = new formattingSettings.TextInput({
        name: "list",
        displayNameKey: "Visual_List",
        placeholder: "",
        value: "[]"
    });

    visible: boolean = false;
    name: string = "collapsedTasks";
    displayNameKey: string = "Visual_CollapsedTasks";
    slices = [this.list];
}

export class CollapsedTasksUpdateIdCardSettings extends Card {

    value = new formattingSettings.TextInput({
        name: "value",
        displayNameKey: "Visual_UpdateId",
        placeholder: "",
        value: ""
    });

    visible: boolean = false;
    name: string = "collapsedTasksUpdateId";
    displayNameKey: string = "Visual_CollapsedTasksUpdateId";
    slices = [this.value];
}

export class MilestonesCardSettings extends Card {

    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayNameKey: "Visual_Fill",
        value: { value: "#FFFA94" }
    });

    labelColor = new formattingSettings.ColorPicker({
        name: "labelColor",
        displayName: "Datum",
        value: { value: "#303030" }
    });

    labelFontSize = new formattingSettings.NumUpDown({
        name: "labelFontSize",
        displayName: "Schriftgrad",
        value: 16,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 8
            }
        }
    });

    shapeType = new formattingSettings.ItemDropdown({
        name: "shapeType",
        displayNameKey: "Visual_Shape",
        items: shapesOptions,
        value: shapesOptions[0]
    });

    name: string = "milestones";
    displayNameKey: string = "Visual_Milestones";
    slices = [this.fill, this.labelColor, this.labelFontSize, this.shapeType];
}

export class TaskLabelsCardSettings extends Card {

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayNameKey: "Visual_Show",
        value: true
    });

    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayNameKey: "Visual_Fill",
        value: { value: "#000000" }
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayNameKey: "Visual_FontSize",
        value: FontSizeSettings.DefaultFontSize,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: FontSizeSettings.MinFontSize,
            },
        }
    });

    width = new formattingSettings.NumUpDown({
        name: "width",
        displayNameKey: "Visual_Width",
        value: WidthSettings.DefaultFontSize,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: WidthSettings.MinFontSize,
            },
        }
    });

    name: string = "taskLabels";
    displayNameKey: string = "Visual_CategoryLabels";
    slices = [this.fill, this.fontSize, this.width];
    topLevelSlice?: formattingSettings.SimpleSlice<any> = this.show;
}

export class TooltipConfigCardSettings extends Card {

    dateFormat = new formattingSettings.TextInput({
        name: "dateFormat",
        displayNameKey: "Visual_TooltipSettings_DateFormat",
        placeholder: "",
        value: ""
    });

    name: string = "tooltipConfig";
    displayNameKey: string = "Visual_TooltipSettings";
    slices = [this.dateFormat];
}

export class TaskConfigCardSettings extends Card {

    criticalFill = new formattingSettings.ColorPicker({
        name: "criticalFill",
        displayName: "Farbe terminkritisch",
        value: { value: "#FF341D" }
    });

    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayNameKey: "Visual_TaskSettings_Color",
        description: "This ONLY takes effect when you have no legend specified",
        descriptionKey: "Visual_Description_TaskSettings_Color",
        value: { value: "#838383" }
    });

    height = new formattingSettings.NumUpDown({
        name: "height",
        displayNameKey: "Visual_TaskSettings_Height",
        value: HeightSettings.DefaultFontSize,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: HeightSettings.MinFontSize,
            },
        }
    });

    name: string = "taskConfig";
    displayNameKey: string = "Visual_TaskSettings";
    slices = [this.fill, this.criticalFill, this.height];
}

export class TaskResourceCardSettings extends Card {

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayNameKey: "Visual_Show",
        value: true
    });

    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayNameKey: "Visual_Color",
        value: { value: "#000000" }
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayNameKey: "Visual_FontSize",
        value: FontSizeSettings.DefaultFontSize,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: FontSizeSettings.MinFontSize,
            },
        }
    });

    position = new formattingSettings.ItemDropdown({
        name: "position",
        displayNameKey: "Visual_Position",
        items: resourcePositionOptions,
        value: resourcePositionOptions[1]
    });

    fullText = new formattingSettings.ToggleSwitch({
        name: "fullText",
        displayNameKey: "Visual_FullText",
        value: false
    });

    widthByTask = new formattingSettings.ToggleSwitch({
        name: "widthByTask",
        displayNameKey: "Visual_WidthByTask",
        value: false
    });

    name: string = "taskResource";
    displayNameKey: string = "Visual_DataLabels";
    slices = [this.fill, this.fontSize, this.position, this.fullText, this.widthByTask];
    topLevelSlice?: formattingSettings.SimpleSlice<any> = this.show;
}

export class DateTypeCardSettings extends Card {

    type = new formattingSettings.ItemDropdown({
        name: "type",
        displayNameKey: "Visual_Type",
        items: dateTypeOptions,
        value: dateTypeOptions[5]
    });

    todayColor = new formattingSettings.ColorPicker({
        name: "todayColor",
        displayNameKey: "Visual_DateType_TodayColor",
        value: { value: "#000000" }
    });

    axisColor = new formattingSettings.ColorPicker({
        name: "axisColor",
        displayNameKey: "Visual_DateType_AxisColor",
        value: { value: "#000000" }
    });

    axisTextColor = new formattingSettings.ColorPicker({
        name: "axisTextColor",
        displayNameKey: "Visual_DateType_AxisTextColor",
        value: { value: "#000000" }
    });

    name: string = "dateType";
    displayNameKey: string = "Visual_DateType";
    slices = [this.type, this.todayColor, this.axisColor, this.axisTextColor];
}

export class GanttChartSettingsModel extends Model { 
    generalCardSettings = new GeneralCardSettings();
    collapsedTasksCardSettings = new CollapsedTasksCardSettings();
    collapsedTasksUpdateIdCardSettings = new CollapsedTasksUpdateIdCardSettings();
    milestonesCardSettings = new MilestonesCardSettings();
    taskLabelsCardSettings = new TaskLabelsCardSettings();
    tooltipConfigCardSettings = new TooltipConfigCardSettings();
    taskConfigCardSettings = new TaskConfigCardSettings();
    taskResourceCardSettings = new TaskResourceCardSettings();
    dateTypeCardSettings = new DateTypeCardSettings();
    
    cards = [this.generalCardSettings, this.collapsedTasksCardSettings, this.collapsedTasksUpdateIdCardSettings,
            this.milestonesCardSettings, this.taskLabelsCardSettings,
            this.tooltipConfigCardSettings, this.taskConfigCardSettings, this.taskResourceCardSettings, this.dateTypeCardSettings];


    setLocalizedOptions(localizationManager: ILocalizationManager) {
        this.setLocalizedDisplayName(shapesOptions, localizationManager);
        this.setLocalizedDisplayName(resourcePositionOptions, localizationManager);
        this.setLocalizedDisplayName(dateTypeOptions, localizationManager);
    }       

    public setLocalizedDisplayName(options: IEnumMember[], localizationManager: ILocalizationManager) {
        options.forEach(option => {
            option.displayName = localizationManager.getDisplayName(option.displayName.toString())
        });
    }
}
