import {
    FormattingSettingsService,
    formattingSettings
} from "powerbi-visuals-utils-formattingmodel";

const getToday = (): Date => {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return today;
};

const addMonths = (base: Date, months: number): Date => {
    const result = new Date(base);
    result.setMonth(result.getMonth() + months);
    return result;
};

export class ZeitachseSettings extends formattingSettings.SimpleCard {
    vergabeVon = new formattingSettings.DatePicker({
        name: "vergabeVon",
        displayName: "Vergabe SOLL von",
        placeholder: "",
        value: getToday()
    });

    vergabeBis = new formattingSettings.DatePicker({
        name: "vergabeBis",
        displayName: "Vergabe SOLL bis",
        placeholder: "",
        value: addMonths(getToday(), 6)
    });

    name = "zeitachse";
    displayName = "Zeitraum Vergabe SOLL";
    slices = [this.vergabeVon, this.vergabeBis];
}

export class TimelineFormattingSettingsModel extends formattingSettings.Model {
    zeitachse = new ZeitachseSettings();
    cards = [this.zeitachse];
}

export const formattingSettingsService = new FormattingSettingsService();
