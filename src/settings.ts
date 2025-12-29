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

"use strict";

import powerbi from "powerbi-visuals-api";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

const animationOptions: powerbi.IEnumMember[] = [
    { value: "none", displayName: "None" },
    { value: "grow", displayName: "Grow" },
    { value: "explode", displayName: "Explode" },
    { value: "spin", displayName: "Spin" },
    { value: "fade", displayName: "Fade" }
];

const hoverAnimationOptions: powerbi.IEnumMember[] = [
    { value: "none", displayName: "None" },
    { value: "grow", displayName: "Grow" },
    { value: "explode", displayName: "Explode" },
    { value: "pulse", displayName: "Pulse" },
    { value: "color", displayName: "Color" },
    { value: "tilt", displayName: "Tilt" }
];

const legendPlacementOptions: powerbi.IEnumMember[] = [
    { value: "auto", displayName: "Auto" },
    { value: "right", displayName: "Right" },
    { value: "left", displayName: "Left" },
    { value: "top", displayName: "Top" },
    { value: "bottom", displayName: "Bottom" }
];

const legendLayoutOptions: powerbi.IEnumMember[] = [
    { value: "list", displayName: "List" },
    { value: "row", displayName: "Row" }
];

const detailLabelOptions: powerbi.IEnumMember[] = [
    { value: "category", displayName: "Category" },
    { value: "dataValue", displayName: "Data value" },
    { value: "percent", displayName: "Percent of total" },
    { value: "categoryDataValue", displayName: "Category, data value" },
    { value: "categoryPercent", displayName: "Category, percent of total" },
    { value: "dataValuePercent", displayName: "Data value, percent of total" },
    { value: "all", displayName: "All detail labels" }
];

const detailLabelPlacementOptions: powerbi.IEnumMember[] = [
    { value: "outside", displayName: "Outside" },
    { value: "inside", displayName: "Inside" }
];

const displayUnitOptions: powerbi.IEnumMember[] = [
    { value: "auto", displayName: "Auto" },
    { value: "none", displayName: "None" },
    { value: "thousands", displayName: "Thousands" },
    { value: "millions", displayName: "Millions" },
    { value: "billions", displayName: "Billions" }
];

/**
 * Data Point Formatting Card
 */
class DataPointCardSettings extends FormattingSettingsCard {
    defaultColor = new formattingSettings.ColorPicker({
        name: "defaultColor",
        displayName: "Default color",
        value: { value: "" }
    });

    showAllDataPoints = new formattingSettings.ToggleSwitch({
        name: "showAllDataPoints",
        displayName: "Show all",
        value: true
    });

    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayName: "Fill",
        value: { value: "" }
    });

    fillRule = new formattingSettings.ColorPicker({
        name: "fillRule",
        displayName: "Color saturation",
        value: { value: "" }
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Text Size",
        value: 12
    });

    name: string = "dataPoint";
    displayName: string = "Data colors";
    slices: Array<FormattingSettingsSlice> = [this.defaultColor, this.showAllDataPoints, this.fill, this.fillRule, this.fontSize];
}

class DetailLabelsCardSettings extends FormattingSettingsCard {
    showDetailLabels = new formattingSettings.ToggleSwitch({ name: "showDetailLabels", displayName: "Show detail labels", value: true });
    detailLabelMode = new formattingSettings.ItemDropdown({ name: "detailLabelMode", displayName: "Detail label mode", value: detailLabelOptions[3], items: detailLabelOptions });
    detailLabelPlacement = new formattingSettings.ItemDropdown({ name: "detailLabelPlacement", displayName: "Label placement", value: detailLabelPlacementOptions[0], items: detailLabelPlacementOptions });
    detailLabelFontColor = new formattingSettings.ColorPicker({ name: "detailLabelFontColor", displayName: "Label color", value: { value: "#1b2233" } });
    detailLabelDisplayUnits = new formattingSettings.ItemDropdown({ name: "detailLabelDisplayUnits", displayName: "Display units", value: displayUnitOptions[0], items: displayUnitOptions });
    detailLabelFont = new formattingSettings.FontControl({
        name: "detailLabelFont",
        displayName: "Font",
        fontFamily: new formattingSettings.FontPicker({ name: "fontFamily", displayName: "Font family", value: "Segoe UI" }),
        fontSize: new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Font size", value: 12 }),
        bold: new formattingSettings.ToggleSwitch({ name: "fontBold", displayName: "Bold", value: false }),
        italic: new formattingSettings.ToggleSwitch({ name: "fontItalic", displayName: "Italic", value: false }),
        underline: new formattingSettings.ToggleSwitch({ name: "fontUnderline", displayName: "Underline", value: false })
    });
    valueDecimalPlaces = new formattingSettings.NumUpDown({ name: "valueDecimalPlaces", displayName: "Value decimal places", value: 2 });
    percentageDecimalPlaces = new formattingSettings.NumUpDown({ name: "percentageDecimalPlaces", displayName: "Percentage decimal places", value: 1 });

    name: string = "detailLabels";
    displayName: string = "Detail labels";
    slices: Array<FormattingSettingsSlice> = [
        this.showDetailLabels,
        this.detailLabelMode,
        this.detailLabelPlacement,
        this.detailLabelFontColor,
        this.detailLabelDisplayUnits,
        this.detailLabelFont,
        this.valueDecimalPlaces,
        this.percentageDecimalPlaces
    ];
}

/**
* visual settings model class
*
*/
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    // Create formatting settings model formatting cards
    dataPointCard = new DataPointCardSettings();
    detailLabelsCard = new DetailLabelsCardSettings();

    // General chart settings
    general = new (class GeneralCardSettings extends FormattingSettingsCard {
        donut = new formattingSettings.ToggleSwitch({ name: "donut", displayName: "Donut mode", value: false });
        innerRadius = new formattingSettings.NumUpDown({ name: "innerRadius", displayName: "Inner radius (%)", value: 50 });
                backgroundColor = new formattingSettings.ColorPicker({ name: "backgroundColor", displayName: "Background color", value: { value: "" } });
        showLegend = new formattingSettings.ToggleSwitch({ name: "showLegend", displayName: "Show legend", value: true });
        legendPlacement = new formattingSettings.ItemDropdown({ name: "legendPlacement", displayName: "Legend placement", value: legendPlacementOptions[0], items: legendPlacementOptions });
        legendLayout = new formattingSettings.ItemDropdown({ name: "legendLayout", displayName: "Legend layout", value: legendLayoutOptions[0], items: legendLayoutOptions });
        sortByValue = new formattingSettings.ToggleSwitch({ name: "sortByValue", displayName: "Sort by value (desc)", value: true });
        animation = new formattingSettings.ToggleSwitch({ name: "animation", displayName: "Animate", value: true });
        showTooltips = new formattingSettings.ToggleSwitch({ name: "showTooltips", displayName: "Show tooltips", value: true });
          showOverview = new formattingSettings.ToggleSwitch({ name: "showOverview", displayName: "Show overview panel", value: true });
           animationType = new formattingSettings.ItemDropdown({ name: "animationType", displayName: "Animation type", value: animationOptions[1], items: animationOptions });
           hoverAnimationType = new formattingSettings.ItemDropdown({ name: "hoverAnimationType", displayName: "Hover animation type", value: hoverAnimationOptions[1], items: hoverAnimationOptions });
        animationDuration = new formattingSettings.NumUpDown({ name: "animationDuration", displayName: "Animation duration (ms)", value: 700 });

        name: string = "general";
        displayName: string = "General";
        slices: Array<FormattingSettingsSlice> = (() => {
            const base: Array<FormattingSettingsSlice> = [
                this.donut,
                this.innerRadius,
                this.backgroundColor,
                this.showLegend,
                this.legendPlacement,
                this.legendLayout,
                this.sortByValue,
                this.animation,
                this.animationType,
                this.hoverAnimationType,
                this.animationDuration,
                this.showTooltips,
                this.showOverview
            ];
            return base;
        })();
    })();

    cards = [this.dataPointCard, this.general, this.detailLabelsCard];
}
