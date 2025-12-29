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

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import DataView = powerbi.DataView;

import { VisualFormattingSettingsModel } from "./settings";
import * as d3 from "d3";

type Selection<T extends d3.BaseType> = d3.Selection<T, any, any, any>;

export class Visual implements IVisual {
    private target: HTMLElement;
    private svg: Selection<SVGElement>;
    private container: Selection<SVGElement>;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    private lastData: any[] | null = null;
    private lastWidth: number = 0;
    private lastHeight: number = 0;
    private host: powerbi.extensibility.visual.IVisualHost;
    private selectionManager: powerbi.extensibility.ISelectionManager;
    private activeSelection: powerbi.extensibility.ISelectionId[] = [];
    private currentSlices: Selection<any> | null = null;
    private currentLegendItems: Selection<HTMLDivElement> | null = null;
    private detailPanel: Selection<HTMLDivElement> | null = null;
    private detailPanelDecimals: number = 2;
    private showDetailPanelSetting: boolean = true;

    private canInteract(): boolean {
        return !!(this.host && (this.host as any).allowInteractions);
    }
    private resolveEnumMemberValue(slice: any): string | undefined {
        const raw = slice && slice.value;
        if (raw === undefined || raw === null) {
            return undefined;
        }
        if (typeof raw === "object" && "value" in raw && raw.value !== undefined) {
            return String(raw.value);
        }
        return String(raw);
    }

    private getAnimationDurationFromFormatting(defaultDuration: number): number {
        if (this.formattingSettings && this.formattingSettings.general) {
            const gen: any = this.formattingSettings.general as any;
            if (gen.animationDuration && typeof gen.animationDuration.value === "number") {
                return gen.animationDuration.value;
            }
        }
        return defaultDuration;
    }

    constructor(options: VisualConstructorOptions) {
        this.formattingSettingsService = new FormattingSettingsService();
        this.target = options.element;
        this.host = options.host;
        this.selectionManager = this.host.createSelectionManager();
        if (this.canInteract() && this.selectionManager.registerOnSelectCallback) {
            this.selectionManager.registerOnSelectCallback((ids) => {
                this.activeSelection = ids || [];
                const total = this.lastData ? d3.sum(this.lastData, d => d.value) : 0;
                this.applySelectionState();
                this.updateDetailPanel(total);
            });
        }
        
        this.svg = d3.select(options.element)
            .append('svg')
            .classed('pie-chart', true);
        
        this.container = this.svg.append("g")
            .classed('container', true);

        // tooltip div (hidden by default)
        d3.select(this.target).style('position', 'relative');
        const existing = d3.select(this.target).select('.adv-tooltip');
        if (!existing.empty()) {
            existing.remove();
        }
        d3.select(this.target).append('div').classed('adv-tooltip', true)
            .style('position', 'absolute')
            .style('pointer-events', 'none')
            .style('display', 'none');
        const existingDetailPanel = d3.select(this.target).select('.detail-panel');
        if (!existingDetailPanel.empty()) {
            existingDetailPanel.remove();
        }
        this.detailPanel = null;
        
            // error overlay (hidden)
            const existingErr = d3.select(this.target).select('.visual-error');
            if (!existingErr.empty()) existingErr.remove();
            d3.select(this.target).append('div').classed('visual-error', true)
                .style('position', 'absolute')
                .style('left', '8px')
                .style('top', '8px')
                .style('right', '8px')
                .style('padding', '10px')
                .style('background', 'rgba(255,230,230,0.95)')
                .style('color', '#800')
                .style('border', '1px solid #f5c2c2')
                .style('border-radius', '6px')
                .style('display', 'none')
                .text('');

            // No in-visual controls: prefer Format pane (General) for animation selections.

        this.svg.on('click', (event: MouseEvent) => {
            if (!this.canInteract()) {
                return;
            }
            if (event && event.target === this.svg.node()) {
                this.selectionManager.clear();
                this.activeSelection = [];
                const total = this.lastData ? d3.sum(this.lastData, d => d.value) : 0;
                this.applySelectionState();
                this.updateDetailPanel(total);
            }
        });
    }

    private ensureDetailPanel(): void {
        d3.select(this.target).selectAll('.detail-panel').remove();
        if (this.showDetailPanelSetting) {
            this.detailPanel = d3.select(this.target).append('div').classed('detail-panel', true);
        } else {
            this.detailPanel = null;
        }
    }

    private getAnimationTypeFromFormatting(): string | undefined {
        if (this.formattingSettings && this.formattingSettings.general) {
            const gen: any = this.formattingSettings.general as any;
            if (gen.animation && typeof gen.animation.value === 'boolean' && !gen.animation.value) {
                return 'none';
            }
            const selectedValue = this.resolveEnumMemberValue(gen.animationType);
            if (selectedValue) {
                return selectedValue;
            }
        }
        return undefined;
    }

    private formatNumericValue(value: number, decimals: number, displayUnits: string): string {
        const absValue = Math.abs(value);
        const sign = value < 0 ? '-' : '';
        if (value === 0) {
            return '0';
        }
        if (displayUnits && displayUnits !== 'auto') {
            const unit = this.getDisplayUnitInfo(displayUnits);
            const scaled = absValue / unit.divisor;
            return `${sign}${this.formatNumberWithFixedDecimals(scaled, decimals)}${unit.suffix}`;
        }

        const formatter = (num: number, minDecimals: number) => this.formatNumberWithFixedDecimals(num, decimals, minDecimals);

        if (absValue >= 1_000_000_000) {
            return `${sign}${formatter(absValue / 1_000_000_000, Math.max(2, decimals))}B`;
        }
        if (absValue >= 1_000_000) {
            return `${sign}${formatter(absValue / 1_000_000, Math.max(2, decimals))}M`;
        }
        if (absValue >= 1_000) {
            return `${sign}${formatter(absValue / 1_000, Math.max(2, decimals))}K`;
        }
        const displayDecimals = Math.max(2, decimals);
        return `${sign}${formatter(absValue, displayDecimals)}`;
    }

    private formatPercentValue(value: number, decimals: number): string {
        const formatter = new Intl.NumberFormat('en-US', {
            minimumFractionDigits: decimals,
            maximumFractionDigits: decimals
        });
        return `${formatter.format(value)}%`;
    }

    private formatNumberWithFixedDecimals(value: number, decimals: number, minDecimals?: number): string {
        const minimum = minDecimals !== undefined ? minDecimals : decimals;
        const maximum = Math.max(minimum, decimals);
        return new Intl.NumberFormat('en-US', {
            minimumFractionDigits: minimum,
            maximumFractionDigits: maximum
        }).format(value);
    }

    private getDisplayUnitInfo(displayUnits: string) {
        switch (displayUnits) {
            case 'none':
                return { divisor: 1, suffix: '' };
            case 'thousands':
                return { divisor: 1_000, suffix: 'K' };
            case 'millions':
                return { divisor: 1_000_000, suffix: 'M' };
            case 'billions':
                return { divisor: 1_000_000_000, suffix: 'B' };
            default:
                return { divisor: 1, suffix: '' };
        }
    }

    private getDetailLabelText(mode: string, category: string, valueText: string, percentText: string): string {
        switch (mode) {
            case 'category':
                return category;
            case 'dataValue':
                return valueText;
            case 'percent':
                return percentText;
            case 'categoryDataValue':
                return `${category}: ${valueText}`;
            case 'categoryPercent':
                return `${category}: ${percentText}`;
            case 'dataValuePercent':
                return `${valueText} (${percentText})`;
            case 'all':
            default:
                return `${category}: ${valueText} (${percentText})`;
        }
    }

    private wrapLabelText(text: string, maxChars: number): string[] {
        if (!text) {
            return [''];
        }
        const words = text.split(/\s+/);
        const lines: string[] = [];
        let current = '';
        words.forEach((word) => {
            if (!current) {
                current = word;
                return;
            }
            if ((current + ' ' + word).length <= maxChars) {
                current = `${current} ${word}`;
            } else {
                lines.push(current);
                current = word;
            }
        });
        if (current) {
            lines.push(current);
        }
        return lines.length ? lines : [''];
    }

    private showError(error: unknown, options?: VisualUpdateOptions): void {
        const message = error instanceof Error
            ? (error.message || String(error))
            : String(error);

        try {
            const overlay = d3.select(this.target).select('.visual-error');
            if (!overlay.empty()) {
                overlay.style('display', 'block').text(message);
            }
        } catch {
            // ignore UI error rendering failures
        }

        try {
            if (options && this.host && this.host.eventService && this.host.eventService.renderingFailed) {
                this.host.eventService.renderingFailed(options, message);
            }
        } catch {
            // ignore event reporting failures
        }
    }

    public update(options: VisualUpdateOptions) {
        try {
            if (this.host && this.host.eventService && this.host.eventService.renderingStarted) {
                this.host.eventService.renderingStarted(options);
            }

            if (!options.dataViews || options.dataViews.length === 0) {
                // clear overlays
                d3.select(this.target).selectAll('.visual-error').style('display', 'none').text('');
                if (this.host && this.host.eventService && this.host.eventService.renderingFinished) {
                    this.host.eventService.renderingFinished(options);
                }
                return;
            }

            const dataView: DataView = options.dataViews[0];
            this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dataView);

            const width: number = options.viewport.width;
            const height: number = options.viewport.height;

            const generalSettings: any = this.formattingSettings && this.formattingSettings.general ? this.formattingSettings.general as any : null;
            this.showDetailPanelSetting = !!(generalSettings && generalSettings.showOverview && generalSettings.showOverview.value);
            this.ensureDetailPanel();

            this.svg
                .attr("width", width)
                .attr("height", height);

            // Clear previous content and remove UI elements (legend/tooltip)
            this.container.selectAll("*").remove();
            d3.select(this.target).selectAll('.legend').remove();

            if (dataView.categorical && dataView.categorical.categories && dataView.categorical.values) {
                const categoryColumn = dataView.categorical.categories[0];
                const valueColumn = dataView.categorical.values[0];

                const data = [];
                const categories = categoryColumn.values;
                const values = valueColumn.values;
                for (let i = 0; i < categories.length; i++) {
                    if (typeof values[i] === "number") {
                        const selectionId = this.host.createSelectionIdBuilder()
                            .withCategory(categoryColumn, i)
                            .createSelectionId();
                        data.push({ label: String(categories[i]), value: values[i], selectionId });
                    }
                }

                if (data.length > 0) {
                    // apply sorting option
                    const sortBy = this.formattingSettings && this.formattingSettings.general && this.formattingSettings.general.sortByValue && this.formattingSettings.general.sortByValue.value;
                    if (sortBy) {
                        data.sort((a, b) => b.value - a.value);
                    }

                    // persist last data & size for in-visual controls redraw
                    this.lastData = data;
                    this.lastWidth = width;
                    this.lastHeight = height;

                    // decide animation type: prefer in-visual user selection, otherwise format pane
                    this.drawPie(data, width, height);
                }
            }

            if (this.host && this.host.eventService && this.host.eventService.renderingFinished) {
                this.host.eventService.renderingFinished(options);
            }
        }
        catch (e) {
            this.showError(e, options);
        }
    }

    private drawPie(data: any[], width: number, height: number) {
        const cx = width / 2;
        const cy = height / 2;
        const baseRadius = Math.min(width, height) / 2 - 16;

        const fallbackColors = ["#4e79a7","#f28e2b","#e15759","#76b7b2","#59a14f","#edc949","#b07aa1","#ff9da7","#9c755f","#bab0ab"];
        const palette = (this.host as any).colorPalette;
        const isHighContrast = !!(palette && palette.isHighContrast);
        const hcForeground = palette && palette.foreground && palette.foreground.value ? String(palette.foreground.value) : "#000";
        const hcBackground = palette && palette.background && palette.background.value ? String(palette.background.value) : "#fff";
        const colors: string[] = data.map((d, i) => {
            const key = d && d.label ? String(d.label) : `slice_${i}`;
            try {
                const color = palette && palette.getColor ? palette.getColor(key) : null;
                return color && color.value ? String(color.value) : fallbackColors[i % fallbackColors.length];
            } catch {
                return fallbackColors[i % fallbackColors.length];
            }
        });
        
        const general = this.formattingSettings && this.formattingSettings.general ? this.formattingSettings.general as any : null;

        const donutMode = general && general.donut && general.donut.value;
        const innerPct = general && general.innerRadius ? general.innerRadius.value : 50;
        const innerRadius = donutMode ? (baseRadius * Math.max(0, Math.min(100, innerPct)) / 100) : 0;

        const showLegend = general && general.showLegend && general.showLegend.value;
        const legendPlacementSetting = general && general.legendPlacement ? this.resolveEnumMemberValue(general.legendPlacement) : undefined;
        const legendLayoutSetting = general && general.legendLayout ? this.resolveEnumMemberValue(general.legendLayout) : undefined;
        const isLegendRowLayout = legendLayoutSetting === 'row';
        const animate = general && general.animation && general.animation.value;
        const animationType = animate ? (this.getAnimationTypeFromFormatting() || 'grow') : 'none';
        const animationDuration = animate ? this.getAnimationDurationFromFormatting(700) : 0;

        const backgroundColor = general && general.backgroundColor && general.backgroundColor.value && general.backgroundColor.value.value
            ? general.backgroundColor.value.value
            : '';
        if (isHighContrast) {
            this.svg.style('--visual-background', hcBackground);
        } else if (backgroundColor) {
            this.svg.style('--visual-background', backgroundColor);
        } else {
            this.svg.style('--visual-background', null);
        }
        const detailSettings: any = this.formattingSettings && this.formattingSettings.detailLabelsCard;
        const showDetailLabels = detailSettings && detailSettings.showDetailLabels && detailSettings.showDetailLabels.value;
        const detailLabelMode = detailSettings && detailSettings.detailLabelMode
            ? this.resolveEnumMemberValue(detailSettings.detailLabelMode)
            : 'categoryDataValue';
        const detailLabelDisplayUnits = detailSettings && detailSettings.detailLabelDisplayUnits
            ? this.resolveEnumMemberValue(detailSettings.detailLabelDisplayUnits)
            : 'auto';
        const detailLabelFont = detailSettings && detailSettings.detailLabelFont;
        const detailLabelFontSize = detailLabelFont && detailLabelFont.fontSize && typeof detailLabelFont.fontSize.value === 'number' ? detailLabelFont.fontSize.value : 12;
        const detailLabelFontFamily = detailLabelFont && detailLabelFont.fontFamily && detailLabelFont.fontFamily.value ? detailLabelFont.fontFamily.value : 'Segoe UI';
        const detailLabelFontBold = detailLabelFont && detailLabelFont.bold && detailLabelFont.bold.value;
        const detailLabelFontItalic = detailLabelFont && detailLabelFont.italic && detailLabelFont.italic.value;
        const detailLabelFontUnderline = detailLabelFont && detailLabelFont.underline && detailLabelFont.underline.value;
        const detailLabelColor = detailSettings && detailSettings.detailLabelFontColor && detailSettings.detailLabelFontColor.value && detailSettings.detailLabelFontColor.value.value ? detailSettings.detailLabelFontColor.value.value : '#ffffff';
        const valueDecimalPlaces = detailSettings && detailSettings.valueDecimalPlaces && typeof detailSettings.valueDecimalPlaces.value === 'number' ? detailSettings.valueDecimalPlaces.value : 2;
        const percentageDecimalPlaces = detailSettings && detailSettings.percentageDecimalPlaces && typeof detailSettings.percentageDecimalPlaces.value === 'number' ? detailSettings.percentageDecimalPlaces.value : 1;
        const detailLabelPlacement = detailSettings && detailSettings.detailLabelPlacement
            ? this.resolveEnumMemberValue(detailSettings.detailLabelPlacement)
            : 'outside';

        this.detailPanelDecimals = valueDecimalPlaces;

        const pie = d3.pie<any>().value((d: any) => d.value);
        const arcs = pie(data);
        const totalValue = d3.sum(data, dd => dd.value) || 0;

        // compute legend size early to reserve space and avoid overlap
        const itemHeight = 20;
        const maxLabel = data.reduce((m, x) => Math.max(m, String(x.label).length), 0);
        const availableLegendHeight = Math.max(height - 24, 0);
        const legendHeightBase = isLegendRowLayout ? 64 : Math.max(100, data.length * itemHeight + 28);
        const legendHeight = showLegend ? legendHeightBase : 0;
        const legendWidthBase = Math.min(320, Math.max(120, maxLabel * 7 + 60));
        const rowLegendAvailableWidth = Math.max(120, width - 32);
        const rowLegendWidth = rowLegendAvailableWidth;
        const legendWidth = showLegend ? (isLegendRowLayout ? rowLegendWidth : legendWidthBase) : 0;

        // calculate legend placement based on user preference and available space
        const padding = 16;
        const minWidthForRightLegend = 520;
        const minPieDiameterWithLegend = 140;
        const canPlaceRight = showLegend && width >= minWidthForRightLegend && (width - legendWidth - (padding * 2)) >= minPieDiameterWithLegend;
        const canPlaceLeft = showLegend && (width - legendWidth - (padding * 2)) >= minPieDiameterWithLegend;
        const canPlaceBottom = showLegend && (height - legendHeight - (padding * 2)) >= minPieDiameterWithLegend;
        const canPlaceTop = showLegend && (height - legendHeight - (padding * 2)) >= minPieDiameterWithLegend;

        const availablePlacements: Record<string, boolean> = {
            right: !!canPlaceRight,
            left: !!canPlaceLeft,
            top: !!canPlaceTop,
            bottom: !!canPlaceBottom
        };
        if (isLegendRowLayout) {
            availablePlacements.left = false;
            availablePlacements.right = false;
        }

        const fallbackPlacement = (): string => {
            if (canPlaceRight) return 'right';
            if (canPlaceLeft) return 'left';
            if (canPlaceBottom) return 'bottom';
            if (canPlaceTop) return 'top';
            return 'bottom';
        };

        let legendPlacement = 'bottom';
        if (showLegend) {
            let requestedPlacement = legendPlacementSetting && legendPlacementSetting !== 'auto' ? legendPlacementSetting : undefined;
            if (isLegendRowLayout && requestedPlacement && (requestedPlacement === 'left' || requestedPlacement === 'right')) {
                requestedPlacement = 'top';
            }
            if (requestedPlacement) {
                legendPlacement = requestedPlacement;
            } else {
                legendPlacement = fallbackPlacement();
            }
            if (!requestedPlacement && !availablePlacements[legendPlacement]) {
                legendPlacement = fallbackPlacement();
            }
        }

        const reservedWidth = showLegend && (legendPlacement === 'left' || legendPlacement === 'right') ? legendWidth + 24 : 0;
        const reservedHeight = showLegend && (legendPlacement === 'top' || legendPlacement === 'bottom') ? legendHeight + 8 : 0;
        const maxPieAvailableWidth = Math.max(0, width - reservedWidth - (padding * 2));
        const maxPieAvailableHeight = Math.max(0, height - (padding * 2) - reservedHeight);
        const maxDiameter = Math.min(maxPieAvailableWidth, maxPieAvailableHeight);
        const radius = Math.max(8, Math.floor(Math.min(baseRadius, maxDiameter / 2)));
        const detailLabelVisibilityThreshold = 240;
        const isDetailLabelOutside = detailLabelPlacement !== 'inside';
        const detailLabelsVisible = showDetailLabels && Math.min(width, height) >= detailLabelVisibilityThreshold && radius >= 60;

        // When the legend is on left/right or the visual is small, the chart group is not centered at width/2.
        // Use the effective chart area (excluding reserved legend space) for label bounds.
        const chartWidth = Math.max(0, width - reservedWidth);
        const chartHeight = Math.max(0, height - reservedHeight);
        const chartHalfWidth = chartWidth / 2;
        const chartHalfHeight = chartHeight / 2;

        // Outside labels need enough side space. If not, fall back to inside labels to avoid broken layouts.
        const availableSideSpace = chartHalfWidth - radius;
        const minOutsideSideSpace = 48;
        const effectiveIsDetailLabelOutside = isDetailLabelOutside && availableSideSpace >= minOutsideSideSpace;

        const arcGen = d3.arc<any>()
            .innerRadius(innerRadius)
            .outerRadius(radius);

        // hover defaults from formatting model (used by mouse handlers)
        let hoverDefaultType = 'grow';
        let hoverDefaultDuration = 300;
        if (this.formattingSettings && this.formattingSettings.general) {
            const gen: any = this.formattingSettings.general as any;
            const hoverType = this.resolveEnumMemberValue(gen.hoverAnimationType);
            if (hoverType) hoverDefaultType = hoverType;
            if (gen.hoverAnimationDuration && typeof gen.hoverAnimationDuration.value === 'number') hoverDefaultDuration = gen.hoverAnimationDuration.value;
        }

        // compute center considering reserved space if legend on right
        let cxAdjusted = Math.floor(width / 2);
        if (showLegend) {
            if (legendPlacement === 'right') {
                cxAdjusted = Math.floor((width - reservedWidth) / 2);
            } else if (legendPlacement === 'left') {
                cxAdjusted = Math.floor(reservedWidth + (width - reservedWidth) / 2);
            }
        }
        let cyAdjusted = Math.floor(height / 2);
        if (showLegend) {
            if (legendPlacement === 'top') {
                cyAdjusted = Math.floor(reservedHeight + (height - reservedHeight) / 2);
            } else if (legendPlacement === 'bottom') {
                cyAdjusted = Math.floor((height - reservedHeight) / 2);
            }
        }
        this.container.attr("transform", `translate(${cxAdjusted}, ${cyAdjusted})`);

        // slices
        const paths = this.container.selectAll('path.slice').data(arcs);

        const pathsEnter = paths.enter().append('path').classed('slice', true)
            .attr('fill', (d, i) => colors[i % colors.length])
            .attr('stroke', isHighContrast ? hcForeground : '#fff')
            .attr('stroke-width', isHighContrast ? 1 : 2)
            .style('cursor', this.canInteract() ? 'pointer' : 'default')
            .attr('d', (d) => {
                const start = Object.assign({}, d, { startAngle: d.startAngle, endAngle: d.startAngle });
                return arcGen(start) as string;
            })
            .on('click', (event, d) => this.handleSliceSelection(event, d.data.selectionId, totalValue))
            .on('contextmenu', (event, d) => this.handleContextMenu(event, d.data.selectionId))
            .on('mouseover', (event, d) => {
                const node = d3.select(event.currentTarget);
                node.attr('opacity', 0.85);

                const baseTransform = (node.attr('data-base-transform') ?? node.attr('transform') ?? '').trim();
                node.attr('data-base-transform', baseTransform);
                const withBase = (t: string) => baseTransform ? `${baseTransform} ${t}` : t;

                // hover animation settings (read defaults captured from drawPie scope)
                let hoverType = hoverDefaultType;
                let hoverDuration = hoverDefaultDuration;

                // apply hover animation
                try {
                    node.interrupt();
                    switch (hoverType) {
                        case 'none':
                            break;
                        case 'explode': {
                            const c = arcGen.centroid(d);
                            const dist = 10;
                            const len = Math.sqrt(c[0] * c[0] + c[1] * c[1]) || 1;
                            const tx = (c[0] / len) * dist;
                            const ty = (c[1] / len) * dist;
                            node.transition().duration(hoverDuration).attr('transform', withBase(`translate(${tx},${ty})`));
                            break;
                        }
                        case 'pulse': {
                            node.transition().duration(hoverDuration).attrTween('transform', function () {
                                return function (t) {
                                    const scale = 1 + 0.08 * Math.sin(t * Math.PI);
                                    return withBase(`scale(${scale})`);
                                };
                            });
                            break;
                        }
                        case 'color': {
                            node.transition().duration(hoverDuration / 2).style('filter', 'brightness(1.15)');
                            break;
                        }
                        case 'tilt': {
                            node.transition().duration(hoverDuration).attrTween('transform', function () {
                                return function (t) {
                                    const ang = (t * 6) - 3; // tilt +-3deg
                                    return withBase(`rotate(${ang})`);
                                };
                            });
                            break;
                        }
                        case 'grow':
                        default:
                            node.transition().duration(hoverDuration).attr('transform', withBase('scale(1.08)'));
                            break;
                    }
                } catch (e) {
                    // ignore animation errors
                }

                const showTooltips = this.formattingSettings && this.formattingSettings.general && this.formattingSettings.general.showTooltips && this.formattingSettings.general.showTooltips.value;
                if (showTooltips) {
                    const tooltip = d3.select(this.target).select('.adv-tooltip');
                    const percent = totalValue === 0 ? 0 : (d.data.value / totalValue) * 100;
                    const pct = this.formatPercentValue(percent, percentageDecimalPlaces);
                    const formatted = this.formatNumericValue(d.data.value, valueDecimalPlaces, 'auto');
                    tooltip.selectAll('*').remove();
                    tooltip.append('div').classed('adv-tooltip__label', true).text(d.data.label);
                    tooltip.append('div').classed('adv-tooltip__value', true).text(`${formatted} (${pct})`);
                    tooltip.style('display', 'block')
                        .style('left', (event.offsetX + 12) + 'px')
                        .style('top', (event.offsetY + 12) + 'px');
                }
            })
            .on('mousemove', (event, d) => {
                const tooltip = d3.select(this.target).select('.adv-tooltip');
                if (!tooltip.empty() && tooltip.style('display') === 'block') {
                    tooltip.style('left', (event.offsetX + 12) + 'px').style('top', (event.offsetY + 12) + 'px');
                }
            })
            .on('mouseout', (event) => {
                const node = d3.select(event.currentTarget);
                node.interrupt();
                const baseTransform = (node.attr('data-base-transform') ?? '').trim();
                node.attr('opacity', 1).style('filter', null)
                    .transition().duration(200)
                    .attr('transform', baseTransform ? baseTransform : null);
                const tooltip = d3.select(this.target).select('.adv-tooltip');
                if (!tooltip.empty()) tooltip.style('display', 'none');
            });

        const sliceSelection = this.container.selectAll('path.slice');
        this.currentSlices = sliceSelection;

        // Helper: interpolate arcs for smooth updates (resize/data changes)
        const arcTween = function (this: any, d: any) {
            const start = this._current || Object.assign({}, d, { endAngle: d.startAngle });
            const i = d3.interpolate(start, d);
            this._current = i(0);
            return function (t: number) {
                const curr = i(t);
                return arcGen(curr) as string;
            };
        };

        // animate to final
        // handle different animation types
        // Clear any previous spin transform so switching animation types doesn't leave stale CSS transforms.
        d3.select(this.svg.node()).style('transform', null).style('transform-origin', null);
        switch (animationType) {
            case 'none':
                pathsEnter.attr('d', (d) => arcGen(d) as string);
                paths.attr('d', (d) => arcGen(d) as string);
                // reset base transform
                this.container.selectAll('path.slice').attr('transform', null).attr('data-base-transform', '');
                break;
            case 'fade':
                pathsEnter.attr('d', (d) => arcGen(d) as string)
                    .style('opacity', 0)
                    .transition().duration(animationDuration).style('opacity', 1);
                paths.attr('d', (d) => arcGen(d) as string).style('opacity', 1);
                this.container.selectAll('path.slice').attr('transform', null).attr('data-base-transform', '');
                break;
            case 'spin':
                // rotate container and grow slices
                d3.select(this.svg.node()).style('transform-origin', `${cxAdjusted}px ${cyAdjusted}px`)
                    .style('transform', 'rotate(-360deg)')
                    .transition().duration(animationDuration).style('transform', 'rotate(0deg)');

                pathsEnter.transition().duration(animationDuration).attrTween('d', arcTween);
                paths.transition().duration(animationDuration).attrTween('d', arcTween);
                this.container.selectAll('path.slice').attr('transform', null).attr('data-base-transform', '');
                break;
            case 'explode':
                // draw and then slightly translate slices outward
                pathsEnter.attr('d', (d) => arcGen(d) as string)
                    .attr('transform', 'translate(0,0)')
                    .style('opacity', 0)
                    .transition().duration(animationDuration)
                    .style('opacity', 1)
                    .attrTween('transform', function (d) {
                        const p = arcGen.centroid(d);
                        const dist = 8; // explode distance
                        const len = Math.sqrt(p[0] * p[0] + p[1] * p[1]) || 1;
                        const tx = (p[0] / len) * dist;
                        const ty = (p[1] / len) * dist;
                        d3.select(this).attr('data-base-transform', `translate(${tx},${ty})`);
                        const interp = d3.interpolateNumber(0, 1);
                        return function (t) {
                            const k = interp(t);
                            return `translate(${tx * k}, ${ty * k})`;
                        };
                    });

                // Update existing slices too (resize/data changes) and keep base transform in sync
                paths.attr('d', (d) => arcGen(d) as string)
                    .transition().duration(animationDuration)
                    .attrTween('transform', function (d) {
                        const p = arcGen.centroid(d);
                        const dist = 8;
                        const len = Math.sqrt(p[0] * p[0] + p[1] * p[1]) || 1;
                        const tx = (p[0] / len) * dist;
                        const ty = (p[1] / len) * dist;
                        d3.select(this).attr('data-base-transform', `translate(${tx},${ty})`);
                        const interp = d3.interpolateNumber(0, 1);
                        return function (t) {
                            const k = interp(t);
                            return `translate(${tx * k}, ${ty * k})`;
                        };
                    });
                break;
            case 'grow':
            default:
                pathsEnter.transition().duration(animationDuration).attrTween('d', arcTween);
                paths.transition().duration(animationDuration).attrTween('d', arcTween);
                this.container.selectAll('path.slice').attr('transform', null).attr('data-base-transform', '');
                break;
        }

        const detailLabelSize = Math.max(10, detailLabelFontSize);
        const fontWeight = detailLabelFontBold ? '700' : '600';
        const fontStyle = detailLabelFontItalic ? 'italic' : 'normal';
        const textDecoration = detailLabelFontUnderline ? 'underline' : 'none';
        const labelOffsetRadius = effectiveIsDetailLabelOutside
            ? radius + Math.max(20, detailLabelSize * 1.8)
            : Math.max(radius - 12, radius * 0.5);
        const labelArc = d3.arc<any>().innerRadius(labelOffsetRadius).outerRadius(labelOffsetRadius);
        const lineHeight = 1.2;
        const charLimitInside = Math.max(8, Math.round(radius / 3.5));
        const charLimitOutside = Math.max(12, Math.round((width - radius * 2) / 10));
        const wrapLimit = effectiveIsDetailLabelOutside ? charLimitOutside : charLimitInside;
        const self = this;

        // Clear any existing connectors and labels when labels are hidden
        if (!detailLabelsVisible) {
            this.container.selectAll('.slice-connector').remove();
            this.container.selectAll('.slice-label').remove();
        }

        // Handle inside labels (simple placement)
        if (detailLabelsVisible && !effectiveIsDetailLabelOutside) {
            this.container.selectAll('.slice-connector').remove();
            
            const insideLabelSelection = this.container.selectAll<SVGTextElement, any>('text.slice-label').data(arcs);
            insideLabelSelection.exit().remove();
            
            const insideLabelEnter = insideLabelSelection.enter().append('text')
                .classed('slice-label', true)
                .attr('dominant-baseline', 'middle')
                .attr('pointer-events', 'none');

            insideLabelEnter.merge(insideLabelSelection as any)
                .attr('transform', (d) => {
                    const centroid = labelArc.centroid(d) as [number, number];
                    return `translate(${centroid[0]}, ${centroid[1]})`;
                })
                .attr('text-anchor', 'middle')
                .attr('fill', detailLabelColor)
                .style('font-size', detailLabelSize + 'px')
                .style('font-family', detailLabelFontFamily)
                .style('font-weight', fontWeight)
                .style('font-style', fontStyle)
                .style('text-decoration', textDecoration)
                .each(function (d) {
                    const percent = totalValue === 0 ? 0 : (d.data.value / totalValue) * 100;
                    const percentText = self.formatPercentValue(percent, percentageDecimalPlaces);
                    const formattedValue = self.formatNumericValue(d.data.value, valueDecimalPlaces, detailLabelDisplayUnits);
                    const labelText = self.getDetailLabelText(detailLabelMode, d.data.label, formattedValue, percentText);
                    const lines = self.wrapLabelText(labelText, wrapLimit);
                    const verticalOffset = ((lines.length - 1) * lineHeight) / 2;
                    const textNode = d3.select(this);
                    const tspans = textNode.selectAll<SVGTSpanElement, string>('tspan').data(lines);
                    tspans.enter().append('tspan')
                        .attr('x', 0)
                        .attr('dy', (_, index) => (index === 0 ? `${-verticalOffset}em` : `${lineHeight}em`))
                        .merge(tspans)
                        .text(line => line);
                    tspans.exit().remove();
                });
        }

        // Handle outside labels (complex placement with connectors)
        if (detailLabelsVisible && effectiveIsDetailLabelOutside) {
            const outsideLabelData: Array<{
                arc: any;
                midAngle: number;
                sliceX: number;
                sliceY: number;
                elbowX: number;
                elbowY: number;
                targetY: number;
                finalY: number;
                finalX: number;
                side: 'left' | 'right';
            }> = arcs.map((d) => {
                const midAngle = (d.startAngle + d.endAngle) / 2;
                const sliceCentroid = arcGen.centroid(d) as [number, number];
                const labelStart = labelArc.centroid(d) as [number, number];
                const side = midAngle > Math.PI ? 'left' : 'right';
                return {
                    arc: d,
                    midAngle: midAngle,
                    sliceX: sliceCentroid[0],
                    sliceY: sliceCentroid[1],
                    elbowX: labelStart[0],
                    elbowY: labelStart[1],
                    targetY: labelStart[1],
                    finalY: labelStart[1],
                    finalX: 0,
                    side: side
                };
            });

            // Separate into left and right sides
            const rightLabels = outsideLabelData.filter(d => d.side === 'right');
            const leftLabels = outsideLabelData.filter(d => d.side === 'left');

            // Distribute labels vertically with adaptive spacing (use chart-area bounds)
            const topLimit = -chartHalfHeight + padding + 14;
            const bottomLimit = chartHalfHeight - padding - 14;
            const availableHeight = bottomLimit - topLimit;

            const distributeVertically = (labels: typeof outsideLabelData) => {
                if (labels.length === 0) return;
                
                // Sort by target Y position
                labels.sort((a, b) => a.targetY - b.targetY);
                
                // Calculate required space and adjust spacing if needed
                const idealSpacing = detailLabelSize * 2.0;
                const requiredHeight = (labels.length - 1) * idealSpacing;
                
                let actualSpacing: number;
                if (requiredHeight > availableHeight) {
                    // Reduce spacing to fit all labels
                    actualSpacing = Math.max(detailLabelSize * 1.1, availableHeight / (labels.length - 1));
                } else {
                    actualSpacing = idealSpacing;
                }
                
                // Distribute evenly if space is constrained, otherwise use target positions
                if (requiredHeight > availableHeight) {
                    // Even distribution when space is tight
                    for (let i = 0; i < labels.length; i++) {
                        labels[i].finalY = topLimit + (i * actualSpacing);
                    }
                } else {
                    // Normal distribution with target positions
                    labels[0].finalY = Math.max(topLimit, labels[0].targetY);
                    for (let i = 1; i < labels.length; i++) {
                        labels[i].finalY = Math.max(
                            labels[i].targetY,
                            labels[i - 1].finalY + actualSpacing
                        );
                    }
                    
                    // Backward pass - compress from bottom if needed
                    labels[labels.length - 1].finalY = Math.min(bottomLimit, labels[labels.length - 1].finalY);
                    for (let i = labels.length - 2; i >= 0; i--) {
                        labels[i].finalY = Math.min(
                            labels[i].finalY,
                            labels[i + 1].finalY - actualSpacing
                        );
                    }
                }
                
                // Final bounds check
                labels.forEach(label => {
                    label.finalY = Math.max(topLimit, Math.min(bottomLimit, label.finalY));
                });
            };

            distributeVertically(rightLabels);
            distributeVertically(leftLabels);

            // Set horizontal positions - responsive to canvas size
            // Since the pie group is centered inside the chart area, bounds are symmetrical:
            const leftBound = -chartHalfWidth + padding;
            const rightBound = chartHalfWidth - padding;

            // Keep labels within the chart area bounds.
            const desiredOffset = radius + 55;
            const minOffset = radius + 32;

            const rightLabelX = Math.min(desiredOffset, rightBound - 6);
            const leftLabelX = Math.max(-desiredOffset, leftBound + 6);

            // If very tight, pull labels closer.
            const clampedRightX = Math.max(minOffset, rightLabelX);
            const clampedLeftX = Math.min(-minOffset, leftLabelX);

            rightLabels.forEach(d => { d.finalX = clampedRightX; });
            leftLabels.forEach(d => { d.finalX = clampedLeftX; });

            // Draw connectors with elbow
            const connectorSelection = this.container.selectAll<SVGPolylineElement, any>('polyline.slice-connector').data(outsideLabelData);
            connectorSelection.exit().remove();
            
            const connectorEnter = connectorSelection.enter().append('polyline').classed('slice-connector', true);
            connectorEnter.merge(connectorSelection as any)
                .attr('points', (d) => {
                    // Three-point connector: slice -> elbow -> label
                    const labelStartX = d.finalX + (d.side === 'right' ? -5 : 5);
                    return `${d.sliceX},${d.sliceY} ${d.elbowX},${d.elbowY} ${labelStartX},${d.finalY}`;
                })
                .style('stroke', 'rgba(100, 100, 100, 0.5)')
                .style('stroke-width', '1px')
                .style('fill', 'none');

            // Draw labels
            const outsideLabelSelection = this.container.selectAll<SVGTextElement, any>('text.slice-label').data(outsideLabelData);
            outsideLabelSelection.exit().remove();
            
            const outsideLabelEnter = outsideLabelSelection.enter().append('text')
                .classed('slice-label', true)
                .classed('slice-label--outside', true)
                .attr('dominant-baseline', 'middle')
                .attr('pointer-events', 'none');

            outsideLabelEnter.merge(outsideLabelSelection as any)
                .attr('transform', (d) => `translate(${d.finalX}, ${d.finalY})`)
                .attr('text-anchor', (d) => d.side === 'right' ? 'start' : 'end')
                .attr('dx', (d) => d.side === 'right' ? '4px' : '-4px')
                .attr('fill', detailLabelColor)
                .style('font-size', detailLabelSize + 'px')
                .style('font-family', detailLabelFontFamily)
                .style('font-weight', fontWeight)
                .style('font-style', fontStyle)
                .style('text-decoration', textDecoration)
                .each(function (d) {
                    const percent = totalValue === 0 ? 0 : (d.arc.data.value / totalValue) * 100;
                    const percentText = self.formatPercentValue(percent, percentageDecimalPlaces);
                    const formattedValue = self.formatNumericValue(d.arc.data.value, valueDecimalPlaces, detailLabelDisplayUnits);
                    const labelText = self.getDetailLabelText(detailLabelMode, d.arc.data.label, formattedValue, percentText);
                    const lines = self.wrapLabelText(labelText, wrapLimit);
                    const verticalOffset = ((lines.length - 1) * lineHeight) / 2;
                    const textNode = d3.select(this);
                    const tspans = textNode.selectAll<SVGTSpanElement, string>('tspan').data(lines);
                    tspans.enter().append('tspan')
                        .attr('x', 0)
                        .attr('dy', (_, index) => (index === 0 ? `${-verticalOffset}em` : `${lineHeight}em`))
                        .merge(tspans)
                        .text(line => line);
                    tspans.exit().remove();
                });
        }

        // legend (placement controlled via Format pane)
        d3.select(this.target).selectAll('.legend').remove();
        if (showLegend) {
            const legendContainer = d3.select(this.target).append('div').classed('legend', true)
                .classed('legend-row-layout', isLegendRowLayout)
                .style('position', 'absolute')
                .style('font-family', 'Segoe UI, Arial')
                .style('font-size', '12px')
                .style('background', 'transparent')
                .style('padding', '0')
                .style('border-radius', '0px')
                .style('border', 'none')
                .style('box-shadow', 'none')
                .style('max-height', `${legendHeight}px`)
                .style('overflow', 'visible')
                .style('z-index', '5');

            if (legendPlacement === 'right') {
                const top = Math.max(8, Math.min(height - 8 - legendHeight, cyAdjusted - legendHeight / 2));
                legendContainer.style('right', '8px').style('top', `${top}px`).style('width', `${legendWidth}px`);
            } else if (legendPlacement === 'left') {
                const top = Math.max(8, Math.min(height - 8 - legendHeight, cyAdjusted - legendHeight / 2));
                legendContainer.style('left', '8px').style('top', `${top}px`).style('width', `${legendWidth}px`);
            } else if (legendPlacement === 'top') {
                const topWidth = Math.min(legendWidth, width);
                const left = Math.max(0, Math.floor((width - topWidth) / 2));
                legendContainer.style('left', `${left}px`).style('top', '8px');
                if (!isLegendRowLayout) {
                    legendContainer.style('width', `${topWidth}px`);
                }
            } else {
                const bottomWidth = Math.min(legendWidth, width);
                const legendTopCandidate = cyAdjusted + radius + 8;
                const legendTop = Math.max(8, Math.min(height - legendHeight - 8, legendTopCandidate));
                const left = Math.max(0, Math.floor((width - bottomWidth) / 2));
                legendContainer.style('left', `${left}px`).style('top', `${legendTop}px`);
                if (!isLegendRowLayout) {
                    legendContainer.style('width', `${bottomWidth}px`);
                }
            }

            // no legend title (lean layout)
            const legendList = legendContainer.append('div').classed('legend-items', true)
                .style('display', 'flex')
                .style('flex-direction', isLegendRowLayout ? 'row' : 'column')
                .style('flex-wrap', isLegendRowLayout ? 'wrap' : 'nowrap')
                .style('gap', '8px');
            const legendItems = legendList.selectAll('.legend-item').data(data).enter().append('div')
                .classed('legend-item', true)
                .style('display', 'flex')
                .style('align-items', 'center')
                .style('gap', '10px')
                .style('padding', '6px 4px')
                .style('border-radius', '6px')
                .style('cursor', 'pointer')
                .style('width', isLegendRowLayout ? 'auto' : '100%')
                .style('margin-bottom', isLegendRowLayout ? '0' : '8px')
                .style('flex', isLegendRowLayout ? '0 1 auto' : '0 0 auto')
                .on('click', (event, d) => this.handleSliceSelection(event, d.selectionId, totalValue))
                .on('contextmenu', (event, d) => this.handleContextMenu(event, d.selectionId));

            legendItems.append('span').classed('legend-pill', true)
                .style('background', (d, i) => colors[i % colors.length])
                .style('box-shadow', '0 0 8px rgba(0,0,0,0.15)');

            const textGroup = legendItems.append('div').classed('legend-text', true);
            textGroup.append('span').classed('legend-label', true).text((d) => d.label);
            this.currentLegendItems = legendList.selectAll('.legend-item');

            if (isLegendRowLayout && (legendPlacement === 'top' || legendPlacement === 'bottom')) {
                const legendNode = legendContainer.node() as HTMLElement | null;
                if (legendNode) {
                    const computedWidth = Math.min(width, legendNode.getBoundingClientRect().width);
                    const left = Math.max(0, Math.floor((width - computedWidth) / 2));
                    legendContainer.style('width', `${computedWidth}px`).style('left', `${left}px`);
                }
            }
        }

            this.applySelectionState();
            this.updateDetailPanel(totalValue);
        }

        private handleSliceSelection(event: MouseEvent, selectionId: powerbi.extensibility.ISelectionId | undefined, totalValue: number): void {
            if (!this.canInteract()) {
                return;
            }
            if (!selectionId) {
                return;
            }
            event.stopPropagation();
            const multiSelect = event.ctrlKey || event.metaKey;
            const selectionResult = this.selectionManager.select(selectionId, multiSelect);
            this.processSelectionResult(selectionResult, totalValue);
        }

        private processSelectionResult(selectionResult: any, totalValue: number): void {
            Promise.resolve(selectionResult).then((ids) => {
                this.activeSelection = ids || [];
                this.applySelectionState();
                this.updateDetailPanel(totalValue);
            });
        }

        private handleContextMenu(event: MouseEvent, selectionId: powerbi.extensibility.ISelectionId | undefined): void {
            if (!this.canInteract()) {
                return;
            }
            if (!selectionId) {
                return;
            }
            event.preventDefault();
            event.stopPropagation();
            this.selectionManager.showContextMenu(selectionId, { x: event.clientX, y: event.clientY });
        }

        private applySelectionState(): void {
            const hasSelection = this.activeSelection && this.activeSelection.length > 0;
            const isSelectionActive = (id?: powerbi.extensibility.ISelectionId) => {
                if (!id || !hasSelection) {
                    return false;
                }
                return this.activeSelection.some((sel) => this.compareSelectionIds(sel, id));
            };

            this.currentSlices?.classed('slice-dimmed', d => hasSelection && !isSelectionActive(d.data.selectionId))
                .classed('slice-selected', d => isSelectionActive(d.data.selectionId))
                .attr('opacity', d => (!hasSelection || isSelectionActive(d.data.selectionId)) ? 1 : 0.35);

            this.currentLegendItems?.classed('legend-item--dimmed', d => hasSelection && !isSelectionActive(d.selectionId))
                .classed('legend-item--selected', d => isSelectionActive(d.selectionId));
        }

        private updateDetailPanel(totalValue: number): void {
            if (!this.detailPanel) {
                return;
            }
            const data = this.lastData || [];
            const selectedItems = data.filter((item) => this.isSelectionActive(item.selectionId));
            const selectedSum = selectedItems.reduce((sum, item) => sum + item.value, 0);
            const sorted = [...data].sort((a, b) => b.value - a.value).slice(0, 3);

            const selectionHint = selectedItems.length
                ? `Selected ${selectedItems.length} slice${selectedItems.length > 1 ? 's' : ''}`
                : 'Click a slice to filter other visuals';

            const detailPanel = this.detailPanel;
            detailPanel.selectAll('*').remove();
            detailPanel.append('div').classed('detail-panel__title', true).text('Overview');
            const metricRow = detailPanel.append('div').classed('detail-panel__metric', true);
            metricRow.append('span').classed('detail-panel__metric-label', true).text('Total');
            metricRow.append('strong').classed('detail-panel__metric-value', true)
                .text(this.formatNumericValue(totalValue, this.detailPanelDecimals, 'auto'));
            const stack = detailPanel.append('div').classed('detail-panel__stack', true);

            sorted.forEach((item) => {
                const row = stack.append('div').classed('detail-panel__stat-row', true);
                row.append('span').classed('detail-panel__stat-label', true).text(item.label);
                row.append('span').classed('detail-panel__stat-value', true).text(this.formatNumericValue(item.value, this.detailPanelDecimals, 'auto'));
            });

            if (selectedItems.length) {
                const filteredRow = stack.append('div').classed('detail-panel__stat-row', true);
                filteredRow.append('span').classed('detail-panel__stat-label', true).text('Filtered sum');
                filteredRow.append('span').classed('detail-panel__stat-value', true).text(this.formatNumericValue(selectedSum, this.detailPanelDecimals, 'auto'));
            } else {
                stack.append('div').classed('detail-panel__prompt', true).text(selectionHint);
            }
        }

        private compareSelectionIds(a?: any, b?: any): boolean {
            if (!a || !b) {
                return false;
            }
            if (typeof a.equals === 'function') {
                return a.equals(b);
            }
            if (typeof b.equals === 'function') {
                return b.equals(a);
            }
            return a === b;
        }

        private isSelectionActive(id?: powerbi.extensibility.ISelectionId): boolean {
            if (!id || !this.activeSelection.length) {
                return false;
            }
            return this.activeSelection.some((sel) => this.compareSelectionIds(sel, id));
        }


    /**
     * Returns properties pane formatting model content hierarchies, properties and latest formatting values, Then populate properties pane.
     * This method is called once every time we open properties pane or when the user edit any format property. 
     */
    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

}