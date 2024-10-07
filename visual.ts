"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;

import { VisualFormattingSettingsModel } from "./settings";

export class Visual implements IVisual {
    private target: HTMLElement;
    private updateCount: number;
    private textNode: Text;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    private table: HTMLTableElement; // Table element for dynamic data

    constructor(options: VisualConstructorOptions) {
        console.log('Visual constructor', options);
        this.formattingSettingsService = new FormattingSettingsService();
        this.target = options.element;
        this.updateCount = 0;

        // Create a scrollable container for the table
        const tableContainer: HTMLElement = document.createElement("div");
        tableContainer.style.overflowY = "auto"; // Enable vertical scrolling
        tableContainer.style.maxHeight = "400px"; // Set a fixed max height for the scrollable area
        this.target.appendChild(tableContainer);

        // Create and append a table to the container
        this.table = document.createElement("table");
        this.table.style.width = "100%";
        this.table.style.borderCollapse = "collapse";
        tableContainer.appendChild(this.table);

        // Create paragraph to show update count
        if (document) {
            const new_p: HTMLElement = document.createElement("p");
            new_p.appendChild(document.createTextNode("Update count:"));
            const new_em: HTMLElement = document.createElement("em");
            this.textNode = document.createTextNode(this.updateCount.toString());
            new_em.appendChild(this.textNode);
            new_p.appendChild(new_em);
            this.target.appendChild(new_p);
        }
    }

    public update(options: VisualUpdateOptions) {
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews[0]);

        console.log('Visual update', options);
        
        // Clear previous table content
        this.table.innerHTML = "";

        // Generate table header
        if (options.dataViews && options.dataViews.length > 0) {
            const dataView = options.dataViews[0];
            const categorical = dataView.categorical;
            if (categorical && categorical.categories) {
                const headerRow: HTMLTableRowElement = document.createElement("tr");

                // Create headers based on categories
                categorical.categories.forEach(category => {
                    const headerCell: HTMLTableCellElement = document.createElement("th");
                    headerCell.textContent = category.source.displayName;
                    headerCell.style.border = "1px solid black"; // Border for header cells
                    headerCell.style.padding = "8px"; // Padding for header cells
                    headerCell.style.textAlign = "left"; // Left align header cells
                    headerRow.appendChild(headerCell);
                });

                // Append header row to the table
                this.table.appendChild(headerRow);

                // Generate table rows based on values
                const rowCount = categorical.categories[0].values.length; // Assume all categories have the same row count
                for (let i = 0; i < rowCount; i++) {
                    const dataRow: HTMLTableRowElement = document.createElement("tr");
                    categorical.categories.forEach(category => {
                        const dataCell: HTMLTableCellElement = document.createElement("td");
                        dataCell.textContent = category.values[i]?.toString() || ""; // Display the category value
                        dataCell.style.border = "1px solid black"; // Border for data cells
                        dataCell.style.padding = "8px"; // Padding for data cells
                        dataRow.appendChild(dataCell);
                    });
                    this.table.appendChild(dataRow);
                }
            }
        }

        // Update the count display
        if (this.textNode) {
            this.textNode.textContent = (this.updateCount++).toString();
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
