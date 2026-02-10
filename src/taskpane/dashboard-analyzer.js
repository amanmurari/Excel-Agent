/**
 * DashboardAnalyzer - Core analysis engine for Excel data
 * Extends ExcelAIAssistant to provide higher-level dashboard and analysis features
 */
class DashboardAnalyzer extends ExcelAIAssistant {
    constructor(apiKey) {
        super(apiKey);
    }

    /**
     * Performs complete analysis and returns structured result
     */
    async performCompleteAnalysis(dataAddress = "selection") {
        try {
            const context = await this.getExcelContext(dataAddress);

            const prompt = `You are a World-Class Data Scientist and Business Consultant. Perform an ELITE-LEVEL analysis of this dataset.

DATA CONTEXT:
Summary: ${JSON.stringify(context.summary)}
Address: ${context.address}
Sample Data: ${JSON.stringify(context.dataSample)}

YOUR ANALYSIS REQUIREMENTS:
1. **Statistical Significance**: Identify patterns, anomalies, and correlations with mathematical rigor.
2. **Segment Performance**: Breakthrough the data to find which categories/segments are driving value or causing losses.
3. **Executive Recommendations**: Provide 3 actionable strategic recommendations based on the findings.
4. **Visualization Strategy**: Recommend high-impact visualizations (Line, Bar, Pie, Scatter) to communicate these insights to a CEO.

RETURN JSON FORMAT (MANDATORY):
{
  "summary": "High-level executive overview (1-2 sentences)",
  "detailedInsights": [
    {
      "title": "Insight Title",
      "observation": "What the data shows",
      "implication": "What this means for the business",
      "confidence": "High/Medium/Low"
    }
  ],
  "recommendations": [
    "Strategic recommendation 1",
    "Strategic recommendation 2"
  ],
  "charts": [
    {
      "type": "LineChart", 
      "title": "Mandatory Professional Title", 
      "columns": ["X-Axis Col", "Y-Axis Col"],
      "rationale": "Why this chart is critical for the executive"
    }
  ],
  "statisticalSummary": {
    "keyMetrics": "...",
    "anomalies": "..."
  }
}`;

            const response = await this.callClaudeAPI(prompt);
            const analysis = this.extractJSON(response);

            if (!analysis) {
                throw new Error("Failed to parse analysis response into structured JSON.");
            }

            return {
                status: "success",
                address: context.address,
                analysis: analysis
            };
        } catch (error) {
            console.error("Analysis error:", error);
            throw error;
        }
    }

    /**
     * Helper: Get data range for specific columns
     */
    async getDataRangeForColumns(columnNames) {
        if (!columnNames || !Array.isArray(columnNames) || columnNames.length === 0) return null;

        const data = await this.getWorksheetData();
        const headers = data.values[0];
        const rowCount = data.rowCount;

        // Find indices
        const indices = [];
        for (const colName of columnNames) {
            if (!colName) continue;
            // Fuzzy match header
            const index = headers.findIndex(h =>
                String(h).toLowerCase().trim() === String(colName).toLowerCase().trim()
            );
            if (index !== -1) indices.push(index);
        }

        if (indices.length === 0) return null;

        // Sort indices to ensure range order
        indices.sort((a, b) => a - b);

        // Helper to convert index to letter (0 -> A, 25 -> Z, 26 -> AA)
        const getColLetter = (colIndex) => {
            let letter = '';
            let temp = colIndex;
            while (temp >= 0) {
                letter = String.fromCharCode((temp % 26) + 65) + letter;
                temp = Math.floor(temp / 26) - 1;
            }
            return letter;
        };

        // If indices are contiguous, return simplest range
        let isContiguous = true;
        for (let i = 0; i < indices.length - 1; i++) {
            if (indices[i + 1] !== indices[i] + 1) {
                isContiguous = false;
                break;
            }
        }

        const startRow = 1; // Assuming headers are row 1
        const endRow = startRow + rowCount - 1;

        try {
            if (isContiguous) {
                const startCol = getColLetter(indices[0]);
                const endCol = getColLetter(indices[indices.length - 1]);
                return `${startCol}${startRow}:${endCol}${endRow}`;
            } else {
                return indices.map(idx => {
                    const col = getColLetter(idx);
                    return `${col}${startRow}:${col}${endRow}`;
                }).join(',');
            }
        } catch (e) {
            console.error("Error constructing range:", e);
            return null;
        }
    }

    /**
     * Helper: Prepare chart data on a temporary sheet (handling disjoint columns)
     */
    async prepareChartDataOnTempSheet(columns, context) {
        if (!columns || columns.length === 0) return null;

        // 1. Get source data
        const activeSheet = context.workbook.worksheets.getActiveWorksheet();
        const sourceRange = activeSheet.getUsedRange();
        sourceRange.load("values");
        await context.sync();

        const headerRow = sourceRange.values[0];

        // 2. Identify column indices
        const indices = [];
        columns.forEach(colName => {
            const idx = headerRow.findIndex(h => String(h).toLowerCase().trim() === String(colName).toLowerCase().trim());
            if (idx !== -1) indices.push(idx);
        });

        if (indices.length === 0) return null;

        // 3. Get or Create Temp Sheet
        let tempSheet = context.workbook.worksheets.getItemOrNullObject("Temp_Chart_Data");
        await context.sync();

        if (tempSheet.isNullObject) {
            tempSheet = context.workbook.worksheets.add("Temp_Chart_Data");
        }

        // Use next empty columns to avoid overwriting if we process multiple charts
        const usedRange = tempSheet.getUsedRangeOrNullObject();
        await context.sync();
        let startColIndex = 0;
        if (!usedRange.isNullObject) {
            usedRange.load(["columnCount", "columnIndex"]);
            await context.sync();
            startColIndex = usedRange.columnIndex + usedRange.columnCount + 1;
        }

        // Helper to convert index to letter
        const getColLetter = (colIndex) => {
            let letter = '';
            let temp = colIndex;
            while (temp >= 0) {
                letter = String.fromCharCode((temp % 26) + 65) + letter;
                temp = Math.floor(temp / 26) - 1;
            }
            return letter;
        };

        // 4. Copy each column to Temp Sheet
        let currentRowCount = sourceRange.values.length;

        for (let i = 0; i < indices.length; i++) {
            const srcColIdx = indices[i];
            const srcColLetter = getColLetter(srcColIdx);
            const srcAddress = `${srcColLetter}1:${srcColLetter}${currentRowCount}`;
            // Use internal helper for safe range access
            const srcRange = await this.getRangeFromString(context, srcAddress);

            const destColLetter = getColLetter(startColIndex + i);
            const destAddress = `Temp_Chart_Data!${destColLetter}1:${destColLetter}${currentRowCount}`;
            const destRange = await this.getRangeFromString(context, destAddress);

            destRange.copyFrom(srcRange, Excel.RangeCopyType.all);
        }

        // Sync to ensure all data is written before returning
        await context.sync();

        // 5. Return the new contiguous range address on Temp Sheet
        const rangeStart = getColLetter(startColIndex);
        const rangeEnd = getColLetter(startColIndex + indices.length - 1);
        const rangeAddress = `Temp_Chart_Data!${rangeStart}1:${rangeEnd}${currentRowCount}`;

        console.log(`✅ Prepared temp sheet data at: ${rangeAddress}`);
        return rangeAddress;
    }

    /**
     * Filter data based on conditions and return results on a new sheet
     */
    async filterData(column, operator, value, dataRange = "worksheet") {
        console.log(`[DashboardAnalyzer] Filtering ${dataRange} by ${column} ${operator} ${value}`);
        const data = await (dataRange === "worksheet" ? this.getWorksheetData() : this.getRangeByAddress(dataRange));
        const headers = data.values[0];
        const colIndex = headers.findIndex(h => String(h).toLowerCase().trim() === String(column).toLowerCase().trim());

        if (colIndex === -1) throw new Error(`Column "${column}" not found in data.`);

        const filteredRows = [headers];
        for (let i = 1; i < data.values.length; i++) {
            const row = data.values[i];
            const cellValue = row[colIndex];
            let match = false;

            const v1 = String(cellValue).toLowerCase();
            const v2 = String(value).toLowerCase();

            switch (operator.toLowerCase()) {
                case 'equals': match = v1 === v2; break;
                case 'not_equals': match = v1 !== v2; break;
                case 'contains': match = v1.includes(v2); break;
                case 'greater_than': match = parseFloat(cellValue) > parseFloat(value); break;
                case 'less_than': match = parseFloat(cellValue) < parseFloat(value); break;
                default: match = v1 === v2;
            }
            if (match) filteredRows.push(row);
        }

        return await this.exportToNewSheet(filteredRows, `Filtered_${column.substring(0, 10)}`);
    }

    /**
     * Sort data by multiple columns
     */
    async sortData(columns, order = 'asc', dataRange = "worksheet") {
        console.log(`[DashboardAnalyzer] Sorting ${dataRange} by ${columns}`);
        const data = await (dataRange === "worksheet" ? this.getWorksheetData() : this.getRangeByAddress(dataRange));
        const headers = data.values[0];
        const rows = data.values.slice(1);

        const colIndices = columns.map(col => ({
            index: headers.findIndex(h => String(h).toLowerCase().trim() === String(col).toLowerCase().trim()),
            name: col
        }));

        rows.sort((a, b) => {
            for (let i = 0; i < colIndices.length; i++) {
                const { index } = colIndices[i];
                if (index === -1) continue;

                const valA = a[index];
                const valB = b[index];

                if (valA < valB) return order === 'asc' ? -1 : 1;
                if (valA > valB) return order === 'asc' ? 1 : -1;
            }
            return 0;
        });

        const sortedData = [headers, ...rows];
        return await this.exportToNewSheet(sortedData, `Sorted_Data`);
    }

    /**
     * Merge multiple ranges/datasets
     */
    async mergeData(ranges, mergeType = 'vertical') {
        console.log(`[DashboardAnalyzer] Merging ${ranges.length} ranges ${mergeType}`);
        let mergedData = [];

        for (const r of ranges) {
            const data = await this.getRangeByAddress(r);
            if (mergeType === 'vertical') {
                if (mergedData.length === 0) {
                    mergedData = data.values;
                } else {
                    // Skip headers for subsequent ranges if it's vertical stacking
                    mergedData = mergedData.concat(data.values.slice(1));
                }
            } else {
                // Horizontal merge
                if (mergedData.length === 0) {
                    mergedData = data.values;
                } else {
                    for (let i = 0; i < Math.min(mergedData.length, data.values.length); i++) {
                        mergedData[i] = mergedData[i].concat(data.values[i]);
                    }
                }
            }
        }

        return await this.exportToNewSheet(mergedData, `Merged_Result`);
    }

    /**
     * Export data to a new sheet and return the address
     */
    async exportToNewSheet(data, sheetName = null) {
        const finalSheetName = sheetName || `Export_${new Date().getTime().toString().slice(-4)}`;
        await this.createNewSheet(finalSheetName);
        const result = await this.writeToRange(`${finalSheetName}!A1`, data);

        // Calculate the full range address
        const colCount = data[0] ? data[0].length : 0;
        const endCol = this.columnIndexToLetter(colCount - 1);
        const fullAddress = `${finalSheetName}!A1:${endCol}${data.length}`;

        return {
            status: 'success',
            address: fullAddress,
            message: `Data exported to ${finalSheetName}`,
            sheetName: finalSheetName
        };
    }

    /**
     * Validate data quality
     */
    async validateData(rules, dataRange = "worksheet") {
        const data = await (dataRange === "worksheet" ? this.getWorksheetData() : this.getRangeByAddress(dataRange));
        const dataStr = this.formatDataForAI(data.values, data.address);

        const prompt = `Perform data quality validation on this dataset based on rules: ${typeof rules === 'string' ? rules : JSON.stringify(rules)}
        
        Dataset:
        ${dataStr}
        
        Return JSON with:
        {
          "dataQualityScore": 1-100,
          "issuesFound": [{"row": N, "column": "Name", "issue": "Missing value", "severity": "High"}],
          "summary": "Overall data health summary",
          "recommendations": ["Action to fix data"]
        }`;

        const response = await this.callClaudeAPI(prompt, "Return only valid JSON.");
        return this.extractJSON(response);
    }

    /**
     * Trace error source
     */
    async traceError(address) {
        console.log(`[DashboardAnalyzer] Tracing error in ${address}`);
        return await Excel.run(async (context) => {
            const range = await this.getRangeFromString(context, address);
            range.load(["formulas", "values", "valueTypes"]);
            await context.sync();

            // Basic check
            if (range.valueTypes[0][0] !== Excel.RangeValueType.error) {
                return { status: "no_error", message: "Cell does not contain an error." };
            }

            // Try to get precedents (if supported)
            let precedentsInfo = [];
            try {
                // This might fail in older Excel versions
                const precedents = range.getPrecedents();
                precedents.load("areas");
                await context.sync();

                for (let area of precedents.areas.items) {
                    area.load(["address", "values"]);
                    await context.sync();
                    precedentsInfo.push({
                        address: area.address,
                        values: area.values
                    });
                }
            } catch (e) {
                console.warn("getPrecedents not supported or failed:", e);
                precedentsInfo = ["Could not automatically trace precedents. Please analyze formula text."];
            }

            return {
                status: "success",
                address: address,
                errorValue: range.values[0][0],
                formula: range.formulas[0][0],
                precedents: precedentsInfo,
                message: "Error traced. Analyze the precedents and formula."
            };
        });
    }

    /**
     * Explain formula
     */
    async explainFormula(address) {
        console.log(`[DashboardAnalyzer] Explaining formula in ${address}`);
        const context = await this.getExcelContext(address);

        // We need the formula, not just values
        const rangeData = await this.getRangeByAddress(address);
        const formula = rangeData.formulas ? rangeData.formulas[0][0] : null;

        if (!formula || !formula.startsWith('=')) {
            return { status: "info", message: "Cell does not contain a formula." };
        }

        const prompt = `Explain this Excel formula in plain English for a business user.
        
        Formula: ${formula}
        Cell: ${address}
        
        Context (Precedents/Data):
        (AI should infer based on standard Excel syntax)
        
        Provide:
        1. Simple English explanation.
        2. Breakdown of logic.
        3. Potential risks or dependencies.`;

        const explanation = await this.callClaudeAPI(prompt, "You are an expert Excel educator.");

        return {
            status: "success",
            address: address,
            formula: formula,
            explanation: explanation
        };
    }

    /**
     * Create scenario
     */
    async createScenario(scenarioName, changes) {
        console.log(`[DashboardAnalyzer] Creating scenario: ${scenarioName}`);

        return await Excel.run(async (context) => {
            const activeSheet = context.workbook.worksheets.getActiveWorksheet();
            activeSheet.load("name");
            await context.sync();

            // Copy sheet
            const newSheet = activeSheet.copy(Excel.WorksheetPositionType.after, activeSheet);
            newSheet.name = `${scenarioName.substring(0, 20)}_${Math.floor(Math.random() * 1000)}`;
            await context.sync();

            // Apply changes
            for (const change of changes) {
                const range = newSheet.getRange(change.address);
                range.values = [[change.value]];
                // Highlight change
                range.format.fill.color = "#FFF2CC"; // Light yellow
                range.format.font.color = "#C00000"; // Red text
            }

            await context.sync();

            return {
                status: "success",
                originalSheet: activeSheet.name,
                scenarioSheet: newSheet.name,
                changesApplied: changes.length,
                message: `Scenario '${scenarioName}' created on new sheet '${newSheet.name}'. Changes highlighted.`
            };
        });
    }

    /**
     * Apply a formula to an entire column
     */
    async applyFormulaToColumn(formula, targetColumn, sourceColumnForHeight = "A") {
        console.log(`[DashboardAnalyzer] Applying formula ${formula} to column ${targetColumn}`);

        // Guardrail: Prevent Table Structured References
        if (formula.includes("[@") || formula.match(/\[.*?\]/)) {
            throw new Error(`Formula Syntax Error: The agent attempted to use a Table Reference (e.g. [@Name] or [Column]). \n\nSuggested Fix: Retry using standard cell references (e.g. assume 'Name' is Column A, so use 'A2'). Formula provided: ${formula}`);
        }

        // Guardrail: Validate Target Column is a specific letter (e.g. "M", not "New Column")
        if (!targetColumn.match(/^[A-Za-z]{1,3}$/)) {
            throw new Error(`Parameter Error: 'targetColumn' must be a valid Excel Column LETTER (e.g. 'M', 'AA'), not a name like "${targetColumn}". Please check the available workspace and pick an empty column letter.`);
        }

        return await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();

            // 1. Determine height based on source column (default A)
            // We use getRange to safely access the column
            const sourceCol = sheet.getRange(`${sourceColumnForHeight}:${sourceColumnForHeight}`);
            const sourceRange = sourceCol.getUsedRange();
            sourceRange.load(["rowCount", "rowIndex"]);
            await context.sync();

            // Calculate absolute end row
            const totalRows = sourceRange.rowIndex + sourceRange.rowCount;
            const startRow = 2; // Assuming Row 1 is header

            if (totalRows < startRow) {
                return { status: "error", message: "No data found to apply formula against." };
            }

            // 2. Set header
            const headerCell = sheet.getRange(`${targetColumn}1`);
            headerCell.values = [[`Calc_${targetColumn}`]];
            headerCell.format.font.bold = true;

            // 3. Set formula in first cell and autofill
            const firstCell = sheet.getRange(`${targetColumn}${startRow}`);
            firstCell.formulas = [[formula]];

            if (totalRows > startRow) {
                const fillRange = sheet.getRange(`${targetColumn}${startRow}:${targetColumn}${totalRows}`);
                firstCell.autoFill(fillRange, Excel.AutoFillType.fillDefault);
            }

            await context.sync();

            return {
                status: "success",
                address: `${targetColumn}${startRow}:${targetColumn}${totalRows}`,
                message: `Applied ${formula} to column ${targetColumn} (Rows ${startRow}-${totalRows})`
            };
        });
    }
}

// Export class
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { DashboardAnalyzer };
}

// Ensure global availability
if (typeof window !== 'undefined') {
    window.DashboardAnalyzer = DashboardAnalyzer;
}
