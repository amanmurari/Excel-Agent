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

            const prompt = `Analyze this Excel data and provide key insights, trends, and suggested visualizations.
      Data Summary: ${JSON.stringify(context.summary)}
      Sample Data: ${JSON.stringify(context.dataSample)}
      
      Return JSON format:
      {
        "summary": "Overall data trend",
        "insights": ["insight 1", "insight 2"],
        "charts": [
          {"type": "LineChart", "title": "Trend", "columns": ["Date", "Value"]}
        ]
      }`;

            const response = await this.callAI(prompt);
            const analysis = JSON.parse(response);

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
            const srcRange = activeSheet.getRange(srcAddress);

            const destColLetter = getColLetter(startColIndex + i);
            const destAddress = `${destColLetter}1:${destColLetter}${currentRowCount}`;
            const destRange = tempSheet.getRange(destAddress);

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
}

// Export class
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { DashboardAnalyzer };
}

// Ensure global availability
if (typeof window !== 'undefined') {
    window.DashboardAnalyzer = DashboardAnalyzer;
}
