// tool-registry.js - Tool Registry for Autonomous Agent (OpenAI Function Calling Format)

/**
 * Tool Registry - Minimal, structured tool definitions for AI function calling
 */

const TOOL_REGISTRY = {
    // Analysis Tools
    analyzeData: {
        name: "analyzeData",
        description: "Performs comprehensive data analysis. Returns strategic insights and success metrics. Always specify a dataRange if referencing a previous step.",
        parameters: {
            type: "object",
            properties: {
                dataRange: {
                    type: "string",
                    description: "Excel range address (e.g., 'A1:C10'), 'selection', or 'result from step X'"
                }
            },
            required: []
        }
    },

    findInsights: {
        name: "findInsights",
        description: "Answers specific questions about the data using AI deep-drill analysis",
        parameters: {
            type: "object",
            properties: {
                question: {
                    type: "string",
                    description: "The specific question to answer"
                },
                dataRange: {
                    type: "string",
                    description: "Excel range address to focus on, or 'result from step X'"
                }
            },
            required: ["question"]
        }
    },







    calculateMetric: {
        name: "calculateMetric",
        description: "Calculates custom metrics based on the data",
        parameters: {
            type: "object",
            properties: {
                metric: {
                    type: "string",
                    description: "Description of the metric to calculate (e.g., 'average order value', 'year-over-year growth')"
                }
            },
            required: ["metric"]
        }
    },

    createMetricTable: {
        name: "createMetricTable",
        description: "Creates a summarized analysis table (e.g., Average Price by Brand). Automatically handles UNIQUE categories and aggregation formulas. Ideal for driving dashboard charts. PREFER placing this on the ACTIVE SHEET (e.g. to the right of data, columns M+) to keep context.",
        parameters: {
            type: "object",
            properties: {
                dataRange: {
                    type: "string",
                    description: "Source data range (e.g., 'A1:Z500' or 'selection')"
                },
                categoryColumn: {
                    type: "string",
                    description: "Column name to group by (e.g., 'Brand', 'Fuel_Type')"
                },
                metricColumn: {
                    type: "string",
                    description: "Column name to aggregate (e.g., 'Price', 'Sales')"
                },
                aggregation: {
                    type: "string",
                    enum: ["Average", "Count", "Sum", "Max", "Min"],
                    description: "Aggregation type"
                },
                targetCell: {
                    type: "string",
                    description: "Starting cell for the summary table (e.g., 'P1', 'Sheet2!A1'). PREFER active sheet cells."
                }
            },
            required: ["dataRange", "categoryColumn", "metricColumn", "aggregation", "targetCell"]
        }
    },

    createPivot: {
        name: "createPivot",
        description: "Creates an advanced pivot table analysis. Transforms raw data into a structured summary with insights. Always specify dataRange.",
        parameters: {
            type: "object",
            properties: {
                dataRange: {
                    type: "string",
                    description: "Source data range (e.g., 'Sheet1!A1:D100', 'selection', or 'result from step X')"
                },
                rows: {
                    type: "array",
                    items: { type: "string" },
                    description: "Column names for pivot rows"
                },
                columns: {
                    type: "array",
                    items: { type: "string" },
                    description: "Column names for pivot columns"
                },
                values: {
                    type: "array",
                    items: { type: "string" },
                    description: "Column names and aggregation for values (e.g., 'sum of Sales')"
                }
            },
            required: ["dataRange", "rows", "values"]
        }
    },

    // Visualization Tools
    moveChart: {
        name: "moveChart",
        description: "Moves a chart to a specific location (top, left) on the active sheet or another sheet.",
        parameters: {
            type: "object",
            properties: {
                chartName: {
                    type: "string",
                    description: "Name of the chart to move (optional if only one chart)"
                },
                targetCell: {
                    type: "string",
                    description: "Top-left cell where the chart should be positioned (e.g., 'Sheet1!D5')"
                }
            },
            required: ["targetCell"]
        }
    },

    createChart: {
        name: "createChart",
        description: "Creates a single chart. STRATEGIC RULE: Always create charts from summarized metric tables (result from createMetricTable) rather than raw data to ensure professional aggregation.",
        parameters: {
            type: "object",
            properties: {
                dataRange: {
                    type: "string",
                    description: "Excel range address (e.g., 'Sheet1!P1:Q10'), 'selection', or 'result from step X'"
                },
                chartType: {
                    type: "string",
                    enum: ["line", "column", "bar", "pie", "area", "scatter", "combo"],
                    description: "Type of chart to create"
                },
                title: {
                    type: "string",
                    description: "Chart title"
                },
                xAxis: {
                    type: "string",
                    description: "Label for X-axis"
                },
                yAxis: {
                    type: "string",
                    description: "Label for Y-axis"
                }
            },
            required: ["dataRange", "chartType", "title"]
        }
    },

    generateDashboard: {
        name: "generateDashboard",
        description: "Architects a complete elite-tier executive dashboard. SOP: 1. Insert columns, 2. Create multiple Metric Tables, 3. Create charts from those tables. Create dashboard elements on the CURRENT SHEET if possible.",
        parameters: {
            type: "object",
            properties: {
                dataRange: {
                    type: "string",
                    description: "The source data for the dashboard (e.g., 'Sheet1!A1:Z500' or 'result from step X')"
                }
            },
            required: []
        }
    },

    // Data Transformation Tools
    cleanData: {
        name: "cleanData",
        description: "Transforms and cleans data based on natural language instructions. Can remove duplicates, fix formatting, fill missing values, standardize text, convert types, split/merge columns. Returns address of cleaned data for use in subsequent steps",
        parameters: {
            type: "object",
            properties: {
                instructions: {
                    type: "string",
                    description: "Natural language description of cleaning/transformation to perform (e.g., 'Remove duplicate rows based on Customer ID', 'Fill missing values in Sales column with 0')"
                },
                dataRange: {
                    type: "string",
                    description: "Excel range to clean (e.g., 'A1:C10', 'Sheet1!A1:B10', 'selection')"
                }
            },
            required: ["instructions"]
        }
    },

    filterData: {
        name: "filterData",
        description: "Filters rows based on conditions. Creates new sheet with filtered data and returns address",
        parameters: {
            type: "object",
            properties: {
                column: {
                    type: "string",
                    description: "Column name to filter on"
                },
                operator: {
                    type: "string",
                    enum: ["equals", "not_equals", "greater_than", "less_than", "contains", "not_contains", "starts_with", "ends_with"],
                    description: "Comparison operator"
                },
                value: {
                    type: ["string", "number", "boolean"],
                    description: "Value to compare against"
                }
            },
            required: ["column", "operator", "value"]
        }
    },

    sortData: {
        name: "sortData",
        description: "Sorts data by one or more columns. Creates new sheet with sorted data and returns address",
        parameters: {
            type: "object",
            properties: {
                columns: {
                    type: "array",
                    items: { type: "string" },
                    description: "Array of column names to sort by (in priority order)"
                },
                order: {
                    type: ["string", "array"],
                    description: "Sort order: 'asc' or 'desc', or array matching columns length",
                    default: "asc"
                }
            },
            required: ["columns"]
        }
    },

    mergeData: {
        name: "mergeData",
        description: "Merges data from multiple ranges horizontally (side-by-side) or vertically (stacked)",
        parameters: {
            type: "object",
            properties: {
                ranges: {
                    type: "array",
                    items: { type: "string" },
                    description: "Array of range addresses to merge"
                },
                mergeType: {
                    type: "string",
                    enum: ["horizontal", "vertical"],
                    description: "How to merge: horizontal (side-by-side) or vertical (stacked)",
                    default: "horizontal"
                }
            },
            required: ["ranges"]
        }
    },

    validateData: {
        name: "validateData",
        description: "Validates data quality and checks for issues like missing values, type mismatches, duplicates, format inconsistencies",
        parameters: {
            type: "object",
            properties: {
                rules: {
                    type: ["string", "object"],
                    description: "Validation rules to check (natural language or structured object)"
                }
            },
            required: ["rules"]
        }
    },

    // Formula & Calculation Tools
    generateFormula: {
        name: "generateFormula",
        description: "Generates Excel formula based on natural language description. If address is provided, formula is automatically applied to that range",
        parameters: {
            type: "object",
            properties: {
                description: {
                    type: "string",
                    description: "Natural language description of what the formula should calculate (e.g., 'Calculate 10% of the value in C2', 'Sum all values in column B')"
                },
                context: {
                    type: ["array", "object"],
                    description: "Current worksheet data for context (auto-populated if not provided)"
                },
                address: {
                    type: "string",
                    description: "Target address where formula will be applied (e.g., 'D2', 'E2:E100', 'selection'). If omitted, formula is just generated and returned"
                }
            },
            required: ["description"]
        }
    },

    insertFormula: {
        name: "insertFormula",
        description: "Inserts a pre-written Excel formula into specified cells. Formula must start with =",
        parameters: {
            type: "object",
            properties: {
                address: {
                    type: "string",
                    description: "Target range where formula will be inserted"
                },
                formula: {
                    type: "string",
                    description: "Excel formula string starting with = (e.g., '=SUM(B2:B100)', '=IF(C2>1000,\"High\",\"Low\")')"
                }
            },
            required: ["address", "formula"]
        }
    },

    applyFormulaToColumn: {
        name: "applyFormulaToColumn",
        description: "Applies a formula to an entire new column, automatically filling it down to match the data height. Use this for creating new metrics, splitting text, or unit conversions.",
        parameters: {
            type: "object",
            properties: {
                formula: {
                    type: "string",
                    description: "Excel formula for the first row using STANDARD CELL REFERENCES (e.g. '=A2*100'). NEVER use Table references like '[@Column]'. Assume data starts at row 2."
                },
                targetColumn: {
                    type: "string",
                    description: "The letter of the new column to create (e.g., 'M', 'Z')"
                }
            },
            required: ["formula", "targetColumn"]
        }
    },

    // Debugging & auditing
    traceError: {
        name: "traceError",
        description: "Traces the source of an error (#REF!, #VALUE!, etc.) in a specific cell. Returns the formula and values of precedent cells to diagnose the root cause.",
        parameters: {
            type: "object",
            properties: {
                address: {
                    type: "string",
                    description: "The cell address containing the error (e.g., 'Sheet1!C5')"
                }
            },
            required: ["address"]
        }
    },

    explainFormula: {
        name: "explainFormula",
        description: "Explains the logic of a complex formula in plain English, breaking down its components and dependencies.",
        parameters: {
            type: "object",
            properties: {
                address: {
                    type: "string",
                    description: "The cell address containing the formula to explain"
                }
            },
            required: ["address"]
        }
    },

    createScenario: {
        name: "createScenario",
        description: "Creates a scenario by copying the current data to a new sheet and updating specific assumptions (input cells) to see the impact on key metrics.",
        parameters: {
            type: "object",
            properties: {
                scenarioName: {
                    type: "string",
                    description: "Name of the scenario (e.g., 'Optimistic Case')"
                },
                changes: {
                    type: "array",
                    items: {
                        type: "object",
                        properties: {
                            address: { type: "string", description: "Cell address to change" },
                            value: { type: ["string", "number"], description: "New value" }
                        }
                    },
                    description: "List of changes to apply in this scenario"
                }
            },
            required: ["scenarioName", "changes"]
        }
    },

    // Sheet & Data Management
    createSheet: {
        name: "createSheet",
        description: "Creates a new worksheet in the workbook",
        parameters: {
            type: "object",
            properties: {
                name: {
                    type: "string",
                    description: "Name for the new sheet"
                }
            },
            required: ["name"]
        }
    },

    insertColumns: {
        name: "insertColumns",
        description: "Inserts one or more columns at a specific location. Use this to create space for dashboards.",
        parameters: {
            type: "object",
            properties: {
                address: {
                    type: "string",
                    description: "Reference column or cell where to insert (e.g., 'P:P' or 'P1')"
                },
                count: {
                    type: "number",
                    description: "Number of columns to insert"
                },
                shift: {
                    type: "string",
                    enum: ["Right"],
                    default: "Right"
                }
            },
            required: ["address", "count"]
        }
    },

    writeData: {
        name: "writeData",
        description: "Writes a 2D array of values to a specific range. Returns address of written data",
        parameters: {
            type: "object",
            properties: {
                address: {
                    type: "string",
                    description: "Target range address where data will be written (e.g., 'A1', 'Sheet2!B5', 'selection')"
                },
                data: {
                    type: "array",
                    items: { type: "array" },
                    description: "2D array of values to write. Each inner array is a row"
                }
            },
            required: ["address", "data"]
        }
    },

    exportToNewSheet: {
        name: "exportToNewSheet",
        description: "Copies data to a new worksheet. Returns address of exported data",
        parameters: {
            type: "object",
            properties: {
                data: {
                    type: "array",
                    items: { type: "array" },
                    description: "2D array of data to export. If not provided, uses current worksheet data"
                },
                sheetName: {
                    type: "string",
                    description: "Name for the new sheet. If not provided, auto-generates name"
                }
            },
            required: []
        }
    },

    createSummary: {
        name: "createSummary",
        description: "Generates a summary sheet with key metrics, aggregations, and insights from the data",
        parameters: {
            type: "object",
            properties: {},
            required: []
        }
    },

    // Formatting
    formatData: {
        name: "formatData",
        description: "Applies formatting to cells including colors, fonts, borders, number formats",
        parameters: {
            type: "object",
            properties: {
                range: {
                    type: "string",
                    description: "Range to format"
                },
                formatting: {
                    type: "object",
                    description: "Formatting options (e.g., {backgroundColor: '#FFE699', fontBold: true, numberFormat: '$#,##0.00'})"
                }
            },
            required: ["range", "formatting"]
        }
    },

    // Forecasting

    // ============================================
    // FINANCIAL ANALYSIS TOOLS
    // ============================================

    calculateComparableCompanyMultiples: {
        name: "calculateComparableCompanyMultiples",
        description: "Calculate trading multiples (EV/Revenue, EV/EBITDA, P/E, EV/FCF) for comparable company analysis. Creates summary table with min/max/median/mean statistics for peer benchmarking.",
        parameters: {
            type: "object",
            properties: {
                dataRange: {
                    type: "string",
                    description: "Range containing company data with columns: Company Name, Market Cap, Enterprise Value, Revenue, EBITDA, Net Income, and optionally FCF. Example: 'A1:G10' or 'CompData!A1:G10'"
                },
                targetCell: {
                    type: "string",
                    description: "Where to place the multiples analysis table (e.g., 'M1', 'Multiples!A1')"
                }
            },
            required: ["dataRange", "targetCell"]
        }
    },

    analyzePrecedentTransactions: {
        name: "analyzePrecedentTransactions",
        description: "Analyze M&A transaction multiples and deal details. Calculates EV/Revenue, EV/EBITDA multiples from historical deals, identifies trends, and highlights premium/discount to trading multiples.",
        parameters: {
            type: "object",
            properties: {
                dataRange: {
                    type: "string",
                    description: "Range with columns: Date, Acquirer, Target, Deal Value, Target Revenue, Target EBITDA, Target Market Cap (if public). Example: 'Transactions!A1:H20'"
                },
                targetCell: {
                    type: "string",
                    description: "Where to place the transaction analysis table"
                },
                includeChronologicalTrend: {
                    type: "boolean",
                    description: "Whether to show how multiples have trended over time",
                    default: true
                }
            },
            required: ["dataRange", "targetCell"]
        }
    },

    buildHistoricalFinancials: {
        name: "buildHistoricalFinancials",
        description: "Aggregate and analyze 3-5 year historical financial data. Calculates growth rates (CAGR), margins, trends, and key performance metrics for financial modeling foundation.",
        parameters: {
            type: "object",
            properties: {
                dataRange: {
                    type: "string",
                    description: "Range containing multi-year financials. Rows = line items (Revenue, COGS, EBITDA, etc.), Columns = years. Include row headers."
                },
                targetCell: {
                    type: "string",
                    description: "Where to place the analyzed historical financials with growth rates and margins"
                },
                includeRatios: {
                    type: "boolean",
                    description: "Calculate profitability margins (Gross Margin %, EBITDA Margin %, Net Margin %)",
                    default: true
                }
            },
            required: ["dataRange", "targetCell"]
        }
    },

    buildThreeStatementModel: {
        name: "buildThreeStatementModel",
        description: "Build integrated three-statement financial model (Income Statement, Balance Sheet, Cash Flow). Links statements via retained earnings, depreciation & amortization, working capital changes, and debt/equity.",
        parameters: {
            type: "object",
            properties: {
                historicalDataRange: {
                    type: "string",
                    description: "Range containing historical financial statements (3-5 years of actual data)"
                },
                assumptionsRange: {
                    type: "string",
                    description: "Range with projection assumptions: Revenue Growth %, Gross Margin %, EBITDA Margin %, Capex % of Revenue, Tax Rate %, Working Capital assumptions, Debt schedule"
                },
                targetCell: {
                    type: "string",
                    description: "Where to place the three-statement model (requires significant space)"
                },
                projectionYears: {
                    type: "number",
                    description: "Number of years to project forward (typically 5)",
                    default: 5
                }
            },
            required: ["historicalDataRange", "assumptionsRange", "targetCell"]
        }
    },

    buildDCFModel: {
        name: "buildDCFModel",
        description: "Build a complete Discounted Cash Flow (DCF) valuation model. Projects unlevered free cash flow, calculates terminal value, applies WACC discount rate, and outputs enterprise and equity value with sensitivity analysis.",
        parameters: {
            type: "object",
            properties: {
                financialsRange: {
                    type: "string",
                    description: "Range containing projected financials (Revenue, EBITDA, D&A, Capex, NWC Change, Tax Rate). Can be output from buildThreeStatementModel."
                },
                assumptionsRange: {
                    type: "string",
                    description: "Range with DCF assumptions: WACC (%), Terminal Growth Rate (%), Net Debt, Shares Outstanding, Current Stock Price (optional)"
                },
                targetCell: {
                    type: "string",
                    description: "Where to place the DCF analysis"
                },
                includeSensitivity: {
                    type: "boolean",
                    description: "Create sensitivity tables for WACC and terminal growth variations",
                    default: true
                }
            },
            required: ["financialsRange", "assumptionsRange", "targetCell"]
        }
    },

    buildLBOModel: {
        name: "buildLBOModel",
        description: "Build a Leveraged Buyout (LBO) analysis model. Calculates sources & uses, debt paydown schedule, exit value scenarios, IRR, and MOIC (Money on Money) returns. Tests multiple exit multiples and hold periods.",
        parameters: {
            type: "object",
            properties: {
                financialsRange: {
                    type: "string",
                    description: "Range with projected financials (Revenue, EBITDA, EBIT, FCF projections)"
                },
                lboAssumptionsRange: {
                    type: "string",
                    description: "Range with: Entry EBITDA Multiple, Purchase Price, % Debt, % Equity, Interest Rate, Debt Term, Exit Year, Exit EBITDA Multiple(s)"
                },
                targetCell: {
                    type: "string",
                    description: "Where to place the LBO model"
                },
                exitScenarios: {
                    type: "array",
                    items: { type: "number" },
                    description: "Array of exit EBITDA multiples to test (e.g., [7, 8, 9, 10])",
                    default: [7, 8, 9, 10]
                }
            },
            required: ["financialsRange", "lboAssumptionsRange", "targetCell"]
        }
    },

    calculateFinancialRatios: {
        name: "calculateFinancialRatios",
        description: "Calculate comprehensive financial ratios across categories: Profitability (ROE, ROA, ROIC, margins), Liquidity (Current Ratio, Quick Ratio), Leverage (Debt/Equity, Interest Coverage), and Efficiency (Asset Turnover, Days Sales Outstanding).",
        parameters: {
            type: "object",
            properties: {
                incomeStatementRange: {
                    type: "string",
                    description: "Range containing income statement data (Revenue, COGS, Operating Income, Net Income, etc.)"
                },
                balanceSheetRange: {
                    type: "string",
                    description: "Range containing balance sheet data (Assets, Liabilities, Equity, Current Assets/Liabilities, Debt, etc.)"
                },
                targetCell: {
                    type: "string",
                    description: "Where to place the ratios analysis table"
                },
                periods: {
                    type: "array",
                    items: { type: "string" },
                    description: "Period labels (e.g., ['2022', '2023', '2024'])"
                }
            },
            required: ["incomeStatementRange", "balanceSheetRange", "targetCell"]
        }
    },

    benchmarkAgainstPeers: {
        name: "benchmarkAgainstPeers",
        description: "Compare target company metrics against peer group. Calculates percentile rankings, identifies outliers, shows variance from median, and provides strategic recommendations for areas of outperformance/underperformance.",
        parameters: {
            type: "object",
            properties: {
                targetCompanyRange: {
                    type: "string",
                    description: "Range with target company metrics (Revenue, EBITDA Margin %, Revenue Growth %, ROE %, etc.)"
                },
                peerGroupRange: {
                    type: "string",
                    description: "Range with peer company data (same metrics as target). Each row = one peer company."
                },
                targetCell: {
                    type: "string",
                    description: "Where to place the benchmarking analysis"
                },
                metricsToCompare: {
                    type: "array",
                    items: { type: "string" },
                    description: "Specific metrics to benchmark (e.g., ['Revenue Growth %', 'EBITDA Margin %', 'ROE %']). If omitted, compares all available metrics."
                }
            },
            required: ["targetCompanyRange", "peerGroupRange", "targetCell"]
        }
    }
};

/**
 * Get tool by name (with fuzzy matching)
 */
function getTool(toolName) {
    const normalizedName = toolName.toLowerCase().replace(/_/g, '');

    for (const [key, tool] of Object.entries(TOOL_REGISTRY)) {
        if (key.toLowerCase() === normalizedName) {
            return tool;
        }
    }

    return null;
}

/**
 * Get tools in OpenAI function calling format
 */
function getToolsForFunctionCalling() {
    return Object.values(TOOL_REGISTRY).map(tool => ({
        type: "function",
        function: {
            name: tool.name,
            description: tool.description,
            parameters: tool.parameters
        }
    }));
}

/**
 * Get formatted tool descriptions for AI prompts (concise version)
 */
function getToolDescriptionsForPrompt() {
    const tools = Object.values(TOOL_REGISTRY);

    let description = "AVAILABLE TOOLS:\n\n";

    tools.forEach((tool, index) => {
        const params = tool.parameters.required.length > 0
            ? `(${tool.parameters.required.join(', ')})`
            : '()';
        description += `${index + 1}. ${tool.name}${params} - ${tool.description}\n`;
    });

    return description;
}

/**
 * Validate tool call parameters
 */
function validateToolCall(toolName, parameters) {
    const tool = getTool(toolName);
    if (!tool) {
        return { valid: false, error: `Unknown tool: ${toolName}` };
    }

    const required = tool.parameters.required || [];
    const missing = required.filter(param => !(param in parameters));

    if (missing.length > 0) {
        return {
            valid: false,
            error: `Missing required parameters: ${missing.join(', ')}`
        };
    }

    return { valid: true };
}

// Export for Node.js
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        TOOL_REGISTRY,
        getToolsForFunctionCalling,
        getTool,
        getToolDescriptionsForPrompt,
        validateToolCall
    };
}

// Ensure global availability for browser
if (typeof window !== 'undefined') {
    window.TOOL_REGISTRY = TOOL_REGISTRY;
    window.getToolsForFunctionCalling = getToolsForFunctionCalling;
    window.getTool = getTool;
    window.getToolDescriptionsForPrompt = getToolDescriptionsForPrompt;
    window.validateToolCall = validateToolCall;
}
