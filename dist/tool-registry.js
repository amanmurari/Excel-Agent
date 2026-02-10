// tool-registry.js - Tool Registry for Autonomous Agent (OpenAI Function Calling Format)

/**
 * Tool Registry - Minimal, structured tool definitions for AI function calling
 */

const TOOL_REGISTRY = {
    // Analysis Tools
    analyzeData: {
        name: "analyzeData",
        description: "Performs comprehensive data analysis including descriptive statistics, trends, correlations, outliers, forecasting, and segmentation",
        parameters: {
            type: "object",
            properties: {},
            required: []
        }
    },

    findInsights: {
        name: "findInsights",
        description: "Answers specific questions about the data using AI analysis",
        parameters: {
            type: "object",
            properties: {
                question: {
                    type: "string",
                    description: "The specific question to answer about the data"
                }
            },
            required: ["question"]
        }
    },

    detectOutliers: {
        name: "detectOutliers",
        description: "Identifies outliers and anomalies in the data using statistical methods",
        parameters: {
            type: "object",
            properties: {},
            required: []
        }
    },

    correlationAnalysis: {
        name: "correlationAnalysis",
        description: "Analyzes correlations between variables to find relationships",
        parameters: {
            type: "object",
            properties: {},
            required: []
        }
    },

    trendAnalysis: {
        name: "trendAnalysis",
        description: "Analyzes trends over time including direction, strength, and patterns",
        parameters: {
            type: "object",
            properties: {},
            required: []
        }
    },

    segmentation: {
        name: "segmentation",
        description: "Performs clustering/segmentation analysis to group similar data points",
        parameters: {
            type: "object",
            properties: {},
            required: []
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

    createPivot: {
        name: "createPivot",
        description: "Creates a pivot table analysis with specified rows, columns, and values",
        parameters: {
            type: "object",
            properties: {
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
            required: ["rows", "values"]
        }
    },

    // Visualization Tools
    createChart: {
        name: "createChart",
        description: "Creates a single chart from specified data range. Supports line, column, bar, pie, area, scatter chart types",
        parameters: {
            type: "object",
            properties: {
                dataRange: {
                    type: "string",
                    description: "Excel range address (e.g., 'A1:C10'), 'selection' for current selection, or 'result from step X' to reference previous step"
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
        description: "Creates a complete dashboard with multiple charts, analysis summary, and professional formatting. Automatically analyzes data and determines optimal charts",
        parameters: {
            type: "object",
            properties: {},
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
    forecast: {
        name: "forecast",
        description: "Predicts future values based on historical data trends using AI analysis",
        parameters: {
            type: "object",
            properties: {
                periods: {
                    type: "number",
                    description: "Number of future periods to forecast",
                    default: 12
                }
            },
            required: []
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
