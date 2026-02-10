// financial-tools.js - Specialized Financial Analysis Tools for Valuation & Modeling

/**
 * Financial Analysis Tool Definitions
 * These tools enable investment banking-grade financial analysis including:
 * - Comparable Company Analysis
 * - Precedent Transaction Analysis
 * - DCF Valuation
 * - LBO Modeling
 * - Three-Statement Financial Models
 */

const FINANCIAL_TOOLS = {
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
 * Export for integration with main tool registry
 */
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { FINANCIAL_TOOLS };
}

if (typeof window !== 'undefined') {
    window.FINANCIAL_TOOLS = FINANCIAL_TOOLS;
}
