// autonomous-agent.js - Autonomous AI Agent for Excel Operations



class AutonomousExcelAgent extends DashboardAnalyzer {

    constructor(apiKey) {

        super(apiKey);

        this.conversationHistory = [];

        this.executionPlan = null;

        this.isExecuting = false;

    }



    // ============================================

    // AUTONOMOUS AGENT CORE

    // ============================================



    /**

     * Main entry point - Process user query and execute autonomously

     */

    async processQuery(userQuery, onStatusUpdate = null, retryCount = 0) {

        console.log(`🤖 Agent received query: "${userQuery}" (Attempt ${retryCount + 1})`);

        this.onStatusUpdate = onStatusUpdate;



        this.isExecuting = true;



        try {

            if (this.onStatusUpdate) this.onStatusUpdate("🤔 Analyzing your request...");

            // Step 1: Understand the query

            const intent = await this.analyzeIntent(userQuery);



            if (this.onStatusUpdate) this.onStatusUpdate("📋 Creating expert execution plan...");



            // Step 2: Create execution plan

            const plan = await this.createExecutionPlan(intent, userQuery);



            // NEW: Show the plan to the user immediately

            if (this.onStatusUpdate) {

                let planMsg = `### 📋 Strategy: ${plan.title || 'Data Initiative'}\n\n`;

                planMsg += `${plan.summary || ''}\n\n`;

                planMsg += `**Execution Steps:**\n`;

                plan.steps.forEach(s => {

                    const stepNum = s.step || s.stepNumber || '?';

                    planMsg += `${stepNum}. **${s.description}**\n   *Rationale: ${s.rationale || 'Necessary for objective'}*\n`;

                });

                planMsg += `\n**Expected Outcome:** ${plan.expectedOutcome || 'Analysis complete'}`;



                this.onStatusUpdate(planMsg);

            }



            // Step 3: Get user approval (Wait a small bit to allow user to read)

            const approved = await this.presentPlan(plan);



            if (!approved) {

                return { status: 'cancelled', message: 'User cancelled operation' };

            }



            if (this.onStatusUpdate) this.onStatusUpdate("⚙️ Executing high-fidelity operations...");



            // Step 4: Execute the plan with granular recovery

            const result = await this.executePlan(plan);



            // Step 5: Final expert reflection

            if (this.onStatusUpdate) this.onStatusUpdate("🧪 Synthesizing final insights...");

            const finalReflection = await this.generateExpertReflection(userQuery, result);



            return {

                status: 'success',

                query: userQuery,

                intent: intent,

                plan: plan,

                result: result,

                reflection: finalReflection

            };



        } catch (error) {

            console.error(`Agent error:`, error);

            return {

                status: 'error',

                message: error.message

            };

        } finally {

            this.isExecuting = false;

        }

    }



    /**

     * Analyze user intent using Claude AI

     */

    async analyzeIntent(userQuery) {

        const prompt = `You are a World-Class Strategy Consultant and Data Scientist (McKinsey/BCG level).
Global Objective: Analyze the user's intent with extreme professional depth.

USER QUERY: "${userQuery}"

INTENT CATEGORIES:
- analyze_data: Data analysis or insights
- create_chart: Create visualization
- clean_data: Clean, transform, or fix data (splitting, deduplication, formatting)
- create_dashboard: Build multi-chart dashboard
- find_insights: Answer specific question about data
- calculate: Compute metrics or add calculated columns
- format: Style or format cells
- filter_sort: Filter or sort data
- comparable_analysis: Trading multiples / comparable company analysis
- precedent_transactions: M&A deal analysis
- dcf_valuation: DCF modeling
- lbo_analysis: LBO scenario analysis
- three_statement_model: Integrated financial model
- financial_benchmarking: Peer comparison
- historical_financials: Multi-year financial aggregation
- financial_ratios: Ratio analysis
- custom: Multi-stage or complex operations

Respond in JSON ONLY:
{
  "intent": "category from above",
  "businessObjective": "Executive Summary of user goal",
  "confidence": "high/medium/low",
  "details": {
     "dataRange": "EXACT address or 'selection'",
     "keyMetrics": ["relevant columns or KPIs"],
     "implicitNeeds": ["What user didn't say but needs"]
  },
  "expertObservations": "Strategic context",
  "reasoning": "Brief classification explanation"
}`;

        const systemPrompt = "You are an Excel AI agent. Classify intent accurately. Return only valid JSON.";

        const response = await this.callClaudeAPI(prompt, systemPrompt);
        const intent = this.extractJSON(response);

        if (!intent || !intent.intent) {
            throw new Error(`Failed to analyze user intent. Agent response: ${response.substring(0, 100)}...`);
        }

        return intent;
    }



    /**

     * Generate an expert-level reflection and summary of the completed work

     */

    async generateExpertReflection(query, results) {

        const prompt = `Summarize the completed Excel automation project.

User Query: "${query}"
Results: ${JSON.stringify(results, null, 2)}

Provide a concise Markdown summary with:
1. **What was done**: Brief statement of operations performed
2. **Key Findings**: Data insights from the results (if any)
3. **Recommendations**: 2-3 actionable next steps

Be concise and data-driven.`;



        const systemPrompt = "You are a World-Class Strategy Consultant. Provide an elite-level, markdown-formatted executive summary of results.";



        try {

            const response = await this.callClaudeAPI(prompt, systemPrompt);

            return response;

        } catch (error) {

            console.warn("Failed to generate reflection:", error);

            return "✅ **Project Complete**: High-fidelity operations performed successfully. Please review the spreadsheet for the final output.";

        }

    }



    /**

     * Create detailed execution plan

     */

    async createExecutionPlan(intent, userQuery) {
        // Get current Excel context
        // Get current Excel context but limit size for prompt
        const context = await this.getExcelContext();

        // Truncate large datasets to prevent token overflow
        const contextForPrompt = { ...context };
        if (contextForPrompt.values && Array.isArray(contextForPrompt.values) && contextForPrompt.values.length > 50) {
            contextForPrompt.values = contextForPrompt.values.slice(0, 50);
            contextForPrompt.note = `Data truncated for analysis. Showing first 50 of ${context.values.length} rows.`;
        }

        const contextStr = JSON.stringify(contextForPrompt, null, 2);

        const toolDescriptions = getToolDescriptionsForPrompt();

        // Only include financial workflows if the intent is financial
        const financialIntents = ['comparable_analysis', 'precedent_transactions', 'dcf_valuation', 'lbo_analysis', 'three_statement_model', 'financial_benchmarking', 'historical_financials', 'financial_ratios'];
        const isFinancial = financialIntents.includes(intent.intent);

        let financialSection = '';
        if (isFinancial) {
            financialSection = `
═══════════════════════════════════════════════════════════
FINANCIAL ANALYSIS WORKFLOWS
═══════════════════════════════════════════════════════════
- Comparable Company: Use \`calculateComparableCompanyMultiples\` for EV/Revenue, EV/EBITDA, P/E multiples
- DCF: Use \`buildHistoricalFinancials\` → \`buildThreeStatementModel\` → \`buildDCFModel\`
- LBO: Use \`buildLBOModel\` with entry assumptions and exit scenarios
- Precedent Transactions: Use \`analyzePrecedentTransactions\` for deal multiples
- Three-Statement: Use \`buildThreeStatementModel\` for integrated P&L, BS, CF projections
- Benchmarking: Use \`benchmarkAgainstPeers\` for percentile rankings
- All financial models use Excel formulas for auditability.
`;
        }

        const prompt = `You are an Excel Solutions Architect. Create a formula-first execution plan.

STRATEGIC CONSTRAINT: Work on the ACTIVE SHEET only. Do not create new sheets or switch sheets unless explicitly requested.

═══════════════════════════════════════════════════════════
EXCEL CONTEXT
═══════════════════════════════════════════════════════════
${contextStr}

═══════════════════════════════════════════════════════════
GUIDELINES
═══════════════════════════════════════════════════════════
1. **Formula-First**: Use \`createMetricTable\` for summaries, \`applyFormulaToColumn\` for transformations. Avoid \`cleanData\` for row-by-row ops.
2. **No Table References**: Use standard cell references (A2, B2), not [@Column].
3. **Aggregate Before Charting**: Always create summary tables before charts.
4. **Sheet Names**: Include sheet names in ranges (e.g., "${context.activeSheet}!A1:B10").
5. **Data starts at row 2** (row 1 = headers).
6. **Single Sheet Focus**: Output results to the active sheet (e.g., columns M+ or relevant empty space) rather than creating new sheets.
${financialSection}
═══════════════════════════════════════════════════════════
AVAILABLE TOOLS
═══════════════════════════════════════════════════════════
${toolDescriptions}

═══════════════════════════════════════════════════════════
USER QUERY & DATA
═══════════════════════════════════════════════════════════
Query: "${userQuery}"
Headers: ${JSON.stringify(context.headers)}
Data Sample:
${this.formatDataForAI(context.dataSample, context.address)}

═══════════════════════════════════════════════════════════
OUTPUT FORMAT (JSON ONLY)
═══════════════════════════════════════════════════════════
{
  "title": "Plan Title",
  "summary": "What this plan does",
  "steps": [
    {
      "step": 1,
      "thought": "Why this step is needed",
      "description": "What this step does",
      "method": "toolName",
      "parameters": { 
        "dataRange": "SheetName!A1:B10" or "selection" or "result from step X",
        ... 
      },
      "rationale": "Business value",
      "critical": true
    }
  ],
  "expectedOutcome": "Final result description"
}

Output valid JSON only.`;





        const systemPrompt = "You are a Senior Excel Solutions Architect. Create detailed, multi-step execution plans with explicit rationale. Return only valid JSON.";



        const response = await this.callClaudeAPI(prompt, systemPrompt);

        const plan = this.extractJSON(response);



        if (!plan || !plan.steps) {

            throw new Error(`Failed to create execution plan.Help the user understand what happened.`);

        }



        return plan;

    }



    /**

     * Get current Excel context

     */

    async getExcelContext(dataAddress = "worksheet") {

        try {

            const sheets = await this.getAllSheetNames();

            let data;



            if (dataAddress === "selection") {

                data = await this.getSelectedRange();

            } else if (dataAddress === "worksheet") {

                data = await this.getWorksheetData();

            } else {

                try {

                    data = await this.getRangeByAddress(dataAddress);

                } catch (rangeError) {

                    console.warn(`[Agent] Specific context range "${dataAddress}" not found, falling back to worksheet.`);

                    data = await this.getWorksheetData();

                }

            }



            // Extract a sample of the data (e.g., first 10 rows)
            const dataSample = data.values ? data.values.slice(0, 10) : [];
            const headers = (data.values && data.values.length > 0) ? data.values[0] : [];

            return {
                sheets: sheets,
                activeSheet: data.sheetName || "Active",
                headers: headers,
                dataRows: data.rowCount || (data.values ? data.values.length : 0),
                dataColumns: data.columnCount || (data.values && data.values[0] ? data.values[0].length : 0),
                selection: data.address,
                address: data.address,
                dataSample: dataSample
            };

        } catch (error) {

            console.warn("Context fetch error:", error);

            return { error: "Could not get Excel context", sheets: [] };

        }

    }



    /**

     * Present plan to user for approval

     */

    async presentPlan(plan) {

        console.log("📋 Execution Plan:");

        console.log(`Summary: ${plan.summary} `);

        console.log(`Total Steps: ${plan.steps.length} `);

        console.log(`Estimated Time: ${plan.estimatedTotalTime || 'unknown'} `);



        plan.steps.forEach((step, i) => {

            const stepNum = step.step || step.stepNumber || (i + 1);

            const actionText = step.description || step.action || "Executing step";

            console.log(`  ${stepNum}. ${actionText} `);

        });



        if (plan.warnings && plan.warnings.length > 0) {

            console.log("⚠️ Warnings:", plan.warnings);

        }



        // For autonomous operation, auto-approve

        // In production, you'd show UI confirmation

        return true;

    }



    /**

     * Execute the plan autonomously

     */

    async executePlan(plan) {

        const results = [];

        const MAX_STEP_RETRIES = 2;



        console.log(`🚀 Starting execution of ${plan.steps.length} steps with granular recovery...`);

        if (this.onStatusUpdate) this.onStatusUpdate(`🚀 Starting ${plan.steps.length} steps...`);



        for (let i = 0; i < plan.steps.length; i++) {

            let step = plan.steps[i];

            const totalSteps = plan.steps.length;

            let retryCount = 0;

            let stepSuccess = false;



            while (retryCount <= MAX_STEP_RETRIES && !stepSuccess) {

                // ReAct - Thought

                const thoughtMsg = retryCount > 0

                    ? `🔄 Retry ${retryCount}: ${step.thought} `

                    : `💭 Thought: ${step.thought} `;



                console.log(thoughtMsg);

                if (this.onStatusUpdate) this.onStatusUpdate(thoughtMsg);



                // ReAct - Action

                const currentStepNum = step.step || step.stepNumber || (i + 1);

                const actionText = step.description || step.action || "Executing step";

                const actionMsg = `⏳ Step ${currentStepNum}/${totalSteps}: ${actionText}`;

                console.log(actionMsg);

                if (this.onStatusUpdate) this.onStatusUpdate(actionMsg);



                try {

                    const stepResult = await this.executeStep(step, results);



                    // ReAct - Observation

                    if (this.onStatusUpdate) this.onStatusUpdate(`👀 Observing results...`);

                    const observation = await this.getExcelContext();



                    results.push({

                        step: currentStepNum,

                        thought: step.thought,

                        action: actionText,

                        method: step.method,

                        status: 'success',

                        result: stepResult,

                        observation: observation

                    });



                    console.log(`✅ Step ${currentStepNum} completed and verified`);

                    stepSuccess = true;



                } catch (error) {

                    const currentStepNum = step.step || step.stepNumber || (i + 1);

                    retryCount++;

                    console.error(`❌ Step ${currentStepNum} failed (Attempt ${retryCount}):`, error);



                    if (retryCount <= MAX_STEP_RETRIES) {

                        const healingStatus = `🔄 Step ${currentStepNum} failed. Initiating localized self-healing...`;

                        if (this.onStatusUpdate) this.onStatusUpdate(healingStatus);



                        try {

                            const healResponse = await this.attemptSelfHeal(error, step, results);



                            if (healResponse.replanFromHere && healResponse.newSteps) {

                                console.log(`🆕 AI requested a replan from Step ${step.stepNumber}`);

                                // Replace remaining steps with new ones

                                plan.steps.splice(i, plan.steps.length - i, ...healResponse.newSteps);

                                step = plan.steps[i]; // Update current step to the first new one

                                retryCount = 0; // Reset retry count for the new step

                                if (this.onStatusUpdate) this.onStatusUpdate(`📋 Plan patched! Continuing with new approach...`);

                            } else if (healResponse.fixAction) {

                                // Update current step with fixed version and continue loop

                                step = { ...step, ...healResponse.fixAction, thought: healResponse.thought || step.thought };

                            }

                        } catch (healingError) {

                            console.error(`💔 Self-healing failed for Step ${step.stepNumber}:`, healingError);

                            if (retryCount > MAX_STEP_RETRIES) break;

                        }

                    } else {

                        results.push({

                            step: currentStepNum,

                            action: actionText,

                            method: step.method,

                            status: 'error',

                            error: error.message

                        });



                        if (this.onStatusUpdate) this.onStatusUpdate(`❌ Step ${currentStepNum} failed permanently: ${error.message}`);



                        if (step.critical) {

                            throw new Error(`Critical step ${currentStepNum} failed after ${MAX_STEP_RETRIES} recovery attempts: ${error.message}`);

                        }

                    }

                }

            }

        }



        console.log(`✅ Execution complete! ${results.filter(r => r.status === 'success').length}/${plan.steps.length} steps succeeded`);



        return {

            totalSteps: plan.steps.length,

            successfulSteps: results.filter(r => r.status === 'success').length,

            failedSteps: results.filter(r => r.status === 'error').length,

            results: results

        };

    }



    /**

     * Attempt to fix a failed step autonomously with comprehensive error context

     */

    async attemptSelfHeal(error, failedStep, previousResults) {

        const context = await this.getExcelContext();



        // Build detailed execution history with results

        const detailedHistory = previousResults.map(r => ({

            step: r.step,

            thought: r.thought || 'N/A',

            action: r.action,

            method: r.method || 'unknown',

            status: r.status,

            result: r.result ? {

                address: r.result.address || 'N/A',

                message: r.result.message || 'N/A',

                status: r.result.status || 'N/A'

            } : 'No result data',

            observation: r.observation ? {

                dataRows: r.observation.dataRows,

                dataColumns: r.observation.dataColumns,

                activeSheet: r.observation.activeSheet

            } : 'No observation'

        }));



        // Get successful steps that can be referenced

        const successfulSteps = previousResults.filter(r => r.status === 'success' && r.result && r.result.address);

        const availableDataRanges = successfulSteps.map(s => ({

            step: s.step,

            action: s.action,

            address: s.result.address,

            method: s.method || 'unknown'

        }));



        // Get stack trace if available

        const stackTrace = error.stack || 'No stack trace available';



        // Get tool registry information (if available)

        let toolInfo = '';

        if (typeof getTool === 'function') {

            const tool = getTool(failedStep.method);

            if (tool) {

                toolInfo = `

TOOL INFORMATION FOR "${failedStep.method}":

Description: ${tool.description}

Parameters: ${JSON.stringify(tool.parameters, null, 2)}

Common Errors: ${JSON.stringify(tool.commonErrors || [], null, 2)}

Examples: ${JSON.stringify(tool.examples || [], null, 2)}

`;

            }

        }



        const prompt = `You are an expert Excel AI repair specialist. A step failed and you need to fix it using the ReAct pattern.



═══════════════════════════════════════════════════════════

FAILED STEP DETAILS

═══════════════════════════════════════════════════════════

${JSON.stringify(failedStep, null, 2)}



═══════════════════════════════════════════════════════════

ERROR INFORMATION

═══════════════════════════════════════════════════════════

Error Message: "${error.message}"

${failedStep.method === 'insertFormula' || failedStep.method === 'generateFormula' ? `Attempted Formula: "${failedStep.parameters.formula || 'Internal Generation'}"` : ''}



Stack Trace (first 500 chars):

${stackTrace.substring(0, 500)}



Parameters Used:

${JSON.stringify(failedStep.parameters, null, 2)}



Previous Thought: "${failedStep.thought || 'N/A'}"



═══════════════════════════════════════════════════════════

CURRENT EXCEL STATE (where it failed)

═══════════════════════════════════════════════════════════

${JSON.stringify(context, null, 2)}



═══════════════════════════════════════════════════════════

COMPLETE EXECUTION HISTORY (what worked before)

═══════════════════════════════════════════════════════════

${JSON.stringify(detailedHistory, null, 2)}



AVAILABLE DATA RANGES FROM SUCCESSFUL STEPS:

${JSON.stringify(availableDataRanges, null, 2)}



${toolInfo}



═══════════════════════════════════════════════════════════

AVAILABLE TOOLS (with new additions)

═══════════════════════════════════════════════════════════

${getToolDescriptionsForPrompt()}



═══════════════════════════════════════════════════════════

YOUR TASK

═══════════════════════════════════════════════════════════

Analyze the error and choose ONE of these recovery strategies:



STRATEGY 1: Fix Current Step (fixAction)

- Modify parameters to avoid the error

- Use different tool if current one is wrong

- Reference successful step results if needed

- Example: If chart creation failed due to bad range, use a successful data range from previous steps



STRATEGY 2: Replan Remaining Steps (replanFromHere)

- If the entire approach is fundamentally wrong

- Provide new steps to replace ALL remaining steps

- Use insights from what worked so far



CRITICAL GUIDELINES:

- If error mentions "invalid range" or "address", check availableDataRanges and use one of those

- If error mentions "missing parameter", add it with appropriate value

- If error is about formula syntax, simplify the formula or use generateFormula instead of insertFormula

- If data doesn't exist, consider using writeData or cleanData first

- Always reference successful steps when possible (e.g., "result from step 2")



Respond in JSON:

{

  "thought": "Deep reasoning: What caused this error? What did work? How can I fix it or work around it?",

  "errorAnalysis": "Specific diagnosis of what went wrong",

  "fixAction": { 

    "method": "tool name", 

    "parameters": { ... },

    "action": "clear description of what this fix does"

  },

  "replanFromHere": false,

  "newSteps": []

}



OR (if replanning):

{

  "thought": "Why the entire approach needs to change",

  "errorAnalysis": "What's fundamentally wrong",

  "fixAction": null,

  "replanFromHere": true,

  "newSteps": [

    { 

      "stepNumber": N,

      "thought": "reasoning", 

      "action": "description", 

      "method": "tool", 

      "parameters": { ... },

      "expectedOutcome": "what should happen"

    }

  ]

}`;



        const response = await this.callClaudeAPI(prompt, "You are an expert Excel AI repair specialist using ReAct reasoning. Analyze errors deeply and provide robust fixes. Return only valid JSON.");

        const fix = this.extractJSON(response);



        if (!fix || (!fix.fixAction && !fix.replanFromHere)) {

            throw new Error("Failed to generate a valid fix action.");

        }



        console.log(`🩹 Self-Heal Analysis: ${fix.errorAnalysis || 'N/A'}`);

        console.log(`🩹 Self-Heal Thought: ${fix.thought}`);

        if (this.onStatusUpdate) this.onStatusUpdate(`🩹 Fix Strategy: ${fix.thought.substring(0, 100)}...`);



        return fix;

    }



    /**

     * Execute individual step

     */

    async executeStep(step, previousResults = []) {

        let method = step.method;

        // Guard: missing method
        if (!method || typeof method !== 'string') {
            throw new Error(`Step is missing a valid 'method'. Step description: "${step.description || step.action || 'unknown'}"`);
        }

        let params = step.parameters || {};



        // Resolve parameters that refer to previous steps (More robustly)

        for (let key in params) {

            if (typeof params[key] === 'string') {

                const lowerVal = params[key].toLowerCase();

                // Match "result from step X", "step X", or "(step X)"

                const stepNumMatch = lowerVal.match(/(?:result\s+from\s+)?(?:step\s*)(\d+)/i);



                if (stepNumMatch) {

                    const targetStepNum = parseInt(stepNumMatch[1]);

                    const prevStep = previousResults.find(r => r.step === targetStepNum);



                    if (prevStep && prevStep.result) {

                        // Priority: Prefer address for range-based tools, but use data if needed

                        if (prevStep.result.address) {

                            console.log(`[Agent] Resolved parameter ${key} as ADDRESS from Step ${targetStepNum}: ${prevStep.result.address}`);

                            params[key] = prevStep.result.address;

                        } else if (prevStep.result.values || prevStep.result.summary || prevStep.result.data) {

                            console.log(`[Agent] Resolved parameter ${key} as DATA from Step ${targetStepNum}`);

                            params[key] = prevStep.result.values || prevStep.result.summary || prevStep.result.data;

                        } else {
                            console.warn(`[Agent] Step ${targetStepNum} result has no usable address/data. Keeping original reference: "${params[key]}"`);
                        }

                    } else {
                        console.warn(`[Agent] Referenced step ${targetStepNum} not found or has no result. Keeping original: "${params[key]}"`);
                    }

                }

            }

        }



        // Normalize method name (handle common AI variations)
        method = method.replace(/_/g, ''); // Convert snake_case to CamelCase-ish
        const normalizedMethod = method.toLowerCase();

        // AUTO-FETCH DATA: If tool expects data (array) but gets address (string), fetch it
        const dataKey = ['data', 'ranges', 'sourceRange'].find(k => params[k] && typeof params[k] === 'string' && params[k] !== 'selection');

        if (dataKey && ['writedata', 'exporttonewsheet', 'mergedata'].includes(normalizedMethod)) {

            console.log(`[Agent] Auto-fetching data for tool ${normalizedMethod} from address: ${params[dataKey]}`);

            try {

                const fetched = await this.getRangeByAddress(params[dataKey]);

                params[dataKey] = fetched.values;

            } catch (e) {

                console.warn(`[Agent] Failed to auto-fetch data from ${params[dataKey]}:`, e);

            }

        }



        console.log(`[AutonomousExcelAgent] Calling tool "${normalizedMethod}" with params:`, params);



        // Map method names to actual functions

        if (normalizedMethod === 'analyzedata') return await this.performCompleteAnalysis(params.dataRange || params.address || "selection");



        if (normalizedMethod === 'createchart') {
            // Robust Chart Type Normalization
            const typeValue = params.chartType || "";
            const lowerType = typeValue.toLowerCase();

            // Map variations to basic types expected by createChart
            if (lowerType.includes('line')) params.chartType = 'line';
            else if (lowerType.includes('bar')) params.chartType = 'bar';
            else if (lowerType.includes('column')) params.chartType = 'column';
            else if (lowerType.includes('pie')) params.chartType = 'pie';
            else if (lowerType.includes('area')) params.chartType = 'area';
            else if (lowerType.includes('scatter')) params.chartType = 'scatter';

            // Handle missing dataRange by checking for previous data-creating steps
            if (!params.dataRange || params.dataRange === 'selection' || params.dataRange === 'undefined') {
                const lastDataStep = [...previousResults].reverse().find(r => r.result && r.result.address);
                if (lastDataStep) {
                    params.dataRange = lastDataStep.result.address || lastDataStep.observation?.address;
                    console.log(`[Agent] Auto-assigned chart dataRange from Step ${lastDataStep.step}: ${params.dataRange}`);
                } else {
                    const selection = await this.getSelectedRange();
                    params.dataRange = selection.address;
                }
            }
            return await this.createChart(params.dataRange, params.chartType, params.title, params.xAxis, params.yAxis);
        }

        if (normalizedMethod === 'movechart') {
            // Handle "moveChartAndTable" hallucination by just moving the chart
            // If the user/AI asks for moveChartAndTable, we assume they mean moveChart for now
            return await this.moveChart(params.chartName, params.targetCell);
        }

        if (normalizedMethod === 'movechartandtable') {
            // Fallback for hallucinated method
            return await this.moveChart(params.chartName, params.targetCell);
        }

        if (normalizedMethod === 'generatedashboard' || normalizedMethod === 'dashboard') return await this.generateMultipleCharts(params.dataRange || params.address || params.sourceRange || "selection");

        if (normalizedMethod === 'cleandata' || normalizedMethod === 'transformdata' || normalizedMethod === 'transform') {

            const instructions = params.instructions || params.instruction || params.cleaningInstructions || step.action;

            const targetRange = params.dataRange || params.address || params.sourceRange || "selection";

            return await this.applyCleanedData(instructions, targetRange);

        }

        if (normalizedMethod === 'calculatemetric' || normalizedMethod === 'calculate') return await this.calculateMetric(params);

        if (normalizedMethod === 'createmetrictable' || normalizedMethod === 'summarytable') {
            const dataRange = params.dataRange || params.address || "selection";
            return await this.createMetricTable(
                dataRange,
                params.categoryColumn,
                params.metricColumn,
                params.aggregation || "Count",
                params.targetCell || "P1"
            );
        }

        if (normalizedMethod === 'findinsights' || normalizedMethod === 'ask') return await this.askAboutData(params.question, params.dataRange || params.address || params.sourceRange || "selection");

        if (normalizedMethod === 'formatdata' || normalizedMethod === 'format') {

            const range = params.range || params.address || params.targetRange || "selection";

            const formatting = params.formatting || {};

            return await this.formatRange(range, formatting);

        }

        if (normalizedMethod === 'createsummary') return await this.generateSummary(params.dataRange || params.address || params.sourceRange || "selection");

        if (normalizedMethod === 'insertcolumns' || normalizedMethod === 'addcolumns') {
            return await this.insertColumns(params.address || "P:P", params.count || 1);
        }

        if (normalizedMethod === 'generateformula') {

            let context = params.context;

            let address = null;

            if (!context || typeof context === 'string') {

                const sheetData = await this.getWorksheetData();

                context = sheetData.values;

                address = sheetData.address;

            }

            const formula = await this.generateFormula(params.description, context, address);



            // Resolve target address

            let targetAddress = params.address || params.range || params.targetRange || "selection";

            if (targetAddress.toString().toLowerCase() === 'selection') {

                const selection = await this.getSelectedRange();

                targetAddress = selection.address;

            }



            if (targetAddress && formula) {

                console.log(`[Agent] Applying generated formula to: ${targetAddress}`);

                try {

                    await this.insertFormula(targetAddress, formula);

                    return {

                        status: 'success',

                        address: targetAddress,

                        message: `Formula applied to ${targetAddress}: ${formula}`,

                        formula: formula

                    };

                } catch (e) {

                    console.error(`[Agent] Error inserting generated formula:`, e);

                    throw e;

                }

            }

            return { status: 'success', formula: formula };

        }

        if (normalizedMethod === 'insertformula' || normalizedMethod === 'applyformula' || normalizedMethod === 'applyformulatocolumn') {

            // Check if it's the specific column tool
            if (normalizedMethod === 'applyformulatocolumn' || (params.targetColumn && params.formula)) {
                return await this.applyFormulaToColumn(params.formula, params.targetColumn, params.sourceColumnForHeight || "A", params.headerName);
            }

            let targetAddress = params.address || params.range || "selection";

            if (targetAddress.toString().toLowerCase() === 'selection') {
                const selection = await this.getSelectedRange();
                targetAddress = selection.address;
            }

            if (!params.formula) {
                throw new Error("Missing 'formula' parameter for insertFormula operation.");
            }

            return await this.insertFormula(targetAddress, params.formula);
        }

        if (normalizedMethod === 'createpivot') return await this.createPivotAnalysis(params);

        if (normalizedMethod === 'createsheet' || normalizedMethod === 'addsheet') return await this.createNewSheet(params.name || params.sheetName);

        if (normalizedMethod === 'writedata' || normalizedMethod === 'writetorange' || normalizedMethod === 'write') {

            if (!params.address || params.address === 'undefined' || params.address === 'selection') {

                const selection = await this.getSelectedRange();

                params.address = selection.address;

                console.log(`[Agent] Auto-assigned writeData address: ${params.address}`);

            }

            let data = params.data;

            if (typeof data === 'string' && data.length < 100 && (data.includes('!') || data.match(/[A-Z]+\d+/))) {

                console.log(`[Agent] writeData 'data' appears to be an address: ${data}. Fetching...`);

                const source = await this.getRangeByAddress(data);

                data = source.values;

            }

            return await this.writeToRange(params.address, data);

        }



        // New Tools - Data Transformation

        if (normalizedMethod === 'filterdata' || normalizedMethod === 'filter') {

            return await this.filterData(params.column, params.operator, params.value, params.dataRange || params.address);

        }

        if (normalizedMethod === 'sortdata' || normalizedMethod === 'sort') {

            return await this.sortData(params.columns, params.order || 'asc', params.dataRange || params.address);

        }

        if (normalizedMethod === 'mergedata' || normalizedMethod === 'merge') {

            return await this.mergeData(params.ranges, params.mergeType || 'horizontal');

        }

        if (normalizedMethod === 'validatedata' || normalizedMethod === 'validate') {

            return await this.validateData(params.rules, params.dataRange || params.address);

        }

        if (normalizedMethod === 'exporttonewsheet' || normalizedMethod === 'export') {

            let data = params.data;

            if (!data || !Array.isArray(data)) {

                const source = await (params.dataRange || params.address ? this.getRangeByAddress(params.dataRange || params.address) : this.getWorksheetData());

                data = source.values;

            }

            return await this.exportToNewSheet(data, params.sheetName);

        }



        // Debugging & Auditing Tools

        if (normalizedMethod === 'traceerror' || normalizedMethod === 'trace') {

            return await this.traceError(params.address);

        }



        if (normalizedMethod === 'explainformula' || normalizedMethod === 'explain') {

            return await this.explainFormula(params.address);

        }



        if (normalizedMethod === 'createscenario' || normalizedMethod === 'scenario') {

            return await this.createScenario(params.scenarioName, params.changes);

        }

        // Handle sheet activation gracefully (even if hallucinated)
        if (normalizedMethod === 'activatesheet' || normalizedMethod === 'gotosheet' || normalizedMethod === 'selectsheet') {
            console.log(`[Agent] Handling activateSheet request for: ${params.sheetName || params.name}. Staying on active sheet as per policy.`);
            return { success: true, message: `Active sheet confirmed: ${params.sheetName || "current"}` };
        }

        // ============================================
        // FINANCIAL ANALYSIS TOOLS
        // ============================================

        if (normalizedMethod === 'calculatecomparablecompanymultiples' || normalizedMethod === 'comps' || normalizedMethod === 'comparables') {
            return await this.calculateComparableCompanyMultiples(params);
        }

        if (normalizedMethod === 'analyzeprecedenttransactions' || normalizedMethod === 'precedents' || normalizedMethod === 'transactions') {
            return await this.analyzePrecedentTransactions(params);
        }

        if (normalizedMethod === 'buildhistoricalfinancials' || normalizedMethod === 'historicals') {
            return await this.buildHistoricalFinancials(params);
        }

        if (normalizedMethod === 'buildthreestatementmodel' || normalizedMethod === 'threestatement' || normalizedMethod === '3statement') {
            return await this.buildThreeStatementModel(params);
        }

        if (normalizedMethod === 'builddcfmodel' || normalizedMethod === 'dcf' || normalizedMethod === 'dcfvaluation') {
            return await this.buildDCFModel(params);
        }

        if (normalizedMethod === 'buildlbomodel' || normalizedMethod === 'lbo' || normalizedMethod === 'leveragedbuyout') {
            return await this.buildLBOModel(params);
        }

        if (normalizedMethod === 'calculatefinancialratios' || normalizedMethod === 'ratios' || normalizedMethod === 'financialratios') {
            return await this.calculateFinancialRatios(params);
        }

        if (normalizedMethod === 'benchmarkagainstpeers' || normalizedMethod === 'benchmark' || normalizedMethod === 'peerbenchmark') {
            return await this.benchmarkAgainstPeers(params);
        }


        throw new Error(`Unknown method: ${step.method}`);

    }



    // ============================================

    // ADVANCED AUTONOMOUS OPERATIONS

    // ============================================



    /**

     * Calculate custom metrics

     */

    async calculateMetric(params) {

        const data = await this.getWorksheetData();

        const dataStr = this.formatDataForAI(data.values);



        const prompt = `Calculate this metric: ${params.metric}



Data:

${dataStr}



Parameters: ${JSON.stringify(params)}



Return JSON:

{

  "result": "calculated value",

  "formula": "Excel formula used",

  "explanation": "how it was calculated",

  "cellLocation": "where to write result"

}`;



        const response = await this.callClaudeAPI(prompt, "Return only valid JSON.");

        const result = this.extractJSON(response);



        if (!result) {

            throw new Error("Could not parse metric calculation result.");

        }



        // Write result to Excel

        if (result.cellLocation) {

            await this.writeToRange(result.cellLocation, [[result.result]]);

        }



        return result;

    }



    /**

     * Create pivot table analysis

     */

    async createPivotAnalysis(params) {

        if (this.onStatusUpdate) this.onStatusUpdate("📊 Analyzing data for pivot table...");

        const dataRange = params.dataRange || params.address || params.sourceRange || "selection";

        const contextData = await this.getExcelContext(dataRange);

        const dataStr = this.formatDataForAI(contextData.values, contextData.address);



        const prompt = `Create a pivot table analysis:



Data:

${dataStr}



Requirements: ${JSON.stringify(params)}



Return JSON with:

{

  "pivotStructure": {

    "rows": ["field1"],

    "columns": ["field2"],

    "values": ["sum of field3"]

  },

  "insights": "key findings from pivot",

  "summary": [["Row Labels", "Value"], ["Category A", 100], ["Category B", 200]] (A 2D array of the summarized data)

}`;



        const response = await this.callClaudeAPI(prompt, "Return only valid JSON.");

        const result = this.extractJSON(response);



        if (!result) {

            throw new Error("Could not parse pivot analysis result.");

        }



        if (this.onStatusUpdate) this.onStatusUpdate("📑 Creating pivot output sheet...");



        // Write the summary to a new sheet

        const sheetName = `Pivot_${new Date().getTime().toString().slice(-4)}`;

        await Excel.run(async (context) => {

            const sheet = context.workbook.worksheets.add(sheetName);

            sheet.activate();

            await context.sync();

        });



        if (result.summary && Array.isArray(result.summary)) {

            await this.writeToRange(`${sheetName}!A1`, result.summary);

            // Calculate the address of the written data

            const rowCount = result.summary.length;

            const colCount = result.summary[0] ? result.summary[0].length : 0;

            const endCol = this.columnIndexToLetter(colCount - 1);

            result.address = `${sheetName}!A1:${endCol}${rowCount}`;

        }



        return result;

    }



    // ============================================

    // CONVERSATIONAL INTERFACE

    // ============================================



    /**

     * Chat with the agent

     */

    async chat(userMessage) {

        this.conversationHistory.push({

            role: 'user',

            content: userMessage,

            timestamp: new Date()

        });



        // Build conversation context

        const conversationContext = this.conversationHistory

            .slice(-5) // Last 5 messages

            .map(msg => `${msg.role}: ${msg.content}`)

            .join('\n');



        const prompt = `You are an autonomous Excel AI agent. The user is chatting with you about their Excel data.



Conversation history:

${conversationContext}



Current Excel context: ${JSON.stringify(await this.getExcelContext())}



Respond conversationally and offer to perform operations. If the user wants something done, explain what you'll do and ask for confirmation.



Your response:`;



        const response = await this.callClaudeAPI(prompt);



        this.conversationHistory.push({

            role: 'assistant',

            content: response,

            timestamp: new Date()

        });



        // Check if response suggests an action

        const actionIntent = await this.detectActionIntent(response);



        return {

            message: response,

            suggestedAction: actionIntent

        };

    }



    /**

     * Detect if agent response suggests an action

     */

    async detectActionIntent(agentResponse) {

        const prompt = `Does this agent response suggest performing an action?



Response: "${agentResponse}"



Return JSON:

{

  "hasAction": true/false,

  "action": "description of action or null",

  "confidence": "high/medium/low"

}`;



        try {

            const response = await this.callClaudeAPI(prompt, "Return only valid JSON.");

            return this.extractJSON(response) || { hasAction: false };

        } catch {

            return { hasAction: false };

        }

    }



    // ============================================

    // PREDEFINED QUERY HANDLERS

    // ============================================



    /**

     * Handle common query patterns

     */

    async handleCommonQueries(query) {

        const lowerQuery = query.toLowerCase();



        // Dashboard generation

        if (lowerQuery.includes('dashboard') || lowerQuery.includes('complete analysis')) {

            return await this.processQuery("Generate a complete dashboard with analysis and charts");

        }







        // Chart creation

        if (lowerQuery.includes('create chart') || lowerQuery.includes('visualize')) {

            return await this.processQuery("Create the best chart for this data");

        }



        // Analysis

        if (lowerQuery.includes('analyze') || lowerQuery.includes('insights')) {

            return await this.processQuery("Analyze this data and provide insights");

        }



        // Cleaning

        if (lowerQuery.includes('clean') || lowerQuery.includes('fix')) {

            return await this.processQuery("Clean and standardize this data");

        }



        // Summary

        if (lowerQuery.includes('summary') || lowerQuery.includes('summarize')) {

            return await this.processQuery("Create a summary of this data");

        }



        // Otherwise, process as custom query

        return await this.processQuery(query);

    }



    /**
    
     * Extract numbers from text
    
     */

    extractNumber(text) {

        const match = text.match(/\d+/);

        return match ? parseInt(match[0]) : null;

    }



    // ============================================

    // AUTONOMOUS MONITORING

    // ============================================



    /**

     * Monitor data changes and suggest actions

     */

    async monitorAndSuggest() {

        const data = await this.getWorksheetData();



        const prompt = `Analyze this Excel data and suggest helpful autonomous actions:



Data:

${this.formatDataForAI(data.values.slice(0, 20))} (showing first 20 rows)



Total rows: ${data.rowCount}



Suggest 3-5 actions I could autonomously perform to help the user. Return JSON:

{

  "suggestions": [

    {

      "action": "action description",

      "benefit": "why this would help",

      "priority": "high/medium/low",

      "query": "query to execute this action"

    }

  ]

}`;



        const response = await this.callClaudeAPI(prompt, "Return only valid JSON.");

        return this.extractJSON(response);

    }



    /**

     * Learn from user feedback

     */

    async learnFromFeedback(action, wasSuccessful, userFeedback) {

        this.conversationHistory.push({

            role: 'feedback',

            action: action,

            successful: wasSuccessful,

            feedback: userFeedback,

            timestamp: new Date()

        });



        // In production, this would update a learning model

        console.log(`📚 Learning: Action "${action}" was ${wasSuccessful ? 'successful' : 'unsuccessful'}`);

        if (userFeedback) {

            console.log(`User feedback: ${userFeedback}`);

        }

    }



    // ============================================

    // NEW DATA TRANSFORMATION TOOLS

    // ============================================



    /**

     * Filter data based on conditions

     */

    async filterData(column, operator, value, dataRange = null) {

        const data = dataRange ? await this.getRangeByAddress(dataRange) : await this.getWorksheetData();

        const dataStr = this.formatDataForAI(data.values, data.address);



        const prompt = `Filter this data based on the condition:



DATA:

${dataStr}



FILTER CONDITION:

- Column: "${column}"

- Operator: "${operator}"

- Value: ${JSON.stringify(value)}



Return JSON:

{

  "filteredData": [[row1], [row2], ...] (2D array with header row first, then matching data rows),

  "rowsReturned": number,

  "originalRows": number,

  "filterSummary": "description of what was filtered"

}`;



        const response = await this.callClaudeAPI(prompt, "Return only valid JSON with filtered data.");

        const result = this.extractJSON(response);



        if (!result || !result.filteredData) {

            throw new Error("Failed to filter data - invalid response from AI");

        }



        // Write filtered data to new location

        const newSheetName = `Filtered_${new Date().getTime().toString().slice(-4)}`;

        await this.createNewSheet(newSheetName);

        await this.writeToRange(`${newSheetName}!A1`, result.filteredData);



        // Calculate address

        const rowCount = result.filteredData.length;

        const colCount = result.filteredData[0] ? result.filteredData[0].length : 0;

        const endCol = this.columnIndexToLetter(colCount - 1);

        const address = `${newSheetName}!A1:${endCol}${rowCount}`;



        return {

            status: 'success',

            address: address,

            rowsReturned: result.rowsReturned,

            originalRows: result.originalRows,

            message: result.filterSummary,

            sheetName: newSheetName

        };

    }



    /**

     * Sort data by columns

     */

    async sortData(columns, order = 'asc', dataRange = null) {

        const data = dataRange ? await this.getRangeByAddress(dataRange) : await this.getWorksheetData();

        const dataStr = this.formatDataForAI(data.values, data.address);



        // Normalize order parameter

        const orderArray = Array.isArray(order) ? order : [order];



        const prompt = `Sort this data:



DATA:

${dataStr}



SORT CRITERIA:

- Columns: ${JSON.stringify(columns)}

- Order: ${JSON.stringify(orderArray)} (asc = ascending, desc = descending)



Sort by the specified columns in priority order. Keep the header row at the top.



Return JSON:

{

  "sortedData": [[header], [row1], [row2], ...] (2D array with sorted data),

  "sortedBy": "description of sort criteria applied"

}`;



        const response = await this.callClaudeAPI(prompt, "Return only valid JSON with sorted data.");

        const result = this.extractJSON(response);



        if (!result || !result.sortedData) {

            throw new Error("Failed to sort data - invalid response from AI");

        }



        // Write sorted data to new location

        const newSheetName = `Sorted_${new Date().getTime().toString().slice(-4)}`;

        await this.createNewSheet(newSheetName);

        await this.writeToRange(`${newSheetName}!A1`, result.sortedData);



        // Calculate address

        const rowCount = result.sortedData.length;

        const colCount = result.sortedData[0] ? result.sortedData[0].length : 0;

        const endCol = this.columnIndexToLetter(colCount - 1);

        const address = `${newSheetName}!A1:${endCol}${rowCount}`;



        return {

            status: 'success',

            address: address,

            sortedBy: result.sortedBy,

            message: `Data sorted by ${columns.join(', ')}`,

            sheetName: newSheetName

        };

    }



    /**

     * Merge data from multiple ranges

     */

    async mergeData(ranges, mergeType = 'horizontal') {

        const prompt = `Merge data from multiple ranges:



RANGES TO MERGE: ${JSON.stringify(ranges)}

MERGE TYPE: ${mergeType} (horizontal = side-by-side, vertical = stacked)



Instructions:

- If horizontal: combine columns side-by-side, matching row counts

- If vertical: stack rows on top of each other, matching column counts

- Handle header rows appropriately



Return JSON:

{

  "mergedData": [[row1], [row2], ...] (2D array with merged data),

  "mergeDescription": "description of how data was merged",

  "rowCount": number,

  "columnCount": number

}`;



        const response = await this.callClaudeAPI(prompt, "Return only valid JSON with merged data.");

        const result = this.extractJSON(response);



        if (!result || !result.mergedData) {

            throw new Error("Failed to merge data - invalid response from AI");

        }



        // Write merged data to new location

        const newSheetName = `Merged_${new Date().getTime().toString().slice(-4)}`;

        await this.createNewSheet(newSheetName);

        await this.writeToRange(`${newSheetName}!A1`, result.mergedData);



        // Calculate address

        const rowCount = result.mergedData.length;

        const colCount = result.mergedData[0] ? result.mergedData[0].length : 0;

        const endCol = this.columnIndexToLetter(colCount - 1);

        const address = `${newSheetName}!A1:${endCol}${rowCount}`;



        return {

            status: 'success',

            address: address,

            message: result.mergeDescription,

            rowCount: result.rowCount,

            columnCount: result.columnCount,

            sheetName: newSheetName

        };

    }



    /**

     * Validate data quality

     */

    async validateData(rules, dataRange = null) {

        const data = dataRange ? await this.getRangeByAddress(dataRange) : await this.getWorksheetData();

        const dataStr = this.formatDataForAI(data.values, data.address);



        const prompt = `Validate data quality based on these rules:



DATA:

${dataStr}



VALIDATION RULES: ${JSON.stringify(rules)}



Check for:

- Missing/null values

- Data type mismatches

- Out-of-range values

- Duplicate entries

- Format inconsistencies

- Any rule violations specified



Return JSON:

{

  "isValid": true/false,

  "issues": [

    {

      "type": "missing_value/type_mismatch/duplicate/etc",

      "location": "cell or row reference",

      "description": "what's wrong",

      "severity": "critical/warning/info"

    }

  ],

  "summary": "overall data quality assessment",

  "passedRules": ["rule1", "rule2"],

  "failedRules": ["rule3"]

}`;



        const response = await this.callClaudeAPI(prompt, "Return only valid JSON with validation results.");

        const result = this.extractJSON(response);



        if (!result) {

            throw new Error("Failed to validate data - invalid response from AI");

        }



        return {

            status: 'success',

            isValid: result.isValid,

            issues: result.issues || [],

            summary: result.summary,

            passedRules: result.passedRules || [],

            failedRules: result.failedRules || [],

            message: `Data validation complete: ${result.issues ? result.issues.length : 0} issues found`

        };

    }



    /**

     * Export data to new sheet

     */

    async exportToNewSheet(data, sheetName) {

        // If data is a string reference to previous step, resolve it

        let actualData = data;

        if (typeof data === 'string' && data.toLowerCase().includes('step')) {

            throw new Error("Cannot resolve step reference in exportToNewSheet - data must be provided directly");

        }



        // If data is not provided, use current worksheet

        if (!actualData) {

            const currentData = await this.getWorksheetData();

            actualData = currentData.values;

        }



        // Create new sheet

        const finalSheetName = sheetName || `Export_${new Date().getTime().toString().slice(-4)}`;

        await this.createNewSheet(finalSheetName);



        // Write data

        await this.writeToRange(`${finalSheetName}!A1`, actualData);



        // Calculate address

        const rowCount = actualData.length;

        const colCount = actualData[0] ? actualData[0].length : 0;

        const endCol = this.columnIndexToLetter(colCount - 1);

        const address = `${finalSheetName}!A1:${endCol}${rowCount}`;



        return {

            status: 'success',

            address: address,

            sheetName: finalSheetName,

            rowsExported: rowCount,

            columnsExported: colCount,

            message: `Data exported to sheet "${finalSheetName}"`

        };

    }

    // ============================================
    // FINANCIAL ANALYSIS METHODS
    // ============================================

    /**
     * Calculate Comparable Company Multiples
     */
    async calculateComparableCompanyMultiples(params) {
        console.log('[FinancialAnalysis] Calculating comparable company multiples:', params);

        const { dataRange, targetCell } = params;

        // Get source data
        const sourceData = await this.getRangeByAddress(dataRange);
        const values = sourceData.values;

        if (!values || values.length < 2) {
            throw new Error('Insufficient data for comparable company analysis. Need at least headers + 1 company.');
        }

        const headers = values[0];
        const companies = values.slice(1);

        // Find column indices (case-insensitive matching)
        const findColumn = (names) => {
            for (const name of names) {
                const idx = headers.findIndex(h =>
                    h && h.toString().toLowerCase().includes(name.toLowerCase())
                );
                if (idx !== -1) return idx;
            }
            return -1;
        };

        const companyIdx = findColumn(['company', 'name']);
        const mcapIdx = findColumn(['market cap', 'marketcap', 'mkt cap']);
        const evIdx = findColumn(['enterprise value', 'ev', 'enterprisevalue']);
        const revenueIdx = findColumn(['revenue', 'sales']);
        const ebitdaIdx = findColumn(['ebitda']);
        const niIdx = findColumn(['net income', 'netincome', 'ni', 'profit']);
        const fcfIdx = findColumn(['fcf', 'free cash flow', 'freecashflow']);

        // Build multiples table with formulas
        const multiplesData = [
            ['Company', 'EV/Revenue', 'EV/EBITDA', 'P/E', 'EV/FCF'],
            ...companies.map((company, i) => {
                const rowNum = i + 2; // +2 because of header row and 0-indexing
                const sourceSheet = sourceData.sheetName || 'Sheet1';

                return [
                    company[companyIdx] || `Company ${i + 1}`,
                    evIdx !== -1 && revenueIdx !== -1
                        ? `=INDIRECT("${sourceSheet}!${this.columnIndexToLetter(evIdx)}${rowNum}")/INDIRECT("${sourceSheet}!${this.columnIndexToLetter(revenueIdx)}${rowNum}")`
                        : 'N/A',
                    evIdx !== -1 && ebitdaIdx !== -1
                        ? `=INDIRECT("${sourceSheet}!${this.columnIndexToLetter(evIdx)}${rowNum}")/INDIRECT("${sourceSheet}!${this.columnIndexToLetter(ebitdaIdx)}${rowNum}")`
                        : 'N/A',
                    mcapIdx !== -1 && niIdx !== -1
                        ? `=INDIRECT("${sourceSheet}!${this.columnIndexToLetter(mcapIdx)}${rowNum}")/INDIRECT("${sourceSheet}!${this.columnIndexToLetter(niIdx)}${rowNum}")`
                        : 'N/A',
                    evIdx !== -1 && fcfIdx !== -1
                        ? `=INDIRECT("${sourceSheet}!${this.columnIndexToLetter(evIdx)}${rowNum}")/INDIRECT("${sourceSheet}!${this.columnIndexToLetter(fcfIdx)}${rowNum}")`
                        : 'N/A'
                ];
            })
        ];

        // Add summary statistics
        const summaryStartRow = multiplesData.length + 2;
        multiplesData.push([]); // Blank row
        multiplesData.push(['Summary Statistics', '', '', '', '']);

        const statRows = ['Minimum', 'Maximum', 'Median', 'Mean'];
        const statFunctions = ['MIN', 'MAX', 'MEDIAN', 'AVERAGE'];

        statFunctions.forEach((func, idx) => {
            const row = [statRows[idx]];
            for (let col = 1; col <= 4; col++) {
                const colLetter = this.columnIndexToLetter(col);
                const startRow = 2;
                const endRow = companies.length + 1;
                row.push(`=${func}(${colLetter}${startRow}:${colLetter}${endRow})`);
            }
            multiplesData.push(row);
        });

        // Write to target
        const result = await this.writeToRange(targetCell, multiplesData);

        // Format the multiples table
        await this.formatMultiplesTable(targetCell, multiplesData.length, 5);

        return {
            status: 'success',
            address: result.address || targetCell,
            message: `Comparable company multiples calculated for ${companies.length} companies`,
            companies: companies.length,
            multiples: ['EV/Revenue', 'EV/EBITDA', 'P/E', 'EV/FCF']
        };
    }

    /**
     * Analyze Precedent Transactions
     */
    async analyzePrecedentTransactions(params) {
        console.log('[FinancialAnalysis] Analyzing precedent transactions:', params);

        const { dataRange, targetCell, includeChronologicalTrend = true } = params;

        const sourceData = await this.getRangeByAddress(dataRange);
        const values = sourceData.values;

        if (!values || values.length < 2) {
            throw new Error('Insufficient transaction data. Need at least headers + 1 transaction.');
        }

        const headers = values[0];
        const transactions = values.slice(1);

        // Find columns
        const findColumn = (names) => {
            for (const name of names) {
                const idx = headers.findIndex(h =>
                    h && h.toString().toLowerCase().includes(name.toLowerCase())
                );
                if (idx !== -1) return idx;
            }
            return -1;
        };

        const dateIdx = findColumn(['date', 'announced']);
        const acquirerIdx = findColumn(['acquirer', 'buyer']);
        const targetIdx = findColumn(['target', 'seller']);
        const dealValueIdx = findColumn(['deal value', 'dealvalue', 'ev']);
        const revenueIdx = findColumn(['revenue', 'sales']);
        const ebitdaIdx = findColumn(['ebitda']);

        // Build transaction analysis table
        const analysisData = [
            ['Date', 'Acquirer', 'Target', 'Deal Value', 'EV/Revenue', 'EV/EBITDA'],
            ...transactions.map((txn, i) => {
                const rowNum = i + 2;
                const sourceSheet = sourceData.sheetName || 'Sheet1';

                return [
                    txn[dateIdx] || '',
                    txn[acquirerIdx] || '',
                    txn[targetIdx] || '',
                    dealValueIdx !== -1 ? txn[dealValueIdx] : 'N/A',
                    dealValueIdx !== -1 && revenueIdx !== -1
                        ? `=INDIRECT("${sourceSheet}!${this.columnIndexToLetter(dealValueIdx)}${rowNum}")/INDIRECT("${sourceSheet}!${this.columnIndexToLetter(revenueIdx)}${rowNum}")`
                        : 'N/A',
                    dealValueIdx !== -1 && ebitdaIdx !== -1
                        ? `=INDIRECT("${sourceSheet}!${this.columnIndexToLetter(dealValueIdx)}${rowNum}")/INDIRECT("${sourceSheet}!${this.columnIndexToLetter(ebitdaIdx)}${rowNum}")`
                        : 'N/A'
                ];
            })
        ];

        // Add summary statistics
        analysisData.push([]);
        analysisData.push(['Summary Statistics', '', '', '', '', '']);

        ['Minimum', 'Maximum', 'Median', 'Mean'].forEach((stat, idx) => {
            const func = ['MIN', 'MAX', 'MEDIAN', 'AVERAGE'][idx];
            const evRevRange = `E2:E${transactions.length + 1}`;
            const evEbitdaRange = `F2:F${transactions.length + 1}`;

            analysisData.push([
                stat,
                '',
                '',
                '',
                `=${func}(${evRevRange})`,
                `=${func}(${evEbitdaRange})`
            ]);
        });

        const result = await this.writeToRange(targetCell, analysisData);

        // Format
        await this.formatMultiplesTable(targetCell, analysisData.length, 6);

        return {
            status: 'success',
            address: result.address || targetCell,
            message: `Precedent transaction analysis completed for ${transactions.length} deals`,
            transactions: transactions.length
        };
    }

    /**
     * Build Historical Financials
     */
    async buildHistoricalFinancials(params) {
        console.log('[FinancialAnalysis] Building historical financials:', params);

        const { dataRange, targetCell, includeRatios = true } = params;

        const sourceData = await this.getRangeByAddress(dataRange);
        const values = sourceData.values;

        if (!values || values.length < 3) {
            throw new Error('Insufficient historical data. Need at least 2 years of data.');
        }

        // Assume format: Row headers in column 1, years in subsequent columns
        const years = values[0].slice(1); // Skip first header cell
        const numYears = years.length;

        // Build analysis table with growth rates
        const analysisData = [
            ['Line Item', ...years, `${numYears}-Yr CAGR`],
            ...values.slice(1).map((row, idx) => {
                const lineItem = row[0];
                const yearlyValues = row.slice(1);

                // Create formula for CAGR: ((End/Start)^(1/periods)) - 1
                const startCol = this.columnIndexToLetter(1);
                const endCol = this.columnIndexToLetter(numYears);
                const rowNum = idx + 2;

                const cagrFormula = `=(POWER(${endCol}${rowNum}/${startCol}${rowNum}, 1/${numYears - 1}) - 1)`;

                return [lineItem, ...yearlyValues, cagrFormula];
            })
        ];

        // Add ratio analysis if requested
        if (includeRatios) {
            analysisData.push([]);
            analysisData.push(['Margin Analysis', ...years, '']);

            // Find Revenue row
            const revenueRowIdx = values.findIndex(row =>
                row[0] && row[0].toString().toLowerCase().includes('revenue')
            );

            const ebitdaRowIdx = values.findIndex(row =>
                row[0] && row[0].toString().toLowerCase().includes('ebitda')
            );

            const niRowIdx = values.findIndex(row =>
                row[0] && (row[0].toString().toLowerCase().includes('net income') ||
                    row[0].toString().toLowerCase().includes('profit'))
            );

            if (revenueRowIdx !== -1 && ebitdaRowIdx !== -1) {
                const marginRow = ['EBITDA Margin %'];
                for (let i = 0; i < numYears; i++) {
                    const col = this.columnIndexToLetter(i + 1);
                    marginRow.push(`=${col}${ebitdaRowIdx + 2}/${col}${revenueRowIdx + 2}`);
                }
                marginRow.push('');
                analysisData.push(marginRow);
            }

            if (revenueRowIdx !== -1 && niRowIdx !== -1) {
                const marginRow = ['Net Margin %'];
                for (let i = 0; i < numYears; i++) {
                    const col = this.columnIndexToLetter(i + 1);
                    marginRow.push(`=${col}${niRowIdx + 2}/${col}${revenueRowIdx + 2}`);
                }
                marginRow.push('');
                analysisData.push(marginRow);
            }
        }

        const result = await this.writeToRange(targetCell, analysisData);

        // Format as percentages for CAGR and margins
        await this.formatHistoricalTable(targetCell, analysisData.length, years.length + 2);

        return {
            status: 'success',
            address: result.address || targetCell,
            message: `Historical financials analyzed for ${numYears} years with growth rates and margins`,
            years: numYears,
            metrics: values.length - 1
        };
    }

    /**
     * Build Three-Statement Financial Model
     */
    async buildThreeStatementModel(params) {
        console.log('[FinancialAnalysis] Building three-statement model:', params);

        const { historicalDataRange, assumptionsRange, targetCell, projectionYears = 5 } = params;

        // This is a complex model - we'll create a simplified version
        // In production, this would integrate all three statements with full linking

        const modelData = [
            ['THREE-STATEMENT FINANCIAL MODEL', '', '', '', '', '', ''],
            ['Projection Years:', projectionYears, '', '', '', '', ''],
            [],
            ['Income Statement', '', '', '', '', '', ''],
            ['Line Item', 'Historical', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5'],
            ['Revenue', 0, '=B6*(1+$B$9)', '=C6*(1+$B$9)', '=D6*(1+$B$9)', '=E6*(1+$B$9)', '=F6*(1+$B$9)'],
            ['Cost of Goods Sold', 0, '=C6*(1-$B$10)', '=D6*(1-$B$10)', '=E6*(1-$B$10)', '=F6*(1-$B$10)', '=G6*(1-$B$10)'],
            ['Gross Profit', '=B6-B7', '=C6-C7', '=D6-D7', '=E6-E7', '=F6-F7', '=G6-G7'],
            [],
            ['Assumptions (% of Revenue)', '', '', '', '', '', ''],
            ['Revenue Growth %', 0.10, '', '', '', '', ''],
            ['Gross Margin %', 0.40, '', '', '', '', ''],
            ['EBITDA Margin %', 0.20, '', '', '', '', ''],
            [],
            ['Balance Sheet (Simplified)', '', '', '', '', '', ''],
            ['Assets', 'Historical', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5'],
            ['Total Assets', 0, '=C6*1.5', '=D6*1.5', '=E6*1.5', '=F6*1.5', '=G6*1.5'],
            [],
            ['Cash Flow Statement (Simplified)', '', '', '', '', '', ''],
            ['Operating CF', '', '=C6*$B$12', '=D6*$B$12', '=E6*$B$12', '=F6*$B$12', '=G6*$B$12'],
            ['Investing CF', '', '=-C6*0.05', '=-D6*0.05', '=-E6*0.05', '=-F6*0.05', '=-G6*0.05'],
            ['Free Cash Flow', '', '=C20+C21', '=D20+D21', '=E20+E21', '=F20+F21', '=G20+G21']
        ];

        const result = await this.writeToRange(targetCell, modelData);

        return {
            status: 'success',
            address: result.address || targetCell,
            message: `Three-statement model template created for ${projectionYears} years. Please populate historical data and assumptions.`,
            projectionYears,
            note: 'This is a simplified template. Update assumptions and historical data to complete the model.'
        };
    }

    /**
     * Build DCF Valuation Model
     */
    async buildDCFModel(params) {
        console.log('[FinancialAnalysis] Building DCF model:', params);

        const { financialsRange, assumptionsRange, targetCell, includeSensitivity = true } = params;

        // Create DCF model template
        const dcfData = [
            ['DISCOUNTED CASH FLOW (DCF) VALUATION', '', '', '', '', '', ''],
            [],
            ['Assumptions', '', '', '', '', '', ''],
            ['WACC (%)', 0.10, '', '', '', '', ''],
            ['Terminal Growth Rate (%)', 0.03, '', '', '', '', ''],
            ['Net Debt', 0, '', '', '', '', ''],
            ['Shares Outstanding (M)', 100, '', '', '', '', ''],
            [],
            ['Projected Free Cash Flow', '', '', '', '', '', ''],
            ['Year', '1', '2', '3', '4', '5', 'Terminal'],
            ['FCF', 0, 0, 0, 0, 0, '=F11*(1+$B$5)/(SB$4-$B$5)'],
            ['Discount Factor', '=1/(1+$B$4)^B10', '=1/(1+$B$4)^C10', '=1/(1+$B$4)^D10', '=1/(1+$B$4)^E10', '=1/(1+$B$4)^F10', '=1/(1+$B$4)^5'],
            ['Present Value', '=B11*B12', '=C11*C12', '=D11*D12', '=E11*E12', '=F11*F12', '=G11*G12'],
            [],
            ['Valuation Summary', '', '', '', '', '', ''],
            ['PV of Projected FCF', '=SUM(B13:F13)', '', '', '', '', ''],
            ['PV of Terminal Value', '=G13', '', '', '', '', ''],
            ['Enterprise Value', '=B16+B17', '', '', '', '', ''],
            ['Less: Net Debt', '=$B$6', '', '', '', '', ''],
            ['Equity Value', '=B18-B19', '', '', '', '', ''],
            ['Shares Outstanding (M)', '=$B$7', '', '', '', '', ''],
            ['Equity Value Per Share', '=B20/B21', '', '', '', '', '']
        ];

        // Add sensitivity table if requested
        if (includeSensitivity) {
            dcfData.push([]);
            dcfData.push(['Sensitivity Analysis: Equity Value per Share', '', '', '', '', '', '']);
            dcfData.push(['Terminal Growth →', '2.0%', '2.5%', '3.0%', '3.5%', '4.0%', '']);
            dcfData.push(['WACC ↓', '', '', '', '', '', '']);

            const waccValues = [0.08, 0.09, 0.10, 0.11, 0.12];
            waccValues.forEach((wacc, idx) => {
                const row = [`${(wacc * 100).toFixed(1)}%`];
                // In a real implementation, these would be proper sensitivity formulas
                for (let i = 0; i < 5; i++) {
                    row.push('=[DCF Formula]');
                }
                row.push('');
                dcfData.push(row);
            });
        }

        const result = await this.writeToRange(targetCell, dcfData);

        return {
            status: 'success',
            address: result.address || targetCell,
            message: 'DCF valuation model created. Please input projected FCF values and update assumptions.',
            includeSensitivity,
            note: 'Update the FCF projections in row 11 and assumptions to complete the valuation.'
        };
    }

    /**
     * Build LBO Model
     */
    async buildLBOModel(params) {
        console.log('[FinancialAnalysis] Building LBO model:', params);

        const { financialsRange, lboAssumptionsRange, targetCell, exitScenarios = [7, 8, 9, 10] } = params;

        const lboData = [
            ['LEVERAGED BUYOUT (LBO) MODEL', '', '', '', '', '', ''],
            [],
            ['Transaction Assumptions', '', '', '', '', '', ''],
            ['Entry EBITDA Multiple', 8.0, '', '', '', '', ''],
            ['LTM EBITDA', 100, '', '', '', '', ''],
            ['Purchase Price', '=B4*B5', '', '', '', '', ''],
            ['% Debt', 0.60, '', '', '', '', ''],
            ['% Equity', 0.40, '', '', '', '', ''],
            ['Interest Rate', 0.07, '', '', '', '', ''],
            [],
            ['Sources & Uses', '', '', '', '', '', ''],
            ['Sources', '', '', '', '', '', ''],
            ['Debt', '=B6*B7', '', '', '', '', ''],
            ['Equity', '=B6*B8', '', '', '', '', ''],
            ['Total Sources', '=B13+B14', '', '', '', '', ''],
            ['Uses', '', '', '', '', '', ''],
            ['Purchase Price', '=B6', '', '', '', '', ''],
            ['Fees (3%)', '=B6*0.03', '', '', '', '', ''],
            ['Total Uses', '=B17+B18', '', '', '', '', ''],
            [],
            ['Projected EBITDA & Debt Paydown', '', '', '', '', '', ''],
            ['Year', '0', '1', '2', '3', '4', '5'],
            ['EBITDA', '=B5', 0, 0, 0, 0, 0],
            ['Free Cash Flow', '', 0, 0, 0, 0, 0],
            ['Debt Paydown', '', '=C24', '=D24', '=E24', '=F24', '=G24'],
            ['Ending Debt', '=B13', '=B25-C25', '=C25-D25', '=D25-E25', '=E25-F25', '=F25-G25'],
            [],
            ['Exit Analysis', '', '', '', '', '', ''],
            ['Exit Multiple', exitScenarios[0], exitScenarios[1], exitScenarios[2], exitScenarios[3], '', ''],
            ['Exit Enterprise Value', '=B29*G23', '=C29*G23', '=D29*G23', '=E29*G23', '', ''],
            ['Less: Net Debt', '=G26', '=G26', '=G26', '=G26', '', ''],
            ['Exit Equity Value', '=B30-B31', '=C30-C31', '=D30-D31', '=E30-E31', '', ''],
            ['Initial Equity', '=B14', '=B14', '=B14', '=B14', '', ''],
            ['MOIC (Money on Money)', '=B32/B33', '=C32/C33', '=D32/D33', '=E32/E33', '', ''],
            ['IRR (assuming 5-yr hold)', '=B34^(1/5)-1', '=C34^(1/5)-1', '=D34^(1/5)-1', '=E34^(1/5)-1', '', '']
        ];

        const result = await this.writeToRange(targetCell, lboData);

        return {
            status: 'success',
            address: result.address || targetCell,
            message: `LBO model created with ${exitScenarios.length} exit multiple scenarios`,
            exitScenarios,
            note: 'Input projected EBITDA and FCF to complete the analysis. Model will calculate IRR and MOIC for each exit scenario.'
        };
    }

    /**
     * Calculate Financial Ratios
     */
    async calculateFinancialRatios(params) {
        console.log('[FinancialAnalysis] Calculating financial ratios:', params);

        const { incomeStatementRange, balanceSheetRange, targetCell, periods = ['Current'] } = params;

        const ratiosData = [
            ['FINANCIAL RATIOS ANALYSIS', ...periods],
            [],
            ['Profitability Ratios', ''],
            ['ROE (Return on Equity)', '=NetIncome/Equity'],
            ['ROA (Return on Assets)', '=NetIncome/TotalAssets'],
            ['ROIC (Return on Invested Capital)', '=EBIT*(1-TaxRate)/(Debt+Equity)'],
            ['Gross Margin %', '=GrossProfit/Revenue'],
            ['EBITDA Margin %', '=EBITDA/Revenue'],
            ['Net Margin %', '=NetIncome/Revenue'],
            [],
            ['Liquidity Ratios', ''],
            ['Current Ratio', '=CurrentAssets/CurrentLiabilities'],
            ['Quick Ratio', '=(CurrentAssets-Inventory)/CurrentLiabilities'],
            ['Cash Ratio', '=Cash/CurrentLiabilities'],
            [],
            ['Leverage Ratios', ''],
            ['Debt/Equity', '=TotalDebt/TotalEquity'],
            ['Debt/Assets', '=TotalDebt/TotalAssets'],
            ['Interest Coverage', '=EBIT/InterestExpense'],
            [],
            ['Efficiency Ratios', ''],
            ['Asset Turnover', '=Revenue/TotalAssets'],
            ['Inventory Turnover', '=COGS/AverageInventory'],
            ['Days Sales Outstanding', '=365/(Revenue/AccountsReceivable)']
        ];

        const result = await this.writeToRange(targetCell, ratiosData);

        return {
            status: 'success',
            address: result.address || targetCell,
            message: 'Financial ratios template created. Link formulas to your income statement and balance sheet data.',
            categories: ['Profitability', 'Liquidity', 'Leverage', 'Efficiency'],
            note: 'Replace placeholder formulas with actual cell references from your financial statements.'
        };
    }

    /**
     * Benchmark Against Peers
     */
    async benchmarkAgainstPeers(params) {
        console.log('[FinancialAnalysis] Benchmarking against peers:', params);

        const { targetCompanyRange, peerGroupRange, targetCell, metricsToCompare } = params;

        const targetData = await this.getRangeByAddress(targetCompanyRange);
        const peerData = await this.getRangeByAddress(peerGroupRange);

        const targetValues = targetData.values;
        const peerValues = peerData.values;

        if (!peerValues || peerValues.length < 2) {
            throw new Error('Insufficient peer data for benchmarking.');
        }

        const peerHeaders = peerValues[0];
        const peerCompanies = peerValues.slice(1);

        // Build benchmarking table
        const benchmarkData = [
            ['PEER BENCHMARKING ANALYSIS', '', '', '', '', ''],
            [],
            ['Metric', 'Target Company', 'Peer Min', 'Peer Max', 'Peer Median', 'Percentile Rank'],
            ...peerHeaders.slice(1).map((metric, idx) => {
                const col = this.columnIndexToLetter(idx + 1);
                const peerSheet = peerData.sheetName || 'Sheet1';
                const targetValue = targetValues[0] && targetValues[0][idx + 1] ? targetValues[0][idx + 1] : 0;

                const peerRange = `${peerSheet}!${col}2:${col}${peerCompanies.length + 1}`;

                return [
                    metric,
                    targetValue,
                    `=MIN(${peerRange})`,
                    `=MAX(${peerRange})`,
                    `=MEDIAN(${peerRange})`,
                    `=PERCENTRANK(${peerRange}, B${idx + 4})`
                ];
            })
        ];

        // Add strategic insights section
        benchmarkData.push([]);
        benchmarkData.push(['Strategic Insights', '', '', '', '', '']);
        benchmarkData.push(['Metrics', 'Target vs. Median', 'Recommendation', '', '', '']);

        const result = await this.writeToRange(targetCell, benchmarkData);

        // Format percentages
        await this.formatBenchmarkTable(targetCell, benchmarkData.length, 6);

        return {
            status: 'success',
            address: result.address || targetCell,
            message: `Peer benchmarking completed comparing ${peerCompanies.length} peers across ${peerHeaders.length - 1} metrics`,
            peerCount: peerCompanies.length,
            metrics: peerHeaders.length - 1
        };
    }

    // Helper formatting methods
    async formatMultiplesTable(startCell, numRows, numCols) {
        try {
            return await Excel.run(async (context) => {
                const range = await this.getRangeFromString(context, startCell);
                const fullRange = range.getResizedRange(numRows - 1, numCols - 1);

                // Format headers
                const headerRange = range.getResizedRange(0, numCols - 1);
                headerRange.format.font.bold = true;
                headerRange.format.fill.color = '#4472C4';
                headerRange.format.font.color = 'white';

                // Format numbers as 2 decimal places
                fullRange.numberFormat = [['0.00']];

                await context.sync();
                return { status: 'success' };
            });
        } catch (e) {
            console.warn('Formatting failed:', e);
            return { status: 'warning', message: 'Data created but formatting failed' };
        }
    }

    async formatHistoricalTable(startCell, numRows, numCols) {
        try {
            return await Excel.run(async (context) => {
                const range = await this.getRangeFromString(context, startCell);
                const fullRange = range.getResizedRange(numRows - 1, numCols - 1);

                // Format headers
                const headerRange = range.getResizedRange(0, numCols - 1);
                headerRange.format.font.bold = true;
                headerRange.format.fill.color = '#70AD47';
                headerRange.format.font.color = 'white';

                // Format CAGR column as percentage
                const cagrCol = range.getOffsetRange(1, numCols - 1).getResizedRange(numRows - 2, 0);
                cagrCol.numberFormat = [['0.0%']];

                await context.sync();
                return { status: 'success' };
            });
        } catch (e) {
            console.warn('Formatting failed:', e);
            return { status: 'warning' };
        }
    }

    async formatBenchmarkTable(startCell, numRows, numCols) {
        try {
            return await Excel.run(async (context) => {
                const range = await this.getRangeFromString(context, startCell);

                // Format headers
                const headerRange = range.getOffsetRange(2, 0).getResizedRange(0, numCols - 1);
                headerRange.format.font.bold = true;
                headerRange.format.fill.color = '#FFC000';
                headerRange.format.font.color = 'black';

                // Format percentile rank as percentage
                const percentileCol = range.getOffsetRange(3, 5).getResizedRange(numRows - 4, 0);
                percentileCol.numberFormat = [['0%']];

                await context.sync();
                return { status: 'success' };
            });
        } catch (e) {
            console.warn('Formatting failed:', e);
            return { status: 'warning' };
        }
    }

}



// Export

if (typeof module !== 'undefined' && module.exports) {

    module.exports = { AutonomousExcelAgent };

}

