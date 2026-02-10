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

            if (this.onStatusUpdate) this.onStatusUpdate("📋 Creating execution plan...");

            // Step 2: Create execution plan
            const plan = await this.createExecutionPlan(intent, userQuery);

            // Step 3: Get user approval (optional)
            const approved = await this.presentPlan(plan);

            if (!approved) {
                return { status: 'cancelled', message: 'User cancelled operation' };
            }

            // Step 4: Execute the plan with granular recovery
            const result = await this.executePlan(plan);

            // Step 5: Return results
            return {
                status: 'success',
                query: userQuery,
                intent: intent,
                plan: plan,
                result: result
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
        const prompt = `You are a senior data analyst with deep expertise in Excel automation. Analyze the user's query to understand their true intent.

USER QUERY: "${userQuery}"

YOUR TASK:
1. Identify the PRIMARY intent (what they explicitly asked for)
2. Identify IMPLICIT needs (what they'll need to accomplish the primary intent)
3. Assess confidence based on query clarity
4. Extract specific entities and parameters

INTENT CATEGORIES:
- analyze_data: General data analysis, insights, patterns
- create_chart: Specific visualization request
- forecast: Predictions, future projections
- clean_data: Data transformation, cleaning, formatting
- create_dashboard: Comprehensive view with multiple charts
- find_insights: Answer specific questions about data
- calculate: Compute metrics, formulas, aggregations
- format: Styling, colors, conditional formatting
- filter_sort: Data filtering, sorting, organizing
- custom: Complex multi-step operations

CONFIDENCE LEVELS:
- high: Query is clear and specific ("Create a line chart of sales over time")
- medium: Query is somewhat vague ("Analyze this data")
- low: Query is ambiguous or unclear ("Do something with this")

Respond in JSON:
{
  "intent": "primary intent category",
  "implicitNeeds": ["list of implied requirements"],
  "confidence": "high/medium/low",
  "confidenceReason": "why this confidence level",
  "entities": {
    "dataRange": "A1:C10 or null",
    "chartType": "column/line/pie/etc or null",
    "timeframe": "12 months/next quarter/etc or null",
    "metric": "sales/revenue/profit/etc or null",
    "operation": "sum/average/count/filter/sort/etc or null"
  },
  "parameters": {},
  "reasoning": "Deep analysis of what user wants and why",
  "alternativeInterpretations": ["other possible meanings if confidence is not high"]
}

EXAMPLES:
Query: "Create a dashboard" → intent: create_dashboard, implicitNeeds: ["analyze data first", "determine best charts", "create summary"]
Query: "Show top 10 products" → intent: filter_sort, implicitNeeds: ["sort by sales", "filter top 10", "possibly create chart"]
Query: "Clean this data" → intent: clean_data, confidence: medium (unclear what cleaning is needed)`;

        const systemPrompt = "You are a senior data analyst. Deeply analyze user intent for Excel operations. Think about implicit needs. Return only valid JSON.";

        const response = await this.callClaudeAPI(prompt, systemPrompt);
        const intent = this.extractJSON(response);

        if (!intent || !intent.intent) {
            throw new Error(`Failed to analyze user intent. Agent response: ${response.substring(0, 100)}...`);
        }

        return intent;
    }

    /**
     * Create detailed execution plan
     */
    async createExecutionPlan(intent, userQuery) {
        // Get current Excel context
        const context = await this.getExcelContext();

        const prompt = `You are an expert Excel automation planner. Create a detailed, robust execution plan.

═══════════════════════════════════════════════════════════
USER REQUEST
═══════════════════════════════════════════════════════════
Query: "${userQuery}"

Intent Analysis:
${JSON.stringify(intent, null, 2)}

═══════════════════════════════════════════════════════════
CURRENT EXCEL CONTEXT
═══════════════════════════════════════════════════════════
${JSON.stringify(context, null, 2)}

Data Sample (First 10 rows):
${this.formatDataForAI(context.dataSample, context.address)}

═══════════════════════════════════════════════════════════
AVAILABLE TOOLS (23 tools total)
═══════════════════════════════════════════════════════════

📊 ANALYSIS TOOLS:
1. analyzeData - Comprehensive analysis (descriptive, trends, correlations, forecasting, etc.)
2. findInsights - Answer specific question (params: question)
3. detectOutliers - Find anomalies
4. correlationAnalysis - Find variable relationships
5. trendAnalysis - Analyze changes over time
6. segmentation - Cluster/group data
7. calculateMetric - Calculate custom metric (params: metric)
8. createPivot - Pivot table (params: rows, columns, values)

📈 VISUALIZATION TOOLS:
9. createChart - Single chart (params: dataRange, chartType, title, xAxis, yAxis)
   - chartType: line, column, bar, pie, area, scatter
   - CRITICAL: dataRange must be valid Excel address or "selection" or "result from step X"
10. generateDashboard - Complete dashboard with multiple charts (auto-analyzes and creates 6-8 charts)

🔧 DATA TRANSFORMATION TOOLS:
11. cleanData - Transform/clean data (params: instructions - natural language)
   - Returns: { address, status, message } - USE THIS ADDRESS in subsequent steps!
12. filterData - Filter rows (params: column, operator, value) [NEW]
   - operator: equals, not_equals, greater_than, less_than, contains, etc.
   - Returns: { address, rowsReturned, sheetName }
13. sortData - Sort data (params: columns, order) [NEW]
   - columns: array of column names, order: "asc" or "desc"
   - Returns: { address, sortedBy, sheetName }
14. mergeData - Combine ranges (params: ranges, mergeType) [NEW]
   - mergeType: "horizontal" or "vertical"
15. validateData - Check data quality (params: rules) [NEW]

📝 FORMULA & CALCULATION TOOLS:
16. generateFormula - AI creates formula (params: description, address)
   - If address provided, formula is auto-applied
   - Returns: { formula, address, status }
17. insertFormula - Insert existing formula (params: address, formula)
   - formula MUST start with =

📋 SHEET & DATA MANAGEMENT:
18. createSheet - New worksheet (params: name)
19. writeData - Write 2D array (params: address, data)
   - Returns: { address, rowsWritten, columnsWritten }
20. exportToNewSheet - Copy data to new sheet (params: data, sheetName) [NEW]
21. createSummary - Generate summary sheet

🎨 FORMATTING:
22. formatData - Apply styling (params: range, formatting)

🔮 FORECASTING:
23. forecast - Predict future (params: periods)

═══════════════════════════════════════════════════════════
PLANNING GUIDELINES
═══════════════════════════════════════════════════════════

✅ BEST PRACTICES:
1. **Tool Chaining**: When a tool creates data, it returns an "address". Reference this in next step:
   - Step 1: cleanData → returns { address: "Cleaned_1234!A1:C100" }
   - Step 2: createChart with dataRange: "result from step 1"

2. **Data Range Specification**:
   - Use specific ranges: "A2:C100" (not "A:C")
   - Use "selection" for user's current selection
   - Use "result from step X" to reference previous step output

3. **Dashboard Creation Pattern**:
   - Step 1: analyzeData (understand data structure)
   - Step 2: cleanData if needed (returns clean data address)
   - Step 3: generateDashboard (creates multiple charts automatically)

4. **Filter/Sort Pattern**:
   - Step 1: filterData or sortData (returns new sheet with filtered/sorted data)
   - Step 2: createChart using the returned address

5. **Error Prevention**:
   - Always specify required parameters
   - For formulas, provide clear descriptions
   - Mark critical steps with "critical": true

📚 SUCCESS PATTERNS:

Pattern: "Create dashboard"
→ Step 1: analyzeData
→ Step 2: generateDashboard

Pattern: "Show top 10 products by sales"
→ Step 1: sortData (columns: ["Sales"], order: "desc")
→ Step 2: filterData (use first 10 rows)
→ Step 3: createChart with sorted/filtered data

Pattern: "Clean data and visualize"
→ Step 1: cleanData (returns address)
→ Step 2: createChart (dataRange: "result from step 1")

═══════════════════════════════════════════════════════════
YOUR TASK
═══════════════════════════════════════════════════════════

Create a detailed execution plan using the ReAct pattern:
- **Thought**: Deep reasoning about WHY this step and WHAT you expect to see
- **Action**: Clear description of what the step does
- **Method**: Exact tool name from the list above
- **Parameters**: All required parameters with correct types
- **ExpectedOutcome**: Specific, verifiable result

Return JSON:
{
  "summary": "Overall strategy in 1-2 sentences",
  "reasoning": "Why this approach will work",
  "steps": [
    {
      "stepNumber": 1,
      "thought": "I need to [deep reasoning]... I expect to observe [specific outcome]...",
      "action": "Clear description of what this step accomplishes",
      "method": "exact tool name",
      "parameters": { /* all required params */ },
      "expectedOutcome": "Specific, verifiable result (e.g., 'Cleaned data in new sheet with address returned')",
      "critical": true  // true if failure would break the entire plan
    }
  ],
  "estimatedTotalTime": "e.g., 45s",
  "fallbackStrategy": "What to do if primary approach fails"
}`;

        const systemPrompt = "You are an expert Excel automation planner. Create detailed, executable plans. Return only valid JSON.";

        const response = await this.callClaudeAPI(prompt, systemPrompt);
        this.executionPlan = this.extractJSON(response);

        if (!this.executionPlan || !this.executionPlan.steps) {
            throw new Error(`Failed to create a valid execution plan. Agent response: ${response.substring(0, 100)}...`);
        }

        return this.executionPlan;
    }

    /**
     * Get current Excel context
     */
    async getExcelContext() {
        try {
            const sheets = await this.getAllSheetNames();
            const activeData = await this.getWorksheetData();
            const selection = await this.getSelectedRange();

            // Extract a sample of the data (e.g., first 10 rows)
            const dataSample = activeData.values ? activeData.values.slice(0, 10) : [];

            return {
                sheets: sheets,
                activeSheet: activeData.sheetName,
                dataRows: activeData.rowCount,
                dataColumns: activeData.columnCount,
                selection: selection.address,
                hasSelection: selection.rowCount > 0,
                address: activeData.address,
                dataSample: dataSample
            };
        } catch (error) {
            return { error: "Could not get Excel context" };
        }
    }

    /**
     * Present plan to user for approval
     */
    async presentPlan(plan) {
        console.log("📋 Execution Plan:");
        console.log(`Summary: ${plan.summary}`);
        console.log(`Total Steps: ${plan.steps.length}`);
        console.log(`Estimated Time: ${plan.estimatedTotalTime || 'unknown'}`);

        plan.steps.forEach(step => {
            console.log(`  ${step.stepNumber}. ${step.action}`);
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
                    ? `🔄 Retry ${retryCount}: ${step.thought}`
                    : `💭 Thought: ${step.thought}`;

                console.log(thoughtMsg);
                if (this.onStatusUpdate) this.onStatusUpdate(thoughtMsg);

                // ReAct - Action
                const currentStepNum = step.stepNumber || (i + 1);
                const actionMsg = `⏳ Step ${currentStepNum}/${totalSteps}: ${step.action}`;
                console.log(actionMsg);
                if (this.onStatusUpdate) this.onStatusUpdate(actionMsg);

                try {
                    const stepResult = await this.executeStep(step, results);

                    // ReAct - Observation
                    if (this.onStatusUpdate) this.onStatusUpdate(`👀 Observing results...`);
                    const observation = await this.getExcelContext();

                    results.push({
                        step: step.stepNumber,
                        thought: step.thought,
                        action: step.action,
                        status: 'success',
                        result: stepResult,
                        observation: observation
                    });

                    console.log(`✅ Step ${step.stepNumber} completed and verified`);
                    stepSuccess = true;

                } catch (error) {
                    const currentStepNum = step.stepNumber || (i + 1);
                    retryCount++;
                    console.error(`❌ Step ${currentStepNum} failed (Attempt ${retryCount}):`, error);

                    if (retryCount <= MAX_STEP_RETRIES) {
                        const healingStatus = `🔄 Step ${step.stepNumber} failed. Initiating localized self-healing...`;
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
                            step: step.stepNumber,
                            action: step.action,
                            status: 'error',
                            error: error.message
                        });

                        if (step.critical) {
                            throw new Error(`Critical step ${step.stepNumber} failed after ${MAX_STEP_RETRIES} recovery attempts: ${error.message}`);
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
1. analyzeData - Comprehensive analysis
2. createChart - Single chart (params: dataRange, chartType, title, xAxis, yAxis)
3. generateDashboard - Complete dashboard with multiple charts
4. forecast - Predict future values (params: periods)
5. cleanData - Transform/clean data (params: instructions)
6. calculateMetric - Calculate custom metrics (params: metric)
7. findInsights - Answer questions (params: question)
8. formatData - Apply formatting (params: range, formatting)
9. createSummary - Generate summary sheet
10. detectOutliers - Find anomalies
11. correlationAnalysis - Find relationships
12. trendAnalysis - Analyze trends
13. segmentation - Cluster data
14. generateFormula - Create formula (params: description, address)
15. createPivot - Pivot table (params: rows, columns, values)
16. insertFormula - Insert formula (params: address, formula)
17. createSheet - New worksheet (params: name)
18. writeData - Write values (params: address, data)
19. filterData - Filter rows (params: column, operator, value) [NEW]
20. sortData - Sort data (params: columns, order) [NEW]
21. mergeData - Combine ranges (params: ranges, mergeType) [NEW]
22. validateData - Check quality (params: rules) [NEW]
23. exportToNewSheet - Copy to new sheet (params: data, sheetName) [NEW]

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
        let params = step.parameters || {};

        // Resolve parameters that refer to previous steps
        // Example: "Pivot results from step 5"
        for (let key in params) {
            if (typeof params[key] === 'string' && params[key].toLowerCase().includes('step')) {
                const stepNumMatch = params[key].match(/step\s*(\d+)/i);
                if (stepNumMatch) {
                    const targetStepNum = parseInt(stepNumMatch[1]);
                    const prevStep = previousResults.find(r => r.step === targetStepNum);
                    if (prevStep && prevStep.result && prevStep.result.address) {
                        console.log(`[Agent] Resolved parameter ${key} from Step ${targetStepNum}: ${prevStep.result.address}`);
                        params[key] = prevStep.result.address;
                    }
                }
            }
        }

        // Normalize method name (handle common AI variations)
        method = method.replace(/_/g, ''); // Convert snake_case to CamelCase-ish
        const normalizedMethod = method.toLowerCase();

        // Map method names to actual functions
        if (normalizedMethod === 'analyzedata') return await this.performCompleteAnalysis();

        if (normalizedMethod === 'createchart') {
            // Handle missing dataRange by checking for previous data-creating steps
            if (!params.dataRange || params.dataRange === 'selection') {
                const lastDataStep = [...previousResults].reverse().find(r => r.result && r.result.address);
                if (lastDataStep) {
                    params.dataRange = lastDataStep.result.address;
                    console.log(`[Agent] Auto-assigned chart dataRange from Step ${lastDataStep.step}: ${params.dataRange}`);
                } else {
                    const selection = await this.getSelectedRange();
                    params.dataRange = selection.address;
                }
            }
            return await this.createChart(params.dataRange, params.chartType, params.title, params.xAxis, params.yAxis);
        }
        if (normalizedMethod === 'generatedashboard') return await this.generateCompleteDashboard();
        if (normalizedMethod === 'forecast') return await this.forecastingAnalysis(params.periods || 12);
        if (normalizedMethod === 'cleandata' || normalizedMethod === 'transformdata' || normalizedMethod === 'transform') {
            const instructions = params.instructions || params.instruction || step.action;
            return await this.applyCleanedData(instructions);
        }
        if (normalizedMethod === 'calculatemetric' || normalizedMethod === 'calculate') return await this.calculateMetric(params);
        if (normalizedMethod === 'findinsights' || normalizedMethod === 'ask') return await this.askAboutData(params.question);
        if (normalizedMethod === 'formatdata' || normalizedMethod === 'format') return await this.formatRange(params.range, params.formatting);
        if (normalizedMethod === 'createsummary') return await this.generateSummary();
        if (normalizedMethod === 'detectoutliers') return await this.outlierAnalysis();
        if (normalizedMethod === 'correlationanalysis') return await this.correlationAnalysis();
        if (normalizedMethod === 'trendanalysis') return await this.trendAnalysis();
        if (normalizedMethod === 'segmentation') return await this.segmentationAnalysis();
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
            let targetAddress = params.address || params.range || params.targetRange;
            if (!targetAddress || targetAddress.toLowerCase() === 'selection') {
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
        if (normalizedMethod === 'insertformula' || normalizedMethod === 'applyformula') {
            let targetAddress = params.address || params.range;
            if (!targetAddress || targetAddress.toLowerCase() === 'selection') {
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
            if (!params.address || params.address === 'undefined') {
                const selection = await this.getSelectedRange();
                params.address = selection.address;
                console.log(`[Agent] Auto-assigned writeData address: ${params.address}`);
            }
            return await this.writeToRange(params.address, params.data);
        }

        // New Tools - Data Transformation
        if (normalizedMethod === 'filterdata' || normalizedMethod === 'filter') {
            return await this.filterData(params.column, params.operator, params.value);
        }
        if (normalizedMethod === 'sortdata' || normalizedMethod === 'sort') {
            return await this.sortData(params.columns, params.order || 'asc');
        }
        if (normalizedMethod === 'mergedata' || normalizedMethod === 'merge') {
            return await this.mergeData(params.ranges, params.mergeType || 'horizontal');
        }
        if (normalizedMethod === 'validatedata' || normalizedMethod === 'validate') {
            return await this.validateData(params.rules);
        }
        if (normalizedMethod === 'exporttonewsheet' || normalizedMethod === 'export') {
            return await this.exportToNewSheet(params.data, params.sheetName);
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
        const data = await this.getWorksheetData();
        const dataStr = this.formatDataForAI(data.values, data.address);

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

        // Forecasting
        if (lowerQuery.includes('forecast') || lowerQuery.includes('predict')) {
            const periods = this.extractNumber(query) || 12;
            return await this.processQuery(`Forecast the next ${periods} periods`);
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
    async filterData(column, operator, value) {
        const data = await this.getWorksheetData();
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
    async sortData(columns, order = 'asc') {
        const data = await this.getWorksheetData();
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
    async validateData(rules) {
        const data = await this.getWorksheetData();
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
}

// Export
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { AutonomousExcelAgent };
}
