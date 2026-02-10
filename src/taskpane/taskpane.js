
class ExcelAIAssistant {
  constructor(apiKey) {
    this.apiKey = apiKey;
    this.claudeApiUrl = 'https://openrouter.ai/api/v1/chat/completions';
    this.model = 'google/gemini-2.5-flash';
  }


  async callClaudeAPI(userMessage, systemPrompt = null) {
    try {
      const headers = {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${this.apiKey.trim()}`,
        'HTTP-Referer': 'https://localhost:3000',
        'X-Title': 'Excel AI Dashboard'
      };

      const body = {
        model: this.model,
        messages: [
          {
            role: 'user',
            content: userMessage
          }
        ]
      };

      if (systemPrompt) {
        body.messages.unshift({
          role: 'system',
          content: systemPrompt
        });
      }

      const response = await fetch(this.claudeApiUrl, {
        method: 'POST',
        headers: headers,
        body: JSON.stringify(body)
      });

      if (!response.ok) {
        const error = await response.json();
        throw new Error(`API Error: ${error.error?.message || response.statusText}`);
      }

      const data = await response.json();
      return this.stripMarkdown(data.choices[0].message.content);
    } catch (error) {
      console.error('Claude API Error:', error);
      throw error;
    }
  }
  stripMarkdown(text) {
    if (!text) return "";
    return text.replace(/```[a-z]*\n?/gi, '').replace(/\n?```/g, '').trim();
  }


  extractJSON(text) {
    if (!text) return null;

    try {

      const trimmed = text.trim();
      // Remove common markdown wrappers if they wrap the whole thing
      let cleaned = trimmed.replace(/^```json\s*/i, '').replace(/\s*```$/i, '').trim();

      try {
        return JSON.parse(cleaned);
      } catch (e) {
        // AI JSON Repair Logic
        try {
          // 1. Fix unescaped control characters in values
          cleaned = cleaned.replace(/":\s*"(.*?)"/gs, (match, p1) => {
            return '": "' + p1.replace(/\n/g, '\\n').replace(/\t/g, '\\t').replace(/\r/g, '\\r') + '"';
          });

          // 1.1 Fix unescaped double quotes inside string values
          // This matches a double quote that is NOT preceded by a backslash AND NOT followed by a comma/brace/bracket
          // It's a heuristic but helps with things like "value with "quotes" inside"
          cleaned = cleaned.replace(/":\s*"(.*?)"\s*([,}])/gs, (match, p1, suffix) => {
            const repairedValue = p1.replace(/(?<!\\)"/g, '\\"');
            return '": "' + repairedValue + '"' + suffix;
          });

          // 2. Fix trailing commas (e.g., [1, 2,] or {a:1,})
          cleaned = cleaned.replace(/,\s*([\]}])/g, '$1');

          // 3. Attempt to repair truncated JSON by adding missing closers
          // Check for balanced braces and brackets
          const stack = [];
          for (let i = 0; i < cleaned.length; i++) {
            const char = cleaned[i];
            if (char === '{' || char === '[') stack.push(char === '{' ? '}' : ']');
            else if (char === '}' || char === ']') {
              if (stack.length > 0 && stack[stack.length - 1] === char) stack.pop();
            }
          }
          // Append missing closers in reverse order
          while (stack.length > 0) {
            cleaned += stack.pop();
          }

          // 3.5. Fix literal newlines in values (Bad control character fix)
          // This must happen BEFORE other repairs to avoid breaking quotes
          cleaned = cleaned.replace(/":\s*"(.*?)"/gs, (match, p1) => {
            return '": "' + p1.replace(/\n/g, '\\n').replace(/\r/g, '\\r').replace(/\t/g, '\\t') + '"';
          });

          // 4. Fix missing commas between properties (handles complex values like objects/arrays)
          // This looks for a value end (quote, bracket, brace, or literal) followed by a new property start (quote)
          cleaned = cleaned.replace(/([\]}]|"(?:[^"\\]|\\.)*"|\d+|true|false|null)\s*"/g, '$1, "');

          // 5. Fix double commas or comma before closing
          cleaned = cleaned.replace(/,(\s*,)+/g, ',');
          cleaned = cleaned.replace(/,\s*([\]}])/g, '$1');

          return JSON.parse(cleaned);
        } catch (innerE) {
          // Fallback to brace finding logic
        }

        // Secondary attempt: find the first { and last }
        const firstBrace = text.indexOf('{');
        const lastBrace = text.lastIndexOf('}');

        if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
          let jsonCandidate = text.substring(firstBrace, lastBrace + 1);
          try {
            return JSON.parse(jsonCandidate);
          } catch (lastE) {
            // One last try on the candidate with trailing comma fix
            jsonCandidate = jsonCandidate.replace(/,\s*([\]}])/g, '$1');
            return JSON.parse(jsonCandidate);
          }
        }

        throw new Error("No JSON structure found in response");
      }
    } catch (error) {
      console.error("JSON Extraction Error:", error, "Original Text (truncated):", text.slice(0, 200) + "...");
      return null;
    }
  }

  // ============================================
  // EXCEL DATA OPERATIONS
  // ============================================

  /**
   * Get selected range data from Excel
   */
  async getSelectedRange() {
    return await Excel.run(async (context) => {
      let range;
      try {
        range = context.workbook.getSelectedRange();
        range.load(['values', 'formulas', 'address', 'rowCount', 'columnCount']);
        await context.sync();
      } catch (e) {
        console.warn("[ExcelAIAssistant] Selection/UsedRange invalid, falling back to A1.");
        try {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          range = sheet.getUsedRange();
          range.load(['values', 'formulas', 'address', 'rowCount', 'columnCount']);
          await context.sync();
        } catch (usedRangeError) {
          // Final fallback for empty sheets
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          range = sheet.getRange("A1");
          range.load(['values', 'formulas', 'address', 'rowCount', 'columnCount']);
          await context.sync();
        }
      }

      return {
        address: range.address,
        values: range.values,
        formulas: range.formulas,
        rowCount: range.rowCount,
        columnCount: range.columnCount
      };
    });
  }

  /**
   * Write data to a specific range
   */
  async writeToRange(address, data) {
    console.log(`[ExcelAIAssistant] Writing to range ${address}`, data);
    if (!data || !data.length || !data[0].length) {
      console.warn("[ExcelAIAssistant] No data to write.");
      return { status: 'error', message: 'No data to write' };
    }

    return await Excel.run(async (context) => {
      // Use helper to resolve start address and find sheet
      const startRange = await this.getRangeFromString(context, address || 'A1');
      const sheet = startRange.worksheet;

      startRange.load(["rowIndex", "columnIndex"]);
      await context.sync();

      // STRICT DIMENSION CALCULATION
      // Ensure we only write to a range that matches the data exactly
      const rowCount = data.length;
      const colCount = data[0] ? data[0].length : 0;

      if (rowCount === 0 || colCount === 0) {
        console.warn("[ExcelAIAssistant] No data points to write.");
        return { status: 'error', message: 'No data to write' };
      }

      // Calculate final range based on ACTUAL data dimensions
      // Use sheet object from the startRange to ensure we are on the right worksheet
      const targetRange = sheet.getRangeByIndexes(
        startRange.rowIndex,
        startRange.columnIndex,
        rowCount,
        colCount
      );

      targetRange.values = data;
      targetRange.load('address');

      console.log(`[ExcelAIAssistant] Syncing writeToRange... Dim: ${rowCount}x${colCount}`);
      await context.sync();

      console.log(`[ExcelAIAssistant] writeToRange complete at ${targetRange.address}`);
      return { status: 'success', address: targetRange.address };
    });
  }

  /**
   * Helper: Get range from string (handles "Sheet1!A1:B10" or "A1:B10")
   */
  async getRangeFromString(context, rangeString) {
    if (!rangeString || typeof rangeString !== 'string' || rangeString.trim() === "") {
      throw new Error(`Invalid range string provided: ${rangeString}`);
    }

    const cleanRange = this.sanitizeRangeString(rangeString.trim());

    // 0. Handle "selection" keyword
    if (cleanRange.toLowerCase() === 'selection') {
      return context.workbook.getSelectedRange();
    }

    // 1. Try workbook-level getRange (Most robust: handles Sheets, Names, and Structured Refs)
    // Wrap in try-catch because it throws if range is not found or syntax is slightly off for workbook level
    try {
      // Check if the method exists (some older environments might not have it on workbook)
      if (typeof context.workbook.getRange === 'function') {
        const range = context.workbook.getRange(cleanRange);
        // We MUST return the range object directly to the Excel.run caller
        return range;
      }
    } catch (e) {
      console.log(`[ExcelAIAssistant] Workbook.getRange failed for "${cleanRange}", trying fallback...`);
    }

    // 2. Fallback: Manually handle Sheet!A1 format
    if (cleanRange.includes('!')) {
      // Handle cases where comma separation might have confused the caller (e.g. "Sheet1!A1,Sheet2!B2")
      // We only take the first part if it looks like a list
      if (cleanRange.includes(',') && cleanRange.indexOf(',') > cleanRange.indexOf('!')) {
        console.warn(`[ExcelAIAssistant] Range "${cleanRange}" looks like a list. Using first part.`);
        cleanRange = cleanRange.split(',')[0].trim();
      }

      let cleanSheetName = null;
      let address = null;

      // Strategy 1: Handle quoted sheet names ('Sheet Name'!A1)
      if (cleanRange.startsWith("'")) {
        const endQuoteIdx = cleanRange.indexOf("'!");
        if (endQuoteIdx !== -1) {
          cleanSheetName = cleanRange.substring(1, endQuoteIdx);
          address = cleanRange.substring(endQuoteIdx + 2);
        }
      }

      // Strategy 2: Handle unquoted sheet names (Sheet1!A1) - taking the FIRST '!' as separator
      if (!cleanSheetName) {
        const firstExclIdx = cleanRange.indexOf('!');
        if (firstExclIdx !== -1) {
          cleanSheetName = cleanRange.substring(0, firstExclIdx);
          address = cleanRange.substring(firstExclIdx + 1);
        }
      }

      // If we successfully parsed a sheet name, clean up the address
      if (cleanSheetName && address) {
        // Fix for "Sheet1!A1:Sheet1!B2" -> "A1:B2"
        // We remove occurrences of the sheet name from the address part to make it a local address
        const escapedSheet = cleanSheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        // Regex to remove "SheetName!" or "'SheetName'!"
        const pattern = new RegExp(`(?:'${escapedSheet}'|${escapedSheet})!`, 'g');
        address = address.replace(pattern, '');
      } else {
        // Fallback to legacy right-to-left split if something went wrong (unlikely)
        const parts = cleanRange.split('!');
        address = parts.pop();
        const sheetNamePart = parts.join('!');
        cleanSheetName = sheetNamePart.replace(/^'|'$/g, '');
      }

      const sheet = context.workbook.worksheets.getItemOrNullObject(cleanSheetName);
      await context.sync();

      if (sheet.isNullObject) {
        throw new Error(`Sheet "${cleanSheetName}" not found in the workbook.`);
      }
      return sheet.getRange(address);
    }

    // 3. Fallback: Table structured reference check (if workbook.getRange wasn't available)
    if (cleanRange.includes('[') && cleanRange.includes(']')) {
      const tableName = cleanRange.split('[')[0];
      const table = context.workbook.tables.getItemOrNullObject(tableName);
      await context.sync();

      if (!table.isNullObject) {
        // This is a simplification; full structured ref parsing is complex
        // But context.workbook.getRange should have handled this if available.
        return table.getRange();
      }
    }

    // 4. Final Fallback: Local address on active sheet
    return context.workbook.worksheets.getActiveWorksheet().getRange(cleanRange);
  }

  /**
   * Write data to selected range
   */
  async writeToSelectedRange(data) {
    console.log(`[ExcelAIAssistant] writeToSelectedRange started`, data);
    if (!data || !data.length || !data[0].length) {
      console.warn("[ExcelAIAssistant] No data to write.");
      return false;
    }

    return await Excel.run(async (context) => {
      const selectedRange = context.workbook.getSelectedRange();
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      selectedRange.load(["rowIndex", "columnIndex"]);
      console.log(`[ExcelAIAssistant] Loading range indices...`);
      await context.sync();

      const rowCount = data.length;
      const colCount = data[0].length;
      console.log(`[ExcelAIAssistant] Target dimensions: ${rowCount}x${colCount} at [${selectedRange.rowIndex}, ${selectedRange.columnIndex}]`);

      const targetRange = sheet.getRangeByIndexes(
        selectedRange.rowIndex,
        selectedRange.columnIndex,
        rowCount,
        colCount
      );

      targetRange.values = data;
      targetRange.load('address');
      console.log(`[ExcelAIAssistant] Syncing writeToSelectedRange...`);
      await context.sync();
      console.log(`[ExcelAIAssistant] writeToSelectedRange complete.`);
      return { status: 'success', address: targetRange.address };
    });
  }

  /**
   * Get entire worksheet data
   */
  async getWorksheetData(sheetName = null) {
    return await Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItemOrNullObject(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      if (sheetName) {
        await context.sync();
        if (sheet.isNullObject) {
          throw new Error(`Sheet "${sheetName}" not found.`);
        }
      }

      const usedRange = sheet.getUsedRangeOrNullObject();
      sheet.load('name');
      await context.sync();

      if (usedRange.isNullObject) {
        console.warn(`[ExcelAIAssistant] Worksheet "${sheet.name}" is empty.`);
        return {
          sheetName: sheet.name,
          address: "A1",
          values: [[]],
          rowCount: 0,
          columnCount: 0
        };
      }

      usedRange.load(['values', 'address', 'rowCount', 'columnCount']);
      await context.sync();

      return {
        sheetName: sheet.name,
        address: usedRange.address,
        values: usedRange.values,
        rowCount: usedRange.rowCount,
        columnCount: usedRange.columnCount
      };
    });
  }

  /**
   * Get specific range by address (e.g., "A1:C10")
   */
  async getRangeByAddress(address) {
    return await Excel.run(async (context) => {
      const range = await this.getRangeFromString(context, address);
      range.load(['values', 'formulas', 'address']);

      await context.sync();

      return {
        address: range.address,
        values: range.values,
        formulas: range.formulas
      };
    });
  }

  /**
   * Create a summarized analysis table (e.g., Average Price by Brand)
   * Uses dynamic formulas (UNIQUE, AVERAGEIF) for automatic updates
   */
  async createMetricTable(dataRange, categoryColumn, metricColumn, aggregation, targetCell) {
    console.log(`[ExcelAIAssistant] createMetricTable started: ${aggregation} of ${metricColumn} by ${categoryColumn} into ${targetCell}`);

    return await Excel.run(async (context) => {
      // 1. Resolve ranges
      const sourceRange = await this.getRangeFromString(context, dataRange);

      // Robust input handling for AI arrays
      if (Array.isArray(categoryColumn)) categoryColumn = categoryColumn[0];
      if (Array.isArray(metricColumn)) metricColumn = metricColumn[0];
      if (Array.isArray(aggregation)) aggregation = aggregation[0];

      const targetRange = await this.getRangeFromString(context, targetCell);
      const sheet = targetRange.worksheet;
      const sourceSheet = sourceRange.worksheet;

      sourceRange.load(["values", "address", "columnCount", "columnIndex"]);
      sheet.load("name");
      sourceSheet.load("name");
      await context.sync();

      const headers = sourceRange.values[0];

      // Helper for fuzzy column matching
      const findColumnIndex = (colName) => {
        if (!colName) return -1;

        // Convert to string and clean
        const originalColName = String(colName).trim();
        const normalize = (s) => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
        const target = normalize(originalColName);

        // 1. Try exact match (trimmed, case-insensitive)
        let idx = headers.findIndex(h => String(h).toLowerCase().trim() === originalColName.toLowerCase());
        if (idx !== -1) return idx;

        // 2. Try normalized match (remove special chars)
        idx = headers.findIndex(h => normalize(h) === target);
        if (idx !== -1) return idx;

        // 3. Handle Column Letters (e.g., "F", "Sheet1!F:F", "F:F")
        // Check if colName looks like a column letter/address
        let potentialColLetter = originalColName;
        if (potentialColLetter.includes('!')) potentialColLetter = potentialColLetter.split('!').pop();
        if (potentialColLetter.includes(':')) potentialColLetter = potentialColLetter.split(':')[0];

        // Simple regex for A, B, ... Z, AA, etc.
        if (/^[A-Z]+$/i.test(potentialColLetter)) {
          const colNum = this.letterToColumnIndex(potentialColLetter); // 0-indexed absolute
          // Convert absolute column number to relative index in sourceRange
          const relativeIdx = colNum - sourceRange.columnIndex;
          if (relativeIdx >= 0 && relativeIdx < sourceRange.columnCount) {
            console.log(`[ExcelAIAssistant] Resolved column letter "${originalColName}" to relative index ${relativeIdx}`);
            return relativeIdx;
          }
        }

        // 4. Try partial match
        if (target.length > 3) {
          idx = headers.findIndex(h => normalize(h).includes(target) || target.includes(normalize(h)));
          if (idx !== -1) {
            console.log(`[ExcelAIAssistant] Fuzzy matched column "${colName}" to "${headers[idx]}"`);
            return idx;
          }
        }

        return -1;
      };

      const catIdx = findColumnIndex(categoryColumn);
      const metIdx = findColumnIndex(metricColumn);

      if (catIdx === -1 || metIdx === -1) {
        throw new Error(`Could not find columns "${categoryColumn}" or "${metricColumn}" in source range. Available headers: ${JSON.stringify(headers)}`);
      }

      // Helper to get column letter from index (within the source range)
      const getSourceColLetter = (idx) => {
        // Need to find absolute column index
        const rangeObj = sourceRange;
        // This is tricky if sourceRange is not starting at A
        // A simpler way is to use address if possible, but let's calculate based on sourceRange.columnIndex
        return this.getColumnLetter(sourceRange.columnIndex + idx);
      };

      const catColLetter = getSourceColLetter(catIdx);
      const metColLetter = getSourceColLetter(metIdx);
      const rowCount = sourceRange.values.length;
      const sourceSheetRef = `'${sourceSheet.name}'!`;

      const catRangeRef = `${sourceSheetRef}${catColLetter}2:${catColLetter}${rowCount}`;
      const metRangeRef = `${sourceSheetRef}${metColLetter}2:${metColLetter}${rowCount}`;

      // 2. Insert Table Headers
      // We must resize targetRange to match the 1x2 header array [[Col1, Col2]]
      const headerRange = targetRange.getResizedRange(0, 1);
      headerRange.values = [[categoryColumn, `${aggregation} of ${metricColumn}`]];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#DDEBF7";

      // 3. Insert UNIQUE Categories Formula
      // targetRange is the top-left cell. uniqueCell is one row below it.
      const uniqueCell = targetRange.getOffsetRange(1, 0);
      uniqueCell.formulas = [[`=SORT(UNIQUE(${catRangeRef}))`]];

      // 4. Insert Aggregation Formula (Spill referencing)
      const aggCell = targetRange.getOffsetRange(1, 1);
      targetRange.load("address");
      uniqueCell.load("address");
      aggCell.load("address");
      await context.sync();

      const uniqueRef = uniqueCell.address.split('!').pop() + "#";

      let aggFormula;
      switch (aggregation.toLowerCase()) {
        case 'average':
          aggFormula = `=IFERROR(AVERAGEIF(${catRangeRef}, ${uniqueRef}, ${metRangeRef}), 0)`;
          break;
        case 'sum':
          aggFormula = `=SUMIF(${catRangeRef}, ${uniqueRef}, ${metRangeRef})`;
          break;
        case 'count':
          aggFormula = `=COUNTIF(${catRangeRef}, ${uniqueRef})`;
          break;
        case 'max':
          aggFormula = `=MAXIFS(${metRangeRef}, ${catRangeRef}, ${uniqueRef})`;
          break;
        case 'min':
          aggFormula = `=MINIFS(${metRangeRef}, ${catRangeRef}, ${uniqueRef})`;
          break;
        default:
          aggFormula = `=COUNTIF(${catRangeRef}, ${uniqueRef})`;
      }

      aggCell.formulas = [[aggFormula]];

      await context.sync();

      // 5. Final Formatting
      // We don't know the exact size of the spilled range yet, let's auto-fit the columns
      targetRange.getResizedRange(0, 1).getEntireColumn().format.autofitColumns();

      const fullTableAddress = `${targetRange.address}:${aggCell.address}#`;
      console.log(`[ExcelAIAssistant] createMetricTable complete at ${fullTableAddress}`);

      return {
        status: 'success',
        address: fullTableAddress,
        sheetName: sheet.name
      };
    });
  }

  /**
   * Insert columns at a specific location
   */
  async insertColumns(address, count) {
    console.log(`[ExcelAIAssistant] insertColumns started: ${count} at ${address}`);
    return await Excel.run(async (context) => {
      const range = await this.getRangeFromString(context, address);
      const entireColumn = range.getEntireColumn();

      // Office.js: insert(shift) on column range
      for (let i = 0; i < count; i++) {
        entireColumn.insert(Excel.InsertShiftDirection.right);
      }

      await context.sync();
      return { status: 'success', insertedCount: count };
    });
  }

  async insertRows(address, count) {
    console.log(`[ExcelAIAssistant] insertRows started: ${count} at ${address}`);
    return await Excel.run(async (context) => {
      const range = await this.getRangeFromString(context, address);
      const entireRow = range.getEntireRow();

      for (let i = 0; i < count; i++) {
        entireRow.insert(Excel.InsertShiftDirection.down);
      }

      await context.sync();
      return { status: 'success', insertedCount: count };
    });
  }

  /**
   * Helper to get column letter from number (0-indexed)
   */
  getColumnLetter(colIndex) {
    let letter = "";
    while (colIndex >= 0) {
      letter = String.fromCharCode((colIndex % 26) + 65) + letter;
      colIndex = Math.floor(colIndex / 26) - 1;
    }
    return letter;
  }

  /**
   * Helper to get column number from letter (e.g., "A" -> 0, "Z" -> 25, "AA" -> 26)
   */
  letterToColumnIndex(letter) {
    if (!letter) return 0;
    const cleanLetter = String(letter).toUpperCase().replace(/[^A-Z]/g, '');
    let colIndex = 0;
    for (let i = 0; i < cleanLetter.length; i++) {
      colIndex = colIndex * 26 + (cleanLetter.charCodeAt(i) - 64);
    }
    return colIndex - 1;
  }

  /**
   * Sanitizes a range string to fix common malformations like "A1:B" -> "A1:B1048576"
   */
  sanitizeRangeString(rangeString) {
    if (!rangeString || typeof rangeString !== 'string') return rangeString;

    let address = rangeString.trim();

    // Fix "Sheet1!K2:L" -> "Sheet1!K2:L1048576" (missing end row in range)
    // This matches a column letter followed by no number after a colon
    const malformedRangePattern = /([A-Z]+[0-9]+):([A-Z]+)(?![0-9])/gi;
    if (malformedRangePattern.test(address)) {
      console.log(`[ExcelAIAssistant] Sanitizing malformed range: ${address}`);
      address = address.replace(malformedRangePattern, (match, start, end) => {
        return `${start}:${end}1048576`;
      });
    }

    // Fix "A:A10" -> "A1:A10" (missing start row)
    const missingStartRowPattern = /(?<![0-9])([A-Z]+):([A-Z]+[0-9]+)/gi;
    if (missingStartRowPattern.test(address)) {
      address = address.replace(missingStartRowPattern, (match, startCol, end) => {
        return `${startCol}1:${end}`;
      });
    }

    return address;
  }

  /**
   * Insert formula into a cell or range
   */
  async insertFormula(address, formula) {
    console.log(`[ExcelAIAssistant] insertFormula started: ${formula} into ${address}`);

    // Aggressive sanitization: strip ALL markdown, control characters, and non-formula text
    if (!formula) formula = "";
    let cleanFormula = String(formula)
      .replace(/```[a-z]*\n?/gi, '') // Remove opening code blocks
      .replace(/\n?```/g, '')        // Remove closing code blocks
      .replace(/[\u0000-\u001F\u007F-\u009F]/g, "") // Remove control characters
      .trim();

    // If AI added extra explanation text before or after the formula, 
    // try to isolate the part that starts with =
    if (cleanFormula.includes('=')) {
      const index = cleanFormula.indexOf('=');
      // If there's content before '=', it's likely explanation text
      // But we should be careful not to cut off legitimate parts of a string if it starts with =
      // Usually, Excel formulas start immediately with = or after some markdown/text.
      cleanFormula = cleanFormula.substring(index);

      // Find the end of the formula (until last closing brace or line end)
      // This is still a bit of a heuristic but safer than the previous regex
      const lastBrace = cleanFormula.lastIndexOf(')');
      if (lastBrace !== -1 && lastBrace < cleanFormula.length - 1) {
        // Check if there's meaningful text after the last brace
        const remainder = cleanFormula.substring(lastBrace + 1).trim();
        if (remainder.length > 0 && !remainder.match(/^[0-9.]+$/)) {
          cleanFormula = cleanFormula.substring(0, lastBrace + 1);
        }
      }
    }

    if (!address || typeof address !== 'string') {
      console.error("[ExcelAIAssistant] Invalid address for insertFormula:", address);
      throw new Error(`Invalid range address: ${address}`);
    }

    return await Excel.run(async (context) => {
      let range;

      try {
        // Use helper to resolve range string (handles "Sheet!A1")
        range = await this.getRangeFromString(context, address);

        // If address is a whole column (e.g., "O:O"), intersect it with Used Range 
        if (address.includes(':') && !address.match(/\d/)) {
          console.log(`[ExcelAIAssistant] Processing whole column range: ${address}`);
          const sheet = range.worksheet; // range already has the right sheet
          const usedRange = sheet.getUsedRange();
          range = range.getIntersection(usedRange);
        }

        // Pre-fetch range to verify it exists before setting formula
        range.load("address");
        await context.sync();
      } catch (rangeError) {
        console.error(`[ExcelAIAssistant] Range error for ${address}:`, rangeError);
        throw new Error(`Excel could not find or access the range "${address}". Please ensure the address is valid (e.g., "A1" or "B2:C10").`);
      }

      const finalFormula = cleanFormula.startsWith('=') ? cleanFormula : '=' + cleanFormula;

      try {
        // Office.js: Setting formulas to a single string applies it to the whole range
        range.formulas = finalFormula;
        await context.sync();
      } catch (formulaError) {
        console.error(`[ExcelAIAssistant] Formula syntax error: ${finalFormula}`, formulaError);
        throw new Error(`Excel rejected the formula "${finalFormula}". Error: ${formulaError.message}`);
      }

      console.log(`[ExcelAIAssistant] insertFormula complete.`);
      return true;
    });
  }

  /**
   * Get all sheet names
   */
  async getAllSheetNames() {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load('items/name');

      await context.sync();

      return sheets.items.map(sheet => sheet.name);
    });
  }

  /**
   * Create a new worksheet
   */
  async createNewSheet(sheetName) {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.add(sheetName);
      sheet.activate();

      await context.sync();
      return sheetName;
    });
  }

  /**
   * Format range (color, bold, etc.)
   */
  async formatRange(address, formatting) {
    return await Excel.run(async (context) => {
      const range = await this.getRangeFromString(context, address);

      if (formatting.bold) range.format.font.bold = true;
      if (formatting.italic) range.format.font.italic = true;
      if (formatting.fontSize) range.format.font.size = formatting.fontSize;
      if (formatting.backgroundColor) range.format.fill.color = formatting.backgroundColor;
      if (formatting.fontColor) range.format.font.color = formatting.fontColor;

      await context.sync();
      return true;
    });
  }

  // ============================================
  // AI-POWERED EXCEL OPERATIONS
  // ============================================

  /**
   * Analyze data with Claude (defaults to selection if no address provided)
   */
  async analyzeData(address = null) {
    const rangeData = address ? await this.getRangeByAddress(address) : await this.getSelectedRange();
    const dataStr = this.formatDataForAI(rangeData.values, rangeData.address);

    const prompt = `Analyze this Excel data and provide insights:\n\n${dataStr}\n\nProvide a concise analysis including patterns, trends, or notable observations.`;

    return await this.callClaudeAPI(prompt);
  }

  /**
   * Generate Excel formula based on description
   */
  async generateFormula(description, contextData = null, address = null) {
    let prompt = `Generate an Excel formula for: ${description}\n\nReturn ONLY the formula, starting with =`;

    if (contextData) {
      const dataStr = this.formatDataForAI(contextData, address);
      prompt += `\n\nContext data (with row/column labels for context):\n${dataStr}`;
    }

    const systemPrompt = "You are an Excel formula expert. Return only the formula without explanation, starting with =";

    return await this.callClaudeAPI(prompt, systemPrompt);
  }

  /**
   * Clean and transform data using AI
   */
  async cleanData(instructions, address = null) {
    const rangeData = address ? await this.getRangeByAddress(address) : await this.getSelectedRange();
    const dataStr = this.formatDataForAI(rangeData.values, rangeData.address);
    const prompt = `Transform this data according to: ${instructions}

Original data (with row/column labels for context):
${dataStr}

IMPORTANT: 
1. Return ONLY the transformed values in the same tabular structure.
2. DO NOT include the row numbers or column letters in your response.
3. Use tabs for columns and newlines for rows.`;

    const systemPrompt = "Return only the cleaned data values as tab-separated values. No labels, no headers, no explanations.";

    const result = await this.callClaudeAPI(prompt, systemPrompt);

    // Parse result back to 2D array
    const cleanedData = this.parseAIDataResponse(result);

    return cleanedData;
  }

  /**
   * Apply cleaned data back to Excel
   */
  async applyCleanedData(instructions, address = null) {
    const cleanedData = await this.cleanData(instructions, address);

    // If we have a specific address, write back to it, otherwise write to selection
    if (address) {
      return await this.writeToRange(address, cleanedData);
    } else {
      return await this.writeToSelectedRange(cleanedData);
    }
  }

  /**
   * Generate summary table from data
   */
  async generateSummary(dataRange = null) {
    if (this.onStatusUpdate) this.onStatusUpdate("📊 Analyzing data for summary...");
    const rangeData = dataRange ? await this.getRangeByAddress(dataRange) : await this.getWorksheetData();
    const dataStr = this.formatDataForAI(rangeData.values, rangeData.address);

    const prompt = `Create a summary table from this data:\n\n${dataStr}\n\nReturn a concise summary table with key metrics as a 2D JSON array: [["Metric", "Value"], ["Total", 100], ...]`;

    const response = await this.callClaudeAPI(prompt, "Return only valid JSON.");
    const cleanResponse = response.replace(/```json\n?|\n?```/g, '').trim();
    const summaryData = JSON.parse(cleanResponse);

    if (this.onStatusUpdate) this.onStatusUpdate("📑 Creating summary output sheet...");

    // Write the summary to a new sheet
    const sheetName = `Summary_${new Date().getTime().toString().slice(-4)}`;
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.add(sheetName);
      sheet.activate();
      await context.sync();
    });

    if (summaryData && Array.isArray(summaryData)) {
      const writeResult = await this.writeToRange(`${sheetName}!A1`, summaryData);
      // Add address to result
      const colCount = summaryData[0] ? summaryData[0].length : 0;
      const endCol = this.columnIndexToLetter(colCount - 1);
      return {
        summary: summaryData,
        address: `${sheetName}!A1:${endCol}${summaryData.length}`
      };
    }

    return summaryData;
  }

  /**
   * Answer questions about the data
   */
  async askAboutData(question, address = null) {
    const rangeData = address ? await this.getRangeByAddress(address) : await this.getWorksheetData();
    const dataStr = this.formatDataForAI(rangeData.values, rangeData.address);

    const prompt = `Based on this Excel data:\n\n${dataStr}\n\nQuestion: ${question}`;

    return await this.callClaudeAPI(prompt);
  }

  /**
   * Generate chart recommendations
   */
  async recommendChart(address = null) {
    const rangeData = address ? await this.getRangeByAddress(address) : await this.getSelectedRange();
    const dataStr = this.formatDataForAI(rangeData.values, rangeData.address);

    const prompt = `Given this data:\n\n${dataStr}\n\nRecommend the best chart type and explain why.`;

    return await this.callClaudeAPI(prompt);
  }

  /**
   * Automatically generate chart from selected data with AI recommendations
   */
  async autoGenerateChart(address = null) {
    const rangeData = address ? await this.getRangeByAddress(address) : await this.getSelectedRange();
    const dataStr = this.formatDataForAI(rangeData.values, rangeData.address);

    // Ask AI for chart recommendations
    const prompt = `Analyze this data and recommend the best chart type:\n\n${dataStr}\n\nRespond in JSON format:
{
  "chartType": "one of: ColumnClustered, ColumnStacked, BarClustered, Line, LineMarkers, Pie, Area, XYScatter, Combo",
  "title": "suggested chart title",
  "xAxisTitle": "x-axis label",
  "yAxisTitle": "y-axis label",
  "reasoning": "brief explanation"
}`;

    const systemPrompt = "You are an expert data visualization specialist. Return only valid JSON.";

    const aiResponse = await this.callClaudeAPI(prompt, systemPrompt);

    // Parse AI recommendation
    let recommendation;
    try {
      // Remove markdown code blocks if present
      const cleanResponse = aiResponse.replace(/```json\n?|\n?```/g, '').trim();
      recommendation = JSON.parse(cleanResponse);
    } catch (error) {
      // Fallback to column chart if parsing fails
      recommendation = {
        chartType: "ColumnClustered",
        title: "Data Visualization",
        xAxisTitle: "Category",
        yAxisTitle: "Value"
      };
    }

    // Create the chart
    const chartResult = await this.createChart(
      rangeData.address,
      recommendation.chartType,
      recommendation.title,
      recommendation.xAxisTitle,
      recommendation.yAxisTitle
    );

    return {
      chartId: chartResult.chartId,
      recommendation: recommendation
    };
  }

  /**
   * Create a chart from a data range
   */
  async createChart(sourceRange, chartType, title, xAxisTitle = "", yAxisTitle = "") {
    return await Excel.run(async (context) => {
      // Support absolute addresses with sheet names or local ones
      // 1. Resolve range
      let cleanSource = sourceRange;
      let isSpilled = false;
      if (typeof cleanSource === 'string' && cleanSource.endsWith('#')) {
        isSpilled = true;
        cleanSource = cleanSource.slice(0, -1);
      }

      const range = await this.getRangeFromString(context, cleanSource);

      // Load and sync to ensure range is valid and we have dimensions
      range.load(["address", "rowCount", "columnCount"]);
      await context.sync();

      // Handle spilled range expansion
      let plotRange = range;
      if (isSpilled) {
        plotRange = range.getSpillingToParent(); // Or similar if available, or just use the range with #
        // Note: In Office.js, if you pass a range to charts.add, it should handle it.
        // If range is a single cell and we know it spills, we might need more logic.
      }

      // Safety check: ensure we have data to plot
      if (range.rowCount < 1) {
        throw new Error(`The range "${sourceRange}" contains no data to plot.`);
      }

      // CRITICAL: Get the worksheet where the data actually resides
      const sheet = range.worksheet;
      sheet.load("name");
      await context.sync();

      console.log(`[ExcelAIAssistant] Creating ${chartType} chart on sheet "${sheet.name}" using range ${range.address}`);

      // Map chart type names to Excel chart types
      const type = (chartType || "").toLowerCase().trim();
      const chartTypeMap = {
        "column": Excel.ChartType.columnClustered,
        "columnclustered": Excel.ChartType.columnClustered,
        "columnstacked": Excel.ChartType.columnStacked,
        "bar": Excel.ChartType.barClustered,
        "barclustered": Excel.ChartType.barClustered,
        "barstacked": Excel.ChartType.barStacked,
        "line": Excel.ChartType.line,
        "linemarkers": Excel.ChartType.lineMarkers,
        "pie": Excel.ChartType.pie,
        "area": Excel.ChartType.area,
        "areastacked": Excel.ChartType.areaStacked,
        "xyscatter": Excel.ChartType.xyscatter,
        "scatter": Excel.ChartType.xyscatter,
        "combo": Excel.ChartType.columnClustered
      };

      const excelChartType = chartTypeMap[type] || Excel.ChartType.columnClustered;

      // Add the chart
      console.log(`[ExcelAIAssistant] Adding chart with type ${excelChartType} for range ${sourceRange}`);

      let chart;
      try {
        // Try adding with auto series
        chart = sheet.charts.add(excelChartType, isSpilled ? `${range.address}#` : range, Excel.ChartSeriesBy.auto);
      } catch (e) {
        console.warn(`[ExcelAIAssistant] Chart creation failed with "auto", trying "columns"...`, e);
        try {
          chart = sheet.charts.add(excelChartType, isSpilled ? `${range.address}#` : range, Excel.ChartSeriesBy.columns);
        } catch (innerE) {
          console.error(`[ExcelAIAssistant] Final chart creation failure:`, innerE);
          throw innerE;
        }
      }

      chart.title.text = title || "Chart";
      chart.title.visible = true;

      // Set axis titles if provided
      if (xAxisTitle) {
        chart.axes.valueAxis.title.text = yAxisTitle; // Usually value axis is Y
        chart.axes.valueAxis.title.visible = true;
        chart.axes.categoryAxis.title.text = xAxisTitle; // Category axis is X
        chart.axes.categoryAxis.title.visible = true;
      }

      // Position: place it 2 rows below the data, same width
      // We need to fetch range dimensions to do this nicely
      const chartTop = range.getOffsetRange(range.rowCount + 2, 0);
      chart.setPosition(chartTop, chartTop.getOffsetRange(15, 8)); // Approx 15 rows high, 8 cols wide

      // Force data binding reset to ensure dynamic arrays are picked up
      try {
        chart.setData(isSpilled ? `${range.address}#` : range, Excel.ChartSeriesBy.auto);
      } catch (dataError) {
        console.warn(`[ExcelAIAssistant] chart.setData failed, continuing with initial binding.`);
      }
      chart.dataLabels.showValue = true; // Ensure values are visible

      await context.sync();

      return { status: 'success', chartId: chart.id, sheetName: sheet.name };
    });
  }

  /**
   * Move a chart to a specific location
   */
  async moveChart(chartName, targetCell) {
    console.log(`[ExcelAIAssistant] moveChart: moving "${chartName || 'active'}" to ${targetCell}`);

    return await Excel.run(async (context) => {
      const targetRange = await this.getRangeFromString(context, targetCell);
      const targetSheet = targetRange.worksheet;

      // Load target info
      targetRange.load(["top", "left"]);
      targetSheet.load(["name", "charts"]);
      await context.sync();

      let chart;

      // If we are moving to a different sheet, we actually need to cut/paste or recreate
      // But Office.js charts are bound to a sheet.
      // Moving across sheets is hard: we have to copy the chart or move it.
      // However, usually "moveChart" in this context implies positioning it on the DASHBOARD sheet.

      // Strategy: Look for the chart on the ACTIVE sheet or Source Data sheet?
      // Since we don't know where the chart IS, we might have to search all sheets or assume active.
      // Simplification: Assume chart is on the active sheet OR on the sheet of the targetCell?
      // Actually, if we just created it, it's on the data sheet. We want it on the Dashboard sheet.

      // Find the chart. If chartName is not provided, take the last added chart?
      // Let's search all sheets for the chart if name provided.
      // Or if chartName is not provided, use the last chart on the ACTIVE sheet.

      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load(["charts", "name"]);
      await context.sync();

      if (activeSheet.charts.count > 0) {
        // Get the last chart (most recently created)
        chart = activeSheet.charts.getItemAt(activeSheet.charts.count - 1);
      } else {
        // Try finding by name across all sheets
        // This is expensive, so let's skip for now and throw error if not found
        throw new Error("No charts found on active sheet to move.");
      }

      // Check if target is on a different sheet
      if (targetSheet.name !== activeSheet.name) {
        // Moving chart to another sheet is tricky in Office.js (no direct .move()).
        // We have to clone it or cut/paste.
        // Workaround: Copy chart as image or recreate it?
        // BETTER: Just update the `createChart` to accept a targetSheet parameter?
        // For now, let's just log a warning and move it within the same sheet if possible.

        console.warn("Moving charts between sheets is limited. Creating new chart reference.");

        // Alternative: If the user wants a dashboard, we should have created the chart ON the dashboard sheet 
        // referencing the data sheet.

        // Let's try to Re-Create the chart on the target sheet using the original data source
        chart.load(["series", "chartType", "title", "height", "width"]);
        await context.sync();

        // We need the data source. This is hard to extract perfectly.
        // FALLBACK: Just tell the user we positioned it on the source sheet.
        throw new Error("Moving existing charts between sheets is not fully supported. Please specify the target sheet when creating the chart.");
      }

      // Same sheet move
      chart.top = targetRange.top;
      chart.left = targetRange.left;

      await context.sync();
      return { status: "success", message: "Chart moved." };
    });
  }

  /**
   * Create multiple charts based on AI analysis
   */
  async generateMultipleCharts(dataRange = null) {
    let worksheetData;
    if (dataRange) {
      worksheetData = await Excel.run(async (context) => {
        const range = await this.getRangeFromString(context, dataRange);
        range.load(['values', 'address']);
        await context.sync();
        return { values: range.values, address: range.address };
      });
    } else {
      worksheetData = await this.getWorksheetData();
    }

    const dataStr = this.formatDataForAI(worksheetData.values, worksheetData.address);

    const prompt = `You are a Lead Data Visualization Architect. Perform a high-fidelity analysis of this dataset and architect an "Elite Executive Dashboard".

DATA CONTEXT:
${dataStr}

DASHBOARD REQUIREMENTS:
1. **High Volume of Insights**: Suggest 5-8 distinct, high-impact visualizations. Do not settle for less than 5.
2. **Contextual Diversity**: Include a mix of Line (trends), Column (comparison), Pie (composition), and Scatter (correlation) where appropriate.
3. **Professional Naming**: Use evocative, corporate-grade chart titles (e.g., "Q4 Revenue Momentum Analysis" instead of "Sales Chart").
4. **Insight Rationale**: Every chart must have a clear "insight" property explaining the business truth it reveals.

RETURN JSON FORMAT (MANDATORY):
{
  "charts": [
    {
      "dataRange": "${worksheetData.address}",
      "chartType": "ColumnClustered",
      "title": "Elite Professional Title",
      "xAxisTitle": "Contextual Label",
      "yAxisTitle": "Metric Label",
      "insight": "Strategic revelation of this chart"
    }
  ]
}`;

    const systemPrompt = "Return only valid JSON with 5-8 professional chart recommendations. No explanations.";

    const aiResponse = await this.callClaudeAPI(prompt, systemPrompt);

    // Parse response using robust repair logic
    let recommendations;
    try {
      recommendations = this.extractJSON(aiResponse);
      if (!recommendations || !recommendations.charts) {
        throw new Error("Invalid structure in chart recommendations");
      }
    } catch (error) {
      console.error("[ExcelAIAssistant] Dashboard JSON parse failed:", error);
      throw new Error("Failed to parse chart recommendations");
    }

    // Create all recommended charts
    const createdCharts = [];
    for (const chartRec of recommendations.charts) {
      const result = await this.createChart(
        chartRec.dataRange,
        chartRec.chartType,
        chartRec.title,
        chartRec.xAxisTitle,
        chartRec.yAxisTitle
      );

      createdCharts.push({
        ...result,
        insight: chartRec.insight
      });
    }

    return {
      charts: createdCharts,
      recommendations: recommendations
    };
  }

  /**
   * Create a specific chart type manually
   */
  async createSpecificChart(chartType, title, xAxisTitle = "", yAxisTitle = "") {
    const rangeData = await this.getSelectedRange();
    return await this.createChart(
      rangeData.address,
      chartType,
      title,
      xAxisTitle,
      yAxisTitle
    );
  }

  /**
   * Update existing chart with new data
   */
  async updateChart(chartId, newDataRange) {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const chart = sheet.charts.getItemOrNullObject(chartId);
      await context.sync();

      if (chart.isNullObject) {
        throw new Error(`Chart with ID "${chartId}" not found.`);
      }

      // Update data range
      const range = await this.getRangeFromString(context, newDataRange);
      chart.setData(range, Excel.ChartSeriesBy.auto);

      await context.sync();
      return true;
    });
  }

  /**
   * Delete a chart
   */
  async deleteChart(chartId) {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const chart = sheet.charts.getItemOrNullObject(chartId);
      await context.sync();

      if (chart.isNullObject) {
        console.warn(`[ExcelAIAssistant] Chart "${chartId}" already deleted or not found.`);
        return true;
      }

      chart.delete();
      await context.sync();
      return true;
    });
  }

  /**
   * Get all charts in worksheet
   */
  async getAllCharts() {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const charts = sheet.charts;
      charts.load('items/name, items/id');

      await context.sync();

      return charts.items.map(chart => ({
        id: chart.id,
        name: chart.name
      }));
    });
  }

  /**
   * Style chart with custom colors
   */
  async styleChart(chartId, styling) {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const chart = sheet.charts.getItem(chartId);

      // Apply styling
      if (styling.titleFontSize) {
        chart.title.format.font.size = styling.titleFontSize;
      }

      if (styling.titleColor) {
        chart.title.format.font.color = styling.titleColor;
      }

      if (styling.legendPosition) {
        chart.legend.position = styling.legendPosition; // "Top", "Bottom", "Left", "Right"
      }

      if (styling.backgroundColor) {
        chart.format.fill.setSolidColor(styling.backgroundColor);
      }

      await context.sync();
      return true;
    });
  }

  /**
   * Fill down pattern with AI
   */
  async fillPattern(examples, targetCount) {
    const dataStr = this.formatDataForAI(examples);

    const prompt = `Continue this pattern for ${targetCount} more rows:\n\n${dataStr}\n\nReturn only the new rows in tabular format.`;

    const systemPrompt = "Return only the data continuation, no explanations.";

    const result = await this.callClaudeAPI(prompt, systemPrompt);

    return this.parseAIDataResponse(result);
  }

  // ============================================
  // HELPER FUNCTIONS
  // ============================================

  /**
   * Format Excel data for AI prompt
   */
  formatDataForAI(data, address = null) {
    if (!data) return "";
    if (!Array.isArray(data)) return String(data);

    let startColIndex = 0;
    let startRowIndex = 1;

    if (address) {
      // Parse address like "Sheet1!A1:C10" or "A1:G100"
      const rangePart = address.includes('!') ? address.split('!')[1] : address;
      const startCell = rangePart.split(':')[0];
      const colMatch = startCell.match(/[A-Z]+/);
      const rowMatch = startCell.match(/\d+/);

      if (colMatch) startColIndex = this.columnLetterToIndex(colMatch[0]);
      if (rowMatch) startRowIndex = parseInt(rowMatch[0]);
    }

    if (data.length === 0) return "";

    // If it's a 2D array, add headers
    if (Array.isArray(data[0])) {
      // Build Header (Column Letters)
      let header = "      \t";
      for (let j = 0; j < data[0].length; j++) {
        header += this.columnIndexToLetter(startColIndex + j) + "\t";
      }

      let rows = [header.trimEnd()];

      data.forEach((row, i) => {
        const rowNum = String(startRowIndex + i).padStart(5, ' ');
        const rowData = row.map(cell => (cell === null || cell === undefined) ? "" : String(cell)).join('\t');
        rows.push(`${rowNum}\t${rowData}`);
      });

      return rows.join('\n');
    }

    // Fallback for 1D array
    return data.join('\t');
  }

  /**
   * Parse AI response back to 2D array
   */
  parseAIDataResponse(response) {
    if (!response) return [[]];

    // Split into lines, filter out empty lines
    const lines = response.trim().split('\n').filter(line => line.trim() !== '');

    // Split each line by tabs or commas (AI might use either if it's hallucinating CSV)
    // and trim each cell
    const rawData = lines.map(line => {
      // Try tab first, then comma if no tabs found
      const delimiter = line.includes('\t') ? '\t' : (line.includes(',') ? ',' : '\t');
      return line.split(delimiter).map(cell => cell.trim());
    });

    // Ensure it's a perfect rectangle (Excel requirement)
    // 1. Find max columns
    const maxCols = Math.max(...rawData.map(row => row.length));

    // 2. Pad rows that are too short AND trim rows that are too long (ensure exact rectangle)
    return rawData.map(row => {
      if (row.length < maxCols) {
        while (row.length < maxCols) row.push("");
      } else if (row.length > maxCols) {
        return row.slice(0, maxCols);
      }
      return row;
    });
  }

  /**
   * Convert column letter to index (A=0, B=1, etc.)
   */
  columnLetterToIndex(letter) {
    let index = 0;
    for (let i = 0; i < letter.length; i++) {
      index = index * 26 + (letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return index - 1;
  }

  /**
   * Convert column index to letter (0=A, 1=B, etc.)
   */
  columnIndexToLetter(index) {
    let letter = '';
    while (index >= 0) {
      letter = String.fromCharCode((index % 26) + 'A'.charCodeAt(0)) + letter;
      index = Math.floor(index / 26) - 1;
    }
    return letter;
  }
}

// ============================================
// USAGE EXAMPLE
// ============================================

// Initialize the assistant
let assistant;

function initializeAssistant(apiKey) {
  assistant = new ExcelAIAssistant(apiKey);
  console.log('Excel AI Assistant initialized');
}

// Example functions to call from UI
async function analyzeCurrentSelection() {
  try {
    const analysis = await assistant.analyzeData();
    console.log('Analysis:', analysis);
    return analysis;
  } catch (error) {
    console.error('Error:', error);
    throw error;
  }
}

async function generateFormulaFromDescription(description) {
  try {
    const formula = await assistant.generateFormula(description);
    console.log('Generated formula:', formula);

    // Optionally insert into selected cell
    const range = await assistant.getSelectedRange();
    await assistant.insertFormula(range.address, formula);

    return formula;
  } catch (error) {
    console.error('Error:', error);
    throw error;
  }
}

async function cleanSelectedData(instructions) {
  try {
    await assistant.applyCleanedData(instructions);
    console.log('Data cleaned successfully');
    return true;
  } catch (error) {
    console.error('Error:', error);
    throw error;
  }
}

async function askQuestion(question) {
  try {
    const answer = await assistant.askAboutData(question);
    console.log('Answer:', answer);
    return answer;
  } catch (error) {
    console.error('Error:', error);
    throw error;
  }
}

// ============================================
// CHART GENERATION EXAMPLES
// ============================================

async function autoCreateChart() {
  try {
    // AI analyzes data and creates the best chart automatically
    const result = await assistant.autoGenerateChart();
    console.log('Chart created:', result);
    console.log('AI Reasoning:', result.recommendation.reasoning);
    return result;
  } catch (error) {
    console.error('Error:', error);
    throw error;
  }
}

async function createColumnChart(title = "Data Visualization") {
  try {
    const result = await assistant.createSpecificChart(
      "ColumnClustered",
      title,
      "Categories",
      "Values"
    );
    console.log('Column chart created:', result);
    return result;
  } catch (error) {
    console.error('Error:', error);
    throw error;
  }
}

async function createLineChart(title = "Trend Analysis") {
  try {
    const result = await assistant.createSpecificChart(
      "LineMarkers",
      title,
      "Time Period",
      "Value"
    );
    console.log('Line chart created:', result);
    return result;
  } catch (error) {
    console.error('Error:', error);
    throw error;
  }
}

async function createPieChart(title = "Distribution") {
  try {
    const result = await assistant.createSpecificChart(
      "Pie",
      title
    );
    console.log('Pie chart created:', result);
    return result;
  } catch (error) {
    console.error('Error:', error);
    throw error;
  }
}

async function generateDashboard() {
  try {
    // Creates multiple charts based on AI analysis
    const result = await assistant.generateMultipleCharts();
    console.log('Dashboard created with', result.charts.length, 'charts');
    result.charts.forEach((chart, i) => {
      console.log(`Chart ${i + 1} insight:`, chart.insight);
    });
    return result;
  } catch (error) {
    console.error('Error:', error);
    throw error;
  }
}

async function customStyleChart(chartId) {
  try {
    await assistant.styleChart(chartId, {
      titleFontSize: 16,
      titleColor: "#2E75B6",
      legendPosition: "Bottom",
      backgroundColor: "#F0F0F0"
    });
    console.log('Chart styled successfully');
    return true;
  } catch (error) {
    console.error('Error:', error);
    throw error;
  }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { ExcelAIAssistant };
}

// Ensure global availability
if (typeof window !== 'undefined') {
  window.ExcelAIAssistant = ExcelAIAssistant;
}