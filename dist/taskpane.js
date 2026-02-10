/******/ (function() { // webpackBootstrap
/******/ 	var __webpack_modules__ = ({

/***/ "./src/taskpane/taskpane.html":
/*!************************************!*\
  !*** ./src/taskpane/taskpane.html ***!
  \************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
// Module
var code = "<!DOCTYPE html>\n<html lang=\"en\">\n\n<head>\n    <meta charset=\"UTF-8\">\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n    <title>AI Dashboard Generator</title>\n\n    <!-- Load Office.js FIRST and ONLY ONCE -->\n    <" + "script src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\"><" + "/script>\n\n    <style>\n        /* Your CSS here - same as before */\n        * {\n            margin: 0;\n            padding: 0;\n            box-sizing: border-box;\n        }\n\n        body {\n            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;\n            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);\n            padding: 20px;\n            min-height: 100vh;\n        }\n\n        .container {\n            max-width: 900px;\n            margin: 0 auto;\n            background: white;\n            border-radius: 15px;\n            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);\n            overflow: hidden;\n        }\n\n        .header {\n            background: linear-gradient(135deg, #2E75B6 0%, #1F4E78 100%);\n            color: white;\n            padding: 30px;\n            text-align: center;\n        }\n\n        .header h1 {\n            font-size: 28px;\n            margin-bottom: 10px;\n        }\n\n        .content {\n            padding: 30px;\n        }\n\n        .setup-section {\n            background: #f8f9fa;\n            padding: 20px;\n            border-radius: 10px;\n            margin-bottom: 30px;\n        }\n\n        .input-group {\n            margin-bottom: 15px;\n        }\n\n        .input-group label {\n            display: block;\n            margin-bottom: 5px;\n            font-weight: 600;\n            color: #333;\n        }\n\n        .input-group input {\n            width: 100%;\n            padding: 12px;\n            border: 2px solid #e0e0e0;\n            border-radius: 8px;\n            font-size: 14px;\n        }\n\n        .btn {\n            padding: 12px 24px;\n            border: none;\n            border-radius: 8px;\n            font-size: 14px;\n            font-weight: 600;\n            cursor: pointer;\n            transition: all 0.3s;\n            margin: 5px;\n        }\n\n        .btn-primary {\n            background: #2E75B6;\n            color: white;\n        }\n\n        .btn-primary:hover {\n            background: #1F4E78;\n        }\n\n        .btn-success {\n            background: #28a745;\n            color: white;\n        }\n\n        .btn-full {\n            width: 100%;\n            margin: 10px 0;\n        }\n\n        .section {\n            margin-bottom: 30px;\n        }\n\n        .section-title {\n            font-size: 20px;\n            font-weight: 700;\n            color: #2E75B6;\n            margin-bottom: 15px;\n            padding-bottom: 10px;\n            border-bottom: 3px solid #2E75B6;\n        }\n\n        .button-grid {\n            display: grid;\n            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));\n            gap: 15px;\n            margin-top: 15px;\n        }\n\n        #dashboardResults {\n            margin-top: 30px;\n            padding: 20px;\n            background: #f8f9fa;\n            border-radius: 10px;\n            min-height: 200px;\n        }\n\n        .loading-spinner {\n            text-align: center;\n            padding: 40px;\n        }\n\n        .spinner {\n            border: 4px solid #f3f3f3;\n            border-top: 4px solid #2E75B6;\n            border-radius: 50%;\n            width: 50px;\n            height: 50px;\n            animation: spin 1s linear infinite;\n            margin: 0 auto 20px;\n        }\n\n        @keyframes spin {\n            0% {\n                transform: rotate(0deg);\n            }\n\n            100% {\n                transform: rotate(360deg);\n            }\n        }\n\n        .success-message {\n            background: #d4edda;\n            border: 1px solid #c3e6cb;\n            color: #155724;\n            padding: 20px;\n            border-radius: 8px;\n        }\n\n        .error-message {\n            background: #f8d7da;\n            border: 1px solid #f5c6cb;\n            color: #721c24;\n            padding: 20px;\n            border-radius: 8px;\n        }\n\n        .info-box {\n            background: #d1ecf1;\n            border: 1px solid #bee5eb;\n            color: #0c5460;\n            padding: 15px;\n            border-radius: 8px;\n            margin-bottom: 20px;\n        }\n\n        .tabs {\n            display: flex;\n            border-bottom: 2px solid #e0e0e0;\n            margin-bottom: 20px;\n        }\n\n        .tab {\n            padding: 12px 24px;\n            cursor: pointer;\n            border: none;\n            background: none;\n            font-size: 14px;\n            font-weight: 600;\n            color: #666;\n        }\n\n        .tab.active {\n            color: #2E75B6;\n            border-bottom: 3px solid #2E75B6;\n        }\n\n        .tab-content {\n            display: none;\n        }\n\n        .tab-content.active {\n            display: block;\n        }\n\n        .analysis-card {\n            background: #f8f9fa;\n            padding: 15px;\n            border-radius: 8px;\n            border-left: 4px solid #2E75B6;\n            margin-bottom: 15px;\n        }\n\n        .analysis-card h4 {\n            color: #2E75B6;\n            margin-bottom: 10px;\n        }\n\n        /* Chat Interface Styles */\n        .chat-container {\n            height: 400px;\n            display: flex;\n            flex-direction: column;\n            border: 1px solid #e0e0e0;\n            border-radius: 8px;\n            background: #fff;\n        }\n\n        .chat-history {\n            flex: 1;\n            overflow-y: auto;\n            padding: 15px;\n            background: #f8f9fa;\n        }\n\n        .message {\n            margin-bottom: 15px;\n            display: flex;\n            align-items: flex-start;\n        }\n\n        .message-icon {\n            width: 30px;\n            height: 30px;\n            border-radius: 50%;\n            background: #e0e0e0;\n            display: flex;\n            align-items: center;\n            justify-content: center;\n            margin-right: 10px;\n            font-size: 16px;\n        }\n\n        .message-content {\n            background: white;\n            padding: 10px 15px;\n            border-radius: 12px;\n            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.05);\n            max-width: 80%;\n            word-wrap: break-word;\n        }\n\n        .agent-message .message-content {\n            background: #E2EFD9;\n            border-bottom-left-radius: 2px;\n        }\n\n        .user-message {\n            flex-direction: row-reverse;\n        }\n\n        .user-message .message-icon {\n            margin-right: 0;\n            margin-left: 10px;\n            background: #2E75B6;\n            color: white;\n        }\n\n        .user-message .message-content {\n            background: #2E75B6;\n            color: white;\n            border-bottom-right-radius: 2px;\n        }\n\n        .error-message .message-content {\n            background: #F8D7DA;\n            color: #721C24;\n        }\n\n        .chat-input-area {\n            padding: 15px;\n            border-top: 1px solid #e0e0e0;\n            display: flex;\n            background: white;\n        }\n\n        .chat-input {\n            flex: 1;\n            padding: 10px;\n            border: 1px solid #ccc;\n            border-radius: 20px;\n            margin-right: 10px;\n            outline: none;\n        }\n\n        .chat-input:focus {\n            border-color: #2E75B6;\n        }\n\n        .chat-send-btn {\n            background: #2E75B6;\n            color: white;\n            border: none;\n            width: 40px;\n            height: 40px;\n            border-radius: 50%;\n            cursor: pointer;\n            display: flex;\n            align-items: center;\n            justify-content: center;\n            transition: background 0.2s;\n        }\n\n        .chat-send-btn:hover {\n            background: #1F4E78;\n        }\n\n        /* Dot animation */\n        .dot {\n            animation: typing 1.4s infinite ease-in-out both;\n        }\n\n        .dot:nth-child(1) {\n            animation-delay: -0.32s;\n        }\n\n        .dot:nth-child(2) {\n            animation-delay: -0.16s;\n        }\n\n        @keyframes typing {\n\n            0%,\n            80%,\n            100% {\n                transform: scale(0);\n            }\n\n            40% {\n                transform: scale(1);\n            }\n        }\n    </style>\n</head>\n\n<body>\n    <div class=\"container\">\n        <div class=\"header\">\n            <h1>📊 AI Dashboard Generator</h1>\n            <p>Comprehensive Analysis • Forecasting • Multiple Visualizations</p>\n        </div>\n\n        <div class=\"content\">\n            <!-- Setup Section -->\n            <div class=\"setup-section\">\n                <div class=\"input-group\">\n                    <label for=\"apiKey\">🔑 Claude API Key</label>\n                    <input type=\"password\" id=\"apiKey\" placeholder=\"sk-ant-...\">\n                </div>\n                <button class=\"btn btn-primary btn-full\" onclick=\"initialize()\">\n                    Initialize AI Agent\n                </button>\n            </div>\n\n            <!-- Agent Chat Interface (Always Visible) -->\n            <div class=\"section\">\n                <h2 class=\"section-title\">🤖 AI Assistant</h2>\n                <p style=\"margin-bottom: 15px; color: #666;\">\n                    Ask me anything! I can analyze data, create charts, build dashboards, filter/sort data, and more.\n                </p>\n\n                <div class=\"info-box\" style=\"margin-bottom: 20px;\">\n                    <strong>💡 Examples:</strong><br>\n                    • \"Analyze this data and create a dashboard\"<br>\n                    • \"Show me the top 10 products by sales\"<br>\n                    • \"Create a trend chart for revenue over time\"<br>\n                    • \"Filter data where sales > 1000\"<br>\n                    • \"Clean this data and remove duplicates\"\n                </div>\n\n                <div class=\"chat-container\">\n                    <div id=\"chatHistory\" class=\"chat-history\">\n                        <!-- Messages will appear here -->\n                    </div>\n                    <div class=\"chat-input-area\">\n                        <input type=\"text\" id=\"agentInput\" class=\"chat-input\" placeholder=\"Type your request here...\"\n                            onkeypress=\"if(event.key==='Enter') document.getElementById('agentSendBtn').click()\">\n                        <button id=\"agentSendBtn\" class=\"chat-send-btn\">\n                            ➤\n                        </button>\n                    </div>\n                </div>\n            </div>\n\n            <!-- Results Section -->\n            <div id=\"dashboardResults\">\n                <p style=\"text-align: center; color: #999; padding: 40px;\">\n                    Results will appear here...\n                </p>\n            </div>\n        </div>\n    </div>\n\n    <!-- \n    CRITICAL: Load scripts in this exact order AFTER Office.js \n    Do NOT include office.js again\n    -->\n    <" + "script>\n        // Global variables\n        let assistant;\n\n        // Wait for Office to be ready before loading other scripts\n        Office.onReady((info) => {\n            if (info.host === Office.HostType.Excel) {\n                console.log(\"Office.js ready, loading application scripts...\");\n\n                // Now it's safe to load other scripts\n                loadScript('taskpane.js')\n                    //.then(() => loadScript('financialAnalysis.js')) // File missing\n                    .then(() => loadScript('dashboard-analyzer.js'))\n                    .then(() => loadScript('tool-registry.js'))  // Load tool registry BEFORE autonomous-agent\n                    .then(() => loadScript('autonomous-agent.js'))\n                    .then(() => loadScript('agent-interface-v2.js'))\n                    .then(() => {\n                        console.log(\"All scripts loaded successfully\");\n                    })\n                    .catch(err => {\n                        console.error(\"Error loading scripts:\", err);\n                        showError(\"Failed to load application scripts. Please refresh the page.\");\n                    });\n            }\n        });\n\n        // Helper function to load scripts dynamically\n        function loadScript(src) {\n            return new Promise((resolve, reject) => {\n                const script = document.createElement('script');\n                script.src = src;\n                script.type = \"text/javascript\";\n                script.onload = resolve;\n                script.onerror = reject;\n                document.body.appendChild(script);\n            });\n        }\n\n        // Initialize dashboard\n        function initialize() {\n            const apiKey = document.getElementById('apiKey').value;\n            if (!apiKey) {\n                console.log(\"Please enter your Claude API key\");\n                return;\n            }\n\n            try {\n                // Check if classes are loaded\n                if (typeof DashboardAnalyzer === 'undefined') {\n                    throw new Error(\"Dashboard classes not loaded yet. Please wait a moment and try again.\");\n                }\n\n                initializeDashboard(apiKey);\n                initializeSimpleAgent();\n                console.log(\"Dashboard Analyzer initialized successfully!\");\n            } catch (error) {\n                console.error(\"Initialization error:\", error);\n            }\n        }\n\n        // UI Helper functions\n        function showLoading(message) {\n            const container = document.getElementById('dashboardResults');\n            if (container) {\n                container.innerHTML = `\n                    <div class=\"loading-spinner\">\n                        <div class=\"spinner\"></div>\n                        <p>${message}</p>\n                    </div>\n                `;\n            }\n        }\n\n        function showSuccess(message) {\n            const container = document.getElementById('dashboardResults');\n            if (container) {\n                container.innerHTML = `\n                    <div class=\"success-message\">\n                        <h3>✅ Success!</h3>\n                        <pre>${message}</pre>\n                    </div>\n                `;\n            }\n        }\n\n        function showError(message) {\n            const container = document.getElementById('dashboardResults');\n            if (container) {\n                container.innerHTML = `\n                    <div class=\"error-message\">\n                        <h3>❌ Error</h3>\n                        <p>${message}</p>\n                    </div>\n                `;\n            }\n        }\n    <" + "/script>\n</body>\n\n</html>";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);

/***/ }),

/***/ "./src/taskpane/taskpane.js":
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.js ***!
  \**********************************/
/***/ (function(module) {

function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _toConsumableArray(r) { return _arrayWithoutHoles(r) || _iterableToArray(r) || _unsupportedIterableToArray(r) || _nonIterableSpread(); }
function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _iterableToArray(r) { if ("undefined" != typeof Symbol && null != r[Symbol.iterator] || null != r["@@iterator"]) return Array.from(r); }
function _arrayWithoutHoles(r) { if (Array.isArray(r)) return _arrayLikeToArray(r); }
function ownKeys(e, r) { var t = Object.keys(e); if (Object.getOwnPropertySymbols) { var o = Object.getOwnPropertySymbols(e); r && (o = o.filter(function (r) { return Object.getOwnPropertyDescriptor(e, r).enumerable; })), t.push.apply(t, o); } return t; }
function _objectSpread(e) { for (var r = 1; r < arguments.length; r++) { var t = null != arguments[r] ? arguments[r] : {}; r % 2 ? ownKeys(Object(t), !0).forEach(function (r) { _defineProperty(e, r, t[r]); }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(e, Object.getOwnPropertyDescriptors(t)) : ownKeys(Object(t)).forEach(function (r) { Object.defineProperty(e, r, Object.getOwnPropertyDescriptor(t, r)); }); } return e; }
function _defineProperty(e, r, t) { return (r = _toPropertyKey(r)) in e ? Object.defineProperty(e, r, { value: t, enumerable: !0, configurable: !0, writable: !0 }) : e[r] = t, e; }
function _createForOfIteratorHelper(r, e) { var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (!t) { if (Array.isArray(r) || (t = _unsupportedIterableToArray(r)) || e && r && "number" == typeof r.length) { t && (r = t); var _n = 0, F = function F() {}; return { s: F, n: function n() { return _n >= r.length ? { done: !0 } : { done: !1, value: r[_n++] }; }, e: function e(r) { throw r; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var o, a = !0, u = !1; return { s: function s() { t = t.call(r); }, n: function n() { var r = t.next(); return a = r.done, r; }, e: function e(r) { u = !0, o = r; }, f: function f() { try { a || null == t.return || t.return(); } finally { if (u) throw o; } } }; }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
function _classCallCheck(a, n) { if (!(a instanceof n)) throw new TypeError("Cannot call a class as a function"); }
function _defineProperties(e, r) { for (var t = 0; t < r.length; t++) { var o = r[t]; o.enumerable = o.enumerable || !1, o.configurable = !0, "value" in o && (o.writable = !0), Object.defineProperty(e, _toPropertyKey(o.key), o); } }
function _createClass(e, r, t) { return r && _defineProperties(e.prototype, r), t && _defineProperties(e, t), Object.defineProperty(e, "prototype", { writable: !1 }), e; }
function _toPropertyKey(t) { var i = _toPrimitive(t, "string"); return "symbol" == _typeof(i) ? i : i + ""; }
function _toPrimitive(t, r) { if ("object" != _typeof(t) || !t) return t; var e = t[Symbol.toPrimitive]; if (void 0 !== e) { var i = e.call(t, r || "default"); if ("object" != _typeof(i)) return i; throw new TypeError("@@toPrimitive must return a primitive value."); } return ("string" === r ? String : Number)(t); }
var ExcelAIAssistant = /*#__PURE__*/function () {
  function ExcelAIAssistant(apiKey) {
    _classCallCheck(this, ExcelAIAssistant);
    this.apiKey = apiKey;
    this.claudeApiUrl = 'https://openrouter.ai/api/v1/chat/completions';
    this.model = 'google/gemini-2.5-flash';
  }
  return _createClass(ExcelAIAssistant, [{
    key: "callClaudeAPI",
    value: function () {
      var _callClaudeAPI = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee(userMessage) {
        var systemPrompt,
          headers,
          body,
          response,
          _error$error,
          error,
          data,
          _args = arguments,
          _t;
        return _regenerator().w(function (_context) {
          while (1) switch (_context.p = _context.n) {
            case 0:
              systemPrompt = _args.length > 1 && _args[1] !== undefined ? _args[1] : null;
              _context.p = 1;
              headers = {
                'Content-Type': 'application/json',
                'Authorization': "Bearer ".concat(this.apiKey.trim()),
                'HTTP-Referer': 'https://localhost:3000',
                'X-Title': 'Excel AI Dashboard'
              };
              body = {
                model: this.model,
                messages: [{
                  role: 'user',
                  content: userMessage
                }]
              };
              if (systemPrompt) {
                // OpenRouter/OpenAI usually handles system prompt as a message with role 'system'
                body.messages.unshift({
                  role: 'system',
                  content: systemPrompt
                });
              }
              _context.n = 2;
              return fetch(this.claudeApiUrl, {
                method: 'POST',
                headers: headers,
                body: JSON.stringify(body)
              });
            case 2:
              response = _context.v;
              if (response.ok) {
                _context.n = 4;
                break;
              }
              _context.n = 3;
              return response.json();
            case 3:
              error = _context.v;
              throw new Error("API Error: ".concat(((_error$error = error.error) === null || _error$error === void 0 ? void 0 : _error$error.message) || response.statusText));
            case 4:
              _context.n = 5;
              return response.json();
            case 5:
              data = _context.v;
              return _context.a(2, this.stripMarkdown(data.choices[0].message.content));
            case 6:
              _context.p = 6;
              _t = _context.v;
              console.error('Claude API Error:', _t);
              throw _t;
            case 7:
              return _context.a(2);
          }
        }, _callee, this, [[1, 6]]);
      }));
      function callClaudeAPI(_x) {
        return _callClaudeAPI.apply(this, arguments);
      }
      return callClaudeAPI;
    }()
    /**
     * Helper to strip markdown backticks from AI responses
     */
  }, {
    key: "stripMarkdown",
    value: function stripMarkdown(text) {
      if (!text) return "";
      return text.replace(/```[a-z]*\n?/gi, '').replace(/\n?```/g, '').trim();
    }

    /**
     * Helper to extract and parse JSON from AI response
     */
  }, {
    key: "extractJSON",
    value: function extractJSON(text) {
      if (!text) return null;
      try {
        // First attempt: direct parse of trimmed text
        var trimmed = text.trim();
        // Remove common markdown wrappers if they wrap the whole thing
        var cleaned = trimmed.replace(/^```json\s*/i, '').replace(/\s*```$/i, '').trim();
        try {
          return JSON.parse(cleaned);
        } catch (e) {
          // Secondary attempt: find the first { and last }
          var firstBrace = text.indexOf('{');
          var lastBrace = text.lastIndexOf('}');
          if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
            var jsonCandidate = text.substring(firstBrace, lastBrace + 1);
            return JSON.parse(jsonCandidate);
          }

          // Tertiary: try square brackets for array-only responses
          var firstSquare = text.indexOf('[');
          var lastSquare = text.lastIndexOf(']');
          if (firstSquare !== -1 && lastSquare !== -1 && lastSquare > firstSquare) {
            var _jsonCandidate = text.substring(firstSquare, lastSquare + 1);
            return JSON.parse(_jsonCandidate);
          }
          throw new Error("No JSON structure found in response");
        }
      } catch (error) {
        console.error("JSON Extraction Error:", error, "Original Text:", text);
        return null;
      }
    }

    // ============================================
    // EXCEL DATA OPERATIONS
    // ============================================

    /**
     * Get selected range data from Excel
     */
  }, {
    key: "getSelectedRange",
    value: function () {
      var _getSelectedRange = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3() {
        return _regenerator().w(function (_context3) {
          while (1) switch (_context3.n) {
            case 0:
              _context3.n = 1;
              return Excel.run(/*#__PURE__*/function () {
                var _ref = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2(context) {
                  var range;
                  return _regenerator().w(function (_context2) {
                    while (1) switch (_context2.n) {
                      case 0:
                        range = context.workbook.getSelectedRange();
                        range.load(['values', 'formulas', 'address', 'rowCount', 'columnCount']);
                        _context2.n = 1;
                        return context.sync();
                      case 1:
                        return _context2.a(2, {
                          address: range.address,
                          values: range.values,
                          formulas: range.formulas,
                          rowCount: range.rowCount,
                          columnCount: range.columnCount
                        });
                    }
                  }, _callee2);
                }));
                return function (_x2) {
                  return _ref.apply(this, arguments);
                };
              }());
            case 1:
              return _context3.a(2, _context3.v);
          }
        }, _callee3);
      }));
      function getSelectedRange() {
        return _getSelectedRange.apply(this, arguments);
      }
      return getSelectedRange;
    }()
    /**
     * Write data to a specific range
     */
  }, {
    key: "writeToRange",
    value: (function () {
      var _writeToRange = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee5(address, data) {
        return _regenerator().w(function (_context5) {
          while (1) switch (_context5.n) {
            case 0:
              console.log("[ExcelAIAssistant] Writing to range ".concat(address), data);
              if (!(!data || !data.length || !data[0].length)) {
                _context5.n = 1;
                break;
              }
              console.warn("[ExcelAIAssistant] No data to write.");
              return _context5.a(2, {
                status: 'error',
                message: 'No data to write'
              });
            case 1:
              _context5.n = 2;
              return Excel.run(/*#__PURE__*/function () {
                var _ref2 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4(context) {
                  var sheet, startAddress, parts, sheetName, startCellAddress, startRange, rowCount, colCount, targetRange;
                  return _regenerator().w(function (_context4) {
                    while (1) switch (_context4.n) {
                      case 0:
                        // Parse address to get sheet and starting cell
                        startAddress = address || 'A1';
                        if (startAddress.includes('!')) {
                          parts = address.split('!');
                          sheetName = parts[0].replace(/'/g, '');
                          sheet = context.workbook.worksheets.getItem(sheetName);
                          startAddress = parts[1];
                        } else {
                          sheet = context.workbook.worksheets.getActiveWorksheet();
                        }

                        // Convert startAddress (e.g., "A1:C10") to its top-left cell
                        startCellAddress = startAddress.split(':')[0];
                        startRange = sheet.getRange(startCellAddress);
                        startRange.load(["rowIndex", "columnIndex"]);
                        _context4.n = 1;
                        return context.sync();
                      case 1:
                        rowCount = data.length;
                        colCount = data[0].length; // Calculate final range based on data dimensions
                        targetRange = sheet.getRangeByIndexes(startRange.rowIndex, startRange.columnIndex, rowCount, colCount);
                        targetRange.values = data;
                        targetRange.load('address');
                        console.log("[ExcelAIAssistant] Syncing writeToRange...");
                        _context4.n = 2;
                        return context.sync();
                      case 2:
                        console.log("[ExcelAIAssistant] writeToRange complete at ".concat(targetRange.address));
                        return _context4.a(2, {
                          status: 'success',
                          address: targetRange.address
                        });
                    }
                  }, _callee4);
                }));
                return function (_x5) {
                  return _ref2.apply(this, arguments);
                };
              }());
            case 2:
              return _context5.a(2, _context5.v);
          }
        }, _callee5);
      }));
      function writeToRange(_x3, _x4) {
        return _writeToRange.apply(this, arguments);
      }
      return writeToRange;
    }()
    /**
     * Write data to selected range
     */
    )
  }, {
    key: "writeToSelectedRange",
    value: (function () {
      var _writeToSelectedRange = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee7(data) {
        return _regenerator().w(function (_context7) {
          while (1) switch (_context7.n) {
            case 0:
              console.log("[ExcelAIAssistant] writeToSelectedRange started", data);
              if (!(!data || !data.length || !data[0].length)) {
                _context7.n = 1;
                break;
              }
              console.warn("[ExcelAIAssistant] No data to write.");
              return _context7.a(2, false);
            case 1:
              _context7.n = 2;
              return Excel.run(/*#__PURE__*/function () {
                var _ref3 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee6(context) {
                  var selectedRange, sheet, rowCount, colCount, targetRange;
                  return _regenerator().w(function (_context6) {
                    while (1) switch (_context6.n) {
                      case 0:
                        selectedRange = context.workbook.getSelectedRange();
                        sheet = context.workbook.worksheets.getActiveWorksheet();
                        selectedRange.load(["rowIndex", "columnIndex"]);
                        console.log("[ExcelAIAssistant] Loading range indices...");
                        _context6.n = 1;
                        return context.sync();
                      case 1:
                        rowCount = data.length;
                        colCount = data[0].length;
                        console.log("[ExcelAIAssistant] Target dimensions: ".concat(rowCount, "x").concat(colCount, " at [").concat(selectedRange.rowIndex, ", ").concat(selectedRange.columnIndex, "]"));
                        targetRange = sheet.getRangeByIndexes(selectedRange.rowIndex, selectedRange.columnIndex, rowCount, colCount);
                        targetRange.values = data;
                        targetRange.load('address');
                        console.log("[ExcelAIAssistant] Syncing writeToSelectedRange...");
                        _context6.n = 2;
                        return context.sync();
                      case 2:
                        console.log("[ExcelAIAssistant] writeToSelectedRange complete.");
                        return _context6.a(2, {
                          status: 'success',
                          address: targetRange.address
                        });
                    }
                  }, _callee6);
                }));
                return function (_x7) {
                  return _ref3.apply(this, arguments);
                };
              }());
            case 2:
              return _context7.a(2, _context7.v);
          }
        }, _callee7);
      }));
      function writeToSelectedRange(_x6) {
        return _writeToSelectedRange.apply(this, arguments);
      }
      return writeToSelectedRange;
    }()
    /**
     * Get entire worksheet data
     */
    )
  }, {
    key: "getWorksheetData",
    value: (function () {
      var _getWorksheetData = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee9() {
        var sheetName,
          _args9 = arguments;
        return _regenerator().w(function (_context9) {
          while (1) switch (_context9.n) {
            case 0:
              sheetName = _args9.length > 0 && _args9[0] !== undefined ? _args9[0] : null;
              _context9.n = 1;
              return Excel.run(/*#__PURE__*/function () {
                var _ref4 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee8(context) {
                  var sheet, usedRange;
                  return _regenerator().w(function (_context8) {
                    while (1) switch (_context8.n) {
                      case 0:
                        sheet = sheetName ? context.workbook.worksheets.getItem(sheetName) : context.workbook.worksheets.getActiveWorksheet();
                        usedRange = sheet.getUsedRange();
                        usedRange.load(['values', 'address', 'rowCount', 'columnCount']);
                        sheet.load('name');
                        _context8.n = 1;
                        return context.sync();
                      case 1:
                        return _context8.a(2, {
                          sheetName: sheet.name,
                          address: usedRange.address,
                          values: usedRange.values,
                          rowCount: usedRange.rowCount,
                          columnCount: usedRange.columnCount
                        });
                    }
                  }, _callee8);
                }));
                return function (_x8) {
                  return _ref4.apply(this, arguments);
                };
              }());
            case 1:
              return _context9.a(2, _context9.v);
          }
        }, _callee9);
      }));
      function getWorksheetData() {
        return _getWorksheetData.apply(this, arguments);
      }
      return getWorksheetData;
    }()
    /**
     * Get specific range by address (e.g., "A1:C10")
     */
    )
  }, {
    key: "getRangeByAddress",
    value: (function () {
      var _getRangeByAddress = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee1(address) {
        var sheetName,
          _args1 = arguments;
        return _regenerator().w(function (_context1) {
          while (1) switch (_context1.n) {
            case 0:
              sheetName = _args1.length > 1 && _args1[1] !== undefined ? _args1[1] : null;
              _context1.n = 1;
              return Excel.run(/*#__PURE__*/function () {
                var _ref5 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee0(context) {
                  var sheet, range;
                  return _regenerator().w(function (_context0) {
                    while (1) switch (_context0.n) {
                      case 0:
                        sheet = sheetName ? context.workbook.worksheets.getItem(sheetName) : context.workbook.worksheets.getActiveWorksheet();
                        range = sheet.getRange(address);
                        range.load(['values', 'formulas', 'address']);
                        _context0.n = 1;
                        return context.sync();
                      case 1:
                        return _context0.a(2, {
                          address: range.address,
                          values: range.values,
                          formulas: range.formulas
                        });
                    }
                  }, _callee0);
                }));
                return function (_x0) {
                  return _ref5.apply(this, arguments);
                };
              }());
            case 1:
              return _context1.a(2, _context1.v);
          }
        }, _callee1);
      }));
      function getRangeByAddress(_x9) {
        return _getRangeByAddress.apply(this, arguments);
      }
      return getRangeByAddress;
    }()
    /**
     * Insert formula into a cell or range
     */
    )
  }, {
    key: "insertFormula",
    value: (function () {
      var _insertFormula = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee11(address, formula) {
        var cleanFormula, formulaMatch;
        return _regenerator().w(function (_context11) {
          while (1) switch (_context11.n) {
            case 0:
              console.log("[ExcelAIAssistant] insertFormula started: ".concat(formula, " into ").concat(address));

              // Aggressive sanitization: strip ALL markdown, control characters, and non-formula text
              if (!formula) formula = "";
              cleanFormula = String(formula).replace(/```[a-z]*\n?/gi, '') // Remove opening code blocks
              .replace(/\n?```/g, '') // Remove closing code blocks
              .replace(/[\u0000-\u001F\u007F-\u009F]/g, "") // Remove control characters
              .trim(); // If AI added extra explanation text before or after the formula, 
              // try to isolate the part that starts with = and ends with a parenthesis or number
              formulaMatch = cleanFormula.match(/=[A-Z]+\(.*?\)|=[A-Z]+[0-9]+|=[0-9.]+/i);
              if (formulaMatch) {
                cleanFormula = formulaMatch[0];
              }
              if (!(!address || typeof address !== 'string')) {
                _context11.n = 1;
                break;
              }
              console.error("[ExcelAIAssistant] Invalid address for insertFormula:", address);
              throw new Error("Invalid range address: ".concat(address));
            case 1:
              _context11.n = 2;
              return Excel.run(/*#__PURE__*/function () {
                var _ref6 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee10(context) {
                  var sheet, range, usedRange, finalFormula, _t2, _t3;
                  return _regenerator().w(function (_context10) {
                    while (1) switch (_context10.p = _context10.n) {
                      case 0:
                        sheet = context.workbook.worksheets.getActiveWorksheet();
                        _context10.p = 1;
                        // If address is a whole column (e.g., "O:O"), intersect it with Used Range 
                        if (address.includes(':') && !address.match(/\d/)) {
                          console.log("[ExcelAIAssistant] Processing whole column range: ".concat(address));
                          usedRange = sheet.getUsedRange();
                          range = sheet.getRange(address).getIntersection(usedRange);
                        } else {
                          range = sheet.getRange(address);
                        }

                        // Pre-fetch range to verify it exists before setting formula
                        range.load("address");
                        _context10.n = 2;
                        return context.sync();
                      case 2:
                        _context10.n = 4;
                        break;
                      case 3:
                        _context10.p = 3;
                        _t2 = _context10.v;
                        console.error("[ExcelAIAssistant] Range error for ".concat(address, ":"), _t2);
                        throw new Error("Excel could not find or access the range \"".concat(address, "\". Please ensure the address is valid (e.g., \"A1\" or \"B2:C10\")."));
                      case 4:
                        finalFormula = cleanFormula.startsWith('=') ? cleanFormula : '=' + cleanFormula;
                        _context10.p = 5;
                        // Office.js: Setting formulas to a single string applies it to the whole range
                        range.formulas = finalFormula;
                        _context10.n = 6;
                        return context.sync();
                      case 6:
                        _context10.n = 8;
                        break;
                      case 7:
                        _context10.p = 7;
                        _t3 = _context10.v;
                        console.error("[ExcelAIAssistant] Formula syntax error: ".concat(finalFormula), _t3);
                        throw new Error("Excel rejected the formula \"".concat(finalFormula, "\". Error: ").concat(_t3.message));
                      case 8:
                        console.log("[ExcelAIAssistant] insertFormula complete.");
                        return _context10.a(2, true);
                    }
                  }, _callee10, null, [[5, 7], [1, 3]]);
                }));
                return function (_x11) {
                  return _ref6.apply(this, arguments);
                };
              }());
            case 2:
              return _context11.a(2, _context11.v);
          }
        }, _callee11);
      }));
      function insertFormula(_x1, _x10) {
        return _insertFormula.apply(this, arguments);
      }
      return insertFormula;
    }()
    /**
     * Get all sheet names
     */
    )
  }, {
    key: "getAllSheetNames",
    value: (function () {
      var _getAllSheetNames = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee13() {
        return _regenerator().w(function (_context13) {
          while (1) switch (_context13.n) {
            case 0:
              _context13.n = 1;
              return Excel.run(/*#__PURE__*/function () {
                var _ref7 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee12(context) {
                  var sheets;
                  return _regenerator().w(function (_context12) {
                    while (1) switch (_context12.n) {
                      case 0:
                        sheets = context.workbook.worksheets;
                        sheets.load('items/name');
                        _context12.n = 1;
                        return context.sync();
                      case 1:
                        return _context12.a(2, sheets.items.map(function (sheet) {
                          return sheet.name;
                        }));
                    }
                  }, _callee12);
                }));
                return function (_x12) {
                  return _ref7.apply(this, arguments);
                };
              }());
            case 1:
              return _context13.a(2, _context13.v);
          }
        }, _callee13);
      }));
      function getAllSheetNames() {
        return _getAllSheetNames.apply(this, arguments);
      }
      return getAllSheetNames;
    }()
    /**
     * Create a new worksheet
     */
    )
  }, {
    key: "createNewSheet",
    value: (function () {
      var _createNewSheet = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee15(sheetName) {
        return _regenerator().w(function (_context15) {
          while (1) switch (_context15.n) {
            case 0:
              _context15.n = 1;
              return Excel.run(/*#__PURE__*/function () {
                var _ref8 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee14(context) {
                  var sheet;
                  return _regenerator().w(function (_context14) {
                    while (1) switch (_context14.n) {
                      case 0:
                        sheet = context.workbook.worksheets.add(sheetName);
                        sheet.activate();
                        _context14.n = 1;
                        return context.sync();
                      case 1:
                        return _context14.a(2, sheetName);
                    }
                  }, _callee14);
                }));
                return function (_x14) {
                  return _ref8.apply(this, arguments);
                };
              }());
            case 1:
              return _context15.a(2, _context15.v);
          }
        }, _callee15);
      }));
      function createNewSheet(_x13) {
        return _createNewSheet.apply(this, arguments);
      }
      return createNewSheet;
    }()
    /**
     * Format range (color, bold, etc.)
     */
    )
  }, {
    key: "formatRange",
    value: (function () {
      var _formatRange = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee17(address, formatting) {
        return _regenerator().w(function (_context17) {
          while (1) switch (_context17.n) {
            case 0:
              _context17.n = 1;
              return Excel.run(/*#__PURE__*/function () {
                var _ref9 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee16(context) {
                  var sheet, range;
                  return _regenerator().w(function (_context16) {
                    while (1) switch (_context16.n) {
                      case 0:
                        sheet = context.workbook.worksheets.getActiveWorksheet();
                        range = sheet.getRange(address);
                        if (formatting.bold) range.format.font.bold = true;
                        if (formatting.italic) range.format.font.italic = true;
                        if (formatting.fontSize) range.format.font.size = formatting.fontSize;
                        if (formatting.backgroundColor) range.format.fill.color = formatting.backgroundColor;
                        if (formatting.fontColor) range.format.font.color = formatting.fontColor;
                        _context16.n = 1;
                        return context.sync();
                      case 1:
                        return _context16.a(2, true);
                    }
                  }, _callee16);
                }));
                return function (_x17) {
                  return _ref9.apply(this, arguments);
                };
              }());
            case 1:
              return _context17.a(2, _context17.v);
          }
        }, _callee17);
      }));
      function formatRange(_x15, _x16) {
        return _formatRange.apply(this, arguments);
      }
      return formatRange;
    }() // ============================================
    // AI-POWERED EXCEL OPERATIONS
    // ============================================
    /**
     * Analyze selected data with Claude
     */
    )
  }, {
    key: "analyzeSelection",
    value: function () {
      var _analyzeSelection = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee18() {
        var rangeData, dataStr, prompt;
        return _regenerator().w(function (_context18) {
          while (1) switch (_context18.n) {
            case 0:
              _context18.n = 1;
              return this.getSelectedRange();
            case 1:
              rangeData = _context18.v;
              dataStr = this.formatDataForAI(rangeData.values, rangeData.address);
              prompt = "Analyze this Excel data and provide insights:\n\n".concat(dataStr, "\n\nProvide a concise analysis including patterns, trends, or notable observations.");
              _context18.n = 2;
              return this.callClaudeAPI(prompt);
            case 2:
              return _context18.a(2, _context18.v);
          }
        }, _callee18, this);
      }));
      function analyzeSelection() {
        return _analyzeSelection.apply(this, arguments);
      }
      return analyzeSelection;
    }()
    /**
     * Generate Excel formula based on description
     */
  }, {
    key: "generateFormula",
    value: (function () {
      var _generateFormula = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee19(description) {
        var contextData,
          address,
          prompt,
          dataStr,
          systemPrompt,
          _args19 = arguments;
        return _regenerator().w(function (_context19) {
          while (1) switch (_context19.n) {
            case 0:
              contextData = _args19.length > 1 && _args19[1] !== undefined ? _args19[1] : null;
              address = _args19.length > 2 && _args19[2] !== undefined ? _args19[2] : null;
              prompt = "Generate an Excel formula for: ".concat(description, "\n\nReturn ONLY the formula, starting with =");
              if (contextData) {
                dataStr = this.formatDataForAI(contextData, address);
                prompt += "\n\nContext data (with row/column labels for context):\n".concat(dataStr);
              }
              systemPrompt = "You are an Excel formula expert. Return only the formula without explanation, starting with =";
              _context19.n = 1;
              return this.callClaudeAPI(prompt, systemPrompt);
            case 1:
              return _context19.a(2, _context19.v);
          }
        }, _callee19, this);
      }));
      function generateFormula(_x18) {
        return _generateFormula.apply(this, arguments);
      }
      return generateFormula;
    }()
    /**
     * Clean and transform data using AI
     */
    )
  }, {
    key: "cleanData",
    value: (function () {
      var _cleanData = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee20(instructions) {
        var rangeData, dataStr, prompt, systemPrompt, result, cleanedData;
        return _regenerator().w(function (_context20) {
          while (1) switch (_context20.n) {
            case 0:
              _context20.n = 1;
              return this.getSelectedRange();
            case 1:
              rangeData = _context20.v;
              dataStr = this.formatDataForAI(rangeData.values, rangeData.address);
              prompt = "Transform this data according to: ".concat(instructions, "\n\nOriginal data (with row/column labels for context):\n").concat(dataStr, "\n\nIMPORTANT: \n1. Return ONLY the transformed values in the same tabular structure.\n2. DO NOT include the row numbers or column letters in your response.\n3. Use tabs for columns and newlines for rows.");
              systemPrompt = "Return only the cleaned data values as tab-separated values. No labels, no headers, no explanations.";
              _context20.n = 2;
              return this.callClaudeAPI(prompt, systemPrompt);
            case 2:
              result = _context20.v;
              // Parse result back to 2D array
              cleanedData = this.parseAIDataResponse(result);
              return _context20.a(2, cleanedData);
          }
        }, _callee20, this);
      }));
      function cleanData(_x19) {
        return _cleanData.apply(this, arguments);
      }
      return cleanData;
    }()
    /**
     * Apply cleaned data back to Excel
     */
    )
  }, {
    key: "applyCleanedData",
    value: (function () {
      var _applyCleanedData = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee21(instructions) {
        var cleanedData;
        return _regenerator().w(function (_context21) {
          while (1) switch (_context21.n) {
            case 0:
              _context21.n = 1;
              return this.cleanData(instructions);
            case 1:
              cleanedData = _context21.v;
              _context21.n = 2;
              return this.writeToSelectedRange(cleanedData);
            case 2:
              return _context21.a(2, _context21.v);
          }
        }, _callee21, this);
      }));
      function applyCleanedData(_x20) {
        return _applyCleanedData.apply(this, arguments);
      }
      return applyCleanedData;
    }()
    /**
     * Generate summary table from data
     */
    )
  }, {
    key: "generateSummary",
    value: (function () {
      var _generateSummary = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee23() {
        var rangeData, dataStr, prompt, response, cleanResponse, summaryData, sheetName, writeResult, colCount, endCol;
        return _regenerator().w(function (_context23) {
          while (1) switch (_context23.n) {
            case 0:
              if (this.onStatusUpdate) this.onStatusUpdate("📊 Analyzing data for summary...");
              _context23.n = 1;
              return this.getWorksheetData();
            case 1:
              rangeData = _context23.v;
              dataStr = this.formatDataForAI(rangeData.values, rangeData.address);
              prompt = "Create a summary table from this data:\n\n".concat(dataStr, "\n\nReturn a concise summary table with key metrics as a 2D JSON array: [[\"Metric\", \"Value\"], [\"Total\", 100], ...]");
              _context23.n = 2;
              return this.callClaudeAPI(prompt, "Return only valid JSON.");
            case 2:
              response = _context23.v;
              cleanResponse = response.replace(/```json\n?|\n?```/g, '').trim();
              summaryData = JSON.parse(cleanResponse);
              if (this.onStatusUpdate) this.onStatusUpdate("📑 Creating summary output sheet...");

              // Write the summary to a new sheet
              sheetName = "Summary_".concat(new Date().getTime().toString().slice(-4));
              _context23.n = 3;
              return Excel.run(/*#__PURE__*/function () {
                var _ref0 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee22(context) {
                  var sheet;
                  return _regenerator().w(function (_context22) {
                    while (1) switch (_context22.n) {
                      case 0:
                        sheet = context.workbook.worksheets.add(sheetName);
                        sheet.activate();
                        _context22.n = 1;
                        return context.sync();
                      case 1:
                        return _context22.a(2);
                    }
                  }, _callee22);
                }));
                return function (_x21) {
                  return _ref0.apply(this, arguments);
                };
              }());
            case 3:
              if (!(summaryData && Array.isArray(summaryData))) {
                _context23.n = 5;
                break;
              }
              _context23.n = 4;
              return this.writeToRange("".concat(sheetName, "!A1"), summaryData);
            case 4:
              writeResult = _context23.v;
              // Add address to result
              colCount = summaryData[0] ? summaryData[0].length : 0;
              endCol = this.columnIndexToLetter(colCount - 1);
              return _context23.a(2, {
                summary: summaryData,
                address: "".concat(sheetName, "!A1:").concat(endCol).concat(summaryData.length)
              });
            case 5:
              return _context23.a(2, summaryData);
          }
        }, _callee23, this);
      }));
      function generateSummary() {
        return _generateSummary.apply(this, arguments);
      }
      return generateSummary;
    }()
    /**
     * Answer questions about the data
     */
    )
  }, {
    key: "askAboutData",
    value: (function () {
      var _askAboutData = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee24(question) {
        var worksheetData, dataStr, prompt;
        return _regenerator().w(function (_context24) {
          while (1) switch (_context24.n) {
            case 0:
              _context24.n = 1;
              return this.getWorksheetData();
            case 1:
              worksheetData = _context24.v;
              dataStr = this.formatDataForAI(worksheetData.values);
              prompt = "Based on this Excel data:\n\n".concat(dataStr, "\n\nQuestion: ").concat(question);
              _context24.n = 2;
              return this.callClaudeAPI(prompt);
            case 2:
              return _context24.a(2, _context24.v);
          }
        }, _callee24, this);
      }));
      function askAboutData(_x22) {
        return _askAboutData.apply(this, arguments);
      }
      return askAboutData;
    }()
    /**
     * Generate chart recommendations
     */
    )
  }, {
    key: "recommendChart",
    value: (function () {
      var _recommendChart = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee25() {
        var rangeData, dataStr, prompt;
        return _regenerator().w(function (_context25) {
          while (1) switch (_context25.n) {
            case 0:
              _context25.n = 1;
              return this.getSelectedRange();
            case 1:
              rangeData = _context25.v;
              dataStr = this.formatDataForAI(rangeData.values);
              prompt = "Given this data:\n\n".concat(dataStr, "\n\nRecommend the best chart type and explain why.");
              _context25.n = 2;
              return this.callClaudeAPI(prompt);
            case 2:
              return _context25.a(2, _context25.v);
          }
        }, _callee25, this);
      }));
      function recommendChart() {
        return _recommendChart.apply(this, arguments);
      }
      return recommendChart;
    }()
    /**
     * Automatically generate chart from selected data with AI recommendations
     */
    )
  }, {
    key: "autoGenerateChart",
    value: (function () {
      var _autoGenerateChart = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee26() {
        var rangeData, dataStr, prompt, systemPrompt, aiResponse, recommendation, cleanResponse, chartResult;
        return _regenerator().w(function (_context26) {
          while (1) switch (_context26.n) {
            case 0:
              _context26.n = 1;
              return this.getSelectedRange();
            case 1:
              rangeData = _context26.v;
              dataStr = this.formatDataForAI(rangeData.values); // Ask AI for chart recommendations
              prompt = "Analyze this data and recommend the best chart type:\n\n".concat(dataStr, "\n\nRespond in JSON format:\n{\n  \"chartType\": \"one of: ColumnClustered, ColumnStacked, BarClustered, Line, LineMarkers, Pie, Area, XYScatter, Combo\",\n  \"title\": \"suggested chart title\",\n  \"xAxisTitle\": \"x-axis label\",\n  \"yAxisTitle\": \"y-axis label\",\n  \"reasoning\": \"brief explanation\"\n}");
              systemPrompt = "You are an expert data visualization specialist. Return only valid JSON.";
              _context26.n = 2;
              return this.callClaudeAPI(prompt, systemPrompt);
            case 2:
              aiResponse = _context26.v;
              try {
                // Remove markdown code blocks if present
                cleanResponse = aiResponse.replace(/```json\n?|\n?```/g, '').trim();
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
              _context26.n = 3;
              return this.createChart(rangeData.address, recommendation.chartType, recommendation.title, recommendation.xAxisTitle, recommendation.yAxisTitle);
            case 3:
              chartResult = _context26.v;
              return _context26.a(2, {
                chartId: chartResult.chartId,
                recommendation: recommendation
              });
          }
        }, _callee26, this);
      }));
      function autoGenerateChart() {
        return _autoGenerateChart.apply(this, arguments);
      }
      return autoGenerateChart;
    }()
    /**
     * Create a chart from a data range
     */
    )
  }, {
    key: "createChart",
    value: (function () {
      var _createChart = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee28(sourceRange, chartType, title) {
        var xAxisTitle,
          yAxisTitle,
          _args28 = arguments;
        return _regenerator().w(function (_context28) {
          while (1) switch (_context28.n) {
            case 0:
              xAxisTitle = _args28.length > 3 && _args28[3] !== undefined ? _args28[3] : "";
              yAxisTitle = _args28.length > 4 && _args28[4] !== undefined ? _args28[4] : "";
              _context28.n = 1;
              return Excel.run(/*#__PURE__*/function () {
                var _ref1 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee27(context) {
                  var range, sheet, type, chartTypeMap, excelChartType, chart;
                  return _regenerator().w(function (_context27) {
                    while (1) switch (_context27.n) {
                      case 0:
                        // Use workbook.getRange to support absolute addresses with sheet names
                        range = context.workbook.getRange(sourceRange);
                        sheet = context.workbook.worksheets.getActiveWorksheet(); // Map chart type names to Excel chart types
                        type = (chartType || "").toLowerCase().trim();
                        chartTypeMap = {
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
                        excelChartType = chartTypeMap[type] || Excel.ChartType.columnClustered; // Create chart
                        chart = sheet.charts.add(excelChartType, range, Excel.ChartSeriesBy.auto); // Set chart properties
                        chart.title.text = title;
                        if (xAxisTitle) {
                          chart.axes.categoryAxis.title.text = xAxisTitle;
                        }
                        if (yAxisTitle) {
                          chart.axes.valueAxis.title.text = yAxisTitle;
                        }

                        // Position chart
                        chart.top = 50;
                        chart.left = 500;
                        chart.height = 300;
                        chart.width = 500;

                        // Load chart ID
                        chart.load('id');
                        _context27.n = 1;
                        return context.sync();
                      case 1:
                        return _context27.a(2, {
                          chartId: chart.id,
                          chartType: chartType
                        });
                    }
                  }, _callee27);
                }));
                return function (_x26) {
                  return _ref1.apply(this, arguments);
                };
              }());
            case 1:
              return _context28.a(2, _context28.v);
          }
        }, _callee28);
      }));
      function createChart(_x23, _x24, _x25) {
        return _createChart.apply(this, arguments);
      }
      return createChart;
    }()
    /**
     * Create multiple charts based on AI analysis
     */
    )
  }, {
    key: "generateMultipleCharts",
    value: (function () {
      var _generateMultipleCharts = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee29() {
        var worksheetData, dataStr, prompt, systemPrompt, aiResponse, recommendations, cleanResponse, createdCharts, _iterator, _step, chartRec, result, _t4, _t5;
        return _regenerator().w(function (_context29) {
          while (1) switch (_context29.p = _context29.n) {
            case 0:
              _context29.n = 1;
              return this.getWorksheetData();
            case 1:
              worksheetData = _context29.v;
              dataStr = this.formatDataForAI(worksheetData.values);
              prompt = "Analyze this dataset and suggest 2-3 different visualizations that would provide valuable insights:\n\n".concat(dataStr, "\n\nRespond in JSON format:\n{\n  \"charts\": [\n    {\n      \"dataRange\": \"A1:C10\",\n      \"chartType\": \"ColumnClustered\",\n      \"title\": \"chart title\",\n      \"xAxisTitle\": \"x-axis\",\n      \"yAxisTitle\": \"y-axis\",\n      \"insight\": \"what this chart reveals\"\n    }\n  ]\n}");
              systemPrompt = "Return only valid JSON with 2-3 chart recommendations.";
              _context29.n = 2;
              return this.callClaudeAPI(prompt, systemPrompt);
            case 2:
              aiResponse = _context29.v;
              _context29.p = 3;
              cleanResponse = aiResponse.replace(/```json\n?|\n?```/g, '').trim();
              recommendations = JSON.parse(cleanResponse);
              _context29.n = 5;
              break;
            case 4:
              _context29.p = 4;
              _t4 = _context29.v;
              throw new Error("Failed to parse chart recommendations");
            case 5:
              // Create all recommended charts
              createdCharts = [];
              _iterator = _createForOfIteratorHelper(recommendations.charts);
              _context29.p = 6;
              _iterator.s();
            case 7:
              if ((_step = _iterator.n()).done) {
                _context29.n = 10;
                break;
              }
              chartRec = _step.value;
              _context29.n = 8;
              return this.createChart(chartRec.dataRange, chartRec.chartType, chartRec.title, chartRec.xAxisTitle, chartRec.yAxisTitle);
            case 8:
              result = _context29.v;
              createdCharts.push(_objectSpread(_objectSpread({}, result), {}, {
                insight: chartRec.insight
              }));
            case 9:
              _context29.n = 7;
              break;
            case 10:
              _context29.n = 12;
              break;
            case 11:
              _context29.p = 11;
              _t5 = _context29.v;
              _iterator.e(_t5);
            case 12:
              _context29.p = 12;
              _iterator.f();
              return _context29.f(12);
            case 13:
              return _context29.a(2, {
                charts: createdCharts,
                recommendations: recommendations
              });
          }
        }, _callee29, this, [[6, 11, 12, 13], [3, 4]]);
      }));
      function generateMultipleCharts() {
        return _generateMultipleCharts.apply(this, arguments);
      }
      return generateMultipleCharts;
    }()
    /**
     * Create a specific chart type manually
     */
    )
  }, {
    key: "createSpecificChart",
    value: (function () {
      var _createSpecificChart = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee30(chartType, title) {
        var xAxisTitle,
          yAxisTitle,
          rangeData,
          _args30 = arguments;
        return _regenerator().w(function (_context30) {
          while (1) switch (_context30.n) {
            case 0:
              xAxisTitle = _args30.length > 2 && _args30[2] !== undefined ? _args30[2] : "";
              yAxisTitle = _args30.length > 3 && _args30[3] !== undefined ? _args30[3] : "";
              _context30.n = 1;
              return this.getSelectedRange();
            case 1:
              rangeData = _context30.v;
              _context30.n = 2;
              return this.createChart(rangeData.address, chartType, title, xAxisTitle, yAxisTitle);
            case 2:
              return _context30.a(2, _context30.v);
          }
        }, _callee30, this);
      }));
      function createSpecificChart(_x27, _x28) {
        return _createSpecificChart.apply(this, arguments);
      }
      return createSpecificChart;
    }()
    /**
     * Update existing chart with new data
     */
    )
  }, {
    key: "updateChart",
    value: (function () {
      var _updateChart = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee32(chartId, newDataRange) {
        return _regenerator().w(function (_context32) {
          while (1) switch (_context32.n) {
            case 0:
              _context32.n = 1;
              return Excel.run(/*#__PURE__*/function () {
                var _ref10 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee31(context) {
                  var sheet, chart, range;
                  return _regenerator().w(function (_context31) {
                    while (1) switch (_context31.n) {
                      case 0:
                        sheet = context.workbook.worksheets.getActiveWorksheet();
                        chart = sheet.charts.getItem(chartId); // Update data range
                        range = sheet.getRange(newDataRange);
                        chart.setData(range, Excel.ChartSeriesBy.auto);
                        _context31.n = 1;
                        return context.sync();
                      case 1:
                        return _context31.a(2, true);
                    }
                  }, _callee31);
                }));
                return function (_x31) {
                  return _ref10.apply(this, arguments);
                };
              }());
            case 1:
              return _context32.a(2, _context32.v);
          }
        }, _callee32);
      }));
      function updateChart(_x29, _x30) {
        return _updateChart.apply(this, arguments);
      }
      return updateChart;
    }()
    /**
     * Delete a chart
     */
    )
  }, {
    key: "deleteChart",
    value: (function () {
      var _deleteChart = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee34(chartId) {
        return _regenerator().w(function (_context34) {
          while (1) switch (_context34.n) {
            case 0:
              _context34.n = 1;
              return Excel.run(/*#__PURE__*/function () {
                var _ref11 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee33(context) {
                  var sheet, chart;
                  return _regenerator().w(function (_context33) {
                    while (1) switch (_context33.n) {
                      case 0:
                        sheet = context.workbook.worksheets.getActiveWorksheet();
                        chart = sheet.charts.getItem(chartId);
                        chart.delete();
                        _context33.n = 1;
                        return context.sync();
                      case 1:
                        return _context33.a(2, true);
                    }
                  }, _callee33);
                }));
                return function (_x33) {
                  return _ref11.apply(this, arguments);
                };
              }());
            case 1:
              return _context34.a(2, _context34.v);
          }
        }, _callee34);
      }));
      function deleteChart(_x32) {
        return _deleteChart.apply(this, arguments);
      }
      return deleteChart;
    }()
    /**
     * Get all charts in worksheet
     */
    )
  }, {
    key: "getAllCharts",
    value: (function () {
      var _getAllCharts = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee36() {
        return _regenerator().w(function (_context36) {
          while (1) switch (_context36.n) {
            case 0:
              _context36.n = 1;
              return Excel.run(/*#__PURE__*/function () {
                var _ref12 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee35(context) {
                  var sheet, charts;
                  return _regenerator().w(function (_context35) {
                    while (1) switch (_context35.n) {
                      case 0:
                        sheet = context.workbook.worksheets.getActiveWorksheet();
                        charts = sheet.charts;
                        charts.load('items/name, items/id');
                        _context35.n = 1;
                        return context.sync();
                      case 1:
                        return _context35.a(2, charts.items.map(function (chart) {
                          return {
                            id: chart.id,
                            name: chart.name
                          };
                        }));
                    }
                  }, _callee35);
                }));
                return function (_x34) {
                  return _ref12.apply(this, arguments);
                };
              }());
            case 1:
              return _context36.a(2, _context36.v);
          }
        }, _callee36);
      }));
      function getAllCharts() {
        return _getAllCharts.apply(this, arguments);
      }
      return getAllCharts;
    }()
    /**
     * Style chart with custom colors
     */
    )
  }, {
    key: "styleChart",
    value: (function () {
      var _styleChart = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee38(chartId, styling) {
        return _regenerator().w(function (_context38) {
          while (1) switch (_context38.n) {
            case 0:
              _context38.n = 1;
              return Excel.run(/*#__PURE__*/function () {
                var _ref13 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee37(context) {
                  var sheet, chart;
                  return _regenerator().w(function (_context37) {
                    while (1) switch (_context37.n) {
                      case 0:
                        sheet = context.workbook.worksheets.getActiveWorksheet();
                        chart = sheet.charts.getItem(chartId); // Apply styling
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
                        _context37.n = 1;
                        return context.sync();
                      case 1:
                        return _context37.a(2, true);
                    }
                  }, _callee37);
                }));
                return function (_x37) {
                  return _ref13.apply(this, arguments);
                };
              }());
            case 1:
              return _context38.a(2, _context38.v);
          }
        }, _callee38);
      }));
      function styleChart(_x35, _x36) {
        return _styleChart.apply(this, arguments);
      }
      return styleChart;
    }()
    /**
     * Fill down pattern with AI
     */
    )
  }, {
    key: "fillPattern",
    value: (function () {
      var _fillPattern = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee39(examples, targetCount) {
        var dataStr, prompt, systemPrompt, result;
        return _regenerator().w(function (_context39) {
          while (1) switch (_context39.n) {
            case 0:
              dataStr = this.formatDataForAI(examples);
              prompt = "Continue this pattern for ".concat(targetCount, " more rows:\n\n").concat(dataStr, "\n\nReturn only the new rows in tabular format.");
              systemPrompt = "Return only the data continuation, no explanations.";
              _context39.n = 1;
              return this.callClaudeAPI(prompt, systemPrompt);
            case 1:
              result = _context39.v;
              return _context39.a(2, this.parseAIDataResponse(result));
          }
        }, _callee39, this);
      }));
      function fillPattern(_x38, _x39) {
        return _fillPattern.apply(this, arguments);
      }
      return fillPattern;
    }() // ============================================
    // HELPER FUNCTIONS
    // ============================================
    /**
     * Format Excel data for AI prompt
     */
    )
  }, {
    key: "formatDataForAI",
    value: function formatDataForAI(data) {
      var address = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : null;
      if (!data) return "";
      if (!Array.isArray(data)) return String(data);
      var startColIndex = 0;
      var startRowIndex = 1;
      if (address) {
        // Parse address like "Sheet1!A1:C10" or "A1:G100"
        var rangePart = address.includes('!') ? address.split('!')[1] : address;
        var startCell = rangePart.split(':')[0];
        var colMatch = startCell.match(/[A-Z]+/);
        var rowMatch = startCell.match(/\d+/);
        if (colMatch) startColIndex = this.columnLetterToIndex(colMatch[0]);
        if (rowMatch) startRowIndex = parseInt(rowMatch[0]);
      }
      if (data.length === 0) return "";

      // If it's a 2D array, add headers
      if (Array.isArray(data[0])) {
        // Build Header (Column Letters)
        var header = "      \t";
        for (var j = 0; j < data[0].length; j++) {
          header += this.columnIndexToLetter(startColIndex + j) + "\t";
        }
        var rows = [header.trimEnd()];
        data.forEach(function (row, i) {
          var rowNum = String(startRowIndex + i).padStart(5, ' ');
          var rowData = row.map(function (cell) {
            return cell === null || cell === undefined ? "" : String(cell);
          }).join('\t');
          rows.push("".concat(rowNum, "\t").concat(rowData));
        });
        return rows.join('\n');
      }

      // Fallback for 1D array
      return data.join('\t');
    }

    /**
     * Parse AI response back to 2D array
     */
  }, {
    key: "parseAIDataResponse",
    value: function parseAIDataResponse(response) {
      if (!response) return [[]];

      // Split into lines, filter out empty lines
      var lines = response.trim().split('\n').filter(function (line) {
        return line.trim() !== '';
      });

      // Split each line by tabs or commas (AI might use either if it's hallucinating CSV)
      // and trim each cell
      var rawData = lines.map(function (line) {
        // Try tab first, then comma if no tabs found
        var delimiter = line.includes('\t') ? '\t' : line.includes(',') ? ',' : '\t';
        return line.split(delimiter).map(function (cell) {
          return cell.trim();
        });
      });

      // Ensure it's a perfect rectangle (Excel requirement)
      // 1. Find max columns
      var maxCols = Math.max.apply(Math, _toConsumableArray(rawData.map(function (row) {
        return row.length;
      })));

      // 2. Pad rows that are too short
      return rawData.map(function (row) {
        while (row.length < maxCols) {
          row.push("");
        }
        return row;
      });
    }

    /**
     * Convert column letter to index (A=0, B=1, etc.)
     */
  }, {
    key: "columnLetterToIndex",
    value: function columnLetterToIndex(letter) {
      var index = 0;
      for (var i = 0; i < letter.length; i++) {
        index = index * 26 + (letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
      }
      return index - 1;
    }

    /**
     * Convert column index to letter (0=A, 1=B, etc.)
     */
  }, {
    key: "columnIndexToLetter",
    value: function columnIndexToLetter(index) {
      var letter = '';
      while (index >= 0) {
        letter = String.fromCharCode(index % 26 + 'A'.charCodeAt(0)) + letter;
        index = Math.floor(index / 26) - 1;
      }
      return letter;
    }
  }]);
}(); // ============================================
// USAGE EXAMPLE
// ============================================
// Initialize the assistant
var assistant;
function initializeAssistant(apiKey) {
  assistant = new ExcelAIAssistant(apiKey);
  console.log('Excel AI Assistant initialized');
}

// Example functions to call from UI
function analyzeCurrentSelection() {
  return _analyzeCurrentSelection.apply(this, arguments);
}
function _analyzeCurrentSelection() {
  _analyzeCurrentSelection = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee40() {
    var analysis, _t6;
    return _regenerator().w(function (_context40) {
      while (1) switch (_context40.p = _context40.n) {
        case 0:
          _context40.p = 0;
          _context40.n = 1;
          return assistant.analyzeSelection();
        case 1:
          analysis = _context40.v;
          console.log('Analysis:', analysis);
          return _context40.a(2, analysis);
        case 2:
          _context40.p = 2;
          _t6 = _context40.v;
          console.error('Error:', _t6);
          throw _t6;
        case 3:
          return _context40.a(2);
      }
    }, _callee40, null, [[0, 2]]);
  }));
  return _analyzeCurrentSelection.apply(this, arguments);
}
function generateFormulaFromDescription(_x40) {
  return _generateFormulaFromDescription.apply(this, arguments);
}
function _generateFormulaFromDescription() {
  _generateFormulaFromDescription = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee41(description) {
    var formula, range, _t7;
    return _regenerator().w(function (_context41) {
      while (1) switch (_context41.p = _context41.n) {
        case 0:
          _context41.p = 0;
          _context41.n = 1;
          return assistant.generateFormula(description);
        case 1:
          formula = _context41.v;
          console.log('Generated formula:', formula);

          // Optionally insert into selected cell
          _context41.n = 2;
          return assistant.getSelectedRange();
        case 2:
          range = _context41.v;
          _context41.n = 3;
          return assistant.insertFormula(range.address, formula);
        case 3:
          return _context41.a(2, formula);
        case 4:
          _context41.p = 4;
          _t7 = _context41.v;
          console.error('Error:', _t7);
          throw _t7;
        case 5:
          return _context41.a(2);
      }
    }, _callee41, null, [[0, 4]]);
  }));
  return _generateFormulaFromDescription.apply(this, arguments);
}
function cleanSelectedData(_x41) {
  return _cleanSelectedData.apply(this, arguments);
}
function _cleanSelectedData() {
  _cleanSelectedData = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee42(instructions) {
    var _t8;
    return _regenerator().w(function (_context42) {
      while (1) switch (_context42.p = _context42.n) {
        case 0:
          _context42.p = 0;
          _context42.n = 1;
          return assistant.applyCleanedData(instructions);
        case 1:
          console.log('Data cleaned successfully');
          return _context42.a(2, true);
        case 2:
          _context42.p = 2;
          _t8 = _context42.v;
          console.error('Error:', _t8);
          throw _t8;
        case 3:
          return _context42.a(2);
      }
    }, _callee42, null, [[0, 2]]);
  }));
  return _cleanSelectedData.apply(this, arguments);
}
function askQuestion(_x42) {
  return _askQuestion.apply(this, arguments);
} // ============================================
// CHART GENERATION EXAMPLES
// ============================================
function _askQuestion() {
  _askQuestion = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee43(question) {
    var answer, _t9;
    return _regenerator().w(function (_context43) {
      while (1) switch (_context43.p = _context43.n) {
        case 0:
          _context43.p = 0;
          _context43.n = 1;
          return assistant.askAboutData(question);
        case 1:
          answer = _context43.v;
          console.log('Answer:', answer);
          return _context43.a(2, answer);
        case 2:
          _context43.p = 2;
          _t9 = _context43.v;
          console.error('Error:', _t9);
          throw _t9;
        case 3:
          return _context43.a(2);
      }
    }, _callee43, null, [[0, 2]]);
  }));
  return _askQuestion.apply(this, arguments);
}
function autoCreateChart() {
  return _autoCreateChart.apply(this, arguments);
}
function _autoCreateChart() {
  _autoCreateChart = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee44() {
    var result, _t0;
    return _regenerator().w(function (_context44) {
      while (1) switch (_context44.p = _context44.n) {
        case 0:
          _context44.p = 0;
          _context44.n = 1;
          return assistant.autoGenerateChart();
        case 1:
          result = _context44.v;
          console.log('Chart created:', result);
          console.log('AI Reasoning:', result.recommendation.reasoning);
          return _context44.a(2, result);
        case 2:
          _context44.p = 2;
          _t0 = _context44.v;
          console.error('Error:', _t0);
          throw _t0;
        case 3:
          return _context44.a(2);
      }
    }, _callee44, null, [[0, 2]]);
  }));
  return _autoCreateChart.apply(this, arguments);
}
function createColumnChart() {
  return _createColumnChart.apply(this, arguments);
}
function _createColumnChart() {
  _createColumnChart = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee45() {
    var title,
      result,
      _args45 = arguments,
      _t1;
    return _regenerator().w(function (_context45) {
      while (1) switch (_context45.p = _context45.n) {
        case 0:
          title = _args45.length > 0 && _args45[0] !== undefined ? _args45[0] : "Data Visualization";
          _context45.p = 1;
          _context45.n = 2;
          return assistant.createSpecificChart("ColumnClustered", title, "Categories", "Values");
        case 2:
          result = _context45.v;
          console.log('Column chart created:', result);
          return _context45.a(2, result);
        case 3:
          _context45.p = 3;
          _t1 = _context45.v;
          console.error('Error:', _t1);
          throw _t1;
        case 4:
          return _context45.a(2);
      }
    }, _callee45, null, [[1, 3]]);
  }));
  return _createColumnChart.apply(this, arguments);
}
function createLineChart() {
  return _createLineChart.apply(this, arguments);
}
function _createLineChart() {
  _createLineChart = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee46() {
    var title,
      result,
      _args46 = arguments,
      _t10;
    return _regenerator().w(function (_context46) {
      while (1) switch (_context46.p = _context46.n) {
        case 0:
          title = _args46.length > 0 && _args46[0] !== undefined ? _args46[0] : "Trend Analysis";
          _context46.p = 1;
          _context46.n = 2;
          return assistant.createSpecificChart("LineMarkers", title, "Time Period", "Value");
        case 2:
          result = _context46.v;
          console.log('Line chart created:', result);
          return _context46.a(2, result);
        case 3:
          _context46.p = 3;
          _t10 = _context46.v;
          console.error('Error:', _t10);
          throw _t10;
        case 4:
          return _context46.a(2);
      }
    }, _callee46, null, [[1, 3]]);
  }));
  return _createLineChart.apply(this, arguments);
}
function createPieChart() {
  return _createPieChart.apply(this, arguments);
}
function _createPieChart() {
  _createPieChart = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee47() {
    var title,
      result,
      _args47 = arguments,
      _t11;
    return _regenerator().w(function (_context47) {
      while (1) switch (_context47.p = _context47.n) {
        case 0:
          title = _args47.length > 0 && _args47[0] !== undefined ? _args47[0] : "Distribution";
          _context47.p = 1;
          _context47.n = 2;
          return assistant.createSpecificChart("Pie", title);
        case 2:
          result = _context47.v;
          console.log('Pie chart created:', result);
          return _context47.a(2, result);
        case 3:
          _context47.p = 3;
          _t11 = _context47.v;
          console.error('Error:', _t11);
          throw _t11;
        case 4:
          return _context47.a(2);
      }
    }, _callee47, null, [[1, 3]]);
  }));
  return _createPieChart.apply(this, arguments);
}
function generateDashboard() {
  return _generateDashboard.apply(this, arguments);
}
function _generateDashboard() {
  _generateDashboard = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee48() {
    var result, _t12;
    return _regenerator().w(function (_context48) {
      while (1) switch (_context48.p = _context48.n) {
        case 0:
          _context48.p = 0;
          _context48.n = 1;
          return assistant.generateMultipleCharts();
        case 1:
          result = _context48.v;
          console.log('Dashboard created with', result.charts.length, 'charts');
          result.charts.forEach(function (chart, i) {
            console.log("Chart ".concat(i + 1, " insight:"), chart.insight);
          });
          return _context48.a(2, result);
        case 2:
          _context48.p = 2;
          _t12 = _context48.v;
          console.error('Error:', _t12);
          throw _t12;
        case 3:
          return _context48.a(2);
      }
    }, _callee48, null, [[0, 2]]);
  }));
  return _generateDashboard.apply(this, arguments);
}
function customStyleChart(_x43) {
  return _customStyleChart.apply(this, arguments);
} // Export for use in other modules
function _customStyleChart() {
  _customStyleChart = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee49(chartId) {
    var _t13;
    return _regenerator().w(function (_context49) {
      while (1) switch (_context49.p = _context49.n) {
        case 0:
          _context49.p = 0;
          _context49.n = 1;
          return assistant.styleChart(chartId, {
            titleFontSize: 16,
            titleColor: "#2E75B6",
            legendPosition: "Bottom",
            backgroundColor: "#F0F0F0"
          });
        case 1:
          console.log('Chart styled successfully');
          return _context49.a(2, true);
        case 2:
          _context49.p = 2;
          _t13 = _context49.v;
          console.error('Error:', _t13);
          throw _t13;
        case 3:
          return _context49.a(2);
      }
    }, _callee49, null, [[0, 2]]);
  }));
  return _customStyleChart.apply(this, arguments);
}
if ( true && module.exports) {
  module.exports = {
    ExcelAIAssistant: ExcelAIAssistant
  };
}

// Ensure global availability
if (typeof window !== 'undefined') {
  window.ExcelAIAssistant = ExcelAIAssistant;
}

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Check if module exists (development only)
/******/ 		if (__webpack_modules__[moduleId] === undefined) {
/******/ 			var e = new Error("Cannot find module '" + moduleId + "'");
/******/ 			e.code = 'MODULE_NOT_FOUND';
/******/ 			throw e;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/************************************************************************/
/******/ 	
/******/ 	// startup
/******/ 	// Load entry module and return exports
/******/ 	// This entry module is referenced by other modules so it can't be inlined
/******/ 	__webpack_require__("./src/taskpane/taskpane.js");
/******/ 	var __webpack_exports__ = __webpack_require__("./src/taskpane/taskpane.html");
/******/ 	
/******/ })()
;
//# sourceMappingURL=taskpane.js.map