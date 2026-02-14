# ExcelAI Agent (Dashbord)

> **Your Autonomous Financial Analyst & Data Scientist inside Excel.**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Office Add-in](https://img.shields.io/badge/Platform-Office%20Add--in-green)](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)
[![Powered By Claude](https://img.shields.io/badge/AI-Claude%203.5%20Sonnet-purple)](https://www.anthropic.com/)

**ExcelAI Agent** is a next-generation Office Add-in that transforms Microsoft Excel into an autonomous intelligent workspace. Powered by **Claude 3.5 Sonnet**, it doesn't just answer questions—it **plans, executes, and self-heals** complex data analysis and financial modeling workflows directly in your spreadsheet.

---

## 🚀 Key Features

### 🧠 Autonomous Agent Architecture
*   **ReAct Loop**: Uses a Reason-Act observation loop to break down complex goals into executable steps.
*   **Self-Healing**: Automatically detects errors (e.g., invalid ranges, formula errors) and patches its own execution plan in real-time.
*   **Formula-First Approach**: Builds live, audit-ready Excel models using native formulas (`UNIQUE`, `AVERAGEIF`, `XLOOKUP`) rather than just pasting static values.

### 💼 Investment Banking-Grade Financial Tools
Built-in specialized agents for high-end financial workflows:
*   **DCF Valuation**: Automated Discounted Cash Flow models with sensitivity analysis.
*   **LBO Modeling**: Leveraged Buyout analysis with multiple exit scenarios.
*   **Comparable Analysis**: Auto-generation of trading comps (EV/EBITDA, P/E).
*   **3-Statement Models**: Integrated Income Statement, Balance Sheet, and Cash Flow projections.
*   **Precedent Transactions**: M&A deal analysis and benchmarking.

### 📊 Intelligent Data analysis
*   **Instant Dashboarding**: "Visualize sales trends by region" creates summary tables and charts instantly.
*   **Smart Formatting**: Automatically applies professional styling to headers, tables, and financial data.
*   **Metric Tables**: Dynamic aggregation of data using spillable array formulas.

---

## 🛠️ Installation & Setup

### Prerequisites
*   **Node.js** (v16 or higher)
*   **Microsoft Excel** (Desktop or Web)
*   **Microsoft 365 Account** (for sideloading add-ins)
*   **OpenRouter/Anthropic API Key**

### Quick Start

1.  **Clone the Repository**
    ```bash
    git clone https://github.com/yourusername/excel-ai-agent.git
    cd excel-ai-agent
    ```

2.  **Install Dependencies**
    ```bash
    npm install
    ```

3.  **Start the Development Server**
    ```bash
    npm start
    ```
    *This command will start the local web server and attempt to sideload the add-in into Excel Desktop automatically.*

---

## 🖥️ Usage

1.  Open Excel and click the **"Show Taskpane"** button in the Home ribbon.
2.  Enter your **API Key** in the settings panel.
3.  **Ask a question** or give a command in natural language.

### Example Prompts

**Financial Modeling:**
> "Build a DCF model for a company with $500M revenue, growing 5% annually, with 20% EBITDA margins."

**Data Analysis:**
> "Analyze the sales data in Sheet1. Create a metric table showing Average Revenue by Region and highlight the top performers."

**Market Research:**
> "Create a comparable company analysis table for 5 tech companies with their EV/Revenue and P/E ratios."

---

## 🏗️ Architecture

The project is built on the **Office JavaScript API** and uses a sophisticated agentic workflow:

1.  **Intent Analysis**: The agent classifies user requests (e.g., `dcf_valuation`, `clean_data`) using a specialized prompt.
2.  **Plan Generation**: It constructs a JSON-based execution plan with critical steps and rationale.
3.  **Execution & Recovery**: The agent executes steps sequentially using `Excel.run()`. If a step fails (e.g., "Sheet not found"), the **Self-Healing** module analyzes the error and modifies the plan dynamically to recover.

---

## 🤝 Contributing

Contributions are welcome! We are looking for help with:
*   Adding new financial modeling templates.
*   Improving the self-healing logic for edge cases.
*   expanding support for PowerPoint and Word.

1.  Fork the project.
2.  Create your feature branch (`git checkout -b feature/AmazingFeature`).
3.  Commit your changes (`git commit -m 'Add some AmazingFeature'`).
4.  Push to the branch (`git push origin feature/AmazingFeature`).
5.  Open a Pull Request.

---

## 📄 License

Distributed under the MIT License. See `LICENSE` for more information.

---

*Built with ❤️ for financial analysts and data scientists everywhere.*
