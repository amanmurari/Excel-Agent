// Simplified Agent Interface - More Reliable Version

class SimpleAgent {
    constructor(analyzer) {
        this.analyzer = analyzer;
        this.isProcessing = false;
    }

    initialize() {
        const input = document.getElementById('agentInput');
        const sendBtn = document.getElementById('agentSendBtn');

        if (input && sendBtn) {
            sendBtn.addEventListener('click', () => this.handleInput());
            input.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    this.handleInput();
                }
            });
        }

        this.showMessage("system", "👋 Hi! I can help you with Excel tasks. Try: 'analyze data', 'create dashboard', 'make a chart'");
    }

    async handleInput() {
        if (this.isProcessing) return;

        const input = document.getElementById('agentInput');
        const message = input.value.trim();

        if (!message) return;

        input.value = '';
        this.showMessage("user", message);
        this.isProcessing = true;
        this.showTyping(true);

        try {
            // Use the advanced Autonomous Agent's processQuery method
            // We pass a callback to show step-by-step progress in the chat
            const response = await this.analyzer.processQuery(message, (status) => {
                this.showMessage("agent", status);
            });

            if (response.status === 'success') {
                this.showMessage("agent", "✅ **High-Fidelity Task complete!**");
                if (response.reflection) {
                    this.showMessage("agent", response.reflection);
                }
            } else if (response.status === 'cancelled') {
                this.showMessage("system", "Operation was cancelled.");
            } else {
                this.showMessage("error", `❌ **Error**: ${response.message || 'An unknown error occurred during execution.'}`);
            }
        } catch (error) {
            console.error("Agent error:", error);
            this.showMessage("error", "❌ **Critical Error**: " + error.message);
        } finally {
            this.isProcessing = false;
            this.showTyping(false);
        }
    }

    showMessage(role, text) {
        const container = document.getElementById('chatHistory');
        if (!container) return;

        const msg = document.createElement('div');
        msg.className = `message ${role}-message`;

        const icon = role === 'user' ? '👤' : role === 'error' ? '❌' : '🤖';

        // Simple Markdown-to-HTML conversion
        let html = text
            .replace(/^### (.*$)/gim, '<h3 style="margin-top: 10px; color: #0078d4;">$1</h3>')
            .replace(/^## (.*$)/gim, '<h2 style="margin-top: 12px; color: #0078d4;">$1</h2>')
            .replace(/\*\*(.*?)\*\*/g, '<strong style="font-weight: 600;">$1</strong>')
            .replace(/\*(.*?)\*/g, '<em style="font-style: italic;">$1</em>')
            .replace(/^\s*\*\s+(.*$)/gim, '<li style="margin-left: 20px;">$1</li>')
            .replace(/^\s*(\d+)\.\s+(.*$)/gim, '<li style="margin-left: 20px; list-style-type: decimal;">$2</li>')
            .replace(/\n/g, '<br>');

        // Wrap lists if they exist
        if (html.includes('<li')) {
            html = html.replace(/(<li.*<\/li>)/gms, '<ul style="margin: 10px 0; padding-left: 0;">$1</ul>');
        }

        msg.innerHTML = `
            <div class="message-icon">${icon}</div>
            <div class="message-content" style="line-height: 1.5; font-size: 14px;">${html}</div>
        `;

        container.appendChild(msg);
        container.scrollTop = container.scrollHeight;
    }

    showTyping(show) {
        const container = document.getElementById('chatHistory');
        let indicator = document.getElementById('typingIndicator');

        if (show && !indicator) {
            indicator = document.createElement('div');
            indicator.id = 'typingIndicator';
            indicator.className = 'message agent-message';
            indicator.innerHTML = `
                <div class="message-icon">🤖</div>
                <div class="message-content">
                    <span class="dot">.</span><span class="dot">.</span><span class="dot">.</span>
                </div>
            `;
            container.appendChild(indicator);
            container.scrollTop = container.scrollHeight;
        } else if (!show && indicator) {
            indicator.remove();
        }
    }
}

// Global initialization
window.simpleAgent = null;

function initializeSimpleAgent() {
    console.log("Initializing Simple Agent...");

    if (!window.dashboardAnalyzer) {
        console.warn("Dashboard analyzer not ready, retrying...");
        setTimeout(initializeSimpleAgent, 500);
        return;
    }

    window.simpleAgent = new SimpleAgent(window.dashboardAnalyzer);
    window.simpleAgent.initialize();
    console.log("✅ Simple Agent ready!");
}

// Auto-initialize when script loads
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeSimpleAgent);
} else {
    initializeSimpleAgent();
}
