/**
 * HTTP Server for MCP
 *
 * This provides an HTTP endpoint that Copilot Studio can connect to.
 * Uses Streamable transport (required for Copilot Studio as of August 2025)
 */
import express from "express";
import dotenv from "dotenv";
import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { requireRole } from "./lib/auth.js";
// Core MCP tools - Email classification, entity extraction, action planning
import { emailTools, handleEmailTool, getClassificationLog, getClassificationStats, clearClassificationLog } from "./tools/email.js";
import { actionPlannerTools, handleActionPlannerTool } from "./tools/action-planner.js";
import { generateReport } from "./services/report-generator.js";
import { deleteMemory, recallMemory, storeMemory } from "./services/memory.js";
import { emailResources, readEmailResource } from "./resources/email-context.js";
// Microsoft Graph API tools - Teams, Outlook, Dataverse, Planner
import { teamsTools, handleTeamsTool } from "./tools/teams.js";
import { outlookTools, handleOutlookTool } from "./tools/outlook.js";
import { dataverseTools, handleDataverseTool } from "./tools/dataverse.js";
import { plannerTools, handlePlannerTool } from "./tools/planner.js";
import { memoryTools, handleMemoryTool } from "./tools/memory.js";
import { cargowiseTools, handleCargowiseTool } from "./tools/cargowise.js";
import { systemTools, handleSystemTool } from "./tools/system.js";
// Category actions for routing and categorization
import { CATEGORY_ACTIONS, getCategoryAction, getCategoriesByPriority, getOutlookCategories, OUTLOOK_CATEGORY_RULES, } from "./config/category-actions.js";
// Action planner tool names for routing
const ACTION_PLANNER_TOOL_NAMES: string[] = [
    "email_plan_actions",
    "email_get_routing_config",
    "email_extract_cw1_entities",
    "email_check_required_date",
    // Outlook rules tools
    "email_apply_outlook_rules",
    "email_list_outlook_rules",
    "email_get_outlook_rule",
    "email_get_outlook_folders",
];
dotenv.config();
// Tool exposure mode: FULL (default) or MIN (essential only)
const MCP_TOOL_EXPOSURE = (process.env.MCP_TOOL_EXPOSURE || "FULL").toUpperCase();
const ESSENTIAL_TOOL_NAMES = new Set<string>([
    // Email Classification (Core 7)
    "email_classify",
    "email_extract_entities",
    "email_get_client_info",
    "email_plan_actions",
    "email_check_required_date",
    "email_extract_cw1_entities",
    "email_suggest_response",
    // Email Sending (for Copilot Studio agent)
    "outlook_send_email",
]);
const ALL_TOOLS: Tool[] = [
    ...plannerTools,
    ...emailTools,
    ...actionPlannerTools,
    ...teamsTools,
    ...outlookTools,
    ...dataverseTools,
    ...memoryTools,
    ...cargowiseTools,
    ...systemTools,
];
const EXPOSURE_TOOLS: Tool[] = MCP_TOOL_EXPOSURE === "MIN"
    ? ALL_TOOLS.filter(t => ESSENTIAL_TOOL_NAMES.has(t.name))
    : ALL_TOOLS;
// All tool name arrays for routing
const TEAMS_TOOL_NAMES: string[] = teamsTools.map((t: { name: string }) => t.name);
const OUTLOOK_TOOL_NAMES: string[] = outlookTools.map((t: { name: string }) => t.name);
const DATAVERSE_TOOL_NAMES: string[] = dataverseTools.map((t: { name: string }) => t.name);
const PLANNER_TOOL_NAMES: string[] = plannerTools.map((t: { name: string }) => t.name);
const MEMORY_TOOL_NAMES: string[] = memoryTools.map((t: { name: string }) => t.name);
const CARGOWISE_TOOL_NAMES: string[] = cargowiseTools.map((t: { name: string }) => t.name);
const SYSTEM_TOOL_NAMES: string[] = systemTools.map((t: { name: string }) => t.name);
const app = express();
app.use(express.json());
// CORS for Copilot Studio
app.use((req, res, next) => {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
    if (req.method === "OPTIONS") {
        return res.status(200).end();
    }
    next();
});
// Health check endpoint
app.get("/health", (req, res) => {
    res.json({ status: "healthy", server: "navia-mcp-server", version: "1.0.0" });
});
// Teams Tab Configuration Page
app.get("/config", (req, res) => {
    res.setHeader("Content-Type", "text/html");
    res.send(`
<!DOCTYPE html>
<html>
<head>
  <title>Navia Triage Agent - Configuration</title>
  <script src="https://res.cdn.office.net/teams-js/2.0.0/js/MicrosoftTeams.min.js"></script>
  <style>
    body { font-family: 'Segoe UI', sans-serif; padding: 20px; background: #f5f5f5; }
    .container { max-width: 500px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
    h1 { color: #6264A7; margin-bottom: 10px; }
    p { color: #666; line-height: 1.6; }
    .feature { display: flex; align-items: center; margin: 15px 0; padding: 10px; background: #f8f9fa; border-radius: 6px; }
    .feature-icon { font-size: 24px; margin-right: 15px; }
    .success { color: #107c10; font-weight: bold; margin-top: 20px; }
  </style>
</head>
<body>
  <div class="container">
    <h1>Navia Triage Agent</h1>
    <p>AI-powered email classification and task routing for Navia Freight.</p>

    <div class="feature">
      <span class="feature-icon">📧</span>
      <div><strong>Email Classification</strong><br/>Automatically classify emails by priority, category, and SLA</div>
    </div>

    <div class="feature">
      <span class="feature-icon">⭐</span>
      <div><strong>VIP Detection</strong><br/>Identify high-priority clients and internal stakeholders</div>
    </div>

    <div class="feature">
      <span class="feature-icon">📋</span>
      <div><strong>Task Routing</strong><br/>Route tasks to appropriate Planner buckets automatically</div>
    </div>

    <div class="feature">
      <span class="feature-icon">🔍</span>
      <div><strong>Entity Extraction</strong><br/>Extract order numbers, tracking IDs, dates, and amounts</div>
    </div>

    <p class="success" id="status">Configuring...</p>
  </div>

  <script>
    microsoftTeams.app.initialize().then(() => {
      microsoftTeams.pages.config.registerOnSaveHandler((saveEvent) => {
        microsoftTeams.pages.config.setConfig({
          suggestedDisplayName: "Navia Triage Agent",
          entityId: "navia-triage-agent",
          contentUrl: "https://navia-mcp.ngrok.io/dashboard",
          websiteUrl: "https://navia-mcp.ngrok.io/dashboard"
        });
        saveEvent.notifySuccess();
      });

      microsoftTeams.pages.config.setValidityState(true);
      document.getElementById('status').textContent = 'Ready to save!';
    }).catch((err) => {
      document.getElementById('status').textContent = 'Configuration ready';
    });
  </script>
</body>
</html>
  `);
});
// Teams Tab Content Page
app.get("/tab", (req, res) => {
    res.setHeader("Content-Type", "text/html");
    res.send(`
<!DOCTYPE html>
<html>
<head>
  <title>Navia Triage Agent</title>
  <script src="https://res.cdn.office.net/teams-js/2.0.0/js/MicrosoftTeams.min.js"></script>
  <style>
    body { font-family: 'Segoe UI', sans-serif; padding: 20px; margin: 0; }
    .header { background: linear-gradient(135deg, #6264A7 0%, #464775 100%); color: white; padding: 20px; margin: -20px -20px 20px -20px; }
    h1 { margin: 0 0 5px 0; }
    .subtitle { opacity: 0.9; }
    .card { background: white; border: 1px solid #e0e0e0; border-radius: 8px; padding: 20px; margin: 15px 0; }
    .card h3 { margin-top: 0; color: #6264A7; }
    .status { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; }
    .status.online { background: #dff6dd; color: #107c10; }
    .tools { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; }
    .tool { padding: 15px; background: #f8f9fa; border-radius: 6px; }
    .tool-name { font-weight: 600; color: #333; }
    .tool-desc { font-size: 13px; color: #666; margin-top: 5px; }
  </style>
</head>
<body>
  <div class="header">
    <h1>Navia Triage Agent</h1>
    <div class="subtitle">AI-powered email classification for Navia Freight</div>
  </div>

  <div class="card">
    <h3>Server Status</h3>
    <span class="status online">● Online</span>
    <p>MCP Server: navia-mcp.ngrok.io</p>
  </div>

  <div class="card">
    <h3>Available Tools</h3>
    <div class="tools">
      <div class="tool">
        <div class="tool-name">email_classify</div>
        <div class="tool-desc">Classify emails by priority, category, and SLA</div>
      </div>
      <div class="tool">
        <div class="tool-name">email_extract_entities</div>
        <div class="tool-desc">Extract order numbers, tracking IDs, dates</div>
      </div>
      <div class="tool">
        <div class="tool-name">email_get_client_info</div>
        <div class="tool-desc">Look up client tier and contact info</div>
      </div>
      <div class="tool">
        <div class="tool-name">email_suggest_response</div>
        <div class="tool-desc">Get response template suggestions</div>
      </div>
    </div>
  </div>

  <div class="card">
    <h3>Quick Actions</h3>
    <p>Use the chat to interact with the Navia Triage Agent. Try:</p>
    <ul>
      <li>"Classify this email from chris@tilesofezra.com about an urgent order"</li>
      <li>"What's the SLA for Medifab?"</li>
      <li>"Extract entities from this email body..."</li>
    </ul>
  </div>

  <script>
    microsoftTeams.app.initialize();
  </script>
</body>
</html>
  `);
});
// Teams Dashboard Page
app.get("/dashboard", (req, res) => {
    res.setHeader("Content-Type", "text/html");
    res.send(`
<!DOCTYPE html>
<html>
<head>
  <title>Navia Triage Dashboard</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Merriweather:wght@300;400;700&family=Sora:wght@400;600;700&display=swap" rel="stylesheet" />
  <script src="https://res.cdn.office.net/teams-js/2.0.0/js/MicrosoftTeams.min.js"></script>
  <style>
    :root {
      --storm-navy: #0c1a2a;
      --midnight: #132638;
      --mist-cream: #f7f3eb;
      --sage-emerald: #2ec4b6;
      --sunset-salmon: #ff6a5c;
      --goldenrod: #f5c057;
      --ice: #e6eef5;
      --ink: #101821;
    }

    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: radial-gradient(1200px 600px at 10% -10%, #20344a 0%, #0c1a2a 60%, #09141f 100%);
      color: var(--mist-cream);
      font-family: "Merriweather", serif;
    }

    .dashboard {
      max-width: 1200px;
      margin: 0 auto;
      padding: 28px clamp(16px, 3vw, 36px) 60px;
      display: grid;
      gap: 24px;
    }

    .hero {
      position: relative;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 24px;
      padding: 28px;
      border-radius: 22px;
      background: linear-gradient(140deg, #1b2f46 0%, #203b59 45%, #122336 100%);
      box-shadow: 0 20px 40px rgba(4, 10, 18, 0.5);
      overflow: hidden;
    }

    .hero::after {
      content: "";
      position: absolute;
      right: -120px;
      top: -60px;
      width: 320px;
      height: 320px;
      background: radial-gradient(circle, rgba(46, 196, 182, 0.35), rgba(46, 196, 182, 0));
      filter: blur(10px);
    }

    .eyebrow {
      font-family: "Sora", sans-serif;
      letter-spacing: 0.14em;
      font-size: 12px;
      text-transform: uppercase;
      color: rgba(247, 243, 235, 0.7);
      margin: 0 0 8px;
    }

    h1 {
      font-family: "Sora", sans-serif;
      font-size: clamp(28px, 3vw, 40px);
      margin: 0 0 10px;
    }

    .hero p {
      margin: 0;
      line-height: 1.5;
      color: rgba(247, 243, 235, 0.8);
    }

    .stat-badge {
      background: rgba(12, 26, 42, 0.7);
      border: 1px solid rgba(255, 255, 255, 0.08);
      padding: 18px 22px;
      border-radius: 16px;
      min-width: 180px;
      text-align: center;
      font-family: "Sora", sans-serif;
    }

    .stat-badge strong {
      display: block;
      font-size: 32px;
      margin: 6px 0;
      color: var(--goldenrod);
    }

    .overview-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 18px;
    }

    .card {
      background: linear-gradient(160deg, rgba(16, 30, 46, 0.95), rgba(18, 35, 54, 0.9));
      border: 1px solid rgba(255, 255, 255, 0.08);
      border-radius: 18px;
      padding: 18px 20px;
      box-shadow: 0 12px 24px rgba(4, 10, 18, 0.4);
    }

    .card header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-family: "Sora", sans-serif;
      font-size: 13px;
      color: rgba(247, 243, 235, 0.7);
    }

    .metric {
      font-family: "Sora", sans-serif;
      font-size: 24px;
      margin: 12px 0 10px;
      color: var(--mist-cream);
    }

    .progress {
      position: relative;
      height: 8px;
      border-radius: 999px;
      background: rgba(255, 255, 255, 0.08);
      overflow: hidden;
    }

    .progress span {
      position: absolute;
      left: 0;
      top: 0;
      bottom: 0;
      width: var(--value, 40%);
      background: linear-gradient(90deg, var(--sage-emerald), var(--goldenrod));
      border-radius: 999px;
      box-shadow: 0 0 12px rgba(46, 196, 182, 0.6);
      transition: width 0.6s ease;
    }

    section header h2 {
      font-family: "Sora", sans-serif;
      font-size: 18px;
      margin: 0;
    }

    .section-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
    }

    .pill-button {
      background: rgba(255, 255, 255, 0.08);
      border: 1px solid rgba(255, 255, 255, 0.12);
      color: var(--mist-cream);
      border-radius: 999px;
      padding: 6px 14px;
      font-family: "Sora", sans-serif;
      font-size: 12px;
    }

    .channels-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
      gap: 14px;
    }

    .channel-card {
      padding: 14px 16px;
      background: rgba(12, 26, 42, 0.7);
      border-radius: 14px;
      border: 1px solid rgba(255, 255, 255, 0.06);
      display: grid;
      gap: 6px;
    }

    .channel-card h3 {
      margin: 0;
      font-family: "Sora", sans-serif;
      font-size: 15px;
    }

    .status-chip {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      font-size: 11px;
      font-family: "Sora", sans-serif;
      color: var(--sage-emerald);
    }

    .planner-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
      gap: 14px;
    }

    .plan-card {
      padding: 16px;
      border-radius: 16px;
      background: rgba(15, 28, 44, 0.8);
      border: 1px solid rgba(255, 255, 255, 0.06);
    }

    .plan-card h3 {
      margin: 0 0 6px;
      font-family: "Sora", sans-serif;
      font-size: 15px;
    }

    .tag {
      display: inline-flex;
      align-items: center;
      padding: 4px 10px;
      border-radius: 999px;
      font-size: 11px;
      font-family: "Sora", sans-serif;
    }

    .tag.blocker { background: rgba(255, 106, 92, 0.15); color: var(--sunset-salmon); }
    .tag.proxy { background: rgba(245, 192, 87, 0.15); color: var(--goldenrod); }

    .list {
      display: grid;
      gap: 12px;
    }

    .list-item {
      padding: 14px 16px;
      border-radius: 14px;
      background: rgba(14, 26, 40, 0.85);
      border: 1px solid rgba(255, 255, 255, 0.06);
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
    }

    .list-item strong {
      font-family: "Sora", sans-serif;
      font-size: 13px;
    }

    .activity {
      display: grid;
      gap: 12px;
    }

    .activity-item {
      padding: 12px 16px;
      border-radius: 12px;
      background: rgba(10, 22, 34, 0.9);
      border: 1px solid rgba(255, 255, 255, 0.05);
      font-size: 13px;
      line-height: 1.5;
    }

    .activity-item time {
      font-family: "Sora", sans-serif;
      color: rgba(247, 243, 235, 0.6);
      font-size: 11px;
      display: block;
      margin-bottom: 6px;
    }

    .subtle {
      color: rgba(247, 243, 235, 0.6);
      font-size: 12px;
    }

    @media (max-width: 720px) {
      .hero {
        flex-direction: column;
        align-items: flex-start;
      }
    }
  </style>
</head>
<body>
  <main class="dashboard">
    <section class="hero">
      <div>
        <p class="eyebrow">Navia Email Triage</p>
        <h1>Live Inbox Command</h1>
        <p>Routing overview, channel health, and triage flow at a glance.</p>
      </div>
      <div class="stat-badge">
        <span class="subtle">SLA confidence</span>
        <strong id="sla-score">--</strong>
        <small class="subtle">proxy based on urgent volume</small>
      </div>
    </section>

    <section>
      <div class="overview-grid">
        <article class="card">
          <header><span>Total Classified</span><span id="metric-updated">--</span></header>
          <div class="metric" id="metric-total">--</div>
          <div class="progress"><span id="progress-total" style="--value: 40%"></span></div>
        </article>
        <article class="card">
          <header><span>Urgent Queue</span><span class="tag blocker">P1/P2</span></header>
          <div class="metric" id="metric-urgent">--</div>
          <div class="progress"><span id="progress-urgent" style="--value: 20%"></span></div>
        </article>
        <article class="card">
          <header><span>VIP Follow-up</span><span class="tag">VIP</span></header>
          <div class="metric" id="metric-vip">--</div>
          <div class="progress"><span id="progress-vip" style="--value: 10%"></span></div>
        </article>
        <article class="card">
          <header><span>Stale 24h</span><span class="tag proxy">proxy</span></header>
          <div class="metric" id="metric-stale">--</div>
          <div class="progress"><span id="progress-stale" style="--value: 30%"></span></div>
        </article>
      </div>
    </section>

    <section>
      <div class="section-header">
        <h2>Connected Teams Channels</h2>
        <button class="pill-button" id="refresh-btn">Refresh</button>
      </div>
      <div class="channels-grid" id="channels-grid"></div>
    </section>

    <section>
      <div class="section-header">
        <h2>Planner Plans and Tasks</h2>
        <span class="subtle">Planner task counts require Graph access</span>
      </div>
      <div class="planner-grid" id="plans-grid"></div>
    </section>

    <section>
      <div class="section-header">
        <h2>Overdue and Blockers</h2>
        <span class="subtle">Derived from recent triggers</span>
      </div>
      <div class="list" id="blocker-list"></div>
    </section>

    <section>
      <div class="section-header">
        <h2>Recent Activity</h2>
        <span class="subtle">Latest classifications</span>
      </div>
      <div class="activity" id="activity-list"></div>
    </section>
  </main>

  <script>
    function setText(id, value) {
      const el = document.getElementById(id);
      if (el) el.textContent = value;
    }

    function formatPercent(value) {
      if (!Number.isFinite(value)) return "--";
      return value.toFixed(0) + "%";
    }

    function buildChannelCard(key, value) {
      const card = document.createElement("div");
      card.className = "channel-card";
      const name = key.replace(/_/g, " ");
      card.innerHTML = [
        "<h3>" + name + "</h3>",
        "<div class=\"subtle\">Channel ID: " + value + "</div>",
        "<div class=\"status-chip\">● Connected</div>",
      ].join("");
      return card;
    }

    function buildPlanCard(key, value) {
      const card = document.createElement("div");
      card.className = "plan-card";
      const name = key.replace(/_/g, " ");
      card.innerHTML = [
        "<h3>" + name + "</h3>",
        "<div class=\"subtle\">Plan ID: " + value + "</div>",
        "<div class=\"subtle\">Tasks: unavailable</div>",
      ].join("");
      return card;
    }

    function buildListItem(label, value, tag) {
      const item = document.createElement("div");
      item.className = "list-item";
      const tagLabel = tag === "blocker" ? "Blocker" : "Proxy";
      item.innerHTML = [
        "<div>",
        "<strong>" + label + "</strong>",
        "<div class=\"subtle\">" + value + "</div>",
        "</div>",
        "<span class=\"tag " + tag + "\">" + tagLabel + "</span>",
      ].join("");
      return item;
    }

    function buildActivityItem(entry) {
      const item = document.createElement("div");
      item.className = "activity-item";
      const time = new Date(entry.timestamp).toLocaleString();
      const subject = entry.subject || "(no subject)";
      const sender = entry.sender || "unknown";
      const bucket = entry.bucket || "Backlog";
      const confidence = entry.confidence || "low";
      item.innerHTML = [
        "<time>" + time + "</time>",
        "<div><strong>" + subject + "</strong></div>",
        "<div class=\"subtle\">" + sender + " | " + bucket + " | " + confidence + "</div>",
      ].join("");
      return item;
    }

    async function loadDashboard() {
      try {
        const [configRes, statsRes, logRes] = await Promise.all([
          fetch("/api/config"),
          fetch("/api/classification-stats"),
          fetch("/api/classification-log?limit=12"),
        ]);

        const config = await configRes.json();
        const stats = await statsRes.json();
        const logPayload = await logRes.json();
        const entries = logPayload.entries || [];

        const total = stats.total || 0;
        const urgent = (stats.byBucket && stats.byBucket.Urgent) || 0;
        const vip = (stats.byBucket && (stats.byBucket["VIP Follow-up"] || stats.byBucket["VIP Follow-up"])) || 0;
        const now = Date.now();
        const stale = entries.filter(e => now - new Date(e.timestamp).getTime() > 24 * 60 * 60 * 1000).length;
        const slaProxy = total ? Math.max(0, Math.min(100, Math.round(((total - urgent) / total) * 100))) : 0;

        setText("metric-total", total.toString());
        setText("metric-urgent", urgent.toString());
        setText("metric-vip", vip.toString());
        setText("metric-stale", stale.toString());
        setText("sla-score", formatPercent(slaProxy));
        setText("metric-updated", new Date().toLocaleTimeString());

        document.getElementById("progress-total").style.setProperty("--value", total ? "80%" : "10%");
        document.getElementById("progress-urgent").style.setProperty("--value", total ? Math.min(100, Math.round((urgent / total) * 100)) + "%" : "5%");
        document.getElementById("progress-vip").style.setProperty("--value", total ? Math.min(100, Math.round((vip / total) * 100)) + "%" : "5%");
        document.getElementById("progress-stale").style.setProperty("--value", entries.length ? Math.min(100, Math.round((stale / entries.length) * 100)) + "%" : "5%");

        const channelGrid = document.getElementById("channels-grid");
        channelGrid.innerHTML = "";
        if (config.teams_channels) {
          Object.entries(config.teams_channels).forEach(([key, value]) => {
            channelGrid.appendChild(buildChannelCard(key, value));
          });
        }

        const plansGrid = document.getElementById("plans-grid");
        plansGrid.innerHTML = "";
        if (config.planner_plans) {
          Object.entries(config.planner_plans).forEach(([key, value]) => {
            plansGrid.appendChild(buildPlanCard(key, value));
          });
        }

        const blockerList = document.getElementById("blocker-list");
        blockerList.innerHTML = "";
        const triggerCounts = stats.triggerCounts || {};
        const topTriggers = Object.entries(triggerCounts)
          .sort((a, b) => b[1] - a[1])
          .slice(0, 5);

        if (topTriggers.length === 0) {
          blockerList.appendChild(buildListItem("No trigger data", "Awaiting classifications", "proxy"));
        } else {
          topTriggers.forEach(([trigger, count]) => {
            blockerList.appendChild(buildListItem(trigger.replace(/_/g, " "), count + " hits", "blocker"));
          });
        }

        const activityList = document.getElementById("activity-list");
        activityList.innerHTML = "";
        if (entries.length === 0) {
          activityList.innerHTML = "<div class=\"activity-item\"><time>--</time>No recent classifications yet.</div>";
        } else {
          entries.forEach(entry => activityList.appendChild(buildActivityItem(entry)));
        }
      } catch (error) {
        console.error("Dashboard error", error);
      }
    }

    document.getElementById("refresh-btn").addEventListener("click", loadDashboard);
    microsoftTeams.app.initialize();
    loadDashboard();
  </script>
</body>
</html>
  `);
});
// Server info endpoint
app.get("/", (req, res) => {
    res.json({
        name: "navia-mcp-server",
        version: "1.3.0",
        description: "MCP Server for Navia Freight Copilot Agent - Email Triage & Action Planning",
        endpoints: {
            mcp: "/mcp",
            health: "/health",
            emailClassify: "/email.classify",
            emailExtract: "/email.extract",
            tools: "/api/tools",
            resources: "/api/resources",
            generateReport: "/api/generate-report",
            classificationLog: "/api/classification-log",
            classificationStats: "/api/classification-stats",
            monthlyMetrics: "/api/metrics/monthly",
            config: "/api/config",
        },
        routing: {
            teams_channels: {
                ops_ap: "19:825f894d32b8473a9739bc5363b37dd8@thread.tacv2",
                wove: "19:11b039e14c124f2cb42155ac01ea80cc@thread.tacv2",
                it_integrations: "19:Wodl4fnvVItdRGcdLQCuYf8NbNMmZWkBUvGceDoSJp01@thread.tacv2",
                email_tasks: "19:dCAbGRmp2ZNxjwjPCVC7iaKRp6XbJWz6fvO9uBEgalM1@thread.tacv2",
            },
            planner_plans: {
                raft_desk_tickets: "6bf7f866-058c-4dbb-a6cf-9396a562ee46",
                wove_plan: "c09fd612-7331-430b-b6ff-e6d6cdf4ad17",
                integrations_plan: "2d5cd8c3-61be-41da-99a7-d3281aaaacf0",
                email_tasks: "isN5lApe9UKwJzUcBXpiWsgABOwz",
            },
            dm_recipients: {
                daniel_breglia: "Danielb@naviafreight.com",
            },
        },
        capabilities: {
            tools: EXPOSURE_TOOLS.map(t => ({ name: t.name, description: t.description })),
            resources: emailResources.map((r: { uri: string; name: string }) => ({ uri: r.uri, name: r.name })),
        },
    });
});
// Full configuration endpoint for Copilot Studio agent
app.get("/api/config", (req, res) => {
    res.json({
        version: "1.3.0",
        environment: {
            MCP_BASE_URL: "https://navia-mcp.ngrok.io",
            DM_TO_UPN: "Danielb@naviafreight.com",
        },
        teams_channels: {
            CH_AP_RAFT: "19:825f894d32b8473a9739bc5363b37dd8@thread.tacv2",
            CH_WOVE: "19:11b039e14c124f2cb42155ac01ea80cc@thread.tacv2",
            CH_INTEGRATIONS: "19:Wodl4fnvVItdRGcdLQCuYf8NbNMmZWkBUvGceDoSJp01@thread.tacv2",
            CH_EMAILTASKS: "19:dCAbGRmp2ZNxjwjPCVC7iaKRp6XbJWz6fvO9uBEgalM1@thread.tacv2",
        },
        planner_plans: {
            PLAN_RAFT: "6bf7f866-058c-4dbb-a6cf-9396a562ee46",
            PLAN_WOVE: "c09fd612-7331-430b-b6ff-e6d6cdf4ad17",
            PLAN_INTEGRATIONS: "2d5cd8c3-61be-41da-99a7-d3281aaaacf0",
            PLAN_EMAILTASKS: "isN5lApe9UKwJzUcBXpiWsgABOwz",
        },
        buckets: [
            "P1.Urgent-Escalation",
            "P1.AP-Posting-Failure",
            "P2.Raft→CW1-NotPosting",
            "P2.Integration-Project",
            "P3.Warehouse-Orders-GHI",
            "P3.Task-Digests",
            "P4.Newsletters-Verification",
        ],
        vip_clients: [
            "Splash Blanket",
            "Yoni Pleasure Palace",
            "BF Global Warehouse",
            "Tiles of Ezra",
            "BDS Animal Health",
        ],
        vip_domains: [
            "tilesofezra.com",
            "medifab.com",
            "raft.ai",
            "wove.com",
            "splashblanket.com",
            "splash.com",
            "yoni.care",
            "yonipleasurepalace.com",
            "bfglobalwarehouse.com",
            "bfglobal.com",
            "bdsanimalhealth.com",
        ],
        patterns: {
            urgent: "(?i)urgent|asap|shipment delay|pickup (today|now)|expo",
            ap_posting: "(?i)unable to post transaction|ATP trigger",
            raft_not_posting: "(?i)pushed but not posting|Awaiting response from CargoWise",
            warehouse: "(?i)Warehouse Order\\s+W\\d+|Goods Handling Instructions",
            planner: "(?i)You have late tasks|Incomplete Tasks|Overdue Task Alert",
            wove: "(?i)\\bWove\\b",
            integrations_generic: "(?i)(Shopify|NaviaFill|Cin7|Odoo|NetSuite|MachShip|Logic App|Webhook|CW1 Integration)",
        },
        dataverse_tables: {
            // Primary tables (nf_ source of truth)
            nf_triagelog: {
                plural: "nf_triagelogs",
                columns: ["bucket", "category", "dmSentTo", "emailId", "entities", "internetMessageId", "movedToFolder", "plannerGroupId", "plannerPlanId", "plannerTaskId", "priority", "receivedTime", "sender", "status", "subject", "teamsChannelId", "teamsGroupId", "triagedBy", "triageTime", "overriddenBy", "overriddenAt", "originalCategory", "originalBucket", "rerunCount", "confidence", "webLink"],
            },
            nf_emaillog: {
                plural: "nf_emaillogs",
                columns: ["attachmentCount", "bodyPreview", "cc", "conversationId", "from", "fromAddress", "hasAttachments", "importance", "messageId", "received", "sizeKB", "subject", "to"],
            },
            nf_routingrule: {
                plural: "nf_routingrules",
                columns: ["domain", "isActive", "keyword", "ruleName", "targetFolder"],
            },
            nf_extractedentity: {
                plural: "nf_extractedentitys",
                columns: ["entityType", "entityValue", "subject"],
            },
            nf_auditlog: {
                plural: "nf_auditlogs",
                columns: ["action", "details", "email", "success", "timestamp"],
            },
            nf_plannertask: {
                plural: "nf_plannertasks",
                columns: ["assignedTo", "details", "dueDate", "planId", "subject", "taskTitle"],
            },
            nf_prioritylevel: {
                plural: "nf_prioritylevels",
                columns: ["priorityName", "description", "default"],
            },
            nf_keywordrule: {
                plural: "nf_keywordrules",
                columns: ["category", "isActive", "keyword"],
            },
            nf_vipdomain: {
                plural: "nf_vipdomains",
                columns: ["description", "domainName", "isActive"],
            },
            nf_foldermapping: {
                plural: "nf_foldermappings",
                columns: ["category", "folderPath"],
            },
            nf_teamschannelmapping: {
                plural: "nf_teamschannelmappings",
                columns: ["category", "teamId", "channelId", "channelName", "isActive"],
            },
            nf_allowedtarget: {
                plural: "nf_allowedtargets",
                columns: ["targetType", "targetValue", "targetDisplayName", "isActive"],
            },
            nf_groupplan: {
                plural: "nf_groupplans",
                columns: ["groupId", "planId", "description", "isActive"],
            },
            // cr4b9_ tables (no nf_ equivalent)
            cr4b9_p1escalation: {
                plural: "cr4b9_p1escalations",
                columns: ["escalatedTo", "escalationStatus", "escalationSubject", "timerStart", "triageLog", "acknowledgedAt", "resolvedAt", "slaMinutes", "slaStatus", "dmMessageId", "teamsPostId"],
            },
            cr4b9_insightsummary: {
                plural: "cr4b9_insightsummaries",
                columns: ["insightDate", "insightSummaryName", "ruleEffectiveness", "sevenDayTrend", "slaMetrics", "summaryText"],
            },
            cr4b9_tasklink: {
                plural: "cr4b9_tasklinks",
                columns: ["assignedTo", "dueDate", "sourceEmail", "taskName", "taskUrl"],
            },
            // Legacy references
            EmailTriageEvents: {
                plural: "nf_emailtriageevents",
                columns: ["message_id", "bucket", "confidence", "subject", "from", "received_at", "vip", "web_link", "planner_task_id", "teams_message_id", "keywords_hit", "entities", "processed"],
            },
            WarehouseDailyDigest: {
                columns: ["digest_date", "message_id", "required_date", "subject", "from", "web_link", "entities", "posted"],
            },
            PlannerDailyDigest: {
                columns: ["digest_date", "message_id", "subject", "web_link", "posted"],
            },
        },
    });
});
// List tools endpoint
app.get("/api/tools", (req, res) => {
    res.json({
        tools: EXPOSURE_TOOLS.map(t => ({ name: t.name, description: t.description })),
    });
});
// List resources endpoint
app.get("/api/resources", (req, res) => {
    res.json({
        resources: emailResources,
    });
});
// Execute tool endpoint (REST API for testing)
app.post("/api/tools/:toolName", async (req, res) => {
    const { toolName } = req.params;
    const args = req.body || {};
    try {
        let result: { content: { type: string; text: string }[] };
        // Route to appropriate handler based on tool name
        if (ACTION_PLANNER_TOOL_NAMES.includes(toolName)) {
            result = await handleActionPlannerTool(toolName, args);
        }
        else if (TEAMS_TOOL_NAMES.includes(toolName)) {
            result = await handleTeamsTool(toolName, args);
        }
        else if (OUTLOOK_TOOL_NAMES.includes(toolName)) {
            result = await handleOutlookTool(toolName, args);
        }
        else if (DATAVERSE_TOOL_NAMES.includes(toolName)) {
            result = await handleDataverseTool(toolName, args);
        }
        else if (PLANNER_TOOL_NAMES.includes(toolName)) {
            result = await handlePlannerTool(toolName, args);
        }
        else if (MEMORY_TOOL_NAMES.includes(toolName)) {
            result = await handleMemoryTool(toolName, args);
        }
        else if (CARGOWISE_TOOL_NAMES.includes(toolName)) {
            result = await handleCargowiseTool(toolName, args);
        }
        else if (SYSTEM_TOOL_NAMES.includes(toolName)) {
            result = await handleSystemTool(toolName, args);
        }
        else if (toolName.startsWith("email_")) {
            result = await handleEmailTool(toolName, args);
        }
        else {
            return res.status(404).json({ error: `Unknown tool: ${toolName}` });
        }
        res.json(result);
    }
    catch (error) {
        const errorMessage = (error as { message?: string }).message || "Unknown error";
        console.error(`Tool error [${toolName}]:`, error);
        res.status(500).json({ error: errorMessage });
    }
});
app.get("/api/memory/:key?", async (req, res) => {
    try {
        const userId = "danielb@naviafreight.com";
        const key = req.params.key;
        const category = typeof req.query.category === "string" ? req.query.category : undefined;
        const memories = await recallMemory(userId, key, category);
        res.json({ count: memories.length, memories });
    }
    catch (error) {
        const errorMessage = (error as { message?: string }).message || "Unknown error";
        res.status(500).json({ error: errorMessage });
    }
});
app.post("/api/memory", async (req, res) => {
    try {
        const userId = "danielb@naviafreight.com";
        const { key, value, category } = req.body || {};
        if (!key || !value || !category) {
            return res.status(400).json({ error: "key, value, and category are required" });
        }

        await storeMemory(userId, String(key), String(value), String(category));
        res.json({ success: true, key: String(key), category: String(category) });
    }
    catch (error) {
        const errorMessage = (error as { message?: string }).message || "Unknown error";
        res.status(500).json({ error: errorMessage });
    }
});
app.delete("/api/memory/:key", async (req, res) => {
    try {
        const userId = "danielb@naviafreight.com";
        const key = req.params.key;
        await deleteMemory(userId, key);
        res.json({ success: true, key });
    }
    catch (error) {
        const errorMessage = (error as { message?: string }).message || "Unknown error";
        res.status(500).json({ error: errorMessage });
    }
});
// Read resource endpoint (REST API for testing)
app.get("/api/resources/:category/:name", async (req, res) => {
    const { category, name } = req.params;
    const uri = `navia://${category}/${name}`;
    try {
        const result = await readEmailResource(uri);
        res.json(result);
    }
    catch (error) {
        const errorMessage = (error as { message?: string }).message || "Unknown error";
        res.status(404).json({ error: errorMessage });
    }
});
// ============================================================================
// Category → Action Mapping Endpoint
// ============================================================================
app.get("/api/category-actions", (req, res) => {
    res.json({
        categories: CATEGORY_ACTIONS,
        summary: getCategoriesByPriority().map(({ category, action }) => ({
            category,
            displayName: action.displayName,
            priority: `P${action.priority}`,
            sla: action.sla,
            bucket: action.bucket,
            teamsChannel: action.teamsChannel || null,
            dmRecipient: action.dmRecipient || null,
            createTask: action.createTask,
            sendNotification: action.sendNotification,
        })),
    });
});
// Get action for a specific category
app.get("/api/category-actions/:category", (req, res) => {
    const { category } = req.params;
    const action = getCategoryAction(category);
    res.json({ category, action });
});
// Get Outlook categories for a classified email
app.post("/api/outlook-categories", (req, res) => {
    const { classification, sender, subject, body } = req.body;
    if (!classification) {
        return res.status(400).json({ error: "classification is required" });
    }
    const categories = getOutlookCategories(classification, sender || "", subject || "", body);
    res.json({
        classification,
        outlookCategories: categories,
        rules: OUTLOOK_CATEGORY_RULES.length,
    });
});
// Classification log endpoints
app.get("/api/classification-log", (req, res) => {
    const limit = parseInt(String(req.query.limit || "100")) || 100;
    res.json({
        entries: getClassificationLog(limit),
        total: getClassificationLog(500).length,
    });
});
app.get("/api/classification-stats", (req, res) => {
    res.json(getClassificationStats());
});
app.delete("/api/classification-log", (req, res) => {
    clearClassificationLog();
    res.json({ success: true, message: "Classification log cleared" });
});
// Simple classify endpoint for Power Automate (no JSON-RPC wrapper needed)
app.post("/api/classify", async (req, res) => {
    try {
        const { sender, subject, body } = req.body;
        if (!sender || !subject) {
            return res.status(400).json({
                error: "Missing required fields: sender and subject are required"
            });
        }
        const result = await handleEmailTool("email_classify", { sender, subject, body: body || "" });
        res.json(result);
    }
    catch (error) {
        const errorMessage = (error as { message?: string }).message || "Unknown error";
        console.error("[CLASSIFY] Error:", error);
        res.status(500).json({ error: errorMessage });
    }
});
// Simple action planning endpoint for Power Automate
app.post("/api/plan-actions", async (req, res) => {
    try {
        const { sender, subject, body } = req.body;
        if (!sender || !subject) {
            return res.status(400).json({
                error: "Missing required fields: sender and subject are required"
            });
        }
        const result = await handleActionPlannerTool("email_plan_actions", { sender, subject, body: body || "" });
        res.json(result);
    }
    catch (error) {
        const errorMessage = (error as { message?: string }).message || "Unknown error";
        console.error("[PLAN] Error:", error);
        res.status(500).json({ error: errorMessage });
    }
});
// Generate and send report endpoint - ONE CALL DOES IT ALL
app.post("/api/generate-report", async (req, res) => {
    try {
        const request = {
            reportType: req.body.reportType || "on_demand",
            dataSources: req.body.dataSources || ["both"],
            output: req.body.output || ["teams"],
            emailTo: req.body.emailTo,
        };
        console.log("[REPORT] Generating report:", JSON.stringify(request));
        const result = await generateReport(request);
        console.log("[REPORT] Result:", result.success ? "SUCCESS" : "PARTIAL", result.deliveredTo);
        res.json(result);
    }
    catch (error) {
        const errorMessage = (error as { message?: string }).message || "Unknown error";
        console.error("[REPORT] Error:", error);
        res.status(500).json({
            success: false,
            error: errorMessage,
        });
    }
});
// Monthly metrics endpoint - aggregates data for reporting
app.get("/api/metrics/monthly", async (req, res) => {
    const stats = getClassificationStats();
    const log = getClassificationLog(500);
    // Calculate date range
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    const monthEntries = log.filter(e => new Date(e.timestamp) >= monthStart);
    // Group by day for trends
    const dailyCounts: Record<string, number> = {};
    for (const entry of monthEntries) {
        const day = entry.timestamp.split("T")[0];
        dailyCounts[day] = (dailyCounts[day] || 0) + 1;
    }
    // Top clients by volume
    const clientCounts: Record<string, number> = {};
    for (const entry of monthEntries) {
        if (entry.domain) {
            clientCounts[entry.domain] = (clientCounts[entry.domain] || 0) + 1;
        }
    }
    const topClients = Object.entries(clientCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10)
        .map(([domain, count]) => ({ domain, count }));
    res.json({
        period: {
            start: monthStart.toISOString(),
            end: now.toISOString(),
        },
        summary: {
            totalClassified: monthEntries.length,
            byCategory: stats.byCategory,
            byBucket: stats.byBucket,
        },
        trends: {
            dailyCounts,
        },
        topClients,
        potentialFalsePositives: stats.potentialFalsePositives.slice(-20),
    });
});
// Auth gate — require AAD bearer with API.Access role for /mcp.
// See runbook docs/runbooks/navia-mcp-auth-rotation.md for the full rationale;
// without this gate the endpoint was publicly callable once the PP-AU CIDR allowlist
// was lifted on 2026-04-20. Leaves /health, /config, / (tab/dashboard) public.
app.use("/mcp", requireRole("API.Access"));

// MCP Protocol endpoint (JSON-RPC over HTTP)
app.post("/mcp", async (req, res) => {
    const { method, params, id } = req.body;
    // Log incoming requests for debugging
    console.log(`[MCP] Method: ${method}, ID: ${id}`);
    if (params)
        console.log(`[MCP] Params:`, JSON.stringify(params).substring(0, 200));
    try {
        let result: unknown;
        switch (method) {
            case "initialize":
                result = {
                    protocolVersion: "2024-11-05",
                    serverInfo: {
                        name: "navia-mcp-server",
                        version: "1.0.0",
                    },
                    capabilities: {
                        tools: { listChanged: true },
                        resources: { listChanged: true },
                    },
                };
                break;
            case "notifications/initialized":
                // Client acknowledges initialization
                result = {};
                break;
            case "tools/list":
                result = {
                    tools: EXPOSURE_TOOLS,
                };
                break;
            case "tools/call": {
                const { name, arguments: args } = params;
                // Route to appropriate handler based on tool name
                if (ACTION_PLANNER_TOOL_NAMES.includes(name)) {
                    result = await handleActionPlannerTool(name, args || {});
                }
                else if (TEAMS_TOOL_NAMES.includes(name)) {
                    result = await handleTeamsTool(name, args || {});
                }
                else if (OUTLOOK_TOOL_NAMES.includes(name)) {
                    result = await handleOutlookTool(name, args || {});
                }
                else if (DATAVERSE_TOOL_NAMES.includes(name)) {
                    result = await handleDataverseTool(name, args || {});
                }
                else if (PLANNER_TOOL_NAMES.includes(name)) {
                    result = await handlePlannerTool(name, args || {});
                }
                else if (MEMORY_TOOL_NAMES.includes(name)) {
                    result = await handleMemoryTool(name, args || {});
                }
                else if (CARGOWISE_TOOL_NAMES.includes(name)) {
                    result = await handleCargowiseTool(name, args || {});
                }
                else if (SYSTEM_TOOL_NAMES.includes(name)) {
                    result = await handleSystemTool(name, args || {});
                }
                else if (name.startsWith("email_")) {
                    result = await handleEmailTool(name, args || {});
                }
                else {
                    throw new Error(`Unknown tool: ${name}`);
                }
                break;
            }
            case "resources/list":
                result = {
                    resources: emailResources,
                };
                break;
            case "resources/read":
                result = await readEmailResource(params.uri);
                break;
            default:
                throw new Error(`Unknown method: ${method}`);
        }
        res.json({
            jsonrpc: "2.0",
            id,
            result,
        });
    }
    catch (error) {
        const errorMessage = (error as { message?: string }).message || "Unknown error";
        console.error("MCP error:", error);
        res.json({
            jsonrpc: "2.0",
            id,
            error: {
                code: -32000,
                message: errorMessage,
            },
        });
    }
});
const PORT = process.env.PORT || 88;
app.listen(PORT, () => {
    const totalTools = emailTools.length + actionPlannerTools.length + teamsTools.length + outlookTools.length + dataverseTools.length + plannerTools.length + memoryTools.length + cargowiseTools.length + systemTools.length;
    console.log(`
╔═══════════════════════════════════════════════════════════╗
║           Navia Unified MCP Server v2.1.0                 ║
╠═══════════════════════════════════════════════════════════╣
║  Local:     http://localhost:${PORT}                          ║
║  Public:    https://navia-mcp.ngrok.io                    ║
║  MCP:       https://navia-mcp.ngrok.io/mcp                ║
║  Health:    https://navia-mcp.ngrok.io/health             ║
╠═══════════════════════════════════════════════════════════╣
║  Email Tools:      ${String(emailTools.length).padEnd(4)}   Planner Tools:   ${String(plannerTools.length).padEnd(4)}  ║
║  Action Planner:   ${String(actionPlannerTools.length).padEnd(4)}   Teams Tools:     ${String(teamsTools.length).padEnd(4)}  ║
║  Outlook Tools:    ${String(outlookTools.length).padEnd(4)}   Dataverse Tools: ${String(dataverseTools.length).padEnd(4)}  ║
║  Memory Tools:     ${String(memoryTools.length).padEnd(4)}   CargoWise Tools: ${String(cargowiseTools.length).padEnd(4)}  ║
║  System Tools:     ${String(systemTools.length).padEnd(4)}   Resources:       ${String(emailResources.length).padEnd(4)}  ║
║  Total Tools:      ${String(totalTools).padEnd(4)}                              ║
╠═══════════════════════════════════════════════════════════╣
║  UNIFIED CONNECTOR - All services in one!                 ║
╚═══════════════════════════════════════════════════════════╝

Endpoints:
  Email:     /email.classify, /email.extract
  Teams:     /teams.postMessage, /teams.postAdaptiveCard
  Outlook:   /outlook.getEmails, /outlook.assignCategory
  Planner:   /planner.createTask, /planner.listTasks
  Dataverse: /dataverse.listRecords, /dataverse.createRecord
  MCP:       /mcp (JSON-RPC)

To start ngrok tunnel:
  M:\\tools\\ngrok\\ngrok.exe http --domain=navia-mcp.ngrok.io 88
  `);
});
