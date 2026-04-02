/* global Chart */

const STATUS = {
  Done: "Done",
  Working: "Working",
  Stuck: "Stuck",
  Testing: "Testing",
  Open: "Open",
};

// Seeded in-memory main tasks with subtasks (POC)
const mainTasks = [
  {
    id: "M201",
    name: "Website Redesign",
    assignee: "Sarah",
    parentOrderId: "O-10021",
    createdDate: "2026-03-10",
    dueDate: "2026-03-28",
    subtasks: [
      { id: "S201-1", name: "Create Wireframes", toolName: "Figma", assignee: "Alice", status: STATUS.Working, stage: STATUS.Working, createdDate: "2026-03-10", dueDate: "2026-03-25" },
      { id: "S201-2", name: "Review UX Flow", toolName: "UX Review", assignee: "Charlie", status: STATUS.Open, stage: STATUS.Open, createdDate: "2026-03-11", dueDate: "2026-03-26" },
      { id: "S201-3", name: "Final Approval", toolName: "Sign-off", assignee: "Sarah", status: STATUS.Open, stage: STATUS.Open, createdDate: "2026-03-12", dueDate: "2026-03-28" },
    ],
  },
  {
    id: "M202",
    name: "Payments",
    assignee: "Bob",
    parentOrderId: "O-10022",
    createdDate: "2026-03-05",
    subtasks: [
      { id: "S202-1", name: "Payment Gateway UI", toolName: "UI", assignee: "Bob", status: STATUS.Working, stage: STATUS.Working, createdDate: "2026-03-06", dueDate: "2026-03-18" },
      { id: "S202-2", name: "Payment Webhooks", toolName: "API", assignee: "James", status: STATUS.Testing, stage: STATUS.Testing, createdDate: "2026-03-07", dueDate: "2026-03-19" },
    ],
  },
  {
    id: "M203",
    name: "Authentication",
    assignee: "Alice",
    parentOrderId: "O-10023",
    createdDate: "2026-02-20",
    subtasks: [
      { id: "S203-1", name: "Login API Integration", toolName: "API", assignee: "Alice", status: STATUS.Done, stage: STATUS.Done, createdDate: "2026-02-20", dueDate: "2026-03-12" },
      { id: "S203-2", name: "Password Reset Flow", toolName: "Auth", assignee: "James", status: STATUS.Done, stage: STATUS.Done, createdDate: "2026-02-22", dueDate: "2026-03-06" },
      { id: "S203-3", name: "Role-based Access Control", toolName: "RBAC", assignee: "Sarah", status: STATUS.Open, stage: STATUS.Open, createdDate: "2026-02-25", dueDate: "2026-04-02" },
    ],
  },
  {
    id: "M204",
    name: "Notifications",
    assignee: "Charlie",
    parentOrderId: "O-10021",
    createdDate: "2026-03-01",
    subtasks: [
      { id: "S204-1", name: "Automated Email Service", toolName: "Email", assignee: "Charlie", status: STATUS.Testing, stage: STATUS.Testing, createdDate: "2026-03-02", dueDate: "2026-03-19" },
    ],
  },
  {
    id: "M205",
    name: "Admin & Observability",
    assignee: "Alice",
    // Edge case: no subtasks -> render main only
    parentOrderId: "O-10021",
    createdDate: "2026-03-03",
    dueDate: "2026-03-23",
    status: STATUS.Working,
    subtasks: [],
  },
  {
    id: "M206",
    name: "Performance pass (frontend)",
    assignee: "Charlie",
    // Edge case: no subtasks and older date (overdue depending on today)
    parentOrderId: "O-10021",
    createdDate: "2026-01-02",
    dueDate: "2026-01-15",
    status: STATUS.Working,
    subtasks: [],
  },
];

// Sample orders required by the dashboard spec (10 orders).
// IMPORTANT: Grid View + Upcoming Deadlines must use the same source of truth.
const TOOL_LIST = [
  "Share Graph",
  "Investment Calculator",
  "Share Price Lookup",
  "Share Series",
  "Subscription Center (Manual Distribution)",
  "Financial Calendar",
  "Ticker (Chart, Static, Graph, Financial Calendar)",
  "Share Price Alert",
  "Fragulizer",
  "IR Meeting Request",
  "Latest Share Trades",
  "Market Overview",
];

const ASSIGNEES = ["Alice", "James", "Sarah", "Charlie", "Bob"];

/** Workflow boards and per-board stage lists (single source of truth for filters + grid + chart). */
const BOARDS = {
  ONBOARDING: "On-boarding",
  PRODUCTION: "Production",
  ANALYST: "Analyst",
  DIVIDENDS: "Dividends",
};

const BOARD_LIST = [BOARDS.ONBOARDING, BOARDS.PRODUCTION, BOARDS.ANALYST, BOARDS.DIVIDENDS];

const STAGES_BY_BOARD = {
  [BOARDS.ONBOARDING]: ["Open", "In-progress", "Final QC", "Client Testing", "Live"],
  [BOARDS.PRODUCTION]: ["Open", "In-progress", "First Level QC", "Completed"],
  [BOARDS.DIVIDENDS]: ["Open", "In-progress", "First Level QC", "Completed"],
  [BOARDS.ANALYST]: ["Open", "In-progress", "First Level QC", "Completed"],
};

function getAllStagesUnionSorted() {
  const set = new Set();
  for (const b of BOARD_LIST) {
    for (const s of STAGES_BY_BOARD[b] || []) set.add(s);
  }
  return [...set].sort((a, b) => a.localeCompare(b));
}

/** Stage options for the Stage dropdown given selected Board filter value ("all" or a board name). */
function getStageOptionsForBoardFilter(boardValue) {
  if (!boardValue || boardValue === "all") return getAllStagesUnionSorted();
  return [...(STAGES_BY_BOARD[boardValue] || [])];
}

/** Sub-task workspace labels (segmentation: IOD, QA, .NET, Design, HR). */
const WORKSPACE_OPTIONS = ["IOD Workspace", "QA", ".NET", "Design", "HR"];

const ORDER_TYPE_OPTIONS = ["New Order", "PLG", "Upsell", "Re-design"];

const LABEL_COLORS = {
  Urgent: "#df2f4a",
  High: "#fdab3d",
  Medium: "#579bfc",
  IPO: "#a855f7",
  Concierge: "#00c875",
  Prospects: "#64748b",
};
const LABEL_OPTIONS = Object.keys(LABEL_COLORS);

/** Global dashboard filters (legacy dropdowns kept but hidden). */
let dashboardFilters = { board: "all", stage: "all", workspace: "all" };

/** Smart filter state (ClickUp-like): AND across categories, OR within category. */
const smartFilters = {
  assignees: new Set(),
  orderTypes: new Set(),
  labels: new Set(),
  boards: new Set(),
  stages: new Set(),
  ageBuckets: new Set(),
  workspaces: new Set(),
  assignedToMe: false,
  unassigned: false,
};

function workspaceFilterAppliesToRole() {
  return activeDashboardRole === "manager-lead" || activeDashboardRole === "member";
}

const LOGGED_IN_USER_NAME = "Alice"; // POC: used for "My Tasks" behavior
let managerLeadScope = "all"; // "all" | "mine" (Manager/Lead only)

function syncKpiModeForRole() {
  const sel = document.getElementById("kpiModeSelect");
  if (!sel) return;
  const mainOpt = sel.querySelector('option[value="main"]');

  const isMember = activeDashboardRole === "member";
  if (mainOpt) {
    mainOpt.disabled = isMember;
    mainOpt.hidden = isMember;
  }

  if (isMember && sel.value === "main") {
    sel.value = "sub";
    kpiMode = "sub";
    renderKpis();
  }
}

function toolMatchesDashboardFilters(tool) {
  // Legacy dropdowns (kept but hidden)
  if (dashboardFilters.board !== "all" && tool.board !== dashboardFilters.board) return false;
  if (dashboardFilters.stage !== "all" && tool.workflowStage !== dashboardFilters.stage) return false;
  if (workspaceFilterAppliesToRole() && dashboardFilters.workspace !== "all" && tool.workspace !== dashboardFilters.workspace) return false;

  // Smart filters
  if (smartFilters.boards.size > 0 && !smartFilters.boards.has(tool.board)) return false;
  if (smartFilters.stages.size > 0 && !smartFilters.stages.has(tool.workflowStage)) return false;
  if (workspaceFilterAppliesToRole() && smartFilters.workspaces.size > 0 && !smartFilters.workspaces.has(tool.workspace)) return false;

  // Age buckets (uses current basis; Stage Aging Analysis forces stageEntered basis)
  if (smartFilters.ageBuckets.size > 0) {
    const basis = stageChartViewMode === "stage_aging" ? "stageEntered" : ageBasisMode;
    const days = computeAgeDaysFromTool(tool, new Date(), basis);
    const bucket = ageBucketForDays(days);
    if (!bucket || !smartFilters.ageBuckets.has(bucket)) return false;
  }

  // Assignee
  if (activeDashboardRole === "manager-lead" && managerLeadScope === "mine") {
    if ((tool.assignee || "").trim() !== LOGGED_IN_USER_NAME) return false;
  }
  if (smartFilters.unassigned && (tool.assignee || "").trim() !== "") return false;
  if (smartFilters.assignees.size > 0 && !smartFilters.assignees.has(tool.assignee || "—")) return false;

  // Order Type
  if (smartFilters.orderTypes.size > 0 && !smartFilters.orderTypes.has(tool.orderType)) return false;

  // Labels (OR within labels)
  if (smartFilters.labels.size > 0) {
    const labels = tool.labels || [];
    const any = labels.some((l) => smartFilters.labels.has(l));
    if (!any) return false;
  }
  return true;
}

function dashboardHasActiveFilters() {
  const ws = workspaceFilterAppliesToRole() && dashboardFilters.workspace !== "all";
  const smartWs = workspaceFilterAppliesToRole() && smartFilters.workspaces.size > 0;
  return (
    dashboardFilters.board !== "all" ||
    dashboardFilters.stage !== "all" ||
    ws ||
    smartFilters.assignees.size > 0 ||
    smartFilters.orderTypes.size > 0 ||
    smartFilters.labels.size > 0 ||
    smartFilters.boards.size > 0 ||
    smartFilters.stages.size > 0 ||
    smartWs ||
    smartFilters.unassigned ||
    smartFilters.assignedToMe
  );
}

function syncWorkspaceFilterVisibility() {
  const wrap = document.getElementById("filterWorkspaceWrap");
  const sel = document.getElementById("filterWorkspace");
  if (!wrap || !sel) return;
  const show = workspaceFilterAppliesToRole();
  wrap.hidden = !show;
  wrap.setAttribute("aria-hidden", show ? "false" : "true");
  if (!show) {
    dashboardFilters.workspace = "all";
    sel.value = "all";
  }
}

function syncManagerLeadScopeVisibility() {
  const wrap = document.getElementById("managerLeadScopeWrap");
  if (!wrap) return;
  const show = activeDashboardRole === "manager-lead";
  wrap.hidden = !show;
  wrap.setAttribute("aria-hidden", show ? "false" : "true");
  if (!show) managerLeadScope = "all";

  const btnAll = document.getElementById("scopeAllTasks");
  const btnMine = document.getElementById("scopeMyTasks");
  if (btnAll && btnMine) {
    const isAll = managerLeadScope === "all";
    btnAll.classList.toggle("active", isAll);
    btnAll.setAttribute("aria-selected", isAll ? "true" : "false");
    btnMine.classList.toggle("active", !isAll);
    btnMine.setAttribute("aria-selected", !isAll ? "true" : "false");
  }
}

function filterToolsForDashboard(tools) {
  return (tools || []).filter(toolMatchesDashboardFilters);
}

function applyDashboardFiltersToFlatTools(flat) {
  return flat.filter((t) => toolMatchesDashboardFilters(t));
}


/** Orders whose tool list is non-empty after dashboard filters (for main-task KPIs). */
function getOrdersMatchingDashboardFilters() {
  if (!dashboardHasActiveFilters()) return orders;
  return orders
    .map((o) => ({ ...o, tools: filterToolsForDashboard(o.tools || []) }))
    .filter((o) => o.tools.length > 0);
}

function makeToolId(orderId, idx) {
  return `${orderId}-T${String(idx + 1).padStart(2, "0")}`;
}

function addDaysIso(base, deltaDays) {
  const d = clampToLocalMidnight(base);
  d.setDate(d.getDate() + deltaDays);
  return toIsoDateStringOrNull(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`);
}

const orders = (() => {
  const today = new Date();
  const names = ["Apple", "Samsung", "Adani", "Tesla", "Reliance", "Infosys", "TCS", "HDFC Bank", "ICICI Bank", "Wipro"];
  const orderBaseId = 10021;

  // Spread due dates across overdue/today/next 7 days/after to exercise UI rules.
  const dueOffsets = [-2, 0, 1, 3, 5, 7, 9, 12];

  return names.map((name, i) => {
    const id = `O-${orderBaseId + i}`;
    const orderType = ORDER_TYPE_OPTIONS[i % ORDER_TYPE_OPTIONS.length];
    const orderLabels = [LABEL_OPTIONS[i % LABEL_OPTIONS.length], LABEL_OPTIONS[(i + 2) % LABEL_OPTIONS.length]].slice(0, (i % 2) + 1);
    // Each Company (Main Task) has the same IR Tools list (Sub Tasks)
    const toolCount = TOOL_LIST.length;
    const tools = [];
    for (let t = 0; t < toolCount; t++) {
      const toolName = TOOL_LIST[(i * 3 + t) % TOOL_LIST.length];
      const assignee = ASSIGNEES[(i + t) % ASSIGNEES.length];
      const createdDate = addDaysIso(today, -20 + ((i + t) % 14)); // seeded for health calc
      const stageEnteredDate = addDaysIso(today, -10 + ((i * 2 + t) % 9)); // time spent in stage proxy
      const dueDate = addDaysIso(today, dueOffsets[(i + t) % dueOffsets.length]);
      const status = (i + t) % 5 === 0 ? STATUS.Done : (i + t) % 3 === 0 ? STATUS.Testing : STATUS.Working;
      const board = BOARD_LIST[(i + t) % BOARD_LIST.length];
      const stages = STAGES_BY_BOARD[board];
      const workflowStage = stages[(i * 3 + t) % stages.length];
      const workspace = WORKSPACE_OPTIONS[(i * 2 + t) % WORKSPACE_OPTIONS.length];
      const labels = t % 4 === 0 ? [LABEL_OPTIONS[(i + t) % LABEL_OPTIONS.length]] : orderLabels;
      tools.push({
        id: makeToolId(id, t),
        name: toolName,
        assignee,
        createdDate,
        stageEnteredDate,
        dueDate,
        status,
        board,
        workflowStage,
        workspace,
        orderType,
        labels,
        stage: workflowStage, // modal "Stage" column = workflow stage
      });
    }
    return { id, name, tools, orderType, labels: orderLabels };
  });
})();

const charts = {
  status: null,
  assignee: null,
  health: null,
};

const smartCharts = {
  workloadDonut: null,
  equalBar: null,
};

let assigneeChartMode = "performance"; // "workload" | "performance" | "smart"
let activeTheme = "dark"; // "dark" | "light"

/** Active dashboard role tab (same UI shell; use for future permissions / data filtering). */
let activeDashboardRole = "super-admin";

const DASHBOARD_ROLE_KEYS = ["super-admin", "manager-lead", "member"];

function applyDashboardRole(role) {
  const next = DASHBOARD_ROLE_KEYS.includes(role) ? role : "super-admin";
  activeDashboardRole = next;
  document.body.dataset.dashboardRole = next;
  const shell = document.getElementById("dashboardShell");
  if (shell) shell.dataset.role = next;

  document.querySelectorAll(".role-tab[data-role]").forEach((btn) => {
    const isActive = btn.dataset.role === next;
    btn.classList.toggle("active", isActive);
    btn.setAttribute("aria-selected", isActive ? "true" : "false");
  });

  try {
    localStorage.setItem("orionDashboardRole", next);
  } catch {
    // ignore
  }

  syncWorkspaceFilterVisibility();
  syncManagerLeadScopeVisibility();
  syncKpiModeForRole();
}

function initRoleTabs() {
  const buttons = document.querySelectorAll(".role-tab[data-role]");
  if (!buttons.length) return;

  let saved = null;
  try {
    saved = localStorage.getItem("orionDashboardRole");
  } catch {
    saved = null;
  }
  applyDashboardRole(DASHBOARD_ROLE_KEYS.includes(saved) ? saved : "super-admin");

  for (const btn of buttons) {
    btn.addEventListener("click", () => {
      applyDashboardRole(btn.dataset.role);
      applyAndRender();
    });
  }
}

let activePrimaryTab = "order-tools"; // "board-overview" | "order-tools" | "individual-board" | "time-tracking"

function mountSharedDashboardInto(panelKey) {
  const shared = document.getElementById("sharedDashboardContent");
  if (!shared) return;
  const host =
    panelKey === "individual-board"
      ? document.getElementById("primaryPanelIndividualBoard")
      : document.getElementById("primaryPanelOrderTools");
  if (!host) return;
  if (shared.parentElement !== host) host.appendChild(shared);
}

function applyPrimaryTab(tab) {
  const next = tab === "board-overview" || tab === "time-tracking" || tab === "individual-board" ? tab : "order-tools";
  activePrimaryTab = next;
  document.body.dataset.primaryTab = next;

  const panels = document.querySelectorAll(".primary-panel[data-primary-panel]");
  panels.forEach((p) => {
    const isActive = p.getAttribute("data-primary-panel") === next;
    p.hidden = !isActive;
    p.setAttribute("aria-hidden", isActive ? "false" : "true");
  });

  document.querySelectorAll(".primary-tab[data-primary-tab]").forEach((btn) => {
    const isActive = btn.getAttribute("data-primary-tab") === next;
    btn.classList.toggle("active", isActive);
    btn.setAttribute("aria-selected", isActive ? "true" : "false");
  });

  // Both "Order and Tools" and "Individual Board" use the same dashboard layout for now.
  if (next === "individual-board") mountSharedDashboardInto("individual-board");
  else if (next === "order-tools") mountSharedDashboardInto("order-tools");
}

function initPrimaryTabs() {
  const tabs = document.querySelectorAll(".primary-tab[data-primary-tab]");
  if (!tabs.length) return;
  applyPrimaryTab("order-tools");
  tabs.forEach((btn) => {
    btn.addEventListener("click", () => {
      applyPrimaryTab(btn.getAttribute("data-primary-tab"));
    });
  });
}

function debugLog(payload) {
  // #region agent log
  fetch("http://127.0.0.1:7251/ingest/88409bff-c33c-4975-a14b-d528fb0dfa59", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ...payload, timestamp: Date.now() }),
  }).catch(() => {});
  // #endregion
}

function getThemeTokens() {
  const styles = getComputedStyle(document.body);
  return {
    textDim: styles.getPropertyValue("--text-dim").trim() || "#b1b1b1",
    gridLine: styles.getPropertyValue("--grid-line").trim() || "rgba(255,255,255,0.08)",
    accent: styles.getPropertyValue("--accent").trim() || "#00a9ff",
    textMain: styles.getPropertyValue("--text-main").trim() || "#ffffff",
  };
}

function applyTheme(theme) {
  activeTheme = theme === "light" ? "light" : "dark";
  document.body.dataset.theme = activeTheme;
  try {
    localStorage.setItem("orionTheme", activeTheme);
  } catch {
    // ignore
  }

  const t = getThemeTokens();
  Chart.defaults.color = t.textDim;

  // Update existing charts to reflect new colors (ticks/grid + redraw)
  for (const c of [charts.status, charts.assignee, charts.health, smartCharts.workloadDonut, smartCharts.equalBar]) {
    if (!c) continue;
    if (c.options?.scales?.x?.ticks) c.options.scales.x.ticks.color = t.textDim;
    if (c.options?.scales?.y?.ticks) c.options.scales.y.ticks.color = t.textDim;
    if (c.options?.scales?.x?.grid) c.options.scales.x.grid.color = t.gridLine;
    if (c.options?.scales?.y?.grid) c.options.scales.y.grid.color = t.gridLine;
    c.update();
  }
}

function $(id) {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing element #${id}`);
  return el;
}

function clampToLocalMidnight(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function parseIsoDate(iso) {
  // Treat date-only ISO as local date (not UTC-shifted)
  const [y, m, d] = iso.split("-").map((n) => Number(n));
  return new Date(y, m - 1, d);
}

function safeParseDueDate(isoOrNull) {
  if (!isoOrNull) return null;
  try {
    return parseIsoDate(isoOrNull);
  } catch {
    return null;
  }
}

function toIsoDateStringOrNull(value) {
  if (!value) return null;
  const d = safeParseDueDate(value);
  if (!d) return null;
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function inDateRange(isoDate, startIso, endIso) {
  const d = safeParseDueDate(isoDate);
  if (!d) return false;
  const t = d.getTime();
  if (startIso) {
    const s = safeParseDueDate(startIso);
    if (s && t < s.getTime()) return false;
  }
  if (endIso) {
    const e = safeParseDueDate(endIso);
    if (e && t > e.getTime()) return false;
  }
  return true;
}

function includesText(haystack, needle) {
  if (!needle) return true;
  return String(haystack || "").toLowerCase().includes(String(needle).toLowerCase());
}

function parseNumberOrNull(v) {
  if (v === "" || v === null || v === undefined) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function formatShortMonthDay(date) {
  return date.toLocaleDateString(undefined, { month: "short", day: "2-digit" });
}

function toDateInputValue(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function addMonths(date, deltaMonths) {
  const d = new Date(date);
  d.setMonth(d.getMonth() + deltaMonths);
  return d;
}

function startOfMonth(date) {
  const d = new Date(date);
  d.setDate(1);
  d.setHours(0, 0, 0, 0);
  return d;
}

function endOfMonth(date) {
  const d = new Date(date);
  d.setMonth(d.getMonth() + 1, 0); // last day of month
  d.setHours(23, 59, 59, 999);
  return d;
}

function getDateRangeFromPreset(preset, startInput, endInput) {
  const now = new Date();
  const todayStart = clampToLocalMidnight(now);

  const presetNormalized = String(preset || "").trim();

  if (presetNormalized === "Custom") {
    const s = startInput?.value ? parseIsoDate(startInput.value) : null;
    const e = endInput?.value ? parseIsoDate(endInput.value) : null;
    if (s) s.setHours(0, 0, 0, 0);
    if (e) e.setHours(23, 59, 59, 999);
    return { start: s, end: e };
  }

  if (presetNormalized === "Current Month") {
    const start = startOfMonth(now);
    const end = endOfMonth(now);
    return { start, end };
  }

  if (presetNormalized === "Last Month") {
    const lastMonth = addMonths(now, -1);
    const start = startOfMonth(lastMonth);
    const end = endOfMonth(lastMonth);
    return { start, end };
  }

  if (presetNormalized === "Last 3 Months") {
    const start = startOfMonth(addMonths(now, -2));
    const end = endOfMonth(now);
    return { start, end };
  }

  if (presetNormalized === "Last 6 Months") {
    const start = startOfMonth(addMonths(now, -5));
    const end = endOfMonth(now);
    return { start, end };
  }

  if (presetNormalized === "Last 7 Days") {
    const start = new Date(todayStart);
    start.setDate(start.getDate() - 6);
    const end = new Date(now);
    end.setHours(23, 59, 59, 999);
    return { start, end };
  }

  if (presetNormalized === "Last 30 Days") {
    const start = new Date(todayStart);
    start.setDate(start.getDate() - 29);
    const end = new Date(now);
    end.setHours(23, 59, 59, 999);
    return { start, end };
  }

  // Fallback: no date filter
  return { start: null, end: null };
}

function filterTasksByDate(allTasks, range) {
  const { start, end } = range || {};
  if (!start && !end) return [...allTasks];

  return allTasks.filter((t) => {
    const due = parseIsoDate(t.dueDate);
    const dueTime = due.getTime();
    if (start && dueTime < start.getTime()) return false;
    if (end && dueTime > end.getTime()) return false;
    return true;
  });
}

function filterMainTasksByDate(allMainTasks, range) {
  const { start, end } = range || {};
  if (!start && !end) return allMainTasks.map((m) => ({ ...m, subtasks: [...(m.subtasks || [])] }));

  return allMainTasks
    .map((m) => {
      const subtasks = (m.subtasks || []).filter((s) => {
        const due = safeParseDueDate(s.dueDate);
        if (!due) return false;
        const dueTime = due.getTime();
        if (start && dueTime < start.getTime()) return false;
        if (end && dueTime > end.getTime()) return false;
        return true;
      });

      const mainDue = safeParseDueDate(m.dueDate);
      const mainInRange =
        !!mainDue &&
        (!start || mainDue.getTime() >= start.getTime()) &&
        (!end || mainDue.getTime() <= end.getTime());

      const includeMain = subtasks.length > 0 || (subtasks.length === 0 && mainInRange);
      if (!includeMain) return null;
      return { ...m, subtasks };
    })
    .filter(Boolean);
}

function flattenForTableAndCharts(filteredMainTasks) {
  const flat = [];
  for (const m of filteredMainTasks) {
    const hasSubtasks = (m.subtasks || []).length > 0;
    if (!hasSubtasks) {
      flat.push({
        id: m.id,
        name: m.name,
        assignee: m.assignee || "—",
        status: m.status || STATUS.Open,
        dueDate: m.dueDate,
        parentId: null,
        parentName: null,
      });
      continue;
    }
    for (const s of m.subtasks) {
      flat.push({
        id: s.id,
        name: `${m.name} ↳ ${s.name}`,
        assignee: s.assignee || m.assignee || "—",
        status: s.status || STATUS.Open,
        dueDate: s.dueDate,
        parentId: m.id,
        parentName: m.name,
      });
    }
  }
  return flat;
}

function isOverdue(task, now = new Date()) {
  if (task.status === STATUS.Done) return false;
  const today = clampToLocalMidnight(now).getTime();
  const due = clampToLocalMidnight(parseIsoDate(task.dueDate)).getTime();
  return due < today;
}

function daysUntil(dueDateIso, now = new Date()) {
  const due = safeParseDueDate(dueDateIso);
  if (!due) return null;
  const today = clampToLocalMidnight(now);
  const dueMid = clampToLocalMidnight(due);
  const diffMs = dueMid.getTime() - today.getTime();
  return Math.round(diffMs / (1000 * 60 * 60 * 24));
}

function dueStateFromDays(days) {
  if (days === null) return { key: "healthy", className: "due-healthy", label: "—" };
  if (days < 0) return { key: "overdue", className: "due-overdue", label: `Overdue ${Math.abs(days)}d` };
  if (days <= 2) return { key: "warning", className: "due-warning", label: `${days}d left` };
  return { key: "healthy", className: "due-healthy", label: `${days}d left` };
}

function computeHealthFromItems(items, now = new Date()) {
  // items: array of {dueDate, stage/status}
  let overdue = 0;
  let nearDue = 0;
  let total = 0;
  for (const it of items) {
    if (!it?.dueDate) continue;
    total += 1;
    const d = daysUntil(it.dueDate, now);
    if (d === null) continue;
    if (d < 0 && it.status !== STATUS.Done) overdue += 1;
    else if (d <= 2 && it.status !== STATUS.Done) nearDue += 1;
  }

  if (overdue > 0) return { key: "overdue", label: "Overdue", className: "pill overdue" };
  if (nearDue > 0) return { key: "warning", label: "Near due", className: "pill warning" };
  return { key: "healthy", label: "Healthy", className: "pill healthy" };
}

function computeHealthFromTool(tool, now = new Date()) {
  if (!tool?.dueDate) return { key: "healthy", label: "Healthy", className: "pill healthy" };
  if (tool.status === STATUS.Done) return { key: "healthy", label: "Healthy", className: "pill healthy" };
  const d = daysUntil(tool.dueDate, now);
  if (d === null) return { key: "healthy", label: "Healthy", className: "pill healthy" };
  if (d < 0) return { key: "overdue", label: "Overdue", className: "pill overdue" };
  if (d <= 2) return { key: "warning", label: "Near due", className: "pill warning" };
  return { key: "healthy", label: "Healthy", className: "pill healthy" };
}

function flattenToolsFromOrders(allOrders) {
  const out = [];
  for (const o of allOrders) {
    for (const tool of o.tools || []) {
      out.push({
        toolId: tool.id,
        orderId: o.id,
        orderName: o.name,
        toolName: tool.name,
        assignee: tool.assignee,
        status: tool.status,
        createdDate: tool.createdDate,
        stageEnteredDate: tool.stageEnteredDate,
        dueDate: tool.dueDate,
        board: tool.board,
        workflowStage: tool.workflowStage,
        workspace: tool.workspace,
        orderType: tool.orderType || o.orderType,
        labels: tool.labels || o.labels || [],
      });
    }
  }
  return out;
}

function getFilteredFlatTools() {
  const scoped = applyDashboardFiltersToFlatTools(flattenToolsFromOrders(orders));
  const range = getActiveDateRangeFromUi();
  const { start, end } = range || {};
  if (!start && !end) return scoped;
  return scoped.filter((t) => {
    const due = safeParseDueDate(t.dueDate);
    if (!due) return false;
    const time = due.getTime();
    if (start && time < start.getTime()) return false;
    if (end && time > end.getTime()) return false;
    return true;
  });
}

function computeKpis(filteredTasks) {
  const totalTasks = filteredTasks.length;
  // With hierarchy enabled, subtasks count will be computed separately (see applyAndRender)
  const subtasks = 0;
  const completed = filteredTasks.filter((t) => t.status === STATUS.Done).length;
  const overdue = filteredTasks.filter((t) => isOverdue(t)).length;

  // Simple, deterministic pulse trend for POC (based on completion ratio)
  const completionRatio = totalTasks > 0 ? completed / totalTasks : 0;
  const pulseTrendPct = Math.round((completionRatio - 0.5) * 100); // centered around 50%
  const pulseTrend = totalTasks === 0 ? "—" : `${pulseTrendPct >= 0 ? "+" : ""}${pulseTrendPct}%`;

  return { totalTasks, subtasks, completed, overdue, pulseTrend };
}

function statusToClass(status) {
  switch (status) {
    case STATUS.Done:
      return "bg-done";
    case STATUS.Working:
      return "bg-working";
    case STATUS.Testing:
      return "bg-testing";
    case STATUS.Stuck:
      return "bg-stuck";
    default:
      return "bg-open";
  }
}

function isInProgressStatus(status) {
  return status === STATUS.Working || status === STATUS.Testing || status === STATUS.Open || status === STATUS.Stuck;
}

function computeAssigneeStats(flatTasks) {
  const by = new Map();
  const now = new Date();

  const ensure = (name) => {
    if (!by.has(name)) {
      by.set(name, {
        name,
        completed: 0,
        inProgress: 0,
        overdue: 0,
        total: 0,
        lastActivity: null, // Date
      });
    }
    return by.get(name);
  };

  for (const t of flatTasks) {
    const who = t.assignee || "—";
    const s = ensure(who);
    s.total += 1;

    if (t.status === STATUS.Done) s.completed += 1;
    else if (isInProgressStatus(t.status)) s.inProgress += 1;

    if (t.dueDate && isOverdue(t, now)) s.overdue += 1;

    // Use dueDate as a simple activity proxy in this POC.
    const activity = safeParseDueDate(t.dueDate);
    if (activity && (!s.lastActivity || activity.getTime() > s.lastActivity.getTime())) s.lastActivity = activity;
  }

  return Array.from(by.values());
}

function performanceScore(stat) {
  return stat.completed * 1.5 + stat.inProgress * 1.0 - stat.overdue * 1.2;
}

function sortAssigneesByPerformance(stats) {
  return [...stats].sort((a, b) => {
    const sa = performanceScore(a);
    const sb = performanceScore(b);
    if (sb !== sa) return sb - sa;
    if (b.completed !== a.completed) return b.completed - a.completed;
    if (a.overdue !== b.overdue) return a.overdue - b.overdue;
    const ta = a.lastActivity ? a.lastActivity.getTime() : -Infinity;
    const tb = b.lastActivity ? b.lastActivity.getTime() : -Infinity;
    if (tb !== ta) return tb - ta;
    return a.name.localeCompare(b.name);
  });
}

function buildAssigneeChartModel(flatTasks) {
  const stats = computeAssigneeStats(flatTasks);
  const memberCount = stats.length;

  const totalValue =
    assigneeChartMode === "workload"
      ? stats.reduce((sum, s) => sum + s.total, 0)
      : stats.reduce((sum, s) => sum + performanceScore(s), 0);

  const sorted =
    assigneeChartMode === "workload"
      ? [...stats].sort((a, b) => b.total - a.total || a.name.localeCompare(b.name))
      : sortAssigneesByPerformance(stats);

  // Behavior rules:
  // - <= 7: show all
  // - 8-12: show all + scroll
  // - > 12: top 8 + Others
  const showAll = memberCount <= 12;
  const needsOthers = memberCount > 12;

  const displayed = needsOthers ? sorted.slice(0, 8) : sorted;
  const others = needsOthers ? sorted.slice(8) : [];

  const labels = displayed.map((s) => s.name).concat(needsOthers ? ["Others"] : []);
  const values = displayed.map((s) => (assigneeChartMode === "workload" ? s.total : performanceScore(s)));
  let othersValue = 0;
  if (needsOthers) {
    othersValue =
      assigneeChartMode === "workload"
        ? others.reduce((sum, s) => sum + s.total, 0)
        : others.reduce((sum, s) => sum + performanceScore(s), 0);
    values.push(othersValue);
  }

  const meta = new Map();
  for (const s of displayed) meta.set(s.name, s);
  if (needsOthers) meta.set("Others", { name: "Others", members: others });

  return {
    mode: assigneeChartMode,
    memberCount,
    labels,
    values,
    totalValue,
    // Keep readability for larger teams: start scroll earlier when vertical.
    // For >10 we switch to horizontal bars, so scroll is less necessary.
    needsScroll: memberCount > 6 && memberCount <= 12,
    useHorizontal: memberCount > 10,
    meta,
  };
}

function setAssigneeCardVisibility() {
  const barWrap = document.getElementById("assigneeChartWrap");
  const smartWrap = document.getElementById("smartDistributionWrap");
  if (!barWrap || !smartWrap) return;

  const showSmart = assigneeChartMode === "smart";
  barWrap.hidden = showSmart;
  smartWrap.hidden = !showSmart;
}

function destroySmartCharts() {
  if (smartCharts.workloadDonut) {
    smartCharts.workloadDonut.destroy();
    smartCharts.workloadDonut = null;
  }
  if (smartCharts.equalBar) {
    smartCharts.equalBar.destroy();
    smartCharts.equalBar = null;
  }
}

function renderSmartInsight(names, values, equalShare) {
  const el = document.getElementById("smartInsight");
  if (!el) return;
  const diffs = names.map((name, i) => ({ name, value: values[i], diff: values[i] - equalShare }));
  diffs.sort((a, b) => b.diff - a.diff);
  const overloaded = diffs[0];
  const underloaded = diffs[diffs.length - 1];
  const x = Math.round((overloaded.value - equalShare) * 10) / 10;
  el.textContent = `Move ${x} tasks from ${overloaded.name} to ${underloaded.name} to balance the team`;
}

function createSmartDistributionCharts() {
  // Use the exact scenario values provided
  const names = ["Alice", "James", "Sarah", "Charlie", "Bob"];
  const values = [3.3, 2.5, 2.0, 1.7, 1.0];
  const colors = ["#378ADD", "#534AB7", "#1D9E75", "#BA7517", "#888780"];

  const total = values.reduce((s, v) => s + v, 0);
  const equalShare = total / names.length;

  // Donut slice labels (name + percent) + center total text
  const sliceLabelPlugin = {
    id: "sliceLabelPlugin",
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      const meta = chart.getDatasetMeta(0);
      if (!meta?.data) return;
      ctx.save();
      ctx.fillStyle = "#ffffff";
      ctx.font = "12px Inter, sans-serif";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";

      meta.data.forEach((arc, i) => {
        const p = arc.getProps(["x", "y", "startAngle", "endAngle", "innerRadius", "outerRadius"], true);
        const angle = (p.startAngle + p.endAngle) / 2;
        const r = (p.innerRadius + p.outerRadius) / 2;
        const x = p.x + Math.cos(angle) * r;
        const y = p.y + Math.sin(angle) * r;
        const pct = total ? (values[i] / total) * 100 : 0;
        ctx.fillText(`${names[i]} ${pct.toFixed(0)}%`, x, y);
      });
      ctx.restore();
    },
  };

  const centerTextPlugin = {
    id: "centerTextPlugin",
    afterDraw(chart) {
      const { ctx, chartArea } = chart;
      if (!chartArea) return;
      const cx = (chartArea.left + chartArea.right) / 2;
      const cy = (chartArea.top + chartArea.bottom) / 2;
      ctx.save();
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillStyle = "#ffffff";
      ctx.font = "600 16px Inter, sans-serif";
      ctx.fillText(total.toFixed(1), cx, cy - 8);
      ctx.fillStyle = "#b1b1b1";
      ctx.font = "12px Inter, sans-serif";
      ctx.fillText("total tasks", cx, cy + 10);
      ctx.restore();
    },
  };

  const donutCanvas = document.getElementById("smartWorkloadDonut");
  const barCanvas = document.getElementById("smartEqualBar");
  if (!donutCanvas || !barCanvas) return;

  smartCharts.workloadDonut = new Chart(donutCanvas, {
    type: "doughnut",
    data: {
      labels: names,
      datasets: [{ data: values, backgroundColor: colors, borderWidth: 0 }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "65%",
      plugins: { legend: { display: false }, tooltip: { enabled: true } },
    },
    plugins: [sliceLabelPlugin, centerTextPlugin],
  });

  // Vertical dashed red line at equal share
  const dashedLinePlugin = {
    id: "dashedLinePlugin",
    afterDraw(chart) {
      const { ctx, chartArea, scales } = chart;
      if (!chartArea) return;
      const xScale = scales.x;
      if (!xScale) return;
      const x = xScale.getPixelForValue(equalShare);
      ctx.save();
      ctx.setLineDash([6, 5]);
      ctx.strokeStyle = "#df2f4a";
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.moveTo(x, chartArea.top);
      ctx.lineTo(x, chartArea.bottom);
      ctx.stroke();
      ctx.restore();
    },
  };

  smartCharts.equalBar = new Chart(barCanvas, {
    type: "bar",
    data: {
      labels: names,
      datasets: [
        {
          label: "Equal share",
          data: names.map(() => Number(equalShare.toFixed(1))),
          backgroundColor: "#639922",
          borderWidth: 0,
          borderRadius: 6,
          maxBarThickness: 18,
        },
      ],
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { enabled: false } },
      scales: {
        x: { min: 0, max: 4, ticks: { color: getThemeTokens().textDim }, grid: { color: getThemeTokens().gridLine } },
        y: { ticks: { color: getThemeTokens().textDim }, grid: { display: false } },
      },
    },
    plugins: [dashedLinePlugin],
  });

  renderSmartInsight(names, values, equalShare);
}

let kpiMode = "sub"; // "main" | "sub"

function trendFor(value, total, { invert = false } = {}) {
  // Deterministic POC trend: compare ratio vs a baseline (50%).
  // invert=true means lower is better (e.g., overdue).
  const denom = total > 0 ? total : 1;
  const ratio = value / denom;
  const pct = Math.round((ratio - 0.5) * 100);
  const signed = invert ? -pct : pct;
  const dir = signed >= 0 ? "up" : "down";
  const arrow = signed >= 0 ? "↑" : "↓";
  return { pct: Math.abs(signed), dir, arrow, signed };
}

function setKpiSlot(slot, { label, value, delta, deltaDir, arrow }) {
  const slotEl = document.getElementById(`kpiSlot${slot}`);
  const labelEl = document.getElementById(`kpi${slot}Label`);
  const valueEl = document.getElementById(`kpi${slot}Value`);
  const deltaEl = document.getElementById(`kpi${slot}Delta`);
  if (!slotEl || !labelEl || !valueEl || !deltaEl) return;

  if (!label) {
    slotEl.hidden = true;
    return;
  }

  slotEl.hidden = false;
  labelEl.textContent = label;
  valueEl.textContent = String(value ?? 0);
  deltaEl.className = `kpi-delta ${deltaDir || ""}`.trim();
  deltaEl.innerHTML = `<span class="arrow">${arrow || "—"}</span><span>${delta ?? "—"}</span>`;
}

function computeMainTaskKpis(now = new Date()) {
  const scoped = getOrdersMatchingDashboardFilters();
  const created = scoped.length;
  let completed = 0;
  for (const o of scoped) {
    const tools = o.tools || [];
    const allDone = tools.length > 0 && tools.every((t) => t.status === STATUS.Done);
    if (allDone) completed += 1;
  }
  const active = Math.max(0, created - completed);

  const tCreated = trendFor(created, Math.max(1, created));
  const tActive = trendFor(active, Math.max(1, created));
  const tCompleted = trendFor(completed, Math.max(1, created));

  return [
    { label: "Created Main Tasks", value: created, trend: tCreated },
    { label: "Active Main Tasks", value: active, trend: tActive },
    { label: "Completed Main Tasks", value: completed, trend: tCompleted },
  ];
}

function computeSubTaskKpis(now = new Date()) {
  const flat = getFilteredFlatTools();
  const created = flat.length;
  const contributed = flat.filter((t) => (t.assignee || "").trim() === LOGGED_IN_USER_NAME).length;
  const completed = flat.filter((t) => t.status === STATUS.Done).length;
  const active = Math.max(0, created - completed);
  const overdue = flat.filter((t) => {
    const days = daysUntil(t.dueDate, now);
    return days !== null && days < 0 && t.status !== STATUS.Done;
  }).length;

  const tCreated = trendFor(created, Math.max(1, created));
  const tContributed = trendFor(contributed, Math.max(1, created));
  const tActive = trendFor(active, Math.max(1, created));
  const tCompleted = trendFor(completed, Math.max(1, created));
  const tOverdue = trendFor(overdue, Math.max(1, created), { invert: true });

  const items = [
    { label: "Created Sub Tasks", value: created, trend: tCreated },
    { label: "Active Sub Tasks", value: active, trend: tActive },
    { label: "Completed Sub Tasks", value: completed, trend: tCompleted },
    { label: "Overdue Count", value: overdue, trend: tOverdue },
  ];
  // Member view requirement: remove "Created Sub Tasks" KPI card
  if (activeDashboardRole === "member") {
    const withoutCreated = items.filter((it) => it.label !== "Created Sub Tasks");
    // Add "Contributed Tasks" as the 4th KPI (POC: contributed = tasks assigned to logged-in user)
    withoutCreated.push({ label: "Contributed Tasks", value: contributed, trend: tContributed });
    return withoutCreated;
  }
  // Manager/Lead + My Tasks scope only: show Contributed Tasks instead of Created (sub-task KPI path only)
  if (activeDashboardRole === "manager-lead" && managerLeadScope === "mine") {
    return items.map((it) =>
      it.label === "Created Sub Tasks"
        ? { label: "Contributed Tasks", value: contributed, trend: tContributed }
        : it
    );
  }
  return items;
}

function renderKpis() {
  const now = new Date();
  const items = kpiMode === "main" ? computeMainTaskKpis(now) : computeSubTaskKpis(now);

  for (let i = 1; i <= 4; i++) {
    const it = items[i - 1];
    if (!it) {
      setKpiSlot(i, { label: null });
      continue;
    }
    setKpiSlot(i, {
      label: it.label,
      value: it.value,
      delta: `${it.trend.signed >= 0 ? "+" : "-"}${it.trend.pct}%`,
      deltaDir: it.trend.dir,
      arrow: it.trend.arrow,
    });
  }
}

function renderTable(filteredTasks) {
  const tbody = $("tasksTbody");
  tbody.innerHTML = "";

  if (filteredTasks.length === 0) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 6;
    td.style.color = "#b1b1b1";
    td.textContent = "No tasks found for the selected range.";
    tr.appendChild(td);
    tbody.appendChild(tr);
    return;
  }

  const now = new Date();
  for (const t of filteredTasks) {
    const tr = document.createElement("tr");

    const tdId = document.createElement("td");
    tdId.textContent = t.id;

    const tdName = document.createElement("td");
    tdName.textContent = t.name;

    const tdAssignee = document.createElement("td");
    tdAssignee.textContent = t.assignee;

    const tdStatus = document.createElement("td");
    const statusSpan = document.createElement("span");
    statusSpan.className = `status-cell ${statusToClass(t.status)}`;
    statusSpan.textContent = t.status;
    tdStatus.appendChild(statusSpan);

    const tdTimeline = document.createElement("td");
    tdTimeline.textContent = formatShortMonthDay(parseIsoDate(t.dueDate));

    const tdOverdue = document.createElement("td");
    const overdue = isOverdue(t, now);
    tdOverdue.className = overdue ? "overdue-yes" : "overdue-no";
    tdOverdue.textContent = overdue ? "Yes" : "No";

    tr.appendChild(tdId);
    tr.appendChild(tdName);
    tr.appendChild(tdAssignee);
    tr.appendChild(tdStatus);
    tr.appendChild(tdTimeline);
    tr.appendChild(tdOverdue);

    tbody.appendChild(tr);
  }
}

function renderHierarchicalTable(filteredMainTasks) {
  const tbody = $("tasksTbody");
  tbody.innerHTML = "";

  if (filteredMainTasks.length === 0) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 6;
    td.style.color = "#b1b1b1";
    td.textContent = "No tasks found for the selected range.";
    tr.appendChild(td);
    tbody.appendChild(tr);
    return;
  }

  const now = new Date();

  // Sort main tasks by their earliest due date (subtask-level), stable by original order.
  const stableMain = filteredMainTasks.map((m, idx) => ({ m, idx }));
  stableMain.sort((a, b) => {
    const da = earliestDueForMainTask(a.m);
    const db = earliestDueForMainTask(b.m);
    const ta = da ? da.getTime() : Number.POSITIVE_INFINITY;
    const tb = db ? db.getTime() : Number.POSITIVE_INFINITY;
    return ta - tb || a.idx - b.idx;
  });

  for (const { m } of stableMain) {
    const hasSubtasks = (m.subtasks || []).length > 0;

    const mainTr = document.createElement("tr");
    mainTr.className = "table-main-row";

    const tdId = document.createElement("td");
    tdId.textContent = m.id;

    const tdName = document.createElement("td");
    tdName.className = "table-main-title";
    tdName.textContent = m.name;

    const tdAssignee = document.createElement("td");
    tdAssignee.textContent = m.assignee || "—";

    const tdStatus = document.createElement("td");
    const mainStatus = m.status || (hasSubtasks ? "—" : STATUS.Open);
    if (mainStatus !== "—") {
      const statusSpan = document.createElement("span");
      statusSpan.className = `status-cell ${statusToClass(mainStatus)}`;
      statusSpan.textContent = mainStatus;
      tdStatus.appendChild(statusSpan);
    } else {
      tdStatus.style.color = "#b1b1b1";
      tdStatus.textContent = "—";
    }

    const tdTimeline = document.createElement("td");
    if (!hasSubtasks) {
      const d = safeParseDueDate(m.dueDate);
      tdTimeline.textContent = d ? formatShortMonthDay(d) : "—";
      tdTimeline.className = "table-main-timeline";
    } else {
      const d = earliestDueForMainTask(m);
      tdTimeline.textContent = d ? formatShortMonthDay(d) : "—";
      tdTimeline.className = "table-main-timeline";
    }

    const tdOverdue = document.createElement("td");
    let mainOverdue = false;
    if (!hasSubtasks) {
      if (m.dueDate) mainOverdue = isOverdue({ dueDate: m.dueDate, status: m.status || STATUS.Open }, now);
    } else {
      mainOverdue = (m.subtasks || []).some((s) => isOverdue(s, now));
    }
    tdOverdue.className = mainOverdue ? "overdue-yes" : "overdue-no";
    tdOverdue.textContent = mainOverdue ? "Yes" : "No";

    mainTr.appendChild(tdId);
    mainTr.appendChild(tdName);
    mainTr.appendChild(tdAssignee);
    mainTr.appendChild(tdStatus);
    mainTr.appendChild(tdTimeline);
    mainTr.appendChild(tdOverdue);
    tbody.appendChild(mainTr);

    if (!hasSubtasks) continue;

    const stableSubs = (m.subtasks || []).map((s, idx) => ({ s, idx, due: safeParseDueDate(s.dueDate) }));
    stableSubs.sort((a, b) => {
      const ta = a.due ? a.due.getTime() : Number.POSITIVE_INFINITY;
      const tb = b.due ? b.due.getTime() : Number.POSITIVE_INFINITY;
      return ta - tb || a.idx - b.idx; // stable for equal dates
    });

    for (const { s, due } of stableSubs) {
      const tr = document.createElement("tr");
      tr.className = "table-sub-row";

      const sId = document.createElement("td");
      sId.textContent = s.id;

      const sName = document.createElement("td");
      sName.className = "table-sub-title";
      sName.textContent = `↳ ${s.name}`;

      const sAssignee = document.createElement("td");
      sAssignee.textContent = s.assignee || m.assignee || "—";

      const sStatus = document.createElement("td");
      const statusSpan = document.createElement("span");
      statusSpan.className = `status-cell ${statusToClass(s.status)}`;
      statusSpan.textContent = s.status;
      sStatus.appendChild(statusSpan);

      const sTimeline = document.createElement("td");
      sTimeline.textContent = due ? formatShortMonthDay(due) : "—";

      const sOverdue = document.createElement("td");
      const overdue = s.dueDate ? isOverdue(s, now) : false;
      sOverdue.className = overdue ? "overdue-yes" : "overdue-no";
      sOverdue.textContent = overdue ? "Yes" : "No";

      tr.appendChild(sId);
      tr.appendChild(sName);
      tr.appendChild(sAssignee);
      tr.appendChild(sStatus);
      tr.appendChild(sTimeline);
      tr.appendChild(sOverdue);
      tbody.appendChild(tr);
    }
  }
}

let activeGridTab = "order"; // "order" | "task"
let expandedOrderId = null;
let expandedTaskId = null;

const gridFilterState = {
  order: {
    orderId: "",
    orderName: "",
    toolsMin: "",
    toolsMax: "",
    health: "any", // any|healthy|warning|overdue
    createdFrom: "",
    createdTo: "",
    dueFrom: "",
    dueTo: "",
  },
  task: {
    taskName: "",
    parentOrder: "",
    subTaskName: "",
    toolName: "",
    toolAssignee: "",
    dueFrom: "",
    dueTo: "",
    status: "any",
    stage: "any",
  },
};

function renderGridTabs() {
  const tabOrder = document.getElementById("tabOrderView");
  const tabTask = document.getElementById("tabTaskView");
  if (!tabOrder || !tabTask) return;

  tabOrder.classList.toggle("active", activeGridTab === "order");
  tabTask.classList.toggle("active", activeGridTab === "task");
  tabOrder.setAttribute("aria-selected", activeGridTab === "order" ? "true" : "false");
  tabTask.setAttribute("aria-selected", activeGridTab === "task" ? "true" : "false");
}

function statusOptionsHtml(includeAny = true) {
  const opts = [STATUS.Done, STATUS.Working, STATUS.Stuck, STATUS.Testing, STATUS.Open];
  const any = includeAny ? `<option value="any">Any</option>` : "";
  return `${any}${opts.map((s) => `<option value="${s}">${s}</option>`).join("")}`;
}

function renderGridFilters() {
  const root = document.getElementById("gridFilters");
  if (!root) return;

  const order = gridFilterState.order;
  const task = gridFilterState.task;

  if (activeGridTab === "order") {
    root.innerHTML = `
      <div class="filters-row">
        <div class="filter-field">
          <label>Order ID</label>
          <input id="fOrderId" value="${order.orderId}" placeholder="e.g. O-10021" />
        </div>
        <div class="filter-field">
          <label>Order Name</label>
          <input id="fOrderName" value="${order.orderName}" placeholder="Search name" />
        </div>
        <div class="filter-field">
          <label>Tools Count (min)</label>
          <input id="fToolsMin" type="number" min="0" value="${order.toolsMin}" placeholder="0" />
        </div>
        <div class="filter-field">
          <label>Tools Count (max)</label>
          <input id="fToolsMax" type="number" min="0" value="${order.toolsMax}" placeholder="10" />
        </div>
        <div class="filter-field">
          <label>Health Status</label>
          <select id="fHealth">
            <option value="any">Any</option>
            <option value="healthy" ${order.health === "healthy" ? "selected" : ""}>Healthy</option>
            <option value="warning" ${order.health === "warning" ? "selected" : ""}>Near due</option>
            <option value="overdue" ${order.health === "overdue" ? "selected" : ""}>Overdue</option>
          </select>
        </div>
        <div class="filter-field">
          <label>Created From</label>
          <input id="fCreatedFrom" type="date" value="${order.createdFrom}" />
        </div>
        <div class="filter-field">
          <label>Created To</label>
          <input id="fCreatedTo" type="date" value="${order.createdTo}" />
        </div>
        <div class="filter-field">
          <label>Due From</label>
          <input id="fDueFrom" type="date" value="${order.dueFrom}" />
        </div>
        <div class="filter-field">
          <label>Due To</label>
          <input id="fDueTo" type="date" value="${order.dueTo}" />
        </div>
        <div class="filters-actions">
          <button id="fApply" class="btn" type="button">Apply</button>
          <button id="fClear" class="btn secondary" type="button">Clear</button>
        </div>
      </div>
    `;
  } else {
    root.innerHTML = `
      <div class="filters-row">
        <div class="filter-field">
          <label>Task Name</label>
          <input id="fTaskName" value="${task.taskName}" placeholder="Search main task" />
        </div>
        <div class="filter-field">
          <label>Parent Order</label>
          <input id="fParentOrder" value="${task.parentOrder}" placeholder="Order ID or name" />
        </div>
        <div class="filter-field">
          <label>Sub-task Name</label>
          <input id="fSubTaskName" value="${task.subTaskName}" placeholder="Search subtask" />
        </div>
        <div class="filter-field">
          <label>Tool Name</label>
          <input id="fToolName" value="${task.toolName}" placeholder="Search tool" />
        </div>
        <div class="filter-field">
          <label>Tool Assignee</label>
          <input id="fToolAssignee" value="${task.toolAssignee}" placeholder="Search assignee" />
        </div>
        <div class="filter-field">
          <label>Due From</label>
          <input id="fDueFrom" type="date" value="${task.dueFrom}" />
        </div>
        <div class="filter-field">
          <label>Due To</label>
          <input id="fDueTo" type="date" value="${task.dueTo}" />
        </div>
        <div class="filter-field">
          <label>Status</label>
          <select id="fStatus">${statusOptionsHtml(true)}</select>
        </div>
        <div class="filter-field">
          <label>Stage</label>
          <select id="fStage">${statusOptionsHtml(true)}</select>
        </div>
        <div class="filters-actions">
          <button id="fApply" class="btn" type="button">Apply</button>
          <button id="fClear" class="btn secondary" type="button">Clear</button>
        </div>
      </div>
    `;
    const statusSel = root.querySelector("#fStatus");
    const stageSel = root.querySelector("#fStage");
    if (statusSel) statusSel.value = task.status || "any";
    if (stageSel) stageSel.value = task.stage || "any";
  }

  const applyBtn = root.querySelector("#fApply");
  const clearBtn = root.querySelector("#fClear");

  const bind = (id, key, tabKey) => {
    const el = root.querySelector(`#${id}`);
    if (!el) return;
    el.addEventListener("input", () => {
      gridFilterState[tabKey][key] = el.value;
    });
    el.addEventListener("change", () => {
      gridFilterState[tabKey][key] = el.value;
    });
  };

  if (activeGridTab === "order") {
    bind("fOrderId", "orderId", "order");
    bind("fOrderName", "orderName", "order");
    bind("fToolsMin", "toolsMin", "order");
    bind("fToolsMax", "toolsMax", "order");
    bind("fHealth", "health", "order");
    bind("fCreatedFrom", "createdFrom", "order");
    bind("fCreatedTo", "createdTo", "order");
    bind("fDueFrom", "dueFrom", "order");
    bind("fDueTo", "dueTo", "order");
  } else {
    bind("fTaskName", "taskName", "task");
    bind("fParentOrder", "parentOrder", "task");
    bind("fSubTaskName", "subTaskName", "task");
    bind("fToolName", "toolName", "task");
    bind("fToolAssignee", "toolAssignee", "task");
    bind("fDueFrom", "dueFrom", "task");
    bind("fDueTo", "dueTo", "task");
    bind("fStatus", "status", "task");
    bind("fStage", "stage", "task");
  }

  if (applyBtn) applyBtn.addEventListener("click", () => applyAndRender());
  if (clearBtn)
    clearBtn.addEventListener("click", () => {
      if (activeGridTab === "order") {
        gridFilterState.order = { orderId: "", orderName: "", toolsMin: "", toolsMax: "", health: "any", createdFrom: "", createdTo: "", dueFrom: "", dueTo: "" };
        expandedOrderId = null;
      } else {
        gridFilterState.task = {
          taskName: "",
          parentOrder: "",
          subTaskName: "",
          toolName: "",
          toolAssignee: "",
          dueFrom: "",
          dueTo: "",
          status: "any",
          stage: "any",
        };
        expandedTaskId = null;
      }
      applyAndRender();
    });
}

function renderOrderView(range) {
  const root = document.getElementById("gridContent");
  if (!root) return;
  root.innerHTML = "";

  const now = new Date();
  const hasDashFilters = dashboardHasActiveFilters();
  const filteredOrders2 = orders
    .map((o) => {
      const tools = filterToolsForDashboard(o.tools || []);
      return { ...o, tools };
    })
    .filter((o) => !hasDashFilters || o.tools.length > 0);

  const table = document.createElement("table");
  table.className = "grid-table";
  table.innerHTML = `
    <thead>
      <tr>
        <th style="width:45%">Order Name</th>
        <th>Order Type</th>
        <th>Labels</th>
        <th>Tool Count</th>
        <th>Task Health</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;
  const tbody = table.querySelector("tbody");

  for (const o of filteredOrders2) {
    const toolsForHealth = o.tools || [];
    const health = computeHealthFromItems(toolsForHealth, now);
    const toolCount = toolsForHealth.length;

    const tr = document.createElement("tr");
    tr.className = "grid-row-clickable";
    tr.dataset.orderId = o.id;
    tr.innerHTML = `
      <td>${o.name}</td>
      <td><span class="badge">${o.orderType || "—"}</span></td>
      <td>${(o.labels || []).map((l) => `<span class="label-chip" data-label="${l}" style="background:${LABEL_COLORS[l] || "#64748b"}20;border-color:${LABEL_COLORS[l] || "#64748b"}55;color:${LABEL_COLORS[l] || "#64748b"}">${l}</span>`).join(" ") || "—"}</td>
      <td>${toolCount}</td>
      <td><span class="${health.className}">${health.label}</span></td>
    `;

    tr.addEventListener("click", () => {
      expandedOrderId = expandedOrderId === o.id ? null : o.id;
      renderActiveGrid(range);
    });

    tbody.appendChild(tr);

    if (expandedOrderId === o.id) {
      const detailTr = document.createElement("tr");
      const td = document.createElement("td");
      td.colSpan = 3;
      td.className = "subgrid";

      const subTitle = document.createElement("div");
      subTitle.className = "subgrid-title";
      subTitle.textContent = "Order → Sub Task Details";

      const subTable = document.createElement("table");
      subTable.className = "grid-table";
      subTable.innerHTML = `
        <thead>
          <tr>
            <th>S.No</th>
            <th style="width:32%">Tool Name</th>
            <th>Assignee</th>
            <th>Labels</th>
            <th>Due Date</th>
            <th>Board</th>
            <th>Stage</th>
            <th>Workspace</th>
            <th>Health</th>
          </tr>
        </thead>
        <tbody></tbody>
      `;
      const subBody = subTable.querySelector("tbody");

      toolsForHealth.forEach((t, i) => {
        const d = safeParseDueDate(t.dueDate);
        const days = daysUntil(t.dueDate, now);
        const state = dueStateFromDays(days);
        const h = computeHealthFromTool(t, now);

        const subTr = document.createElement("tr");
        subTr.dataset.toolId = t.id;
        subTr.innerHTML = `
          <td>${i + 1}</td>
          <td>${t.name}</td>
          <td class="assignee-cell" data-assignee="${(t.assignee || "—").replace(/"/g, "&quot;")}"><span class="avatar">${String(t.assignee || "—").slice(0, 1).toUpperCase()}</span> ${t.assignee || "—"}</td>
          <td>${(t.labels || o.labels || []).map((l) => `<span class="label-chip" data-label="${l}" style="background:${LABEL_COLORS[l] || "#64748b"}20;border-color:${LABEL_COLORS[l] || "#64748b"}55;color:${LABEL_COLORS[l] || "#64748b"}">${l}</span>`).join(" ") || "—"}</td>
          <td class="${state.className}">${d ? formatShortMonthDay(d) : "—"}</td>
          <td>${t.board || "—"}</td>
          <td>${t.workflowStage || "—"}</td>
          <td>${t.workspace || "—"}</td>
          <td><span class="${h.className}">${h.label}</span></td>
        `;
        subBody.appendChild(subTr);
      });

      td.appendChild(subTitle);
      td.appendChild(subTable);
      detailTr.appendChild(td);
      tbody.appendChild(detailTr);
    }
  }

  if (filteredOrders2.length === 0) {
    const empty = document.createElement("div");
    empty.style.color = "#b1b1b1";
    empty.innerHTML = `No tasks match your filters. <button type="button" id="resetFiltersInlineOrders" class="smart-filter-clear" style="margin-left:8px;">Reset Filters</button>`;
    root.appendChild(empty);
    empty.querySelector("#resetFiltersInlineOrders")?.addEventListener("click", () => {
      clearSmartFilters();
      syncSmartFilterUiFromState();
      applyAndRender();
    });
    return;
  }

  root.appendChild(table);
}

let highlightedToolId = null;

function renderTaskViewFromOrders(range) {
  const root = document.getElementById("gridContent");
  if (!root) return;
  root.innerHTML = "";

  const now = new Date();

  const flat = getFilteredFlatTools();

  const table = document.createElement("table");
  table.className = "grid-table";
  table.innerHTML = `
    <thead>
      <tr>
        <th style="width:22%">Order Name</th>
        <th style="width:26%">Tool Name</th>
        <th>Assignee</th>
        <th>Order Type</th>
        <th>Labels</th>
        <th>Due Date</th>
        <th>Status</th>
        <th>Board</th>
        <th>Stage</th>
        <th>Workspace</th>
        <th>Health</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;
  const tbody = table.querySelector("tbody");

  for (const tool of flat) {
    const d = safeParseDueDate(tool.dueDate);
    const days = daysUntil(tool.dueDate, now);
    const state = dueStateFromDays(days);
    const overdue = days !== null && days < 0 && tool.status !== STATUS.Done;
    const h = computeHealthFromTool(tool, now);

    const tr = document.createElement("tr");
    tr.dataset.toolId = tool.toolId;
    tr.className = "grid-row";
    if (highlightedToolId && highlightedToolId === tool.toolId) tr.classList.add("row-highlight");
    tr.innerHTML = `
      <td>${tool.orderName}</td>
      <td>${tool.toolName}</td>
      <td class="assignee-cell" data-assignee="${(tool.assignee || "—").replace(/"/g, "&quot;")}"><span class="avatar">${String(tool.assignee || "—").slice(0, 1).toUpperCase()}</span> ${tool.assignee || "—"}</td>
      <td><span class="badge">${tool.orderType || "—"}</span></td>
      <td>${(tool.labels || []).map((l) => `<span class="label-chip" data-label="${l}" style="background:${LABEL_COLORS[l] || "#64748b"}20;border-color:${LABEL_COLORS[l] || "#64748b"}55;color:${LABEL_COLORS[l] || "#64748b"}">${l}</span>`).join(" ") || "—"}</td>
      <td class="${state.className}">${d ? formatShortMonthDay(d) : "—"}</td>
      <td class="${overdue ? "overdue-yes" : "overdue-no"}">${tool.status}</td>
      <td>${tool.board || "—"}</td>
      <td>${tool.workflowStage || "—"}</td>
      <td>${tool.workspace || "—"}</td>
      <td><span class="${h.className}">${h.label}</span></td>
    `;
    tbody.appendChild(tr);
  }

  if (flat.length === 0) {
    const empty = document.createElement("div");
    empty.style.color = "#b1b1b1";
    empty.innerHTML = `No tasks match your filters. <button type="button" id="resetFiltersInline" class="smart-filter-clear" style="margin-left:8px;">Reset Filters</button>`;
    root.appendChild(empty);
    empty.querySelector("#resetFiltersInline")?.addEventListener("click", () => {
      clearSmartFilters();
      syncSmartFilterUiFromState();
      applyAndRender();
    });
    return;
  }

  root.appendChild(table);
}

function renderActiveGrid(range, filteredMainTasks = null) {
  renderGridTabs();
  if (activeGridTab === "order") {
    renderOrderView(range);
  } else {
    renderTaskViewFromOrders(range);
  }
}

function clearSmartFilters() {
  smartFilters.assignees.clear();
  smartFilters.orderTypes.clear();
  smartFilters.labels.clear();
  smartFilters.boards.clear();
  smartFilters.stages.clear();
  smartFilters.ageBuckets.clear();
  smartFilters.workspaces.clear();
  smartFilters.assignedToMe = false;
  smartFilters.unassigned = false;
}

function smartFiltersToChips() {
  const chips = [];
  for (const a of smartFilters.assignees) chips.push({ key: "assignee", value: a, label: `Assignee: ${a}` });
  for (const t of smartFilters.orderTypes) chips.push({ key: "orderType", value: t, label: `Order Type: ${t}` });
  for (const l of smartFilters.labels) chips.push({ key: "label", value: l, label: `Label: ${l}` });
  for (const b of smartFilters.boards) chips.push({ key: "board", value: b, label: `Board: ${b}` });
  for (const s of smartFilters.stages) chips.push({ key: "stage", value: s, label: `Stage: ${s}` });
  for (const ab of smartFilters.ageBuckets) {
    const label = AGE_BUCKETS.find((x) => x.key === ab)?.label || ab;
    chips.push({ key: "ageBucket", value: ab, label: `Age: ${label}` });
  }
  if (workspaceFilterAppliesToRole()) for (const w of smartFilters.workspaces) chips.push({ key: "workspace", value: w, label: `Workspace: ${w}` });
  if (smartFilters.assignedToMe) chips.push({ key: "assignedToMe", value: "1", label: "Assigned to Me" });
  if (smartFilters.unassigned) chips.push({ key: "unassigned", value: "1", label: "Unassigned" });
  return chips;
}

function renderActiveFilterChips() {
  const bar = document.getElementById("activeFiltersBar");
  const host = document.getElementById("activeFilterChips");
  const clearBtn = document.getElementById("activeFiltersClear");
  if (!bar || !host || !clearBtn) return;

  const popWrap = document.getElementById("popoverAppliedFilters");
  const popHost = document.getElementById("popoverAppliedFilterChips");

  const chips = smartFiltersToChips();
  const show = chips.length > 0;
  bar.hidden = !show;
  bar.setAttribute("aria-hidden", show ? "false" : "true");
  host.innerHTML = "";

  if (popWrap && popHost) {
    popWrap.hidden = !show;
    popWrap.setAttribute("aria-hidden", show ? "false" : "true");
    popHost.innerHTML = "";
  }

  const makeChipEl = (c) => {
    const el = document.createElement("span");
    el.className = "active-chip";
    el.innerHTML = `<span>${c.label}</span><button type="button" aria-label="Remove filter">×</button>`;
    el.querySelector("button")?.addEventListener("click", () => {
      if (c.key === "assignee") smartFilters.assignees.delete(c.value);
      if (c.key === "orderType") smartFilters.orderTypes.delete(c.value);
      if (c.key === "label") smartFilters.labels.delete(c.value);
      if (c.key === "board") smartFilters.boards.delete(c.value);
      if (c.key === "stage") smartFilters.stages.delete(c.value);
      if (c.key === "workspace") smartFilters.workspaces.delete(c.value);
      if (c.key === "ageBucket") smartFilters.ageBuckets.delete(c.value);
      if (c.key === "assignedToMe") smartFilters.assignedToMe = false;
      if (c.key === "unassigned") smartFilters.unassigned = false;
      syncSmartFilterUiFromState();
      applyAndRender();
    });
    return el;
  };

  for (const c of chips) {
    host.appendChild(makeChipEl(c));
    if (popHost) popHost.appendChild(makeChipEl(c));
  }

  clearBtn.onclick = () => {
    clearSmartFilters();
    syncSmartFilterUiFromState();
    applyAndRender();
  };
}

function makePillButton(text, isActive, onClick) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = `pill-btn${isActive ? " active" : ""}`;
  btn.textContent = text;
  btn.addEventListener("click", onClick);
  return btn;
}

function syncSmartFilterUiFromState() {
  // Update dropdown summary metas
  const setMeta = (id, text) => {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
  };
  setMeta(
    "assigneeSummaryMeta",
    smartFilters.unassigned ? "Unassigned" : smartFilters.assignees.size ? `${smartFilters.assignees.size} selected` : "Any"
  );
  setMeta("orderTypeSummaryMeta", smartFilters.orderTypes.size ? `${smartFilters.orderTypes.size} selected` : "Any");
  setMeta("labelsSummaryMeta", smartFilters.labels.size ? `${smartFilters.labels.size} selected` : "Any");
  setMeta("boardSummaryMeta", smartFilters.boards.size ? `${smartFilters.boards.size} selected` : "Any");
  setMeta("stageSummaryMeta", smartFilters.stages.size ? `${smartFilters.stages.size} selected` : "Any");
  setMeta(
    "workspaceSummaryMeta",
    workspaceFilterAppliesToRole() ? (smartFilters.workspaces.size ? `${smartFilters.workspaces.size} selected` : "Any") : "Hidden"
  );

  // pills
  const pillAreas = [
    ["filterOrderTypePills", ORDER_TYPE_OPTIONS, smartFilters.orderTypes],
    ["filterBoardMulti", BOARD_LIST, smartFilters.boards],
  ];
  for (const [id, opts, set] of pillAreas) {
    const host = document.getElementById(id);
    if (!host) continue;
    host.innerHTML = "";
    for (const o of opts) {
      host.appendChild(
        makePillButton(o, set.has(o), () => {
          if (set.has(o)) set.delete(o);
          else set.add(o);
          // Stage list depends on board selection
          if (id === "filterBoardMulti") renderStagePills();
          renderActiveFilterChips();
          applyAndRender();
        })
      );
    }
  }

  // labels
  const labelsHost = document.getElementById("filterLabelChips");
  if (labelsHost) {
    labelsHost.innerHTML = "";
    for (const l of LABEL_OPTIONS) {
      const chip = document.createElement("span");
      const active = smartFilters.labels.has(l);
      chip.className = "label-chip";
      chip.style.background = `${LABEL_COLORS[l]}20`;
      chip.style.border = `1px solid ${LABEL_COLORS[l]}55`;
      chip.style.color = LABEL_COLORS[l];
      chip.style.opacity = active ? "1" : "0.75";
      chip.textContent = l;
      chip.addEventListener("click", () => {
        if (smartFilters.labels.has(l)) smartFilters.labels.delete(l);
        else smartFilters.labels.add(l);
        renderActiveFilterChips();
        syncSmartFilterUiFromState();
        applyAndRender();
      });
      labelsHost.appendChild(chip);
    }
  }

  // workspace section visibility
  const wsSection = document.getElementById("filterWorkspaceSection");
  if (wsSection) wsSection.hidden = !workspaceFilterAppliesToRole();

  const wsHost = document.getElementById("filterWorkspaceMulti");
  if (wsHost) {
    wsHost.innerHTML = "";
    for (const w of WORKSPACE_OPTIONS) {
      wsHost.appendChild(
        makePillButton(w, smartFilters.workspaces.has(w), () => {
          if (smartFilters.workspaces.has(w)) smartFilters.workspaces.delete(w);
          else smartFilters.workspaces.add(w);
          renderActiveFilterChips();
          applyAndRender();
        })
      );
    }
  }

  // assignee list (search applied elsewhere)
  renderAssigneeList();
  renderActiveFilterChips();
}

function getStageOptionsForCurrentBoardSelection() {
  const selectedBoards = [...smartFilters.boards];
  const boards = selectedBoards.length ? selectedBoards : BOARD_LIST;
  const set = new Set();
  for (const b of boards) for (const s of STAGES_BY_BOARD[b] || []) set.add(s);
  return [...set];
}

function renderStagePills() {
  const host = document.getElementById("filterStageMulti");
  if (!host) return;
  const options = getStageOptionsForCurrentBoardSelection().sort((a, b) => a.localeCompare(b));
  host.innerHTML = "";
  for (const s of options) {
    host.appendChild(
      makePillButton(s, smartFilters.stages.has(s), () => {
        if (smartFilters.stages.has(s)) smartFilters.stages.delete(s);
        else smartFilters.stages.add(s);
        renderActiveFilterChips();
        applyAndRender();
      })
    );
  }
}

function renderAssigneeList() {
  const host = document.getElementById("filterAssigneeList");
  if (!host) return;
  const q = (document.getElementById("filterAssigneeSearch")?.value || "").trim().toLowerCase();
  host.innerHTML = "";
  const list = ASSIGNEES.filter((a) => !q || a.toLowerCase().includes(q));
  for (const a of list) {
    const row = document.createElement("label");
    row.className = "smart-filter-item";
    const checked = smartFilters.assignees.has(a);
    row.innerHTML = `<input type="checkbox" ${checked ? "checked" : ""} /> <span class="avatar">${a.slice(0, 1).toUpperCase()}</span> <span>${a}</span>`;
    row.querySelector("input")?.addEventListener("change", (e) => {
      const isChecked = Boolean(e.target?.checked);
      if (isChecked) smartFilters.assignees.add(a);
      else smartFilters.assignees.delete(a);
      smartFilters.unassigned = false;
      renderActiveFilterChips();
      applyAndRender();
    });
    host.appendChild(row);
  }
}

function initSmartFilterPopover() {
  const btn = document.getElementById("smartFilterBtn");
  const pop = document.getElementById("smartFilterPopover");
  const closeBtn = document.getElementById("smartFilterCloseBtn");
  if (!btn || !pop) return;

  const open = () => {
    pop.hidden = false;
    pop.setAttribute("aria-hidden", "false");
    btn.setAttribute("aria-expanded", "true");
    syncSmartFilterUiFromState();
  };
  const close = () => {
    pop.hidden = true;
    pop.setAttribute("aria-hidden", "true");
    btn.setAttribute("aria-expanded", "false");
  };

  btn.addEventListener("click", () => (pop.hidden ? open() : close()));
  closeBtn?.addEventListener("click", close);

  document.addEventListener("click", (e) => {
    if (pop.hidden) return;
    const target = e.target;
    if (target === btn || btn.contains(target)) return;
    if (pop.contains(target)) return;
    close();
  });

  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;
    if (pop.hidden) return;
    close();
  });

  document.getElementById("filterAssigneeSearch")?.addEventListener("input", () => renderAssigneeList());
  document.getElementById("clearAllFiltersBtn")?.addEventListener("click", () => {
    clearSmartFilters();
    syncSmartFilterUiFromState();
    applyAndRender();
  });
  document.getElementById("filterAssignedToMe")?.addEventListener("click", () => {
    smartFilters.assignedToMe = !smartFilters.assignedToMe;
    // POC: no real user context -> treat as filtering to Alice
    if (smartFilters.assignedToMe) smartFilters.assignees.add("Alice");
    else smartFilters.assignees.delete("Alice");
    renderActiveFilterChips();
    syncSmartFilterUiFromState();
    applyAndRender();
  });
  document.getElementById("filterUnassigned")?.addEventListener("click", () => {
    smartFilters.unassigned = !smartFilters.unassigned;
    if (smartFilters.unassigned) smartFilters.assignees.clear();
    renderActiveFilterChips();
    syncSmartFilterUiFromState();
    applyAndRender();
  });

  renderStagePills();
  syncSmartFilterUiFromState();
}

function applyLabelFilter(label) {
  if (!label) return;
  smartFilters.labels.add(label);
  renderActiveFilterChips();
  syncSmartFilterUiFromState();
  applyAndRender();
}

function applyAssigneeFilter(assignee) {
  if (!assignee) return;
  smartFilters.assignees.add(assignee);
  smartFilters.unassigned = false;
  renderActiveFilterChips();
  syncSmartFilterUiFromState();
  applyAndRender();
}

function csvEscape(value) {
  const s = String(value ?? "");
  if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function tableToCsv(table) {
  const lines = [];
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  if (thead) {
    const ths = Array.from(thead.querySelectorAll("th")).map((th) => csvEscape(th.textContent.trim()));
    if (ths.length) lines.push(ths.join(","));
  }

  if (tbody) {
    const rows = Array.from(tbody.querySelectorAll("tr"));
    for (const tr of rows) {
      const tds = Array.from(tr.querySelectorAll("td"));
      if (tds.length === 0) continue;
      // Skip wrapper rows that just contain an embedded sub-table (colspan)
      const hasNestedTable = tr.querySelector("table");
      if (hasNestedTable) continue;
      lines.push(tds.map((td) => csvEscape(td.textContent.trim())).join(","));
    }
  }

  return lines.join("\n");
}

function exportCurrentGridToCsv() {
  const grid = document.getElementById("gridContent");
  if (!grid) return;

  const tables = Array.from(grid.querySelectorAll("table"));
  if (tables.length === 0) return;

  const parts = [];
  // Export the main grid table first
  parts.push(tableToCsv(tables[0]));

  // If an expanded subgrid exists, export its sub-table too
  const subTable = grid.querySelector(".subgrid table");
  if (subTable) {
    parts.push(""); // blank line
    parts.push("Order Sub Task Details");
    parts.push(tableToCsv(subTable));
  }

  const csv = parts.filter((p) => p !== null && p !== undefined).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  const today = new Date();
  const stamp = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-${String(today.getDate()).padStart(2, "0")}`;
  a.href = url;
  a.download = `grid-export-${activeGridTab}-${stamp}.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();

  setTimeout(() => URL.revokeObjectURL(url), 0);
}

function deadlineBadgeClass(dueDate, now = new Date()) {
  const today = clampToLocalMidnight(now);
  const due = clampToLocalMidnight(parseIsoDate(dueDate));
  const diffDays = Math.round((due.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));

  if (diffDays <= 1) return "urgent";
  if (diffDays <= 3) return "warning";
  return "normal";
}

function renderDeadlines(filteredTasks) {
  // Render is now handled by renderGroupedDeadlines()
  // This function remains for compatibility with previous calls.
  renderGroupedDeadlines(filteredTasks);
}

function earliestDueForMainTask(mainTask) {
  const candidates = [];
  const mainDue = safeParseDueDate(mainTask.dueDate);
  if (mainDue) candidates.push({ due: mainDue, kind: "main", index: 0 });

  const subs = mainTask.subtasks || [];
  for (let i = 0; i < subs.length; i++) {
    const d = safeParseDueDate(subs[i].dueDate);
    if (!d) continue;
    candidates.push({ due: d, kind: "sub", index: i });
  }
  if (candidates.length === 0) return null;

  candidates.sort((a, b) => a.due.getTime() - b.due.getTime() || a.kind.localeCompare(b.kind) || a.index - b.index);
  return candidates[0].due;
}

function findNearestUpcomingDueDate(filteredMainTasks, now = new Date()) {
  const today = clampToLocalMidnight(now).getTime();
  let best = null;

  for (const m of filteredMainTasks) {
    const subs = m.subtasks || [];
    if (subs.length === 0) {
      const d = safeParseDueDate(m.dueDate);
      if (!d) continue;
      const t = clampToLocalMidnight(d).getTime();
      if (t < today) continue;
      if (!best || t < best.getTime()) best = new Date(t);
      continue;
    }

    for (const s of subs) {
      const d = safeParseDueDate(s.dueDate);
      if (!d) continue;
      const t = clampToLocalMidnight(d).getTime();
      if (t < today) continue;
      if (!best || t < best.getTime()) best = new Date(t);
    }
  }

  return best;
}

function renderGroupedDeadlines(filteredMainTasks) {
  const list = $("deadlineList");
  list.innerHTML = "";

  const now = new Date();
  const nearestUpcoming = findNearestUpcomingDueDate(filteredMainTasks, now);
  const nearestUpcomingTime = nearestUpcoming ? nearestUpcoming.getTime() : null;

  // Sort main tasks by earliest due among subtasks (or main due if no subtasks)
  const sortedMain = [...filteredMainTasks].sort((a, b) => {
    const da = earliestDueForMainTask(a);
    const db = earliestDueForMainTask(b);
    const ta = da ? da.getTime() : Number.POSITIVE_INFINITY;
    const tb = db ? db.getTime() : Number.POSITIVE_INFINITY;
    return ta - tb;
  });

  if (sortedMain.length === 0) {
    const li = document.createElement("li");
    li.style.color = "#b1b1b1";
    li.textContent = "No upcoming deadlines in this range.";
    list.appendChild(li);
    return;
  }

  // Persist user expand/collapse choices across re-renders.
  if (!renderGroupedDeadlines.expandedIds) renderGroupedDeadlines.expandedIds = new Set();
  const expandedIds = renderGroupedDeadlines.expandedIds;

  for (const m of sortedMain) {
    const groupLi = document.createElement("li");
    groupLi.className = "deadline-group";

    const header = document.createElement("div");
    header.className = "deadline-main";
    header.setAttribute("role", "button");
    header.tabIndex = 0;

    const left = document.createElement("div");
    left.className = "deadline-main-left";

    const title = document.createElement("div");
    title.className = "deadline-main-title";
    title.textContent = m.name;

    const count = document.createElement("div");
    count.className = "deadline-main-count";
    const subCount = (m.subtasks || []).length;
    count.textContent = subCount > 0 ? `${subCount} tasks` : "0 tasks";

    left.appendChild(title);
    left.appendChild(count);

    const right = document.createElement("div");
    right.className = "deadline-main-right";

    const mainDue = safeParseDueDate(m.dueDate);
    if (subCount === 0 && mainDue) {
      const dueBadge = document.createElement("span");
      dueBadge.className = `deadline ${deadlineBadgeClass(m.dueDate, now)}`;
      dueBadge.textContent = formatShortMonthDay(mainDue);
      right.appendChild(dueBadge);
    } else {
      const placeholder = document.createElement("span");
      placeholder.className = "deadline-main-placeholder";
      placeholder.textContent = subCount > 0 ? "" : "—";
      right.appendChild(placeholder);
    }

    header.appendChild(left);
    header.appendChild(right);

    const subsWrap = document.createElement("div");
    subsWrap.className = "deadline-subtasks";

    const subs = (m.subtasks || []).map((s, index) => ({ ...s, __index: index, __due: safeParseDueDate(s.dueDate) }));
    subs.sort((a, b) => {
      const ta = a.__due ? a.__due.getTime() : Number.POSITIVE_INFINITY;
      const tb = b.__due ? b.__due.getTime() : Number.POSITIVE_INFINITY;
      return ta - tb || a.__index - b.__index; // stable for equal dates
    });

    for (const s of subs) {
      const row = document.createElement("div");
      row.className = "deadline-subtask";

      const sLeft = document.createElement("div");
      sLeft.className = "deadline-subtask-name";
      sLeft.textContent = `↳ ${s.name}`;

      const sRight = document.createElement("div");
      sRight.className = "deadline-subtask-due";

      const dueBadge = document.createElement("span");
      dueBadge.className = `deadline ${deadlineBadgeClass(s.dueDate, now)}`;
      dueBadge.textContent = s.__due ? formatShortMonthDay(s.__due) : "—";

      if (nearestUpcomingTime && s.__due && clampToLocalMidnight(s.__due).getTime() === nearestUpcomingTime) {
        row.classList.add("deadline-nearest");
      }

      sRight.appendChild(dueBadge);
      row.appendChild(sLeft);
      row.appendChild(sRight);
      subsWrap.appendChild(row);
    }

    // Collapsible groups: default collapsed (user decides when to expand)
    const isExpanded = expandedIds.has(m.id);
    subsWrap.style.display = isExpanded ? "block" : "none";
    header.setAttribute("aria-expanded", isExpanded ? "true" : "false");

    const toggle = () => {
      const currentlyExpanded = expandedIds.has(m.id);
      const nextExpanded = !currentlyExpanded;
      if (nextExpanded) expandedIds.add(m.id);
      else expandedIds.delete(m.id);

      subsWrap.style.display = nextExpanded ? "block" : "none";
      header.setAttribute("aria-expanded", nextExpanded ? "true" : "false");
    };

    header.addEventListener("click", toggle);
    header.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        toggle();
      }
    });

    groupLi.appendChild(header);
    groupLi.appendChild(subsWrap);
    list.appendChild(groupLi);
  }
}

const DEADLINE_CARD_SVG = {
  building:
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="2" width="16" height="20" rx="1"/><path d="M8 6h2M14 6h2M8 10h2M14 10h2M8 14h2M14 14h2M8 18h8"/></svg>',
  clock:
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>',
  doc:
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M7 3h8l4 4v14a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z"/><path d="M14 3v4h4M8 13h8M8 17h6"/></svg>',
  check:
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22a10 10 0 1 0 0-20 10 10 0 0 0 0 20z"/><path d="M8 12l2.5 2.5L16 9"/></svg>',
};

function computeOrderDeadlineSummary(order, now = new Date()) {
  const tools = order.tools || [];
  const total = tools.length;
  let completed = 0;
  let overdueCount = 0;
  let earliestPendingMs = Number.POSITIVE_INFINITY;
  let earliestPendingIso = null;

  for (const t of tools) {
    if (t.status === STATUS.Done) {
      completed += 1;
      continue;
    }
    const due = safeParseDueDate(t.dueDate);
    if (due) {
      const ms = due.getTime();
      if (ms < earliestPendingMs) {
        earliestPendingMs = ms;
        earliestPendingIso = t.dueDate || null;
      }
    }
    const days = daysUntil(t.dueDate, now);
    if (days !== null && days < 0) overdueCount += 1;
  }

  const pendingCount = total - completed;
  return {
    overdueCount,
    total,
    completed,
    pendingCount,
    earliestPendingMs: earliestPendingMs === Number.POSITIVE_INFINITY ? null : earliestPendingMs,
    earliestPendingIso,
  };
}

function monthAbbrevUpper(date) {
  return date.toLocaleString(undefined, { month: "short" }).toUpperCase();
}

function badgeStateForDueIso(dueIso, now = new Date()) {
  const days = daysUntil(dueIso, now);
  if (days === null) return { key: "neutral", className: "deadline-date-badge--neutral", urgent: false };
  if (days < 0 || days <= 1) return { key: "urgent", className: "deadline-date-badge--urgent", urgent: true };
  if (days <= 3) return { key: "warning", className: "deadline-date-badge--warning", urgent: false };
  return { key: "neutral", className: "deadline-date-badge--neutral", urgent: false };
}

function setDateBadge(el, dueIso, now = new Date()) {
  if (!el) return;
  const monthEl = el.querySelector(".deadline-date-badge__month");
  const dayEl = el.querySelector(".deadline-date-badge__day");

  const d = safeParseDueDate(dueIso);
  if (!d || !monthEl || !dayEl) {
    if (monthEl) monthEl.textContent = "—";
    if (dayEl) dayEl.textContent = "—";
    el.classList.remove("deadline-date-badge--urgent", "deadline-date-badge--warning", "deadline-date-badge--neutral");
    el.classList.add("deadline-date-badge--neutral");
    return;
  }

  monthEl.textContent = monthAbbrevUpper(d);
  dayEl.textContent = String(d.getDate()).padStart(2, "0");
  const st = badgeStateForDueIso(dueIso, now);
  el.classList.remove("deadline-date-badge--urgent", "deadline-date-badge--warning", "deadline-date-badge--neutral");
  el.classList.add(st.className);
}

function sortOrdersForDeadlineWidget(ordersList, now = new Date()) {
  return [...ordersList].sort((a, b) => {
    const sa = computeOrderDeadlineSummary(a, now);
    const sb = computeOrderDeadlineSummary(b, now);
    if (sb.overdueCount !== sa.overdueCount) return sb.overdueCount - sa.overdueCount;
    if (sa.pendingCount === 0 && sb.pendingCount === 0) return a.name.localeCompare(b.name);
    if (sa.pendingCount === 0) return 1;
    if (sb.pendingCount === 0) return -1;
    const ta = sa.earliestPendingMs ?? Number.POSITIVE_INFINITY;
    const tb = sb.earliestPendingMs ?? Number.POSITIVE_INFINITY;
    if (ta !== tb) return ta - tb;
    return a.name.localeCompare(b.name);
  });
}

function makeDeadlineMetaChip(svgHtml, text) {
  const chip = document.createElement("span");
  chip.className = "deadline-card__chip";
  const iconHolder = document.createElement("span");
  iconHolder.innerHTML = svgHtml;
  const label = document.createElement("span");
  label.textContent = text;
  chip.append(iconHolder.firstElementChild || iconHolder, label);
  return chip;
}

function closeDeadlineModal() {
  const root = document.getElementById("deadlineModal");
  if (!root) return;
  root.hidden = true;
  root.setAttribute("aria-hidden", "true");
}

function renderDeadlineModalOrderSummary(order, now = new Date()) {
  const titleEl = document.getElementById("deadlineModalTitle");
  const summaryEl = document.getElementById("deadlineModalSummary");
  const badgeEl = document.getElementById("deadlineModalDateBadge");
  if (!titleEl || !summaryEl) return;

  const s = computeOrderDeadlineSummary(order, now);
  titleEl.textContent = order.name;
  setDateBadge(badgeEl, s.earliestPendingIso, now);

  summaryEl.innerHTML = "";
  const p1 = document.createElement("p");
  p1.textContent = `${s.total} total tools · ${s.completed} completed · ${s.overdueCount} overdue`;
  const p2 = document.createElement("p");
  p2.textContent =
    s.pendingCount === 0
      ? "All tools are complete."
      : `${s.pendingCount} tool${s.pendingCount === 1 ? "" : "s"} still in progress.`;
  summaryEl.append(p1, p2);
}

function renderDeadlineModalToolsForOrder(order, now = new Date()) {
  const tbody = document.getElementById("deadlineModalToolsBody");
  if (!tbody) return;

  tbody.innerHTML = "";
  const tools = [...(order.tools || [])].sort((a, b) => {
    const da = safeParseDueDate(a.dueDate);
    const db = safeParseDueDate(b.dueDate);
    const ta = da ? da.getTime() : Number.POSITIVE_INFINITY;
    const tb = db ? db.getTime() : Number.POSITIVE_INFINITY;
    return ta - tb;
  });

  for (const t of tools) {
    const tr = document.createElement("tr");
    const days = daysUntil(t.dueDate, now);
    const overdue = days !== null && days < 0 && t.status !== STATUS.Done;
    if (overdue) tr.classList.add("deadline-modal__row--overdue");

    const tdName = document.createElement("td");
    tdName.textContent = t.name;

    const tdDue = document.createElement("td");
    const d = safeParseDueDate(t.dueDate);
    tdDue.textContent = d ? formatShortMonthDay(d) : "—";

    const tdStatus = document.createElement("td");
    tdStatus.textContent = t.status;

    const tdStage = document.createElement("td");
    tdStage.textContent = t.stage || "—";

    tr.append(tdName, tdDue, tdStatus, tdStage);
    tr.addEventListener("click", () => {
      closeDeadlineModal();
      highlightedToolId = t.id;
      activeGridTab = "task";
      expandedOrderId = null;
      renderActiveGrid({ start: null, end: null });
      requestAnimationFrame(() => {
        const gridRow = document.querySelector(`tr[data-tool-id="${CSS.escape(t.id)}"]`);
        if (gridRow?.scrollIntoView) gridRow.scrollIntoView({ block: "center", behavior: "smooth" });
      });
    });
    tbody.appendChild(tr);
  }
}

function renderDeadlineModalOrdersList(selectedOrderId, now = new Date()) {
  const tbody = document.getElementById("deadlineModalOrdersBody");
  if (!tbody) return;

  tbody.innerHTML = "";
  for (const order of orders) {
    const s = computeOrderDeadlineSummary(order, now);
    const tr = document.createElement("tr");
    tr.dataset.orderId = order.id;
    if (order.id === selectedOrderId) tr.classList.add("deadline-modal__order-row--active");
    if (s.overdueCount > 0) tr.classList.add("deadline-modal__order-row--risk");

    const tdName = document.createElement("td");
    tdName.textContent = order.name;

    const tdOverdue = document.createElement("td");
    tdOverdue.textContent = s.overdueCount === 0 ? "0 Overdue" : `${s.overdueCount} Overdue`;

    tr.append(tdName, tdOverdue);
    tr.addEventListener("click", () => {
      // Update active row styling (without re-rendering the whole list)
      tbody.querySelectorAll(".deadline-modal__order-row--active").forEach((row) => row.classList.remove("deadline-modal__order-row--active"));
      tr.classList.add("deadline-modal__order-row--active");
      renderDeadlineModalOrderSummary(order, now);
      renderDeadlineModalToolsForOrder(order, now);
    });

    tbody.appendChild(tr);
  }
}

function openDeadlineModal(order, now = new Date()) {
  const root = document.getElementById("deadlineModal");
  if (!root) return;

  renderDeadlineModalOrdersList(order.id, now);
  renderDeadlineModalOrderSummary(order, now);
  renderDeadlineModalToolsForOrder(order, now);

  root.hidden = false;
  root.setAttribute("aria-hidden", "false");
  document.getElementById("deadlineModalClose")?.focus();
}

function initDeadlineModal() {
  const root = document.getElementById("deadlineModal");
  if (!root) return;

  document.getElementById("deadlineModalClose")?.addEventListener("click", closeDeadlineModal);
  root.querySelectorAll("[data-deadline-modal-dismiss]").forEach((el) => {
    el.addEventListener("click", closeDeadlineModal);
  });

  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;
    if (root.hidden) return;
    closeDeadlineModal();
  });
}

function renderUpcomingDeadlinesFromOrders() {
  const list = $("deadlineList");
  list.innerHTML = "";

  const now = new Date();
  const range = getActiveDateRangeFromUi();
  const startIso = range.start ? toDateInputValue(range.start) : null;
  const endIso = range.end ? toDateInputValue(new Date(range.end.getFullYear(), range.end.getMonth(), range.end.getDate())) : null;
  const scopedOrders = orders
    .map((o) => ({
      ...o,
      tools: (o.tools || []).filter((t) => toolMatchesDashboardFilters(t) && inDateRange(t.dueDate, startIso, endIso)),
    }))
    .filter((o) => (o.tools || []).length > 0);
  const sorted = sortOrdersForDeadlineWidget(scopedOrders, now);
  const top = sorted.slice(0, 4);

  if (scopedOrders.length === 0) {
    const empty = document.createElement("p");
    empty.className = "deadline-card-empty";
    empty.textContent = "No orders to show for the selected filters.";
    list.appendChild(empty);
    return;
  }

  for (const order of top) {
    const s = computeOrderDeadlineSummary(order, now);
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `deadline-card${s.overdueCount > 0 ? " deadline-card--risk" : ""}`;
    btn.setAttribute("role", "listitem");

    const badge = document.createElement("div");
    badge.className = "deadline-date-badge deadline-date-badge--neutral";
    badge.innerHTML = '<div class="deadline-date-badge__month">—</div><div class="deadline-date-badge__day">—</div>';
    setDateBadge(badge, s.earliestPendingIso, now);

    const body = document.createElement("div");
    body.className = "deadline-card__body";

    const header = document.createElement("div");
    header.className = "deadline-card__header";

    const title = document.createElement("div");
    title.className = "deadline-card__title";
    title.textContent = order.name;

    const st = badgeStateForDueIso(s.earliestPendingIso, now);
    const tag = document.createElement("div");
    tag.className = "deadline-card__tag";
    tag.textContent = st.urgent || s.overdueCount > 0 ? "URGENT" : "";

    const meta = document.createElement("div");
    meta.className = "deadline-card__meta";
    const overdueLabel = s.overdueCount === 1 ? "1 Overdue" : `${s.overdueCount} Overdue`;

    const subline = document.createElement("div");
    subline.className = "deadline-card__subline";
    subline.textContent = `Total tool ${s.total}/${s.completed}`;
    meta.append(
      makeDeadlineMetaChip(DEADLINE_CARD_SVG.clock, s.overdueCount === 0 ? "0 Overdue" : overdueLabel),
      makeDeadlineMetaChip(DEADLINE_CARD_SVG.doc, `${s.total} Total tasks`),
      makeDeadlineMetaChip(DEADLINE_CARD_SVG.check, `${s.completed} Completed`)
    );

    header.append(title, tag);
    body.append(header, subline, meta);
    btn.append(badge, body);
    btn.addEventListener("click", () => openDeadlineModal(order, now));
    list.appendChild(btn);
  }
}

function buildAggregates(filteredTasks) {
  const statusCounts = {
    [STATUS.Done]: 0,
    [STATUS.Working]: 0,
    [STATUS.Stuck]: 0,
    [STATUS.Testing]: 0,
    [STATUS.Open]: 0,
  };

  const assigneeCounts = {};
  for (const t of filteredTasks) {
    statusCounts[t.status] = (statusCounts[t.status] || 0) + 1;
    assigneeCounts[t.assignee] = (assigneeCounts[t.assignee] || 0) + 1;
  }

  const overdueCount = filteredTasks.filter((t) => isOverdue(t)).length;
  const stuckCount = filteredTasks.filter((t) => t.status === STATUS.Stuck).length;
  const doneCount = filteredTasks.filter((t) => t.status === STATUS.Done).length;

  // Basic health segmentation for POC
  const healthy = Math.max(0, doneCount);
  const atRisk = Math.max(0, overdueCount + stuckCount);
  const warning = Math.max(0, filteredTasks.length - healthy - atRisk);

  const assigneeModel = buildAssigneeChartModel(filteredTasks);
  const stageDistribution = buildStageDistributionModel(filteredTasks);
  return { statusCounts, assigneeCounts, assigneeModel, health: { healthy, warning, atRisk }, stageDistribution };
}

function buildBoardHealthKpisFromFilteredFlatTools(flatTools, now = new Date()) {
  // "Board" here maps to the tool.board field (On-boarding / Production / Analyst / Dividends)
  const byBoard = new Map();
  for (const t of flatTools) {
    const b = t.board || "—";
    if (!byBoard.has(b)) byBoard.set(b, []);
    byBoard.get(b).push(t);
  }

  let healthy = 0;
  let needsAttention = 0;
  let atRisk = 0;
  for (const [board, items] of byBoard.entries()) {
    if (board === "—") continue;
    const h = computeHealthFromItems(items, now);
    if (h.key === "overdue") atRisk += 1;
    else if (h.key === "warning") needsAttention += 1;
    else healthy += 1;
  }
  const total = healthy + needsAttention + atRisk;
  return { total, healthy, needsAttention, atRisk };
}

function renderGridHealthKpis() {
  const wrap = document.getElementById("gridHealthKpis");
  const totalEl = document.getElementById("kpiTotalBoards");
  const healthyEl = document.getElementById("kpiHealthyBoards");
  const warnEl = document.getElementById("kpiNeedsAttentionBoards");
  const riskEl = document.getElementById("kpiAtRiskBoards");
  if (!wrap || !totalEl || !healthyEl || !warnEl || !riskEl) return;

  // Only show this KPI strip in Order and Tools tab (per requirement to replace the chart there)
  const show = document.body.dataset.primaryTab === "order-tools";
  wrap.hidden = !show;
  wrap.setAttribute("aria-hidden", show ? "false" : "true");
  if (!show) return;

  const now = new Date();
  const flat = getFilteredFlatTools();
  const k = buildBoardHealthKpisFromFilteredFlatTools(flat, now);
  totalEl.textContent = String(k.total);
  healthyEl.textContent = String(k.healthy);
  warnEl.textContent = String(k.needsAttention);
  riskEl.textContent = String(k.atRisk);
}

function buildStageDistributionModel(tasks) {
  const palette = ["#00c875", "#fdab3d", "#579bfc", "#a855f7", "#f472b6", "#14b8a6", "#f59e0b", "#64748b", "#df2f4a"];
  const counts = {};
  for (const t of tasks) {
    const st = t.workflowStage || "—";
    counts[st] = (counts[st] || 0) + 1;
  }
  // Sort by count desc (then name asc for stability)
  const labels = Object.keys(counts).sort((a, b) => (counts[b] || 0) - (counts[a] || 0) || a.localeCompare(b));
  const values = labels.map((l) => counts[l] || 0);
  const colors = labels.map((_, i) => palette[i % palette.length]);
  return { labels, values, colors };
}

const AGE_BUCKETS = [
  { key: "0-2", label: "0–2 Days (New)", min: 0, max: 2 },
  { key: "3-5", label: "3–5 Days (Early)", min: 3, max: 5 },
  { key: "6-10", label: "6–10 Days (Aging)", min: 6, max: 10 },
  { key: "11-20", label: "11–20 Days (Delayed)", min: 11, max: 20 },
  { key: "21+", label: "21+ Days (Critical)", min: 21, max: Number.POSITIVE_INFINITY },
];

let stageChartViewMode = "stage"; // "stage" | "age" | "stage_age" | "stage_aging"
let ageBasisMode = "created"; // "created" | "stageEntered"

function computeAgeDaysFromTool(tool, now = new Date(), basisOverride = null) {
  const basis = basisOverride || ageBasisMode;
  const iso = basis === "stageEntered" ? tool.stageEnteredDate : tool.createdDate;
  if (!iso) return null;
  const d = safeParseDueDate(iso);
  if (!d) return null;
  const start = clampToLocalMidnight(d).getTime();
  const today = clampToLocalMidnight(now).getTime();
  const diff = today - start;
  return Math.max(0, Math.round(diff / (1000 * 60 * 60 * 24)));
}

function ageBucketForDays(days) {
  if (days === null) return null;
  for (const b of AGE_BUCKETS) {
    if (days >= b.min && days <= b.max) return b.key;
  }
  return AGE_BUCKETS[AGE_BUCKETS.length - 1].key;
}

function buildAgeBucketModel(tasks) {
  const counts = {};
  const now = new Date();
  for (const b of AGE_BUCKETS) counts[b.key] = 0;
  for (const t of tasks) {
    const days = computeAgeDaysFromTool(t, now);
    const bucket = ageBucketForDays(days);
    if (!bucket) continue;
    counts[bucket] += 1;
  }
  const labels = AGE_BUCKETS.map((b) => b.label);
  const values = AGE_BUCKETS.map((b) => counts[b.key] || 0);
  const colors = ["#579bfc", "#00c875", "#fdab3d", "#f59e0b", "#df2f4a"];
  return { labels, values, colors, keys: AGE_BUCKETS.map((b) => b.key) };
}

function buildStageByAgeStackedModel(tasks, basisOverride = null) {
  // Y-axis: stages, X-axis: task count, stacked by age buckets.
  const now = new Date();
  const stageCounts = {};
  for (const t of tasks) {
    const st = t.workflowStage || "—";
    if (!stageCounts[st]) stageCounts[st] = {};
    const days = computeAgeDaysFromTool(t, now, basisOverride);
    const bucket = ageBucketForDays(days);
    if (!bucket) continue;
    stageCounts[st][bucket] = (stageCounts[st][bucket] || 0) + 1;
  }
  const stages = Object.keys(stageCounts).sort((a, b) => {
    const ta = Object.values(stageCounts[a] || {}).reduce((s, n) => s + Number(n || 0), 0);
    const tb = Object.values(stageCounts[b] || {}).reduce((s, n) => s + Number(n || 0), 0);
    return tb - ta || a.localeCompare(b);
  });

  const bucketColors = {
    "0-2": "#579bfc",
    "3-5": "#00c875",
    "6-10": "#fdab3d",
    "11-20": "#f59e0b",
    "21+": "#df2f4a",
  };

  const datasets = AGE_BUCKETS.map((b) => ({
    label: b.label,
    data: stages.map((s) => stageCounts[s]?.[b.key] || 0),
    backgroundColor: bucketColors[b.key] || "#64748b",
    borderRadius: 4,
    maxBarThickness: 36,
  }));

  return { labels: stages.length ? stages : ["—"], datasets };
}

function getStageChartModelFromAggregatesAndTasks(aggregates, filteredTasks) {
  if (stageChartViewMode === "age") {
    const m = buildAgeBucketModel(filteredTasks);
    return { kind: "age", labels: m.labels, datasets: [{ label: "Tasks", data: m.values, backgroundColor: m.colors, borderRadius: 6, maxBarThickness: 36 }] };
  }
  if (stageChartViewMode === "stage_aging") {
    // Force stage-based aging (Stage Entered Date)
    const m = buildStageByAgeStackedModel(filteredTasks, "stageEntered");
    return { kind: "stage_aging", labels: m.labels, datasets: m.datasets };
  }
  if (stageChartViewMode === "stage_age") {
    const m = buildStageByAgeStackedModel(filteredTasks);
    return { kind: "stage_age", labels: m.labels, datasets: m.datasets };
  }
  // default stage view
  const sd = aggregates.stageDistribution || buildStageDistributionModel([]);
  return { kind: "stage", labels: sd.labels.length ? sd.labels : ["—"], datasets: [{ label: "Tasks", data: sd.labels.length ? sd.values : [0], backgroundColor: sd.labels.length ? sd.colors : ["#797e93"], borderRadius: 6, maxBarThickness: 36 }] };
}

function initCharts(aggregates) {
  Chart.defaults.color = getThemeTokens().textDim;

  const tokens = getThemeTokens();
  const filteredTasks = getFilteredFlatTools();
  const model = getStageChartModelFromAggregatesAndTasks(aggregates, filteredTasks);

  if (charts.status) {
    charts.status.destroy();
    charts.status = null;
  }

  // Stage chart scalability:
  // - horizontal bars
  // - if > 10 categories, scroll container and grow canvas height to fit labels (keeps labels fully visible)
  const stageWrap = document.getElementById("stageChartWrap");
  const stageCount = model.labels.length;
  const noScroll = stageCount <= 10;
  if (stageWrap) {
    stageWrap.classList.toggle("stage-chart-wrap--no-scroll", noScroll);
    const rowPx = 28; // label+bar row height
    const minPx = 240;
    const desired = Math.max(minPx, stageCount * rowPx);
    const canvas = document.getElementById("statusChart");
    if (canvas) canvas.style.height = `${desired}px`;
  }

  charts.status = new Chart($("statusChart"), {
    type: "bar",
    data: {
      labels: model.labels,
      datasets: model.datasets,
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: stageChartViewMode === "stage_age" || stageChartViewMode === "stage_aging", position: "bottom" },
        tooltip: {
          callbacks: {
            label: (ctx) => ` ${ctx.dataset.label || "Tasks"}: ${ctx.parsed.x}`,
          },
        },
      },
      onClick: (_evt, elements) => {
        if (!elements?.length) return;
        if (stageChartViewMode !== "stage_age" && stageChartViewMode !== "stage_aging") return;
        const el = elements[0];
        const stage = charts.status?.data?.labels?.[el.index];
        const bucketLabel = charts.status?.data?.datasets?.[el.datasetIndex]?.label;
        const bucketKey = AGE_BUCKETS.find((b) => b.label === bucketLabel)?.key;
        if (!stage || !bucketKey) return;

        smartFilters.stages.clear();
        smartFilters.ageBuckets.clear();
        smartFilters.stages.add(stage);
        smartFilters.ageBuckets.add(bucketKey);
        syncSmartFilterUiFromState();
        applyAndRender();
      },
      scales: {
        x: {
          beginAtZero: true,
          ticks: { color: tokens.textDim, precision: 0, stepSize: 1 },
          grid: { color: tokens.gridLine },
          stacked: stageChartViewMode === "stage_age" || stageChartViewMode === "stage_aging",
        },
        y: {
          ticks: {
            color: tokens.textDim,
            autoSkip: false, // keep all stage names visible
            crossAlign: "far",
          },
          grid: { color: tokens.gridLine },
          stacked: stageChartViewMode === "stage_age" || stageChartViewMode === "stage_aging",
        },
      },
      layout: {
        padding: { left: 8, right: 8 },
      },
    },
  });

  // Bar chart is only for workload/performance modes
  if (assigneeChartMode !== "smart") {
    charts.assignee = createOrRecreateAssigneeChart(aggregates.assigneeModel, null);
  }

  if (charts.health) {
    charts.health.destroy();
    charts.health = null;
  }

  charts.health = new Chart($("healthChart"), {
    type: "pie",
    data: {
      labels: ["Healthy", "Warning", "At Risk"],
      datasets: [
        {
          data: [aggregates.health.healthy, aggregates.health.warning, aggregates.health.atRisk],
          backgroundColor: ["#00c875", "#ffcb00", "#f44336"],
        },
      ],
    },
  });
}

function setAssigneeChartWrapBehavior(model) {
  const wrap = document.getElementById("assigneeChartWrap");
  if (!wrap) return;
  if (model.needsScroll) wrap.classList.add("scroll-x");
  else wrap.classList.remove("scroll-x");

  // Scale min-width with member count when scrolling is enabled
  if (model.needsScroll) {
    const pxPerUser = 90;
    const minWidth = Math.max(720, model.labels.length * pxPerUser);
    const canvas = document.getElementById("assigneeChart");
    if (canvas) canvas.style.minWidth = `${minWidth}px`;
  } else {
    const canvas = document.getElementById("assigneeChart");
    if (canvas) canvas.style.minWidth = "";
  }
}

function createOrRecreateAssigneeChart(model, existingChart) {
  setAssigneeChartWrapBehavior(model);

  const indexAxis = model.useHorizontal ? "y" : "x";
  const label = model.mode === "workload" ? "Workload (tasks)" : "Performance Score";
  const tokens = getThemeTokens();

  const tooltipCallbacks = {
    title: (items) => (items?.[0]?.label ? [items[0].label] : []),
    label: (ctx) => {
      const name = ctx.label;
      const value = ctx.parsed?.[model.useHorizontal ? "x" : "y"] ?? ctx.parsed;

      if (name === "Others") {
        const others = model.meta.get("Others")?.members || [];
        const lines = [];
        if (model.mode === "workload") {
          lines.push(`${label}: ${Number(value)}`);
          lines.push(`Members: ${others.length}`);
        } else {
          const contributed = others.reduce((sum, s) => sum + (s.total || 0), 0);
          const completed = others.reduce((sum, s) => sum + (s.completed || 0), 0);
          const totalScore = others.reduce((sum, s) => sum + performanceScore(s), 0);
          lines.push(`Contributed Tasks: ${contributed}`);
          lines.push(`Completed Tasks: ${completed}`);
          lines.push(`Total: ${Math.round(totalScore)}`);
          lines.push(`Members: ${others.length}`);
        }
        // Show breakdown (up to 12 lines)
        const top = others.slice(0, 12);
        for (const s of top) {
          const score = performanceScore(s);
          lines.push(
            `${s.name}: C${s.completed}, IP${s.inProgress}, OD${s.overdue}, Score ${score.toFixed(1)}, Tasks ${s.total}`
          );
        }
        if (others.length > top.length) lines.push(`…and ${others.length - top.length} more`);
        return lines;
      }

      const s = model.meta.get(name);
      if (!s) return `${label}: ${Number(value).toFixed(1)}`;

      const score = performanceScore(s);
      const pct = model.totalValue !== 0 ? (Number(value) / model.totalValue) * 100 : 0;

      if (model.mode === "workload") {
        return [
          `Total Tasks: ${s.total}`,
          `Completed Tasks: ${s.completed}`,
          `Overdue Tasks: ${s.overdue}`,
          `Contribution: ${pct.toFixed(1)}%`,
        ];
      }

      // View by Performance tooltip format requirement
      return [
        `Contributed Tasks: ${s.total}`,
        `Completed Tasks: ${s.completed}`,
        `Total: ${Math.round(score)}`,
      ];
    },
  };

  const config = {
    type: "bar",
    data: {
      labels: model.labels,
      datasets: [
        {
          label,
          data: model.values,
          backgroundColor: "#00a9ff",
          borderRadius: 6,
          maxBarThickness: 26,
        },
      ],
    },
    options: {
      indexAxis,
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: tooltipCallbacks },
      },
      scales: {
        x: {
          ticks: {
            maxRotation: model.useHorizontal ? 0 : model.memberCount > 6 ? 45 : 0,
            minRotation: model.useHorizontal ? 0 : model.memberCount > 6 ? 45 : 0,
            autoSkip: model.useHorizontal ? true : model.memberCount > 10 ? true : false,
            color: tokens.textDim,
          },
          grid: { color: tokens.gridLine },
        },
        y: { ticks: { autoSkip: true, color: tokens.textDim }, grid: { color: tokens.gridLine } },
      },
    },
  };

  // If axis orientation or labels count changed a lot, recreate
  if (existingChart) {
    const sameAxis = (existingChart.options.indexAxis || "x") === indexAxis;
    if (!sameAxis) {
      existingChart.destroy();
      return new Chart($("assigneeChart"), config);
    }

    existingChart.data.labels = model.labels;
    existingChart.data.datasets[0].label = label;
    existingChart.data.datasets[0].data = model.values;
    existingChart.options.indexAxis = indexAxis;
    existingChart.update();
    return existingChart;
  }

  // Ensure the canvas has room when maintainAspectRatio false
  const wrap = document.getElementById("assigneeChartWrap");
  if (wrap) wrap.style.height = model.useHorizontal ? "320px" : "260px";
  return new Chart($("assigneeChart"), config);
}

function updateCharts(aggregates) {
  if (!charts.status || !charts.assignee || !charts.health) {
    initCharts(aggregates);
    return;
  }

  const filteredTasks = getFilteredFlatTools();
  const model = getStageChartModelFromAggregatesAndTasks(aggregates, filteredTasks);
  const stageWrap = document.getElementById("stageChartWrap");
  const stageCount = model.labels.length;
  const noScroll = stageCount <= 10;
  if (stageWrap) {
    stageWrap.classList.toggle("stage-chart-wrap--no-scroll", noScroll);
    const rowPx = 28;
    const minPx = 240;
    const desired = Math.max(minPx, stageCount * rowPx);
    const canvas = document.getElementById("statusChart");
    if (canvas) canvas.style.height = `${desired}px`;
  }

  if (charts.status.config.type !== "bar") {
    charts.status.destroy();
    charts.status = null;
    initCharts(aggregates);
  } else {
    charts.status.data.labels = model.labels;
    charts.status.data.datasets = model.datasets;
    charts.status.options.indexAxis = "y";
    charts.status.options.plugins.legend.display = stageChartViewMode === "stage_age";
    charts.status.options.scales.x.stacked = stageChartViewMode === "stage_age";
    charts.status.options.scales.y.stacked = stageChartViewMode === "stage_age";
    charts.status.update();
  }

  if (assigneeChartMode !== "smart") {
    charts.assignee = createOrRecreateAssigneeChart(aggregates.assigneeModel, charts.assignee);
  }

  charts.health.data.datasets[0].data = [aggregates.health.healthy, aggregates.health.warning, aggregates.health.atRisk];
  charts.health.update();
}

function setDateInputsEnabled(isEnabled) {
  const start = $("startDateInput");
  const end = $("endDateInput");
  start.disabled = !isEnabled;
  end.disabled = !isEnabled;
}

function syncInputsForPreset(preset) {
  const startEl = $("startDateInput");
  const endEl = $("endDateInput");
  const customWrap = document.getElementById("customDateRange");
  if (preset === "Custom") {
    setDateInputsEnabled(true);
    if (customWrap) {
      customWrap.hidden = false;
      customWrap.setAttribute("aria-hidden", "false");
    }
    return;
  }

  setDateInputsEnabled(false);
  if (customWrap) {
    customWrap.hidden = true;
    customWrap.setAttribute("aria-hidden", "true");
  }
  const range = getDateRangeFromPreset(preset, startEl, endEl);
  if (range.start) startEl.value = toDateInputValue(range.start);
  if (range.end) endEl.value = toDateInputValue(new Date(range.end.getFullYear(), range.end.getMonth(), range.end.getDate()));
}

function getActiveDateRangeFromUi() {
  const dateFilter = document.getElementById("dateFilter");
  const startEl = document.getElementById("startDateInput");
  const endEl = document.getElementById("endDateInput");
  if (!dateFilter || !startEl || !endEl) return { start: null, end: null };
  return getDateRangeFromPreset(dateFilter.value, startEl, endEl);
}

function applyAndRender() {
  const range = getActiveDateRangeFromUi();
  const filtered = getFilteredFlatTools().map((t) => ({
    id: t.toolId,
    name: `${t.orderName} ↳ ${t.toolName}`,
    assignee: t.assignee,
    status: t.status,
    dueDate: t.dueDate,
    parentId: t.orderId,
    workflowStage: t.workflowStage,
    board: t.board,
    workspace: t.workspace,
    orderType: t.orderType,
    labels: t.labels,
  }));
  debugLog({
    runId: "pre-fix",
    hypothesisId: "H_filters",
    location: "app.js:applyAndRender",
    message: "Apply filters",
    data: {
      start: range.start ? range.start.toISOString() : null,
      end: range.end ? range.end.toISOString() : null,
      filteredCount: filtered.length,
    },
  });

  // KPI cards are driven by dropdown mode now
  renderKpis();
  renderActiveGrid(range);
  renderUpcomingDeadlinesFromOrders();
  renderActiveFilterChips();
  renderGridHealthKpis();

  const aggregates = buildAggregates(filtered);
  updateCharts(aggregates);

  // Smart Distribution is UI-local: use fixed sample dataset
  setAssigneeCardVisibility();
  if (assigneeChartMode === "smart") {
    if (charts.assignee) {
      charts.assignee.destroy();
      charts.assignee = null;
    }
    destroySmartCharts();
    createSmartDistributionCharts();
  } else {
    destroySmartCharts();
  }
}

function populateDashboardFilterControls() {
  const boardEl = document.getElementById("filterBoard");
  const stageEl = document.getElementById("filterStage");
  if (!boardEl || !stageEl) return;

  boardEl.innerHTML = "";
  const bAll = document.createElement("option");
  bAll.value = "all";
  bAll.textContent = "All boards";
  boardEl.appendChild(bAll);
  for (const b of BOARD_LIST) {
    const opt = document.createElement("option");
    opt.value = b;
    opt.textContent = b;
    boardEl.appendChild(opt);
  }
  boardEl.value = dashboardFilters.board;

  function syncStageDropdown() {
    const stageOpts = getStageOptionsForBoardFilter(boardEl.value);
    const prev = dashboardFilters.stage;
    stageEl.innerHTML = "";
    const sAll = document.createElement("option");
    sAll.value = "all";
    sAll.textContent = "All stages";
    stageEl.appendChild(sAll);
    for (const s of stageOpts) {
      const opt = document.createElement("option");
      opt.value = s;
      opt.textContent = s;
      stageEl.appendChild(opt);
    }
    if (prev !== "all" && stageOpts.includes(prev)) dashboardFilters.stage = prev;
    else dashboardFilters.stage = "all";
    stageEl.value = dashboardFilters.stage;
  }

  syncStageDropdown();

  boardEl.addEventListener("change", () => {
    dashboardFilters.board = boardEl.value;
    expandedOrderId = null;
    syncStageDropdown();
    applyAndRender();
  });

  stageEl.addEventListener("change", () => {
    dashboardFilters.stage = stageEl.value;
    expandedOrderId = null;
    applyAndRender();
  });

  const workspaceEl = document.getElementById("filterWorkspace");
  if (workspaceEl) {
    workspaceEl.innerHTML = "";
    const wAll = document.createElement("option");
    wAll.value = "all";
    wAll.textContent = "All workspaces";
    workspaceEl.appendChild(wAll);
    for (const w of WORKSPACE_OPTIONS) {
      const opt = document.createElement("option");
      opt.value = w;
      opt.textContent = w;
      workspaceEl.appendChild(opt);
    }
    workspaceEl.value = dashboardFilters.workspace;
    workspaceEl.addEventListener("change", () => {
      dashboardFilters.workspace = workspaceEl.value;
      expandedOrderId = null;
      applyAndRender();
    });
  }

  syncWorkspaceFilterVisibility();
}

function init() {
  const dateFilter = document.getElementById("dateFilter");
  const applyBtn = document.getElementById("applyFiltersBtn");
  const modeInputs = document.querySelectorAll('input[name="assigneeChartMode"]');
  const tabOrder = document.getElementById("tabOrderView");
  const tabTask = document.getElementById("tabTaskView");
  const themeToggle = document.getElementById("themeToggle");
  const kpiModeSelect = document.getElementById("kpiModeSelect");
  const exportBtn = document.getElementById("exportGridBtn");

  // Filter bar may be removed in some layouts; default to no range filtering.
  if (dateFilter) {
    // Default preset
    dateFilter.value = "Current Month";
    syncInputsForPreset(dateFilter.value);
    debugLog({
      runId: "pre-fix",
      hypothesisId: "H_dom",
      location: "app.js:init",
      message: "Dashboard init",
      data: { preset: dateFilter.value, mainTasksCount: mainTasks.length },
    });
  } else {
    debugLog({
      runId: "pre-fix",
      hypothesisId: "H_dom",
      location: "app.js:init",
      message: "Dashboard init (no filter bar)",
      data: { mainTasksCount: mainTasks.length },
    });
  }

  populateDashboardFilterControls();

  initPrimaryTabs();
  initRoleTabs();
  initSmartFilterPopover();
  const viewSel = document.getElementById("stageChartViewSelect");
  const basisSel = document.getElementById("ageBasisSelect");
  if (viewSel) {
    stageChartViewMode = viewSel.value || "stage";
    viewSel.addEventListener("change", () => {
      stageChartViewMode = viewSel.value || "stage";
      // Stage Aging Analysis uses stage-entered basis only
      if (stageChartViewMode === "stage_aging") {
        ageBasisMode = "stageEntered";
        const basisSel = document.getElementById("ageBasisSelect");
        if (basisSel) basisSel.value = "stageEntered";
      }
      applyAndRender();
    });
  }
  if (basisSel) {
    ageBasisMode = basisSel.value === "stageEntered" ? "stageEntered" : "created";
    basisSel.addEventListener("change", () => {
      ageBasisMode = basisSel.value === "stageEntered" ? "stageEntered" : "created";
      applyAndRender();
    });
  }

  document.getElementById("scopeAllTasks")?.addEventListener("click", () => {
    managerLeadScope = "all";
    syncManagerLeadScopeVisibility();
    applyAndRender();
  });
  document.getElementById("scopeMyTasks")?.addEventListener("click", () => {
    managerLeadScope = "mine";
    syncManagerLeadScopeVisibility();
    applyAndRender();
  });

  // Click-to-filter interactions (labels + assignees) from grid
  document.addEventListener("click", (e) => {
    const t = e.target;
    if (!t) return;
    const labelEl = t.closest?.(".label-chip[data-label]");
    if (labelEl && labelEl.getAttribute) {
      applyLabelFilter(labelEl.getAttribute("data-label"));
      return;
    }
    const assigneeEl = t.closest?.(".assignee-cell[data-assignee]");
    if (assigneeEl && assigneeEl.getAttribute) {
      const a = assigneeEl.getAttribute("data-assignee");
      if (a && a !== "—") applyAssigneeFilter(a);
    }
  });

  // Theme init + toggle
  let saved = null;
  try {
    saved = localStorage.getItem("orionTheme");
  } catch {
    saved = null;
  }
  applyTheme(saved === "light" ? "light" : "dark");
  if (themeToggle) {
    themeToggle.checked = activeTheme === "light";
    themeToggle.addEventListener("change", () => {
      applyTheme(themeToggle.checked ? "light" : "dark");
    });
  }

  if (kpiModeSelect) {
    kpiMode = kpiModeSelect.value === "main" ? "main" : "sub";
    kpiModeSelect.addEventListener("change", () => {
      if (activeDashboardRole === "member" && kpiModeSelect.value === "main") {
        kpiModeSelect.value = "sub";
      }
      kpiMode = kpiModeSelect.value === "main" ? "main" : "sub";
      renderKpis();
    });
  }

  if (exportBtn) {
    exportBtn.addEventListener("click", () => exportCurrentGridToCsv());
  }

  for (const input of modeInputs) {
    input.addEventListener("change", () => {
      if (!input.checked) return;
      if (input.value === "smart") assigneeChartMode = "smart";
      else assigneeChartMode = input.value === "workload" ? "workload" : "performance";
      applyAndRender();
    });
  }

  if (tabOrder && tabTask) {
    tabOrder.addEventListener("click", () => {
      activeGridTab = "order";
      expandedOrderId = null;
      renderGridTabs();
      applyAndRender();
    });
    tabTask.addEventListener("click", () => {
      activeGridTab = "task";
      expandedTaskId = null;
      renderGridTabs();
      applyAndRender();
    });
  }

  if (dateFilter) {
    dateFilter.addEventListener("change", () => {
      syncInputsForPreset(dateFilter.value);
      applyAndRender();
    });
  }

  // Custom range changes should reflect immediately
  const startEl = document.getElementById("startDateInput");
  const endEl = document.getElementById("endDateInput");
  if (startEl) startEl.addEventListener("change", () => applyAndRender());
  if (endEl) endEl.addEventListener("change", () => applyAndRender());

  if (applyBtn) {
    applyBtn.addEventListener("click", () => {
      applyAndRender();
    });
  }

  initDeadlineModal();

  // Initial render
  applyAndRender();
}

document.addEventListener("DOMContentLoaded", init);

