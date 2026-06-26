import { useState, useRef, useEffect } from "react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, LineChart, Line, CartesianGrid, PieChart, Pie, Cell } from "recharts";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";

// ── Constants ─────────────────────────────────────────────────────────────────
// Generate Term 1/2/3 for a range of years so rollover/promotion always has valid next-terms
const TERM_YEARS = ["2024", "2025", "2026", "2027", "2028"];
const TERMS = TERM_YEARS.flatMap(y => [`Term 1, ${y}`, `Term 2, ${y}`, `Term 3, ${y}`]);

// ── School Types & Class Lists ───────────────────────────────────────────────
const NURSERY_CLASSES = ["Baby Class", "Middle Class", "Top Class"];
const PRIMARY_CLASSES = ["P1", "P2", "P3", "P4", "P5", "P6", "P7"];
const SECONDARY_CLASSES = ["S1", "S2", "S3", "S4", "S5", "S6"];

const SCHOOL_TYPES = {
  nursery: { label: "Nursery", classes: [...NURSERY_CLASSES] },
  primary: { label: "Primary", classes: [...PRIMARY_CLASSES] },
  secondary: { label: "Secondary", classes: [...SECONDARY_CLASSES] },
  nursery_primary: { label: "Nursery & Primary", classes: [...NURSERY_CLASSES, ...PRIMARY_CLASSES] },
  primary_secondary: { label: "Primary & Secondary", classes: [...PRIMARY_CLASSES, ...SECONDARY_CLASSES] },
  full: { label: "Nursery, Primary & Secondary", classes: [...NURSERY_CLASSES, ...PRIMARY_CLASSES, ...SECONDARY_CLASSES] },
};

// Get the class list for a given school (falls back to Secondary for legacy/demo schools)
const getSchoolClasses = (school) => SCHOOL_TYPES[school?.schoolType]?.classes || SECONDARY_CLASSES;

// Build display label for a class + optional stream, e.g. "S1 - East"
const classLabel = (cls, stream) => stream ? `${cls} - ${stream}` : cls;

// Identify "transition" classes — the last class of each education level within a
// school's type (e.g. P7 for primary, S4 for O-level). Students in these classes
// need an explicit "continue or leave" decision during promotion, since many
// transfer institutions at these points. The very last class overall always
// graduates (no decision needed — they're leaving regardless).
const getTransitionClasses = (school) => {
  const type = school?.schoolType || "secondary";
  const levelEnds = {
    nursery: ["Top Class"],
    primary: ["P7"],
    secondary: ["S4"], // S6 is the final graduation, not a transition
    nursery_primary: ["Top Class", "P7"],
    primary_secondary: ["P7", "S4"],
    full: ["Top Class", "P7", "S4"],
  };
  return levelEnds[type] || ["S4"];
};

const METHODS = ["Cash", "MTN MoMo", "Airtel Money", "Bank"];
const METHOD_ICON = { Cash: "💵", "MTN MoMo": "📱", "Airtel Money": "📲", Bank: "🏦" };
const STAFF = ["Mugisha R.", "Nakato B.", "Ssali J.", "Admin"];

// ── Default Fee Structure (editable per school via state) ────────────────────
const DEFAULT_FEE_STRUCTURE = {
  "Day Scholar": {
    // Nursery
    "Baby Class": { tuition: 300000, lunch: 100000, development: 50000, pta: 10000 },
    "Middle Class": { tuition: 320000, lunch: 100000, development: 50000, pta: 10000 },
    "Top Class": { tuition: 350000, lunch: 100000, development: 50000, pta: 10000 },
    // Primary
    P1: { tuition: 380000, lunch: 110000, development: 60000, pta: 15000 },
    P2: { tuition: 380000, lunch: 110000, development: 60000, pta: 15000 },
    P3: { tuition: 400000, lunch: 110000, development: 60000, pta: 15000 },
    P4: { tuition: 420000, lunch: 110000, development: 60000, pta: 15000 },
    P5: { tuition: 450000, lunch: 110000, development: 70000, pta: 15000 },
    P6: { tuition: 480000, lunch: 110000, development: 70000, pta: 15000 },
    P7: { tuition: 520000, lunch: 120000, development: 80000, pta: 15000 },
    // Secondary
    S1: { tuition: 650000, lunch: 120000, development: 80000, pta: 20000 },
    S2: { tuition: 650000, lunch: 120000, development: 80000, pta: 20000 },
    S3: { tuition: 720000, lunch: 120000, development: 80000, pta: 20000 },
    S4: { tuition: 720000, lunch: 120000, development: 80000, pta: 20000 },
    S5: { tuition: 850000, lunch: 120000, development: 100000, pta: 20000 },
    S6: { tuition: 850000, lunch: 120000, development: 100000, pta: 20000 },
  },
  "Boarder": {
    // Nursery (boarding rarely offered, but included for completeness)
    "Baby Class": { tuition: 300000, boarding: 350000, lunch: 150000, development: 50000, pta: 10000 },
    "Middle Class": { tuition: 320000, boarding: 350000, lunch: 150000, development: 50000, pta: 10000 },
    "Top Class": { tuition: 350000, boarding: 350000, lunch: 150000, development: 50000, pta: 10000 },
    // Primary
    P1: { tuition: 380000, boarding: 380000, lunch: 160000, development: 60000, pta: 15000 },
    P2: { tuition: 380000, boarding: 380000, lunch: 160000, development: 60000, pta: 15000 },
    P3: { tuition: 400000, boarding: 380000, lunch: 160000, development: 60000, pta: 15000 },
    P4: { tuition: 420000, boarding: 400000, lunch: 160000, development: 60000, pta: 15000 },
    P5: { tuition: 450000, boarding: 400000, lunch: 170000, development: 70000, pta: 15000 },
    P6: { tuition: 480000, boarding: 420000, lunch: 170000, development: 70000, pta: 15000 },
    P7: { tuition: 520000, boarding: 420000, lunch: 180000, development: 80000, pta: 15000 },
    // Secondary
    S1: { tuition: 650000, boarding: 450000, lunch: 180000, development: 80000, pta: 20000 },
    S2: { tuition: 650000, boarding: 450000, lunch: 180000, development: 80000, pta: 20000 },
    S3: { tuition: 720000, boarding: 450000, lunch: 180000, development: 80000, pta: 20000 },
    S4: { tuition: 720000, boarding: 450000, lunch: 180000, development: 80000, pta: 20000 },
    S5: { tuition: 850000, boarding: 500000, lunch: 180000, development: 100000, pta: 20000 },
    S6: { tuition: 850000, boarding: 500000, lunch: 180000, development: 100000, pta: 20000 },
  },
};

// Keep a global FEE_STRUCTURE reference — overridden by component state
let FEE_STRUCTURE = JSON.parse(JSON.stringify(DEFAULT_FEE_STRUCTURE));

const STUDENT_CATEGORIES = ["Day Scholar", "Boarder"];

// Default school requirements (items + costs — editable)
const DEFAULT_REQUIREMENTS = [
  { id: "r1", name: "School Uniform (2 sets)", cost: 80000, appliesTo: ["Day Scholar", "Boarder"], mandatory: true },
  { id: "r2", name: "Exercise Books (set)", cost: 35000, appliesTo: ["Day Scholar", "Boarder"], mandatory: true },
  { id: "r3", name: "Mattress & Bedding", cost: 120000, appliesTo: ["Boarder"], mandatory: true },
  { id: "r4", name: "Sports Kit", cost: 45000, appliesTo: ["Day Scholar", "Boarder"], mandatory: false },
  { id: "r5", name: "Library Fee", cost: 20000, appliesTo: ["Day Scholar", "Boarder"], mandatory: true },
];

// getStudentFee is redefined inside component to access state — see below
const getStudentFeeStatic = (student, feeStructure, requirements) => {
  if (!student) return 0; // safety guard — should never be undefined but prevents white screen
  // Requirements that apply to this student (mandatory only added to fee)
  const reqCost = requirements
    .filter(r => r.mandatory && r.appliesTo.includes(student.category || "Day Scholar"))
    .reduce((a, r) => a + r.cost, 0);

  if (student.customFee && student.customFee > 0) return student.customFee + reqCost;
  const cat = student.category || "Day Scholar";
  const structure = feeStructure[cat] || feeStructure["Day Scholar"];
  const baseFee = Object.values(structure[student.class] || {}).reduce((a, b) => a + b, 0);
  let fee = baseFee;
  if (student.bursary) {
    if (student.bursary.type === "percent") fee = Math.round(baseFee * (1 - student.bursary.value / 100));
    if (student.bursary.type === "fixed") fee = Math.max(0, baseFee - student.bursary.value);
  }
  return fee + reqCost;
};

// ── SINGLE SOURCE OF TRUTH for what a student owes ──────────────────────────
// Every screen in the app (dashboard, balances report, receipts, promotion,
// rollover, parent portal, exports) MUST call this instead of recomputing the
// formula inline. This is what previously caused balances to look different
// depending on which screen you were on — some places forgot to include
// arrears carried over from earlier terms. There is now exactly one formula:
//
//   Total Due (this term)  =  this term's fee  +  arrears carried in
//   Balance                =  Total Due  −  payments made in this term
//
// "Arrears" already represents the FULL unpaid balance from every term and
// year before the current one (it is itself recalculated the same way, via
// this function, every time a term ends — see handleRollover/handlePromotion).
// So nothing from past terms is ever left out: it's either fully reflected in
// "arrears", or it's the current term's fee, and this function adds both.
const getBalanceStatic = (student, term, feeStructure, requirements) => {
  const termFee = getStudentFeeStatic(student, feeStructure, requirements);
  const arrears = student.arrears || 0;
  const totalDue = termFee + arrears;
  const paidThisTerm = student.payments.filter(p => !term || p.term === term).reduce((a, p) => a + p.amount, 0);
  const balance = Math.max(0, totalDue - paidThisTerm);
  const status = paidThisTerm >= totalDue ? "Paid" : paidThisTerm > 0 ? "Partial" : "Unpaid";
  return { termFee, arrears, totalDue, paidThisTerm, balance, status };
};

// ── Multi-School Data ─────────────────────────────────────────────────────────
// ── Subscription Plans ────────────────────────────────────────────────────────
const PLANS = {
  Starter: { name: "Starter", price: 100000, maxStudents: 200, features: ["Fee tracking", "Receipts", "Basic reports"] },
  Standard: { name: "Standard", price: 200000, maxStudents: 500, features: ["Everything in Starter", "SMS notifications", "Excel export"] },
  Premium: { name: "Premium", price: 350000, maxStudents: Infinity, features: ["Everything in Standard", "Bank reconciliation", "Multi-campus", "Bulk import"] },
};
// Per-term price = 3 months' worth at a 10% discount, rewarding schools that pay upfront
// for the whole term instead of monthly. A term is treated as 90 days for billing-date math.
const TERM_DISCOUNT = 0.10;
const termPriceFor = (planName) => Math.round((PLANS[planName]?.price || 0) * 3 * (1 - TERM_DISCOUNT));
const BILLING_CYCLE_DAYS = { monthly: 30, term: 90 };
// Returns { price, cycleLabel, periodLabel, isCustom } for a given plan + billing cycle.
// Falls back gracefully to Starter pricing if an invalid/missing plan name is passed.
// `customPrice` is an optional Super-Admin-set override for ONE specific school — when
// present, it replaces the calculated price but everything else (cycle, period label,
// cycleDays) still comes from the school's actual plan/cycle, so "Per Term"/"Monthly"
// labeling and renewal-date math stay correct even for a discounted school.
const getBillingInfo = (planName, billingCycle, customPrice) => {
  const cycle = billingCycle === "term" ? "term" : "monthly";
  const calculatedPrice = cycle === "term" ? termPriceFor(planName) : (PLANS[planName]?.price || 0);
  const hasOverride = customPrice !== null && customPrice !== undefined && customPrice !== "" && !isNaN(customPrice);
  const price = hasOverride ? Number(customPrice) : calculatedPrice;
  return { price, cycle, periodLabel: cycle === "term" ? "/term" : "/mo", cycleDays: BILLING_CYCLE_DAYS[cycle], isCustom: hasOverride, standardPrice: calculatedPrice };
};
const GRACE_PERIOD_DAYS = 7;

// Suggest a plan based on the school's stated number of students at signup
const suggestPlan = (studentsEstimate) => {
  const n = parseInt(String(studentsEstimate).replace(/[^0-9]/g, "")) || 0;
  if (n <= 0) return "Starter"; // no estimate given — default to lowest tier
  if (n <= PLANS.Starter.maxStudents) return "Starter";
  if (n <= PLANS.Standard.maxStudents) return "Standard";
  return "Premium";
};

// Given an actual enrolled-student count, return the minimum plan that supports it
const minPlanForCount = (count) => {
  if (count <= PLANS.Starter.maxStudents) return "Starter";
  if (count <= PLANS.Standard.maxStudents) return "Standard";
  return "Premium";
};

// Plan order for comparing "is this an upgrade" — never auto-downgrade
const PLAN_ORDER = ["Starter", "Standard", "Premium"];

// Populated at runtime from the real Supabase `schools` table (see the
// loadSchoolsFromSupabase effect below) — starts empty, not with fake demo
// data, since this now reflects real schools that have actually signed up.
// Kept as a plain mutable object (not React state) because many existing
// functions throughout this file already mutate it directly and then bump
// `subscriptionRefresh` to force a re-render — changing that pattern now
// would mean touching dozens of call sites for no real benefit.
let SCHOOLS_DATA = {};

// ── Initial Students ──────────────────────────────────────────────────────────
const makeStudents = (schoolId) => [
  { id: `${schoolId}-1`, schoolId, name: "Nakato Aisha", class: "S1", gender: "F", category: "Day Scholar", parent: "Nakato Mary", phone: "0772-441-823", arrears: 0, bursary: null, customFee: null, payments: [{ id: "RCP-1001", date: "2025-05-05", amount: 500000, method: "MTN MoMo", receivedBy: "Mugisha R.", term: "Term 2, 2025" }, { id: "RCP-1018", date: "2025-05-14", amount: 220000, method: "Cash", receivedBy: "Nakato B.", term: "Term 2, 2025" }] },
  { id: `${schoolId}-2`, schoolId, name: "Ssemakula Brian", class: "S2", gender: "M", category: "Boarder", parent: "Ssemakula John", phone: "0701-234-567", arrears: 150000, bursary: null, customFee: null, payments: [{ id: "RCP-1003", date: "2025-05-06", amount: 400000, method: "Airtel Money", receivedBy: "Mugisha R.", term: "Term 2, 2025" }] },
  { id: `${schoolId}-3`, schoolId, name: "Namukasa Grace", class: "S1", gender: "F", category: "Day Scholar", parent: "Namukasa Ruth", phone: "0752-345-678", arrears: 200000, bursary: { type: "percent", value: 50, reason: "Orphan Bursary" }, customFee: null, payments: [] },
  { id: `${schoolId}-4`, schoolId, name: "Okello David", class: "S3", gender: "M", category: "Boarder", parent: "Okello James", phone: "0783-456-789", arrears: 0, bursary: null, customFee: null, payments: [{ id: "RCP-1005", date: "2025-05-02", amount: 500000, method: "Bank", receivedBy: "Nakato B.", term: "Term 2, 2025" }, { id: "RCP-1009", date: "2025-05-08", amount: 420000, method: "MTN MoMo", receivedBy: "Mugisha R.", term: "Term 2, 2025" }] },
  { id: `${schoolId}-5`, schoolId, name: "Nabirye Fatuma", class: "S2", gender: "F", category: "Day Scholar", parent: "Nabirye Hawa", phone: "0712-567-890", arrears: 0, bursary: { type: "fixed", value: 200000, reason: "Staff Child Discount" }, customFee: null, payments: [{ id: "RCP-1007", date: "2025-05-07", amount: 300000, method: "Cash", receivedBy: "Nakato B.", term: "Term 2, 2025" }, { id: "RCP-1022", date: "2025-05-20", amount: 300000, method: "Cash", receivedBy: "Nakato B.", term: "Term 2, 2025" }] },
  { id: `${schoolId}-6`, schoolId, name: "Mugisha Patrick", class: "S3", gender: "M", category: "Boarder", parent: "Mugisha Fred", phone: "0700-678-901", arrears: 100000, bursary: null, customFee: null, payments: [{ id: "RCP-1010", date: "2025-05-09", amount: 200000, method: "Cash", receivedBy: "Mugisha R.", term: "Term 2, 2025" }] },
  { id: `${schoolId}-7`, schoolId, name: "Apio Christine", class: "S4", gender: "F", category: "Day Scholar", parent: "Apio Agnes", phone: "0776-789-012", arrears: 0, bursary: null, customFee: null, payments: [{ id: "RCP-1002", date: "2025-05-05", amount: 920000, method: "Bank", receivedBy: "Nakato B.", term: "Term 2, 2025" }] },
  { id: `${schoolId}-8`, schoolId, name: "Byamugisha Kevin", class: "S5", gender: "M", category: "Boarder", parent: "Byamugisha Paul", phone: "0753-890-123", arrears: 0, bursary: null, customFee: null, payments: [{ id: "RCP-1011", date: "2025-05-10", amount: 500000, method: "MTN MoMo", receivedBy: "Mugisha R.", term: "Term 2, 2025" }, { id: "RCP-1025", date: "2025-05-22", amount: 300000, method: "Airtel Money", receivedBy: "Nakato B.", term: "Term 2, 2025" }] },
  { id: `${schoolId}-9`, schoolId, name: "Nalubega Sandra", class: "S6", gender: "F", category: "Day Scholar", parent: "Nalubega Prossy", phone: "0704-901-234", arrears: 0, bursary: null, customFee: 800000, payments: [{ id: "RCP-1004", date: "2025-05-06", amount: 800000, method: "Bank", receivedBy: "Nakato B.", term: "Term 2, 2025" }] },
  { id: `${schoolId}-10`, schoolId, name: "Ochieng Moses", class: "S4", gender: "M", category: "Day Scholar", parent: "Ochieng Peter", phone: "0789-012-345", arrears: 300000, bursary: null, customFee: null, payments: [] },
  { id: `${schoolId}-11`, schoolId, name: "Nansubuga Irene", class: "S2", gender: "F", category: "Boarder", parent: "Nansubuga Sarah", phone: "0771-123-456", arrears: 0, bursary: null, customFee: null, payments: [{ id: "RCP-1006", date: "2025-05-06", amount: 490000, method: "MTN MoMo", receivedBy: "Mugisha R.", term: "Term 2, 2025" }, { id: "RCP-1013", date: "2025-05-12", amount: 380000, method: "Cash", receivedBy: "Nakato B.", term: "Term 2, 2025" }] },
  { id: `${schoolId}-12`, schoolId, name: "Tumwine Alex", class: "S5", gender: "M", category: "Day Scholar", parent: "Tumwine Robert", phone: "0702-234-567", arrears: 80000, bursary: { type: "percent", value: 30, reason: "Academic Scholarship" }, customFee: null, payments: [{ id: "RCP-1015", date: "2025-05-13", amount: 500000, method: "Airtel Money", receivedBy: "Mugisha R.", term: "Term 2, 2025" }] },
  { id: `${schoolId}-13`, schoolId, name: "Nakato Brenda", class: "S4", gender: "F", category: "Day Scholar", parent: "Nakato Mary", phone: "0772-441-823", arrears: 0, bursary: null, customFee: null, payments: [{ id: "RCP-1030", date: "2025-05-09", amount: 500000, method: "MTN MoMo", receivedBy: "Mugisha R.", term: "Term 2, 2025" }] },
];

const makeExpenses = (schoolId) => [
  { id: `${schoolId}-e1`, schoolId, category: "Salaries", description: "Teaching staff salaries", amount: 8500000, date: "2025-05-01", term: "Term 2, 2025", paidBy: "Admin" },
  { id: `${schoolId}-e2`, schoolId, category: "Utilities", description: "Electricity & Water", amount: 450000, date: "2025-05-05", term: "Term 2, 2025", paidBy: "Mugisha R." },
  { id: `${schoolId}-e3`, schoolId, category: "Food", description: "Lunch provisions", amount: 1200000, date: "2025-05-08", term: "Term 2, 2025", paidBy: "Nakato B." },
  { id: `${schoolId}-e4`, schoolId, category: "Maintenance", description: "Classroom repairs", amount: 320000, date: "2025-05-15", term: "Term 2, 2025", paidBy: "Admin" },
  { id: `${schoolId}-e5`, schoolId, category: "Stationery", description: "Exercise books & pens", amount: 180000, date: "2025-05-18", term: "Term 2, 2025", paidBy: "Nakato B." },
];

// Sample staff/workers per school — pay is recorded flexibly (no fixed daily/monthly
// constraint per worker), so defaultRate/defaultRateType here are just a reference
// shown to make recording a payment faster, not an enforced rule.
const makeStaff = (schoolId) => [
  { id: `${schoolId}-st1`, schoolId, name: "Mugisha Robert", role: "Security Guard", phone: "0772-100-001", defaultRate: 15000, defaultRateType: "daily", active: true, photo: null },
  { id: `${schoolId}-st2`, schoolId, name: "Nakato Betty", role: "Cook", phone: "0772-100-002", defaultRate: 12000, defaultRateType: "daily", active: true, photo: null },
  { id: `${schoolId}-st3`, schoolId, name: "Okello Patrick", role: "Groundskeeper", phone: "0772-100-003", defaultRate: 10000, defaultRateType: "daily", active: true, photo: null },
  { id: `${schoolId}-st4`, schoolId, name: "Atim Sarah", role: "Matron", phone: "0772-100-004", defaultRate: 350000, defaultRateType: "monthly", active: true, photo: null },
  { id: `${schoolId}-st5`, schoolId, name: "Kato Emmanuel", role: "Driver", phone: "0772-100-005", defaultRate: 400000, defaultRateType: "monthly", active: true, photo: null },
];
const makeStaffPayments = (schoolId) => [
  { id: `${schoolId}-sp1`, schoolId, staffId: `${schoolId}-st1`, staffName: "Mugisha Robert", amount: 15000, payType: "daily", periodLabel: "12 May 2025", date: "2025-05-12", term: "Term 2, 2025", paidBy: "Admin" },
  { id: `${schoolId}-sp2`, schoolId, staffId: `${schoolId}-st4`, staffName: "Atim Sarah", amount: 350000, payType: "monthly", periodLabel: "May 2025", date: "2025-05-01", term: "Term 2, 2025", paidBy: "Admin" },
];

// ── Helpers ───────────────────────────────────────────────────────────────────
const fmt = (n) => "UGX " + Number(n || 0).toLocaleString();
const fmtShort = (n) => n >= 1000000 ? `${(n / 1000000).toFixed(1)}M` : n >= 1000 ? `${(n / 1000).toFixed(0)}K` : String(n);
const fmtDate = (d) => new Date(d).toLocaleDateString("en-UG", { day: "2-digit", month: "short", year: "numeric" });
const fmtDateTime = (d) => new Date(d).toLocaleString("en-UG", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });

// Compute live subscription status: Active / Grace Period / Suspended
// ── Session persistence ──────────────────────────────────────────
// Saves just enough to know who was logged in, in the browser's own
// localStorage — survives a page refresh, but NOT real authentication on
// its own (no password is re-checked on restore). On restore, the app
// re-fetches fresh real data from Supabase for whichever school/role was
// saved, rather than trusting any cached data — only the *fact* of being
// logged in is remembered, not stale information.
const SESSION_KEY = "feetrack_session";
// After this many minutes with no real activity (clicks, typing, touches,
// scrolling) on the page, the session is treated as expired and the person
// is logged out automatically — similar to how banking/finance apps behave,
// so a phone left open and unattended doesn't stay logged in indefinitely.
const INACTIVITY_TIMEOUT_MINUTES = 30;
const saveSession = (session) => {
  try { localStorage.setItem(SESSION_KEY, JSON.stringify({ ...session, lastActivity: Date.now() })); } catch { /* storage unavailable — fail silently, just won't persist */ }
};
const loadSession = () => {
  try {
    const raw = localStorage.getItem(SESSION_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch { return null; }
};
const clearSession = () => {
  try { localStorage.removeItem(SESSION_KEY); } catch { /* ignore */ }
};
// Bumps just the lastActivity timestamp on an already-saved session, without
// touching role/schoolId/tab — called frequently (on clicks/keys/etc.), so
// this needs to be cheap and must NOT call saveSession (which would re-stamp
// activity unconditionally even when nobody is logged in).
const touchSessionActivity = () => {
  try {
    const raw = localStorage.getItem(SESSION_KEY);
    if (!raw) return;
    const session = JSON.parse(raw);
    session.lastActivity = Date.now();
    localStorage.setItem(SESSION_KEY, JSON.stringify(session));
  } catch { /* ignore */ }
};
const isSessionExpired = (session) => {
  if (!session || !session.lastActivity) return false; // older/missing timestamp — don't punish, just let it through once
  return Date.now() - session.lastActivity > INACTIVITY_TIMEOUT_MINUTES * 60 * 1000;
};

const getSubscriptionInfo = (school) => {
  if (!school || !school.nextBillingDate) return { status: "Active", daysOverdue: 0 };
  const due = new Date(school.nextBillingDate);
  const now = new Date();
  const diffDays = Math.floor((now - due) / (1000 * 60 * 60 * 24));
  if (diffDays <= 0) return { status: "Active", daysOverdue: 0, daysUntilDue: -diffDays };

  // If a payment notice is pending review, freeze the countdown — never auto-suspend
  if (school.paymentNoticeFreeze) {
    const frozenDays = Math.min(diffDays, GRACE_PERIOD_DAYS);
    return { status: "Grace Period", daysOverdue: frozenDays, daysRemaining: Math.max(1, GRACE_PERIOD_DAYS - frozenDays), frozen: true };
  }

  if (diffDays <= GRACE_PERIOD_DAYS) return { status: "Grace Period", daysOverdue: diffDays, daysRemaining: GRACE_PERIOD_DAYS - diffDays };
  return { status: "Suspended", daysOverdue: diffDays };
};
const totalPaid = (s, term) => s.payments.filter(p => !term || p.term === term).reduce((a, p) => a + p.amount, 0);
// getStatus is redefined inside the component to use live feeStructure/requirements state
const getStatusStatic = (s, term, feeStructure, requirements) => getBalanceStatic(s, term, feeStructure, requirements).status;

// NOTE: This is a simple in-memory counter, sufficient for this single-session
// prototype. When migrating to the real backend, this MUST become a Postgres
// sequence (per-school or global) — see FeeTrack_Database_Schema.md, Migration
// Notes #2 — to avoid duplicate receipt numbers across concurrent bursars/devices.
let rcptN = 1050;
const nextRcpt = () => `RCP-${++rcptN}`;

// ── Components ────────────────────────────────────────────────────────────────
const Badge = ({ status }) => {
  const map = { Paid: ["#d1fae5", "#065f46"], Partial: ["#fef3c7", "#92400e"], Unpaid: ["#fee2e2", "#991b1b"] };
  const [bg, col] = map[status] || ["#f1f5f9", "#475569"];
  return <span style={{ background: bg, color: col, padding: "2px 9px", borderRadius: 99, fontSize: 11, fontWeight: 700 }}>{status}</span>;
};

const Pill = ({ text, bg = "#eff6ff", col = "#2563eb" }) => (
  <span style={{ background: bg, color: col, padding: "2px 9px", borderRadius: 99, fontSize: 11, fontWeight: 700 }}>{text}</span>
);

// ── SMS Log Store ─────────────────────────────────────────────────────────────
const buildSMS = (student, payment, school, balance) =>
  `Dear ${student.parent}, UGX ${payment.amount.toLocaleString()} received for ${student.name} (${student.class}) on ${fmtDate(payment.date)}. Balance: UGX ${balance.toLocaleString()}. Ref: ${payment.id}. ${school.name.split(" ").slice(0, 2).join(" ")}.`;

// ════════════════════════════════════════════════════════════════
export default function App() {
  // ── State ──────────────────────────────────────────────────────
  const [activeSchoolId, setActiveSchoolId] = useState(null);
  const school = SCHOOLS_DATA[activeSchoolId];
  const schoolClasses = getSchoolClasses(school); // class list for this school's type (Nursery/Primary/Secondary/etc.)
  const [currentUser, setCurrentUser] = useState(null); // null=admin, else {studentId}
  const [mobileSidebarOpen, setMobileSidebarOpen] = useState(false);
  const [isMobile, setIsMobile] = useState(typeof window !== "undefined" && window.innerWidth < 900);

  // Track viewport width for responsive layouts
  useEffect(() => {
    if (typeof window === "undefined") return;
    const onResize = () => setIsMobile(window.innerWidth < 900);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);
  const [loginScreen, setLoginScreen] = useState("admin"); // "admin"|"parent"|"school-select"|"signup"|"forgot-password"
  const [forgotPasswordEmail, setForgotPasswordEmail] = useState("");
  const [forgotPasswordStatus, setForgotPasswordStatus] = useState(null); // null | "sending" | "sent" | error message string
  const [resetPasswordToken, setResetPasswordToken] = useState(null); // set if the URL contains ?reset_token=...
  const [resetNewPassword, setResetNewPassword] = useState("");
  const [resetConfirmPassword, setResetConfirmPassword] = useState("");
  const [resetStatus, setResetStatus] = useState(null); // null | "submitting" | "success" | error message string
  const [loginInput, setLoginInput] = useState({ user: "", pass: "" });
  const [loginError, setLoginError] = useState("");
  const [schoolsLoading, setSchoolsLoading] = useState(true);
  // Checked once, synchronously, on first render — before the async schools
  // fetch even starts — purely so we know whether to show a brief "loading"
  // state instead of flashing the login screen while we wait to find out if
  // there's a session to restore. This does NOT log anyone in by itself;
  // loadSchools() still does the real restoration once real data has loaded.
  const [hasSavedSession] = useState(() => !!loadSession());
  const [schoolsLoadError, setSchoolsLoadError] = useState("");
  const [studentsLoading, setStudentsLoading] = useState(false);

  // Start empty — students/expenses/staff/staff-payments tables aren't migrated to
  // Supabase yet (only schools + students are, as of this step). Once each table is
  // migrated, this will load from Supabase the same way schools does below. Until
  // then, a real school correctly starts with nothing rather than fake demo data.
  const [allStudents, setAllStudents] = useState({});
  const [allExpenses, setAllExpenses] = useState({});
  const [allStaff, setAllStaff] = useState({});
  const [allStaffPayments, setAllStaffPayments] = useState({});
  const [smsLog, setSmsLog] = useState([]);

  const [tab, setTab] = useState("dashboard");
  const [currentTerm, setCurrentTerm] = useState(() => {
    try { return localStorage.getItem("feetrack_current_term") || "Term 2, 2025"; } catch { return "Term 2, 2025"; }
  });
  const [search, setSearch] = useState("");
  const [paymentsTermFilter, setPaymentsTermFilter] = useState("current"); // "current" | "all" | specific term string
  const [paymentsSearch, setPaymentsSearch] = useState("");
  const [filterClass, setFilterClass] = useState("All");
  const [filterStatus, setFilterStatus] = useState("All");

  const [showPay, setShowPay] = useState(null);
  const [payAmt, setPayAmt] = useState("");
  const [payMethod, setPayMethod] = useState("Cash");
  const [payDate, setPayDate] = useState(new Date().toISOString().split("T")[0]);
  const [receivedBy, setReceivedBy] = useState("Mugisha R.");
  const [sendSMS, setSendSMS] = useState(true);

  const [showAdd, setShowAdd] = useState(false);
  const [newS, setNewS] = useState({ name: "", class: "S1", stream: "", gender: "M", category: "Day Scholar", parent: "", phone: "", bursary: null, customFee: "" });
  const [returningMatch, setReturningMatch] = useState(null); // matched alumni record
  const [confirmReturning, setConfirmReturning] = useState(false);

  // First-time setup wizard
  const [setupStep, setSetupStep] = useState(1);
  const [setupDismissed, setSetupDismissed] = useState(false);

  // Bulk student import
  const [bulkRows, setBulkRows] = useState([]);
  const [bulkFileName, setBulkFileName] = useState("");
  const [bulkImportDone, setBulkImportDone] = useState(false);
  const [showBulkImport, setShowBulkImport] = useState(false);

  // Bulk payment import
  const [showBulkPayments, setShowBulkPayments] = useState(false);
  const [bulkPayRows, setBulkPayRows] = useState([]);
  const [bulkPayFileName, setBulkPayFileName] = useState("");
  const [bulkPayImportDone, setBulkPayImportDone] = useState(false);
  const bulkPayFileRef = useRef(null);
  const bulkFileRef = useRef(null);
  const [showBulkStaffPay, setShowBulkStaffPay] = useState(false);
  const [bulkStaffPayRows, setBulkStaffPayRows] = useState([]);
  const [bulkStaffPayFileName, setBulkStaffPayFileName] = useState("");
  const [bulkStaffPayImportDone, setBulkStaffPayImportDone] = useState(false);
  const bulkStaffPayFileRef = useRef(null);
  const [showFeeEdit, setShowFeeEdit] = useState(null); // student being fee-edited
  const [feeEditData, setFeeEditData] = useState({ mode: "category", bursaryType: "percent", bursaryValue: "", bursaryReason: "", customFee: "" });

  const [showAddExp, setShowAddExp] = useState(false);
  const [newExp, setNewExp] = useState({ category: "Salaries", description: "", amount: "", date: new Date().toISOString().split("T")[0] });
  const [showAddStaff, setShowAddStaff] = useState(false);
  const [showCustomPriceEdit, setShowCustomPriceEdit] = useState(null); // school being given a custom price
  const [customPriceForm, setCustomPriceForm] = useState({ price: "", note: "" });
  const [newStaff, setNewStaff] = useState({ name: "", role: "", phone: "", defaultRate: "", defaultRateType: "daily" });
  const [showEditStaff, setShowEditStaff] = useState(null);
  const [showPayStaff, setShowPayStaff] = useState(null);
  const [payStaffForm, setPayStaffForm] = useState({ amount: "", payType: "daily", periodLabel: "" });
  const [selectedStaffId, setSelectedStaffId] = useState(null);
  const [staffSearch, setStaffSearch] = useState("");

  const [showReceipt, setShowReceipt] = useState(null);
  const [showRollover, setShowRollover] = useState(false);
  const [showPromotion, setShowPromotion] = useState(false);
  const [promotionYear, setPromotionYear] = useState("2026");
  // S4 decisions: { studentId: "continue" | "leave" }
  const [transitionDecisions, setTransitionDecisions] = useState({}); // { studentId: "continue" | "leave" } — for end-of-level classes (P7, S4, Top Class, etc.)
  // "did not return" decisions for ALL students: { studentId: true }
  const [dnrDecisions, setDnrDecisions] = useState({});
  const [repeatDecisions, setRepeatDecisions] = useState({}); // { studentId: true } — student stays in same class (repeating the year)
  const [rolloverDnr, setRolloverDnr] = useState({});
  const [rolloverStep, setRolloverStep] = useState(1);
  const [rolloverTerm, setRolloverTerm] = useState("");
  const [showEditExpense, setShowEditExpense] = useState(null);
  const [showRecoverDebt, setShowRecoverDebt] = useState(null); // alumni record being collected from
  const [recoverAmt, setRecoverAmt] = useState("");
  const [showBalancesReport, setShowBalancesReport] = useState(false);
  // Generic confirm dialog (replaces window.confirm, which doesn't work in this sandbox)
  const [confirmDialog, setConfirmDialog] = useState(null); // { title, message, danger, onConfirm }
  const [activityLog, setActivityLog] = useState([]); // super-admin action history
  const [editExpAmt, setEditExpAmt] = useState("");

  // Bank reconciliation
  const [bankRows, setBankRows] = useState([]); // parsed from uploaded CSV/Excel

  // Subscription payment auto-reconciliation (super admin)
  const [subPayRows, setSubPayRows] = useState([]);
  const [subPayFileName, setSubPayFileName] = useState("");
  const [subPayResults, setSubPayResults] = useState(null);

  // "I've Sent Payment" confirmations (school → super admin notice)
  const [paymentNotices, setPaymentNotices] = useState([]);
  const [showPaymentConfirm, setShowPaymentConfirm] = useState(false);
  const [paymentConfirmForm, setPaymentConfirmForm] = useState({ method: "MTN MoMo", amount: "", date: new Date().toISOString().split("T")[0], note: "" });
  const subPayFileRef = useRef(null);
  const [bankMatched, setBankMatched] = useState({}); // { rowIndex: studentId }
  const [bankImporting, setBankImporting] = useState(false);
  const [bankFileName, setBankFileName] = useState("");
  const [bankImportDone, setBankImportDone] = useState(false);
  const bankFileRef = useRef(null);

  // ── School signup / approval system ────────────────────────────
  // Starts empty — populated from the real Supabase pending_signups table by
  // the loadPendingSignups effect below, not hardcoded fake demo entries.
  const [pendingSchools, setPendingSchools] = useState([]);
  const [signupForm, setSignupForm] = useState({ schoolName: "", location: "", principal: "", phone: "", email: "", students: "", schoolType: "secondary", billingCycle: "monthly", username: "", password: "", confirmPassword: "" });
  const [signupSubmitted, setSignupSubmitted] = useState(false);
  const [isSuperAdmin, setIsSuperAdmin] = useState(false);
  const [subscriptionRefresh, setSubscriptionRefresh] = useState(0); // bump to force re-render after SCHOOLS_DATA mutation

  // ── Load real schools from Supabase ────────────────────────────
  // Runs once when the app first loads. Populates the module-level SCHOOLS_DATA
  // object (declared near the top of this file) with real rows from the
  // database, converting each row's snake_case columns (e.g. admin_username)
  // into the camelCase shape the rest of the app already expects (adminUsername)
  // — so every existing function that reads SCHOOLS_DATA[id].adminUsername etc.
  // keeps working completely unchanged.
  useEffect(() => {
    let cancelled = false;
    async function loadSchools() {
      setSchoolsLoading(true);
      setSchoolsLoadError("");

      // Restore the Supabase Auth session from localStorage first.
      // This is fast (local read) and ensures RLS policies see a valid
      // auth.uid() for the subsequent schools query.
      const { data: sessionData } = await supabase.auth.getSession();
      const hasAuthSession = !!sessionData?.session;

      let data, error;

      if (hasAuthSession) {
        // Authenticated: RLS will filter to only this school's row automatically.
        ({ data, error } = await supabase.from("schools").select("*"));
      } else {
        // No auth session yet — check if there's a saved custom session.
        // If there is, we need to sign back into Supabase Auth first so RLS works.
        const saved = loadSession();
        if (saved && saved.role === "admin") {
          // Can't restore auth session without credentials — clear the saved
          // session and show the login screen. The user will log in fresh
          // which will establish a new Supabase Auth session via signInWithPassword.
          clearSession();
          setSchoolsLoading(false);
          return;
        }
        // Super admin or parent login doesn't need school rows pre-auth.
        // Just load an empty set and proceed to login screen.
        SCHOOLS_DATA = {};
        setSchoolsLoading(false);
        setSubscriptionRefresh(r => r + 1);
        return;
      }
      if (cancelled) return;
      if (error) {
        setSchoolsLoadError(error.message);
        setSchoolsLoading(false);
        return;
      }
      const loaded = {};
      (data || []).forEach(row => {
        loaded[row.id] = {
          id: row.id,
          userId: row.user_id || null,
          name: row.name,
          location: row.location,
          principal: row.principal,
          phone: row.phone,
          notifyEmail: row.notify_email || "",
          logo: row.logo || "🏫",
          schoolType: row.school_type || "secondary",
          streams: row.streams || {},
          setupComplete: row.setup_complete,
          adminUsername: row.admin_username,
          adminPassword: row.admin_password,
          plan: row.plan || "Starter",
          billingCycle: row.billing_cycle || "monthly",
          customPrice: row.custom_price,
          customPriceNote: row.custom_price_note || "",
          subscriptionStatus: row.subscription_status || "Active",
          isTrial: row.is_trial,
          trialActivated: row.trial_activated,
          trialStartDate: row.trial_start_date,
          nextBillingDate: row.next_billing_date,
          lastPaymentDate: row.last_payment_date,
          paymentNoticeFreeze: row.payment_notice_freeze || false,
          billingRef: row.billing_ref,
        };
      });
      SCHOOLS_DATA = loaded;
      setSchoolsLoading(false);
      setSubscriptionRefresh(r => r + 1);

      // ── Restore a saved login session, if there is one ────────────
      // Only restores if the saved school/role still genuinely exists in
      // the real data we just loaded — e.g. if a school was somehow
      // removed since the session was saved, this safely does nothing
      // rather than restoring into a broken state.
      const session = loadSession();
      if (session && isSessionExpired(session)) {
        // Too long since the last real activity — treat exactly like a
        // manual logout rather than restoring a stale session.
        clearSession();
      } else if (session) {
        if (session.role === "superadmin") {
          setIsSuperAdmin(true);
          setCurrentUser({ role: "superadmin" });
        } else if (session.role === "admin" && loaded[session.schoolId]) {
          setActiveSchoolId(session.schoolId);
          setCurrentUser({ role: "admin" });
          if (session.tab) setTab(session.tab);
          loadStudentsForSchool(session.schoolId);
        } else if (session.role === "parent" && loaded[session.schoolId]) {
          setActiveSchoolId(session.schoolId);
          setCurrentUser(session);
          if (session.tab) setTab(session.tab);
          loadStudentsForSchool(session.schoolId);
        } else {
          // Saved session no longer matches anything real — clear it
          // rather than leaving a stale, unusable session sitting around.
          clearSession();
        }
      }
    }
    loadSchools();
    return () => { cancelled = true; };
  }, []);

  // ── Detect a password-reset link being opened ───────────────────
  // The email sent by request-password-reset links back here with
  // ?reset_token=... in the URL. Checked once on load — if present, the
  // reset-completion screen takes over regardless of login state, since
  // the person clicking this link is, by definition, not logged in.
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const token = params.get("reset_token");
    if (token) setResetPasswordToken(token);
  }, []);

  // ── Auto-logout after a period of inactivity ────────────────────
  // Two parts: (1) any real activity (click/key/touch/scroll) bumps the
  // saved session's timestamp, and (2) a periodic check, while the app
  // stays open, logs the person out the moment too much time has passed
  // since the last real activity — not just on the next reload. logout()
  // itself is declared further down the file as a plain const, but since
  // this effect's callback only runs later (on a timer/event, after the
  // whole component body has already evaluated once), referencing it here
  // is safe — same pattern already used elsewhere in this file.
  useEffect(() => {
    if (!currentUser) return;
    const bump = () => touchSessionActivity();
    const events = ["click", "keydown", "touchstart", "scroll", "mousemove"];
    events.forEach(e => window.addEventListener(e, bump, { passive: true }));

    const interval = setInterval(() => {
      const session = loadSession();
      if (isSessionExpired(session)) {
        notify("You were logged out after a period of inactivity", "err");
        // Brief delay before actually switching screens — gives the toast
        // a real chance to paint before the view jumps to the login
        // screen, which has no toast element of its own to show it in.
        setTimeout(logout, 600);
      }
    }, 60 * 1000); // checked once a minute — frequent enough to feel responsive, cheap enough not to matter

    return () => {
      events.forEach(e => window.removeEventListener(e, bump));
      clearInterval(interval);
    };
  }, [currentUser]);

  // ── Keep the saved session's "tab" in sync with the real one ───
  // So that refreshing the page while on, say, the Students screen brings
  // you back to Students rather than always landing on the Dashboard.
  // Only writes once someone is actually logged in — merges onto whatever
  // session is already saved (role/schoolId/etc.) rather than overwriting
  // those.
  useEffect(() => {
    if (!currentUser) return;
    const existing = loadSession();
    if (existing) saveSession({ ...existing, tab });
  }, [tab, currentUser]);

  // ── Load real pending signups from Supabase ────────────────────
  // Same pattern as loadSchools above — runs once on mount, replaces the
  // fake demo signup requests with whatever's actually been submitted
  // through the real signup form and stored in Supabase.
  useEffect(() => {
    let cancelled = false;
    async function loadSuperAdminData() {
      // Load pending signups
      const { data: signups, error: signupsError } = await supabase.from("pending_signups").select("*").order("submitted_at", { ascending: false });
      if (cancelled) return;
      if (!signupsError) {
        setPendingSchools((signups || []).map(row => ({
          id: row.id, schoolName: row.school_name, location: row.location,
          principal: row.principal, phone: row.phone, email: row.email,
          students: row.students_estimate, schoolType: row.school_type,
          billingCycle: row.billing_cycle, username: row.requested_username,
          password: row.requested_password,
          submittedAt: row.submitted_at ? row.submitted_at.split("T")[0] : "",
          status: row.status,
        })));
      }
      // Load payment notices
      const { data: notices, error: noticesError } = await supabase.from("payment_notices").select("*").order("submitted_at", { ascending: false });
      if (cancelled) return;
      if (!noticesError) {
        setPaymentNotices((notices || []).map(row => ({
          id: row.id, schoolId: row.school_id, schoolName: row.school_name,
          billingRef: row.billing_ref, method: row.method, amount: row.amount,
          date: row.date, note: row.note, submittedAt: row.submitted_at,
          status: row.status,
        })));
      }
    }
    loadSuperAdminData();
    return () => { cancelled = true; };
  }, []);

  // Note: students and payments for the logged-in school are loaded by
  // loadStudentsForSchool (called from handleLogin and from the session-
  // restoration effect above), not by a separate effect watching
  // activeSchoolId here. An earlier version of this had two such effects;
  // they were removed because they duplicated loadStudentsForSchool and,
  // once activeSchoolId's default changed from the placeholder number 1 to
  // null (as part of fixing login-session persistence), began firing a
  // real Supabase query with school_id=null on every page load before
  // login — harmless (Supabase correctly rejected it) but noisy and
  // pointless, since the real loading path was already handling this
  // correctly elsewhere.

  // Persist currentTerm to localStorage so it survives page refreshes
  useEffect(() => {
    try { localStorage.setItem("feetrack_current_term", currentTerm); } catch {}
  }, [currentTerm]);

  const [superAdminTab, setSuperAdminTab] = useState("signups");
  const [expandedId, setExpandedId] = useState(null);
  const [expandedAlumni, setExpandedAlumni] = useState(null);
  const [toast, setToast] = useState(null);
  const [parentPin, setParentPin] = useState("");
  // Alumni: { schoolId: [...leavers] }
  // Starts empty — populated from the real Supabase students table (rows with
  // status != "active") by the loadStudentsForSchool function, not hardcoded
  // demo data. Alumni share the same students table as active students (see
  // the 04_alumni_status.sql migration) — split apart here by status, purely
  // for the UI's convenience, since the Alumni and Students pages are
  // separate screens that each expect their own list.
  const [allAlumni, setAllAlumni] = useState({});

  // Settings
  const [adminCreds, setAdminCreds] = useState({ username: "admin", password: "admin123" });
  const [pwForm, setPwForm] = useState({ currentPw: "", newPw: "", confirmPw: "" });
  const [pwError, setPwError] = useState("");

  // Super admin credentials & notification email
  const [superAdminCreds, setSuperAdminCreds] = useState({ username: "superadmin", password: "super123", notifyEmail: "" });

  // Platform payment details — where schools send subscription payments (super admin configurable)
  const [platformPayInfo, setPlatformPayInfo] = useState({
    momoNumber: "0772 000 000", momoName: "FeeTrack UG Ltd",
    bankName: "Stanbic Bank Uganda", bankAccount: "9030 0123 4567 8",
  });
  const [platformPayForm, setPlatformPayForm] = useState(null); // null = not editing
  const [superPwForm, setSuperPwForm] = useState({ currentPw: "", newPw: "", confirmPw: "", newUsername: "" });
  const [superPwError, setSuperPwError] = useState("");
  const [notifyEmailInput, setNotifyEmailInput] = useState("");
  const [schoolProfile, setSchoolProfile] = useState({ ...SCHOOLS_DATA[activeSchoolId] });
  const [streamsForm, setStreamsForm] = useState(null); // local edit buffer for stream config, null = not editing
  const [newStreamInput, setNewStreamInput] = useState({}); // { [className]: "text being typed" }
  const [backendUrl, setBackendUrl] = useState(""); // in-memory only — browser storage APIs aren't available in this sandbox
  const [backendStatus, setBackendStatus] = useState("unknown");

  // Editable fee structure & requirements
  const [feeStructure, setFeeStructure] = useState(JSON.parse(JSON.stringify(DEFAULT_FEE_STRUCTURE)));
  const [requirements, setRequirements] = useState(DEFAULT_REQUIREMENTS);
  const [feeEditCell, setFeeEditCell] = useState(null); // { cat, cls, field }
  const [feeEditVal, setFeeEditVal] = useState("");
  const [showAddReq, setShowAddReq] = useState(false);
  const [newReq, setNewReq] = useState({ name: "", cost: "", appliesTo: ["Day Scholar", "Boarder"], mandatory: true });
  const [showAddFeeItem, setShowAddFeeItem] = useState(null); // category being edited
  const [newFeeItemName, setNewFeeItemName] = useState("");
  const [newFeeItemAmt, setNewFeeItemAmt] = useState("");
  const [showEditPayment, setShowEditPayment] = useState(null); // { student, payment }
  const [showEditStudent, setShowEditStudent] = useState(null); // student being edited
  const [editPayAmt, setEditPayAmt] = useState("");
  const [showMoveAlumni, setShowMoveAlumni] = useState(null); // student being moved manually
  const [moveAlumniReason, setMoveAlumniReason] = useState("Transferred to another school");
  const [showBulkAlumni, setShowBulkAlumni] = useState(false);
  const [bulkAlumniSelected, setBulkAlumniSelected] = useState({}); // { [studentId]: true }
  const [bulkAlumniReason, setBulkAlumniReason] = useState("Transferred to another school");
  const [bulkAlumniMode, setBulkAlumniMode] = useState("checklist"); // "checklist" | "excel"
  const [bulkAlumniExcelRows, setBulkAlumniExcelRows] = useState([]);
  const [bulkAlumniExcelFileName, setBulkAlumniExcelFileName] = useState("");
  const [bulkAlumniExcelDone, setBulkAlumniExcelDone] = useState(false);
  const bulkAlumniFileRef = useRef(null);
  const [showPhotoUpload, setShowPhotoUpload] = useState(null); // student or staff record being photo-edited
  const [photoUploadType, setPhotoUploadType] = useState("student"); // "student" | "staff" — which list savePhoto/removePhoto writes to
  const [cameraActive, setCameraActive] = useState(false);
  const [cameraStream, setCameraStream] = useState(null);
  const videoRef = useRef(null);
  const canvasRef = useRef(null);
  const fileInputRef = useRef(null);

  // Redefine getStudentFee and getStatus using current state
  const getStudentFee = (student) => getStudentFeeStatic(student, feeStructure, requirements);
  const getStatus = (s, term) => getStatusStatic(s, term, feeStructure, requirements);
  // Single source of truth for balance — returns { termFee, arrears, totalDue, paidThisTerm, balance, status }
  const getBalance = (s, term) => getBalanceStatic(s, term, feeStructure, requirements);
  const totalFee = (cls, cat = "Day Scholar") => Object.values((feeStructure[cat] || feeStructure["Day Scholar"])[cls] || {}).reduce((a, b) => a + b, 0);

  const receiptRef = useRef();

  const notify = (msg, type = "ok") => { setToast({ msg, type }); setTimeout(() => setToast(null), 3500); };

  // ── Super Admin Activity Log ─────────────────────────────────────
  const logActivity = (action, detail) => {
    setActivityLog(prev => [{ id: `act-${Date.now()}-${Math.random().toString(36).slice(2,6)}`, action, detail, at: new Date().toISOString() }, ...prev].slice(0, 200));
  };

  // ── Real SMS sender — calls send-sms Edge Function ────────────
  const sendRealSMS = async (student, payment, schoolObj, balance) => {
    const message = `Dear ${student.parent}, UGX ${payment.amount.toLocaleString()} received for ${student.name} (${student.class}) on ${fmtDate(payment.date)}. Balance: UGX ${balance.toLocaleString()}. Ref: ${payment.id}. ${schoolObj.name.split(" ").slice(0,2).join(" ")}.`;
    const time = new Date().toLocaleTimeString();

    const rawPhone = student.phone || "";
    const phone = rawPhone.startsWith("+") ? rawPhone : "+256" + rawPhone.replace(/^0/, "");
    try {
      const { data, error } = await supabase.functions.invoke("send-sms", {
        body: {
          to: phone,
          message,
          student_name: student.name,
          school_id: activeSchoolId,
        },
      });
      const status = error ? `Error: ${error.message}` : (data?.status || "Sent ✓");
      setSmsLog(prev => [{ id: payment.id, to: student.phone, student: student.name, message, time, status }, ...prev]);
    } catch (err) {
      setSmsLog(prev => [{ id: payment.id, to: student.phone, student: student.name, message, time, status: `Error: ${err.message}` }, ...prev]);
    }
  };

  // ── Persist fee structure + requirements to Supabase ────────────
  // Uses upsert (insert or update) since each school has exactly one
  // config row. Called after any change to feeStructure or requirements.
  const saveSchoolConfig = async (newFeeStructure, newRequirements) => {
    const { error } = await supabase.from("school_config").upsert({
      school_id: activeSchoolId,
      fee_structure: newFeeStructure,
      requirements: newRequirements,
      updated_at: new Date().toISOString(),
    }, { onConflict: "school_id" });
    if (error) console.error("Could not save school config:", error.message);
  };

  // ── Add / Remove Fee Structure Line Items ──────────────────────
  const handleAddFeeItem = () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const name = newFeeItemName.trim().toLowerCase().replace(/\s+/g, "_");
    const amt = parseInt(newFeeItemAmt);
    if (!newFeeItemName.trim() || !amt || amt < 0) return notify("Enter a valid item name and amount", "err");
    const cat = showAddFeeItem;
    setFeeStructure(prev => {
      const updated = { ...prev, [cat]: { ...prev[cat] } };
      schoolClasses.forEach(cls => {
        updated[cat][cls] = { ...updated[cat][cls], [name]: amt };
      });
      saveSchoolConfig(updated, requirements);
      return updated;
    });
    notify(`"${newFeeItemName}" added to ${cat} fee structure (${fmt(amt)} for all classes)`);
    setShowAddFeeItem(null);
    setNewFeeItemName("");
    setNewFeeItemAmt("");
  };

  const handleRemoveFeeItem = (cat, field) => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    setConfirmDialog({
      title: "Remove Fee Item",
      message: `Remove "${field}" from ${cat} fee structure for all classes?`,
      danger: true,
      onConfirm: () => {
        setFeeStructure(prev => {
          const updated = { ...prev, [cat]: {} };
          schoolClasses.forEach(cls => {
            const { [field]: removed, ...rest } = prev[cat][cls] || {};
            updated[cat][cls] = rest;
          });
          saveSchoolConfig(updated, requirements);
          return updated;
        });
        notify(`"${field}" removed from ${cat} fee structure`);
      },
    });
  };

  // ── Fee Structure Edit ────────────────────────────────────────
  const saveFeeCell = () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    if (!feeEditCell) return;
    const { cat, cls, field } = feeEditCell;
    const val = parseInt(feeEditVal);
    if (!val || val < 0) return notify("Enter a valid amount", "err");
    setFeeStructure(prev => {
      const updated = {
        ...prev,
        [cat]: { ...prev[cat], [cls]: { ...prev[cat][cls], [field]: val } }
      };
      saveSchoolConfig(updated, requirements);
      return updated;
    });
    setFeeEditCell(null); setFeeEditVal("");
    notify(`${cat} ${cls} ${field} → ${fmt(val)}`);
  };

  // ── Requirements ──────────────────────────────────────────────
  const handleAddReq = () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    if (!newReq.name.trim() || !newReq.cost) return notify("Fill all fields", "err");
    setRequirements(prev => {
      const updated = [...prev, { id: `r${Date.now()}`, ...newReq, cost: parseInt(newReq.cost) }];
      saveSchoolConfig(feeStructure, updated);
      return updated;
    });
    setShowAddReq(false);
    setNewReq({ name: "", cost: "", appliesTo: ["Day Scholar", "Boarder"], mandatory: true });
    notify("Requirement added");
  };

  const deleteReq = (id) => {
    if (isReadOnly) return notify("Account is in read-only mode.", "err");
    setRequirements(prev => {
      const updated = prev.filter(r => r.id !== id);
      saveSchoolConfig(feeStructure, updated);
      return updated;
    });
    notify("Requirement removed");
  };

  const handleSignupSubmit = async () => {
    const { schoolName, location, principal, phone, email, students, schoolType, billingCycle, username, password, confirmPassword } = signupForm;
    if (!schoolName.trim() || !principal.trim() || !phone.trim() || !username.trim() || !password.trim()) {
      return notify("Please fill all required fields", "err");
    }
    if (password.length < 6) return notify("Password must be at least 6 characters", "err");
    if (password !== confirmPassword) return notify("Passwords do not match", "err");

    const { data, error } = await supabase.from("pending_signups").insert({
      school_name: schoolName, location, principal, phone, email,
      students_estimate: students, school_type: schoolType, billing_cycle: billingCycle || "monthly",
      requested_username: username, requested_password: password,
      status: "pending",
    }).select().single();

    if (error) {
      return notify(`Could not submit signup: ${error.message}`, "err");
    }

    const newPending = {
      id: data.id,
      schoolName, location, principal, phone, email, students, schoolType,
      username, password,
      submittedAt: data.submitted_at ? data.submitted_at.split("T")[0] : new Date().toISOString().split("T")[0],
      status: "pending",
    };
    setPendingSchools(prev => [...prev, newPending]);
    setSignupSubmitted(true);
  };

  const approveSchool = async (id) => {
    const school = pendingSchools.find(p => p.id === id);
    if (!school) return;
    const autoPlan = suggestPlan(school.students);
    const billingRef = `FT-${(school.schoolName || "SCHOOL").replace(/[^A-Za-z]/g, "").toUpperCase().slice(0, 6)}-${Date.now().toString().slice(-4)}`;

    // Create the real school row in Supabase first — id is generated by the
    // database (a uuid), not made up locally, so it's guaranteed to be a
    // real, valid, unique school id from the moment it exists.
    const { data: newSchoolRow, error: insertError } = await supabase.from("schools").insert({
      name: school.schoolName,
      location: school.location || "Uganda",
      principal: school.principal,
      phone: school.phone,
      notify_email: school.email || "",
      logo: "🏫",
      school_type: school.schoolType || "secondary",
      setup_complete: false,
      admin_username: school.username || "admin",
      admin_password: school.password || "school123",
      plan: autoPlan,
      billing_cycle: school.billingCycle || "monthly",
      subscription_status: "Active",
      next_billing_date: null, // trial starts on first login, not at approval
      last_payment_date: null,
      is_trial: true,
      trial_activated: false,
      billing_ref: billingRef,
    }).select().single();

    if (insertError) {
      return notify(`Could not approve school: ${insertError.message}`, "err");
    }

    // Mark the signup request as approved in Supabase too, so it doesn't
    // show up as still-pending if the page is reloaded.
    const { error: updateError } = await supabase.from("pending_signups").update({ status: "approved" }).eq("id", id);
    if (updateError) {
      // The school itself was created successfully — this second failure is
      // less critical, so warn rather than block, since the real work is done.
      notify(`School created, but couldn't update signup status: ${updateError.message}`, "err");
    }

    const newSchoolId = newSchoolRow.id;

    // Create a Supabase Auth account for this school immediately at approval
    // time, so they can log in right away without any manual migration step.
    // We reuse the migrate-schools-to-auth Edge Function which handles this
    // safely — it only processes schools with notify_email set and no user_id.
    if (school.email) {
      const { error: authMigrateError } = await supabase.functions.invoke("migrate-schools-to-auth");
      if (authMigrateError) {
        // Non-blocking — the school row was created successfully. They can
        // still log in via the old password fallback until the auth account
        // is created manually.
        console.error("Could not create auth account for new school:", authMigrateError.message);
      }
    }

    SCHOOLS_DATA[newSchoolId] = {
      id: newSchoolId, name: school.schoolName, location: school.location || "Uganda",
      principal: school.principal, phone: school.phone, logo: "🏫",
      notifyEmail: school.email || "",
      userId: null, // will be populated on next loadSchools() after auth account is created
      adminUsername: school.username || "admin", adminPassword: school.password || "school123",
      setupComplete: false,
      schoolType: school.schoolType || "secondary",
      plan: autoPlan, billingCycle: school.billingCycle || "monthly", customPrice: null, customPriceNote: "", subscriptionStatus: "Active",
      nextBillingDate: null, lastPaymentDate: null, isTrial: true, trialActivated: false,
      billingRef,
    };
    setSubscriptionRefresh(r => r + 1); // force re-render now that SCHOOLS_DATA changed
    setAllStudents(prev => ({ ...prev, [newSchoolId]: [] }));
    setAllExpenses(prev => ({ ...prev, [newSchoolId]: [] }));
    setAllAlumni(prev => ({ ...prev, [newSchoolId]: [] }));
    setPendingSchools(prev => prev.map(p => p.id === id ? { ...p, status: "approved" } : p));
    logActivity("School Approved", `${school.schoolName} approved on ${autoPlan} plan (${SCHOOL_TYPES[school.schoolType]?.label || "Secondary"}, ~${school.students || "0"} students)`);
    notify(`✓ ${school.schoolName} approved on the ${autoPlan} plan (based on ${school.students || "0"} students)! Login: ${school.username || "admin"} · 30-day trial starts on first login`);
  };

  const rejectSchool = (id) => {
    const school = pendingSchools.find(p => p.id === id);
    setConfirmDialog({
      title: "Reject Signup Request",
      message: `Reject ${school ? `"${school.schoolName}"'s` : "this school's"} signup request? They will not be notified automatically.`,
      danger: true,
      onConfirm: () => {
        setPendingSchools(prev => prev.map(p => p.id === id ? { ...p, status: "rejected" } : p));
        logActivity("Signup Rejected", school ? school.schoolName : id);
        notify("School signup rejected");
      },
    });
  };

  // ── Subscription / Billing Handlers (Super Admin) ──────────────
  const markSubscriptionPaid = async (schoolId) => {
    let sch = SCHOOLS_DATA[schoolId];
    if (!sch) {
      // Super admin — school data not in memory, fetch from Supabase
      const { data, error } = await supabase.from("schools").select("*").eq("id", schoolId).single();
      if (error || !data) return notify("Could not load school data", "err");
      sch = {
        id: data.id, name: data.name, plan: data.plan || "Starter",
        billingCycle: data.billing_cycle || "monthly",
        customPrice: data.custom_price,
      };
    }
    const today = new Date();
    const nextDue = new Date();
    const billing = getBillingInfo(sch.plan, sch.billingCycle, sch.customPrice);
    nextDue.setDate(nextDue.getDate() + billing.cycleDays);
    const updates = {
      subscription_status: "Active",
      last_payment_date: today.toISOString().split("T")[0],
      next_billing_date: nextDue.toISOString().split("T")[0],
      is_trial: false,
      payment_notice_freeze: false,
    };
    const { error: updateError } = await supabase.from("schools").update(updates).eq("id", schoolId);
    if (updateError) return notify(`Could not update school: ${updateError.message}`, "err");
    if (SCHOOLS_DATA[schoolId]) {
      SCHOOLS_DATA[schoolId] = {
        ...SCHOOLS_DATA[schoolId],
        subscriptionStatus: "Active",
        lastPaymentDate: today.toISOString().split("T")[0],
        nextBillingDate: nextDue.toISOString().split("T")[0],
        isTrial: false,
        paymentNoticeFreeze: false,
      };
    }
    setSubscriptionRefresh(r => r + 1);
    // Mark payment notices as confirmed in Supabase
    supabase.from("payment_notices").update({ status: "confirmed" })
      .eq("school_id", schoolId).eq("status", "pending")
      .then(({ error }) => { if (error) console.error("Could not update payment notices:", error.message); });
    setPaymentNotices(prev => prev.map(n => n.schoolId === schoolId && n.status === "pending" ? { ...n, status: "confirmed" } : n));
    logActivity("Payment Confirmed", `${sch.name} marked as paid (${sch.plan}, ${fmt(billing.price)}${billing.periodLabel}) — next due ${fmtDate(nextDue.toISOString().split("T")[0])}`);
    notify(`✓ ${sch.name} marked as paid — next billing date ${fmtDate(nextDue.toISOString().split("T")[0])}`);
  };

  // Manual plan change by Super Admin. Clears any custom negotiated price, since a
  // discount agreed for one plan's price shouldn't silently carry over to a different
  // plan — Super Admin can set a new custom price for the new plan if needed.
  const changeSchoolPlan = (schoolId, newPlan) => {
    const sch = SCHOOLS_DATA[schoolId];
    if (!sch) return;
    SCHOOLS_DATA[schoolId] = { ...sch, plan: newPlan, customPrice: null, customPriceNote: "" };
    setSubscriptionRefresh(r => r + 1);
    const billing = getBillingInfo(newPlan, sch.billingCycle);
    logActivity("Plan Changed", `${sch.name}: ${sch.plan} → ${newPlan} (${fmt(billing.price)}${billing.periodLabel})`);
    notify(`${sch.name} plan changed to ${newPlan} (${fmt(billing.price)}${billing.periodLabel})`);
  };

  // Lets a school switch between Monthly and Per Term billing for themselves.
  // If they had a custom negotiated price, it's cleared here — a discount agreed for
  // one cycle (e.g. a monthly rate) shouldn't silently carry over to a different cycle
  // with different standard pricing. Super Admin can always re-apply a new custom price
  // for the new cycle from the Billing tab if the same discount should continue.
  const changeBillingCycle = (schoolId, newCycle) => {
    const sch = SCHOOLS_DATA[schoolId];
    if (!sch) return;
    const hadCustomPrice = sch.customPrice !== null && sch.customPrice !== undefined;
    SCHOOLS_DATA[schoolId] = { ...sch, billingCycle: newCycle, customPrice: null, customPriceNote: "" };
    setSubscriptionRefresh(r => r + 1);
    const billing = getBillingInfo(sch.plan, newCycle);
    notify(`Billing cycle changed to ${newCycle === "term" ? "Per Term" : "Monthly"} (${fmt(billing.price)}${billing.periodLabel})${hadCustomPrice ? " — your previous custom rate was cleared; contact support if you'd like a new one set for this cycle" : ""} — takes effect on your next payment`);
  };

  // ── Super Admin: set or clear a one-off negotiated price for a single school ──
  // This is completely independent of PLANS — it never changes what any other school
  // pays. The school keeps their plan and its features exactly as before; only the
  // amount they're billed (per their existing billing cycle) is overridden.
  const setCustomPrice = (schoolId, price, note) => {
    const sch = SCHOOLS_DATA[schoolId];
    if (!sch) return;
    const amount = parseInt(String(price).replace(/,/g, ""));
    if (!amount || amount <= 0) return notify("Enter a valid custom price", "err");
    const standard = getBillingInfo(sch.plan, sch.billingCycle).price;
    const periodLabel = getBillingInfo(sch.plan, sch.billingCycle).periodLabel;
    SCHOOLS_DATA[schoolId] = { ...sch, customPrice: amount, customPriceNote: note || "" };
    setSubscriptionRefresh(r => r + 1);
    logActivity("Custom Price Set", `${sch.name}: ${fmt(amount)}${periodLabel} (standard price is ${fmt(standard)})${note ? ` — ${note}` : ""}`);
    notify(`✓ ${sch.name} will now be billed ${fmt(amount)} instead of the standard ${fmt(standard)} — this does not affect any other school`);
  };

  const clearCustomPrice = (schoolId) => {
    const sch = SCHOOLS_DATA[schoolId];
    if (!sch) return;
    const standard = getBillingInfo(sch.plan, sch.billingCycle).price;
    const periodLabel = getBillingInfo(sch.plan, sch.billingCycle).periodLabel;
    SCHOOLS_DATA[schoolId] = { ...sch, customPrice: null, customPriceNote: "" };
    setSubscriptionRefresh(r => r + 1);
    logActivity("Custom Price Removed", `${sch.name}: reverted to standard ${sch.plan} pricing (${fmt(standard)}${periodLabel})`);
    notify(`${sch.name}'s custom price removed — back to standard ${sch.plan} pricing (${fmt(standard)})`);
  };

  // Auto-upgrade a school's plan if their enrolled student count now exceeds their current plan's limit.
  // Never downgrades automatically — only upgrades when a real limit is crossed.
  const checkAutoUpgrade = (schoolId, newStudentCount) => {
    const sch = SCHOOLS_DATA[schoolId];
    if (!sch) return;
    const required = minPlanForCount(newStudentCount);
    if (PLAN_ORDER.indexOf(required) > PLAN_ORDER.indexOf(sch.plan)) {
      SCHOOLS_DATA[schoolId] = { ...sch, plan: required };
      setSubscriptionRefresh(r => r + 1);
      const billing = getBillingInfo(required, sch.billingCycle);
      logActivity("Auto-Upgrade", `${sch.name}: ${sch.plan} → ${required} (${newStudentCount} students exceeded ${sch.plan} limit of ${PLANS[sch.plan].maxStudents})`);
      notify(`📈 ${sch.name} automatically upgraded to ${required} plan (${fmt(billing.price)}${billing.periodLabel}) — now has ${newStudentCount} students, exceeding the ${sch.plan} limit of ${PLANS[sch.plan].maxStudents}`);
    }
  };

  // ── "I've Sent Payment" Confirmation (school notifies super admin) ──
  const submitPaymentNotice = async () => {
    const amt = parseInt(paymentConfirmForm.amount.replace(/,/g, ""));
    if (!amt || amt <= 0) return notify("Enter the amount you sent", "err");
    const { data: inserted, error } = await supabase.from("payment_notices").insert({
      school_id: activeSchoolId, school_name: school.name, billing_ref: school.billingRef || "",
      method: paymentConfirmForm.method, amount: amt, date: paymentConfirmForm.date,
      note: paymentConfirmForm.note, status: "pending",
    }).select().single();
    if (error) return notify(`Could not submit notice: ${error.message}`, "err");
    const notice = {
      id: inserted.id, schoolId: activeSchoolId, schoolName: school.name, billingRef: school.billingRef,
      method: inserted.method, amount: inserted.amount, date: inserted.date,
      note: inserted.note, submittedAt: inserted.submitted_at, status: "pending",
    };
    setPaymentNotices(prev => [notice, ...prev]);
    // Freeze countdown while notice is under review
    await supabase.from("schools").update({ payment_notice_freeze: true }).eq("id", activeSchoolId);
    SCHOOLS_DATA[activeSchoolId] = { ...SCHOOLS_DATA[activeSchoolId], paymentNoticeFreeze: true };
    setSubscriptionRefresh(r => r + 1);
    setShowPaymentConfirm(false);
    setPaymentConfirmForm({ method: "MTN MoMo", amount: "", date: new Date().toISOString().split("T")[0], note: "" });
    notify("✓ Thanks! Your account access is protected while we confirm this payment.");
  };

  // Payment was NOT found / not received — unfreeze and flag it so the school sees it
  const rejectPaymentNotice = async (id) => {
    const notice = paymentNotices.find(n => n.id === id);
    if (!notice) return;
    await supabase.from("payment_notices").update({ status: "not_found" }).eq("id", id);
    await supabase.from("schools").update({ payment_notice_freeze: false }).eq("id", notice.schoolId);
    setPaymentNotices(prev => prev.map(n => n.id === id ? { ...n, status: "not_found" } : n));
    if (SCHOOLS_DATA[notice.schoolId]) {
      SCHOOLS_DATA[notice.schoolId] = { ...SCHOOLS_DATA[notice.schoolId], paymentNoticeFreeze: false };
    }
    setSubscriptionRefresh(r => r + 1);
    logActivity("Payment Notice Rejected", `${notice.schoolName} — ${fmt(notice.amount)} via ${notice.method} not found`);
    notify(`${notice.schoolName} notified — payment not found, protection removed`);
  };

  // ── Automatic Subscription Renewal via Bank/MoMo Statement ─────
  const handleSubPayFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setSubPayFileName(file.name);
    setSubPayResults(null);

    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

        let headerIdx = 0;
        for (let i = 0; i < Math.min(10, raw.length); i++) {
          const row = raw[i].map(c => String(c || "").toLowerCase());
          if (row.some(c => c.includes("amount") || c.includes("credit") || c.includes("deposit"))) {
            headerIdx = i; break;
          }
        }
        const headers = raw[headerIdx].map(h => String(h || "").toLowerCase().trim());
        const dataRows = raw.slice(headerIdx + 1).filter(r => r.some(c => c));
        const colIdx = (keywords) => headers.findIndex(h => keywords.some(k => h.includes(k)));
        const dateCol = colIdx(["date", "trans date", "value date"]);
        const amtCol = colIdx(["credit", "deposit", "amount", "cr"]);
        const refCol = colIdx(["reference", "ref", "narration", "description", "particulars", "details"]);

        const parsed = dataRows.map((row, i) => {
          const amtRaw = String(row[amtCol] || "").replace(/[^0-9.]/g, "");
          const amount = parseFloat(amtRaw) || 0;
          const reference = String(row[refCol] || "").trim();
          const date = row[dateCol] ? String(row[dateCol]).trim() : "";
          return { rowIndex: i, date, amount, reference };
        }).filter(r => r.amount > 0);

        setSubPayRows(parsed);
        notify(`✓ ${parsed.length} credit entries loaded from statement`);
      } catch (err) {
        notify("Could not read file. Make sure it's a valid Excel or CSV file.", "err");
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleSubPayReconcile = () => {
    const results = { matched: [], unmatched: [] };

    subPayRows.forEach(row => {
      const refUpper = row.reference.toUpperCase().replace(/\s+/g, "");
      // Find a school whose billingRef appears in the reference
      const matchedSchool = Object.values(SCHOOLS_DATA).find(s =>
        s.billingRef && refUpper.includes(s.billingRef.toUpperCase().replace(/\s+/g, ""))
      );

      if (matchedSchool) {
        const today = new Date();
        const nextDue = new Date();
        const prevStatus = getSubscriptionInfo(matchedSchool).status;
        const prevPlan = matchedSchool.plan;
        const cycle = matchedSchool.billingCycle || "monthly";

        // ── Auto-detect plan switch: if amount paid matches a DIFFERENT plan's STANDARD
        // price exactly, switch automatically. This check intentionally ignores any custom
        // price the school might have — a custom rate is a discount on their CURRENT plan,
        // not a signal they're trying to switch plans. If the school has a custom price and
        // pays exactly that amount, we keep them on their current plan further down. ──
        let newPlan = matchedSchool.plan;
        const exactPlanMatch = Object.keys(PLANS).find(p => getBillingInfo(p, cycle).price === row.amount);
        if (exactPlanMatch && exactPlanMatch !== matchedSchool.plan) {
          newPlan = exactPlanMatch;
        }
        // Expected amount: if staying on their current plan AND they have a custom price,
        // use that custom price. Otherwise use the new/standard plan price.
        const billing = newPlan === matchedSchool.plan
          ? getBillingInfo(newPlan, cycle, matchedSchool.customPrice)
          : getBillingInfo(newPlan, cycle);
        const expectedAmount = billing.price;
        nextDue.setDate(nextDue.getDate() + billing.cycleDays);

        const planSwitched = newPlan !== matchedSchool.plan;
        SCHOOLS_DATA[matchedSchool.id] = {
          ...matchedSchool,
          plan: newPlan,
          subscriptionStatus: "Active",
          lastPaymentDate: today.toISOString().split("T")[0],
          nextBillingDate: nextDue.toISOString().split("T")[0],
          isTrial: false,
          paymentNoticeFreeze: false,
          ...(planSwitched ? { customPrice: null, customPriceNote: "" } : {}),
        };
        // Persist subscription update to Supabase
        supabase.from("schools").update({
          plan: newPlan,
          subscription_status: "Active",
          last_payment_date: today.toISOString().split("T")[0],
          next_billing_date: nextDue.toISOString().split("T")[0],
          is_trial: false,
          payment_notice_freeze: false,
          ...(planSwitched ? { custom_price: null, custom_price_note: "" } : {}),
        }).eq("id", matchedSchool.id).then(({ error }) => {
          if (error) console.error("Could not persist subscription update for", matchedSchool.name, error.message);
        });
        setPaymentNotices(prev => prev.map(n => n.schoolId === matchedSchool.id && n.status === "pending" ? { ...n, status: "confirmed" } : n));

        results.matched.push({
          ...row, school: matchedSchool.name, schoolId: matchedSchool.id,
          plan: newPlan, prevPlan, planChanged: newPlan !== prevPlan,
          expectedAmount, prevStatus, billingCycle: cycle,
          newDueDate: nextDue.toISOString().split("T")[0],
          amountMatch: row.amount >= expectedAmount,
        });
      } else {
        results.unmatched.push(row);
      }
    });

    setSubscriptionRefresh(r => r + 1);
    setSubPayResults(results);
    if (results.matched.length > 0) {
      const planChanges = results.matched.filter(m => m.planChanged).length;
      logActivity("Bank Reconciliation", `${results.matched.length} school(s) reactivated${planChanges > 0 ? `, ${planChanges} plan change(s)` : ""}: ${results.matched.map(m => m.school).join(", ")}`);
      notify(`✓ ${results.matched.length} school(s) automatically reactivated!${planChanges > 0 ? ` ${planChanges} plan change(s) applied.` : ""}`);
    } else {
      notify("No schools matched — check billing reference codes", "err");
    }
  };

  const handleBankFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setBankFileName(file.name);
    setBankImportDone(false);
    setBankMatched({});

    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

        // Find header row (look for keywords: date, amount, reference, narration)
        let headerIdx = 0;
        for (let i = 0; i < Math.min(10, raw.length); i++) {
          const row = raw[i].map(c => String(c || "").toLowerCase());
          if (row.some(c => c.includes("amount") || c.includes("credit") || c.includes("deposit"))) {
            headerIdx = i; break;
          }
        }

        const headers = raw[headerIdx].map(h => String(h || "").toLowerCase().trim());
        const dataRows = raw.slice(headerIdx + 1).filter(r => r.some(c => c));

        // Map columns
        const colIdx = (keywords) => headers.findIndex(h => keywords.some(k => h.includes(k)));
        const dateCol   = colIdx(["date", "trans date", "value date"]);
        const amtCol    = colIdx(["credit", "deposit", "amount", "cr"]);
        const refCol    = colIdx(["reference", "ref", "narration", "description", "particulars", "details"]);

        const parsed = dataRows.map((row, i) => {
          const amtRaw = String(row[amtCol] || "").replace(/[^0-9.]/g, "");
          const amount = parseFloat(amtRaw) || 0;
          const reference = String(row[refCol] || "").trim();
          const date = row[dateCol] ? String(row[dateCol]).trim() : "";
          return { rowIndex: i, date, amount, reference, raw: row };
        }).filter(r => r.amount > 0); // only rows with a credit amount

        // Auto-match: try to find student by name in reference
        const matched = {};
        parsed.forEach(r => {
          const refLower = r.reference.toLowerCase();
          const match = termStudents.find(s =>
            s.name.toLowerCase().split(" ").some(part => part.length > 2 && refLower.includes(part))
          );
          if (match) matched[r.rowIndex] = match.id;
        });

        setBankRows(parsed);
        setBankMatched(matched);
        notify(`✓ ${parsed.length} bank entries loaded · ${Object.keys(matched).length} auto-matched`);
      } catch (err) {
        notify("Could not read file. Make sure it is an Excel (.xlsx) or CSV file.", "err");
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleBankImport = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const toInsert = [];
    const currentStudents = allStudents[activeSchoolId] || [];

    bankRows.forEach(row => {
      const studentId = bankMatched[row.rowIndex];
      if (!studentId) return;
      const student = currentStudents.find(s => s.id === studentId);
      if (!student) return;
      const alreadyExists = student.payments.some(p =>
        p.amount === row.amount && p.method === "Bank" && p.term === currentTerm
      );
      if (alreadyExists) return;
      toInsert.push({
        school_id: activeSchoolId, student_id: studentId,
        student_name: student.name, term: currentTerm,
        amount: row.amount, method: "Bank",
        received_by: "Bank Import",
        payment_date: row.date || new Date().toISOString().split("T")[0],
        receipt_no: nextRcpt(),
      });
    });

    if (toInsert.length === 0) return notify("No new payments to import", "err");

    const { data: inserted, error } = await supabase.from("payments").insert(toInsert).select();
    if (error) return notify(`Could not import payments: ${error.message}`, "err");

    setAllStudents(prev => ({
      ...prev,
      [activeSchoolId]: prev[activeSchoolId].map(s => {
        const newPays = (inserted || []).filter(p => p.student_id === s.id).map(p => ({
          id: p.receipt_no, dbId: p.id, date: p.payment_date,
          amount: p.amount, method: p.method, receivedBy: p.received_by, term: p.term,
        }));
        return newPays.length > 0 ? { ...s, payments: [...s.payments, ...newPays] } : s;
      })
    }));
    setBankImportDone(true);
    notify(`✓ ${inserted.length} bank payments imported successfully`);
  };
  const handleFilePhoto = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      savePhoto(showPhotoUpload.id, ev.target.result);
    };
    reader.readAsDataURL(file);
  };

  const startCamera = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "user" }, audio: false });
      setCameraStream(stream);
      setCameraActive(true);
      setTimeout(() => { if (videoRef.current) videoRef.current.srcObject = stream; }, 100);
    } catch (err) {
      notify("Camera not available: " + err.message, "err");
    }
  };

  const stopCamera = () => {
    if (cameraStream) { cameraStream.getTracks().forEach(t => t.stop()); setCameraStream(null); }
    setCameraActive(false);
  };

  const capturePhoto = () => {
    if (!videoRef.current || !canvasRef.current) return;
    const canvas = canvasRef.current;
    canvas.width = videoRef.current.videoWidth;
    canvas.height = videoRef.current.videoHeight;
    canvas.getContext("2d").drawImage(videoRef.current, 0, 0);
    const dataUrl = canvas.toDataURL("image/jpeg", 0.8);
    savePhoto(showPhotoUpload.id, dataUrl);
    stopCamera();
  };

  // Uploads a student's photo to Supabase Storage and saves the resulting
  // public URL on their row — rather than saving the raw base64 image data
  // directly on the row, which would make every single students-list load
  // (which happens on login, after every enrollment, etc.) drag along every
  // student's full photo data even when no photo is being shown yet.
  //
  // Staff photos are NOT included in this — there's no real Supabase table
  // for staff yet (that's still a separate, not-yet-migrated piece), so
  // staff photos deliberately keep the old in-memory-only behavior for now
  // and will still reset on refresh until staff itself gets migrated.
  const savePhoto = async (recordId, dataUrl) => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");

    // Convert base64 data URL to binary Blob for Supabase Storage upload
    const blob = await (await fetch(dataUrl)).blob();
    const ext = blob.type === "image/png" ? "png" : "jpg";

    if (photoUploadType === "staff") {
      const path = `staff-${recordId}-${Date.now()}.${ext}`;
      const { error: uploadError } = await supabase.storage.from("photos").upload(path, blob, {
        contentType: blob.type, upsert: true,
      });
      if (uploadError) return notify(`Could not upload photo: ${uploadError.message}`, "err");
      const { data: urlData } = supabase.storage.from("photos").getPublicUrl(path);
      const publicUrl = urlData.publicUrl;
      const { error: updateError } = await supabase.from("staff").update({ photo_url: publicUrl }).eq("id", recordId);
      if (updateError) return notify(`Photo uploaded but could not be linked: ${updateError.message}`, "err");
      setAllStaff(prev => ({
        ...prev,
        [activeSchoolId]: prev[activeSchoolId].map(s => s.id === recordId ? { ...s, photo: publicUrl } : s)
      }));
      setShowPhotoUpload(null);
      setCameraActive(false);
      notify("Photo saved successfully ✓");
      return;
    }

    const path = `students-${recordId}-${Date.now()}.${ext}`;
    const { error: uploadError } = await supabase.storage.from("photos").upload(path, blob, {
      contentType: blob.type, upsert: true,
    });
    if (uploadError) {
      return notify(`Could not upload photo: ${uploadError.message}`, "err");
    }
    const { data: urlData } = supabase.storage.from("photos").getPublicUrl(path);
    const publicUrl = urlData.publicUrl;

    const { error: updateError } = await supabase.from("students").update({ photo_url: publicUrl }).eq("id", recordId);
    if (updateError) {
      return notify(`Photo uploaded but could not be linked to the record: ${updateError.message}`, "err");
    }

    setAllStudents(prev => ({
      ...prev,
      [activeSchoolId]: prev[activeSchoolId].map(s => s.id === recordId ? { ...s, photo: publicUrl } : s)
    }));
    setShowPhotoUpload(null);
    setCameraActive(false);
    notify("Photo saved successfully ✓");
  };

  const removePhoto = async (recordId) => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");

    if (photoUploadType === "staff") {
      const { error } = await supabase.from("staff").update({ photo_url: null }).eq("id", recordId);
      if (error) return notify(`Could not remove photo: ${error.message}`, "err");
      setAllStaff(prev => ({
        ...prev,
        [activeSchoolId]: prev[activeSchoolId].map(s => s.id === recordId ? { ...s, photo: null } : s)
      }));
      setShowPhotoUpload(null);
      setCameraActive(false);
      notify("Photo removed");
      return;
    }

    const { error: updateError } = await supabase.from("students").update({ photo_url: null }).eq("id", recordId);
    if (updateError) {
      return notify(`Could not remove photo: ${updateError.message}`, "err");
    }
    setAllStudents(prev => ({
      ...prev,
      [activeSchoolId]: prev[activeSchoolId].map(s => s.id === recordId ? { ...s, photo: null } : s)
    }));
    setShowPhotoUpload(null);
    notify("Photo removed");
  };

  // ── Edit / Delete Payment ─────────────────────────────────────
  const handleEditPayment = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const amt = parseInt(editPayAmt);
    if (!amt || amt <= 0) return notify("Enter valid amount", "err");
    const { student, payment } = showEditPayment;

    const { error: updateError } = await supabase.from("payments")
      .update({ amount: amt })
      .eq("school_id", activeSchoolId)
      .eq("receipt_no", payment.id);
    if (updateError) {
      return notify(`Could not correct payment: ${updateError.message}`, "err");
    }

    setAllStudents(prev => ({
      ...prev,
      [activeSchoolId]: prev[activeSchoolId].map(s =>
        s.id === student.id
          ? { ...s, payments: s.payments.map(p => p.id === payment.id ? { ...p, amount: amt } : p) }
          : s
      )
    }));
    notify(`Payment ${payment.id} corrected to ${fmt(amt)}`);
    setShowEditPayment(null); setEditPayAmt("");
  };

  const handleDeletePayment = (student, paymentId) => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    setConfirmDialog({
      title: "Delete Payment",
      message: `Delete payment ${paymentId}? This cannot be undone.`,
      danger: true,
      onConfirm: async () => {
        const { error: deleteError } = await supabase.from("payments")
          .delete()
          .eq("school_id", activeSchoolId)
          .eq("receipt_no", paymentId);
        if (deleteError) {
          return notify(`Could not delete payment: ${deleteError.message}`, "err");
        }
        setAllStudents(prev => ({
          ...prev,
          [activeSchoolId]: prev[activeSchoolId].map(s =>
            s.id === student.id ? { ...s, payments: s.payments.filter(p => p.id !== paymentId) } : s
          )
        }));
        notify(`Payment ${paymentId} deleted`);
      },
    });
  };


  // ── Manual Move to Alumni (mid-term transfer/dropout) ─────────
  const handleMoveToAlumni = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    if (!showMoveAlumni) return;
    const s = showMoveAlumni;
    const debt = getBalance(s, currentTerm).balance;
    const leftYear = currentTerm.split(", ")[1] || promotionYear;
    const leftNote = moveAlumniReason || "Left mid-term";

    const { error: updateError } = await supabase.from("students").update({
      status: "Transferred",
      left_class: s.class,
      left_year: leftYear,
      outstanding_debt: debt,
      left_note: leftNote,
    }).eq("id", s.id);
    if (updateError) {
      return notify(`Could not move student to Alumni: ${updateError.message}`, "err");
    }

    const alumniRecord = {
      ...s,
      status: "Transferred",
      leftClass: s.class,
      leftYear,
      outstandingDebt: debt,
      leftNote,
    };
    setAllStudents(prev => ({
      ...prev,
      [activeSchoolId]: prev[activeSchoolId].filter(st => st.id !== s.id),
    }));
    setAllAlumni(prev => ({
      ...prev,
      [activeSchoolId]: [...(prev[activeSchoolId] || []), alumniRecord],
    }));
    notify(`${s.name} moved to Alumni — ${moveAlumniReason}`);
    setShowMoveAlumni(null);
    setMoveAlumniReason("Transferred to another school");
  };

  // ── Bulk Move to Alumni ──────────────────────────────────────────
  const handleBulkMoveToAlumni = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const ids = Object.keys(bulkAlumniSelected).filter(id => bulkAlumniSelected[id]);
    if (ids.length === 0) return notify("Select at least one student", "err");
    const sid = activeSchoolId;
    const current = allStudents[sid] || [];
    const toMove = current.filter(s => ids.includes(String(s.id)));
    const staying = current.filter(s => !ids.includes(String(s.id)));
    const leftYear = currentTerm.split(", ")[1] || promotionYear;
    const leftNote = bulkAlumniReason || "Left mid-term";

    const updateResults = await Promise.all(toMove.map(s => {
      const debt = getBalance(s, currentTerm).balance;
      return supabase.from("students").update({
        status: "Transferred",
        left_class: s.class,
        left_year: leftYear,
        outstanding_debt: debt,
        left_note: leftNote,
      }).eq("id", s.id);
    }));
    const failedCount = updateResults.filter(r => r.error).length;
    if (failedCount > 0) {
      notify(`${failedCount} of ${toMove.length} students could not be moved — please try again for those`, "err");
    }

    const alumniRecords = toMove.map(s => {
      const debt = getBalance(s, currentTerm).balance;
      return {
        ...s,
        status: "Transferred",
        leftClass: s.class,
        leftYear,
        outstandingDebt: debt,
        leftNote,
      };
    });

    setAllStudents(prev => ({ ...prev, [sid]: staying }));
    setAllAlumni(prev => ({ ...prev, [sid]: [...(prev[sid] || []), ...alumniRecords] }));
    notify(`✓ ${alumniRecords.length} student(s) moved to Alumni — ${bulkAlumniReason}`);
    setShowBulkAlumni(false);
    setBulkAlumniSelected({});
    setBulkAlumniReason("Transferred to another school");
  };

  // ── Bulk Move to Alumni via Excel ────────────────────────────────
  const downloadBulkAlumniTemplate = () => {
    const wb = XLSX.utils.book_new();
    const sample = termStudents.slice(0, 10).map(s => ({
      Name: s.name, Class: classLabel(s.class, s.stream), "Move? (Yes/No)": "", Reason: "",
    }));
    const ws = XLSX.utils.json_to_sheet(sample.length > 0 ? sample : [
      { Name: "Okello James", Class: "S2", "Move? (Yes/No)": "Yes", Reason: "Transferred to another school" },
    ]);
    XLSX.utils.book_append_sheet(wb, ws, "Students");

    const instructions = [
      { Column: "Name", Instructions: "Must match the student's name exactly as it appears in Student Register" },
      { Column: "Class", Instructions: "For reference only — not used for matching" },
      { Column: "Move? (Yes/No)", Instructions: 'Type "Yes" for every student who should be moved to Alumni. Leave blank or "No" to keep them active.' },
      { Column: "Reason", Instructions: "Optional — e.g. Transferred, Dropped out, Withdrawn. If blank, defaults to the reason selected in the app." },
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(instructions), "Instructions");

    XLSX.writeFile(wb, `FeeTrack_Bulk_Alumni_${currentTerm.replace(/[, ]/g, "_")}.xlsx`);
    notify("Template downloaded ✓ (pre-filled with your current student list)");
  };

  const handleBulkAlumniFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setBulkAlumniExcelFileName(file.name);
    setBulkAlumniExcelDone(false);

    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });

        const parsed = raw.map((row, i) => {
          const get = (...keys) => {
            for (const k of keys) {
              const found = Object.keys(row).find(h => h.toLowerCase().trim() === k.toLowerCase());
              if (found && String(row[found]).trim() !== "") return String(row[found]).trim();
            }
            return "";
          };
          const name = get("name", "student name", "full name");
          const cls = get("class", "form");
          const moveFlag = get("move? (yes/no)", "move", "move?", "transfer");
          const reason = get("reason", "note");

          const candidates = termStudents.filter(s => s.name.toLowerCase().trim() === name.toLowerCase().trim());
          let matchedStudent = null;
          let matchStatus = "not_found";
          if (candidates.length === 1) { matchedStudent = candidates[0]; matchStatus = "matched"; }
          else if (candidates.length > 1) {
            if (cls) {
              const exact = candidates.find(s => classLabel(s.class, s.stream).toUpperCase() === cls.toUpperCase())
                || candidates.find(s => s.class.toUpperCase() === cls.toUpperCase());
              if (exact) { matchedStudent = exact; matchStatus = "matched"; }
              else matchStatus = "ambiguous";
            } else {
              matchStatus = "ambiguous";
            }
          }

          const shouldMove = /^(yes|y|true|1)$/i.test(moveFlag);

          return {
            rowIndex: i, name, moveFlag, reason,
            matchedStudent, matchStatus, shouldMove,
            valid: matchStatus === "matched" && shouldMove,
          };
        }).filter(r => r.name);

        setBulkAlumniExcelRows(parsed);
        const toMove = parsed.filter(r => r.valid).length;
        notify(`✓ ${parsed.length} rows found — ${toMove} marked to move to Alumni`);
      } catch (err) {
        notify("Could not read file. Make sure it's a valid Excel (.xlsx) or CSV file.", "err");
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleBulkAlumniExcelImport = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const validRows = bulkAlumniExcelRows.filter(r => r.valid);
    if (validRows.length === 0) return notify("No students marked to move", "err");
    const sid = activeSchoolId;
    const current = allStudents[sid] || [];
    const moveIds = validRows.map(r => r.matchedStudent.id);
    const toMove = current.filter(s => moveIds.includes(s.id));
    const staying = current.filter(s => !moveIds.includes(s.id));

    const alumniRecords = toMove.map(s => {
      const row = validRows.find(r => r.matchedStudent.id === s.id);
      const debt = getBalance(s, currentTerm).balance;
      return {
        ...s, status: "Transferred",
        leftClass: s.class,
        leftYear: currentTerm.split(", ")[1] || promotionYear,
        outstandingDebt: debt,
        leftNote: row.reason || bulkAlumniReason || "Left mid-term",
      };
    });

    // Persist status changes to Supabase
    await Promise.all(alumniRecords.map(s =>
      supabase.from("students").update({ status: "Transferred", arrears: s.outstandingDebt || 0 }).eq("id", s.id)
    ));

    setAllStudents(prev => ({ ...prev, [sid]: staying }));
    setAllAlumni(prev => ({ ...prev, [sid]: [...(prev[sid] || []), ...alumniRecords] }));
    setBulkAlumniExcelDone(true);
    notify(`✓ ${alumniRecords.length} student(s) moved to Alumni`);
  };

  // ── Alumni Debt Recovery ─────────────────────────────────────────
  const handleRecoverDebt = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    if (!showRecoverDebt) return;
    const amt = parseInt(recoverAmt.replace(/,/g, ""));
    if (!amt || amt <= 0) return notify("Enter a valid amount", "err");
    if (amt > showRecoverDebt.outstandingDebt) return notify(`Max amount is ${fmt(showRecoverDebt.outstandingDebt)}`, "err");
    const rcpt = nextRcpt();
    const today = new Date().toISOString().split("T")[0];
    const { data: inserted, error } = await supabase.from("payments").insert({
      school_id: activeSchoolId, student_id: showRecoverDebt.id,
      student_name: showRecoverDebt.name, term: `Recovery-${showRecoverDebt.leftYear}`,
      amount: amt, method: "Cash", received_by: adminCreds.username,
      payment_date: today, receipt_no: rcpt,
    }).select().single();
    if (error) return notify(`Could not record recovery payment: ${error.message}`, "err");
    const recoveryPay = { id: rcpt, dbId: inserted?.id, date: today, amount: amt, method: "Cash", receivedBy: adminCreds.username, term: `Recovery-${showRecoverDebt.leftYear}` };
    setAllAlumni(prev => ({
      ...prev,
      [activeSchoolId]: prev[activeSchoolId].map(al => al.id === showRecoverDebt.id
        ? { ...al, outstandingDebt: Math.max(0, al.outstandingDebt - amt), payments: [...(al.payments || []), recoveryPay] }
        : al)
    }));
    notify(`✓ ${fmt(amt)} recovery payment recorded for ${showRecoverDebt.name} — ${rcpt}`);
    setShowRecoverDebt(null);
    setRecoverAmt("");
  };

  // ── Derived Data ───────────────────────────────────────────────
  const subInfo = getSubscriptionInfo(school);
  const isReadOnly = subInfo.status === "Suspended";
  const getClassStreams = (cls) => (school?.streams && school.streams[cls]) || [];
  const students = allStudents[activeSchoolId] || [];
  const expenses = allExpenses[activeSchoolId] || [];
  const staff = allStaff[activeSchoolId] || [];
  const staffPayments = allStaffPayments[activeSchoolId] || [];

  const termStudents = students;
  const totalExpected = termStudents.reduce((s, st) => s + getBalance(st, currentTerm).totalDue, 0);
  const totalCollected = termStudents.reduce((s, st) => s + totalPaid(st, currentTerm), 0);
  const totalBalance = totalExpected - totalCollected;
  const collRate = totalExpected > 0 ? Math.round((totalCollected / totalExpected) * 100) : 0;
  const paidCount = termStudents.filter(s => getStatus(s, currentTerm) === "Paid").length;
  const partialCount = termStudents.filter(s => getStatus(s, currentTerm) === "Partial").length;
  const unpaidCount = termStudents.filter(s => getStatus(s, currentTerm) === "Unpaid").length;
  const totalExpensesAmt = expenses.filter(e => e.term === currentTerm).reduce((a, e) => a + e.amount, 0);
  const netSurplus = totalCollected - totalExpensesAmt;

  // Alumni debt — money owed by students who have already LEFT the school (graduated, transferred,
  // dropped out, etc.). Kept entirely separate from "Outstanding" above, which is only current,
  // actively-enrolled students. The two should never be added together: one is collectible through
  // normal term billing, the other requires separate debt-recovery follow-up with someone no longer
  // attending the school.
  const alumniDebtTotal = (allAlumni[activeSchoolId] || []).reduce((a, al) => a + Math.max(0, al.outstandingDebt || 0), 0);
  const alumniDebtorCount = (allAlumni[activeSchoolId] || []).filter(al => (al.outstandingDebt || 0) > 0).length;

  const classStats = schoolClasses.map(c => {
    const cls = termStudents.filter(s => s.class === c);
    const exp = cls.reduce((a, s) => a + getBalance(s, currentTerm).totalDue, 0);
    const col = cls.reduce((a, s) => a + totalPaid(s, currentTerm), 0);
    return { class: c, count: cls.length, expected: exp, collected: col, rate: exp > 0 ? Math.round((col / exp) * 100) : 0 };
  }).filter(g => g.count > 0);

  // Full payment ledger across active students AND alumni (so past-term/past-year records are never lost)
  const alumniForLedger = allAlumni[activeSchoolId] || [];
  const allStudentsAndAlumni = [...termStudents, ...alumniForLedger];

  // Every term that actually has at least one payment recorded, sorted most-recent-first
  const termsWithPayments = Array.from(new Set(
    allStudentsAndAlumni.flatMap(s => (s.payments || []).map(p => p.term))
  )).sort((a, b) => TERMS.indexOf(b) - TERMS.indexOf(a));

  const allPayments = allStudentsAndAlumni.flatMap(s =>
    (s.payments || [])
      .filter(p => paymentsTermFilter === "all" ? true : paymentsTermFilter === "current" ? p.term === currentTerm : p.term === paymentsTermFilter)
      .map(p => ({ ...p, studentName: s.name, studentClass: s.class, studentParent: s.parent, studentPhone: s.phone, isAlumni: !!s.status }))
  ).sort((a, b) => new Date(b.date) - new Date(a.date));

  const expenseCategories = ["Salaries & Wages", "Salaries", "Utilities", "Food", "Maintenance", "Stationery", "Other"];
  const expByCategory = expenseCategories.map(cat => ({
    name: cat,
    value: expenses.filter(e => e.category === cat && e.term === currentTerm).reduce((a, e) => a + e.amount, 0),
  })).filter(e => e.value > 0);
  const PIE_COLORS = ["#3b82f6", "#10b981", "#f59e0b", "#ef4444", "#8b5cf6", "#ec4899"];

  const filtered = termStudents.filter(s => {
    const st = getStatus(s, currentTerm);
    const matchesClass = filterClass === "All"
      || s.class === filterClass
      || classLabel(s.class, s.stream) === filterClass;
    return s.name.toLowerCase().includes(search.toLowerCase()) &&
      matchesClass &&
      (filterStatus === "All" || st === filterStatus);
  });

  // ── Handlers ──────────────────────────────────────────────────
  // ── Load real students + their payments from Supabase ──────────
  // Called right after a school admin successfully logs in. Fetches both
  // tables, then merges each student's payments into a `payments` array on
  // the student object — matching the exact in-app shape every existing
  // student function (add/edit/promote/etc.) already expects, so none of
  // that existing logic needs to change. Computes `arrears` fresh from
  // each student's actual saved value (now a real Supabase column) rather
  // than the old hardcoded demo arrears.
  const loadStudentsForSchool = async (schoolId) => {
    setStudentsLoading(true);
    // Use SECURITY DEFINER Postgres functions instead of direct table queries —
    // this works for both authenticated school admins AND unauthenticated parents
    // (who have no Supabase Auth session and would be blocked by RLS otherwise).
    const [studentsResult, paymentsResult, staffResult, staffPaymentsResult, expensesResult, configResult, smsResult] = await Promise.all([
      supabase.rpc("get_students_for_school", { p_school_id: schoolId }),
      supabase.rpc("get_payments_for_school", { p_school_id: schoolId }),
      supabase.rpc("get_staff_for_school", { p_school_id: schoolId }),
      supabase.rpc("get_staff_payments_for_school", { p_school_id: schoolId }),
      supabase.rpc("get_expenses_for_school", { p_school_id: schoolId }),
      supabase.rpc("get_school_config", { p_school_id: schoolId }),
      supabase.rpc("get_sms_log_for_school", { p_school_id: schoolId }),
    ]);
    if (studentsResult.error) {
      notify(`Could not load students: ${studentsResult.error.message}`, "err");
      setStudentsLoading(false);
      return;
    }
    if (paymentsResult.error) {
      notify(`Could not load payment history: ${paymentsResult.error.message}`, "err");
    }
    const paymentsByStudent = {};
    (paymentsResult.data || []).filter(Boolean).forEach(p => {
      if (!paymentsByStudent[p.student_id]) paymentsByStudent[p.student_id] = [];
      paymentsByStudent[p.student_id].push({
        id: p.receipt_no, dbId: p.id, date: p.payment_date, amount: p.amount,
        method: p.method, receivedBy: p.received_by || "", term: p.term,
      });
    });
    // Advance the in-memory receipt counter past the highest real receipt
    // number already saved for this school, so it can't generate a number
    // that collides with one that already exists in the database — the
    // counter itself has no memory between page loads, only the database does.
    (paymentsResult.data || []).filter(Boolean).forEach(p => {
      const match = /^RCP-(\d+)$/.exec(p.receipt_no || "");
      if (match) {
        const n = parseInt(match[1], 10);
        if (n >= rcptN) rcptN = n + 1;
      }
    });
    const loadedStudents = [];
    const loadedAlumni = [];
    (studentsResult.data || []).filter(Boolean).forEach(row => {
      if (!row || !row.id) return; // skip any malformed rows entirely
      const base = {
        id: row.id, schoolId: row.school_id || "",
        name: row.name || "", class: row.class || "", stream: row.stream || "", gender: row.gender || "",
        category: row.category || "Day Scholar", parent: row.parent_name || "", phone: row.phone || "",
        arrears: row.arrears || 0,
        bursary: row.bursary_type ? { type: row.bursary_type, value: row.bursary_value || 0, reason: row.bursary_reason || "" } : null,
        customFee: row.custom_fee || null,
        photo: row.photo_url || null,
        payments: paymentsByStudent[row.id] || [],
        status: row.status || "active",
      };
      if (!row.status || row.status === "active") {
        loadedStudents.push(base);
      } else {
        loadedAlumni.push({
          ...base,
          status: row.status,
          leftClass: row.left_class || "",
          leftYear: row.left_year || "",
          outstandingDebt: row.outstanding_debt || 0,
          leftNote: row.left_note || "",
        });
      }
    });
    setAllStudents(prev => ({ ...prev, [schoolId]: loadedStudents }));
    setAllAlumni(prev => ({ ...prev, [schoolId]: loadedAlumni }));

    // Load staff
    const loadedStaff = (staffResult.data || []).filter(Boolean).map(row => ({
      id: row.id, schoolId: row.school_id,
      name: row.name, role: row.role, phone: row.phone || "",
      defaultRate: row.default_rate || 0,
      defaultRateType: row.default_rate_type || "daily",
      active: row.active !== false, photo: row.photo_url || null,
    }));
    setAllStaff(prev => ({ ...prev, [schoolId]: loadedStaff }));

    // Load staff payments
    const loadedStaffPayments = (staffPaymentsResult.data || []).filter(Boolean).map(row => ({
      id: row.id, schoolId: row.school_id, staffId: row.staff_id, staffName: row.staff_name,
      amount: row.amount, payType: row.pay_type, periodLabel: row.period_label,
      date: row.payment_date, term: row.term, paidBy: row.paid_by || "",
    }));
    setAllStaffPayments(prev => ({ ...prev, [schoolId]: loadedStaffPayments }));

    // Load expenses
    const loadedExpenses = (expensesResult.data || []).filter(Boolean).map(row => ({
      id: row.id, schoolId: row.school_id, category: row.category,
      description: row.description, amount: row.amount,
      date: row.date, term: row.term, paidBy: row.paid_by || "",
      staffPaymentId: row.staff_payment_id || null,
    }));
    setAllExpenses(prev => ({ ...prev, [schoolId]: loadedExpenses }));

    // Load fee structure and requirements from school_config
    if (configResult.data && configResult.data.length > 0) {
      const config = configResult.data[0];
      if (config.fee_structure && Object.keys(config.fee_structure).length > 0) {
        setFeeStructure(config.fee_structure);
      }
      if (config.requirements && config.requirements.length > 0) {
        setRequirements(config.requirements);
      }
    }

    // Load SMS log
    const loadedSmsLog = (smsResult.data || []).filter(Boolean).map(row => ({
      id: row.id, to: row.phone, student: row.student_name,
      message: row.message, status: row.status,
      time: new Date(row.sent_at).toLocaleTimeString(),
    }));
    setSmsLog(loadedSmsLog);

    setStudentsLoading(false);
  };

  // ── Request a password reset email ───────────────────────────────
  // Calls the request-password-reset Edge Function, which looks up the
  // school by email, generates a token, and sends the reset link via
  // Resend. Always shows the same success message regardless of whether
  // the email actually matched a school — the Edge Function deliberately
  // doesn't reveal that, to avoid letting this form be used to check
  // which emails are registered.
  const handleForgotPassword = async () => {
    if (!forgotPasswordEmail.trim()) return;
    setForgotPasswordStatus("sending");
    try {
      const { error } = await supabase.functions.invoke("request-password-reset", {
        body: { email: forgotPasswordEmail.trim() },
      });
      if (error) {
        setForgotPasswordStatus("Something went wrong. Please try again.");
        return;
      }
      setForgotPasswordStatus("sent");
    } catch {
      setForgotPasswordStatus("Something went wrong. Please try again.");
    }
  };

  // ── Complete a password reset using the token from the email link ──
  const handleCompletePasswordReset = async () => {
    if (resetNewPassword.length < 6) return setResetStatus("Password must be at least 6 characters");
    if (resetNewPassword !== resetConfirmPassword) return setResetStatus("Passwords do not match");
    setResetStatus("submitting");
    try {
      const { data, error } = await supabase.functions.invoke("complete-password-reset", {
        body: { token: resetPasswordToken, newPassword: resetNewPassword },
      });
      if (error) {
        // Edge Function errors (4xx/5xx) surface here without the response
        // body by default — fall back to a generic message in that case.
        setResetStatus(data?.error || "Could not reset password. The link may be invalid or expired.");
        return;
      }
      setResetStatus("success");
    } catch {
      setResetStatus("Something went wrong. Please try again.");
    }
  };

  const handleLogin = async () => {
    if (loginScreen === "admin") {
      // Super admin login
      if (loginInput.user === superAdminCreds.username && loginInput.pass === superAdminCreds.password) {
        setIsSuperAdmin(true);
        setCurrentUser({ role: "superadmin" }); setLoginError("");
        saveSession({ role: "superadmin" });
        return;
      }
      // ── Supabase Auth login ──────────────────────────────────────
      // Step 1: look up the school's email from their username via
      // a safe server-side Postgres function (username is typed by
      // the user; the function returns only the matching email).
      const { data: emailData, error: emailLookupError } = await supabase
        .rpc("get_school_email_by_username", { p_username: loginInput.user });
      if (emailLookupError || !emailData) {
        setLoginError("Invalid credentials. If your school just signed up, wait for approval.");
        return;
      }
      // Step 2: sign in with Supabase Auth using the real email + password.
      const { data: authData, error: authError } = await supabase.auth.signInWithPassword({
        email: emailData,
        password: loginInput.pass,
      });
      if (authError || !authData?.user) {
        setLoginError("Invalid credentials. If your school just signed up, wait for approval.");
        return;
      }
      // Step 3: fetch this school's row directly — now that we have a valid
      // Supabase Auth session, RLS will return exactly this school's row.
      // We can't rely on SCHOOLS_DATA here since it was empty on page load
      // (no auth session existed yet when loadSchools ran).
      const { data: schoolRows, error: schoolFetchError } = await supabase
        .from("schools")
        .select("*")
        .eq("user_id", authData.user.id)
        .single();
      if (schoolFetchError || !schoolRows) {
        setLoginError("School account not found. Please contact support.");
        return;
      }
      const fetchedSchool = schoolRows;
      const newSchoolId = fetchedSchool.id;
      // Populate SCHOOLS_DATA with this school so the rest of the app works
      SCHOOLS_DATA[newSchoolId] = {
        id: newSchoolId,
        userId: fetchedSchool.user_id || null,
        name: fetchedSchool.name,
        location: fetchedSchool.location,
        principal: fetchedSchool.principal,
        phone: fetchedSchool.phone,
        notifyEmail: fetchedSchool.notify_email || "",
        logo: fetchedSchool.logo || "🏫",
        schoolType: fetchedSchool.school_type || "secondary",
        streams: fetchedSchool.streams || {},
        setupComplete: fetchedSchool.setup_complete,
        adminUsername: fetchedSchool.admin_username,
        adminPassword: fetchedSchool.admin_password,
        plan: fetchedSchool.plan || "Starter",
        billingCycle: fetchedSchool.billing_cycle || "monthly",
        customPrice: fetchedSchool.custom_price,
        customPriceNote: fetchedSchool.custom_price_note || "",
        subscriptionStatus: fetchedSchool.subscription_status || "Active",
        isTrial: fetchedSchool.is_trial,
        trialActivated: fetchedSchool.trial_activated,
        trialStartDate: fetchedSchool.trial_start_date,
        nextBillingDate: fetchedSchool.next_billing_date,
        lastPaymentDate: fetchedSchool.last_payment_date,
        paymentNoticeFreeze: fetchedSchool.payment_notice_freeze || false,
        billingRef: fetchedSchool.billing_ref,
      };
      // Start the 30-day free trial on first login
      if (fetchedSchool.is_trial && !fetchedSchool.trial_activated) {
        const trialEnd = new Date();
        trialEnd.setDate(trialEnd.getDate() + 30);
        SCHOOLS_DATA[newSchoolId].trialActivated = true;
        SCHOOLS_DATA[newSchoolId].trialStartDate = new Date().toISOString().split("T")[0];
        SCHOOLS_DATA[newSchoolId].nextBillingDate = trialEnd.toISOString().split("T")[0];
        await supabase.from("schools").update({
          trial_activated: true,
          trial_start_date: new Date().toISOString().split("T")[0],
          next_billing_date: trialEnd.toISOString().split("T")[0],
        }).eq("id", newSchoolId);
      }
      setActiveSchoolId(newSchoolId);
      setCurrentUser({ role: "admin" }); setLoginError("");
      saveSession({ role: "admin", schoolId: newSchoolId });
      loadStudentsForSchool(newSchoolId);
      setSubscriptionRefresh(r => r + 1);
      return;
      // ── End Supabase Auth login ──────────────────────────────────
    } else {
      // ── Parent login: phone number as username, last 4 digits as PIN ──
      // A parent logging in fresh has no school selected yet — unlike an
      // admin, who picks their school implicitly via their own username/
      // password. So rather than searching only the currently-loaded
      // school's students (which would be empty for a parent who hasn't
      // selected anything), this searches the real students table directly
      // by phone number, across all schools — a phone number naturally
      // identifies one family, which in practice means one school.
      const cleanPhone = loginInput.user.replace(/-/g, "");
      const { data: matches, error } = await supabase.rpc("get_students_by_phone", { p_phone: cleanPhone });
      if (error) {
        setLoginError("Could not check phone number — please try again.");
        return;
      }
      const myChildrenRows = (matches || []).filter(row => (row.phone || "").replace(/-/g, "") === cleanPhone);
      const firstRow = myChildrenRows[0];
      if (firstRow && loginInput.pass === (firstRow.phone || "").slice(-4)) {
        const schoolId = firstRow.school_id;
        // Load this school's real data so the parent's dashboard (which
        // reads from the same allStudents/allAlumni state as admins) has
        // something to show, exactly like an admin login already does.
        await loadStudentsForSchool(schoolId);
        const childIds = myChildrenRows.filter(r => r.school_id === schoolId).map(r => r.id);
        const session = { role: "parent", schoolId, studentId: firstRow.id, childIds };
        setActiveSchoolId(schoolId);
        setCurrentUser(session);
        setLoginError("");
        saveSession(session);
      } else {
        setLoginError("Phone not found or wrong PIN. PIN = last 4 digits of phone.");
      }
    }
  };

  // ── Shared logout — clears both in-memory state and the saved session ──
  const logout = () => {
    supabase.auth.signOut();
    setCurrentUser(null);
    setIsSuperAdmin(false);
    setActiveSchoolId(null);
    setTab("dashboard");
    setLoginInput({ user: "", pass: "" });
    clearSession();
  };

  const handlePay = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const amount = parseInt(payAmt.replace(/,/g, ""));
    if (!amount || amount <= 0) return notify("Enter valid amount", "err");
    const bal = getBalance(showPay, currentTerm).balance;
    if (amount > bal) return notify(`Max payable is ${fmt(bal)}`, "err");
    const rcpt = nextRcpt();

    const { error: insertError } = await supabase.from("payments").insert({
      receipt_no: rcpt, school_id: activeSchoolId, student_id: showPay.id, student_name: showPay.name,
      term: currentTerm, amount, method: payMethod, received_by: receivedBy, payment_date: payDate,
    });
    if (insertError) {
      return notify(`Could not record payment: ${insertError.message}`, "err");
    }

    const newPay = { id: rcpt, date: payDate, amount, method: payMethod, receivedBy, term: currentTerm };
    const newBal = bal - amount;
    setAllStudents(prev => ({
      ...prev,
      [activeSchoolId]: prev[activeSchoolId].map(s => s.id === showPay.id
        ? { ...s, payments: [...s.payments, newPay] } : s)
    }));
    const updatedStudent = { ...showPay, payments: [...showPay.payments, newPay] };
    setShowReceipt({ payment: newPay, student: updatedStudent, school, newBalance: newBal });
    notify(`${fmt(amount)} recorded — ${rcpt}`);
    setShowPay(null); setPayAmt(""); setPayMethod("Cash");
  };

  const handleSaveEditStudent = async () => {
    if (!showEditStudent) return;
    const { id, name, class: cls, stream, gender, category, parent, phone } = showEditStudent;
    if (!name.trim()) return notify("Student name is required", "err");
    const { error } = await supabase.from("students").update({
      name: name.trim(), class: cls, stream: stream || "",
      gender, category, parent_name: parent || "", phone: phone || "",
    }).eq("id", id);
    if (error) return notify(`Could not update student: ${error.message}`, "err");
    setAllStudents(prev => ({
      ...prev,
      [activeSchoolId]: prev[activeSchoolId].map(s => s.id === id ? { ...s, name: name.trim(), class: cls, stream: stream || "", gender, category, parent: parent || "", phone: phone || "" } : s),
    }));
    notify(`${name} updated ✓`);
    setShowEditStudent(null);
  };

  const handleAddStudent = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    if (!newS.name.trim() || !newS.parent.trim()) return notify("Fill required fields", "err");

    // ── Returning student: confirmed match with an alumni record ──
    if (confirmReturning && returningMatch) {
      const { error: updateError } = await supabase.from("students").update({
        status: "active",
        name: newS.name, class: newS.class, stream: newS.stream || "", gender: newS.gender,
        category: newS.category, parent_name: newS.parent, phone: newS.phone,
        arrears: returningMatch.outstandingDebt || 0, // old debt carries forward as arrears
        custom_fee: newS.customFee ? parseInt(newS.customFee) : (returningMatch.customFee || null),
        left_class: null, left_year: null, outstanding_debt: null, left_note: null,
      }).eq("id", returningMatch.id);

      if (updateError) {
        return notify(`Could not re-enrol returning student: ${updateError.message}`, "err");
      }

      const s = {
        ...returningMatch,
        schoolId: activeSchoolId,
        name: newS.name, class: newS.class, stream: newS.stream || "", gender: newS.gender,
        category: newS.category, parent: newS.parent, phone: newS.phone,
        arrears: returningMatch.outstandingDebt || 0, // old debt carries forward as arrears
        bursary: returningMatch.bursary || null,
        customFee: newS.customFee ? parseInt(newS.customFee) : (returningMatch.customFee || null),
        payments: returningMatch.payments || [], // full payment history preserved
      };
      // remove status/leftClass/leftYear/outstandingDebt/leftNote fields from the alumni record
      delete s.status; delete s.leftClass; delete s.leftYear; delete s.outstandingDebt; delete s.leftNote;

      setAllStudents(prev => ({ ...prev, [activeSchoolId]: [...(prev[activeSchoolId] || []), s] }));
      setAllAlumni(prev => ({ ...prev, [activeSchoolId]: prev[activeSchoolId].filter(a => a.id !== returningMatch.id) }));
      checkAutoUpgrade(activeSchoolId, students.length + 1);
      notify(`✓ ${newS.name} re-enrolled as returning student! Old records restored — arrears: ${fmt(s.arrears)}`);
      setShowAdd(false);
      setNewS({ name: "", class: schoolClasses[0], stream: "", gender: "M", category: "Day Scholar", parent: "", phone: "", bursary: null, customFee: "" });
      setReturningMatch(null);
      setConfirmReturning(false);
      return;
    }

    // ── Normal new enrolment ──
    const { data: insertedRow, error: insertError } = await supabase.from("students").insert({
      school_id: activeSchoolId,
      name: newS.name, class: newS.class, stream: newS.stream || "", gender: newS.gender,
      category: newS.category, parent_name: newS.parent, phone: newS.phone,
      arrears: 0, custom_fee: newS.customFee ? parseInt(newS.customFee) : null,
    }).select().single();

    if (insertError) {
      return notify(`Could not enroll student: ${insertError.message}`, "err");
    }

    const s = {
      id: insertedRow.id, schoolId: activeSchoolId,
      name: newS.name, class: newS.class, stream: newS.stream || "", gender: newS.gender,
      category: newS.category, parent: newS.parent, phone: newS.phone,
      arrears: 0, payments: [],
      bursary: null,
      customFee: newS.customFee ? parseInt(newS.customFee) : null,
    };
    setAllStudents(prev => ({ ...prev, [activeSchoolId]: [...(prev[activeSchoolId] || []), s] }));
    checkAutoUpgrade(activeSchoolId, students.length + 1);
    notify(`${newS.name} enrolled — ${fmt(getStudentFee(s))}/term`);
    setShowAdd(false);
    setNewS({ name: "", class: schoolClasses[0], stream: "", gender: "M", category: "Day Scholar", parent: "", phone: "", bursary: null, customFee: "" });
    setReturningMatch(null);
    setConfirmReturning(false);
  };

  // Check alumni for a matching name as the bursar types
  const checkReturningStudent = (name) => {
    if (!name || name.trim().length < 3) { setReturningMatch(null); setConfirmReturning(false); return; }
    const alumni = allAlumni[activeSchoolId] || [];
    const match = alumni.find(a => a.name.toLowerCase().trim() === name.toLowerCase().trim());
    setReturningMatch(match || null);
    if (!match) setConfirmReturning(false);
  };

  // ── First-time setup wizard ─────────────────────────────────
  const markSetupComplete = async () => {
    SCHOOLS_DATA[activeSchoolId].setupComplete = true;
    setSetupDismissed(true);
    notify("✓ Setup complete! Welcome to FeeTrack UG");
    const { error } = await supabase.from("schools").update({ setup_complete: true }).eq("id", activeSchoolId);
    if (error) console.error("Failed to save setup-complete status:", error.message);
    // Deliberately not blocking on this or showing an error toast for a
    // failure here — the in-memory flag already dismissed the wizard for
    // this session, which is the part that actually matters to the
    // person right now. Worst case if this write fails: the wizard
    // reappears once on their next reload, which is a minor inconvenience
    // rather than something that should interrupt their flow.
  };

  // ── Class Streams Configuration ──────────────────────────────
  const addStream = (cls) => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const name = (newStreamInput[cls] || "").trim();
    if (!name) return;
    const current = (streamsForm ?? school.streams ?? {});
    const existing = current[cls] || [];
    if (existing.some(s => s.toLowerCase() === name.toLowerCase())) return notify(`"${name}" already exists for ${cls}`, "err");
    const updated = { ...current, [cls]: [...existing, name] };
    setStreamsForm(updated);
    setNewStreamInput(prev => ({ ...prev, [cls]: "" }));
  };

  const removeStream = (cls, streamName) => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const current = (streamsForm ?? school.streams ?? {});
    const updated = { ...current, [cls]: (current[cls] || []).filter(s => s !== streamName) };
    setStreamsForm(updated);
  };

  const saveStreams = async () => {
    if (!streamsForm) return notify("No changes to save");
    const { data, error } = await supabase.from("schools").update({ streams: streamsForm }).eq("id", activeSchoolId).select();
    if (error) {
      console.error("Failed to save class streams:", error.message);
      return notify("Could not save changes — please try again", "err");
    }
    if (!data || data.length === 0) {
      console.error("saveStreams — update matched zero rows for activeSchoolId:", activeSchoolId);
      return notify("Could not save changes — school not found", "err");
    }
    SCHOOLS_DATA[activeSchoolId].streams = streamsForm;
    setSubscriptionRefresh(r => r + 1);
    setStreamsForm(null);
    notify("✓ Class streams updated");
  };

  // ── Bulk Student Import (Excel/CSV) ──────────────────────────
  const handleBulkFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setBulkFileName(file.name);
    setBulkImportDone(false);

    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });

        const parsed = raw.map((row, i) => {
          // Flexible header matching — accept common variations
          const get = (...keys) => {
            for (const k of keys) {
              const found = Object.keys(row).find(h => h.toLowerCase().trim() === k.toLowerCase());
              if (found && String(row[found]).trim() !== "") return String(row[found]).trim();
            }
            return "";
          };
          const name = get("name", "student name", "full name");
          const cls = get("class", "form");
          const gender = get("gender", "sex");
          const category = get("category", "type");
          const parent = get("parent", "guardian", "parent name", "parent/guardian");
          const phone = get("phone", "parent phone", "contact", "phone number");
          const arrears = get("arrears", "balance brought forward", "opening balance");

          // ── Fee arrangement: Full Fee / Bursary (%) / Bursary (Fixed) / Custom Fee ──
          const feeType = get("fee type", "feetype", "arrangement", "payment type").toLowerCase();
          const bursaryPercent = get("bursary %", "bursary percent", "discount %", "scholarship %");
          const bursaryFixed = get("bursary amount", "discount amount", "bursary fixed");
          const bursaryReason = get("bursary reason", "scholarship reason", "reason");
          const customFeeAmt = get("custom fee", "custom amount", "agreed fee");

          let bursary = null;
          let customFee = null;
          if (feeType.includes("custom") || customFeeAmt) {
            customFee = parseInt(customFeeAmt) || null;
          } else if (feeType.includes("bursary") || feeType.includes("scholarship") || bursaryPercent || bursaryFixed) {
            if (bursaryPercent) {
              bursary = { type: "percent", value: parseInt(bursaryPercent) || 0, reason: bursaryReason || "Bursary" };
            } else if (bursaryFixed) {
              bursary = { type: "fixed", value: parseInt(bursaryFixed) || 0, reason: bursaryReason || "Bursary" };
            }
          }

          return {
            rowIndex: i,
            name,
            class: schoolClasses.map(c=>c.toUpperCase()).includes(cls.toUpperCase()) ? schoolClasses.find(c=>c.toUpperCase()===cls.toUpperCase()) : schoolClasses[0],
            gender: gender.toUpperCase().startsWith("F") ? "F" : "M",
            category: category.toLowerCase().includes("board") ? "Boarder" : "Day Scholar",
            parent: parent || "Unknown",
            phone: phone.replace(/\s+/g, ""),
            arrears: parseInt(arrears) || 0,
            bursary, customFee,
            valid: !!name,
          };
        }).filter(r => r.name);

        setBulkRows(parsed);
        notify(`✓ ${parsed.length} students found in ${file.name}`);
      } catch (err) {
        notify("Could not read file. Make sure it's a valid Excel (.xlsx) or CSV file.", "err");
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleBulkImport = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const validRows = bulkRows.filter(r => r.valid);
    if (validRows.length === 0) return notify("No valid rows to import", "err");

    const { data: insertedRows, error: insertError } = await supabase.from("students").insert(
      validRows.map(r => ({
        school_id: activeSchoolId,
        name: r.name, class: r.class, gender: r.gender, category: r.category,
        parent_name: r.parent, phone: r.phone,
        arrears: r.arrears || 0,
        bursary_type: r.bursary?.type || null, bursary_value: r.bursary?.value || null, bursary_reason: r.bursary?.reason || null,
        custom_fee: r.customFee || null,
      }))
    ).select();

    if (insertError) {
      return notify(`Could not import students: ${insertError.message}`, "err");
    }

    const newStudents = (insertedRows || []).map(row => ({
      id: row.id, schoolId: row.school_id,
      name: row.name, class: row.class, gender: row.gender, category: row.category || "Day Scholar",
      parent: row.parent_name || "", phone: row.phone || "",
      arrears: row.arrears || 0,
      bursary: row.bursary_type ? { type: row.bursary_type, value: row.bursary_value, reason: row.bursary_reason || "" } : null,
      customFee: row.custom_fee, payments: [],
    }));
    setAllStudents(prev => ({ ...prev, [activeSchoolId]: [...(prev[activeSchoolId] || []), ...newStudents] }));
    checkAutoUpgrade(activeSchoolId, students.length + newStudents.length);
    setBulkImportDone(true);
    notify(`✓ ${newStudents.length} students imported successfully!`);
  };

  const downloadBulkTemplate = () => {
    const wb = XLSX.utils.book_new();
    const sample = [
      { Name: "Nakato Sarah", Class: "S1", Gender: "F", Category: "Day Scholar", Parent: "Nakato Mary", Phone: "0772441823", Arrears: 0, "Fee Type": "Full Fee", "Bursary %": "", "Bursary Amount": "", "Bursary Reason": "", "Custom Fee": "" },
      { Name: "Okello James", Class: "S2", Gender: "M", Category: "Boarder", Parent: "Okello Peter", Phone: "0701234567", Arrears: 50000, "Fee Type": "Full Fee", "Bursary %": "", "Bursary Amount": "", "Bursary Reason": "", "Custom Fee": "" },
      { Name: "Namukasa Grace", Class: "S1", Gender: "F", Category: "Day Scholar", Parent: "Namukasa Ruth", Phone: "0752345678", Arrears: 0, "Fee Type": "Bursary", "Bursary %": 50, "Bursary Amount": "", "Bursary Reason": "Orphan Bursary", "Custom Fee": "" },
      { Name: "Tumwine Alex", Class: "S5", Gender: "M", Category: "Day Scholar", Parent: "Tumwine Robert", Phone: "0702234567", Arrears: 0, "Fee Type": "Custom Fee", "Bursary %": "", "Bursary Amount": "", "Bursary Reason": "", "Custom Fee": 600000 },
    ];
    const ws = XLSX.utils.json_to_sheet(sample);
    XLSX.utils.book_append_sheet(wb, ws, "Students");

    // Add an instructions sheet
    const instructions = [
      { Column: "Name", Instructions: "Student's full name (required)" },
      { Column: "Class", Instructions: "S1, S2, S3, S4, S5, or S6" },
      { Column: "Gender", Instructions: "M or F" },
      { Column: "Category", Instructions: "Day Scholar or Boarder" },
      { Column: "Parent", Instructions: "Parent/guardian full name" },
      { Column: "Phone", Instructions: "Parent's phone number (used for Parent Portal login)" },
      { Column: "Arrears", Instructions: "Any balance brought forward from a previous term (0 if none)" },
      { Column: "Fee Type", Instructions: "Full Fee = pays standard fee · Bursary = gets a discount · Custom Fee = pays a specific agreed amount" },
      { Column: "Bursary %", Instructions: "If Fee Type is Bursary and discount is a percentage, e.g. 50 for 50% off" },
      { Column: "Bursary Amount", Instructions: "If Fee Type is Bursary and discount is a fixed UGX amount instead of %" },
      { Column: "Bursary Reason", Instructions: "e.g. Orphan Bursary, Staff Child Discount, Academic Scholarship" },
      { Column: "Custom Fee", Instructions: "If Fee Type is Custom Fee, the exact UGX amount this student pays per term" },
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(instructions), "Instructions");

    XLSX.writeFile(wb, "FeeTrack_Student_Import_Template.xlsx");
    notify("Template downloaded ✓");
  };

  // ── Bulk Payment Import (Excel/CSV) ──────────────────────────
  const handleBulkPayFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setBulkPayFileName(file.name);
    setBulkPayImportDone(false);

    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });

        const parsed = raw.map((row, i) => {
          const get = (...keys) => {
            for (const k of keys) {
              const found = Object.keys(row).find(h => h.toLowerCase().trim() === k.toLowerCase());
              if (found && String(row[found]).trim() !== "") return String(row[found]).trim();
            }
            return "";
          };
          const name = get("name", "student name", "full name");
          const cls = get("class", "form");
          const amount = get("amount", "amount paid", "paid");
          const method = get("method", "payment method");
          const date = get("date", "payment date");
          const receivedBy = get("received by", "staff", "receivedby");

          // Find matching student in current roster — match by name (and class/stream if provided, for disambiguation)
          const candidates = termStudents.filter(s => s.name.toLowerCase().trim() === name.toLowerCase().trim());
          let matchedStudent = null;
          let matchStatus = "not_found";
          if (candidates.length === 1) {
            matchedStudent = candidates[0];
            matchStatus = "matched";
          } else if (candidates.length > 1) {
            if (cls) {
              // Try exact class+stream label first (e.g. "S1 - East"), then plain class as fallback
              const exact = candidates.find(s => classLabel(s.class, s.stream).toUpperCase() === cls.toUpperCase())
                || candidates.find(s => s.class.toUpperCase() === cls.toUpperCase());
              if (exact) { matchedStudent = exact; matchStatus = "matched"; }
              else matchStatus = "ambiguous";
            } else {
              matchStatus = "ambiguous";
            }
          }

          const parsedAmount = parseInt(String(amount).replace(/,/g, "")) || 0;
          let balance = null, overpay = false;
          if (matchedStudent) {
            balance = getBalance(matchedStudent, currentTerm).balance;
            overpay = parsedAmount > balance;
          }

          const methodMatch = METHODS.find(m => m.toLowerCase() === method.toLowerCase());

          return {
            rowIndex: i,
            name, class: cls, amount: parsedAmount,
            method: methodMatch || "Cash",
            date: date || new Date().toISOString().split("T")[0],
            receivedBy: receivedBy || adminCreds.username,
            matchedStudent, matchStatus, balance, overpay,
            valid: matchStatus === "matched" && parsedAmount > 0,
          };
        }).filter(r => r.name);

        setBulkPayRows(parsed);
        const matched = parsed.filter(r => r.matchStatus === "matched").length;
        notify(`✓ ${parsed.length} rows found — ${matched} matched to students`);
      } catch (err) {
        notify("Could not read file. Make sure it's a valid Excel (.xlsx) or CSV file.", "err");
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleBulkPayImport = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const validRows = bulkPayRows.filter(r => r.valid);
    if (validRows.length === 0) return notify("No valid payments to import", "err");

    const inserts = validRows.map(r => ({
      school_id: activeSchoolId, student_id: r.matchedStudent.id,
      student_name: r.matchedStudent.name, term: currentTerm,
      amount: r.amount, method: r.method || "Cash",
      received_by: r.receivedBy || adminCreds.username,
      payment_date: r.date || new Date().toISOString().split("T")[0],
      receipt_no: nextRcpt(),
    }));

    const { data: inserted, error } = await supabase.from("payments").insert(inserts).select();
    if (error) return notify(`Could not record payments: ${error.message}`, "err");

    setAllStudents(prev => ({
      ...prev,
      [activeSchoolId]: prev[activeSchoolId].map(s => {
        const rowsForStudent = (inserted || []).filter(p => p.student_id === s.id);
        if (rowsForStudent.length === 0) return s;
        const newPayments = rowsForStudent.map(p => ({
          id: p.receipt_no, dbId: p.id, date: p.payment_date,
          amount: p.amount, method: p.method, receivedBy: p.received_by, term: p.term,
        }));
        return { ...s, payments: [...s.payments, ...newPayments] };
      })
    }));

    setBulkPayImportDone(true);
    const total = validRows.reduce((a, r) => a + r.amount, 0);
    notify(`✓ ${validRows.length} payments recorded — total ${fmt(total)}`);
  };

  const downloadBulkPayTemplate = () => {
    const wb = XLSX.utils.book_new();
    const sample = termStudents.slice(0, 5).map(s => ({
      Name: s.name, Class: s.class, Amount: "", Method: "Cash", Date: new Date().toISOString().split("T")[0], "Received By": adminCreds.username,
    }));
    const ws = XLSX.utils.json_to_sheet(sample.length > 0 ? sample : [
      { Name: "Nakato Sarah", Class: "S1", Amount: 200000, Method: "Cash", Date: new Date().toISOString().split("T")[0], "Received By": "Mugisha R." },
    ]);
    XLSX.utils.book_append_sheet(wb, ws, "Payments");

    const instructions = [
      { Column: "Name", Instructions: "Must match the student's name exactly as it appears in Student Register" },
      { Column: "Class", Instructions: "Recommended if two students share the same name — helps match the right one" },
      { Column: "Amount", Instructions: "Amount paid in UGX (required)" },
      { Column: "Method", Instructions: `One of: ${METHODS.join(", ")} (defaults to Cash if blank or unrecognized)` },
      { Column: "Date", Instructions: "Date of payment, format YYYY-MM-DD (defaults to today if blank)" },
      { Column: "Received By", Instructions: "Staff member who received the payment (optional)" },
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(instructions), "Instructions");

    XLSX.writeFile(wb, `FeeTrack_Bulk_Payments_${currentTerm.replace(/[, ]/g, "_")}.xlsx`);
    notify("Template downloaded ✓ (pre-filled with your current student list)");
  };

  // ── Bulk Pay Staff (Excel) ──────────────────────────────────────
  const downloadBulkStaffPayTemplate = () => {
    const wb = XLSX.utils.book_new();
    const activeStaffList = staff.filter(s => s.active);
    const sample = activeStaffList.slice(0, 10).map(s => ({
      Name: s.name, Role: s.role, Amount: "", Type: s.defaultRateType === "monthly" ? "Monthly" : "Daily",
      Period: "", Date: new Date().toISOString().split("T")[0],
    }));
    const ws = XLSX.utils.json_to_sheet(sample.length > 0 ? sample : [
      { Name: "Mugisha Robert", Role: "Security Guard", Amount: 15000, Type: "Daily", Period: "16 Jun 2026", Date: new Date().toISOString().split("T")[0] },
    ]);
    XLSX.utils.book_append_sheet(wb, ws, "Staff Payments");

    const instructions = [
      { Column: "Name", Instructions: "Must match the worker's name exactly as it appears in Staff & Wages" },
      { Column: "Role", Instructions: "Recommended if two workers share the same name — helps match the right one" },
      { Column: "Amount", Instructions: "Amount paid in UGX (required)" },
      { Column: "Type", Instructions: "Either \"Daily\" or \"Monthly\" (defaults to Daily if blank or unrecognized)" },
      { Column: "Period", Instructions: "What this payment covers, e.g. \"16 Jun 2026\" for a day or \"June 2026\" for a month (required)" },
      { Column: "Date", Instructions: "Date the payment was made, format YYYY-MM-DD (defaults to today if blank)" },
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(instructions), "Instructions");

    XLSX.writeFile(wb, `FeeTrack_Bulk_Staff_Payments_${currentTerm.replace(/[, ]/g, "_")}.xlsx`);
    notify("Template downloaded ✓ (pre-filled with your active staff list)");
  };

  const handleBulkStaffPayFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setBulkStaffPayFileName(file.name);
    setBulkStaffPayImportDone(false);

    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });

        const parsed = raw.map((row, i) => {
          const get = (...keys) => {
            for (const k of keys) {
              const found = Object.keys(row).find(h => h.toLowerCase().trim() === k.toLowerCase());
              if (found && String(row[found]).trim() !== "") return String(row[found]).trim();
            }
            return "";
          };
          const name = get("name", "worker name", "full name");
          const role = get("role");
          const amount = get("amount", "amount paid", "paid");
          const typeRaw = get("type", "pay type");
          const period = get("period", "period label");
          const date = get("date", "payment date");

          // Find matching worker — match by name (and role, if provided, for disambiguation)
          const candidates = staff.filter(s => s.active && s.name.toLowerCase().trim() === name.toLowerCase().trim());
          let matchedStaff = null;
          let matchStatus = "not_found";
          if (candidates.length === 1) {
            matchedStaff = candidates[0];
            matchStatus = "matched";
          } else if (candidates.length > 1) {
            if (role) {
              const exact = candidates.find(s => s.role.toLowerCase() === role.toLowerCase());
              if (exact) { matchedStaff = exact; matchStatus = "matched"; }
              else matchStatus = "ambiguous";
            } else {
              matchStatus = "ambiguous";
            }
          }

          const parsedAmount = parseInt(String(amount).replace(/,/g, "")) || 0;
          const payType = typeRaw.toLowerCase() === "monthly" ? "monthly" : "daily";

          return {
            rowIndex: i,
            name, role, amount: parsedAmount, payType,
            periodLabel: period,
            date: date || new Date().toISOString().split("T")[0],
            matchedStaff, matchStatus,
            valid: matchStatus === "matched" && parsedAmount > 0 && !!period,
          };
        }).filter(r => r.name);

        setBulkStaffPayRows(parsed);
        const matched = parsed.filter(r => r.matchStatus === "matched").length;
        notify(`✓ ${parsed.length} rows found — ${matched} matched to staff`);
      } catch (err) {
        notify("Could not read file. Make sure it's a valid Excel (.xlsx) or CSV file.", "err");
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleBulkStaffPayImport = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const validRows = bulkStaffPayRows.filter(r => r.valid);
    if (validRows.length === 0) return notify("No valid payments to import", "err");

    const staffPayInserts = validRows.map(r => ({
      school_id: activeSchoolId, staff_id: r.matchedStaff.id, staff_name: r.matchedStaff.name,
      amount: r.amount, pay_type: r.payType, period_label: r.periodLabel,
      payment_date: r.date, term: currentTerm, paid_by: adminCreds.username,
    }));

    const { data: insertedPayments, error: spError } = await supabase
      .from("staff_payments").insert(staffPayInserts).select();
    if (spError) return notify(`Could not record staff payments: ${spError.message}`, "err");

    const newStaffPayments = (insertedPayments || []).map((p) => ({
      id: p.id, schoolId: p.school_id, staffId: p.staff_id, staffName: p.staff_name,
      amount: p.amount, payType: p.pay_type, periodLabel: p.period_label,
      date: p.payment_date, term: p.term, paidBy: p.paid_by || "",
    }));
    const newExpenseEntries = validRows.map((r, idx) => ({
      id: `${activeSchoolId}-e${Date.now()}_${idx}`, schoolId: activeSchoolId, category: "Salaries & Wages",
      description: `${r.matchedStaff.name} (${r.matchedStaff.role}) — ${r.periodLabel}`,
      amount: r.amount, date: r.date, term: currentTerm, paidBy: adminCreds.username,
    }));

    setAllStaffPayments(prev => ({ ...prev, [activeSchoolId]: [...(prev[activeSchoolId] || []), ...newStaffPayments] }));
    setAllExpenses(prev => ({ ...prev, [activeSchoolId]: [...(prev[activeSchoolId] || []), ...newExpenseEntries] }));

    setBulkStaffPayImportDone(true);
    const total = validRows.reduce((a, r) => a + r.amount, 0);
    notify(`✓ ${validRows.length} staff payments recorded — total ${fmt(total)}`);
  };

  const handleFeeEdit = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const { mode, bursaryType, bursaryValue, bursaryReason, customFee } = feeEditData;
    let updates = {};
    if (mode === "custom") updates = { bursary_type: null, bursary_value: null, bursary_reason: null, custom_fee: parseInt(customFee) || null };
    else if (mode === "bursary") updates = { custom_fee: null, bursary_type: bursaryType, bursary_value: parseFloat(bursaryValue), bursary_reason: bursaryReason };
    else updates = { bursary_type: null, bursary_value: null, bursary_reason: null, custom_fee: null };
    const { error } = await supabase.from("students").update(updates).eq("id", showFeeEdit.id);
    if (error) return notify(`Could not update fee: ${error.message}`, "err");
    setAllStudents(prev => ({
      ...prev,
      [activeSchoolId]: prev[activeSchoolId].map(s => {
        if (s.id !== showFeeEdit.id) return s;
        if (mode === "custom") return { ...s, bursary: null, customFee: parseInt(customFee) || null };
        if (mode === "bursary") return { ...s, customFee: null, bursary: { type: bursaryType, value: parseFloat(bursaryValue), reason: bursaryReason } };
        return { ...s, bursary: null, customFee: null };
      })
    }));
    notify(`Fee updated for ${showFeeEdit.name}`);
    setShowFeeEdit(null);
  };

  const handleAddExpense = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    if (!newExp.description || !newExp.amount) return notify("Fill all fields", "err");
    const { data: inserted, error } = await supabase.from("expenses").insert({
      school_id: activeSchoolId, category: newExp.category, description: newExp.description,
      amount: parseInt(newExp.amount), date: newExp.date, term: currentTerm, paid_by: "Admin",
    }).select().single();
    if (error) return notify(`Could not save expense: ${error.message}`, "err");
    const e = {
      id: inserted.id, schoolId: activeSchoolId, category: inserted.category,
      description: inserted.description, amount: inserted.amount,
      date: inserted.date, term: inserted.term, paidBy: inserted.paid_by || "",
    };
    setAllExpenses(prev => ({ ...prev, [activeSchoolId]: [...(prev[activeSchoolId] || []), e] }));
    notify("Expense recorded");
    setShowAddExp(false); setNewExp({ category: "Salaries", description: "", amount: "", date: new Date().toISOString().split("T")[0] });
  };

  // ── Staff / Workers ──────────────────────────────────────────────
  const handleAddStaff = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    if (!newStaff.name || !newStaff.role) return notify("Enter at least a name and role", "err");
    const { data: inserted, error } = await supabase.from("staff").insert({
      school_id: activeSchoolId,
      name: newStaff.name, role: newStaff.role, phone: newStaff.phone || "",
      default_rate: parseInt(newStaff.defaultRate) || 0,
      default_rate_type: newStaff.defaultRateType || "daily",
      active: true,
    }).select().single();
    if (error) return notify(`Could not add staff: ${error.message}`, "err");
    const st = {
      id: inserted.id, schoolId: activeSchoolId,
      name: inserted.name, role: inserted.role, phone: inserted.phone || "",
      defaultRate: inserted.default_rate || 0,
      defaultRateType: inserted.default_rate_type || "daily",
      active: true, photo: null,
    };
    setAllStaff(prev => ({ ...prev, [activeSchoolId]: [...(prev[activeSchoolId] || []), st] }));
    notify(`${st.name} added to staff`);
    setShowAddStaff(false);
    setNewStaff({ name: "", role: "", phone: "", defaultRate: "", defaultRateType: "daily" });
  };

  const handleEditStaff = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    if (!showEditStaff) return;
    const { error } = await supabase.from("staff").update({
      name: showEditStaff.name, role: showEditStaff.role, phone: showEditStaff.phone,
      default_rate: parseInt(showEditStaff.defaultRate) || 0,
      default_rate_type: showEditStaff.defaultRateType,
    }).eq("id", showEditStaff.id);
    if (error) return notify(`Could not update staff: ${error.message}`, "err");
    setAllStaff(prev => ({
      ...prev,
      [activeSchoolId]: prev[activeSchoolId].map(s => s.id === showEditStaff.id ? {
        ...s, name: showEditStaff.name, role: showEditStaff.role, phone: showEditStaff.phone,
        defaultRate: parseInt(showEditStaff.defaultRate) || 0, defaultRateType: showEditStaff.defaultRateType,
      } : s),
    }));
    notify(`${showEditStaff.name}'s details updated`);
    setShowEditStaff(null);
  };

  const handleToggleStaffActive = async (staffId) => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const current = (allStaff[activeSchoolId] || []).find(s => s.id === staffId);
    if (!current) return;
    const newActive = !current.active;
    const { error } = await supabase.from("staff").update({ active: newActive }).eq("id", staffId);
    if (error) return notify(`Could not update staff status: ${error.message}`, "err");
    setAllStaff(prev => ({
      ...prev,
      [activeSchoolId]: prev[activeSchoolId].map(s => s.id === staffId ? { ...s, active: newActive } : s),
    }));
  };

  // Recording a staff payment creates BOTH a staff payment record (for the Staff tab's
  // history) AND a matching expense under "Salaries & Wages" (so Dashboard/Net Surplus
  // figures stay accurate without the bursar having to enter the same amount twice).
  const handlePayStaff = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    if (!showPayStaff) return;
    const amount = parseInt(String(payStaffForm.amount).replace(/,/g, ""));
    if (!amount || amount <= 0) return notify("Enter a valid amount", "err");
    if (!payStaffForm.periodLabel) return notify("Describe what period this payment covers (e.g. \"May 2026\" or \"12 Jun 2026\")", "err");

    const today = new Date().toISOString().split("T")[0];
    const { data: inserted, error } = await supabase.from("staff_payments").insert({
      school_id: activeSchoolId, staff_id: showPayStaff.id, staff_name: showPayStaff.name,
      amount, pay_type: payStaffForm.payType, period_label: payStaffForm.periodLabel,
      payment_date: today, term: currentTerm, paid_by: "Admin",
    }).select().single();
    if (error) return notify(`Could not record payment: ${error.message}`, "err");

    const sp = {
      id: inserted.id, schoolId: activeSchoolId, staffId: showPayStaff.id, staffName: showPayStaff.name,
      amount, payType: payStaffForm.payType, periodLabel: payStaffForm.periodLabel,
      date: today, term: currentTerm, paidBy: "Admin",
    };
    // Also insert a linked expense so Net Surplus figures stay accurate
    const { data: expInserted } = await supabase.from("expenses").insert({
      school_id: activeSchoolId, category: "Salaries & Wages",
      description: `${showPayStaff.name} (${showPayStaff.role}) — ${payStaffForm.periodLabel}`,
      amount, date: today, term: currentTerm, paid_by: "Admin",
      staff_payment_id: inserted.id,
    }).select().single();
    const exp = expInserted ? {
      id: expInserted.id, schoolId: activeSchoolId, category: expInserted.category,
      description: expInserted.description, amount: expInserted.amount,
      date: expInserted.date, term: expInserted.term, paidBy: expInserted.paid_by || "",
      staffPaymentId: inserted.id,
    } : null;
    setAllStaffPayments(prev => ({ ...prev, [activeSchoolId]: [...(prev[activeSchoolId] || []), sp] }));
    if (exp) setAllExpenses(prev => ({ ...prev, [activeSchoolId]: [...(prev[activeSchoolId] || []), exp] }));
    notify(`✓ ${fmt(amount)} paid to ${showPayStaff.name}`);
    setShowPayStaff(null);
    setPayStaffForm({ amount: "", payType: "daily", periodLabel: "" });
  };

  const handleDeleteStaffPayment = async (paymentId) => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const sp = staffPayments.find(p => p.id === paymentId);
    if (!sp) return;
    // Delete the staff payment from Supabase
    const { error } = await supabase.from("staff_payments").delete().eq("id", paymentId);
    if (error) return notify(`Could not delete payment: ${error.message}`, "err");
    // Delete the linked expense too (cascade isn't set, so delete explicitly)
    await supabase.from("expenses").delete().eq("staff_payment_id", paymentId);
    setAllStaffPayments(prev => ({ ...prev, [activeSchoolId]: prev[activeSchoolId].filter(p => p.id !== paymentId) }));
    setAllExpenses(prev => ({ ...prev, [activeSchoolId]: (prev[activeSchoolId] || []).filter(e => e.staffPaymentId !== paymentId) }));
    notify("Staff payment deleted");
  };

  const handleRollover = async (newTerm) => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const sid = activeSchoolId;
    const current = allStudents[sid] || [];
    const dnrIds = Object.keys(rolloverDnr).filter(id => rolloverDnr[id]);
    const newAlumni = [];

    // Carry forward arrears for everyone first
    const withArrears = current.map(s => ({ ...s, arrears: getBalance(s, currentTerm).balance }));

    // Split out students marked as "Did Not Return"
    const staying = [];
    withArrears.forEach(s => {
      if (dnrIds.includes(String(s.id))) {
        newAlumni.push({
          ...s, status: "Did Not Return",
          leftClass: s.class,
          leftYear: currentTerm.split(", ")[1] || promotionYear,
          outstandingDebt: s.arrears || 0,
          leftNote: `Did not return for ${newTerm}`,
        });
      } else {
        staying.push(s);
      }
    });

    // Persist arrears for staying students
    const arrearUpdates = staying.map(s =>
      supabase.from("students").update({ arrears: s.arrears || 0 }).eq("id", s.id)
    );
    // Mark DNR students as alumni in Supabase
    const alumniUpdates = newAlumni.map(s =>
      supabase.from("students").update({ status: "Did Not Return", arrears: s.outstandingDebt || 0 }).eq("id", s.id)
    );
    await Promise.all([...arrearUpdates, ...alumniUpdates]);

    setAllStudents(prev => ({ ...prev, [sid]: staying }));
    if (newAlumni.length > 0) {
      setAllAlumni(prev => ({ ...prev, [sid]: [...(prev[sid] || []), ...newAlumni] }));
    }
    setCurrentTerm(newTerm);
    setShowRollover(false);
    setRolloverDnr({});
    setRolloverStep(1);
    notify(`Rolled over to ${newTerm}. Unpaid balances carried forward as arrears.${newAlumni.length > 0 ? ` ${newAlumni.length} student(s) moved to Alumni.` : ""}`);
  };

  // ── Promotion Handler ─────────────────────────────────────────
  const handlePromotion = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const sid = activeSchoolId;
    const current = allStudents[sid] || [];
    const newAlumni = [];
    const promoted = [];
    const lastClass = schoolClasses[schoolClasses.length - 1];
    const transitionClasses = getTransitionClasses(school);
    // Each entry: { studentId, supabasePayload } — collected as we decide each
    // student's outcome below, then sent to Supabase together in one batch
    // right before updating local state, so the database and the screen
    // never disagree about what happened.
    const dbUpdates = [];

    current.forEach(s => {
      // Outstanding balance = this term's fee + any arrears already carried in,
      // minus what's been paid this term. Same single formula used everywhere
      // else in the app (see getBalanceStatic) — not summed across every term
      // the school has ever had, which would massively overstate debt.
      const totalDebt = getBalance(s, currentTerm).balance;

      // Did Not Return — applies to ANY class
      if (dnrDecisions[s.id]) {
        const leftNote = `Did not return for ${promotionYear}`;
        newAlumni.push({ ...s, status: "Did Not Return", leftClass: s.class, leftYear: promotionYear, outstandingDebt: totalDebt, leftNote });
        dbUpdates.push({ id: s.id, status: "Did Not Return", left_class: s.class, left_year: promotionYear, outstanding_debt: totalDebt, left_note: leftNote });
        return;
      }

      // Repeating — stays in the same class, does not advance and is not treated as a transition/graduation case
      if (repeatDecisions[s.id]) {
        promoted.push({ ...s, class: s.class, stream: s.stream, arrears: totalDebt, payments: s.payments, isRepeating: true, repeatYear: promotionYear });
        dbUpdates.push({ id: s.id, class: s.class, arrears: totalDebt });
        return;
      }

      const clsIdx = schoolClasses.indexOf(s.class);

      if (s.class === lastClass) {
        // Final graduation point — always leaves (repeat case already handled above)
        const leftNote = `Completed ${s.class}`;
        newAlumni.push({ ...s, status: "Graduate", leftClass: s.class, leftYear: promotionYear, outstandingDebt: totalDebt, leftNote });
        dbUpdates.push({ id: s.id, status: "Graduate", left_class: s.class, left_year: promotionYear, outstanding_debt: totalDebt, left_note: leftNote });
      } else if (transitionClasses.includes(s.class)) {
        // End-of-level class (P7, S4, Top Class, etc.) — explicit continue/leave decision
        const decision = transitionDecisions[s.id] || "leave";
        if (decision === "leave") {
          const leftNote = `Left after ${s.class}`;
          newAlumni.push({ ...s, status: "Leaver", leftClass: s.class, leftYear: promotionYear, outstandingDebt: totalDebt, leftNote });
          dbUpdates.push({ id: s.id, status: "Leaver", left_class: s.class, left_year: promotionYear, outstanding_debt: totalDebt, left_note: leftNote });
        } else {
          const nextClass = schoolClasses[clsIdx + 1] || s.class;
          promoted.push({ ...s, class: nextClass, stream: "", arrears: totalDebt, payments: s.payments });
          dbUpdates.push({ id: s.id, class: nextClass, stream: "", arrears: totalDebt });
        }
      } else {
        // Normal promotion to next class in sequence
        const nextClass = schoolClasses[clsIdx + 1] || s.class;
        promoted.push({ ...s, class: nextClass, stream: "", arrears: totalDebt, payments: s.payments });
        dbUpdates.push({ id: s.id, class: nextClass, stream: "", arrears: totalDebt });
      }
    });

    // Send every student's update to Supabase as one batch before touching
    // local state — if some fail, we report exactly how many and don't
    // pretend the whole promotion succeeded.
    const updateResults = await Promise.all(dbUpdates.map(({ id, ...payload }) =>
      supabase.from("students").update(payload).eq("id", id)
    ));
    const failedCount = updateResults.filter(r => r.error).length;
    if (failedCount > 0) {
      notify(`${failedCount} of ${dbUpdates.length} students could not be updated — please check and retry for those`, "err");
    }

    setAllStudents(prev => ({ ...prev, [sid]: promoted }));
    setAllAlumni(prev => ({ ...prev, [sid]: [...(prev[sid] || []), ...newAlumni] }));
    setCurrentTerm(`Term 1, ${promotionYear}`);
    setShowPromotion(false);
    setTransitionDecisions({});
    setDnrDecisions({});
    const dnrCount = Object.values(dnrDecisions).filter(Boolean).length;
    const repeatCount = Object.values(repeatDecisions).filter(Boolean).length;
    setRepeatDecisions({});
    notify(`✓ ${promoted.length - repeatCount} promoted · ${repeatCount} repeating · ${newAlumni.length - dnrCount} to Alumni · ${dnrCount} Did Not Return`);
  };

  // ── Edit Expense Amount ───────────────────────────────────────
  const handleEditExpense = async () => {
    if (isReadOnly) return notify("Account is in read-only mode. Please renew your subscription to make changes.", "err");
    const amt = parseInt(editExpAmt);
    if (!amt || amt <= 0) return notify("Enter valid amount", "err");
    const { error } = await supabase.from("expenses").update({ amount: amt }).eq("id", showEditExpense.id);
    if (error) return notify(`Could not update expense: ${error.message}`, "err");
    setAllExpenses(prev => ({
      ...prev,
      [activeSchoolId]: prev[activeSchoolId].map(e =>
        e.id === showEditExpense.id ? { ...e, amount: amt } : e
      )
    }));
    notify(`Expense corrected to ${fmt(amt)}`);
    setShowEditExpense(null);
    setEditExpAmt("");
  };

  // ── Excel Export ──────────────────────────────────────────────
  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    // Students sheet
    const stuData = termStudents.map(s => {
      const bal = getBalance(s, currentTerm);
      return {
        Name: s.name, Class: s.class, Gender: s.gender === "M" ? "Male" : "Female",
        Parent: s.parent, Phone: s.phone,
        "Term Fee": bal.termFee, "Arrears": bal.arrears,
        "Total Due": bal.totalDue,
        "Amount Paid": bal.paidThisTerm,
        "Balance": bal.balance,
        Status: bal.status,
      };
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(stuData), "Students");
    // Payments sheet
    const payData = allPayments.map(p => ({
      "Receipt No": p.id, Student: p.studentName, Class: p.studentClass, Term: p.term,
      Date: fmtDate(p.date), Amount: p.amount, Method: p.method, "Received By": p.receivedBy,
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(payData), "Payments");
    // Expenses sheet
    const expData = expenses.filter(e => e.term === currentTerm).map(e => ({
      Category: e.category, Description: e.description, Amount: e.amount,
      Date: fmtDate(e.date), "Paid By": e.paidBy,
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(expData), "Expenses");
    XLSX.writeFile(wb, `FeeTrack_${school.name.replace(/ /g, "_")}_${currentTerm.replace(/ /g, "_")}.xlsx`);
    notify("Excel file downloaded ✓");
  };

  // ── Print Balances Report (by class) ────────────────────────────
  // ── Balances Report (printed inline via modal, since window.open is blocked) ──
  const getBalancesReportData = () => {
    const groups = {};
    filtered.forEach(s => {
      const key = classLabel(s.class, s.stream);
      if (!groups[key]) groups[key] = [];
      groups[key].push(s);
    });
    const sortedKeys = Object.keys(groups).sort((a, b) => {
      const baseA = a.split(" - ")[0], baseB = b.split(" - ")[0];
      const idxA = schoolClasses.indexOf(baseA), idxB = schoolClasses.indexOf(baseB);
      if (idxA !== idxB) return idxA - idxB;
      return a.localeCompare(b);
    });

    let grandExpected = 0, grandPaid = 0, grandBalance = 0;
    const sections = sortedKeys.map(key => {
      const students = groups[key];
      let clsExpected = 0, clsPaid = 0, clsBalance = 0;
      const rows = students.map(s => {
        const bal = getBalance(s, currentTerm);
        clsExpected += bal.totalDue; clsPaid += bal.paidThisTerm; clsBalance += bal.balance;
        return { name: s.name, parent: s.parent, phone: s.phone, due: bal.totalDue, paid: bal.paidThisTerm, balance: bal.balance, status: bal.status };
      });
      grandExpected += clsExpected; grandPaid += clsPaid; grandBalance += clsBalance;
      return { key, rows, clsExpected, clsPaid, clsBalance };
    });

    const filterLabel = [
      filterClass !== "All" ? filterClass : null,
      filterStatus !== "All" ? filterStatus : null,
      search ? `"${search}"` : null,
    ].filter(Boolean).join(" · ") || "All Students";

    return { sections, grandExpected, grandPaid, grandBalance, filterLabel, totalStudents: filtered.length };
  };

  // Triggers the browser print dialog, scoped to print only the given element id.
  // Each printable modal (receipt, balances report) has its OWN unique id — using one
  // shared id across multiple modals caused the print stylesheet to sometimes target the
  // wrong element (or nothing at all) since two DOM nodes can't validly share one id.
  const triggerPrint = (targetId = "feetrack-print-area") => {
    const styleId = "feetrack-print-style";
    let styleEl = document.getElementById(styleId);
    if (!styleEl) {
      styleEl = document.createElement("style");
      styleEl.id = styleId;
      document.head.appendChild(styleEl);
    }
    styleEl.innerHTML = `
      @media print {
        body * { visibility: hidden; }
        #${targetId}, #${targetId} * { visibility: visible; }
        #${targetId} { position: absolute; top: 0; left: 0; width: 100%; }
        .feetrack-no-print { display: none !important; }
      }
    `;
    window.print();
  };

  // ── Print Receipt ─────────────────────────────────────────────
  // Uses an inline print stylesheet + window.print() on the current page,
  // since window.open() for a new tab/window is blocked in this sandbox.
  const printReceipt = () => triggerPrint("feetrack-print-receipt");

  const card = { background: "#fff", borderRadius: 14, padding: isMobile ? "16px 16px" : "20px 22px", border: "1px solid #e8edf3", boxShadow: "0 2px 8px rgba(15,23,42,0.05)" };
  const inp = { width: "100%", padding: "9px 13px", borderRadius: 9, border: "1px solid #e2e8f0", fontSize: 13, outline: "none", boxSizing: "border-box", background: "#fff" };
  const lbl = { fontSize: 11, fontWeight: 700, color: "#64748b", display: "block", marginBottom: 5, textTransform: "uppercase", letterSpacing: 0.4 };
  // Responsive grid: collapses to fewer columns on mobile
  const grid = (desktopCols, mobileCols = 1) => ({ display: "grid", gridTemplateColumns: isMobile ? `repeat(${mobileCols}, 1fr)` : `repeat(${desktopCols}, 1fr)` });

  // ════════════════════════ LOGIN SCREEN ════════════════════════
  if (resetPasswordToken) {
    // Someone opened a password-reset email link — this takes priority
    // over everything else (including a saved session), since the whole
    // point of this screen is to let them in regardless of whether
    // they're currently logged in anywhere.
    return (
      <div style={{ minHeight: "100vh", background: "linear-gradient(135deg,#0f172a 0%,#1e3a5f 100%)", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Plus Jakarta Sans',sans-serif", padding: 20 }}>
        <div style={{ background: "#fff", borderRadius: 20, padding: isMobile ? 24 : 40, width: isMobile ? "100%" : 380, maxWidth: 380, boxShadow: "0 32px 80px rgba(0,0,0,0.3)", boxSizing: "border-box" }}>
          <div style={{ textAlign: "center", marginBottom: 28 }}>
            <div style={{ fontSize: 44, marginBottom: 8 }}>🏫</div>
            <div style={{ fontSize: 22, fontWeight: 800, color: "#0f172a" }}>FeeTrack UG</div>
            <div style={{ fontSize: 13, color: "#64748b", marginTop: 4 }}>Reset Your Password</div>
          </div>
          {resetStatus === "success" ? (
            <div style={{ textAlign: "center", padding: "10px 0" }}>
              <div style={{ fontSize: 48, marginBottom: 12 }}>✅</div>
              <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 8 }}>Password Updated</div>
              <div style={{ fontSize: 13, color: "#64748b", lineHeight: 1.6, marginBottom: 20 }}>
                You can now log in with your new password.
              </div>
              <button onClick={() => { window.location.href = window.location.pathname; }} style={{ width: "100%", padding: 12, borderRadius: 10, border: "none", background: "#0f172a", color: "#fff", fontSize: 14, fontWeight: 800, cursor: "pointer" }}>
                Go to Login
              </button>
              <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 10, textAlign: "center" }}>
                (This will refresh the page so your new password takes effect)
              </div>
            </div>
          ) : (
            <div>
              <div style={{ marginBottom: 16 }}>
                <label style={lbl}>New Password</label>
                <input type="password" value={resetNewPassword} onChange={e => setResetNewPassword(e.target.value)} placeholder="min 6 characters" style={inp} />
              </div>
              <div style={{ marginBottom: 16 }}>
                <label style={lbl}>Confirm New Password</label>
                <input type="password" value={resetConfirmPassword} onChange={e => setResetConfirmPassword(e.target.value)} onKeyDown={e => e.key === "Enter" && handleCompletePasswordReset()} placeholder="re-enter password" style={inp} />
              </div>
              {resetStatus && resetStatus !== "submitting" && (
                <div style={{ background: "#fef2f2", color: "#b91c1c", padding: "10px 14px", borderRadius: 9, fontSize: 12, marginBottom: 16, fontWeight: 600 }}>{resetStatus}</div>
              )}
              <button onClick={handleCompletePasswordReset} disabled={resetStatus === "submitting"} style={{ width: "100%", padding: 13, borderRadius: 10, border: "none", background: "linear-gradient(135deg,#f59e0b,#ef4444)", color: "#fff", fontSize: 15, fontWeight: 800, cursor: resetStatus === "submitting" ? "default" : "pointer", opacity: resetStatus === "submitting" ? 0.7 : 1 }}>
                {resetStatus === "submitting" ? "Updating..." : "Update Password"}
              </button>
            </div>
          )}
        </div>
      </div>
    );
  }
  if (!currentUser && hasSavedSession && schoolsLoading) {
    // We know there's a saved session waiting to be checked, but the real
    // schools data hasn't loaded yet — show a brief, neutral loading state
    // instead of flashing the login form, which would otherwise show for a
    // moment even though the person is about to be logged back in anyway.
    return (
      <div style={{ minHeight: "100vh", background: "linear-gradient(135deg,#0f172a 0%,#1e3a5f 100%)", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Plus Jakarta Sans',sans-serif" }}>
        <div style={{ color: "#fff", fontSize: 15, fontWeight: 600, opacity: 0.85 }}>🏫 FeeTrack UG</div>
      </div>
    );
  }
  if (!currentUser) {
    return (
      <div style={{ minHeight: "100vh", background: "linear-gradient(135deg,#0f172a 0%,#1e3a5f 100%)", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Plus Jakarta Sans',sans-serif", padding: 20 }}>
        <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet" />
        <div style={{ background: "#fff", borderRadius: 20, padding: isMobile ? 24 : 40, width: isMobile ? "100%" : (loginScreen === "signup" ? 460 : 380), maxWidth: loginScreen === "signup" ? 460 : 380, boxShadow: "0 32px 80px rgba(0,0,0,0.3)", maxHeight: "90vh", overflowY: "auto", boxSizing: "border-box" }}>
          <div style={{ textAlign: "center", marginBottom: 28 }}>
            <div style={{ fontSize: 44, marginBottom: 8 }}>🏫</div>
            <div style={{ fontSize: 22, fontWeight: 800, color: "#0f172a" }}>FeeTrack UG</div>
            <div style={{ fontSize: 13, color: "#64748b", marginTop: 4 }}>School Finance Management System</div>
          </div>

          <div style={{ display: "flex", gap: 6, marginBottom: 24, background: "#f1f5f9", borderRadius: 10, padding: 4 }}>
            {["admin", "parent", "signup"].map(t => (
              <button key={t} onClick={() => { setLoginScreen(t); setLoginError(""); setSignupSubmitted(false); }} style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: "none", cursor: "pointer", fontWeight: 700, fontSize: 12, background: loginScreen === t ? "#fff" : "transparent", color: loginScreen === t ? "#0f172a" : "#64748b", boxShadow: loginScreen === t ? "0 1px 4px rgba(0,0,0,0.1)" : "none" }}>
                {t === "admin" ? "🔑 Admin" : t === "parent" ? "👨‍👩‍👧 Parent" : "🏫 Sign Up School"}
              </button>
            ))}
          </div>

          {/* ── SIGNUP FORM ── */}
          {loginScreen === "signup" && (
            signupSubmitted ? (
              <div style={{ textAlign: "center", padding: "20px 0" }}>
                <div style={{ fontSize: 48, marginBottom: 12 }}>✅</div>
                <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 8 }}>Application Submitted!</div>
                <div style={{ fontSize: 13, color: "#64748b", lineHeight: 1.6, marginBottom: 20 }}>
                  Thank you for registering <strong>{signupForm.schoolName}</strong>. Our team will review your application within 24 hours.
                  You'll be notified via SMS/email once approved, and can then log in with the username and password you created.
                </div>
                <button onClick={() => { setLoginScreen("admin"); setSignupForm({ schoolName: "", location: "", principal: "", phone: "", email: "", students: "", schoolType: "secondary", username: "", password: "", confirmPassword: "" }); setSignupSubmitted(false); }}
                  style={{ width: "100%", padding: 12, borderRadius: 10, border: "none", background: "#0f172a", color: "#fff", fontSize: 14, fontWeight: 800, cursor: "pointer" }}>
                  Back to Login
                </button>
              </div>
            ) : (
              <div>
                <div style={{ fontSize: 13, color: "#64748b", marginBottom: 18, textAlign: "center" }}>
                  Register your school below. An admin will review and approve your account within 24 hours.
                </div>
                {[
                  { label: "School Name *", key: "schoolName", placeholder: "e.g. Jinja Progressive College" },
                  { label: "Location *", key: "location", placeholder: "e.g. Jinja" },
                  { label: "Principal / Headteacher *", key: "principal", placeholder: "e.g. Mrs. Akello Joan" },
                  { label: "School Phone *", key: "phone", placeholder: "e.g. 0772-555-101" },
                  { label: "School Email (for password reset)", key: "email", placeholder: "e.g. info@yourschool.ac.ug" },
                  { label: "Approx. Number of Students", key: "students", placeholder: "e.g. 350" },
                ].map(f => (
                  <div key={f.key} style={{ marginBottom: 12 }}>
                    <label style={lbl}>{f.label}</label>
                    <input value={signupForm[f.key]} onChange={e => setSignupForm(p => ({ ...p, [f.key]: e.target.value }))} placeholder={f.placeholder} style={inp} />
                  </div>
                ))}
                <div style={{ marginBottom: 12 }}>
                  <label style={lbl}>School Type *</label>
                  <select value={signupForm.schoolType} onChange={e => setSignupForm(p => ({ ...p, schoolType: e.target.value }))} style={inp}>
                    {Object.entries(SCHOOL_TYPES).map(([key, t]) => <option key={key} value={key}>{t.label}</option>)}
                  </select>
                  <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 4 }}>This determines which classes/levels appear in your account (e.g. Nursery: Baby–Top Class, Primary: P1–P7, Secondary: S1–S6).</div>
                </div>
                <div style={{ marginBottom: 12 }}>
                  <label style={lbl}>Billing Cycle *</label>
                  <div style={{ display: "flex", gap: 10 }}>
                    {(() => {
                      const previewPlan = suggestPlan(signupForm.students);
                      return ["monthly", "term"].map(cyc => {
                        const info = getBillingInfo(previewPlan, cyc);
                        const selected = signupForm.billingCycle === cyc;
                        return (
                          <button key={cyc} type="button" onClick={() => setSignupForm(p => ({ ...p, billingCycle: cyc }))}
                            style={{ flex: 1, padding: "12px 14px", borderRadius: 10, border: `2px solid ${selected ? "#f59e0b" : "#e2e8f0"}`, background: selected ? "#fffbeb" : "#fff", cursor: "pointer", textAlign: "left" }}>
                            <div style={{ fontSize: 13, fontWeight: 800, color: selected ? "#92400e" : "#374151" }}>{cyc === "monthly" ? "Monthly" : "Per Term"}</div>
                            <div style={{ fontSize: 15, fontWeight: 800, color: "#0f172a", marginTop: 2 }}>{fmt(info.price)}<span style={{ fontSize: 11, color: "#94a3b8" }}>{info.periodLabel}</span></div>
                            {cyc === "term" && <div style={{ fontSize: 10, color: "#15803d", fontWeight: 700, marginTop: 2 }}>Save 10% vs. paying monthly</div>}
                          </button>
                        );
                      });
                    })()}
                  </div>
                  <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 4 }}>Pay every month, or pay once per school term (about every 3 months) at a 10% discount. Estimated plan: <strong>{suggestPlan(signupForm.students)}</strong> — confirmed on approval.</div>
                </div>
                <div style={{ borderTop: "1px dashed #e2e8f0", margin: "16px 0", paddingTop: 14 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 10 }}>Choose your login credentials (used after approval)</div>
                  <div style={{ ...grid(2, 1), gap: 12 }}>
                    <div>
                      <label style={lbl}>Username *</label>
                      <input value={signupForm.username} onChange={e => setSignupForm(p => ({ ...p, username: e.target.value }))} placeholder="e.g. jinjapc_admin" style={inp} />
                    </div>
                    <div>
                      <label style={lbl}>Password *</label>
                      <input type="password" value={signupForm.password} onChange={e => setSignupForm(p => ({ ...p, password: e.target.value }))} placeholder="min 6 characters" style={inp} />
                    </div>
                    <div>
                      <label style={lbl}>Confirm Password *</label>
                      <input type="password" value={signupForm.confirmPassword} onChange={e => setSignupForm(p => ({ ...p, confirmPassword: e.target.value }))} placeholder="re-enter password" style={inp} />
                      {signupForm.confirmPassword && signupForm.password !== signupForm.confirmPassword && (
                        <div style={{ fontSize: 11, color: "#dc2626", marginTop: 4, fontWeight: 600 }}>Passwords do not match</div>
                      )}
                    </div>
                  </div>
                </div>
                {loginError && <div style={{ background: "#fef2f2", color: "#b91c1c", padding: "10px 14px", borderRadius: 9, fontSize: 12, marginBottom: 14, fontWeight: 600 }}>{loginError}</div>}
                <button onClick={handleSignupSubmit} style={{ width: "100%", padding: 13, borderRadius: 10, border: "none", background: "linear-gradient(135deg,#10b981,#059669)", color: "#fff", fontSize: 15, fontWeight: 800, cursor: "pointer", marginTop: 6 }}>
                  Submit Application
                </button>
                <div style={{ marginTop: 12, textAlign: "center", fontSize: 11, color: "#94a3b8" }}>
                  Already approved? Switch to the Admin tab to log in.
                </div>
              </div>
            )
          )}

          {/* ── ADMIN / PARENT LOGIN ── */}
          {loginScreen !== "signup" && loginScreen !== "forgot-password" && (
            <>
              <div style={{ marginBottom: 14 }}>
                <label style={lbl}>{loginScreen === "admin" ? "Username" : "Phone Number"}</label>
                <input value={loginInput.user} onChange={e => setLoginInput({ ...loginInput, user: e.target.value })} placeholder={loginScreen === "admin" ? "admin" : "e.g. 0772441823"} style={inp} />
              </div>
              <div style={{ marginBottom: 20 }}>
                <label style={lbl}>{loginScreen === "admin" ? "Password" : "PIN (last 4 digits of phone)"}</label>
                <input type="password" value={loginInput.pass} onChange={e => setLoginInput({ ...loginInput, pass: e.target.value })} onKeyDown={e => e.key === "Enter" && handleLogin()} placeholder={loginScreen === "admin" ? "admin123" : "e.g. 1823"} style={inp} />
              </div>
              {loginError && <div style={{ background: "#fef2f2", color: "#b91c1c", padding: "10px 14px", borderRadius: 9, fontSize: 12, marginBottom: 16, fontWeight: 600 }}>{loginError}</div>}
              <button onClick={handleLogin} style={{ width: "100%", padding: 13, borderRadius: 10, border: "none", background: "linear-gradient(135deg,#f59e0b,#ef4444)", color: "#fff", fontSize: 15, fontWeight: 800, cursor: "pointer" }}>
                {loginScreen === "admin" ? "Login to Dashboard" : "View My Child's Account"}
              </button>
              {loginScreen === "admin" && (
                <div style={{ marginTop: 10, textAlign: "center" }}>
                  <button onClick={() => { setLoginScreen("forgot-password"); setForgotPasswordEmail(""); setForgotPasswordStatus(null); }} style={{ background: "none", border: "none", color: "#64748b", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                    Forgot password?
                  </button>
                </div>
              )}
              <div style={{ marginTop: 16, textAlign: "center", fontSize: 11, color: "#94a3b8" }}>
                {loginScreen === "admin" ? "Demo school login: admin / admin123" : "Enter your phone number & last 4 digits as PIN"}
              </div>
              {loginScreen === "admin" && (
                <div style={{ marginTop: 10, textAlign: "center" }}>
                  <button onClick={() => setLoginScreen("signup")} style={{ background: "none", border: "none", color: "#3b82f6", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                    Don't have an account? Sign up your school →
                  </button>
                </div>
              )}
            </>
          )}

          {/* ── FORGOT PASSWORD ── */}
          {loginScreen === "forgot-password" && (
            forgotPasswordStatus === "sent" ? (
              <div style={{ textAlign: "center", padding: "20px 0" }}>
                <div style={{ fontSize: 48, marginBottom: 12 }}>📧</div>
                <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 8 }}>Check your email</div>
                <div style={{ fontSize: 13, color: "#64748b", lineHeight: 1.6, marginBottom: 20 }}>
                  If that email is registered with a school, a reset link has been sent. It expires in 1 hour. Check your spam folder if you don't see it within a few minutes.
                </div>
                <button onClick={() => setLoginScreen("admin")} style={{ width: "100%", padding: 12, borderRadius: 10, border: "none", background: "#0f172a", color: "#fff", fontSize: 14, fontWeight: 800, cursor: "pointer" }}>
                  Back to Login
                </button>
              </div>
            ) : (
              <div>
                <div style={{ fontSize: 13, color: "#64748b", marginBottom: 18, textAlign: "center" }}>
                  Enter the email address registered to your school. We'll send a link to reset your password.
                </div>
                <div style={{ marginBottom: 16 }}>
                  <label style={lbl}>School Email</label>
                  <input value={forgotPasswordEmail} onChange={e => setForgotPasswordEmail(e.target.value)} onKeyDown={e => e.key === "Enter" && handleForgotPassword()} placeholder="e.g. info@yourschool.ac.ug" style={inp} />
                </div>
                {forgotPasswordStatus && forgotPasswordStatus !== "sending" && (
                  <div style={{ background: "#fef2f2", color: "#b91c1c", padding: "10px 14px", borderRadius: 9, fontSize: 12, marginBottom: 16, fontWeight: 600 }}>{forgotPasswordStatus}</div>
                )}
                <button onClick={handleForgotPassword} disabled={forgotPasswordStatus === "sending"} style={{ width: "100%", padding: 13, borderRadius: 10, border: "none", background: "linear-gradient(135deg,#f59e0b,#ef4444)", color: "#fff", fontSize: 15, fontWeight: 800, cursor: forgotPasswordStatus === "sending" ? "default" : "pointer", opacity: forgotPasswordStatus === "sending" ? 0.7 : 1 }}>
                  {forgotPasswordStatus === "sending" ? "Sending..." : "Send Reset Link"}
                </button>
                <div style={{ marginTop: 10, textAlign: "center" }}>
                  <button onClick={() => setLoginScreen("admin")} style={{ background: "none", border: "none", color: "#64748b", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                    ← Back to Login
                  </button>
                </div>
              </div>
            )
          )}
        </div>
      </div>
    );
  }

  // ════════════════════════ SUPER ADMIN DASHBOARD ════════════════════════
  if (currentUser.role === "superadmin") {
    const pending = pendingSchools.filter(p => p.status === "pending");
    const approved = pendingSchools.filter(p => p.status === "approved");
    const rejected = pendingSchools.filter(p => p.status === "rejected");
    const allSchoolsList = Object.values(SCHOOLS_DATA).map(s => ({ ...s, subInfo: getSubscriptionInfo(s) }));
    // MRR = Monthly Recurring Revenue. For schools paying per term, normalize their
    // termly price down to a monthly-equivalent (÷3) so all schools compare on the
    // same basis regardless of which billing cycle they chose.
    const totalMRR = allSchoolsList.reduce((a, s) => {
      const billing = getBillingInfo(s.plan, s.billingCycle, s.customPrice);
      const monthlyEquivalent = billing.cycle === "term" ? billing.price / 3 : billing.price;
      return a + monthlyEquivalent;
    }, 0);
    const overdueSchools = allSchoolsList.filter(s => s.subInfo.status !== "Active");

    return (
      <div style={{ minHeight: "100vh", background: "#f4f6fb", fontFamily: "'Plus Jakarta Sans',sans-serif" }}>
        <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet" />

        {toast && (
          <div style={{ position: "fixed", top: 20, right: 20, zIndex: 999, padding: "12px 20px", borderRadius: 12, fontWeight: 600, fontSize: 13, boxShadow: "0 8px 24px rgba(0,0,0,0.15)", background: toast.type === "err" ? "#fef2f2" : "#f0fdf4", color: toast.type === "err" ? "#b91c1c" : "#15803d", border: `1px solid ${toast.type === "err" ? "#fca5a5" : "#86efac"}` }}>
            {toast.type === "err" ? "✗ " : "✓ "}{toast.msg}
          </div>
        )}

        <div style={{ background: "#0f172a", padding: "16px 24px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 24 }}>👑</span>
            <div>
              <div style={{ color: "#f1f5f9", fontWeight: 800, fontSize: 14 }}>FeeTrack UG — Super Admin</div>
              <div style={{ color: "#64748b", fontSize: 11 }}>Manage school signups & subscriptions</div>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            {paymentNotices.filter(n => n.status === "pending").length > 0 && (
              <button onClick={() => setSuperAdminTab("billing")} style={{ position: "relative", background: "#1e293b", border: "none", borderRadius: 8, padding: "8px 10px", cursor: "pointer" }}>
                <span style={{ fontSize: 16 }}>🔔</span>
                <span style={{ position: "absolute", top: -4, right: -4, background: "#ef4444", color: "#fff", borderRadius: 99, fontSize: 9, fontWeight: 800, padding: "1px 5px", minWidth: 14, textAlign: "center" }}>
                  {paymentNotices.filter(n => n.status === "pending").length}
                </span>
              </button>
            )}
            <button onClick={logout} style={{ background: "#1e293b", color: "#94a3b8", border: "none", borderRadius: 8, padding: "7px 14px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>Logout</button>
          </div>
        </div>

        <div style={{ maxWidth: 1000, margin: "32px auto", padding: "0 20px" }}>

          {/* Tab navigation */}
          <div style={{ display: "flex", gap: 8, marginBottom: 20, background: "#fff", borderRadius: 10, padding: 4, border: "1px solid #e8edf3" }}>
            {[{ id: "signups", label: "📋 Signups", badge: pending.length }, { id: "billing", label: "💳 Billing & Subscriptions", badge: overdueSchools.length + paymentNotices.filter(n => n.status === "pending").length }, { id: "settings", label: "⚙ Settings", badge: 0 }].map(t => (
              <button key={t.id} onClick={() => setSuperAdminTab(t.id)} style={{ flex: 1, padding: "10px", borderRadius: 8, border: "none", cursor: "pointer", fontWeight: 700, fontSize: 13, background: superAdminTab === t.id ? "#0f172a" : "transparent", color: superAdminTab === t.id ? "#fff" : "#64748b", position: "relative" }}>
                {t.label}
                {t.badge > 0 && <span style={{ marginLeft: 6, background: "#ef4444", color: "#fff", borderRadius: 99, fontSize: 10, fontWeight: 800, padding: "1px 6px" }}>{t.badge}</span>}
              </button>
            ))}
          </div>

          {superAdminTab === "signups" && (
          <div>
          {/* Summary cards */}
          <div style={{ ...grid(4, 2), gap: 14, marginBottom: 24 }}>
            {[
              { label: "Pending Approval", value: pending.length, color: "#f59e0b", icon: "⏳" },
              { label: "Approved Schools", value: Object.keys(SCHOOLS_DATA).length, color: "#10b981", icon: "✅" },
              { label: "Rejected", value: rejected.length, color: "#ef4444", icon: "✗" },
              { label: "Total Applications", value: pendingSchools.length, color: "#3b82f6", icon: "📋" },
            ].map((c, i) => (
              <div key={i} style={{ ...card, borderTop: `3px solid ${c.color}`, padding: "16px 18px" }}>
                <div style={{ display: "flex", justifyContent: "space-between" }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase" }}>{c.label}</div>
                  <span style={{ fontSize: 18 }}>{c.icon}</span>
                </div>
                <div style={{ fontSize: 26, fontWeight: 800, color: "#0f172a", marginTop: 6 }}>{c.value}</div>
              </div>
            ))}
          </div>

          {/* Pending approvals */}
          <div style={{ ...card, marginBottom: 20 }}>
            <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a", marginBottom: 4 }}>⏳ Pending School Signups</div>
            <div style={{ fontSize: 12, color: "#64748b", marginBottom: 16 }}>Review each application and approve or reject</div>

            {pending.length === 0 ? (
              <div style={{ textAlign: "center", padding: 40, color: "#94a3b8" }}>
                <div style={{ fontSize: 36, marginBottom: 8 }}>📭</div>
                No pending applications
              </div>
            ) : pending.map(p => (
              <div key={p.id} style={{ border: "1px solid #fde68a", background: "#fffbeb", borderRadius: 12, padding: 18, marginBottom: 12 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 800, fontSize: 16, color: "#0f172a", marginBottom: 6 }}>🏫 {p.schoolName}</div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, fontSize: 12, color: "#374151" }}>
                      <div><span style={{ color: "#94a3b8" }}>Location:</span> {p.location || "—"}</div>
                      <div><span style={{ color: "#94a3b8" }}>Principal:</span> {p.principal}</div>
                      <div><span style={{ color: "#94a3b8" }}>Phone:</span> {p.phone}</div>
                      <div><span style={{ color: "#94a3b8" }}>Email:</span> {p.email || "—"}</div>
                      <div><span style={{ color: "#94a3b8" }}>Students:</span> {p.students || "—"}</div>
                      <div><span style={{ color: "#94a3b8" }}>Type:</span> {SCHOOL_TYPES[p.schoolType]?.label || "Secondary"}</div>
                      <div><span style={{ color: "#94a3b8" }}>Billing:</span> {p.billingCycle === "term" ? "Per Term" : "Monthly"}</div>
                      <div><span style={{ color: "#94a3b8" }}>Submitted:</span> {fmtDate(p.submittedAt)}</div>
                    </div>
                    <div style={{ marginTop: 8 }}>
                      <span style={{ fontSize: 11, color: "#94a3b8" }}>Plan on approval: </span>
                      <Pill text={`${suggestPlan(p.students)} (${fmt(getBillingInfo(suggestPlan(p.students), p.billingCycle).price)}${getBillingInfo(suggestPlan(p.students), p.billingCycle).periodLabel})`} bg="#f5f3ff" col="#7c3aed" />
                    </div>
                    <div style={{ marginTop: 10, padding: "8px 12px", background: "#fff", borderRadius: 8, fontSize: 12 }}>
                      <span style={{ color: "#94a3b8" }}>Login username they chose:</span> <strong style={{ fontFamily: "monospace" }}>{p.username || "admin"}</strong>
                    </div>
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 8, marginLeft: 16 }}>
                    <button onClick={() => approveSchool(p.id)} style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 9, padding: "10px 20px", fontWeight: 800, fontSize: 13, cursor: "pointer" }}>✓ Approve</button>
                    <button onClick={() => rejectSchool(p.id)} style={{ background: "#fef2f2", color: "#dc2626", border: "1px solid #fca5a5", borderRadius: 9, padding: "8px 20px", fontWeight: 700, fontSize: 12, cursor: "pointer" }}>✗ Reject</button>
                  </div>
                </div>
              </div>
            ))}
          </div>

          {/* Approved schools list */}
          <div style={{ ...card }}>
            <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a", marginBottom: 16 }}>✅ Active Schools on Platform</div>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr style={{ background: "#f8fafc" }}>
                {["School", "Location", "Principal", "Login Username", "Plan"].map(h => (
                  <th key={h} style={{ padding: "9px 12px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase" }}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {Object.values(SCHOOLS_DATA).map((s, i) => (
                  <tr key={s.id} style={{ borderTop: "1px solid #f1f5f9", background: i % 2 === 0 ? "#fff" : "#fafbfc" }}>
                    <td style={{ padding: "10px 12px", fontWeight: 600, fontSize: 13 }}>{s.logo} {s.name}</td>
                    <td style={{ padding: "10px 12px", fontSize: 12, color: "#64748b" }}>{s.location}</td>
                    <td style={{ padding: "10px 12px", fontSize: 12, color: "#374151" }}>{s.principal}</td>
                    <td style={{ padding: "10px 12px", fontSize: 12, fontFamily: "monospace", color: "#3b82f6" }}>{s.adminUsername || "admin"}</td>
                    <td style={{ padding: "10px 12px" }}><Pill text={s.plan || "Starter"} bg="#eff6ff" col="#2563eb" /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          </div>
          )}

          {superAdminTab === "billing" && (
          <div>
            {/* MRR + overdue summary */}
            <div style={{ ...grid(3, 1), gap: 14, marginBottom: 24 }}>
              <div style={{ ...card, borderTop: "3px solid #10b981", padding: "16px 18px" }}>
                <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase" }}>Monthly Recurring Revenue</div>
                <div style={{ fontSize: 24, fontWeight: 800, color: "#0f172a", marginTop: 6 }}>{fmt(totalMRR)}</div>
                <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 2 }}>Across {Object.keys(SCHOOLS_DATA).length} schools</div>
              </div>
              <div style={{ ...card, borderTop: "3px solid #f59e0b", padding: "16px 18px" }}>
                <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase" }}>Grace Period</div>
                <div style={{ fontSize: 24, fontWeight: 800, color: "#0f172a", marginTop: 6 }}>{allSchoolsList.filter(s => s.subInfo.status === "Grace Period").length}</div>
                <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 2 }}>Schools in {GRACE_PERIOD_DAYS}-day grace window</div>
              </div>
              <div style={{ ...card, borderTop: "3px solid #ef4444", padding: "16px 18px" }}>
                <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase" }}>Suspended</div>
                <div style={{ fontSize: 24, fontWeight: 800, color: "#0f172a", marginTop: 6 }}>{allSchoolsList.filter(s => s.subInfo.status === "Suspended").length}</div>
                <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 2 }}>Read-only mode — payment required</div>
              </div>
            </div>

            {/* Payment Notices from Schools */}
            {paymentNotices.filter(n => n.status === "pending").length > 0 && (
              <div style={{ ...card, marginBottom: 20, border: "2px solid #bfdbfe", background: "#eff6ff" }}>
                <div style={{ fontWeight: 800, fontSize: 15, color: "#1d4ed8", marginBottom: 4 }}>📨 Payment Notices from Schools</div>
                <div style={{ fontSize: 12, color: "#64748b", marginBottom: 14 }}>
                  Schools have notified you they've sent payment and their access is temporarily protected from suspension while you check. Verify against your statement, then confirm or reject below.
                </div>
                {paymentNotices.filter(n => n.status === "pending").map(n => (
                  <div key={n.id} style={{ background: "#fff", border: "1px solid #bfdbfe", borderRadius: 10, padding: 14, marginBottom: 8, display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
                    <div>
                      <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a" }}>{n.schoolName}</div>
                      <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>
                        {n.method} · {fmt(n.amount)} · {fmtDate(n.date)} · Ref: <strong style={{ fontFamily: "monospace" }}>{n.billingRef}</strong>
                      </div>
                      {n.note && <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 2 }}>Note: {n.note}</div>}
                      <div style={{ fontSize: 11, color: "#2563eb", marginTop: 4, fontWeight: 600 }}>📨 Currently protected from suspension</div>
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                      <button onClick={() => markSubscriptionPaid(n.schoolId)} style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 9, padding: "8px 14px", fontWeight: 800, fontSize: 12, cursor: "pointer" }}>✓ Confirm (+{getBillingInfo(SCHOOLS_DATA[n.schoolId]?.plan, SCHOOLS_DATA[n.schoolId]?.billingCycle, SCHOOLS_DATA[n.schoolId]?.customPrice).cycleDays} days)</button>
                      <button onClick={() => setConfirmDialog({
                        title: "Payment Not Received",
                        message: `Confirm payment from ${n.schoolName} was NOT received? This removes their protection and resumes normal grace period / suspension rules.`,
                        danger: true,
                        onConfirm: () => rejectPaymentNotice(n.id),
                      })} style={{ background: "#fef2f2", color: "#dc2626", border: "1px solid #fca5a5", borderRadius: 9, padding: "8px 14px", fontWeight: 700, fontSize: 12, cursor: "pointer" }}>✗ Not Received</button>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {/* Recently Rejected Payment Notices */}
            {paymentNotices.filter(n => n.status === "not_found").length > 0 && (
              <div style={{ ...card, marginBottom: 20, border: "1px solid #fca5a5", background: "#fef2f2" }}>
                <div style={{ fontWeight: 800, fontSize: 14, color: "#b91c1c", marginBottom: 4 }}>✗ Payment Not Found — School Notified</div>
                <div style={{ fontSize: 12, color: "#64748b", marginBottom: 10 }}>These schools' protection has been removed. Their normal grace period / suspension countdown resumed from where it left off.</div>
                {paymentNotices.filter(n => n.status === "not_found").map(n => (
                  <div key={n.id} style={{ fontSize: 12, color: "#374151", padding: "6px 0", borderTop: "1px solid #fecaca" }}>
                    {n.schoolName} · {n.method} · {fmt(n.amount)} · {fmtDate(n.date)} · Ref: {n.billingRef}
                  </div>
                ))}
              </div>
            )}

            {/* Automatic Reconciliation */}
            <div style={{ ...card, marginBottom: 20, border: "2px solid #8b5cf6", background: "linear-gradient(135deg,#f5f3ff,#fff)" }}>
              <div style={{ fontWeight: 800, fontSize: 15, color: "#6d28d9", marginBottom: 4 }}>⚡ Automatic Subscription Renewal</div>
              <div style={{ fontSize: 12, color: "#64748b", marginBottom: 14, lineHeight: 1.6 }}>
                Upload your MTN MoMo or bank statement (Excel/CSV). The system reads each school's <strong>billing reference code</strong> from the payment narration and automatically reactivates matching accounts — extending their subscription by a month or a term, depending on each school's chosen billing cycle, with no manual work.
              </div>

              {subPayRows.length === 0 ? (
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                  <button onClick={() => subPayFileRef.current?.click()} style={{ background: "#8b5cf6", color: "#fff", border: "none", borderRadius: 9, padding: "10px 18px", fontWeight: 800, fontSize: 13, cursor: "pointer" }}>
                    📂 Upload Payment Statement
                  </button>
                  <input ref={subPayFileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleSubPayFileUpload} style={{ display: "none" }} />
                </div>
              ) : !subPayResults ? (
                <div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
                    <div style={{ fontSize: 13, fontWeight: 700, color: "#0f172a" }}>📄 {subPayFileName} — {subPayRows.length} credit entries found</div>
                    <button onClick={() => { setSubPayRows([]); setSubPayFileName(""); }} style={{ background: "#f1f5f9", color: "#64748b", border: "none", borderRadius: 8, padding: "6px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>✕ Clear</button>
                  </div>
                  <button onClick={handleSubPayReconcile} style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 9, padding: "10px 18px", fontWeight: 800, fontSize: 13, cursor: "pointer" }}>
                    ⚡ Run Auto-Reconciliation
                  </button>
                </div>
              ) : (
                <div>
                  {/* Results */}
                  {subPayResults.matched.length > 0 && (
                    <div style={{ marginBottom: 14 }}>
                      <div style={{ fontSize: 13, fontWeight: 700, color: "#15803d", marginBottom: 8 }}>✅ {subPayResults.matched.length} School(s) Automatically Reactivated</div>
                      {subPayResults.matched.map((m, i) => (
                        <div key={i} style={{ background: m.planChanged ? "#f5f3ff" : "#f0fdf4", border: `1px solid ${m.planChanged ? "#c4b5fd" : "#86efac"}`, borderRadius: 10, padding: 12, marginBottom: 8 }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                            <div>
                              <div style={{ fontWeight: 700, fontSize: 13, color: "#0f172a" }}>{m.school}</div>
                              <div style={{ fontSize: 11, color: "#64748b" }}>Ref: {m.reference} · Paid {fmt(m.amount)} · Was: {m.prevStatus}</div>
                            </div>
                            <div style={{ textAlign: "right" }}>
                              <Pill text="Active ✓" bg="#d1fae5" col="#065f46" />
                              <div style={{ fontSize: 11, color: "#64748b", marginTop: 2 }}>New due: {fmtDate(m.newDueDate)}</div>
                            </div>
                          </div>
                          {m.planChanged && (
                            <div style={{ fontSize: 12, color: "#6d28d9", marginTop: 8, fontWeight: 700, background: "#fff", borderRadius: 7, padding: "6px 10px" }}>
                              🔄 Plan automatically changed: {m.prevPlan} → {m.plan} (based on amount paid)
                            </div>
                          )}
                          {!m.amountMatch && (
                            <div style={{ fontSize: 11, color: "#92400e", marginTop: 6, fontWeight: 600 }}>
                              ⚠ Amount paid ({fmt(m.amount)}) is less than the {m.plan} plan price ({fmt(m.expectedAmount)}). Account reactivated anyway — review if needed.
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                  {subPayResults.unmatched.length > 0 && (
                    <div>
                      <div style={{ fontSize: 13, fontWeight: 700, color: "#dc2626", marginBottom: 8 }}>⚠ {subPayResults.unmatched.length} Unmatched Entries</div>
                      <div style={{ fontSize: 11, color: "#64748b", marginBottom: 8 }}>These deposits didn't contain a recognized billing reference code. Review manually below.</div>
                      {subPayResults.unmatched.map((u, i) => (
                        <div key={i} style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 10, padding: 10, marginBottom: 6, fontSize: 12, display: "flex", justifyContent: "space-between" }}>
                          <span>{u.date} · {u.reference || "(no reference)"}</span>
                          <strong>{fmt(u.amount)}</strong>
                        </div>
                      ))}
                    </div>
                  )}
                  <button onClick={() => { setSubPayRows([]); setSubPayFileName(""); setSubPayResults(null); }} style={{ marginTop: 12, background: "#0f172a", color: "#fff", border: "none", borderRadius: 9, padding: "9px 16px", fontWeight: 700, fontSize: 12, cursor: "pointer" }}>Done</button>
                </div>
              )}
            </div>

            {/* Per-school billing cards */}
            <div style={{ ...card }}>
              <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a", marginBottom: 4 }}>💳 School Subscriptions</div>
              <div style={{ fontSize: 12, color: "#64748b", marginBottom: 16 }}>Schools pay via MTN MoMo or bank transfer. Once you receive payment, click "Mark as Paid" to extend their subscription by a month or a term, depending on their billing cycle.</div>

              {allSchoolsList.map(s => {
                const info = s.subInfo;
                const statusColor = info.status === "Active" ? "#10b981" : info.status === "Grace Period" ? "#f59e0b" : "#ef4444";
                const statusBg = info.status === "Active" ? "#f0fdf4" : info.status === "Grace Period" ? "#fffbeb" : "#fef2f2";
                return (
                  <div key={s.id} style={{ border: `1px solid ${statusColor}40`, background: statusBg, borderRadius: 12, padding: 16, marginBottom: 12 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
                      <div style={{ flex: "1 1 250px" }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                          <span style={{ fontWeight: 800, fontSize: 15, color: "#0f172a" }}>{s.logo} {s.name}</span>
                          <Pill text={info.status} bg="#fff" col={statusColor} />
                          {s.isTrial && <Pill text="Free Trial" bg="#eff6ff" col="#2563eb" />}
                        </div>
                        <div style={{ display: "flex", gap: 16, fontSize: 12, color: "#374151", flexWrap: "wrap" }}>
                          <span>Plan: <strong>{s.plan}</strong> ({fmt(getBillingInfo(s.plan, s.billingCycle, s.customPrice).price)}{getBillingInfo(s.plan, s.billingCycle, s.customPrice).periodLabel}, {s.billingCycle === "term" ? "Per Term" : "Monthly"})</span>
                          <span>Students: <strong>{(allStudents[s.id] || []).length}</strong> / {PLANS[s.plan]?.maxStudents === Infinity ? "∞" : PLANS[s.plan]?.maxStudents}</span>
                          {s.nextBillingDate
                            ? <span>Next due: <strong>{fmtDate(s.nextBillingDate)}</strong></span>
                            : <span style={{ color: "#2563eb", fontWeight: 700 }}>⏳ Trial not yet started — awaiting first login</span>}
                          {s.lastPaymentDate && <span>Last paid: <strong>{fmtDate(s.lastPaymentDate)}</strong></span>}
                          <span>Ref: <strong style={{ fontFamily: "monospace" }}>{s.billingRef}</strong></span>
                        </div>
                        {s.customPrice != null && (
                          <div style={{ marginTop: 8, display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                            <Pill text={`💲 Custom price: ${fmt(s.customPrice)}${getBillingInfo(s.plan, s.billingCycle).periodLabel} (standard is ${fmt(getBillingInfo(s.plan, s.billingCycle).price)})`} bg="#f5f3ff" col="#7c3aed" />
                            {s.customPriceNote && <span style={{ fontSize: 11, color: "#6d28d9", fontStyle: "italic" }}>"{s.customPriceNote}"</span>}
                          </div>
                        )}
                        {info.status === "Grace Period" && (
                          <div style={{ marginTop: 8, fontSize: 12, color: "#92400e", fontWeight: 600 }}>
                            ⚠ {info.daysOverdue} day(s) overdue · {info.daysRemaining} day(s) left before suspension
                          </div>
                        )}
                        {info.status === "Suspended" && (
                          <div style={{ marginTop: 8, fontSize: 12, color: "#b91c1c", fontWeight: 600 }}>
                            🔒 {info.daysOverdue} days overdue · Account in read-only mode
                          </div>
                        )}
                      </div>
                      <div style={{ display: "flex", flexDirection: "column", gap: 8, alignItems: "flex-end" }}>
                        <select value={s.plan} onChange={e => changeSchoolPlan(s.id, e.target.value)}
                          style={{ padding: "6px 10px", borderRadius: 7, border: "1px solid #e2e8f0", fontSize: 12, fontWeight: 700, cursor: "pointer", background: "#fff" }}>
                          {Object.keys(PLANS).map(p => { const b = getBillingInfo(p, s.billingCycle); return <option key={p} value={p}>{p} — {fmt(b.price)}{b.periodLabel}</option>; })}
                        </select>
                        <button onClick={() => markSubscriptionPaid(s.id)} style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 9, padding: "8px 16px", fontWeight: 800, fontSize: 12, cursor: "pointer" }}>
                          ✓ Mark as Paid (+{getBillingInfo(s.plan, s.billingCycle, s.customPrice).cycleDays} days)
                        </button>
                        {s.customPrice != null ? (
                          <div style={{ display: "flex", gap: 6 }}>
                            <button onClick={() => { setShowCustomPriceEdit(s); setCustomPriceForm({ price: String(s.customPrice), note: s.customPriceNote || "" }); }}
                              style={{ background: "#f5f3ff", color: "#7c3aed", border: "1px solid #ddd6fe", borderRadius: 8, padding: "6px 11px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>✏ Edit Price</button>
                            <button onClick={() => setConfirmDialog({
                              title: "Remove Custom Price",
                              message: `Remove ${s.name}'s custom price of ${fmt(s.customPrice)} and revert to the standard ${s.plan} price of ${fmt(getBillingInfo(s.plan, s.billingCycle).price)}${getBillingInfo(s.plan, s.billingCycle).periodLabel}?`,
                              danger: false,
                              onConfirm: () => clearCustomPrice(s.id),
                            })}
                              style={{ background: "#fef2f2", color: "#dc2626", border: "1px solid #fca5a5", borderRadius: 8, padding: "6px 11px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>✕ Remove</button>
                          </div>
                        ) : (
                          <button onClick={() => { setShowCustomPriceEdit(s); setCustomPriceForm({ price: "", note: "" }); }}
                            style={{ background: "#fff", color: "#64748b", border: "1px solid #e2e8f0", borderRadius: 8, padding: "6px 11px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>💲 Set Custom Price</button>
                        )}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Plans reference */}
            <div style={{ ...grid(3, 1), gap: 14, marginTop: 20 }}>
              {Object.values(PLANS).map(p => (
                <div key={p.name} style={{ ...card }}>
                  <div style={{ fontWeight: 800, fontSize: 15, color: "#0f172a" }}>{p.name}</div>
                  <div style={{ fontSize: 20, fontWeight: 800, color: "#8b5cf6", marginTop: 4 }}>{fmt(p.price)}<span style={{ fontSize: 11, color: "#94a3b8", fontWeight: 600 }}>/month</span></div>
                  <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 2, marginBottom: 10 }}>Up to {p.maxStudents === Infinity ? "unlimited" : p.maxStudents} students</div>
                  {p.features.map(f => <div key={f} style={{ fontSize: 12, color: "#374151", padding: "3px 0" }}>✓ {f}</div>)}
                </div>
              ))}
            </div>
          </div>
          )}

          {superAdminTab === "settings" && (
          <div>
            {/* Change Super Admin Password */}
            <div style={{ ...card, marginBottom: 20 }}>
              <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a", marginBottom: 4 }}>🔐 Super Admin Login</div>
              <div style={{ fontSize: 12, color: "#64748b", marginBottom: 20 }}>Change your super admin username and/or password</div>

              <div style={{ marginBottom: 14 }}>
                <label style={lbl}>New Username (optional)</label>
                <input value={superPwForm.newUsername} onChange={e => { setSuperPwForm(p => ({ ...p, newUsername: e.target.value })); setSuperPwError(""); }}
                  placeholder={superAdminCreds.username} style={inp} />
              </div>
              <div style={{ marginBottom: 14 }}>
                <label style={lbl}>Current Password</label>
                <input type="password" value={superPwForm.currentPw} onChange={e => { setSuperPwForm(p => ({ ...p, currentPw: e.target.value })); setSuperPwError(""); }} style={inp} />
              </div>
              <div style={{ marginBottom: 14 }}>
                <label style={lbl}>New Password</label>
                <input type="password" value={superPwForm.newPw} onChange={e => { setSuperPwForm(p => ({ ...p, newPw: e.target.value })); setSuperPwError(""); }} style={inp} />
              </div>
              <div style={{ marginBottom: 14 }}>
                <label style={lbl}>Confirm New Password</label>
                <input type="password" value={superPwForm.confirmPw} onChange={e => { setSuperPwForm(p => ({ ...p, confirmPw: e.target.value })); setSuperPwError(""); }} style={inp} />
              </div>
              {superPwError && <div style={{ color: "#dc2626", fontSize: 12, fontWeight: 600, marginBottom: 12 }}>{superPwError}</div>}
              <button onClick={() => {
                if (superPwForm.currentPw !== superAdminCreds.password) return setSuperPwError("Current password is incorrect.");
                if (!superPwForm.newPw || superPwForm.newPw.length < 6) return setSuperPwError("New password must be at least 6 characters.");
                if (superPwForm.newPw !== superPwForm.confirmPw) return setSuperPwError("New passwords do not match.");
                setSuperAdminCreds(prev => ({ ...prev, password: superPwForm.newPw, username: superPwForm.newUsername.trim() || prev.username }));
                setSuperPwForm({ currentPw: "", newPw: "", confirmPw: "", newUsername: "" });
                notify("✓ Super admin login updated");
              }} style={{ background: "#0f172a", color: "#fff", border: "none", borderRadius: 9, padding: "10px 20px", fontWeight: 800, fontSize: 13, cursor: "pointer" }}>
                Save Changes
              </button>
            </div>

            {/* Payment Account Details */}
            <div style={{ ...card, marginBottom: 20 }}>
              <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a", marginBottom: 4 }}>💰 Subscription Payment Details</div>
              <div style={{ fontSize: 12, color: "#64748b", marginBottom: 16 }}>
                These MTN MoMo and bank account details are shown to every school when they renew their subscription (Resubscribe page & Settings). Update them here if your payment accounts change.
              </div>

              <div style={{ ...grid(2, 1), gap: 16 }}>
                <div>
                  <div style={{ fontSize: 12, fontWeight: 800, color: "#374151", marginBottom: 10 }}>📱 MTN MoMo</div>
                  <div style={{ marginBottom: 12 }}>
                    <label style={lbl}>MoMo Number</label>
                    <input value={(platformPayForm ?? platformPayInfo).momoNumber} onChange={e => setPlatformPayForm(p => ({ ...(p ?? platformPayInfo), momoNumber: e.target.value }))} style={inp} />
                  </div>
                  <div>
                    <label style={lbl}>Account Name</label>
                    <input value={(platformPayForm ?? platformPayInfo).momoName} onChange={e => setPlatformPayForm(p => ({ ...(p ?? platformPayInfo), momoName: e.target.value }))} style={inp} />
                  </div>
                </div>
                <div>
                  <div style={{ fontSize: 12, fontWeight: 800, color: "#374151", marginBottom: 10 }}>🏦 Bank Transfer</div>
                  <div style={{ marginBottom: 12 }}>
                    <label style={lbl}>Bank Name</label>
                    <input value={(platformPayForm ?? platformPayInfo).bankName} onChange={e => setPlatformPayForm(p => ({ ...(p ?? platformPayInfo), bankName: e.target.value }))} style={inp} />
                  </div>
                  <div>
                    <label style={lbl}>Account Number</label>
                    <input value={(platformPayForm ?? platformPayInfo).bankAccount} onChange={e => setPlatformPayForm(p => ({ ...(p ?? platformPayInfo), bankAccount: e.target.value }))} style={inp} />
                  </div>
                </div>
              </div>

              <button onClick={() => {
                if (platformPayForm) setPlatformPayInfo(platformPayForm);
                setPlatformPayForm(null);
                notify("✓ Payment details updated — schools will see the new details immediately");
              }} style={{ marginTop: 16, background: "#8b5cf6", color: "#fff", border: "none", borderRadius: 9, padding: "10px 20px", fontWeight: 800, fontSize: 13, cursor: "pointer" }}>
                Save Payment Details
              </button>
            </div>

            {/* Notification Email */}
            <div style={{ ...card }}>
              <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a", marginBottom: 4 }}>📧 Payment Notice Email Alerts</div>
              <div style={{ fontSize: 12, color: "#64748b", marginBottom: 16 }}>
                When a school submits "I've Sent Payment", you'll get an email alert here in addition to the in-app notification on the Billing tab.
              </div>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <input value={notifyEmailInput} onChange={e => setNotifyEmailInput(e.target.value)}
                  placeholder={superAdminCreds.notifyEmail || "you@example.com"} style={{ ...inp, flex: "1 1 240px" }} />
                <button onClick={() => {
                  if (!notifyEmailInput.trim() || !notifyEmailInput.includes("@")) return notify("Enter a valid email address", "err");
                  setSuperAdminCreds(prev => ({ ...prev, notifyEmail: notifyEmailInput.trim() }));
                  notify("✓ Notification email saved");
                }} style={{ background: "#8b5cf6", color: "#fff", border: "none", borderRadius: 9, padding: "10px 20px", fontWeight: 800, fontSize: 13, cursor: "pointer" }}>
                  Save Email
                </button>
              </div>
              {superAdminCreds.notifyEmail && (
                <div style={{ marginTop: 10, fontSize: 12, color: "#10b981", fontWeight: 600 }}>✓ Currently sending alerts to: {superAdminCreds.notifyEmail}</div>
              )}
              <div style={{ marginTop: 14, padding: "10px 14px", background: "#f8fafc", borderRadius: 9, fontSize: 11, color: "#94a3b8", lineHeight: 1.6 }}>
                Note: this prototype displays in-app notifications. Sending real emails requires connecting an email service (e.g. SendGrid, Resend) to the backend — let your developer know once you're ready to deploy that.
              </div>
            </div>

            {/* Activity Log */}
            <div style={{ ...card, marginTop: 20 }}>
              <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a", marginBottom: 4 }}>📜 Activity Log</div>
              <div style={{ fontSize: 12, color: "#64748b", marginBottom: 16 }}>
                A record of approvals, plan changes, payment confirmations, and other administrative actions taken in this session.
              </div>
              {activityLog.length === 0 ? (
                <div style={{ textAlign: "center", padding: 24, color: "#94a3b8", fontSize: 13 }}>No activity recorded yet this session.</div>
              ) : (
                <div style={{ maxHeight: 320, overflowY: "auto" }}>
                  {activityLog.map(a => (
                    <div key={a.id} style={{ display: "flex", gap: 12, padding: "10px 0", borderBottom: "1px solid #f1f5f9" }}>
                      <div style={{ fontSize: 11, color: "#94a3b8", whiteSpace: "nowrap", minWidth: 110 }}>{fmtDateTime(a.at)}</div>
                      <div>
                        <span style={{ fontSize: 12, fontWeight: 700, color: "#374151" }}>{a.action}</span>
                        <div style={{ fontSize: 12, color: "#64748b", marginTop: 1 }}>{a.detail}</div>
                      </div>
                    </div>
                  ))}
                </div>
              )}
              <div style={{ marginTop: 12, padding: "10px 14px", background: "#f8fafc", borderRadius: 9, fontSize: 11, color: "#94a3b8", lineHeight: 1.6 }}>
                Note: this log resets when the page is reloaded. Once the real backend is connected, this becomes a persistent audit trail stored in the database.
              </div>
            </div>
          </div>
          )}
        </div>

        {/* ════════ CUSTOM PRICE MODAL ════════ */}
        {showCustomPriceEdit && (() => {
          const sch = showCustomPriceEdit;
          const standard = getBillingInfo(sch.plan, sch.billingCycle);
          return (
          <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 300, padding: 16 }}>
            <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 420, maxWidth: 420, boxShadow: "0 24px 60px rgba(0,0,0,0.25)" }}>
              <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>💲 Custom Price for {sch.name}</div>
              <div style={{ fontSize: 12, color: "#64748b", marginBottom: 18 }}>
                This only changes what <strong>{sch.name}</strong> pays. Their plan ({sch.plan}) and all its features stay exactly the same, and no other school is affected.
              </div>

              <div style={{ background: "#f8fafc", borderRadius: 10, padding: 12, marginBottom: 16, fontSize: 12, color: "#374151" }}>
                Standard {sch.plan} price ({sch.billingCycle === "term" ? "Per Term" : "Monthly"}): <strong>{fmt(standard.price)}{standard.periodLabel}</strong>
              </div>

              <div style={{ marginBottom: 13 }}>
                <label style={lbl}>Custom Price (UGX{standard.periodLabel})</label>
                <input type="number" value={customPriceForm.price} onChange={e => setCustomPriceForm(p => ({ ...p, price: e.target.value }))}
                  placeholder={`e.g. ${Math.round(standard.price * 0.8)}`} style={inp} />
              </div>
              <div style={{ marginBottom: 20 }}>
                <label style={lbl}>Reason / Note (for your records)</label>
                <input value={customPriceForm.note} onChange={e => setCustomPriceForm(p => ({ ...p, note: e.target.value }))}
                  placeholder='e.g. "Negotiated rate, approved by [name]"' style={inp} />
              </div>

              <div style={{ display: "flex", gap: 10 }}>
                <button onClick={() => setShowCustomPriceEdit(null)} style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
                <button onClick={() => { setCustomPrice(sch.id, customPriceForm.price, customPriceForm.note); setShowCustomPriceEdit(null); }}
                  style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#7c3aed", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Save Custom Price</button>
              </div>
            </div>
          </div>
          );
        })()}
      </div>
    );
  }

  // ════════════════════════ PARENT PORTAL ════════════════════════
  if (currentUser.role === "parent") {
    const myChildren = (currentUser.childIds || [currentUser.studentId])
      .map(id => students.find(s => s.id === id))
      .filter(Boolean);
    const activeChildId = currentUser.activeChildId || currentUser.studentId;
    const myStudent = myChildren.find(s => s.id === activeChildId) || myChildren[0] || termStudents[0];
    const myBalance = getBalance(myStudent, currentTerm);
    const myPaid = myBalance.paidThisTerm;
    const myFee = myBalance.totalDue;
    const myBal = myBalance.balance;
    const myStatus = myBalance.status;
    const myPayments = myStudent.payments.filter(p => p.term === currentTerm);

    return (
      <div style={{ minHeight: "100vh", background: "#f4f6fb", fontFamily: "'Plus Jakarta Sans',sans-serif" }}>
        <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet" />
        <div style={{ background: "#0f172a", padding: "16px 24px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 24 }}>🏫</span>
            <div>
              <div style={{ color: "#f1f5f9", fontWeight: 800, fontSize: 14 }}>{school.name}</div>
              <div style={{ color: "#64748b", fontSize: 11 }}>Parent Portal · {currentTerm}</div>
            </div>
          </div>
          <button onClick={logout} style={{ background: "#1e293b", color: "#94a3b8", border: "none", borderRadius: 8, padding: "7px 14px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>Logout</button>
        </div>

        <div style={{ maxWidth: 700, margin: "32px auto", padding: "0 20px" }}>

          {/* Child switcher — only shown if parent has multiple children */}
          {myChildren.length > 1 && (
            <div style={{ ...card, marginBottom: 16, padding: "14px 16px" }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b", textTransform: "uppercase", marginBottom: 10 }}>
                👨‍👩‍👧‍👦 You have {myChildren.length} children at this school — select one to view
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                {myChildren.map(child => {
                  const isActive = child.id === myStudent.id;
                  const childBal = getBalance(child, currentTerm).balance;
                  return (
                    <button key={child.id}
                      onClick={() => setCurrentUser(prev => ({ ...prev, activeChildId: child.id }))}
                      style={{
                        display: "flex", alignItems: "center", gap: 8, padding: "8px 14px", borderRadius: 10,
                        border: `2px solid ${isActive ? "#f59e0b" : "#e2e8f0"}`,
                        background: isActive ? "#fffbeb" : "#fff", cursor: "pointer"
                      }}>
                      {child.photo
                        ? <img src={child.photo} alt={child.name} style={{ width: 28, height: 28, borderRadius: 8, objectFit: "cover" }} />
                        : <span style={{ fontSize: 18 }}>{child.gender === "F" ? "👩" : "👨"}</span>}
                      <div style={{ textAlign: "left" }}>
                        <div style={{ fontSize: 12, fontWeight: 700, color: isActive ? "#92400e" : "#0f172a" }}>{child.name}</div>
                        <div style={{ fontSize: 10, color: "#94a3b8" }}>{child.class} · {childBal > 0 ? `Owes ${fmt(childBal)}` : "Fully paid"}</div>
                      </div>
                    </button>
                  );
                })}
              </div>
            </div>
          )}

          <div style={{ ...card, marginBottom: 20, display: "flex", flexWrap: "wrap", gap: 16, alignItems: "center" }}>
            <div style={{ width: 60, height: 60, borderRadius: 16, background: "#eff6ff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 32, overflow: "hidden", flexShrink: 0 }}>
              {myStudent.photo
                ? <img src={myStudent.photo} alt={myStudent.name} style={{ width: "100%", height: "100%", objectFit: "cover" }} />
                : (myStudent.gender === "F" ? "👩" : "👨")}
            </div>
            <div style={{ flex: "1 1 200px" }}>
              <div style={{ fontSize: 20, fontWeight: 800, color: "#0f172a" }}>{myStudent.name}</div>
              <div style={{ fontSize: 13, color: "#64748b" }}>Class {myStudent.class} · Parent: {myStudent.parent} · {myStudent.phone}</div>
              <div style={{ marginTop: 6 }}><Badge status={myStatus} /></div>
            </div>
            <div style={{ textAlign: isMobile ? "left" : "right" }}>
              <div style={{ fontSize: 11, color: "#94a3b8", fontWeight: 600 }}>BALANCE DUE</div>
              <div style={{ fontSize: 26, fontWeight: 800, color: myBal > 0 ? "#dc2626" : "#15803d" }}>{fmt(myBal)}</div>
            </div>
          </div>

          <div style={{ ...grid(3, 1), gap: 14, marginBottom: 20 }}>
            {[
              { label: "Term Fee", value: fmt(myBalance.termFee), color: "#3b82f6" },
              { label: "Arrears", value: fmt(myBalance.arrears), color: "#f59e0b" },
              { label: "Total Due", value: fmt(myFee), color: "#8b5cf6" },
            ].map((c, i) => (
              <div key={i} style={{ ...card, padding: "14px 16px", borderTop: `3px solid ${c.color}` }}>
                <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase" }}>{c.label}</div>
                <div style={{ fontSize: 18, fontWeight: 800, color: "#0f172a", marginTop: 6 }}>{c.value}</div>
              </div>
            ))}
          </div>

          <div style={{ ...card }}>
            <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a", marginBottom: 16 }}>📋 Payment History — {currentTerm}</div>
            {myPayments.length === 0
              ? <div style={{ textAlign: "center", padding: 30, color: "#94a3b8" }}>No payments recorded this term</div>
              : myPayments.map((p, i) => (
                <div key={p.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 0", borderBottom: i < myPayments.length - 1 ? "1px solid #f1f5f9" : "none" }}>
                  <div>
                    <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 4 }}>
                      <span style={{ fontFamily: "monospace", fontSize: 12, color: "#3b82f6", fontWeight: 700, background: "#eff6ff", padding: "1px 7px", borderRadius: 5 }}>{p.id}</span>
                      <span style={{ fontSize: 12, color: "#374151" }}>{METHOD_ICON[p.method]} {p.method}</span>
                    </div>
                    <div style={{ fontSize: 12, color: "#94a3b8", marginTop: 4 }}>📅 {fmtDate(p.date)} · Received by {p.receivedBy}</div>
                  </div>
                  <div style={{ fontSize: 18, fontWeight: 800, color: "#15803d" }}>{fmt(p.amount)}</div>
                </div>
              ))}
            <div style={{ marginTop: 16, padding: "12px 16px", background: "#f0fdf4", borderRadius: 10, display: "flex", justifyContent: "space-between" }}>
              <span style={{ fontSize: 13, fontWeight: 600, color: "#15803d" }}>Total Paid This Term</span>
              <span style={{ fontSize: 16, fontWeight: 800, color: "#15803d" }}>{fmt(myPaid)}</span>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ════════════════════════ ADMIN APP ════════════════════════════
  // Everyday items — what a bursar uses day to day, shown first and un-grouped.
  const NAV_MAIN = [
    { id: "dashboard", icon: "◼", label: "Dashboard" },
    ...(subInfo.status !== "Active" ? [{ id: "resubscribe", icon: "💰", label: "Resubscribe", highlight: true }] : []),
    { id: "students", icon: "◉", label: "Students" },
    { id: "payments", icon: "◈", label: "Payments" },
    { id: "expenses", icon: "◇", label: "Expenses" },
    { id: "staff", icon: "👷", label: "Staff & Wages" },
    { id: "reports", icon: "▤", label: "Reports" },
    { id: "alumni", icon: "🎓", label: "Alumni / Leavers" },
  ];
  // Setup / admin items — touched occasionally (once a term or less), grouped
  // separately under a "Settings & Setup" divider so the everyday list above
  // isn't cluttered with low-frequency items.
  const NAV_SETUP = [
    { id: "fees", icon: "◎", label: "Fee Structure" },
    { id: "bank", icon: "🏦", label: "Bank Reconciliation" },
    { id: "settings", icon: "⚙", label: "Settings" },
  ];


  return (
    <div style={{ fontFamily: "'Plus Jakarta Sans',sans-serif", background: "#f4f6fb", minHeight: "100vh", display: "flex" }}>
      <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet" />

      {/* Toast */}
      {toast && (
        <div style={{ position: "fixed", top: 20, right: 20, zIndex: 9999, padding: "12px 20px", borderRadius: 12, fontWeight: 600, fontSize: 13, boxShadow: "0 8px 24px rgba(0,0,0,0.15)", background: toast.type === "err" ? "#fef2f2" : "#f0fdf4", color: toast.type === "err" ? "#b91c1c" : "#15803d", border: `1px solid ${toast.type === "err" ? "#fca5a5" : "#86efac"}` }}>
          {toast.type === "err" ? "✗ " : "✓ "}{toast.msg}
        </div>
      )}

      {/* ── Mobile Top Bar ── */}
      {isMobile && (
        <div style={{ position: "fixed", top: 0, left: 0, right: 0, height: 56, background: "#0f172a", display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 14px", zIndex: 60, boxShadow: "0 2px 8px rgba(0,0,0,0.2)" }}>
          <button onClick={() => setMobileSidebarOpen(true)} style={{ background: "none", border: "none", color: "#f1f5f9", fontSize: 22, cursor: "pointer", padding: 6 }}>
            ☰
          </button>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 18 }}>🏫</span>
            <div style={{ color: "#f1f5f9", fontWeight: 800, fontSize: 13 }}>FeeTrack UG</div>
          </div>
          <div style={{ width: 34 }} />
        </div>
      )}

      {/* ── Mobile Sidebar Overlay Backdrop ── */}
      {isMobile && mobileSidebarOpen && (
        <div onClick={() => setMobileSidebarOpen(false)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 55 }} />
      )}

      {/* ── Sidebar ── */}
      <aside style={{
        width: 224, background: "#0f172a", minHeight: "100vh", position: "fixed", top: 0, left: 0, bottom: 0,
        display: "flex", flexDirection: "column", zIndex: 56,
        transform: isMobile ? (mobileSidebarOpen ? "translateX(0)" : "translateX(-100%)") : "none",
        transition: "transform 0.25s ease",
      }}>
        <div style={{ padding: "22px 18px 14px" }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{ width: 36, height: 36, borderRadius: 10, background: "linear-gradient(135deg,#f59e0b,#ef4444)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>🏫</div>
              <div>
                <div style={{ color: "#f1f5f9", fontWeight: 800, fontSize: 13 }}>FeeTrack UG</div>
                <div style={{ color: "#475569", fontSize: 10 }}>School Finance</div>
              </div>
            </div>
            {isMobile && (
              <button onClick={() => setMobileSidebarOpen(false)} style={{ background: "#1e293b", border: "none", color: "#94a3b8", fontSize: 16, width: 28, height: 28, borderRadius: 7, cursor: "pointer" }}>✕</button>
            )}
          </div>

          {/* School identity — fixed, not switchable (privacy) */}
          <div style={{ background: "#1e293b", borderRadius: 8, padding: "8px 10px", marginBottom: 4 }}>
            <div style={{ color: "#475569", fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5 }}>Your School</div>
            <div style={{ color: "#f59e0b", fontSize: 12, fontWeight: 700, marginTop: 2 }}>{school.logo} {school.name}</div>
            {subInfo.status !== "Active" && (
              <div style={{ marginTop: 4, fontSize: 10, fontWeight: 700, color: subInfo.frozen ? "#60a5fa" : subInfo.status === "Grace Period" ? "#fbbf24" : "#f87171" }}>
                {subInfo.frozen ? "📨 Payment under review" : subInfo.status === "Grace Period" ? `⚠ Renew within ${subInfo.daysRemaining}d` : "🔒 Read-only — renew now"}
              </div>
            )}
          </div>

          {/* Term selector */}
          <div style={{ marginTop: 8 }}>
            <select value={currentTerm} onChange={e => setCurrentTerm(e.target.value)} style={{ width: "100%", background: "#1e293b", color: "#94a3b8", border: "1px solid #334155", borderRadius: 8, padding: "7px 10px", fontSize: 11, fontWeight: 600, outline: "none", cursor: "pointer" }}>
              {TERMS.map(t => <option key={t} value={t}>{t}</option>)}
            </select>
          </div>
        </div>

        <nav style={{ padding: "0 10px", flex: 1, overflowY: "auto" }}>
          <div style={{ color: "#334155", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, padding: "6px 8px 4px" }}>Menu</div>
          {NAV_MAIN.map(item => (
            <button key={item.id} onClick={() => { setTab(item.id); setSelectedStaffId(null); if (isMobile) setMobileSidebarOpen(false); }} style={{ width: "100%", display: "flex", alignItems: "center", gap: 9, padding: "9px 11px", borderRadius: 9, border: item.highlight && tab !== item.id ? "1px solid #f87171" : "none", cursor: "pointer", marginBottom: 2, textAlign: "left", fontSize: 13, fontWeight: tab === item.id ? 700 : (item.highlight ? 700 : 500), background: tab === item.id ? "#f59e0b" : item.highlight ? "#450a0a" : "transparent", color: tab === item.id ? "#0f172a" : item.highlight ? "#fca5a5" : "#64748b" }}>
              <span style={{ fontSize: 10 }}>{item.icon}</span>{item.label}
              {item.id === "sms" && smsLog.length > 0 && <span style={{ marginLeft: "auto", background: "#ef4444", color: "#fff", borderRadius: 99, fontSize: 9, fontWeight: 700, padding: "1px 6px" }}>{smsLog.length}</span>}
            </button>
          ))}

          <div style={{ color: "#334155", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, padding: "14px 8px 4px", marginTop: 8, borderTop: "1px solid #1e293b" }}>Settings & Setup</div>
          {NAV_SETUP.map(item => (
            <button key={item.id} onClick={() => { setTab(item.id); setSelectedStaffId(null); if (isMobile) setMobileSidebarOpen(false); }} style={{ width: "100%", display: "flex", alignItems: "center", gap: 9, padding: "9px 11px", borderRadius: 9, border: "none", cursor: "pointer", marginBottom: 2, textAlign: "left", fontSize: 13, fontWeight: tab === item.id ? 700 : 500, background: tab === item.id ? "#f59e0b" : "transparent", color: tab === item.id ? "#0f172a" : "#64748b" }}>
              <span style={{ fontSize: 10 }}>{item.icon}</span>{item.label}
            </button>
          ))}
        </nav>

        <div style={{ padding: "12px 18px", borderTop: "1px solid #1e293b" }}>
          <div style={{ color: "#475569", fontSize: 10, marginBottom: 2 }}>Logged in as</div>
          <div style={{ color: "#94a3b8", fontWeight: 600, fontSize: 11, marginBottom: 8 }}>{adminCreds.username} · {school.name.split(" ")[0]}</div>
          <button onClick={logout} style={{ width: "100%", background: "#450a0a", color: "#f87171", border: "none", borderRadius: 7, padding: "7px 0", fontSize: 11, cursor: "pointer", fontWeight: 700 }}>🚪 Logout</button>
        </div>
      </aside>

      {/* ── Main Content ── */}
      <main style={{ marginLeft: isMobile ? 0 : 224, marginTop: isMobile ? 56 : 0, flex: 1, padding: isMobile ? 14 : 26, minWidth: 0, width: "100%", boxSizing: "border-box" }}>

        {/* ════ SUBSCRIPTION STATUS BANNERS ════ */}
        {subInfo.status === "Grace Period" && (
          <div style={{ background: subInfo.frozen ? "#eff6ff" : "#fffbeb", border: `2px solid ${subInfo.frozen ? "#bfdbfe" : "#fde68a"}`, borderRadius: 12, padding: "14px 18px", marginBottom: 18, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
            <div>
              {subInfo.frozen ? (
                <>
                  <div style={{ fontWeight: 800, fontSize: 14, color: "#1d4ed8" }}>📨 Payment Under Review — Access Protected</div>
                  <div style={{ fontSize: 12, color: "#1e40af", marginTop: 2 }}>
                    We've received your payment notice and your account won't be suspended while we confirm it. This usually takes a few hours. Reference: <strong style={{ fontFamily: "monospace" }}>{school.billingRef}</strong>
                  </div>
                </>
              ) : (
                <>
                  <div style={{ fontWeight: 800, fontSize: 14, color: "#92400e" }}>⚠ Subscription Payment Overdue</div>
                  <div style={{ fontSize: 12, color: "#78350f", marginTop: 2 }}>
                    Your payment was due on {fmtDate(school.nextBillingDate)} ({subInfo.daysOverdue} day{subInfo.daysOverdue !== 1 ? "s" : ""} ago). You have <strong>{subInfo.daysRemaining} day{subInfo.daysRemaining !== 1 ? "s" : ""}</strong> left before your account becomes read-only. Reference: <strong style={{ fontFamily: "monospace" }}>{school.billingRef}</strong>
                  </div>
                </>
              )}
            </div>
            {!subInfo.frozen && (
              <button onClick={() => setTab("resubscribe")} style={{ background: "#92400e", color: "#fff", border: "none", borderRadius: 9, padding: "9px 16px", fontSize: 12, fontWeight: 800, cursor: "pointer" }}>
                💰 Pay {fmt(getBillingInfo(school.plan, school.billingCycle, school.customPrice).price)} Now →
              </button>
            )}
          </div>
        )}
        {subInfo.status === "Suspended" && (
          <div style={{ background: "#fef2f2", border: "2px solid #fca5a5", borderRadius: 12, padding: "14px 18px", marginBottom: 18, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
            <div>
              <div style={{ fontWeight: 800, fontSize: 14, color: "#b91c1c" }}>🔒 Account in Read-Only Mode — Payment Required</div>
              <div style={{ fontSize: 12, color: "#7f1d1d", marginTop: 2 }}>
                Your subscription was due on {fmtDate(school.nextBillingDate)} ({subInfo.daysOverdue} days overdue). You can view all your data, but cannot record payments, edit records, or make changes until payment is made. Reference: <strong style={{ fontFamily: "monospace" }}>{school.billingRef}</strong>
              </div>
            </div>
            <button onClick={() => setTab("resubscribe")} style={{ background: "#b91c1c", color: "#fff", border: "none", borderRadius: 9, padding: "9px 16px", fontSize: 12, fontWeight: 800, cursor: "pointer" }}>
              💰 Pay {fmt(getBillingInfo(school.plan, school.billingCycle, school.customPrice).price)} Now →
            </button>
          </div>
        )}

        {/* ════════ DASHBOARD ════════ */}
        {tab === "dashboard" && (
          <div>
            {/* ════ FIRST-TIME SETUP WIZARD ════ */}
            {!school.setupComplete && !setupDismissed && (
              <div style={{ ...card, marginBottom: 22, border: "2px solid #8b5cf6", background: "linear-gradient(135deg,#f5f3ff,#fff)" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16, flexWrap: "wrap", gap: 8 }}>
                  <div>
                    <div style={{ fontSize: 18, fontWeight: 800, color: "#6d28d9" }}>👋 Welcome to FeeTrack UG, {school.name}!</div>
                    <div style={{ fontSize: 13, color: "#64748b", marginTop: 4 }}>Let's get your account set up in 3 quick steps. This takes about 10 minutes.</div>
                  </div>
                  <button onClick={() => setSetupDismissed(true)} style={{ background: "#f1f5f9", color: "#64748b", border: "none", borderRadius: 7, padding: "6px 12px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>Skip for now</button>
                </div>

                {/* Step indicators */}
                <div style={{ display: "flex", gap: 8, marginBottom: 20, flexWrap: "wrap" }}>
                  {[
                    { n: 1, label: "Set Your Fees", icon: "💰" },
                    { n: 2, label: "Add Requirements", icon: "📋" },
                    { n: 3, label: "Enrol Students", icon: "👨‍🎓" },
                  ].map(s => (
                    <button key={s.n} onClick={() => setSetupStep(s.n)}
                      style={{ flex: "1 1 30%", padding: "12px 14px", borderRadius: 10, border: `2px solid ${setupStep === s.n ? "#8b5cf6" : "#e2e8f0"}`, background: setupStep === s.n ? "#f5f3ff" : "#fff", cursor: "pointer", textAlign: "left" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                        <span style={{ width: 22, height: 22, borderRadius: "50%", background: setupStep === s.n ? "#8b5cf6" : "#f1f5f9", color: setupStep === s.n ? "#fff" : "#94a3b8", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 800, flexShrink: 0 }}>{s.n}</span>
                        <span style={{ fontSize: 16 }}>{s.icon}</span>
                      </div>
                      <div style={{ fontSize: 12, fontWeight: 700, color: setupStep === s.n ? "#6d28d9" : "#374151" }}>{s.label}</div>
                    </button>
                  ))}
                </div>

                {/* Step content */}
                {setupStep === 1 && (
                  <div style={{ background: "#fff", borderRadius: 12, padding: 18, border: "1px solid #e8edf3" }}>
                    <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a", marginBottom: 6 }}>💰 Step 1: Set Your School Fees</div>
                    <div style={{ fontSize: 13, color: "#64748b", marginBottom: 14, lineHeight: 1.6 }}>
                      We've loaded sample fee amounts for Day Scholar and Boarder students (S1–S6). Go to <strong>Fee Structure</strong> and click any amount to update it to match your actual fees. You can also add or remove fee items (e.g. Exam Fee, Transport).
                    </div>
                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                      <button onClick={() => setTab("fees")} style={{ background: "#8b5cf6", color: "#fff", border: "none", borderRadius: 9, padding: "10px 18px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>Go to Fee Structure →</button>
                      <button onClick={() => setSetupStep(2)} style={{ background: "#f1f5f9", color: "#374151", border: "none", borderRadius: 9, padding: "10px 18px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>Next: Requirements →</button>
                    </div>
                  </div>
                )}
                {setupStep === 2 && (
                  <div style={{ background: "#fff", borderRadius: 12, padding: 18, border: "1px solid #e8edf3" }}>
                    <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a", marginBottom: 6 }}>📋 Step 2: Add School Requirements</div>
                    <div style={{ fontSize: 13, color: "#64748b", marginBottom: 14, lineHeight: 1.6 }}>
                      Set up items students must bring (uniform, books, mattress for boarders, etc.) and their costs. Mandatory items are automatically added to each student's total fee. Found at the bottom of the <strong>Fee Structure</strong> tab.
                    </div>
                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                      <button onClick={() => setTab("fees")} style={{ background: "#8b5cf6", color: "#fff", border: "none", borderRadius: 9, padding: "10px 18px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>Go to Requirements →</button>
                      <button onClick={() => setSetupStep(3)} style={{ background: "#f1f5f9", color: "#374151", border: "none", borderRadius: 9, padding: "10px 18px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>Next: Enrol Students →</button>
                    </div>
                  </div>
                )}
                {setupStep === 3 && (
                  <div style={{ background: "#fff", borderRadius: 12, padding: 18, border: "1px solid #e8edf3" }}>
                    <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a", marginBottom: 6 }}>👨‍🎓 Step 3: Enrol Your Students</div>
                    <div style={{ fontSize: 13, color: "#64748b", marginBottom: 14, lineHeight: 1.6 }}>
                      Two ways to add students: enrol them <strong>one by one</strong> using the "+ Enrol Student" button, or use <strong>"📋 Bulk Enrol Students"</strong> to upload your entire register from an Excel/CSV file — much faster if you already have a list.
                    </div>
                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                      <button onClick={() => setTab("students")} style={{ background: "#8b5cf6", color: "#fff", border: "none", borderRadius: 9, padding: "10px 18px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>Go to Students →</button>
                      <button onClick={markSetupComplete} style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 9, padding: "10px 18px", fontWeight: 800, fontSize: 13, cursor: "pointer" }}>✓ Finish Setup</button>
                    </div>
                  </div>
                )}
              </div>
            )}

            <div style={{ display: "flex", flexDirection: isMobile ? "column" : "row", justifyContent: "space-between", alignItems: isMobile ? "stretch" : "flex-start", gap: 12, marginBottom: 22 }}>
              <div>
                <div style={{ fontSize: isMobile ? 18 : 22, fontWeight: 800, color: "#0f172a" }}>Dashboard — {school.name}</div>
                <div style={{ color: "#64748b", fontSize: 13, marginTop: 2 }}>{currentTerm} · {termStudents.length} students</div>
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <button onClick={() => { setShowRollover(true); setRolloverStep(1); setRolloverDnr({}); setRolloverTerm(""); }} style={{ background: "#f1f5f9", color: "#374151", border: "1px solid #e2e8f0", borderRadius: 9, padding: "9px 14px", fontWeight: 700, fontSize: 12, cursor: "pointer", flex: isMobile ? "1 1 45%" : "none" }}>🔄 Term Rollover</button>
                <button onClick={() => setShowPromotion(true)} style={{ background: "#8b5cf6", color: "#fff", border: "none", borderRadius: 9, padding: "9px 14px", fontWeight: 700, fontSize: 12, cursor: "pointer", flex: isMobile ? "1 1 45%" : "none" }}>🎓 New Academic Year</button>
                <button onClick={exportExcel} style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 9, padding: "9px 14px", fontWeight: 700, fontSize: 12, cursor: "pointer", flex: isMobile ? "1 1 45%" : "none" }}>📊 Export Excel</button>
                <button onClick={() => setTab("students")} style={{ background: "#f59e0b", color: "#0f172a", border: "none", borderRadius: 9, padding: "9px 14px", fontWeight: 700, fontSize: 12, cursor: "pointer", flex: isMobile ? "1 1 45%" : "none" }}>+ Record Payment</button>
              </div>
            </div>

            {/* KPIs */}
            <div style={{ ...grid(7, 2), gap: 12, marginBottom: 18 }}>
              {[
                { label: "Students", value: termStudents.length, sub: `${paidCount} paid`, icon: "👨‍🎓", accent: "#3b82f6" },
                { label: "Expected", value: fmtShort(totalExpected), sub: fmt(totalExpected), icon: "🎯", accent: "#8b5cf6" },
                { label: "Collected", value: fmtShort(totalCollected), sub: `${collRate}% rate`, icon: "✅", accent: "#10b981" },
                { label: "Outstanding (Active)", value: fmtShort(totalBalance), sub: `${unpaidCount + partialCount} students`, icon: "⏳", accent: "#f59e0b" },
                { label: "Alumni Debt", value: fmtShort(alumniDebtTotal), sub: alumniDebtorCount > 0 ? `${alumniDebtorCount} former students` : "None owing", icon: "🎓", accent: "#7c3aed" },
                { label: "Expenses", value: fmtShort(totalExpensesAmt), sub: currentTerm, icon: "💸", accent: "#ef4444" },
                { label: "Net Surplus", value: fmtShort(netSurplus), sub: netSurplus >= 0 ? "Positive" : "Deficit", icon: "📈", accent: netSurplus >= 0 ? "#10b981" : "#ef4444" },
              ].map((k, i) => (
                <div key={i} style={{ ...card, borderTop: `3px solid ${k.accent}`, padding: "14px 16px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <div style={{ fontSize: 9, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", letterSpacing: 0.4 }}>{k.label}</div>
                    <span style={{ fontSize: 16 }}>{k.icon}</span>
                  </div>
                  <div style={{ fontSize: 20, fontWeight: 800, color: "#0f172a", margin: "6px 0 2px" }}>{k.value}</div>
                  <div style={{ fontSize: 10, color: "#94a3b8" }}>{k.sub}</div>
                </div>
              ))}
            </div>

            <div style={{ ...grid(2, 1), gap: 16, marginBottom: 16 }}>
              <div style={{ ...card }}>
                <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a", marginBottom: 14 }}>Collection Rate</div>
                <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
                  <div style={{ position: "relative", width: 96, height: 96, flexShrink: 0 }}>
                    <svg viewBox="0 0 36 36" style={{ width: 96, height: 96, transform: "rotate(-90deg)" }}>
                      <circle cx="18" cy="18" r="15.9" fill="none" stroke="#f1f5f9" strokeWidth="3" />
                      <circle cx="18" cy="18" r="15.9" fill="none" stroke={collRate >= 80 ? "#10b981" : collRate >= 50 ? "#f59e0b" : "#ef4444"} strokeWidth="3" strokeDasharray={`${collRate} ${100 - collRate}`} strokeLinecap="round" />
                    </svg>
                    <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
                      <div style={{ fontSize: 18, fontWeight: 800, color: "#0f172a" }}>{collRate}%</div>
                    </div>
                  </div>
                  <div style={{ flex: 1 }}>
                    {[{ label: "Fully Paid", count: paidCount, color: "#10b981" }, { label: "Partial", count: partialCount, color: "#f59e0b" }, { label: "Unpaid", count: unpaidCount, color: "#ef4444" }].map((r, i) => (
                      <div key={i} style={{ marginBottom: 9 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 3 }}>
                          <span style={{ color: "#374151" }}>{r.label}</span>
                          <span style={{ fontWeight: 700, color: r.color }}>{r.count}</span>
                        </div>
                        <div style={{ height: 5, background: "#f1f5f9", borderRadius: 99 }}>
                          <div style={{ height: "100%", width: `${Math.round((r.count / termStudents.length) * 100)}%`, background: r.color, borderRadius: 99 }} />
                        </div>
                      </div>
                    ))}
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginTop: 10 }}>
                      <div style={{ background: "#f0fdf4", borderRadius: 8, padding: "7px 10px" }}>
                        <div style={{ fontSize: 9, color: "#16a34a", fontWeight: 700, textTransform: "uppercase" }}>Collected</div>
                        <div style={{ fontSize: 13, fontWeight: 800, color: "#15803d" }}>{fmt(totalCollected)}</div>
                      </div>
                      <div style={{ background: "#fef2f2", borderRadius: 8, padding: "7px 10px" }}>
                        <div style={{ fontSize: 9, color: "#dc2626", fontWeight: 700, textTransform: "uppercase" }}>Balance</div>
                        <div style={{ fontSize: 13, fontWeight: 800, color: "#b91c1c" }}>{fmt(totalBalance)}</div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              {/* Expense breakdown pie */}
              <div style={{ ...card }}>
                <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a", marginBottom: 4 }}>Expense Breakdown</div>
                <div style={{ fontSize: 11, color: "#94a3b8", marginBottom: 10 }}>Total: {fmt(totalExpensesAmt)}</div>
                {expByCategory.length === 0
                  ? <div style={{ textAlign: "center", padding: 30, color: "#94a3b8" }}>No expenses recorded</div>
                  : <div style={{ display: "flex", alignItems: "center" }}>
                    <ResponsiveContainer width={140} height={140}>
                      <PieChart>
                        <Pie data={expByCategory} cx="50%" cy="50%" innerRadius={35} outerRadius={60} dataKey="value" paddingAngle={2}>
                          {expByCategory.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                        </Pie>
                      </PieChart>
                    </ResponsiveContainer>
                    <div style={{ flex: 1, paddingLeft: 12 }}>
                      {expByCategory.map((e, i) => (
                        <div key={e.name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                            <div style={{ width: 8, height: 8, borderRadius: 2, background: PIE_COLORS[i % PIE_COLORS.length], flexShrink: 0 }} />
                            <span style={{ fontSize: 11, color: "#374151" }}>{e.name}</span>
                          </div>
                          <span style={{ fontSize: 11, fontWeight: 700, color: "#0f172a" }}>{fmtShort(e.value)}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                }
              </div>
            </div>

            {/* Class table + recent payments */}
            <div style={{ ...grid(2, 1), gridTemplateColumns: isMobile ? "1fr" : "1.2fr 1fr", gap: 16, marginBottom: 16 }}>
              <div style={{ ...card, overflowX: "auto" }}>
                <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a", marginBottom: 12 }}>Collection by Class</div>
                <table style={{ width: "100%", borderCollapse: "collapse", minWidth: isMobile ? 480 : "auto" }}>
                  <thead><tr>{["Class", "Students", "Expected", "Collected", "Balance", "Rate"].map(h => <th key={h} style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", textAlign: "left", padding: "0 6px 8px" }}>{h}</th>)}</tr></thead>
                  <tbody>
                    {classStats.map(g => (
                      <tr key={g.class} style={{ borderTop: "1px solid #f8fafc" }}>
                        <td style={{ padding: "8px 6px" }}><Pill text={g.class} /></td>
                        <td style={{ padding: "8px 6px", fontSize: 12, fontWeight: 600 }}>{g.count}</td>
                        <td style={{ padding: "8px 6px", fontSize: 11, color: "#64748b" }}>{fmtShort(g.expected)}</td>
                        <td style={{ padding: "8px 6px", fontSize: 11, color: "#15803d", fontWeight: 600 }}>{fmtShort(g.collected)}</td>
                        <td style={{ padding: "8px 6px", fontSize: 11, color: "#b91c1c", fontWeight: 600 }}>{fmtShort(g.expected - g.collected)}</td>
                        <td style={{ padding: "8px 6px" }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                            <div style={{ flex: 1, height: 5, background: "#f1f5f9", borderRadius: 99, minWidth: 40 }}>
                              <div style={{ height: "100%", width: `${g.rate}%`, background: g.rate >= 80 ? "#10b981" : g.rate >= 50 ? "#f59e0b" : "#ef4444", borderRadius: 99 }} />
                            </div>
                            <span style={{ fontSize: 10, fontWeight: 700 }}>{g.rate}%</span>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div style={{ ...card }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                  <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a" }}>🕐 Recent Payments</div>
                  <button onClick={() => setTab("payments")} style={{ fontSize: 11, color: "#3b82f6", background: "none", border: "none", cursor: "pointer", fontWeight: 600 }}>View All →</button>
                </div>
                {allStudentsAndAlumni.flatMap(s =>
                  (s.payments || []).map(p => ({ ...p, studentName: s.name }))
                ).sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 6).map((p, i) => (
                  <div key={p.id + i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: i < 5 ? "1px solid #f1f5f9" : "none" }}>
                    <div>
                      <div style={{ fontSize: 12, fontWeight: 600, color: "#0f172a" }}>{p.studentName}</div>
                      <div style={{ fontSize: 11, color: "#94a3b8" }}>{fmtDate(p.date)} · {METHOD_ICON[p.method]} {p.method}</div>
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <div style={{ fontSize: 13, fontWeight: 800, color: "#15803d" }}>{fmt(p.amount)}</div>
                      <div style={{ fontSize: 10, fontFamily: "monospace", color: "#3b82f6" }}>{p.id}</div>
                    </div>
                  </div>
                ))}
                {allPayments.length === 0 && <div style={{ textAlign: "center", padding: 20, color: "#94a3b8" }}>No payments yet</div>}
              </div>
            </div>
          </div>
        )}

        {/* ════════ RESUBSCRIBE ════════ */}
        {tab === "resubscribe" && (
          <div>
            <div style={{ marginBottom: 20 }}>
              <div style={{ fontSize: isMobile ? 18 : 22, fontWeight: 800, color: "#0f172a" }}>💰 Resubscribe to FeeTrack UG</div>
              <div style={{ color: "#64748b", fontSize: 13, marginTop: 2 }}>
                {subInfo.status === "Suspended"
                  ? "Your account is currently read-only. Renew now to restore full access immediately."
                  : `Your subscription is overdue. Renew within ${subInfo.daysRemaining} day(s) to avoid losing access.`}
              </div>
            </div>

            {/* Current status */}
            <div style={{ ...card, marginBottom: 20, border: `2px solid ${subInfo.status === "Suspended" ? "#fca5a5" : "#fde68a"}`, background: subInfo.status === "Suspended" ? "#fef2f2" : "#fffbeb" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
                <div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                    <span style={{ fontWeight: 800, fontSize: 16, color: "#0f172a" }}>{school.plan} Plan</span>
                    <Pill text={subInfo.status} bg="#fff" col={subInfo.status === "Grace Period" ? "#f59e0b" : "#ef4444"} />
                  </div>
                  <div style={{ fontSize: 13, color: "#374151" }}>Was due: <strong>{fmtDate(school.nextBillingDate)}</strong> · {subInfo.daysOverdue} day(s) overdue</div>
                  {subInfo.status === "Suspended" && <div style={{ fontSize: 12, color: "#b91c1c", fontWeight: 700, marginTop: 6 }}>🔒 Payments, enrolments, edits and bulk imports are disabled until you renew</div>}
                  {subInfo.status === "Grace Period" && <div style={{ fontSize: 12, color: "#92400e", fontWeight: 700, marginTop: 6 }}>⚠ {subInfo.daysRemaining} day(s) left in your grace period</div>}
                </div>
              </div>
            </div>

            {/* Payment options */}
            <div style={{ ...card, border: "2px solid #8b5cf6", background: "linear-gradient(135deg,#f5f3ff,#fff)", marginBottom: 20 }}>
              <div style={{ fontWeight: 800, fontSize: 16, color: "#6d28d9", marginBottom: 4 }}>💳 Renew Your Subscription</div>
              <div style={{ fontSize: 13, color: "#64748b", marginBottom: 16 }}>Pay the amount below using your unique reference code — your account reactivates automatically, usually within minutes.</div>

              <div style={{ ...grid(2, 1), gap: 14, marginBottom: 14 }}>
                <div style={{ background: "#fff", borderRadius: 10, padding: 14 }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", marginBottom: 8 }}>📱 MTN MoMo</div>
                  <div style={{ fontSize: 14, color: "#374151", marginBottom: 4 }}>Send to: <strong style={{ fontFamily: "monospace" }}>{platformPayInfo.momoNumber}</strong></div>
                  <div style={{ fontSize: 14, color: "#374151" }}>Name: <strong>{platformPayInfo.momoName}</strong></div>
                </div>
                <div style={{ background: "#fff", borderRadius: 10, padding: 14 }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", marginBottom: 8 }}>🏦 Bank Transfer</div>
                  <div style={{ fontSize: 14, color: "#374151", marginBottom: 4 }}>Bank: <strong>{platformPayInfo.bankName}</strong></div>
                  <div style={{ fontSize: 14, color: "#374151" }}>Account: <strong style={{ fontFamily: "monospace" }}>{platformPayInfo.bankAccount}</strong></div>
                </div>
              </div>

              <div style={{ background: "#fff", borderRadius: 10, padding: 16, marginBottom: 14, textAlign: "center" }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", marginBottom: 6 }}>Amount Due — {school.plan} Plan ({school.billingCycle === "term" ? "Per Term" : "30 days"})</div>
                <div style={{ fontSize: 32, fontWeight: 800, color: "#0f172a" }}>{fmt(getBillingInfo(school.plan, school.billingCycle, school.customPrice).price)}</div>
              </div>

              <div style={{ background: "#fef3c7", border: "2px solid #fde68a", borderRadius: 10, padding: 16, textAlign: "center" }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: "#92400e", textTransform: "uppercase", marginBottom: 6 }}>⚠ Required — Your Reference Code</div>
                <div style={{ fontSize: 24, fontWeight: 800, color: "#0f172a", fontFamily: "monospace", letterSpacing: 2, marginBottom: 8 }}>{school.billingRef}</div>
                <div style={{ fontSize: 12, color: "#78350f", lineHeight: 1.6 }}>
                  Type this code into the payment reference/narration field. Our system scans for it automatically and restores your access — no need to contact anyone.
                </div>
              </div>

              {/* Step-by-step */}
              <div style={{ marginTop: 14, background: "#fff", borderRadius: 10, padding: 14 }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 8 }}>How it works:</div>
                <div style={{ fontSize: 12, color: "#64748b", lineHeight: 1.9 }}>
                  <strong>1.</strong> Open your MTN MoMo or banking app on your phone<br />
                  <strong>2.</strong> Send {fmt(getBillingInfo(school.plan, school.billingCycle, school.customPrice).price)} to the number/account above, with reference <strong style={{ fontFamily: "monospace" }}>{school.billingRef}</strong><br />
                  <strong>3.</strong> Come back here and click <strong>"✓ I've Sent Payment"</strong> below so our team knows to expect it<br />
                  <strong>4.</strong> Your access restores automatically — usually within minutes, sometimes up to a few hours
                </div>
              </div>

              {!isReadOnly || subInfo.status !== "Active" ? (
                <button onClick={() => setShowPaymentConfirm(true)} style={{ width: "100%", marginTop: 14, padding: 14, borderRadius: 10, border: "none", background: "#10b981", color: "#fff", fontSize: 15, fontWeight: 800, cursor: "pointer" }}>
                  ✓ I've Sent Payment
                </button>
              ) : null}

              {/* Show pending confirmations from this school */}
              {paymentNotices.filter(n => n.schoolId === activeSchoolId && n.status === "pending").length > 0 && (
                <div style={{ marginTop: 14, background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 10, padding: 14 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#1d4ed8", marginBottom: 6 }}>📨 Payment Notice Sent</div>
                  {paymentNotices.filter(n => n.schoolId === activeSchoolId && n.status === "pending").map(n => (
                    <div key={n.id} style={{ fontSize: 12, color: "#374151", marginBottom: 4 }}>
                      {fmtDate(n.date)} · {n.method} · {fmt(n.amount)} — we'll confirm shortly. Your account will reactivate automatically once payment is verified.
                    </div>
                  ))}
                </div>
              )}

              {/* Show "not received" notices so school knows to follow up */}
              {paymentNotices.filter(n => n.schoolId === activeSchoolId && n.status === "not_found").length > 0 && (
                <div style={{ marginTop: 14, background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 10, padding: 14 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#b91c1c", marginBottom: 6 }}>⚠ Payment Not Found</div>
                  {paymentNotices.filter(n => n.schoolId === activeSchoolId && n.status === "not_found").map(n => (
                    <div key={n.id} style={{ fontSize: 12, color: "#374151", marginBottom: 4 }}>
                      We couldn't find your payment of {fmt(n.amount)} sent on {fmtDate(n.date)} (Ref: {n.billingRef}) in our records. Please double-check the transaction went through with the correct reference code, or contact us if you believe this is an error.
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Want a different plan? */}
            <div style={{ ...card }}>
              <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a", marginBottom: 4 }}>Want to switch plans while renewing?</div>
              <div style={{ fontSize: 12, color: "#64748b", marginBottom: 14 }}>Pay the exact monthly or per-term price of a different plan (with your reference code) and your plan switches automatically — no approval needed.</div>
              <div style={{ ...grid(3, 1), gap: 10 }}>
                {Object.values(PLANS).map(p => {
                  const monthly = getBillingInfo(p.name, "monthly");
                  const term = getBillingInfo(p.name, "term");
                  return (
                  <div key={p.name} style={{ border: `2px solid ${p.name === school.plan ? "#8b5cf6" : "#e2e8f0"}`, borderRadius: 10, padding: 12, background: p.name === school.plan ? "#f5f3ff" : "#fff" }}>
                    <div style={{ fontWeight: 800, fontSize: 13, color: "#0f172a" }}>{p.name}{p.name === school.plan && <span style={{ color: "#7c3aed", fontSize: 11, marginLeft: 6 }}>Current</span>}</div>
                    <div style={{ marginTop: 4 }}>
                      <div style={{ fontSize: 15, fontWeight: 800, color: "#7c3aed" }}>{fmt(monthly.price)}<span style={{ fontSize: 10, color: "#94a3b8" }}>{monthly.periodLabel}</span></div>
                      <div style={{ fontSize: 12, fontWeight: 700, color: "#15803d", marginTop: 1 }}>{fmt(term.price)}<span style={{ fontSize: 10, color: "#94a3b8" }}>{term.periodLabel}</span> <span style={{ fontSize: 9, color: "#94a3b8" }}>(save 10%)</span></div>
                    </div>
                    <div style={{ fontSize: 10, color: "#94a3b8", marginTop: 6, marginBottom: 8 }}>Up to {p.maxStudents === Infinity ? "unlimited" : p.maxStudents} students</div>
                    {p.features.map(f => <div key={f} style={{ fontSize: 10, color: "#64748b", padding: "1px 0" }}>✓ {f}</div>)}
                  </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}

        {/* ════════ STUDENTS ════════ */}
        {tab === "students" && (
          <div>
            <div style={{ display: "flex", flexDirection: isMobile ? "column" : "row", justifyContent: "space-between", alignItems: isMobile ? "stretch" : "center", gap: 12, marginBottom: 20 }}>
              <div>
                <div style={{ fontSize: isMobile ? 18 : 22, fontWeight: 800, color: "#0f172a" }}>Student Register</div>
                <div style={{ color: "#64748b", fontSize: 13 }}>Click a row to see payment history · {termStudents.length} students</div>
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <button onClick={() => setShowBalancesReport(true)} style={{ background: "#fff", color: "#374151", border: "1px solid #e2e8f0", borderRadius: 9, padding: "9px 16px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>🖨️ Print Balances</button>
                {!isReadOnly && (
                <>
                <button onClick={() => setShowBulkImport(true)} style={{ background: "#fff", color: "#374151", border: "1px solid #e2e8f0", borderRadius: 9, padding: "9px 16px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>📋 Bulk Enrol Students</button>
                <button onClick={() => { setShowBulkAlumni(true); setBulkAlumniSelected({}); setBulkAlumniReason("Transferred to another school"); setBulkAlumniMode("checklist"); setBulkAlumniExcelRows([]); setBulkAlumniExcelFileName(""); setBulkAlumniExcelDone(false); }} style={{ background: "#f5f3ff", color: "#7c3aed", border: "1px solid #ddd6fe", borderRadius: 9, padding: "9px 16px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>🎓 Bulk Move to Alumni</button>
                <button onClick={() => setShowAdd(true)} style={{ background: "#f59e0b", color: "#0f172a", border: "none", borderRadius: 9, padding: "9px 16px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>+ Enrol Student</button>
                </>
                )}
              </div>
            </div>

            <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 14 }}>
              <input value={search} onChange={e => setSearch(e.target.value)} placeholder="🔍 Search student name..."
                style={{ flex: isMobile ? "1 1 100%" : 2, padding: "10px 14px", borderRadius: 9, border: "1px solid #e2e8f0", fontSize: 13, outline: "none", background: "#fff", boxSizing: "border-box" }} />
              <select value={filterClass} onChange={e => setFilterClass(e.target.value)}
                style={{ flex: isMobile ? "1 1 30%" : 1, padding: "10px 12px", borderRadius: 9, border: "1px solid #e2e8f0", fontSize: 13, outline: "none", background: "#fff", cursor: "pointer" }}>
                <option value="All">All Classes</option>
                {schoolClasses.map(c => {
                  const streams = getClassStreams(c);
                  if (streams.length === 0) return <option key={c} value={c}>{c}</option>;
                  return (
                    <optgroup key={c} label={c}>
                      <option value={c}>All {c}</option>
                      {streams.map(s => (
                        <option key={c + s} value={classLabel(c, s)}>{classLabel(c, s)}</option>
                      ))}
                    </optgroup>
                  );
                })}
              </select>
              <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)}
                style={{ flex: isMobile ? "1 1 30%" : 1, padding: "10px 12px", borderRadius: 9, border: "1px solid #e2e8f0", fontSize: 13, outline: "none", background: "#fff", cursor: "pointer" }}>
                <option value="All">All Status</option>
                <option>Paid</option><option>Partial</option><option>Unpaid</option>
              </select>
              <select value={filterClass === "All" && filterStatus === "All" ? "All" : "All"}
                onChange={e => { setFilterClass("All"); setFilterStatus("All"); setSearch(""); }}
                style={{ flex: isMobile ? "1 1 30%" : "none", padding: "10px 12px", borderRadius: 9, border: "1px solid #e2e8f0", fontSize: 13, outline: "none", background: "#f1f5f9", cursor: "pointer", color: "#64748b", fontWeight: 600 }}>
                <option value="All">Clear Filters</option>
              </select>
            </div>

            <div style={{ ...card, padding: 0, overflow: "hidden" }}>
            <div style={{ overflowX: "auto" }}>
              <div style={{ minWidth: isMobile ? 760 : "auto" }}>
              <div style={{ display: "grid", gridTemplateColumns: "2fr 0.5fr 1.1fr 0.9fr 0.9fr 0.9fr 0.6fr 0.6fr", gap: 4, padding: "10px 16px", background: "#f8fafc", borderBottom: "1px solid #e8edf3" }}>
                {["Student", "Class", "Parent / Phone", "Term Fee", "Paid", "Balance", "Status", "Action"].map(h => (
                  <div key={h} style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", letterSpacing: 0.3 }}>{h}</div>
                ))}
              </div>
              {filtered.map((s, i) => {
                const rowBal = getBalance(s, currentTerm);
                const tf = rowBal.totalDue;
                const paid = rowBal.paidThisTerm;
                const bal = rowBal.balance;
                const st = rowBal.status;
                const isOpen = expandedId === s.id;
                return (
                  <div key={s.id}>
                    <div onClick={() => setExpandedId(isOpen ? null : s.id)}
                      style={{ display: "grid", gridTemplateColumns: "2fr 0.5fr 1.1fr 0.9fr 0.9fr 0.9fr 0.6fr 0.8fr", gap: 4, alignItems: "center", padding: "11px 16px", borderTop: i > 0 ? "1px solid #f1f5f9" : "none", background: isOpen ? "#f0f7ff" : i % 2 === 0 ? "#fff" : "#fafbfc", cursor: "pointer" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
                        {/* Student photo / avatar */}
                        <div style={{ position: "relative", flexShrink: 0 }}>
                          {s.photo
                            ? <img src={s.photo} alt={s.name} style={{ width: 38, height: 38, borderRadius: 10, objectFit: "cover", border: "2px solid #e2e8f0" }} />
                            : <div style={{ width: 38, height: 38, borderRadius: 10, background: s.gender === "F" ? "#fce7f3" : "#dbeafe", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20, border: "2px solid #e2e8f0" }}>
                                {s.gender === "F" ? "👩" : "👨"}
                              </div>
                          }
                          <button onClick={e => { e.stopPropagation(); setShowPhotoUpload(s); setPhotoUploadType("student"); setCameraActive(false); }}
                            title="Add / change photo"
                            style={{ position: "absolute", bottom: -4, right: -4, width: 16, height: 16, borderRadius: "50%", background: "#f59e0b", border: "2px solid #fff", cursor: "pointer", fontSize: 8, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, color: "#0f172a" }}>
                            📷
                          </button>
                        </div>
                        <div>
                          <div style={{ fontWeight: 600, fontSize: 13, color: "#0f172a" }}>{s.name}</div>
                          <div style={{ display: "flex", gap: 4, marginTop: 2, flexWrap: "wrap" }}>
                            <Pill text={s.category || "Day Scholar"} bg="#eff6ff" col="#2563eb" />
                            {s.bursary && <Pill text={`${s.bursary.type === "percent" ? s.bursary.value + "% off" : fmt(s.bursary.value) + " off"} · ${s.bursary.reason}`} bg="#fef3c7" col="#92400e" />}
                            {s.customFee && <Pill text="Custom Fee" bg="#f5f3ff" col="#7c3aed" />}
                            {s.arrears > 0 && <Pill text={`+${fmt(s.arrears)} arrears`} bg="#fef2f2" col="#dc2626" />}
                          </div>
                        </div>
                      </div>
                      <div><Pill text={classLabel(s.class, s.stream)} /></div>
                      <div>
                        <div style={{ fontSize: 12, color: "#374151", fontWeight: 500 }}>{s.parent}</div>
                        <div style={{ fontSize: 10, color: "#94a3b8", fontFamily: "monospace" }}>{s.phone}</div>
                      </div>
                      <div style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>{fmt(tf)}</div>
                      <div style={{ fontSize: 12, fontWeight: 700, color: "#15803d" }}>{fmt(paid)}</div>
                      <div style={{ fontSize: 12, fontWeight: 700, color: bal > 0 ? "#dc2626" : "#15803d" }}>{fmt(bal)}</div>
                      <div><Badge status={st} /></div>
                      <div style={{ display: "flex", gap: 4, alignItems: "center", flexWrap: "wrap" }}>
                        {st !== "Paid" && !isReadOnly && <button onClick={e => { e.stopPropagation(); setShowPay(s); }} style={{ background: "#0f172a", color: "#fff", border: "none", borderRadius: 7, padding: "4px 9px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>Pay</button>}
                        {!isReadOnly && <button onClick={e => { e.stopPropagation(); setShowEditStudent({ ...s }); }} style={{ background: "#fffbeb", color: "#92400e", border: "1px solid #fde68a", borderRadius: 7, padding: "4px 9px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>✏ Edit</button>}
                        <button onClick={e => { e.stopPropagation(); setShowFeeEdit(s); setFeeEditData({ mode: s.customFee ? "custom" : s.bursary ? "bursary" : "category", bursaryType: s.bursary?.type || "percent", bursaryValue: s.bursary?.value || "", bursaryReason: s.bursary?.reason || "", customFee: s.customFee || "" }); }} style={{ background: "#f1f5f9", color: "#374151", border: "1px solid #e2e8f0", borderRadius: 7, padding: "4px 9px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>⚙ Fee</button>
                        <button onClick={e => { e.stopPropagation(); setShowMoveAlumni(s); setMoveAlumniReason("Transferred to another school"); }} title="Move to Alumni (transfer/dropout)" style={{ background: "#f5f3ff", color: "#7c3aed", border: "1px solid #ddd6fe", borderRadius: 7, padding: "4px 9px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>🎓 Alumni</button>
                        <span style={{ color: "#94a3b8", fontSize: 13 }}>{isOpen ? "▲" : "▼"}</span>
                      </div>
                    </div>
                    {/* Expanded payment timeline */}
                    {isOpen && (
                      <div style={{ background: "#f0f7ff", borderTop: "1px solid #bfdbfe", borderBottom: "1px solid #bfdbfe", padding: "16px 24px" }}>
                        <div style={{ fontWeight: 700, fontSize: 13, color: "#1e3a8a", marginBottom: 12 }}>📋 Payment History — {s.name} ({s.class}) · {currentTerm}</div>
                        {s.payments.filter(p => p.term === currentTerm).length === 0
                          ? <div style={{ color: "#ef4444", fontSize: 13, fontWeight: 600 }}>No payments this term. Balance due: {fmt(tf)}</div>
                          : (
                            <div style={{ position: "relative", paddingLeft: 26 }}>
                              <div style={{ position: "absolute", left: 8, top: 0, bottom: 0, width: 2, background: "#bfdbfe" }} />
                              {s.payments.filter(p => p.term === currentTerm).map((p, pi) => (
                                <div key={p.id} style={{ position: "relative", marginBottom: 12 }}>
                                  <div style={{ position: "absolute", left: -22, top: 4, width: 10, height: 10, borderRadius: "50%", background: "#3b82f6", border: "2px solid #fff" }} />
                                  <div style={{ background: "#fff", borderRadius: 9, padding: "10px 14px", border: "1px solid #bfdbfe", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                                    <div>
                                      <div style={{ display: "flex", gap: 7, alignItems: "center", marginBottom: 3 }}>
                                        <span style={{ fontFamily: "monospace", fontSize: 11, color: "#3b82f6", fontWeight: 700, background: "#eff6ff", padding: "1px 6px", borderRadius: 5 }}>{p.id}</span>
                                        <span style={{ fontSize: 12, color: "#374151", fontWeight: 600 }}>{METHOD_ICON[p.method]} {p.method}</span>
                                        <span style={{ fontSize: 11, color: "#94a3b8" }}>· {p.receivedBy}</span>
                                      </div>
                                      <div style={{ fontSize: 11, color: "#64748b" }}>📅 {fmtDate(p.date)}</div>
                                    </div>
                                    <div style={{ textAlign: "right" }}>
                                      <div style={{ fontSize: 15, fontWeight: 800, color: "#15803d" }}>{fmt(p.amount)}</div>
                                      <div style={{ display: "flex", gap: 5, marginTop: 4 }}>
                                        <button onClick={() => setShowReceipt({ payment: p, student: s, school, newBalance: tf - s.payments.filter(px => px.term === currentTerm && px.id <= p.id).reduce((a, px) => a + px.amount, 0) })} style={{ fontSize: 10, color: "#3b82f6", background: "none", border: "1px solid #bfdbfe", borderRadius: 5, padding: "2px 7px", cursor: "pointer" }}>Receipt</button>
                                        <button onClick={() => { setShowEditPayment({ student: s, payment: p }); setEditPayAmt(String(p.amount)); }} style={{ fontSize: 10, color: "#f59e0b", background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 5, padding: "2px 7px", cursor: "pointer", fontWeight: 700 }}>✏ Edit</button>
                                        <button onClick={() => handleDeletePayment(s, p.id)} style={{ fontSize: 10, color: "#ef4444", background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 5, padding: "2px 7px", cursor: "pointer", fontWeight: 700 }}>🗑 Delete</button>
                                      </div>
                                    </div>
                                  </div>
                                </div>
                              ))}
                            </div>
                          )}
                        <div style={{ display: "flex", gap: 10, marginTop: 10, padding: "9px 13px", background: "#fff", borderRadius: 9, border: "1px solid #bfdbfe", fontSize: 12, color: "#64748b" }}>
                          <span>Paid: <strong style={{ color: "#15803d" }}>{fmt(paid)}</strong></span>
                          <span>|</span>
                          <span>Balance: <strong style={{ color: bal > 0 ? "#dc2626" : "#15803d" }}>{fmt(bal)}</strong></span>
                          <span>|</span>
                          <span>Payments: <strong style={{ color: "#0f172a" }}>{s.payments.filter(p => p.term === currentTerm).length}</strong></span>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
              {filtered.length === 0 && <div style={{ textAlign: "center", padding: 40, color: "#94a3b8" }}>No students match filters</div>}
              </div>
              </div>
            </div>
          </div>
        )}

        {/* ════════ PAYMENTS ════════ */}
        {tab === "payments" && (
          <div>
            <div style={{ display: "flex", flexDirection: isMobile ? "column" : "row", justifyContent: "space-between", alignItems: isMobile ? "stretch" : "center", gap: 12, marginBottom: 20 }}>
              <div>
                <div style={{ fontSize: isMobile ? 18 : 22, fontWeight: 800, color: "#0f172a" }}>Payment Ledger</div>
                <div style={{ color: "#64748b", fontSize: 13 }}>
                  {allPayments.length} transactions · {paymentsTermFilter === "all" ? "All Terms" : paymentsTermFilter === "current" ? currentTerm : paymentsTermFilter}
                </div>
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <input value={paymentsSearch || ""} onChange={e => setPaymentsSearch(e.target.value)}
                  placeholder="🔍 Search student name..."
                  style={{ padding: "9px 14px", borderRadius: 9, border: "1px solid #e2e8f0", fontSize: 13, outline: "none", background: "#fff", minWidth: 180 }} />
                <select value={paymentsTermFilter} onChange={e => setPaymentsTermFilter(e.target.value)}
                  style={{ padding: "9px 12px", borderRadius: 9, border: "1px solid #e2e8f0", fontSize: 12, fontWeight: 600, outline: "none", background: "#fff", cursor: "pointer" }}>
                  <option value="current">Current Term ({currentTerm})</option>
                  <option value="all">All Terms (All Years)</option>
                  {termsWithPayments.filter(t => t !== currentTerm).map(t => (
                    <option key={t} value={t}>{t}</option>
                  ))}
                </select>
                {!isReadOnly && <button onClick={() => { setShowBulkPayments(true); setBulkPayRows([]); setBulkPayFileName(""); setBulkPayImportDone(false); }} style={{ background: "#ecfdf5", color: "#15803d", border: "1px solid #86efac", borderRadius: 9, padding: "9px 16px", fontWeight: 700, fontSize: 12, cursor: "pointer" }}>💰 Bulk Record Payments</button>}
                <button onClick={exportExcel} style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 9, padding: "9px 16px", fontWeight: 700, fontSize: 12, cursor: "pointer" }}>📊 Export Excel</button>
              </div>
            </div>

            <div style={{ ...grid(4, 2), gap: 12, marginBottom: 18 }}>
              {[
                { label: "Transactions", value: allPayments.length, color: "#3b82f6" },
                { label: "Cash", value: allPayments.filter(p => p.method === "Cash").length, color: "#10b981" },
                { label: "MTN MoMo", value: allPayments.filter(p => p.method === "MTN MoMo").length, color: "#f59e0b" },
                { label: "Bank / Airtel", value: allPayments.filter(p => ["Bank", "Airtel Money"].includes(p.method)).length, color: "#8b5cf6" },
              ].map((c, i) => (
                <div key={i} style={{ ...card, borderTop: `3px solid ${c.color}`, padding: "12px 16px" }}>
                  <div style={{ fontSize: 9, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", marginBottom: 4 }}>{c.label}</div>
                  <div style={{ fontSize: 26, fontWeight: 800, color: "#0f172a" }}>{c.value}</div>
                </div>
              ))}
            </div>

            <div style={{ ...card, padding: 0, overflow: "hidden" }}>
            <div style={{ overflowX: "auto" }}>
            <div style={{ minWidth: isMobile ? (paymentsTermFilter === "all" ? 860 : 760) : "auto" }}>
              <div style={{ display: "grid", gridTemplateColumns: paymentsTermFilter === "all" ? "1fr 1.3fr 0.5fr 1fr 1fr 1fr 1fr 0.9fr 0.6fr" : "1fr 1.3fr 0.5fr 1fr 1fr 1fr 0.9fr 0.6fr", gap: 4, padding: "10px 16px", background: "#f8fafc", borderBottom: "1px solid #e8edf3" }}>
                {(paymentsTermFilter === "all"
                  ? ["Receipt #", "Student", "Class", "Term", "Date", "Amount", "Method", "Received By", "Action"]
                  : ["Receipt #", "Student", "Class", "Date", "Amount", "Method", "Received By", "Action"]
                ).map(h => (
                  <div key={h} style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", letterSpacing: 0.3 }}>{h}</div>
                ))}
              </div>
              {allPayments.filter(p => !paymentsSearch || p.studentName.toLowerCase().includes(paymentsSearch.toLowerCase())).map((p, i) => (
                <div key={p.id + i} style={{ display: "grid", gridTemplateColumns: paymentsTermFilter === "all" ? "1fr 1.3fr 0.5fr 1fr 1fr 1fr 1fr 0.9fr 0.6fr" : "1fr 1.3fr 0.5fr 1fr 1fr 1fr 0.9fr 0.6fr", gap: 4, alignItems: "center", padding: "10px 16px", borderTop: "1px solid #f1f5f9", background: i % 2 === 0 ? "#fff" : "#fafbfc" }}>
                  <div style={{ fontFamily: "monospace", fontSize: 11, color: "#3b82f6", fontWeight: 700 }}>{p.id}</div>
                  <div style={{ fontSize: 12, fontWeight: 600, color: "#0f172a" }}>
                    {p.studentName}
                    {p.isAlumni && <span style={{ marginLeft: 6, fontSize: 9, fontWeight: 700, color: "#7c3aed", background: "#f5f3ff", borderRadius: 5, padding: "1px 5px" }}>ALUMNI</span>}
                  </div>
                  <div><Pill text={p.studentClass} /></div>
                  {paymentsTermFilter === "all" && <div style={{ fontSize: 11, color: "#64748b", fontWeight: 600 }}>{p.term}</div>}
                  <div style={{ fontSize: 12, color: "#374151", fontWeight: 600 }}>{fmtDate(p.date)}</div>
                  <div style={{ fontSize: 13, fontWeight: 800, color: "#15803d" }}>{fmt(p.amount)}</div>
                  <div style={{ fontSize: 12, color: "#374151" }}>{METHOD_ICON[p.method] || "💰"} {p.method}</div>
                  <div style={{ fontSize: 12, color: "#64748b" }}>{p.receivedBy}</div>
                  <button onClick={() => {
                    const student = allStudentsAndAlumni.find(s => s.payments && s.payments.some(px => px.id === p.id));
                    let newBalance = 0;
                    if (student) {
                      const due = getBalance(student, p.term).totalDue;
                      const paidUpToThis = student.payments
                        .filter(px => px.term === p.term && px.id <= p.id)
                        .reduce((a, px) => a + px.amount, 0);
                      newBalance = Math.max(0, due - paidUpToThis);
                    }
                    setShowReceipt({ payment: p, student, school, newBalance });
                  }} style={{ fontSize: 11, color: "#3b82f6", background: "#eff6ff", border: "none", borderRadius: 6, padding: "4px 9px", cursor: "pointer", fontWeight: 600 }}>Receipt</button>
                </div>
              ))}
              {allPayments.length === 0 && <div style={{ textAlign: "center", padding: 40, color: "#94a3b8" }}>No payments recorded</div>}
            </div>
            </div>
            </div>
          </div>
        )}

        {/* ════════ EXPENSES ════════ */}
        {tab === "expenses" && (
          <div>
            <div style={{ display: "flex", flexDirection: isMobile ? "column" : "row", justifyContent: "space-between", alignItems: isMobile ? "stretch" : "center", gap: 12, marginBottom: 20 }}>
              <div>
                <div style={{ fontSize: isMobile ? 18 : 22, fontWeight: 800, color: "#0f172a" }}>Expense Tracker</div>
                <div style={{ color: "#64748b", fontSize: 13 }}>{currentTerm} · Total: {fmt(totalExpensesAmt)}</div>
              </div>
              {!isReadOnly && <button onClick={() => setShowAddExp(true)} style={{ background: "#ef4444", color: "#fff", border: "none", borderRadius: 9, padding: "9px 16px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>+ Add Expense</button>}
            </div>

            <div style={{ ...grid(2, 1), gap: 16, marginBottom: 18 }}>
              <div style={{ ...card }}>
                <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a", marginBottom: 12 }}>Fee Income vs Expenses</div>
                <div style={{ display: "flex", gap: 12, marginBottom: 14 }}>
                  {[
                    { label: "Fee Income", value: totalCollected, color: "#10b981", bg: "#f0fdf4" },
                    { label: "Total Expenses", value: totalExpensesAmt, color: "#ef4444", bg: "#fef2f2" },
                    { label: "Net Surplus", value: netSurplus, color: netSurplus >= 0 ? "#3b82f6" : "#ef4444", bg: netSurplus >= 0 ? "#eff6ff" : "#fef2f2" },
                  ].map((c, i) => (
                    <div key={i} style={{ flex: 1, background: c.bg, borderRadius: 10, padding: "10px 12px" }}>
                      <div style={{ fontSize: 9, fontWeight: 700, color: c.color, textTransform: "uppercase", marginBottom: 4 }}>{c.label}</div>
                      <div style={{ fontSize: 15, fontWeight: 800, color: c.color }}>{fmt(c.value)}</div>
                    </div>
                  ))}
                </div>
                <div style={{ height: 6, background: "#f1f5f9", borderRadius: 99, overflow: "hidden" }}>
                  <div style={{ height: "100%", width: `${totalCollected > 0 ? Math.min(Math.round((totalExpensesAmt / totalCollected) * 100), 100) : 0}%`, background: "#ef4444", borderRadius: 99 }} />
                </div>
                <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 4 }}>Expenses = {totalCollected > 0 ? Math.round((totalExpensesAmt / totalCollected) * 100) : 0}% of income</div>
              </div>

              <div style={{ ...card }}>
                <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a", marginBottom: 12 }}>By Category</div>
                {expByCategory.length === 0
                  ? <div style={{ textAlign: "center", padding: 20, color: "#94a3b8" }}>No expenses yet</div>
                  : expByCategory.map((e, i) => (
                    <div key={e.name} style={{ marginBottom: 10 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 4 }}>
                        <span style={{ color: "#374151", fontWeight: 500 }}>{e.name}</span>
                        <span style={{ fontWeight: 700, color: "#0f172a" }}>{fmt(e.value)}</span>
                      </div>
                      <div style={{ height: 5, background: "#f1f5f9", borderRadius: 99 }}>
                        <div style={{ height: "100%", width: `${totalExpensesAmt > 0 ? Math.round((e.value / totalExpensesAmt) * 100) : 0}%`, background: PIE_COLORS[i % PIE_COLORS.length], borderRadius: 99 }} />
                      </div>
                    </div>
                  ))}
              </div>
            </div>

            <div style={{ ...card, padding: 0, overflow: "hidden" }}>
            <div style={{ overflowX: "auto" }}>
            <div style={{ minWidth: isMobile ? 640 : "auto" }}>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr 1fr 1fr 1fr 0.5fr", gap: 4, padding: "10px 16px", background: "#f8fafc", borderBottom: "1px solid #e8edf3" }}>
                {["Category", "Description", "Amount", "Date", "Paid By", ""].map(h => <div key={h} style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", letterSpacing: 0.3 }}>{h}</div>)}
              </div>
              {expenses.filter(e => e.term === currentTerm).map((e, i) => (
                <div key={e.id} style={{ display: "grid", gridTemplateColumns: "1fr 2fr 1fr 1fr 1fr 0.5fr", gap: 4, alignItems: "center", padding: "11px 16px", borderTop: "1px solid #f1f5f9", background: i % 2 === 0 ? "#fff" : "#fafbfc" }}>
                  <div><Pill text={e.category} bg="#fef3c7" col="#92400e" /></div>
                  <div style={{ fontSize: 13, color: "#374151" }}>{e.description}</div>
                  <div style={{ fontSize: 13, fontWeight: 800, color: "#ef4444" }}>{fmt(e.amount)}</div>
                  <div style={{ fontSize: 12, color: "#64748b" }}>{fmtDate(e.date)}</div>
                  <div style={{ fontSize: 12, color: "#64748b" }}>{e.paidBy}</div>
                  <div style={{ display: "flex", gap: 5 }}>
                    <button onClick={() => { setShowEditExpense(e); setEditExpAmt(String(e.amount)); }}
                      style={{ background: "#fffbeb", color: "#92400e", border: "1px solid #fde68a", borderRadius: 7, padding: "4px 9px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>✏</button>
                    <button onClick={() => setConfirmDialog({
                      title: "Delete Expense",
                      message: `Delete this ${e.category.toLowerCase()} expense of ${fmt(e.amount)}? This cannot be undone.`,
                      danger: true,
                      onConfirm: async () => {
                        const { error } = await supabase.from("expenses").delete().eq("id", e.id);
                        if (error) return notify(`Could not delete expense: ${error.message}`, "err");
                        setAllExpenses(prev => ({ ...prev, [activeSchoolId]: (prev[activeSchoolId] || []).filter(ex => ex.id !== e.id) }));
                        notify("Expense deleted");
                      },
                    })}
                      style={{ background: "#fef2f2", color: "#dc2626", border: "1px solid #fca5a5", borderRadius: 7, padding: "4px 9px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>🗑</button>
                  </div>
                </div>
              ))}
              {expenses.filter(e => e.term === currentTerm).length === 0 && <div style={{ textAlign: "center", padding: 40, color: "#94a3b8" }}>No expenses for {currentTerm}</div>}
            </div>
            </div>
            </div>
          </div>
        )}

        {/* ════════ STAFF & WAGES ════════ */}
        {tab === "staff" && (() => {
          const termStaffPayments = staffPayments.filter(p => p.term === currentTerm);
          const totalPaidThisTerm = termStaffPayments.reduce((a, p) => a + p.amount, 0);
          const matchesSearch = (s) => !staffSearch.trim() || s.name.toLowerCase().includes(staffSearch.trim().toLowerCase());
          const activeStaff = staff.filter(s => s.active && matchesSearch(s));
          const inactiveStaff = staff.filter(s => !s.active && matchesSearch(s));
          const paidThisTermByStaff = (staffId) => termStaffPayments.filter(p => p.staffId === staffId).reduce((a, p) => a + p.amount, 0);
          const lastPaidDate = (staffId) => {
            const ps = staffPayments.filter(p => p.staffId === staffId).sort((a, b) => new Date(b.date) - new Date(a.date));
            return ps.length > 0 ? ps[0].date : null;
          };

          // ── Per-worker detail view ──────────────────────────────────────
          const selectedStaff = selectedStaffId ? staff.find(s => s.id === selectedStaffId) : null;
          if (selectedStaff) {
            const allPaymentsForStaff = staffPayments.filter(p => p.staffId === selectedStaff.id).sort((a, b) => new Date(b.date) - new Date(a.date));
            const dailyPayments = allPaymentsForStaff.filter(p => p.payType === "daily");
            const monthlyPayments = allPaymentsForStaff.filter(p => p.payType === "monthly");
            const dailyTotal = dailyPayments.reduce((a, p) => a + p.amount, 0);
            const monthlyTotal = monthlyPayments.reduce((a, p) => a + p.amount, 0);
            const paymentRow = (p) => (
              <div key={p.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "11px 16px", borderTop: "1px solid #f1f5f9" }}>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 700, color: "#374151" }}>{p.periodLabel}</div>
                  <div style={{ fontSize: 11, color: "#94a3b8" }}>{fmtDate(p.date)} · {p.term}</div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ fontSize: 14, fontWeight: 800, color: "#15803d" }}>{fmt(p.amount)}</div>
                  {!isReadOnly && (
                    <button onClick={() => setConfirmDialog({
                      title: "Delete Staff Payment",
                      message: `Delete this payment of ${fmt(p.amount)} to ${p.staffName}? This will also remove the matching expense entry. This cannot be undone.`,
                      danger: true,
                      onConfirm: () => handleDeleteStaffPayment(p.id),
                    })}
                      style={{ background: "#fef2f2", color: "#dc2626", border: "1px solid #fca5a5", borderRadius: 7, padding: "4px 9px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>🗑</button>
                  )}
                </div>
              </div>
            );
            return (
              <div>
                <button onClick={() => setSelectedStaffId(null)} style={{ background: "none", border: "none", color: "#64748b", fontSize: 13, fontWeight: 600, cursor: "pointer", marginBottom: 16, padding: 0 }}>← Back to Staff & Wages</button>

                <div style={{ ...card, marginBottom: 18, display: "flex", flexDirection: isMobile ? "column" : "row", justifyContent: "space-between", alignItems: isMobile ? "stretch" : "center", gap: 14 }}>
                  <div>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
                      <div style={{ position: "relative", flexShrink: 0 }}>
                        {selectedStaff.photo
                          ? <img src={selectedStaff.photo} alt={selectedStaff.name} style={{ width: 52, height: 52, borderRadius: 13, objectFit: "cover", border: "2px solid #e2e8f0" }} />
                          : <div style={{ width: 52, height: 52, borderRadius: 13, background: "#e0e7ff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 26, border: "2px solid #e2e8f0" }}>👷</div>
                        }
                        {!isReadOnly && (
                          <button onClick={() => { setShowPhotoUpload(selectedStaff); setPhotoUploadType("staff"); setCameraActive(false); }}
                            title="Add / change photo"
                            style={{ position: "absolute", bottom: -4, right: -4, width: 19, height: 19, borderRadius: "50%", background: "#f59e0b", border: "2px solid #fff", cursor: "pointer", fontSize: 9, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, color: "#0f172a" }}>
                            📷
                          </button>
                        )}
                      </div>
                      <div>
                        <div style={{ fontWeight: 800, fontSize: 18, color: "#0f172a" }}>{selectedStaff.name}</div>
                        <Pill text={selectedStaff.role} bg="#eff6ff" col="#2563eb" />
                      </div>
                    </div>
                    <div style={{ display: "flex", gap: 14, fontSize: 12, color: "#64748b", flexWrap: "wrap", marginTop: 8 }}>
                      <span>Usual rate: <strong>{fmt(selectedStaff.defaultRate)}</strong> / {selectedStaff.defaultRateType === "daily" ? "day" : "month"}</span>
                      {selectedStaff.phone && <span>📞 {selectedStaff.phone}</span>}
                      {!selectedStaff.active && <Pill text="Inactive" bg="#fef2f2" col="#dc2626" />}
                    </div>
                  </div>
                  {!isReadOnly && (
                    <div style={{ display: "flex", gap: 6 }}>
                      <button onClick={() => { setShowPayStaff(selectedStaff); setPayStaffForm({ amount: String(selectedStaff.defaultRate || ""), payType: selectedStaff.defaultRateType || "daily", periodLabel: "" }); }}
                        style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 8, padding: "9px 16px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>💵 Record Payment</button>
                      <button onClick={() => setShowEditStaff({ ...selectedStaff, defaultRate: String(selectedStaff.defaultRate) })}
                        style={{ background: "#f1f5f9", color: "#374151", border: "1px solid #e2e8f0", borderRadius: 8, padding: "9px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>✏ Edit</button>
                    </div>
                  )}
                </div>

                <div style={{ ...grid(2, 1), gap: 16 }}>
                  <div style={{ ...card, padding: 0, overflow: "hidden" }}>
                    <div style={{ padding: "14px 16px", borderBottom: "1px solid #f1f5f9", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a" }}>☀ Daily Payments</div>
                      <div style={{ fontSize: 13, fontWeight: 800, color: "#15803d" }}>{fmt(dailyTotal)}</div>
                    </div>
                    {dailyPayments.length === 0
                      ? <div style={{ textAlign: "center", padding: 30, color: "#94a3b8", fontSize: 13 }}>No daily payments recorded yet</div>
                      : dailyPayments.map(paymentRow)}
                  </div>

                  <div style={{ ...card, padding: 0, overflow: "hidden" }}>
                    <div style={{ padding: "14px 16px", borderBottom: "1px solid #f1f5f9", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a" }}>🗓 Monthly Payments</div>
                      <div style={{ fontSize: 13, fontWeight: 800, color: "#15803d" }}>{fmt(monthlyTotal)}</div>
                    </div>
                    {monthlyPayments.length === 0
                      ? <div style={{ textAlign: "center", padding: 30, color: "#94a3b8", fontSize: 13 }}>No monthly payments recorded yet</div>
                      : monthlyPayments.map(paymentRow)}
                  </div>
                </div>
              </div>
            );
          }

          // ── Main staff list view ────────────────────────────────────────
          return (
          <div>
            <div style={{ display: "flex", flexDirection: isMobile ? "column" : "row", justifyContent: "space-between", alignItems: isMobile ? "stretch" : "center", gap: 12, marginBottom: 20 }}>
              <div>
                <div style={{ fontSize: isMobile ? 18 : 22, fontWeight: 800, color: "#0f172a" }}>Staff & Wages</div>
                <div style={{ color: "#64748b", fontSize: 13 }}>{currentTerm} · Paid so far: {fmt(totalPaidThisTerm)}</div>
              </div>
              {!isReadOnly && (
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={() => setShowBulkStaffPay(true)} style={{ background: "#f1f5f9", color: "#374151", border: "1px solid #e2e8f0", borderRadius: 9, padding: "9px 16px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>💰 Bulk Pay</button>
                  <button onClick={() => setShowAddStaff(true)} style={{ background: "#0f172a", color: "#fff", border: "none", borderRadius: 9, padding: "9px 16px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>+ Add Worker</button>
                </div>
              )}
            </div>

            <div style={{ marginBottom: 14 }}>
              <input value={staffSearch} onChange={e => setStaffSearch(e.target.value)} placeholder="🔍 Search worker name..."
                style={{ width: "100%", padding: "10px 14px", borderRadius: 9, border: "1px solid #e2e8f0", fontSize: 13, outline: "none", background: "#fff", boxSizing: "border-box" }} />
            </div>

            <div style={{ ...card, marginBottom: 18, background: "#fffbeb", border: "1px solid #fde68a" }}>
              <div style={{ fontSize: 12, color: "#92400e" }}>💡 Every payment you record here is automatically added to <strong>Expenses</strong> under "Salaries & Wages" — so your Dashboard and Net Surplus figures stay accurate without entering anything twice. Tap a worker to see their full daily and monthly payment history.</div>
            </div>

            {staff.length === 0 ? (
              <div style={{ ...card, textAlign: "center", padding: 60 }}>
                <div style={{ fontSize: 40, marginBottom: 12 }}>👷</div>
                <div style={{ fontSize: 16, fontWeight: 700, color: "#0f172a", marginBottom: 6 }}>No workers added yet</div>
                <div style={{ fontSize: 13, color: "#64748b" }}>Add cooks, security guards, drivers, matrons, and other support staff to track their pay.</div>
              </div>
            ) : activeStaff.length === 0 && inactiveStaff.length === 0 ? (
              <div style={{ ...card, textAlign: "center", padding: 40 }}>
                <div style={{ fontSize: 13, color: "#94a3b8" }}>No workers match "{staffSearch}"</div>
              </div>
            ) : (
              <>
                <div style={{ ...grid(1, 1), gap: 12, marginBottom: 18 }}>
                  {activeStaff.map(s => {
                    const paidTerm = paidThisTermByStaff(s.id);
                    const lastPaid = lastPaidDate(s.id);
                    return (
                      <div key={s.id} onClick={() => setSelectedStaffId(s.id)} style={{ ...card, cursor: "pointer", display: "flex", flexDirection: isMobile ? "column" : "row", justifyContent: "space-between", alignItems: isMobile ? "stretch" : "center", gap: 12 }}>
                        <div style={{ flex: 1 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                            <div style={{ position: "relative", flexShrink: 0 }}>
                              {s.photo
                                ? <img src={s.photo} alt={s.name} style={{ width: 34, height: 34, borderRadius: 9, objectFit: "cover", border: "2px solid #e2e8f0" }} />
                                : <div style={{ width: 34, height: 34, borderRadius: 9, background: "#e0e7ff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, border: "2px solid #e2e8f0" }}>👷</div>
                              }
                              {!isReadOnly && (
                                <button onClick={e => { e.stopPropagation(); setShowPhotoUpload(s); setPhotoUploadType("staff"); setCameraActive(false); }}
                                  title="Add / change photo"
                                  style={{ position: "absolute", bottom: -4, right: -4, width: 15, height: 15, borderRadius: "50%", background: "#f59e0b", border: "2px solid #fff", cursor: "pointer", fontSize: 7, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, color: "#0f172a" }}>
                                  📷
                                </button>
                              )}
                            </div>
                            <span style={{ fontWeight: 800, fontSize: 14, color: "#0f172a" }}>{s.name}</span>
                            <Pill text={s.role} bg="#eff6ff" col="#2563eb" />
                          </div>
                          <div style={{ display: "flex", gap: 14, fontSize: 12, color: "#64748b", flexWrap: "wrap" }}>
                            <span>Rate: <strong>{fmt(s.defaultRate)}</strong> / {s.defaultRateType === "daily" ? "day" : "month"}</span>
                            <span>Paid this term: <strong style={{ color: "#15803d" }}>{fmt(paidTerm)}</strong></span>
                            {lastPaid && <span>Last paid: <strong>{fmtDate(lastPaid)}</strong></span>}
                            {s.phone && <span>📞 {s.phone}</span>}
                          </div>
                        </div>
                        {!isReadOnly && (
                          <div style={{ display: "flex", gap: 6 }} onClick={e => e.stopPropagation()}>
                            <button onClick={() => { setShowPayStaff(s); setPayStaffForm({ amount: String(s.defaultRate || ""), payType: s.defaultRateType || "daily", periodLabel: "" }); }}
                              style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>💵 Pay</button>
                            <button onClick={() => setShowEditStaff({ ...s, defaultRate: String(s.defaultRate) })}
                              style={{ background: "#f1f5f9", color: "#374151", border: "1px solid #e2e8f0", borderRadius: 8, padding: "8px 11px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>✏</button>
                            <button onClick={() => setConfirmDialog({
                              title: "Mark Worker as Inactive",
                              message: `Mark ${s.name} as no longer working here? Their payment history is kept, and you can reactivate them later.`,
                              danger: false,
                              onConfirm: () => handleToggleStaffActive(s.id),
                            })}
                              style={{ background: "#fef2f2", color: "#dc2626", border: "1px solid #fca5a5", borderRadius: 8, padding: "8px 11px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>🚪</button>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>

                {inactiveStaff.length > 0 && (
                  <div style={{ marginBottom: 18 }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", marginBottom: 8 }}>Inactive / No Longer Working Here</div>
                    {inactiveStaff.map(s => (
                      <div key={s.id} onClick={() => setSelectedStaffId(s.id)} style={{ ...card, cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", opacity: 0.6, marginBottom: 8 }}>
                        <div style={{ fontSize: 13, color: "#64748b" }}>{s.name} · {s.role}</div>
                        {!isReadOnly && <button onClick={e => { e.stopPropagation(); handleToggleStaffActive(s.id); }} style={{ background: "#f0fdf4", color: "#15803d", border: "1px solid #86efac", borderRadius: 8, padding: "6px 11px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>↩ Reactivate</button>}
                      </div>
                    ))}
                  </div>
                )}
              </>
            )}
          </div>
          );
        })()}

        {/* ════════ SMS LOG ════════ */}
        {/* ════════ REPORTS ════════ */}
        {tab === "reports" && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
              <div>
                <div style={{ fontSize: 22, fontWeight: 800, color: "#0f172a" }}>Financial Reports</div>
                <div style={{ color: "#64748b", fontSize: 13 }}>{school.name} · {currentTerm}</div>
              </div>
              <button onClick={exportExcel} style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 9, padding: "9px 16px", fontWeight: 700, fontSize: 12, cursor: "pointer" }}>📊 Download Excel Report</button>
            </div>
            <div style={{ ...card, marginBottom: 16 }}>
              <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a", marginBottom: 14 }}>Expected vs Collected by Class</div>
              <ResponsiveContainer width="100%" height={210}>
                <BarChart data={classStats}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" vertical={false} />
                  <XAxis dataKey="class" tick={{ fontSize: 12, fill: "#64748b" }} axisLine={false} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: "#94a3b8" }} axisLine={false} tickLine={false} tickFormatter={v => `${(v / 1000000).toFixed(1)}M`} />
                  <Tooltip formatter={(v, n) => [fmt(v), n === "expected" ? "Expected" : "Collected"]} contentStyle={{ borderRadius: 8, border: "1px solid #e2e8f0", fontSize: 12 }} />
                  <Bar dataKey="expected" fill="#e2e8f0" radius={[4, 4, 0, 0]} name="expected" />
                  <Bar dataKey="collected" fill="#f59e0b" radius={[4, 4, 0, 0]} name="collected" />
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div style={{ ...grid(3, 1), gap: 14 }}>
              {classStats.map(g => (
                <div key={g.class} style={{ ...card }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
                    <span style={{ fontSize: 20, fontWeight: 800, color: "#0f172a" }}>{g.class}</span>
                    <span style={{ fontSize: 22, fontWeight: 800, color: g.rate >= 80 ? "#10b981" : g.rate >= 50 ? "#f59e0b" : "#ef4444" }}>{g.rate}%</span>
                  </div>
                  <div style={{ height: 5, background: "#f1f5f9", borderRadius: 99, marginBottom: 12 }}>
                    <div style={{ height: "100%", width: `${g.rate}%`, background: g.rate >= 80 ? "#10b981" : g.rate >= 50 ? "#f59e0b" : "#ef4444", borderRadius: 99 }} />
                  </div>
                  {[["Students", g.count], ["Expected", fmt(g.expected)], ["Collected", fmt(g.collected)], ["Balance", fmt(g.expected - g.collected)]].map(([l, v]) => (
                    <div key={l} style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 5 }}>
                      <span style={{ color: "#64748b" }}>{l}</span>
                      <span style={{ fontWeight: 600, color: l === "Balance" ? "#dc2626" : l === "Collected" ? "#15803d" : "#0f172a" }}>{v}</span>
                    </div>
                  ))}
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ════════ FEE STRUCTURE ════════ */}
        {tab === "fees" && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
              <div>
                <div style={{ fontSize: 22, fontWeight: 800, color: "#0f172a" }}>Fee Structure & Requirements</div>
                <div style={{ color: "#64748b", fontSize: 13, marginTop: 2 }}>Click any amount to edit it · Changes apply immediately to all students</div>
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={() => setConfirmDialog({
                  title: "Reset Fee Structure",
                  message: "Reset all fees to original defaults? Any custom fee amounts you've set for this school will be lost.",
                  danger: true,
                  onConfirm: async () => {
                    const defaultFees = JSON.parse(JSON.stringify(DEFAULT_FEE_STRUCTURE));
                    setFeeStructure(defaultFees);
                    await saveSchoolConfig(defaultFees, requirements);
                    notify("Fees reset to defaults");
                  },
                })}
                  style={{ background: "#f1f5f9", color: "#374151", border: "1px solid #e2e8f0", borderRadius: 9, padding: "9px 14px", fontWeight: 700, fontSize: 12, cursor: "pointer" }}>
                  ↺ Reset to Defaults
                </button>
              </div>
            </div>

            {/* ── Editable Fee Tables ── */}
            {STUDENT_CATEGORIES.map(cat => (
              <div key={cat} style={{ ...card, padding: 0, overflow: "hidden", marginBottom: 16 }}>
                <div style={{ padding: "13px 18px", background: cat === "Boarder" ? "#0f172a" : "#1e3a5f", display: "flex", alignItems: "center", gap: 10 }}>
                  <span style={{ fontSize: 16 }}>{cat === "Boarder" ? "🏠" : "🚶"}</span>
                  <div style={{ color: "#f1f5f9", fontWeight: 800, fontSize: 14 }}>{cat} Fee Structure</div>
                  <span style={{ color: "#94a3b8", fontSize: 11, marginLeft: 8 }}>Click any cell to edit</span>
                  <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center" }}>
                    <button onClick={() => { setShowAddFeeItem(cat); setNewFeeItemName(""); setNewFeeItemAmt(""); }}
                      style={{ background: "#3b82f6", color: "#fff", border: "none", borderRadius: 7, padding: "5px 12px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                      + Add Fee Item
                    </button>
                    <span style={{ background: "#f59e0b", color: "#0f172a", padding: "2px 10px", borderRadius: 99, fontSize: 11, fontWeight: 700 }}>
                      {termStudents.filter(s => (s.category || "Day Scholar") === cat).length} students
                    </span>
                  </div>
                </div>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr style={{ background: "#f8fafc" }}>
                      <th style={{ padding: "10px 16px", textAlign: "left", fontSize: 11, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase" }}>Class</th>
                      {Object.keys(feeStructure[cat][schoolClasses[0]] || {}).map(field => (
                        <th key={field} style={{ padding: "10px 16px", textAlign: "left", fontSize: 11, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase" }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                            {field}
                            <button onClick={() => handleRemoveFeeItem(cat, field)} title={`Remove ${field}`}
                              style={{ background: "#fef2f2", color: "#dc2626", border: "none", borderRadius: 5, width: 16, height: 16, fontSize: 10, fontWeight: 800, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", lineHeight: 1 }}>
                              ✕
                            </button>
                          </div>
                        </th>
                      ))}
                      <th style={{ padding: "10px 16px", textAlign: "left", fontSize: 11, fontWeight: 700, color: "#10b981", textTransform: "uppercase" }}>Total / Term</th>
                    </tr>
                  </thead>
                  <tbody>
                    {schoolClasses.map((cls, i) => {
                      const f = feeStructure[cat][cls] || {};
                      const total = Object.values(f).reduce((a, b) => a + b, 0);
                      return (
                        <tr key={cls} style={{ borderTop: "1px solid #f1f5f9", background: i % 2 === 0 ? "#fff" : "#fafbfc" }}>
                          <td style={{ padding: "10px 16px" }}><Pill text={cls} /></td>
                          {Object.entries(f).map(([field, val]) => {
                            const isEditing = feeEditCell?.cat === cat && feeEditCell?.cls === cls && feeEditCell?.field === field;
                            return (
                              <td key={field} style={{ padding: "6px 10px" }}>
                                {isEditing ? (
                                  <div style={{ display: "flex", gap: 4 }}>
                                    <input type="number" value={feeEditVal} onChange={e => setFeeEditVal(e.target.value)}
                                      onKeyDown={e => { if (e.key === "Enter") saveFeeCell(); if (e.key === "Escape") setFeeEditCell(null); }}
                                      autoFocus style={{ width: 110, padding: "5px 8px", borderRadius: 7, border: "2px solid #f59e0b", fontSize: 13, fontWeight: 700, outline: "none" }} />
                                    <button onClick={saveFeeCell} style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 6, padding: "4px 8px", fontSize: 11, cursor: "pointer", fontWeight: 700 }}>✓</button>
                                    <button onClick={() => setFeeEditCell(null)} style={{ background: "#f1f5f9", color: "#64748b", border: "none", borderRadius: 6, padding: "4px 8px", fontSize: 11, cursor: "pointer" }}>✕</button>
                                  </div>
                                ) : (
                                  <button onClick={() => { setFeeEditCell({ cat, cls, field }); setFeeEditVal(String(val)); }}
                                    style={{ background: "transparent", border: "1px dashed #e2e8f0", borderRadius: 7, padding: "5px 10px", fontSize: 13, color: "#374151", cursor: "pointer", fontWeight: 500, width: "100%", textAlign: "left" }}
                                    title="Click to edit">
                                    {fmt(val)} <span style={{ fontSize: 9, color: "#94a3b8" }}>✏</span>
                                  </button>
                                )}
                              </td>
                            );
                          })}
                          <td style={{ padding: "10px 16px", fontSize: 14, fontWeight: 800, color: "#10b981" }}>{fmt(total)}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            ))}

            {/* ── School Requirements ── */}
            <div style={{ ...card, marginBottom: 16 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a" }}>📋 School Requirements</div>
                  <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>Items students must bring. Mandatory items are added to the total fee automatically.</div>
                </div>
                <button onClick={() => setShowAddReq(true)} style={{ background: "#3b82f6", color: "#fff", border: "none", borderRadius: 9, padding: "8px 14px", fontWeight: 700, fontSize: 12, cursor: "pointer" }}>+ Add Item</button>
              </div>

              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr style={{ background: "#f8fafc" }}>
                    {["Item", "Cost (UGX)", "Applies To", "Added to Fee?", "Actions"].map(h => (
                      <th key={h} style={{ padding: "9px 14px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {requirements.map((r, i) => (
                    <tr key={r.id} style={{ borderTop: "1px solid #f1f5f9", background: i % 2 === 0 ? "#fff" : "#fafbfc" }}>
                      <td style={{ padding: "10px 14px", fontWeight: 600, fontSize: 13, color: "#0f172a" }}>{r.name}</td>
                      <td style={{ padding: "10px 14px" }}>
                        {/* Inline cost edit */}
                        <input type="number" defaultValue={r.cost}
                          onBlur={e => {
                            const val = parseInt(e.target.value);
                            if (val > 0 && val !== r.cost) {
                              setRequirements(prev => {
                                const updated = prev.map(req => req.id === r.id ? { ...req, cost: val } : req);
                                saveSchoolConfig(feeStructure, updated);
                                return updated;
                              });
                              notify(`${r.name} cost updated to ${fmt(val)}`);
                            }
                          }}
                          style={{ width: 110, padding: "4px 8px", borderRadius: 7, border: "1px dashed #e2e8f0", fontSize: 12, fontWeight: 700, outline: "none", background: "transparent" }}
                          title="Click to edit cost" />
                      </td>
                      <td style={{ padding: "10px 14px" }}>
                        <div style={{ display: "flex", gap: 4 }}>
                          {STUDENT_CATEGORIES.map(cat => (
                            <button key={cat} onClick={() => setRequirements(prev => {
                              const updated = prev.map(req => req.id === r.id ? { ...req, appliesTo: req.appliesTo.includes(cat) ? req.appliesTo.filter(a => a !== cat) : [...req.appliesTo, cat] } : req);
                              saveSchoolConfig(feeStructure, updated);
                              return updated;
                            })}
                              style={{ padding: "2px 8px", borderRadius: 99, border: "none", cursor: "pointer", fontSize: 11, fontWeight: 700, background: r.appliesTo.includes(cat) ? (cat === "Boarder" ? "#f0fdf4" : "#eff6ff") : "#f1f5f9", color: r.appliesTo.includes(cat) ? (cat === "Boarder" ? "#15803d" : "#2563eb") : "#94a3b8" }}>
                              {cat}
                            </button>
                          ))}
                        </div>
                      </td>
                      <td style={{ padding: "10px 14px" }}>
                        <button onClick={() => setRequirements(prev => {
                          const updated = prev.map(req => req.id === r.id ? { ...req, mandatory: !req.mandatory } : req);
                          saveSchoolConfig(feeStructure, updated);
                          return updated;
                        })}
                          style={{ padding: "4px 12px", borderRadius: 99, border: "none", cursor: "pointer", fontWeight: 700, fontSize: 11, background: r.mandatory ? "#d1fae5" : "#fef3c7", color: r.mandatory ? "#065f46" : "#92400e" }}>
                          {r.mandatory ? "✓ Yes — included in fee" : "○ Optional — not in fee"}
                        </button>
                      </td>
                      <td style={{ padding: "10px 14px" }}>
                        <button onClick={() => deleteReq(r.id)} style={{ background: "#fef2f2", color: "#dc2626", border: "none", borderRadius: 7, padding: "4px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>Remove</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>

              {/* Requirements cost summary */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginTop: 16 }}>
                {STUDENT_CATEGORIES.map(cat => {
                  const reqTotal = requirements.filter(r => r.mandatory && r.appliesTo.includes(cat)).reduce((a, r) => a + r.cost, 0);
                  return (
                    <div key={cat} style={{ background: "#f8fafc", borderRadius: 10, padding: "12px 16px" }}>
                      <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b", textTransform: "uppercase", marginBottom: 4 }}>{cat} — Mandatory Requirements Total</div>
                      <div style={{ fontSize: 18, fontWeight: 800, color: "#0f172a" }}>{fmt(reqTotal)}</div>
                      <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 2 }}>Added on top of tuition fees</div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Bursary summary */}
            {termStudents.some(s => s.bursary || s.customFee) && (
              <div style={{ ...card, marginBottom: 16 }}>
                <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a", marginBottom: 14 }}>🎓 Students with Special Fee Arrangements</div>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead><tr>{["Student", "Class", "Category", "Arrangement", "Standard Fee", "Their Fee", "Saving"].map(h => (
                    <th key={h} style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", textAlign: "left", padding: "0 8px 8px" }}>{h}</th>
                  ))}</tr></thead>
                  <tbody>
                    {termStudents.filter(s => s.bursary || s.customFee).map((s, i) => {
                      const standardFee = totalFee(s.class, s.category || "Day Scholar");
                      const actualFee = getStudentFee(s);
                      return (
                        <tr key={s.id} style={{ borderTop: "1px solid #f1f5f9", background: i % 2 === 0 ? "#fff" : "#fafbfc" }}>
                          <td style={{ padding: "10px 8px", fontWeight: 600, fontSize: 13 }}>{s.name}</td>
                          <td style={{ padding: "10px 8px" }}><Pill text={classLabel(s.class, s.stream)} /></td>
                          <td style={{ padding: "10px 8px", fontSize: 12, color: "#64748b" }}>{s.category || "Day Scholar"}</td>
                          <td style={{ padding: "10px 8px" }}>
                            {s.customFee ? <Pill text="Custom Fee" bg="#f5f3ff" col="#7c3aed" />
                              : <Pill text={`${s.bursary.type === "percent" ? s.bursary.value + "% off" : fmt(s.bursary.value) + " off"} · ${s.bursary.reason}`} bg="#fef3c7" col="#92400e" />}
                          </td>
                          <td style={{ padding: "10px 8px", fontSize: 13 }}>{fmt(standardFee)}</td>
                          <td style={{ padding: "10px 8px", fontSize: 13, fontWeight: 800, color: "#7c3aed" }}>{fmt(actualFee)}</td>
                          <td style={{ padding: "10px 8px", fontSize: 13, fontWeight: 700, color: "#10b981" }}>{fmt(standardFee - actualFee)}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            <div style={{ ...card, background: "#0f172a", color: "#f1f5f9" }}>
              <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 12 }}>📌 Fee Policy</div>
              {["Click any fee amount in the table above to edit it — changes apply to all students immediately.", "Boarders pay an additional boarding levy on top of tuition.", "Mandatory requirements are automatically added to every student's total fee.", "Optional requirements appear on the student's checklist but are not charged.", "Bursaries and scholarships are approved by the school board each term.", "All fee changes are logged and auditable."].map((p, i) => (
                <div key={i} style={{ display: "flex", gap: 10, marginBottom: 7, fontSize: 13, color: "#94a3b8" }}>
                  <span style={{ color: "#f59e0b", fontWeight: 700 }}>{i + 1}.</span>{p}
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ════════ SETTINGS ════════ */}
        {tab === "settings" && (
          <div>
            <div style={{ fontSize: 22, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>Settings</div>
            <div style={{ color: "#64748b", fontSize: 13, marginBottom: 24 }}>Manage your account, school profile and security</div>

            <div style={{ ...grid(2, 1), gap: 20 }}>

              {/* ── Change Password ── */}
              <div style={{ ...card }}>
                <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a", marginBottom: 4 }}>🔐 Change Password</div>
                <div style={{ fontSize: 12, color: "#64748b", marginBottom: 20 }}>Update the admin/bursar login password</div>

                <div style={{ marginBottom: 14 }}>
                  <label style={lbl}>Current Password</label>
                  <input type="password" value={pwForm.currentPw} onChange={e => { setPwForm(p => ({ ...p, currentPw: e.target.value })); setPwError(""); }}
                    placeholder="Enter current password" style={inp} />
                </div>
                <div style={{ marginBottom: 14 }}>
                  <label style={lbl}>New Password</label>
                  <input type="password" value={pwForm.newPw} onChange={e => { setPwForm(p => ({ ...p, newPw: e.target.value })); setPwError(""); }}
                    placeholder="Enter new password (min 6 chars)" style={inp} />
                </div>
                <div style={{ marginBottom: 16 }}>
                  <label style={lbl}>Confirm New Password</label>
                  <input type="password" value={pwForm.confirmPw} onChange={e => { setPwForm(p => ({ ...p, confirmPw: e.target.value })); setPwError(""); }}
                    placeholder="Re-enter new password" style={inp} />
                </div>

                {pwError && (
                  <div style={{ background: "#fef2f2", color: "#b91c1c", padding: "10px 13px", borderRadius: 9, fontSize: 12, marginBottom: 14, fontWeight: 600 }}>
                    ✗ {pwError}
                  </div>
                )}

                <button onClick={async () => {
                  if (pwForm.newPw.length < 6) return setPwError("New password must be at least 6 characters.");
                  if (pwForm.newPw !== pwForm.confirmPw) return setPwError("New passwords do not match.");
                  // Verify current password by re-authenticating with Supabase Auth
                  const school = SCHOOLS_DATA[activeSchoolId];
                  const notifyEmail = school?.notifyEmail;
                  if (!notifyEmail) return setPwError("No email set for this school. Set one in School Profile first.");
                  const { error: verifyError } = await supabase.auth.signInWithPassword({
                    email: notifyEmail, password: pwForm.currentPw,
                  });
                  if (verifyError) return setPwError("Current password is incorrect.");
                  // Update password in Supabase Auth
                  const { error: updateError } = await supabase.auth.updateUser({ password: pwForm.newPw });
                  if (updateError) return setPwError(`Could not update password: ${updateError.message}`);
                  // Also update admin_password column so it stays in sync
                  await supabase.from("schools").update({ admin_password: pwForm.newPw }).eq("id", activeSchoolId);
                  SCHOOLS_DATA[activeSchoolId].adminPassword = pwForm.newPw;
                  setPwForm({ currentPw: "", newPw: "", confirmPw: "" });
                  setPwError("");
                  notify("Password changed successfully ✓");
                }} style={{ width: "100%", padding: 11, borderRadius: 9, border: "none", background: "#0f172a", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>
                  Update Password
                </button>

                <div style={{ marginTop: 14, padding: "10px 13px", background: "#f8fafc", borderRadius: 9, fontSize: 11, color: "#94a3b8" }}>
                  Current username: <strong style={{ color: "#0f172a" }}>{adminCreds.username}</strong>
                </div>
              </div>

              {/* ── Change Username ── */}
              <div style={{ ...card }}>
                <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a", marginBottom: 4 }}>👤 Change Username</div>
                <div style={{ fontSize: 12, color: "#64748b", marginBottom: 20 }}>Update the login username for this account</div>

                <div style={{ marginBottom: 14 }}>
                  <label style={lbl}>Current Username</label>
                  <input value={adminCreds.username} disabled style={{ ...inp, background: "#f8fafc", color: "#94a3b8" }} />
                </div>
                <div style={{ marginBottom: 20 }}>
                  <label style={lbl}>New Username</label>
                  <input id="new-username" placeholder="Enter new username" style={inp} />
                </div>

                <button onClick={async () => {
                  const val = document.getElementById("new-username").value.trim();
                  if (!val) return notify("Enter a username", "err");
                  if (val.length < 3) return notify("Username must be at least 3 characters", "err");
                  const { error } = await supabase.from("schools").update({ admin_username: val }).eq("id", activeSchoolId);
                  if (error) return notify(`Could not update username: ${error.message}`, "err");
                  SCHOOLS_DATA[activeSchoolId].adminUsername = val;
                  setAdminCreds(prev => ({ ...prev, username: val }));
                  document.getElementById("new-username").value = "";
                  notify(`Username changed to "${val}" ✓`);
                }} style={{ width: "100%", padding: 11, borderRadius: 9, border: "none", background: "#3b82f6", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>
                  Update Username
                </button>
              </div>

              {/* ── School Profile ── */}
              <div style={{ ...card }}>
                <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a", marginBottom: 4 }}>🏫 School Profile</div>
                <div style={{ fontSize: 12, color: "#64748b", marginBottom: 20 }}>Update school information shown on receipts and reports</div>

                {[
                  { label: "School Name", key: "name", placeholder: "e.g. Kampala Senior Secondary School" },
                  { label: "Location", key: "location", placeholder: "e.g. Kampala" },
                  { label: "Principal Name", key: "principal", placeholder: "e.g. Mr. Ssempijja Robert" },
                  { label: "School Phone", key: "phone", placeholder: "e.g. 0772-000-001" },
                  { label: "Notification Email", key: "notifyEmail", placeholder: "e.g. bursar@kampalass.ac.ug" },
                ].map(f => (
                  <div key={f.key} style={{ marginBottom: 13 }}>
                    <label style={lbl}>{f.label}</label>
                    <input value={schoolProfile[f.key] || ""} onChange={e => setSchoolProfile(p => ({ ...p, [f.key]: e.target.value }))}
                      placeholder={f.placeholder} style={inp} />
                    {f.key === "notifyEmail" && <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 4 }}>Used for subscription reminders, receipts, and other important alerts from FeeTrack UG.</div>}
                  </div>
                ))}

                <button onClick={async () => {
                  if (schoolProfile.notifyEmail && !schoolProfile.notifyEmail.includes("@")) return notify("Enter a valid email address", "err");
                  const updates = {
                    name: schoolProfile.name || SCHOOLS_DATA[activeSchoolId].name,
                    location: schoolProfile.location || SCHOOLS_DATA[activeSchoolId].location,
                    principal: schoolProfile.principal || SCHOOLS_DATA[activeSchoolId].principal,
                    phone: schoolProfile.phone || SCHOOLS_DATA[activeSchoolId].phone,
                    notify_email: schoolProfile.notifyEmail || "",
                  };
                  const { data, error } = await supabase.from("schools").update(updates).eq("id", activeSchoolId).select();
                  if (error) {
                    console.error("Failed to save school profile:", error.message);
                    return notify("Could not save changes — please try again", "err");
                  }
                  if (!data || data.length === 0) {
                    console.error("saveProfile — update matched zero rows for activeSchoolId:", activeSchoolId);
                    return notify("Could not save changes — school not found", "err");
                  }
                  // Mirror the same change into the in-memory copy too, so the
                  // rest of the app (which still reads from SCHOOLS_DATA in
                  // several places) reflects the update immediately without
                  // needing a full page reload.
                  SCHOOLS_DATA[activeSchoolId].name = updates.name;
                  SCHOOLS_DATA[activeSchoolId].location = updates.location;
                  SCHOOLS_DATA[activeSchoolId].principal = updates.principal;
                  SCHOOLS_DATA[activeSchoolId].phone = updates.phone;
                  SCHOOLS_DATA[activeSchoolId].notifyEmail = updates.notify_email;
                  setSubscriptionRefresh(r => r + 1);
                  notify("School profile updated ✓");
                }} style={{ width: "100%", padding: 11, borderRadius: 9, border: "none", background: "#10b981", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>
                  Save Profile
                </button>
              </div>

              {/* ── Class Streams ── */}
              <div style={{ ...card, marginTop: 20 }}>
                <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a", marginBottom: 4 }}>🏷 Class Streams</div>
                <div style={{ fontSize: 12, color: "#64748b", marginBottom: 18, lineHeight: 1.6 }}>
                  If your classes are split into streams (e.g. "East" / "West", "A" / "B", "Science" / "Arts"), add them here. Classes with no streams are treated as one undivided group — most schools can skip this entirely.
                </div>

                {schoolClasses.map(cls => {
                  const streams = (streamsForm ?? school.streams ?? {})[cls] || [];
                  return (
                    <div key={cls} style={{ marginBottom: 14, paddingBottom: 14, borderBottom: "1px solid #f1f5f9" }}>
                      <div style={{ fontSize: 13, fontWeight: 700, color: "#374151", marginBottom: 8 }}>{cls}</div>
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 8 }}>
                        {streams.length === 0 && <span style={{ fontSize: 12, color: "#94a3b8" }}>No streams — treated as one class</span>}
                        {streams.map(s => (
                          <span key={s} style={{ display: "flex", alignItems: "center", gap: 6, background: "#f5f3ff", color: "#7c3aed", borderRadius: 7, padding: "5px 10px", fontSize: 12, fontWeight: 700 }}>
                            {classLabel(cls, s)}
                            <button onClick={() => removeStream(cls, s)} style={{ background: "none", border: "none", color: "#7c3aed", cursor: "pointer", fontSize: 13, padding: 0, lineHeight: 1 }}>✕</button>
                          </span>
                        ))}
                      </div>
                      <div style={{ display: "flex", gap: 8 }}>
                        <input value={newStreamInput[cls] || ""} onChange={e => setNewStreamInput(prev => ({ ...prev, [cls]: e.target.value }))}
                          onKeyDown={e => { if (e.key === "Enter") addStream(cls); }}
                          placeholder="e.g. East, A, Science" style={{ ...inp, flex: 1 }} />
                        <button onClick={() => addStream(cls)} style={{ background: "#f1f5f9", color: "#374151", border: "1px solid #e2e8f0", borderRadius: 9, padding: "9px 16px", fontWeight: 700, fontSize: 12, cursor: "pointer", whiteSpace: "nowrap" }}>+ Add Stream</button>
                      </div>
                    </div>
                  );
                })}

                {streamsForm && (
                  <button onClick={saveStreams} style={{ width: "100%", padding: 11, borderRadius: 9, border: "none", background: "#10b981", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer", marginTop: 4 }}>
                    Save Stream Changes
                  </button>
                )}
              </div>

              {/* ── Session & Logout ── */}
              <div style={{ ...card }}>
                <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a", marginBottom: 4 }}>🔒 Session & Security</div>
                <div style={{ fontSize: 12, color: "#64748b", marginBottom: 20 }}>Manage your current login session</div>

                <div style={{ background: "#f0fdf4", borderRadius: 10, padding: "14px 16px", marginBottom: 16 }}>
                  <div style={{ fontSize: 11, color: "#15803d", fontWeight: 700, textTransform: "uppercase", marginBottom: 6 }}>Currently Logged In</div>
                  <div style={{ fontSize: 13, fontWeight: 600, color: "#0f172a" }}>👤 {adminCreds.username}</div>
                  <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>Role: Admin / Bursar · {school.name}</div>
                  <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>Session started this browser session</div>
                </div>

                <div style={{ background: "#fef3c7", borderRadius: 10, padding: "12px 14px", marginBottom: 16, fontSize: 12, color: "#92400e" }}>
                  <strong>⚠ Security Tips:</strong>
                  <div style={{ marginTop: 6, lineHeight: 1.7 }}>
                    • Always log out when leaving the computer<br />
                    • Never share your password with others<br />
                    • Change your password every term
                  </div>
                </div>

                <button onClick={logout}
                  style={{ width: "100%", padding: 12, borderRadius: 9, border: "none", background: "#ef4444", color: "#fff", fontSize: 14, fontWeight: 800, cursor: "pointer", marginBottom: 10 }}>
                  🚪 Log Out Now
                </button>

                <button onClick={() => {
                  clearSession();
                  setCurrentUser(null);
                  setActiveSchoolId(null);
                  setTab("dashboard");
                  setLoginInput({ user: "", pass: "" });
                  notify("Screen locked — log in again to continue");
                }}
                  style={{ width: "100%", padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", color: "#374151", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                  🔒 Lock Screen
                </button>
              </div>

            </div>

            {/* Subscription & Billing */}
            <div style={{ ...card, marginTop: 20 }}>
              <div style={{ fontWeight: 700, fontSize: 15, color: "#0f172a", marginBottom: 4 }}>💳 Subscription & Billing</div>
              <div style={{ fontSize: 12, color: "#64748b", marginBottom: 16 }}>Your current plan and payment status</div>

              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12, padding: 16, borderRadius: 12, background: subInfo.status === "Active" ? "#f0fdf4" : subInfo.status === "Grace Period" ? "#fffbeb" : "#fef2f2", border: `1px solid ${subInfo.status === "Active" ? "#86efac" : subInfo.status === "Grace Period" ? "#fde68a" : "#fca5a5"}` }}>
                <div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                    <span style={{ fontWeight: 800, fontSize: 16, color: "#0f172a" }}>{school.plan} Plan</span>
                    <Pill text={subInfo.status} bg="#fff" col={subInfo.status === "Active" ? "#10b981" : subInfo.status === "Grace Period" ? "#f59e0b" : "#ef4444"} />
                    {school.isTrial && <Pill text="Free Trial" bg="#eff6ff" col="#2563eb" />}
                  </div>
                  <div style={{ fontSize: 13, color: "#374151" }}>{fmt(getBillingInfo(school.plan, school.billingCycle, school.customPrice).price)}{getBillingInfo(school.plan, school.billingCycle, school.customPrice).periodLabel} ({school.billingCycle === "term" ? "Per Term" : "Monthly"}) · Up to {PLANS[school.plan]?.maxStudents === Infinity ? "unlimited" : PLANS[school.plan]?.maxStudents} students</div>
                  <div style={{ fontSize: 12, color: "#64748b", marginTop: 4 }}>
                    {school.nextBillingDate ? <>Next payment due: <strong>{fmtDate(school.nextBillingDate)}</strong></> : "Trial starting..."}
                    {school.lastPaymentDate && <> · Last paid: <strong>{fmtDate(school.lastPaymentDate)}</strong></>}
                  </div>
                  {subInfo.status === "Grace Period" && <div style={{ fontSize: 12, color: "#92400e", fontWeight: 700, marginTop: 6 }}>⚠ {subInfo.daysRemaining} day(s) left before your account becomes read-only</div>}
                  {subInfo.status === "Suspended" && <div style={{ fontSize: 12, color: "#b91c1c", fontWeight: 700, marginTop: 6 }}>🔒 Account is read-only — renew to restore full access</div>}
                </div>
              </div>

              {/* Billing Cycle Switcher */}
              <div style={{ marginTop: 14, border: "1px solid #e2e8f0", borderRadius: 12, padding: 16 }}>
                <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a", marginBottom: 4 }}>Billing Cycle</div>
                <div style={{ fontSize: 12, color: "#64748b", marginBottom: 12 }}>Pay every month, or pay once per school term (about every 3 months) and save 10%.</div>
                <div style={{ display: "flex", gap: 10 }}>
                  {["monthly", "term"].map(cyc => {
                    const info = getBillingInfo(school.plan, cyc);
                    const selected = (school.billingCycle || "monthly") === cyc;
                    return (
                      <button key={cyc} onClick={() => changeBillingCycle(activeSchoolId, cyc)}
                        style={{ flex: 1, padding: "12px 14px", borderRadius: 10, border: `2px solid ${selected ? "#8b5cf6" : "#e2e8f0"}`, background: selected ? "#f5f3ff" : "#fff", cursor: "pointer", textAlign: "left" }}>
                        <div style={{ fontSize: 13, fontWeight: 800, color: selected ? "#6d28d9" : "#374151" }}>{cyc === "monthly" ? "Monthly" : "Per Term"} {selected && <span style={{ fontSize: 10, fontWeight: 700 }}>✓ Current</span>}</div>
                        <div style={{ fontSize: 15, fontWeight: 800, color: "#0f172a", marginTop: 2 }}>{fmt(info.price)}<span style={{ fontSize: 11, color: "#94a3b8" }}>{info.periodLabel}</span></div>
                        {cyc === "term" && <div style={{ fontSize: 10, color: "#15803d", fontWeight: 700, marginTop: 2 }}>Save 10% vs. monthly</div>}
                      </button>
                    );
                  })}
                </div>
                <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 8 }}>Changing this only affects your next payment amount and due date — it doesn't change your current balance or due date.</div>
              </div>

              {/* Pay Now — Payment Instructions */}
              <div style={{ marginTop: 14, border: "2px solid #8b5cf6", borderRadius: 12, padding: 16, background: "#f5f3ff" }}>
                <div style={{ fontWeight: 800, fontSize: 14, color: "#6d28d9", marginBottom: 10 }}>💰 Pay Now to Renew</div>
                <div style={{ ...grid(2, 1), gap: 14, marginBottom: 12 }}>
                  <div style={{ background: "#fff", borderRadius: 10, padding: 12 }}>
                    <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", marginBottom: 6 }}>📱 MTN MoMo</div>
                    <div style={{ fontSize: 13, color: "#374151" }}>Send to: <strong style={{ fontFamily: "monospace" }}>{platformPayInfo.momoNumber}</strong></div>
                    <div style={{ fontSize: 13, color: "#374151" }}>Name: <strong>{platformPayInfo.momoName}</strong></div>
                  </div>
                  <div style={{ background: "#fff", borderRadius: 10, padding: 12 }}>
                    <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", marginBottom: 6 }}>🏦 Bank Transfer</div>
                    <div style={{ fontSize: 13, color: "#374151" }}>Bank: <strong>{platformPayInfo.bankName}</strong></div>
                    <div style={{ fontSize: 13, color: "#374151" }}>Account: <strong style={{ fontFamily: "monospace" }}>{platformPayInfo.bankAccount}</strong></div>
                  </div>
                </div>
                <div style={{ background: "#fff", borderRadius: 10, padding: 12, marginBottom: 12 }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", marginBottom: 4 }}>Amount Due ({school.billingCycle === "term" ? "Per Term" : "Monthly"})</div>
                  <div style={{ fontSize: 20, fontWeight: 800, color: "#0f172a" }}>{fmt(getBillingInfo(school.plan, school.billingCycle, school.customPrice).price)}</div>
                </div>
                <div style={{ background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 10, padding: 12 }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "#92400e", textTransform: "uppercase", marginBottom: 4 }}>⚠ Important — Reference Code (REQUIRED)</div>
                  <div style={{ fontSize: 18, fontWeight: 800, color: "#0f172a", fontFamily: "monospace", letterSpacing: 1 }}>{school.billingRef}</div>
                  <div style={{ fontSize: 11, color: "#78350f", marginTop: 6, lineHeight: 1.6 }}>
                    Always include this code in the payment reference/narration when sending money. This allows our system to automatically detect your payment and restore your account access — usually within minutes, with no need to contact us.
                    <br /><br />
                    <strong>Want to upgrade or downgrade?</strong> Simply pay the exact price of the plan you want (with your reference code) — your plan switches automatically too, no approval needed.
                  </div>
                </div>
                <button onClick={() => setShowPaymentConfirm(true)} style={{ width: "100%", marginTop: 12, padding: 12, borderRadius: 10, border: "none", background: "#10b981", color: "#fff", fontSize: 14, fontWeight: 800, cursor: "pointer" }}>
                  ✓ I've Sent Payment
                </button>
              </div>

              <div style={{ marginTop: 12, fontSize: 13, fontWeight: 700, color: "#0f172a" }}>Available Plans</div>
              <div style={{ marginTop: 10, ...grid(3, 1), gap: 10 }}>
                {Object.values(PLANS).map(p => {
                  const monthly = getBillingInfo(p.name, "monthly");
                  const term = getBillingInfo(p.name, "term");
                  return (
                  <div key={p.name} style={{ border: `2px solid ${p.name === school.plan ? "#8b5cf6" : "#e2e8f0"}`, borderRadius: 10, padding: 12, background: p.name === school.plan ? "#f5f3ff" : "#fff" }}>
                    <div style={{ fontWeight: 800, fontSize: 13, color: "#0f172a" }}>{p.name}{p.name === school.plan && <span style={{ color: "#7c3aed", fontSize: 11, marginLeft: 6 }}>Current</span>}</div>
                    <div style={{ marginTop: 4 }}>
                      <div style={{ fontSize: 15, fontWeight: 800, color: "#7c3aed" }}>{fmt(monthly.price)}<span style={{ fontSize: 10, color: "#94a3b8" }}>{monthly.periodLabel}</span></div>
                      <div style={{ fontSize: 12, fontWeight: 700, color: "#15803d", marginTop: 1 }}>{fmt(term.price)}<span style={{ fontSize: 10, color: "#94a3b8" }}>{term.periodLabel}</span> <span style={{ fontSize: 9, color: "#94a3b8" }}>(save 10%)</span></div>
                    </div>
                    <div style={{ fontSize: 10, color: "#94a3b8", marginTop: 6, marginBottom: 8 }}>Up to {p.maxStudents === Infinity ? "unlimited" : p.maxStudents} students</div>
                    {p.features.map(f => <div key={f} style={{ fontSize: 10, color: "#64748b", padding: "1px 0" }}>✓ {f}</div>)}
                  </div>
                  );
                })}
              </div>
            </div>

            {/* App Info */}
            <div style={{ ...card, marginTop: 20, background: "#0f172a", color: "#f1f5f9" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
                <div>
                  <div style={{ fontWeight: 800, fontSize: 15 }}>🏫 FeeTrack UG — School Finance System</div>
                  <div style={{ color: "#64748b", fontSize: 12, marginTop: 4 }}>Version 1.0 · Built for Uganda Schools · Supports MTN MoMo, Airtel Money, Cash, Bank</div>
                </div>
                <div style={{ textAlign: isMobile ? "left" : "right" }}>
                  <div style={{ color: "#f59e0b", fontWeight: 700, fontSize: 13 }}>{school.name}</div>
                  <div style={{ color: "#64748b", fontSize: 11, marginTop: 2 }}>{termStudents.length} students · {currentTerm}</div>
                </div>
              </div>
            </div>

          </div>
        )}

        {/* ════════ BANK RECONCILIATION ════════ */}
        {tab === "bank" && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 22 }}>
              <div>
                <div style={{ fontSize: 22, fontWeight: 800, color: "#0f172a" }}>🏦 Bank Reconciliation</div>
                <div style={{ color: "#64748b", fontSize: 13, marginTop: 2 }}>Upload your bank statement to automatically record student payments</div>
              </div>
              <button onClick={() => bankFileRef.current?.click()}
                style={{ background: "#0f172a", color: "#fff", border: "none", borderRadius: 9, padding: "10px 18px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                📂 Upload Bank Statement
              </button>
              <input ref={bankFileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleBankFileUpload} style={{ display: "none" }} />
            </div>

            {/* How it works */}
            {bankRows.length === 0 && (
              <div>
                <div style={{ ...grid(3, 1), gap: 16, marginBottom: 20 }}>
                  {[
                    { step: "1", title: "Download Statement", desc: "Log into your school's online banking (Stanbic, Centenary, DFCU, PostBank). Download the account statement as Excel or CSV.", icon: "🏦" },
                    { step: "2", title: "Upload Here", desc: "Click 'Upload Bank Statement' above. The app reads it automatically and tries to match each deposit to a student.", icon: "📂" },
                    { step: "3", title: "Review & Import", desc: "Check the matches, fix any that are wrong, then click 'Import All Matched'. Payments are recorded instantly.", icon: "✅" },
                  ].map(s => (
                    <div key={s.step} style={{ ...card, textAlign: "center", padding: "28px 20px" }}>
                      <div style={{ fontSize: 36, marginBottom: 10 }}>{s.icon}</div>
                      <div style={{ width: 28, height: 28, borderRadius: "50%", background: "#0f172a", color: "#f59e0b", fontWeight: 800, fontSize: 13, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 10px" }}>{s.step}</div>
                      <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a", marginBottom: 6 }}>{s.title}</div>
                      <div style={{ fontSize: 12, color: "#64748b", lineHeight: 1.6 }}>{s.desc}</div>
                    </div>
                  ))}
                </div>

                {/* Supported banks */}
                <div style={{ ...card }}>
                  <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a", marginBottom: 14 }}>🇺🇬 Supported Uganda Banks</div>
                  <div style={{ ...grid(4, 2), gap: 12 }}>
                    {[
                      { name: "Stanbic Bank", format: "Excel (.xlsx)", color: "#1e40af" },
                      { name: "Centenary Bank", format: "Excel (.xlsx)", color: "#15803d" },
                      { name: "DFCU Bank", format: "CSV or Excel", color: "#b45309" },
                      { name: "PostBank Uganda", format: "Excel (.xlsx)", color: "#7c3aed" },
                      { name: "Equity Bank", format: "CSV or Excel", color: "#dc2626" },
                      { name: "Absa Uganda", format: "CSV (.csv)", color: "#b91c1c" },
                      { name: "Housing Finance", format: "Excel (.xlsx)", color: "#0369a1" },
                      { name: "Any Bank", format: "CSV or Excel", color: "#374151" },
                    ].map(b => (
                      <div key={b.name} style={{ padding: "12px 14px", borderRadius: 10, border: `2px solid ${b.color}20`, background: `${b.color}08` }}>
                        <div style={{ fontWeight: 700, fontSize: 13, color: b.color }}>{b.name}</div>
                        <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 2 }}>{b.format}</div>
                      </div>
                    ))}
                  </div>
                  <div style={{ marginTop: 16, padding: "12px 16px", background: "#fffbeb", borderRadius: 10, fontSize: 12, color: "#92400e" }}>
                    <strong>💡 Tip for parents:</strong> When depositing at the bank, they must write the <strong>student's full name</strong> in the reference/narration field so the app can match the payment automatically.
                  </div>
                </div>
              </div>
            )}

            {/* Results after upload */}
            {bankRows.length > 0 && (
              <div>
                {/* Summary */}
                <div style={{ ...grid(4, 2), gap: 14, marginBottom: 20 }}>
                  {[
                    { label: "Total Entries", value: bankRows.length, color: "#3b82f6" },
                    { label: "Auto-Matched", value: Object.keys(bankMatched).filter(k => bankMatched[k]).length, color: "#10b981" },
                    { label: "Unmatched", value: bankRows.length - Object.keys(bankMatched).filter(k => bankMatched[k]).length, color: "#ef4444" },
                    { label: "Total Amount", value: fmtShort(bankRows.reduce((a, r) => a + r.amount, 0)), color: "#f59e0b" },
                  ].map((c, i) => (
                    <div key={i} style={{ ...card, borderTop: `3px solid ${c.color}`, padding: "14px 16px" }}>
                      <div style={{ fontSize: 9, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", marginBottom: 6 }}>{c.label}</div>
                      <div style={{ fontSize: 22, fontWeight: 800, color: "#0f172a" }}>{c.value}</div>
                    </div>
                  ))}
                </div>

                {/* File info + actions */}
                <div style={{ ...card, marginBottom: 16, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a" }}>📄 {bankFileName}</div>
                    <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>{bankRows.length} credit entries found · {currentTerm}</div>
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <button onClick={() => { setBankRows([]); setBankMatched({}); setBankFileName(""); setBankImportDone(false); }}
                      style={{ background: "#f1f5f9", color: "#374151", border: "1px solid #e2e8f0", borderRadius: 9, padding: "9px 14px", fontWeight: 700, fontSize: 12, cursor: "pointer" }}>
                      ✕ Clear
                    </button>
                    {!bankImportDone && (
                      <button onClick={handleBankImport}
                        style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 9, padding: "9px 18px", fontWeight: 800, fontSize: 13, cursor: "pointer" }}>
                        ✓ Import All Matched ({Object.keys(bankMatched).filter(k => bankMatched[k]).length})
                      </button>
                    )}
                    {bankImportDone && (
                      <div style={{ background: "#f0fdf4", color: "#15803d", border: "1px solid #86efac", borderRadius: 9, padding: "9px 14px", fontWeight: 700, fontSize: 12 }}>
                        ✅ Import Complete
                      </div>
                    )}
                  </div>
                </div>

                {/* Rows table */}
                <div style={{ ...card, padding: 0, overflow: "hidden" }}>
                <div style={{ overflowX: "auto" }}>
                <div style={{ minWidth: isMobile ? 760 : "auto" }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1.5fr 1.8fr 0.8fr", gap: 4, padding: "10px 16px", background: "#f8fafc", borderBottom: "1px solid #e8edf3" }}>
                    {["Date", "Amount (UGX)", "Bank Reference / Name", "Matched Student", "Status"].map(h => (
                      <div key={h} style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", letterSpacing: 0.3 }}>{h}</div>
                    ))}
                  </div>

                  {bankRows.map((row, i) => {
                    const matchedStudentId = bankMatched[row.rowIndex];
                    const matchedStudent = matchedStudentId ? termStudents.find(s => s.id === matchedStudentId) : null;
                    const isMatched = !!matchedStudent;

                    return (
                      <div key={row.rowIndex} style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1.5fr 1.8fr 0.8fr", gap: 4, alignItems: "center", padding: "11px 16px", borderTop: "1px solid #f1f5f9", background: isMatched ? "#f0fdf4" : "#fef2f2" }}>
                        {/* Date */}
                        <div style={{ fontSize: 12, color: "#374151", fontWeight: 500 }}>{row.date}</div>

                        {/* Amount */}
                        <div style={{ fontSize: 13, fontWeight: 800, color: "#15803d" }}>{fmt(row.amount)}</div>

                        {/* Reference */}
                        <div style={{ fontSize: 12, color: "#374151" }}>{row.reference || "—"}</div>

                        {/* Student match — dropdown to assign */}
                        <div>
                          <select value={bankMatched[row.rowIndex] || ""}
                            onChange={e => setBankMatched(prev => ({ ...prev, [row.rowIndex]: e.target.value || null }))}
                            style={{ width: "100%", padding: "5px 8px", borderRadius: 7, border: `1px solid ${isMatched ? "#86efac" : "#fca5a5"}`, fontSize: 12, outline: "none", background: isMatched ? "#f0fdf4" : "#fff", fontWeight: isMatched ? 700 : 400 }}>
                            <option value="">— Select Student —</option>
                            {termStudents.map(s => (
                              <option key={s.id} value={s.id}>{s.name} ({s.class})</option>
                            ))}
                          </select>
                        </div>

                        {/* Status */}
                        <div>
                          {isMatched
                            ? <Pill text="✓ Matched" bg="#d1fae5" col="#065f46" />
                            : <Pill text="⚠ Unmatched" bg="#fee2e2" col="#991b1b" />}
                        </div>
                      </div>
                    );
                  })}
                </div>
                </div>
                </div>

                <div style={{ marginTop: 14, padding: "12px 16px", background: "#fffbeb", borderRadius: 10, fontSize: 12, color: "#92400e" }}>
                  <strong>💡 How to fix unmatched rows:</strong> Click the dropdown in the "Matched Student" column and manually select the correct student. Then click "Import All Matched".
                </div>
              </div>
            )}
          </div>
        )}

        {/* ════════ ALUMNI / LEAVERS ════════ */}
        {tab === "alumni" && (() => {
          const alumni = allAlumni[activeSchoolId] || [];
          const debtors = alumni.filter(a => a.outstandingDebt > 0);
          const cleared = alumni.filter(a => !a.outstandingDebt || a.outstandingDebt <= 0);
          const totalDebt = debtors.reduce((a, s) => a + s.outstandingDebt, 0);
          return (
            <div>
              <div style={{ fontSize: 22, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>Alumni & Leavers</div>
              <div style={{ color: "#64748b", fontSize: 13, marginBottom: 20 }}>All historical records preserved — {alumni.length} total leavers from {school.name}</div>

              {/* Summary cards */}
              <div style={{ ...grid(6, 2), gap: 14, marginBottom: 20 }}>
                {[
                  { label: "Total Leavers", value: alumni.length, color: "#8b5cf6", icon: "🎓" },
                  { label: "Graduates (S6)", value: alumni.filter(a => a.status === "Graduate").length, color: "#10b981", icon: "✅" },
                  { label: "Left after S4", value: alumni.filter(a => a.status === "Leaver").length, color: "#f59e0b", icon: "🚶" },
                  { label: "Transferred/Other", value: alumni.filter(a => a.status === "Transferred").length, color: "#7c3aed", icon: "↪️" },
                  { label: "Did Not Return", value: alumni.filter(a => a.status === "Did Not Return").length, color: "#6b7280", icon: "❓" },
                  { label: "Total Debt Owed", value: fmt(totalDebt), color: "#ef4444", icon: "💰" },
                ].map((c, i) => (
                  <div key={i} style={{ ...card, borderTop: `3px solid ${c.color}`, padding: "14px 18px" }}>
                    <div style={{ display: "flex", justifyContent: "space-between" }}>
                      <div style={{ fontSize: 9, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase" }}>{c.label}</div>
                      <span style={{ fontSize: 18 }}>{c.icon}</span>
                    </div>
                    <div style={{ fontSize: 20, fontWeight: 800, color: "#0f172a", marginTop: 6 }}>{c.value}</div>
                  </div>
                ))}
              </div>

              {/* Debt Recovery Section */}
              {debtors.length > 0 && (
                <div style={{ ...card, borderLeft: "4px solid #ef4444", marginBottom: 20 }}>
                  <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a", marginBottom: 14 }}>🚨 Debt Recovery — Outstanding Balances from Leavers</div>
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr style={{ background: "#fef2f2" }}>
                        {["Student", "Left Class", "Year Left", "Reason", "Debt Owed", "Contact", "Action"].map(h => (
                          <th key={h} style={{ padding: "9px 12px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase" }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {debtors.map((a, i) => (
                        <tr key={a.id} style={{ borderTop: "1px solid #fef2f2", background: i % 2 === 0 ? "#fff" : "#fff9f9" }}>
                          <td style={{ padding: "10px 12px" }}>
                            <div style={{ fontWeight: 600, fontSize: 13, color: "#0f172a" }}>{a.name}</div>
                            <div style={{ fontSize: 11, color: "#94a3b8" }}>{a.parent}</div>
                          </td>
                          <td style={{ padding: "10px 12px" }}><Pill text={a.leftClass} /></td>
                          <td style={{ padding: "10px 12px", fontSize: 12, color: "#374151", fontWeight: 600 }}>{a.leftYear}</td>
                          <td style={{ padding: "10px 12px" }}><Pill text={a.leftNote} bg="#fef3c7" col="#92400e" /></td>
                          <td style={{ padding: "10px 12px", fontSize: 13, fontWeight: 800, color: "#dc2626" }}>{fmt(a.outstandingDebt)}</td>
                          <td style={{ padding: "10px 12px", fontSize: 12, color: "#64748b", fontFamily: "monospace" }}>{a.phone}</td>
                          <td style={{ padding: "10px 12px" }}>
                            <button onClick={() => { setShowRecoverDebt(a); setRecoverAmt(""); }} style={{ background: "#ef4444", color: "#fff", border: "none", borderRadius: 7, padding: "5px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                              Collect Debt
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {/* All Alumni List */}
              <div style={{ ...card, padding: 0, overflow: "hidden" }}>
                <div style={{ padding: "14px 18px", background: "#f8fafc", borderBottom: "1px solid #e8edf3", fontWeight: 700, fontSize: 14, color: "#0f172a" }}>
                  📚 Full Alumni Record ({alumni.length} students)
                </div>
                {alumni.length === 0 ? (
                  <div style={{ textAlign: "center", padding: 50, color: "#94a3b8" }}>
                    <div style={{ fontSize: 36, marginBottom: 10 }}>🎓</div>
                    <div style={{ fontSize: 15, fontWeight: 700, color: "#374151", marginBottom: 6 }}>No alumni records yet</div>
                    <div style={{ fontSize: 13 }}>Run "New Academic Year" to promote students. Graduates and leavers will appear here.</div>
                  </div>
                ) : (
                <div style={{ overflowX: "auto" }}>
                <div style={{ minWidth: isMobile ? 800 : "auto" }}>
                {alumni.map((a, i) => {
                  const isOpen = expandedAlumni === a.id;
                  const totalPayments = (a.payments || []).reduce((acc, p) => acc + p.amount, 0);
                  return (
                    <div key={a.id}>
                      <div onClick={() => setExpandedAlumni(isOpen ? null : a.id)}
                        style={{ display: "grid", gridTemplateColumns: "2fr 0.6fr 0.6fr 0.6fr 1fr 1fr 0.6fr", gap: 8, alignItems: "center", padding: "12px 18px", borderTop: i > 0 ? "1px solid #f1f5f9" : "none", background: isOpen ? "#f8f4ff" : i % 2 === 0 ? "#fff" : "#fafbfc", cursor: "pointer" }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          <span style={{ fontSize: 18 }}>{a.gender === "F" ? "👩" : "👨"}</span>
                          <div>
                            <div style={{ fontWeight: 600, fontSize: 13, color: "#0f172a" }}>{a.name}</div>
                            <div style={{ fontSize: 11, color: "#94a3b8" }}>{a.parent} · {a.phone}</div>
                          </div>
                        </div>
                        <div><Pill text={a.leftClass} /></div>
                        <div style={{ fontSize: 12, color: "#374151", fontWeight: 600 }}>{a.leftYear}</div>
                        <div>
                          <Pill text={a.status}
                            bg={a.status === "Graduate" ? "#d1fae5" : a.status === "Did Not Return" ? "#f1f5f9" : a.status === "Transferred" ? "#f5f3ff" : "#fef3c7"}
                            col={a.status === "Graduate" ? "#065f46" : a.status === "Did Not Return" ? "#374151" : a.status === "Transferred" ? "#7c3aed" : "#92400e"} />
                        </div>
                        <div style={{ fontSize: 12, color: "#64748b" }}>{a.leftNote}</div>
                        <div style={{ fontSize: 13, fontWeight: 800, color: a.outstandingDebt > 0 ? "#dc2626" : "#15803d" }}>
                          {a.outstandingDebt > 0 ? `Owes ${fmt(a.outstandingDebt)}` : "✓ Cleared"}
                        </div>
                        <div style={{ fontSize: 12, color: "#94a3b8" }}>{isOpen ? "▲" : "▼"}</div>
                      </div>
                      {isOpen && (
                        <div style={{ background: "#f5f3ff", borderTop: "1px solid #e9d5ff", borderBottom: "1px solid #e9d5ff", padding: "16px 24px" }}>
                          <div style={{ fontWeight: 700, fontSize: 13, color: "#6d28d9", marginBottom: 12 }}>📋 Full Payment History — {a.name}</div>
                          {(a.payments || []).length === 0
                            ? <div style={{ color: "#94a3b8", fontSize: 13 }}>No payments ever recorded.</div>
                            : (
                              <div style={{ position: "relative", paddingLeft: 24 }}>
                                <div style={{ position: "absolute", left: 8, top: 0, bottom: 0, width: 2, background: "#c4b5fd" }} />
                                {a.payments.map((p, pi) => (
                                  <div key={p.id} style={{ position: "relative", marginBottom: 10 }}>
                                    <div style={{ position: "absolute", left: -20, top: 4, width: 10, height: 10, borderRadius: "50%", background: "#7c3aed", border: "2px solid #fff" }} />
                                    <div style={{ background: "#fff", borderRadius: 8, padding: "9px 13px", border: "1px solid #c4b5fd", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                                      <div>
                                        <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 3 }}>
                                          <span style={{ fontFamily: "monospace", fontSize: 11, color: "#7c3aed", fontWeight: 700, background: "#f5f3ff", padding: "1px 6px", borderRadius: 5 }}>{p.id}</span>
                                          <span style={{ fontSize: 12, color: "#374151" }}>{METHOD_ICON[p.method] || "💵"} {p.method}</span>
                                          <span style={{ fontSize: 11, color: "#94a3b8" }}>· {p.term}</span>
                                        </div>
                                        <div style={{ fontSize: 11, color: "#64748b" }}>📅 {fmtDate(p.date)} · {p.receivedBy}</div>
                                      </div>
                                      <div style={{ fontSize: 15, fontWeight: 800, color: "#15803d" }}>{fmt(p.amount)}</div>
                                    </div>
                                  </div>
                                ))}
                              </div>
                            )}
                          <div style={{ display: "flex", gap: 12, marginTop: 10, padding: "9px 13px", background: "#fff", borderRadius: 9, border: "1px solid #c4b5fd", fontSize: 12, color: "#64748b" }}>
                            <span>Total Ever Paid: <strong style={{ color: "#15803d" }}>{fmt(totalPayments)}</strong></span>
                            <span>|</span>
                            <span>Outstanding: <strong style={{ color: a.outstandingDebt > 0 ? "#dc2626" : "#15803d" }}>{fmt(a.outstandingDebt)}</strong></span>
                            <span>|</span>
                            <span>Left: <strong style={{ color: "#0f172a" }}>{a.leftYear}</strong></span>
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
                </div>
                </div>
                )}
              </div>
            </div>
          );
        })()}
      </main>

      {/* ════════ EDIT PAYMENT MODAL ════════ */}
      {showEditPayment && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 250 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 400, maxWidth: 400, boxShadow: "0 24px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>✏️ Edit Payment</div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 18 }}>Correcting a wrong amount for {showEditPayment.student?.name}</div>
            <div style={{ background: "#f8fafc", borderRadius: 11, padding: 13, marginBottom: 16 }}>
              {[
                ["Receipt No.", showEditPayment.payment?.id],
                ["Date", fmtDate(showEditPayment.payment?.date)],
                ["Method", showEditPayment.payment?.method],
                ["Current Amount", fmt(showEditPayment.payment?.amount)],
              ].map(([l, v]) => (
                <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: "1px solid #f1f5f9", fontSize: 13 }}>
                  <span style={{ color: "#64748b" }}>{l}</span>
                  <span style={{ fontWeight: 600, color: "#0f172a" }}>{v}</span>
                </div>
              ))}
            </div>
            <div style={{ marginBottom: 20 }}>
              <label style={lbl}>Corrected Amount (UGX)</label>
              <input type="number" value={editPayAmt} onChange={e => setEditPayAmt(e.target.value)}
                placeholder={`Was: ${showEditPayment.payment?.amount}`}
                style={{ width: "100%", padding: "11px 14px", borderRadius: 9, border: "2px solid #f59e0b", fontSize: 16, fontWeight: 800, outline: "none", boxSizing: "border-box" }} />
              <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 5 }}>⚠ This corrects a data entry error. Use this only if the wrong amount was typed.</div>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => { setShowEditPayment(null); setEditPayAmt(""); }}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
              <button onClick={handleEditPayment}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#f59e0b", color: "#0f172a", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Save Correction</button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ EDIT STUDENT MODAL ════════ */}
      {showEditStudent && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 300, padding: 16 }}>
          <div style={{ background: "#fff", borderRadius: 16, padding: isMobile ? 18 : 26, width: isMobile ? "calc(100vw - 32px)" : 480, maxWidth: 480, boxShadow: "0 24px 60px rgba(0,0,0,0.25)", maxHeight: "90vh", overflowY: "auto" }}>
            <div style={{ fontSize: 16, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>✏ Edit Student</div>
            <div style={{ fontSize: 12, color: "#64748b", marginBottom: 20 }}>Update {showEditStudent.name}'s details</div>

            {[
              { label: "Full Name", key: "name", type: "text", required: true },
              { label: "Parent / Guardian Name", key: "parent", type: "text" },
              { label: "Phone Number", key: "phone", type: "tel" },
            ].map(({ label, key, type, required }) => (
              <div key={key} style={{ marginBottom: 14 }}>
                <label style={lbl}>{label}{required && " *"}</label>
                <input type={type} value={showEditStudent[key] || ""} onChange={e => setShowEditStudent(prev => ({ ...prev, [key]: e.target.value }))}
                  style={{ width: "100%", padding: "10px 14px", borderRadius: 9, border: "1px solid #e2e8f0", fontSize: 13, outline: "none", boxSizing: "border-box" }} />
              </div>
            ))}

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 14 }}>
              <div>
                <label style={lbl}>Class</label>
                <select value={showEditStudent.class || ""} onChange={e => setShowEditStudent(prev => ({ ...prev, class: e.target.value, stream: "" }))}
                  style={{ width: "100%", padding: "10px 12px", borderRadius: 9, border: "1px solid #e2e8f0", fontSize: 13, outline: "none", background: "#fff" }}>
                  {schoolClasses.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              <div>
                <label style={lbl}>Stream</label>
                <select value={showEditStudent.stream || ""} onChange={e => setShowEditStudent(prev => ({ ...prev, stream: e.target.value }))}
                  style={{ width: "100%", padding: "10px 12px", borderRadius: 9, border: "1px solid #e2e8f0", fontSize: 13, outline: "none", background: "#fff" }}>
                  <option value="">No Stream</option>
                  {getClassStreams(showEditStudent.class || "").map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
              <div>
                <label style={lbl}>Gender</label>
                <select value={showEditStudent.gender || "M"} onChange={e => setShowEditStudent(prev => ({ ...prev, gender: e.target.value }))}
                  style={{ width: "100%", padding: "10px 12px", borderRadius: 9, border: "1px solid #e2e8f0", fontSize: 13, outline: "none", background: "#fff" }}>
                  <option value="M">Male</option>
                  <option value="F">Female</option>
                </select>
              </div>
              <div>
                <label style={lbl}>Category</label>
                <select value={showEditStudent.category || "Day Scholar"} onChange={e => setShowEditStudent(prev => ({ ...prev, category: e.target.value }))}
                  style={{ width: "100%", padding: "10px 12px", borderRadius: 9, border: "1px solid #e2e8f0", fontSize: 13, outline: "none", background: "#fff" }}>
                  <option value="Day Scholar">Day Scholar</option>
                  <option value="Boarder">Boarder</option>
                </select>
              </div>
            </div>

            <div style={{ display: "flex", gap: 10, marginTop: 4 }}>
              <button onClick={() => setShowEditStudent(null)}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
              <button onClick={handleSaveEditStudent}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#0f172a", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Save Changes</button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ GENERIC CONFIRM DIALOG ════════ */}
      {confirmDialog && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 300, padding: 16 }}>
          <div style={{ background: "#fff", borderRadius: 16, padding: isMobile ? 18 : 26, width: isMobile ? "calc(100vw - 32px)" : 400, maxWidth: 400, boxShadow: "0 24px 60px rgba(0,0,0,0.25)" }}>
            <div style={{ fontSize: 16, fontWeight: 800, color: "#0f172a", marginBottom: 8 }}>{confirmDialog.title}</div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 22, lineHeight: 1.6 }}>{confirmDialog.message}</div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => setConfirmDialog(null)}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
              <button onClick={() => { const fn = confirmDialog.onConfirm; setConfirmDialog(null); fn(); }}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: confirmDialog.danger ? "#dc2626" : "#0f172a", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>
                {confirmDialog.danger ? "Yes, Continue" : "Confirm"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ I'VE SENT PAYMENT MODAL ════════ */}
      {showPaymentConfirm && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 250, padding: 16 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 440, maxWidth: 440, boxShadow: "0 24px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>✓ Confirm Your Payment</div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 18 }}>Let us know you've sent the money so we can confirm it quickly. Your account will still reactivate automatically once we verify the transaction.</div>

            <div style={{ background: "#f8fafc", borderRadius: 10, padding: 12, marginBottom: 14, fontSize: 12, color: "#374151" }}>
              Reference code: <strong style={{ fontFamily: "monospace" }}>{school.billingRef}</strong>
            </div>

            <div style={{ marginBottom: 14 }}>
              <label style={lbl}>Payment Method</label>
              <div style={{ display: "flex", gap: 8 }}>
                {["MTN MoMo", "Bank Transfer"].map(m => (
                  <button key={m} onClick={() => setPaymentConfirmForm(p => ({ ...p, method: m }))}
                    style={{ flex: 1, padding: "9px 12px", borderRadius: 9, border: `2px solid ${paymentConfirmForm.method === m ? "#10b981" : "#e2e8f0"}`, background: paymentConfirmForm.method === m ? "#f0fdf4" : "#fff", color: paymentConfirmForm.method === m ? "#15803d" : "#94a3b8", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                    {m === "MTN MoMo" ? "📱 MTN MoMo" : "🏦 Bank Transfer"}
                  </button>
                ))}
              </div>
            </div>

            <div style={{ marginBottom: 14 }}>
              <label style={lbl}>Amount Sent (UGX)</label>
              <input type="number" value={paymentConfirmForm.amount} onChange={e => setPaymentConfirmForm(p => ({ ...p, amount: e.target.value }))}
                placeholder={String(getBillingInfo(school.plan, school.billingCycle, school.customPrice).price)} style={inp} />
              <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 4 }}>Expected: {fmt(getBillingInfo(school.plan, school.billingCycle, school.customPrice).price)} for {school.plan} plan. Pay a different plan's price to switch plans too.</div>
            </div>

            <div style={{ marginBottom: 14 }}>
              <label style={lbl}>Date Sent</label>
              <input type="date" value={paymentConfirmForm.date} onChange={e => setPaymentConfirmForm(p => ({ ...p, date: e.target.value }))} style={inp} />
            </div>

            <div style={{ marginBottom: 20 }}>
              <label style={lbl}>Note (optional)</label>
              <input value={paymentConfirmForm.note} onChange={e => setPaymentConfirmForm(p => ({ ...p, note: e.target.value }))}
                placeholder="e.g. Transaction ID, sender's name" style={inp} />
            </div>

            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => setShowPaymentConfirm(false)}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
              <button onClick={submitPaymentNotice}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#10b981", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Submit</button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ ALUMNI DEBT RECOVERY MODAL ════════ */}
      {showRecoverDebt && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 250, padding: 16 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 420, maxWidth: 420, boxShadow: "0 24px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>💰 Collect Debt Payment</div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 18 }}>
              Record a payment from <strong>{showRecoverDebt.name}</strong> ({showRecoverDebt.parent}) towards their outstanding balance from {showRecoverDebt.leftClass}, {showRecoverDebt.leftYear}.
            </div>

            <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 10, padding: 12, marginBottom: 16, textAlign: "center" }}>
              <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", marginBottom: 4 }}>Total Owed</div>
              <div style={{ fontSize: 22, fontWeight: 800, color: "#dc2626" }}>{fmt(showRecoverDebt.outstandingDebt)}</div>
            </div>

            <div style={{ marginBottom: 20 }}>
              <label style={lbl}>Amount Received (UGX)</label>
              <input type="number" value={recoverAmt} onChange={e => setRecoverAmt(e.target.value)}
                placeholder={`Max ${showRecoverDebt.outstandingDebt}`} style={inp} autoFocus />
              {recoverAmt && parseInt(recoverAmt) > 0 && (
                <div style={{ fontSize: 12, color: "#64748b", marginTop: 6 }}>
                  Remaining balance after this payment: <strong>{fmt(Math.max(0, showRecoverDebt.outstandingDebt - (parseInt(recoverAmt.replace(/,/g, "")) || 0)))}</strong>
                </div>
              )}
            </div>

            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => { setShowRecoverDebt(null); setRecoverAmt(""); }}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
              <button onClick={handleRecoverDebt}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#ef4444", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Record Payment</button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ BULK MOVE TO ALUMNI MODAL ════════ */}
      {showBulkAlumni && (() => {
        const selectedIds = Object.keys(bulkAlumniSelected).filter(id => bulkAlumniSelected[id]);
        const selectedCount = selectedIds.length;
        const totalDebt = termStudents.filter(s => selectedIds.includes(String(s.id)))
          .reduce((a, s) => a + getBalance(s, currentTerm).balance, 0);
        const allSelected = termStudents.length > 0 && selectedCount === termStudents.length;
        return (
          <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 250, padding: 16 }}>
            <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 560, maxWidth: 560, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 24px 60px rgba(0,0,0,0.25)" }}>
              <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>🎓 Bulk Move to Alumni</div>
              <div style={{ fontSize: 13, color: "#64748b", marginBottom: 16 }}>Select multiple students to move to Alumni at once — useful for transfers, dropouts, or end-of-term cleanups.</div>

              <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
                <button onClick={() => setBulkAlumniMode("checklist")} style={{ flex: 1, padding: 9, borderRadius: 9, border: `2px solid ${bulkAlumniMode === "checklist" ? "#7c3aed" : "#e2e8f0"}`, background: bulkAlumniMode === "checklist" ? "#f5f3ff" : "#fff", color: bulkAlumniMode === "checklist" ? "#7c3aed" : "#94a3b8", fontWeight: 700, fontSize: 12, cursor: "pointer" }}>✓ Checklist (in-app)</button>
                <button onClick={() => setBulkAlumniMode("excel")} style={{ flex: 1, padding: 9, borderRadius: 9, border: `2px solid ${bulkAlumniMode === "excel" ? "#7c3aed" : "#e2e8f0"}`, background: bulkAlumniMode === "excel" ? "#f5f3ff" : "#fff", color: bulkAlumniMode === "excel" ? "#7c3aed" : "#94a3b8", fontWeight: 700, fontSize: 12, cursor: "pointer" }}>📊 Upload Excel</button>
              </div>

              <div style={{ background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 10, padding: 12, marginBottom: 14, fontSize: 12, color: "#78350f", lineHeight: 1.7 }}>
                <strong>⚠ This moves selected students out of the active class list immediately.</strong> Full payment history is preserved, and any outstanding balances are tracked in Alumni Debt Recovery.
              </div>

              {bulkAlumniMode === "checklist" ? (
                <>
              <div style={{ marginBottom: 14 }}>
                <label style={lbl}>Reason for Leaving (applies to all selected)</label>
                <select value={bulkAlumniReason} onChange={e => setBulkAlumniReason(e.target.value)} style={inp}>
                  <option>Transferred to another school</option>
                  <option>Dropped out</option>
                  <option>Withdrawn by parent</option>
                  <option>Expelled / Disciplinary</option>
                  <option>Deceased</option>
                  <option>Other</option>
                </select>
              </div>

              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: "#374151" }}>
                  {selectedCount} of {termStudents.length} selected
                  {selectedCount > 0 && totalDebt > 0 && <span style={{ color: "#dc2626", marginLeft: 8 }}>· {fmt(totalDebt)} total outstanding</span>}
                </div>
                <button onClick={() => {
                  if (allSelected) { setBulkAlumniSelected({}); }
                  else { const all = {}; termStudents.forEach(s => { all[String(s.id)] = true; }); setBulkAlumniSelected(all); }
                }} style={{ background: "#f1f5f9", color: "#374151", border: "none", borderRadius: 7, padding: "5px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                  {allSelected ? "Deselect All" : "Select All"}
                </button>
              </div>

              <div style={{ border: "1px solid #e8edf3", borderRadius: 10, overflow: "hidden", marginBottom: 18, maxHeight: 320, overflowY: "auto" }}>
                {schoolClasses.map(cls => {
                  const clsStudents = termStudents.filter(s => s.class === cls);
                  if (clsStudents.length === 0) return null;
                  return (
                    <div key={cls}>
                      <div style={{ padding: "6px 12px", background: "#f8fafc", fontSize: 11, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", position: "sticky", top: 0 }}>{cls}</div>
                      {clsStudents.map(s => {
                        const checked = !!bulkAlumniSelected[String(s.id)];
                        const debt = getBalance(s, currentTerm).balance;
                        return (
                          <label key={s.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, padding: "8px 12px", borderTop: "1px solid #f1f5f9", cursor: "pointer", background: checked ? "#f5f3ff" : "#fff" }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                              <input type="checkbox" checked={checked} onChange={e => setBulkAlumniSelected(prev => ({ ...prev, [String(s.id)]: e.target.checked }))} style={{ width: 16, height: 16, cursor: "pointer" }} />
                              <span style={{ fontSize: 13, fontWeight: 600, color: "#0f172a" }}>{s.name}</span>
                              {s.stream && <Pill text={s.stream} bg="#f5f3ff" col="#7c3aed" />}
                            </div>
                            {debt > 0 && <span style={{ fontSize: 11, color: "#f59e0b", fontWeight: 600 }}>owes {fmt(debt)}</span>}
                          </label>
                        );
                      })}
                    </div>
                  );
                })}
                {termStudents.length === 0 && <div style={{ textAlign: "center", padding: 30, color: "#94a3b8", fontSize: 13 }}>No students to display</div>}
              </div>

              <div style={{ display: "flex", gap: 10 }}>
                <button onClick={() => { setShowBulkAlumni(false); setBulkAlumniSelected({}); }}
                  style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
                <button onClick={handleBulkMoveToAlumni} disabled={selectedCount === 0}
                  style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: selectedCount > 0 ? "#7c3aed" : "#e2e8f0", color: selectedCount > 0 ? "#fff" : "#94a3b8", fontSize: 13, fontWeight: 800, cursor: selectedCount > 0 ? "pointer" : "not-allowed" }}>
                  🎓 Move {selectedCount > 0 ? selectedCount : ""} to Alumni
                </button>
              </div>
                </>
              ) : (
                <>
                  {bulkAlumniExcelRows.length === 0 ? (
                    <div>
                      <div style={{ background: "#f8fafc", borderRadius: 12, padding: 16, marginBottom: 16 }}>
                        <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 10 }}>How it works:</div>
                        <div style={{ fontSize: 12, color: "#64748b", lineHeight: 1.8 }}>
                          1. Download the template — pre-filled with your current student list<br />
                          2. Type "Yes" in the "Move? (Yes/No)" column for every student leaving, and optionally a Reason<br />
                          3. Upload the file — we'll match each row to a student and show a preview<br />
                          4. Review, then click "Move to Alumni"
                        </div>
                      </div>

                      <div style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
                        <button onClick={downloadBulkAlumniTemplate} style={{ flex: "1 1 45%", background: "#f1f5f9", color: "#374151", border: "1px solid #e2e8f0", borderRadius: 9, padding: "11px 14px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                          📄 Download Template
                        </button>
                        <button onClick={() => bulkAlumniFileRef.current?.click()} style={{ flex: "1 1 45%", background: "#7c3aed", color: "#fff", border: "none", borderRadius: 9, padding: "11px 14px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                          📂 Choose File to Upload
                        </button>
                        <input ref={bulkAlumniFileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleBulkAlumniFileUpload} style={{ display: "none" }} />
                      </div>

                      <div style={{ display: "flex", gap: 10 }}>
                        <button onClick={() => setShowBulkAlumni(false)}
                          style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Close</button>
                      </div>
                    </div>
                  ) : (
                    <div>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
                        <div>
                          <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a" }}>📄 {bulkAlumniExcelFileName}</div>
                          <div style={{ fontSize: 12, color: "#64748b" }}>{bulkAlumniExcelRows.filter(r => r.valid).length} student(s) will be moved to Alumni</div>
                        </div>
                        {!bulkAlumniExcelDone && (
                          <button onClick={() => { setBulkAlumniExcelRows([]); setBulkAlumniExcelFileName(""); }} style={{ background: "#f1f5f9", color: "#64748b", border: "none", borderRadius: 8, padding: "7px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>✕ Clear</button>
                        )}
                      </div>

                      {bulkAlumniExcelRows.some(r => r.matchStatus !== "matched" && r.shouldMove) && (
                        <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 10, padding: 12, marginBottom: 12 }}>
                          <div style={{ fontSize: 12, fontWeight: 700, color: "#b91c1c", marginBottom: 6 }}>⚠ Some rows need attention:</div>
                          {bulkAlumniExcelRows.filter(r => r.matchStatus === "not_found" && r.shouldMove).map((r, i) => (
                            <div key={"nf" + i} style={{ fontSize: 11, color: "#7f1d1d", padding: "2px 0" }}>• "{r.name}" — no matching student found, will be skipped</div>
                          ))}
                          {bulkAlumniExcelRows.filter(r => r.matchStatus === "ambiguous" && r.shouldMove).map((r, i) => (
                            <div key={"am" + i} style={{ fontSize: 11, color: "#7f1d1d", padding: "2px 0" }}>• "{r.name}" — multiple students with this name; will be skipped</div>
                          ))}
                        </div>
                      )}

                      <div style={{ border: "1px solid #e8edf3", borderRadius: 10, overflow: "hidden", marginBottom: 16, maxHeight: 300, overflowY: "auto" }}>
                        <div style={{ overflowX: "auto" }}>
                        <div style={{ minWidth: 480 }}>
                          <div style={{ display: "grid", gridTemplateColumns: "1.6fr 0.6fr 1.2fr 0.8fr", gap: 4, padding: "8px 12px", background: "#f8fafc", borderBottom: "1px solid #e8edf3", position: "sticky", top: 0 }}>
                            {["Name", "Class", "Reason", "Status"].map(h => (
                              <div key={h} style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase" }}>{h}</div>
                            ))}
                          </div>
                          {bulkAlumniExcelRows.map(r => (
                            <div key={r.rowIndex} style={{ display: "grid", gridTemplateColumns: "1.6fr 0.6fr 1.2fr 0.8fr", gap: 4, alignItems: "center", padding: "7px 12px", borderTop: "1px solid #f1f5f9", background: r.valid ? "#f5f3ff" : "#fff", fontSize: 12 }}>
                              <div style={{ fontWeight: 600, color: "#0f172a" }}>{r.name}</div>
                              <div>{r.matchedStudent ? <Pill text={classLabel(r.matchedStudent.class, r.matchedStudent.stream)} /> : "—"}</div>
                              <div style={{ fontSize: 11, color: "#64748b" }}>{r.reason || (r.shouldMove ? bulkAlumniReason : "—")}</div>
                              <div>
                                {!r.shouldMove && <Pill text="Stay" bg="#f1f5f9" col="#94a3b8" />}
                                {r.shouldMove && r.matchStatus === "matched" && <Pill text="🎓 Move" bg="#f5f3ff" col="#7c3aed" />}
                                {r.shouldMove && r.matchStatus === "not_found" && <Pill text="✗ Not found" bg="#fef2f2" col="#dc2626" />}
                                {r.shouldMove && r.matchStatus === "ambiguous" && <Pill text="⚠ Ambiguous" bg="#fffbeb" col="#92400e" />}
                              </div>
                            </div>
                          ))}
                        </div>
                        </div>
                      </div>

                      {bulkAlumniExcelDone ? (
                        <div style={{ background: "#f5f3ff", border: "1px solid #ddd6fe", borderRadius: 10, padding: 14, textAlign: "center", color: "#7c3aed", fontWeight: 700, fontSize: 13 }}>
                          ✅ {bulkAlumniExcelRows.filter(r => r.valid).length} student(s) moved to Alumni successfully!
                        </div>
                      ) : (
                        <div style={{ display: "flex", gap: 10 }}>
                          <button onClick={() => { setShowBulkAlumni(false); setBulkAlumniExcelRows([]); setBulkAlumniExcelFileName(""); }}
                            style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
                          <button onClick={handleBulkAlumniExcelImport} disabled={bulkAlumniExcelRows.filter(r => r.valid).length === 0}
                            style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: bulkAlumniExcelRows.filter(r => r.valid).length > 0 ? "#7c3aed" : "#e2e8f0", color: bulkAlumniExcelRows.filter(r => r.valid).length > 0 ? "#fff" : "#94a3b8", fontSize: 13, fontWeight: 800, cursor: bulkAlumniExcelRows.filter(r => r.valid).length > 0 ? "pointer" : "not-allowed" }}>
                            🎓 Move to Alumni ({bulkAlumniExcelRows.filter(r => r.valid).length})
                          </button>
                        </div>
                      )}
                      {bulkAlumniExcelDone && (
                        <div style={{ display: "flex", gap: 10, marginTop: 10 }}>
                          <button onClick={() => { setShowBulkAlumni(false); setBulkAlumniExcelRows([]); setBulkAlumniExcelFileName(""); setBulkAlumniExcelDone(false); }}
                            style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#0f172a", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Done</button>
                        </div>
                      )}
                    </div>
                  )}
                </>
              )}
            </div>
          </div>
        );
      })()}

      {/* ════════ MOVE TO ALUMNI MODAL (manual mid-term transfer) ════════ */}
      {showMoveAlumni && (() => {
        const debt = getBalance(showMoveAlumni, currentTerm).balance;
        return (
          <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 250 }}>
            <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 420, maxWidth: 420, boxShadow: "0 24px 60px rgba(0,0,0,0.2)" }}>
              <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>🎓 Move to Alumni</div>
              <div style={{ fontSize: 13, color: "#64748b", marginBottom: 18 }}>{showMoveAlumni.name} · {showMoveAlumni.class} · {showMoveAlumni.parent}</div>

              <div style={{ background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 10, padding: 14, marginBottom: 18, fontSize: 12, color: "#78350f", lineHeight: 1.7 }}>
                <strong>⚠ This moves the student out of the active class list immediately.</strong><br />
                • Full payment history is preserved<br />
                • {debt > 0 ? <>Outstanding balance of <strong>{fmt(debt)}</strong> will be tracked in Alumni Debt Recovery</> : "No outstanding balance — fully cleared"}<br />
                • Use this for mid-term transfers, withdrawals, or dropouts
              </div>

              <div style={{ marginBottom: 20 }}>
                <label style={lbl}>Reason for Leaving</label>
                <select value={moveAlumniReason} onChange={e => setMoveAlumniReason(e.target.value)} style={inp}>
                  <option>Transferred to another school</option>
                  <option>Dropped out</option>
                  <option>Withdrawn by parent</option>
                  <option>Expelled / Disciplinary</option>
                  <option>Deceased</option>
                  <option>Other</option>
                </select>
              </div>

              <div style={{ display: "flex", gap: 10 }}>
                <button onClick={() => setShowMoveAlumni(null)}
                  style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
                <button onClick={handleMoveToAlumni}
                  style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#7c3aed", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>🎓 Move to Alumni</button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* ════════ ADD FEE ITEM MODAL ════════ */}
      {showAddFeeItem && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 250 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 400, maxWidth: 400, boxShadow: "0 24px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>+ Add Fee Item</div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 20 }}>Add a new fee component to the <strong>{showAddFeeItem}</strong> structure — applied to all classes (S1–S6)</div>
            <div style={{ marginBottom: 14 }}>
              <label style={lbl}>Item Name *</label>
              <input value={newFeeItemName} onChange={e => setNewFeeItemName(e.target.value)}
                placeholder="e.g. Exam Fee, Transport, Computer Lab" style={inp} />
            </div>
            <div style={{ marginBottom: 18 }}>
              <label style={lbl}>Amount per Term (UGX) *</label>
              <input type="number" value={newFeeItemAmt} onChange={e => setNewFeeItemAmt(e.target.value)}
                placeholder="e.g. 30000" style={{ ...inp, fontSize: 16, fontWeight: 800 }} />
              <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 5 }}>This amount will be added for S1 through S6. You can edit individual class amounts afterwards by clicking the cell.</div>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => setShowAddFeeItem(null)}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
              <button onClick={handleAddFeeItem}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#3b82f6", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Add Fee Item</button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ ADD REQUIREMENT MODAL ════════ */}
      {showAddReq && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 250 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 420, maxWidth: 420, boxShadow: "0 24px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>📋 Add School Requirement</div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 20 }}>Add an item students must bring or pay for</div>
            <div style={{ marginBottom: 14 }}>
              <label style={lbl}>Item Name *</label>
              <input value={newReq.name} onChange={e => setNewReq(p => ({ ...p, name: e.target.value }))}
                placeholder="e.g. School Uniform, Mattress, Lab Fee" style={inp} />
            </div>
            <div style={{ marginBottom: 14 }}>
              <label style={lbl}>Cost (UGX)</label>
              <input type="number" value={newReq.cost} onChange={e => setNewReq(p => ({ ...p, cost: e.target.value }))}
                placeholder="e.g. 80000" style={inp} />
            </div>
            <div style={{ marginBottom: 14 }}>
              <label style={lbl}>Applies To</label>
              <div style={{ display: "flex", gap: 8 }}>
                {STUDENT_CATEGORIES.map(cat => (
                  <button key={cat} onClick={() => setNewReq(p => ({
                    ...p,
                    appliesTo: p.appliesTo.includes(cat) ? p.appliesTo.filter(a => a !== cat) : [...p.appliesTo, cat]
                  }))} style={{ flex: 1, padding: "9px 12px", borderRadius: 9, border: `2px solid ${newReq.appliesTo.includes(cat) ? "#3b82f6" : "#e2e8f0"}`, background: newReq.appliesTo.includes(cat) ? "#eff6ff" : "#fff", color: newReq.appliesTo.includes(cat) ? "#2563eb" : "#94a3b8", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                    {cat}
                  </button>
                ))}
              </div>
            </div>
            <div style={{ marginBottom: 22 }}>
              <label style={lbl}>Add to Fee?</label>
              <div style={{ display: "flex", gap: 8 }}>
                {[["mandatory", "✓ Mandatory — add to fee", "#10b981", "#f0fdf4"], ["optional", "○ Optional — checklist only", "#f59e0b", "#fffbeb"]].map(([val, label, col, bg]) => (
                  <button key={val} onClick={() => setNewReq(p => ({ ...p, mandatory: val === "mandatory" }))}
                    style={{ flex: 1, padding: "9px 8px", borderRadius: 9, border: `2px solid ${(val === "mandatory") === newReq.mandatory ? col : "#e2e8f0"}`, background: (val === "mandatory") === newReq.mandatory ? bg : "#fff", color: (val === "mandatory") === newReq.mandatory ? col : "#94a3b8", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                    {label}
                  </button>
                ))}
              </div>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => setShowAddReq(false)}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
              <button onClick={handleAddReq}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#3b82f6", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Add Requirement</button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ PHOTO UPLOAD MODAL ════════ */}
      {showPhotoUpload && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 300 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 28, width: isMobile ? "calc(100vw - 32px)" : 420, maxWidth: 420, boxShadow: "0 24px 60px rgba(0,0,0,0.25)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>📷 {photoUploadType === "staff" ? "Worker Photo" : "Student Photo"}</div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 20 }}>{showPhotoUpload.name} · {photoUploadType === "staff" ? showPhotoUpload.role : showPhotoUpload.class}</div>

            {/* Current photo preview */}
            <div style={{ display: "flex", justifyContent: "center", marginBottom: 20 }}>
              {cameraActive ? (
                <div style={{ position: "relative" }}>
                  <video ref={videoRef} autoPlay playsInline style={{ width: isMobile ? "100%" : 280, maxWidth: 280, height: 210, borderRadius: 12, objectFit: "cover", background: "#000" }} />
                  <canvas ref={canvasRef} style={{ display: "none" }} />
                </div>
              ) : (
                showPhotoUpload.photo
                  ? <img src={showPhotoUpload.photo} alt="Current" style={{ width: 140, height: 140, borderRadius: 14, objectFit: "cover", border: "3px solid #e2e8f0" }} />
                  : <div style={{ width: 140, height: 140, borderRadius: 14, background: photoUploadType === "staff" ? "#e0e7ff" : (showPhotoUpload.gender === "F" ? "#fce7f3" : "#dbeafe"), display: "flex", alignItems: "center", justifyContent: "center", fontSize: 64 }}>
                      {photoUploadType === "staff" ? "👷" : (showPhotoUpload.gender === "F" ? "👩" : "👨")}
                    </div>
              )}
            </div>

            {/* Action buttons */}
            {cameraActive ? (
              <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
                <button onClick={capturePhoto} style={{ flex: 1, padding: 12, borderRadius: 10, border: "none", background: "#10b981", color: "#fff", fontSize: 14, fontWeight: 800, cursor: "pointer" }}>
                  📸 Capture Photo
                </button>
                <button onClick={stopCamera} style={{ flex: 1, padding: 12, borderRadius: 10, border: "1px solid #e2e8f0", background: "#fff", color: "#64748b", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
                  Cancel Camera
                </button>
              </div>
            ) : (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 14 }}>
                {/* Upload from file */}
                <button onClick={() => fileInputRef.current?.click()} style={{ padding: 12, borderRadius: 10, border: "2px solid #3b82f6", background: "#eff6ff", color: "#2563eb", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                  📁 Upload from Phone
                </button>
                <input ref={fileInputRef} type="file" accept="image/*" onChange={handleFilePhoto} style={{ display: "none" }} />

                {/* Take photo with camera */}
                <button onClick={startCamera} style={{ padding: 12, borderRadius: 10, border: "2px solid #f59e0b", background: "#fffbeb", color: "#92400e", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                  📷 Take Photo
                </button>
              </div>
            )}

            {/* Remove + close */}
            <div style={{ display: "flex", gap: 10 }}>
              {showPhotoUpload.photo && !cameraActive && (
                <button onClick={() => removePhoto(showPhotoUpload.id)} style={{ flex: 1, padding: 10, borderRadius: 9, border: "none", background: "#fef2f2", color: "#dc2626", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                  🗑 Remove Photo
                </button>
              )}
              <button onClick={() => { stopCamera(); setShowPhotoUpload(null); }} style={{ flex: 1, padding: 10, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", color: "#64748b", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ RECEIPT MODAL ════════ */}
      {showReceipt && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 300 }}>
          <div ref={receiptRef} id="feetrack-print-receipt" style={{ background: "#fff", borderRadius: 18, width: 410, overflow: "hidden", boxShadow: "0 24px 60px rgba(0,0,0,0.25)" }}>
            <div style={{ background: "#0f172a", padding: "18px 22px", textAlign: "center" }}>
              <div style={{ fontSize: 24 }}>🏫</div>
              <div style={{ color: "#f1f5f9", fontWeight: 800, fontSize: 14, marginTop: 4 }}>{showReceipt.school?.name}</div>
              <div style={{ color: "#64748b", fontSize: 11, marginTop: 2 }}>OFFICIAL FEE RECEIPT</div>
            </div>
            <div style={{ padding: "18px 22px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 14, paddingBottom: 12, borderBottom: "2px dashed #e2e8f0" }}>
                <div><div style={{ fontSize: 10, color: "#94a3b8", fontWeight: 600 }}>RECEIPT NO.</div>
                  <div style={{ fontFamily: "monospace", fontSize: 16, fontWeight: 800, color: "#3b82f6" }}>{showReceipt.payment?.id}</div></div>
                <div style={{ textAlign: "right" }}><div style={{ fontSize: 10, color: "#94a3b8", fontWeight: 600 }}>DATE</div>
                  <div style={{ fontSize: 13, fontWeight: 700 }}>{fmtDate(showReceipt.payment?.date)}</div></div>
              </div>
              {[
                ["Student", showReceipt.student?.name],
                ["Class", showReceipt.student?.class],
                ["Parent / Guardian", showReceipt.student?.parent],
                ["Term", currentTerm],
                ["Payment Method", `${METHOD_ICON[showReceipt.payment?.method]} ${showReceipt.payment?.method}`],
                ["Received By", showReceipt.payment?.receivedBy],
                ["Remaining Balance", fmt(showReceipt.newBalance)],
              ].map(([label, val]) => (
                <div key={label} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderBottom: "1px solid #f8fafc", fontSize: 13 }}>
                  <span style={{ color: "#64748b" }}>{label}</span>
                  <span style={{ fontWeight: 600, color: "#0f172a" }}>{val}</span>
                </div>
              ))}
              <div style={{ background: "#f0fdf4", borderRadius: 10, padding: "14px", marginTop: 14, textAlign: "center" }}>
                <div style={{ fontSize: 11, color: "#16a34a", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 3 }}>Amount Paid</div>
                <div style={{ fontSize: 28, fontWeight: 800, color: "#15803d" }}>{fmt(showReceipt.payment?.amount)}</div>
              </div>
              <div style={{ marginTop: 12, textAlign: "center", fontSize: 11, color: "#94a3b8" }}>This is an official receipt. Please keep it safe.</div>
            </div>
            <div style={{ padding: "0 22px 18px", display: "flex", gap: 8 }}>
              <button onClick={printReceipt} style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#10b981", color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>🖨️ Print Receipt</button>
              <button onClick={() => setShowReceipt(null)} style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#0f172a", color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>Close</button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ BALANCES REPORT MODAL (printable) ════════ */}
      {showBalancesReport && (() => {
        const data = getBalancesReportData();
        return (
          <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 300, padding: isMobile ? 8 : 20 }}>
            <div style={{ background: "#fff", borderRadius: 14, width: "100%", maxWidth: 900, maxHeight: "92vh", overflowY: "auto", boxShadow: "0 24px 60px rgba(0,0,0,0.3)" }}>
              <div id="feetrack-print-balances" style={{ padding: isMobile ? 16 : 28 }}>
                <div style={{ textAlign: "center", borderBottom: "3px solid #0f172a", paddingBottom: 14, marginBottom: 16 }}>
                  <div style={{ fontSize: 20, fontWeight: 800, color: "#0f172a" }}>🏫 {school.name}</div>
                  <div style={{ fontSize: 13, color: "#64748b", marginTop: 2 }}>Student Fee Balances Report — {currentTerm}</div>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "#64748b", marginBottom: 16 }}>
                  <span>Filter: {data.filterLabel}</span>
                  <span>Generated: {fmtDate(new Date().toISOString())}</span>
                </div>

                {data.sections.length === 0 ? (
                  <div style={{ textAlign: "center", padding: 30, color: "#94a3b8", fontSize: 13 }}>No students match the current filter.</div>
                ) : data.sections.map(sec => (
                  <div key={sec.key} style={{ marginBottom: 22, pageBreakInside: "avoid" }}>
                    <div style={{ fontSize: 14, fontWeight: 800, background: "#f1f5f9", padding: "8px 12px", borderRadius: "8px 8px 0 0" }}>
                      {sec.key} <span style={{ fontWeight: 400, color: "#94a3b8", fontSize: 12 }}>({sec.rows.length} students)</span>
                    </div>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                      <thead>
                        <tr style={{ background: "#fafbfc" }}>
                          {["Name", "Parent", "Phone", "Fee Due", "Paid", "Balance", "Status"].map((h, i) => (
                            <th key={h} style={{ padding: "6px 10px", textAlign: i >= 3 && i <= 5 ? "right" : "left", fontSize: 10, textTransform: "uppercase", color: "#94a3b8", borderBottom: "2px solid #e2e8f0" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {sec.rows.map((r, i) => (
                          <tr key={i}>
                            <td style={{ padding: "6px 10px", borderBottom: "1px solid #e2e8f0" }}>{r.name}</td>
                            <td style={{ padding: "6px 10px", borderBottom: "1px solid #e2e8f0" }}>{r.parent}</td>
                            <td style={{ padding: "6px 10px", borderBottom: "1px solid #e2e8f0" }}>{r.phone}</td>
                            <td style={{ padding: "6px 10px", borderBottom: "1px solid #e2e8f0", textAlign: "right", fontFamily: "monospace" }}>{fmt(r.due)}</td>
                            <td style={{ padding: "6px 10px", borderBottom: "1px solid #e2e8f0", textAlign: "right", fontFamily: "monospace" }}>{fmt(r.paid)}</td>
                            <td style={{ padding: "6px 10px", borderBottom: "1px solid #e2e8f0", textAlign: "right", fontFamily: "monospace", color: r.balance > 0 ? "#dc2626" : "#0f172a", fontWeight: r.balance > 0 ? 700 : 400 }}>{fmt(r.balance)}</td>
                            <td style={{ padding: "6px 10px", borderBottom: "1px solid #e2e8f0" }}>
                              <span style={{
                                fontSize: 10, padding: "2px 8px", borderRadius: 99, fontWeight: 700,
                                background: r.status === "Paid" ? "#dcfce7" : r.status === "Partial" ? "#fef3c7" : "#fee2e2",
                                color: r.status === "Paid" ? "#15803d" : r.status === "Partial" ? "#92400e" : "#b91c1c",
                              }}>{r.status}</span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                      <tfoot>
                        <tr style={{ background: "#f8fafc", fontWeight: 700, borderTop: "2px solid #0f172a" }}>
                          <td style={{ padding: "6px 10px" }} colSpan={3}>Class Total</td>
                          <td style={{ padding: "6px 10px", textAlign: "right", fontFamily: "monospace" }}>{fmt(sec.clsExpected)}</td>
                          <td style={{ padding: "6px 10px", textAlign: "right", fontFamily: "monospace" }}>{fmt(sec.clsPaid)}</td>
                          <td style={{ padding: "6px 10px", textAlign: "right", fontFamily: "monospace", color: sec.clsBalance > 0 ? "#dc2626" : "#0f172a" }}>{fmt(sec.clsBalance)}</td>
                          <td></td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                ))}

                {data.sections.length > 0 && (
                  <div style={{ marginTop: 10, padding: 14, background: "#0f172a", color: "#fff", borderRadius: 8, display: "flex", justifyContent: "space-between", fontSize: 13, fontWeight: 700, flexWrap: "wrap", gap: 6 }}>
                    <span>GRAND TOTAL ({data.totalStudents} students)</span>
                    <span>Expected: {fmt(data.grandExpected)} &nbsp;·&nbsp; Paid: {fmt(data.grandPaid)} &nbsp;·&nbsp; Balance: {fmt(data.grandBalance)}</span>
                  </div>
                )}
                <div style={{ textAlign: "center", fontSize: 11, color: "#94a3b8", marginTop: 18 }}>Generated by FeeTrack UG — {school.name}</div>
              </div>

              <div className="feetrack-no-print" style={{ display: "flex", gap: 10, padding: isMobile ? 16 : 28, paddingTop: 0 }}>
                <button onClick={() => setShowBalancesReport(false)} style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Close</button>
                <button onClick={() => triggerPrint("feetrack-print-balances")} style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#10b981", color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>🖨️ Print</button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* ════════ PAY MODAL ════════ */}
      {showPay && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 430, maxWidth: 430, boxShadow: "0 24px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 3 }}>Record Payment</div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 4 }}>{showPay.name} · {showPay.class} · {showPay.parent}</div>
            <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
              <Pill text={showPay.category || "Day Scholar"} bg="#eff6ff" col="#2563eb" />
              {showPay.customFee && <Pill text="Custom Fee" bg="#f5f3ff" col="#7c3aed" />}
              {showPay.bursary && <Pill text={showPay.bursary.reason} bg="#fef3c7" col="#92400e" />}
            </div>
            <div style={{ background: "#f8fafc", borderRadius: 11, padding: 13, marginBottom: 16 }}>
              {(() => {
                const payBal = getBalance(showPay, currentTerm);
                return [
                  ["Term Fee", fmt(payBal.termFee), "#374151"],
                  ["Arrears", fmt(payBal.arrears), "#f59e0b"],
                  ["Paid This Term", fmt(payBal.paidThisTerm), "#15803d"],
                  ["Balance Due", fmt(payBal.balance), "#dc2626"],
                ];
              })().map(([l, v, c]) => (
                <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: l !== "Balance Due" ? "1px solid #f1f5f9" : "none", fontSize: 13 }}>
                  <span style={{ color: "#64748b" }}>{l}</span>
                  <span style={{ fontWeight: 700, color: c }}>{v}</span>
                </div>
              ))}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
              <div><label style={lbl}>Amount (UGX)</label>
                <input type="number" value={payAmt} onChange={e => setPayAmt(e.target.value)} placeholder="Enter amount" style={inp} /></div>
              <div><label style={lbl}>Payment Date</label>
                <input type="date" value={payDate} onChange={e => setPayDate(e.target.value)} style={inp} /></div>
            </div>
            <div style={{ marginBottom: 12 }}>
              <label style={lbl}>Received By</label>
              <select value={receivedBy} onChange={e => setReceivedBy(e.target.value)} style={inp}>
                {STAFF.map(s => <option key={s}>{s}</option>)}
              </select>
            </div>
            <div style={{ marginBottom: 14 }}>
              <label style={lbl}>Payment Method</label>
              <div style={{ display: "flex", gap: 6 }}>
                {METHODS.map(m => (
                  <button key={m} onClick={() => setPayMethod(m)} style={{ flex: 1, padding: "8px 3px", borderRadius: 8, border: `2px solid ${payMethod === m ? "#f59e0b" : "#e2e8f0"}`, background: payMethod === m ? "#fffbeb" : "#fff", color: payMethod === m ? "#92400e" : "#64748b", fontSize: 10, fontWeight: 700, cursor: "pointer" }}>
                    {METHOD_ICON[m]}<br />{m}
                  </button>
                ))}
              </div>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => { setShowPay(null); setPayAmt(""); }} style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
              <button onClick={handlePay} style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#0f172a", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>✓ Record & Issue Receipt</button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ BULK PAYMENTS IMPORT MODAL ════════ */}
      {showBulkPayments && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 250, padding: 16 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 640, maxWidth: 640, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 24px 60px rgba(0,0,0,0.25)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>💰 Bulk Record Payments</div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 18 }}>Upload a list of students who paid and their amounts — record dozens of payments in one go for {currentTerm}.</div>

            {bulkPayRows.length === 0 ? (
              <div>
                <div style={{ background: "#f8fafc", borderRadius: 12, padding: 16, marginBottom: 16 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 10 }}>How it works:</div>
                  <div style={{ fontSize: 12, color: "#64748b", lineHeight: 1.8 }}>
                    1. Download the template — it's pre-filled with your current student list<br />
                    2. Fill in the Amount column for everyone who paid (leave blank for those who didn't)<br />
                    3. Upload the file — we'll match each row to a student and show a preview<br />
                    4. Review, then click "Record All Payments"
                  </div>
                </div>

                <div style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
                  <button onClick={downloadBulkPayTemplate} style={{ flex: "1 1 45%", background: "#f1f5f9", color: "#374151", border: "1px solid #e2e8f0", borderRadius: 9, padding: "11px 14px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                    📄 Download Template
                  </button>
                  <button onClick={() => bulkPayFileRef.current?.click()} style={{ flex: "1 1 45%", background: "#3b82f6", color: "#fff", border: "none", borderRadius: 9, padding: "11px 14px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                    📂 Choose File to Upload
                  </button>
                  <input ref={bulkPayFileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleBulkPayFileUpload} style={{ display: "none" }} />
                </div>

                <div style={{ background: "#fffbeb", borderRadius: 10, padding: "10px 14px", fontSize: 12, color: "#92400e" }}>
                  <strong>Columns:</strong> Name (must match exactly), Class (helps if names repeat), Amount, Method, Date, Received By. Rows with a blank or zero Amount are skipped automatically — only people who paid need an amount filled in.
                </div>
              </div>
            ) : (
              <div>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
                  <div>
                    <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a" }}>📄 {bulkPayFileName}</div>
                    <div style={{ fontSize: 12, color: "#64748b" }}>
                      {bulkPayRows.filter(r => r.valid).length} payment(s) ready · {fmt(bulkPayRows.filter(r => r.valid).reduce((a, r) => a + r.amount, 0))} total
                    </div>
                  </div>
                  {!bulkPayImportDone && (
                    <button onClick={() => { setBulkPayRows([]); setBulkPayFileName(""); }} style={{ background: "#f1f5f9", color: "#64748b", border: "none", borderRadius: 8, padding: "7px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>✕ Clear</button>
                  )}
                </div>

                {/* Warnings for unmatched/ambiguous/overpay rows */}
                {bulkPayRows.some(r => r.matchStatus !== "matched" || r.overpay) && (
                  <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 10, padding: 12, marginBottom: 12 }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: "#b91c1c", marginBottom: 6 }}>⚠ Some rows need attention:</div>
                    {bulkPayRows.filter(r => r.matchStatus === "not_found").map((r, i) => (
                      <div key={"nf" + i} style={{ fontSize: 11, color: "#7f1d1d", padding: "2px 0" }}>• "{r.name}" — no matching student found, this row will be skipped</div>
                    ))}
                    {bulkPayRows.filter(r => r.matchStatus === "ambiguous").map((r, i) => (
                      <div key={"am" + i} style={{ fontSize: 11, color: "#7f1d1d", padding: "2px 0" }}>• "{r.name}" — multiple students with this name, add a Class to disambiguate; this row will be skipped</div>
                    ))}
                    {bulkPayRows.filter(r => r.matchStatus === "matched" && r.overpay).map((r, i) => (
                      <div key={"ov" + i} style={{ fontSize: 11, color: "#7f1d1d", padding: "2px 0" }}>• "{r.name}" — amount {fmt(r.amount)} exceeds their remaining balance of {fmt(r.balance)}; will still be recorded but check for typos</div>
                    ))}
                  </div>
                )}

                <div style={{ border: "1px solid #e8edf3", borderRadius: 10, overflow: "hidden", marginBottom: 16, maxHeight: 300, overflowY: "auto" }}>
                  <div style={{ overflowX: "auto" }}>
                  <div style={{ minWidth: 560 }}>
                    <div style={{ display: "grid", gridTemplateColumns: "1.6fr 0.6fr 1fr 1fr 1fr 0.8fr", gap: 4, padding: "8px 12px", background: "#f8fafc", borderBottom: "1px solid #e8edf3", position: "sticky", top: 0 }}>
                      {["Name", "Class", "Amount", "Method", "Date", "Status"].map(h => (
                        <div key={h} style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase" }}>{h}</div>
                      ))}
                    </div>
                    {bulkPayRows.map(r => (
                      <div key={r.rowIndex} style={{ display: "grid", gridTemplateColumns: "1.6fr 0.6fr 1fr 1fr 1fr 0.8fr", gap: 4, alignItems: "center", padding: "7px 12px", borderTop: "1px solid #f1f5f9", background: r.valid ? "#fff" : "#fef2f2", fontSize: 12 }}>
                        <div style={{ fontWeight: 600, color: "#0f172a" }}>{r.name}</div>
                        <div>{r.matchedStudent ? <Pill text={r.matchedStudent.class} /> : (r.class ? <Pill text={r.class} /> : "—")}</div>
                        <div style={{ fontWeight: 800, color: r.amount > 0 ? "#15803d" : "#94a3b8" }}>{r.amount > 0 ? fmt(r.amount) : "—"}</div>
                        <div style={{ fontSize: 11 }}>{r.amount > 0 ? `${METHOD_ICON[r.method] || "💰"} ${r.method}` : "—"}</div>
                        <div style={{ fontSize: 11 }}>{r.amount > 0 ? fmtDate(r.date) : "—"}</div>
                        <div>
                          {r.matchStatus === "matched" && r.amount > 0 && <Pill text="✓ Ready" bg="#f0fdf4" col="#15803d" />}
                          {r.matchStatus === "matched" && r.amount === 0 && <Pill text="Skip" bg="#f1f5f9" col="#94a3b8" />}
                          {r.matchStatus === "not_found" && <Pill text="✗ Not found" bg="#fef2f2" col="#dc2626" />}
                          {r.matchStatus === "ambiguous" && <Pill text="⚠ Ambiguous" bg="#fffbeb" col="#92400e" />}
                        </div>
                      </div>
                    ))}
                  </div>
                  </div>
                </div>

                {bulkPayImportDone ? (
                  <div style={{ background: "#f0fdf4", border: "1px solid #86efac", borderRadius: 10, padding: 14, textAlign: "center", color: "#15803d", fontWeight: 700, fontSize: 13 }}>
                    ✅ {bulkPayRows.filter(r => r.valid).length} payments recorded successfully!
                  </div>
                ) : (
                  <div style={{ display: "flex", gap: 10 }}>
                    <button onClick={() => { setShowBulkPayments(false); setBulkPayRows([]); setBulkPayFileName(""); }}
                      style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
                    <button onClick={handleBulkPayImport} disabled={bulkPayRows.filter(r => r.valid).length === 0}
                      style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: bulkPayRows.filter(r => r.valid).length > 0 ? "#10b981" : "#e2e8f0", color: bulkPayRows.filter(r => r.valid).length > 0 ? "#fff" : "#94a3b8", fontSize: 13, fontWeight: 800, cursor: bulkPayRows.filter(r => r.valid).length > 0 ? "pointer" : "not-allowed" }}>
                      ✓ Record All Payments ({bulkPayRows.filter(r => r.valid).length})
                    </button>
                  </div>
                )}
              </div>
            )}

            {bulkPayRows.length === 0 && (
              <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
                <button onClick={() => setShowBulkPayments(false)}
                  style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Close</button>
              </div>
            )}
            {bulkPayImportDone && (
              <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
                <button onClick={() => { setShowBulkPayments(false); setBulkPayRows([]); setBulkPayFileName(""); setBulkPayImportDone(false); }}
                  style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#0f172a", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Done</button>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ════════ BULK PAY STAFF MODAL ════════ */}
      {showBulkStaffPay && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 250, padding: 16 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 640, maxWidth: 640, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 24px 60px rgba(0,0,0,0.25)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>💰 Bulk Pay Staff</div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 18 }}>Upload a list of workers and what they're owed — record many staff payments in one go for {currentTerm}.</div>

            {bulkStaffPayRows.length === 0 ? (
              <div>
                <div style={{ background: "#f8fafc", borderRadius: 12, padding: 16, marginBottom: 16 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 10 }}>How it works:</div>
                  <div style={{ fontSize: 12, color: "#64748b", lineHeight: 1.8 }}>
                    1. Download the template — it's pre-filled with your active staff list<br />
                    2. Fill in Amount, Type (Daily/Monthly), and Period for everyone being paid (leave blank to skip someone)<br />
                    3. Upload the file — we'll match each row to a worker and show a preview<br />
                    4. Review, then click "Record All Payments"
                  </div>
                </div>

                <div style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
                  <button onClick={downloadBulkStaffPayTemplate} style={{ flex: "1 1 45%", background: "#f1f5f9", color: "#374151", border: "1px solid #e2e8f0", borderRadius: 9, padding: "11px 14px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                    📄 Download Template
                  </button>
                  <button onClick={() => bulkStaffPayFileRef.current?.click()} style={{ flex: "1 1 45%", background: "#3b82f6", color: "#fff", border: "none", borderRadius: 9, padding: "11px 14px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                    📂 Choose File to Upload
                  </button>
                  <input ref={bulkStaffPayFileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleBulkStaffPayFileUpload} style={{ display: "none" }} />
                </div>

                <div style={{ background: "#fffbeb", borderRadius: 10, padding: "10px 14px", fontSize: 12, color: "#92400e" }}>
                  <strong>Columns:</strong> Name (must match exactly), Role (helps if names repeat), Amount, Type (Daily or Monthly), Period (e.g. "16 Jun 2026" or "June 2026"), Date. Rows with a blank Amount or Period are skipped automatically.
                </div>
              </div>
            ) : (
              <div>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
                  <div>
                    <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a" }}>📄 {bulkStaffPayFileName}</div>
                    <div style={{ fontSize: 12, color: "#64748b" }}>
                      {bulkStaffPayRows.filter(r => r.valid).length} payment(s) ready · {fmt(bulkStaffPayRows.filter(r => r.valid).reduce((a, r) => a + r.amount, 0))} total
                    </div>
                  </div>
                  {!bulkStaffPayImportDone && (
                    <button onClick={() => { setBulkStaffPayRows([]); setBulkStaffPayFileName(""); }} style={{ background: "#f1f5f9", color: "#64748b", border: "none", borderRadius: 8, padding: "7px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>✕ Clear</button>
                  )}
                </div>

                {/* Warnings for unmatched/ambiguous/incomplete rows */}
                {bulkStaffPayRows.some(r => !r.valid) && (
                  <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 10, padding: 12, marginBottom: 12 }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: "#b91c1c", marginBottom: 6 }}>⚠ Some rows need attention:</div>
                    {bulkStaffPayRows.filter(r => r.matchStatus === "not_found").map((r, i) => (
                      <div key={"nf" + i} style={{ fontSize: 11, color: "#7f1d1d", padding: "2px 0" }}>• "{r.name}" — no matching worker found, this row will be skipped</div>
                    ))}
                    {bulkStaffPayRows.filter(r => r.matchStatus === "ambiguous").map((r, i) => (
                      <div key={"am" + i} style={{ fontSize: 11, color: "#7f1d1d", padding: "2px 0" }}>• "{r.name}" — multiple workers with this name, add a Role to disambiguate; this row will be skipped</div>
                    ))}
                    {bulkStaffPayRows.filter(r => r.matchStatus === "matched" && r.amount > 0 && !r.periodLabel).map((r, i) => (
                      <div key={"np" + i} style={{ fontSize: 11, color: "#7f1d1d", padding: "2px 0" }}>• "{r.name}" — missing a Period (e.g. "16 Jun 2026"), this row will be skipped</div>
                    ))}
                  </div>
                )}

                <div style={{ border: "1px solid #e8edf3", borderRadius: 10, overflow: "hidden", marginBottom: 16, maxHeight: 300, overflowY: "auto" }}>
                  <div style={{ overflowX: "auto" }}>
                  <div style={{ minWidth: 600 }}>
                    <div style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr 0.9fr 0.8fr 1fr 0.8fr", gap: 4, padding: "8px 12px", background: "#f8fafc", borderBottom: "1px solid #e8edf3", position: "sticky", top: 0 }}>
                      {["Name", "Period", "Amount", "Type", "Date", "Status"].map(h => (
                        <div key={h} style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase" }}>{h}</div>
                      ))}
                    </div>
                    {bulkStaffPayRows.map(r => (
                      <div key={r.rowIndex} style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr 0.9fr 0.8fr 1fr 0.8fr", gap: 4, alignItems: "center", padding: "7px 12px", borderTop: "1px solid #f1f5f9", background: r.valid ? "#fff" : "#fef2f2", fontSize: 12 }}>
                        <div style={{ fontWeight: 600, color: "#0f172a" }}>{r.name}</div>
                        <div style={{ fontSize: 11 }}>{r.periodLabel || "—"}</div>
                        <div style={{ fontWeight: 800, color: r.amount > 0 ? "#15803d" : "#94a3b8" }}>{r.amount > 0 ? fmt(r.amount) : "—"}</div>
                        <div><Pill text={r.payType === "monthly" ? "Monthly" : "Daily"} bg="#f1f5f9" col="#475569" /></div>
                        <div style={{ fontSize: 11 }}>{fmtDate(r.date)}</div>
                        <div>
                          {r.valid && <Pill text="✓ Ready" bg="#f0fdf4" col="#15803d" />}
                          {!r.valid && r.matchStatus === "not_found" && <Pill text="✗ Not found" bg="#fef2f2" col="#dc2626" />}
                          {!r.valid && r.matchStatus === "ambiguous" && <Pill text="⚠ Ambiguous" bg="#fffbeb" col="#92400e" />}
                          {!r.valid && r.matchStatus === "matched" && <Pill text="Skip" bg="#f1f5f9" col="#94a3b8" />}
                        </div>
                      </div>
                    ))}
                  </div>
                  </div>
                </div>

                {bulkStaffPayImportDone ? (
                  <div style={{ background: "#f0fdf4", border: "1px solid #86efac", borderRadius: 10, padding: 14, textAlign: "center", color: "#15803d", fontWeight: 700, fontSize: 13 }}>
                    ✅ {bulkStaffPayRows.filter(r => r.valid).length} staff payments recorded successfully!
                  </div>
                ) : (
                  <div style={{ display: "flex", gap: 10 }}>
                    <button onClick={() => { setShowBulkStaffPay(false); setBulkStaffPayRows([]); setBulkStaffPayFileName(""); }}
                      style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
                    <button onClick={handleBulkStaffPayImport} disabled={bulkStaffPayRows.filter(r => r.valid).length === 0}
                      style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: bulkStaffPayRows.filter(r => r.valid).length > 0 ? "#10b981" : "#e2e8f0", color: bulkStaffPayRows.filter(r => r.valid).length > 0 ? "#fff" : "#94a3b8", fontSize: 13, fontWeight: 800, cursor: bulkStaffPayRows.filter(r => r.valid).length > 0 ? "pointer" : "not-allowed" }}>
                      ✓ Record All Payments ({bulkStaffPayRows.filter(r => r.valid).length})
                    </button>
                  </div>
                )}
              </div>
            )}

            {bulkStaffPayRows.length === 0 && (
              <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
                <button onClick={() => setShowBulkStaffPay(false)}
                  style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Close</button>
              </div>
            )}
            {bulkStaffPayImportDone && (
              <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
                <button onClick={() => { setShowBulkStaffPay(false); setBulkStaffPayRows([]); setBulkStaffPayFileName(""); setBulkStaffPayImportDone(false); }}
                  style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#0f172a", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Done</button>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ════════ BULK IMPORT MODAL ════════ */}
      {showBulkImport && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 250, padding: 16 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 560, maxWidth: 560, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 24px 60px rgba(0,0,0,0.25)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>📋 Bulk Enrol Students</div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 18 }}>Upload an Excel or CSV file with your student list — faster than enrolling one by one</div>

            {bulkRows.length === 0 ? (
              <div>
                {/* Steps */}
                <div style={{ background: "#f8fafc", borderRadius: 12, padding: 16, marginBottom: 16 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 10 }}>How it works:</div>
                  <div style={{ fontSize: 12, color: "#64748b", lineHeight: 1.8 }}>
                    1. Download our template below (or use your own spreadsheet)<br />
                    2. Fill in: Name, Class, Gender, Category, Parent, Phone, Arrears<br />
                    3. Upload the file — we'll show a preview before importing<br />
                    4. Click "Import All" to add everyone at once
                  </div>
                </div>

                <div style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
                  <button onClick={downloadBulkTemplate} style={{ flex: "1 1 45%", background: "#f1f5f9", color: "#374151", border: "1px solid #e2e8f0", borderRadius: 9, padding: "11px 14px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                    📄 Download Template
                  </button>
                  <button onClick={() => bulkFileRef.current?.click()} style={{ flex: "1 1 45%", background: "#3b82f6", color: "#fff", border: "none", borderRadius: 9, padding: "11px 14px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                    📂 Choose File to Upload
                  </button>
                  <input ref={bulkFileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleBulkFileUpload} style={{ display: "none" }} />
                </div>

                <div style={{ background: "#fffbeb", borderRadius: 10, padding: "10px 14px", fontSize: 12, color: "#92400e" }}>
                  <strong>Accepted column names:</strong> Name, Class (S1–S6), Gender (M/F), Category (Day Scholar/Boarder), Parent, Phone, Arrears. Column names are flexible — "Student Name", "Form", "Guardian" etc. also work.
                  <div style={{ marginTop: 6 }}>
                    <strong>Fee arrangements:</strong> add a "Fee Type" column (Full Fee / Bursary / Custom Fee). For Bursary students, add "Bursary %" or "Bursary Amount" plus "Bursary Reason". For Custom Fee students, add "Custom Fee" with the exact UGX amount. Leave blank for students paying the standard fee.
                  </div>
                </div>
              </div>
            ) : (
              <div>
                {/* Preview */}
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
                  <div>
                    <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a" }}>📄 {bulkFileName}</div>
                    <div style={{ fontSize: 12, color: "#64748b" }}>{bulkRows.filter(r => r.valid).length} valid students ready to import</div>
                  </div>
                  {!bulkImportDone && (
                    <button onClick={() => { setBulkRows([]); setBulkFileName(""); }} style={{ background: "#f1f5f9", color: "#64748b", border: "none", borderRadius: 8, padding: "7px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>✕ Clear</button>
                  )}
                </div>

                <div style={{ border: "1px solid #e8edf3", borderRadius: 10, overflow: "hidden", marginBottom: 16, maxHeight: 300, overflowY: "auto" }}>
                  <div style={{ overflowX: "auto" }}>
                  <div style={{ minWidth: 620 }}>
                    <div style={{ display: "grid", gridTemplateColumns: "1.4fr 0.5fr 0.5fr 0.9fr 1.1fr 1.1fr 0.7fr 1.2fr", gap: 4, padding: "8px 12px", background: "#f8fafc", borderBottom: "1px solid #e8edf3", position: "sticky", top: 0 }}>
                      {["Name", "Class", "Gender", "Category", "Parent", "Phone", "Arrears", "Fee Arrangement"].map(h => (
                        <div key={h} style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase" }}>{h}</div>
                      ))}
                    </div>
                    {bulkRows.map((r, i) => (
                      <div key={r.rowIndex} style={{ display: "grid", gridTemplateColumns: "1.4fr 0.5fr 0.5fr 0.9fr 1.1fr 1.1fr 0.7fr 1.2fr", gap: 4, alignItems: "center", padding: "7px 12px", borderTop: "1px solid #f1f5f9", background: i % 2 === 0 ? "#fff" : "#fafbfc", fontSize: 12 }}>
                        <div style={{ fontWeight: 600, color: "#0f172a" }}>{r.name}</div>
                        <div><Pill text={r.class} /></div>
                        <div>{r.gender === "F" ? "👩" : "👨"}</div>
                        <div style={{ fontSize: 11, color: "#64748b" }}>{r.category}</div>
                        <div style={{ fontSize: 11 }}>{r.parent}</div>
                        <div style={{ fontSize: 11, fontFamily: "monospace" }}>{r.phone || "—"}</div>
                        <div style={{ fontSize: 11, color: r.arrears > 0 ? "#dc2626" : "#94a3b8" }}>{r.arrears > 0 ? fmt(r.arrears) : "—"}</div>
                        <div>
                          {r.customFee
                            ? <Pill text={`Custom: ${fmt(r.customFee)}`} bg="#f5f3ff" col="#7c3aed" />
                            : r.bursary
                              ? <Pill text={`${r.bursary.type === "percent" ? r.bursary.value + "% off" : fmt(r.bursary.value) + " off"} · ${r.bursary.reason}`} bg="#fef3c7" col="#92400e" />
                              : <Pill text="Full Fee" bg="#f0fdf4" col="#15803d" />}
                        </div>
                      </div>
                    ))}
                  </div>
                  </div>
                </div>

                {bulkImportDone ? (
                  <div style={{ background: "#f0fdf4", border: "1px solid #86efac", borderRadius: 10, padding: 14, textAlign: "center", color: "#15803d", fontWeight: 700, fontSize: 13 }}>
                    ✅ {bulkRows.filter(r => r.valid).length} students imported successfully!
                  </div>
                ) : (
                  <div style={{ display: "flex", gap: 10 }}>
                    <button onClick={() => { setShowBulkImport(false); setBulkRows([]); setBulkFileName(""); }}
                      style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
                    <button onClick={handleBulkImport}
                      style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#10b981", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>✓ Import All ({bulkRows.filter(r => r.valid).length})</button>
                  </div>
                )}
              </div>
            )}

            {bulkRows.length === 0 && (
              <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
                <button onClick={() => setShowBulkImport(false)}
                  style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Close</button>
              </div>
            )}
            {bulkImportDone && (
              <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
                <button onClick={() => { setShowBulkImport(false); setBulkRows([]); setBulkFileName(""); setBulkImportDone(false); }}
                  style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#0f172a", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Done</button>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ════════ ADD STUDENT MODAL ════════ */}
      {showAdd && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 460, maxWidth: 460, boxShadow: "0 24px 60px rgba(0,0,0,0.2)", maxHeight: "90vh", overflowY: "auto" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 20 }}>Enrol New Student</div>

            <div style={{ marginBottom: 13 }}>
              <label style={lbl}>Full Name *</label>
              <input value={newS.name || ""} onChange={e => { setNewS({ ...newS, name: e.target.value }); checkReturningStudent(e.target.value); }} placeholder="e.g. Nakato Sarah" style={inp} />
            </div>

            {/* Returning student detected banner */}
            {returningMatch && (
              <div style={{ background: "#f5f3ff", border: "2px solid #c4b5fd", borderRadius: 12, padding: 14, marginBottom: 14 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                  {returningMatch.photo
                    ? <img src={returningMatch.photo} alt="" style={{ width: 36, height: 36, borderRadius: 8, objectFit: "cover" }} />
                    : <span style={{ fontSize: 24 }}>{returningMatch.gender === "F" ? "👩" : "👨"}</span>}
                  <div>
                    <div style={{ fontWeight: 800, fontSize: 13, color: "#6d28d9" }}>🎓 Returning Student Found in Alumni!</div>
                    <div style={{ fontSize: 11, color: "#7c3aed" }}>
                      Left as {returningMatch.leftClass} in {returningMatch.leftYear} · {returningMatch.leftNote}
                    </div>
                  </div>
                </div>
                <div style={{ display: "flex", gap: 12, fontSize: 12, color: "#5b21b6", marginBottom: 10 }}>
                  <span>Past payments: <strong>{(returningMatch.payments || []).length}</strong></span>
                  <span>·</span>
                  <span>Outstanding balance: <strong style={{ color: returningMatch.outstandingDebt > 0 ? "#dc2626" : "#15803d" }}>{fmt(returningMatch.outstandingDebt || 0)}</strong></span>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", background: "#fff", borderRadius: 9, cursor: "pointer" }} onClick={() => setConfirmReturning(!confirmReturning)}>
                  <div style={{ width: 18, height: 18, borderRadius: 4, border: `2px solid ${confirmReturning ? "#7c3aed" : "#e2e8f0"}`, background: confirmReturning ? "#7c3aed" : "#fff", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                    {confirmReturning && <span style={{ color: "#fff", fontSize: 11, fontWeight: 800 }}>✓</span>}
                  </div>
                  <span style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>
                    Yes, this is the same student — restore their records and remove from Alumni
                    {returningMatch.outstandingDebt > 0 && <span style={{ color: "#dc2626" }}> (old balance of {fmt(returningMatch.outstandingDebt)} will carry forward as arrears)</span>}
                  </span>
                </div>
              </div>
            )}

            {[{ label: "Parent / Guardian *", key: "parent", placeholder: "e.g. Nakato Mary" }, { label: "Parent Phone", key: "phone", placeholder: "e.g. 0772-441-823" }].map(f => (
              <div key={f.key} style={{ marginBottom: 13 }}>
                <label style={lbl}>{f.label}</label>
                <input value={newS[f.key] || ""} onChange={e => setNewS({ ...newS, [f.key]: e.target.value })} placeholder={f.placeholder} style={inp} />
              </div>
            ))}
            <div style={{ display: "grid", gridTemplateColumns: getClassStreams(newS.class).length > 0 ? "0.8fr 0.8fr 1fr 1fr" : "1fr 1fr 1fr", gap: 12, marginBottom: 13 }}>
              <div><label style={lbl}>Class</label>
                <select value={newS.class} onChange={e => setNewS({ ...newS, class: e.target.value, stream: "" })} style={inp}>
                  {schoolClasses.map(c => <option key={c} value={c}>{c}</option>)}
                </select></div>
              {getClassStreams(newS.class).length > 0 && (
                <div><label style={lbl}>Stream</label>
                  <select value={newS.stream || ""} onChange={e => setNewS({ ...newS, stream: e.target.value })} style={inp}>
                    <option value="">— Select —</option>
                    {getClassStreams(newS.class).map(s => <option key={s} value={s}>{s}</option>)}
                  </select></div>
              )}
              <div><label style={lbl}>Category</label>
                <select value={newS.category} onChange={e => setNewS({ ...newS, category: e.target.value })} style={inp}>
                  {STUDENT_CATEGORIES.map(c => <option key={c}>{c}</option>)}
                </select></div>
              <div><label style={lbl}>Gender</label>
                <select value={newS.gender} onChange={e => setNewS({ ...newS, gender: e.target.value })} style={inp}>
                  <option value="M">Male</option><option value="F">Female</option>
                </select></div>
            </div>
            <div style={{ marginBottom: 13 }}>
              <label style={lbl}>Custom Fee Override (optional)</label>
              <input type="number" value={newS.customFee} onChange={e => setNewS({ ...newS, customFee: e.target.value })} placeholder="Leave blank to use category fee" style={inp} />
              <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 4 }}>
                {newS.category} {newS.class} base fee: <strong>{fmt(Object.values((FEE_STRUCTURE[newS.category] || FEE_STRUCTURE["Day Scholar"])[newS.class] || {}).reduce((a, b) => a + b, 0))}</strong>
                {newS.customFee && <span style={{ color: "#7c3aed", marginLeft: 6 }}>→ Custom: <strong>{fmt(parseInt(newS.customFee))}</strong></span>}
              </div>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => { setShowAdd(false); setReturningMatch(null); setConfirmReturning(false); }} style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
              <button onClick={handleAddStudent} style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: confirmReturning ? "#7c3aed" : "#f59e0b", color: confirmReturning ? "#fff" : "#0f172a", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>
                {confirmReturning ? "🎓 Re-enrol Returning Student" : "Enrol Student"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ FEE EDIT MODAL ════════ */}
      {showFeeEdit && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 460, maxWidth: 460, boxShadow: "0 24px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>⚙ Edit Fee — {showFeeEdit.name}</div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 20 }}>{showFeeEdit.class} · {showFeeEdit.category || "Day Scholar"} · Current fee: <strong>{fmt(getStudentFee(showFeeEdit))}</strong></div>

            {/* Mode tabs */}
            <div style={{ display: "flex", gap: 6, marginBottom: 20, background: "#f1f5f9", borderRadius: 10, padding: 4 }}>
              {[["category", "📋 Category Rate"], ["bursary", "🎓 Bursary / Discount"], ["custom", "✏️ Custom Fee"]].map(([mode, label]) => (
                <button key={mode} onClick={() => setFeeEditData(prev => ({ ...prev, mode }))} style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 12, fontWeight: 700, background: feeEditData.mode === mode ? "#fff" : "transparent", color: feeEditData.mode === mode ? "#0f172a" : "#64748b", boxShadow: feeEditData.mode === mode ? "0 1px 4px rgba(0,0,0,0.1)" : "none" }}>
                  {label}
                </button>
              ))}
            </div>

            {feeEditData.mode === "category" && (
              <div style={{ background: "#f0fdf4", border: "1px solid #86efac", borderRadius: 12, padding: 16, marginBottom: 20 }}>
                <div style={{ fontWeight: 700, color: "#15803d", fontSize: 13, marginBottom: 8 }}>✓ Use Standard Category Rate</div>
                <div style={{ fontSize: 12, color: "#374151" }}>Removes any bursary or custom fee. Student pays the standard {showFeeEdit.category || "Day Scholar"} rate for {showFeeEdit.class}.</div>
                <div style={{ fontSize: 14, fontWeight: 800, color: "#15803d", marginTop: 10 }}>
                  {fmt(Object.values((FEE_STRUCTURE[showFeeEdit.category || "Day Scholar"])[showFeeEdit.class] || {}).reduce((a, b) => a + b, 0))} / term
                </div>
              </div>
            )}

            {feeEditData.mode === "bursary" && (
              <div style={{ marginBottom: 20 }}>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
                  <div><label style={lbl}>Discount Type</label>
                    <select value={feeEditData.bursaryType} onChange={e => setFeeEditData(p => ({ ...p, bursaryType: e.target.value }))} style={inp}>
                      <option value="percent">Percentage (%)</option>
                      <option value="fixed">Fixed Amount (UGX)</option>
                    </select></div>
                  <div><label style={lbl}>{feeEditData.bursaryType === "percent" ? "Discount %" : "Discount Amount"}</label>
                    <input type="number" value={feeEditData.bursaryValue} onChange={e => setFeeEditData(p => ({ ...p, bursaryValue: e.target.value }))} placeholder={feeEditData.bursaryType === "percent" ? "e.g. 50" : "e.g. 200000"} style={inp} /></div>
                </div>
                <div style={{ marginBottom: 12 }}><label style={lbl}>Reason / Scholarship Name</label>
                  <input value={feeEditData.bursaryReason} onChange={e => setFeeEditData(p => ({ ...p, bursaryReason: e.target.value }))} placeholder="e.g. Orphan Bursary, Staff Child, Academic Scholarship" style={inp} /></div>
                {feeEditData.bursaryValue && (
                  <div style={{ background: "#fef3c7", borderRadius: 10, padding: 12, fontSize: 12, color: "#92400e" }}>
                    Base fee: {fmt(Object.values((FEE_STRUCTURE[showFeeEdit.category || "Day Scholar"])[showFeeEdit.class] || {}).reduce((a, b) => a + b, 0))}
                    {" → "}After discount: <strong>{fmt(
                      feeEditData.bursaryType === "percent"
                        ? Math.round(Object.values((FEE_STRUCTURE[showFeeEdit.category || "Day Scholar"])[showFeeEdit.class] || {}).reduce((a, b) => a + b, 0) * (1 - feeEditData.bursaryValue / 100))
                        : Math.max(0, Object.values((FEE_STRUCTURE[showFeeEdit.category || "Day Scholar"])[showFeeEdit.class] || {}).reduce((a, b) => a + b, 0) - parseInt(feeEditData.bursaryValue || 0))
                    )}</strong>
                  </div>
                )}
              </div>
            )}

            {feeEditData.mode === "custom" && (
              <div style={{ marginBottom: 20 }}>
                <label style={lbl}>Custom Fee Amount (UGX / term)</label>
                <input type="number" value={feeEditData.customFee} onChange={e => setFeeEditData(p => ({ ...p, customFee: e.target.value }))} placeholder="Enter exact amount this student pays" style={{ ...inp, fontSize: 18, fontWeight: 800 }} />
                <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 6 }}>This completely overrides category and bursary. Use for special agreements with parents.</div>
              </div>
            )}

            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => setShowFeeEdit(null)} style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
              <button onClick={handleFeeEdit} style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#8b5cf6", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Save Fee Settings</button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ EDIT EXPENSE MODAL ════════ */}
      {showEditExpense && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 250 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 380, maxWidth: 380, boxShadow: "0 24px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>✏️ Correct Expense Amount</div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 18 }}>Fixing a wrong amount entry</div>
            <div style={{ background: "#f8fafc", borderRadius: 11, padding: 13, marginBottom: 16 }}>
              {[
                ["Category", showEditExpense.category],
                ["Description", showEditExpense.description],
                ["Date", fmtDate(showEditExpense.date)],
                ["Current Amount", fmt(showEditExpense.amount)],
              ].map(([l, v]) => (
                <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: "1px solid #f1f5f9", fontSize: 13 }}>
                  <span style={{ color: "#64748b" }}>{l}</span>
                  <span style={{ fontWeight: 600, color: "#0f172a" }}>{v}</span>
                </div>
              ))}
            </div>
            <div style={{ marginBottom: 20 }}>
              <label style={lbl}>Corrected Amount (UGX)</label>
              <input type="number" value={editExpAmt} onChange={e => setEditExpAmt(e.target.value)}
                placeholder={`Was: ${showEditExpense.amount}`}
                style={{ width: "100%", padding: "11px 14px", borderRadius: 9, border: "2px solid #ef4444", fontSize: 16, fontWeight: 800, outline: "none", boxSizing: "border-box" }} />
              <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 5 }}>⚠ Only use this to fix a data entry error.</div>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => { setShowEditExpense(null); setEditExpAmt(""); }}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
              <button onClick={handleEditExpense}
                style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#ef4444", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Save Correction</button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ ADD EXPENSE MODAL ════════ */}
      {showAddExp && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 400, maxWidth: 400, boxShadow: "0 24px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 20 }}>Record Expense</div>
            <div style={{ marginBottom: 13 }}><label style={lbl}>Category</label>
              <select value={newExp.category} onChange={e => setNewExp({ ...newExp, category: e.target.value })} style={inp}>
                {expenseCategories.map(c => <option key={c}>{c}</option>)}
              </select></div>
            <div style={{ marginBottom: 13 }}><label style={lbl}>Description</label>
              <input value={newExp.description} onChange={e => setNewExp({ ...newExp, description: e.target.value })} placeholder="e.g. Staff salaries for May" style={inp} /></div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 20 }}>
              <div><label style={lbl}>Amount (UGX)</label>
                <input type="number" value={newExp.amount} onChange={e => setNewExp({ ...newExp, amount: e.target.value })} placeholder="e.g. 500000" style={inp} /></div>
              <div><label style={lbl}>Date</label>
                <input type="date" value={newExp.date} onChange={e => setNewExp({ ...newExp, date: e.target.value })} style={inp} /></div>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => setShowAddExp(false)} style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
              <button onClick={handleAddExpense} style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#ef4444", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Save Expense</button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ ADD WORKER MODAL ════════ */}
      {showAddStaff && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 16 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 400, maxWidth: 400, boxShadow: "0 24px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 20 }}>👷 Add Worker</div>
            <div style={{ marginBottom: 13 }}><label style={lbl}>Full Name</label>
              <input value={newStaff.name} onChange={e => setNewStaff({ ...newStaff, name: e.target.value })} placeholder="e.g. Nakato Betty" style={inp} /></div>
            <div style={{ marginBottom: 13 }}><label style={lbl}>Role</label>
              <input value={newStaff.role} onChange={e => setNewStaff({ ...newStaff, role: e.target.value })} placeholder="e.g. Cook, Security Guard, Driver" style={inp} /></div>
            <div style={{ marginBottom: 13 }}><label style={lbl}>Phone (optional)</label>
              <input value={newStaff.phone} onChange={e => setNewStaff({ ...newStaff, phone: e.target.value })} placeholder="e.g. 0772-100-001" style={inp} /></div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 20 }}>
              <div><label style={lbl}>Usual Rate (UGX)</label>
                <input type="number" value={newStaff.defaultRate} onChange={e => setNewStaff({ ...newStaff, defaultRate: e.target.value })} placeholder="e.g. 15000" style={inp} /></div>
              <div><label style={lbl}>Rate Type</label>
                <select value={newStaff.defaultRateType} onChange={e => setNewStaff({ ...newStaff, defaultRateType: e.target.value })} style={inp}>
                  <option value="daily">Daily</option>
                  <option value="monthly">Monthly</option>
                </select></div>
            </div>
            <div style={{ fontSize: 11, color: "#94a3b8", marginBottom: 16 }}>This is just a reference to speed up recording payments later — you can always pay a different amount or frequency.</div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => setShowAddStaff(false)} style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
              <button onClick={handleAddStaff} style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#0f172a", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Add Worker</button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ EDIT WORKER MODAL ════════ */}
      {showEditStaff && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 16 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 400, maxWidth: 400, boxShadow: "0 24px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 20 }}>✏ Edit {showEditStaff.name}</div>
            <div style={{ marginBottom: 13 }}><label style={lbl}>Full Name</label>
              <input value={showEditStaff.name} onChange={e => setShowEditStaff({ ...showEditStaff, name: e.target.value })} style={inp} /></div>
            <div style={{ marginBottom: 13 }}><label style={lbl}>Role</label>
              <input value={showEditStaff.role} onChange={e => setShowEditStaff({ ...showEditStaff, role: e.target.value })} style={inp} /></div>
            <div style={{ marginBottom: 13 }}><label style={lbl}>Phone</label>
              <input value={showEditStaff.phone} onChange={e => setShowEditStaff({ ...showEditStaff, phone: e.target.value })} style={inp} /></div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 20 }}>
              <div><label style={lbl}>Usual Rate (UGX)</label>
                <input type="number" value={showEditStaff.defaultRate} onChange={e => setShowEditStaff({ ...showEditStaff, defaultRate: e.target.value })} style={inp} /></div>
              <div><label style={lbl}>Rate Type</label>
                <select value={showEditStaff.defaultRateType} onChange={e => setShowEditStaff({ ...showEditStaff, defaultRateType: e.target.value })} style={inp}>
                  <option value="daily">Daily</option>
                  <option value="monthly">Monthly</option>
                </select></div>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => setShowEditStaff(null)} style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
              <button onClick={handleEditStaff} style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#0f172a", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Save Changes</button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ PAY WORKER MODAL ════════ */}
      {showPayStaff && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 16 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 30, width: isMobile ? "calc(100vw - 32px)" : 400, maxWidth: 400, boxShadow: "0 24px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>💵 Pay {showPayStaff.name}</div>
            <div style={{ fontSize: 12, color: "#64748b", marginBottom: 20 }}>{showPayStaff.role} · Usual rate: {fmt(showPayStaff.defaultRate)} / {showPayStaff.defaultRateType === "daily" ? "day" : "month"}</div>
            <div style={{ marginBottom: 13 }}>
              <label style={lbl}>Pay Type</label>
              <div style={{ display: "flex", gap: 8 }}>
                {[["daily", "Daily"], ["monthly", "Monthly"]].map(([val, label]) => (
                  <button key={val} onClick={() => setPayStaffForm(p => ({ ...p, payType: val }))}
                    style={{ flex: 1, padding: "9px 0", borderRadius: 8, border: `2px solid ${payStaffForm.payType === val ? "#10b981" : "#e2e8f0"}`, background: payStaffForm.payType === val ? "#f0fdf4" : "#fff", color: payStaffForm.payType === val ? "#15803d" : "#64748b", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                    {label}
                  </button>
                ))}
              </div>
            </div>
            <div style={{ marginBottom: 13 }}><label style={lbl}>Period This Payment Covers</label>
              <input value={payStaffForm.periodLabel} onChange={e => setPayStaffForm({ ...payStaffForm, periodLabel: e.target.value })}
                placeholder={payStaffForm.payType === "daily" ? "e.g. 16 Jun 2026" : "e.g. June 2026"} style={inp} /></div>
            <div style={{ marginBottom: 20 }}><label style={lbl}>Amount (UGX)</label>
              <input type="number" value={payStaffForm.amount} onChange={e => setPayStaffForm({ ...payStaffForm, amount: e.target.value })} placeholder="e.g. 15000" style={inp} /></div>
            <div style={{ background: "#f8fafc", borderRadius: 10, padding: 11, marginBottom: 16, fontSize: 11, color: "#64748b" }}>
              This will also be added to <strong>Expenses</strong> under "Salaries & Wages" automatically.
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => setShowPayStaff(null)} style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
              <button onClick={handlePayStaff} style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#10b981", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Record Payment</button>
            </div>
          </div>
        </div>
      )}

      {/* ════════ TERM ROLLOVER MODAL ════════ */}
      {showRollover && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 16 }}>
          <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 32, width: isMobile ? "calc(100vw - 32px)" : 480, maxWidth: 480, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 24px 60px rgba(0,0,0,0.2)" }}>

            {rolloverStep === 1 && (
              <>
                <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 6 }}>🔄 Term Rollover</div>
                <div style={{ fontSize: 13, color: "#64748b", marginBottom: 20 }}>End the current term and start a new one. Unpaid balances will be carried forward as arrears.</div>
                <div style={{ background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 10, padding: 14, marginBottom: 20 }}>
                  <div style={{ fontWeight: 700, color: "#92400e", fontSize: 13, marginBottom: 6 }}>⚠ What will happen:</div>
                  <div style={{ fontSize: 12, color: "#78350f", lineHeight: 1.7 }}>
                    • {termStudents.filter(s => getStatus(s, currentTerm) !== "Paid").length} students with outstanding balances will have their balance added as arrears<br />
                    • New term fee collection starts fresh<br />
                    • All previous payment records are preserved
                  </div>
                </div>
                <div style={{ marginBottom: 20 }}>
                  <label style={lbl}>Roll over to which term?</label>
                  <select value={rolloverTerm || TERMS[TERMS.indexOf(currentTerm) + 1] || TERMS[0]} onChange={e => setRolloverTerm(e.target.value)} style={inp}>
                    {TERMS.filter(t => t !== currentTerm).map(t => <option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
                <div style={{ display: "flex", gap: 10 }}>
                  <button onClick={() => { setShowRollover(false); setRolloverDnr({}); setRolloverStep(1); }} style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
                  <button onClick={() => {
                    if (!rolloverTerm) setRolloverTerm(TERMS[TERMS.indexOf(currentTerm) + 1] || TERMS[0]);
                    setRolloverStep(2);
                  }} style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#f59e0b", color: "#0f172a", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Next →</button>
                </div>
              </>
            )}

            {rolloverStep === 2 && (() => {
              const nextTerm = rolloverTerm || TERMS[TERMS.indexOf(currentTerm) + 1] || TERMS[0];
              const dnrCount = Object.values(rolloverDnr).filter(Boolean).length;
              return (
                <>
                  <div style={{ fontSize: 17, fontWeight: 800, color: "#0f172a", marginBottom: 6 }}>🚶 Any Students Not Returning for {nextTerm}?</div>
                  <div style={{ fontSize: 13, color: "#64748b", marginBottom: 16 }}>
                    Mark any students who won't be continuing into {nextTerm} (e.g. transferred, dropped out, or stopped attending mid-year). They'll move to Alumni with their outstanding balance recorded as debt. Leave everyone as "Returning" if all students are continuing.
                  </div>

                  <div style={{ background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 12, padding: 16, marginBottom: 20, maxHeight: 320, overflowY: "auto" }}>
                    {schoolClasses.map(cls => {
                      const clsStudents = termStudents.filter(s => s.class === cls);
                      if (clsStudents.length === 0) return null;
                      return (
                        <div key={cls} style={{ marginBottom: 12 }}>
                          <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 6 }}>{cls} ({clsStudents.length} students)</div>
                          {clsStudents.map(s => {
                            const isDNR = rolloverDnr[s.id] || false;
                            const unpaid = getBalance(s, currentTerm).balance;
                            return (
                              <div key={s.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 10px", borderRadius: 8, marginBottom: 4, background: isDNR ? "#fef2f2" : "#fff", border: `1px solid ${isDNR ? "#fca5a5" : "#e2e8f0"}` }}>
                                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                  {s.photo ? <img src={s.photo} alt="" style={{ width: 24, height: 24, borderRadius: 6, objectFit: "cover" }} /> : <span style={{ fontSize: 16 }}>{s.gender === "F" ? "👩" : "👨"}</span>}
                                  <div>
                                    <span style={{ fontSize: 12, fontWeight: 600, color: isDNR ? "#b91c1c" : "#0f172a" }}>{s.name}</span>
                                    {s.stream && <span style={{ marginLeft: 6 }}><Pill text={s.stream} bg="#f5f3ff" col="#7c3aed" /></span>}
                                    {unpaid > 0 && <span style={{ fontSize: 10, color: "#f59e0b", marginLeft: 6, fontWeight: 600 }}>owes {fmt(unpaid)}</span>}
                                  </div>
                                </div>
                                <button onClick={() => setRolloverDnr(prev => ({ ...prev, [s.id]: !isDNR }))}
                                  style={{ padding: "4px 10px", borderRadius: 7, border: "none", cursor: "pointer", fontSize: 11, fontWeight: 700, background: isDNR ? "#ef4444" : "#f1f5f9", color: isDNR ? "#fff" : "#64748b" }}>
                                  {isDNR ? "✗ Did Not Return" : "Returning ✓"}
                                </button>
                              </div>
                            );
                          })}
                        </div>
                      );
                    })}
                  </div>

                  {dnrCount > 0 && (
                    <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 10, padding: 12, marginBottom: 16, fontSize: 12, color: "#b91c1c", fontWeight: 600 }}>
                      ⚠ {dnrCount} student(s) will be moved to Alumni with status "Did Not Return for {nextTerm}"
                    </div>
                  )}

                  <div style={{ display: "flex", gap: 10 }}>
                    <button onClick={() => setRolloverStep(1)} style={{ flex: 1, padding: 11, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>← Back</button>
                    <button onClick={() => handleRollover(nextTerm)} style={{ flex: 1, padding: 11, borderRadius: 9, border: "none", background: "#f59e0b", color: "#0f172a", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Confirm Rollover</button>
                  </div>
                </>
              );
            })()}
          </div>
        </div>
      )}

      {/* ════════ NEW ACADEMIC YEAR / PROMOTION MODAL ════════ */}
      {showPromotion && (() => {
        const lastClass = schoolClasses[schoolClasses.length - 1];
        const transitionClasses = getTransitionClasses(school);
        const graduateStudents = termStudents.filter(s => s.class === lastClass);
        const transitionGroups = transitionClasses
          .map(cls => ({ cls, students: termStudents.filter(s => s.class === cls) }))
          .filter(g => g.students.length > 0);
        const autoPromoteClasses = schoolClasses.filter(c => c !== lastClass && !transitionClasses.includes(c));
        return (
          <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, overflowY: "auto", padding: "20px 0" }}>
            <div style={{ background: "#fff", borderRadius: 18, padding: isMobile ? 18 : 32, width: isMobile ? "calc(100vw - 32px)" : 560, maxWidth: 560, boxShadow: "0 24px 60px rgba(0,0,0,0.25)", margin: "auto" }}>
              <div style={{ fontSize: 19, fontWeight: 800, color: "#0f172a", marginBottom: 4 }}>🎓 New Academic Year Promotion</div>
              <div style={{ fontSize: 13, color: "#64748b", marginBottom: 22 }}>Promote all students to the next class. All records and payment history are permanently preserved.</div>

              {/* Year picker */}
              <div style={{ marginBottom: 20 }}>
                <label style={lbl}>New Academic Year</label>
                <input value={promotionYear} onChange={e => setPromotionYear(e.target.value)} placeholder="e.g. 2026" style={{ ...inp, fontSize: 18, fontWeight: 800, color: "#0f172a" }} />
              </div>

              {/* Auto-promotions with DNR / Repeat options per class */}
              {autoPromoteClasses.some(cls => termStudents.some(s => s.class === cls)) && (
              <div style={{ background: "#f0fdf4", border: "1px solid #86efac", borderRadius: 12, padding: 16, marginBottom: 20 }}>
                <div style={{ fontWeight: 700, color: "#15803d", fontSize: 13, marginBottom: 4 }}>✓ Auto-Promoted Students</div>
                <div style={{ fontSize: 11, color: "#166534", marginBottom: 10 }}>Mark "Repeating" for any student who failed and must repeat the class, or "Did Not Return" if they've left.</div>
                {autoPromoteClasses.map(cls => {
                  const clsStudents = termStudents.filter(s => s.class === cls);
                  if (clsStudents.length === 0) return null;
                  const clsIdx = schoolClasses.indexOf(cls);
                  const nextCls = schoolClasses[clsIdx + 1];
                  return (
                    <div key={cls} style={{ marginBottom: 12 }}>
                      <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 6 }}>{cls} → {nextCls} ({clsStudents.length} students)</div>
                      {clsStudents.map(s => {
                        const isDNR = dnrDecisions[s.id] || false;
                        const isRepeating = repeatDecisions[s.id] || false;
                        const debt = getBalanceStatic(s, currentTerm, feeStructure, requirements).balance;
                        const rowBg = isDNR ? "#fef2f2" : isRepeating ? "#fffbeb" : "#fff";
                        const rowBorder = isDNR ? "#fca5a5" : isRepeating ? "#fde68a" : "#e2e8f0";
                        const nameColor = isDNR ? "#b91c1c" : isRepeating ? "#92400e" : "#0f172a";
                        return (
                          <div key={s.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 10px", borderRadius: 8, marginBottom: 4, background: rowBg, border: `1px solid ${rowBorder}`, flexWrap: "wrap", gap: 6 }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                              {s.photo ? <img src={s.photo} alt="" style={{ width: 24, height: 24, borderRadius: 6, objectFit: "cover" }} /> : <span style={{ fontSize: 16 }}>{s.gender === "F" ? "👩" : "👨"}</span>}
                              <div>
                                <span style={{ fontSize: 12, fontWeight: 600, color: nameColor }}>{s.name}{s.stream ? ` (${classLabel(cls, s.stream)})` : ""}</span>
                                {isRepeating && <span style={{ fontSize: 10, color: "#92400e", marginLeft: 6, fontWeight: 700 }}>repeating {cls}</span>}
                                {debt > 0 && <span style={{ fontSize: 10, color: "#f59e0b", marginLeft: 6, fontWeight: 600 }}>owes {fmt(debt)}</span>}
                              </div>
                            </div>
                            <div style={{ display: "flex", gap: 4 }}>
                              <button onClick={() => { setDnrDecisions(prev => ({ ...prev, [s.id]: false })); setRepeatDecisions(prev => ({ ...prev, [s.id]: false })); }}
                                style={{ padding: "4px 9px", borderRadius: 7, border: "none", cursor: "pointer", fontSize: 10, fontWeight: 700, background: (!isDNR && !isRepeating) ? "#10b981" : "#f1f5f9", color: (!isDNR && !isRepeating) ? "#fff" : "#64748b" }}>
                                Returning
                              </button>
                              <button onClick={() => { setRepeatDecisions(prev => ({ ...prev, [s.id]: true })); setDnrDecisions(prev => ({ ...prev, [s.id]: false })); }}
                                style={{ padding: "4px 9px", borderRadius: 7, border: "none", cursor: "pointer", fontSize: 10, fontWeight: 700, background: isRepeating ? "#f59e0b" : "#f1f5f9", color: isRepeating ? "#fff" : "#64748b" }}>
                                ↻ Repeating
                              </button>
                              <button onClick={() => { setDnrDecisions(prev => ({ ...prev, [s.id]: true })); setRepeatDecisions(prev => ({ ...prev, [s.id]: false })); }}
                                style={{ padding: "4px 9px", borderRadius: 7, border: "none", cursor: "pointer", fontSize: 10, fontWeight: 700, background: isDNR ? "#ef4444" : "#f1f5f9", color: isDNR ? "#fff" : "#64748b" }}>
                                ✗ Did Not Return
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  );
                })}
              </div>
              )}

              {/* Final-class graduates */}
              {graduateStudents.length > 0 && (
                <div style={{ background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 12, padding: 16, marginBottom: 20 }}>
                  <div style={{ fontWeight: 700, color: "#1d4ed8", fontSize: 13, marginBottom: 4 }}>🎓 {lastClass} Students</div>
                  <div style={{ fontSize: 11, color: "#1e40af", marginBottom: 10 }}>These students graduate by default. Mark "Repeating" for anyone who must repeat {lastClass} instead.</div>
                  {graduateStudents.map(s => {
                    const debt = getBalance(s, currentTerm).balance;
                    const isRepeating = repeatDecisions[s.id] || false;
                    return (
                      <div key={s.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 12, color: "#374151", padding: "6px 0", borderBottom: "1px solid #dbeafe", flexWrap: "wrap", gap: 6 }}>
                        <div>
                          <span style={{ fontWeight: 600 }}>{s.gender === "F" ? "👩" : "👨"} {s.name}</span>
                          {isRepeating && <span style={{ fontSize: 10, color: "#92400e", marginLeft: 6, fontWeight: 700 }}>repeating {lastClass}</span>}
                        </div>
                        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                          {!isRepeating && (debt > 0
                            ? <span style={{ color: "#dc2626", fontWeight: 700, fontSize: 11 }}>⚠ Owes {fmt(debt)} — recorded in Alumni</span>
                            : <span style={{ color: "#15803d", fontWeight: 700, fontSize: 11 }}>✓ Fully cleared</span>)}
                          <button onClick={() => setRepeatDecisions(prev => ({ ...prev, [s.id]: !isRepeating }))}
                            style={{ padding: "4px 9px", borderRadius: 7, border: "none", cursor: "pointer", fontSize: 10, fontWeight: 700, background: isRepeating ? "#f59e0b" : "#f1f5f9", color: isRepeating ? "#fff" : "#64748b" }}>
                            {isRepeating ? `↻ Repeating ${lastClass}` : "Mark as Repeating"}
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}

              {/* Transition-class decisions (P7, S4, Top Class, etc.) */}
              {transitionGroups.map(({ cls, students }) => {
                const clsIdx = schoolClasses.indexOf(cls);
                const nextCls = schoolClasses[clsIdx + 1];
                return (
                  <div key={cls} style={{ background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 12, padding: 16, marginBottom: 20 }}>
                    <div style={{ fontWeight: 700, color: "#92400e", fontSize: 13, marginBottom: 4 }}>⚠ {cls} Students — Decision Required</div>
                    <div style={{ fontSize: 11, color: "#78350f", marginBottom: 12 }}>For each {cls} student, choose whether they are continuing to {nextCls}, leaving (e.g. transferring to another school), or repeating {cls}. If they leave with a balance, it is permanently tracked in Alumni.</div>
                    {students.map(s => {
                      const debt = getBalance(s, currentTerm).balance;
                      const isRepeating = repeatDecisions[s.id] || false;
                      const decision = isRepeating ? "repeat" : (transitionDecisions[s.id] || "leave");
                      return (
                        <div key={s.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: "1px solid #fde68a", flexWrap: "wrap", gap: 8 }}>
                          <div>
                            <div style={{ fontWeight: 600, fontSize: 13, color: "#0f172a" }}>{s.gender === "F" ? "👩" : "👨"} {s.name}</div>
                            {isRepeating
                              ? <div style={{ fontSize: 11, color: "#92400e", fontWeight: 600 }}>↻ Repeating {cls}{debt > 0 ? ` — owes ${fmt(debt)}, carried as arrears` : ""}</div>
                              : debt > 0
                                ? <div style={{ fontSize: 11, color: "#dc2626", fontWeight: 600 }}>Owes {fmt(debt)}{decision === "leave" ? " — will be tracked in Alumni" : ` — will become arrears in ${nextCls}`}</div>
                                : <div style={{ fontSize: 11, color: "#15803d", fontWeight: 600 }}>✓ No outstanding balance</div>}
                          </div>
                          <div style={{ display: "flex", gap: 6 }}>
                            {[["continue", `→ ${nextCls}`, "#10b981", "#f0fdf4"], ["repeat", `↻ Repeat ${cls}`, "#f59e0b", "#fffbeb"], ["leave", "Leaving", "#ef4444", "#fef2f2"]].map(([val, label, col, bg]) => (
                              <button key={val} onClick={() => {
                                if (val === "repeat") { setRepeatDecisions(prev => ({ ...prev, [s.id]: true })); }
                                else { setRepeatDecisions(prev => ({ ...prev, [s.id]: false })); setTransitionDecisions(prev => ({ ...prev, [s.id]: val })); }
                              }}
                                style={{ padding: "6px 10px", borderRadius: 8, border: `2px solid ${decision === val ? col : "#e2e8f0"}`, background: decision === val ? bg : "#fff", color: decision === val ? col : "#94a3b8", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                                {label}
                              </button>
                            ))}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                );
              })}

              {/* Final warning */}
              <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 10, padding: 12, marginBottom: 22, fontSize: 12, color: "#991b1b" }}>
                <strong>⚠ This action cannot be undone.</strong> All students will be promoted to {promotionYear}. Current term is set to Term 1, {promotionYear}. All historical data stays permanently.
              </div>

              <div style={{ display: "flex", gap: 10 }}>
                <button onClick={() => { setShowPromotion(false); setTransitionDecisions({}); setDnrDecisions({}); setRepeatDecisions({}); }} style={{ flex: 1, padding: 12, borderRadius: 9, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#64748b" }}>Cancel</button>
                <button onClick={handlePromotion} style={{ flex: 1, padding: 12, borderRadius: 9, border: "none", background: "linear-gradient(135deg,#8b5cf6,#6d28d9)", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>
                  🎓 Confirm Promotion to {promotionYear}
                </button>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}
