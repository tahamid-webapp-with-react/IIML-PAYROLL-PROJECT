import { useEffect, useState, useMemo, type FormEvent, type ReactNode } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { SafeAreaTopScrim, fileToBase64 } from "@hatch/space-sdk/client";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
} from "recharts";
import { api } from "./api";
import { DataTable, type Column } from "./components/DataTable";
import { ConfirmDialog } from "./components/ConfirmDialog";
import { formatDate, formatDateTime, humanize, printDocument } from "./utils/formatters";
import * as XLSX from "xlsx";
import { loginWithGoogle, logoutFirebase, db } from "./firebase";
import { doc, setDoc } from "firebase/firestore";

type PageKey = "dashboard" | "periods" | "projects" | "configuration" | "rates" | "employees" | "users" | "roles" | "recycle-bin" | "audit" | "imports" | "production-records" | "production-work" | "production-shipments" | "production-settings" | "production-exceptions" | "batch-mapping" | "reconciliation" | "reconciliation-exceptions" | "salary-calculation" | "salary-details" | "master-salary" | "review-approval" | "export-center" | "planned-settings";
type ProductionStream = "doctype" | "proof" | "transmitted_batch" | "shipdet";
type SourceType = "employee_information" | "record_details" | "doctype" | "proof" | "transmitted_batch" | "shipdet";

type Notice = { kind: "success" | "error"; text: string } | null;
type EmployeeDraft = { id?: number; employeeId: string; name: string; previousIds: string; joiningDate: string; status: "active" | "inactive" | "leave" | "separated"; designation: string; contractType: string; team: string; location: string };

type CustomTierDraft = {
  id: string;
  label: string;
  minHours: string;
  maxHours: string;
  multiplier: string;
};

type SlabConfigDraft = {
  thresholdLow: string;
  thresholdHigh: string;
  multiplierLow: string;
  multiplierMid: string;
  multiplierHigh: string;
  customTiers: CustomTierDraft[];
};

type ProjectDraft = {
  id?: number;
  projectId: string;
  name: string;
  mainProjectId: string;
  subProjectName: string;
  client: string;
  description: string;
  startDate: string;
  endDate: string;
  status: "active" | "inactive" | "archived";
  currency: string;
  manager: string;
  notes: string;
  language: "english" | "nonenglish";
  countries: string[];
  averageWorkingHour: boolean;
  slabFormula: boolean;
  slabConfig: SlabConfigDraft;
  subProjectIncentives: boolean;
  docType: boolean;
  proof: boolean;
  paymentBasis: "record" | "character" | "batch";
  characterTypes: string[];
};
type CsvEmployee = Omit<EmployeeDraft, "id" | "previousIds"> & { previousIds: string[] };

const emptyEmployee: EmployeeDraft = { employeeId: "", name: "", previousIds: "", joiningDate: "", status: "active", designation: "", contractType: "", team: "", location: "" };
const emptyProject: ProjectDraft = {
  projectId: "",
  name: "",
  mainProjectId: "hw",
  subProjectName: "",
  client: "",
  description: "",
  startDate: new Date().toISOString().slice(0, 10),
  endDate: "",
  status: "active",
  currency: "BDT",
  manager: "",
  notes: "",
  language: "english",
  countries: ["BD"],
  averageWorkingHour: false,
  slabFormula: false,
  slabConfig: {
    thresholdLow: "3.99",
    thresholdHigh: "5.99",
    multiplierLow: "0.60",
    multiplierMid: "0.80",
    multiplierHigh: "1.00",
    customTiers: [],
  },
  subProjectIncentives: false,
  docType: false,
  proof: false,
  paymentBasis: "record",
  characterTypes: ["Header Character", "Date Character"],
};

const NAV: Array<{
  group: string;
  icon: string;
  directPage?: PageKey;
  items: Array<{ key: PageKey; label: string; icon: string; stage: 1 | 2 | 3 | 4 | 5 | 6 | "planned" }>;
}> = [
  {
    group: "Dashboard & Overview",
    icon: "database",
    directPage: "dashboard",
    items: [],
  },
  {
    group: "Project Control and Configuration",
    icon: "sliders",
    items: [
      { key: "periods", label: "Payroll Period", icon: "calendar", stage: 1 },
      { key: "projects", label: "Projects", icon: "folder", stage: 1 },
      { key: "rates", label: "Rates", icon: "calculator", stage: 1 },
      { key: "employees", label: "Employee Management", icon: "users", stage: 1 },
      { key: "batch-mapping", label: "Batch Mapping", icon: "link", stage: 1 },
    ],
  },
  {
    group: "Data and Production",
    icon: "inbox",
    items: [
      { key: "imports", label: "Data Import Center", icon: "upload", stage: 2 },
      { key: "production-records", label: "Combined Record Details", icon: "fileText", stage: 3 },
      { key: "production-work", label: "Doctype & Proof", icon: "check", stage: 3 },
      { key: "production-shipments", label: "Transmitted & Shipdet", icon: "truck", stage: 3 },
      { key: "production-exceptions", label: "Production Exceptions", icon: "alert", stage: 3 },
      { key: "production-settings", label: "Production Settings", icon: "sliders", stage: 3 },
    ],
  },
  {
    group: "Payroll",
    icon: "calculator",
    items: [
      { key: "reconciliation", label: "Reconciliation Center", icon: "check", stage: 4 },
      { key: "reconciliation-exceptions", label: "Reconciliation Exceptions", icon: "alert", stage: 4 },
      { key: "salary-calculation", label: "Salary Calculation", icon: "calculator", stage: 5 },
      { key: "salary-details", label: "Employee Salary Details", icon: "users", stage: 5 },
      { key: "master-salary", label: "Master Salary", icon: "database", stage: 6 },
      { key: "review-approval", label: "Review & Approval", icon: "shield", stage: 6 },
      { key: "export-center", label: "Export Center", icon: "download", stage: 6 },
    ],
  },
  {
    group: "Administration",
    icon: "shield",
    items: [
      { key: "audit", label: "Audit Logs", icon: "fileText", stage: 1 },
      { key: "users", label: "User Management", icon: "users", stage: 1 },
      { key: "roles", label: "Role Permissions", icon: "shield", stage: 1 },
      { key: "recycle-bin", label: "Recycle Bin", icon: "trash", stage: 1 },
      { key: "planned-settings", label: "System Settings", icon: "settings", stage: "planned" },
    ],
  },
];

const PAGE_TITLES: Record<PageKey, { title: string; subtitle: string }> = {
  dashboard: { title: "Dashboard", subtitle: "Operational readiness, safeguards, and cycle overview." },
  periods: { title: "Payroll periods", subtitle: "Monthly payroll cycles and project setup." },
  projects: { title: "Projects", subtitle: "Simplified project creation and configuration." },
  configuration: { title: "Projects", subtitle: "Simplified project creation and configuration." },
  rates: { title: "Rates", subtitle: "Rate and calculation rules." },
  employees: { title: "Employee management", subtitle: "Maintain stable Employee IDs, historical IDs, and employment attributes." },
  users: { title: "User management", subtitle: "Create accounts with server-enforced roles." },
  roles: { title: "Role permissions", subtitle: "Configure dynamic server-enforced roles and custom permission mappings." },
  "recycle-bin": { title: "Recycle bin", subtitle: "Restore or permanently delete previously removed entities." },
  audit: { title: "Audit log", subtitle: "Review persistent changes to foundation records and access administration." },
  imports: { title: "Data Import Center", subtitle: "Upload, map, validate, and commit traceable payroll source files to persistent storage." },
  "production-records": { title: "Combined Record Details", subtitle: "Filter, inspect, and export normalized Record Details, Review Key, Full Key, and category quantities." },
  "production-work": { title: "Doctype and Proof", subtitle: "Review separately normalized Doctype and Proof work with explicit employee attribution." },
  "production-shipments": { title: "Transmitted Batch and Shipdet", subtitle: "Inspect transmission states and confirm normalized shipment records before downstream reconciliation." },
  "production-settings": { title: "Production settings", subtitle: "Configure project-specific Doctype, Proof, key-variant, proof-type, and status-normalization policies." },
  "production-exceptions": { title: "Production exceptions", subtitle: "Review unmatched employees, unknown statuses, key overlaps, and other actionable normalization exceptions." },
  "batch-mapping": { title: "Batch mapping", subtitle: "Connect internal, client, and shipment identifiers through reviewable, approved relationships." },
  reconciliation: { title: "Reconciliation center", subtitle: "Run traceable stage comparisons, apply payable policies, and allocate authorized quantities." },
  "reconciliation-exceptions": { title: "Reconciliation exceptions", subtitle: "Resolve blocking attribution, mapping, rate, unit, and control failures before approval." },
  "salary-calculation": { title: "Salary calculation", subtitle: "Create preserved calculation versions from approved payable quantities and effective rates." },
  "salary-details": { title: "Employee salary details", subtitle: "Trace every component from payable quantity and applied rate through employee and project totals." },
  "master-salary": { title: "Master Salary", subtitle: "Consolidate approved project salary versions by stable Employee ID and reconcile every total." },
  "review-approval": { title: "Payroll review and approval", subtitle: "Move payroll through review, approval, locking, and exported status with separate role authority." },
  "export-center": { title: "Export center", subtitle: "Generate controlled Excel workbooks from the exact approved payroll database snapshot." },
  "planned-settings": { title: "System settings", subtitle: "Planned controls that do not belong to the Stage 1 foundation." },
};

function message(error: unknown) {
  const raw = (error instanceof Error ? error.message : String(error)).replace(/^Error:\s*/, "");
  const jsonStart = raw.indexOf("{");
  if (jsonStart >= 0) { try { const parsed = JSON.parse(raw.slice(jsonStart)) as { error?: string; message?: string }; if (parsed.error || parsed.message) return parsed.error ?? parsed.message ?? raw; } catch { /* keep the readable fallback below */ } }
  return raw.replace(/^action\s+\w+\s+failed:\s*\d+\s*/i, "");
}

function hasPermission(permissions: string[], permission: string) {
  return permissions.includes(permission);
}


function field(form: FormData, key: string) {
  return String(form.get(key) ?? "").trim();
}

function Icon({ name }: { name: string }) {
  const paths: Record<string, ReactNode> = {
    menu: <><path d="M4 7h16M4 12h16M4 17h16" /></>,
    close: <><path d="m6 6 12 12M18 6 6 18" /></>,
    chevron: <path d="m9 18 6-6-6-6" />,
    logout: <><path d="M10 17l5-5-5-5M15 12H3M15 5h4a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2h-4" /></>,
    plus: <path d="M12 5v14M5 12h14" />,
    database: <><ellipse cx="12" cy="5" rx="8" ry="3" /><path d="M4 5v6c0 1.7 3.6 3 8 3s8-1.3 8-3V5M4 11v6c0 1.7 3.6 3 8 3s8-1.3 8-3v-6" /></>,
    shield: <path d="M12 3 5 6v5c0 4.6 2.8 8 7 10 4.2-2 7-5.4 7-10V6l-7-3Z" />,
    check: <path d="M20 6 9 17l-5-5" />,
    lock: <><rect width="18" height="11" x="3" y="11" rx="2" ry="2" /><path d="M7 11V7a5 5 0 0 1 10 0v4" /></>,
    download: <><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" x2="12" y1="15" y2="3" /></>,
    upload: <><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="17 8 12 3 7 8" /><line x1="12" x2="12" y1="3" y2="15" /></>,
    calculator: <><rect width="16" height="20" x="4" y="2" rx="2" /><line x1="8" x2="16" y1="6" y2="6" /><path d="M8 10h.01M12 10h.01M16 10h.01M8 14h.01M12 14h.01M16 14h.01M8 18h.01M12 18h.01M16 18h.01" /></>,
    sliders: <><line x1="4" x2="4" y1="21" y2="14" /><line x1="4" x2="4" y1="10" y2="3" /><line x1="12" x2="12" y1="21" y2="12" /><line x1="12" x2="12" y1="8" y2="3" /><line x1="20" x2="20" y1="21" y2="16" /><line x1="20" x2="20" y1="12" y2="3" /><line x1="1" x2="7" y1="14" y2="14" /><line x1="9" x2="15" y1="8" y2="8" /><line x1="17" x2="23" y1="16" y2="16" /></>,
    calendar: <><rect width="18" height="18" x="3" y="4" rx="2" /><line x1="16" x2="16" y1="2" y2="6" /><line x1="8" x2="8" y1="2" y2="6" /><line x1="3" x2="21" y1="10" y2="10" /></>,
    folder: <><path d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z" /></>,
    settings: <><circle cx="12" cy="12" r="3" /><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1Z" /></>,
    users: <><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" /><circle cx="9" cy="7" r="4" /><path d="M22 21v-2a4 4 0 0 0-3-3.87" /><path d="M16 3.13a4 4 0 0 1 0 7.75" /></>,
    link: <><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" /><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" /></>,
    inbox: <><polyline points="22 12 16 12 14 15 10 15 8 12 2 12" /><path d="M5.45 5.11 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z" /></>,
    fileText: <><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="16" x2="8" y1="13" y2="13" /><line x1="16" x2="8" y1="17" y2="17" /><polyline points="10 9 9 9 8 9" /></>,
    truck: <><path d="M14 18V6a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2v11a1 1 0 0 0 1 1h2" /><path d="M15 18H9" /><path d="M19 18h2a1 1 0 0 0 1-1v-3.65a1 1 0 0 0-.22-.62l-3.48-4.35A1 1 0 0 0 17.52 8H14v10Z" /><circle cx="17" cy="18" r="2" /><circle cx="7" cy="18" r="2" /></>,
    alert: <><circle cx="12" cy="12" r="10" /><line x1="12" x2="12" y1="8" y2="12" /><line x1="12" x2="12.01" y1="16" y2="16" /></>,
    trash: <><polyline points="3 6 5 6 21 6" /><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" /></>,
    restore: <><polyline points="1 4 1 10 7 10" /><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10" /></>,
    edit: <><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" /><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" /></>,
    eye: <><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z" /><circle cx="12" cy="12" r="3" /></>,
    pencil: <><path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z" /><path d="m15 5 4 4" /></>,
    history: <><path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" /><path d="M3 3v5h5" /><path d="M12 7v5l4 2" /></>,
    printer: <><path d="M6 9V2h12v7" /><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2" /><rect width="12" height="8" x="6" y="14" /></>,
    copy: <><rect width="14" height="14" x="8" y="8" rx="2" ry="2" /><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2" /></>,
  };
  return <svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">{paths[name] ?? paths.database}</svg>;
}

function Loading() {
  return <div className="loading"><span /><p>Loading secure workspace…</p></div>;
}

function Empty({ title, body, action }: { title: string; body: string; action?: ReactNode }) {
  return <div className="empty-state"><div className="empty-mark"><Icon name="database" /></div><h3>{title}</h3><p>{body}</p>{action}</div>;
}

function Status({ value }: { value: string }) {
  const tone = ["active", "Approved", "Locked", "Exported"].includes(value) ? "good" : ["inactive", "archived", "separated"].includes(value) ? "quiet" : ["Draft", "Importing"].includes(value) ? "warn" : "info";
  return <span className={`status ${tone}`}>{value}</span>;
}

const PERIOD_LIFECYCLE_STAGES = [
  { key: "Draft", label: "Draft", step: 1, percent: 12, color: "#64748b" },
  { key: "Importing", label: "Ingestion", step: 2, percent: 28, color: "#0284c7" },
  { key: "Reconciliation Pending", label: "Reconciliation", step: 3, percent: 45, color: "#d97706" },
  { key: "Ready for Calculation", label: "Ready to Calc", step: 4, percent: 58, color: "#7c3aed" },
  { key: "Calculated", label: "Calculated", step: 5, percent: 72, color: "#9333ea" },
  { key: "Under Review", label: "Review", step: 6, percent: 85, color: "#2563eb" },
  { key: "Approved", label: "Approved", step: 7, percent: 93, color: "#059669" },
  { key: "Locked", label: "Locked", step: 8, percent: 98, color: "#0f5b4a" },
  { key: "Exported", label: "Exported", step: 9, percent: 100, color: "#047857" },
];

const PERIOD_CORE_MILESTONES = [
  { label: "Draft", minStep: 1, color: "#64748b" },
  { label: "Ingestion", minStep: 2, color: "#0284c7" },
  { label: "Reconciliation", minStep: 3, color: "#d97706" },
  { label: "Calculation", minStep: 4, color: "#9333ea" },
  { label: "Review", minStep: 6, color: "#2563eb" },
  { label: "Approved", minStep: 7, color: "#059669" },
  { label: "Locked", minStep: 8, color: "#0f5b4a" },
];

function PeriodLifecycleProgress({ status }: { status: string }) {
  const stage = PERIOD_LIFECYCLE_STAGES.find((s) => s.key === status) ?? {
    key: status,
    label: status,
    step: 1,
    percent: 15,
    color: "#64748b",
  };

  return (
    <div className="period-progress-cell">
      <div className="period-progress-header">
        <Status value={status} />
        <span className="period-progress-percent" title={`${stage.percent}% workflow progress`}>
          {stage.percent}%
        </span>
      </div>

      <div className="period-progress-bar-track" title={`${status}: ${stage.percent}% workflow complete`}>
        <div
          className="period-progress-bar-fill"
          style={{
            width: `${stage.percent}%`,
            backgroundColor: stage.color,
          }}
        />
      </div>

      <div className="period-milestone-dots" role="group" aria-label={`Milestones for ${status}`}>
        {PERIOD_CORE_MILESTONES.map((m) => {
          const isCompleted = stage.step > m.minStep;
          const isCurrent = stage.step === m.minStep || (m.minStep === 4 && stage.step === 5);
          const stateLabel = isCompleted ? "Completed" : isCurrent ? "Current active stage" : "Upcoming stage";
          return (
            <span
              key={m.label}
              className={`period-dot-wrap ${isCurrent ? "current" : isCompleted ? "completed" : "pending"}`}
              title={`${m.label}: ${stateLabel}`}
            >
              <span
                className="period-dot"
                style={{
                  backgroundColor: isCompleted ? "#059669" : isCurrent ? m.color : "var(--border)",
                  boxShadow: isCurrent ? `0 0 0 3px ${m.color}33` : undefined,
                }}
              />
            </span>
          );
        })}
      </div>
    </div>
  );
}


function Modal({ title, children, onClose, wide = false }: { title: string; children: ReactNode; onClose: () => void; wide?: boolean }) {
  return <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.currentTarget === event.target) onClose(); }}><section className={`modal ${wide ? "modal-wide" : ""}`} role="dialog" aria-modal="true" aria-labelledby="modal-title"><div className="modal-head"><h2 id="modal-title">{title}</h2><button className="icon-button" aria-label="Close dialog" onClick={onClose}><Icon name="close" /></button></div>{children}</section></div>;
}

function AuthGate({ onAuthenticated }: { onAuthenticated: (token: string) => void }) {
  const [mode, setMode] = useState<"loading" | "setup" | "login">("loading");
  const [notice, setNotice] = useState<Notice>(null);
  const [firebaseConnecting, setFirebaseConnecting] = useState(false);
  const status = useQuery({ queryKey: ["auth-status"], queryFn: () => api.authStatus({}) });
  useEffect(() => { if (status.data) setMode(status.data.needsSetup ? "setup" : "login"); }, [status.data]);
  const setup = useMutation({ mutationFn: api.bootstrapAdmin, onSuccess: (data) => onAuthenticated(data.token), onError: (error) => setNotice({ kind: "error", text: message(error) }) });
  const login = useMutation({ mutationFn: api.login, onSuccess: (data) => onAuthenticated(data.token), onError: (error) => setNotice({ kind: "error", text: message(error) }) });
  const firebaseLoginMutation = useMutation({
    mutationFn: api.firebaseLogin,
    onSuccess: (data) => onAuthenticated(data.token),
    onError: (error) => setNotice({ kind: "error", text: message(error) }),
  });

  if (mode === "loading") return <Loading />;

  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault(); setNotice(null); const form = new FormData(event.currentTarget);
    const payload = { email: field(form, "email"), password: field(form, "password") };
    if (mode === "setup") setup.mutate({ ...payload, name: field(form, "name") }); else login.mutate(payload);
  };

  const handleGoogleSignIn = async () => {
    try {
      setFirebaseConnecting(true);
      setNotice(null);
      const user = await loginWithGoogle();
      if (!user || !user.email) {
        throw new Error("Unable to retrieve account details from Google Sign-In.");
      }
      try {
        await setDoc(
          doc(db, "users", user.uid),
          {
            uid: user.uid,
            email: user.email,
            displayName: user.displayName || user.email.split("@")[0],
            photoURL: user.photoURL || null,
            lastLoginAt: new Date().toISOString(),
          },
          { merge: true }
        );
      } catch (e) {
        console.warn("Firestore user sync note:", e);
      }
      await firebaseLoginMutation.mutateAsync({
        uid: user.uid,
        email: user.email,
        name: user.displayName || user.email.split("@")[0],
        avatar: user.photoURL || null,
      });
    } catch (err) {
      setNotice({ kind: "error", text: message(err) });
    } finally {
      setFirebaseConnecting(false);
    }
  };

  const pending = setup.isPending || login.isPending || firebaseConnecting;
  return (
    <div className="auth-shell">
      <SafeAreaTopScrim backgroundColor="var(--evergreen)" />
      <section className="auth-intro">
        <div className="auth-seal"><Icon name="shield" /></div>
        <p className="overline">Secure month-end operations</p>
        <h1>Build the payroll foundation before the numbers move.</h1>
        <p>Persistent employee identities, effective-dated rates, controlled periods, and server-enforced roles are the starting point for every calculation.</p>
        <div className="auth-proof"><span>01</span><div><strong>Persistent records</strong><small>Foundation data is written to the platform database.</small></div></div>
        <div className="auth-proof"><span>02</span><div><strong>Seven defined roles</strong><small>Every write is checked by the server.</small></div></div>
        <div className="auth-proof"><span>03</span><div><strong>No payroll assumptions</strong><small>Projects, categories, and rates start empty until configured.</small></div></div>
      </section>
      <section className="auth-panel">
        <div className="auth-form-wrap">
          <p className="overline">{mode === "setup" ? "First-time setup" : "Authorized access"}</p>
          <h2>{mode === "setup" ? "Create the administrator account" : "Sign in to continue"}</h2>
          <p className="form-lead">{mode === "setup" ? "This account initializes the role matrix and owns the first audit entry." : "Use an active account created by an administrator or sign in with Google."}</p>
          {notice && <div className={`notice ${notice.kind}`}>{notice.text}</div>}
          <form className="stack-form" onSubmit={submit}>
            {mode === "setup" && <label>Full name<input name="name" autoComplete="name" minLength={2} required /></label>}
            <label>Work email<input name="email" type="email" autoComplete="email" required /></label>
            <label>Password<input name="password" type="password" autoComplete={mode === "setup" ? "new-password" : "current-password"} minLength={mode === "setup" ? 10 : 1} required /></label>
            {mode === "setup" && <p className="field-help">Use at least 10 characters. Passwords are stored as one-way hashes.</p>}
            <button className="primary-button" disabled={pending}>{pending ? "Working…" : mode === "setup" ? "Create administrator" : "Sign in"}</button>
          </form>

          <div style={{ display: "flex", flexDirection: "column", gap: "0.75rem", marginTop: "1.25rem" }}>
            <div style={{ display: "flex", alignItems: "center", gap: "0.5rem", color: "var(--muted)", fontSize: "0.8125rem" }}>
              <div style={{ flex: 1, height: "1px", background: "var(--border)" }} />
              <span>or sign in with</span>
              <div style={{ flex: 1, height: "1px", background: "var(--border)" }} />
            </div>
            <button
              type="button"
              className="secondary-button"
              onClick={handleGoogleSignIn}
              disabled={pending}
              style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "0.625rem", width: "100%", padding: "0.625rem 1rem", fontWeight: 500 }}
            >
              <svg width="18" height="18" viewBox="0 0 24 24" aria-hidden="true">
                <path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" />
                <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" />
                <path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.06H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.94l2.85-2.22.81-.63z" />
                <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.06l3.66 2.84c.87-2.6 3.3-4.52 6.16-4.52z" />
              </svg>
              {firebaseConnecting ? "Connecting to Firebase…" : "Sign in with Google"}
            </button>
          </div>
        </div>
      </section>
    </div>
  );
}

function ProfileModal({
  isOpen,
  onClose,
  token,
  currentName,
  currentAvatar,
  notify,
}: {
  isOpen: boolean;
  onClose: () => void;
  token: string;
  currentName: string;
  currentAvatar: string | null | undefined;
  notify: (notice: Notice) => void;
}) {
  const qc = useQueryClient();
  const [name, setName] = useState(currentName);
  const [avatar, setAvatar] = useState(currentAvatar ?? "");
  const update = useMutation({
    mutationFn: api.updateProfile,
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ["session"] });
      onClose();
      notify({ kind: "success", text: "Profile updated successfully." });
    },
    onError: (err) => notify({ kind: "error", text: message(err) }),
  });

  const handleFileUpload = async (file: File) => {
    try {
      const b64 = await fileToBase64(file);
      setAvatar(`data:${b64.mimeType};base64,${b64.dataBase64}`);
    } catch {
      notify({ kind: "error", text: "Could not read image file." });
    }
  };

  if (!isOpen) return null;
  return (
    <Modal title="User profile & avatar" onClose={onClose}>
      <form
        className="stack-form"
        onSubmit={(e) => {
          e.preventDefault();
          update.mutate({ token, name, avatar: avatar || "" });
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 16, marginBottom: 12 }}>
          <div style={{ width: 64, height: 64, borderRadius: "50%", background: "#d8ece4", display: "grid", placeItems: "center", overflow: "hidden", flexShrink: 0, border: "2px solid #0f5b4a" }}>
            {avatar ? (
              <img src={avatar} alt="Profile" style={{ width: "100%", height: "100%", objectFit: "cover" }} />
            ) : (
              <span style={{ fontSize: 24, fontWeight: 800, color: "#073b33" }}>{name.slice(0, 1).toUpperCase()}</span>
            )}
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            <label className="secondary-button compact" style={{ cursor: "pointer" }}>
              <Icon name="upload" /> Upload photo
              <input type="file" accept="image/*" style={{ display: "none" }} onChange={(e) => { const f = e.target.files?.[0]; if (f) void handleFileUpload(f); }} />
            </label>
            {avatar && (
              <button type="button" className="ghost-button compact" style={{ color: "var(--danger)" }} onClick={() => setAvatar("")}>
                <Icon name="trash" /> Remove photo
              </button>
            )}
          </div>
        </div>
        <label>
          Full name
          <input value={name} onChange={(e) => setName(e.target.value)} required />
        </label>
        <label>
          Avatar Image URL (or upload)
          <input value={avatar} onChange={(e) => setAvatar(e.target.value)} />
        </label>
        <div className="form-actions">
          <button type="button" className="ghost-button" onClick={onClose}>Cancel</button>
          <button className="primary-button" disabled={update.isPending}>
            {update.isPending ? "Saving…" : "Save profile"}
          </button>
        </div>
      </form>
    </Modal>
  );
}

function BrandingModal({
  isOpen,
  onClose,
  token,
  currentAppName,
  currentSubtitle,
  currentLogoUrl,
  notify,
}: {
  isOpen: boolean;
  onClose: () => void;
  token: string;
  currentAppName: string;
  currentSubtitle: string;
  currentLogoUrl: string;
  notify: (notice: Notice) => void;
}) {
  const qc = useQueryClient();
  const [appName, setAppName] = useState(currentAppName);
  const [subtitle, setSubtitle] = useState(currentSubtitle);
  const [logoUrl, setLogoUrl] = useState(currentLogoUrl);
  const update = useMutation({
    mutationFn: api.saveAppSettings,
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ["app-settings"] });
      onClose();
      notify({ kind: "success", text: "Application branding updated." });
    },
    onError: (err) => notify({ kind: "error", text: message(err) }),
  });

  const handleLogoUpload = async (file: File) => {
    try {
      const b64 = await fileToBase64(file);
      setLogoUrl(`data:${b64.mimeType};base64,${b64.dataBase64}`);
    } catch {
      notify({ kind: "error", text: "Could not read logo file." });
    }
  };

  if (!isOpen) return null;
  return (
    <Modal title="Edit application branding" onClose={onClose}>
      <form
        className="stack-form"
        onSubmit={(e) => {
          e.preventDefault();
          update.mutate({ token, appName, subtitle, logoUrl });
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 10 }}>
          <div className="app-logo-badge" style={{ width: 48, height: 48, fontSize: 20 }}>
            {logoUrl ? <img src={logoUrl} alt="Logo" /> : <Icon name="database" />}
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            <label className="secondary-button compact" style={{ cursor: "pointer" }}>
              <Icon name="upload" /> Upload logo image
              <input type="file" accept="image/*" style={{ display: "none" }} onChange={(e) => { const f = e.target.files?.[0]; if (f) void handleLogoUpload(f); }} />
            </label>
            {logoUrl && (
              <button type="button" className="ghost-button compact" style={{ color: "var(--danger)" }} onClick={() => setLogoUrl("")}>
                <Icon name="trash" /> Reset logo
              </button>
            )}
          </div>
        </div>
        <label>
          Application name
          <input value={appName} onChange={(e) => setAppName(e.target.value)} required />
        </label>
        <label>
          Subtitle
          <input value={subtitle} onChange={(e) => setSubtitle(e.target.value)} required />
        </label>
        <label>
          Logo Image URL (or upload)
          <input value={logoUrl} onChange={(e) => setLogoUrl(e.target.value)} />
        </label>
        <div className="form-actions">
          <button type="button" className="ghost-button" onClick={onClose}>Cancel</button>
          <button className="primary-button" disabled={update.isPending}>
            {update.isPending ? "Saving…" : "Save branding"}
          </button>
        </div>
      </form>
    </Modal>
  );
}

export function App() {
  const queryClient = useQueryClient();
  const [token, setToken] = useState(() => sessionStorage.getItem("payroll-session") ?? "");
  const [page, setPage] = useState<PageKey>("dashboard");
  const [navOpen, setNavOpen] = useState(false);
  const [collapsed, setCollapsed] = useState(false);
  const [notice, setNotice] = useState<Notice>(null);
  const [openGroup, setOpenGroup] = useState<string>("Control");
  const [profileModalOpen, setProfileModalOpen] = useState(false);
  const [brandModalOpen, setBrandModalOpen] = useState(false);

  const session = useQuery({ queryKey: ["session", token], queryFn: () => api.sessionInfo({ token }), enabled: Boolean(token), retry: false });
  const appSettings = useQuery({ queryKey: ["app-settings", token], queryFn: () => api.getAppSettings({ token }), enabled: Boolean(token) });

  useEffect(() => {
    if (session.error) { sessionStorage.removeItem("payroll-session"); setToken(""); }
  }, [session.error]);

  const authenticate = (nextToken: string) => { sessionStorage.setItem("payroll-session", nextToken); setToken(nextToken); };
  const logout = async () => { try { await api.logout({ token }); } finally { logoutFirebase().catch(() => {}); sessionStorage.removeItem("payroll-session"); setToken(""); queryClient.clear(); } };

  if (!token) return <AuthGate onAuthenticated={authenticate} />;
  if (session.isPending || !session.data) return <Loading />;
  const permissions = session.data.permissions;

  const toggleGroup = (groupName: string) => {
    setOpenGroup((prev) => (prev === groupName ? "" : groupName));
  };

  const go = (key: PageKey) => {
    setPage(key);
    setNavOpen(false);
    setNotice(null);
    const found = NAV.find((g) => g.items.some((i) => i.key === key));
    if (found) {
      setOpenGroup(found.group);
    }
  };

  const title = PAGE_TITLES[page];
  const brandAppName = appSettings.data?.appName || "IIML PAYROLL";
  const brandSubtitle = appSettings.data?.subtitle || "Month-end Control & Payroll System";
  const brandLogo = appSettings.data?.logoUrl || "";

  return (
    <div className={`app-shell ${collapsed ? "is-collapsed" : ""}`}>
      <SafeAreaTopScrim backgroundColor="var(--evergreen)" />
      <aside className={`sidebar ${navOpen ? "open" : ""}`} style={{ width: collapsed ? undefined : "365px" }}>
        <div className="side-context">
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
            <button
              type="button"
              className="app-branding-box"
              onClick={() => go("dashboard")}
              title="Navigate to Dashboard"
            >
              <div className="app-logo-badge">
                {brandLogo ? <img src={brandLogo} alt="Logo" /> : <Icon name="database" />}
              </div>
              {!collapsed && (
                <div className="app-branding-text">
                  <h2>{brandAppName}</h2>
                  <p>{brandSubtitle}</p>
                </div>
              )}
            </button>
            {!collapsed && (
              <button
                type="button"
                className="brand-edit-btn"
                title="Edit branding"
                onClick={() => setBrandModalOpen(true)}
              >
                <Icon name="settings" />
              </button>
            )}
            <button
              className="icon-button side-collapse"
              aria-label={collapsed ? "Expand sidebar" : "Collapse sidebar"}
              onClick={() => setCollapsed(!collapsed)}
            >
              <Icon name="chevron" />
            </button>
          </div>
        </div>

        <nav aria-label="Primary navigation">
          {NAV.map((group) => {
            if (group.directPage) {
              const isActive = page === group.directPage;
              return (
                <div className="nav-group" key={group.group}>
                  <button
                    type="button"
                    className={`nav-group-header direct-nav-btn ${isActive ? "active open" : ""}`}
                    onClick={() => go(group.directPage!)}
                    title={group.group}
                    style={{
                      cursor: "pointer",
                      background: isActive ? "rgba(15, 91, 74, 0.45)" : "transparent",
                      color: isActive ? "#ffffff" : "#8bbcb0",
                    }}
                  >
                    <div className="group-title-wrap">
                      <Icon name={group.icon} />
                      {!collapsed && <span>{group.group}</span>}
                    </div>
                  </button>
                </div>
              );
            }
            const isOpen = openGroup === group.group;
            return (
              <div className="nav-group" key={group.group}>
                <button
                  type="button"
                  className={`nav-group-header ${isOpen ? "open" : ""}`}
                  onClick={() => toggleGroup(group.group)}
                  title={group.group}
                >
                  <div className="group-title-wrap">
                    <Icon name={group.icon} />
                    {!collapsed && <span>{group.group}</span>}
                  </div>
                  {!collapsed && (
                    <span className="nav-group-chevron">
                      <Icon name="chevron" />
                    </span>
                  )}
                </button>
                {isOpen && (
                  <div className="nav-sub-items">
                    {group.items.map((item, index) => (
                      <button
                        key={`${item.key}-${index}`}
                        className={page === item.key ? "active" : ""}
                        onClick={() => go(item.key)}
                        title={collapsed ? item.label : undefined}
                      >
                        <span className="nav-index">
                          <Icon name={item.icon} />
                        </span>
                        <span className="nav-label">{item.label}</span>
                        {item.stage === "planned" && <span className="planned-dot" aria-label="Planned" />}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </nav>

        <div className="side-user">
          <button
            type="button"
            className="side-user-avatar-btn"
            onClick={() => setProfileModalOpen(true)}
            title="Profile & avatar"
          >
            {session.data.avatar ? (
              <img src={session.data.avatar} alt={session.data.name} className="user-avatar-img" />
            ) : (
              <span>{session.data.name.slice(0, 1).toUpperCase()}</span>
            )}
          </button>
          <div>
            <strong>{session.data.name}</strong>
            <small>{session.data.roleName}</small>
          </div>
          <button className="icon-button" aria-label="Sign out" onClick={logout} title="Sign out">
            <Icon name="logout" />
          </button>
        </div>
      </aside>

      {navOpen && <button className="nav-scrim" aria-label="Close navigation" onClick={() => setNavOpen(false)} />}

      <main className="workspace">
        <header className="workspace-head">
          <button className="icon-button mobile-menu" aria-label="Open navigation" onClick={() => setNavOpen(true)}>
            <Icon name="menu" />
          </button>
          <div>
            <p className="breadcrumb">
              {page === "imports" ? "Stage 2" : page.startsWith("production-") ? "Stage 3" : ["batch-mapping", "reconciliation", "reconciliation-exceptions"].includes(page) ? "Stage 4" : ["salary-calculation", "salary-details"].includes(page) ? "Stage 5" : ["master-salary", "review-approval", "export-center"].includes(page) ? "Stage 6" : "Stage 1"} / {title.title}
            </p>
            <h1>{title.title}</h1>
            <p>{title.subtitle}</p>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: "0.75rem", marginLeft: "auto" }}>
            <span
              className="status good"
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: "0.375rem",
                fontSize: "0.75rem",
                padding: "0.25rem 0.625rem",
                borderRadius: "9999px",
                fontWeight: 500,
              }}
              title="Firebase Firestore & Auth Connected (gen-lang-client-0103137916)"
            >
              <span style={{ width: "6px", height: "6px", borderRadius: "50%", background: "currentColor" }} />
              Firebase Connected
            </span>
            <div className="identity-chip" style={{ cursor: "pointer" }} onClick={() => setProfileModalOpen(true)} title="Edit profile">
              {session.data.avatar ? (
                <img src={session.data.avatar} alt={session.data.name} className="user-avatar-img" />
              ) : (
                <span>{session.data.name.slice(0, 1).toUpperCase()}</span>
              )}
              <div>
                <strong>{session.data.name}</strong>
                <small>{session.data.roleName}</small>
              </div>
            </div>
          </div>
        </header>

        {notice && (
          <div className={`notice page-notice ${notice.kind}`}>
            {notice.text}
            <button aria-label="Dismiss message" onClick={() => setNotice(null)}>×</button>
          </div>
        )}

        <div className="page-body">
          {page === "dashboard" && <Dashboard token={token} permissions={permissions} go={go} />}
          {page === "employees" && <Employees token={token} permissions={permissions} notify={setNotice} />}
          {page === "projects" && <Projects token={token} permissions={permissions} notify={setNotice} />}
          {page === "periods" && <Periods token={token} permissions={permissions} notify={setNotice} />}
          {page === "configuration" && <Projects token={token} permissions={permissions} notify={setNotice} />}
          {page === "rates" && <Configuration token={token} permissions={permissions} notify={setNotice} mode="rates" />}
          {page === "imports" && <ImportCenter token={token} permissions={permissions} notify={setNotice} />}
          {page === "production-records" && <RecordDetails token={token} permissions={permissions} notify={setNotice} />}
          {page === "production-work" && <ProductionWork token={token} permissions={permissions} notify={setNotice} />}
          {page === "production-shipments" && <ProductionShipments token={token} permissions={permissions} notify={setNotice} />}
          {page === "production-settings" && <ProductionSettings token={token} permissions={permissions} notify={setNotice} />}
          {page === "production-exceptions" && <ProductionExceptions token={token} permissions={permissions} />}
          {page === "batch-mapping" && <BatchMapping token={token} permissions={permissions} notify={setNotice} />}
          {page === "reconciliation" && <ReconciliationCenter token={token} permissions={permissions} notify={setNotice} />}
          {page === "reconciliation-exceptions" && <ReconciliationExceptions token={token} permissions={permissions} notify={setNotice} />}
          {page === "salary-calculation" && <SalaryCalculation token={token} permissions={permissions} notify={setNotice} />}
          {page === "salary-details" && <SalaryDetails token={token} permissions={permissions} notify={setNotice} />}
          {page === "master-salary" && <MasterSalary token={token} permissions={permissions} notify={setNotice} />}
          {page === "review-approval" && <ReviewApproval token={token} permissions={permissions} notify={setNotice} />}
          {page === "export-center" && <ExportCenter token={token} permissions={permissions} notify={setNotice} />}
          {page === "users" && <Users token={token} permissions={permissions} notify={setNotice} />}
          {page === "roles" && <Roles token={token} permissions={permissions} notify={setNotice} />}
          {page === "recycle-bin" && <RecycleBin token={token} permissions={permissions} notify={setNotice} />}
          {page === "audit" && <Audit token={token} permissions={permissions} notify={setNotice} />}
          {page.startsWith("planned-") && <Planned page={page} go={go} />}
        </div>
      </main>

      <ProfileModal
        isOpen={profileModalOpen}
        onClose={() => setProfileModalOpen(false)}
        token={token}
        currentName={session.data.name}
        currentAvatar={session.data.avatar}
        notify={setNotice}
      />

      <BrandingModal
        isOpen={brandModalOpen}
        onClose={() => setBrandModalOpen(false)}
        token={token}
        currentAppName={brandAppName}
        currentSubtitle={brandSubtitle}
        currentLogoUrl={brandLogo}
        notify={setNotice}
      />
    </div>
  );
}






function Dashboard({ token, permissions, go }: { token: string; permissions: string[]; go: (page: PageKey) => void }) {
  const query = useQuery({ queryKey: ["dashboard", token], queryFn: () => api.getDashboard({ token }), enabled: hasPermission(permissions, "dashboard.read") });
  if (!hasPermission(permissions, "dashboard.read")) return <Empty title="Dashboard access is restricted" body="Your role does not include dashboard access." />;
  if (query.isPending) return <Loading />;
  if (query.error || !query.data) return <Empty title="Dashboard could not load" body={message(query.error)} />;
  const d = query.data;

  const totalPayrollFormatted = `${d.currency || "BDT"} ${Number(d.totalPayrollAmount || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  const activeRatePercent = d.employees > 0 ? Math.round((d.activeEmployees / d.employees) * 100) : 0;

  return (
    <div className="dashboard-summary-page">
      <div className="dashboard-summary-header">
        <div className="dashboard-summary-header-left">
          <h2>Overview & Summary Metrics</h2>
          <p>Key operational workforce, project, and payroll indicators for the active operational cycle.</p>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          {d.latestPeriod && (
            <span className="summary-card-badge success">
              <span style={{ width: 6, height: 6, borderRadius: "50%", background: "#059669" }} />
              Active Period: <strong>{d.latestPeriod.periodId}</strong>
            </span>
          )}
        </div>
      </div>

      <div className="dashboard-summary-grid">
        {/* 1. Total Number of Employee */}
        <div
          className="summary-card theme-teal"
          onClick={() => go("employees")}
          role="button"
          tabIndex={0}
          onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); go("employees"); } }}
        >
          <div className="summary-card-top">
            <div className="summary-card-title-group">
              <span className="summary-card-label">Total Number of Employee</span>
              <span className="summary-card-badge neutral">Master Register</span>
            </div>
            <div className="summary-card-icon-wrap teal">
              <Icon name="users" />
            </div>
          </div>
          <div className="summary-card-body">
            <div className="summary-card-value">{d.employees}</div>
          </div>
          <div className="summary-card-foot">
            <span>Total registered workforce records</span>
            <span className="summary-card-foot-link">View employees →</span>
          </div>
        </div>

        {/* 2. Active Payroll Period */}
        <div
          className="summary-card theme-blue"
          onClick={() => go("periods")}
          role="button"
          tabIndex={0}
          onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); go("periods"); } }}
        >
          <div className="summary-card-top">
            <div className="summary-card-title-group">
              <span className="summary-card-label">Active Payroll Period</span>
              <span className={`summary-card-badge ${d.latestPeriod ? "info" : "neutral"}`}>
                {d.latestPeriod ? d.latestPeriod.status : "No Active Cycle"}
              </span>
            </div>
            <div className="summary-card-icon-wrap blue">
              <Icon name="calendar" />
            </div>
          </div>
          <div className="summary-card-body">
            <div className="summary-card-value" style={{ fontSize: d.latestPeriod ? "clamp(19px, 1.8vw, 25px)" : undefined }}>
              {d.latestPeriod ? d.latestPeriod.periodId : "None"}
            </div>
          </div>
          <div className="summary-card-foot">
            <span>{d.latestPeriod ? `${formatDate(d.latestPeriod.startDate)} – ${formatDate(d.latestPeriod.endDate)}` : "Create new cycle"}</span>
            <span className="summary-card-foot-link">Manage periods →</span>
          </div>
        </div>

        {/* 3. Number of Project Active */}
        <div
          className="summary-card theme-indigo"
          onClick={() => go("projects")}
          role="button"
          tabIndex={0}
          onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); go("projects"); } }}
        >
          <div className="summary-card-top">
            <div className="summary-card-title-group">
              <span className="summary-card-label">Number of Project Active</span>
              <span className="summary-card-badge info">
                {d.activeProjects} / {d.totalProjects ?? d.activeProjects} Active
              </span>
            </div>
            <div className="summary-card-icon-wrap indigo">
              <Icon name="folder" />
            </div>
          </div>
          <div className="summary-card-body">
            <div className="summary-card-value">{d.activeProjects}</div>
          </div>
          <div className="summary-card-foot">
            <span>Operational projects in scope</span>
            <span className="summary-card-foot-link">View projects →</span>
          </div>
        </div>

        {/* 4. Number of Employees */}
        <div
          className="summary-card theme-emerald"
          onClick={() => go("employees")}
          role="button"
          tabIndex={0}
          onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); go("employees"); } }}
        >
          <div className="summary-card-top">
            <div className="summary-card-title-group">
              <span className="summary-card-label">Number of Employees</span>
              <span className="summary-card-badge success">
                {activeRatePercent}% Active
              </span>
            </div>
            <div className="summary-card-icon-wrap emerald">
              <Icon name="check" />
            </div>
          </div>
          <div className="summary-card-body">
            <div className="summary-card-value">{d.activeEmployees}</div>
          </div>
          <div className="summary-card-foot">
            <span>Active employees on payroll</span>
            <span className="summary-card-foot-link">View active →</span>
          </div>
        </div>

        {/* 5. Total Payroll Amount */}
        <div
          className="summary-card theme-amber"
          onClick={() => go("master-salary")}
          role="button"
          tabIndex={0}
          onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); go("master-salary"); } }}
        >
          <div className="summary-card-top">
            <div className="summary-card-title-group">
              <span className="summary-card-label">Total Payroll Amount</span>
              <span className="summary-card-badge warning">
                {d.latestPeriod ? d.latestPeriod.periodId : "Current"}
              </span>
            </div>
            <div className="summary-card-icon-wrap amber">
              <Icon name="calculator" />
            </div>
          </div>
          <div className="summary-card-body">
            <div className="summary-card-value" style={{ fontSize: "clamp(18px, 1.6vw, 24px)" }}>
              {totalPayrollFormatted}
            </div>
          </div>
          <div className="summary-card-foot">
            <span>Calculated / approved payroll gross</span>
            <span className="summary-card-foot-link">View salary →</span>
          </div>
        </div>
      </div>
    </div>
  );
}

function Employees({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const queryClient = useQueryClient();
  const [modal, setModal] = useState(false);
  const [viewEmployee, setViewEmployee] = useState<any | null>(null);
  const [deleteTarget, setDeleteTarget] = useState<any | null>(null);
  const [draft, setDraft] = useState<EmployeeDraft>(emptyEmployee);
  const [csvRows, setCsvRows] = useState<CsvEmployee[]>([]);
  const [csvName, setCsvName] = useState("");
  const canManage = hasPermission(permissions, "employees.manage");
  const canImport = hasPermission(permissions, "employees.import");

  const query = useQuery({ queryKey: ["employees", token, ""], queryFn: () => api.listEmployees({ token, search: "" }), enabled: hasPermission(permissions, "employees.read") });
  const save = useMutation({
    mutationFn: api.saveEmployee,
    onSuccess: (data) => {
      queryClient.invalidateQueries({ queryKey: ["employees"] });
      queryClient.invalidateQueries({ queryKey: ["dashboard"] });
      setModal(false);
      setDraft(emptyEmployee);
      notify({ kind: "success", text: `Employee ${data.mode} and persisted.` });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const removeMutation = useMutation({
    mutationFn: api.deleteEmployee,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["employees"] });
      queryClient.invalidateQueries({ queryKey: ["recycle-bin"] });
      queryClient.invalidateQueries({ queryKey: ["dashboard"] });
      setDeleteTarget(null);
      notify({ kind: "success", text: "Employee moved to Recycle Bin." });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const importer = useMutation({
    mutationFn: api.importEmployees,
    onSuccess: (data) => {
      queryClient.invalidateQueries({ queryKey: ["employees"] });
      queryClient.invalidateQueries({ queryKey: ["dashboard"] });
      setCsvRows([]);
      setCsvName("");
      notify({ kind: data.rejected.length ? "error" : "success", text: `${data.created} created, ${data.updated} updated, ${data.rejected.length} rejected.` });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const openEdit = (employee: NonNullable<typeof query.data>["employees"][number]) => {
    setDraft({
      id: employee.id,
      employeeId: employee.employeeId,
      name: employee.name,
      previousIds: employee.previousIds.join(", "),
      joiningDate: employee.joiningDate ?? "",
      status: employee.status as EmployeeDraft["status"],
      designation: employee.designation ?? "",
      contractType: employee.contractType ?? "",
      team: employee.team ?? "",
      location: employee.location ?? "",
    });
    setModal(true);
  };

  const parseCsv = async (file: File) => {
    let matrix: string[][] = [];
    const isExcel = /\.(xlsx|xls)$/i.test(file.name);
    try {
      if (isExcel) {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        if (!sheetName) throw new Error("Excel workbook contains no sheets.");
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) throw new Error("Excel sheet could not be read.");
        matrix = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1, raw: false, defval: "" }).map((row) =>
          Array.isArray(row) ? row.map((cell) => String(cell ?? "").trim()) : []
        );
      } else {
        const text = await file.text();
        matrix = parseCsvText(text);
      }
    } catch (err) {
      notify({ kind: "error", text: `Failed to read file: ${message(err)}` });
      return;
    }

    if (matrix.length < 2) {
      notify({ kind: "error", text: "The file contains no readable data rows." });
      return;
    }
    const headers = (matrix[0] ?? []).map(normalizeHeader);
    const value = (row: string[], aliases: string[]) => {
      const index = headers.findIndex((header) => aliases.includes(header));
      return index >= 0 ? String(row[index] ?? "").trim() : "";
    };
    const rows: CsvEmployee[] = matrix
      .slice(1)
      .filter((row) => row.some((cell) => cell && String(cell).trim()))
      .map((row) => {
        const empId = value(row, ["employeeid", "empid", "userid", "productionid", "id", "staffid", "workerid", "employeecode", "empcode", "employee", "empno", "badgeno", "badgeid", "code", "idno", "empidno"]);
        const rawName = value(row, ["employeename", "name", "username", "fullname", "empname", "staffname", "workername", "employee", "personname"]);
        const name = rawName || empId;
        const prev = value(row, ["previousemployeeid", "previousids", "previousemployeeids", "oldid", "oldids", "previousid", "previouscode", "priorid", "prevempid", "oldempid", "previd"]);
        const prevIds = prev ? prev.split(/[;,|]/).map((item) => item.trim()).filter(Boolean) : [];
        const rawDate = value(row, ["joiningdate", "dateofjoining", "doj", "joindate", "startdate", "hiredate", "employmentdate", "datejoined", "joining"]);
        return {
          employeeId: empId,
          name: name,
          previousIds: prevIds,
          joiningDate: normalizeDate(rawDate) || "",
          status: normalizeEmployeeStatus(value(row, ["status", "employmentstatus", "employeestatus", "state"])),
          designation: value(row, ["designation", "jobtitle", "title", "role", "position", "post"]),
          contractType: value(row, ["contracttype", "employmenttype", "type", "jobtype", "contract", "employment"]),
          team: value(row, ["team", "department", "dept", "unit", "division", "group", "section"]),
          location: value(row, ["location", "worklocation", "country", "city", "branch", "site", "project"]),
        };
      })
      .filter((row) => row.employeeId);

    if (!rows.length) {
      notify({ kind: "error", text: "No valid rows found. Please ensure the file has an 'Employee ID' (or 'ID', 'Emp ID') column." });
      return;
    }
    setCsvRows(rows.slice(0, 10000));
    setCsvName(file.name);
  };

  const downloadBlankTemplate = () => {
    const csvContent = "Employee ID,Employee Name,Previous IDs,Joining Date,Status,Designation,Contract Type,Team,Location\nEMP-1001,John Doe,,2025-01-15,active,Data Specialist,Full-time,Operations,Dhaka\nEMP-1002,Jane Smith,EMP-9002,2025-02-01,active,Team Lead,Full-time,Operations,Chattogram\n";
    downloadText("employee_template.csv", csvContent);
  };

  if (!hasPermission(permissions, "employees.read")) return <Empty title="Employee access is restricted" body="Your role does not include employee records." />;

  const columns: Column<any>[] = [
    {
      key: "employeeId",
      header: "Employee ID",
      isImportant: true,
      render: (e) => <span className="mono" style={{ fontWeight: 700, color: "#073b33", border: "1px solid #99c2b6", padding: "2px 6px", borderRadius: 4, background: "#f0f8f5" }}>{e.employeeId}</span>,
    },
    {
      key: "name",
      header: "Employee Name",
      render: (e) => <div><strong>{e.name}</strong>{e.previousIds?.length ? <small style={{ display: "block", color: "var(--dim)", fontFamily: "monospace" }}>Prev: {e.previousIds.join(", ")}</small> : null}</div>,
    },
    {
      key: "designation",
      header: "Designation",
      render: (e) => e.designation || "—",
    },
    {
      key: "contractType",
      header: "Contract Type",
      render: (e) => e.contractType || "—",
    },
    {
      key: "team",
      header: "Team / Dept",
      render: (e) => e.team || "—",
    },
    {
      key: "location",
      header: "Location",
      render: (e) => e.location || "—",
    },
    {
      key: "joiningDate",
      header: "Joining Date",
      render: (e) => formatDate(e.joiningDate),
    },
    {
      key: "status",
      header: "Status",
      render: (e) => <Status value={e.status} />,
    },
  ];

  return (
    <>
      <div style={{ marginBottom: 12, display: "flex", justifyContent: "flex-end", gap: 10, alignItems: "center" }}>
        {canImport && (
          <>
            <button
              type="button"
              className="secondary-button compact"
              onClick={downloadBlankTemplate}
              title="Download employee template CSV"
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" x2="12" y1="15" y2="3"/></svg>
              <span>Download Template</span>
            </button>
            <label className="secondary-button compact" style={{ cursor: "pointer" }}>
              <Icon name="upload" /> Import CSV / Excel
              <input
                type="file"
                accept=".csv,text/csv,.xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
                style={{ display: "none" }}
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (file) void parseCsv(file);
                  e.currentTarget.value = "";
                }}
              />
            </label>
          </>
        )}
        {canManage && (
          <button
            className="primary-button compact"
            onClick={() => {
              setDraft(emptyEmployee);
              setModal(true);
            }}
          >
            <Icon name="plus" /> Add employee
          </button>
        )}
      </div>

      {csvRows.length > 0 && (
        <section className="import-preview" style={{ marginBottom: 16 }}>
          <div>
            <p className="overline">CSV preview</p>
            <h3>{csvName}</h3>
            <p>{csvRows.length} valid rows found. Matching uses Employee ID only; existing IDs will be updated.</p>
          </div>
          <div className="preview-sample">
            {csvRows.slice(0, 3).map((row) => (
              <span key={row.employeeId}>
                <b>{row.employeeId}</b>{row.name}
              </span>
            ))}
          </div>
          <div className="toolbar-actions">
            <button className="ghost-button" onClick={() => { setCsvRows([]); setCsvName(""); }}>
              Cancel
            </button>
            <button
              className="primary-button compact"
              disabled={importer.isPending}
              onClick={() => importer.mutate({ token, rows: csvRows })}
            >
              {importer.isPending ? "Importing…" : `Import ${csvRows.length} rows`}
            </button>
          </div>
        </section>
      )}

      {query.isPending ? (
        <Loading />
      ) : query.error ? (
        <Empty title="Employees could not load" body={message(query.error)} />
      ) : (
        <DataTable
          data={query.data?.employees ?? []}
          columns={columns}
          title="Employee Master Register"
          subtitle="Governed employee identities, designations, joining dates, and historical identifiers"
          searchPlaceholder="Search employees by ID, name, designation, or team..."
          searchFilter={(e, term) =>
            e.employeeId.toLowerCase().includes(term) ||
            e.name.toLowerCase().includes(term) ||
            (e.designation ?? "").toLowerCase().includes(term) ||
            (e.team ?? "").toLowerCase().includes(term) ||
            (e.location ?? "").toLowerCase().includes(term) ||
            (e.previousIds ?? []).some((id: string) => id.toLowerCase().includes(term))
          }
          onDownloadTemplate={downloadBlankTemplate}
          onView={(e) => setViewEmployee(e)}
          onEdit={canManage ? (e) => openEdit(e) : undefined}
          onDelete={canManage ? (e) => setDeleteTarget(e) : undefined}
        />
      )}

      {modal && (
        <Modal title={draft.id ? "Edit employee" : "Add employee"} onClose={() => setModal(false)} wide>
          <form
            className="form-grid"
            onSubmit={(event) => {
              event.preventDefault();
              const form = new FormData(event.currentTarget);
              save.mutate({
                token,
                id: draft.id,
                employee: {
                  employeeId: field(form, "employeeId"),
                  name: field(form, "name"),
                  previousIds: field(form, "previousIds").split(/[;,|]/).map((item) => item.trim()).filter(Boolean),
                  joiningDate: field(form, "joiningDate") || null,
                  status: field(form, "status") as EmployeeDraft["status"],
                  designation: field(form, "designation") || null,
                  contractType: field(form, "contractType") || null,
                  team: field(form, "team") || null,
                  location: field(form, "location") || null,
                },
              });
            }}
          >
            <label>
              Employee ID
              <input name="employeeId" defaultValue={draft.employeeId} required />
            </label>
            <label>
              Employee name
              <input name="name" defaultValue={draft.name} required />
            </label>
            <label className="full">
              Previous Employee IDs
              <input name="previousIds" defaultValue={draft.previousIds} />
            </label>
            <label>
              Joining date
              <input name="joiningDate" type="date" defaultValue={draft.joiningDate} />
            </label>
            <label>
              Status
              <select name="status" defaultValue={draft.status}>
                <option value="active">Active</option>
                <option value="inactive">Inactive</option>
                <option value="leave">On leave</option>
                <option value="separated">Separated</option>
              </select>
            </label>
            <label>
              Designation
              <input name="designation" defaultValue={draft.designation} />
            </label>
            <label>
              Contract type
              <input name="contractType" defaultValue={draft.contractType} />
            </label>
            <label>
              Team or department
              <input name="team" defaultValue={draft.team} />
            </label>
            <label>
              Work location
              <input name="location" defaultValue={draft.location} />
            </label>
            <div className="form-actions full">
              <button type="button" className="ghost-button" onClick={() => setModal(false)}>
                Cancel
              </button>
              <button className="primary-button compact" disabled={save.isPending}>
                <Icon name="check" /> {save.isPending ? "Saving…" : "Save employee"}
              </button>
            </div>
          </form>
        </Modal>
      )}

      {viewEmployee && (
        <Modal title={`Employee Details: ${viewEmployee.name}`} onClose={() => setViewEmployee(null)}>
          <div style={{ display: "grid", gap: 12 }}>
            <p><strong>Employee ID:</strong> <span className="mono">{viewEmployee.employeeId}</span></p>
            <p><strong>Full Name:</strong> {viewEmployee.name}</p>
            <p><strong>Status:</strong> <Status value={viewEmployee.status} /></p>
            <p><strong>Designation:</strong> {viewEmployee.designation || "—"}</p>
            <p><strong>Contract Type:</strong> {viewEmployee.contractType || "—"}</p>
            <p><strong>Team / Dept:</strong> {viewEmployee.team || "—"}</p>
            <p><strong>Work Location:</strong> {viewEmployee.location || "—"}</p>
            <p><strong>Joining Date:</strong> {formatDate(viewEmployee.joiningDate)}</p>
            <p><strong>Previous IDs:</strong> {viewEmployee.previousIds?.join(", ") || "—"}</p>
          </div>
          <div className="form-actions" style={{ marginTop: 16 }}>
            <button className="secondary-button" onClick={() => setViewEmployee(null)}>Close</button>
          </div>
        </Modal>
      )}

      <ConfirmDialog
        isOpen={Boolean(deleteTarget)}
        title="Move Employee to Recycle Bin?"
        message={`Are you sure you want to delete employee "${deleteTarget?.name}" (${deleteTarget?.employeeId})? The record will be moved to the Recycle Bin and can be restored at any time.`}
        confirmText="Move to Recycle Bin"
        danger
        onConfirm={() => {
          if (deleteTarget) removeMutation.mutate({ token, id: deleteTarget.id });
        }}
        onCancel={() => setDeleteTarget(null)}
      />
    </>
  );
}

const DEFAULT_COUNTRIES = ["BD", "Kolkata", "KTM", "IDR", "LKR"];
const DEFAULT_DOC_CATEGORIES = [
  "Header Character",
  "Date Character",
  "Birth Character",
  "Marriage Character",
  "Death Character",
  "Review Key",
  "Header",
];

function Projects({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const queryClient = useQueryClient();
  const [modal, setModal] = useState(false);
  const [viewProject, setViewProject] = useState<any | null>(null);
  const [deleteTarget, setDeleteTarget] = useState<any | null>(null);
  const [draft, setDraft] = useState<ProjectDraft>(emptyProject);
  const [customCountry, setCustomCountry] = useState("");
  const [customCharType, setCustomCharType] = useState("");
  const canManage = hasPermission(permissions, "projects.manage");

  const query = useQuery({ queryKey: ["projects", token], queryFn: () => api.listProjects({ token }), enabled: hasPermission(permissions, "projects.read") });
  const save = useMutation({
    mutationFn: api.saveProject,
    onSuccess: (data) => {
      queryClient.invalidateQueries({ queryKey: ["projects"] });
      queryClient.invalidateQueries({ queryKey: ["dashboard"] });
      setModal(false);
      setDraft(emptyProject);
      notify({ kind: "success", text: `Project ${data.mode} and persisted.` });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const removeMutation = useMutation({
    mutationFn: api.deleteProject,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["projects"] });
      queryClient.invalidateQueries({ queryKey: ["recycle-bin"] });
      queryClient.invalidateQueries({ queryKey: ["dashboard"] });
      setDeleteTarget(null);
      notify({ kind: "success", text: "Project moved to Recycle Bin." });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  if (!hasPermission(permissions, "projects.read")) return <Empty title="Project access is restricted" body="Your role does not include project records." />;

  const openEdit = (p: any) => {
    let subName = p.name || p.projectId;
    if (p.mainProjectId && subName.startsWith(p.mainProjectId + "-")) {
      subName = subName.slice(p.mainProjectId.length + 1);
    }
    setDraft({
      id: p.id,
      projectId: p.projectId,
      name: p.name,
      mainProjectId: p.mainProjectId || "hw",
      subProjectName: subName,
      client: p.client ?? "",
      description: p.description ?? "",
      startDate: p.startDate ?? "",
      endDate: p.endDate ?? "",
      status: p.status as ProjectDraft["status"],
      currency: p.currency || "BDT",
      manager: p.manager ?? "",
      notes: p.notes ?? "",
      language: p.language || "english",
      countries: Array.isArray(p.countries) && p.countries.length ? p.countries : ["BD"],
      averageWorkingHour: Boolean(p.averageWorkingHour),
      slabFormula: Boolean(p.averageWorkingHour && p.slabFormula),
      slabConfig: p.slabConfig ? {
        thresholdLow: String(p.slabConfig.thresholdLow ?? "3.99"),
        thresholdHigh: String(p.slabConfig.thresholdHigh ?? "5.99"),
        multiplierLow: String(p.slabConfig.multiplierLow ?? "0.60"),
        multiplierMid: String(p.slabConfig.multiplierMid ?? "0.80"),
        multiplierHigh: String(p.slabConfig.multiplierHigh ?? "1.00"),
        customTiers: Array.isArray(p.slabConfig.customTiers)
          ? p.slabConfig.customTiers.map((t: any, idx: number) => ({
              id: t.id || `tier-${idx + 1}`,
              label: t.label || `Tier ${idx + 1}`,
              minHours: t.minHours != null ? String(t.minHours) : "",
              maxHours: t.maxHours != null ? String(t.maxHours) : "",
              multiplier: String(t.multiplier ?? "1.00"),
            }))
          : [],
      } : {
        thresholdLow: "3.99",
        thresholdHigh: "5.99",
        multiplierLow: "0.60",
        multiplierMid: "0.80",
        multiplierHigh: "1.00",
        customTiers: [],
      },
      subProjectIncentives: Boolean(p.subProjectIncentives),
      docType: Boolean(p.docType),
      proof: Boolean(p.proof),
      paymentBasis: p.paymentBasis || (p.projectType as any) || "record",
      characterTypes: Array.isArray(p.characterTypes) && p.characterTypes.length ? p.characterTypes : ["Header Character", "Date Character"],
    });
    setModal(true);
  };

  const addCountryTag = (code: string) => {
    const trimmed = code.trim();
    if (!trimmed) return;
    if (!draft.countries.includes(trimmed)) {
      setDraft((prev) => ({ ...prev, countries: [...prev.countries, trimmed] }));
    }
    setCustomCountry("");
  };

  const removeCountryTag = (code: string) => {
    setDraft((prev) => ({
      ...prev,
      countries: prev.countries.filter((c) => c !== code),
    }));
  };

  const addCharTag = (name: string) => {
    const trimmed = name.trim();
    if (!trimmed) return;
    if (!draft.characterTypes.includes(trimmed)) {
      setDraft((prev) => ({ ...prev, characterTypes: [...prev.characterTypes, trimmed] }));
    }
    setCustomCharType("");
  };

  const removeCharTag = (name: string) => {
    setDraft((prev) => ({ ...prev, characterTypes: prev.characterTypes.filter((t) => t !== name) }));
  };

  const columns: Column<any>[] = [
    {
      key: "projectId",
      header: "Project ID",
      isImportant: true,
      render: (p) => (
        <span className="mono" style={{ fontWeight: 750, color: "#073b33", border: "1px solid #99c2b6", padding: "2px 7px", borderRadius: 4, background: "#f0f8f5", fontSize: "12px" }}>
          {p.projectId}
        </span>
      ),
    },
    {
      key: "mainProjectId",
      header: "Main Project",
      render: (p) => <span className="mono" style={{ fontWeight: 650 }}>{p.mainProjectId || "—"}</span>,
    },
    {
      key: "name",
      header: "Sub Project",
      render: (p) => <strong>{p.name}</strong>,
    },
    {
      key: "startDate",
      header: "Start Date",
      render: (p) => formatDate(p.startDate),
    },
    {
      key: "language",
      header: "Language",
      render: (p) => (
        <span style={{ fontSize: "11px", fontWeight: 700, padding: "2px 8px", borderRadius: 4, background: p.language === "nonenglish" ? "#fef3c7" : "#e0f2fe", color: p.language === "nonenglish" ? "#92400e" : "#0369a1", textTransform: "capitalize" }}>
          {p.language === "nonenglish" ? "Non-English" : "English"}
        </span>
      ),
    },
    {
      key: "countries",
      header: "Country",
      render: (p) => {
        const list: string[] = Array.isArray(p.countries) && p.countries.length ? p.countries : ["BD"];
        return (
          <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
            {list.map((c) => (
              <span key={c} className="mono" style={{ fontSize: "10.5px", fontWeight: 700, background: "#ecfdf5", color: "#065f46", border: "1px solid #a7f3d0", padding: "1px 5px", borderRadius: 3 }}>
                {c}
              </span>
            ))}
          </div>
        );
      },
    },
    {
      key: "paymentBasis",
      header: "Payment Basis",
      render: (p) => {
        const pb = p.paymentBasis || p.projectType || "record";
        return (
          <span style={{ display: "inline-flex", alignItems: "center", padding: "2px 8px", borderRadius: 4, fontSize: "11px", fontWeight: 700, background: pb === "character" ? "#f3e8ff" : pb === "batch" ? "#ffedd5" : "#ccfbf1", color: pb === "character" ? "#6b21a8" : pb === "batch" ? "#9a3412" : "#0f766e", textTransform: "capitalize" }}>
            {pb}
          </span>
        );
      },
    },
    {
      key: "features",
      header: "Config Setup",
      render: (p) => (
        <div style={{ display: "flex", gap: 4, flexWrap: "wrap", fontSize: "10px" }}>
          <span style={{ padding: "1px 5px", borderRadius: 3, background: p.docType ? "#ecfdf5" : "#f3f4f6", color: p.docType ? "#065f46" : "#6b7280", border: `1px solid ${p.docType ? "#a7f3d0" : "#e5e7eb"}` }}>
            DocType: {p.docType ? "Yes" : "No"}
          </span>
          <span style={{ padding: "1px 5px", borderRadius: 3, background: p.proof ? "#ecfdf5" : "#f3f4f6", color: p.proof ? "#065f46" : "#6b7280", border: `1px solid ${p.proof ? "#a7f3d0" : "#e5e7eb"}` }}>
            Proof: {p.proof ? "Yes" : "No"}
          </span>
          <span style={{ padding: "1px 5px", borderRadius: 3, background: p.averageWorkingHour ? "#ecfdf5" : "#f3f4f6", color: p.averageWorkingHour ? "#065f46" : "#6b7280", border: `1px solid ${p.averageWorkingHour ? "#a7f3d0" : "#e5e7eb"}` }}>
            AvgHr: {p.averageWorkingHour ? "Yes" : "No"}
          </span>
          <span style={{ padding: "1px 5px", borderRadius: 3, background: p.subProjectIncentives ? "#ecfdf5" : "#f3f4f6", color: p.subProjectIncentives ? "#065f46" : "#6b7280", border: `1px solid ${p.subProjectIncentives ? "#a7f3d0" : "#e5e7eb"}` }}>
            Inc: {p.subProjectIncentives ? "Yes" : "No"}
          </span>
        </div>
      ),
    },
    { key: "manager", header: "Project Manager", render: (p) => p.manager || "—" },
    {
      key: "status",
      header: "Status",
      render: (p) => <Status value={p.status} />,
    },
  ];

  return (
    <>
      <div style={{ marginBottom: 12, display: "flex", justifyContent: "flex-end" }}>
        {canManage && (
          <button
            className="primary-button compact"
            onClick={() => {
              setDraft(emptyProject);
              setModal(true);
            }}
          >
            <Icon name="plus" /> Create project
          </button>
        )}
      </div>

      {query.isPending ? (
        <Loading />
      ) : query.error ? (
        <Empty title="Projects could not load" body={message(query.error)} />
      ) : (
        <DataTable
          data={query.data?.projects ?? []}
          columns={columns}
          title="Projects"
          subtitle="Unified project setup and configuration"
          searchPlaceholder="Search projects..."
          searchFilter={(p, term) =>
            p.projectId.toLowerCase().includes(term) ||
            p.name.toLowerCase().includes(term) ||
            (p.mainProjectId ?? "").toLowerCase().includes(term) ||
            (p.manager ?? "").toLowerCase().includes(term) ||
            (p.paymentBasis ?? "").toLowerCase().includes(term)
          }
          onView={(p) => setViewProject(p)}
          onEdit={canManage ? (p) => openEdit(p) : undefined}
          onDelete={canManage ? (p) => setDeleteTarget(p) : undefined}
        />
      )}

      {modal && (
        <Modal title={draft.id ? "Edit Project" : "Create Project"} onClose={() => setModal(false)} wide>
          <form
            className="form-grid"
            onSubmit={(event) => {
              event.preventDefault();
              const sub = draft.subProjectName.trim();
              if (!sub) {
                notify({ kind: "error", text: "Sub Project Name is required." });
                return;
              }
              const main = draft.mainProjectId.trim().toLowerCase();
              const derivedProjectId = main ? `${main}-${sub}` : sub;
              save.mutate({
                token,
                project: {
                  id: draft.id,
                  projectId: derivedProjectId,
                  name: sub,
                  mainProjectId: main || null,
                  client: draft.client || null,
                  description: draft.description || null,
                  startDate: draft.startDate || null,
                  endDate: draft.endDate || null,
                  status: draft.status,
                  currency: draft.currency || "BDT",
                  manager: draft.manager.trim() || null,
                  notes: draft.notes || null,
                  projectType: draft.paymentBasis === "character" ? "character" : "record",
                  characterTypes: draft.characterTypes,
                  language: draft.language,
                  countries: draft.countries,
                  averageWorkingHour: draft.averageWorkingHour,
                  slabFormula: draft.averageWorkingHour && draft.slabFormula,
                  slabConfig: draft.averageWorkingHour && draft.slabFormula ? {
                    thresholdLow: draft.slabConfig.thresholdLow || "3.99",
                    thresholdHigh: draft.slabConfig.thresholdHigh || "5.99",
                    multiplierLow: draft.slabConfig.multiplierLow || "0.60",
                    multiplierMid: draft.slabConfig.multiplierMid || "0.80",
                    multiplierHigh: draft.slabConfig.multiplierHigh || "1.00",
                    customTiers: draft.slabConfig.customTiers.map((t) => ({
                      id: t.id,
                      label: t.label,
                      minHours: t.minHours ? t.minHours : null,
                      maxHours: t.maxHours ? t.maxHours : null,
                      multiplier: t.multiplier,
                    })),
                  } : null,
                  subProjectIncentives: draft.subProjectIncentives,
                  docType: draft.docType,
                  proof: draft.proof,
                  paymentBasis: draft.paymentBasis,
                },
              });
            }}
          >
            {/* First Row: 3 manual input cells */}
            <label>
              Main Project
              <input
                value={draft.mainProjectId}
                onChange={(e) => setDraft((prev) => ({ ...prev, mainProjectId: e.target.value }))}
                placeholder="hw"
                required
              />
            </label>
            <label>
              Sub Project
              <input
                value={draft.subProjectName}
                onChange={(e) => setDraft((prev) => ({ ...prev, subProjectName: e.target.value }))}
                placeholder="na0424"
                required
              />
            </label>
            <label>
              Project Manager
              <input
                value={draft.manager}
                onChange={(e) => setDraft((prev) => ({ ...prev, manager: e.target.value }))}
                placeholder="Manager name"
              />
            </label>

            {/* Second Row: Start Date, Language, Payment Basis */}
            <label>
              Start Date
              <input
                type="date"
                value={draft.startDate}
                onChange={(e) => setDraft((prev) => ({ ...prev, startDate: e.target.value }))}
                required
              />
            </label>
            <label>
              Language
              <select
                value={draft.language}
                onChange={(e) => setDraft((prev) => ({ ...prev, language: e.target.value as any }))}
              >
                <option value="english">English</option>
                <option value="nonenglish">Non-English</option>
              </select>
            </label>
            <label>
              Payment Basis
              <select
                value={draft.paymentBasis}
                onChange={(e) => setDraft((prev) => ({ ...prev, paymentBasis: e.target.value as any }))}
              >
                <option value="record">Record</option>
                <option value="character">Character</option>
                <option value="batch">Batch</option>
              </select>
            </label>

            {/* Third Row: Average Working Hour + Conditional Slab Formula Dropdown */}
            <label>
              Average Working Hour
              <select
                value={draft.averageWorkingHour ? "yes" : "no"}
                onChange={(e) => {
                  const val = e.target.value === "yes";
                  setDraft((prev) => ({
                    ...prev,
                    averageWorkingHour: val,
                    slabFormula: val ? prev.slabFormula : false,
                  }));
                }}
              >
                <option value="no">No</option>
                <option value="yes">Yes</option>
              </select>
            </label>

            {draft.averageWorkingHour && (
              <label>
                Slab Formula / Tier
                <select
                  value={draft.slabFormula ? "yes" : "no"}
                  onChange={(e) => setDraft((prev) => ({ ...prev, slabFormula: e.target.value === "yes" }))}
                >
                  <option value="no">No</option>
                  <option value="yes">Yes</option>
                </select>
              </label>
            )}

            <label>
              Sub Project Incentive
              <select
                value={draft.subProjectIncentives ? "yes" : "no"}
                onChange={(e) => setDraft((prev) => ({ ...prev, subProjectIncentives: e.target.value === "yes" }))}
              >
                <option value="no">No</option>
                <option value="yes">Yes</option>
              </select>
            </label>
            <label>
              Doc Type
              <select
                value={draft.docType ? "yes" : "no"}
                onChange={(e) => setDraft((prev) => ({ ...prev, docType: e.target.value === "yes" }))}
              >
                <option value="no">No</option>
                <option value="yes">Yes</option>
              </select>
            </label>
            <label>
              Proof
              <select
                value={draft.proof ? "yes" : "no"}
                onChange={(e) => setDraft((prev) => ({ ...prev, proof: e.target.value === "yes" }))}
              >
                <option value="no">No</option>
                <option value="yes">Yes</option>
              </select>
            </label>

            {/* Custom Slab Formula Configuration Box */}
            {draft.averageWorkingHour && draft.slabFormula && (
              <div className="full" style={{ background: "#f0fdf9", border: "1px solid #99c2b6", borderRadius: 8, padding: "16px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
                  <div>
                    <strong style={{ fontSize: "13.5px", color: "#073b33", display: "flex", alignItems: "center", gap: 6 }}>
                      <Icon name="calculator" /> Slab Formula & Tier Setup
                    </strong>
                    <span style={{ fontSize: "11.5px", color: "var(--dim)" }}>
                      Configure working hour thresholds, payout multipliers, and custom slab tiers for salary calculation.
                    </span>
                  </div>
                  <button
                    type="button"
                    className="secondary-button compact"
                    style={{ fontSize: "11.5px" }}
                    onClick={() => {
                      const newId = `tier-${Date.now()}`;
                      setDraft((prev) => ({
                        ...prev,
                        slabConfig: {
                          ...prev.slabConfig,
                          customTiers: [
                            ...prev.slabConfig.customTiers,
                            {
                              id: newId,
                              label: `Tier ${prev.slabConfig.customTiers.length + 1}`,
                              minHours: "",
                              maxHours: "",
                              multiplier: "1.00",
                            },
                          ],
                        },
                      }));
                    }}
                  >
                    <Icon name="plus" /> Add Custom Tier
                  </button>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: 10, marginBottom: 14 }}>
                  <label>
                    <span style={{ fontSize: "11px", fontWeight: 700, color: "#92400e" }}>Low Hours Threshold</span>
                    <input
                      value={draft.slabConfig.thresholdLow}
                      placeholder="3.99"
                      onChange={(e) => {
                        const val = e.target.value;
                        setDraft((prev) => ({ ...prev, slabConfig: { ...prev.slabConfig, thresholdLow: val } }));
                      }}
                    />
                  </label>
                  <label>
                    <span style={{ fontSize: "11px", fontWeight: 700, color: "#1e40af" }}>High Hours Threshold</span>
                    <input
                      value={draft.slabConfig.thresholdHigh}
                      placeholder="5.99"
                      onChange={(e) => {
                        const val = e.target.value;
                        setDraft((prev) => ({ ...prev, slabConfig: { ...prev.slabConfig, thresholdHigh: val } }));
                      }}
                    />
                  </label>
                  <label>
                    <span style={{ fontSize: "11px", fontWeight: 700, color: "#92400e" }}>Low Multiplier (&lt; Low)</span>
                    <input
                      value={draft.slabConfig.multiplierLow}
                      placeholder="0.60"
                      onChange={(e) => {
                        const val = e.target.value;
                        setDraft((prev) => ({ ...prev, slabConfig: { ...prev.slabConfig, multiplierLow: val } }));
                      }}
                    />
                  </label>
                  <label>
                    <span style={{ fontSize: "11px", fontWeight: 700, color: "#1e40af" }}>Mid Multiplier (Low - High)</span>
                    <input
                      value={draft.slabConfig.multiplierMid}
                      placeholder="0.80"
                      onChange={(e) => {
                        const val = e.target.value;
                        setDraft((prev) => ({ ...prev, slabConfig: { ...prev.slabConfig, multiplierMid: val } }));
                      }}
                    />
                  </label>
                  <label>
                    <span style={{ fontSize: "11px", fontWeight: 700, color: "#065f46" }}>High Multiplier (≥ High)</span>
                    <input
                      value={draft.slabConfig.multiplierHigh}
                      placeholder="1.00"
                      onChange={(e) => {
                        const val = e.target.value;
                        setDraft((prev) => ({ ...prev, slabConfig: { ...prev.slabConfig, multiplierHigh: val } }));
                      }}
                    />
                  </label>
                </div>

                {draft.slabConfig.customTiers.length > 0 && (
                  <div style={{ marginTop: 10, background: "#fff", border: "1px solid #d1fae5", borderRadius: 6, padding: "10px 12px" }}>
                    <span style={{ fontSize: "12px", fontWeight: 750, color: "#065f46", display: "block", marginBottom: 8 }}>
                      Custom Tier Overrides ({draft.slabConfig.customTiers.length})
                    </span>
                    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                      {draft.slabConfig.customTiers.map((t, idx) => (
                        <div key={t.id} style={{ display: "grid", gridTemplateColumns: "1.5fr 1fr 1fr 1fr auto", gap: 8, alignItems: "center" }}>
                          <input
                            placeholder="Tier Name (e.g. Standard)"
                            value={t.label}
                            style={{ fontSize: "12px", padding: "6px 8px" }}
                            onChange={(e) => {
                              const val = e.target.value;
                              setDraft((prev) => ({
                                ...prev,
                                slabConfig: {
                                  ...prev.slabConfig,
                                  customTiers: prev.slabConfig.customTiers.map((item, i) => i === idx ? { ...item, label: val } : item),
                                },
                              }));
                            }}
                          />
                          <input
                            placeholder="Min Hours (≥)"
                            type="number"
                            step="any"
                            value={t.minHours}
                            style={{ fontSize: "12px", padding: "6px 8px" }}
                            onChange={(e) => {
                              const val = e.target.value;
                              setDraft((prev) => ({
                                ...prev,
                                slabConfig: {
                                  ...prev.slabConfig,
                                  customTiers: prev.slabConfig.customTiers.map((item, i) => i === idx ? { ...item, minHours: val } : item),
                                },
                              }));
                            }}
                          />
                          <input
                            placeholder="Max Hours (≤)"
                            type="number"
                            step="any"
                            value={t.maxHours}
                            style={{ fontSize: "12px", padding: "6px 8px" }}
                            onChange={(e) => {
                              const val = e.target.value;
                              setDraft((prev) => ({
                                ...prev,
                                slabConfig: {
                                  ...prev.slabConfig,
                                  customTiers: prev.slabConfig.customTiers.map((item, i) => i === idx ? { ...item, maxHours: val } : item),
                                },
                              }));
                            }}
                          />
                          <input
                            placeholder="Multiplier"
                            type="number"
                            step="any"
                            value={t.multiplier}
                            style={{ fontSize: "12px", padding: "6px 8px", fontWeight: 700 }}
                            onChange={(e) => {
                              const val = e.target.value;
                              setDraft((prev) => ({
                                ...prev,
                                slabConfig: {
                                  ...prev.slabConfig,
                                  customTiers: prev.slabConfig.customTiers.map((item, i) => i === idx ? { ...item, multiplier: val } : item),
                                },
                              }));
                            }}
                          />
                          <button
                            type="button"
                            className="ghost-button"
                            style={{ color: "#dc2626", padding: "4px 8px" }}
                            onClick={() => {
                              setDraft((prev) => ({
                                ...prev,
                                slabConfig: {
                                  ...prev.slabConfig,
                                  customTiers: prev.slabConfig.customTiers.filter((_, i) => i !== idx),
                                },
                              }));
                            }}
                            title="Remove Tier"
                          >
                            ×
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Country Section with Dropdown + Manual Add */}
            <div className="full" style={{ background: "var(--surface-alt)", padding: "14px 16px", borderRadius: 8, border: "1px solid var(--border)" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, flexWrap: "wrap", gap: 6 }}>
                <span style={{ fontSize: "12.5px", fontWeight: 750, color: "#073b33" }}>
                  Country
                </span>
                <span style={{ fontSize: "11px", color: "var(--dim)" }}>
                  Select from dropdown or type custom country code
                </span>
              </div>
              <div style={{ display: "flex", gap: 8, marginBottom: 10, flexWrap: "wrap", alignItems: "center" }}>
                <select
                  style={{ minWidth: 180, padding: "7px 10px", fontSize: "13px" }}
                  defaultValue=""
                  onChange={(e) => {
                    if (e.target.value) {
                      addCountryTag(e.target.value);
                      e.target.value = "";
                    }
                  }}
                >
                  <option value="" disabled>Select Country from list…</option>
                  {DEFAULT_COUNTRIES.map((c) => (
                    <option key={c} value={c} disabled={draft.countries.includes(c)}>
                      {c} {draft.countries.includes(c) ? "(Added)" : ""}
                    </option>
                  ))}
                </select>

                <div style={{ display: "flex", gap: 6, flex: 1, minWidth: 200 }}>
                  <input
                    style={{ flex: 1, padding: "7px 10px", fontSize: "13px" }}
                    value={customCountry}
                    placeholder="Or type manual country code..."
                    onChange={(e) => setCustomCountry(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter") {
                        e.preventDefault();
                        addCountryTag(customCountry);
                      }
                    }}
                  />
                  <button
                    type="button"
                    className="secondary-button compact"
                    onClick={() => addCountryTag(customCountry)}
                  >
                    <Icon name="plus" /> Add
                  </button>
                </div>
              </div>

              <div style={{ display: "flex", flexWrap: "wrap", gap: 6, minHeight: 24 }}>
                {draft.countries.length === 0 ? (
                  <span style={{ fontSize: "11.5px", color: "var(--dim)" }}>No country added yet. Select from dropdown or enter manually.</span>
                ) : (
                  draft.countries.map((c) => (
                    <span
                      key={c}
                      style={{
                        display: "inline-flex",
                        alignItems: "center",
                        gap: 6,
                        background: "#e8f3ef",
                        color: "#073b33",
                        border: "1px solid #99c2b6",
                        padding: "3px 9px",
                        borderRadius: 4,
                        fontSize: "12px",
                        fontWeight: 700,
                        fontFamily: "monospace",
                      }}
                    >
                      {c}
                      <button
                        type="button"
                        onClick={() => removeCountryTag(c)}
                        style={{ background: "none", border: "none", cursor: "pointer", color: "#c0392b", fontWeight: 800, padding: 0 }}
                        title={`Remove ${c}`}
                      >
                        ×
                      </button>
                    </span>
                  ))
                )}
              </div>
            </div>

            {/* Categories Doc Type Section with Dropdown + Manual Add */}
            <div className="full" style={{ background: "var(--surface-alt)", padding: "14px 16px", borderRadius: 8, border: "1px solid var(--border)" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, flexWrap: "wrap", gap: 6 }}>
                <span style={{ fontSize: "12.5px", fontWeight: 750, color: "#073b33" }}>
                  Categories Doc Type
                </span>
                <span style={{ fontSize: "11px", color: "var(--dim)" }}>
                  Select from dropdown or type custom category
                </span>
              </div>
              <div style={{ display: "flex", gap: 8, marginBottom: 10, flexWrap: "wrap", alignItems: "center" }}>
                <select
                  style={{ minWidth: 200, padding: "7px 10px", fontSize: "13px" }}
                  defaultValue=""
                  onChange={(e) => {
                    if (e.target.value) {
                      addCharTag(e.target.value);
                      e.target.value = "";
                    }
                  }}
                >
                  <option value="" disabled>Select Category from list…</option>
                  {DEFAULT_DOC_CATEGORIES.map((cat) => (
                    <option key={cat} value={cat} disabled={draft.characterTypes.includes(cat)}>
                      {cat} {draft.characterTypes.includes(cat) ? "(Added)" : ""}
                    </option>
                  ))}
                </select>

                <div style={{ display: "flex", gap: 6, flex: 1, minWidth: 220 }}>
                  <input
                    style={{ flex: 1, padding: "7px 10px", fontSize: "13px" }}
                    value={customCharType}
                    placeholder="Or type custom category name..."
                    onChange={(e) => setCustomCharType(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter") {
                        e.preventDefault();
                        addCharTag(customCharType);
                      }
                    }}
                  />
                  <button
                    type="button"
                    className="secondary-button compact"
                    onClick={() => addCharTag(customCharType)}
                  >
                    <Icon name="plus" /> Add
                  </button>
                </div>
              </div>

              <div style={{ display: "flex", flexWrap: "wrap", gap: 6, minHeight: 24 }}>
                {draft.characterTypes.length === 0 ? (
                  <span style={{ fontSize: "11.5px", color: "var(--dim)" }}>No categories added. Select from dropdown or enter manually.</span>
                ) : (
                  draft.characterTypes.map((t) => (
                    <span
                      key={t}
                      style={{
                        display: "inline-flex",
                        alignItems: "center",
                        gap: 6,
                        background: "#e8f3ef",
                        color: "#073b33",
                        border: "1px solid #99c2b6",
                        padding: "3px 9px",
                        borderRadius: 4,
                        fontSize: "12px",
                        fontWeight: 650,
                      }}
                    >
                      {t}
                      <button
                        type="button"
                        onClick={() => removeCharTag(t)}
                        style={{ background: "none", border: "none", cursor: "pointer", color: "#c0392b", fontWeight: 800, padding: 0 }}
                        title={`Remove ${t}`}
                      >
                        ×
                      </button>
                    </span>
                  ))
                )}
              </div>
            </div>

            <div className="form-actions full">
              <button type="button" className="ghost-button" onClick={() => setModal(false)}>
                Cancel
              </button>
              <button className="primary-button compact" disabled={save.isPending}>
                <Icon name="check" /> {save.isPending ? "Saving…" : draft.id ? "Save project" : "Create project"}
              </button>
            </div>
          </form>
        </Modal>
      )}

      {viewProject && (
        <Modal title={`Project ${viewProject.projectId} Details`} onClose={() => setViewProject(null)} wide>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
            <div>
              <p><strong>Main Project:</strong> <span className="mono">{viewProject.mainProjectId || "—"}</span></p>
              <p><strong>Sub Project Name:</strong> {viewProject.name}</p>
              <p><strong>Project ID:</strong> <span className="mono">{viewProject.projectId}</span></p>
              <p><strong>Start Date:</strong> {formatDate(viewProject.startDate)}</p>
              <p><strong>Language:</strong> {viewProject.language === "nonenglish" ? "Non-English" : "English"}</p>
              <p><strong>Payment Basis:</strong> {viewProject.paymentBasis || viewProject.projectType || "record"}</p>
            </div>
            <div>
              <p><strong>Project Manager:</strong> {viewProject.manager || "—"}</p>
              <p><strong>Status:</strong> <Status value={viewProject.status} /></p>
              <p><strong>DocType:</strong> {viewProject.docType ? "Enabled" : "Disabled"}</p>
              <p><strong>Proof:</strong> {viewProject.proof ? "Enabled" : "Disabled"}</p>
              <p><strong>Average Working Hour:</strong> {viewProject.averageWorkingHour ? "Enabled" : "Disabled"}</p>
              <p><strong>Sub Project Incentives:</strong> {viewProject.subProjectIncentives ? "Enabled" : "Disabled"}</p>
            </div>
            <div className="full" style={{ gridColumn: "span 2" }}>
              <p><strong>Country Scopes:</strong></p>
              <div style={{ display: "flex", gap: 6, marginTop: 4 }}>
                {(viewProject.countries || ["BD"]).map((c: string) => (
                  <span key={c} className="mono" style={{ background: "#ecfdf5", color: "#065f46", border: "1px solid #a7f3d0", padding: "2px 8px", borderRadius: 4, fontSize: "12px", fontWeight: 700 }}>
                    {c}
                  </span>
                ))}
              </div>
              <p style={{ marginTop: 12 }}><strong>Record / Character Categories:</strong></p>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 4 }}>
                {Array.isArray(viewProject.characterTypes) && viewProject.characterTypes.length ? (
                  viewProject.characterTypes.map((t: string) => (
                    <span key={t} style={{ background: "#ecfdf5", color: "#065f46", border: "1px solid #a7f3d0", padding: "2px 8px", borderRadius: 4, fontSize: "12px" }}>
                      {t}
                    </span>
                  ))
                ) : (
                  <span>None</span>
                )}
              </div>
            </div>
          </div>
          <div className="form-actions" style={{ marginTop: 16 }}>
            <button className="secondary-button" onClick={() => setViewProject(null)}>Close</button>
          </div>
        </Modal>
      )}

      <ConfirmDialog
        isOpen={Boolean(deleteTarget)}
        title="Move Project to Recycle Bin?"
        message={`Are you sure you want to delete project "${deleteTarget?.projectId} - ${deleteTarget?.name}"? It will be safely moved to the Recycle Bin and can be restored whenever needed.`}
        confirmText="Move to Recycle Bin"
        danger
        onConfirm={() => {
          if (deleteTarget) removeMutation.mutate({ token, id: deleteTarget.id });
        }}
        onCancel={() => setDeleteTarget(null)}
      />
    </>
  );
}


type PeriodRow = NonNullable<Awaited<ReturnType<typeof api.listPeriods>>>["periods"][number];

function getPeriodDefaults(month: number, year: number) {
  const mm = String(month).padStart(2, "0");
  const startDate = `${year}-${mm}-01`;
  const lastDay = new Date(year, month, 0).getDate();
  const endDate = `${year}-${mm}-${String(lastDay).padStart(2, "0")}`;
  const suggestedId = `PR-${year}-${mm}`;
  return { startDate, endDate, suggestedId };
}

type AugmentedPeriod = PeriodRow & { monthName: string; projectNames: string };

function Periods({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const queryClient = useQueryClient();
  const allowed = hasPermission(permissions, "periods.read");
  const canManage = hasPermission(permissions, "periods.manage");

  const [modal, setModal] = useState(false);
  const [editPeriod, setEditPeriod] = useState<AugmentedPeriod | null>(null);
  const [viewPeriod, setViewPeriod] = useState<AugmentedPeriod | null>(null);
  const [deleteTarget, setDeleteTarget] = useState<AugmentedPeriod | null>(null);
  const [statusPeriod, setStatusPeriod] = useState<number | null>(null);
  const [statusFilter, setStatusFilter] = useState("all");
  const [yearFilter, setYearFilter] = useState("all");

  // Create form state
  const [createMonth, setCreateMonth] = useState(new Date().getMonth() + 1);
  const [createYear, setCreateYear] = useState(new Date().getFullYear());
  const [createPeriodId, setCreatePeriodId] = useState("");
  const [createStartDate, setCreateStartDate] = useState("");
  const [createEndDate, setCreateEndDate] = useState("");
  const [createCopyFromId, setCreateCopyFromId] = useState<number | null>(null);
  const [createSelectedProjects, setCreateSelectedProjects] = useState<number[]>([]);
  const [createProjectSearch, setCreateProjectSearch] = useState("");

  // Edit form state
  const [editSelectedProjects, setEditSelectedProjects] = useState<number[]>([]);
  const [editProjectSearch, setEditProjectSearch] = useState("");

  const periods = useQuery({
    queryKey: ["periods", token],
    queryFn: () => api.listPeriods({ token }),
    enabled: allowed,
  });

  const projects = useQuery({
    queryKey: ["projects", token],
    queryFn: () => api.listProjects({ token }),
    enabled: hasPermission(permissions, "projects.read"),
  });

  const create = useMutation({
    mutationFn: api.createPeriod,
    onSuccess: (data) => {
      queryClient.invalidateQueries({ queryKey: ["periods"] });
      queryClient.invalidateQueries({ queryKey: ["dashboard"] });
      setModal(false);
      notify({ kind: "success", text: `Payroll period created successfully with ${data.copiedProjects} active project configuration(s).` });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const update = useMutation({
    mutationFn: api.updatePeriod,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["periods"] });
      queryClient.invalidateQueries({ queryKey: ["dashboard"] });
      setEditPeriod(null);
      notify({ kind: "success", text: "Payroll period updated successfully." });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const statusMutation = useMutation({
    mutationFn: api.updatePeriodStatus,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["periods"] });
      queryClient.invalidateQueries({ queryKey: ["dashboard"] });
      setStatusPeriod(null);
      notify({ kind: "success", text: "Payroll period lifecycle status advanced successfully." });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const removeMutation = useMutation({
    mutationFn: api.deletePeriod,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["periods"] });
      queryClient.invalidateQueries({ queryKey: ["recycle-bin"] });
      queryClient.invalidateQueries({ queryKey: ["dashboard"] });
      setDeleteTarget(null);
      notify({ kind: "success", text: "Payroll period moved to Recycle Bin. It can be restored at any time." });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  if (!allowed) {
    return <Empty title="Payroll period access is restricted" body="Your role does not include payroll periods authority." />;
  }

  const transitions: Record<string, string[]> = {
    Draft: ["Importing"],
    Importing: ["Reconciliation Pending"],
    "Reconciliation Pending": ["Ready for Calculation"],
    "Ready for Calculation": ["Calculated"],
    Calculated: ["Under Review"],
    "Under Review": ["Approved", "Calculated"],
    Approved: ["Locked"],
    Locked: ["Exported"],
    Exported: [],
  };

  const projectMap = new Map((projects.data?.projects ?? []).map((p) => [p.id, p]));

  // Augment periods with project names and formatted month for search & display
  const augmentedPeriods: AugmentedPeriod[] = (periods.data?.periods ?? []).map((p) => {
    const monthName = new Intl.DateTimeFormat("en-US", { month: "long" }).format(new Date(p.year, p.month - 1, 1));
    const projectNames = p.activeProjectIds
      .map((id) => {
        const proj = projectMap.get(id);
        return proj ? `${proj.projectId} (${proj.name})` : `#${id}`;
      })
      .join(", ");
    return {
      ...p,
      monthName,
      projectNames,
    };
  });

  // Apply filters
  const filteredPeriods = augmentedPeriods.filter((p) => {
    if (statusFilter !== "all" && p.status !== statusFilter) return false;
    if (yearFilter !== "all" && String(p.year) !== yearFilter) return false;
    return true;
  });

  // Unique years for filter
  const availableYears = Array.from(new Set(augmentedPeriods.map((p) => String(p.year)))).sort((a, b) => b.localeCompare(a));

  // Initialize create modal with defaults
  const openCreateModal = () => {
    const currentMonth = new Date().getMonth() + 1;
    const currentYear = new Date().getFullYear();
    const defaults = getPeriodDefaults(currentMonth, currentYear);
    setCreateMonth(currentMonth);
    setCreateYear(currentYear);
    setCreatePeriodId(defaults.suggestedId);
    setCreateStartDate(defaults.startDate);
    setCreateEndDate(defaults.endDate);
    setCreateCopyFromId(null);
    setCreateSelectedProjects(
      projects.data?.projects.filter((p) => p.status === "active").map((p) => p.id) ?? []
    );
    setCreateProjectSearch("");
    setModal(true);
  };

  const handleCreateMonthYearChange = (newMonth: number, newYear: number) => {
    setCreateMonth(newMonth);
    setCreateYear(newYear);
    const def = getPeriodDefaults(newMonth, newYear);
    setCreateStartDate(def.startDate);
    setCreateEndDate(def.endDate);
    if (!createPeriodId || /^PR-\d{4}-\d{2}$/i.test(createPeriodId) || /^PAY-\d{4}-\d{2}$/i.test(createPeriodId)) {
      setCreatePeriodId(def.suggestedId);
    }
  };

  const handleCopyFromChange = (copyId: number | null) => {
    setCreateCopyFromId(copyId);
    if (copyId) {
      const source = periods.data?.periods.find((p) => p.id === copyId);
      if (source && source.activeProjectIds.length > 0) {
        setCreateSelectedProjects(source.activeProjectIds);
      }
    }
  };

  const openEdit = (p: AugmentedPeriod) => {
    setEditPeriod(p);
    setEditSelectedProjects(p.activeProjectIds);
    setEditProjectSearch("");
  };

  const columns: Column<AugmentedPeriod>[] = [
    {
      key: "periodId",
      header: "Period ID",
      isImportant: true,
      render: (p) => (
        <span
          className="mono"
          style={{
            fontWeight: 750,
            color: "#073b33",
            border: "1px solid #99c2b6",
            padding: "3px 8px",
            borderRadius: 5,
            background: "#f0f8f5",
            letterSpacing: "0.02em",
            display: "inline-block",
          }}
        >
          {p.periodId}
        </span>
      ),
    },
    {
      key: "monthName",
      header: "Payroll Month & Year",
      render: (p) => (
        <div>
          <strong style={{ fontSize: "13.5px" }}>{p.monthName} {p.year}</strong>
          <small style={{ display: "block", color: "var(--dim)", fontSize: "11px", marginTop: "1px" }}>
            Month {p.month} · Calendar Year {p.year}
          </small>
        </div>
      ),
    },
    {
      key: "startDate",
      header: "Cycle Dates",
      render: (p) => {
        const start = new Date(p.startDate);
        const end = new Date(p.endDate);
        const days = Math.round((end.getTime() - start.getTime()) / 86400000) + 1;
        return (
          <div>
            <span style={{ fontWeight: 600 }}>{formatDate(p.startDate)} – {formatDate(p.endDate)}</span>
            <small style={{ display: "block", color: "var(--dim)", fontSize: "11px", marginTop: "1px" }}>
              {days > 0 ? `${days} calendar days` : "Active window"}
            </small>
          </div>
        );
      },
    },
    {
      key: "status",
      header: "Lifecycle Status & Progress",
      render: (p) => <PeriodLifecycleProgress status={p.status} />,
    },
    {
      key: "calculationVersion",
      header: "Calc Version",
      render: (p) =>
        p.calculationVersion > 0 ? (
          <span
            className="mono"
            style={{
              fontWeight: 750,
              color: "var(--accent-deep)",
              background: "var(--accent-pale)",
              border: "1px solid color-mix(in srgb, var(--accent) 30%, transparent)",
              padding: "2px 7px",
              borderRadius: 4,
              fontSize: "12px",
            }}
          >
            v{p.calculationVersion}
          </span>
        ) : (
          <span style={{ color: "var(--dim)", fontSize: "12px" }}>Uncalculated</span>
        ),
    },
    {
      key: "activeProjectIds",
      header: "Active Projects",
      render: (p) => {
        const count = p.activeProjectIds.length;
        return (
          <button
            type="button"
            className="text-button"
            style={{
              padding: "2px 8px",
              fontSize: "12px",
              fontWeight: 650,
              background: "var(--surface-alt)",
              border: "1px solid var(--border)",
              borderRadius: 5,
              cursor: "pointer",
            }}
            title={p.projectNames || "No active projects"}
            onClick={(e) => {
              e.stopPropagation();
              setViewPeriod(p);
            }}
          >
            📁 {count} Project{count === 1 ? "" : "s"}
          </button>
        );
      },
    },
    {
      key: "createdAt",
      header: "Created",
      render: (p) => formatDate(p.createdAt),
    },
  ];

  return (
    <>
      {/* Action Toolbar */}
      {canManage && (
        <div style={{ marginBottom: 14, display: "flex", justifyContent: "flex-end" }}>
          <button
            type="button"
            className="primary-button compact"
            onClick={openCreateModal}
            title="Create a new payroll period"
          >
            <Icon name="plus" /> Create Period
          </button>
        </div>
      )}

      {/* Main Table View */}
      {periods.isPending ? (
        <Loading />
      ) : periods.error ? (
        <Empty title="Payroll periods could not load" body={message(periods.error)} />
      ) : !periods.data?.periods.length ? (
        <Empty
          title="No payroll periods yet"
          body="Create a period and choose its active projects. Production quantities and salary results are never copied."
          action={canManage ? (
            <button className="primary-button compact" onClick={openCreateModal}>
              <Icon name="plus" /> Create the first period
            </button>
          ) : undefined}
        />
      ) : (
        <DataTable
          data={filteredPeriods}
          columns={columns}
          title="Payroll Periods Register"
          subtitle="Authoritative calendar cycles, lifecycle progression, and active project scopes"
          searchPlaceholder="Search periods by ID, month, year, status, or project..."
          searchFilter={(p, term) =>
            p.periodId.toLowerCase().includes(term) ||
            p.status.toLowerCase().includes(term) ||
            String(p.year).includes(term) ||
            String(p.month).includes(term) ||
            p.monthName.toLowerCase().includes(term) ||
            p.projectNames.toLowerCase().includes(term)
          }
          extraFilters={
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
              <label className="select-field" style={{ margin: 0 }}>
                <span style={{ fontSize: "11.5px", fontWeight: 650, color: "var(--dim)", marginRight: 6 }}>Status:</span>
                <select
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value)}
                  style={{ fontSize: "12px", padding: "4px 8px", height: "auto" }}
                >
                  <option value="all">All Statuses ({augmentedPeriods.length})</option>
                  <option value="Draft">Draft</option>
                  <option value="Importing">Importing</option>
                  <option value="Reconciliation Pending">Reconciliation Pending</option>
                  <option value="Ready for Calculation">Ready for Calculation</option>
                  <option value="Calculated">Calculated</option>
                  <option value="Under Review">Under Review</option>
                  <option value="Approved">Approved</option>
                  <option value="Locked">Locked</option>
                  <option value="Exported">Exported</option>
                </select>
              </label>

              {availableYears.length > 1 && (
                <label className="select-field" style={{ margin: 0 }}>
                  <span style={{ fontSize: "11.5px", fontWeight: 650, color: "var(--dim)", marginRight: 6 }}>Year:</span>
                  <select
                    value={yearFilter}
                    onChange={(e) => setYearFilter(e.target.value)}
                    style={{ fontSize: "12px", padding: "4px 8px", height: "auto" }}
                  >
                    <option value="all">All Years</option>
                    {availableYears.map((yr) => (
                      <option key={yr} value={yr}>{yr}</option>
                    ))}
                  </select>
                </label>
              )}
            </div>
          }
          onView={(p) => setViewPeriod(p)}
          onEdit={canManage ? (p) => openEdit(p) : undefined}
          onDelete={
            canManage
              ? (p) => {
                  if (["Approved", "Locked", "Exported"].includes(p.status)) {
                    notify({ kind: "error", text: `Cannot delete period in ${p.status} status. Approved and locked periods protect audit integrity.` });
                    return;
                  }
                  setDeleteTarget(p);
                }
              : undefined
          }
          customActions={(p) =>
            canManage && (transitions[p.status] ?? []).length > 0 ? (
              <button
                type="button"
                className="secondary-button compact"
                onClick={() => setStatusPeriod(p.id)}
                title={`Advance status from ${p.status} to next lifecycle phase`}
                style={{ fontSize: "11.5px", padding: "3px 8px", whiteSpace: "nowrap" }}
              >
                Advance
              </button>
            ) : null
          }
        />
      )}

      {/* CREATE PERIOD MODAL */}
      {modal && (
        <Modal title="Create payroll period" onClose={() => setModal(false)} wide>
          <form
            className="form-grid"
            onSubmit={(event) => {
              event.preventDefault();
              if (createEndDate < createStartDate) {
                notify({ kind: "error", text: "End date must be on or after start date." });
                return;
              }
              create.mutate({
                token,
                periodId: createPeriodId.trim().toUpperCase(),
                month: Number(createMonth),
                year: Number(createYear),
                startDate: createStartDate,
                endDate: createEndDate,
                copyFromPeriodId: createCopyFromId,
                activeProjectIds: createSelectedProjects,
              });
            }}
          >
            <label>
              Month
              <select
                value={createMonth}
                onChange={(e) => handleCreateMonthYearChange(Number(e.target.value), createYear)}
                required
              >
                {Array.from({ length: 12 }, (_, index) => (
                  <option key={index + 1} value={index + 1}>
                    {new Intl.DateTimeFormat("en-US", { month: "long" }).format(new Date(2026, index, 1))} (Month {index + 1})
                  </option>
                ))}
              </select>
            </label>
            <label>
              Year
              <input
                type="number"
                min="2020"
                max="2100"
                value={createYear}
                onChange={(e) => handleCreateMonthYearChange(createMonth, Number(e.target.value))}
                required
              />
            </label>
            <label>
              Start Date
              <input
                type="date"
                value={createStartDate}
                onChange={(e) => setCreateStartDate(e.target.value)}
                required
              />
            </label>
            <label>
              End Date
              <input
                type="date"
                value={createEndDate}
                onChange={(e) => setCreateEndDate(e.target.value)}
                required
              />
            </label>
            <label className="full">
              Copy active project setup from
              <select
                value={createCopyFromId ?? ""}
                onChange={(e) => handleCopyFromChange(e.target.value ? Number(e.target.value) : null)}
              >
                <option value="">Do not copy (Use default active projects)</option>
                {periods.data?.periods.map((p) => (
                  <option value={p.id} key={p.id}>
                    {p.periodId} ({p.status}) · {p.activeProjectIds.length} projects
                  </option>
                ))}
              </select>
            </label>

            <details className="full" style={{ fontSize: "12px", border: "1px dashed var(--border)", borderRadius: 6, padding: "8px 12px" }}>
              <summary style={{ cursor: "pointer", fontWeight: 650, color: "var(--dim)" }}>
                Advanced: Custom Project Scope ({createSelectedProjects.length} of {projects.data?.projects.length || 0} active)
              </summary>
              <div style={{ marginTop: 8, display: "flex", gap: 8, alignItems: "center", marginBottom: 6, flexWrap: "wrap" }}>
                <button
                  type="button"
                  className="secondary-button compact"
                  style={{ fontSize: "11px", padding: "2px 8px" }}
                  onClick={() => setCreateSelectedProjects(projects.data?.projects.map((p) => p.id) ?? [])}
                >
                  Select All
                </button>
                <button
                  type="button"
                  className="secondary-button compact"
                  style={{ fontSize: "11px", padding: "2px 8px" }}
                  onClick={() => setCreateSelectedProjects([])}
                >
                  Deselect All
                </button>
                <input
                  type="search"
                  placeholder="Filter project list..."
                  value={createProjectSearch}
                  onChange={(e) => setCreateProjectSearch(e.target.value)}
                  style={{ flex: 1, minWidth: "150px", fontSize: "11.5px", padding: "3px 8px" }}
                />
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 6, maxHeight: "160px", overflowY: "auto", marginTop: 6 }}>
                {projects.data?.projects
                  .filter((p) =>
                    !createProjectSearch ||
                    p.projectId.toLowerCase().includes(createProjectSearch.toLowerCase()) ||
                    p.name.toLowerCase().includes(createProjectSearch.toLowerCase())
                  )
                  .map((p) => (
                    <label key={p.id} style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", fontSize: "12px" }}>
                      <input
                        type="checkbox"
                        checked={createSelectedProjects.includes(p.id)}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setCreateSelectedProjects((prev) => [...prev, p.id]);
                          } else {
                            setCreateSelectedProjects((prev) => prev.filter((id) => id !== p.id));
                          }
                        }}
                      />
                      <span>
                        <b className="mono">{p.projectId}</b> {p.name}
                      </span>
                    </label>
                  ))}
              </div>
            </details>

            <div className="form-actions full">
              <button type="button" className="ghost-button" onClick={() => setModal(false)}>
                Cancel
              </button>
              <button className="primary-button compact" disabled={create.isPending}>
                {create.isPending ? "Creating…" : "Create Period"}
              </button>
            </div>
          </form>
        </Modal>
      )}

      {/* EDIT PERIOD MODAL */}
      {editPeriod && (
        <Modal title={`Edit Payroll Period: ${editPeriod.periodId}`} onClose={() => setEditPeriod(null)} wide>
          <form
            className="form-grid"
            onSubmit={(event) => {
              event.preventDefault();
              const form = new FormData(event.currentTarget);
              const start = field(form, "startDate");
              const end = field(form, "endDate");
              if (end < start) {
                notify({ kind: "error", text: "End date must be on or after start date." });
                return;
              }
              update.mutate({
                token,
                id: editPeriod.id,
                periodId: field(form, "periodId"),
                month: Number(field(form, "month")),
                year: Number(field(form, "year")),
                startDate: start,
                endDate: end,
                status: (field(form, "status") || undefined) as any,
                activeProjectIds: editSelectedProjects,
              });
            }}
          >
            <label>
              Unique period ID
              <input name="periodId" defaultValue={editPeriod.periodId} required />
            </label>
            <label>
              Month
              <select name="month" defaultValue={editPeriod.month} required>
                {Array.from({ length: 12 }, (_, index) => (
                  <option key={index + 1} value={index + 1}>
                    {new Intl.DateTimeFormat("en-US", { month: "long" }).format(new Date(2026, index, 1))} (Month {index + 1})
                  </option>
                ))}
              </select>
            </label>
            <label>
              Year
              <input name="year" type="number" min="2020" max="2100" defaultValue={editPeriod.year} required />
            </label>
            <label>
              Current Status
              <select
                name="status"
                defaultValue={editPeriod.status}
                disabled={["Locked", "Exported"].includes(editPeriod.status)}
              >
                <option value="Draft">Draft</option>
                <option value="Importing">Importing</option>
                <option value="Reconciliation Pending">Reconciliation Pending</option>
                <option value="Ready for Calculation">Ready for Calculation</option>
                <option value="Calculated">Calculated</option>
                <option value="Under Review">Under Review</option>
                <option value="Approved">Approved</option>
                <option value="Locked">Locked</option>
                <option value="Exported">Exported</option>
              </select>
            </label>
            <label>
              Start date
              <input name="startDate" type="date" defaultValue={editPeriod.startDate} required />
            </label>
            <label>
              End date
              <input name="endDate" type="date" defaultValue={editPeriod.endDate} required />
            </label>

            <fieldset className="full checklist">
              <legend>
                Active projects for this period ({editSelectedProjects.length} selected of {projects.data?.projects.length || 0})
              </legend>
              <div style={{ gridColumn: "1 / -1", display: "flex", gap: 8, alignItems: "center", marginBottom: 6, flexWrap: "wrap" }}>
                <button
                  type="button"
                  className="secondary-button compact"
                  style={{ fontSize: "11px", padding: "2px 8px" }}
                  disabled={["Approved", "Locked", "Exported"].includes(editPeriod.status)}
                  onClick={() => setEditSelectedProjects(projects.data?.projects.map((p) => p.id) ?? [])}
                >
                  Select All
                </button>
                <button
                  type="button"
                  className="secondary-button compact"
                  style={{ fontSize: "11px", padding: "2px 8px" }}
                  disabled={["Approved", "Locked", "Exported"].includes(editPeriod.status)}
                  onClick={() => setEditSelectedProjects([])}
                >
                  Deselect All
                </button>
                <input
                  type="search"
                  placeholder="Filter project list..."
                  value={editProjectSearch}
                  onChange={(e) => setEditProjectSearch(e.target.value)}
                  style={{ flex: 1, minWidth: "150px", fontSize: "11.5px", padding: "3px 8px" }}
                />
              </div>
              {projects.data?.projects.length ? (
                projects.data.projects
                  .filter((p) =>
                    !editProjectSearch ||
                    p.projectId.toLowerCase().includes(editProjectSearch.toLowerCase()) ||
                    p.name.toLowerCase().includes(editProjectSearch.toLowerCase())
                  )
                  .map((p) => (
                    <label key={p.id} style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }}>
                      <input
                        type="checkbox"
                        checked={editSelectedProjects.includes(p.id)}
                        disabled={["Approved", "Locked", "Exported"].includes(editPeriod.status)}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setEditSelectedProjects((prev) => [...prev, p.id]);
                          } else {
                            setEditSelectedProjects((prev) => prev.filter((id) => id !== p.id));
                          }
                        }}
                      />
                      <span>
                        <b className="mono">{p.projectId}</b>
                        {p.name} {p.status !== "active" && `(${p.status})`}
                      </span>
                    </label>
                  ))
              ) : (
                <p>No project definitions available.</p>
              )}
            </fieldset>

            {["Locked", "Exported"].includes(editPeriod.status) && (
              <p className="field-help full" style={{ color: "var(--danger)" }}>
                Note: This period is {editPeriod.status}. Core configuration is locked to protect historical and financial integrity.
              </p>
            )}

            <div className="form-actions full">
              <button type="button" className="ghost-button" onClick={() => setEditPeriod(null)}>
                Cancel
              </button>
              <button
                className="primary-button compact"
                disabled={update.isPending || ["Locked", "Exported"].includes(editPeriod.status)}
              >
                {update.isPending ? "Saving…" : "Save Changes"}
              </button>
            </div>
          </form>
        </Modal>
      )}

      {/* VIEW PERIOD DETAILS MODAL */}
      {viewPeriod && (
        <Modal title={`Payroll Period Details: ${viewPeriod.periodId}`} onClose={() => setViewPeriod(null)} wide>
          <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                padding: "16px 20px",
                background: "var(--surface-alt)",
                border: "1px solid var(--border)",
                borderRadius: 8,
                flexWrap: "wrap",
                gap: 12,
              }}
            >
              <div>
                <span className="mono" style={{ fontSize: "16px", fontWeight: 800, color: "#073b33" }}>
                  {viewPeriod.periodId}
                </span>
                <p style={{ margin: "4px 0 0", color: "var(--dim)", fontSize: "13px" }}>
                  {new Intl.DateTimeFormat("en-US", { month: "long" }).format(new Date(viewPeriod.year, viewPeriod.month - 1, 1))} {viewPeriod.year}
                </p>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <Status value={viewPeriod.status} />
                {viewPeriod.calculationVersion > 0 && (
                  <span className="mono" style={{ fontWeight: 700, color: "var(--accent-deep)", background: "var(--accent-pale)", padding: "3px 8px", borderRadius: 4 }}>
                    Calc Version {viewPeriod.calculationVersion}
                  </span>
                )}
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
              <div style={{ padding: "12px 14px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--surface)" }}>
                <span style={{ fontSize: "11px", color: "var(--dim)", textTransform: "uppercase", fontWeight: 700 }}>Cycle Start</span>
                <p style={{ margin: "4px 0 0", fontWeight: 650 }}>{formatDate(viewPeriod.startDate)}</p>
              </div>
              <div style={{ padding: "12px 14px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--surface)" }}>
                <span style={{ fontSize: "11px", color: "var(--dim)", textTransform: "uppercase", fontWeight: 700 }}>Cycle End</span>
                <p style={{ margin: "4px 0 0", fontWeight: 650 }}>{formatDate(viewPeriod.endDate)}</p>
              </div>
              <div style={{ padding: "12px 14px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--surface)" }}>
                <span style={{ fontSize: "11px", color: "var(--dim)", textTransform: "uppercase", fontWeight: 700 }}>Cycle Duration</span>
                <p style={{ margin: "4px 0 0", fontWeight: 650 }}>
                  {Math.round((new Date(viewPeriod.endDate).getTime() - new Date(viewPeriod.startDate).getTime()) / 86400000) + 1} Calendar Days
                </p>
              </div>
              <div style={{ padding: "12px 14px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--surface)" }}>
                <span style={{ fontSize: "11px", color: "var(--dim)", textTransform: "uppercase", fontWeight: 700 }}>Active Projects</span>
                <p style={{ margin: "4px 0 0", fontWeight: 650 }}>{viewPeriod.activeProjectIds.length} Assigned</p>
              </div>
              <div style={{ padding: "12px 14px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--surface)" }}>
                <span style={{ fontSize: "11px", color: "var(--dim)", textTransform: "uppercase", fontWeight: 700 }}>Created</span>
                <p style={{ margin: "4px 0 0", fontWeight: 650 }}>{formatDate(viewPeriod.createdAt)}</p>
              </div>
            </div>

            <div>
              <h3 style={{ margin: "0 0 10px", fontSize: "14px", fontWeight: 750 }}>Assigned Active Projects ({viewPeriod.activeProjectIds.length})</h3>
              {viewPeriod.activeProjectIds.length === 0 ? (
                <p style={{ color: "var(--dim)", fontStyle: "italic", fontSize: "13px" }}>No projects currently active for this period.</p>
              ) : (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(240px, 1fr))", gap: 8, maxHeight: "280px", overflowY: "auto", paddingRight: 4 }}>
                  {viewPeriod.activeProjectIds.map((pId) => {
                    const proj = projectMap.get(pId);
                    return (
                      <div
                        key={pId}
                        style={{
                          padding: "8px 12px",
                          border: "1px solid var(--border)",
                          borderRadius: 6,
                          background: "var(--surface-alt)",
                          display: "flex",
                          alignItems: "center",
                          gap: 10,
                        }}
                      >
                        <span className="mono" style={{ fontWeight: 750, color: "var(--accent-deep)" }}>
                          {proj?.projectId ?? `#${pId}`}
                        </span>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <strong style={{ fontSize: "12px", display: "block", textOverflow: "ellipsis", overflow: "hidden", whiteSpace: "nowrap" }}>
                            {proj?.name ?? "Unknown Project"}
                          </strong>
                          <small style={{ color: "var(--dim)", fontSize: "11px" }}>{proj?.client || proj?.manager || "Active"}</small>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 10, borderTop: "1px solid var(--border)", paddingTop: 14 }}>
              {canManage && (transitions[viewPeriod.status] ?? []).length > 0 && (
                <button
                  type="button"
                  className="secondary-button"
                  onClick={() => {
                    const id = viewPeriod.id;
                    setViewPeriod(null);
                    setStatusPeriod(id);
                  }}
                >
                  Advance Lifecycle Status
                </button>
              )}
              {canManage && (
                <button
                  type="button"
                  className="primary-button"
                  onClick={() => {
                    const p = viewPeriod;
                    setViewPeriod(null);
                    openEdit(p);
                  }}
                >
                  Edit Period
                </button>
              )}
              <button type="button" className="ghost-button" onClick={() => setViewPeriod(null)}>
                Close
              </button>
            </div>
          </div>
        </Modal>
      )}

      {/* ADVANCE STATUS MODAL */}
      {statusPeriod && (
        <Modal title="Advance payroll lifecycle status" onClose={() => setStatusPeriod(null)}>
          <form
            className="stack-form"
            onSubmit={(event) => {
              event.preventDefault();
              const form = new FormData(event.currentTarget);
              statusMutation.mutate({
                token,
                id: statusPeriod,
                status: field(form, "status") as any,
                reason: field(form, "reason"),
              });
            }}
          >
            <label>
              Next lifecycle stage
              <select name="status">
                {(transitions[periods.data?.periods.find((p) => p.id === statusPeriod)?.status ?? ""] ?? []).map((val) => (
                  <option key={val} value={val}>
                    {val}
                  </option>
                ))}
              </select>
            </label>
            <label>
              Reason for status transition
              <textarea name="reason" minLength={3} rows={3} placeholder="Explain authorization or verification outcome..." required />
            </label>
            <button className="primary-button" disabled={statusMutation.isPending}>
              {statusMutation.isPending ? "Updating…" : "Confirm status change"}
            </button>
          </form>
        </Modal>
      )}

      {/* DELETE CONFIRMATION DIALOG */}
      <ConfirmDialog
        isOpen={Boolean(deleteTarget)}
        title="Move Payroll Period to Recycle Bin?"
        message={`Are you sure you want to delete payroll period "${deleteTarget?.periodId}"? The period and its configurations will be safely moved to the Recycle Bin and can be restored at any time.`}
        confirmText="Move to Recycle Bin"
        danger
        isLoading={removeMutation.isPending}
        onConfirm={() => {
          if (deleteTarget) {
            removeMutation.mutate({ token, id: deleteTarget.id });
          }
        }}
        onCancel={() => setDeleteTarget(null)}
      />
    </>
  );
}

interface RateRuleItem {
  id: number;
  countryCode: string;
  workTypeId: number;
  workTypeName: string;
  categoryId: number | null;
  categoryName: string | null;
  ruleName: string;
  ruleType: string;
  unitRate: string;
  thresholdLow: string | null;
  thresholdHigh: string | null;
  multiplierLow: string | null;
  multiplierMid: string | null;
  multiplierHigh: string | null;
  priority: number;
  effectiveFrom: string;
  effectiveTo: string | null;
  status: string;
  supersedesRateRuleId: number | null;
  createdAt: string;
}

function RateRules({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const queryClient = useQueryClient();
  const canManage = hasPermission(permissions, "rates.manage");
  const projects = useQuery({
    queryKey: ["projects", token],
    queryFn: () => api.listProjects({ token }),
    enabled: hasPermission(permissions, "projects.read"),
  });

  const [projectId, setProjectId] = useState<number | null>(null);
  useEffect(() => {
    if (!projectId && projects.data?.projects[0]) {
      setProjectId(projects.data.projects[0].id);
    }
  }, [projectId, projects.data]);

  const config = useQuery({
    queryKey: ["project-config", token, projectId],
    queryFn: () => api.getProjectConfiguration({ token, projectId: projectId ?? 0 }),
    enabled: Boolean(projectId) && hasPermission(permissions, "rates.read"),
  });

  // Modals & action states
  const [createModal, setCreateModal] = useState(false);
  const [editRule, setEditRule] = useState<RateRuleItem | null>(null);
  const [viewRule, setViewRule] = useState<RateRuleItem | null>(null);
  const [historyRuleId, setHistoryRuleId] = useState<number | null>(null);
  const [deleteTarget, setDeleteTarget] = useState<RateRuleItem | null>(null);
  const [simulatorOpen, setSimulatorOpen] = useState(false);
  const [simInitial, setSimInitial] = useState<{
    quantity: string;
    rate: string;
    averageHours: string;
    thresholdLow: string;
    thresholdHigh: string;
    multiplierLow: string;
    multiplierMid: string;
    multiplierHigh: string;
  } | null>(null);

  // Filters
  const [modelFilter, setModelFilter] = useState("all");
  const [countryFilter, setCountryFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");
  const [categoryFilter, setCategoryFilter] = useState("all");

  // Category Rates Quick Setup / Matrix state
  const [categoryModalOpen, setCategoryModalOpen] = useState(false);
  const [categoryCountryCode, setCategoryCountryCode] = useState("");
  const [categoryWorkTypeId, setCategoryWorkTypeId] = useState<number | string>("");
  const [categoryEffectiveFrom, setCategoryEffectiveFrom] = useState(new Date().toISOString().slice(0, 10));
  const [categoryRuleType, setCategoryRuleType] = useState<"unit-rate" | "slab" | "fixed" | "hybrid">("unit-rate");
  const [categoryDraftRates, setCategoryDraftRates] = useState<Record<number, string>>({});

  // Create form state with preset & clone helpers
  const [formRuleName, setFormRuleName] = useState("Production Standard 3-Tier Slab Rate");
  const [formCountryCode, setFormCountryCode] = useState("");
  const [formWorkTypeId, setFormWorkTypeId] = useState<number | string>("");
  const [formCategoryId, setFormCategoryId] = useState<number | string>("");
  const [formRuleType, setFormRuleType] = useState<"unit-rate" | "slab" | "fixed" | "hybrid">("unit-rate");
  const [formUnitRate, setFormUnitRate] = useState("0.0225");
  const [formThresholdLow, setFormThresholdLow] = useState("3.99");
  const [formThresholdHigh, setFormThresholdHigh] = useState("5.99");
  const [formMultiplierLow, setFormMultiplierLow] = useState("0.60");
  const [formMultiplierMid, setFormMultiplierMid] = useState("0.80");
  const [formMultiplierHigh, setFormMultiplierHigh] = useState("1.00");
  const [formPriority, setFormPriority] = useState(100);
  const [formEffectiveFrom, setFormEffectiveFrom] = useState(new Date().toISOString().slice(0, 10));
  const [formEffectiveTo, setFormEffectiveTo] = useState("");
  const [formSupersedesRateRuleId, setFormSupersedesRateRuleId] = useState<number | string>("");
  const [clonedFromRule, setClonedFromRule] = useState<RateRuleItem | null>(null);

  // Edit form state
  const [editRuleType, setEditRuleType] = useState<"unit-rate" | "slab" | "fixed" | "hybrid">("unit-rate");

  // View modal interactive mini-simulator state
  const [viewSimQty, setViewSimQty] = useState("1000");
  const [viewSimHours, setViewSimHours] = useState("6.0");

  const invalidateRates = (text: string) => {
    queryClient.invalidateQueries({ queryKey: ["project-config"] });
    queryClient.invalidateQueries({ queryKey: ["dashboard"] });
    queryClient.invalidateQueries({ queryKey: ["recycle-bin"] });
    notify({ kind: "success", text });
  };

  const saveCategoryMutation = useMutation({
    mutationFn: api.saveCategoryRateRules,
    onSuccess: (data) => {
      invalidateRates(`Category rates saved successfully (${data.created} created, ${data.updated} updated).`);
      setCategoryModalOpen(false);
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const createMutation = useMutation({
    mutationFn: api.addRateRule,
    onSuccess: () => {
      invalidateRates(clonedFromRule ? "Cloned rate rule version created successfully." : "New rate rule created successfully.");
      setCreateModal(false);
      setClonedFromRule(null);
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const updateMutation = useMutation({
    mutationFn: api.updateRateRule,
    onSuccess: () => {
      invalidateRates("Rate rule updated successfully.");
      setEditRule(null);
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const removeMutation = useMutation({
    mutationFn: api.deleteRateRule,
    onSuccess: () => {
      invalidateRates("Rate rule moved to Recycle Bin.");
      setDeleteTarget(null);
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const [previewResult, setPreviewResult] = useState<{ baseAmount: string; multiplier: string; payableAmount: string } | null>(null);
  const previewMutation = useMutation({
    mutationFn: api.previewCalculationKernel,
    onSuccess: setPreviewResult,
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const historyQuery = useQuery({
    queryKey: ["rate-rule-history", token, historyRuleId],
    queryFn: () => api.getRateRuleHistory({ token, id: historyRuleId! }),
    enabled: Boolean(historyRuleId) && hasPermission(permissions, "rates.read"),
  });

  const handleExportCsv = () => {
    const headers = [
      "Rule ID",
      "Rule Name",
      "Country Code",
      "Work Type",
      "Category",
      "Calculation Model",
      "Base Unit Rate",
      "Priority",
      "Effective From",
      "Effective To",
      "Status",
      "Low Hour Threshold",
      "High Hour Threshold",
      "Low Multiplier",
      "Mid Multiplier",
      "High Multiplier",
      "Supersedes Rule ID",
      "Created Timestamp",
    ];

    const formatCsvCell = (val: unknown) => {
      let text = String(val ?? "");
      if (/^[=+\-@]/.test(text)) text = `'${text}`;
      return `"${text.replace(/"/g, '""')}"`;
    };

    const rows = filteredRules.map((r) => [
      r.id,
      r.ruleName,
      r.countryCode,
      r.workTypeName,
      r.categoryName || "All Categories",
      r.ruleType,
      r.unitRate,
      r.priority,
      r.effectiveFrom,
      r.effectiveTo || "Open-ended",
      r.status,
      r.thresholdLow || "",
      r.thresholdHigh || "",
      r.multiplierLow || "",
      r.multiplierMid || "",
      r.multiplierHigh || "",
      r.supersedesRateRuleId || "",
      r.createdAt,
    ]);

    const csvContent = [
      headers.map(formatCsvCell).join(","),
      ...rows.map((row) => row.map(formatCsvCell).join(",")),
    ].join("\r\n");

    const blob = new Blob(["\uFEFF" + csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const projectCode = currentProject?.projectId || "PRJ";
    link.href = url;
    link.download = `rate_rules_${projectCode}_${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    notify({
      kind: "success",
      text: `Exported ${filteredRules.length} rate rules to CSV file (${link.download}).`,
    });
  };

  const openSimulatorForRule = (r: RateRuleItem) => {
    setSimInitial({
      quantity: "1000",
      rate: r.unitRate,
      averageHours: "6.00",
      thresholdLow: r.thresholdLow || "3.99",
      thresholdHigh: r.thresholdHigh || "5.99",
      multiplierLow: r.multiplierLow || "0.60",
      multiplierMid: r.multiplierMid || "0.80",
      multiplierHigh: r.multiplierHigh || "1.00",
    });
    setPreviewResult(null);
    setSimulatorOpen(true);
  };

  const openEditModal = (r: RateRuleItem) => {
    setEditRule(r);
    setEditRuleType((r.ruleType as any) || "unit-rate");
  };


  const openNewRateRuleModal = () => {
    setClonedFromRule(null);
    setFormRuleName("Production Standard 3-Tier Slab Rate");
    setFormCountryCode(config.data?.countries.find((c) => c.status === "active")?.countryCode || "");
    setFormWorkTypeId(config.data?.workTypes.find((w) => w.status === "active")?.id || "");
    setFormCategoryId("");
    setFormRuleType("slab");
    setFormUnitRate("0.0225");
    setFormThresholdLow("3.99");
    setFormThresholdHigh("5.99");
    setFormMultiplierLow("0.60");
    setFormMultiplierMid("0.80");
    setFormMultiplierHigh("1.00");
    setFormPriority(100);
    setFormEffectiveFrom(new Date().toISOString().slice(0, 10));
    setFormEffectiveTo("");
    setFormSupersedesRateRuleId("");
    setCreateModal(true);
  };

  const handleCloneRateRule = (r: RateRuleItem) => {
    setClonedFromRule(r);
    setFormRuleName(`${r.ruleName} (Revision)`);
    setFormCountryCode(r.countryCode);
    setFormWorkTypeId(r.workTypeId);
    setFormCategoryId(r.categoryId ? String(r.categoryId) : "");
    setFormRuleType((r.ruleType as any) || "unit-rate");
    setFormUnitRate(r.unitRate);
    setFormPriority(r.priority);
    setFormEffectiveFrom(new Date().toISOString().slice(0, 10));
    setFormEffectiveTo("");
    setFormSupersedesRateRuleId(String(r.id));
    setFormThresholdLow(r.thresholdLow || "3.99");
    setFormThresholdHigh(r.thresholdHigh || "5.99");
    setFormMultiplierLow(r.multiplierLow || "0.60");
    setFormMultiplierMid(r.multiplierMid || "0.80");
    setFormMultiplierHigh(r.multiplierHigh || "1.00");
    setCreateModal(true);
  };

  const openCategoryRatesModal = () => {
    const defaultCountry = config.data?.countries.find((c) => c.status === "active")?.countryCode || "BD";
    const defaultWorkType = config.data?.workTypes.find((w) => w.status === "active")?.id || (config.data?.workTypes[0]?.id ?? "");
    setCategoryCountryCode(defaultCountry);
    setCategoryWorkTypeId(defaultWorkType);
    setCategoryEffectiveFrom(new Date().toISOString().slice(0, 10));
    setCategoryRuleType("unit-rate");
    const drafts: Record<number, string> = {};
    (config.data?.categories || []).forEach((cat) => {
      const existing = (config.data?.rateRules || []).find(
        (r) => r.status === "active" && r.categoryId === cat.id && r.countryCode === defaultCountry && (!defaultWorkType || r.workTypeId === Number(defaultWorkType))
      );
      drafts[cat.id] = existing?.unitRate || "";
    });
    setCategoryDraftRates(drafts);
    setCategoryModalOpen(true);
  };

  const handleCategoryScopeChange = (newCountry: string, newWorkType: number | string) => {
    setCategoryCountryCode(newCountry);
    setCategoryWorkTypeId(newWorkType);
    const drafts: Record<number, string> = {};
    (config.data?.categories || []).forEach((cat) => {
      const existing = (config.data?.rateRules || []).find(
        (r) => r.status === "active" && r.categoryId === cat.id && r.countryCode === newCountry && (!newWorkType || r.workTypeId === Number(newWorkType))
      );
      drafts[cat.id] = existing?.unitRate || "";
    });
    setCategoryDraftRates(drafts);
  };

  if (!hasPermission(permissions, "rates.read")) {
    return <Empty title="Rate and calculation access restricted" body="Your role does not include rate rules and calculation engine policies." />;
  }

  if (!projects.data?.projects.length) {
    return <Empty title="Create a project first" body="Rate rules and country models must belong to an active project definition." />;
  }

  const rateRules = config.data?.rateRules || [];
  const activeCountries = config.data?.countries || [];
  const activeWorkTypes = config.data?.workTypes || [];
  const activeCategories = config.data?.categories || [];

  const projectCountryList = useMemo(() => {
    const list: string[] = [];
    if (config.data?.project?.countries && Array.isArray(config.data.project.countries)) {
      list.push(...config.data.project.countries);
    }
    for (const c of activeCountries) {
      if (!list.includes(c.countryCode)) list.push(c.countryCode);
    }
    return list.length > 0 ? list : ["BD"];
  }, [config.data?.project?.countries, activeCountries]);

  const [activeCountryTab, setActiveCountryTab] = useState<string>("BD");
  const [countryDraftRates, setCountryDraftRates] = useState<Record<string, Record<string, string>>>({});

  useEffect(() => {
    if (projectCountryList.length > 0 && !projectCountryList.includes(activeCountryTab)) {
      setActiveCountryTab(projectCountryList[0] || "BD");
    }
  }, [projectCountryList, activeCountryTab]);

  const uniqueCountries = Array.from(new Set(rateRules.map((r) => r.countryCode))).filter(Boolean);

  const filteredRules = rateRules.filter((r) => {
    if (modelFilter !== "all" && r.ruleType !== modelFilter) return false;
    if (countryFilter !== "all" && r.countryCode !== countryFilter) return false;
    if (statusFilter !== "all" && r.status !== statusFilter) return false;
    if (categoryFilter !== "all") {
      if (categoryFilter === "general" && r.categoryId !== null) return false;
      if (categoryFilter !== "general" && String(r.categoryId) !== categoryFilter) return false;
    }
    return true;
  });

  // Metrics
  const totalRules = rateRules.length;
  const activeRulesCount = rateRules.filter((r) => r.status === "active").length;
  const slabRulesCount = rateRules.filter((r) => r.ruleType === "slab" || r.ruleType === "hybrid").length;
  const unitRulesCount = rateRules.filter((r) => r.ruleType === "unit-rate" || r.ruleType === "fixed").length;
  const coveredCountriesCount = new Set(rateRules.filter((r) => r.status === "active").map((r) => r.countryCode)).size;

  const currentProject = projects.data.projects.find((p) => p.id === projectId);

  const todayIso = new Date().toISOString().slice(0, 10);

  const columns: Column<RateRuleItem>[] = [
    {
      key: "ruleName",
      header: "Rule Name & Scope",
      isImportant: true,
      render: (r) => {
        const isPast = Boolean(r.effectiveTo && r.effectiveTo < todayIso);
        return (
          <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
              <span
                className="mono"
                style={{
                  fontWeight: 750,
                  fontSize: "11px",
                  color: isPast ? "#4b5563" : "#073b33",
                  background: isPast ? "#f3f4f6" : "#f0f8f5",
                  border: isPast ? "1px solid #d1d5db" : "1px solid #99c2b6",
                  padding: "2px 6px",
                  borderRadius: 4,
                }}
              >
                {r.countryCode}
              </span>
              <strong style={{ fontSize: "13px", color: isPast ? "#6b7280" : "var(--text)" }}>{r.ruleName}</strong>
              {isPast && (
                <span
                  className="mono"
                  style={{
                    fontSize: "10px",
                    fontWeight: 800,
                    padding: "1px 6px",
                    borderRadius: 3,
                    background: "#fef3c7",
                    color: "#92400e",
                    border: "1px solid #fde68a",
                    letterSpacing: "0.04em",
                  }}
                  title={`Superseded rule: Effective validity window closed on ${formatDate(r.effectiveTo!)}`}
                >
                  SUPERSEDED
                </span>
              )}
            </div>
            <div style={{ fontSize: "11.5px", color: "var(--dim)" }}>
              <span>{r.workTypeName}</span>
              {r.categoryName && <span> · Cat: {r.categoryName}</span>}
            </div>
            {r.supersedesRateRuleId && (
              <small style={{ fontSize: "10.5px", color: "#6b7280" }}>
                Supersedes Rule #{r.supersedesRateRuleId}
              </small>
            )}
          </div>
        );
      },
    },
    {
      key: "ruleType",
      header: "Calculation Model",
      render: (r) => {
        const isSlab = r.ruleType === "slab";
        const isUnit = r.ruleType === "unit-rate";
        const isFixed = r.ruleType === "fixed";
        const isHybrid = r.ruleType === "hybrid";

        const bg = isSlab ? "#f3e8ff" : isUnit ? "#ccfbf1" : isFixed ? "#fef3c7" : isHybrid ? "#dbeafe" : "#f1f5f9";
        const color = isSlab ? "#6b21a8" : isUnit ? "#0f766e" : isFixed ? "#92400e" : isHybrid ? "#1e40af" : "#475569";
        const border = isSlab ? "#d8b4fe" : isUnit ? "#99f6e4" : isFixed ? "#fde68a" : isHybrid ? "#bfdbfe" : "#cbd5e1";

        return (
          <div>
            <span
              className="mono"
              style={{
                display: "inline-block",
                padding: "2px 7px",
                fontSize: "11px",
                fontWeight: 700,
                borderRadius: 4,
                background: bg,
                color,
                border: `1px solid ${border}`,
                textTransform: "uppercase",
              }}
            >
              {r.ruleType}
            </span>
            <small style={{ display: "block", color: "var(--dim)", fontSize: "11px", marginTop: 2 }}>
              {isSlab ? "Productivity Slab" : isUnit ? "Linear Unit Rate" : isFixed ? "Fixed Lump Sum" : "Composite Hybrid"}
            </small>
          </div>
        );
      },
    },
    {
      key: "unitRate",
      header: "Base Rate",
      isNumeric: true,
      render: (r) => (
        <div>
          <strong className="mono" style={{ fontSize: "13.5px", color: "#073b33" }}>
            {r.unitRate}
          </strong>
          <small style={{ display: "block", color: "var(--dim)", fontSize: "10.5px" }}>
            {r.ruleType === "fixed" ? "flat amount" : "per unit / char"}
          </small>
        </div>
      ),
    },
    {
      key: "slabFormula",
      header: "Slab Formula / Tiers",
      render: (r) => {
        if (r.ruleType === "slab" || r.ruleType === "hybrid") {
          return (
            <div style={{ display: "flex", flexDirection: "column", gap: 2, fontSize: "11px" }}>
              <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
                <span className="mono" style={{ color: "#b45309", background: "#fef3c7", padding: "1px 5px", borderRadius: 3, fontSize: "10.5px" }}>
                  &lt;{r.thresholdLow}h: {r.multiplierLow}×
                </span>
                <span className="mono" style={{ color: "#2563eb", background: "#dbeafe", padding: "1px 5px", borderRadius: 3, fontSize: "10.5px" }}>
                  {r.thresholdLow}-{r.thresholdHigh}h: {r.multiplierMid}×
                </span>
                <span className="mono" style={{ color: "#059669", background: "#d1fae5", padding: "1px 5px", borderRadius: 3, fontSize: "10.5px" }}>
                  ≥{r.thresholdHigh}h: {r.multiplierHigh}×
                </span>
              </div>
            </div>
          );
        }
        return <span style={{ color: "var(--dim)", fontSize: "11.5px" }}>Single factor (1.00×)</span>;
      },
    },
    {
      key: "priority",
      header: "Priority",
      render: (r) => (
        <span
          className="mono"
          style={{
            fontWeight: 700,
            padding: "2px 6px",
            background: "var(--surface-alt)",
            border: "1px solid var(--border)",
            borderRadius: 4,
            fontSize: "11.5px",
          }}
        >
          P{r.priority}
        </span>
      ),
    },
    {
      key: "effectiveWindow",
      header: "Validity Window",
      render: (r) => {
        const isPast = Boolean(r.effectiveTo && r.effectiveTo < todayIso);
        return (
          <div style={{ fontSize: "11.5px" }}>
            <span style={{ color: isPast ? "var(--dim)" : "var(--text)" }}>{formatDate(r.effectiveFrom)}</span>
            <small
              style={{
                display: "block",
                color: isPast ? "#b45309" : "var(--dim)",
                fontSize: "10.5px",
                fontWeight: isPast ? 700 : 400,
              }}
            >
              {r.effectiveTo ? (isPast ? `Expired: ${formatDate(r.effectiveTo)}` : `to ${formatDate(r.effectiveTo)}`) : "open-ended"}
            </small>
          </div>
        );
      },
    },
    {
      key: "status",
      header: "Status",
      render: (r) => {
        const isPast = Boolean(r.effectiveTo && r.effectiveTo < todayIso);
        return (
          <div style={{ display: "flex", flexDirection: "column", gap: 3, alignItems: "flex-start" }}>
            <Status value={r.status} />
            {isPast && (
              <span
                className="mono"
                style={{
                  fontSize: "9.5px",
                  fontWeight: 750,
                  color: "#92400e",
                  background: "#fffbeb",
                  padding: "1px 5px",
                  borderRadius: 3,
                  border: "1px solid #fde68a",
                }}
              >
                PAST WINDOW
              </span>
            )}
          </div>
        );
      },
    },
  ];

  return (
    <>
      {/* 1. PROJECT SELECTOR & TOP CONTROLS */}
      <div className="toolbar" style={{ flexWrap: "wrap", gap: 12 }}>
        <div style={{ display: "flex", gap: 12, alignItems: "flex-end", flex: "1 1 340px" }}>
          <label className="select-field" style={{ margin: 0 }}>
            Active Project
            <select
              value={projectId ?? ""}
              onChange={(e) => setProjectId(Number(e.target.value))}
              style={{ fontWeight: 650 }}
            >
              {projects.data.projects.map((project) => (
                <option key={project.id} value={project.id}>
                  {project.projectId} · {project.name}
                </option>
              ))}
            </select>
          </label>
        </div>

        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <button
            type="button"
            className="secondary-button compact"
            onClick={handleExportCsv}
            title="Download current filtered list of rate rules as a CSV file"
          >
            <Icon name="download" />
            <span>Export to CSV</span>
          </button>
          <button
            type="button"
            className="secondary-button compact"
            onClick={() => {
              setSimInitial(null);
              setPreviewResult(null);
              setSimulatorOpen(true);
            }}
            title="Open calculation kernel simulator"
          >
            <Icon name="calculator" />
            <span>Calculation Simulator</span>
          </button>
          {canManage && (
            <button
              type="button"
              className="primary-button compact"
              onClick={openCategoryRatesModal}
              title="Set rates for Header Character, Birth Character, Marriage Character, Death Character, Proof (+/-), etc."
              style={{ background: "#0f766e", borderColor: "#0d5f58" }}
            >
              <Icon name="tag" />
              <span>Category Rates Matrix</span>
            </button>
          )}
          {canManage && (
            <button
              type="button"
              className="primary-button compact"
              onClick={openNewRateRuleModal}
            >
              <Icon name="plus" />
              <span>Add Rate Rule</span>
            </button>
          )}
        </div>
      </div>

      {/* 2. FIVE-METRIC LEDGER */}
      <div className="metric-ledger" style={{ marginTop: 12 }}>
        <article>
          <p>Total Rules</p>
          <strong>{totalRules}</strong>
          <small>Project versions</small>
        </article>
        <article>
          <p>Active Production</p>
          <strong style={{ color: "var(--accent-deep)" }}>{activeRulesCount}</strong>
          <small>Authoritative scope</small>
        </article>
        <article>
          <p>Slab Tier Rules</p>
          <strong>{slabRulesCount}</strong>
          <small>3-tier productivity</small>
        </article>
        <article>
          <p>Flat / Unit Rules</p>
          <strong>{unitRulesCount}</strong>
          <small>Linear & fixed payout</small>
        </article>
        <article>
          <p>Country Scopes</p>
          <strong>{coveredCountriesCount}</strong>
          <small>Active geographic tiers</small>
        </article>
      </div>

      {/* 3. COUNTRY-WISE RATE CONFIGURATION (Only shows options enabled on the selected project) */}
      <div className="category-rates-strip" style={{ marginTop: 14, marginBottom: 14, background: "var(--surface-alt)", border: "1px solid var(--border)", borderRadius: 8, padding: "16px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12, flexWrap: "wrap", gap: 10 }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ color: "#0f766e", display: "flex" }}>
                <Icon name="calculator" />
              </span>
              <strong style={{ fontSize: "14px", color: "var(--accent-deep)" }}>
                Country-Wise Rate Configuration
              </strong>
            </div>
            <p style={{ margin: "3px 0 0", fontSize: "12px", color: "var(--dim)" }}>
              Configure unit rates per country. Only features enabled for <b>{currentProject?.projectId || "this project"}</b> (Character types, Proof (+/-), DocType, Incentives) are available.
            </p>
          </div>
          {canManage && (
            <div style={{ display: "flex", gap: 8 }}>
              <button
                type="button"
                className="secondary-button compact"
                style={{ fontSize: "11.5px", padding: "4px 10px" }}
                onClick={openCategoryRatesModal}
              >
                <Icon name="pencil" />
                <span>Open Matrix Modal</span>
              </button>
            </div>
          )}
        </div>

        {/* Country Selector Tabs */}
        <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 16, flexWrap: "wrap", borderBottom: "1px solid var(--border)", paddingBottom: 10 }}>
          <span style={{ fontSize: "12px", fontWeight: 700, color: "var(--dim)", marginRight: 6 }}>
            Select Country:
          </span>
          {projectCountryList.map((cCode) => {
            const isSelected = activeCountryTab === cCode;
            const countryRules = rateRules.filter((r) => r.status === "active" && r.countryCode === cCode);
            return (
              <button
                key={cCode}
                type="button"
                onClick={() => setActiveCountryTab(cCode)}
                style={{
                  padding: "6px 14px",
                  fontSize: "12.5px",
                  fontWeight: isSelected ? 800 : 600,
                  borderRadius: "6px",
                  border: isSelected ? "1.5px solid #0f766e" : "1px solid var(--border)",
                  background: isSelected ? "#0f766e" : "var(--surface)",
                  color: isSelected ? "#ffffff" : "var(--text)",
                  cursor: "pointer",
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 6,
                  transition: "all 0.15s ease",
                }}
              >
                <span>{cCode}</span>
                <span
                  style={{
                    fontSize: "10.5px",
                    padding: "1px 5px",
                    borderRadius: "3px",
                    background: isSelected ? "rgba(255,255,255,0.25)" : "var(--surface-alt)",
                    color: isSelected ? "#fff" : "var(--dim)",
                    fontWeight: 700,
                  }}
                >
                  {countryRules.length} set
                </span>
              </button>
            );
          })}
        </div>

        {/* Rates Grid for activeCountryTab */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(260px, 1fr))", gap: 12 }}>
          {/* Base Work Type Rate */}
          {(() => {
            const baseRule = rateRules.find((r) => r.status === "active" && r.countryCode === activeCountryTab && r.categoryId === null);
            const draftVal = countryDraftRates[activeCountryTab]?.["base"] ?? "";
            const currentVal = draftVal !== "" ? draftVal : (baseRule?.unitRate || "");
            return (
              <div
                style={{
                  background: "var(--surface)",
                  border: "1px solid var(--border)",
                  borderRadius: 6,
                  padding: "12px",
                  display: "flex",
                  flexDirection: "column",
                  justifyContent: "space-between",
                  gap: 8,
                }}
              >
                <div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <strong style={{ fontSize: "12.5px", color: "var(--text)" }}>Base Unit Rate</strong>
                    <span style={{ fontSize: "10.5px", padding: "1px 6px", background: "#e0f2fe", color: "#0369a1", borderRadius: 3, fontWeight: 700 }}>
                      General
                    </span>
                  </div>
                  <small style={{ fontSize: "11px", color: "var(--dim)", display: "block", marginTop: 2 }}>
                    Base rate for {currentProject?.paymentBasis || "record"} entries in {activeCountryTab}
                  </small>
                </div>
                <div>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: "11px", color: "var(--dim)", marginBottom: 4 }}>
                    <span>Active: <b className="mono" style={{ color: baseRule ? "#073b33" : "var(--dim)" }}>{baseRule ? baseRule.unitRate : "Not set"}</b></span>
                  </div>
                  <input
                    inputMode="decimal"
                    placeholder="e.g. 0.0225"
                    disabled={!canManage}
                    value={currentVal}
                    onChange={(e) => {
                      const val = e.target.value;
                      setCountryDraftRates((prev) => ({
                        ...prev,
                        [activeCountryTab]: { ...(prev[activeCountryTab] || {}), base: val },
                      }));
                    }}
                    style={{ fontSize: "13px", padding: "6px 8px", width: "100%", fontFamily: "var(--mono)", fontWeight: 700 }}
                  />
                </div>
              </div>
            );
          })()}

          {/* Enabled Categories (Character types, Proof (+), Proof (-), DocType, Incentive) */}
          {activeCategories.map((cat) => {
            const activeRule = rateRules.find(
              (r) => r.status === "active" && r.countryCode === activeCountryTab && (r.categoryId === cat.id || r.categoryName === cat.name)
            );
            const isProofMinus = cat.name.toLowerCase().includes("proof (-)");
            const isProofPlus = cat.name.toLowerCase().includes("proof (+)");
            const isDocTypeCat = cat.name.toLowerCase() === "doctype";
            const isIncentiveCat = cat.name.toLowerCase().includes("incentive");

            const draftVal = countryDraftRates[activeCountryTab]?.[String(cat.id)] ?? "";
            const currentVal = draftVal !== "" ? draftVal : (activeRule?.unitRate || "");

            let badgeColor = "#e8f3ef";
            let badgeTextColor = "#073b33";
            let badgeText = "Character";
            if (isProofPlus) {
              badgeColor = "#dcfce7";
              badgeTextColor = "#15803d";
              badgeText = "Proof (+)";
            } else if (isProofMinus) {
              badgeColor = "#fee2e2";
              badgeTextColor = "#b91c1c";
              badgeText = "Proof (-)";
            } else if (isDocTypeCat) {
              badgeColor = "#fef3c7";
              badgeTextColor = "#92400e";
              badgeText = "DocType";
            } else if (isIncentiveCat) {
              badgeColor = "#f3e8ff";
              badgeTextColor = "#6b21a8";
              badgeText = "Incentive";
            }

            return (
              <div
                key={cat.id}
                style={{
                  background: isProofMinus ? "#fffbfb" : isProofPlus ? "#fbfdfb" : "var(--surface)",
                  border: isProofMinus ? "1px solid #fca5a5" : isProofPlus ? "1px solid #86efac" : "1px solid var(--border)",
                  borderRadius: 6,
                  padding: "12px",
                  display: "flex",
                  flexDirection: "column",
                  justifyContent: "space-between",
                  gap: 8,
                }}
              >
                <div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <strong style={{ fontSize: "12.5px", color: "var(--text)" }}>{cat.name}</strong>
                    <span style={{ fontSize: "10.5px", padding: "1px 6px", background: badgeColor, color: badgeTextColor, borderRadius: 3, fontWeight: 700 }}>
                      {badgeText}
                    </span>
                  </div>
                  <small style={{ fontSize: "11px", color: "var(--dim)", display: "block", marginTop: 2 }}>
                    {isProofMinus
                      ? "Granular Proof deduction adjustment rate (can be negative)"
                      : isProofPlus
                        ? "Granular Proof positive adjustment rate"
                        : `${cat.name} unit rate for ${activeCountryTab}`}
                  </small>
                </div>
                <div>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: "11px", color: "var(--dim)", marginBottom: 4 }}>
                    <span>Active: <b className="mono" style={{ color: activeRule ? (isProofMinus ? "#b91c1c" : "#073b33") : "var(--dim)" }}>{activeRule ? activeRule.unitRate : "Not set"}</b></span>
                    <span className="mono" style={{ fontSize: "10.5px" }}>{cat.code}</span>
                  </div>
                  <input
                    inputMode="decimal"
                    placeholder={isProofMinus ? "e.g. -0.0150" : "e.g. 0.0250"}
                    disabled={!canManage}
                    value={currentVal}
                    onChange={(e) => {
                      const val = e.target.value;
                      setCountryDraftRates((prev) => ({
                        ...prev,
                        [activeCountryTab]: { ...(prev[activeCountryTab] || {}), [String(cat.id)]: val },
                      }));
                    }}
                    style={{
                      fontSize: "13px",
                      padding: "6px 8px",
                      width: "100%",
                      fontFamily: "var(--mono)",
                      fontWeight: 700,
                      color: isProofMinus && currentVal.startsWith("-") ? "#b91c1c" : undefined,
                    }}
                  />
                </div>
              </div>
            );
          })}
        </div>

        {/* Save Bar for activeCountryTab */}
        {canManage && (
          <div style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", marginTop: 14, paddingTop: 10, borderTop: "1px solid var(--border)", gap: 10 }}>
            <span style={{ fontSize: "12px", color: "var(--dim)" }}>
              Applying updates to <b>{activeCountryTab}</b>
            </span>
            <button
              type="button"
              className="primary-button"
              disabled={saveCategoryMutation.isPending}
              onClick={() => {
                const ratesToSave: { categoryId: number | null; unitRate: string; ruleType: "unit-rate"; ruleName?: string }[] = [];
                // Save category rates
                activeCategories.forEach((cat) => {
                  const draftVal = countryDraftRates[activeCountryTab]?.[String(cat.id)];
                  const activeRule = rateRules.find(
                    (r) => r.status === "active" && r.countryCode === activeCountryTab && (r.categoryId === cat.id || r.categoryName === cat.name)
                  );
                  const effectiveVal = draftVal !== undefined ? draftVal.trim() : (activeRule?.unitRate || "");
                  if (effectiveVal && !isNaN(Number(effectiveVal))) {
                    ratesToSave.push({
                      categoryId: cat.id,
                      unitRate: effectiveVal,
                      ruleType: "unit-rate",
                      ruleName: `${activeCountryTab} - ${cat.name} Rate`,
                    });
                  }
                });
                // Save base rate
                const draftBase = countryDraftRates[activeCountryTab]?.["base"];
                const baseRule = rateRules.find((r) => r.status === "active" && r.countryCode === activeCountryTab && r.categoryId === null);
                const effectiveBase = draftBase !== undefined ? draftBase.trim() : (baseRule?.unitRate || "");
                if (effectiveBase && !isNaN(Number(effectiveBase))) {
                  ratesToSave.push({
                    categoryId: null,
                    unitRate: effectiveBase,
                    ruleType: "unit-rate",
                    ruleName: `${activeCountryTab} - Base ${currentProject?.paymentBasis || "Record"} Rate`,
                  });
                }

                if (ratesToSave.length === 0) {
                  notify({ kind: "error", text: "Please enter at least one valid rate before saving." });
                  return;
                }

                saveCategoryMutation.mutate({
                  token,
                  projectId: projectId!,
                  countryCode: activeCountryTab,
                  workTypeId: Number(activeWorkTypes[0]?.id || 1),
                  effectiveFrom: todayIso,
                  effectiveTo: null,
                  rates: ratesToSave,
                });
              }}
            >
              <Icon name="check" />
              <span>{saveCategoryMutation.isPending ? "Saving..." : `Save Rates for ${activeCountryTab}`}</span>
            </button>
          </div>
        )}
      </div>

      {/* 3. MAIN DATA TABLE */}
      {config.isPending ? (
        <Loading />
      ) : config.error || !config.data ? (
        <Empty title="Rate rules could not load" body={message(config.error)} />
      ) : (
        <DataTable
          data={filteredRules}
          columns={columns}
          title={`Rates (${currentProject?.projectId || "Project"})`}
          subtitle="Payment bases, country rates, and version history"
          searchPlaceholder="Search rules by name, country, work type, model, rate, or priority..."
          searchFilter={(r, term) =>
            r.ruleName.toLowerCase().includes(term) ||
            r.countryCode.toLowerCase().includes(term) ||
            r.workTypeName.toLowerCase().includes(term) ||
            (r.categoryName || "").toLowerCase().includes(term) ||
            r.ruleType.toLowerCase().includes(term) ||
            r.unitRate.toLowerCase().includes(term) ||
            String(r.priority).includes(term) ||
            r.status.toLowerCase().includes(term)
          }
          extraFilters={
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
              <label className="select-field" style={{ margin: 0 }}>
                <span style={{ fontSize: "11.5px", fontWeight: 650, color: "var(--dim)", marginRight: 6 }}>Model:</span>
                <select
                  value={modelFilter}
                  onChange={(e) => setModelFilter(e.target.value)}
                  style={{ fontSize: "12px", padding: "4px 8px", height: "auto" }}
                >
                  <option value="all">All Models</option>
                  <option value="unit-rate">Unit Rate</option>
                  <option value="slab">Slab-Based</option>
                  <option value="fixed">Fixed Amount</option>
                  <option value="hybrid">Hybrid</option>
                </select>
              </label>

              {uniqueCountries.length > 1 && (
                <label className="select-field" style={{ margin: 0 }}>
                  <span style={{ fontSize: "11.5px", fontWeight: 650, color: "var(--dim)", marginRight: 6 }}>Country:</span>
                  <select
                    value={countryFilter}
                    onChange={(e) => setCountryFilter(e.target.value)}
                    style={{ fontSize: "12px", padding: "4px 8px", height: "auto" }}
                  >
                    <option value="all">All Countries</option>
                    {uniqueCountries.map((c) => (
                      <option key={c} value={c}>{c}</option>
                    ))}
                  </select>
                </label>
              )}

              <label className="select-field" style={{ margin: 0 }}>
                <span style={{ fontSize: "11.5px", fontWeight: 650, color: "var(--dim)", marginRight: 6 }}>Category:</span>
                <select
                  value={categoryFilter}
                  onChange={(e) => setCategoryFilter(e.target.value)}
                  style={{ fontSize: "12px", padding: "4px 8px", height: "auto" }}
                >
                  <option value="all">All Categories</option>
                  <option value="general">General (No Category)</option>
                  {activeCategories.map((c) => (
                    <option key={c.id} value={String(c.id)}>{c.name}</option>
                  ))}
                </select>
              </label>

              <label className="select-field" style={{ margin: 0 }}>
                <span style={{ fontSize: "11.5px", fontWeight: 650, color: "var(--dim)", marginRight: 6 }}>Status:</span>
                <select
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value)}
                  style={{ fontSize: "12px", padding: "4px 8px", height: "auto" }}
                >
                  <option value="all">All Statuses ({rateRules.length})</option>
                  <option value="active">Active Only ({activeRulesCount})</option>
                  <option value="inactive">Inactive</option>
                </select>
              </label>

              <button
                type="button"
                className="secondary-button compact"
                onClick={handleExportCsv}
                style={{ fontSize: "12px", padding: "4px 10px", height: "auto" }}
                title="Download current filtered list of rate rules as a CSV file"
              >
                <Icon name="download" />
                <span>Export CSV</span>
              </button>
            </div>
          }
          rowClassName={(r) => (r.effectiveTo && r.effectiveTo < todayIso ? "rate-rule-superseded-row" : "")}
          onView={(r) => {
            setViewRule(r);
            setViewSimQty("1000");
            setViewSimHours("6.0");
          }}
          onEdit={canManage ? (r) => openEditModal(r) : undefined}
          onDelete={canManage ? (r) => setDeleteTarget(r) : undefined}
          customActions={(r) => (
            <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
              {canManage && (
                <button
                  type="button"
                  className="secondary-button compact"
                  onClick={() => handleCloneRateRule(r)}
                  title="Clone this rate rule to pre-fill a new versioned rule"
                  style={{ fontSize: "11px", padding: "3px 8px", whiteSpace: "nowrap" }}
                >
                  <Icon name="copy" />
                  <span>Clone</span>
                </button>
              )}
              <button
                type="button"
                className="secondary-button compact"
                onClick={() => setHistoryRuleId(r.id)}
                title="View version history and all superseded versions of this rule"
                style={{ fontSize: "11px", padding: "3px 8px", whiteSpace: "nowrap" }}
              >
                <Icon name="history" />
                <span>History</span>
              </button>
              <button
                type="button"
                className="secondary-button compact"
                onClick={() => openSimulatorForRule(r)}
                title="Test payout kernel with this rule's rate and slab formula"
                style={{ fontSize: "11px", padding: "3px 8px", whiteSpace: "nowrap" }}
              >
                Simulate
              </button>
            </div>
          )}
        />
      )}

      {/* CREATE RATE RULE MODAL */}
      {createModal && projectId && (
        <Modal
          title={clonedFromRule ? `Clone Rate Rule #${clonedFromRule.id} → New Version` : "Add Rate Rule Version"}
          onClose={() => {
            setCreateModal(false);
            setClonedFromRule(null);
          }}
          wide
        >
          {clonedFromRule && (
            <div
              style={{
                background: "var(--accent-pale)",
                border: "1px solid #99c2b6",
                borderRadius: 6,
                padding: "10px 14px",
                marginBottom: 14,
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                gap: 12,
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ color: "var(--accent-deep)", display: "flex" }}>
                  <Icon name="copy" />
                </span>
                <div>
                  <strong style={{ fontSize: "12.5px", color: "#073b33" }}>
                    Cloning from Rule #{clonedFromRule.id}: {clonedFromRule.ruleName}
                  </strong>
                  <p style={{ margin: "2px 0 0", fontSize: "11.5px", color: "var(--dim)" }}>
                    Pre-populated from Rule #{clonedFromRule.id} ({clonedFromRule.countryCode} · {clonedFromRule.workTypeName}). Configured with today's effective start date and supersession link.
                  </p>
                </div>
              </div>
              <button
                type="button"
                className="secondary-button compact"
                onClick={() => {
                  setClonedFromRule(null);
                  setFormSupersedesRateRuleId("");
                }}
                style={{ fontSize: "11px", padding: "3px 8px", whiteSpace: "nowrap" }}
                title="Disconnect clone reference and create as independent rule"
              >
                Clear clone link
              </button>
            </div>
          )}


          <form
            className="form-grid"
            onSubmit={(event) => {
              event.preventDefault();
              const f = new FormData(event.currentTarget);
              const from = field(f, "effectiveFrom");
              const to = field(f, "effectiveTo") || null;
              if (to && to < from) {
                notify({ kind: "error", text: "Effective To date must be on or after Effective From date." });
                return;
              }

              const rType = field(f, "ruleType") as "unit-rate" | "slab" | "fixed" | "hybrid";
              if (rType === "slab") {
                const lowT = field(f, "thresholdLow");
                const highT = field(f, "thresholdHigh");
                const lowM = field(f, "multiplierLow");
                const midM = field(f, "multiplierMid");
                const highM = field(f, "multiplierHigh");
                if (!lowT || !highT || !lowM || !midM || !highM) {
                  notify({ kind: "error", text: "Slab rule requires both thresholds and all three multipliers." });
                  return;
                }
              }

              createMutation.mutate({
                token,
                projectId,
                countryCode: field(f, "countryCode"),
                workTypeId: Number(field(f, "workTypeId")),
                categoryId: field(f, "categoryId") ? Number(field(f, "categoryId")) : null,
                ruleName: field(f, "ruleName"),
                ruleType: rType,
                unitRate: field(f, "unitRate"),
                thresholdLow: field(f, "thresholdLow") || null,
                thresholdHigh: field(f, "thresholdHigh") || null,
                multiplierLow: field(f, "multiplierLow") || null,
                multiplierMid: field(f, "multiplierMid") || null,
                multiplierHigh: field(f, "multiplierHigh") || null,
                priority: Number(field(f, "priority")),
                effectiveFrom: from,
                effectiveTo: to,
                supersedesRateRuleId: field(f, "supersedesRateRuleId") ? Number(field(f, "supersedesRateRuleId")) : null,
              });
            }}
          >
            <label className="full">
              Rule Name
              <input
                name="ruleName"
                placeholder="e.g. Bangladesh Review Key 3-Tier Slab Rate"
                value={formRuleName}
                onChange={(e) => setFormRuleName(e.target.value)}
                required
              />
            </label>

            <label>
              Country Scope
              <select
                name="countryCode"
                value={formCountryCode}
                onChange={(e) => setFormCountryCode(e.target.value)}
                required
              >
                <option value="">Select country</option>
                {formCountryCode && !activeCountries.some((c) => c.countryCode === formCountryCode) && (
                  <option value={formCountryCode}>{formCountryCode} (Pre-selected)</option>
                )}
                {activeCountries
                  .filter((c) => c.status === "active")
                  .map((c) => (
                    <option key={c.id} value={c.countryCode}>
                      {c.countryCode} · {c.countryName}
                    </option>
                  ))}
              </select>
            </label>

            <label>
              Work Type
              <select
                name="workTypeId"
                value={formWorkTypeId}
                onChange={(e) => setFormWorkTypeId(e.target.value)}
                required
              >
                <option value="">Select work type</option>
                {formWorkTypeId && !activeWorkTypes.some((w) => w.id === Number(formWorkTypeId)) && (
                  <option value={formWorkTypeId}>Work Type #{formWorkTypeId} (Pre-selected)</option>
                )}
                {activeWorkTypes
                  .filter((w) => w.status === "active")
                  .map((w) => (
                    <option key={w.id} value={w.id}>
                      {w.name} ({w.paymentBasis} · {w.unit})
                    </option>
                  ))}
              </select>
            </label>

            <div className="full" style={{ marginBottom: 2 }}>
              <span style={{ fontSize: "11.5px", fontWeight: 650, color: "var(--dim)", display: "block", marginBottom: 6 }}>
                Quick Select Category (Header / Birth / Marriage / Death Character):
              </span>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {activeCategories.map((c) => {
                  const isSelected = String(formCategoryId) === String(c.id);
                  return (
                    <button
                      key={c.id}
                      type="button"
                      className={`category-chip-btn ${isSelected ? "active" : ""}`}
                      onClick={() => {
                        setFormCategoryId(c.id);
                        setFormRuleName(`${formCountryCode || "BD"} - ${c.name} Rate`);
                      }}
                    >
                      {c.name}
                    </button>
                  );
                })}
                <button
                  type="button"
                  className={`category-chip-btn ${formCategoryId === "" ? "active" : ""}`}
                  onClick={() => {
                    setFormCategoryId("");
                    setFormRuleName("Production Standard 3-Tier Slab Rate");
                  }}
                >
                  All Categories (General)
                </button>
              </div>
            </div>

            <label>
              Character Category (Optional)
              <select
                name="categoryId"
                value={formCategoryId}
                onChange={(e) => {
                  const val = e.target.value;
                  setFormCategoryId(val);
                  if (val) {
                    const c = activeCategories.find((cat) => cat.id === Number(val));
                    if (c) setFormRuleName(`${formCountryCode || "BD"} - ${c.name} Rate`);
                  }
                }}
              >
                <option value="">All Categories (General basis)</option>
                {formCategoryId && !activeCategories.some((c) => c.id === Number(formCategoryId)) && (
                  <option value={formCategoryId}>Category #{formCategoryId} (Pre-selected)</option>
                )}
                {activeCategories
                  .filter((c) => c.status === "active")
                  .map((c) => (
                    <option key={c.id} value={c.id}>
                      {c.name} ({c.code})
                    </option>
                  ))}
              </select>
            </label>

            <label>
              Calculation Model
              <select
                name="ruleType"
                value={formRuleType}
                onChange={(e) => setFormRuleType(e.target.value as any)}
                required
              >
                <option value="unit-rate">Unit Rate (Linear)</option>
                <option value="slab">Slab-Based (Hour Thresholds)</option>
                <option value="fixed">Fixed Lump-Sum</option>
                <option value="hybrid">Hybrid</option>
              </select>
            </label>

            <label>
              Base Unit Rate
              <input
                name="unitRate"
                inputMode="decimal"
                value={formUnitRate}
                onChange={(e) => setFormUnitRate(e.target.value)}
                placeholder="0.0225"
                required
              />
            </label>

            <label>
              Priority (1-999)
              <input
                name="priority"
                type="number"
                min="1"
                max="999"
                value={formPriority}
                onChange={(e) => setFormPriority(Number(e.target.value))}
                required
              />
            </label>

            <label>
              Effective From
              <input
                name="effectiveFrom"
                type="date"
                value={formEffectiveFrom}
                onChange={(e) => setFormEffectiveFrom(e.target.value)}
                required
              />
            </label>

            <label>
              Effective To (Optional)
              <input
                name="effectiveTo"
                type="date"
                value={formEffectiveTo}
                onChange={(e) => setFormEffectiveTo(e.target.value)}
              />
            </label>

            <label className="full">
              Supersedes Existing Active Rule
              <select
                name="supersedesRateRuleId"
                value={formSupersedesRateRuleId}
                onChange={(e) => setFormSupersedesRateRuleId(e.target.value)}
              >
                <option value="">No previous rule (New standalone policy)</option>
                {rateRules.map((r) => (
                  <option key={r.id} value={r.id}>
                    #{r.id} · {r.ruleName} ({r.countryCode} · {r.workTypeName} · {r.unitRate}) {r.status !== "active" ? `[${r.status}]` : ""}
                  </option>
                ))}
              </select>
            </label>

            {(formRuleType === "slab" || formRuleType === "hybrid") && (
              <div
                className="full"
                style={{
                  background: "var(--surface-alt)",
                  border: "1px solid var(--border)",
                  borderRadius: 8,
                  padding: "16px",
                  display: "flex",
                  flexDirection: "column",
                  gap: 12,
                }}
              >
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <strong style={{ fontSize: "13px", color: "var(--accent-deep)" }}>
                    Slab Thresholds & Multipliers (Hours vs. Payout Rate)
                  </strong>
                  <span style={{ fontSize: "11px", color: "var(--dim)" }}>Preserved Server Kernel</span>
                </div>

                <div className="slab-fields" style={{ margin: 0 }}>
                  <label>
                    Low Hour Threshold
                    <input
                      name="thresholdLow"
                      inputMode="decimal"
                      value={formThresholdLow}
                      onChange={(e) => setFormThresholdLow(e.target.value)}
                      placeholder="3.99"
                      required
                    />
                  </label>
                  <label>
                    High Hour Threshold
                    <input
                      name="thresholdHigh"
                      inputMode="decimal"
                      value={formThresholdHigh}
                      onChange={(e) => setFormThresholdHigh(e.target.value)}
                      placeholder="5.99"
                      required
                    />
                  </label>
                  <label>
                    Low Multiplier (&lt; Low)
                    <input
                      name="multiplierLow"
                      inputMode="decimal"
                      value={formMultiplierLow}
                      onChange={(e) => setFormMultiplierLow(e.target.value)}
                      placeholder="0.60"
                      required
                    />
                  </label>
                  <label>
                    Mid Multiplier (Between)
                    <input
                      name="multiplierMid"
                      inputMode="decimal"
                      value={formMultiplierMid}
                      onChange={(e) => setFormMultiplierMid(e.target.value)}
                      placeholder="0.80"
                      required
                    />
                  </label>
                  <label>
                    High Multiplier (≥ High)
                    <input
                      name="multiplierHigh"
                      inputMode="decimal"
                      value={formMultiplierHigh}
                      onChange={(e) => setFormMultiplierHigh(e.target.value)}
                      placeholder="1.00"
                      required
                    />
                  </label>
                </div>

                {/* VISUAL PREVIEW OF THE 3 TIERS */}
                <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10, marginTop: 4 }}>
                  <div style={{ background: "#fff", border: "1px solid #fde68a", borderRadius: 6, padding: "10px", textAlign: "center" }}>
                    <span style={{ fontSize: "10px", textTransform: "uppercase", fontWeight: 700, color: "#b45309", display: "block" }}>
                      Tier 1: Under-Hour
                    </span>
                    <strong style={{ fontSize: "14px", color: "#b45309" }}>&lt; {formThresholdLow || 0} hrs</strong>
                    <div style={{ fontSize: "12px", color: "#b45309", fontWeight: 650, marginTop: 2 }}>
                      {formMultiplierLow || 0}× ({Math.round(Number(formMultiplierLow || 0) * 100)}% Payout)
                    </div>
                  </div>
                  <div style={{ background: "#fff", border: "1px solid #bfdbfe", borderRadius: 6, padding: "10px", textAlign: "center" }}>
                    <span style={{ fontSize: "10px", textTransform: "uppercase", fontWeight: 700, color: "#2563eb", display: "block" }}>
                      Tier 2: Intermediate
                    </span>
                    <strong style={{ fontSize: "14px", color: "#2563eb" }}>
                      {formThresholdLow || 0} to {formThresholdHigh || 0} hrs
                    </strong>
                    <div style={{ fontSize: "12px", color: "#2563eb", fontWeight: 650, marginTop: 2 }}>
                      {formMultiplierMid || 0}× ({Math.round(Number(formMultiplierMid || 0) * 100)}% Payout)
                    </div>
                  </div>
                  <div style={{ background: "#fff", border: "1px solid #a7f3d0", borderRadius: 6, padding: "10px", textAlign: "center" }}>
                    <span style={{ fontSize: "10px", textTransform: "uppercase", fontWeight: 700, color: "#059669", display: "block" }}>
                      Tier 3: Full Standard
                    </span>
                    <strong style={{ fontSize: "14px", color: "#059669" }}>≥ {formThresholdHigh || 0} hrs</strong>
                    <div style={{ fontSize: "12px", color: "#059669", fontWeight: 650, marginTop: 2 }}>
                      {formMultiplierHigh || 0}× ({Math.round(Number(formMultiplierHigh || 0) * 100)}% Payout)
                    </div>
                  </div>
                </div>
              </div>
            )}

            <div className="form-actions full" style={{ marginTop: 8 }}>
              <button
                type="button"
                className="ghost-button"
                onClick={() => {
                  setCreateModal(false);
                  setClonedFromRule(null);
                }}
              >
                Cancel
              </button>
              <button className="primary-button compact" disabled={createMutation.isPending}>
                {createMutation.isPending
                  ? "Saving Version…"
                  : clonedFromRule
                  ? "Create Cloned Rule Version"
                  : "Save Immutable Version"}
              </button>
            </div>
          </form>
        </Modal>
      )}

      {/* EDIT RATE RULE MODAL */}
      {editRule && (
        <Modal title={`Edit Rate Rule: ${editRule.ruleName}`} onClose={() => setEditRule(null)} wide>
          <form
            className="form-grid"
            onSubmit={(event) => {
              event.preventDefault();
              const f = new FormData(event.currentTarget);
              const from = field(f, "effectiveFrom");
              const to = field(f, "effectiveTo") || null;
              if (to && to < from) {
                notify({ kind: "error", text: "Effective To date must be on or after Effective From date." });
                return;
              }

              const rType = field(f, "ruleType") as "unit-rate" | "slab" | "fixed" | "hybrid";
              if (rType === "slab") {
                const lowT = field(f, "thresholdLow");
                const highT = field(f, "thresholdHigh");
                const lowM = field(f, "multiplierLow");
                const midM = field(f, "multiplierMid");
                const highM = field(f, "multiplierHigh");
                if (!lowT || !highT || !lowM || !midM || !highM) {
                  notify({ kind: "error", text: "Slab rule requires both thresholds and all three multipliers." });
                  return;
                }
              }

              updateMutation.mutate({
                token,
                id: editRule.id,
                ruleName: field(f, "ruleName"),
                ruleType: rType,
                unitRate: field(f, "unitRate"),
                thresholdLow: field(f, "thresholdLow") || null,
                thresholdHigh: field(f, "thresholdHigh") || null,
                multiplierLow: field(f, "multiplierLow") || null,
                multiplierMid: field(f, "multiplierMid") || null,
                multiplierHigh: field(f, "multiplierHigh") || null,
                priority: Number(field(f, "priority")),
                effectiveFrom: from,
                effectiveTo: to,
                status: field(f, "status") as "active" | "inactive",
                categoryId: field(f, "categoryId") ? Number(field(f, "categoryId")) : null,
              });
            }}
          >
            <label className="full">
              Rule Name
              <input name="ruleName" defaultValue={editRule.ruleName} required />
            </label>

            <label>
              Country Scope (Fixed)
              <input value={`${editRule.countryCode}`} disabled style={{ background: "var(--surface-alt)" }} />
            </label>

            <label>
              Work Type (Fixed)
              <input value={`${editRule.workTypeName}`} disabled style={{ background: "var(--surface-alt)" }} />
            </label>

            <label>
              Character Category
              <select name="categoryId" defaultValue={editRule.categoryId ?? ""}>
                <option value="">All Categories (General basis)</option>
                {activeCategories
                  .filter((c) => c.status === "active")
                  .map((c) => (
                    <option key={c.id} value={c.id}>
                      {c.name} ({c.code})
                    </option>
                  ))}
              </select>
            </label>

            <label>
              Calculation Model
              <select
                name="ruleType"
                value={editRuleType}
                onChange={(e) => setEditRuleType(e.target.value as any)}
                required
              >
                <option value="unit-rate">Unit Rate (Linear)</option>
                <option value="slab">Slab-Based (Hour Thresholds)</option>
                <option value="fixed">Fixed Lump-Sum</option>
                <option value="hybrid">Hybrid</option>
              </select>
            </label>

            <label>
              Base Unit Rate
              <input name="unitRate" inputMode="decimal" defaultValue={editRule.unitRate} required />
            </label>

            <label>
              Current Status
              <select name="status" defaultValue={editRule.status} required>
                <option value="active">Active (Included in calculations)</option>
                <option value="inactive">Inactive (Archived)</option>
              </select>
            </label>

            <label>
              Priority (1-999)
              <input name="priority" type="number" min="1" max="999" defaultValue={editRule.priority} required />
            </label>

            <label>
              Effective From
              <input name="effectiveFrom" type="date" defaultValue={editRule.effectiveFrom} required />
            </label>

            <label>
              Effective To (Optional)
              <input name="effectiveTo" type="date" defaultValue={editRule.effectiveTo ?? ""} />
            </label>

            {(editRuleType === "slab" || editRuleType === "hybrid") && (
              <div
                className="full"
                style={{
                  background: "var(--surface-alt)",
                  border: "1px solid var(--border)",
                  borderRadius: 8,
                  padding: "16px",
                  display: "flex",
                  flexDirection: "column",
                  gap: 12,
                }}
              >
                <strong style={{ fontSize: "13px", color: "var(--accent-deep)" }}>
                  Slab Thresholds & Multipliers
                </strong>
                <div className="slab-fields" style={{ margin: 0 }}>
                  <label>
                    Low Hour Threshold
                    <input name="thresholdLow" inputMode="decimal" defaultValue={editRule.thresholdLow ?? "3.99"} required />
                  </label>
                  <label>
                    High Hour Threshold
                    <input name="thresholdHigh" inputMode="decimal" defaultValue={editRule.thresholdHigh ?? "5.99"} required />
                  </label>
                  <label>
                    Low Multiplier (&lt; Low)
                    <input name="multiplierLow" inputMode="decimal" defaultValue={editRule.multiplierLow ?? "0.60"} required />
                  </label>
                  <label>
                    Mid Multiplier (Between)
                    <input name="multiplierMid" inputMode="decimal" defaultValue={editRule.multiplierMid ?? "0.80"} required />
                  </label>
                  <label>
                    High Multiplier (≥ High)
                    <input name="multiplierHigh" inputMode="decimal" defaultValue={editRule.multiplierHigh ?? "1.00"} required />
                  </label>
                </div>
              </div>
            )}

            <div className="form-actions full" style={{ marginTop: 8 }}>
              <button type="button" className="ghost-button" onClick={() => setEditRule(null)}>
                Cancel
              </button>
              <button className="primary-button compact" disabled={updateMutation.isPending}>
                {updateMutation.isPending ? "Saving Changes…" : "Save Changes"}
              </button>
            </div>
          </form>
        </Modal>
      )}

      {/* VIEW RATE RULE MODAL WITH INTERACTIVE SANDBOX */}
      {viewRule && (
        <Modal title={`Rate Rule Details: ${viewRule.ruleName}`} onClose={() => setViewRule(null)} wide>
          <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
            {/* Header info */}
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                padding: "16px 20px",
                background: "var(--surface-alt)",
                border: "1px solid var(--border)",
                borderRadius: 8,
                flexWrap: "wrap",
                gap: 12,
              }}
            >
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span
                    className="mono"
                    style={{
                      fontWeight: 800,
                      fontSize: "13px",
                      color: "#073b33",
                      background: "#f0f8f5",
                      border: "1px solid #99c2b6",
                      padding: "2px 8px",
                      borderRadius: 4,
                    }}
                  >
                    {viewRule.countryCode}
                  </span>
                  <strong style={{ fontSize: "16px", color: "var(--text)" }}>{viewRule.ruleName}</strong>
                </div>
                <p style={{ margin: "4px 0 0", color: "var(--dim)", fontSize: "12px" }}>
                  {viewRule.workTypeName} {viewRule.categoryName ? `· Category: ${viewRule.categoryName}` : "· All Categories"}
                </p>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <Status value={viewRule.status} />
                <span
                  className="mono"
                  style={{
                    fontWeight: 700,
                    padding: "3px 8px",
                    background: "var(--accent-pale)",
                    color: "var(--accent-deep)",
                    borderRadius: 4,
                    fontSize: "12px",
                  }}
                >
                  Priority P{viewRule.priority}
                </span>
              </div>
            </div>

            {/* Metadata Grid */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 10 }}>
              <div style={{ padding: "10px 12px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--surface)" }}>
                <span style={{ fontSize: "11px", color: "var(--dim)", textTransform: "uppercase", fontWeight: 700 }}>Rule ID</span>
                <p style={{ margin: "3px 0 0", fontWeight: 700 }} className="mono">#{viewRule.id}</p>
              </div>
              <div style={{ padding: "10px 12px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--surface)" }}>
                <span style={{ fontSize: "11px", color: "var(--dim)", textTransform: "uppercase", fontWeight: 700 }}>Calculation Model</span>
                <p style={{ margin: "3px 0 0", fontWeight: 700 }} className="mono">{viewRule.ruleType.toUpperCase()}</p>
              </div>
              <div style={{ padding: "10px 12px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--surface)" }}>
                <span style={{ fontSize: "11px", color: "var(--dim)", textTransform: "uppercase", fontWeight: 700 }}>Base Unit Rate</span>
                <p style={{ margin: "3px 0 0", fontWeight: 750, color: "#073b33" }} className="mono">{viewRule.unitRate}</p>
              </div>
              <div style={{ padding: "10px 12px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--surface)" }}>
                <span style={{ fontSize: "11px", color: "var(--dim)", textTransform: "uppercase", fontWeight: 700 }}>Effective Window</span>
                <p style={{ margin: "3px 0 0", fontWeight: 650 }}>
                  {formatDate(viewRule.effectiveFrom)} {viewRule.effectiveTo ? `to ${formatDate(viewRule.effectiveTo)}` : "(Open-ended)"}
                </p>
              </div>
              <div style={{ padding: "10px 12px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--surface)" }}>
                <span style={{ fontSize: "11px", color: "var(--dim)", textTransform: "uppercase", fontWeight: 700 }}>Supersedes</span>
                <p style={{ margin: "3px 0 0", fontWeight: 650 }}>
                  {viewRule.supersedesRateRuleId ? `Rule #${viewRule.supersedesRateRuleId}` : "None"}
                </p>
              </div>
              <div style={{ padding: "10px 12px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--surface)" }}>
                <span style={{ fontSize: "11px", color: "var(--dim)", textTransform: "uppercase", fontWeight: 700 }}>Created Timestamp</span>
                <p style={{ margin: "3px 0 0", fontWeight: 650 }}>{formatDate(viewRule.createdAt)}</p>
              </div>
            </div>

            {/* Slab breakdown card if slab */}
            {(viewRule.ruleType === "slab" || viewRule.ruleType === "hybrid") && (
              <div style={{ border: "1px solid var(--border)", borderRadius: 8, padding: "14px 16px", background: "var(--surface-alt)" }}>
                <strong style={{ fontSize: "13px", color: "var(--accent-deep)", display: "block", marginBottom: 8 }}>
                  Configured Productivity Hour Tiers
                </strong>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10 }}>
                  <div style={{ background: "#fff", border: "1px solid #fde68a", borderRadius: 6, padding: "10px" }}>
                    <span style={{ fontSize: "10.5px", textTransform: "uppercase", fontWeight: 700, color: "#b45309" }}>Tier 1 (&lt; Low)</span>
                    <strong style={{ display: "block", fontSize: "14px", color: "#b45309", marginTop: 2 }}>
                      &lt; {viewRule.thresholdLow} Hours
                    </strong>
                    <span style={{ fontSize: "12px", color: "#b45309", fontWeight: 700 }}>
                      Multiplier: {viewRule.multiplierLow}×
                    </span>
                  </div>
                  <div style={{ background: "#fff", border: "1px solid #bfdbfe", borderRadius: 6, padding: "10px" }}>
                    <span style={{ fontSize: "10.5px", textTransform: "uppercase", fontWeight: 700, color: "#2563eb" }}>Tier 2 (Mid)</span>
                    <strong style={{ display: "block", fontSize: "14px", color: "#2563eb", marginTop: 2 }}>
                      {viewRule.thresholdLow} to {viewRule.thresholdHigh} Hours
                    </strong>
                    <span style={{ fontSize: "12px", color: "#2563eb", fontWeight: 700 }}>
                      Multiplier: {viewRule.multiplierMid}×
                    </span>
                  </div>
                  <div style={{ background: "#fff", border: "1px solid #a7f3d0", borderRadius: 6, padding: "10px" }}>
                    <span style={{ fontSize: "10.5px", textTransform: "uppercase", fontWeight: 700, color: "#059669" }}>Tier 3 (High)</span>
                    <strong style={{ display: "block", fontSize: "14px", color: "#059669", marginTop: 2 }}>
                      ≥ {viewRule.thresholdHigh} Hours
                    </strong>
                    <span style={{ fontSize: "12px", color: "#059669", fontWeight: 700 }}>
                      Multiplier: {viewRule.multiplierHigh}×
                    </span>
                  </div>
                </div>
              </div>
            )}

            {/* LIVE INTERACTIVE MINI-PAYOUT SANDBOX */}
            <div
              style={{
                border: "1px solid #99c2b6",
                borderRadius: 8,
                padding: "16px",
                background: "#f0f8f5",
                display: "flex",
                flexDirection: "column",
                gap: 12,
              }}
            >
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <strong style={{ fontSize: "13.5px", color: "#073b33" }}>Live Payout Sandbox for this Rule</strong>
                  <span style={{ display: "block", fontSize: "11.5px", color: "var(--dim)" }}>
                    Test how quantity and average hours translate into final compensation with this exact rule formula
                  </span>
                </div>
                <span className="mono" style={{ fontSize: "11px", color: "var(--accent-deep)", fontWeight: 700 }}>
                  Live Kernel Eval
                </span>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr auto", gap: 12, alignItems: "flex-end" }}>
                <label style={{ margin: 0 }}>
                  <span style={{ fontSize: "11.5px", fontWeight: 650, color: "var(--text)", display: "block", marginBottom: 4 }}>
                    Sample Quantity (units/characters)
                  </span>
                  <input
                    type="number"
                    value={viewSimQty}
                    onChange={(e) => setViewSimQty(e.target.value)}
                    placeholder="1000"
                    style={{ background: "#fff" }}
                  />
                </label>
                <label style={{ margin: 0 }}>
                  <span style={{ fontSize: "11.5px", fontWeight: 650, color: "var(--text)", display: "block", marginBottom: 4 }}>
                    Average Daily Working Hours
                  </span>
                  <input
                    type="number"
                    step="0.1"
                    value={viewSimHours}
                    onChange={(e) => setViewSimHours(e.target.value)}
                    placeholder="6.0"
                    style={{ background: "#fff" }}
                  />
                </label>
                {/* Result box */}
                {(() => {
                  const qty = Number(viewSimQty) || 0;
                  const rate = Number(viewRule.unitRate) || 0;
                  const hours = Number(viewSimHours) || 0;
                  let multiplier = 1.0;
                  let tierLabel = "Fixed Multiplier";

                  if (viewRule.ruleType === "slab" || viewRule.ruleType === "hybrid") {
                    const low = Number(viewRule.thresholdLow || 3.99);
                    const high = Number(viewRule.thresholdHigh || 5.99);
                    const mLow = Number(viewRule.multiplierLow || 0.60);
                    const mMid = Number(viewRule.multiplierMid || 0.80);
                    const mHigh = Number(viewRule.multiplierHigh || 1.00);

                    if (hours < low) {
                      multiplier = mLow;
                      tierLabel = `Tier 1 (< ${low}h)`;
                    } else if (hours < high) {
                      multiplier = mMid;
                      tierLabel = `Tier 2 (${low}-${high}h)`;
                    } else {
                      multiplier = mHigh;
                      tierLabel = `Tier 3 (≥ ${high}h)`;
                    }
                  }

                  const baseAmount = viewRule.ruleType === "fixed" ? rate : qty * rate;
                  const payable = baseAmount * multiplier;

                  return (
                    <div
                      style={{
                        padding: "8px 14px",
                        background: "#fff",
                        border: "1px solid #99c2b6",
                        borderRadius: 6,
                        display: "flex",
                        gap: 16,
                        alignItems: "center",
                      }}
                    >
                      <div>
                        <span style={{ fontSize: "10.5px", color: "var(--dim)", textTransform: "uppercase", display: "block" }}>
                          Base Amount
                        </span>
                        <strong className="mono" style={{ fontSize: "13px" }}>
                          ${baseAmount.toFixed(2)}
                        </strong>
                      </div>
                      <div>
                        <span style={{ fontSize: "10.5px", color: "var(--dim)", textTransform: "uppercase", display: "block" }}>
                          Factor ({tierLabel})
                        </span>
                        <strong className="mono" style={{ fontSize: "13px", color: "#2563eb" }}>
                          {multiplier.toFixed(2)}×
                        </strong>
                      </div>
                      <div>
                        <span style={{ fontSize: "10.5px", color: "var(--dim)", textTransform: "uppercase", display: "block" }}>
                          Calculated Payout
                        </span>
                        <strong className="mono" style={{ fontSize: "16px", color: "#073b33" }}>
                          ${payable.toFixed(2)}
                        </strong>
                      </div>
                    </div>
                  );
                })()}
              </div>
            </div>

            {/* Action buttons */}
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, borderTop: "1px solid var(--border)", paddingTop: 14, flexWrap: "wrap" }}>
              <button
                type="button"
                className="secondary-button"
                onClick={() => {
                  const id = viewRule.id;
                  setViewRule(null);
                  setHistoryRuleId(id);
                }}
                title="Inspect all previous superseded versions of this rate rule"
              >
                <Icon name="history" />
                <span>Version History</span>
              </button>
              {canManage && (
                <button
                  type="button"
                  className="secondary-button"
                  onClick={() => {
                    const r = viewRule;
                    setViewRule(null);
                    openEditModal(r);
                  }}
                >
                  <Icon name="pencil" />
                  <span>Edit Rule</span>
                </button>
              )}
              {canManage && (
                <button
                  type="button"
                  className="danger-button"
                  onClick={() => {
                    const r = viewRule;
                    setViewRule(null);
                    setDeleteTarget(r);
                  }}
                >
                  <Icon name="trash" />
                  <span>Move to Recycle Bin</span>
                </button>
              )}
              <button type="button" className="ghost-button" onClick={() => setViewRule(null)}>
                Close
              </button>
            </div>
          </div>
        </Modal>
      )}

      {/* DELETE CONFIRMATION DIALOG */}
      <ConfirmDialog
        isOpen={Boolean(deleteTarget)}
        title="Move Rate Rule to Recycle Bin?"
        message={`Are you sure you want to delete rate rule "${deleteTarget?.ruleName}" (${deleteTarget?.countryCode} · ${deleteTarget?.unitRate})? The rule will be safely backed up to the Recycle Bin and can be restored at any time.`}
        confirmText="Move to Recycle Bin"
        danger
        isLoading={removeMutation.isPending}
        onConfirm={() => {
          if (deleteTarget) {
            removeMutation.mutate({ token, id: deleteTarget.id });
          }
        }}
        onCancel={() => setDeleteTarget(null)}
      />

      {/* VERSION HISTORY MODAL */}
      {historyRuleId && (
        <Modal
          title={`Rate Rule Version History (Rule #${historyRuleId})`}
          onClose={() => setHistoryRuleId(null)}
          wide
        >
          {historyQuery.isPending ? (
            <Loading />
          ) : historyQuery.error || !historyQuery.data ? (
            <Empty title="Version history could not load" body={message(historyQuery.error)} />
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
              {/* Lineage Overview Card */}
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  padding: "16px 20px",
                  background: "var(--surface-alt)",
                  border: "1px solid var(--border)",
                  borderRadius: 8,
                  flexWrap: "wrap",
                  gap: 12,
                }}
              >
                <div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                    <span
                      className="mono"
                      style={{
                        fontWeight: 800,
                        fontSize: "13px",
                        color: "#073b33",
                        background: "#f0f8f5",
                        border: "1px solid #99c2b6",
                        padding: "2px 8px",
                        borderRadius: 4,
                      }}
                    >
                      {historyQuery.data.versions[0]?.countryCode || "SCOPE"}
                    </span>
                    <strong style={{ fontSize: "15px", color: "var(--text)" }}>
                      {historyQuery.data.versions[historyQuery.data.versions.length - 1]?.ruleName || "Rule"}
                    </strong>
                  </div>
                  <p style={{ margin: "4px 0 0", color: "var(--dim)", fontSize: "12px" }}>
                    {historyQuery.data.versions[0]?.workTypeName}
                    {historyQuery.data.versions[0]?.categoryName
                      ? ` · Category: ${historyQuery.data.versions[0]?.categoryName}`
                      : " · All Categories"}
                  </p>
                </div>

                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span
                    className="mono"
                    style={{
                      fontWeight: 750,
                      padding: "4px 10px",
                      background: "var(--accent-pale)",
                      color: "var(--accent-deep)",
                      borderRadius: 4,
                      fontSize: "12px",
                    }}
                  >
                    {historyQuery.data.versions.length} Version{historyQuery.data.versions.length === 1 ? "" : "s"} in Lineage
                  </span>
                </div>
              </div>

              {/* Version History Timeline Cards */}
              <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                {historyQuery.data.versions.map((v, idx) => {
                  const isSelected = v.id === historyRuleId;
                  const isCurrentActive = v.status === "active";
                  const prevVersion = idx > 0 ? historyQuery.data!.versions[idx - 1] : null;

                  const currentRateNum = parseFloat(v.unitRate);
                  const prevRateNum = prevVersion ? parseFloat(prevVersion.unitRate) : null;
                  const diffRate = prevRateNum !== null ? currentRateNum - prevRateNum : null;
                  const diffPercent =
                    prevRateNum !== null && prevRateNum !== 0
                      ? ((diffRate! / prevRateNum) * 100).toFixed(1)
                      : null;

                  return (
                    <div
                      key={v.id}
                      style={{
                        position: "relative",
                        border: isSelected ? "2px solid var(--accent-deep)" : "1px solid var(--border)",
                        borderRadius: 8,
                        background: isSelected ? "rgba(7, 59, 51, 0.02)" : "var(--surface)",
                        boxShadow: isSelected ? "0 4px 14px rgba(7, 59, 51, 0.08)" : "none",
                        padding: "16px 20px",
                        display: "flex",
                        flexDirection: "column",
                        gap: 12,
                      }}
                    >
                      {/* Top Version bar */}
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                          <span
                            className="mono"
                            style={{
                              fontSize: "12px",
                              fontWeight: 800,
                              padding: "2px 8px",
                              borderRadius: 4,
                              background: isCurrentActive ? "#d1fae5" : "#f3f4f6",
                              color: isCurrentActive ? "#065f46" : "#4b5563",
                              border: `1px solid ${isCurrentActive ? "#a7f3d0" : "#d1d5db"}`,
                            }}
                          >
                            Version {v.versionIndex} {idx === 0 ? "(Initial)" : isCurrentActive ? "(Current Active)" : "(Superseded)"}
                          </span>
                          <span className="mono" style={{ fontSize: "12px", color: "var(--dim)" }}>
                            Rule #{v.id}
                          </span>
                          {isSelected && (
                            <span
                              style={{
                                fontSize: "11px",
                                fontWeight: 700,
                                color: "var(--accent-deep)",
                                background: "var(--accent-pale)",
                                padding: "1px 6px",
                                borderRadius: 3,
                              }}
                            >
                              Selected
                            </span>
                          )}
                        </div>

                        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          <Status value={v.status} />
                          <span
                            className="mono"
                            style={{
                              fontSize: "11px",
                              fontWeight: 650,
                              background: "var(--surface-alt)",
                              border: "1px solid var(--border)",
                              padding: "2px 6px",
                              borderRadius: 4,
                            }}
                          >
                            P{v.priority}
                          </span>
                        </div>
                      </div>

                      {/* Rule Name & Supersession pointers */}
                      <div>
                        <strong style={{ fontSize: "14px", color: "var(--text)" }}>{v.ruleName}</strong>
                        <div style={{ display: "flex", gap: 14, flexWrap: "wrap", fontSize: "11.5px", color: "var(--dim)", marginTop: 4 }}>
                          {v.supersedesRateRuleId ? (
                            <span>
                              ↳ Supersedes <b>Rule #{v.supersedesRateRuleId}</b>
                            </span>
                          ) : (
                            <span>↳ Baseline version (Initial rule definition)</span>
                          )}
                          {v.supersededByRateRuleId ? (
                            <span style={{ color: "#b45309" }}>
                              ↳ Superseded by <b>Rule #{v.supersededByRateRuleId}</b>
                            </span>
                          ) : isCurrentActive ? (
                            <span style={{ color: "#059669", fontWeight: 650 }}>
                              ↳ Current Active Production Rule
                            </span>
                          ) : null}
                        </div>
                      </div>

                      {/* Grid of Key Properties: Rates, Dates, Model */}
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 10 }}>
                        <div style={{ padding: "8px 12px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--surface-alt)" }}>
                          <span style={{ fontSize: "10.5px", color: "var(--dim)", textTransform: "uppercase", fontWeight: 700, display: "block" }}>
                            Unit Rate
                          </span>
                          <div style={{ display: "flex", alignItems: "baseline", gap: 6, marginTop: 2 }}>
                            <strong className="mono" style={{ fontSize: "14px", color: "#073b33" }}>
                              {v.unitRate}
                            </strong>
                            {diffRate !== null && diffRate !== 0 && (
                              <span
                                className="mono"
                                style={{
                                  fontSize: "11px",
                                  fontWeight: 700,
                                  color: diffRate > 0 ? "#059669" : "#dc2626",
                                }}
                              >
                                {diffRate > 0 ? `+${diffRate.toFixed(4)}` : diffRate.toFixed(4)} ({diffPercent}%)
                              </span>
                            )}
                          </div>
                        </div>

                        <div style={{ padding: "8px 12px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--surface-alt)" }}>
                          <span style={{ fontSize: "10.5px", color: "var(--dim)", textTransform: "uppercase", fontWeight: 700, display: "block" }}>
                            Calculation Model
                          </span>
                          <span className="mono" style={{ fontSize: "12.5px", fontWeight: 700, color: "var(--text)", textTransform: "uppercase", marginTop: 2, display: "block" }}>
                            {v.ruleType}
                          </span>
                        </div>

                        <div style={{ padding: "8px 12px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--surface-alt)" }}>
                          <span style={{ fontSize: "10.5px", color: "var(--dim)", textTransform: "uppercase", fontWeight: 700, display: "block" }}>
                            Effective Window
                          </span>
                          <span style={{ fontSize: "12px", fontWeight: 650, marginTop: 2, display: "block" }}>
                            {formatDate(v.effectiveFrom)} {v.effectiveTo ? `to ${formatDate(v.effectiveTo)}` : "(Open-ended)"}
                          </span>
                        </div>

                        <div style={{ padding: "8px 12px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--surface-alt)" }}>
                          <span style={{ fontSize: "10.5px", color: "var(--dim)", textTransform: "uppercase", fontWeight: 700, display: "block" }}>
                            Created / Modified
                          </span>
                          <span style={{ fontSize: "11.5px", color: "var(--dim)", marginTop: 2, display: "block" }}>
                            {formatDateTime(v.createdAt)}
                          </span>
                          <small style={{ fontSize: "10.5px", color: "var(--dim)" }}>by {v.createdByName}</small>
                        </div>
                      </div>

                      {/* Slab Details if applicable */}
                      {(v.ruleType === "slab" || v.ruleType === "hybrid") && (
                        <div
                          style={{
                            background: "#fff",
                            border: "1px solid var(--border)",
                            borderRadius: 6,
                            padding: "8px 12px",
                            display: "flex",
                            alignItems: "center",
                            gap: 12,
                            flexWrap: "wrap",
                            fontSize: "11px",
                          }}
                        >
                          <span style={{ fontWeight: 700, color: "var(--dim)", textTransform: "uppercase", fontSize: "10.5px" }}>
                            Hour Slabs:
                          </span>
                          <span className="mono" style={{ color: "#b45309", background: "#fef3c7", padding: "1px 6px", borderRadius: 3 }}>
                            &lt;{v.thresholdLow}h: {v.multiplierLow}×
                          </span>
                          <span className="mono" style={{ color: "#2563eb", background: "#dbeafe", padding: "1px 6px", borderRadius: 3 }}>
                            {v.thresholdLow}-{v.thresholdHigh}h: {v.multiplierMid}×
                          </span>
                          <span className="mono" style={{ color: "#059669", background: "#d1fae5", padding: "1px 6px", borderRadius: 3 }}>
                            ≥{v.thresholdHigh}h: {v.multiplierHigh}×
                          </span>
                        </div>
                      )}

                      {/* Audit Entries Trail */}
                      {v.auditEntries.length > 0 && (
                        <div style={{ borderTop: "1px dashed var(--border)", paddingTop: 8 }}>
                          <span style={{ fontSize: "10.5px", textTransform: "uppercase", fontWeight: 700, color: "var(--dim)", display: "block", marginBottom: 4 }}>
                            Modification Log ({v.auditEntries.length} event{v.auditEntries.length === 1 ? "" : "s"}):
                          </span>
                          <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                            {v.auditEntries.map((a, aIdx) => (
                              <div key={aIdx} style={{ fontSize: "11px", color: "var(--dim)", display: "flex", gap: 8, alignItems: "center" }}>
                                <span style={{ fontWeight: 650, color: "var(--text)" }}>{humanize(a.action)}</span>
                                <span>by {a.actor}</span>
                                <span>· {formatDateTime(a.timestamp)}</span>
                                {a.reason && <span style={{ fontStyle: "italic" }}>— &ldquo;{a.reason}&rdquo;</span>}
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              {/* Close Button */}
              <div style={{ display: "flex", justifyContent: "flex-end", borderTop: "1px solid var(--border)", paddingTop: 14 }}>
                <button type="button" className="ghost-button" onClick={() => setHistoryRuleId(null)}>
                  Close
                </button>
              </div>
            </div>
          )}
        </Modal>
      )}

      {/* KERNEL SIMULATOR MODAL */}
      {simulatorOpen && (
        <Modal title="Calculation Kernel Simulator" onClose={() => setSimulatorOpen(null as any)} wide>
          <form
            className="stack-form"
            onSubmit={(event) => {
              event.preventDefault();
              const f = new FormData(event.currentTarget);
              previewMutation.mutate({
                token,
                quantityRates: [{ quantity: field(f, "quantity"), rate: field(f, "rate") }],
                averageHours: field(f, "averageHours"),
                thresholdLow: field(f, "thresholdLow"),
                thresholdHigh: field(f, "thresholdHigh"),
                multiplierLow: field(f, "multiplierLow"),
                multiplierMid: field(f, "multiplierMid"),
                multiplierHigh: field(f, "multiplierHigh"),
              });
            }}
          >
            <p className="form-lead" style={{ margin: 0 }}>
              The server-side kernel accepts normalized quantity-rate pairs and evaluates the average working hours against
              the configured productivity slab thresholds.
            </p>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
              <label>
                Quantity Units
                <input name="quantity" inputMode="decimal" defaultValue={simInitial?.quantity || "1000"} required />
              </label>
              <label>
                Base Unit Rate
                <input name="rate" inputMode="decimal" defaultValue={simInitial?.rate || "0.0225"} required />
              </label>
              <label>
                Average Working Hours
                <input name="averageHours" inputMode="decimal" defaultValue={simInitial?.averageHours || "6.0"} required />
              </label>
            </div>

            <div className="mini-grid">
              <label>
                Low Threshold (hrs)
                <input name="thresholdLow" defaultValue={simInitial?.thresholdLow || "3.99"} required />
              </label>
              <label>
                High Threshold (hrs)
                <input name="thresholdHigh" defaultValue={simInitial?.thresholdHigh || "5.99"} required />
              </label>
              <label>
                Low Multiplier
                <input name="multiplierLow" defaultValue={simInitial?.multiplierLow || "0.60"} required />
              </label>
              <label>
                Mid Multiplier
                <input name="multiplierMid" defaultValue={simInitial?.multiplierMid || "0.80"} required />
              </label>
              <label>
                High Multiplier
                <input name="multiplierHigh" defaultValue={simInitial?.multiplierHigh || "1.00"} required />
              </label>
            </div>

            <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
              <button className="primary-button compact" disabled={previewMutation.isPending}>
                {previewMutation.isPending ? "Evaluating Kernel…" : "Run Kernel Evaluation"}
              </button>
              <button type="button" className="ghost-button" onClick={() => setSimulatorOpen(false)}>
                Close
              </button>
            </div>

            {previewResult && (
              <div className="preview-result" style={{ marginTop: 14 }}>
                <span>
                  Base Amount <b className="mono">${previewResult.baseAmount}</b>
                </span>
                <span>
                  Applied Multiplier <b className="mono">{previewResult.multiplier}×</b>
                </span>
                <span>
                  Payable Amount <b className="mono">${previewResult.payableAmount}</b>
                </span>
              </div>
            )}
          </form>
        </Modal>
      )}

      {/* CATEGORY RATES MATRIX SETUP MODAL */}
      {categoryModalOpen && projectId && (
        <Modal
          title="Set Category Rates (Header / Birth / Marriage / Death Character)"
          onClose={() => setCategoryModalOpen(false)}
          wide
        >
          <form
            className="stack-form"
            onSubmit={(e) => {
              e.preventDefault();
              if (!categoryCountryCode || !categoryWorkTypeId) {
                notify({ kind: "error", text: "Please select both Country and Work Type." });
                return;
              }
              const ratesToSave: { categoryId: number; unitRate: string; ruleType: "unit-rate" | "slab" | "fixed" | "hybrid"; ruleName?: string }[] = [];
              for (const [catIdStr, rateVal] of Object.entries(categoryDraftRates)) {
                if (rateVal && rateVal.trim() !== "") {
                  const catId = Number(catIdStr);
                  const cat = activeCategories.find((c) => c.id === catId);
                  ratesToSave.push({
                    categoryId: catId,
                    unitRate: rateVal.trim(),
                    ruleType: categoryRuleType,
                    ruleName: `${categoryCountryCode} - ${cat?.name || `Category ${catId}`} Rate`,
                  });
                }
              }
              if (ratesToSave.length === 0) {
                notify({ kind: "error", text: "Please enter at least one category rate." });
                return;
              }
              saveCategoryMutation.mutate({
                token,
                projectId,
                countryCode: categoryCountryCode,
                workTypeId: Number(categoryWorkTypeId),
                effectiveFrom: categoryEffectiveFrom,
                effectiveTo: null,
                rates: ratesToSave,
              });
            }}
          >
            <p className="form-lead" style={{ margin: 0, fontSize: "12.5px" }}>
              Configure specific unit rates for character & record categories (e.g. <b>Header Character</b>, <b>Birth Character</b>, <b>Marriage Character</b>, <b>Date Character</b>, <b>Review Key</b>, <b>Header</b>). Category rates take precedence over the generic base rate for matching records during salary calculation.
            </p>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
              <label>
                Country Scope
                <select
                  value={categoryCountryCode}
                  onChange={(e) => handleCategoryScopeChange(e.target.value, categoryWorkTypeId)}
                  required
                >
                  {activeCountries.filter((c) => c.status === "active").map((c) => (
                    <option key={c.id} value={c.countryCode}>
                      {c.countryCode} · {c.countryName}
                    </option>
                  ))}
                  {!activeCountries.some((c) => c.countryCode === "BD") && <option value="BD">BD · Bangladesh</option>}
                  {!activeCountries.some((c) => c.countryCode === "Kolkata") && <option value="Kolkata">Kolkata</option>}
                  {!activeCountries.some((c) => c.countryCode === "KTM") && <option value="KTM">KTM · Kathmandu</option>}
                </select>
              </label>

              <label>
                Work Type
                <select
                  value={categoryWorkTypeId}
                  onChange={(e) => handleCategoryScopeChange(categoryCountryCode, e.target.value)}
                  required
                >
                  {activeWorkTypes.filter((w) => w.status === "active").map((w) => (
                    <option key={w.id} value={w.id}>
                      {w.name} ({w.paymentBasis} · {w.unit})
                    </option>
                  ))}
                </select>
              </label>

              <label>
                Effective From
                <input
                  type="date"
                  value={categoryEffectiveFrom}
                  onChange={(e) => setCategoryEffectiveFrom(e.target.value)}
                  required
                />
              </label>
            </div>

            <div style={{ display: "flex", alignItems: "center", gap: 14, margin: "2px 0" }}>
              <span style={{ fontSize: "12px", fontWeight: 650, color: "var(--dim)" }}>Rate Calculation Model:</span>
              <label style={{ display: "inline-flex", alignItems: "center", gap: 5, cursor: "pointer", fontSize: "12px" }}>
                <input
                  type="radio"
                  name="categoryRuleType"
                  checked={categoryRuleType === "unit-rate"}
                  onChange={() => setCategoryRuleType("unit-rate")}
                />
                Linear Unit Rate
              </label>
              <label style={{ display: "inline-flex", alignItems: "center", gap: 5, cursor: "pointer", fontSize: "12px" }}>
                <input
                  type="radio"
                  name="categoryRuleType"
                  checked={categoryRuleType === "slab"}
                  onChange={() => setCategoryRuleType("slab")}
                />
                Slab Tier Rate
              </label>
            </div>

            {/* Matrix Table */}
            <div style={{ border: "1px solid var(--border)", borderRadius: 6, overflow: "hidden", maxHeight: "360px", overflowY: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
                <thead>
                  <tr style={{ background: "var(--surface-alt)", borderBottom: "1px solid var(--border)", textAlign: "left", position: "sticky", top: 0, zIndex: 1 }}>
                    <th style={{ padding: "8px 12px", fontWeight: 700 }}>Category Name</th>
                    <th style={{ padding: "8px 12px", fontWeight: 700 }}>Code</th>
                    <th style={{ padding: "8px 12px", fontWeight: 700 }}>Current Rate ({categoryCountryCode})</th>
                    <th style={{ padding: "8px 12px", fontWeight: 700, width: "180px" }}>New Unit Rate</th>
                    <th style={{ padding: "8px 12px", fontWeight: 700, width: "140px" }}>Presets</th>
                  </tr>
                </thead>
                <tbody>
                  {activeCategories.map((cat) => {
                    const currentRule = rateRules.find(
                      (r) => r.status === "active" && r.categoryId === cat.id && r.countryCode === categoryCountryCode && (!categoryWorkTypeId || r.workTypeId === Number(categoryWorkTypeId))
                    );
                    const currentVal = categoryDraftRates[cat.id] ?? "";
                    const isKeyRequested = ["Header Character", "Birth Character", "Marriage Character"].includes(cat.name);
                    return (
                      <tr key={cat.id} style={{ borderBottom: "1px solid var(--border)", background: isKeyRequested ? "var(--accent-pale)" : "transparent" }}>
                        <td style={{ padding: "8px 12px" }}>
                          <strong>{cat.name}</strong>
                          {isKeyRequested && (
                            <span style={{ marginLeft: 6, fontSize: "10px", padding: "1px 5px", background: "#fef3c7", color: "#92400e", borderRadius: 3, fontWeight: 700 }}>
                              Priority
                            </span>
                          )}
                        </td>
                        <td style={{ padding: "8px 12px" }} className="mono">
                          {cat.code}
                        </td>
                        <td style={{ padding: "8px 12px" }}>
                          {currentRule ? (
                            <span className="mono" style={{ fontWeight: 700, color: "#073b33" }}>
                              {currentRule.unitRate} ({currentRule.ruleType})
                            </span>
                          ) : (
                            <span style={{ color: "var(--dim)", fontStyle: "italic" }}>Not configured</span>
                          )}
                        </td>
                        <td style={{ padding: "8px 12px" }}>
                          <input
                            inputMode="decimal"
                            placeholder="e.g. 0.0250"
                            value={currentVal}
                            onChange={(e) => {
                              const val = e.target.value;
                              setCategoryDraftRates((prev) => ({ ...prev, [cat.id]: val }));
                            }}
                            style={{ fontSize: "12.5px", padding: "4px 8px", width: "100%", fontFamily: "var(--mono)" }}
                          />
                        </td>
                        <td style={{ padding: "8px 12px" }}>
                          <div style={{ display: "flex", gap: 4 }}>
                            <button
                              type="button"
                              className="secondary-button compact"
                              style={{ fontSize: "10px", padding: "2px 5px" }}
                              onClick={() => setCategoryDraftRates((prev) => ({ ...prev, [cat.id]: "0.0225" }))}
                            >
                              0.0225
                            </button>
                            <button
                              type="button"
                              className="secondary-button compact"
                              style={{ fontSize: "10px", padding: "2px 5px" }}
                              onClick={() => setCategoryDraftRates((prev) => ({ ...prev, [cat.id]: "0.0250" }))}
                            >
                              0.0250
                            </button>
                            <button
                              type="button"
                              className="secondary-button compact"
                              style={{ fontSize: "10px", padding: "2px 5px" }}
                              onClick={() => setCategoryDraftRates((prev) => ({ ...prev, [cat.id]: "0.0300" }))}
                            >
                              0.0300
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 8 }}>
              <button type="button" className="ghost-button" onClick={() => setCategoryModalOpen(false)}>
                Cancel
              </button>
              <button className="primary-button compact" disabled={saveCategoryMutation.isPending}>
                {saveCategoryMutation.isPending ? "Saving Rates…" : "Save All Category Rates"}
              </button>
            </div>
          </form>
        </Modal>
      )}
    </>
  );
}

function Configuration({ token, permissions, notify, mode }: { token: string; permissions: string[]; notify: (notice: Notice) => void; mode: "configuration" | "rates" }) {
  if (mode === "rates") {
    return <RateRules token={token} permissions={permissions} notify={notify} />;
  }

  const queryClient = useQueryClient();
  const canManage = hasPermission(permissions, "rates.manage");
  const projects = useQuery({
    queryKey: ["projects", token],
    queryFn: () => api.listProjects({ token }),
    enabled: hasPermission(permissions, "projects.read"),
  });
  const [projectId, setProjectId] = useState<number | null>(null);
  const [modal, setModal] = useState<"country" | "work" | "category" | null>(null);

  useEffect(() => {
    if (!projectId && projects.data?.projects[0]) {
      setProjectId(projects.data.projects[0].id);
    }
  }, [projectId, projects.data]);

  const config = useQuery({
    queryKey: ["project-config", token, projectId],
    queryFn: () => api.getProjectConfiguration({ token, projectId: projectId ?? 0 }),
    enabled: Boolean(projectId) && hasPermission(permissions, "rates.read"),
  });

  const onSuccess = (text: string) => {
    queryClient.invalidateQueries({ queryKey: ["project-config"] });
    queryClient.invalidateQueries({ queryKey: ["projects"] });
    queryClient.invalidateQueries({ queryKey: ["dashboard"] });
    setModal(null);
    notify({ kind: "success", text });
  };

  const country = useMutation({
    mutationFn: api.saveCountryConfiguration,
    onSuccess: () => onSuccess("Country configuration saved."),
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const work = useMutation({
    mutationFn: api.saveWorkType,
    onSuccess: () => onSuccess("Work type saved."),
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const category = useMutation({
    mutationFn: api.saveCategory,
    onSuccess: () => onSuccess("Category saved."),
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  if (!hasPermission(permissions, "rates.read")) {
    return <Empty title="Configuration access is restricted" body="Your role does not include rates and project configuration." />;
  }

  if (!projects.data?.projects.length) {
    return <Empty title="Create a project first" body="Countries, work types, categories, and rate rules must belong to a project." />;
  }

  return (
    <>
      <div className="toolbar">
        <label className="select-field">
          Project
          <select value={projectId ?? ""} onChange={(e) => setProjectId(Number(e.target.value))}>
            {projects.data.projects.map((project) => (
              <option key={project.id} value={project.id}>
                {project.projectId} · {project.name}
              </option>
            ))}
          </select>
        </label>
      </div>

      {config.isPending ? (
        <Loading />
      ) : config.error || !config.data ? (
        <Empty title="Configuration could not load" body={message(config.error)} />
      ) : (
        <div className="config-columns">
          <ConfigSection
            title="Countries"
            description="Country scopes are effective-dated."
            action={canManage ? () => setModal("country") : undefined}
            actionLabel="Add country"
          >
            {config.data?.countries.length ? (
              config.data.countries.map((item) => (
                <div className="config-item" key={item.id}>
                  <div>
                    <strong className="mono">{item.countryCode}</strong>
                    <span>{item.countryName}</span>
                  </div>
                  <Status value={item.status} />
                  <small>From {formatDate(item.effectiveFrom)}</small>
                </div>
              ))
            ) : (
              <p className="inline-empty">No countries configured.</p>
            )}
          </ConfigSection>

          <ConfigSection
            title="Work types"
            description="Payment basis and unit stay project-specific."
            action={canManage ? () => setModal("work") : undefined}
            actionLabel="Add work type"
          >
            {config.data?.workTypes.length ? (
              config.data.workTypes.map((item) => (
                <div className="config-item" key={item.id}>
                  <div>
                    <strong>{item.name}</strong>
                    <span className="mono">{item.code}</span>
                  </div>
                  <Status value={item.status} />
                  <small>
                    {item.paymentBasis} · {item.unit}
                  </small>
                </div>
              ))
            ) : (
              <p className="inline-empty">No work types configured.</p>
            )}
          </ConfigSection>

          <ConfigSection
            title="Character categories"
            description="Category records replace fixed spreadsheet columns."
            action={canManage ? () => setModal("category") : undefined}
            actionLabel="Add category"
          >
            {config.data?.categories.length ? (
              config.data.categories.map((item) => (
                <div className="config-item" key={item.id}>
                  <div>
                    <strong>{item.name}</strong>
                    <span className="mono">{item.code}</span>
                  </div>
                  <Status value={item.status} />
                  <small>
                    {item.paymentBasis} · {item.quantitySource}
                  </small>
                </div>
              ))
            ) : (
              <p className="inline-empty">No categories configured.</p>
            )}
          </ConfigSection>
        </div>
      )}

      {modal === "country" && projectId && (
        <Modal title="Add country configuration" onClose={() => setModal(null)}>
          <form
            className="stack-form"
            onSubmit={(event) => {
              event.preventDefault();
              const f = new FormData(event.currentTarget);
              country.mutate({
                token,
                projectId,
                countryCode: field(f, "countryCode"),
                countryName: field(f, "countryName"),
                status: "active",
                effectiveFrom: field(f, "effectiveFrom"),
                effectiveTo: null,
              });
            }}
          >
            <label>
              Country code
              <input name="countryCode" placeholder="BD" required />
            </label>
            <label>
              Country name
              <input name="countryName" placeholder="Bangladesh" required />
            </label>
            <label>
              Effective from
              <input name="effectiveFrom" type="date" required />
            </label>
            <button className="primary-button" disabled={country.isPending}>
              Save country
            </button>
          </form>
        </Modal>
      )}

      {modal === "work" && projectId && (
        <Modal title="Add work type" onClose={() => setModal(null)}>
          <form
            className="stack-form"
            onSubmit={(event) => {
              event.preventDefault();
              const f = new FormData(event.currentTarget);
              work.mutate({
                token,
                projectId,
                code: field(f, "code"),
                name: field(f, "name"),
                unit: field(f, "unit") as "record" | "character" | "image" | "batch" | "fixed",
                paymentBasis: field(f, "paymentBasis") as
                  | "record-wise"
                  | "character-wise"
                  | "record-and-character"
                  | "category-wise"
                  | "document-image"
                  | "fixed"
                  | "slab-based"
                  | "hybrid"
                  | "custom",
                status: "active",
              });
            }}
          >
            <label>
              Work type code
              <input name="code" placeholder="REVIEW-KEY" required />
            </label>
            <label>
              Work type name
              <input name="name" placeholder="Review Key" required />
            </label>
            <label>
              Quantity unit
              <select name="unit">
                <option value="record">Record</option>
                <option value="character">Character</option>
                <option value="image">Image</option>
                <option value="batch">Batch</option>
                <option value="fixed">Fixed amount</option>
              </select>
            </label>
            <label>
              Payment basis
              <select name="paymentBasis">
                <option value="record-wise">Record-wise</option>
                <option value="character-wise">Character-wise</option>
                <option value="record-and-character">Record and character</option>
                <option value="category-wise">Category-wise</option>
                <option value="document-image">Document or image</option>
                <option value="fixed">Fixed</option>
                <option value="slab-based">Slab-based</option>
                <option value="hybrid">Hybrid</option>
                <option value="custom">Custom</option>
              </select>
            </label>
            <button className="primary-button" disabled={work.isPending}>
              Save work type
            </button>
          </form>
        </Modal>
      )}

      {modal === "category" && projectId && (
        <Modal title="Add character category" onClose={() => setModal(null)}>
          <form
            className="stack-form"
            onSubmit={(event) => {
              event.preventDefault();
              const f = new FormData(event.currentTarget);
              category.mutate({
                token,
                projectId,
                countryCode: field(f, "countryCode") || null,
                code: field(f, "code"),
                name: field(f, "name"),
                description: field(f, "description") || null,
                paymentBasis: field(f, "paymentBasis") as "record" | "character",
                quantitySource: field(f, "quantitySource"),
                attributionRule: field(f, "attributionRule"),
                effectiveFrom: field(f, "effectiveFrom"),
                status: "active",
                displayOrder: Number(field(f, "displayOrder")),
              });
            }}
          >
            <label>
              Category code
              <input name="code" placeholder="HEADER" required />
            </label>
            <label>
              Category name
              <input name="name" placeholder="Header" required />
            </label>
            <label>
              Country code (optional)
              <input name="countryCode" />
            </label>
            <label>
              Payment basis
              <select name="paymentBasis">
                <option value="record">Record</option>
                <option value="character">Character</option>
              </select>
            </label>
            <label>
              Quantity source
              <input name="quantitySource" placeholder="Normalized Header Char" required />
            </label>
            <label>
              Employee attribution rule
              <input name="attributionRule" placeholder="Source EMID" required />
            </label>
            <label>
              Effective from
              <input name="effectiveFrom" type="date" required />
            </label>
            <label>
              Display order
              <input name="displayOrder" type="number" min="0" defaultValue="0" />
            </label>
            <label>
              Description
              <textarea name="description" rows={2} />
            </label>
            <button className="primary-button" disabled={category.isPending}>
              Save category
            </button>
          </form>
        </Modal>
      )}
    </>
  );
}

function ConfigSection({ title, description, action, actionLabel, children }: { title: string; description: string; action?: () => void; actionLabel: string; children: ReactNode }) {
  return <section className="config-section"><header><div><h2>{title}</h2><p>{description}</p></div>{action && <button className="text-button" onClick={action}>{actionLabel}</button>}</header><div>{children}</div></section>;
}

const SOURCE_LABELS: Record<SourceType, string> = {
  employee_information: "Employee Information",
  record_details: "Record Details",
  doctype: "Doctype",
  proof: "Proof",
  transmitted_batch: "Transmitted Batch List",
  shipdet: "Shipdet",
};

async function sha256File(file: File) {
  const digest = await crypto.subtle.digest("SHA-256", await file.arrayBuffer());
  return Array.from(new Uint8Array(digest), (value) => value.toString(16).padStart(2, "0")).join("");
}
function downloadText(filename: string, content: string, type = "text/csv;charset=utf-8") {
  const url = URL.createObjectURL(new Blob([content], { type }));
  const anchor = document.createElement("a"); anchor.href = url; anchor.download = filename; anchor.click(); URL.revokeObjectURL(url);
}
function downloadBase64(filename: string, base64: string, type: string) {
  const binary = atob(base64); const bytes = new Uint8Array(binary.length); for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  const url = URL.createObjectURL(new Blob([bytes], { type })); const anchor = document.createElement("a"); anchor.href = url; anchor.download = filename; anchor.click(); URL.revokeObjectURL(url);
}

type ImportHistoryBatch = {
  id: number;
  batchId: string;
  periodCode: string;
  projectCode: string;
  sourceType: string;
  status: string;
  totalFiles: number;
  totalRows: number;
  validRows: number;
  rejectedRows: number;
  duplicateRows: number;
  uploadedBy: string;
  createdAt: string;
  updatedAt?: string | null;
  committedAt?: string | null;
};

function ImportBatchTimeline({ batch }: { batch: ImportHistoryBatch }) {
  const isCommitted = Boolean(batch.committedAt) || ["committed", "reversed"].includes(batch.status);
  const isReversed = batch.status === "reversed";
  const isValidated = ["ready", "committed", "reversed"].includes(batch.status) || (batch.validRows + batch.rejectedRows + batch.duplicateRows > 0);
  const isProcessing = batch.status === "processing";
  const isMappingRequired = batch.status === "mapping_required";

  // Derive logical validation timestamp
  const validatedAt = batch.updatedAt || batch.committedAt || batch.createdAt;

  const steps = [
    {
      id: "upload",
      title: "1. Uploaded & Ingested",
      status: "completed" as const,
      badge: "Completed",
      icon: "upload",
      timestamp: formatDateTime(batch.createdAt),
      details: `${batch.totalFiles} source file${batch.totalFiles === 1 ? "" : "s"} (${batch.totalRows.toLocaleString()} raw rows) ingested by ${batch.uploadedBy}. Chunks validated and staged in isolated session.`,
      chips: null,
    },
    {
      id: "validation",
      title: "2. Column Mapping & Validation",
      status: isValidated ? ("completed" as const) : isProcessing ? ("active" as const) : isMappingRequired ? ("action-required" as const) : ("pending" as const),
      badge: isValidated ? "Validated" : isProcessing ? "Validating" : isMappingRequired ? "Needs Mapping" : "Pending",
      icon: "shield",
      timestamp: isValidated ? formatDateTime(validatedAt) : isProcessing ? "Processing chunks…" : "Awaiting user mapping",
      details: isValidated
        ? "Canonical column schema confirmed. Business key collision check and type conversions complete."
        : isProcessing
        ? "Chunk-level validation actively evaluating source rows against business keys."
        : "Source columns require mapping to target schema before row validation.",
      chips: isValidated ? (
        <div className="import-timeline-metrics-chips">
          <span className="import-timeline-chip valid">✓ {batch.validRows.toLocaleString()} valid</span>
          {batch.rejectedRows > 0 && <span className="import-timeline-chip rejected">✗ {batch.rejectedRows.toLocaleString()} rejected</span>}
          {batch.duplicateRows > 0 && <span className="import-timeline-chip duplicate">▲ {batch.duplicateRows.toLocaleString()} duplicates</span>}
        </div>
      ) : null,
    },
    {
      id: "commit",
      title: "3. Committed to Production",
      status: isCommitted ? ("completed" as const) : batch.status === "ready" ? ("active" as const) : ("pending" as const),
      badge: isCommitted ? "Committed" : batch.status === "ready" ? "Ready to Commit" : "Pending",
      icon: "database",
      timestamp: batch.committedAt ? formatDateTime(batch.committedAt) : batch.status === "ready" ? "Awaiting admin commit" : "Pending validation",
      details: isCommitted
        ? `${batch.validRows.toLocaleString()} normalized rows written to persistent database records.`
        : batch.status === "ready"
        ? "Validation passed. Staged records ready for authoritative write into production ledger."
        : "Will commit normalized rows upon validation approval.",
      chips: null,
    },
    {
      id: "lifecycle",
      title: isReversed ? "4. Reversed / Superseded" : isCommitted ? "4. Active in Production" : "4. Pipeline Finalization",
      status: isReversed ? ("reversed" as const) : isCommitted ? ("completed" as const) : ("pending" as const),
      badge: isReversed ? "Reversed" : isCommitted ? "Active" : "Pending",
      icon: isReversed ? "restore" : isCommitted ? "check" : "history",
      timestamp: isReversed ? formatDateTime(batch.updatedAt || batch.committedAt) : isCommitted ? formatDateTime(batch.committedAt) : "Pending",
      details: isReversed
        ? "Normalized rows deactivated. Audit log record and original source files remain traceable."
        : isCommitted
        ? "Records actively participating in period reconciliation and calculation rules."
        : "Awaiting upstream validation and commit completion.",
      chips: null,
    },
  ];

  return (
    <div className="import-timeline-steps">
      {steps.map((step) => (
        <div
          key={step.id}
          className={`import-timeline-step-card ${step.status}`}
        >
          <div className="import-timeline-step-header">
            <div className="import-timeline-node-wrap">
              <div className={`import-timeline-node-icon ${step.status}`}>
                <Icon name={step.icon} />
              </div>
              <span className="import-timeline-step-title">{step.title}</span>
            </div>
            <span className={`import-timeline-step-badge ${step.status}`}>{step.badge}</span>
          </div>

          <div className="import-timeline-timestamp">
            <Icon name="calendar" />
            <span>{step.timestamp}</span>
          </div>

          <div className="import-timeline-details">
            {step.details}
          </div>

          {step.chips}
        </div>
      ))}
    </div>
  );
}

function ImportCenter({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const queryClient = useQueryClient(); const canManage = hasPermission(permissions, "imports.manage");
  const [tab, setTab] = useState<"new" | "history">("new"); const [sourceType, setSourceType] = useState<SourceType>("record_details");
  const [periodId, setPeriodId] = useState<number | null>(null); const [projectId, setProjectId] = useState<number | null>(null); const [delimiter, setDelimiter] = useState<"auto" | "," | ";" | "tab" | "|">("auto");
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]); const [activeBatchId, setActiveBatchId] = useState<number | null>(null); const [mapping, setMapping] = useState<Record<string, string>>({}); const [businessKeys, setBusinessKeys] = useState<string[]>([]); const [templateName, setTemplateName] = useState("");
  const [progress, setProgress] = useState(0); const [previewStatus, setPreviewStatus] = useState<"all" | "valid" | "rejected" | "duplicate">("all"); const [previewPage, setPreviewPage] = useState(1); const [includeDuplicates, setIncludeDuplicates] = useState(false);
  const [historyPeriod, setHistoryPeriod] = useState(""); const [historyProject, setHistoryProject] = useState(""); const [historySource, setHistorySource] = useState(""); const [historyStatus, setHistoryStatus] = useState(""); const [reverseTarget, setReverseTarget] = useState<number | null>(null); const [replacementOfFileId, setReplacementOfFileId] = useState<number | null>(null);
  const [selectedBatchForTimeline, setSelectedBatchForTimeline] = useState<ImportHistoryBatch | null>(null);
  const periods = useQuery({ queryKey: ["periods", token], queryFn: () => api.listPeriods({ token }), enabled: hasPermission(permissions, "periods.read") });
  const projects = useQuery({ queryKey: ["projects", token], queryFn: () => api.listProjects({ token }), enabled: hasPermission(permissions, "projects.read") });
  useEffect(() => { if (!periodId && periods.data?.periods[0]) setPeriodId(periods.data.periods[0].id); }, [periodId, periods.data]);
  useEffect(() => { if (!projectId && projects.data?.projects[0]) setProjectId(projects.data.projects[0].id); }, [projectId, projects.data]);
  const setup = useQuery({ queryKey: ["import-setup", token, sourceType, projectId], queryFn: () => api.getImportSetup({ token, sourceType, projectId }), enabled: hasPermission(permissions, "imports.read") });
  const detail = useQuery({ queryKey: ["import-batch", token, activeBatchId], queryFn: () => api.getImportBatch({ token, importBatchId: activeBatchId ?? 0 }), enabled: Boolean(activeBatchId) });
  const preview = useQuery({ queryKey: ["import-preview", token, activeBatchId, previewStatus, previewPage], queryFn: () => api.listImportPreview({ token, importBatchId: activeBatchId ?? 0, status: previewStatus, page: previewPage, pageSize: 12 }), enabled: Boolean(activeBatchId) && ["ready", "committed"].includes(detail.data?.status ?? "") });
  const history = useQuery({ queryKey: ["import-history", token, historyPeriod, historyProject, historySource, historyStatus], queryFn: () => api.listImportHistory({ token, payrollPeriodId: historyPeriod ? Number(historyPeriod) : null, projectId: historyProject ? Number(historyProject) : null, sourceType: historySource ? historySource as SourceType : null, status: historyStatus || null }), enabled: hasPermission(permissions, "imports.read") });
  const upload = useMutation({ mutationFn: async () => {
    if (!periodId || !projectId || !selectedFiles.length) throw new Error("Select a payroll period, project, and at least one file.");
    setProgress(1); const meta = [];
    for (const file of selectedFiles) { if (!file.size) throw new Error(`${file.name} is empty.`); meta.push({ originalFilename: file.name, fileHash: await sha256File(file), mimeType: file.type || "application/octet-stream", sizeBytes: file.size, expectedChunks: Math.ceil(file.size / (512 * 1024)), replacementOfFileId }); }
    const created = await api.createImportBatch({ token, payrollPeriodId: periodId, projectId, sourceType, delimiter, files: meta }); let completedChunks = 0; const totalChunks = meta.reduce((sum, file) => sum + file.expectedChunks, 0);
    for (let fileIndex = 0; fileIndex < selectedFiles.length; fileIndex += 1) { const file = selectedFiles[fileIndex]; const registered = created.files[fileIndex]; if (!file || !registered) throw new Error("The registered file list does not match the selected files."); for (let chunkIndex = 0; chunkIndex < meta[fileIndex]!.expectedChunks; chunkIndex += 1) { const chunk = file.slice(chunkIndex * 512 * 1024, Math.min(file.size, (chunkIndex + 1) * 512 * 1024)); const encoded = await fileToBase64(chunk); await api.uploadImportChunk({ token, sourceFileId: registered.id, chunkIndex, dataBase64: encoded.dataBase64 }); completedChunks += 1; setProgress(Math.max(2, Math.round((completedChunks / totalChunks) * 82))); } await api.finalizeImportFile({ token, sourceFileId: registered.id }); }
    const batch = await api.getImportBatch({ token, importBatchId: created.batchId }); setMapping(batch.mapping); setBusinessKeys(batch.businessKeys); setProgress(100); return created;
  }, onSuccess: (data) => { setActiveBatchId(data.batchId); setSelectedFiles([]); setReplacementOfFileId(null); queryClient.invalidateQueries({ queryKey: ["import-history"] }); queryClient.invalidateQueries({ queryKey: ["periods"] }); notify({ kind: "success", text: `${data.batchCode} uploaded and parsed on the server. Review the column mapping before validation.` }); }, onError: (error) => { setProgress(0); notify({ kind: "error", text: message(error) }); } });
  const validate = useMutation({ mutationFn: async () => { if (!activeBatchId) throw new Error("No import batch is open."); await api.beginImportValidation({ token, importBatchId: activeBatchId, mapping, businessKeys, saveTemplateName: templateName || null }); let cursor = 0; let done = false; let counts = { totalRows: detail.data?.totalRows ?? 0, validRows: 0, rejectedRows: 0, duplicateRows: 0 }; while (!done) { const result = await api.validateImportChunk({ token, importBatchId: activeBatchId, cursor, limit: 250 }); cursor = result.nextCursor; done = result.done; counts = result; setProgress(result.totalRows ? Math.round(((result.validRows + result.rejectedRows + result.duplicateRows) / result.totalRows) * 100) : 100); } return counts; }, onSuccess: (counts) => { queryClient.invalidateQueries({ queryKey: ["import-batch"] }); queryClient.invalidateQueries({ queryKey: ["import-preview"] }); queryClient.invalidateQueries({ queryKey: ["import-history"] }); setPreviewStatus("all"); setPreviewPage(1); notify({ kind: counts.rejectedRows ? "error" : "success", text: `${counts.validRows} valid, ${counts.rejectedRows} rejected, and ${counts.duplicateRows} duplicate candidates. Review the preview before committing.` }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  const commit = useMutation({ mutationFn: async () => { if (!activeBatchId) throw new Error("No import batch is open."); let cursor = 0; let done = false; let inserted = 0; while (!done) { const result = await api.commitImportChunk({ token, importBatchId: activeBatchId, cursor, limit: 100, includeDuplicateCandidates: includeDuplicates }); cursor = result.nextCursor; done = result.done; inserted += result.inserted; setProgress(detail.data?.totalRows ? Math.min(99, Math.round((cursor / Math.max(cursor, detail.data.totalRows)) * 100)) : 90); } setProgress(100); return inserted; }, onSuccess: (inserted) => { queryClient.invalidateQueries({ queryKey: ["import-batch"] }); queryClient.invalidateQueries({ queryKey: ["import-history"] }); queryClient.invalidateQueries({ queryKey: ["audit"] }); notify({ kind: "success", text: `${inserted} normalized source rows were committed to the persistent database.` }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  const reprocess = useMutation({ mutationFn: ({ importBatchId, reason }: { importBatchId: number; reason: string }) => api.reprocessImportBatch({ token, importBatchId, reason }), onSuccess: async (data) => { setTab("new"); setActiveBatchId(data.batchId); const batch = await api.getImportBatch({ token, importBatchId: data.batchId }); setPeriodId(batch.payrollPeriodId); setProjectId(batch.projectId); setSourceType(batch.sourceType as SourceType); setMapping(batch.mapping); setBusinessKeys(batch.businessKeys); queryClient.invalidateQueries({ queryKey: ["import-history"] }); notify({ kind: "success", text: "A retry-safe reprocessing batch is ready. The committed source remains active until this replacement commits." }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  const reverse = useMutation({ mutationFn: ({ importBatchId, reason }: { importBatchId: number; reason: string }) => api.reverseImportBatch({ token, importBatchId, reason }), onSuccess: () => { setReverseTarget(null); queryClient.invalidateQueries({ queryKey: ["import-history"] }); queryClient.invalidateQueries({ queryKey: ["import-batch"] }); queryClient.invalidateQueries({ queryKey: ["audit"] }); notify({ kind: "success", text: "The import was reversed. Its source rows remain traceable but are no longer active." }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  const inspectHistory = async (id: number) => { const batch = await api.getImportBatch({ token, importBatchId: id }); setActiveBatchId(id); setPeriodId(batch.payrollPeriodId); setProjectId(batch.projectId); setSourceType(batch.sourceType as SourceType); setMapping(batch.mapping); setBusinessKeys(batch.businessKeys); setTab("new"); };
  const requestReplacement = async (batchId: number, fileId: number) => { const batch = await api.getImportBatch({ token, importBatchId: batchId }); setPeriodId(batch.payrollPeriodId); setProjectId(batch.projectId); setSourceType(batch.sourceType as SourceType); setReplacementOfFileId(fileId); setActiveBatchId(null); setSelectedFiles([]); setTab("new"); notify({ kind: "success", text: "Replacement mode is active. Choose the corrected file; the current data stays active until the replacement commits." }); };
  const selectTemplate = (id: string) => { const template = setup.data?.templates.find((item) => item.id === Number(id)); if (template) { setMapping(template.mapping); setBusinessKeys(template.businessKeys); setDelimiter(template.delimiter as typeof delimiter); } };
  const addFiles = (files: FileList | File[]) => { const next = Array.from(files).filter((file) => /\.(csv|xlsx|xls)$/i.test(file.name)); if (!next.length) notify({ kind: "error", text: "Choose CSV, XLSX, or XLS files." }); else setSelectedFiles(replacementOfFileId ? next.slice(0, 1) : next.slice(0, 250)); };
  if (!hasPermission(permissions, "imports.read")) return <Empty title="Import access is restricted" body="Your role does not include the Data Import Center." />;
  const batch = detail.data;
  return <div className="import-center"><div className="tab-bar" role="tablist"><button role="tab" aria-selected={tab === "new"} className={tab === "new" ? "active" : ""} onClick={() => setTab("new")}>New import</button><button role="tab" aria-selected={tab === "history"} className={tab === "history" ? "active" : ""} onClick={() => setTab("history")}>Import history</button></div>
    {tab === "new" ? <>
      <section className="import-scope"><div><p className="overline">1 · Source scope</p><h2>Define the controlled import</h2></div><div className="scope-grid"><label>Source type<select value={sourceType} disabled={Boolean(replacementOfFileId)} onChange={(e) => { setSourceType(e.target.value as SourceType); setActiveBatchId(null); setMapping({}); setBusinessKeys([]); }}><option value="employee_information">Employee Information CSV</option><option value="record_details">Record Details</option><option value="doctype">Doctype</option><option value="proof">Proof</option><option value="transmitted_batch">Transmitted Batch List</option><option value="shipdet">Shipdet</option></select></label><label>Payroll period<select value={periodId ?? ""} disabled={Boolean(replacementOfFileId)} onChange={(e) => setPeriodId(Number(e.target.value))}><option value="">Select period</option>{periods.data?.periods.map((period) => <option key={period.id} value={period.id}>{period.periodId} · {period.status}</option>)}</select></label><label>Project<select value={projectId ?? ""} disabled={Boolean(replacementOfFileId)} onChange={(e) => setProjectId(Number(e.target.value))}><option value="">Select project</option>{projects.data?.projects.map((project) => <option key={project.id} value={project.id}>{project.projectId} · {project.name}</option>)}</select></label><label>CSV delimiter<select value={delimiter} onChange={(e) => setDelimiter(e.target.value as typeof delimiter)}><option value="auto">Detect automatically</option><option value=",">Comma (,)</option><option value=";">Semicolon (;)</option><option value="tab">Tab</option><option value="|">Pipe (|)</option></select></label></div></section>
      {canManage && !batch && <section className={`drop-zone ${selectedFiles.length ? "has-files" : ""}`} onDragOver={(event) => event.preventDefault()} onDrop={(event) => { event.preventDefault(); addFiles(event.dataTransfer.files); }}><p className="overline">2 · Files</p><h2>{replacementOfFileId ? "Choose the corrected replacement file" : "Drop files here, or choose a file source"}</h2><p>CSV, XLSX, and XLS are processed in 512 KB server-upload chunks. Multiple Record Details files become one logical import batch.</p><div className="drop-actions"><label className="primary-button file-button">Select files<input aria-label="Select multiple import files" type="file" accept=".csv,.xlsx,.xls" multiple={!replacementOfFileId} onChange={(e) => { if (e.target.files) addFiles(e.target.files); e.currentTarget.value = ""; }} /></label>{!replacementOfFileId && <label className="secondary-button file-button">Select folder<input aria-label="Select import folder" type="file" accept=".csv,.xlsx,.xls" multiple {...({ webkitdirectory: "", directory: "" } as Record<string, string>)} onChange={(e) => { if (e.target.files) addFiles(e.target.files); e.currentTarget.value = ""; }} /></label>}</div>{replacementOfFileId && <div className="replacement-note">Replacement target: source file #{replacementOfFileId}. The current committed rows remain active until the new batch passes validation and commits.</div>}{selectedFiles.length > 0 && <div className="file-queue"><div><strong>{selectedFiles.length} file{selectedFiles.length === 1 ? "" : "s"} selected</strong><span>{formatBytes(selectedFiles.reduce((sum, file) => sum + file.size, 0))} total</span></div><ul>{selectedFiles.slice(0, 8).map((file) => <li key={`${file.name}-${file.size}`}><span>{file.name}</span><b>{formatBytes(file.size)}</b></li>)}</ul>{selectedFiles.length > 8 && <small>And {selectedFiles.length - 8} more files</small>}<button className="primary-button compact" disabled={upload.isPending || !periodId || !projectId} onClick={() => upload.mutate()}>{upload.isPending ? `Uploading and processing · ${progress}%` : "Upload and inspect headers"}</button></div>}</section>}
      {batch && <><section className="batch-banner"><div><p className="overline">{batch.batchId}</p><h2>{SOURCE_LABELS[batch.sourceType as SourceType]} import</h2><p>{batch.totalFiles} files · {batch.totalRows.toLocaleString()} source rows · server-side status <b>{batch.status}</b></p></div><div className="count-strip"><span><b>{batch.validRows}</b>Valid</span><span className="bad"><b>{batch.rejectedRows}</b>Rejected</span><span className="warn"><b>{batch.duplicateRows}</b>Duplicates</span></div></section><section className="source-file-list"><header><h3>Source file metadata</h3><span>Every file keeps its own ID, hash, row counts, status, and timestamps.</span></header>{batch.files.map((file) => <article key={file.id}><div><strong>{file.originalFilename}</strong><small className="mono">{file.sourceFileId} · SHA-256 {file.fileHash.slice(0, 12)}…</small></div><div><Status value={file.status} /><small>{file.totalRows} rows · {formatBytes(file.sizeBytes)}</small></div><div><span>{file.importedRows} imported</span><small>{file.rejectedRows} rejected · {file.duplicateRows} duplicate</small></div>{file.status === "committed" && canManage && <label className="text-button file-button">Replace<input aria-label={`Replace ${file.originalFilename}`} type="file" accept=".csv,.xlsx,.xls" onChange={(e) => { const next = e.target.files?.[0]; if (next) { void requestReplacement(batch.id, file.id).then(() => setSelectedFiles([next])); } e.currentTarget.value = ""; }} /></label>}</article>)}</section>
        {batch.status === "mapping_required" || batch.status === "failed" ? <section className="mapping-panel"><div className="section-heading"><div><p className="overline">3 · Mapping</p><h2>Map source columns to canonical fields</h2><p>Required fields must be mapped. Unmapped source columns remain preserved in the original payload.</p></div>{setup.data?.templates.length ? <label>Saved template<select defaultValue="" onChange={(e) => selectTemplate(e.target.value)}><option value="">Choose a template</option>{setup.data.templates.map((template) => <option key={template.id} value={template.id}>{template.name}</option>)}</select></label> : null}</div><div className="mapping-grid">{setup.data?.fields.map((canonical) => <label key={canonical.key}><span>{canonical.label}{canonical.required && <b>Required</b>}<small>{canonical.type}</small></span><select value={mapping[canonical.key] ?? ""} onChange={(e) => setMapping((current) => ({ ...current, [canonical.key]: e.target.value }))}><option value="">Not mapped</option>{batch.headers.map((header) => <option key={header} value={header}>{header}</option>)}</select><input type="checkbox" checked={businessKeys.includes(canonical.key)} onChange={(e) => setBusinessKeys((keys) => e.target.checked ? [...new Set([...keys, canonical.key])] : keys.filter((key) => key !== canonical.key))} aria-label={`Use ${canonical.label} in duplicate business key`} title="Use in duplicate business key" /></label>)}</div><div className="mapping-actions"><label>Save mapping template (optional)<input value={templateName} onChange={(e) => setTemplateName(e.target.value)} placeholder="e.g. Monthly Record Details" /></label><div><span>Business key: {businessKeys.map((key) => setup.data?.fields.find((field) => field.key === key)?.label ?? key).join(" + ") || "None selected"}</span><button className="primary-button compact" disabled={validate.isPending || !businessKeys.length} onClick={() => validate.mutate()}>{validate.isPending ? `Validating in chunks · ${progress}%` : "Validate and build preview"}</button></div></div></section> : null}
        {batch.status === "ready" || batch.status === "committed" ? <section className="preview-panel"><div className="section-heading"><div><p className="overline">4 · Preview</p><h2>Review every outcome before commit</h2><p>Rejected rows remain stored with reasons. Duplicate candidates are never included unless explicitly selected.</p></div><div className="preview-tabs">{(["all", "valid", "rejected", "duplicate"] as const).map((value) => <button key={value} className={previewStatus === value ? "active" : ""} onClick={() => { setPreviewStatus(value); setPreviewPage(1); }}>{humanize(value)}</button>)}</div></div>{preview.isPending ? <Loading /> : preview.data?.rows.length ? <><div className="preview-table"><div className="preview-head"><span>Source</span><span>Normalized values</span><span>Outcome</span></div>{preview.data.rows.map((row) => <article key={row.id}><div><strong>{row.sourceFileName}</strong><small>Row {row.sourceRowNumber}</small></div><div>{Object.entries(row.normalized).filter(([, value]) => value !== null && !String(value).startsWith("source:")).slice(0, 4).map(([key, value]) => <span key={key}><b>{humanize(key)}</b>{value}</span>)}</div><div><Status value={row.status} />{row.errors.map((error) => <small key={error}>{error}</small>)}</div></article>)}</div><div className="pager"><button className="secondary-button compact" disabled={previewPage === 1} onClick={() => setPreviewPage((page) => page - 1)}>Previous</button><span>Page {previewPage} · {preview.data.total} rows</span><button className="secondary-button compact" disabled={previewPage * 12 >= preview.data.total} onClick={() => setPreviewPage((page) => page + 1)}>Next</button></div></> : <Empty title="No rows in this view" body="Choose another preview status or validate the source mapping again." />}{batch.rejectedRows > 0 && <button className="secondary-button compact" onClick={() => void api.downloadImportErrors({ token, importBatchId: batch.id }).then((report) => downloadText(report.filename, report.content)).catch((error) => notify({ kind: "error", text: message(error) }))}>Download {batch.rejectedRows} rejected rows</button>}{batch.status === "ready" && canManage && <div className="commit-bar"><label><input type="checkbox" checked={includeDuplicates} onChange={(e) => setIncludeDuplicates(e.target.checked)} /> Include duplicate candidates after review</label><button className="primary-button" disabled={commit.isPending} onClick={() => commit.mutate()}>{commit.isPending ? `Committing normalized rows · ${progress}%` : `Commit ${batch.validRows + (includeDuplicates ? batch.duplicateRows : 0)} rows`}</button></div>}{batch.status === "committed" && <div className="committed-note"><Status value="committed" /><span>Normalized rows are active in persistent storage. Reprocess or reverse only through Import History.</span></div>}</section> : null}</>}</>
      : <section className="history-panel">
          <div className="history-filters">
            <label>Payroll period<select value={historyPeriod} onChange={(e) => setHistoryPeriod(e.target.value)}><option value="">All periods</option>{periods.data?.periods.map((period) => <option key={period.id} value={period.id}>{period.periodId}</option>)}</select></label>
            <label>Project<select value={historyProject} onChange={(e) => setHistoryProject(e.target.value)}><option value="">All projects</option>{projects.data?.projects.map((project) => <option key={project.id} value={project.id}>{project.projectId}</option>)}</select></label>
            <label>Source type<select value={historySource} onChange={(e) => setHistorySource(e.target.value)}><option value="">All sources</option>{Object.entries(SOURCE_LABELS).map(([value, label]) => <option key={value} value={value}>{label}</option>)}</select></label>
            <label>Status<select value={historyStatus} onChange={(e) => setHistoryStatus(e.target.value)}><option value="">All statuses</option>{["uploading", "processing", "mapping_required", "ready", "committed", "failed", "reversed"].map((status) => <option key={status} value={status}>{humanize(status)}</option>)}</select></label>
          </div>

          {selectedBatchForTimeline && (
            <div className="import-timeline-container">
              <div className="import-timeline-header">
                <div className="import-timeline-title">
                  <div>
                    <p className="overline" style={{ margin: 0 }}>Sequence of Status Changes</p>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 3 }}>
                      <strong className="mono" style={{ fontSize: 16 }}>{selectedBatchForTimeline.batchId}</strong>
                      <Status value={selectedBatchForTimeline.status} />
                    </div>
                  </div>
                  <div className="timeline-badge-group">
                    <span className="timeline-meta-pill">📂 {SOURCE_LABELS[selectedBatchForTimeline.sourceType as SourceType] || selectedBatchForTimeline.sourceType}</span>
                    <span className="timeline-meta-pill">🗓 {selectedBatchForTimeline.periodCode}</span>
                    <span className="timeline-meta-pill">🏢 {selectedBatchForTimeline.projectCode}</span>
                    <span className="timeline-meta-pill">📄 {selectedBatchForTimeline.totalFiles} File{selectedBatchForTimeline.totalFiles === 1 ? "" : "s"}</span>
                    <span className="timeline-meta-pill">🔢 {selectedBatchForTimeline.totalRows.toLocaleString()} Rows</span>
                  </div>
                </div>
                <button
                  type="button"
                  className="icon-button"
                  aria-label="Close timeline"
                  onClick={() => setSelectedBatchForTimeline(null)}
                  title="Close timeline"
                >
                  <Icon name="close" />
                </button>
              </div>
              <ImportBatchTimeline batch={selectedBatchForTimeline} />
            </div>
          )}

          {history.isPending ? (
            <Loading />
          ) : !history.data?.batches.length ? (
            <Empty title="No import history" body="Upload a source file to begin the durable import history." />
          ) : (
            <div className="history-list">
              <div className="history-head">
                <span>Batch and scope</span>
                <span>Source</span>
                <span>Row outcome</span>
                <span>Audit</span>
                <span />
              </div>
              {history.data.batches.map((item) => (
                <article key={item.id} className={selectedBatchForTimeline?.id === item.id ? "selected-batch" : ""}>
                  <div>
                    <strong className="mono">{item.batchId}</strong>
                    <small>{item.periodCode} · {item.projectCode}</small>
                  </div>
                  <div>
                    <span>{SOURCE_LABELS[item.sourceType as SourceType]}</span>
                    <small>{item.totalFiles} file{item.totalFiles === 1 ? "" : "s"}</small>
                  </div>
                  <div>
                    <span>{item.validRows.toLocaleString()} valid</span>
                    <small>{item.rejectedRows.toLocaleString()} rejected · {item.duplicateRows.toLocaleString()} duplicate</small>
                  </div>
                  <div>
                    <Status value={item.status} />
                    <small>{item.uploadedBy} · {formatDate(item.createdAt)}</small>
                  </div>
                  <div className="row-actions">
                    <button
                      type="button"
                      className={`text-button ${selectedBatchForTimeline?.id === item.id ? "active-timeline-btn" : ""}`}
                      onClick={() => setSelectedBatchForTimeline(selectedBatchForTimeline?.id === item.id ? null : item)}
                      title="Inspect sequence of status changes"
                    >
                      ⏱ Timeline
                    </button>
                    <button className="text-button" onClick={() => void inspectHistory(item.id)}>Inspect</button>
                    {item.status === "committed" && canManage && (
                      <>
                        <button className="text-button" onClick={() => reprocess.mutate({ importBatchId: item.id, reason: "Reprocess with revised mapping or validation rules" })}>Reprocess</button>
                        <button className="text-button danger-text" onClick={() => setReverseTarget(item.id)}>Reverse</button>
                      </>
                    )}
                  </div>
                </article>
              ))}
            </div>
          )}
        </section>}
    {reverseTarget && <Modal title="Reverse committed import" onClose={() => setReverseTarget(null)}><form className="stack-form" onSubmit={(event) => { event.preventDefault(); const form = new FormData(event.currentTarget); reverse.mutate({ importBatchId: reverseTarget, reason: field(form, "reason") }); }}><p className="form-lead">The source files and audit trail remain. Only the normalized rows become inactive.</p><label>Reason<textarea name="reason" minLength={3} rows={3} required /></label><div className="form-actions"><button type="button" className="ghost-button" onClick={() => setReverseTarget(null)}>Cancel</button><button className="primary-button" disabled={reverse.isPending}>Reverse import</button></div></form></Modal>}
  </div>;
}

function ProductionScope({ token, periodId, projectId, setPeriodId, setProjectId }: { token: string; periodId: number | null; projectId: number | null; setPeriodId: (value: number | null) => void; setProjectId: (value: number | null) => void }) {
  const periods = useQuery({ queryKey: ["periods", token], queryFn: () => api.listPeriods({ token }) });
  const projects = useQuery({ queryKey: ["projects", token], queryFn: () => api.listProjects({ token }) });
  return <div className="production-scope"><label>Payroll period<select value={periodId ?? ""} onChange={(e) => setPeriodId(e.target.value ? Number(e.target.value) : null)}><option value="">All periods</option>{periods.data?.periods.map((period) => <option key={period.id} value={period.id}>{period.periodId}</option>)}</select></label><label>Project<select value={projectId ?? ""} onChange={(e) => setProjectId(e.target.value ? Number(e.target.value) : null)}><option value="">All projects</option>{projects.data?.projects.map((project) => <option key={project.id} value={project.id}>{project.projectId} · {project.name}</option>)}</select></label></div>;
}

function ProductionOverview({ token, periodId, projectId }: { token: string; periodId: number | null; projectId: number | null }) {
  const overview = useQuery({ queryKey: ["production-overview", token, periodId, projectId], queryFn: () => api.getProductionOverview({ token, payrollPeriodId: periodId, projectId }) });
  if (!overview.data) return null;
  const values = [["Record Details", overview.data.recordDetails], ["Categories", overview.data.categoryQuantities], ["Doctype", overview.data.doctype], ["Proof", overview.data.proof], ["Transmitted", overview.data.transmitted], ["Shipdet", overview.data.shipdet], ["Open exceptions", overview.data.openExceptions]];
  return <div className="production-metrics">{values.map(([label, value]) => <span key={String(label)}><b>{value}</b>{label}</span>)}</div>;
}

function RecordDetails({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const allowed = hasPermission(permissions, "production.read"); const [periodId, setPeriodId] = useState<number | null>(null); const [projectId, setProjectId] = useState<number | null>(null); const [page, setPage] = useState(1); const [detailId, setDetailId] = useState<number | null>(null);
  const [filters, setFilters] = useState({ country: "", employee: "", batch: "", workType: "", keyType: "", category: "", sourceFile: "", dateFrom: "", dateTo: "", sortBy: "productionDate" as "productionDate" | "employeeIdentifier" | "internalBatchId" | "recordCount" | "characterCount" | "sourceRowNumber", direction: "desc" as "asc" | "desc" });
  const query = useQuery({ queryKey: ["record-details", token, periodId, projectId, filters, page], queryFn: () => api.listRecordDetails({ token, payrollPeriodId: periodId, projectId, ...filters, page, pageSize: 25 }), enabled: allowed });
  const detail = useQuery({ queryKey: ["record-detail", token, detailId], queryFn: () => api.getRecordDetail({ token, id: detailId ?? 0 }), enabled: Boolean(detailId) });
  if (!allowed) return <Empty title="Production data access is restricted" body="Your role does not include normalized production records." />;
  const setFilter = (key: keyof typeof filters, value: string) => { setPage(1); setFilters((current) => ({ ...current, [key]: value })); };
  const exportCsv = () => void api.exportProductionCsv({ token, stream: "record_details", payrollPeriodId: periodId, projectId }).then((report) => { downloadText(report.filename, report.content); notify({ kind: "success", text: `${report.rowCount} Record Details rows exported.` }); }).catch((error) => notify({ kind: "error", text: message(error) }));
  return <><ProductionScope token={token} periodId={periodId} projectId={projectId} setPeriodId={(value) => { setPeriodId(value); setPage(1); }} setProjectId={(value) => { setProjectId(value); setPage(1); }} /><ProductionOverview token={token} periodId={periodId} projectId={projectId} /><section className="filter-panel"><div className="filter-grid"><label>Employee<input value={filters.employee} onChange={(e) => setFilter("employee", e.target.value)} placeholder="ID or name" /></label><label>Batch<input value={filters.batch} onChange={(e) => setFilter("batch", e.target.value)} placeholder="Internal or client ID" /></label><label>Country<input value={filters.country} onChange={(e) => setFilter("country", e.target.value)} /></label><label>Work type<input value={filters.workType} onChange={(e) => setFilter("workType", e.target.value)} /></label><label>Key type<input value={filters.keyType} onChange={(e) => setFilter("keyType", e.target.value)} /></label><label>Category<input value={filters.category} onChange={(e) => setFilter("category", e.target.value)} /></label><label>Source file<input value={filters.sourceFile} onChange={(e) => setFilter("sourceFile", e.target.value)} /></label><label>From date<input type="date" value={filters.dateFrom} onChange={(e) => setFilter("dateFrom", e.target.value)} /></label><label>To date<input type="date" value={filters.dateTo} onChange={(e) => setFilter("dateTo", e.target.value)} /></label><label>Sort<select value={filters.sortBy} onChange={(e) => setFilter("sortBy", e.target.value)}><option value="productionDate">Production date</option><option value="employeeIdentifier">Employee ID</option><option value="internalBatchId">Batch ID</option><option value="recordCount">Record count</option><option value="characterCount">Character count</option><option value="sourceRowNumber">Source row</option></select></label><label>Order<select value={filters.direction} onChange={(e) => setFilter("direction", e.target.value)}><option value="desc">Descending</option><option value="asc">Ascending</option></select></label></div><button className="secondary-button compact" onClick={exportCsv}>Export current scope</button></section>{query.isPending ? <Loading /> : query.error ? <Empty title="Record Details could not load" body={message(query.error)} /> : !query.data?.rows.length ? <Empty title="No normalized Record Details" body="Commit a Record Details import, or adjust the active filters." /> : <div className="production-list">{query.data.rows.map((row) => <article key={row.id}><header><div><span className="mono">{row.internalBatchId || row.clientBatchId || "No batch ID"}</span><h3>{row.employeeIdentifier || "Unmatched employee"} · {row.employeeName || "Name unavailable"}</h3></div><div><Status value={row.sourceVariant} />{row.payableState !== "eligible" && <Status value={row.payableState} />}</div></header><div className="production-facts"><span><b>Period / project</b>{row.periodCode} · {row.projectCode}</span><span><b>Production date</b>{formatDate(row.productionDate)}</span><span><b>Country / work</b>{[row.country, row.workType, row.keyType].filter(Boolean).join(" · ") || "—"}</span><span><b>Records / characters</b>{row.recordCount ?? "—"} / {row.characterCount ?? "—"}</span><span><b>Source</b>{row.sourceFile} · row {row.sourceRowNumber}</span><span><b>Employee match</b>{row.employeeMatched ? "Matched" : "Exception"}</span></div>{row.categories.length > 0 && <div className="category-chips">{row.categories.map((item) => <span key={`${item.code}-${item.unit}`}>{item.name} <b>{item.quantity}</b> {item.unit}</span>)}</div>}<button className="text-button" onClick={() => setDetailId(row.id)}>Inspect source row</button></article>)}</div>}<div className="pager"><button className="secondary-button compact" disabled={page === 1} onClick={() => setPage((value) => value - 1)}>Previous</button><span>Page {page} · {query.data?.total ?? 0} rows</span><button className="secondary-button compact" disabled={page * 25 >= (query.data?.total ?? 0)} onClick={() => setPage((value) => value + 1)}>Next</button></div>{detailId && <Modal title="Source row traceability" onClose={() => setDetailId(null)} wide>{detail.isPending ? <Loading /> : detail.error || !detail.data ? <Empty title="Source row unavailable" body={message(detail.error)} /> : <div className="source-inspector"><p><b>{detail.data.sourceFile}</b> · row {detail.data.sourceRowNumber} · {humanize(detail.data.sourceVariant)}</p>{detail.data.categories.length > 0 && <div className="category-chips">{detail.data.categories.map((item) => <span key={`${item.code}-${item.unit}`}>{item.name}: <b>{item.quantity}</b> {item.unit}</span>)}</div>}<dl>{Object.entries(detail.data.originalPayload).map(([key, value]) => <div key={key}><dt>{key}</dt><dd>{String(value ?? "") || "—"}</dd></div>)}</dl></div>}</Modal>}</>;
}

function ProductionWork({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const allowed = hasPermission(permissions, "production.read"); const [stream, setStream] = useState<"doctype" | "proof">("doctype"); const [periodId, setPeriodId] = useState<number | null>(null); const [projectId, setProjectId] = useState<number | null>(null); const [search, setSearch] = useState(""); const [page, setPage] = useState(1);
  const query = useQuery({ queryKey: ["production-stream", token, stream, periodId, projectId, search, page], queryFn: () => api.listProductionStream({ token, stream, payrollPeriodId: periodId, projectId, search, page, pageSize: 25 }), enabled: allowed });
  if (!allowed) return <Empty title="Production data access is restricted" body="Your role does not include Doctype and Proof records." />;
  const exportCsv = () => void api.exportProductionCsv({ token, stream, payrollPeriodId: periodId, projectId }).then((report) => { downloadText(report.filename, report.content); notify({ kind: "success", text: `${report.rowCount} ${stream === "doctype" ? "Doctype" : "Proof"} rows exported.` }); }).catch((error) => notify({ kind: "error", text: message(error) }));
  return <><ProductionScope token={token} periodId={periodId} projectId={projectId} setPeriodId={(value) => { setPeriodId(value); setPage(1); }} setProjectId={(value) => { setProjectId(value); setPage(1); }} /><ProductionOverview token={token} periodId={periodId} projectId={projectId} /><div className="toolbar"><div className="tab-bar"><button className={stream === "doctype" ? "active" : ""} onClick={() => { setStream("doctype"); setPage(1); }}>Doctype</button><button className={stream === "proof" ? "active" : ""} onClick={() => { setStream("proof"); setPage(1); }}>Proof</button></div><label className="search-field">Search current stream<input value={search} onChange={(e) => { setSearch(e.target.value); setPage(1); }} placeholder="Batch, employee, type, or country" /></label><button className="secondary-button compact" onClick={exportCsv}>Export current scope</button></div><StreamRows query={query} stream={stream} page={page} setPage={setPage} /></>;
}

function StreamRows({ query, stream, page, setPage, selected = [], setSelected }: { query: { isPending: boolean; error: unknown; data?: { total: number; page: number; pageSize: number; rows: Array<{ id: number; sourceFile: string; sourceRowNumber: number; periodCode: string; projectCode: string; fields: Record<string, string | null> }> } }; stream: ProductionStream; page: number; setPage: (value: number | ((value: number) => number)) => void; selected?: number[]; setSelected?: (ids: number[]) => void }) {
  if (query.isPending) return <Loading />; if (query.error) return <Empty title="Production stream could not load" body={message(query.error)} />; if (!query.data?.rows.length) return <Empty title={`No ${humanize(stream)} rows`} body="Commit the corresponding source import, or adjust the active filters." />;
  return <><div className="production-list compact-list">{query.data.rows.map((row: { id: number; sourceFile: string; sourceRowNumber: number; periodCode: string; projectCode: string; fields: Record<string, string | null> }) => <article key={row.id}>{setSelected && <label className="row-select"><input type="checkbox" checked={selected.includes(row.id)} onChange={(e) => setSelected(e.target.checked ? [...selected, row.id] : selected.filter((id) => id !== row.id))} />Select row</label>}<header><div><span className="mono">{row.periodCode} · {row.projectCode}</span><h3>{row.fields.batchId ?? row.fields.internalBatchId ?? row.fields.shipmentBatchId ?? row.fields.proofType ?? humanize(stream)}</h3></div>{row.fields.reviewStatus && <Status value={row.fields.reviewStatus} />}</header><div className="production-facts">{Object.entries(row.fields).map(([key, value]) => <span key={key}><b>{humanize(key)}</b>{value ?? "—"}</span>)}<span><b>Source</b>{row.sourceFile} · row {row.sourceRowNumber}</span></div></article>)}</div><div className="pager"><button className="secondary-button compact" disabled={page === 1} onClick={() => setPage((value) => typeof value === "number" ? value - 1 : 1)}>Previous</button><span>Page {page} · {query.data.total} rows</span><button className="secondary-button compact" disabled={page * 25 >= query.data.total} onClick={() => setPage((value) => typeof value === "number" ? value + 1 : 1)}>Next</button></div></>;
}

function ProductionShipments({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const queryClient = useQueryClient(); const allowed = hasPermission(permissions, "production.read"); const canManage = hasPermission(permissions, "production.manage"); const [stream, setStream] = useState<"transmitted_batch" | "shipdet">("transmitted_batch"); const [periodId, setPeriodId] = useState<number | null>(null); const [projectId, setProjectId] = useState<number | null>(null); const [search, setSearch] = useState(""); const [page, setPage] = useState(1); const [selected, setSelected] = useState<number[]>([]); const [confirming, setConfirming] = useState(false);
  const query = useQuery({ queryKey: ["production-stream", token, stream, periodId, projectId, search, page], queryFn: () => api.listProductionStream({ token, stream, payrollPeriodId: periodId, projectId, search, page, pageSize: 25 }), enabled: allowed });
  const confirm = useMutation({ mutationFn: ({ reason }: { reason: string }) => api.confirmShipdetRows({ token, ids: selected, reason }), onSuccess: (data) => { queryClient.invalidateQueries({ queryKey: ["production-stream"] }); queryClient.invalidateQueries({ queryKey: ["production-overview"] }); setSelected([]); setConfirming(false); notify({ kind: "success", text: `${data.confirmed} Shipdet rows confirmed after inspection.` }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  if (!allowed) return <Empty title="Production data access is restricted" body="Your role does not include transmission and shipment records." />;
  const exportCsv = () => void api.exportProductionCsv({ token, stream, payrollPeriodId: periodId, projectId }).then((report) => { downloadText(report.filename, report.content); notify({ kind: "success", text: `${report.rowCount} rows exported.` }); }).catch((error) => notify({ kind: "error", text: message(error) }));
  return <><ProductionScope token={token} periodId={periodId} projectId={projectId} setPeriodId={(value) => { setPeriodId(value); setSelected([]); }} setProjectId={(value) => { setProjectId(value); setSelected([]); }} /><ProductionOverview token={token} periodId={periodId} projectId={projectId} /><div className="toolbar"><div className="tab-bar"><button className={stream === "transmitted_batch" ? "active" : ""} onClick={() => { setStream("transmitted_batch"); setSelected([]); setPage(1); }}>Transmitted Batch</button><button className={stream === "shipdet" ? "active" : ""} onClick={() => { setStream("shipdet"); setSelected([]); setPage(1); }}>Shipdet</button></div><label className="search-field">Search current stream<input value={search} onChange={(e) => { setSearch(e.target.value); setPage(1); }} placeholder="Batch, status, country, or work type" /></label><div className="toolbar-actions"><button className="secondary-button compact" onClick={exportCsv}>Export current scope</button>{stream === "shipdet" && canManage && <button className="primary-button compact" disabled={!selected.length} onClick={() => setConfirming(true)}>Confirm selected ({selected.length})</button>}</div></div><section className="state-principle"><b>Business states remain distinct.</b><span>Produced, transmitted, accepted, shipped, billable, and approved for salary are stored independently. Blank transmission status remains unknown.</span></section><StreamRows query={query} stream={stream} page={page} setPage={setPage} selected={selected} setSelected={stream === "shipdet" && canManage ? setSelected : undefined} />{confirming && <Modal title="Confirm inspected Shipdet rows" onClose={() => setConfirming(false)}><form className="stack-form" onSubmit={(event) => { event.preventDefault(); confirm.mutate({ reason: field(new FormData(event.currentTarget), "reason") }); }}><p className="form-lead">Confirmation records that normalized shipment fields were inspected. It does not mark the rows billable or approved for salary.</p><label>Review reason<textarea name="reason" minLength={3} rows={3} required /></label><button className="primary-button" disabled={confirm.isPending}>{confirm.isPending ? "Confirming…" : `Confirm ${selected.length} rows`}</button></form></Modal>}</>;
}

function ProductionSettings({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const queryClient = useQueryClient();
  const allowed = hasPermission(permissions, "production.read");
  const canManage = hasPermission(permissions, "production.manage") || hasPermission(permissions, "admin.manage") || permissions.includes("*");
  const projects = useQuery({ queryKey: ["projects", token], queryFn: () => api.listProjects({ token }), enabled: allowed });
  const [projectId, setProjectId] = useState<number | null>(null);

  useEffect(() => {
    if (!projectId && projects.data?.projects[0]) setProjectId(projects.data.projects[0].id);
  }, [projectId, projects.data]);

  const config = useQuery({
    queryKey: ["production-config", token, projectId],
    queryFn: () => api.getProductionConfiguration({ token, projectId: projectId ?? 0 }),
    enabled: allowed && Boolean(projectId),
  });

  const appSettings = useQuery({
    queryKey: ["app-settings", token],
    queryFn: () => api.getAppSettings({ token }),
    enabled: allowed,
  });

  const [sensitivityThreshold, setSensitivityThreshold] = useState<number>(() => {
    try {
      const local = localStorage.getItem("anomaly_spike_threshold");
      if (local && !isNaN(Number(local))) return Number(local);
    } catch {}
    return 2.0;
  });

  useEffect(() => {
    if (appSettings.data?.anomalySpikeThreshold !== undefined) {
      setSensitivityThreshold(appSettings.data.anomalySpikeThreshold);
    }
  }, [appSettings.data?.anomalySpikeThreshold]);

  const saveThresholdMutation = useMutation({
    mutationFn: (threshold: number) =>
      api.saveAppSettings({ token, anomalySpikeThreshold: threshold }),
    onSuccess: (_, threshold) => {
      localStorage.setItem("anomaly_spike_threshold", String(threshold));
      queryClient.invalidateQueries({ queryKey: ["app-settings"] });
      queryClient.invalidateQueries({ queryKey: ["audit-spike-detector"] });
      notify({ kind: "success", text: `Anomaly spike alert sensitivity updated to ${threshold}× baseline.` });
    },
    onError: (err) => notify({ kind: "error", text: message(err) }),
  });

  const save = useMutation({
    mutationFn: api.saveProductionConfiguration,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["production-config"] });
      notify({ kind: "success", text: "Project production settings saved and audited." });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  if (!allowed) return <Empty title="Production settings are restricted" body="Your role does not include production configuration." />;
  if (!projects.data?.projects.length) return <Empty title="Create a project first" body="Production policies belong to a project." />;

  return (
    <>
      <div className="toolbar">
        <label className="select-field">
          Project
          <select value={projectId ?? ""} onChange={(e) => setProjectId(Number(e.target.value))}>
            {projects.data.projects.map((project) => (
              <option key={project.id} value={project.id}>
                {project.projectId} · {project.name}
              </option>
            ))}
          </select>
        </label>
      </div>

      {config.isPending ? (
        <Loading />
      ) : config.error || !config.data || !projectId ? (
        <Empty title="Production settings could not load" body={message(config.error)} />
      ) : (
        <form
          key={`${projectId}-${JSON.stringify(config.data)}`}
          className="settings-sheet"
          onSubmit={(event) => {
            event.preventDefault();
            const form = new FormData(event.currentTarget);
            const statusMappings: Record<string, string> = {};
            field(form, "statusMappings")
              .split(/\r?\n/)
              .forEach((line) => {
                const [source = "", target = ""] = line.split("=").map((value) => value.trim());
                if (source && target) statusMappings[source] = target;
              });

            save.mutate({
              token,
              projectId,
              doctypeEnabled: form.get("doctypeEnabled") === "on",
              proofEnabled: form.get("proofEnabled") === "on",
              allowKeyVariantDoubleCount: form.get("allowKeyVariantDoubleCount") === "on",
              proofTypes: field(form, "proofTypes")
                .split(/[,\n]/)
                .map((value) => value.trim())
                .filter(Boolean),
              statusMappings,
            });

            if (canManage) {
              saveThresholdMutation.mutate(sensitivityThreshold);
            }
          }}
        >
          <section>
            <p className="overline">Work streams</p>
            <h2>Enable only applicable project work</h2>
            <label className="setting-toggle">
              <input name="doctypeEnabled" type="checkbox" defaultChecked={config.data.doctypeEnabled} disabled={!canManage} />
              <span>
                <b>Doctype enabled</b>
                <small>A disabled project cannot accept or normalize Doctype imports.</small>
              </span>
            </label>
            <label className="setting-toggle">
              <input name="proofEnabled" type="checkbox" defaultChecked={config.data.proofEnabled} disabled={!canManage} />
              <span>
                <b>Proof enabled</b>
                <small>A disabled project cannot accept or normalize Proof imports.</small>
              </span>
            </label>
            <label>
              Approved proof types
              <textarea
                name="proofTypes"
                rows={4}
                defaultValue={config.data.proofTypes.join(", ")}
                placeholder="Full Key Proof, Review Key Proof, Record Proof"
                disabled={!canManage}
              />
            </label>
          </section>

          <section>
            <p className="overline">Financial safeguard</p>
            <h2>Review Key and Full Key treatment</h2>
            <label className="setting-toggle warning-toggle">
              <input
                name="allowKeyVariantDoubleCount"
                type="checkbox"
                defaultChecked={config.data.allowKeyVariantDoubleCount}
                disabled={!canManage}
              />
              <span>
                <b>Permit both variants to be payable</b>
                <small>Off by default. Keep off unless an approved project rule explicitly permits both Review Key and Full Key quantities.</small>
              </span>
            </label>
          </section>

          <section>
            <p className="overline">Status normalization</p>
            <h2>Map source values explicitly</h2>
            <p>
              Enter one mapping per line as <span className="mono">transmission:Yes=transmitted</span> or{" "}
              <span className="mono">acceptance:Accepted=accepted</span>. Allowed targets: transmitted, not_transmitted, accepted, rejected, pending. Blank values always remain unknown.
            </p>
            <textarea
              name="statusMappings"
              rows={8}
              defaultValue={Object.entries(config.data.statusMappings)
                .map(([key, value]) => `${key}=${value}`)
                .join("\n")}
              disabled={!canManage}
            />
          </section>

          {/* New Section: Anomaly Spike Alert Sensitivity Threshold */}
          <section className="anomaly-sensitivity-card">
            <div className="section-head-row">
              <div>
                <p className="overline">Operational Safeguards & Monitoring</p>
                <h2>Anomaly Spike Alert Sensitivity</h2>
              </div>
              <span className="anomaly-threshold-badge">
                Active Threshold: <strong>{sensitivityThreshold.toFixed(1)}×</strong> Baseline
              </span>
            </div>

            <p className="settings-section-desc">
              Define the sensitivity threshold multiplier against the 30-day baseline average for triggering the{" "}
              <strong>'Spike in Anomalies'</strong> dashboard alert. A lower threshold flags early shifts, while a higher threshold minimizes false positives by triggering only on severe deviations.
            </p>

            <div className="anomaly-threshold-controls">
              <div className="anomaly-preset-chips">
                {[
                  { value: 1.5, label: "1.5×", tag: "High Sensitivity", hint: "Triggers on minor surges (+50% above 30d avg)" },
                  { value: 2.0, label: "2.0×", tag: "Standard (Recommended)", hint: "Standard baseline (double the 30d average)" },
                  { value: 2.5, label: "2.5×", tag: "Moderate", hint: "Filters low-level variance (+150% above 30d avg)" },
                  { value: 3.0, label: "3.0×", tag: "High Volume / Conservative", hint: "Triggers only on massive surges (triple the 30d average)" },
                ].map((preset) => (
                  <button
                    key={preset.value}
                    type="button"
                    className={`anomaly-preset-chip ${Math.abs(sensitivityThreshold - preset.value) < 0.05 ? "active" : ""}`}
                    onClick={() => {
                      setSensitivityThreshold(preset.value);
                      if (canManage) {
                        saveThresholdMutation.mutate(preset.value);
                      }
                    }}
                    disabled={!canManage || saveThresholdMutation.isPending}
                    title={preset.hint}
                  >
                    <span className="preset-chip-val">{preset.label}</span>
                    <span className="preset-chip-tag">{preset.tag}</span>
                  </button>
                ))}
              </div>

              <div className="anomaly-custom-slider-row">
                <label className="anomaly-slider-label">
                  <span>Custom Sensitivity Multiplier:</span>
                  <div className="anomaly-slider-inputs">
                    <input
                      type="range"
                      min="1.0"
                      max="4.0"
                      step="0.1"
                      value={sensitivityThreshold}
                      onChange={(e) => setSensitivityThreshold(parseFloat(e.target.value))}
                      disabled={!canManage || saveThresholdMutation.isPending}
                      className="anomaly-range-slider"
                    />
                    <div className="anomaly-num-input-wrap">
                      <input
                        type="number"
                        min="1.0"
                        max="10.0"
                        step="0.1"
                        value={sensitivityThreshold}
                        onChange={(e) => {
                          const val = parseFloat(e.target.value);
                          if (!isNaN(val) && val >= 1.0 && val <= 10.0) {
                            setSensitivityThreshold(val);
                          }
                        }}
                        disabled={!canManage || saveThresholdMutation.isPending}
                        className="anomaly-num-input"
                      />
                      <span className="anomaly-unit">×</span>
                    </div>
                    {canManage && (
                      <button
                        type="button"
                        className="secondary-button compact"
                        onClick={() => saveThresholdMutation.mutate(sensitivityThreshold)}
                        disabled={saveThresholdMutation.isPending}
                      >
                        {saveThresholdMutation.isPending ? "Saving..." : "Apply Threshold"}
                      </button>
                    )}
                  </div>
                </label>
              </div>

              <div className="anomaly-preview-callout">
                <div className="preview-callout-icon">ℹ</div>
                <div className="preview-callout-text">
                  <strong>Trigger Behavior:</strong> The Dashboard Anomaly Detector monitors the rolling 30-day daily average of audit log events. When audit anomalies or total activity on any recent day exceed{" "}
                  <strong>{sensitivityThreshold.toFixed(1)}× the daily average</strong> (minimum 2 events), the{" "}
                  <em>'Spike in Anomalies'</em> alert is triggered with a mini 7-day sparkline trend chart.
                </div>
              </div>
            </div>
          </section>

          {canManage && (
            <button className="primary-button" disabled={save.isPending || saveThresholdMutation.isPending}>
              {save.isPending ? "Saving…" : "Save production settings"}
            </button>
          )}
        </form>
      )}
    </>
  );
}

function ProductionExceptions({ token, permissions }: { token: string; permissions: string[] }) {
  const allowed = hasPermission(permissions, "production.read"); const [periodId, setPeriodId] = useState<number | null>(null); const [projectId, setProjectId] = useState<number | null>(null); const [stream, setStream] = useState(""); const [status, setStatus] = useState<"open" | "resolved">("open"); const [page, setPage] = useState(1);
  const query = useQuery({ queryKey: ["production-exceptions", token, periodId, projectId, stream, status, page], queryFn: () => api.listProductionExceptions({ token, payrollPeriodId: periodId, projectId, stream: stream || null, status, page, pageSize: 25 }), enabled: allowed });
  if (!allowed) return <Empty title="Production exceptions are restricted" body="Your role does not include normalization exceptions." />;
  return <><ProductionScope token={token} periodId={periodId} projectId={projectId} setPeriodId={(value) => { setPeriodId(value); setPage(1); }} setProjectId={(value) => { setProjectId(value); setPage(1); }} /><div className="toolbar"><label>Stream<select value={stream} onChange={(e) => { setStream(e.target.value); setPage(1); }}><option value="">All streams</option><option value="record_details">Record Details</option><option value="doctype">Doctype</option><option value="proof">Proof</option><option value="transmitted_batch">Transmitted Batch</option><option value="shipdet">Shipdet</option></select></label><div className="tab-bar"><button className={status === "open" ? "active" : ""} onClick={() => setStatus("open")}>Open</button><button className={status === "resolved" ? "active" : ""} onClick={() => setStatus("resolved")}>Resolved</button></div></div>{query.isPending ? <Loading /> : query.error ? <Empty title="Exceptions could not load" body={message(query.error)} /> : !query.data?.rows.length ? <Empty title={`No ${status} production exceptions`} body="Exceptions are created only from real normalization outcomes." /> : <div className="exception-list">{query.data.rows.map((row) => <article key={row.id} className={row.severity === "blocking" ? "blocking" : ""}><header><div><Status value={row.severity} /><span className="mono">{row.code}</span></div><small>{row.periodCode} · {row.projectCode} · {humanize(row.stream)}</small></header><h3>{row.explanation}</h3><p><b>Recommended action:</b> {row.recommendedAction}</p><div><span>{row.employeeIdentifier ? `Employee ${row.employeeIdentifier}` : "No employee identifier"}</span><span>{row.batchIdentifier ? `Batch ${row.batchIdentifier}` : "No batch identifier"}</span><span>{row.sourceFile ? `${row.sourceFile}${row.sourceRowNumber ? ` · row ${row.sourceRowNumber}` : ""}` : "No source file"}</span></div></article>)}</div>}<div className="pager"><button className="secondary-button compact" disabled={page === 1} onClick={() => setPage((value) => value - 1)}>Previous</button><span>Page {page} · {query.data?.total ?? 0} exceptions</span><button className="secondary-button compact" disabled={page * 25 >= (query.data?.total ?? 0)} onClick={() => setPage((value) => value + 1)}>Next</button></div></>;
}

function BatchMapping({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const queryClient = useQueryClient();
  const allowed = hasPermission(permissions, "reconciliation.read");
  const canManage = hasPermission(permissions, "reconciliation.manage");

  const [periodId, setPeriodId] = useState<number | null>(null);
  const [projectId, setProjectId] = useState<number | null>(null);
  const [search, setSearch] = useState("");
  const [status, setStatus] = useState("pending"); // Default to pending / requiring approval for immediate workflow convenience
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");

  const [modal, setModal] = useState(false);
  const [editTarget, setEditTarget] = useState<any | null>(null);
  const [viewTarget, setViewTarget] = useState<any | null>(null);
  const [deleteTarget, setDeleteTarget] = useState<any | null>(null);
  const [selected, setSelected] = useState<number[]>([]);
  const [review, setReview] = useState<"approved" | "rejected" | null>(null);

  // CSV Import states
  const [importModal, setImportModal] = useState(false);
  const [csvText, setCsvText] = useState("");
  const [csvFileName, setCsvFileName] = useState("");
  const [csvPreviewRows, setCsvPreviewRows] = useState<string[][]>([]);
  const [importPeriodId, setImportPeriodId] = useState<number | null>(null);
  const [importProjectId, setImportProjectId] = useState<number | null>(null);

  const periods = useQuery({ queryKey: ["periods", token], queryFn: () => api.listPeriods({ token }) });
  const projects = useQuery({ queryKey: ["projects", token], queryFn: () => api.listProjects({ token }) });

  const rows = useQuery({
    queryKey: ["batch-mappings", token, periodId, projectId, status, search, startDate, endDate],
    queryFn: () => api.listBatchMappings({
      token,
      payrollPeriodId: periodId,
      projectId,
      status: status || null,
      search,
      startDate: startDate || null,
      endDate: endDate || null,
    }),
    enabled: allowed,
  });

  const save = useMutation({
    mutationFn: api.saveBatchMapping,
    onSuccess: () => {
      setModal(false);
      setEditTarget(null);
      queryClient.invalidateQueries({ queryKey: ["batch-mappings"] });
      notify({ kind: "success", text: "Batch mapping saved and queued for review." });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const remove = useMutation({
    mutationFn: api.deleteBatchMapping,
    onSuccess: () => {
      setDeleteTarget(null);
      queryClient.invalidateQueries({ queryKey: ["batch-mappings"] });
      notify({ kind: "success", text: "Batch mapping deleted successfully." });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const decide = useMutation({
    mutationFn: api.reviewBatchMappings,
    onSuccess: (data) => {
      setReview(null);
      setSelected([]);
      queryClient.invalidateQueries({ queryKey: ["batch-mappings"] });
      notify({ kind: "success", text: `${data.updated} batch mappings reviewed.` });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const bulkImport = useMutation({
    mutationFn: api.importBatchMappingsCsv,
    onSuccess: (data) => {
      setImportModal(false);
      setCsvText("");
      setCsvFileName("");
      setCsvPreviewRows([]);
      queryClient.invalidateQueries({ queryKey: ["batch-mappings"] });
      if (data.errors && data.errors.length > 0) {
        notify({
          kind: "error",
          text: `Import finished with notices: ${data.created} created, ${data.updated} updated. ${data.errors.length} rows had errors.`,
        });
      } else {
        notify({
          kind: "success",
          text: `CSV Import completed successfully: ${data.created} created, ${data.updated} updated.`,
        });
      }
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const autoMatch = useMutation({
    mutationFn: () =>
      api.autoMatchBatchMappings({
        token,
        payrollPeriodId: periodId,
        projectId,
        autoApprove: false,
      }),
    onSuccess: (data) => {
      queryClient.invalidateQueries({ queryKey: ["batch-mappings"] });
      notify({
        kind: "success",
        text: data.message,
      });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  // Export Filtered Batch Mappings to Excel
  const exportToExcel = () => {
    const list = rows.data?.mappings || [];
    if (!list.length) {
      notify({ kind: "error", text: "No batch mappings found matching current filters to export." });
      return;
    }

    const exportData = list.map((m) => ({
      "Period": m.periodCode,
      "Project": m.projectCode,
      "Internal Batch ID": m.internalBatchId || "—",
      "Client Batch ID": m.clientBatchId || "—",
      "Shipment Batch ID": m.shipmentBatchId || "—",
      "Country": m.country || "—",
      "Match Status": humanize(m.matchStatus),
      "Match Method": humanize(m.matchMethod),
      "Approval Status": humanize(m.approvalStatus),
      "Approved By": m.approvedBy || "—",
      "Approved At": m.approvedAt ? formatDate(m.approvedAt) : "—",
      "Created At": m.createdAt ? formatDate(m.createdAt) : "—",
      "Updated At": m.updatedAt ? formatDate(m.updatedAt) : "—",
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);

    const keys = Object.keys(exportData[0] || {});
    worksheet["!cols"] = keys.map((key) => {
      let maxLen = key.length;
      for (const row of exportData) {
        const valStr = String((row as any)[key] ?? "");
        if (valStr.length > maxLen) maxLen = valStr.length;
      }
      return { wch: Math.min(Math.max(maxLen + 3, 12), 40) };
    });
    worksheet["!autofilter"] = { ref: worksheet["!ref"] ?? "A1:M1" };

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Batch Mappings");

    const pCode = projects.data?.projects.find((p) => p.id === projectId)?.projectId || "all_projects";
    const perCode = periods.data?.periods.find((p) => p.id === periodId)?.periodId || "all_periods";
    const statusPart = status ? `_${status}` : "";
    const fileName = `batch_mappings_${perCode}_${pCode}${statusPart}_${new Date().toISOString().slice(0, 10)}.xlsx`;

    XLSX.writeFile(workbook, fileName);
    notify({ kind: "success", text: `Exported ${list.length} batch mappings to ${fileName}` });
  };

  // Download Blank CSV Template
  const downloadTemplate = () => {
    const headers = ["Internal Batch ID", "Client Batch ID", "Shipment Batch ID", "Project Code", "Country", "Payroll Period", "Approval Status"];
    const pCode = projects.data?.projects.find((p) => p.id === projectId)?.projectId || projects.data?.projects[0]?.projectId || "PRJ-01";
    const perCode = periods.data?.periods.find((p) => p.id === periodId)?.periodId || periods.data?.periods[0]?.periodId || "2026-03";
    const row1 = ["INT-BATCH-001", "CLT-BATCH-101", "SHP-BATCH-201", pCode, "BD", perCode, "pending"];
    const row2 = ["INT-BATCH-002", "CLT-BATCH-102", "SHP-BATCH-202", pCode, "MY", perCode, "approved"];
    const csvContent = [headers.join(","), row1.join(","), row2.join(",")].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `batch_mappings_template_${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const handleFileSelect = (file: File) => {
    setCsvFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      const content = String(e.target?.result || "");
      setCsvText(content);
      const lines = content.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n").filter((l) => l.trim().length > 0);
      const preview = lines.slice(0, 6).map((line) => line.split(",").map((c) => c.replace(/^"|"$/g, "").trim()));
      setCsvPreviewRows(preview);
    };
    reader.readAsText(file);
  };

  if (!allowed) return <Empty title="Batch mapping is restricted" body="Your role does not include reconciliation access." />;

  const columns: Column<any>[] = [
    {
      key: "select",
      header: "Select",
      width: "48px",
      render: (row) => (
        canManage ? (
          <input
            type="checkbox"
            aria-label={`Select mapping ${row.id}`}
            checked={selected.includes(row.id)}
            onChange={(e) => {
              e.stopPropagation();
              setSelected((curr) => e.target.checked ? [...curr, row.id] : curr.filter((id) => id !== row.id));
            }}
          />
        ) : null
      ),
    },
    {
      key: "internalBatchId",
      header: "Internal Batch",
      isImportant: true,
      render: (row) => <span className="mono font-semibold text-emerald-800">{row.internalBatchId || "—"}</span>,
    },
    {
      key: "clientBatchId",
      header: "Client Batch",
      render: (row) => <span className="mono">{row.clientBatchId || "—"}</span>,
    },
    {
      key: "shipmentBatchId",
      header: "Shipment Batch",
      render: (row) => <span className="mono">{row.shipmentBatchId || "—"}</span>,
    },
    {
      key: "projectCode",
      header: "Project",
      isImportant: true,
      render: (row) => <strong className="text-sm">{row.projectCode}</strong>,
    },
    {
      key: "country",
      header: "Country",
      render: (row) => (
        row.country ? (
          <span className="inline-block px-2 py-0.5 text-xs font-bold rounded bg-emerald-100 text-emerald-800 border border-emerald-300">
            {row.country}
          </span>
        ) : (
          <span className="text-dim">—</span>
        )
      ),
    },
    {
      key: "matchStatus",
      header: "Match Status",
      render: (row) => (
        <div>
          {row.matchStatus === "matched" ? (
            <span className="status-chip chip-matched">
              <span className="chip-dot" />
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12" /></svg>
              <span>Matched</span>
            </span>
          ) : row.matchStatus === "review" ? (
            <span className="status-chip chip-pending">
              <span className="chip-dot animate-pulse" />
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><circle cx="12" cy="12" r="10" /><line x1="12" y1="8" x2="12" y2="12" /><line x1="12" y1="16" x2="12.01" y2="16" /></svg>
              <span>Review</span>
            </span>
          ) : (
            <span className="status-chip chip-unmatched">
              <span className="chip-dot" />
              <span>Unmatched</span>
            </span>
          )}
          <small className="block text-xs text-dim mt-1">{humanize(row.matchMethod)}</small>
        </div>
      ),
    },
    {
      key: "approvalStatus",
      header: "Approval Status",
      isImportant: true,
      render: (row) => (
        <div>
          {row.approvalStatus === "approved" ? (
            <span className="status-chip chip-approved">
              <span className="chip-dot" />
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12" /></svg>
              <span>Approved</span>
            </span>
          ) : row.approvalStatus === "rejected" ? (
            <span className="status-chip chip-rejected">
              <span className="chip-dot" />
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" /></svg>
              <span>Rejected</span>
            </span>
          ) : (
            <span className="status-chip chip-pending">
              <span className="chip-dot animate-pulse" />
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><circle cx="12" cy="12" r="10" /><polyline points="12 6 12 12 16 14" /></svg>
              <span>Pending</span>
            </span>
          )}
          <small className="block text-xs text-dim mt-1">
            {row.approvedBy ? `${row.approvedBy} (${formatDate(row.approvedAt)})` : "Awaiting review"}
          </small>
        </div>
      ),
    },
    {
      key: "createdAt",
      header: "Created / Updated",
      render: (row) => (
        <div className="text-xs">
          <span>{formatDate(row.createdAt)}</span>
          {row.updatedAt && row.updatedAt !== row.createdAt && (
            <small className="block text-dim">Upd: {formatDate(row.updatedAt)}</small>
          )}
        </div>
      ),
    },
  ];

  const counts = rows.data?.counts ?? {
    total: 0,
    pending: 0,
    approved: 0,
    rejected: 0,
    matched: 0,
    unmatched: 0,
  };

  return (
    <>
      <ProductionScope
        token={token}
        periodId={periodId}
        projectId={projectId}
        setPeriodId={(value) => { setPeriodId(value); setSelected([]); }}
        setProjectId={(value) => { setProjectId(value); setSelected([]); }}
      />

      {/* Quick Status Monitoring Count Summary Strip */}
      <div className="status-summary-strip" role="region" aria-label="Batch Mapping status summary">
        <button
          type="button"
          className={`status-summary-card ${status === "" ? "active" : ""}`}
          onClick={() => setStatus("")}
          title="Show all batch mappings"
        >
          <div>
            <span className="summary-card-title">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18"/><path d="M9 21V9"/></svg>
              Total Mappings
            </span>
            <div className="summary-card-count">{counts.total}</div>
          </div>
          <span className="summary-card-badge">All</span>
        </button>

        <button
          type="button"
          className={`status-summary-card card-pending ${status === "pending" ? "active" : ""}`}
          onClick={() => setStatus(status === "pending" ? "" : "pending")}
          title="Filter by Pending Review"
        >
          <div>
            <span className="summary-card-title text-amber-800">
              <span className="chip-dot" style={{ background: "#d97706" }} />
              Pending Review
            </span>
            <div className="summary-card-count text-amber-700">{counts.pending}</div>
          </div>
          <span className="status-chip chip-pending text-[10px] py-0.5 px-2">Action Req</span>
        </button>

        <button
          type="button"
          className={`status-summary-card card-approved ${status === "approved" ? "active" : ""}`}
          onClick={() => setStatus(status === "approved" ? "" : "approved")}
          title="Filter by Approved"
        >
          <div>
            <span className="summary-card-title text-emerald-800">
              <span className="chip-dot" style={{ background: "#059669" }} />
              Approved
            </span>
            <div className="summary-card-count text-emerald-700">{counts.approved}</div>
          </div>
          <span className="status-chip chip-approved text-[10px] py-0.5 px-2">Ready</span>
        </button>

        <button
          type="button"
          className={`status-summary-card card-rejected ${status === "rejected" ? "active" : ""}`}
          onClick={() => setStatus(status === "rejected" ? "" : "rejected")}
          title="Filter by Rejected"
        >
          <div>
            <span className="summary-card-title text-rose-800">
              <span className="chip-dot" style={{ background: "#e11d48" }} />
              Rejected
            </span>
            <div className="summary-card-count text-rose-700">{counts.rejected}</div>
          </div>
          <span className="status-chip chip-rejected text-[10px] py-0.5 px-2">Attention</span>
        </button>
      </div>

      {/* Filter and Action Bar */}
      <div className="toolbar flex flex-wrap items-center gap-3 bg-surface p-3 border border-border rounded-lg mt-2">
        <label className="flex items-center gap-2 text-xs font-bold text-dim">
          <span>Search:</span>
          <input
            type="search"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Batch ID, Project, Country..."
            className="border border-border rounded px-2 py-1 text-xs bg-surface"
            title="Search batch identifiers, project, or country"
          />
        </label>

        <label className="flex items-center gap-2 text-xs font-bold text-dim">
          <span>Status:</span>
          <select
            value={status}
            onChange={(e) => setStatus(e.target.value)}
            className="border border-border rounded px-2 py-1 text-sm bg-surface"
          >
            <option value="">All statuses</option>
            <option value="pending">Pending (Requiring Approval)</option>
            <option value="approved">Approved</option>
            <option value="rejected">Rejected</option>
            <option value="matched">Matched</option>
            <option value="review">Review</option>
            <option value="unmatched">Unmatched</option>
          </select>
        </label>

        <label className="flex items-center gap-2 text-xs font-bold text-dim">
          <span>From:</span>
          <input
            type="date"
            value={startDate}
            onChange={(e) => setStartDate(e.target.value)}
            className="border border-border rounded px-2 py-1 text-xs bg-surface"
            title="Filter start date"
          />
        </label>

        <label className="flex items-center gap-2 text-xs font-bold text-dim">
          <span>To:</span>
          <input
            type="date"
            value={endDate}
            onChange={(e) => setEndDate(e.target.value)}
            className="border border-border rounded px-2 py-1 text-xs bg-surface"
            title="Filter end date"
          />
        </label>

        {(startDate || endDate) && (
          <button
            type="button"
            className="ghost-button compact text-xs"
            onClick={() => { setStartDate(""); setEndDate(""); }}
            title="Clear date range filters"
          >
            Clear dates
          </button>
        )}

        <div className="ml-auto flex items-center gap-2">
          <button
            type="button"
            className="secondary-button compact font-semibold text-emerald-800 bg-white hover:bg-emerald-50 border border-emerald-300 shadow-2xs"
            onClick={exportToExcel}
            title="Download filtered batch mappings report as Excel (.xlsx)"
          >
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
              <polyline points="14 2 14 8 20 8" />
              <path d="M8 13h8M8 17h8M10 9h4" />
            </svg>
            <span>Export to Excel</span>
          </button>

          {canManage && (
            <>
              <button
                type="button"
                className="secondary-button compact font-semibold text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-300"
                onClick={() => autoMatch.mutate()}
                disabled={autoMatch.isPending}
                title="Automatically link pending mappings where internal, client, and shipment batch IDs share a common identifier prefix"
              >
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2">
                  <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" />
                  <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" />
                </svg>
                <span>{autoMatch.isPending ? "Auto-matching..." : "Run auto-match"}</span>
              </button>

              <button
                type="button"
                className="secondary-button compact"
                onClick={downloadTemplate}
                title="Download blank CSV template for batch mapping upload"
              >
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" x2="12" y1="15" y2="3"/></svg>
                <span>Blank Template</span>
              </button>

              <button
                type="button"
                className="secondary-button compact"
                onClick={() => {
                  setImportPeriodId(periodId);
                  setImportProjectId(projectId);
                  setImportModal(true);
                }}
                title="Upload CSV for bulk creating or updating batch mappings"
              >
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" x2="12" y1="3" y2="15"/></svg>
                <span>CSV Upload</span>
              </button>

              <button
                type="button"
                className="primary-button compact"
                onClick={() => setModal(true)}
                title="Add single batch mapping"
              >
                <Icon name="plus" />
                <span>Add mapping</span>
              </button>
            </>
          )}
        </div>
      </div>

      {/* Multi-selection Bulk Action Bar */}
      {selected.length > 0 && canManage && (
        <div className="selection-bar flex items-center justify-between bg-emerald-50 border border-emerald-300 p-2.5 rounded-lg my-2">
          <span className="text-sm font-semibold text-emerald-950">
            {selected.length} batch mapping{selected.length === 1 ? "" : "s"} selected for review
          </span>
          <div className="flex items-center gap-2">
            <button
              type="button"
              className="primary-button compact"
              onClick={() => setReview("approved")}
              disabled={decide.isPending}
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="20 6 9 17 4 12"/></svg>
              <span>Approve Selected</span>
            </button>
            <button
              type="button"
              className="secondary-button compact"
              onClick={() => setReview("rejected")}
              disabled={decide.isPending}
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
              <span>Reject Selected</span>
            </button>
            <button
              type="button"
              className="ghost-button compact text-xs"
              onClick={() => setSelected([])}
            >
              Clear
            </button>
          </div>
        </div>
      )}

      {/* Main Enterprise Data Table */}
      {rows.isPending ? (
        <Loading />
      ) : rows.error ? (
        <Empty title="Batch mappings could not load" body={message(rows.error)} />
      ) : (
        <DataTable
          data={rows.data?.mappings || []}
          columns={columns}
          title="Batch Relationships & Mappings"
          subtitle={`Scoped records: ${rows.data?.mappings.length || 0} total mappings`}
          searchPlaceholder="Filter batch IDs, project, or country..."
          onView={(row) => setViewTarget(row)}
          onEdit={canManage ? (row) => setEditTarget(row) : undefined}
          onDelete={canManage ? (row) => setDeleteTarget(row) : undefined}
          onImportClick={canManage ? () => {
            setImportPeriodId(periodId);
            setImportProjectId(projectId);
            setImportModal(true);
          } : undefined}
          onDownloadTemplate={downloadTemplate}
          emptyMessage="No batch mappings match the specified filters."
        />
      )}

      {/* CSV Bulk Upload Modal */}
      {importModal && (
        <Modal title="Upload CSV for Batch Mapping" onClose={() => setImportModal(false)}>
          <div className="stack-form">
            <p className="form-lead text-sm text-dim">
              Bulk-create or update batch relationships, including internal, client, and shipment batch IDs, project code, and country mappings.
            </p>

            <div className="grid grid-cols-2 gap-3 mb-2">
              <label>
                Default Payroll Period
                <select
                  value={importPeriodId ?? ""}
                  onChange={(e) => setImportPeriodId(e.target.value ? Number(e.target.value) : null)}
                >
                  <option value="">Select fallback period</option>
                  {periods.data?.periods.map((p) => (
                    <option key={p.id} value={p.id}>{p.periodId}</option>
                  ))}
                </select>
                <small className="field-help text-xs text-dim">Used if not provided in CSV column.</small>
              </label>

              <label>
                Default Project
                <select
                  value={importProjectId ?? ""}
                  onChange={(e) => setImportProjectId(e.target.value ? Number(e.target.value) : null)}
                >
                  <option value="">Select fallback project</option>
                  {projects.data?.projects.map((p) => (
                    <option key={p.id} value={p.id}>{p.projectId} · {p.name}</option>
                  ))}
                </select>
                <small className="field-help text-xs text-dim">Used if not provided in CSV column.</small>
              </label>
            </div>

            <div className="file-dropzone border-2 border-dashed border-border rounded-lg p-5 text-center bg-surface-alt hover:border-accent transition-colors">
              <input
                type="file"
                accept=".csv,text/csv"
                id="batchMappingCsvInput"
                className="hidden"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) handleFileSelect(f);
                }}
              />
              <label htmlFor="batchMappingCsvInput" className="cursor-pointer block">
                <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="mx-auto mb-2 text-accent">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" x2="12" y1="3" y2="15"/>
                </svg>
                <span className="font-semibold block text-sm">
                  {csvFileName ? `Selected: ${csvFileName}` : "Click to select or drop CSV file"}
                </span>
                <span className="text-xs text-dim block mt-1">Accepts standard .csv with column headers</span>
              </label>
            </div>

            {/* CSV Preview Section */}
            {csvPreviewRows.length > 0 && (
              <div className="csv-preview mt-3">
                <span className="text-xs font-bold text-dim mb-1 block">Preview of detected rows:</span>
                <div className="overflow-x-auto border border-border rounded text-xs">
                  <table className="w-full text-left">
                    <thead>
                      <tr className="bg-surface-alt border-b border-border">
                        {(csvPreviewRows[0] ?? []).map((h, idx) => (
                          <th key={idx} className="p-1.5 font-bold">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {csvPreviewRows.slice(1).map((r, rIdx) => (
                        <tr key={rIdx} className="border-b border-border/50">
                          {r.map((c, cIdx) => (
                            <td key={cIdx} className="p-1.5 font-mono">{c || "—"}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            <div className="flex items-center justify-between mt-4 pt-3 border-t border-border">
              <button
                type="button"
                className="ghost-button"
                onClick={() => setImportModal(false)}
                disabled={bulkImport.isPending}
              >
                Cancel
              </button>
              <div className="flex gap-2">
                <button
                  type="button"
                  className="secondary-button compact"
                  onClick={downloadTemplate}
                >
                  Download Template
                </button>
                <button
                  type="button"
                  className="primary-button"
                  disabled={!csvText.trim() || bulkImport.isPending}
                  onClick={() => {
                    bulkImport.mutate({
                      token,
                      payrollPeriodId: importPeriodId,
                      projectId: importProjectId,
                      csvContent: csvText,
                    });
                  }}
                >
                  {bulkImport.isPending ? "Uploading & Processing..." : "Upload & Update Relationships"}
                </button>
              </div>
            </div>
          </div>
        </Modal>
      )}

      {/* Add Single Mapping Modal */}
      {modal && (
        <Modal title="Add Batch Mapping" onClose={() => setModal(false)}>
          <form
            className="stack-form"
            onSubmit={(event) => {
              event.preventDefault();
              const form = new FormData(event.currentTarget);
              const targetPerId = periodId || Number(field(form, "payrollPeriodId"));
              const targetPrjId = projectId || Number(field(form, "projectId"));
              if (!targetPerId || !targetPrjId) {
                notify({ kind: "error", text: "Please select both a Payroll Period and a Project." });
                return;
              }
              save.mutate({
                token,
                payrollPeriodId: targetPerId,
                projectId: targetPrjId,
                internalBatchId: field(form, "internalBatchId") || null,
                clientBatchId: field(form, "clientBatchId") || null,
                shipmentBatchId: field(form, "shipmentBatchId") || null,
                country: field(form, "country") || null,
                approvalStatus: "pending",
              });
            }}
          >
            <p className="form-lead text-sm text-dim">
              Define the relationship between internal, client, and shipment batch IDs.
            </p>

            {!periodId && (
              <label>
                Payroll Period
                <select name="payrollPeriodId" required>
                  <option value="">Select period</option>
                  {periods.data?.periods.map((p) => (
                    <option key={p.id} value={p.id}>{p.periodId}</option>
                  ))}
                </select>
              </label>
            )}

            {!projectId && (
              <label>
                Project
                <select name="projectId" required>
                  <option value="">Select project</option>
                  {projects.data?.projects.map((p) => (
                    <option key={p.id} value={p.id}>{p.projectId} · {p.name}</option>
                  ))}
                </select>
              </label>
            )}

            <label>
              Internal Batch ID
              <input name="internalBatchId" placeholder="e.g. INT-00123" />
            </label>
            <label>
              Client Batch ID
              <input name="clientBatchId" placeholder="e.g. CLT-99881" />
            </label>
            <label>
              Shipment Batch ID
              <input name="shipmentBatchId" placeholder="e.g. SHP-44551" />
            </label>
            <label>
              Country Code
              <input name="country" placeholder="e.g. BD or MY" maxLength={10} />
            </label>

            <div className="flex justify-end gap-2 mt-4 pt-3 border-t border-border">
              <button type="button" className="ghost-button" onClick={() => setModal(false)}>Cancel</button>
              <button className="primary-button" disabled={save.isPending}>
                {save.isPending ? "Saving..." : "Save for Review"}
              </button>
            </div>
          </form>
        </Modal>
      )}

      {/* Edit Mapping Modal */}
      {editTarget && (
        <Modal title="Edit Batch Mapping" onClose={() => setEditTarget(null)}>
          <form
            className="stack-form"
            onSubmit={(event) => {
              event.preventDefault();
              const form = new FormData(event.currentTarget);
              save.mutate({
                token,
                id: editTarget.id,
                payrollPeriodId: editTarget.payrollPeriodId,
                projectId: editTarget.projectId,
                internalBatchId: field(form, "internalBatchId") || null,
                clientBatchId: field(form, "clientBatchId") || null,
                shipmentBatchId: field(form, "shipmentBatchId") || null,
                country: field(form, "country") || null,
                approvalStatus: field(form, "approvalStatus") as "pending" | "approved" | "rejected",
              });
            }}
          >
            <div className="text-xs text-dim mb-2 flex gap-4">
              <span><b>Period:</b> {editTarget.periodCode}</span>
              <span><b>Project:</b> {editTarget.projectCode}</span>
            </div>

            <label>
              Internal Batch ID
              <input name="internalBatchId" defaultValue={editTarget.internalBatchId || ""} />
            </label>
            <label>
              Client Batch ID
              <input name="clientBatchId" defaultValue={editTarget.clientBatchId || ""} />
            </label>
            <label>
              Shipment Batch ID
              <input name="shipmentBatchId" defaultValue={editTarget.shipmentBatchId || ""} />
            </label>
            <label>
              Country Code
              <input name="country" defaultValue={editTarget.country || ""} maxLength={10} />
            </label>
            <label>
              Approval Status
              <select name="approvalStatus" defaultValue={editTarget.approvalStatus || "pending"}>
                <option value="pending">Pending Review</option>
                <option value="approved">Approved</option>
                <option value="rejected">Rejected</option>
              </select>
            </label>

            <div className="flex justify-end gap-2 mt-4 pt-3 border-t border-border">
              <button type="button" className="ghost-button" onClick={() => setEditTarget(null)}>Cancel</button>
              <button className="primary-button" disabled={save.isPending}>
                {save.isPending ? "Saving..." : "Update Mapping"}
              </button>
            </div>
          </form>
        </Modal>
      )}

      {/* View Details Modal */}
      {viewTarget && (
        <Modal title="Batch Mapping Details" onClose={() => setViewTarget(null)}>
          <div className="stack-form text-sm">
            <div className="grid grid-cols-2 gap-3 pb-3 border-b border-border">
              <div>
                <small className="text-dim block">Mapping ID</small>
                <strong>#{viewTarget.id}</strong>
              </div>
              <div>
                <small className="text-dim block">Payroll Period</small>
                <strong>{viewTarget.periodCode}</strong>
              </div>
              <div>
                <small className="text-dim block">Project</small>
                <strong>{viewTarget.projectCode}</strong>
              </div>
              <div>
                <small className="text-dim block">Country</small>
                <strong>{viewTarget.country || "—"}</strong>
              </div>
            </div>

            <div className="grid grid-cols-3 gap-2 py-2 border-b border-border">
              <div>
                <small className="text-dim block">Internal Batch</small>
                <span className="mono font-semibold">{viewTarget.internalBatchId || "—"}</span>
              </div>
              <div>
                <small className="text-dim block">Client Batch</small>
                <span className="mono">{viewTarget.clientBatchId || "—"}</span>
              </div>
              <div>
                <small className="text-dim block">Shipment Batch</small>
                <span className="mono">{viewTarget.shipmentBatchId || "—"}</span>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-3 py-2 border-b border-border">
              <div>
                <small className="text-dim block">Match Status / Method</small>
                <div className="flex items-center gap-2 mt-1">
                  <Status value={viewTarget.matchStatus} />
                  <span className="text-xs text-dim">({humanize(viewTarget.matchMethod)})</span>
                </div>
              </div>
              <div>
                <small className="text-dim block">Approval Status</small>
                <div className="mt-1">
                  <Status value={viewTarget.approvalStatus} />
                </div>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-3 text-xs text-dim pt-2">
              <div>
                <span>Reviewed By: </span>
                <strong>{viewTarget.approvedBy || "Awaiting review"}</strong>
              </div>
              <div>
                <span>Reviewed At: </span>
                <strong>{formatDate(viewTarget.approvedAt)}</strong>
              </div>
              <div>
                <span>Created At: </span>
                <strong>{formatDate(viewTarget.createdAt)}</strong>
              </div>
              <div>
                <span>Last Updated: </span>
                <strong>{formatDate(viewTarget.updatedAt)}</strong>
              </div>
            </div>

            <div className="flex justify-end mt-4">
              <button type="button" className="secondary-button" onClick={() => setViewTarget(null)}>Close</button>
            </div>
          </div>
        </Modal>
      )}

      {/* Delete Confirmation Dialog */}
      <ConfirmDialog
        isOpen={Boolean(deleteTarget)}
        title="Delete Batch Mapping"
        message={`Are you sure you want to delete the mapping for internal batch "${deleteTarget?.internalBatchId || deleteTarget?.id}"? This will permanently remove the relationship.`}
        confirmLabel="Delete Mapping"
        cancelLabel="Cancel"
        isDangerous={true}
        isLoading={remove.isPending}
        onConfirm={() => {
          if (deleteTarget) {
            remove.mutate({ token, id: deleteTarget.id, reason: "Deleted from batch mapping management" });
          }
        }}
        onCancel={() => setDeleteTarget(null)}
      />

      {/* Review Selected Modal */}
      {review && (
        <Modal title={`${review === "approved" ? "Approve" : "Reject"} Selected Mappings`} onClose={() => setReview(null)}>
          <form
            className="stack-form"
            onSubmit={(event) => {
              event.preventDefault();
              decide.mutate({
                token,
                ids: selected,
                decision: review,
                reason: field(new FormData(event.currentTarget), "reason"),
              });
            }}
          >
            <p className="form-lead text-sm text-dim">
              {review === "approved"
                ? `Approving ${selected.length} batch mapping(s) connects these batches for downstream reconciliation and salary calculations.`
                : `Rejecting ${selected.length} batch mapping(s) marks them as unmatched and excludes them from payable credit.`}
            </p>
            <label>
              Audit Review Reason
              <textarea name="reason" minLength={3} rows={3} required placeholder="State reason for audit history..." />
            </label>
            <div className="flex justify-end gap-2 mt-4">
              <button type="button" className="ghost-button" onClick={() => setReview(null)}>Cancel</button>
              <button className="primary-button" disabled={decide.isPending}>
                {decide.isPending ? "Processing..." : (review === "approved" ? "Approve Selected" : "Reject Selected")}
              </button>
            </div>
          </form>
        </Modal>
      )}
    </>
  );
}

function ReconciliationCenter({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const queryClient = useQueryClient(); const allowed = hasPermission(permissions, "reconciliation.read"); const canManage = hasPermission(permissions, "reconciliation.manage"); const canApprove = hasPermission(permissions, "reconciliation.approve"); const canExport = hasPermission(permissions, "reconciliation.export"); const [periodId, setPeriodId] = useState<number | null>(null); const [projectId, setProjectId] = useState<number | null>(null); const [tab, setTab] = useState<"report" | "policies">("report"); const [search, setSearch] = useState(""); const [unit, setUnit] = useState(""); const [exceptionStatus, setExceptionStatus] = useState(""); const [sortBy, setSortBy] = useState<"batchKey" | "employeeIdentifier" | "workType" | "entryQuantity" | "employeeAllocatedQuantity" | "quantityDifference">("batchKey"); const [direction, setDirection] = useState<"asc" | "desc">("asc"); const [page, setPage] = useState(1); const [approvalModal, setApprovalModal] = useState(false); const [manualRow, setManualRow] = useState<null | { batchKey: string; employeeId: number; employeeIdentifier: string | null; workType: string; categoryCode: string | null; quantityUnit: "record" | "character" | "image" | "document" }>(null);
  const report = useQuery({ queryKey: ["reconciliation-report", token, periodId, projectId, search, unit, exceptionStatus, sortBy, direction, page], queryFn: () => api.listReconciliationReport({ token, payrollPeriodId: periodId, projectId, search, country: "", workType: "", unit, exceptionStatus, resolutionStatus: "", sortBy, direction, page, pageSize: 25 }), enabled: allowed });
  const policies = useQuery({ queryKey: ["payable-policies", token, projectId], queryFn: () => api.getPayablePolicies({ token, projectId: projectId ?? 0 }), enabled: allowed && Boolean(projectId) });
  const run = useMutation({ mutationFn: api.runReconciliation, onSuccess: (data) => { queryClient.invalidateQueries({ queryKey: ["reconciliation-report"] }); queryClient.invalidateQueries({ queryKey: ["reconciliation-exceptions"] }); queryClient.invalidateQueries({ queryKey: ["batch-mappings"] }); notify({ kind: data.blockingCount ? "error" : "success", text: `${data.runId}: ${data.resultCount} result rows, ${data.warningCount} warnings, ${data.blockingCount} blocking exceptions.` }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  const approve = useMutation({ mutationFn: api.approveReconciliation, onSuccess: (data) => { setApprovalModal(false); queryClient.invalidateQueries({ queryKey: ["reconciliation-report"] }); notify({ kind: "success", text: `${data.approved} reconciliation rows approved for salary calculation.` }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  const savePolicy = useMutation({ mutationFn: api.savePayablePolicy, onSuccess: () => { queryClient.invalidateQueries({ queryKey: ["payable-policies"] }); notify({ kind: "success", text: "Payable quantity policy saved." }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  const saveManual = useMutation({ mutationFn: api.saveManualAllocation, onSuccess: () => { setManualRow(null); notify({ kind: "success", text: "Manual allocation approved and saved. Run reconciliation again to apply it." }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  if (!allowed) return <Empty title="Reconciliation is restricted" body="Your role does not include reconciliation access." />;
  const exportWorkbook = () => { if (!periodId || !projectId) return; void api.exportReconciliationWorkbook({ token, payrollPeriodId: periodId, projectId }).then((data) => { downloadBase64(data.filename, data.base64, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); notify({ kind: "success", text: `${data.rowCount} reconciliation rows exported to Excel.` }); }).catch((error) => notify({ kind: "error", text: message(error) })); };
  return <><ProductionScope token={token} periodId={periodId} projectId={projectId} setPeriodId={(value) => { setPeriodId(value); setPage(1); }} setProjectId={(value) => { setProjectId(value); setPage(1); }} /><div className="stage-principles"><span><b>Produced</b>Source work</span><i>→</i><span><b>Transmitted</b>Sent onward</span><i>→</i><span><b>Accepted</b>Client state</span><i>→</i><span><b>Shipped</b>Confirmed Shipdet</span><i>→</i><span><b>Billable</b>Authorized source</span><i>→</i><span><b>Salary</b>Approved allocation</span></div><div className="tab-bar" role="tablist"><button role="tab" aria-selected={tab === "report"} className={tab === "report" ? "active" : ""} onClick={() => setTab("report")}>Reconciliation report</button><button role="tab" aria-selected={tab === "policies"} className={tab === "policies" ? "active" : ""} onClick={() => setTab("policies")}>Payable policies</button></div>{tab === "report" ? <><div className="toolbar"><label className="search-field">Search report<input value={search} onChange={(e) => { setSearch(e.target.value); setPage(1); }} placeholder="Batch, employee, work type, category" /></label><label>Unit<select value={unit} onChange={(e) => { setUnit(e.target.value); setPage(1); }}><option value="">All compatible units</option><option value="record">Record</option><option value="character">Character</option><option value="image">Image</option><option value="document">Document</option></select></label><label>Exception<select value={exceptionStatus} onChange={(e) => { setExceptionStatus(e.target.value); setPage(1); }}><option value="">All statuses</option><option value="clear">Clear</option><option value="warning">Warning</option><option value="blocking">Blocking</option></select></label><label>Sort<select value={sortBy} onChange={(e) => setSortBy(e.target.value as typeof sortBy)}><option value="batchKey">Batch</option><option value="employeeIdentifier">Employee</option><option value="workType">Work type</option><option value="entryQuantity">Entry quantity</option><option value="employeeAllocatedQuantity">Allocated quantity</option><option value="quantityDifference">Difference</option></select></label><button className="secondary-button compact" onClick={() => setDirection((value) => value === "asc" ? "desc" : "asc")}>{direction === "asc" ? "Ascending" : "Descending"}</button></div><div className="run-command"><div><p className="overline">Controlled computation</p><h2>{report.data?.latestRun ? report.data.latestRun.runId : "No reconciliation run"}</h2><p>{report.data?.latestRun ? `${report.data.latestRun.resultCount} results · ${report.data.latestRun.warningCount} warnings · ${report.data.latestRun.blockingCount} blocking` : "Select one payroll period and project, then run the server-side engine."}</p></div><div className="toolbar-actions">{canExport && <button className="secondary-button compact" disabled={!periodId || !projectId || !report.data?.latestRun} onClick={exportWorkbook}>Export Excel</button>}{canManage && <button className="secondary-button compact" disabled={!periodId || !projectId || run.isPending} onClick={() => periodId && projectId && run.mutate({ token, payrollPeriodId: periodId, projectId })}>{run.isPending ? "Reconciling…" : "Run reconciliation"}</button>}{canApprove && report.data?.latestRun?.status === "completed" && <button className="primary-button compact" disabled={report.data.latestRun.blockingCount > 0} onClick={() => setApprovalModal(true)}>Approve for salary</button>}</div></div>{report.isPending ? <Loading /> : report.error ? <Empty title="Reconciliation report could not load" body={message(report.error)} /> : !report.data?.rows.length ? <Empty title="No reconciliation results" body="Committed production rows become reportable after a reconciliation run." /> : <div className="reconciliation-list">{report.data.rows.map((row) => <article key={row.id}><header><div><strong className="mono">{row.batchKey}</strong><span>{row.employeeIdentifier || "Unattributed"}{row.employeeName ? ` · ${row.employeeName}` : ""}</span></div><div><Status value={row.exceptionStatus} /><Status value={row.salaryApprovalStatus} /></div></header><div className="recon-identity"><span>{row.periodCode}</span><span>{row.projectCode}</span><span>{row.country || "No country"}</span><span>{row.workType}</span><span>{row.categoryName || row.categoryCode || "No category"}</span><span>{row.quantityUnit}</span></div><div className="quantity-track"><span><small>Entry</small><b>{row.entryQuantity}</b></span><span><small>Transmitted</small><b>{row.transmittedQuantity}</b></span><span><small>Accepted</small><b>{row.acceptedQuantity}</b></span><span><small>Shipdet</small><b>{row.shipdetQuantity}</b></span><span className="authorized"><small>Authorized cap</small><b>{row.authorizedPayableQuantity}</b></span><span className="allocated"><small>Employee allocated</small><b>{row.employeeAllocatedQuantity}</b></span><span><small>Non-payable</small><b>{row.nonPayableQuantity}</b></span><span><small>Difference</small><b>{row.quantityDifference}</b></span></div><footer><span>{humanize(row.payableSource)} · {humanize(row.allocationPolicy)} · {humanize(row.matchStatus)}</span>{canManage && row.allocationPolicy === "manual" && row.employeeId && <button className="text-button" onClick={() => setManualRow({ batchKey: row.batchKey, employeeId: row.employeeId!, employeeIdentifier: row.employeeIdentifier, workType: row.workType, categoryCode: row.categoryCode, quantityUnit: row.quantityUnit as "record" | "character" | "image" | "document" })}>Enter manual allocation</button>}</footer></article>)}</div>}<div className="pager"><button className="secondary-button compact" disabled={page === 1} onClick={() => setPage((value) => value - 1)}>Previous</button><span>Page {page} · {report.data?.total ?? 0} rows</span><button className="secondary-button compact" disabled={page * 25 >= (report.data?.total ?? 0)} onClick={() => setPage((value) => value + 1)}>Next</button></div></> : <>{!projectId ? <Empty title="Select a project" body="Payable policies are configured independently for each project and work type." /> : policies.isPending ? <Loading /> : policies.error ? <Empty title="Policies could not load" body={message(policies.error)} /> : <div className="policy-board"><div className="policy-note"><h2>Produced quantity remains the default.</h2><p>Without an explicit policy, produced quantity is allocated directly and stage differences remain warnings. Select a stricter source only when the project rule requires it.</p></div>{policies.data?.workTypes.map((work) => { const current = policies.data.policies.find((row) => row.workType === work.code && row.quantityUnit === work.unit); return <form key={`${work.code}-${work.unit}-${current?.id ?? 0}`} className="policy-row" onSubmit={(event) => { event.preventDefault(); const form = new FormData(event.currentTarget); savePolicy.mutate({ token, projectId, workType: work.code, quantityUnit: field(form, "quantityUnit") as "record" | "character" | "image" | "document", payableSource: field(form, "payableSource") as "produced" | "confirmed_shipped" | "client_billable" | "accepted_transmitted", allocationPolicy: field(form, "allocationPolicy") as "direct_employee" | "record_attribution" | "category_proportional" | "manual", statusRequirement: field(form, "statusRequirement") as "none" | "transmitted" | "accepted" | "confirmed_shipped" | "client_billable", active: true }); }}><div><strong>{work.name}</strong><span className="mono">{work.code}</span></div><label>Unit<select name="quantityUnit" defaultValue={current?.quantityUnit ?? (work.unit === "batch" || work.unit === "fixed" ? "record" : work.unit)}><option value="record">Record</option><option value="character">Character</option><option value="image">Image</option><option value="document">Document</option></select></label><label>Payable source<select name="payableSource" defaultValue={current?.payableSource ?? "produced"}><option value="produced">Produced quantity</option><option value="confirmed_shipped">Confirmed shipped quantity</option><option value="client_billable">Client-confirmed billable quantity</option><option value="accepted_transmitted">Accepted transmitted quantity</option></select></label><label>Allocation policy<select name="allocationPolicy" defaultValue={current?.allocationPolicy ?? "direct_employee"}><option value="direct_employee">Direct employee-level</option><option value="record_attribution">Record-level attribution</option><option value="category_proportional">Category proportional</option><option value="manual">Approved manual allocation</option></select></label><label>Status requirement<select name="statusRequirement" defaultValue={current?.statusRequirement ?? "none"}><option value="none">None</option><option value="transmitted">Transmitted</option><option value="accepted">Accepted</option><option value="confirmed_shipped">Confirmed shipped</option><option value="client_billable">Client billable</option></select></label>{canManage && <button className="primary-button compact" disabled={savePolicy.isPending}>Save</button>}</form>; })}</div>}</>}{approvalModal && report.data?.latestRun && <Modal title="Approve reconciliation for salary" onClose={() => setApprovalModal(false)}><form className="stack-form" onSubmit={(event) => { event.preventDefault(); approve.mutate({ token, reconciliationRunId: report.data!.latestRun!.id, reason: field(new FormData(event.currentTarget), "reason") }); }}><p className="form-lead">Approval marks every result in this run as approved-for-salary. Unresolved blocking exceptions are rechecked by the server.</p><label>Approval reason<textarea name="reason" minLength={3} rows={3} required /></label><button className="primary-button" disabled={approve.isPending}>Approve reconciliation</button></form></Modal>}{manualRow && periodId && projectId && <Modal title="Approved manual allocation" onClose={() => setManualRow(null)}><form className="stack-form" onSubmit={(event) => { event.preventDefault(); const form = new FormData(event.currentTarget); saveManual.mutate({ token, payrollPeriodId: periodId, projectId, batchKey: manualRow.batchKey, employeeId: manualRow.employeeId, workType: manualRow.workType, categoryCode: manualRow.categoryCode, quantityUnit: manualRow.quantityUnit, allocatedQuantity: Number(field(form, "allocatedQuantity")), reason: field(form, "reason") }); }}><p className="form-lead">{manualRow.batchKey} · {manualRow.employeeIdentifier || `Employee record ${manualRow.employeeId}`} · {manualRow.workType}</p><label>Allocated quantity<input name="allocatedQuantity" type="number" min="0" step="0.000001" required /></label><label>Reviewer reason<textarea name="reason" minLength={3} rows={3} required /></label><button className="primary-button" disabled={saveManual.isPending}>Approve allocation input</button></form></Modal>}</>;
}

function ReconciliationExceptions({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const queryClient = useQueryClient(); const allowed = hasPermission(permissions, "reconciliation.read"); const canManage = hasPermission(permissions, "reconciliation.manage"); const [periodId, setPeriodId] = useState<number | null>(null); const [projectId, setProjectId] = useState<number | null>(null); const [severity, setSeverity] = useState(""); const [status, setStatus] = useState<"open" | "resolved">("open"); const [search, setSearch] = useState(""); const [page, setPage] = useState(1); const [selected, setSelected] = useState<number[]>([]); const [modal, setModal] = useState(false);
  const query = useQuery({ queryKey: ["reconciliation-exceptions", token, periodId, projectId, severity, status, search, page], queryFn: () => api.listReconciliationExceptions({ token, payrollPeriodId: periodId, projectId, severity, status, search, page, pageSize: 25 }), enabled: allowed }); const resolve = useMutation({ mutationFn: api.resolveReconciliationExceptions, onSuccess: (data) => { setModal(false); setSelected([]); queryClient.invalidateQueries({ queryKey: ["reconciliation-exceptions"] }); queryClient.invalidateQueries({ queryKey: ["reconciliation-report"] }); notify({ kind: "success", text: `${data.resolved} exceptions resolved with an audit reason.` }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  if (!allowed) return <Empty title="Reconciliation exceptions are restricted" body="Your role does not include reconciliation access." />;
  return <><ProductionScope token={token} periodId={periodId} projectId={projectId} setPeriodId={(value) => { setPeriodId(value); setSelected([]); setPage(1); }} setProjectId={(value) => { setProjectId(value); setSelected([]); setPage(1); }} /><div className="exception-summary"><strong>{query.data?.blockingOpen ?? 0}</strong><div><h2>Unresolved blocking exceptions</h2><p>Final reconciliation approval remains locked until this count reaches zero.</p></div></div><div className="toolbar"><label className="search-field">Search exceptions<input value={search} onChange={(e) => { setSearch(e.target.value); setPage(1); }} placeholder="Batch, employee, type, explanation" /></label><label>Severity<select value={severity} onChange={(e) => { setSeverity(e.target.value); setPage(1); }}><option value="">All severities</option><option value="blocking">Blocking</option><option value="warning">Warning</option></select></label><div className="tab-bar"><button className={status === "open" ? "active" : ""} onClick={() => { setStatus("open"); setPage(1); }}>Open</button><button className={status === "resolved" ? "active" : ""} onClick={() => { setStatus("resolved"); setPage(1); }}>Resolved</button></div></div>{selected.length > 0 && canManage && status === "open" && <div className="selection-bar"><span>{selected.length} selected for controlled resolution</span><button className="primary-button compact" onClick={() => setModal(true)}>Resolve selected</button></div>}{query.isPending ? <Loading /> : query.error ? <Empty title="Exceptions could not load" body={message(query.error)} /> : !query.data?.rows.length ? <Empty title={`No ${status} reconciliation exceptions`} body="Exceptions are produced only by completed reconciliation runs." /> : <div className="exception-list">{query.data.rows.map((row) => <article key={row.id} className={row.severity === "blocking" ? "blocking" : ""}><header><div>{canManage && status === "open" && <input type="checkbox" aria-label={`Select exception ${row.id}`} checked={selected.includes(row.id)} onChange={(e) => setSelected((current) => e.target.checked ? [...current, row.id] : current.filter((id) => id !== row.id))} />}<Status value={row.severity} /><span className="mono">{humanize(row.exceptionType)}</span></div><small>{row.periodCode} · {row.projectCode}</small></header><h3>{row.explanation}</h3><p><b>Recommended action:</b> {row.recommendedAction}</p><div><span>{row.batchKey ? `Batch ${row.batchKey}` : "No batch"}</span><span>{row.employeeIdentifier ? `Employee ${row.employeeIdentifier}` : "No employee"}</span><span>{row.sourceFile || "No source file"}</span>{row.resolvedAt && <span>Resolved {formatDate(row.resolvedAt)} by {row.resolvedBy}</span>}</div>{row.resolutionReason && <p className="resolution-note"><b>Resolution:</b> {row.resolutionReason}</p>}</article>)}</div>}<div className="pager"><button className="secondary-button compact" disabled={page === 1} onClick={() => setPage((value) => value - 1)}>Previous</button><span>Page {page} · {query.data?.total ?? 0} exceptions</span><button className="secondary-button compact" disabled={page * 25 >= (query.data?.total ?? 0)} onClick={() => setPage((value) => value + 1)}>Next</button></div>{modal && <Modal title="Resolve selected exceptions" onClose={() => setModal(false)}><form className="stack-form" onSubmit={(event) => { event.preventDefault(); resolve.mutate({ token, ids: selected, reason: field(new FormData(event.currentTarget), "reason") }); }}><p className="form-lead">Resolution records the reviewer, time, and reason. It does not alter source quantities; rerun reconciliation after correcting source data.</p><label>Resolution reason<textarea name="reason" minLength={3} rows={4} required /></label><button className="primary-button" disabled={resolve.isPending}>Resolve {selected.length} exceptions</button></form></Modal>}</>;
}


function SalaryRunLedger({ token, runId, permissions, notify, employeeId }: { token: string; runId: number; permissions: string[]; notify: (notice: Notice) => void; employeeId?: number }) {
  const queryClient = useQueryClient(); const details = useQuery({ queryKey: ["salary-run", token, runId], queryFn: () => api.getSalaryRun({ token, salaryRunId: runId }) }); const [adjusting, setAdjusting] = useState<null | { employeeId: number; reconciliationResultId: number; label: string }>(null);
  const adjust = useMutation({ mutationFn: api.applyManualAdjustment, onSuccess: (data) => { setAdjusting(null); queryClient.invalidateQueries({ queryKey: ["salary-runs"] }); queryClient.invalidateQueries({ queryKey: ["salary-run"] }); queryClient.invalidateQueries({ queryKey: ["employee-salary"] }); notify({ kind: data.status === "completed" ? "success" : "error", text: `Adjustment recorded. Calculation version ${data.version} ${data.status} with ${data.exceptionCount} exceptions.` }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  if (details.isPending) return <Loading />; if (details.error || !details.data) return <Empty title="Salary run could not load" body={message(details.error)} />;
  const rows = employeeId ? details.data.components.filter((row) => row.employeeId === employeeId) : details.data.components;
  return <div className="salary-run-detail"><section className="salary-total-strip"><div><small>Calculation version</small><strong>V{details.data.run.version}</strong></div><div><small>Gross</small><strong>{details.data.run.currency} {details.data.run.grossTotal}</strong></div><div><small>Adjustments</small><strong>{details.data.run.currency} {details.data.run.adjustmentTotal}</strong></div><div><small>Final</small><strong>{details.data.run.currency} {details.data.run.finalTotal}</strong></div></section>{details.data.exceptions.length > 0 && <div className="salary-exceptions"><h3>Blocking calculation exceptions</h3>{details.data.exceptions.map((item) => <article key={item.id}><Status value="blocking" /><div><b>{humanize(item.type)}</b><p>{item.explanation}</p><small>{item.recommendedAction}</small></div></article>)}</div>}{rows.length > 0 ? <div className="salary-components">{rows.map((row) => <article key={row.id}><header><div><strong>{row.employeeIdentifier} · {row.employeeName}</strong><span>{row.workType}{row.category ? ` · ${row.category}` : ""} · {row.country}</span></div><strong>{details.data.run.currency} {row.finalAmount}</strong></header><div className="salary-formula"><span><small>Entry</small><b>{row.entryQuantity}</b></span><span><small>Transmitted</small><b>{row.transmittedQuantity}</b></span><span><small>Accepted</small><b>{row.acceptedQuantity}</b></span><span><small>Shipped</small><b>{row.shippedQuantity}</b></span><span className="salary-payable"><small>Payable {row.quantityUnit}</small><b>{row.approvedPayableQuantity}</b></span><span><small>Rate</small><b>{row.appliedRate}</b></span><span><small>Multiplier</small><b>{row.multiplier}</b></span><span><small>Gross</small><b>{row.grossAmount}</b></span><span><small>Adjustment</small><b>{row.approvedAdjustment}</b></span></div><footer><details><summary>Trace source and rule</summary><p>Rate rule #{row.rateRuleId} · formula {row.formulaRuleId} · {humanize(row.quantitySource)} quantity · {humanize(row.allocationMethod)} allocation</p><code>{row.sourceTraceJson}</code></details>{hasPermission(permissions, "salary.adjust") && details.data.run.status === "completed" && <button className="text-button" onClick={() => setAdjusting({ employeeId: row.employeeId, reconciliationResultId: row.reconciliationResultId, label: `${row.employeeIdentifier} · ${row.workType}` })}>Add adjustment</button>}</footer></article>)}</div> : !details.data.exceptions.length && <Empty title="No components in this view" body="Only completed, traceable salary components appear here." />}{adjusting && <Modal title="Authorize manual adjustment" onClose={() => setAdjusting(null)}><form className="stack-form" onSubmit={(event) => { event.preventDefault(); const form = new FormData(event.currentTarget); adjust.mutate({ token, payrollPeriodId: details.data!.run.payrollPeriodId, projectId: details.data!.run.projectId, employeeId: adjusting.employeeId, reconciliationResultId: adjusting.reconciliationResultId, amount: field(form, "amount"), reason: field(form, "reason") }); }}><p className="form-lead">{adjusting.label}. A new immutable calculation version will be created.</p><label>Adjustment amount<input name="amount" inputMode="decimal" placeholder="-125.00 or 250.00" required /></label><label>Authorization reason<textarea name="reason" minLength={3} rows={3} required /></label><button className="primary-button" disabled={adjust.isPending}>Authorize and recalculate</button></form></Modal>}</div>;
}

function SalaryCalculation({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const queryClient = useQueryClient(); const allowed = hasPermission(permissions, "salary.read"); const [periodId, setPeriodId] = useState<number | null>(null); const [projectId, setProjectId] = useState<number | null>(null); const [runId, setRunId] = useState<number | null>(null); const [inputModal, setInputModal] = useState(false);
  const runs = useQuery({ queryKey: ["salary-runs", token, periodId, projectId], queryFn: () => api.listSalaryRuns({ token, payrollPeriodId: periodId, projectId }), enabled: allowed });
  const setup = useQuery({ queryKey: ["salary-setup", token, periodId, projectId], queryFn: () => api.getSalarySetup({ token, payrollPeriodId: periodId ?? 0, projectId: projectId ?? 0 }), enabled: allowed && Boolean(periodId && projectId) });
  const calculate = useMutation({ mutationFn: api.runSalaryCalculation, onSuccess: (data) => { queryClient.invalidateQueries({ queryKey: ["salary-runs"] }); queryClient.invalidateQueries({ queryKey: ["periods"] }); notify({ kind: data.status === "completed" ? "success" : "error", text: `${data.runId}: ${data.componentCount} traceable components, ${data.exceptionCount} blocking exceptions.` }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  const savePolicy = useMutation({ mutationFn: api.saveSalaryPolicy, onSuccess: () => { queryClient.invalidateQueries({ queryKey: ["salary-setup"] }); notify({ kind: "success", text: "Salary rounding policy saved and audited." }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  const saveInput = useMutation({ mutationFn: api.saveSalaryInput, onSuccess: () => { setInputModal(false); queryClient.invalidateQueries({ queryKey: ["salary-setup"] }); notify({ kind: "success", text: "Authorized average-hours input saved." }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  if (!allowed) return <Empty title="Salary calculation is restricted" body="Your role does not include salary access." />;
  return <><ProductionScope token={token} periodId={periodId} projectId={projectId} setPeriodId={(value) => { setPeriodId(value); setRunId(null); }} setProjectId={(value) => { setProjectId(value); setRunId(null); }} /><section className="salary-command"><div><p className="overline">Server-side calculation</p><h2>Approved payable quantity × one effective rate</h2><p>Every run is immutable. Missing rates, ambiguous precedence, incomplete slabs, unit mismatches, and unapproved quantities block the run.</p></div><div className="toolbar-actions">{hasPermission(permissions, "salary.adjust") && <button className="secondary-button compact" disabled={!periodId || !projectId} onClick={() => setInputModal(true)}>Authorized hours</button>}{hasPermission(permissions, "salary.calculate") && <button className="primary-button compact" disabled={!periodId || !projectId || calculate.isPending} onClick={() => periodId && projectId && calculate.mutate({ token, payrollPeriodId: periodId, projectId })}>{calculate.isPending ? "Calculating…" : "Create new version"}</button>}</div></section>{periodId && projectId && setup.data && hasPermission(permissions, "salary.adjust") && <form className="rounding-policy" onSubmit={(event) => { event.preventDefault(); const form = new FormData(event.currentTarget); savePolicy.mutate({ token, projectId, roundingScope: field(form, "roundingScope") as "per_component" | "per_project" | "at_total", decimalPlaces: Number(field(form, "decimalPlaces")) }); }}><div><b>Rounding policy</b><small>Internal calculations preserve six decimal places.</small></div><label>Apply<select name="roundingScope" defaultValue={setup.data.policy.roundingScope}><option value="per_component">Per component</option><option value="per_project">Per project</option><option value="at_total">At total</option></select></label><label>Decimals<input name="decimalPlaces" type="number" min="0" max="6" defaultValue={setup.data.policy.decimalPlaces} /></label><button className="secondary-button compact">Save policy</button></form>}{runs.isPending ? <Loading /> : runs.error ? <Empty title="Salary runs could not load" body={message(runs.error)} /> : !runs.data?.runs.length ? <Empty title="No salary calculation versions" body="Select a period and project with an approved reconciliation run, then create the first version." /> : <div className="salary-runs"><div className="salary-runs-head"><span>Version</span><span>Scope</span><span>Components</span><span>Totals</span><span>Result</span></div>{runs.data.runs.map((run) => <button key={run.id} onClick={() => setRunId(run.id)}><span><b>V{run.version}</b><small>{formatDate(run.startedAt)}</small></span><span><b>{run.periodCode}</b><small>{run.projectCode} · {run.projectName}</small></span><span><b>{run.componentCount}</b><small>{run.exceptionCount} exceptions</small></span><span><b>{run.currency} {run.finalTotal}</b><small>Gross {run.grossTotal} · adjustments {run.adjustmentTotal}</small></span><span><Status value={run.status} /><small>{humanize(run.roundingScope)} · {run.decimalPlaces} dp</small></span></button>)}</div>}{runId && <Modal title="Calculation trace" wide onClose={() => setRunId(null)}><SalaryRunLedger token={token} runId={runId} permissions={permissions} notify={notify} /></Modal>}{inputModal && periodId && projectId && <Modal title="Authorize slab input" onClose={() => setInputModal(false)}><form className="stack-form" onSubmit={(event) => { event.preventDefault(); const form = new FormData(event.currentTarget); saveInput.mutate({ token, payrollPeriodId: periodId, projectId, employeeId: Number(field(form, "employeeId")), averageHours: field(form, "averageHours"), reason: field(form, "reason") }); }}><p className="form-lead">Slab rates cannot run until average working hours are explicitly authorized for each affected employee.</p><label>Employee<select name="employeeId" required><option value="">Select employee</option>{setup.data?.employees.map((item) => <option key={item.id} value={item.id}>{item.employeeId} · {item.name}{item.averageHours ? ` · current ${item.averageHours}` : ""}</option>)}</select></label><label>Average working hours<input name="averageHours" inputMode="decimal" required /></label><label>Authorization reason<textarea name="reason" minLength={3} rows={3} required /></label><button className="primary-button" disabled={saveInput.isPending}>Save authorized input</button></form></Modal>}</>;
}

function SalaryDetails({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const allowed = hasPermission(permissions, "salary.read"); const [periodId, setPeriodId] = useState<number | null>(null); const [projectId, setProjectId] = useState<number | null>(null); const [employeeId, setEmployeeId] = useState<number | null>(null); const [drill, setDrill] = useState<null | { runId: number; employeeId: number; label: string }>(null);
  const details = useQuery({ queryKey: ["employee-salary", token, periodId, employeeId], queryFn: () => api.getEmployeeSalaryDetails({ token, payrollPeriodId: periodId ?? 0, employeeId }), enabled: allowed && Boolean(periodId) });
  if (!allowed) return <Empty title="Employee salary details are restricted" body="Your role does not include salary access." />;
  return <><ProductionScope token={token} periodId={periodId} projectId={projectId} setPeriodId={(value) => { setPeriodId(value); setEmployeeId(null); }} setProjectId={setProjectId} /><div className="toolbar"><label>Employee<select value={employeeId ?? ""} onChange={(event) => setEmployeeId(event.target.value ? Number(event.target.value) : null)}><option value="">All employees</option>{details.data?.employees.map((item) => <option key={item.id} value={item.id}>{item.employeeId} · {item.name}</option>)}</select></label></div>{!periodId ? <Empty title="Select a payroll period" body="Employee totals use only the latest completed calculation version for each project." /> : details.isPending ? <Loading /> : details.error ? <Empty title="Salary details could not load" body={message(details.error)} /> : !details.data?.employees.length ? <Empty title="No completed salary details" body="Complete at least one unblocked project calculation for this payroll period." /> : <div className="employee-salary-list">{details.data.employees.map((employee) => <article key={employee.id}><header><div><strong>{employee.employeeId}</strong><span>{employee.name}</span></div><div><small>Grand total across projects</small><b>{employee.grandTotal}</b></div></header><div className="employee-projects">{employee.projects.filter((project) => !projectId || project.projectId === projectId).map((project) => <button key={project.projectId} onClick={() => setDrill({ runId: project.salaryRunId, employeeId: employee.id, label: `${employee.employeeId} · ${project.projectCode}` })}><span><b>{project.projectCode}</b><small>{project.projectName}</small></span><span><b>{project.currency} {project.subtotal}</b><small>{project.componentCount} components · version {project.runVersion}</small></span><em>Trace figures →</em></button>)}</div></article>)}</div>}{drill && <Modal title={drill.label} wide onClose={() => setDrill(null)}><SalaryRunLedger token={token} runId={drill.runId} employeeId={drill.employeeId} permissions={permissions} notify={notify} /></Modal>}</>;
}

function Users({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const queryClient = useQueryClient();
  const [modal, setModal] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState<any | null>(null);
  const [viewUser, setViewUser] = useState<any | null>(null);
  const allowed = hasPermission(permissions, "users.manage");

  const query = useQuery({ queryKey: ["users", token], queryFn: () => api.listUsers({ token }), enabled: allowed });
  const save = useMutation({
    mutationFn: api.addUser,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["users"] });
      queryClient.invalidateQueries({ queryKey: ["audit"] });
      setModal(false);
      notify({ kind: "success", text: "User account created with server-enforced role permissions." });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const removeMutation = useMutation({
    mutationFn: api.deleteUser,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["users"] });
      queryClient.invalidateQueries({ queryKey: ["recycle-bin"] });
      queryClient.invalidateQueries({ queryKey: ["audit"] });
      setDeleteTarget(null);
      notify({ kind: "success", text: "User account moved to Recycle Bin." });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  if (!allowed) return <Empty title="User management is restricted" body="Only administrators can create and review application accounts." />;

  const columns: Column<any>[] = [
    {
      key: "name",
      header: "User",
      isImportant: true,
      render: (u) => (
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 32, height: 32, borderRadius: "50%", background: "#d8ece4", color: "#073b33", display: "grid", placeItems: "center", fontWeight: 700, fontSize: "12px", border: "1px solid #0f5b4a", flexShrink: 0, overflow: "hidden" }}>
            {u.avatar ? <img src={u.avatar} alt={u.name} style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : u.name.slice(0, 1).toUpperCase()}
          </div>
          <div>
            <strong>{u.name}</strong>
          </div>
        </div>
      ),
    },
    { key: "email", header: "Email Address" },
    {
      key: "roleName",
      header: "Role",
      render: (u) => (
        <span style={{ padding: "2px 8px", background: "#f0f8f5", border: "1px solid #99c2b6", borderRadius: 4, fontSize: "11.5px", fontWeight: 600, color: "#073b33" }}>
          {u.roleName}
        </span>
      ),
    },
    {
      key: "status",
      header: "Status",
      render: (u) => <Status value={u.status} />,
    },
  ];

  return (
    <>
      <div style={{ marginBottom: 12, display: "flex", justifyContent: "flex-end" }}>
        <button className="primary-button compact" onClick={() => setModal(true)}>
          <Icon name="plus" /> Create user
        </button>
      </div>

      {query.isPending ? (
        <Loading />
      ) : query.error ? (
        <Empty title="Users could not load" body={message(query.error)} />
      ) : (
        <DataTable
          data={query.data?.users ?? []}
          columns={columns}
          title="User Account Management"
          subtitle="System operators, managers, reviewers, and administrative roles"
          searchPlaceholder="Search users by name or email..."
          onView={(u) => setViewUser(u)}
          onDelete={(u) => setDeleteTarget(u)}
        />
      )}

      {modal && (
        <Modal title="Create user account" onClose={() => setModal(false)}>
          <form
            className="stack-form"
            onSubmit={(event) => {
              event.preventDefault();
              const f = new FormData(event.currentTarget);
              save.mutate({
                token,
                name: field(f, "name"),
                email: field(f, "email"),
                password: field(f, "password"),
                role: field(f, "role") as any,
              });
            }}
          >
            <label>
              Full name
              <input name="name" required />
            </label>
            <label>
              Work email
              <input name="email" type="email" required />
            </label>
            <label>
              Temporary password
              <input name="password" type="password" minLength={10} required />
            </label>
            <label>
              Role
              <select name="role" required>
                {query.data?.roles.map((role) => (
                  <option key={role.id} value={role.slug}>
                    {role.name}
                  </option>
                ))}
              </select>
            </label>
            <div className="form-actions">
              <button type="button" className="ghost-button" onClick={() => setModal(false)}>
                Cancel
              </button>
              <button className="primary-button" disabled={save.isPending}>
                <Icon name="check" /> {save.isPending ? "Creating…" : "Create user"}
              </button>
            </div>
          </form>
        </Modal>
      )}

      {viewUser && (
        <Modal title="User Account Details" onClose={() => setViewUser(null)}>
          <div style={{ display: "grid", gap: 12 }}>
            <p><strong>Name:</strong> {viewUser.name}</p>
            <p><strong>Email:</strong> {viewUser.email}</p>
            <p><strong>Role:</strong> {viewUser.roleName}</p>
            <p><strong>Status:</strong> <Status value={viewUser.status} /></p>
          </div>
          <div className="form-actions" style={{ marginTop: 16 }}>
            <button className="secondary-button" onClick={() => setViewUser(null)}>Close</button>
          </div>
        </Modal>
      )}

      <ConfirmDialog
        isOpen={Boolean(deleteTarget)}
        title="Move User Account to Recycle Bin?"
        message={`Are you sure you want to deactivate and remove account "${deleteTarget?.name}" (${deleteTarget?.email})? It will be moved to the Recycle Bin and can be restored whenever needed.`}
        confirmText="Move to Recycle Bin"
        danger
        onConfirm={() => {
          if (deleteTarget) removeMutation.mutate({ token, id: deleteTarget.id });
        }}
        onCancel={() => setDeleteTarget(null)}
      />
    </>
  );
}

function Roles({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const queryClient = useQueryClient();
  const [modal, setModal] = useState(false);
  const [viewRole, setViewRole] = useState<any | null>(null);
  const [deleteTarget, setDeleteTarget] = useState<any | null>(null);
  const [editingRole, setEditingRole] = useState<{ id?: number; name: string; slug: string; permissionCodes: string[] } | null>(null);
  const allowed = hasPermission(permissions, "users.manage");

  const query = useQuery({ queryKey: ["roles", token], queryFn: () => api.listRoles({ token }), enabled: allowed });

  const save = useMutation({
    mutationFn: api.saveRole,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["roles"] });
      queryClient.invalidateQueries({ queryKey: ["users"] });
      queryClient.invalidateQueries({ queryKey: ["audit"] });
      setModal(false);
      setEditingRole(null);
      notify({ kind: "success", text: "Role permissions saved successfully." });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const removeMutation = useMutation({
    mutationFn: api.deleteRole,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["roles"] });
      queryClient.invalidateQueries({ queryKey: ["audit"] });
      setDeleteTarget(null);
      notify({ kind: "success", text: "Role deleted successfully." });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  if (!allowed) return <Empty title="Role management is restricted" body="Only administrators can manage roles and permissions." />;

  const openCreate = () => {
    setEditingRole({ name: "", slug: "", permissionCodes: [] });
    setModal(true);
  };

  const openEdit = (r: any) => {
    setEditingRole({
      id: r.id,
      name: r.name,
      slug: r.slug,
      permissionCodes: [...r.permissionCodes],
    });
    setModal(true);
  };

  const togglePermission = (code: string) => {
    if (!editingRole) return;
    const exists = editingRole.permissionCodes.includes(code);
    const updated = exists
      ? editingRole.permissionCodes.filter((c) => c !== code)
      : [...editingRole.permissionCodes, code];
    setEditingRole({ ...editingRole, permissionCodes: updated });
  };

  const columns: Column<any>[] = [
    {
      key: "name",
      header: "Role Name",
      isImportant: true,
      render: (r) => (
        <div>
          <strong>{r.name}</strong>
          {r.slug === "administrator" && <small style={{ display: "block", color: "var(--accent)" }}>System Administrator (Protected)</small>}
        </div>
      ),
    },
    {
      key: "slug",
      header: "Role Identifier",
      render: (r) => <span className="mono" style={{ color: "#073b33" }}>{r.slug}</span>,
    },
    {
      key: "permissions",
      header: "Granted Permissions",
      render: (r) => (
        <span style={{ fontWeight: 600, background: "#f0f8f5", padding: "2px 8px", borderRadius: 4, border: "1px solid #99c2b6", color: "#073b33" }}>
          {r.permissionCodes.length} permissions
        </span>
      ),
    },
    {
      key: "type",
      header: "Type",
      render: (r) => (
        <span className={`status ${r.slug === "administrator" ? "good" : "info"}`}>
          {r.slug === "administrator" ? "Core System" : "Custom"}
        </span>
      ),
    },
  ];

  return (
    <>
      <div style={{ marginBottom: 12, display: "flex", justifyContent: "flex-end" }}>
        <button className="primary-button compact" onClick={openCreate}>
          <Icon name="plus" /> Add role
        </button>
      </div>

      {query.isPending ? (
        <Loading />
      ) : query.error ? (
        <Empty title="Roles could not load" body={message(query.error)} />
      ) : (
        <DataTable
          data={query.data?.roles ?? []}
          columns={columns}
          title="Role Permissions Matrix"
          subtitle="Define custom role capabilities and server-enforced access authority"
          searchPlaceholder="Search roles..."
          onView={(r) => setViewRole(r)}
          onEdit={(r) => openEdit(r)}
          onDelete={(r) => {
            if (r.slug === "administrator") {
              notify({ kind: "error", text: "The Administrator role cannot be deleted." });
              return;
            }
            setDeleteTarget(r);
          }}
        />
      )}

      {modal && editingRole && (
        <Modal title={editingRole.id ? `Edit Role: ${editingRole.name}` : "Create New Role"} onClose={() => { setModal(false); setEditingRole(null); }} wide>
          <form
            className="stack-form"
            onSubmit={(event) => {
              event.preventDefault();
              save.mutate({
                token,
                id: editingRole.id,
                name: editingRole.name,
                slug: editingRole.slug,
                permissionCodes: editingRole.permissionCodes,
              });
            }}
          >
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
              <label>
                Role Name
                <input
                  value={editingRole.name}
                  onChange={(e) => {
                    const name = e.target.value;
                    const autoSlug = editingRole.id ? editingRole.slug : name.toLowerCase().replace(/[^a-z0-9]+/g, "-");
                    setEditingRole({ ...editingRole, name, slug: autoSlug });
                  }}
                  required
                />
              </label>
              <label>
                Role Slug / Identifier
                <input
                  value={editingRole.slug}
                  disabled={editingRole.slug === "administrator"}
                  onChange={(e) => setEditingRole({ ...editingRole, slug: e.target.value.toLowerCase().replace(/[^a-z0-9-]+/g, "") })}
                  required
                />
              </label>
            </div>

            <div style={{ marginTop: 12 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <strong>Assigned Permissions ({editingRole.permissionCodes.length})</strong>
                <div style={{ display: "flex", gap: 8 }}>
                  <button
                    type="button"
                    className="ghost-button compact"
                    onClick={() => setEditingRole({ ...editingRole, permissionCodes: query.data?.allPermissions.map((p) => p.code) ?? [] })}
                  >
                    Select All
                  </button>
                  <button
                    type="button"
                    className="ghost-button compact"
                    onClick={() => setEditingRole({ ...editingRole, permissionCodes: [] })}
                  >
                    Clear All
                  </button>
                </div>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(260px, 1fr))", gap: 8, maxHeight: 360, overflowY: "auto", padding: 12, background: "#f8faf9", border: "1px solid #d4ded9", borderRadius: 8 }}>
                {query.data?.allPermissions.map((perm) => {
                  const checked = editingRole.permissionCodes.includes(perm.code);
                  return (
                    <label key={perm.code} style={{ display: "flex", alignItems: "flex-start", gap: 8, fontSize: "12px", cursor: "pointer", padding: "4px 6px", borderRadius: 4, background: checked ? "#e8f3ef" : "transparent" }}>
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={() => togglePermission(perm.code)}
                      />
                      <div>
                        <strong>{perm.label}</strong>
                        <small style={{ display: "block", color: "var(--dim)", fontFamily: "monospace" }}>{perm.code}</small>
                      </div>
                    </label>
                  );
                })}
              </div>
            </div>

            <div className="form-actions" style={{ marginTop: 16 }}>
              <button type="button" className="ghost-button" onClick={() => { setModal(false); setEditingRole(null); }}>
                Cancel
              </button>
              <button className="primary-button" disabled={save.isPending}>
                <Icon name="check" /> {save.isPending ? "Saving…" : "Save role"}
              </button>
            </div>
          </form>
        </Modal>
      )}

      {viewRole && (
        <Modal title={`Role: ${viewRole.name}`} onClose={() => setViewRole(null)} wide>
          <div style={{ display: "grid", gap: 14 }}>
            <div>
              <p><strong>Name:</strong> {viewRole.name}</p>
              <p><strong>Identifier:</strong> <span className="mono">{viewRole.slug}</span></p>
              <p><strong>Granted Permissions:</strong> {viewRole.permissionCodes.length}</p>
            </div>
            <div>
              <strong>Permission List:</strong>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 8, maxHeight: 280, overflowY: "auto" }}>
                {viewRole.permissionCodes.map((code: string) => (
                  <span key={code} style={{ background: "#e8f3ef", border: "1px solid #99c2b6", padding: "2px 8px", borderRadius: 4, fontSize: "11px", fontFamily: "monospace", color: "#073b33" }}>
                    {code}
                  </span>
                ))}
              </div>
            </div>
          </div>
          <div className="form-actions" style={{ marginTop: 16 }}>
            <button className="secondary-button" onClick={() => setViewRole(null)}>Close</button>
          </div>
        </Modal>
      )}

      <ConfirmDialog
        isOpen={Boolean(deleteTarget)}
        title="Delete Role?"
        message={`Are you sure you want to permanently delete role "${deleteTarget?.name}"? Users assigned to this role will need to be reassigned.`}
        confirmText="Delete Role"
        danger
        onConfirm={() => {
          if (deleteTarget) removeMutation.mutate({ token, id: deleteTarget.id });
        }}
        onCancel={() => setDeleteTarget(null)}
      />
    </>
  );
}

function RecycleBin({ token, permissions: _permissions, notify }: { token: string; permissions?: string[]; notify: (notice: Notice) => void }) {
  const queryClient = useQueryClient();
  const [purgeTarget, setPurgeTarget] = useState<any | null>(null);
  const [emptyConfirm, setEmptyConfirm] = useState(false);
  const [viewItem, setViewItem] = useState<any | null>(null);

  const query = useQuery({ queryKey: ["recycle-bin", token], queryFn: () => api.listRecycleBin({ token }) });

  const restore = useMutation({
    mutationFn: api.restoreFromRecycleBin,
    onSuccess: (data) => {
      queryClient.invalidateQueries({ queryKey: ["recycle-bin"] });
      queryClient.invalidateQueries({ queryKey: ["projects"] });
      queryClient.invalidateQueries({ queryKey: ["employees"] });
      queryClient.invalidateQueries({ queryKey: ["periods"] });
      queryClient.invalidateQueries({ queryKey: ["assignments"] });
      queryClient.invalidateQueries({ queryKey: ["users"] });
      queryClient.invalidateQueries({ queryKey: ["dashboard"] });
      notify({ kind: "success", text: `Successfully restored ${data.entityType} to active workspace.` });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const purge = useMutation({
    mutationFn: api.emptyRecycleBin,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["recycle-bin"] });
      setPurgeTarget(null);
      setEmptyConfirm(false);
      notify({ kind: "success", text: "Recycle bin items permanently purged." });
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const columns: Column<any>[] = [
    {
      key: "entityType",
      header: "Entity Type",
      isImportant: true,
      render: (item) => (
        <span style={{ textTransform: "capitalize", fontWeight: 700, padding: "2px 8px", background: "#f0f8f5", border: "1px solid #99c2b6", borderRadius: 4, color: "#073b33", fontSize: "11.5px" }}>
          {item.entityType}
        </span>
      ),
    },
    {
      key: "title",
      header: "Item Title / Reference",
      render: (item) => (
        <div>
          <strong>{item.title}</strong>
          <small style={{ display: "block", color: "var(--dim)", fontFamily: "monospace" }}>ID: {item.entityId}</small>
        </div>
      ),
    },
    { key: "deletedBy", header: "Deleted By" },
    {
      key: "deletedAt",
      header: "Deleted At",
      render: (item) => formatDateTime(item.deletedAt),
    },
  ];

  return (
    <>
      <div style={{ marginBottom: 12, display: "flex", justifyContent: "flex-end", gap: 10 }}>
        {Boolean(query.data?.items.length) && (
          <button className="ghost-button compact" style={{ color: "var(--danger)" }} onClick={() => setEmptyConfirm(true)}>
            <Icon name="trash" /> Empty Recycle Bin
          </button>
        )}
      </div>

      {query.isPending ? (
        <Loading />
      ) : query.error ? (
        <Empty title="Recycle bin could not load" body={message(query.error)} />
      ) : !query.data?.items.length ? (
        <Empty title="Recycle bin is empty" body="Deleted projects, employees, and periods will appear here for safe recovery." />
      ) : (
        <DataTable
          data={query.data.items}
          columns={columns}
          title="System Recycle Bin"
          subtitle="Safe recovery area for previously deleted projects, employees, periods, and records"
          searchPlaceholder="Search recycle bin items..."
          onView={(item) => setViewItem(item)}
          customActions={(item) => (
            <div style={{ display: "flex", gap: 4 }}>
              <button
                type="button"
                className="table-action-btn view-action"
                title="Restore item"
                onClick={() => restore.mutate({ token, id: item.id })}
              >
                <Icon name="restore" />
              </button>
              <button
                type="button"
                className="table-action-btn delete-action"
                title="Purge permanently"
                onClick={() => setPurgeTarget(item)}
              >
                <Icon name="trash" />
              </button>
            </div>
          )}
        />
      )}

      {viewItem && (
        <Modal title={`Recycled ${viewItem.entityType}: ${viewItem.title}`} onClose={() => setViewItem(null)} wide>
          <div style={{ display: "grid", gap: 12 }}>
            <p><strong>Entity Type:</strong> {viewItem.entityType}</p>
            <p><strong>Entity ID:</strong> {viewItem.entityId}</p>
            <p><strong>Title:</strong> {viewItem.title}</p>
            <p><strong>Deleted By:</strong> {viewItem.deletedBy}</p>
            <p><strong>Deleted At:</strong> {formatDateTime(viewItem.deletedAt)}</p>
            <div>
              <strong>Original Payload Snapshot:</strong>
              <pre style={{ maxHeight: 220, overflowY: "auto", background: "#f8faf9", padding: 12, borderRadius: 6, fontSize: "11px", border: "1px solid #d4ded9" }}>
                {(() => {
                  try {
                    return JSON.stringify(JSON.parse(viewItem.payloadJson), null, 2);
                  } catch {
                    return viewItem.payloadJson;
                  }
                })()}
              </pre>
            </div>
          </div>
          <div className="form-actions" style={{ marginTop: 16 }}>
            <button className="secondary-button" onClick={() => setViewItem(null)}>Close</button>
            <button
              className="primary-button"
              onClick={() => {
                restore.mutate({ token, id: viewItem.id });
                setViewItem(null);
              }}
            >
              <Icon name="restore" /> Restore this item
            </button>
          </div>
        </Modal>
      )}

      <ConfirmDialog
        isOpen={Boolean(purgeTarget)}
        title="Permanently Delete Item?"
        message={`This will permanently purge "${purgeTarget?.title}" from the database. This action cannot be undone.`}
        confirmText="Permanently Delete"
        danger
        onConfirm={() => {
          if (purgeTarget) purge.mutate({ token, id: purgeTarget.id });
        }}
        onCancel={() => setPurgeTarget(null)}
      />

      <ConfirmDialog
        isOpen={emptyConfirm}
        title="Empty Entire Recycle Bin?"
        message="Are you sure you want to permanently delete all items in the Recycle Bin? This action is irreversible."
        confirmText="Empty Recycle Bin"
        danger
        onConfirm={() => purge.mutate({ token })}
        onCancel={() => setEmptyConfirm(false)}
      />
    </>
  );
}


type AuditSeverity = "info" | "warning" | "critical";

function computeLogSeverity(log: { action: string; entityType?: string; reason?: string | null; severity?: string }): AuditSeverity {
  if (log.severity === "critical" || log.severity === "warning" || log.severity === "info") {
    return log.severity;
  }
  const act = (log.action ?? "").toLowerCase();
  const reason = (log.reason ?? "").toLowerCase();
  if (
    act.includes("delete") ||
    act.includes("purge") ||
    act.includes("empty") ||
    act.includes("lock") ||
    act.includes("freeze") ||
    act.includes("block") ||
    reason.includes("security") ||
    reason.includes("breach")
  ) {
    return "critical";
  }
  if (
    act.includes("adjust") ||
    act.includes("restore") ||
    act.includes("reverse") ||
    act.includes("override") ||
    act.includes("reject") ||
    act.includes("reconcil") ||
    reason.includes("correction") ||
    reason.includes("discrepancy")
  ) {
    return "warning";
  }
  return "info";
}

function Audit({ token, permissions, notify }: { token: string; permissions: string[]; notify?: (notice: Notice) => void }) {
  const allowed = hasPermission(permissions, "audit.read");
  const query = useQuery({
    queryKey: ["audit", token],
    queryFn: () => api.listAuditLogs({ token, limit: 200 }),
    enabled: allowed,
  });

  const [filterText, setFilterText] = useState("");
  const [severityFilter, setSeverityFilter] = useState<"all" | AuditSeverity>("all");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [hoveredBarIndex, setHoveredBarIndex] = useState<number | null>(null);
  const [progressStatus, setProgressStatus] = useState<{
    mode: "indexing" | "exporting";
    percent: number;
    title: string;
    subtitle: string;
  }>({
    mode: "indexing",
    percent: 100,
    title: "Log Indexing Complete",
    subtitle: "All persistent audit records synchronized and indexed",
  });

  // Keep indexing status updated when query state changes
  useEffect(() => {
    if (query.isPending) {
      setProgressStatus({
        mode: "indexing",
        percent: 45,
        title: "Indexing Audit Log Stream…",
        subtitle: "Fetching persistent event journal from database",
      });
    } else if (query.data?.logs) {
      setProgressStatus({
        mode: "indexing",
        percent: 100,
        title: "Log Indexing Complete",
        subtitle: `${query.data.logs.length} immutable events verified and indexed`,
      });
    }
  }, [query.isPending, query.data?.logs]);

  if (!allowed) return <Empty title="Audit access is restricted" body="Your role does not include audit history." />;
  if (query.isPending && !query.data) return <Loading />;
  if (query.error) return <Empty title="Audit log could not load" body={message(query.error)} />;
  if (!query.data?.logs.length) return <Empty title="No audit entries yet" body="Create or change a foundation record to begin the audit history." />;

  const rawLogs = query.data.logs;

  // Severity counts across all logs
  const countBySeverity = {
    all: rawLogs.length,
    info: rawLogs.filter((l) => computeLogSeverity(l) === "info").length,
    warning: rawLogs.filter((l) => computeLogSeverity(l) === "warning").length,
    critical: rawLogs.filter((l) => computeLogSeverity(l) === "critical").length,
  };

  // 30-day frequency time-series segmented by severity
  const chartData = useMemo(() => {
    const now = new Date();
    const days: Array<{
      date: string;
      displayDate: string;
      fullDate: string;
      info: number;
      warning: number;
      critical: number;
      total: number;
    }> = [];

    for (let i = 29; i >= 0; i--) {
      const d = new Date(now.getTime() - i * 86400000);
      const iso = d.toISOString().slice(0, 10);
      const displayDate = d.toLocaleDateString("en-US", { month: "short", day: "numeric" });
      const fullDate = d.toLocaleDateString("en-US", { weekday: "short", year: "numeric", month: "short", day: "numeric" });
      days.push({
        date: iso,
        displayDate,
        fullDate,
        info: 0,
        warning: 0,
        critical: 0,
        total: 0,
      });
    }

    const dayMap = new Map(days.map((item) => [item.date, item]));

    for (const log of rawLogs) {
      const logDate = log.createdAt.slice(0, 10);
      const entry = dayMap.get(logDate);
      if (entry) {
        const sev = computeLogSeverity(log);
        entry[sev] += 1;
        entry.total += 1;
      }
    }

    return days;
  }, [rawLogs]);

  const chartStats = useMemo(() => {
    let total30d = 0;
    let critical30d = 0;
    let warning30d = 0;
    let info30d = 0;
    for (const d of chartData) {
      total30d += d.total;
      critical30d += d.critical;
      warning30d += d.warning;
      info30d += d.info;
    }
    const dailyAvg = (total30d / 30).toFixed(1);
    return { total30d, critical30d, warning30d, info30d, dailyAvg };
  }, [chartData]);

  const applyDatePreset = (preset: "today" | "7d" | "30d" | "all") => {
    if (preset === "all") {
      setStartDate("");
      setEndDate("");
      return;
    }
    const now = new Date();
    const endIso = now.toISOString().slice(0, 10);
    if (preset === "today") {
      setStartDate(endIso);
      setEndDate(endIso);
      return;
    }
    const days = preset === "7d" ? 6 : 29;
    const start = new Date(now.getTime() - days * 86400000);
    setStartDate(start.toISOString().slice(0, 10));
    setEndDate(endIso);
  };

  const queryTerm = filterText.trim().toLowerCase();
  const filteredLogs = rawLogs.filter((log) => {
    const sev = computeLogSeverity(log);

    // 1. Severity filter check
    if (severityFilter !== "all" && sev !== severityFilter) {
      return false;
    }

    // 2. Date range check (entry's timestamp)
    const logDate = log.createdAt.slice(0, 10);
    if (startDate && logDate < startDate) {
      return false;
    }
    if (endDate && logDate > endDate) {
      return false;
    }

    // 3. Text filter check (actor name or action description)
    if (!queryTerm) return true;
    const actor = (log.actor ?? "").toLowerCase();
    const action = (log.action ?? "").toLowerCase();
    const humanAction = humanize(log.action ?? "").toLowerCase();
    const entityType = (log.entityType ?? "").toLowerCase();
    const humanEntity = humanize(log.entityType ?? "").toLowerCase();
    const entityId = (log.entityId ?? "").toLowerCase();
    const reason = (log.reason ?? "").toLowerCase();
    const fullActionDesc = `${actor} ${action} ${humanAction} ${entityType} ${humanEntity} #${entityId} ${reason} ${sev}`;

    return actor.includes(queryTerm) || fullActionDesc.includes(queryTerm);
  });

  const exportCsv = () => {
    const logsToExport = filteredLogs.length > 0 ? filteredLogs : rawLogs;
    if (!logsToExport.length) return;

    // Trigger visual export progress
    setProgressStatus({
      mode: "exporting",
      percent: 30,
      title: "Preparing CSV Export…",
      subtitle: `Formatting ${logsToExport.length} audit records…`,
    });

    setTimeout(() => {
      setProgressStatus({
        mode: "exporting",
        percent: 75,
        title: "Generating Audit Workbook…",
        subtitle: "Compiling comma-separated data stream…",
      });

      setTimeout(() => {
        const headers = [
          "Log ID",
          "Timestamp",
          "Actor",
          "Action",
          "Entity Type",
          "Entity ID",
          "Severity",
          "Action Description",
          "Reason",
        ];

        const escapeCsv = (val: unknown) => {
          if (val === null || val === undefined) return '""';
          const str = String(val);
          return `"${str.replace(/"/g, '""')}"`;
        };

        const rows = logsToExport.map((log) => {
          const sev = computeLogSeverity(log);
          const desc = `${log.actor} ${humanize(log.action)} ${humanize(log.entityType)} #${log.entityId}`;
          return [
            escapeCsv(log.id),
            escapeCsv(formatDateTime(log.createdAt)),
            escapeCsv(log.actor),
            escapeCsv(log.action),
            escapeCsv(log.entityType),
            escapeCsv(log.entityId),
            escapeCsv(sev.toUpperCase()),
            escapeCsv(desc),
            escapeCsv(log.reason ?? ""),
          ].join(",");
        });

        const csvContent = [headers.join(","), ...rows].join("\r\n");
        const dateStr = new Date().toISOString().slice(0, 10);
        const filename = `audit-logs-${dateStr}.csv`;
        downloadText(filename, csvContent, "text/csv;charset=utf-8");

        setProgressStatus({
          mode: "exporting",
          percent: 100,
          title: "CSV Export Complete",
          subtitle: `Successfully exported ${logsToExport.length} records to ${filename}`,
        });

        notify?.({
          kind: "success",
          text: `${logsToExport.length} audit log entries exported to ${filename}.`,
        });

        // Smoothly transition back to indexing status after 3.5s
        setTimeout(() => {
          setProgressStatus({
            mode: "indexing",
            percent: 100,
            title: "Log Indexing Complete",
            subtitle: `${rawLogs.length} verified immutable events indexed in ledger`,
          });
        }, 3500);
      }, 250);
    }, 250);
  };

  return (
    <div className="audit-page-wrap">
      {/* 1. Audit Component Title Header & Visual Progress Bar */}
      <div className="audit-header-panel">
        <div className="audit-header-title-row">
          <div className="audit-header-title-text">
            <span className="audit-pretitle">Security & Immutability Ledger</span>
            <h2>Audit Trail & System Journal</h2>
            <p>
              Persistent chronological record of security operations, foundation updates, payroll runs, and user actions.
            </p>
          </div>
          <div className="audit-header-badge-wrap">
            <span className="audit-ledger-verified-badge">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                <polyline points="20 6 9 17 4 12" />
              </svg>
              <span>Ledger Verified</span>
            </span>
          </div>
        </div>

        {/* Visual Progress Bar beneath the 'Audit' component title */}
        <div className="audit-progress-card" role="region" aria-label="Log indexing and export progress">
          <div className="audit-progress-meta-row">
            <div className="audit-progress-status-line">
              <span className={`audit-progress-status-dot ${progressStatus.mode} ${progressStatus.percent === 100 ? "complete" : ""}`} />
              <div className="audit-progress-text-group">
                <strong className="audit-progress-title">{progressStatus.title}</strong>
                <span className="audit-progress-subtitle">{progressStatus.subtitle}</span>
              </div>
            </div>
            <div className="audit-progress-numeric-wrap">
              <strong className="audit-progress-numeric-value">{progressStatus.percent}%</strong>
              <small className="audit-progress-numeric-label">
                {progressStatus.mode === "exporting" ? "Export Process" : "Indexing Status"}
              </small>
            </div>
          </div>
          <div className="audit-progress-bar-track">
            <div
              className={`audit-progress-bar-fill ${progressStatus.mode} ${progressStatus.percent === 100 ? "complete" : ""}`}
              style={{ width: `${progressStatus.percent}%` }}
            />
          </div>
        </div>
      </div>

      {/* 2. Recharts 30-Day Activity & Severity Frequency Chart */}
      <div className="audit-chart-card">
        <div className="audit-chart-head">
          <div className="audit-chart-title-group">
            <div className="audit-chart-header-pre">Time-Series Analytics</div>
            <h3>30-Day Activity & Severity Distribution</h3>
            <p>Frequency of immutable audit entries recorded across the last 30 days, segmented by severity level.</p>
          </div>
          <div className="audit-chart-metrics">
            <div className="audit-chart-metric-item">
              <span className="audit-chart-metric-num">{chartStats.total30d}</span>
              <span className="audit-chart-metric-lbl">30D Events</span>
            </div>
            <div className="audit-chart-metric-item critical">
              <span className="audit-chart-metric-dot critical" />
              <span className="audit-chart-metric-num">{chartStats.critical30d}</span>
              <span className="audit-chart-metric-lbl">Critical</span>
            </div>
            <div className="audit-chart-metric-item warning">
              <span className="audit-chart-metric-dot warning" />
              <span className="audit-chart-metric-num">{chartStats.warning30d}</span>
              <span className="audit-chart-metric-lbl">Warning</span>
            </div>
            <div className="audit-chart-metric-item info">
              <span className="audit-chart-metric-dot info" />
              <span className="audit-chart-metric-num">{chartStats.info30d}</span>
              <span className="audit-chart-metric-lbl">Info</span>
            </div>
          </div>
        </div>

        <div className="audit-chart-visual-wrap" style={{ width: "100%", height: 250, minHeight: 250 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart
              data={chartData}
              margin={{ top: 14, right: 14, left: -20, bottom: 4 }}
              onClick={(state) => {
                if (state && state.activePayload && state.activePayload.length) {
                  const clicked = (state.activePayload[0].payload as { date: string }).date;
                  setStartDate(clicked);
                  setEndDate(clicked);
                }
              }}
              onMouseMove={(state) => {
                if (state && typeof state.activeTooltipIndex === "number") {
                  setHoveredBarIndex(state.activeTooltipIndex);
                } else {
                  setHoveredBarIndex(null);
                }
              }}
              onMouseLeave={() => setHoveredBarIndex(null)}
              style={{ cursor: "pointer" }}
            >
              <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" opacity={0.6} vertical={false} />
              <XAxis
                dataKey="displayDate"
                tick={{ fill: "var(--dim)", fontSize: 11, fontWeight: 500 }}
                tickLine={false}
                axisLine={{ stroke: "var(--border)" }}
                interval={4}
              />
              <YAxis
                allowDecimals={false}
                tick={{ fill: "var(--dim)", fontSize: 11, fontWeight: 500 }}
                tickLine={false}
                axisLine={false}
              />
              <Tooltip
                content={<CustomAuditTooltip />}
                animationDuration={180}
                animationEasing="ease-out"
                cursor={{ fill: "rgba(15, 91, 74, 0.08)", rx: 6, ry: 6 }}
              />
              <Legend
                verticalAlign="top"
                align="right"
                iconType="circle"
                iconSize={8}
                wrapperStyle={{ paddingBottom: 10, fontSize: 12 }}
                formatter={(value: string) => (
                  <span style={{ color: "var(--text)", fontWeight: 600, fontSize: 12, marginRight: 8 }}>
                    {value.charAt(0).toUpperCase() + value.slice(1)}
                  </span>
                )}
              />
              <Bar dataKey="critical" name="Critical" stackId="sev" fill="#dc2626" isAnimationActive animationDuration={350}>
                {chartData.map((_, index) => (
                  <Cell
                    key={`crit-cell-${index}`}
                    fill="#dc2626"
                    className={`audit-chart-bar-cell ${hoveredBarIndex === index ? "hovered" : hoveredBarIndex !== null ? "dimmed" : ""}`}
                  />
                ))}
              </Bar>
              <Bar dataKey="warning" name="Warning" stackId="sev" fill="#d97706" isAnimationActive animationDuration={350}>
                {chartData.map((_, index) => (
                  <Cell
                    key={`warn-cell-${index}`}
                    fill="#d97706"
                    className={`audit-chart-bar-cell ${hoveredBarIndex === index ? "hovered" : hoveredBarIndex !== null ? "dimmed" : ""}`}
                  />
                ))}
              </Bar>
              <Bar dataKey="info" name="Info" stackId="sev" fill="#0f5b4a" radius={[3, 3, 0, 0]} isAnimationActive animationDuration={350}>
                {chartData.map((_, index) => (
                  <Cell
                    key={`info-cell-${index}`}
                    fill="#0f5b4a"
                    className={`audit-chart-bar-cell ${hoveredBarIndex === index ? "hovered" : hoveredBarIndex !== null ? "dimmed" : ""}`}
                  />
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div className="audit-chart-foot">
          {hoveredBarIndex !== null && chartData[hoveredBarIndex] ? (
            <span className="audit-chart-tip active-hover">
              <strong style={{ color: "var(--accent-deep)", marginRight: 6 }}>
                📅 {chartData[hoveredBarIndex].fullDate}:
              </strong>
              <span>
                {chartData[hoveredBarIndex].total} event{chartData[hoveredBarIndex].total === 1 ? "" : "s"} (
                <span style={{ color: "#dc2626", fontWeight: 700 }}>{chartData[hoveredBarIndex].critical} critical</span>,{" "}
                <span style={{ color: "#d97706", fontWeight: 700 }}>{chartData[hoveredBarIndex].warning} warning</span>,{" "}
                <span style={{ color: "#0f5b4a", fontWeight: 700 }}>{chartData[hoveredBarIndex].info} info</span>)
              </span>
            </span>
          ) : (
            <span className="audit-chart-tip">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" style={{ marginRight: 5, verticalAlign: "-2px" }}>
                <circle cx="12" cy="12" r="10" />
                <line x1="12" y1="16" x2="12" y2="12" />
                <line x1="12" y1="8" x2="12.01" y2="8" />
              </svg>
              Tip: Hover over bars to inspect daily distribution · Click any bar to filter audit logs to that date.
            </span>
          )}
          <span className="audit-chart-avg">
            30-day average: <strong>{chartStats.dailyAvg}</strong> entries / day
          </span>
        </div>
      </div>

      {/* 3. Controls: Severity Filter Chips + Date Range Picker + Search + CSV Export */}
      <div className="audit-controls-panel">
        {/* Severity Filter Chips */}
        <div className="audit-filter-chips-row" role="group" aria-label="Filter audit logs by severity">
          <span className="audit-chips-label">Severity Level:</span>
          <div className="audit-filter-chips-group">
            <button
              type="button"
              className={`audit-filter-chip all ${severityFilter === "all" ? "active" : ""}`}
              onClick={() => setSeverityFilter("all")}
              title="Show all severity levels"
            >
              <span>All Levels</span>
              <span className="audit-chip-count">{countBySeverity.all}</span>
            </button>

            <button
              type="button"
              className={`audit-filter-chip info ${severityFilter === "info" ? "active" : ""}`}
              onClick={() => setSeverityFilter("info")}
              title="Filter to Info level events"
            >
              <span className="audit-chip-dot info" />
              <span>Info</span>
              <span className="audit-chip-count">{countBySeverity.info}</span>
            </button>

            <button
              type="button"
              className={`audit-filter-chip warning ${severityFilter === "warning" ? "active" : ""}`}
              onClick={() => setSeverityFilter("warning")}
              title="Filter to Warning level events"
            >
              <span className="audit-chip-dot warning" />
              <span>Warning</span>
              <span className="audit-chip-count">{countBySeverity.warning}</span>
            </button>

            <button
              type="button"
              className={`audit-filter-chip critical ${severityFilter === "critical" ? "active" : ""}`}
              onClick={() => setSeverityFilter("critical")}
              title="Filter to Critical level events"
            >
              <span className="audit-chip-dot critical" />
              <span>Critical</span>
              <span className="audit-chip-count">{countBySeverity.critical}</span>
            </button>
          </div>
        </div>

        {/* Date Range Picker: Start and End Date Inputs with Presets */}
        <div className="audit-date-picker-row" role="group" aria-label="Filter audit logs by timestamp range">
          <span className="audit-chips-label">Timestamp Range:</span>
          <div className="audit-date-inputs-group">
            <label className="audit-date-label">
              <span className="audit-date-sublabel">From</span>
              <div className="audit-date-input-wrap">
                <input
                  type="date"
                  value={startDate}
                  onChange={(e) => setStartDate(e.target.value)}
                  max={endDate || undefined}
                  aria-label="Filter audit logs starting from this date"
                />
                {startDate && (
                  <button
                    type="button"
                    className="audit-date-clear-btn"
                    onClick={() => setStartDate("")}
                    title="Clear start date"
                    aria-label="Clear start date"
                  >
                    ✕
                  </button>
                )}
              </div>
            </label>

            <span className="audit-date-divider">to</span>

            <label className="audit-date-label">
              <span className="audit-date-sublabel">To</span>
              <div className="audit-date-input-wrap">
                <input
                  type="date"
                  value={endDate}
                  onChange={(e) => setEndDate(e.target.value)}
                  min={startDate || undefined}
                  aria-label="Filter audit logs up to this date"
                />
                {endDate && (
                  <button
                    type="button"
                    className="audit-date-clear-btn"
                    onClick={() => setEndDate("")}
                    title="Clear end date"
                    aria-label="Clear end date"
                  >
                    ✕
                  </button>
                )}
              </div>
            </label>
          </div>

          <div className="audit-date-presets">
            <button
              type="button"
              className={`audit-date-preset-btn ${!startDate && !endDate ? "active" : ""}`}
              onClick={() => applyDatePreset("all")}
            >
              All Time
            </button>
            <button
              type="button"
              className="audit-date-preset-btn"
              onClick={() => applyDatePreset("today")}
            >
              Today
            </button>
            <button
              type="button"
              className="audit-date-preset-btn"
              onClick={() => applyDatePreset("7d")}
            >
              Last 7 Days
            </button>
            <button
              type="button"
              className="audit-date-preset-btn"
              onClick={() => applyDatePreset("30d")}
            >
              Last 30 Days
            </button>
            {(startDate || endDate) && (
              <button
                type="button"
                className="audit-date-reset-link"
                onClick={() => {
                  setStartDate("");
                  setEndDate("");
                }}
                title="Reset date filter"
              >
                Clear Dates
              </button>
            )}
          </div>
        </div>

        {/* Toolbar: Text Search + Counter + Export Button */}
        <div className="toolbar audit-toolbar">
          <label className="search-field audit-search-field">
            Filter audit log stream
            <div className="audit-search-input-wrap">
              <input
                type="text"
                value={filterText}
                onChange={(e) => setFilterText(e.target.value)}
                placeholder="Filter by actor name or action description…"
                aria-label="Filter audit logs by actor name or action description"
              />
              {filterText && (
                <button
                  type="button"
                  className="audit-search-clear-btn"
                  onClick={() => setFilterText("")}
                  title="Clear search text"
                  aria-label="Clear search text"
                >
                  ✕
                </button>
              )}
            </div>
          </label>

          <div className="toolbar-actions audit-toolbar-actions">
            <span className="audit-count-badge">
              {filteredLogs.length} of {rawLogs.length} event{rawLogs.length === 1 ? "" : "s"}
            </span>
            <button
              type="button"
              className="secondary-button compact"
              onClick={exportCsv}
              disabled={filteredLogs.length === 0}
              title={
                filteredLogs.length === 0
                  ? "No logs available to export"
                  : `Export ${filteredLogs.length} audit log${filteredLogs.length === 1 ? "" : "s"} to CSV`
              }
            >
              <Icon name="download" />
              <span>Export to CSV</span>
            </button>
          </div>
        </div>
      </div>

      {/* 4. Empty Search, Date Range, or Severity State */}
      {filteredLogs.length === 0 ? (
        <div className="audit-empty-filtered">
          <p>
            No audit log entries matched the current filters
            {severityFilter !== "all" && <span> (Severity: <strong>{severityFilter}</strong>)</span>}
            {(startDate || endDate) && (
              <span>
                {" "}
                (Timestamp: <strong>{startDate || "earliest"}</strong> to <strong>{endDate || "latest"}</strong>)
              </span>
            )}
            {filterText && <span> matching <strong>"{filterText}"</strong></span>}.
          </p>
          <button
            type="button"
            className="secondary-button compact"
            onClick={() => {
              setFilterText("");
              setSeverityFilter("all");
              setStartDate("");
              setEndDate("");
            }}
          >
            Reset All Filters
          </button>
        </div>
      ) : (
        /* 5. Stream List with Severity Badges and Rail Markers */
        <div className="audit-stream">
          {filteredLogs.map((log) => {
            const sev = computeLogSeverity(log);

            return (
              <article key={log.id} className={`audit-log-item severity-${sev}`}>
                <div className="audit-rail">
                  <span className={`audit-rail-dot severity-${sev}`} />
                </div>
                <div className="audit-log-body">
                  <div className="audit-log-row-top">
                    <p className="audit-log-action-line">
                      <strong>{log.actor}</strong> {humanize(log.action)}{" "}
                      <b>{humanize(log.entityType)}</b>{" "}
                      <span className="mono">#{log.entityId}</span>
                    </p>
                    <span className={`audit-severity-tag ${sev}`}>
                      {sev === "critical" && "● "}
                      {sev === "warning" && "▲ "}
                      {sev === "info" && "ℹ "}
                      {sev.toUpperCase()}
                    </span>
                  </div>
                  <small className="audit-log-meta-line">
                    {formatDateTime(log.createdAt)}
                    {log.reason ? ` · Reason: ${log.reason}` : ""}
                  </small>
                </div>
              </article>
            );
          })}
        </div>
      )}
    </div>
  );
}

function CustomAuditTooltip({
  active,
  payload,
  label,
}: {
  active?: boolean;
  payload?: Array<{
    payload: {
      date: string;
      displayDate: string;
      fullDate: string;
      info: number;
      warning: number;
      critical: number;
      total: number;
    };
  }>;
  label?: string;
}) {
  if (active && payload && payload.length > 0 && payload[0]) {
    const data = payload[0].payload;
    return (
      <div className="audit-chart-tooltip">
        <div className="audit-chart-tooltip-header">
          <strong>{data.fullDate || data.displayDate || label}</strong>
          <span className="audit-chart-tooltip-total">
            {data.total} event{data.total === 1 ? "" : "s"}
          </span>
        </div>
        <div className="audit-chart-tooltip-rows">
          <div className="audit-chart-tooltip-row critical">
            <span className="audit-chart-tooltip-dot critical" />
            <span className="audit-chart-tooltip-lbl">Critical Safeguards</span>
            <strong className="audit-chart-tooltip-val">{data.critical}</strong>
          </div>
          <div className="audit-chart-tooltip-row warning">
            <span className="audit-chart-tooltip-dot warning" />
            <span className="audit-chart-tooltip-lbl">Warnings & Adjustments</span>
            <strong className="audit-chart-tooltip-val">{data.warning}</strong>
          </div>
          <div className="audit-chart-tooltip-row info">
            <span className="audit-chart-tooltip-dot info" />
            <span className="audit-chart-tooltip-lbl">Routine Info</span>
            <strong className="audit-chart-tooltip-val">{data.info}</strong>
          </div>
        </div>
      </div>
    );
  }
  return null;
}

function PayrollPeriodPicker({ token, value, onChange }: { token: string; value: number | null; onChange: (value: number | null) => void }) {
  const periods = useQuery({ queryKey: ["periods", token], queryFn: () => api.listPeriods({ token }) });
  return <div className="production-scope stage6-scope"><label>Payroll period<select aria-label="Payroll period" value={value ?? ""} onChange={(event) => onChange(event.target.value ? Number(event.target.value) : null)}><option value="">Select a payroll period</option>{periods.data?.periods.map((period) => <option key={period.id} value={period.id}>{period.periodId} · {period.status}</option>)}</select></label></div>;
}

function MasterSalary({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const qc = useQueryClient(); const [periodId, setPeriodId] = useState<number | null>(null); const [adjusting, setAdjusting] = useState<number | null>(null); const allowed = hasPermission(permissions, "master.read");
  const query = useQuery({ queryKey: ["master-salary", token, periodId], queryFn: () => api.getMasterSalary({ token, payrollPeriodId: periodId ?? 0 }), enabled: allowed && Boolean(periodId) });
  const generate = useMutation({ mutationFn: () => api.generateMasterSalary({ token, payrollPeriodId: periodId ?? 0 }), onSuccess: (data) => { notify({ kind: data.status === "completed" ? "success" : "error", text: data.status === "completed" ? `Master Salary V${data.version} completed for ${data.employeeCount} employees.` : `Master Salary V${data.version} is blocked by ${data.exceptionCount} exceptions.` }); qc.invalidateQueries({ queryKey: ["master-salary"] }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  const adjust = useMutation({ mutationFn: (payload: { adjustmentType: "bonus" | "deduction" | "other"; amount: string; reason: string }) => api.addMasterAdjustment({ token, payrollPeriodId: periodId ?? 0, employeeId: adjusting ?? 0, ...payload }), onSuccess: (data) => { setAdjusting(null); notify({ kind: "success", text: data.status === "pending_revision" ? "Correction request recorded for the locked period." : "Adjustment authorized. Generate a new Master Salary version to apply it." }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  if (!allowed) return <Empty title="Master Salary access is restricted" body="Your role does not include consolidated payroll results." />;
  const run = query.data?.run;
  return <><PayrollPeriodPicker token={token} value={periodId} onChange={setPeriodId} />{!periodId ? <Empty title="Select a payroll period" body="Master Salary consolidates one completed salary version per active project." /> : query.isPending ? <Loading /> : query.error ? <Empty title="Master Salary could not load" body={message(query.error)} /> : <><section className="run-command stage6-command"><div><p className="overline">Stable Employee ID consolidation</p><h2>{run ? `${run.runId} · Version ${run.version}` : "No Master Salary version"}</h2><p>{run ? `${run.projectCount} projects · ${run.employeeCount} employees · generated ${formatDate(run.generatedAt)}` : "Complete project salary calculations, then generate the consolidated period snapshot."}</p></div>{hasPermission(permissions, "master.generate") && <button className="primary-button" disabled={generate.isPending || ["Approved", "Locked", "Exported"].includes(query.data?.period.status ?? "")} onClick={() => generate.mutate()}>{generate.isPending ? "Consolidating…" : run ? "Generate new version" : "Generate Master Salary"}</button>}</section>{run && <div className="master-metrics"><span><b>{run.grossTotal}</b>Gross salary</span><span><b>{run.bonusTotal}</b>Bonuses</span><span><b>{run.deductionTotal}</b>Deductions</span><span className="payable"><b>{run.finalTotal}</b>Final payable</span><span className={run.exceptionCount ? "risk" : "clear"}><b>{run.exceptionCount}</b>Exceptions</span></div>}{query.data?.exceptions.length ? <section className="stage6-exceptions"><h2>Blocking reconciliation exceptions</h2>{query.data.exceptions.map((item) => <article key={item.id}><Status value="blocking" /><div><strong>{humanize(item.type)}</strong><p>{item.explanation}</p><small>{[item.projectCode, item.employeeCode].filter(Boolean).join(" · ")}</small></div></article>)}</section> : run ? <div className="success-strip">Project, employee, adjustment, rounding, and master totals reconcile for this version.</div> : null}{query.data?.rows.length ? <div className="master-table-wrap"><table className="master-table"><thead><tr><th>Employee</th>{query.data.projects.map((project) => <th key={project.id}>{project.projectId}</th>)}<th>Production</th><th>Doctype</th><th>Proof</th><th>Other</th><th>Bonuses</th><th>Deductions</th><th>Gross</th><th>Final payable</th><th /></tr></thead><tbody>{query.data.rows.map((row) => <tr key={row.id}><th><strong>{row.employeeCode}</strong><span>{row.employeeName}</span>{row.previousIds.length > 0 && <small>Previous: {row.previousIds.join(", ")}</small>}</th>{query.data!.projects.map((project) => <td key={project.id}>{row.projectAmounts[String(project.id)] ?? "0.00"}</td>)}<td>{row.productionSalary}</td><td>{row.doctypeSalary}</td><td>{row.proofSalary}</td><td>{row.otherComponents}</td><td>{row.bonuses}</td><td>{row.deductions}</td><td>{row.grossSalary}</td><td className="final-cell">{row.finalPayableSalary}</td><td>{hasPermission(permissions, "salary.adjust") && <button className="text-button" onClick={() => setAdjusting(row.employeeId)}>Adjust</button>}</td></tr>)}</tbody></table></div> : run ? <Empty title="No employee salary rows" body="The selected approved project salary versions do not contain employee components." /> : null}</>}{adjusting && <Modal title="Controlled salary adjustment" onClose={() => setAdjusting(null)}><form className="stack-form" onSubmit={(event) => { event.preventDefault(); const form = new FormData(event.currentTarget); adjust.mutate({ adjustmentType: field(form, "adjustmentType") as "bonus" | "deduction" | "other", amount: field(form, "amount"), reason: field(form, "reason") }); }}><p className="form-lead">Every adjustment records the authorizer, reason, and timestamp. Locked periods create a correction request instead of changing payroll.</p><label>Adjustment type<select name="adjustmentType"><option value="bonus">Bonus</option><option value="deduction">Deduction</option><option value="other">Other approved component</option></select></label><label>Amount<input name="amount" inputMode="decimal" required placeholder="0.00" /></label><label>Reason<textarea name="reason" minLength={3} rows={4} required /></label><button className="primary-button" disabled={adjust.isPending}>Record adjustment</button></form></Modal>}</>;
}

function PayrollPeriodTimeline({
  status,
  periodId,
  latestMaster,
  records,
  approvalAt,
  lockAt,
  blocking,
}: {
  status: string;
  periodId: string;
  latestMaster: { id: number; runId: string; version: number; status: string; exceptionCount: number; finalTotal: string } | null;
  records: Array<{ id: number; action: string; fromStatus: string; toStatus: string; reason: string; actor: string; actedAt: string }>;
  approvalAt?: string | null;
  lockAt?: string | null;
  blocking: number;
}) {
  const rankMap: Record<string, number> = {
    Draft: 0,
    Calculated: 1,
    "Under Review": 2,
    Approved: 3,
    Locked: 4,
    Exported: 5,
  };
  const currentRank = rankMap[status] ?? (status === "active" ? 0 : 1);
  const progressPercent = currentRank === 0 ? 10 : currentRank === 1 ? 25 : currentRank === 2 ? 50 : currentRank === 3 ? 75 : currentRank === 4 ? 90 : 100;
  const railFillWidth = Math.min(100, Math.max(0, ((currentRank - 1) / 4) * 100));

  const reviewRecord = records.find((r) => r.toStatus === "Under Review" || r.action === "submitted_for_review");
  const approvedRecord = records.find((r) => r.toStatus === "Approved" || r.action === "approved");
  const lockedRecord = records.find((r) => r.toStatus === "Locked" || r.action === "locked");
  const exportedRecord = records.find((r) => r.toStatus === "Exported" || r.action === "exported");
  const returnedRecord = records.find((r) => r.action === "returned");

  const milestones = [
    {
      key: "calculation",
      num: 1,
      label: "Calculation",
      sub: "Master salary run",
      state: currentRank >= 1 ? ("completed" as const) : ("current" as const),
      tag: currentRank >= 1 ? "Completed" : "In Progress",
      meta: latestMaster ? `Master ${latestMaster.runId} (v${latestMaster.version})` : currentRank >= 1 ? "Calculation complete" : "Calculation pending",
      icon: "calculator" as const,
    },
    {
      key: "review",
      num: 2,
      label: "Review",
      sub: "Submitted for verification",
      state: currentRank >= 3 ? ("completed" as const) : currentRank === 2 ? ("current" as const) : ("upcoming" as const),
      tag: currentRank >= 3 ? "Reviewed" : currentRank === 2 ? "In Review" : currentRank === 1 ? "Ready" : "Upcoming",
      meta: reviewRecord ? `Submitted by ${reviewRecord.actor}` : currentRank === 2 ? "Under active review" : currentRank === 1 ? "Ready to submit" : "Pending submission",
      icon: "shield" as const,
    },
    {
      key: "approved",
      num: 3,
      label: "Approved",
      sub: "Sign-off & authorization",
      state: currentRank >= 4 ? ("completed" as const) : currentRank === 3 ? ("current" as const) : ("upcoming" as const),
      tag: currentRank >= 4 ? "Approved" : currentRank === 3 ? "Signed Off" : "Upcoming",
      meta: approvedRecord ? `Approved by ${approvedRecord.actor}` : approvalAt ? `Approved ${formatDate(approvalAt)}` : blocking > 0 ? `${blocking} blockers pending` : "Requires zero blockers",
      icon: "check" as const,
    },
    {
      key: "locked",
      num: 4,
      label: "Locked",
      sub: "Frozen against changes",
      state: currentRank >= 5 ? ("completed" as const) : currentRank === 4 ? ("current" as const) : ("upcoming" as const),
      tag: currentRank >= 5 ? "Locked" : currentRank === 4 ? "Frozen" : "Upcoming",
      meta: lockedRecord ? `Locked by ${lockedRecord.actor}` : lockAt ? `Locked ${formatDate(lockAt)}` : "Frozen before export",
      icon: "lock" as const,
    },
    {
      key: "exported",
      num: 5,
      label: "Exported",
      sub: "Release & audit workbooks",
      state: currentRank >= 5 ? ("completed" as const) : ("upcoming" as const),
      tag: currentRank >= 5 ? "Exported" : "Upcoming",
      meta: exportedRecord ? `Exported by ${exportedRecord.actor}` : "Final distribution",
      icon: "download" as const,
    },
  ];

  return (
    <section className="period-timeline-card" aria-label="Payroll period progress and status milestones">
      <div className="timeline-header">
        <div className="timeline-title-wrap">
          <div>
            <p className="overline" style={{ margin: 0, letterSpacing: "0.06em" }}>Lifecycle progression</p>
            <h2 style={{ margin: "2px 0 0", fontSize: "19px", fontWeight: 750 }}>
              Payroll Period Milestone Tracker
            </h2>
          </div>
          <span className={`timeline-step-badge ${currentRank === 5 ? "is-complete" : "is-active"}`}>
            {currentRank === 5 ? "100% Finalized" : `Milestone ${Math.min(5, Math.max(1, currentRank))} of 5: ${status}`}
          </span>
        </div>
        <div className="timeline-progress-summary">
          <div className="timeline-progress-bar-wrap" title={`Overall completion: ${progressPercent}%`}>
            <div className="timeline-progress-bar-fill" style={{ width: `${progressPercent}%` }} />
          </div>
          <span className="timeline-progress-percent">{progressPercent}%</span>
        </div>
      </div>

      <div className="timeline-track-container">
        <div className="timeline-rail" aria-hidden="true">
          <div className="timeline-rail-fill" style={{ width: `${railFillWidth}%` }} />
        </div>
        <div className="timeline-steps">
          {milestones.map((m) => (
            <div key={m.key} className={`timeline-step ${m.state}`}>
              <div className="timeline-node" aria-label={`Milestone ${m.num}: ${m.label} (${m.state})`}>
                {m.state === "completed" ? (
                  <Icon name="check" />
                ) : (
                  <span>{m.num}</span>
                )}
              </div>
              <div className="timeline-step-content">
                <span className="timeline-step-label">{m.label}</span>
                <span className="timeline-step-tag">{m.tag}</span>
                <span className="timeline-step-sub">{m.sub}</span>
                <span className="timeline-step-meta">{m.meta}</span>
              </div>
            </div>
          ))}
        </div>
      </div>

      <div className={`timeline-callout ${status === "Under Review" && blocking > 0 ? "has-blocker" : currentRank === 5 ? "is-complete" : ""}`}>
        <div>
          {status === "Calculated" && returnedRecord ? (
            <span>
              <strong>Returned for correction:</strong> &ldquo;{returnedRecord.reason}&rdquo; by {returnedRecord.actor}. Revise calculations and submit for review.
            </span>
          ) : status === "Calculated" ? (
            <span>
              <strong>Calculation complete:</strong> Ready for Reviewer to sign off and submit this period for official review.
            </span>
          ) : status === "Under Review" && blocking > 0 ? (
            <span>
              <strong>Blocking exceptions detected:</strong> {blocking} unresolved exception{blocking > 1 ? "s" : ""} must be resolved in Reconciliation / Salary before Approver can approve.
            </span>
          ) : status === "Under Review" ? (
            <span>
              <strong>Clear for approval:</strong> All reconciliation, salary, and master exceptions are resolved. Ready for Approver sign-off.
            </span>
          ) : status === "Approved" ? (
            <span>
              <strong>Period approved:</strong> Authorized by approver. Administrator can now lock the payroll period to guarantee immutability.
            </span>
          ) : status === "Locked" ? (
            <span>
              <strong>Locked and protected:</strong> The payroll period is immutable. Head to Export Center to generate and release official workbooks.
            </span>
          ) : status === "Exported" ? (
            <span>
              <strong>Period fully exported:</strong> All calculations, approvals, locks, and report distributions are finalized.
            </span>
          ) : (
            <span>
              <strong>Preparation in progress:</strong> Complete salary calculations to advance to the review milestone.
            </span>
          )}
        </div>
        <small style={{ color: "var(--dim)", whiteSpace: "nowrap" }}>
          Period {periodId}
        </small>
      </div>
    </section>
  );
}

function ReviewApproval({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const qc = useQueryClient(); const [periodId, setPeriodId] = useState<number | null>(null); const [reason, setReason] = useState(""); const allowed = hasPermission(permissions, "master.read"); const query = useQuery({ queryKey: ["payroll-approval", token, periodId], queryFn: () => api.getPayrollApproval({ token, payrollPeriodId: periodId ?? 0 }), enabled: allowed && Boolean(periodId) });
  const advance = useMutation({ mutationFn: (action: "submitted_for_review" | "approved" | "returned" | "locked" | "exported") => api.advancePayrollApproval({ token, payrollPeriodId: periodId ?? 0, action, reason }), onSuccess: (data) => { setReason(""); notify({ kind: "success", text: `Payroll moved to ${data.status}.` }); qc.invalidateQueries({ queryKey: ["payroll-approval"] }); qc.invalidateQueries({ queryKey: ["periods"] }); }, onError: (error) => notify({ kind: "error", text: message(error) }) });
  if (!allowed) return <Empty title="Approval access is restricted" body="Your role does not include payroll review records." />;
  const status = query.data?.period.status ?? ""; const blocking = query.data ? query.data.blocking.reconciliation + query.data.blocking.salary + query.data.blocking.master : 0;
  return <><PayrollPeriodPicker token={token} value={periodId} onChange={setPeriodId} />{!periodId ? <Empty title="Select a payroll period" body="The governed path is Under Review, Approved, Locked, then Exported." /> : query.isPending ? <Loading /> : query.error ? <Empty title="Approval state could not load" body={message(query.error)} /> : <><PayrollPeriodTimeline status={status} periodId={query.data?.period.periodId ?? String(periodId)} latestMaster={query.data?.latestMaster ?? null} records={query.data?.records ?? []} approvalAt={query.data?.period.approvalAt} lockAt={query.data?.period.lockAt} blocking={blocking} /><div className="approval-layout"><section className="approval-gate"><div><p className="overline">Current state</p><div className="approval-status"><Status value={status} /><strong>{query.data?.period.periodId}</strong></div><p>{query.data?.latestMaster ? `Master ${query.data.latestMaster.runId} · ${query.data.latestMaster.finalTotal} payable` : "A completed Master Salary version is required before review."}</p></div><div className="blocker-grid"><span className={query.data?.blocking.reconciliation ? "risk" : "clear"}><b>{query.data?.blocking.reconciliation}</b>Reconciliation blockers</span><span className={query.data?.blocking.salary ? "risk" : "clear"}><b>{query.data?.blocking.salary}</b>Salary blockers</span><span className={query.data?.blocking.master ? "risk" : "clear"}><b>{query.data?.blocking.master}</b>Master blockers</span></div><label>Decision reason<textarea value={reason} onChange={(event) => setReason(event.target.value)} rows={3} placeholder="Required for every transition" /></label><div className="approval-actions">{status === "Calculated" && hasPermission(permissions, "payroll.review") && <button className="primary-button" disabled={reason.trim().length < 3 || !query.data?.latestMaster || advance.isPending} onClick={() => advance.mutate("submitted_for_review")}>Submit for review</button>}{status === "Under Review" && hasPermission(permissions, "payroll.review") && <button className="secondary-button" disabled={reason.trim().length < 3 || advance.isPending} onClick={() => advance.mutate("returned")}>Return for correction</button>}{status === "Under Review" && hasPermission(permissions, "payroll.approve") && <button className="primary-button" disabled={reason.trim().length < 3 || blocking > 0 || advance.isPending} onClick={() => advance.mutate("approved")}>Approve payroll</button>}{status === "Approved" && hasPermission(permissions, "payroll.lock") && <button className="primary-button" disabled={reason.trim().length < 3 || advance.isPending} onClick={() => advance.mutate("locked")}>Lock payroll</button>}{status === "Locked" && hasPermission(permissions, "payroll.lock") && <button className="primary-button" disabled={reason.trim().length < 3 || advance.isPending} onClick={() => advance.mutate("exported")}>Mark exported</button>}</div><p className="approval-help">Reviewer submits, Approver approves, and Administrator locks. Approval is unavailable while any blocker remains.</p></section><section className="approval-history"><h2>Decision history</h2>{query.data?.records.length ? query.data.records.map((item) => <article key={item.id}><div><strong>{humanize(item.action)}</strong><Status value={item.toStatus} /></div><p>{item.reason}</p><small>{item.actor} · {formatDate(item.actedAt)} · {item.fromStatus} → {item.toStatus}</small></article>) : <Empty title="No approval decisions" body="The first signed record appears when a reviewer submits payroll." />}</section></div></>}</>;
}

const EXPORT_TYPES = [
  ["project_salary", "Project-wise Salary Workbook"], ["employee_salary", "Employee-wise Salary Details"], ["master_salary", "Master Salary Workbook"], ["production_reconciliation", "Production Reconciliation Report"], ["transmitted_reconciliation", "Transmitted Batch Reconciliation Report"], ["shipdet_reconciliation", "Shipdet Reconciliation Report"], ["doctype_salary", "Doctype Salary Report"], ["proof_salary", "Proof Salary Report"], ["exception_report", "Exception Report"], ["payroll_audit", "Payroll Audit Report"],
] as const;

function ExportCenter({ token, permissions, notify }: { token: string; permissions: string[]; notify: (notice: Notice) => void }) {
  const qc = useQueryClient();
  const [periodId, setPeriodId] = useState<number | null>(null);
  const [pdfLoading, setPdfLoading] = useState<string | null>(null);
  const allowed = hasPermission(permissions, "master.read");

  const jobs = useQuery({
    queryKey: ["export-jobs", token, periodId],
    queryFn: () => api.listExportJobs({ token, payrollPeriodId: periodId ?? 0 }),
    enabled: allowed && Boolean(periodId),
    refetchInterval: (query) => query.state.data?.jobs.some((job) => job.status === "queued" || job.status === "processing") ? 1500 : false,
  });

  const create = useMutation({
    mutationFn: (exportType: typeof EXPORT_TYPES[number][0]) => api.createExportJob({ token, payrollPeriodId: periodId ?? 0, exportType }),
    onSuccess: (data) => {
      notify({ kind: "success", text: `${data.filename} generated successfully with ${data.rowCount} rows.` });
      qc.invalidateQueries({ queryKey: ["export-jobs"] });
      if (data.downloadUrl) window.open(data.downloadUrl, "_blank", "noopener,noreferrer");
    },
    onError: (error) => notify({ kind: "error", text: message(error) }),
  });

  const handlePrintPdf = async (exportType: typeof EXPORT_TYPES[number][0], reportLabel: string) => {
    if (!periodId) return;
    setPdfLoading(exportType);
    try {
      const data = await api.getReportPreviewData({
        token,
        payrollPeriodId: periodId,
        exportType,
      });

      printDocument(
        data.title,
        data.subtitle,
        data.headers,
        data.rows,
        {
          "Payroll Period": String(periodId),
          "Report Type": reportLabel,
          "Total Records": String(data.rows.length),
        }
      );
      notify({ kind: "success", text: `Ready-to-print PDF preview opened for ${reportLabel}.` });
    } catch (err) {
      notify({ kind: "error", text: message(err) });
    } finally {
      setPdfLoading(null);
    }
  };

  if (!allowed) return <Empty title="Export access is restricted" body="Your role does not include controlled payroll reports." />;

  return (
    <>
      <PayrollPeriodPicker token={token} value={periodId} onChange={setPeriodId} />
      {!periodId ? (
        <Empty title="Select a payroll period" body="Exports are generated from one completed Master Salary snapshot and its underlying approved versions." />
      ) : (
        <>
          <section className="export-catalog">
            <div className="export-intro">
              <p className="overline">Controlled distribution</p>
              <h2>English Workbooks & Print-Ready PDF Documents</h2>
              <p>
                Each export provides frozen headers, formatted currency and numeric values with comma separators and zero decimals, left alignment for text, right alignment for numbers, export date and time, and clean print styles.
              </p>
            </div>
            <div className="export-grid">
              {EXPORT_TYPES.map(([key, label]) => {
                const isExcelPending = create.isPending && create.variables === key;
                const isPdfPending = pdfLoading === key;
                return (
                  <div key={key} className="export-item-card" style={{ background: "#fff", border: "1px solid #d4ded9", borderRadius: 8, padding: "16px", display: "flex", flexDirection: "column", justifyContent: "space-between", gap: 12 }}>
                    <div>
                      <strong style={{ fontSize: "14px", color: "#073b33", display: "block", marginBottom: 4 }}>{label}</strong>
                      <span style={{ fontSize: "12px", color: "var(--dim)" }}>Standard A4 print & Excel workbook</span>
                    </div>
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                      <button
                        type="button"
                        className="secondary-button compact"
                        disabled={!hasPermission(permissions, "exports.manage") || isExcelPending || isPdfPending}
                        onClick={() => create.mutate(key)}
                        title="Generate XLSX Spreadsheet"
                      >
                        <Icon name="download" />
                        <span>{isExcelPending ? "Generating…" : "Excel (XLSX)"}</span>
                      </button>
                      <button
                        type="button"
                        className="primary-button compact"
                        disabled={isExcelPending || isPdfPending}
                        onClick={() => handlePrintPdf(key, label)}
                        title="Open Ready-to-Print PDF Document"
                      >
                        <Icon name="printer" />
                        <span>{isPdfPending ? "Preparing PDF…" : "PDF / Print"}</span>
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          </section>

          <section className="job-board">
            <h2>Export job history</h2>
            {jobs.isPending ? (
              <Loading />
            ) : jobs.error ? (
              <Empty title="Export jobs could not load" body={message(jobs.error)} />
            ) : jobs.data?.jobs.length ? (
              <div className="job-list">
                {jobs.data.jobs.map((job) => (
                  <article key={job.id}>
                    <div>
                      <strong>{EXPORT_TYPES.find(([key]) => key === job.exportType)?.[1] ?? humanize(job.exportType)}</strong>
                      <span>{job.jobId}</span>
                    </div>
                    <div className="job-progress">
                      <i style={{ width: `${job.progress}%` }} />
                      <span>{job.progress}%</span>
                    </div>
                    <div>
                      <Status value={job.status} />
                      <span>{job.rowCount} rows · {job.generatedBy} · {formatDate(job.completedAt ?? job.createdAt)}</span>
                      {job.downloadUrl && (
                        <a className="text-button" href={job.downloadUrl} target="_blank" rel="noreferrer" download={job.filename ?? undefined}>
                          <Icon name="download" /> Download XLSX
                        </a>
                      )}
                    </div>
                    {job.errorMessage && <p className="error-text">{job.errorMessage}</p>}
                  </article>
                ))}
              </div>
            ) : (
              <Empty title="No export jobs" body="Choose a report above to create the first controlled workbook." />
            )}
          </section>
        </>
      )}
    </>
  );
}


function Planned({ page, go }: { page: PageKey; go: (page: PageKey) => void }) {
  const stage = "Stage 6";
  return <section className="planned-page"><span>{stage}</span><h2>{PAGE_TITLES[page].title} is intentionally not active yet.</h2><p>{PAGE_TITLES[page].subtitle} No control on this page claims to import, calculate, approve, or export payroll before the durable workflow exists.</p><button className="secondary-button" onClick={() => go("dashboard")}>Return to foundation overview</button></section>;
}

function parseCsvText(text: string) {
  const source = text.replace(/^\uFEFF/, "").trim();
  if (!source) return [];
  const firstLine = source.split(/\r?\n/).find((l) => l.trim().length > 0) ?? "";
  const counts = {
    ",": (firstLine.match(/,/g) || []).length,
    ";": (firstLine.match(/;/g) || []).length,
    "\t": (firstLine.match(/\t/g) || []).length,
    "|": (firstLine.match(/\|/g) || []).length,
  };
  const delim = Object.entries(counts).sort((a, b) => b[1] - a[1])[0]?.[0] || ",";

  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  for (let i = 0; i < source.length; i += 1) {
    const char = source[i] ?? "";
    const next = source[i + 1] ?? "";
    if (char === '"' && quoted && next === '"') {
      cell += '"';
      i += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === delim && !quoted) {
      row.push(cell.trim());
      cell = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") i += 1;
      row.push(cell.trim());
      rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }
  if (cell || row.length) {
    row.push(cell.trim());
    rows.push(row);
  }
  return rows.filter((r) => r.some((c) => c && c.trim().length > 0));
}
function normalizeHeader(value: string) { return value.toLowerCase().replace(/[^a-z0-9]+/g, ""); }
function normalizeEmployeeStatus(value: string): EmployeeDraft["status"] {
  const v = value.toLowerCase().trim();
  if (v.includes("separ") || v.includes("left") || v.includes("resign") || v.includes("term")) return "separated";
  if (v.includes("leave")) return "leave";
  if (v.includes("inact")) return "inactive";
  return "active";
}
function normalizeDate(value: string): string {
  if (!value) return "";
  const s = value.trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const dmy = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  if (dmy) {
    const p1 = Number(dmy[1]);
    const p2 = Number(dmy[2]);
    const year = dmy[3];
    let day = p1;
    let month = p2;
    if (p1 <= 12 && p2 > 12) {
      month = p1;
      day = p2;
    }
    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }
  const ymd = s.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
  if (ymd) {
    return `${ymd[1]}-${String(Number(ymd[2])).padStart(2, "0")}-${String(Number(ymd[3])).padStart(2, "0")}`;
  }
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return parsed.toISOString().slice(0, 10);
  }
  return "";
}
function formatBytes(value: number) { if (value < 1024) return `${value} B`; if (value < 1024 * 1024) return `${(value / 1024).toFixed(1)} KB`; return `${(value / 1024 / 1024).toFixed(1)} MB`; }
