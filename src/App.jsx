import { useEffect, useMemo, useRef, useState } from "react";
import { brandLogo } from "./brandLogo";
import {
  MdAdd,
  MdArrowOutward,
  MdAssignment,
  MdCalendarMonth,
  MdCheckCircle,
  MdContentPasteSearch,
  MdDownload,
  MdDelete,
  MdEdit,
  MdLock,
  MdLogout,
  MdNotificationsNone,
  MdRemove,
  MdReplay,
  MdSearch,
  MdScience,
  MdTune,
} from "react-icons/md";
import * as XLSX from "xlsx-js-style";

const equipment = {
  name: "生物显微镜",
  model: "CX23",
  assetNo: "EQP-2024-000123",
  sn: "SN202407200001",
  maker: "OLYMPUS",
  location: "A区-显微镜室-01架",
  keeper: "李四",
  boughtAt: "2024-03-15",
};

// ---- 演示模式：不连接任何云端，数据只保存在本机浏览器 ----
const DEMO_ADMIN_ID = "admin";
const DEMO_ADMIN_PASSWORD = "demo123";
const STORE_KEY = "lab-equipment-demo-v1";

const seedGroups = ["理化组", "电气组", "材料组"];
const seedDevices = [
  ["理化组", "电子天平", "BAL-001"], ["理化组", "移液器套装 P1000", "PIP-003"], ["理化组", "pH 计", "PH-002"],
  ["理化组", "生物显微镜 CX23", "MIC-001"], ["电气组", "数字万用表", "DMM-007"], ["电气组", "示波器", "OSC-002"],
  ["电气组", "直流稳压电源", "PSU-004"], ["电气组", "绝缘电阻测试仪", "IRT-001"], ["材料组", "游标卡尺", "CAL-011"],
  ["材料组", "红外测温仪", "IRT-T05"], ["材料组", "恒温干燥箱", "OVN-001"], ["材料组", "邵氏硬度计", "HRD-002"],
].map(([group_name, device_name, device_no], index) => ({ id: index + 1, group_name, device_name, device_no, created_at: index }));

function demoRecord(id, type, person, group, asset, deviceNo, borrowedMinutesAgo, dueInMinutes, note = "") {
  const borrowed = new Date(Date.now() - borrowedMinutesAgo * 60_000);
  const due = new Date(Date.now() + dueInMinutes * 60_000);
  return { id, type, person, group, asset, deviceNo, when: formatDate(borrowed), borrowedAt: formatDate(borrowed), dueAt: type === "领用" ? formatDate(due) : "", note, status: "已提交" };
}

function buildSeed() {
  return {
    groups: seedGroups,
    devices: seedDevices,
    records: [
      demoRecord(6, "领用", "E1003", "电气组", "示波器", "OSC-002", 60 * 30, -60 * 6, "电源模块测试"),
      demoRecord(5, "领用", "E1001", "理化组", "电子天平", "BAL-001", 50, 6, "称量样品"),
      demoRecord(4, "领用", "E1005", "材料组", "游标卡尺", "CAL-011", 120, 60 * 24, "来料尺寸抽检"),
      demoRecord(3, "归还", "E1002", "电气组", "数字万用表", "DMM-007", 180, 0),
      demoRecord(2, "领用", "E1002", "电气组", "数字万用表", "DMM-007", 60 * 26, -60 * 2),
      demoRecord(1, "领用", "E1004", "理化组", "pH 计", "PH-002", 60 * 5, 60 * 48, "溶液配制"),
    ],
  };
}

function loadStore() {
  try {
    const cached = JSON.parse(window.localStorage.getItem(STORE_KEY) || "null");
    if (cached?.records && cached?.devices && cached?.groups) return cached;
  } catch { /* 浏览器禁用存储时使用示例数据 */ }
  return buildSeed();
}

function formatDate(date) {
  return new Intl.DateTimeFormat("zh-CN", {
    year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false,
  }).format(date).replaceAll("/", "-");
}

function calculateDueAt(startTime, duration, unit) {
  const due = new Date(startTime);
  if (unit === "天") due.setDate(due.getDate() + duration);
  else due.setMinutes(due.getMinutes() + duration);
  return formatDate(due);
}

function formatHeaderDate(date) {
  return new Intl.DateTimeFormat("zh-CN", { year: "numeric", month: "2-digit", day: "2-digit", weekday: "long" }).format(date);
}

function toAppRecord(row) {
  return {
    id: row.id,
    type: row.operation_type,
    person: row.staff_id,
    asset: row.asset_name,
    when: formatDate(new Date(row.borrowed_at || row.created_at)),
    group: row.group_name,
    deviceNo: row.serial_no || "",
    borrowedAt: formatDate(new Date(row.borrowed_at || row.created_at)),
    dueAt: row.due_at ? formatDate(new Date(row.due_at)) : "",
    note: row.note || "",
    status: "已提交",
  };
}

function recordTime(record) {
  return new Date((record.borrowedAt || record.when || "").replace(" ", "T")).getTime();
}

function hasLaterReturn(record, records) {
  if (record.type !== "领用") return false;
  const borrowedTime = recordTime(record);
  const hasDeviceNo = Boolean(record.deviceNo);
  return records.some((candidate) => candidate.type === "归还"
    && candidate.asset === record.asset
    && (!record.group || !candidate.group || candidate.group === record.group)
    && (!hasDeviceNo || candidate.deviceNo === record.deviceNo)
    && recordTime(candidate) >= borrowedTime);
}

function getReminderLevel(record, now = Date.now(), records = []) {
  if (record.type !== "领用" || !record.dueAt) return null;
  if (hasLaterReturn(record, records)) return null;
  const dueTime = new Date(record.dueAt.replace(" ", "T")).getTime();
  if (!Number.isFinite(dueTime)) return null;
  const remaining = dueTime - now;
  if (remaining < 0) return "overdue";
  if (remaining <= 10 * 60 * 1000) return "warning";
  return null;
}

function getRecordStatus(record, now = Date.now(), records = []) {
  if (record.type === "归还") return { label: "已归还", kind: "returned" };
  if (hasLaterReturn(record, records)) return { label: "已归还", kind: "returned" };
  const dueTime = record.dueAt ? new Date(record.dueAt.replace(" ", "T")).getTime() : NaN;
  if (Number.isFinite(dueTime) && dueTime < now) return { label: "逾期未还", kind: "overdue" };
  return { label: "\u501f\u7528\u4e2d", kind: "borrowing" };
}

function exportExcel(records, now = Date.now()) {
  const header = ["操作类型", "领用人工号", "所属组别", "设备名称", "设备编号", "借用时间", "预计归还时间", "备注", "实时状态"];
  const rows = records.map((item) => {
    const status = getRecordStatus(item, now, records);
    return [item.type, item.person, item.group || "", item.asset, item.deviceNo || "", item.borrowedAt || item.when, item.dueAt || "", item.note || "", status.label];
  });
  const sheet = XLSX.utils.aoa_to_sheet([header, ...rows]);
  sheet["!cols"] = [{ wch: 12 }, { wch: 16 }, { wch: 18 }, { wch: 24 }, { wch: 18 }, { wch: 20 }, { wch: 20 }, { wch: 32 }, { wch: 14 }];
  const headerStyle = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "1F5FBF" } }, alignment: { horizontal: "center" } };
  header.forEach((_, index) => { sheet[XLSX.utils.encode_cell({ r: 0, c: index })].s = headerStyle; });
  records.forEach((record, index) => {
    const status = getRecordStatus(record, now, records);
    const color = status.kind === "overdue" ? "FCE4E4" : status.kind === "returned" ? "E2F0D9" : "FFF2CC";
    header.forEach((_, column) => {
      const cell = sheet[XLSX.utils.encode_cell({ r: index + 1, c: column })];
      cell.s = { fill: { fgColor: { rgb: color } }, alignment: { vertical: "center" } };
    });
  });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "借还记录");
  XLSX.writeFile(workbook, "实验室设备借还记录.xlsx", { compression: true });
}

export function App() {
  const [mode, setMode] = useState("领用");
  const [unit, setUnit] = useState("天");
  const [duration, setDuration] = useState(2);
  const [initial] = useState(loadStore);
  const [staffId, setStaffId] = useState("E1006");
  const [group, setGroup] = useState("材料组");
  const [groups, setGroups] = useState(initial.groups);
  const groupRows = useMemo(() => groups.map((name) => ({ id: name, name })), [groups]);
  const [asset, setAsset] = useState("恒温干燥箱");
  const [sn, setSn] = useState("OVN-001");
  const [note, setNote] = useState("");
  const [records, setRecords] = useState(initial.records);
  const [view, setView] = useState("form");
  const [recordFilter, setRecordFilter] = useState("");
  const [notice, setNotice] = useState("");
  const [errors, setErrors] = useState({});
  const syncState = "演示模式 · 数据仅存本机";
  const [now, setNow] = useState(Date.now());
  const [showReminders, setShowReminders] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [showFireworks, setShowFireworks] = useState(false);
  const fireworksTimerRef = useRef(null);
  const [isAdmin, setIsAdmin] = useState(false);
  const [adminLoginOpen, setAdminLoginOpen] = useState(false);
  const [devices, setDevices] = useState(initial.devices);
  const [showDevicePicker, setShowDevicePicker] = useState(false);
  const [deviceFilter, setDeviceFilter] = useState("");


  const reminders = useMemo(() => records
    .map((record) => ({ ...record, level: getReminderLevel(record, now, records) }))
    .filter((record) => record.level)
    .sort((a, b) => new Date(a.dueAt) - new Date(b.dueAt)), [records, now]);

  const matchingDevices = useMemo(() => {
    const query = deviceFilter.trim().toLowerCase();
    return devices
      .filter((device) => !query || `${device.device_name} ${device.device_no} ${device.group_name}`.toLowerCase().includes(query))
      .sort((a, b) => Number(b.group_name === group) - Number(a.group_name === group));
  }, [devices, group, deviceFilter]);

  const borrowedDeviceKeys = useMemo(() => new Map(records
    .filter((record) => record.type === "领用" && !hasLaterReturn(record, records))
    .map((record) => [`${record.group || ""}::${record.asset}::${record.deviceNo || ""}`, record.person])), [records]);

  const selectedDeviceBorrower = useMemo(() => {
    if (!asset.trim() || !sn.trim()) return "";
    const borrowed = records.find((record) => record.type === "领用"
      && !hasLaterReturn(record, records)
      && record.asset.trim() === asset.trim()
      && record.deviceNo && record.deviceNo.trim() === sn.trim());
    return borrowed?.person || "";
  }, [asset, records, sn]);

  useEffect(() => {
    const timer = window.setInterval(() => setNow(Date.now()), 30_000);
    return () => window.clearInterval(timer);
  }, []);

  useEffect(() => () => {
    if (fireworksTimerRef.current) window.clearTimeout(fireworksTimerRef.current);
  }, []);

  useEffect(() => {
    try { window.localStorage.setItem(STORE_KEY, JSON.stringify({ records, devices, groups })); } catch { /* 忽略存储错误 */ }
  }, [records, devices, groups]);

  const dueAt = useMemo(() => calculateDueAt(new Date(now), duration, unit), [duration, unit, now]);

  const submit = async (event) => {
    event.preventDefault();
    if (submitting) return;
    const nextErrors = {};
    if (!staffId.trim()) nextErrors.staffId = "请输入领用人工号。";
    if (!group) nextErrors.group = "请选择所属组别。";
    if (!asset.trim()) nextErrors.asset = "请选择或输入设备名称。";
    if (!sn.trim()) nextErrors.sn = "请输入设备编号。";
    if (mode === "领用" && selectedDeviceBorrower) nextErrors.asset = `设备已被借走（工号 ${selectedDeviceBorrower}），请选择其他设备。`;
    if (mode === "归还" && !selectedDeviceBorrower) nextErrors.asset = "该设备当前没有待归还的借用记录。";
    if (mode === "归还" && selectedDeviceBorrower && staffId.trim() !== selectedDeviceBorrower) nextErrors.staffId = "请输入正确借用人的工号。";
    setErrors(nextErrors);
    if (Object.keys(nextErrors).length) {
      setNotice(mode === "归还" && selectedDeviceBorrower && staffId.trim() !== selectedDeviceBorrower ? `归还失败：该设备当前由工号 ${selectedDeviceBorrower} 借用。归还工号必须与借用人一致，其他人不能归还；请输入正确借用人的工号。` : selectedDeviceBorrower ? `设备已被借走（工号 ${selectedDeviceBorrower}），请选择其他设备。` : "请修正标记的字段后再提交。");
      return;
    }
    const submittingStartedAt = Date.now();
    setSubmitting(true);
    try {
    const submittedAt = new Date();
    const submittedDueAt = calculateDueAt(submittedAt, duration, unit);
    const record = {
      id: Date.now(), type: mode, person: staffId, asset, deviceNo: sn, when: formatDate(submittedAt), group, note,
      borrowedAt: formatDate(submittedAt), dueAt: submittedDueAt, status: "已提交",
    };
    setRecords((current) => [record, ...current]);
    setNotice(`${mode}登记已提交，已生成操作记录。`);
    // Keep the celebration isolated from the cloud request and reuse one light
    // animation layer on desktop and mobile so a submit never paints a blank page.
    if (fireworksTimerRef.current) window.clearTimeout(fireworksTimerRef.current);
    setShowFireworks(false);
    window.requestAnimationFrame(() => {
      setShowFireworks(true);
      fireworksTimerRef.current = window.setTimeout(() => {
        setShowFireworks(false);
        fireworksTimerRef.current = null;
      }, 1550);
    });
    } finally {
      const remaining = 800 - (Date.now() - submittingStartedAt);
      if (remaining > 0) await new Promise((resolve) => window.setTimeout(resolve, remaining));
      setSubmitting(false);
    }
  };

  const updateField = (field, setter) => (value) => {
    setter(value);
    if (errors[field]) setErrors((current) => ({ ...current, [field]: undefined }));
  };

  const loginAdmin = async (loginStaffId, password) => {
    if (loginStaffId.trim() !== DEMO_ADMIN_ID || password !== DEMO_ADMIN_PASSWORD) return "工号或密码错误（演示账号：admin / demo123）。";
    setIsAdmin(true);
    setAdminLoginOpen(false);
    setView("admin");
    return "";
  };

  const logoutAdmin = async () => {
    setIsAdmin(false);
    setView("form");
    setNotice("管理员已退出登录。");
  };

  const saveDevice = async (draft) => {
    if (!isAdmin) return "请先以管理员身份登录。";
    const payload = { group_name: draft.group_name.trim(), device_name: draft.device_name.trim(), device_no: draft.device_no.trim() };
    if (!payload.group_name || !payload.device_name || !payload.device_no) return "请填写组别、设备名称和设备编号。";
    if (devices.some((device) => device.device_no === payload.device_no && device.id !== draft.id)) return "保存失败：设备编号不能重复。";
    setDevices((current) => draft.id
      ? current.map((device) => device.id === draft.id ? { ...device, ...payload } : device)
      : [...current, { ...payload, id: Date.now(), created_at: Date.now() }]);
    return "";
  };

  const deleteDevice = async (device) => {
    if (!isAdmin) return "请先以管理员身份登录。";
    if (borrowedDeviceKeys.has(`${device.group_name}::${device.device_name}::${device.device_no}`)) return "该设备尚未归还，不能删除。";
    setDevices((current) => current.filter((item) => item.id !== device.id));
    return "";
  };

  const saveGroup = async (draft) => {
    if (!isAdmin) return "请先以管理员身份登录。";
    const name = draft.name.trim();
    if (!name) return "请输入组别名称。";
    if (groups.includes(name) && name !== draft.id) return "保存失败：组别名称不能重复。";
    if (draft.id) {
      setGroups((current) => current.map((item) => item === draft.id ? name : item));
      setDevices((current) => current.map((device) => device.group_name === draft.id ? { ...device, group_name: name } : device));
      if (group === draft.id) setGroup(name);
    } else {
      setGroups((current) => [...current, name]);
    }
    return "";
  };

  const deleteGroup = async (groupRow) => {
    if (!isAdmin) return "请先以管理员身份登录。";
    if (devices.some((device) => device.group_name === groupRow.name)) return "该组别仍有关联设备，请先修改或删除设备。";
    setGroups((current) => current.filter((item) => item !== groupRow.name));
    return "";
  };

  const deleteRecord = async (id) => {
    if (!isAdmin) return "请先以管理员身份登录。";
    setRecords((current) => current.filter((record) => record.id !== id));
    return "";
  };

  const resetDemo = () => {
    if (!window.confirm("恢复成初始示例数据？当前浏览器里的演示数据会被清空。")) return;
    const seed = buildSeed();
    setRecords(seed.records); setDevices(seed.devices); setGroups(seed.groups); setGroup(seed.groups[0]);
    setNotice("已恢复初始示例数据。");
  };

  const nav = [
    ["借用 / 归还", MdReplay, "form"],
    ["领用记录", MdAssignment, "records"], ["归还记录", MdContentPasteSearch, "records"],
  ];

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <img className="brand-logo" src={brandLogo} alt="设备通" />
        <nav className="side-nav" aria-label="主导航">
          {nav.map(([label, Icon, target]) => (
            <button key={label} className={(target === "form" ? view === "form" : view === "records" && label.includes(recordFilter)) ? "nav-item active" : "nav-item"} onClick={() => { setView(target); setRecordFilter(label.includes("\u9886\u7528") ? "\u9886\u7528" : label.includes("\u5f52\u8fd8") ? "\u5f52\u8fd8" : ""); setNotice(""); }}>
              <Icon aria-hidden="true" /><span>{label}</span>
            </button>
          ))}
        </nav>
        <div className="side-bottom">
          <button className="nav-item" onClick={() => isAdmin ? setView("admin") : setAdminLoginOpen(true)}><MdTune /><span>{isAdmin ? "管理员后台" : "管理员登录"}</span></button>
        </div>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div><img className="mobile-brand-logo" src={brandLogo} alt="设备通" /><div><h1>实验室设备借还系统</h1><span>演示版</span></div></div>
          <div className="top-actions">
            <span className="sync-state"><i />{syncState}</span>
            <div className="reminder-wrap">
              <button className="reminder-button" type="button" onClick={() => setShowReminders((open) => !open)} aria-expanded={showReminders}>
                <MdNotificationsNone />超时提醒{reminders.length > 0 && <b>{reminders.length}</b>}
              </button>
              {showReminders && <ReminderPanel reminders={reminders} />}
            </div>
            <button onClick={() => { setRecordFilter(""); setView("records"); }}><MdAssignment />记录</button>
            <button onClick={() => exportExcel(records, now)}><MdDownload />导出记录（Excel）</button>
            <span className="date"><MdCalendarMonth />{formatHeaderDate(new Date(now))}</span>
          </div>
        </header>

        {view === "admin" ? (
          isAdmin ? <AdminPanel records={records} devices={devices} groups={groups} groupRows={groupRows} onSave={saveDevice} onDeleteDevice={deleteDevice} onReset={resetDemo} onSaveGroup={saveGroup} onDeleteGroup={deleteGroup} onDeleteRecord={deleteRecord} onLogout={logoutAdmin} /> : <AdminLogin onLogin={loginAdmin} />
        ) : view === "records" ? (
          <FixedRecordList records={records} recordType={recordFilter} onBack={() => setView("form")} />
        ) : view === "assets" ? (
          <Assets onUse={() => setView("form")} />
        ) : (
          <>
            <div className="content-grid">
              <form className="transaction-form" onSubmit={submit} noValidate>
                <section className="form-section">
                  <h2><span>1</span>借用人信息</h2>
                  <div className="two-col">
                    <Field label="领用人工号" required error={errors.staffId}><input value={staffId} onChange={(e) => updateField("staffId", setStaffId)(e.target.value)} aria-label="领用人工号" aria-invalid={Boolean(errors.staffId)} /></Field>
                    <Field label="领用人所属组别" required error={errors.group}><select value={group} onChange={(e) => updateField("group", setGroup)(e.target.value)} aria-label="领用人所属组别" aria-invalid={Boolean(errors.group)}>{groups.map((item) => <option key={item}>{item}</option>)}</select></Field>
                  </div>
                </section>
                <section className="form-section">
                  <h2><span>2</span>借用时长</h2>
                  <div className="two-col duration-row">
                    <Field label="借用时长单位" required><div className="segmented"><button type="button" className={unit === "天" ? "selected" : ""} onClick={() => setUnit("天")}><MdCalendarMonth />天</button><button type="button" className={unit === "分钟" ? "selected" : ""} onClick={() => setUnit("分钟")}>◷ 分钟</button></div></Field>
                    <Field label="借用时长" required><div className="stepper"><button type="button" onClick={() => setDuration((d) => Math.max(1, d - 1))}><MdRemove /></button><input type="number" min="1" value={duration} aria-label="借用时长数值" onChange={(e) => setDuration(Math.max(1, Number(e.target.value) || 1))} /><button type="button" onClick={() => setDuration((d) => d + 1)}><MdAdd /></button><b>{unit}</b></div><small>预计归还时间：{dueAt}</small></Field>
                  </div>
                </section>
                <section className="form-section">
                  <h2><span>3</span>借用 / 归还选择</h2>
                  <Field label="操作类型" required><div className="wide-toggle"><button type="button" className={mode === "领用" ? "selected" : ""} onClick={() => setMode("领用")}><MdDownload />领用</button><button type="button" className={mode === "归还" ? "selected" : ""} onClick={() => setMode("归还")}><MdArrowOutward />归还</button></div></Field>
                </section>
                <section className="form-section">
                  <h2><span>4</span>设备与设备编号</h2>
                  <div className="two-col">
                    <Field label="选择设备" required error={errors.asset}><div className="search-field device-search"><input value={asset} onChange={(e) => { updateField("asset", setAsset)(e.target.value); setDeviceFilter(e.target.value); setShowDevicePicker(true); }} onFocus={() => setShowDevicePicker(true)} aria-label="选择设备" aria-invalid={Boolean(errors.asset)} /><button type="button" className="search-trigger" aria-label="搜索设备" onClick={() => { setDeviceFilter(""); setShowDevicePicker(true); }}><MdSearch /></button>{showDevicePicker && <div className="device-picker">{matchingDevices.length ? matchingDevices.map((device) => { const borrower = borrowedDeviceKeys.get(`${device.group_name}::${device.device_name}::${device.device_no}`); const borrowed = Boolean(borrower); return <button type="button" className={borrowed ? "device-option borrowed" : "device-option"} key={device.id} onMouseDown={(event) => event.preventDefault()} onClick={() => { if (borrowed && mode === "领用") { setErrors((current) => ({ ...current, asset: `设备已被借走（工号 ${borrower}），请选择其他设备。` })); setNotice("该设备已被借走，请选择其他设备。"); return; } setAsset(device.device_name); setSn(device.device_no); setGroup(device.group_name); setShowDevicePicker(false); setDeviceFilter(""); }}>{device.device_name}<small>{device.group_name} · {device.device_no}{borrowed ? ` · 已借出 · 工号 ${borrower}` : ""}</small></button>; }) : <span>管理员后台暂无匹配设备，可手动输入。</span>}</div>}</div></Field>
                    <Field label="设备编号" required error={errors.sn}><input value={sn} onChange={(e) => updateField("sn", setSn)(e.target.value)} aria-label="设备编号" aria-invalid={Boolean(errors.sn)} /></Field>
                  </div>
                </section>
                <section className="form-section note-section">
                  <h2><span>5</span>备注 <i>（选填）</i></h2>
                  <textarea value={note} onChange={(e) => setNote(e.target.value)} maxLength="200" placeholder="请输入备注信息（如借用用途、注意事项等）" />
                  <small className="count">{note.length}/200</small>
                </section>
              </form>
              <Summary mode={mode} staffId={staffId} group={group} asset={asset} sn={sn} note={note} unit={unit} duration={duration} dueAt={dueAt} borrowedAt={formatDate(new Date(now))} submitting={submitting} onSubmit={() => document.querySelector(".transaction-form").requestSubmit()} />
            </div>
            {notice && <div role="status" aria-live="polite" className="toast"><MdCheckCircle />{notice}</div>}
            {showFireworks && <Fireworks />}
          </>
        )}
        {adminLoginOpen && <AdminLogin onClose={() => setAdminLoginOpen(false)} onLogin={loginAdmin} />}
      </section>
    </main>
  );
}

function Fireworks() {
  return <div className="fireworks" role="status" aria-live="polite" aria-label="提交成功，已生成操作记录"><span className="firework firework-one" /><span className="firework firework-two" /><strong>提交成功</strong></div>;
}

function AdminLogin({ onClose, onLogin }) {
  const [staffId, setStaffId] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const submit = async (event) => {
    event.preventDefault();
    setSaving(true);
    const message = await onLogin(staffId, password);
    setSaving(false);
    setError(message);
  };
  return <div className="admin-overlay" role="dialog" aria-modal="true" aria-label="管理员登录"><form className="admin-login" onSubmit={submit}><MdLock /><h2>管理员登录</h2><p>演示账号：工号 <b>admin</b>，密码 <b>demo123</b></p><label>工号<input value={staffId} onChange={(event) => setStaffId(event.target.value)} autoComplete="username" /></label><label>密码<input type="password" value={password} onChange={(event) => setPassword(event.target.value)} autoComplete="current-password" /></label>{error && <small className="field-error">{error}</small>}<div className="admin-actions"><button type="submit" disabled={saving}>{saving ? "登录中…" : "登录"}</button>{onClose && <button type="button" className="secondary-button" onClick={onClose}>取消</button>}</div></form></div>;
}

function AdminPanel({ records, devices, groups, groupRows, onSave, onDeleteDevice, onReset, onSaveGroup, onDeleteGroup, onDeleteRecord, onLogout }) {
  const emptyDevice = { id: "", group_name: groups[0] || "", device_name: "", device_no: "" };
  const [draft, setDraft] = useState(emptyDevice);
  const [groupDraft, setGroupDraft] = useState({ id: "", name: "" });
  const [message, setMessage] = useState("");
  const [groupMessage, setGroupMessage] = useState("");
  const save = async (event) => { event.preventDefault(); const error = await onSave(draft); if (error) { setMessage(error); return; } setMessage("设备已保存。"); setDraft({ ...emptyDevice, group_name: groups[0] || "" }); };
  const saveGroup = async (event) => { event.preventDefault(); const error = await onSaveGroup(groupDraft); if (error) { setGroupMessage(error); return; } setGroupMessage("组别已保存。"); setGroupDraft({ id: "", name: "" }); };
  const removeGroup = async (row) => { if (!window.confirm(`确定删除组别“${row.name}”吗？`)) return; const error = await onDeleteGroup(row); setGroupMessage(error || "组别已删除。"); };
  const removeDevice = async (device) => { if (!window.confirm(`确定删除设备“${device.device_name}（${device.device_no}）”吗？`)) return; const error = await onDeleteDevice(device); setMessage(error || "设备已删除。"); };
  const removeRecord = async (record) => { if (!window.confirm("确定删除该借还数据吗？此操作无法恢复。")) return; const error = await onDeleteRecord(record.id); setMessage(error || "借还数据已删除。"); };
  return <section className="admin-page"><div className="page-heading"><div><h2>管理员后台</h2><p>查看和删除借还数据；维护组别与设备资料。演示版数据只保存在当前浏览器。</p></div><div className="admin-actions"><button className="secondary-button" onClick={onReset}><MdReplay />恢复示例数据</button><button className="secondary-button" onClick={onLogout}><MdLogout />退出登录</button></div></div><div className="admin-grid"><section className="admin-card"><h3>{draft.id ? "编辑设备" : "新增设备"}</h3><form className="device-form" onSubmit={save}><label>所属组别<select value={draft.group_name} onChange={(event) => setDraft({ ...draft, group_name: event.target.value })}>{groups.map((item) => <option key={item}>{item}</option>)}</select></label><label>设备名称<input value={draft.device_name} onChange={(event) => setDraft({ ...draft, device_name: event.target.value })} /></label><label>设备编号<input value={draft.device_no} onChange={(event) => setDraft({ ...draft, device_no: event.target.value })} /></label><div className="admin-actions"><button type="submit">{draft.id ? "保存修改" : "新增设备"}</button>{draft.id && <button type="button" className="secondary-button" onClick={() => setDraft({ ...emptyDevice, group_name: groups[0] || "" })}>取消编辑</button>}</div></form></section><section className="admin-card"><h3>设备清单</h3><div className="table-wrap"><table><thead><tr><th>组别</th><th>设备名称</th><th>设备编号</th><th></th></tr></thead><tbody>{devices.map((device) => <tr key={device.id}><td>{device.group_name}</td><td>{device.device_name}</td><td>{device.device_no}</td><td><button className="table-button" onClick={() => setDraft(device)}><MdEdit />编辑</button><button className="danger-button" onClick={() => removeDevice(device)}><MdDelete />删除</button></td></tr>)}</tbody></table></div></section></div><section className="admin-card"><h3>组别管理</h3><form className="group-form" onSubmit={saveGroup}><input value={groupDraft.name} onChange={(event) => setGroupDraft({ ...groupDraft, name: event.target.value })} placeholder="输入组别名称" /><button type="submit">{groupDraft.id ? "保存组别" : "新增组别"}</button>{groupDraft.id && <button type="button" className="secondary-button" onClick={() => setGroupDraft({ id: "", name: "" })}>取消</button>}</form>{groupMessage && <small className={groupMessage.includes("已") ? "success" : "field-error"}>{groupMessage}</small>}<div className="group-list">{groupRows.map((row) => <div key={row.id}><span>{row.name}</span><button className="table-button" onClick={() => setGroupDraft(row)}><MdEdit />修改</button><button className="danger-button" onClick={() => removeGroup(row)}><MdDelete />删除</button></div>)}</div></section><section className="admin-card admin-records"><h3>借还后台数据</h3>{message && <small className={message.includes("已") ? "success" : "field-error"}>{message}</small>}<div className="table-wrap"><table><thead><tr><th>操作</th><th>工号</th><th>组别</th><th>设备</th><th>借用时间</th><th>预计归还</th><th></th></tr></thead><tbody>{records.map((record) => <tr key={record.id}><td>{record.type}</td><td>{record.person}</td><td>{record.group || "—"}</td><td>{record.asset}</td><td>{record.borrowedAt || record.when}</td><td>{record.dueAt || "—"}</td><td><button className="danger-button" onClick={() => removeRecord(record)}><MdDelete />删除</button></td></tr>)}</tbody></table></div></section></section>;
}

function Field({ label, required, error, children }) { return <label className={error ? "field has-error" : "field"}><span>{label}{required && <em> *</em>}</span>{children}{error && <small className="field-error" role="alert">{error}</small>}</label>; }

function Summary({ mode, staffId, group, asset, sn, note, unit, duration, dueAt, borrowedAt, submitting, onSubmit }) {
  return <aside className="summary"><h2>借还信息</h2><dl><dt>设备名称</dt><dd>{asset || "—"}</dd><dt>设备编号</dt><dd>{sn || "—"}</dd><dt>领用人工号</dt><dd>{staffId || "—"}</dd><dt>所属组别</dt><dd>{group || "—"}</dd><dt>借用时长</dt><dd>{duration} {unit}</dd><dt>借用时间</dt><dd>{borrowedAt}</dd><dt>预计归还时间</dt><dd>{dueAt}</dd><dt>备注</dt><dd>{note.trim() || "无"}</dd></dl>{mode === "归还" && <p className="return-rule-hint">归还时将核对此设备编号当前借用记录：归还工号必须与借用人一致；其他人不能归还。请输入正确借用人的工号。</p>}<button className={submitting ? "submit-button is-submitting" : "submit-button"} disabled={submitting} onClick={onSubmit}>{submitting ? <><span className="spinner" />正在提交…</> : <><MdArrowOutward />提交登记</>}</button><small>{submitting ? "正在提交，请勿重复提交" : "提交后将生成操作记录"}</small></aside>;
}

function ReminderPanel({ reminders }) {
  return <section className="reminder-panel" role="status" aria-live="polite">
    <header><strong>超时提醒</strong><span>{reminders.length} 条待处理</span></header>
    {reminders.length ? <div className="reminder-list">{reminders.map((item) => (
      <article className={`reminder-card ${item.level}`} key={item.id}>
        <div className="reminder-card-head"><strong>{item.level === "overdue" ? "已超时" : "即将到期"}</strong><span>{item.asset}</span></div>
        <dl>
          <dt>借用人</dt><dd>{item.person}</dd><dt>组别</dt><dd>{item.group || "—"}</dd>
          <dt>借用时间</dt><dd>{item.borrowedAt}</dd><dt>预计归还</dt><dd>{item.dueAt}</dd>
        </dl>
      </article>
    ))}</div> : <div className="reminder-empty"><MdCheckCircle />当前没有即将到期或超时的借用</div>}
    <footer><span className="warning-dot" />黄色：距离归还不足10分钟 <span className="overdue-dot" />红色：已超时</footer>
  </section>;
}

function FixedRecordList({ records, recordType, onBack }) {
  const [query, setQuery] = useState("");
  const normalizedQuery = query.trim().toLowerCase();
  const visible = records.filter((record) => (!recordType || record.type === recordType)
    && `${record.person} ${record.asset} ${record.deviceNo || ""} ${record.group || ""}`.toLowerCase().includes(normalizedQuery));
  const title = recordType ? `${recordType}记录` : "设备操作记录";
  const description = recordType ? `仅展示${recordType}登记。` : "全部领用与归还登记记录";
  return (
    <section className="records-page">
      <div className="page-heading"><div><h2>{title}</h2><p>{description}</p></div><button onClick={onBack}>返回登记</button></div>
      <div className="record-tools"><label>搜索<input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="工号、设备名称或设备编号" /></label><span>{visible.length} 条记录</span></div>
      <div className="table-wrap"><table><thead><tr><th>操作类型</th><th>领用人工号</th><th>设备</th><th>设备编号</th><th>提交时间</th><th>实时状态</th></tr></thead><tbody>
        {visible.length ? visible.map((record) => { const status = getRecordStatus(record, Date.now(), records); return <tr key={record.id} className={`record-row ${status.kind}`}><td><span className={record.type === "领用" ? "tag use" : "tag return"}>{record.type}</span></td><td>{record.person}</td><td>{record.asset}</td><td>{record.deviceNo || "—"}</td><td>{record.when}</td><td><span className={`tag status-${status.kind}`}>{status.label}</span></td></tr>; }) : <tr><td colSpan="6" className="empty-state">没有匹配的{recordType || "操作"}记录。</td></tr>}
      </tbody></table></div>
    </section>
  );
}

function Assets({ onUse }) { return <section className="assets-page"><div className="page-heading"><div><h2>设备管理</h2><p>当前可领用设备</p></div><button onClick={onUse}>新建领用</button></div><article className="asset-card"><img src="/assets/biological-microscope.png" alt="生物显微镜" /><div><span className="available"><i />可用</span><h2>生物显微镜 <b>CX23</b></h2><p>{equipment.assetNo} · {equipment.location}</p><button onClick={onUse}>选择此设备</button></div></article></section>; }
