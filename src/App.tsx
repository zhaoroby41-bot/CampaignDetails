import { ChangeEvent, useMemo, useRef, useState } from "react";
import Button from "devextreme-react/button";
import DataGrid, { Column, FilterRow, HeaderFilter, Paging, SearchPanel, Selection } from "devextreme-react/data-grid";
import Popup from "devextreme-react/popup";
import SelectBox from "devextreme-react/select-box";
import TextBox from "devextreme-react/text-box";
import { CheckCircle2, Download, FileArchive, UploadCloud } from "lucide-react";
import * as XLSX from "xlsx";
import { initialMaterials, participatingDealers } from "./data";
import {
  type ActivityDataRow,
  type ActivityDataSplitResult,
  getDealerActivityData,
  sanitizeExcelFileName,
  splitActivityDataRows,
} from "./lib/activityData";
import { confirmAllDealerMatch, type DealerMatchRow, matchFilesToDealers } from "./lib/dealerMatcher";
import { getDealerNames, getVisibleMaterials, inferFileType } from "./lib/materials";
import type { ActivityMaterial, PendingFile, UserRole } from "./types";

type DataTab = "storeList" | "materials" | "logistics" | "kol" | "activityData" | "ledCollection";
type LedCollectionStatus = "submitted" | "pending";

interface LedStoreRow {
  id: string;
  dealerId: string;
  dealerName: string;
  posAppleId: string;
  posName: string;
  screenAddress: string;
  ledSize: string;
  format: string;
  remarks: string;
  logo: string;
  submittedAt: string;
}

interface LedDraft {
  posAppleId: string;
  posName: string;
  screenAddress: string;
  ledSize: string;
  format: string;
  remarks: string;
  logo: string;
}

const roleOptions = [
  { id: "appleLead", name: "Apple Lead" },
  { id: "dealer", name: "经销商视角" },
] satisfies Array<{ id: UserRole; name: string }>;

const dataTabs: Array<{ id: DataTab; label: string }> = [
  { id: "storeList", label: "活动店单" },
  { id: "materials", label: "上传素材到经销商" },
  { id: "logistics", label: "物流及其他上传" },
  { id: "kol", label: "KOL/KOC探店" },
  { id: "activityData", label: "活动数据" },
  { id: "ledCollection", label: "门店LED信息" },
];

const smartAliases = {
  上海仲璇: "dealer-shanghai",
  深圳酷果: "dealer-shenzhen",
  西藏酷爱: "dealer-tibet",
};

const ALL_DEALERS_ID = "__all_participating_dealers__";
const ALL_DEALERS_OPTION = { id: ALL_DEALERS_ID, name: "全部参与活动经销商" };
const emptyLedDraft: LedDraft = {
  posAppleId: "",
  posName: "",
  screenAddress: "",
  ledSize: "",
  format: "",
  remarks: "",
  logo: "",
};

const dealerNameById = (dealerId: string) => participatingDealers.find((dealer) => dealer.id === dealerId)?.name ?? "";

const initialLedRows: LedStoreRow[] = [
  {
    id: "led-demo-shanghai-1",
    dealerId: "dealer-shanghai",
    dealerName: dealerNameById("dealer-shanghai"),
    posAppleId: "R4845187-1",
    posName: "上海仲璇环球港 Apple 授权店",
    screenAddress: "商场外立面西侧 LED 大屏",
    ledSize: "19209mm * 16720mm",
    format: "横屏 / 双面转角屏",
    remarks: "需适配转角屏，两侧画面比例不同",
    logo: "shanghai-led-logo.png",
    submittedAt: "2026-07-22 10:18:24",
  },
  {
    id: "led-demo-shenzhen-1",
    dealerId: "dealer-shenzhen",
    dealerName: dealerNameById("dealer-shenzhen"),
    posAppleId: "R4845187-2",
    posName: "深圳酷果星创万象天地店",
    screenAddress: "门店入口右侧柱面 LED",
    ledSize: "20650mm * 15428mm",
    format: "竖屏",
    remarks: "建议提供安全区版本，底部 300mm 避免放关键信息",
    logo: "shenzhen-kuguo-logo.png",
    submittedAt: "2026-07-22 14:36:02",
  },
  {
    id: "led-demo-tibet-1",
    dealerId: "dealer-tibet",
    dealerName: dealerNameById("dealer-tibet"),
    posAppleId: "R4845187-3",
    posName: "西藏酷爱拉萨中心店",
    screenAddress: "商场中庭二层环形屏",
    ledSize: "18492mm * 17500mm",
    format: "横屏",
    remarks: "现场亮度偏高，素材需提升对比度",
    logo: "tibet-kuai-logo.png",
    submittedAt: "2026-07-23 09:12:41",
  },
];

function formatBytes(bytes: number) {
  if (!bytes) return "0 KB";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / 1024 ** index).toFixed(index === 0 ? 0 : 1)} ${units[index]}`;
}

function getMatchStatusLabel(status: DealerMatchRow["status"]) {
  const labels: Record<DealerMatchRow["status"], string> = {
    exact: "精确匹配",
    alias: "别名识别",
    fuzzy: "模糊匹配",
    multiple: "多个可能",
    unmatched: "未匹配",
    manual: "手动指定",
    manualAll: "确认全部",
  };
  return labels[status];
}

function FilePreview({ type }: { type: string }) {
  const lowerType = type.toLowerCase();

  if (["jpg", "jpeg", "png", "webp"].includes(lowerType)) {
    return (
      <span className="real-file-preview image-file-preview">
        <span className="image-sky" />
        <span className="image-hero" />
        <span className="image-copy image-copy-one" />
        <span className="image-copy image-copy-two" />
      </span>
    );
  }

  if (lowerType === "pdf") {
    return (
      <span className="real-file-preview document-file-preview pdf-file-preview">
        <span className="document-fold" />
        <span className="pdf-badge">PDF</span>
      </span>
    );
  }

  return (
    <span className="real-file-preview document-file-preview zip-file-preview">
      <span className="document-fold" />
      <span className="document-line document-line-one" />
      <span className="document-line document-line-two" />
      <span className="document-line document-line-three" />
    </span>
  );
}

function exportRowsAsExcel(rows: ActivityDataRow[], fileName: string) {
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "POS List");
  XLSX.writeFile(workbook, `${sanitizeExcelFileName(fileName)}.xlsx`);
}

function exportLedRowsAsExcel(rows: LedStoreRow[]) {
  const header = ["Name of Reseller（客户）", "POS Apple ID", "Name of POS（门店）", "屏幕显示地址", "LED屏幕分辨率(宽*高)", "格式", "备注", "logo", "提交时间"];
  const body = rows.map((row) => [row.dealerName, row.posAppleId, row.posName, row.screenAddress, row.ledSize, row.format, row.remarks, row.logo, row.submittedAt]);
  const worksheet = XLSX.utils.aoa_to_sheet([header, ...body]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "LED Collection");
  XLSX.writeFile(workbook, "门店LED信息汇总.xlsx");
}

export default function App() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const activityDataInputRef = useRef<HTMLInputElement | null>(null);
  const [role, setRole] = useState<UserRole>("appleLead");
  const [currentDealerId, setCurrentDealerId] = useState(participatingDealers[0].id);
  const [materials, setMaterials] = useState<ActivityMaterial[]>(initialMaterials);
  const [isUploadOpen, setIsUploadOpen] = useState(false);
  const [customTag, setCustomTag] = useState("LED");
  const [pendingFiles, setPendingFiles] = useState<PendingFile[]>([]);
  const [dealerMatchRows, setDealerMatchRows] = useState<DealerMatchRow[]>([]);
  const [activeDataTab, setActiveDataTab] = useState<DataTab>("materials");
  const [activityDataResult, setActivityDataResult] = useState<ActivityDataSplitResult | null>(null);
  const [activityDataFileName, setActivityDataFileName] = useState("");
  const [activityDataUpdatedAt, setActivityDataUpdatedAt] = useState("");
  const [activityDataError, setActivityDataError] = useState("");
  const [ledRows, setLedRows] = useState<LedStoreRow[]>(initialLedRows);
  const [ledDraft, setLedDraft] = useState<LedDraft>(emptyLedDraft);

  const dealerSelectOptions = useMemo(() => [ALL_DEALERS_OPTION, ...participatingDealers], []);
  const visibleMaterials = useMemo(() => getVisibleMaterials(materials, role, currentDealerId), [materials, role, currentDealerId]);
  const currentDealer = participatingDealers.find((dealer) => dealer.id === currentDealerId);
  const currentDealerLedRows = useMemo(() => ledRows.filter((row) => row.dealerId === currentDealerId), [currentDealerId, ledRows]);
  const ledDealerStatuses = useMemo(() => participatingDealers.map((dealer) => {
    const dealerRows = ledRows.filter((row) => row.dealerId === dealer.id);
    return {
      dealerId: dealer.id,
      dealerName: dealer.name,
      storeCount: dealer.storeCount,
      submittedCount: dealerRows.length,
      status: dealerRows.length > 0 ? "submitted" as LedCollectionStatus : "pending" as LedCollectionStatus,
      latestSubmittedAt: dealerRows[dealerRows.length - 1]?.submittedAt ?? "-",
    };
  }), [ledRows]);
  const currentDealerActivityData = getDealerActivityData(activityDataResult, currentDealer?.name ?? "");
  const visibleActivityRows = useMemo(() => {
    if (!activityDataResult) return [];
    if (role === "dealer") return currentDealerActivityData?.rows.slice(1) ?? [];
    return activityDataResult.dealerFiles.flatMap((file) => file.rows.slice(1));
  }, [activityDataResult, currentDealerActivityData, role]);
  const visibleActivityExportRows = useMemo(() => {
    if (!activityDataResult) return [];
    return [activityDataResult.header, ...visibleActivityRows];
  }, [activityDataResult, visibleActivityRows]);
  const activityDataColumns = useMemo(() => activityDataResult?.header.map((header, index) => ({
    field: `column_${index}`,
    caption: String(header || `列 ${index + 1}`),
    minWidth: index === 0 ? 220 : 140,
  })) ?? [], [activityDataResult]);
  const activityDataTableRows = useMemo(() => visibleActivityRows.map((row, rowIndex) => {
    const rowObject: Record<string, string | number> = { id: rowIndex + 1 };
    activityDataColumns.forEach((column, columnIndex) => {
      rowObject[column.field] = String(row[columnIndex] ?? "");
    });
    return rowObject;
  }), [activityDataColumns, visibleActivityRows]);
  const uploadReady = pendingFiles.length > 0 && dealerMatchRows.length === pendingFiles.length && dealerMatchRows.every((row) => !row.needsReview);

  const resetUploadForm = () => {
    setCustomTag("LED");
    setPendingFiles([]);
    setDealerMatchRows([]);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const closeUploadDialog = () => {
    setIsUploadOpen(false);
    resetUploadForm();
  };

  const buildPendingFiles = (files: File[]) => files.map((file, index) => ({
    id: `${file.name}-${index}`,
    fileName: file.name,
    fileType: inferFileType(file.name),
    size: formatBytes(file.size),
  }));

  const handleFilesChange = (event: ChangeEvent<HTMLInputElement>) => {
    const nextFiles = buildPendingFiles(Array.from(event.target.files ?? []));
    setPendingFiles(nextFiles);
    setDealerMatchRows(matchFilesToDealers(nextFiles, participatingDealers, smartAliases));
  };

  const updateSmartMatchedDealer = (fileId: string, dealerId: string) => {
    if (!dealerId) return;
    if (dealerId === ALL_DEALERS_ID) {
      confirmAllDealersForFile(fileId);
      return;
    }
    const dealer = participatingDealers.find((item) => item.id === dealerId);
    if (!dealer) return;
    setDealerMatchRows((rows) => rows.map((row) => row.id === fileId ? {
      ...row,
      dealerId: dealer.id,
      dealerName: dealer.name,
      status: "manual",
      confidence: 100,
      needsReview: false,
      candidates: [{ dealerId: dealer.id, dealerName: dealer.name, confidence: 100 }],
    } : row));
  };

  const confirmAllDealersForFile = (fileId: string) => {
    setDealerMatchRows((rows) => rows.map((row) => row.id === fileId ? confirmAllDealerMatch(row) : row));
  };

  const handleActivityDataUpload = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<ActivityDataRow>(worksheet, { header: 1, defval: "" });
      const result = splitActivityDataRows(rows);
      setActivityDataResult(result);
      setActivityDataFileName(file.name);
      setActivityDataUpdatedAt(new Date().toLocaleString("zh-CN", { hour12: false }));
      setActivityDataError("");
    } catch {
      setActivityDataError("Excel 解析失败，请确认上传的是小程序导出的活动数据 .xlsx 文件。");
    }
  };

  const handleUploadConfirm = () => {
    if (!uploadReady) return;
    const uploadedAt = new Date().toLocaleString("zh-CN", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
    });
    const newMaterials = dealerMatchRows.map((row, index) => ({
      id: `smart-upload-${Date.now()}-${index + 1}`,
      fileName: row.fileName,
      fileType: row.fileType,
      size: row.size,
      uploadedAt,
      customTag: customTag || "未分类",
      dealerIds: row.dealerId ? [row.dealerId] : [],
    }));
    setMaterials((current) => [...newMaterials, ...current]);
    closeUploadDialog();
  };

  const updateLedDraft = (field: keyof LedDraft, value: string) => {
    setLedDraft((draft) => ({ ...draft, [field]: value }));
  };

  const ledDraftReady = Boolean(ledDraft.posAppleId.trim() && ledDraft.posName.trim() && ledDraft.screenAddress.trim() && ledDraft.ledSize.trim() && ledDraft.format.trim());

  const submitLedDraft = () => {
    if (!currentDealer || !ledDraftReady) return;
    const submittedAt = new Date().toLocaleString("zh-CN", { hour12: false });
    setLedRows((rows) => [{
      id: `led-${Date.now()}`,
      dealerId: currentDealer.id,
      dealerName: currentDealer.name,
      posAppleId: ledDraft.posAppleId.trim(),
      posName: ledDraft.posName.trim(),
      screenAddress: ledDraft.screenAddress.trim(),
      ledSize: ledDraft.ledSize.trim(),
      format: ledDraft.format.trim(),
      remarks: ledDraft.remarks.trim(),
      logo: ledDraft.logo.trim(),
      submittedAt,
    }, ...rows]);
    setLedDraft(emptyLedDraft);
  };

  return (
    <main className="app-shell">
      <section className="page-heading">
        <div>
          <p className="eyebrow">Channel Marketing</p>
          <h1>测试活动 - xxxx7-21：活动页面</h1>
        </div>
        <div className="view-controls" aria-label="当前视角">
          <SelectBox dataSource={roleOptions} valueExpr="id" displayExpr="name" value={role} onValueChanged={(event) => setRole(event.value)} />
          <SelectBox dataSource={participatingDealers} valueExpr="id" displayExpr="name" value={currentDealerId} disabled={role === "appleLead"} onValueChanged={(event) => setCurrentDealerId(event.value)} />
        </div>
      </section>

      <section className="info-panel compact-info">
        <div className="section-title"><h2>活动基础信息</h2><button type="button" className="collapse-button">展开</button></div>
      </section>
      <section className="info-panel compact-info">
        <div className="section-title"><h2>活动KV图</h2><button type="button" className="collapse-button">展开</button></div>
      </section>

      <section className="material-panel">
        <div className="data-section-heading"><h2>数据信息</h2></div>
        <nav className="data-tabs" aria-label="数据信息 tabs">
          {dataTabs.map((tab) => <button key={tab.id} type="button" className={activeDataTab === tab.id ? "active" : ""} onClick={() => setActiveDataTab(tab.id)}>{tab.label}</button>)}
        </nav>

        {activeDataTab === "materials" ? (
          <>
            <div className="material-toolbar">
              <div>
                <h2>{role === "appleLead" ? "上传素材到经销商" : `${currentDealer?.name ?? ""}：素材下载`}</h2>
                <p>{role === "appleLead" ? "批量上传文件后，系统会根据文件名识别分发对象，未识别文件默认全部分发并需要确认。" : "当前列表已按经销商权限过滤。"}</p>
              </div>
              <div className="toolbar-actions">
                {role === "appleLead" ? <><Button text="上传素材" icon="upload" type="default" stylingMode="outlined" onClick={() => setIsUploadOpen(true)} /><Button text="批量下发" type="default" /></> : null}
                <Button text="批量下载" icon="download" type="default" />
              </div>
            </div>
            <DataGrid className="material-grid" dataSource={visibleMaterials} keyExpr="id" showBorders={false} columnAutoWidth hoverStateEnabled rowAlternationEnabled={false} noDataText="暂无数据">
              <SearchPanel visible width={240} placeholder="搜索素材" />
              <Selection mode="multiple" showCheckBoxesMode="always" />
              <Paging defaultPageSize={10} />
              <Column dataField="id" caption="编号" width={270} />
              <Column dataField="fileName" caption="文件名" minWidth={280} />
              <Column caption="预览" width={100} alignment="center" cellRender={({ data }: { data: ActivityMaterial }) => <span className={`file-preview file-preview-${data.fileType.toLowerCase()}`}><FilePreview type={data.fileType} /></span>} />
              <Column dataField="fileType" caption="文件类型" width={120} />
              {role === "appleLead" ? <Column caption="分发经销商" minWidth={320} cellRender={({ data }: { data: ActivityMaterial }) => <span className="dealer-list">{getDealerNames(data.dealerIds, participatingDealers)}</span>} /> : null}
              <Column dataField="customTag" caption="自定义标签" width={150} />
              <Column dataField="uploadedAt" caption="上传时间" width={180} />
              <Column caption="操作" width={110} alignment="center" cellRender={() => <button type="button" className="icon-command" aria-label="下载"><Download size={18} /></button>} />
            </DataGrid>
          </>
        ) : null}

        {activeDataTab === "activityData" ? (
          <section className="activity-data-workspace">
            <div className="activity-data-hero">
              <div>
                <h2>{role === "appleLead" ? "上传活动数据 Excel" : "活动数据"}</h2>
                <p>{role === "appleLead" ? "导入后直接预览全量活动数据，经销商端会自动按第 1 列 VAD/ T1 过滤。" : "当前表格仅展示本经销商可见的活动数据。"}</p>
              </div>
              <div className="activity-data-actions">
                {role === "appleLead" ? <><Button text="上传活动数据" icon="upload" type="default" onClick={() => activityDataInputRef.current?.click()} /><input ref={activityDataInputRef} type="file" accept=".xlsx,.xls" className="visually-hidden" onChange={handleActivityDataUpload} /></> : null}
                <Button text="导出" icon="download" type="default" stylingMode={role === "appleLead" ? "outlined" : "contained"} disabled={visibleActivityRows.length === 0} onClick={() => exportRowsAsExcel(visibleActivityExportRows, role === "appleLead" ? "活动数据-全量" : `${currentDealer?.name ?? "经销商"}-活动数据`)} />
              </div>
            </div>
            {activityDataError ? <p className="error-text">{activityDataError}</p> : null}
            <div className="activity-data-summary">
              <article><span>源文件</span><strong>{activityDataFileName || "未上传"}</strong></article>
              <article><span>{role === "appleLead" ? "全量数据行数" : "可见数据行数"}</span><strong>{visibleActivityRows.length}</strong></article>
              <article><span>更新时间</span><strong>{activityDataUpdatedAt || "-"}</strong></article>
            </div>
            <DataGrid className="material-grid activity-data-grid" dataSource={activityDataTableRows} keyExpr="id" showBorders={false} columnAutoWidth hoverStateEnabled height={610} scrolling={{ mode: "standard", showScrollbar: "always" }} noDataText={role === "appleLead" ? "上传 Excel 后在这里预览活动数据" : "Apple Lead 上传活动数据后，这里会显示本经销商数据"}>
              <SearchPanel visible width={280} placeholder="搜索活动数据" />
              <FilterRow visible applyFilter="auto" />
              <HeaderFilter visible />
              {activityDataColumns.map((column) => <Column key={column.field} dataField={column.field} caption={column.caption} minWidth={column.minWidth} />)}
            </DataGrid>
          </section>
        ) : null}

        {activeDataTab === "ledCollection" ? (
          <section className="led-workspace">
            <div className="activity-data-hero led-hero">
              <div>
                <h2>{role === "appleLead" ? "门店 LED 信息汇总" : "门店 LED 信息提交"}</h2>
                <p>{role === "appleLead" ? "查看所有参与经销商的提交状态和门店 LED 明细，可导出汇总 Excel。" : "请按 LED 收集模板填写门店屏幕资料，提交后 Apple Lead 可在汇总页查看。"}</p>
              </div>
              <div className="activity-data-actions">
                {role === "appleLead" ? <Button text="导出LED汇总" icon="download" type="default" disabled={ledRows.length === 0} onClick={() => exportLedRowsAsExcel(ledRows)} /> : null}
              </div>
            </div>

            {role === "appleLead" ? (
              <>
                <div className="activity-data-summary">
                  <article><span>参与经销商</span><strong>{participatingDealers.length}</strong></article>
                  <article><span>已提交经销商</span><strong>{ledDealerStatuses.filter((dealer) => dealer.status === "submitted").length}</strong></article>
                  <article><span>LED门店记录</span><strong>{ledRows.length}</strong></article>
                </div>
                <div className="led-grid-stack">
                  <section className="led-card">
                    <div className="led-card-heading"><h3>经销商提交状态</h3><span>{ledDealerStatuses.filter((dealer) => dealer.status === "pending").length} 家待提交</span></div>
                    <DataGrid className="material-grid led-status-grid" dataSource={ledDealerStatuses} keyExpr="dealerId" showBorders={false} columnAutoWidth hoverStateEnabled noDataText="暂无经销商">
                      <SearchPanel visible width={260} placeholder="搜索经销商" />
                      <Paging defaultPageSize={10} />
                      <Column dataField="dealerName" caption="经销商名称" minWidth={320} />
                      <Column dataField="storeCount" caption="参与门店数" width={120} />
                      <Column dataField="submittedCount" caption="已提交记录" width={120} />
                      <Column dataField="latestSubmittedAt" caption="最近提交时间" width={190} />
                      <Column caption="提交状态" width={120} cellRender={({ data }: { data: { status: LedCollectionStatus } }) => <span className={`led-status led-status-${data.status}`}>{data.status === "submitted" ? "已提交" : "未提交"}</span>} />
                    </DataGrid>
                  </section>
                  <section className="led-card">
                    <div className="led-card-heading"><h3>LED 信息汇总</h3><span>{ledRows.length} 条记录</span></div>
                    <DataGrid className="material-grid led-detail-grid" dataSource={ledRows} keyExpr="id" showBorders={false} columnAutoWidth hoverStateEnabled height={460} scrolling={{ mode: "standard", showScrollbar: "always" }} noDataText="经销商提交后，这里会显示汇总明细">
                      <SearchPanel visible width={280} placeholder="搜索LED信息" />
                      <FilterRow visible applyFilter="auto" />
                      <HeaderFilter visible />
                      <Column dataField="dealerName" caption="Name of Reseller（客户）" minWidth={280} />
                      <Column dataField="posAppleId" caption="POS Apple ID" minWidth={150} />
                      <Column dataField="posName" caption="Name of POS（门店）" minWidth={220} />
                      <Column dataField="screenAddress" caption="屏幕显示地址" minWidth={260} />
                      <Column dataField="ledSize" caption="LED屏幕分辨率(宽*高)" minWidth={190} />
                      <Column dataField="format" caption="格式" minWidth={140} />
                      <Column dataField="remarks" caption="备注" minWidth={180} />
                      <Column dataField="logo" caption="logo" minWidth={140} />
                    </DataGrid>
                  </section>
                </div>
              </>
            ) : (
              <div className="led-dealer-layout">
                <section className="led-card led-form-card">
                  <div className="led-card-heading"><h3>填写 LED 收集表</h3><span>{currentDealer?.name}</span></div>
                  <div className="led-form-grid">
                    <label><span>POS Apple ID *</span><TextBox value={ledDraft.posAppleId} placeholder="如 R4845187" onValueChanged={(event) => updateLedDraft("posAppleId", event.value)} /></label>
                    <label><span>Name of POS（门店） *</span><TextBox value={ledDraft.posName} placeholder="请输入门店名称" onValueChanged={(event) => updateLedDraft("posName", event.value)} /></label>
                    <label className="field-wide"><span>屏幕显示地址 *</span><TextBox value={ledDraft.screenAddress} placeholder="请输入 LED 展示位置/屏幕地址" onValueChanged={(event) => updateLedDraft("screenAddress", event.value)} /></label>
                    <label><span>LED屏幕分辨率(宽*高) *</span><TextBox value={ledDraft.ledSize} placeholder="如 1920mm * 1080mm" onValueChanged={(event) => updateLedDraft("ledSize", event.value)} /></label>
                    <label><span>格式 *</span><TextBox value={ledDraft.format} placeholder="如 横屏 / 竖屏 / 异形屏" onValueChanged={(event) => updateLedDraft("format", event.value)} /></label>
                    <label className="field-wide"><span>备注</span><TextBox value={ledDraft.remarks} placeholder="补充说明，如播放限制、亮屏时间等" onValueChanged={(event) => updateLedDraft("remarks", event.value)} /></label>
                    <label className="field-wide"><span>logo</span><TextBox value={ledDraft.logo} placeholder="填写 logo 文件名或链接" onValueChanged={(event) => updateLedDraft("logo", event.value)} /></label>
                  </div>
                  <div className="led-submit-row"><Button text="提交LED信息" type="default" disabled={!ledDraftReady} onClick={submitLedDraft} /></div>
                </section>
                <section className="led-card">
                  <div className="led-card-heading"><h3>我的提交记录</h3><span>{currentDealerLedRows.length} 条</span></div>
                  <DataGrid className="material-grid led-detail-grid" dataSource={currentDealerLedRows} keyExpr="id" showBorders={false} columnAutoWidth hoverStateEnabled height={420} scrolling={{ mode: "standard", showScrollbar: "always" }} noDataText="提交后，这里会显示本经销商 LED 信息">
                    <SearchPanel visible width={240} placeholder="搜索我的LED信息" />
                    <Column dataField="posAppleId" caption="POS Apple ID" minWidth={150} />
                    <Column dataField="posName" caption="Name of POS（门店）" minWidth={220} />
                    <Column dataField="screenAddress" caption="屏幕显示地址" minWidth={260} />
                    <Column dataField="ledSize" caption="LED屏幕分辨率(宽*高)" minWidth={190} />
                    <Column dataField="format" caption="格式" minWidth={140} />
                    <Column dataField="remarks" caption="备注" minWidth={180} />
                    <Column dataField="logo" caption="logo" minWidth={140} />
                    <Column dataField="submittedAt" caption="提交时间" minWidth={180} />
                  </DataGrid>
                </section>
              </div>
            )}
          </section>
        ) : null}

        {activeDataTab !== "materials" && activeDataTab !== "activityData" && activeDataTab !== "ledCollection" ? <div className="empty-tab">该模块沿用原活动页面能力，本原型聚焦素材、活动数据与 LED 信息收集。</div> : null}
      </section>

      <Popup visible={isUploadOpen} onHiding={closeUploadDialog} showTitle={false} width={1120} height="auto" dragEnabled={false} hideOnOutsideClick={false}>
        <div className="upload-dialog">
          <div className="dialog-header"><h2>上传素材</h2><Button icon="close" stylingMode="text" onClick={closeUploadDialog} /></div>
          <div className="dialog-body">
            <div className="smart-upload-intro"><div><h3>根据文件名智能识别经销商</h3><p>系统会从文件名中识别经销商。无法识别时，默认分发给全部参与活动经销商，但需要人工确认后才能提交。</p></div><span>{dealerMatchRows.filter((row) => !row.needsReview).length}/{pendingFiles.length} 已确认</span></div>
            <label className="dialog-field compact"><span>自定义标签</span><TextBox value={customTag} placeholder="如 LED、授权文件、门店台卡" onValueChanged={(event) => setCustomTag(event.value)} /></label>
            <div className="upload-layout"><div><span className="upload-label"><b>*</b> 上传文件</span><div className="upload-empty-panel" aria-hidden="true"><FileArchive size={48} strokeWidth={1.6} /><span>请上传文件</span></div></div><div className="upload-picker"><div className="upload-picker-visual" aria-hidden="true"><UploadCloud size={34} /><span>点击或拖拽文件到此处上传</span><small>文件大小不超过 2GB</small></div><input ref={fileInputRef} className="upload-native-input" type="file" multiple accept=".jpg,.jpeg,.png,.pdf,.zip,.rar" aria-label="点击或拖拽文件到此处上传" onChange={handleFilesChange} /></div></div>
            {pendingFiles.length > 0 ? <div className="pending-strip"><div className="pending-title"><CheckCircle2 size={18} /><span>待上传文件 {pendingFiles.length} 个</span></div><ul>{pendingFiles.map((file) => <li key={file.id}><FilePreview type={file.fileType} /><span>{file.fileName}</span><small>{file.size}</small></li>)}</ul></div> : null}
            {dealerMatchRows.length > 0 ? (
              <div className="smart-match-panel">
                <div className="smart-match-heading"><div><h3>识别结果确认</h3><p>未匹配文件已默认指向“全部参与活动经销商”，请点击“确认全部分发”或手动改选某一家经销商。</p></div><span>{dealerMatchRows.filter((row) => row.needsReview).length === 0 ? "全部文件已确认" : `${dealerMatchRows.filter((row) => row.needsReview).length} 个文件待确认`}</span></div>
                <DataGrid className="smart-match-grid" dataSource={dealerMatchRows} keyExpr="id" showBorders={false} columnAutoWidth hoverStateEnabled noDataText="上传文件后自动识别经销商">
                  <Column dataField="fileName" caption="文件名" minWidth={280} />
                  <Column caption="分发对象" minWidth={280} cellRender={({ data }: { data: DealerMatchRow }) => <div className="match-target-cell"><SelectBox dataSource={dealerSelectOptions} valueExpr="id" displayExpr="name" value={data.dealerId || (data.status === "manualAll" ? ALL_DEALERS_ID : null)} placeholder={data.dealerName || "全部参与活动经销商"} searchEnabled searchExpr="name" searchMode="contains" minSearchLength={0} showDataBeforeSearch searchTimeout={80} onValueChanged={(event) => updateSmartMatchedDealer(data.id, event.value)} />{data.needsReview ? <Button text="确认全部分发" stylingMode="text" onClick={() => confirmAllDealersForFile(data.id)} /> : null}</div>} />
                  <Column caption="匹配状态" width={120} cellRender={({ data }: { data: DealerMatchRow }) => <span className={`match-status match-status-${data.status}`}>{getMatchStatusLabel(data.status)}</span>} />
                </DataGrid>
              </div>
            ) : null}
          </div>
          <div className="dialog-actions"><Button text="取消" stylingMode="outlined" onClick={closeUploadDialog} /><Button text="确认" type="default" disabled={!uploadReady} onClick={handleUploadConfirm} /></div>
        </div>
      </Popup>
    </main>
  );
}
