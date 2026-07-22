import { ChangeEvent, useMemo, useRef, useState } from "react";
import Button from "devextreme-react/button";
import DataGrid, { Column, Paging, SearchPanel, Selection } from "devextreme-react/data-grid";
import Popup from "devextreme-react/popup";
import SelectBox from "devextreme-react/select-box";
import TagBox from "devextreme-react/tag-box";
import TextBox from "devextreme-react/text-box";
import { CheckCircle2, Download, FileArchive, FileSpreadsheet, UploadCloud } from "lucide-react";
import * as XLSX from "xlsx";
import { initialMaterials, participatingDealers } from "./data";
import {
  type ActivityDataRow,
  type ActivityDataSplitResult,
  getDealerActivityData,
  sanitizeExcelFileName,
  splitActivityDataRows,
} from "./lib/activityData";
import { createMaterialRecords, getDealerNames, getVisibleMaterials, inferFileType } from "./lib/materials";
import type { ActivityMaterial, PendingFile, UserRole } from "./types";

const roleOptions = [
  { id: "appleLead", name: "Apple Lead" },
  { id: "dealer", name: "经销商视角" },
] satisfies Array<{ id: UserRole; name: string }>;

type DataTab = "storeList" | "materials" | "activityData" | "logistics" | "kol";

const dataTabs: Array<{ id: DataTab; label: string }> = [
  { id: "storeList", label: "活动店单" },
  { id: "materials", label: "上传素材到经销商" },
  { id: "logistics", label: "物流及其他上传" },
  { id: "kol", label: "KOL/KOC探店" },
  { id: "activityData", label: "活动数据" },
];

function formatBytes(bytes: number) {
  if (!bytes) {
    return "0 KB";
  }

  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / 1024 ** index).toFixed(index === 0 ? 0 : 1)} ${units[index]}`;
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

  if (lowerType === "zip" || lowerType === "rar") {
    return (
      <span className="real-file-preview document-file-preview zip-file-preview">
        <span className="document-fold" />
        <span className="document-line document-line-one" />
        <span className="document-line document-line-two" />
        <span className="document-line document-line-three" />
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
    <span className="real-file-preview document-file-preview">
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

export default function App() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const activityDataInputRef = useRef<HTMLInputElement | null>(null);
  const [role, setRole] = useState<UserRole>("appleLead");
  const [currentDealerId, setCurrentDealerId] = useState(participatingDealers[0].id);
  const [materials, setMaterials] = useState<ActivityMaterial[]>(initialMaterials);
  const [isUploadOpen, setIsUploadOpen] = useState(false);
  const [selectedDealerIds, setSelectedDealerIds] = useState<string[]>([]);
  const [customTag, setCustomTag] = useState("LED");
  const [pendingFiles, setPendingFiles] = useState<PendingFile[]>([]);
  const [activeDataTab, setActiveDataTab] = useState<DataTab>("materials");
  const [activityDataResult, setActivityDataResult] = useState<ActivityDataSplitResult | null>(null);
  const [activityDataFileName, setActivityDataFileName] = useState("");
  const [activityDataUpdatedAt, setActivityDataUpdatedAt] = useState("");
  const [activityDataError, setActivityDataError] = useState("");
  const [selectedActivityDealerName, setSelectedActivityDealerName] = useState("");

  const visibleMaterials = useMemo(
    () => getVisibleMaterials(materials, role, currentDealerId),
    [materials, role, currentDealerId],
  );

  const currentDealer = participatingDealers.find((dealer) => dealer.id === currentDealerId);
  const activeDealerName = selectedActivityDealerName || activityDataResult?.dealerFiles[0]?.dealerName || "";
  const currentDealerActivityData = getDealerActivityData(activityDataResult, activeDealerName);
  const uploadReady = pendingFiles.length > 0;

  const resetUploadForm = () => {
    setSelectedDealerIds([]);
    setCustomTag("LED");
    setPendingFiles([]);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const closeUploadDialog = () => {
    setIsUploadOpen(false);
    resetUploadForm();
  };

  const handleFilesChange = (event: ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(event.target.files ?? []);
    setPendingFiles(files.map((file, index) => ({
      id: `${file.name}-${index}`,
      fileName: file.name,
      fileType: inferFileType(file.name),
      size: formatBytes(file.size),
    })));
  };

  const handleActivityDataUpload = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];

    if (!file) {
      return;
    }

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
      setSelectedActivityDealerName(result.dealerFiles[0]?.dealerName ?? "");
    } catch {
      setActivityDataError("Excel 解析失败，请确认上传的是小程序导出的活动数据 .xlsx 文件。");
    }
  };

  const handleUploadConfirm = () => {
    if (!uploadReady) {
      return;
    }

    const newMaterials = createMaterialRecords(pendingFiles, selectedDealerIds, customTag || "未分类");
    setMaterials((current) => [...newMaterials, ...current]);
    closeUploadDialog();
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
          <SelectBox
            dataSource={participatingDealers}
            valueExpr="id"
            displayExpr="name"
            value={currentDealerId}
            disabled={role === "appleLead"}
            onValueChanged={(event) => setCurrentDealerId(event.value)}
          />
        </div>
      </section>

      <section className="info-panel compact-info">
        <div className="section-title">
          <h2>活动基础信息</h2>
          <button type="button" className="collapse-button">展开</button>
        </div>
      </section>

      <section className="info-panel compact-info">
        <div className="section-title">
          <h2>活动KV图</h2>
          <button type="button" className="collapse-button">展开</button>
        </div>
      </section>

      <section className="material-panel">
        <div className="data-section-heading">
          <h2>数据信息</h2>
        </div>
        <nav className="data-tabs" aria-label="数据信息 tabs">
          {dataTabs.map((tab) => (
            <button key={tab.id} type="button" className={activeDataTab === tab.id ? "active" : ""} onClick={() => setActiveDataTab(tab.id)}>
              {tab.label}
            </button>
          ))}
        </nav>

        {activeDataTab === "materials" ? (
          <>
            <div className="material-toolbar">
              <div>
                <h2>{role === "appleLead" ? "上传素材到经销商" : `${currentDealer?.name ?? ""}：素材下载`}</h2>
                <p>{role === "appleLead" ? "上传时可指定一个或多个参与活动经销商；不指定则默认分发给全部参与活动经销商。" : "当前列表已按经销商权限过滤。"}</p>
              </div>
              <div className="toolbar-actions">
                {role === "appleLead" ? (
                  <>
                    <Button text="上传素材" icon="upload" type="default" stylingMode="outlined" onClick={() => setIsUploadOpen(true)} />
                    <Button text="批量下发" type="default" />
                  </>
                ) : null}
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
          role === "appleLead" ? (
            <section className="activity-data-workspace">
              <div className="activity-data-hero">
                <div>
                  <h2>上传活动数据 Excel</h2>
                  <p>系统按第 1 列 VAD/ T1 自动拆分，经销商只能下载自己的活动数据。</p>
                </div>
                <Button text="上传活动数据" icon="upload" type="default" onClick={() => activityDataInputRef.current?.click()} />
                <input ref={activityDataInputRef} type="file" accept=".xlsx,.xls" className="visually-hidden" onChange={handleActivityDataUpload} />
              </div>

              {activityDataError ? <p className="error-text">{activityDataError}</p> : null}

              <div className="activity-data-summary">
                <article>
                  <span>源文件</span>
                  <strong>{activityDataFileName || "未上传"}</strong>
                </article>
                <article>
                  <span>数据行数</span>
                  <strong>{activityDataResult?.totalRows ?? 0}</strong>
                </article>
                <article>
                  <span>拆分经销商</span>
                  <strong>{activityDataResult?.dealerFiles.length ?? 0}</strong>
                </article>
              </div>

              <DataGrid
                className="material-grid"
                dataSource={activityDataResult?.dealerFiles ?? []}
                keyExpr="dealerName"
                showBorders={false}
                columnAutoWidth
                hoverStateEnabled
                noDataText="上传 Excel 后自动生成经销商拆分文件"
              >
                <SearchPanel visible width={260} placeholder="搜索经销商" />
                <Paging defaultPageSize={10} />
                <Column dataField="dealerName" caption="经销商名称" minWidth={360} />
                <Column dataField="rowCount" caption="数据行数" width={120} />
                <Column caption="上传时间" width={190} cellRender={() => <span>{activityDataUpdatedAt || "-"}</span>} />
                <Column
                  caption="操作"
                  width={140}
                  alignment="center"
                  cellRender={({ data }: { data: { dealerName: string; rows: ActivityDataRow[] } }) => (
                    <Button text="下载" icon="download" stylingMode="text" onClick={() => exportRowsAsExcel(data.rows, `${data.dealerName}-活动数据`)} />
                  )}
                />
              </DataGrid>
            </section>
          ) : (
            <section className="dealer-task-board">
              <div className="task-lane">
                <h3>未开始 (1)</h3>
                <article className="task-card">
                  <FileSpreadsheet size={26} />
                  <div>
                    <strong>测试活动 - xxxx7-21 - 活动数据</strong>
                    <span>自动创建</span>
                  </div>
                </article>
              </div>

              <div className="dealer-download-panel">
                <h2>活动数据下载</h2>
                <p>经销商只能下载 VAD/ T1 为自己经销商名称的数据文件。</p>
                {activityDataResult ? (
                  <>
                    <SelectBox
                      dataSource={activityDataResult.dealerFiles}
                      valueExpr="dealerName"
                      displayExpr="dealerName"
                      value={activeDealerName}
                      onValueChanged={(event) => setSelectedActivityDealerName(event.value)}
                    />
                    <div className="dealer-download-card">
                      <FileSpreadsheet size={38} />
                      <div>
                        <strong>{activeDealerName || "暂无经销商数据"}</strong>
                        <span>{currentDealerActivityData ? `${currentDealerActivityData.rowCount} 行活动数据` : "未匹配到当前经销商数据"}</span>
                      </div>
                      <Button
                        text="下载Excel"
                        icon="download"
                        type="default"
                        disabled={!currentDealerActivityData}
                        onClick={() => currentDealerActivityData ? exportRowsAsExcel(currentDealerActivityData.rows, `${currentDealerActivityData.dealerName}-活动数据`) : undefined}
                      />
                    </div>
                  </>
                ) : (
                  <div className="empty-task">Apple Lead 上传活动数据后，这里会出现可下载 Excel。</div>
                )}
              </div>
            </section>
          )
        ) : null}

        {activeDataTab !== "materials" && activeDataTab !== "activityData" ? <div className="empty-tab">该模块沿用原活动页面能力，本原型聚焦素材与活动数据分发。</div> : null}
      </section>

      <Popup visible={isUploadOpen} onHiding={closeUploadDialog} showTitle={false} width={980} height="auto" dragEnabled={false} hideOnOutsideClick={false}>
        <div className="upload-dialog">
          <div className="dialog-header">
            <h2>上传素材</h2>
            <Button icon="close" stylingMode="text" onClick={closeUploadDialog} />
          </div>

          <div className="dialog-body">
            <div className="dialog-field">
              <span>分发经销商<small>（请选择素材需要分发的经销商，如不选择则默认该素材分发给全部选择参加活动经销商）</small></span>
              <TagBox
                dataSource={participatingDealers}
                value={selectedDealerIds}
                valueExpr="id"
                displayExpr="name"
                searchEnabled
                showSelectionControls
                applyValueMode="instantly"
                placeholder="默认分发给全部参与活动经销商"
                selectAllMode="allPages"
                showClearButton
                maxDisplayedTags={3}
                dropDownOptions={{ container: ".upload-dialog", hideOnOutsideClick: true, hideOnParentScroll: false, shading: false }}
                onValueChanged={(event) => setSelectedDealerIds(event.value ?? [])}
              />
            </div>

            <div className="dealer-summary"><span>当前分发范围：{selectedDealerIds.length > 0 ? `已选择 ${selectedDealerIds.length} 家经销商` : "全部参与活动经销商"}</span></div>

            <label className="dialog-field compact">
              <span>自定义标签</span>
              <TextBox value={customTag} placeholder="如 LED、授权文件、门店台卡" onValueChanged={(event) => setCustomTag(event.value)} />
            </label>

            <div className="upload-layout">
              <div>
                <span className="upload-label"><b>*</b> 上传文件</span>
                <div className="upload-empty-panel" aria-hidden="true"><FileArchive size={48} strokeWidth={1.6} /><span>请上传文件</span></div>
              </div>

              <div className="upload-picker">
                <div className="upload-picker-visual" aria-hidden="true"><UploadCloud size={34} /><span>点击或拖拽文件到此处上传</span><small>文件大小不超过 2GB</small></div>
                <input ref={fileInputRef} className="upload-native-input" type="file" multiple accept=".jpg,.jpeg,.png,.pdf,.zip,.rar" aria-label="点击或拖拽文件到此处上传" onChange={handleFilesChange} />
              </div>
            </div>

            {pendingFiles.length > 0 ? (
              <div className="pending-strip">
                <div className="pending-title"><CheckCircle2 size={18} /><span>待上传文件 {pendingFiles.length} 个</span></div>
                <ul>{pendingFiles.map((file) => <li key={file.id}><FilePreview type={file.fileType} /><span>{file.fileName}</span><small>{file.size}</small></li>)}</ul>
              </div>
            ) : null}
          </div>

          <div className="dialog-actions">
            <Button text="取消" stylingMode="outlined" onClick={closeUploadDialog} />
            <Button text="确认" type="default" disabled={!uploadReady} onClick={handleUploadConfirm} />
          </div>
        </div>
      </Popup>
    </main>
  );
}
