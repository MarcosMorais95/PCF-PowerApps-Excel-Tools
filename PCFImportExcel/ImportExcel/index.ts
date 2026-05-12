import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as XLSX from "xlsx";

type JsonRecord = Record<string, unknown>;
type TraceLevel = "INFO" | "WARN" | "ERROR";
type DisplayMode = "Card" | "Button";

interface TraceEntry {
  ts: string;
  level: TraceLevel;
  step: string;
  message: string;
  fileName?: string;
  details?: Record<string, unknown>;
}

interface FileMeta {
  sheetUsed: string | null;
  sheetIndex: number;
  usedRange: string | null;
  mode: "table" | "range";
  tableName?: string | null;
  headerRowIndex?: number | null;
  detectedColumns?: number | null;
  detectedRows?: number | null;
}

interface FileResult {
  fileName: string;
  data: JsonRecord[];
  meta: FileMeta;
}

interface GlobalMeta {
  filesSelected: number;
  filesParsed: number;
  allowMultipleFiles: boolean;
  hasTable: boolean;
  sheetIndex: number;
  tableName: string | null;
  maxRowsToScan: number;
  enableTrace: boolean;
}

class PcfError extends Error {
  public code: number;
  constructor(code: number, message: string) {
    super(`${code}: ${message}`);
    this.code = code;
  }
}

const ACCEPTED_EXTENSIONS = [".xlsx", ".xlsm", ".xlsb", ".xls", ".csv"];
const ACCEPT_ATTR =
  ".xlsx,.xlsm,.xlsb,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv";

export class ImportExcel implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private context!: ComponentFramework.Context<IInputs>;
  private hostContainer!: HTMLDivElement;
  private notifyOutputChanged!: () => void;

  // Shared
  private fileInput!: HTMLInputElement;
  private currentMode: DisplayMode = "Card";

  // Card-mode elements
  private cardRoot?: HTMLDivElement;
  private cardHeaderRow?: HTMLDivElement;
  private cardTitleEl?: HTMLDivElement;
  private cardGuidelinesEl?: HTMLDivElement;
  private dropzoneEl?: HTMLDivElement;
  private dropzonePrimaryEl?: HTMLDivElement;
  private dropzoneHintEl?: HTMLDivElement;
  private fileChipEl?: HTMLDivElement;
  private fileChipNameEl?: HTMLDivElement;
  private fileChipSubEl?: HTMLDivElement;
  private fileChipRemoveBtn?: HTMLButtonElement;
  private statusEl?: HTMLDivElement;

  // Button-mode elements
  private buttonRoot?: HTMLButtonElement;
  private buttonIconEl?: HTMLSpanElement;
  private buttonTextEl?: HTMLSpanElement;
  private buttonResetTimer: number | null = null;

  // Outputs / state
  private _jsonResult = "";
  private _meta = "";
  private _trace = "";
  private _isValid = false;
  private _errorMessage = "";

  private selectedFiles: File[] = [];
  private traceEntries: TraceEntry[] = [];
  private isParsing = false;
  private lastResultSummary: { rows: number; cols: number | null } | null = null;

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.context = context;
    this.notifyOutputChanged = notifyOutputChanged;
    this.hostContainer = container;

    this.injectStylesOnce();

    this.fileInput = document.createElement("input");
    this.fileInput.type = "file";
    this.fileInput.accept = ACCEPT_ATTR;
    this.fileInput.style.display = "none";
    this.fileInput.addEventListener("change", () => this.onFilesPicked());

    this.currentMode = this.getDisplayMode();
    this.renderForMode(this.currentMode);
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this.context = context;

    const allowMultiple = this.getAllowMultipleFiles();
    this.fileInput.multiple = allowMultiple;

    if (!allowMultiple && this.selectedFiles.length > 1) {
      this.selectedFiles = [this.selectedFiles[0]];
      this.fileInput.value = "";
    }

    const newMode = this.getDisplayMode();
    if (newMode !== this.currentMode) {
      this.currentMode = newMode;
      this.renderForMode(newMode);
      return;
    }

    this.applyTheme();
    this.applyLabels();
  }

  public getOutputs(): IOutputs {
    return {
      jsonResult: this._jsonResult,
      meta: this._meta,
      trace: this._trace,
      isValid: this._isValid,
      errorMessage: this._errorMessage
    };
  }

  public destroy(): void {
    if (this.buttonResetTimer !== null) {
      window.clearTimeout(this.buttonResetTimer);
      this.buttonResetTimer = null;
    }
  }

  // --------------------------
  // Rendering
  // --------------------------

  private injectStylesOnce(): void {
    if (document.getElementById("pcf-import-excel-styles")) return;
    const style = document.createElement("style");
    style.id = "pcf-import-excel-styles";
    style.textContent = `
@keyframes pcfImportExcelSpin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
.pcf-iex-spinner-arc {
  transform-origin: 8px 8px;
  animation: pcfImportExcelSpin 0.9s linear infinite;
}
.pcf-iex-dropzone {
  border: 1.5px dashed #c8c6c4;
  border-radius: 4px;
  padding: 22px 16px;
  text-align: center;
  cursor: pointer;
  background: #fafafa;
  transition: border-color .15s ease, background .15s ease;
  outline: none;
}
.pcf-iex-dropzone:hover, .pcf-iex-dropzone--drag {
  border-color: var(--pcf-iex-accent, #605e5c);
  background: #f3f2f1;
}
.pcf-iex-dropzone:focus-visible {
  border-color: var(--pcf-iex-accent, #323130);
  box-shadow: 0 0 0 2px rgba(50, 49, 48, 0.15);
}
.pcf-iex-button {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 7px 14px;
  border-radius: 4px;
  cursor: pointer;
  font-size: 13px;
  font-weight: 500;
  font-family: inherit;
  line-height: 1;
  background: var(--pcf-iex-bg, #ffffff);
  color: #201f1e;
  border: 1px solid var(--pcf-iex-border, #c8c6c4);
  transition: background .12s ease, border-color .12s ease;
}
.pcf-iex-button:hover:not(:disabled) {
  background: #f3f2f1;
  border-color: var(--pcf-iex-accent, #605e5c);
}
.pcf-iex-button:disabled { cursor: default; opacity: 0.75; }
.pcf-iex-button--success {
  background: #ffffff;
  color: #107C10;
  border-color: #d1ead1;
}
`;
    document.head.appendChild(style);
  }

  private renderForMode(mode: DisplayMode): void {
    // Reset stored element refs
    this.cardRoot = undefined;
    this.cardHeaderRow = undefined;
    this.cardTitleEl = undefined;
    this.cardGuidelinesEl = undefined;
    this.dropzoneEl = undefined;
    this.dropzonePrimaryEl = undefined;
    this.dropzoneHintEl = undefined;
    this.fileChipEl = undefined;
    this.fileChipNameEl = undefined;
    this.fileChipSubEl = undefined;
    this.fileChipRemoveBtn = undefined;
    this.statusEl = undefined;
    this.buttonRoot = undefined;
    this.buttonIconEl = undefined;
    this.buttonTextEl = undefined;

    // Clear container
    while (this.hostContainer.firstChild) {
      this.hostContainer.removeChild(this.hostContainer.firstChild);
    }

    this.hostContainer.appendChild(this.fileInput);

    if (mode === "Button") {
      this.buildButtonUI();
    } else {
      this.buildCardUI();
    }

    this.applyTheme();
    this.applyLabels();
    this.reflectSelectionState();
  }

  private buildCardUI(): void {
    const card = document.createElement("div");
    card.style.fontFamily = "Segoe UI, -apple-system, BlinkMacSystemFont, Arial, sans-serif";
    card.style.boxSizing = "border-box";
    card.style.padding = "16px";
    card.style.borderRadius = "8px";
    card.style.boxShadow = "0 1.6px 3.6px rgba(0,0,0,0.06), 0 0.3px 0.9px rgba(0,0,0,0.04)";
    card.style.fontSize = "14px";
    card.style.color = "#201f1e";

    // Header (icon + title)
    const headerRow = document.createElement("div");
    headerRow.style.display = "flex";
    headerRow.style.alignItems = "center";
    headerRow.style.gap = "10px";
    headerRow.style.marginBottom = "4px";

    const iconWrap = document.createElement("span");
    iconWrap.style.width = "26px";
    iconWrap.style.height = "26px";
    iconWrap.style.borderRadius = "5px";
    iconWrap.style.background = "#f3f2f1";
    iconWrap.style.display = "inline-flex";
    iconWrap.style.alignItems = "center";
    iconWrap.style.justifyContent = "center";
    iconWrap.style.flex = "none";
    iconWrap.innerHTML = this.excelIconSvg(16);

    const titleEl = document.createElement("div");
    titleEl.style.fontSize = "14px";
    titleEl.style.fontWeight = "600";

    headerRow.appendChild(iconWrap);
    headerRow.appendChild(titleEl);

    // Guidelines
    const guidelinesEl = document.createElement("div");
    guidelinesEl.style.fontSize = "12px";
    guidelinesEl.style.color = "#605e5c";
    guidelinesEl.style.margin = "0 0 12px";

    // Dropzone
    const dz = document.createElement("div");
    dz.className = "pcf-iex-dropzone";
    dz.setAttribute("role", "button");
    dz.setAttribute("tabindex", "0");

    const dzIcon = document.createElement("div");
    dzIcon.style.margin = "0 auto 8px";
    dzIcon.style.width = "28px";
    dzIcon.style.height = "28px";
    dzIcon.style.display = "flex";
    dzIcon.style.alignItems = "center";
    dzIcon.style.justifyContent = "center";
    dzIcon.style.color = "#605e5c";
    dzIcon.innerHTML = `<svg width="26" height="26" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M12 16V4M12 4l-4 4M12 4l4 4" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
      <path d="M4 14v4a2 2 0 002 2h12a2 2 0 002-2v-4" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
    </svg>`;

    const dzPrimary = document.createElement("div");
    dzPrimary.style.fontSize = "13px";
    dzPrimary.style.fontWeight = "500";
    dzPrimary.style.color = "#323130";

    const dzHint = document.createElement("div");
    dzHint.style.fontSize = "11px";
    dzHint.style.color = "#8a8886";
    dzHint.style.marginTop = "4px";

    dz.appendChild(dzIcon);
    dz.appendChild(dzPrimary);
    dz.appendChild(dzHint);

    dz.addEventListener("click", () => this.fileInput.click());
    dz.addEventListener("keydown", (e: KeyboardEvent) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        this.fileInput.click();
      }
    });

    // Drag-and-drop
    const onDragOver = (e: DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
      dz.classList.add("pcf-iex-dropzone--drag");
    };
    const onDragLeave = (e: DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
      dz.classList.remove("pcf-iex-dropzone--drag");
    };
    const onDrop = (e: DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
      dz.classList.remove("pcf-iex-dropzone--drag");
      const dropped = e.dataTransfer?.files;
      if (dropped && dropped.length > 0) {
        this.applyPickedFiles(Array.from(dropped));
      }
    };
    dz.addEventListener("dragover", onDragOver);
    dz.addEventListener("dragenter", onDragOver);
    dz.addEventListener("dragleave", onDragLeave);
    dz.addEventListener("drop", onDrop);

    // File chip (hidden by default)
    const chip = document.createElement("div");
    chip.style.display = "none";
    chip.style.alignItems = "center";
    chip.style.gap = "10px";
    chip.style.padding = "10px 12px";
    chip.style.borderRadius = "4px";

    const chipIcon = document.createElement("span");
    chipIcon.style.width = "22px";
    chipIcon.style.height = "22px";
    chipIcon.style.flex = "none";
    chipIcon.style.display = "inline-flex";
    chipIcon.style.alignItems = "center";
    chipIcon.style.justifyContent = "center";
    chipIcon.innerHTML = this.excelIconSvg(20);

    const chipMeta = document.createElement("div");
    chipMeta.style.flex = "1";
    chipMeta.style.minWidth = "0";

    const chipName = document.createElement("div");
    chipName.style.fontSize = "13px";
    chipName.style.fontWeight = "500";
    chipName.style.color = "#201f1e";
    chipName.style.overflow = "hidden";
    chipName.style.textOverflow = "ellipsis";
    chipName.style.whiteSpace = "nowrap";

    const chipSub = document.createElement("div");
    chipSub.style.fontSize = "11px";
    chipSub.style.color = "#605e5c";
    chipSub.style.marginTop = "2px";

    chipMeta.appendChild(chipName);
    chipMeta.appendChild(chipSub);

    const chipRemove = document.createElement("button");
    chipRemove.type = "button";
    chipRemove.style.background = "transparent";
    chipRemove.style.border = "none";
    chipRemove.style.color = "#605e5c";
    chipRemove.style.cursor = "pointer";
    chipRemove.style.padding = "4px 8px";
    chipRemove.style.borderRadius = "4px";
    chipRemove.style.fontSize = "12px";
    chipRemove.style.fontFamily = "inherit";
    chipRemove.addEventListener("click", (e) => {
      e.stopPropagation();
      this.clearSelection();
    });

    chip.appendChild(chipIcon);
    chip.appendChild(chipMeta);
    chip.appendChild(chipRemove);

    // Status
    const status = document.createElement("div");
    status.style.marginTop = "10px";
    status.style.fontSize = "12px";
    status.style.padding = "8px 10px";
    status.style.borderRadius = "4px";
    status.style.background = "#f3f2f1";
    status.style.color = "#605e5c";
    status.style.borderLeft = "3px solid #c8c6c4";

    card.appendChild(headerRow);
    card.appendChild(guidelinesEl);
    card.appendChild(dz);
    card.appendChild(chip);
    card.appendChild(status);

    this.hostContainer.appendChild(card);

    this.cardRoot = card;
    this.cardHeaderRow = headerRow;
    this.cardTitleEl = titleEl;
    this.cardGuidelinesEl = guidelinesEl;
    this.dropzoneEl = dz;
    this.dropzonePrimaryEl = dzPrimary;
    this.dropzoneHintEl = dzHint;
    this.fileChipEl = chip;
    this.fileChipNameEl = chipName;
    this.fileChipSubEl = chipSub;
    this.fileChipRemoveBtn = chipRemove;
    this.statusEl = status;
  }

  private buildButtonUI(): void {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "pcf-iex-button";

    const iconSpan = document.createElement("span");
    iconSpan.style.width = "14px";
    iconSpan.style.height = "14px";
    iconSpan.style.display = "inline-flex";
    iconSpan.style.alignItems = "center";
    iconSpan.style.justifyContent = "center";
    iconSpan.innerHTML = this.excelIconSvg(14);

    const textSpan = document.createElement("span");

    btn.appendChild(iconSpan);
    btn.appendChild(textSpan);

    btn.addEventListener("click", () => {
      if (this.isParsing) return;
      this.fileInput.click();
    });

    this.hostContainer.appendChild(btn);

    this.buttonRoot = btn;
    this.buttonIconEl = iconSpan;
    this.buttonTextEl = textSpan;
  }

  private applyTheme(): void {
    const accent = this.getAccentColor();
    const bg = this.getBackgroundColor();
    const border = this.getBorderColor();

    if (this.cardRoot) {
      this.cardRoot.style.setProperty("--pcf-iex-accent", accent);
      this.cardRoot.style.background = bg;
      this.cardRoot.style.border = `1px solid ${border}`;
    }
    if (this.buttonRoot) {
      this.buttonRoot.style.setProperty("--pcf-iex-accent", accent);
      this.buttonRoot.style.setProperty("--pcf-iex-bg", bg);
      this.buttonRoot.style.setProperty("--pcf-iex-border", border);
    }
  }

  private applyLabels(): void {
    if (this.cardTitleEl) {
      this.cardTitleEl.innerText = this.getTitleText();
      this.cardHeaderRow!.style.display = this.getShowTitle() ? "flex" : "none";
    }
    if (this.cardGuidelinesEl) {
      this.cardGuidelinesEl.innerText = this.getUserGuidelinesText();
      this.cardGuidelinesEl.style.display =
        this.getShowUserGuidelines() && this.getUserGuidelinesText().trim() !== "" ? "block" : "none";
    }
    if (this.dropzonePrimaryEl) {
      this.dropzonePrimaryEl.innerHTML = this.escapeHtmlWithAccent(this.getDropzonePrimaryText());
    }
    if (this.dropzoneHintEl) {
      this.dropzoneHintEl.innerText = this.getDropzoneHintText();
    }
    if (this.fileChipRemoveBtn) {
      this.fileChipRemoveBtn.innerText = this.getRemoveButtonText();
    }
    if (this.statusEl) {
      this.statusEl.style.display = this.getShowStatus() ? "block" : "none";
    }
    if (this.buttonTextEl && !this.isParsing && !this.lastResultSummary) {
      this.buttonTextEl.innerText = this.getButtonText();
    }
  }

  private escapeHtmlWithAccent(text: string): string {
    // Plain text — no HTML interpretation, fully escaped
    const div = document.createElement("div");
    div.innerText = text;
    return div.innerHTML;
  }

  private reflectSelectionState(): void {
    if (this.currentMode === "Card") {
      if (this.selectedFiles.length === 0) {
        if (this.dropzoneEl) this.dropzoneEl.style.display = "block";
        if (this.fileChipEl) this.fileChipEl.style.display = "none";
        this.setCardStatus("idle", "Aguardando seleção do arquivo…");
      } else {
        if (this.dropzoneEl) this.dropzoneEl.style.display = "none";
        if (this.fileChipEl) this.fileChipEl.style.display = "flex";
        this.updateChipForCurrentState();
      }
    }
  }

  // --------------------------
  // File pick + parse orchestration
  // --------------------------

  private onFilesPicked(): void {
    const files = this.fileInput.files ? Array.from(this.fileInput.files) : [];
    this.applyPickedFiles(files);
  }

  private applyPickedFiles(files: File[]): void {
    if (files.length === 0) {
      this.selectedFiles = [];
      this.fileInput.value = "";
      this.reflectSelectionState();
      return;
    }

    const allowMultiple = this.getAllowMultipleFiles();
    let chosen = files;
    if (!allowMultiple) chosen = [files[0]];

    const invalid = chosen.find(f => !this.isAcceptedFile(f));
    if (invalid) {
      this.selectedFiles = chosen;
      this.fileInput.value = "";
      this.handleInvalidFile(invalid);
      return;
    }

    this.selectedFiles = chosen;
    this.fileInput.value = "";
    void this.tryParseSelectedFiles();
  }

  private isAcceptedFile(file: File): boolean {
    const name = file.name.toLowerCase();
    return ACCEPTED_EXTENSIONS.some(ext => name.endsWith(ext));
  }

  private handleInvalidFile(file: File): void {
    this.resetOutputsAndTrace();
    this._isValid = false;
    this._errorMessage = "File format not supported. Use .xlsx, .xls, .xlsm, .xlsb or .csv.";

    if (this.currentMode === "Card") {
      if (this.dropzoneEl) this.dropzoneEl.style.display = "none";
      if (this.fileChipEl) this.fileChipEl.style.display = "flex";
      this.renderChipError(file.name, "Formato não suportado");
      this.setCardStatus("error", `Erro: ${this._errorMessage}`);
    } else {
      this.flashButtonError();
    }

    this.notifyOutputChanged();
  }

  private clearSelection(): void {
    if (this.buttonResetTimer !== null) {
      window.clearTimeout(this.buttonResetTimer);
      this.buttonResetTimer = null;
    }
    this.resetOutputsAndTrace();
    this.selectedFiles = [];
    this.fileInput.value = "";
    this.lastResultSummary = null;
    this.isParsing = false;
    this.reflectSelectionState();
    if (this.currentMode === "Button") {
      this.setButtonIdle();
    }
    this.notifyOutputChanged();
  }

  private async tryParseSelectedFiles(): Promise<void> {
    this.resetOutputsAndTrace();
    this.isParsing = true;

    // UI: show parsing state
    if (this.currentMode === "Card") {
      if (this.dropzoneEl) this.dropzoneEl.style.display = "none";
      if (this.fileChipEl) this.fileChipEl.style.display = "flex";
      this.renderChipParsing(this.selectedFiles[0].name);
      this.setCardStatus("idle", "Importando…");
    } else {
      this.setButtonLoading();
    }

    if (this.selectedFiles.length === 0) {
      this.isParsing = false;
      this.failWithError(new PcfError(404, "No file chosen."));
      this.notifyOutputChanged();
      return;
    }

    const hasTable = this.getHasTable();
    const allowMultipleFiles = this.getAllowMultipleFiles();
    const sheetIndex = this.getSheetIndex();
    const tableName = this.getTableName();
    const maxRowsToScan = this.getMaxRowsToScan();
    const enableTrace = this.getEnableTrace();

    const filesToParse = allowMultipleFiles ? this.selectedFiles : [this.selectedFiles[0]];

    const globalMeta: GlobalMeta = {
      filesSelected: this.selectedFiles.length,
      filesParsed: 0,
      allowMultipleFiles,
      hasTable,
      sheetIndex,
      tableName,
      maxRowsToScan,
      enableTrace
    };

    this.trace("INFO", "START", "Parsing started.", undefined, {
      allowMultipleFiles,
      hasTable,
      sheetIndex,
      tableName,
      maxRowsToScan,
      filesToParse: filesToParse.map(f => f.name)
    });

    try {
      const results: FileResult[] = [];
      const includeFileName = this.getIncludeFileName();

      for (const file of filesToParse) {
        this.trace("INFO", "FILE", "Reading file.", file.name);
        const r = await this.parseSingleFile(file);
        results.push(r);
        globalMeta.filesParsed++;
      }

      const flattenedRows: JsonRecord[] = [];
      for (const fileResult of results) {
        for (const row of fileResult.data) {
          const copy: JsonRecord = { ...row };
          if (includeFileName) {
            copy.fileName = fileResult.fileName;
          }
          flattenedRows.push(copy);
        }
      }

      this._jsonResult = JSON.stringify(flattenedRows);
      this._meta = JSON.stringify(globalMeta);
      this._isValid = true;
      this._errorMessage = "";

      const totalRows = results.reduce((a, r) => a + r.data.length, 0);
      const cols = results[0]?.meta.detectedColumns ?? null;
      this.lastResultSummary = { rows: totalRows, cols };

      this.trace("INFO", "END", "Parsing completed successfully.", undefined, {
        filesParsed: results.length,
        totalRows
      });

      this.flushTrace();
      this.isParsing = false;
      this.notifyOutputChanged();

      // UI: success
      if (this.currentMode === "Card") {
        this.renderChipSuccess(this.selectedFiles[0].name, totalRows, cols);
        this.setCardStatus("success", "Importação concluída — dados enviados para o app.");
      } else {
        this.flashButtonSuccess();
      }
    } catch (err: unknown) {
      this._meta = JSON.stringify(globalMeta);
      this.failWithError(err);
      this.flushTrace();
      this.isParsing = false;
      this.notifyOutputChanged();

      if (this.currentMode === "Card") {
        this.renderChipError(this.selectedFiles[0]?.name ?? "", "Falha ao processar");
        this.setCardStatus("error", `Erro: ${this._errorMessage}`);
      } else {
        this.flashButtonError();
      }
    }
  }

  private async parseSingleFile(file: File): Promise<FileResult> {
    const hasTable = this.getHasTable();
    const tableName = this.getTableName();
    const sheetIndex = this.getSheetIndex();
    const maxRowsToScan = this.getMaxRowsToScan();

    if (hasTable && (!tableName || tableName.trim() === "")) {
      throw new PcfError(404, "Table name was not provided.");
    }

    const buffer = await this.readFileAsArrayBuffer(file);
    this.trace("INFO", "WORKBOOK", "Loading workbook.", file.name);

    const wb = XLSX.read(buffer, { type: "array", cellDates: true, dense: false, bookFiles: true });

    const sheetNames = wb.SheetNames || [];
    if (sheetNames.length === 0) {
      throw new PcfError(404, "No worksheets found in the workbook.");
    }

    const safeSheetIndex = Math.max(0, Math.min(sheetIndex, sheetNames.length - 1));

    const meta: FileMeta = {
      sheetUsed: null,
      sheetIndex: safeSheetIndex,
      usedRange: null,
      mode: hasTable ? "table" : "range",
      tableName: tableName || null,
      headerRowIndex: null,
      detectedColumns: null,
      detectedRows: null
    };

    let records: JsonRecord[] = [];

    if (hasTable) {
      this.trace("INFO", "TABLE", `Searching for table "${tableName}".`, file.name);
      records = this.parseTableAcrossWorkbookOrThrow(wb, tableName!, meta, file.name);
    } else {
      const sheetName = sheetNames[safeSheetIndex];
      const sheet = wb.Sheets[sheetName];
      if (!sheet) {
        throw new PcfError(404, `Worksheet not found at index ${safeSheetIndex}.`);
      }

      meta.sheetUsed = sheetName;
      meta.usedRange = (sheet["!ref"] as string | undefined) || null;
      meta.mode = "range";

      this.trace("INFO", "RANGE", `Parsing worksheet index ${safeSheetIndex} (${sheetName}).`, file.name, {
        usedRange: meta.usedRange,
        maxRowsToScan
      });

      records = this.parseRangeOrThrow(sheet, meta, file.name);
    }

    meta.detectedRows = records.length;

    this.trace("INFO", "FILE_DONE", "File parsed successfully.", file.name, {
      rows: records.length,
      columns: meta.detectedColumns,
      sheetUsed: meta.sheetUsed,
      mode: meta.mode
    });

    return { fileName: file.name, data: records, meta };
  }

  // --------------------------
  // Table parsing (entire workbook)
  // --------------------------

  private parseTableAcrossWorkbookOrThrow(
    workbook: XLSX.WorkBook,
    tableName: string,
    meta: FileMeta,
    fileName: string
  ): JsonRecord[] {
    const found = this.locateTableRange(workbook, tableName);
    if (!found) {
      this.trace("ERROR", "TABLE_NOT_FOUND", `Table not found: ${tableName}`, fileName);
      throw new PcfError(404, `Table not found: ${tableName}`);
    }

    const { refSheetName, rangeAddress } = found;
    const sheet = workbook.Sheets[refSheetName];
    if (!sheet) {
      this.trace("ERROR", "TABLE_SHEET", `Sheet referenced by table not found: ${refSheetName}`, fileName);
      throw new PcfError(404, `Sheet referenced by table not found: ${refSheetName}`);
    }

    this.trace("INFO", "TABLE_FOUND", `Table found in sheet "${refSheetName}" range "${rangeAddress}".`, fileName);

    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, range: rangeAddress, raw: true }) as unknown[][];
    if (!aoa || aoa.length === 0) {
      this.trace("ERROR", "TABLE_EMPTY", `No data found in table: ${tableName}`, fileName);
      throw new PcfError(404, `No data found in table: ${tableName}`);
    }

    const headers = this.buildHeaders((aoa[0] as unknown[]) || []);
    const records: JsonRecord[] = [];

    for (let r = 1; r < aoa.length; r++) {
      const row = (aoa[r] as unknown[]) || [];
      const rowSlice = row.slice(0, headers.length);
      if (this.isRowEmpty(rowSlice)) continue;

      const obj: JsonRecord = {};
      for (let c = 0; c < headers.length; c++) {
        obj[headers[c]] = this.normalizeCell(rowSlice[c]);
      }
      records.push(obj);
    }

    if (records.length === 0) {
      this.trace("ERROR", "TABLE_NO_ROWS", `No data rows found in table: ${tableName}`, fileName);
      throw new PcfError(404, `No data rows found in table: ${tableName}`);
    }

    meta.mode = "table";
    meta.sheetUsed = refSheetName;
    meta.usedRange = (sheet["!ref"] as string | undefined) || null;
    meta.headerRowIndex = 0;
    meta.detectedColumns = headers.length;
    meta.detectedRows = records.length;

    return records;
  }

  // --------------------------
  // Range parsing
  // --------------------------

  private parseRangeOrThrow(sheet: XLSX.WorkSheet, meta: FileMeta, fileName: string): JsonRecord[] {
    const maxRowsToScan = this.getMaxRowsToScan();

    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true }) as unknown[][];
    if (!aoa || aoa.length === 0) {
      this.trace("ERROR", "RANGE_EMPTY_SHEET", "No cells found in worksheet used range.", fileName, { usedRange: meta.usedRange });
      throw new PcfError(404, "No data found in the first worksheet (index 0).");
    }

    const headerRowIndex = this.detectHeaderRowIndex(aoa, maxRowsToScan);
    if (headerRowIndex < 0) {
      this.trace("ERROR", "RANGE_HEADER", "Unable to detect a header row.", fileName, { maxRowsToScan });
      throw new PcfError(404, "No data found in the first worksheet (index 0).");
    }

    const rawHeaders = (aoa[headerRowIndex] as unknown[]) || [];
    const headers = this.buildHeaders(rawHeaders);

    const records: JsonRecord[] = [];
    const maxConsecutiveEmpty = 20;
    let emptyStreak = 0;

    for (let r = headerRowIndex + 1; r < aoa.length; r++) {
      const row = (aoa[r] as unknown[]) || [];
      const rowSlice = row.slice(0, headers.length);

      if (this.isRowEmpty(rowSlice)) {
        emptyStreak++;
        if (emptyStreak >= maxConsecutiveEmpty) break;
        continue;
      }
      emptyStreak = 0;

      const obj: JsonRecord = {};
      for (let c = 0; c < headers.length; c++) {
        obj[headers[c]] = this.normalizeCell(rowSlice[c]);
      }
      records.push(obj);
    }

    if (records.length === 0) {
      this.trace("ERROR", "RANGE_NO_ROWS", "Header detected but no data rows found.", fileName, {
        headerRowIndex,
        detectedColumns: headers.length
      });
      throw new PcfError(404, "No data found in the first worksheet (index 0).");
    }

    meta.headerRowIndex = headerRowIndex;
    meta.detectedColumns = headers.length;
    meta.detectedRows = records.length;

    return records;
  }

  // --------------------------
  // Chip / status rendering
  // --------------------------

  private renderChipParsing(fileName: string): void {
    if (!this.fileChipEl || !this.fileChipNameEl || !this.fileChipSubEl) return;
    this.fileChipEl.style.background = "#f3f2f1";
    this.fileChipEl.style.border = "1px solid #edebe9";
    this.fileChipNameEl.innerText = fileName;
    this.fileChipNameEl.style.color = "#201f1e";
    this.fileChipSubEl.innerText = "Importando…";
    this.fileChipSubEl.style.color = "#605e5c";
  }

  private renderChipSuccess(fileName: string, rows: number, cols: number | null): void {
    if (!this.fileChipEl || !this.fileChipNameEl || !this.fileChipSubEl) return;
    this.fileChipEl.style.background = "#f1faf1";
    this.fileChipEl.style.border = "1px solid #d1ead1";
    this.fileChipNameEl.innerText = fileName;
    this.fileChipNameEl.style.color = "#201f1e";
    const colsTxt = cols !== null ? ` · ${cols} ${cols === 1 ? "coluna" : "colunas"}` : "";
    this.fileChipSubEl.innerText = `${rows} ${rows === 1 ? "linha" : "linhas"}${colsTxt}`;
    this.fileChipSubEl.style.color = "#605e5c";
  }

  private renderChipError(fileName: string, sub: string): void {
    if (!this.fileChipEl || !this.fileChipNameEl || !this.fileChipSubEl) return;
    this.fileChipEl.style.background = "#fdf6f6";
    this.fileChipEl.style.border = "1px solid #efd0d2";
    this.fileChipNameEl.innerText = fileName;
    this.fileChipNameEl.style.color = "#201f1e";
    this.fileChipSubEl.innerText = sub;
    this.fileChipSubEl.style.color = "#a4262c";
  }

  private updateChipForCurrentState(): void {
    if (!this.selectedFiles[0]) return;
    if (this._isValid && this.lastResultSummary) {
      this.renderChipSuccess(this.selectedFiles[0].name, this.lastResultSummary.rows, this.lastResultSummary.cols);
    } else if (this._errorMessage) {
      this.renderChipError(this.selectedFiles[0].name, "Falha ao processar");
    } else if (this.isParsing) {
      this.renderChipParsing(this.selectedFiles[0].name);
    }
  }

  private setCardStatus(kind: "idle" | "success" | "error", text: string): void {
    if (!this.statusEl) return;
    this.statusEl.innerText = "";
    let leftBorder = "#c8c6c4";
    let bg = "#f3f2f1";
    if (kind === "success") {
      leftBorder = "#107C10";
      bg = "#f3f2f1";
      const check = document.createElement("span");
      check.style.display = "inline-flex";
      check.style.verticalAlign = "-3px";
      check.style.marginRight = "4px";
      check.innerHTML = `<svg width="13" height="13" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <path d="M3.5 8.5l3 3 6-7" stroke="#107C10" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>`;
      this.statusEl.appendChild(check);
    } else if (kind === "error") {
      leftBorder = "#a4262c";
      bg = "#fdf6f6";
    }
    this.statusEl.appendChild(document.createTextNode(text));
    this.statusEl.style.background = bg;
    this.statusEl.style.borderLeft = `3px solid ${leftBorder}`;
    this.statusEl.style.color = kind === "idle" ? "#605e5c" : "#201f1e";
  }

  // --------------------------
  // Button mode state transitions
  // --------------------------

  private setButtonIdle(): void {
    if (!this.buttonRoot || !this.buttonIconEl || !this.buttonTextEl) return;
    if (this.buttonResetTimer !== null) {
      window.clearTimeout(this.buttonResetTimer);
      this.buttonResetTimer = null;
    }
    this.buttonRoot.disabled = false;
    this.buttonRoot.classList.remove("pcf-iex-button--success");
    this.buttonIconEl.innerHTML = this.excelIconSvg(14);
    this.buttonTextEl.innerText = this.getButtonText();
  }

  private setButtonLoading(): void {
    if (!this.buttonRoot || !this.buttonIconEl || !this.buttonTextEl) return;
    this.buttonRoot.disabled = true;
    this.buttonRoot.classList.remove("pcf-iex-button--success");
    this.buttonIconEl.innerHTML = this.spinnerSvg(14);
    this.buttonTextEl.innerText = "Importing…";
  }

  private flashButtonSuccess(): void {
    if (!this.buttonRoot || !this.buttonIconEl || !this.buttonTextEl) return;
    this.buttonRoot.disabled = false;
    this.buttonRoot.classList.add("pcf-iex-button--success");
    this.buttonIconEl.innerHTML = this.checkSvg(14, "#107C10");
    this.buttonTextEl.innerText = "Imported";

    if (this.buttonResetTimer !== null) window.clearTimeout(this.buttonResetTimer);
    this.buttonResetTimer = window.setTimeout(() => {
      this.lastResultSummary = null;
      this.setButtonIdle();
    }, 2000);
  }

  private flashButtonError(): void {
    if (!this.buttonRoot || !this.buttonIconEl || !this.buttonTextEl) return;
    this.buttonRoot.disabled = false;
    this.buttonRoot.classList.remove("pcf-iex-button--success");
    this.buttonIconEl.innerHTML = this.errorSvg(14, "#a4262c");
    this.buttonTextEl.innerText = "Failed";
    this.buttonRoot.style.color = "#a4262c";

    if (this.buttonResetTimer !== null) window.clearTimeout(this.buttonResetTimer);
    this.buttonResetTimer = window.setTimeout(() => {
      if (this.buttonRoot) this.buttonRoot.style.color = "";
      this.setButtonIdle();
    }, 2500);
  }

  // --------------------------
  // SVG fragments
  // --------------------------

  private excelIconSvg(size: number): string {
    return `<svg width="${size}" height="${size}" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
      <path d="M3 1.25h6.75L13 4.5v8.25c0 .69-.56 1.25-1.25 1.25H3c-.69 0-1.25-.56-1.25-1.25V2.5C1.75 1.81 2.31 1.25 3 1.25Z" fill="#107C41"/>
      <path d="M9.75 1.25V4c0 .69.56 1.25 1.25 1.25H13" stroke="#33C481" stroke-width="0.9"/>
      <path d="m5.9 11.8-.83-2.31-.81 2.31H2.93l1.6-3.8-1.5-3.63h1.33l.77 2.16.77-2.16h1.33l-1.5 3.63 1.6 3.8H5.9Z" fill="#fff"/>
    </svg>`;
  }

  private spinnerSvg(size: number): string {
    return `<svg width="${size}" height="${size}" viewBox="0 0 16 16" fill="none" aria-hidden="true">
      <circle cx="8" cy="8" r="6" stroke="#c8c6c4" stroke-width="2" fill="none"/>
      <path class="pcf-iex-spinner-arc" d="M14 8a6 6 0 00-6-6" stroke="#323130" stroke-width="2" fill="none" stroke-linecap="round"/>
    </svg>`;
  }

  private checkSvg(size: number, color: string): string {
    return `<svg width="${size}" height="${size}" viewBox="0 0 16 16" fill="none" aria-hidden="true">
      <path d="M3 8.5l3 3 6.5-7" stroke="${color}" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
    </svg>`;
  }

  private errorSvg(size: number, color: string): string {
    return `<svg width="${size}" height="${size}" viewBox="0 0 16 16" fill="none" aria-hidden="true">
      <circle cx="8" cy="8" r="6.3" fill="none" stroke="${color}" stroke-width="1.3"/>
      <path d="M8 4.5v4M8 10.5v.6" stroke="${color}" stroke-width="1.5" stroke-linecap="round"/>
    </svg>`;
  }

  // --------------------------
  // Helpers
  // --------------------------

  private readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new PcfError(404, "Unable to read the file."));
      reader.onload = () => {
        const result = reader.result;
        if (result instanceof ArrayBuffer) resolve(result);
        else reject(new PcfError(404, "Unexpected file reader result type."));
      };
      reader.readAsArrayBuffer(file);
    });
  }

  private extractSheetAndRange(ref: string): { refSheetName: string; rangeAddress: string } | null {
    const parts = ref.split("!");
    if (parts.length !== 2) return null;

    let refSheetName = parts[0].trim();
    if (refSheetName.startsWith("'") && refSheetName.endsWith("'")) {
      refSheetName = refSheetName.slice(1, -1);
    }

    const rangeAddress = parts[1].replace(/\$/g, "").trim();
    if (!rangeAddress) return null;

    return { refSheetName, rangeAddress };
  }

  private locateTableRange(
    workbook: XLSX.WorkBook,
    tableName: string
  ): { refSheetName: string; rangeAddress: string } | null {
    const normalizedTarget = this.normalizeTableName(tableName);
    const names = workbook.Workbook?.Names || [];

    for (const n of names) {
      const nameValue = n.Name || "";
      const normalizedName = this.normalizeTableName(nameValue);
      const normalizedLocalName = this.normalizeTableName(nameValue.split("!").pop() || "");

      if (normalizedName === normalizedTarget || normalizedLocalName === normalizedTarget) {
        if (n.Ref) {
          const parsed = this.extractSheetAndRange(n.Ref);
          if (parsed) return parsed;
        }
      }
    }

    return this.findTableRangeFromFiles(workbook, normalizedTarget);
  }

  private findTableRangeFromFiles(
    workbook: XLSX.WorkBook,
    normalizedTarget: string
  ): { refSheetName: string; rangeAddress: string } | null {
    const wbWithFiles = workbook as unknown as { files?: Record<string, { content?: unknown }>; keys?: string[] };
    const files = wbWithFiles.files;
    const keys = wbWithFiles.keys;
    if (!files || !keys) return null;

    const parser = new DOMParser();
    const tableDefs = new Map<string, { name: string; ref: string }>();

    for (const path of keys) {
      const normalizedPath = path.replace(/\\/g, "/");
      const lower = normalizedPath.toLowerCase();
      if (!lower.startsWith("xl/tables/") || !lower.endsWith(".xml")) continue;

      const raw = files[normalizedPath]?.content;
      const xml = this.bufferToText(raw);
      if (!xml) continue;

      const doc = parser.parseFromString(xml, "application/xml");
      const tableEl = doc.getElementsByTagName("table")[0];
      if (!tableEl) continue;

      const nameAttr = tableEl.getAttribute("name") || tableEl.getAttribute("displayName") || "";
      const refAttr = (tableEl.getAttribute("ref") || "").replace(/\$/g, "");
      if (!nameAttr || !refAttr) continue;

      tableDefs.set(normalizedPath, { name: nameAttr, ref: refAttr });
    }

    if (tableDefs.size === 0) return null;

    const sheetNames = workbook.SheetNames || [];

    for (let idx = 0; idx < sheetNames.length; idx++) {
      const relPath = `xl/worksheets/_rels/sheet${idx + 1}.xml.rels`;
      const relRaw = files[relPath]?.content;
      if (!relRaw) continue;

      const relXml = this.bufferToText(relRaw);
      if (!relXml) continue;

      const relDoc = parser.parseFromString(relXml, "application/xml");
      const relationships = Array.from(relDoc.getElementsByTagName("Relationship"));

      for (const rel of relationships) {
        const typeAttr = rel.getAttribute("Type") || "";
        if (!typeAttr.toLowerCase().includes("/table")) continue;

        const targetAttr = rel.getAttribute("Target") || "";
        const tablePath = this.normalizeTablePath(targetAttr);
        const def = tableDefs.get(tablePath);
        if (!def) continue;

        if (this.normalizeTableName(def.name) === normalizedTarget) {
          return { refSheetName: sheetNames[idx], rangeAddress: def.ref };
        }
      }
    }

    if (sheetNames.length === 1) {
      for (const def of tableDefs.values()) {
        if (this.normalizeTableName(def.name) === normalizedTarget) {
          return { refSheetName: sheetNames[0], rangeAddress: def.ref };
        }
      }
    }

    return null;
  }

  private normalizeTableName(name: string): string {
    return name
      .replace(/^.*!/, "")
      .replace(/\[#.*\]/g, "")
      .replace(/['"]/g, "")
      .trim()
      .toLowerCase();
  }

  private normalizeTablePath(target: string): string {
    const cleaned = target.replace(/\\/g, "/");
    if (cleaned.startsWith("../")) return `xl/${cleaned.slice(3)}`;
    if (cleaned.startsWith("/")) return cleaned.slice(1);
    if (!cleaned.startsWith("xl/")) return `xl/${cleaned}`;
    return cleaned;
  }

  private decodeBufferToText(bytes: Uint8Array): string {
    if (typeof TextDecoder !== "undefined") {
      try {
        return new TextDecoder("utf-8").decode(bytes);
      } catch {
        // Ignore and fall back
      }
    }

    let out = "";
    // eslint-disable-next-line @typescript-eslint/prefer-for-of
    for (let i = 0; i < bytes.length; i++) {
      out += String.fromCharCode(bytes[i]);
    }
    return out;
  }

  private bufferToText(data: unknown): string {
    if (!data) return "";
    if (typeof data === "string") return data;

    if (typeof Buffer !== "undefined" && typeof (Buffer as unknown as { isBuffer?(value: unknown): boolean }).isBuffer === "function") {
      const maybeBuffer = data as Buffer;
      if ((Buffer as unknown as { isBuffer(value: unknown): boolean }).isBuffer(maybeBuffer)) {
        return maybeBuffer.toString("utf8");
      }
    }

    if (data instanceof ArrayBuffer) {
      return this.decodeBufferToText(new Uint8Array(data));
    }

    if (typeof ArrayBuffer !== "undefined" && ArrayBuffer.isView(data)) {
      const view = data as ArrayBufferView;
      return this.decodeBufferToText(new Uint8Array(view.buffer));
    }

    return "";
  }

  private detectHeaderRowIndex(aoa: unknown[][], maxRowsToScan: number): number {
    const limit = Math.min(maxRowsToScan, aoa.length);

    let bestIndex = -1;
    let bestScore = -Infinity;

    for (let i = 0; i < limit; i++) {
      const row = aoa[i] || [];
      const score = this.scoreHeaderCandidate(row, aoa[i + 1] || []);
      if (score > bestScore) {
        bestScore = score;
        bestIndex = i;
      }
    }

    return bestScore >= 3 ? bestIndex : -1;
  }

  private scoreHeaderCandidate(row: unknown[], nextRow: unknown[]): number {
    const cells = row.map(v => (v === null || v === undefined ? "" : String(v).trim()));
    const nonEmpty = cells.filter(c => c !== "");
    if (nonEmpty.length < 2) return -10;

    const unique = new Set(nonEmpty.map(c => c.toLowerCase()));
    const uniquenessRatio = unique.size / nonEmpty.length;

    const numericCount = nonEmpty.filter(v => this.looksNumeric(v)).length;
    const numericRatio = numericCount / nonEmpty.length;

    const nextNonEmpty = nextRow
      .slice(0, cells.length)
      .map(v => (v === null || v === undefined ? "" : String(v).trim()))
      .filter(c => c !== "").length;

    let score = 0;
    score += Math.min(nonEmpty.length, 10);
    score += uniquenessRatio * 5;
    score -= numericRatio * 5;
    score += Math.min(nextNonEmpty, 10) / 2;

    const avgLen = nonEmpty.reduce((a, c) => a + c.length, 0) / nonEmpty.length;
    if (avgLen > 30) score -= 3;

    return score;
  }

  private buildHeaders(rawHeaders: unknown[]): string[] {
    const base = rawHeaders.map((h, idx) => {
      const txt = h === null || h === undefined ? "" : String(h).trim();
      return txt !== "" ? txt : `Column${idx + 1}`;
    });

    const seen = new Map<string, number>();
    return base.map(h => {
      const lower = h.toLowerCase();
      const count = seen.get(lower) || 0;
      seen.set(lower, count + 1);
      return count === 0 ? h : `${h}_${count + 1}`;
    });
  }

  private isRowEmpty(row: unknown[]): boolean {
    for (const cell of row) {
      if (cell === null || cell === undefined) continue;
      if (typeof cell === "string" && cell.trim() === "") continue;
      return false;
    }
    return true;
  }

  private looksNumeric(value: string): boolean {
    const v = value.replace(/,/g, "");
    return /^-?\d+(\.\d+)?$/.test(v);
  }

  private normalizeCell(value: unknown): unknown {
    if (value instanceof Date) return value.toISOString();
    if (typeof value === "number") return value;
    if (typeof value === "boolean") return value;
    if (value === null || value === undefined) return null;
    if (typeof value === "string") {
      const t = value.trim();
      return t === "" ? null : t;
    }
    return value as unknown;
  }

  // --------------------------
  // Inputs
  // --------------------------

  private getDisplayMode(): DisplayMode {
    const raw = this.context.parameters.displayMode?.raw;
    return raw === "Button" ? "Button" : "Card";
  }

  private getShowTitle(): boolean {
    return this.context.parameters.showTitle?.raw !== false;
  }

  private getTitleText(): string {
    const v = this.context.parameters.title?.raw;
    const s = v === null || v === undefined ? "" : String(v);
    return s.trim() || "Import Excel";
  }

  private getShowUserGuidelines(): boolean {
    return this.context.parameters.showUserGuidelines?.raw !== false;
  }

  private getUserGuidelinesText(): string {
    const v = this.context.parameters.userGuidelinesText?.raw;
    return v === null || v === undefined ? "" : String(v);
  }

  private getShowStatus(): boolean {
    return this.context.parameters.showStatus?.raw !== false;
  }

  private getIncludeFileName(): boolean {
    return this.context.parameters.includeFileName?.raw === true;
  }

  private getDropzonePrimaryText(): string {
    const v = this.context.parameters.dropzonePrimaryText?.raw;
    const s = v === null || v === undefined ? "" : String(v);
    return s.trim() || "Drag your file here or click to browse";
  }

  private getDropzoneHintText(): string {
    const v = this.context.parameters.dropzoneHintText?.raw;
    const s = v === null || v === undefined ? "" : String(v);
    return s.trim() || ".xlsx · .xls · .csv";
  }

  private getButtonText(): string {
    const v = this.context.parameters.buttonText?.raw;
    const s = v === null || v === undefined ? "" : String(v);
    return s.trim() || "Import";
  }

  private getRemoveButtonText(): string {
    const v = this.context.parameters.removeButtonText?.raw;
    const s = v === null || v === undefined ? "" : String(v);
    return s.trim() || "Remove";
  }

  private getAccentColor(): string {
    const v = this.context.parameters.accentColor?.raw;
    const s = v === null || v === undefined ? "" : String(v).trim();
    return s || "#323130";
  }

  private getBackgroundColor(): string {
    const v = this.context.parameters.backgroundColor?.raw;
    const s = v === null || v === undefined ? "" : String(v).trim();
    return s || "#ffffff";
  }

  private getBorderColor(): string {
    const v = this.context.parameters.borderColor?.raw;
    const s = v === null || v === undefined ? "" : String(v).trim();
    return s || "#e1dfdd";
  }

  private getAllowMultipleFiles(): boolean {
    return this.context.parameters.allowMultipleFiles?.raw === true;
  }

  private getHasTable(): boolean {
    return this.context.parameters.hasTable?.raw === true;
  }

  private getTableName(): string | null {
    const v = this.context.parameters.tableName?.raw;
    if (v === null || v === undefined) return null;
    const s = String(v).trim();
    return s.length ? s : null;
  }

  private getSheetIndex(): number {
    const raw = this.context.parameters.sheetIndex?.raw;
    const n = typeof raw === "number" ? raw : Number(raw);
    return Number.isFinite(n) ? n : 0;
  }

  private getMaxRowsToScan(): number {
    const raw = this.context.parameters.maxRowsToScan?.raw;
    const n = typeof raw === "number" ? raw : Number(raw);
    const val = Number.isFinite(n) ? n : 50;
    return Math.max(5, Math.min(val, 200));
  }

  private getEnableTrace(): boolean {
    return this.context.parameters.enableTrace?.raw === true;
  }

  // --------------------------
  // Trace + outputs
  // --------------------------

  private resetOutputsAndTrace(): void {
    this._jsonResult = "";
    this._meta = "";
    this._trace = "";
    this._isValid = false;
    this._errorMessage = "";
    this.traceEntries = [];
    this.lastResultSummary = null;
  }

  private trace(level: TraceLevel, step: string, message: string, fileName?: string, details?: Record<string, unknown>): void {
    if (!this.isTraceEnabled()) return;

    this.traceEntries.push({
      ts: new Date().toISOString(),
      level,
      step,
      message,
      fileName,
      details
    });
  }

  private flushTrace(): void {
    this._trace = this.isTraceEnabled() ? JSON.stringify(this.traceEntries) : "";
  }

  private failWithError(err: unknown): void {
    const msg = err instanceof Error ? err.message : String(err);
    this._jsonResult = "";
    this._isValid = false;
    this._errorMessage = msg;
    this.trace("ERROR", "FAIL", msg);
  }

  private isTraceEnabled(): boolean {
    return this.getEnableTrace();
  }
}
