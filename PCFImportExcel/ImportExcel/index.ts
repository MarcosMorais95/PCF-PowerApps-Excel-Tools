import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as XLSX from "xlsx";

type JsonRecord = Record<string, unknown>;
type TraceLevel = "INFO" | "WARN" | "ERROR";

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

export class ImportExcel implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private context!: ComponentFramework.Context<IInputs>;
  private container!: HTMLDivElement;
  private notifyOutputChanged!: () => void;

  private fileInput!: HTMLInputElement;
  private filePickerRow!: HTMLDivElement;
  private filePickerButton!: HTMLButtonElement;
  private filePickerLabel!: HTMLDivElement;
  private statusEl!: HTMLDivElement;
  private titleEl!: HTMLDivElement;
  private guidelinesEl!: HTMLDivElement;
  private actionsRow!: HTMLDivElement;
  private btnParse!: HTMLButtonElement;
  private btnClear!: HTMLButtonElement;

  private _jsonResult = "";
  private _meta = "";
  private _trace = "";
  private _isValid = false;
  private _errorMessage = "";

  private selectedFiles: File[] = [];
  private traceEntries: TraceEntry[] = [];

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.context = context;
    this.notifyOutputChanged = notifyOutputChanged;

    this.container = document.createElement("div");
    this.container.style.fontFamily = "Segoe UI, Arial, sans-serif";
    this.container.style.width = "100%";
    this.container.style.boxSizing = "border-box";

    this.titleEl = document.createElement("div");
    this.titleEl.innerText = "Import Excel";
    this.titleEl.style.fontSize = "14px";
    this.titleEl.style.fontWeight = "600";
    this.titleEl.style.marginBottom = "6px";

    this.guidelinesEl = document.createElement("div");
    this.guidelinesEl.style.fontSize = "12px";
    this.guidelinesEl.style.opacity = "0.85";
    this.guidelinesEl.style.marginBottom = "10px";

    this.fileInput = document.createElement("input");
    this.fileInput.type = "file";
    this.fileInput.accept =
      ".xlsx,.xlsm,.xlsb,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv";
    this.fileInput.style.display = "none";

    this.filePickerRow = document.createElement("div");
    this.filePickerRow.style.display = "flex";
    this.filePickerRow.style.gap = "8px";
    this.filePickerRow.style.alignItems = "center";
    this.filePickerRow.style.marginBottom = "8px";

    this.filePickerButton = document.createElement("button");
    this.filePickerButton.type = "button";
    this.filePickerButton.style.padding = "6px 10px";
    this.filePickerButton.style.cursor = "pointer";
    this.filePickerButton.innerText = this.getChooseFilesText();

    this.filePickerLabel = document.createElement("div");
    this.filePickerLabel.style.fontSize = "12px";
    this.filePickerLabel.style.opacity = "0.85";
    this.filePickerLabel.innerText = this.getNoFileChosenText();

    this.filePickerRow.appendChild(this.filePickerButton);
    this.filePickerRow.appendChild(this.filePickerLabel);
    this.filePickerRow.appendChild(this.fileInput);

    this.actionsRow = document.createElement("div");
    this.actionsRow.style.display = "flex";
    this.actionsRow.style.gap = "8px";
    this.actionsRow.style.alignItems = "center";
    this.actionsRow.style.marginBottom = "10px";

    this.btnParse = document.createElement("button");
    this.btnParse.type = "button";
    this.btnParse.innerText = "Parse";
    this.btnParse.style.padding = "6px 10px";
    this.btnParse.style.cursor = "pointer";

    this.btnClear = document.createElement("button");
    this.btnClear.type = "button";
    this.btnClear.innerText = "Clear";
    this.btnClear.style.padding = "6px 10px";
    this.btnClear.style.cursor = "pointer";

    this.actionsRow.appendChild(this.btnParse);
    this.actionsRow.appendChild(this.btnClear);

    this.statusEl = document.createElement("div");
    this.statusEl.style.fontSize = "12px";
    this.statusEl.style.whiteSpace = "pre-wrap";
    this.statusEl.style.padding = "8px";
    this.statusEl.style.border = "1px solid #e1e1e1";
    this.statusEl.style.borderRadius = "6px";
    this.statusEl.style.background = "#fafafa";

    this.container.appendChild(this.titleEl);
    this.container.appendChild(this.guidelinesEl);
    this.container.appendChild(this.filePickerRow);
    this.container.appendChild(this.actionsRow);
    this.container.appendChild(this.statusEl);

    container.appendChild(this.container);

    this.filePickerButton.addEventListener("click", () => this.fileInput.click());

    this.fileInput.addEventListener("change", () => {
      const files = this.fileInput.files ? Array.from(this.fileInput.files) : [];
      this.selectedFiles = files;

      if (this.selectedFiles.length === 0) {
        this.updateFilePickerLabel();
        this.setStatus(this.getNoFileChosenText(), false, "");
        return;
      }

      if (!this.getAllowMultipleFiles() && this.selectedFiles.length > 1) {
        this.selectedFiles = [this.selectedFiles[0]];
        this.setStatus(
          `Multiple files selected, but the control is set to single-file mode.\nUsing: ${this.selectedFiles[0].name}`,
          true,
          ""
        );
      } else {
        const names = this.selectedFiles.map(f => f.name).join("\n- ");
        this.setStatus(`Selected ${this.selectedFiles.length} file(s):\n- ${names}`, true, "");
      }

      this.updateFilePickerLabel();
    });

    this.btnParse.addEventListener("click", () => void this.tryParseSelectedFiles());

    this.btnClear.addEventListener("click", () => {
      this.resetOutputsAndTrace();
      this.selectedFiles = [];
      this.fileInput.value = "";
      this.notifyOutputChanged();
      this.setStatus("Cleared.", true, "Select an Excel file to parse.");
      this.updateFilePickerLabel();
    });

    this.setStatus("Ready.", true, "Select an Excel file to parse.");
    this.updateFilePickerLabel();
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this.context = context;

    const allowMultiple = this.getAllowMultipleFiles();
    this.fileInput.multiple = allowMultiple;

    if (!allowMultiple && this.selectedFiles.length > 1) {
      this.selectedFiles = [this.selectedFiles[0]];
      this.fileInput.value = "";
      this.setStatus("The control switched to single-file mode. Please reselect the file.", true, "");
      this.updateFilePickerLabel();
    }

    this.titleEl.style.display = this.getShowTitle() ? "block" : "none";
    this.titleEl.innerText = this.getTitleText();

    this.guidelinesEl.style.display = this.getShowUserGuidelines() ? "block" : "none";
    this.guidelinesEl.innerText = this.getUserGuidelinesText();

    this.filePickerButton.innerText = this.getChooseFilesText();
    this.updateFilePickerLabel();

    this.statusEl.style.display = this.getShowStatus() ? "block" : "none";

    this.btnParse.innerText = this.getParseButtonText();
    this.btnClear.innerText = this.getClearButtonText();
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
    // No-op
  }

  // --------------------------
  // Orchestration
  // --------------------------

  private async tryParseSelectedFiles(): Promise<void> {
    this.resetOutputsAndTrace();

    if (this.selectedFiles.length === 0) {
      this.failWithError(new PcfError(404, this.getNoFileChosenText()));
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

      this.trace("INFO", "END", "Parsing completed successfully.", undefined, {
        filesParsed: results.length,
        totalRows: results.reduce((a, r) => a + r.data.length, 0)
      });

      this.flushTrace();
      this.notifyOutputChanged();

      this.setStatus(
        `Success.\nFiles parsed: ${results.length}\nTotal rows: ${results.reduce((a, r) => a + r.data.length, 0)}`,
        true,
        ""
      );
    } catch (err: unknown) {
      this._meta = JSON.stringify(globalMeta);
      this.failWithError(err);
      this.flushTrace();
      this.notifyOutputChanged();
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
  // Table parsing (entire workbook) - MUST succeed when hasTable=true
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
  // Range parsing - MUST succeed when hasTable=false
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

  private updateFilePickerLabel(): void {
    if (!this.filePickerLabel) return;

    if (this.selectedFiles.length === 0) {
      this.filePickerLabel.innerText = this.getNoFileChosenText();
      return;
    }

    const names = this.selectedFiles.map(f => f.name).join(", ");
    this.filePickerLabel.innerText = names;
  }

  // --------------------------
  // Inputs
  // --------------------------

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

  private getParseButtonText(): string {
    const v = this.context.parameters.parseButtonText?.raw;
    const s = v === null || v === undefined ? "" : String(v);
    return s.trim() || "Parse";
  }

  private getClearButtonText(): string {
    const v = this.context.parameters.clearButtonText?.raw;
    const s = v === null || v === undefined ? "" : String(v);
    return s.trim() || "Clear";
  }

  private getChooseFilesText(): string {
    const v = this.context.parameters.chooseFilesText?.raw;
    const s = v === null || v === undefined ? "" : String(v);
    const fallback = this.getAllowMultipleFiles() ? "Choose files" : "Choose file";
    return s.trim() || fallback;
  }

  private getNoFileChosenText(): string {
    const v = this.context.parameters.noFileChosenText?.raw;
    const s = v === null || v === undefined ? "" : String(v);
    const fallback = this.getAllowMultipleFiles() ? "No files chosen." : "No file chosen.";
    return s.trim() || fallback;
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
  // Trace + status
  // --------------------------

  private resetOutputsAndTrace(): void {
    this._jsonResult = "";
    this._meta = "";
    this._trace = "";
    this._isValid = false;
    this._errorMessage = "";
    this.traceEntries = [];
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
    this.setStatus(`Error: ${msg}`, false, "");
  }

  private setStatus(main: string, ok: boolean, hint: string): void {
    const statusTitle = ok ? "Status" : "Status (error)";
    this.statusEl.innerText = `${statusTitle}\n${main}${hint ? `\n\n${hint}` : ""}`;
    this.statusEl.style.borderColor = ok ? "#e1e1e1" : "#d13438";
    this.statusEl.style.background = ok ? "#fafafa" : "#fff5f5";
  }

  private isTraceEnabled(): boolean {
    return this.getEnableTrace();
  }
}
