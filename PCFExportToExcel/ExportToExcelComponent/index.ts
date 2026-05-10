import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as XLSX from "xlsx";

/**
 * Strong JSON types (no `any`)
 */
type JsonPrimitive = string | number | boolean | null;
type JsonValue = JsonPrimitive | JsonObject | JsonValue[];
interface JsonObject {
  [key: string]: JsonValue;
}

interface ExplicitMappedColumn {
  name: string;
  field: string;
}

export class ExportToExcelComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private _context!: ComponentFramework.Context<IInputs>;
  private _notifyOutputChanged!: () => void;

  private _container!: HTMLDivElement;
  private _button!: HTMLButtonElement;
  private _iconSpan!: HTMLSpanElement;
  private _textSpan!: HTMLSpanElement;

  private _onButtonClickBound!: () => void;

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this._context = context;
    this._notifyOutputChanged = notifyOutputChanged;

    this._container = document.createElement("div");
    this._container.style.width = "100%";
    this._container.style.height = "100%";
    this._container.style.display = "flex";
    this._container.style.alignItems = "center";
    this._container.style.justifyContent = "flex-start";

    this._button = document.createElement("button");
    this._button.type = "button";
    this._button.style.display = "inline-flex";
    this._button.style.alignItems = "center";
    this._button.style.justifyContent = "center";
    this._button.style.gap = "8px";
    this._button.style.cursor = "pointer";
    this._button.style.userSelect = "none";
    this._button.style.whiteSpace = "nowrap";
    this._button.style.padding = "10px 14px";
    this._button.style.lineHeight = "1";
    this._button.style.outline = "none";
    this._button.style.boxSizing = "border-box";

    // Icon
    this._iconSpan = document.createElement("span");
    this._iconSpan.style.display = "inline-flex";
    this._iconSpan.style.alignItems = "center";
    this._iconSpan.style.justifyContent = "center";
    this._iconSpan.style.width = "16px";
    this._iconSpan.style.height = "16px";

    // Inline Excel-like SVG icon (no external assets)
    this._iconSpan.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
        <path d="M3 1.25h6.75L13 4.5v8.25c0 .69-.56 1.25-1.25 1.25H3c-.69 0-1.25-.56-1.25-1.25V2.5C1.75 1.81 2.31 1.25 3 1.25Z" fill="#107C41"/>
        <path d="M9.75 1.25V4c0 .69.56 1.25 1.25 1.25H13" stroke="#33C481" stroke-width="0.9"/>
        <path d="m5.9 11.8-.83-2.31-.81 2.31H2.93l1.6-3.8-1.5-3.63h1.33l.77 2.16.77-2.16h1.33l-1.5 3.63 1.6 3.8H5.9Z" fill="#ffffff"/>
      </svg>
    `;

    // Text
    this._textSpan = document.createElement("span");
    this._textSpan.textContent = "Export Excel";

    this._button.appendChild(this._iconSpan);
    this._button.appendChild(this._textSpan);

    this._onButtonClickBound = this.onButtonClick.bind(this);
    this._button.addEventListener("click", this._onButtonClickBound);

    this._container.appendChild(this._button);
    container.appendChild(this._container);

    this.updateView(context);
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this._context = context;

    const buttonText = (context.parameters.buttonText.raw ?? "Export Excel").trim();
    this._textSpan.textContent = buttonText.length ? buttonText : "Export Excel";

    const showIcon = this.asBoolean(context.parameters.showIcon.raw);
    this._iconSpan.style.display = showIcon ? "inline-flex" : "none";

    const isDisabled = this.asBoolean(context.parameters.isDisabled.raw);
    this._button.disabled = isDisabled;

    this.applyButtonStyles(isDisabled);
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    if (this._button && this._onButtonClickBound) {
      this._button.removeEventListener("click", this._onButtonClickBound);
    }
  }

  private onButtonClick(): void {
    try {
      const dataText = this._context.parameters.dataJson.raw;
      const json = this.parseJsonValue(dataText);
      const useFieldMapping = this.asBoolean(this._context.parameters.useFieldMapping.raw);

      const sheetName = (this._context.parameters.sheetName.raw ?? "Export").trim() || "Export";
      const fileNameRaw = (this._context.parameters.fileName.raw ?? "").trim();
      const fileName = this.normalizeFileName(fileNameRaw || `export_${this.formatDateForFileName(new Date())}.xlsx`);

      const { aoa, headers } = this.convertToAOA(json, useFieldMapping);
      if (aoa.length === 0) {
        this.notify("No data to export.", "warning");
        return;
      }

      const worksheet = XLSX.utils.aoa_to_sheet(aoa);

      if (headers.length > 0) {
        const range = XLSX.utils.decode_range(worksheet["!ref"] ?? "A1:A1");
        worksheet["!autofilter"] = { ref: XLSX.utils.encode_range(range) };
      }

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

      const wbout: ArrayBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });

      this.downloadArrayBuffer(wbout, fileName);
      this.notify("Export completed.", "success");
    } catch (err) {
      const message = err instanceof Error ? err.message : "Unexpected error.";
      this.notify(`Export failed: ${message}`, "error");
    }
  }

  /**
   * Applies styles using only backgroundColor/fontColor properties.
   * Also hard-forces no transparency (common host CSS issue).
   */
  private applyButtonStyles(isDisabled: boolean): void {
    const bgRaw = (this._context.parameters.backgroundColor.raw ?? "").trim();
    const fgRaw = (this._context.parameters.fontColor.raw ?? "").trim();

    // Fallbacks (any colors)
    const bg = bgRaw.length ? bgRaw : "#0078d4";
    const fg = fgRaw.length ? fgRaw : "#ffffff";

    const fontSizePx = this._context.parameters.fontSizePx.raw ?? 12;
    const borderRadiusPx = this._context.parameters.borderRadiusPx.raw ?? 6;
    const widthPx = this.toPositiveNumber(this._context.parameters.buttonWidthPx.raw, 120);
    const heightPx = this.toPositiveNumber(this._context.parameters.buttonHeightPx.raw, 32);

    // Force visibility (no transparency)
    this._button.style.setProperty("opacity", isDisabled ? "0.65" : "1", "important");
    this._button.style.setProperty("background", bg, "important");
    this._button.style.setProperty("background-color", bg, "important");
    this._button.style.setProperty("color", fg, "important");

    this._button.style.setProperty("border", `1px solid ${bg}`, "important");
    this._button.style.setProperty("border-radius", `${borderRadiusPx}px`, "important");

    this._button.style.setProperty("font-size", `${fontSizePx}px`, "important");
    this._button.style.setProperty("font-family", "inherit", "important");

    this._button.style.setProperty("width", `${widthPx}px`, "important");
    this._button.style.setProperty("height", `${heightPx}px`, "important");
    this._button.style.setProperty("min-width", `${widthPx}px`, "important");
    this._button.style.setProperty("min-height", `${heightPx}px`, "important");
    this._button.style.setProperty("cursor", isDisabled ? "not-allowed" : "pointer", "important");

    // Neutralize host/browser styles that can affect visuals
    this._button.style.setProperty("box-shadow", "none", "important");
    this._button.style.setProperty("filter", "none", "important");
    this._button.style.setProperty("-webkit-appearance", "none", "important");
  }

  private convertToAOA(
    value: JsonValue,
    useFieldMapping: boolean
  ): { aoa: (string | number | boolean | null)[][]; headers: string[] } {
    if (useFieldMapping) {
      return this.tryConvertMappedColumnsTable(value);
    }

    const explicit = this.tryConvertExplicitColumnsTable(value);
    if (explicit) return explicit;

    if (Array.isArray(value)) {
      if (value.length === 0) return { aoa: [], headers: [] };

      // Array of arrays -> assume already tabular
      if (Array.isArray(value[0])) {
        const rows: (string | number | boolean | null)[][] = (value as JsonValue[])
          .filter(Array.isArray)
          .map(row => (row as JsonValue[]).map(c => this.toCellValue(c)));
        return { aoa: rows, headers: rows.length > 0 ? rows[0].map(v => String(v ?? "")) : [] };
      }

      // Array of objects -> union of keys, preserving first input-object order
      const objs = value.filter(v => this.isPlainObject(v)) as JsonObject[];
      if (objs.length === 0) return { aoa: [], headers: [] };

      const headers = this.unionKeysPreserveInputOrder(objs);
      const aoa: (string | number | boolean | null)[][] = [];
      aoa.push(headers);

      for (const obj of objs) {
        const row = headers.map(h => this.toCellValue(obj[h]));
        aoa.push(row);
      }

      return { aoa, headers };
    }

    // Single object -> 2-column key/value table
    if (this.isPlainObject(value)) {
      const obj = value as JsonObject;
      const keys = Object.keys(obj);
      if (keys.length === 0) return { aoa: [], headers: [] };

      const aoa: (string | number | boolean | null)[][] = [["Field", "Value"]];
      for (const k of keys) {
        aoa.push([k, this.toCellValue(obj[k])]);
      }
      return { aoa, headers: ["Field", "Value"] };
    }

    // Primitive -> single cell
    return { aoa: [[this.toCellValue(value)]], headers: [] };
  }

  private tryConvertMappedColumnsTable(
    value: JsonValue
  ): { aoa: (string | number | boolean | null)[][]; headers: string[] } {
    if (!this.isPlainObject(value)) {
      throw new Error("Mapped mode requires an object payload with 'columns' and 'rows'/'data'/'items'.");
    }

    const obj = value as JsonObject;
    const columnsRaw = obj["columns"];
    const rowsRaw = obj["rows"] ?? obj["data"] ?? obj["items"];

    if (!Array.isArray(columnsRaw) || !Array.isArray(rowsRaw)) {
      throw new Error("Mapped mode requires an object payload with 'columns' and 'rows'/'data'/'items'.");
    }

    const columns: ExplicitMappedColumn[] = columnsRaw.map(column => {
      if (!this.isPlainObject(column)) {
        throw new Error("Each mapped column must include non-empty 'name' and 'field' properties.");
      }

      const name = column["name"];
      const field = column["field"];

      if (
        typeof name !== "string" ||
        name.trim().length === 0 ||
        typeof field !== "string" ||
        field.trim().length === 0
      ) {
        throw new Error("Each mapped column must include non-empty 'name' and 'field' properties.");
      }

      return {
        name,
        field
      };
    });

    const rowObjects = rowsRaw.map(row => {
      if (!this.isPlainObject(row)) {
        throw new Error("Mapped mode rows must be objects.");
      }

      return row;
    });

    const headers = columns.map(column => column.name);
    const fieldKeys = columns.map(column => column.field);
    const aoa: (string | number | boolean | null)[][] = [headers];

    for (const rowObj of rowObjects) {
      aoa.push(fieldKeys.map(fieldKey => this.toCellValue(rowObj[fieldKey])));
    }

    return { aoa, headers };
  }

  private tryConvertExplicitColumnsTable(
    value: JsonValue
  ): { aoa: (string | number | boolean | null)[][]; headers: string[] } | null {
    if (!this.isPlainObject(value)) return null;

    const obj = value as JsonObject;
    const columnsRaw = obj["columns"];
    const rowsRaw = obj["rows"] ?? obj["data"] ?? obj["items"];

    if (!Array.isArray(columnsRaw) || !Array.isArray(rowsRaw)) return null;

    const headers = columnsRaw
      .map(c => {
        if (typeof c === "string") return c;
        if (this.isPlainObject(c) && typeof c["name"] === "string") return c["name"];
        return null;
      })
      .filter((h): h is string => typeof h === "string" && h.trim().length > 0);

    if (headers.length === 0) return { aoa: [], headers: [] };

    const rows = rowsRaw.filter(r => this.isPlainObject(r)) as JsonObject[];
    const aoa: (string | number | boolean | null)[][] = [headers];

    for (const rowObj of rows) {
      aoa.push(headers.map(h => this.toCellValue(rowObj[h])));
    }

    return { aoa, headers };
  }

  private unionKeysPreserveInputOrder(objs: JsonObject[]): string[] {
    const headers: string[] = [];
    const seen = new Set<string>();

    for (const obj of objs) {
      for (const key of Object.keys(obj)) {
        if (!seen.has(key)) {
          seen.add(key);
          headers.push(key);
        }
      }
    }

    return headers;
  }

  private toCellValue(v: JsonValue): string | number | boolean | null {
    if (v === null) return null;
    if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return v;

    try {
      return JSON.stringify(v);
    } catch {
      return String(v);
    }
  }

  private isPlainObject(v: JsonValue): v is JsonObject {
    return typeof v === "object" && v !== null && !Array.isArray(v);
  }

  private parseJsonValue(text: string | null | undefined): JsonValue {
    const raw = (text ?? "").trim();
    if (!raw) throw new Error("dataJson is empty.");

    try {
      return JSON.parse(raw) as JsonValue;
    } catch {
      throw new Error("dataJson is not valid JSON.");
    }
  }

  private downloadArrayBuffer(data: ArrayBuffer, fileName: string): void {
    const blob = new Blob([data], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const url = URL.createObjectURL(blob);
    try {
      const a = document.createElement("a");
      a.href = url;
      a.download = fileName;
      a.style.display = "none";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  private normalizeFileName(name: string): string {
    const trimmed = name.trim();
    if (!trimmed) return "export.xlsx";
    if (trimmed.toLowerCase().endsWith(".xlsx")) return trimmed;
    return `${trimmed}.xlsx`;
  }

  private formatDateForFileName(d: Date): string {
    const yyyy = String(d.getFullYear());
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    const hh = String(d.getHours()).padStart(2, "0");
    const mi = String(d.getMinutes()).padStart(2, "0");
    return `${yyyy}${mm}${dd}_${hh}${mi}`;
  }

  private toPositiveNumber(value: number | null | undefined, fallback: number): number {
    if (typeof value === "number" && Number.isFinite(value) && value > 0) {
      return value;
    }
    return fallback;
  }

  private asBoolean(v: boolean | number | string | null | undefined): boolean {
    if (typeof v === "boolean") return v;
    if (typeof v === "number") return v !== 0;
    if (typeof v === "string") {
      const s = v.trim().toLowerCase();
      return s === "true" || s === "1" || s === "yes";
    }
    return false;
  }

  private notify(message: string, level: "success" | "warning" | "error"): void {
    try {
      // eslint-disable-next-line @typescript-eslint/no-unsafe-function-type
      const api = (this._context as unknown as { navigation?: { openAlertDialog?: Function } }).navigation;
      if (api && typeof api.openAlertDialog === "function" && level === "error") {
        api.openAlertDialog({ text: message, title: "Export to Excel" });
        return;
      }
    } catch {
      // ignore
    }

    // Fallback
    if (level === "error") alert(message);
    else console.log(message);
  }
}
