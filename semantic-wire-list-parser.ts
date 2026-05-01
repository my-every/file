/**
 * Semantic Wire List Parser
 * 
 * This module provides robust parsing of wire list sheets into a semantic data model.
 * It handles:
 * - Intro/metadata rows before the wire list
 * - From/To grouping rows
 * - Multiline header cells (e.g., "Cable (W)\nConductor (SC)\nJumper Clip (JC)")
 * - Footer detection (SolarTools ver., Report date)
 * - Semantic column mapping with stable keys
 */

// ============================================================================
// Types
// ============================================================================

/**
 * Semantic wire list row with stable, named fields.
 */
export interface SemanticWireListRow {
  /** Row index in the original sheet (for debugging/tracing) */
  __rowIndex: number;
  /** Unique row ID for table rendering */
  __rowId: string;

  // From section
  /** From Device ID (e.g., "XT0170:1", "AF0500:12") */
  fromDeviceId: string;
  /** Wire type code (SC, JC, or cable number like WC1242) */
  wireType: string;
  /** Wire number (e.g., "0V", "FU0172", "XT05002", "UA42011") */
  wireNo: string;
  /** Wire ID / color code (e.g., "CLIP", "GRN", "WHT", "RED") */
  wireId: string;
  /** Gauge/Size (e.g., "16", "14", "20", "---", "CABLE") */
  gaugeSize: string;
  /** From Location (e.g., "BOP CTRL", "FG&E CTRL,SMT130") */
  fromLocation: string;
  /** From Page/Zone (e.g., "40.A2", "39.B5") */
  fromPageZone: string;

  // To section
  /** To Device ID (e.g., "XT0170:2", "XT0500:2") */
  toDeviceId: string;
  /** To Location (e.g., "BOP CTRL", "GEN CMPNT") */
  toLocation: string;
  /** To Page/Zone (e.g., "40.A2", "39.C5") */
  toPageZone: string;

  /** @deprecated Use fromLocation or toLocation instead */
  location?: string;
}

/**
 * Parser diagnostics for debugging and troubleshooting.
 */
export interface WireListParserDiagnostics {
  /** Index of the From/To grouping row, if found */
  groupingRowIndex: number | null;
  /** Index of the actual header row */
  headerRowIndex: number;
  /** Index where footer starts, if detected */
  footerStartIndex: number | null;
  /** Raw header values as detected */
  rawHeaders: string[];
  /** Normalized header values */
  normalizedHeaders: string[];
  /** Confidence level of header detection */
  confidence: "high" | "medium" | "low";
  /** Column mapping from raw index to semantic key */
  columnMap: Record<number, keyof SemanticWireListRow>;
  /** Any warnings during parsing */
  warnings: string[];
}

/**
 * Complete result of semantic wire list parsing.
 */
export interface SemanticWireListParseResult {
  /** The semantic wire list rows */
  semanticRows: SemanticWireListRow[];
  /** Raw rows (original parsed data) */
  rawRows: Record<string, string | number | boolean | Date | null>[];
  /** Intro/metadata rows before the wire list */
  introRows: (string | number | boolean | Date | null)[][];
  /** Footer rows after the wire list */
  footerRows: (string | number | boolean | Date | null)[][];
  /** Extracted metadata from intro rows */
  metadata: WireListMetadata;
  /** Parser diagnostics */
  diagnostics: WireListParserDiagnostics;
}

/**
 * Metadata extracted from wire list preamble.
 */
export interface WireListMetadata {
  sheetTitle?: string;
  projectNumber?: string;
  projectName?: string;
  revision?: string;
  controlsDE?: string;
  controlsME?: string;
  phoneDe?: string;
  phoneMe?: string;
  [key: string]: string | undefined;
}

// ============================================================================
// Constants
// ============================================================================

/**
 * Expected wire list header patterns for strong matching.
 * These match the canonical header row.
 */
const STRONG_HEADER_PATTERNS = [
  { pattern: /^device\s*id$/i, weight: 3 },
  { pattern: /^wire\s*no\.?$/i, weight: 3 },
  { pattern: /^wire\s*id$/i, weight: 3 },
  { pattern: /^gauge\/?size$/i, weight: 2 },
  { pattern: /^page\/?zone$/i, weight: 2 },
  { pattern: /^location$/i, weight: 2 },
  { pattern: /cable.*conductor.*jumper/i, weight: 3 },
];

/**
 * Minimum score to consider a row as the header.
 */
const MIN_HEADER_SCORE = 8;

/**
 * Footer detection patterns.
 */
const FOOTER_PATTERNS = [
  /^solartools\s+ver\./i,
  /^report\s+date:/i,
  /^\s*$/,
];

/**
 * Grouping row detection (From ... To).
 */
const GROUPING_ROW_PATTERNS = [
  { col0: /^from$/i, col6: /^to$/i },
];

// ============================================================================
// Header Cell Normalization
// ============================================================================

/**
 * Normalize a header cell value.
 * Handles multiline cells, trims whitespace, and normalizes spacing.
 * 
 * @param value - The raw cell value
 * @returns Normalized string
 */
export function normalizeHeaderCell(value: string | number | boolean | Date | null): string {
  if (value === null || value === undefined) {
    return "";
  }

  // Convert to string
  let str = String(value);

  // Replace newlines and multiple spaces with single space
  str = str.replace(/[\r\n]+/g, " ").replace(/\s+/g, " ").trim();

  return str;
}

// ============================================================================
// Header Row Detection
// ============================================================================

/**
 * Score a row as a potential header row.
 * Higher scores indicate stronger matches.
 * 
 * @param row - The row to score
 * @returns The score (0 = not a header, higher = more likely)
 */
export function scoreHeaderCandidate(
  row: (string | number | boolean | Date | null)[]
): number {
  if (!row || row.length < 3) return 0;

  let score = 0;
  const normalizedCells = row.map(normalizeHeaderCell);

  // Check each strong pattern
  for (const { pattern, weight } of STRONG_HEADER_PATTERNS) {
    for (const cell of normalizedCells) {
      if (pattern.test(cell)) {
        score += weight;
        break; // Only count each pattern once
      }
    }
  }

  // Bonus for having multiple non-empty cells (typical header row)
  const nonEmptyCount = normalizedCells.filter(c => c.length > 0).length;
  if (nonEmptyCount >= 7) score += 2;
  else if (nonEmptyCount >= 5) score += 1;

  return score;
}

/**
 * Detect the From/To grouping row.
 * This row appears before the actual header and contains "From" and "To".
 * 
 * @param rows - All rows to search
 * @param maxRows - Maximum rows to search
 * @returns Index of the grouping row, or null if not found
 */
export function detectGroupingRow(
  rows: (string | number | boolean | Date | null)[][],
  maxRows: number = 20
): number | null {
  for (let i = 0; i < Math.min(rows.length, maxRows); i++) {
    const row = rows[i];
    if (!row || row.length < 7) continue;

    const cell0 = normalizeHeaderCell(row[0]).toLowerCase();
    const cell6 = normalizeHeaderCell(row[6]).toLowerCase();

    // Check for "From" in first column and "To" around column 6-7
    if (cell0 === "from" && cell6 === "to") {
      return i;
    }
    // Also check if "To" is in a different position
    if (cell0 === "from") {
      for (let j = 5; j < Math.min(row.length, 9); j++) {
        if (normalizeHeaderCell(row[j]).toLowerCase() === "to") {
          return i;
        }
      }
    }
  }

  return null;
}

/**
 * Detect the actual wire list header row.
 * Uses scoring to find the best candidate.
 * 
 * @param rows - All rows to search
 * @param startFromRow - Row index to start searching from
 * @param maxRows - Maximum rows to search
 * @returns Header detection result
 */
export function detectWireListHeader(
  rows: (string | number | boolean | Date | null)[][],
  startFromRow: number = 0,
  maxRows: number = 25
): { headerRowIndex: number; confidence: "high" | "medium" | "low"; score: number } {
  let bestScore = 0;
  let bestIndex = startFromRow;

  for (let i = startFromRow; i < Math.min(rows.length, maxRows); i++) {
    const row = rows[i];
    const score = scoreHeaderCandidate(row);

    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }

    // If we find a very strong match, stop searching
    if (score >= 15) break;
  }

  const confidence = bestScore >= 12 ? "high" : bestScore >= 8 ? "medium" : "low";

  return { headerRowIndex: bestIndex, confidence, score: bestScore };
}

// ============================================================================
// Footer Detection
// ============================================================================

/**
 * Check if a row is a footer row.
 * 
 * @param row - The row to check
 * @returns True if this is a footer row
 */
export function isFooterRow(row: (string | number | boolean | Date | null)[]): boolean {
  if (!row || row.length === 0) return false;

  // Check if row is mostly empty
  const nonEmptyCount = row.filter(c => c !== null && c !== "").length;
  if (nonEmptyCount === 0) return true;

  // Check first cell against footer patterns
  const firstCell = normalizeHeaderCell(row[0]);
  for (const pattern of FOOTER_PATTERNS) {
    if (pattern.test(firstCell)) {
      return true;
    }
  }

  // Check any cell for SolarTools or Report date
  for (const cell of row) {
    const normalized = normalizeHeaderCell(cell);
    if (/solartools\s+ver\./i.test(normalized) || /report\s+date:/i.test(normalized)) {
      return true;
    }
  }

  return false;
}

/**
 * Find where footer rows start.
 * Scans from the end backwards.
 * 
 * @param rows - All data rows (after header)
 * @returns Index of first footer row, or null if no footer
 */
export function detectFooterStart(
  rows: (string | number | boolean | Date | null)[][]
): number | null {
  if (rows.length === 0) return null;

  // Scan from end backwards
  let footerStart: number | null = null;

  for (let i = rows.length - 1; i >= Math.max(0, rows.length - 10); i--) {
    if (isFooterRow(rows[i])) {
      footerStart = i;
    } else {
      // Stop scanning once we hit a non-footer row
      break;
    }
  }

  return footerStart;
}

/**
 * Trim trailing non-data rows (footer rows) from the data.
 * 
 * @param rows - The data rows
 * @returns Object with trimmed data rows and footer rows
 */
export function trimTrailingNonDataRows(
  rows: (string | number | boolean | Date | null)[][]
): { dataRows: (string | number | boolean | Date | null)[][]; footerRows: (string | number | boolean | Date | null)[][] } {
  const footerStart = detectFooterStart(rows);

  if (footerStart === null) {
    return { dataRows: rows, footerRows: [] };
  }

  return {
    dataRows: rows.slice(0, footerStart),
    footerRows: rows.slice(footerStart),
  };
}

// ============================================================================
// Metadata Extraction
// ============================================================================

/**
 * Extract metadata from intro rows.
 * 
 * @param introRows - Rows before the header
 * @returns Extracted metadata
 */
export function extractWireListMetadata(
  introRows: (string | number | boolean | Date | null)[][]
): WireListMetadata {
  const metadata: WireListMetadata = {};

  // First non-empty row is usually the sheet title
  for (const row of introRows) {
    const firstCell = normalizeHeaderCell(row[0]);
    if (firstCell) {
      metadata.sheetTitle = firstCell;
      break;
    }
  }

  // Scan for key-value pairs
  for (const row of introRows) {
    if (!row || row.length < 2) continue;

    const key = normalizeHeaderCell(row[0]).toLowerCase().replace(/[:#]?\s*$/, "");
    const value = normalizeHeaderCell(row[1]);

    if (!value) continue;

    switch (key) {
      case "project #":
      case "project":
        metadata.projectNumber = value;
        break;
      case "project name":
        metadata.projectName = value;
        break;
      case "revision":
        metadata.revision = value;
        break;
      case "controls de":
        metadata.controlsDE = value;
        break;
      case "controls me":
        metadata.controlsME = value;
        break;
      case "phone":
        // This could be DE or ME phone - context dependent
        break;
    }
  }

  return metadata;
}

// ============================================================================
// Semantic Row Building
// ============================================================================

/**
 * Build the column map from header row.
 * Maps column indices to semantic field names.
 * 
 * Expected header sequence:
 * 0: Device ID (From)
 * 1: Cable (W) Conductor (SC) Jumper Clip (JC)
 * 2: Wire No.
 * 3: Wire ID
 * 4: Gauge/Size
 * 5: Page/Zone (From)
 * 6: Device ID (To)
 * 7: Location
 * 8: Page/Zone (To)
 * 
 * @param headerRow - The header row
 * @returns Column map
 */
export function buildColumnMap(
  headerRow: (string | number | boolean | Date | null)[]
): Record<number, keyof SemanticWireListRow> {
  const map: Record<number, keyof SemanticWireListRow> = {};
  const normalizedHeaders = headerRow.map(normalizeHeaderCell);

  // Track which Device ID and Page/Zone we've seen
  let deviceIdCount = 0;
  let pageZoneCount = 0;

  for (let i = 0; i < normalizedHeaders.length; i++) {
    const header = normalizedHeaders[i].toLowerCase();

    if (/^device\s*id$/i.test(header)) {
      if (deviceIdCount === 0) {
        map[i] = "fromDeviceId";
        deviceIdCount++;
      } else {
        map[i] = "toDeviceId";
        deviceIdCount++;
      }
    } else if (/cable.*conductor.*jumper/i.test(header) || /^type$/i.test(header)) {
      map[i] = "wireType";
    } else if (/^wire\s*no\.?$/i.test(header)) {
      map[i] = "wireNo";
    } else if (/^wire\s*id$/i.test(header)) {
      map[i] = "wireId";
    } else if (/^gauge\/?size$/i.test(header)) {
      map[i] = "gaugeSize";
    } else if (/^page\/?zone$/i.test(header)) {
      if (pageZoneCount === 0) {
        map[i] = "fromPageZone";
        pageZoneCount++;
      } else {
        map[i] = "toPageZone";
        pageZoneCount++;
      }
    } else if (/^location$/i.test(header)) {
      // Location column - typically appears only once in the "To" section
      // Map to toLocation (this is the standard wire list format)
      // If there's a second Location column, map it to fromLocation
      if (!Object.values(map).includes("toLocation")) {
        map[i] = "toLocation";
      } else if (!Object.values(map).includes("fromLocation")) {
        map[i] = "fromLocation";
      }
    }
  }

  // If we couldn't map semantically, fall back to positional mapping
  // This handles sheets that don't have standard headers
  // Standard format: Device ID | Type | Wire No | Wire ID | Gauge | Page/Zone | Device ID | Location | Page/Zone
  // Location is in column 7 (To section)
  if (Object.keys(map).length < 5 && normalizedHeaders.length >= 9) {
    return {
      0: "fromDeviceId",
      1: "wireType",
      2: "wireNo",
      3: "wireId",
      4: "gaugeSize",
      5: "fromPageZone",
      6: "toDeviceId",
      7: "toLocation",  // Location is in the To section
      8: "toPageZone",
    };
  }

  return map;
}

/**
 * Build semantic wire list rows from raw data.
 * 
 * @param dataRows - The raw data rows (after header, before footer)
 * @param columnMap - The column mapping
 * @param headerRowIndex - Original header row index for offset calculation
 * @returns Array of semantic wire list rows
 */
export function buildSemanticWireListRows(
  dataRows: (string | number | boolean | Date | null)[][],
  columnMap: Record<number, keyof SemanticWireListRow>,
  headerRowIndex: number
): SemanticWireListRow[] {
  const semanticRows: SemanticWireListRow[] = [];

  for (let i = 0; i < dataRows.length; i++) {
    const rawRow = dataRows[i];

    // Skip empty rows
    const hasContent = rawRow.some(cell => cell !== null && cell !== "");
    if (!hasContent) continue;

    // Build semantic row
    const semanticRow: SemanticWireListRow = {
      __rowIndex: headerRowIndex + 1 + i,
      __rowId: `row-${i}`,
      fromDeviceId: "",
      wireType: "",
      wireNo: "",
      wireId: "",
      gaugeSize: "",
      fromLocation: "",
      fromPageZone: "",
      toDeviceId: "",
      toLocation: "",
      toPageZone: "",
      // Deprecated - kept for backward compatibility
      location: "",
    };

    // Map columns to semantic fields
    for (const [colIndexStr, fieldName] of Object.entries(columnMap)) {
      const colIndex = parseInt(colIndexStr, 10);
      if (colIndex < rawRow.length) {
        const value = rawRow[colIndex];
        const strValue = normalizeSemanticWireListValue(fieldName, value);
        (semanticRow as Record<string, string | number>)[fieldName] = strValue;
      }
    }

    // Only add rows that have meaningful content (at least a device ID)
    if (semanticRow.fromDeviceId || semanticRow.toDeviceId) {
      semanticRows.push(semanticRow);
    }
  }

  return semanticRows;
}

function normalizeSemanticWireListValue(
  fieldName: keyof SemanticWireListRow,
  value: string | number | boolean | Date | null
): string {
  const strValue = value !== null && value !== undefined ? String(value).trim() : "";

  if (fieldName === "wireNo" && /^\d/.test(strValue)) {
    return `-${strValue}`;
  }

  return strValue;
}

// ============================================================================
// Main Parser
// ============================================================================

/**
 * Parse a wire list sheet into semantic format.
 * 
 * @param rawData - The raw 2D array of sheet data
 * @returns Complete parse result with semantic rows and diagnostics
 */
export function parseSemanticWireList(
  rawData: (string | number | boolean | Date | null)[][]
): SemanticWireListParseResult {
  const warnings: string[] = [];

  if (!rawData || rawData.length === 0) {
    return {
      semanticRows: [],
      rawRows: [],
      introRows: [],
      footerRows: [],
      metadata: {},
      diagnostics: {
        groupingRowIndex: null,
        headerRowIndex: 0,
        footerStartIndex: null,
        rawHeaders: [],
        normalizedHeaders: [],
        confidence: "low",
        columnMap: {},
        warnings: ["Sheet is empty"],
      },
    };
  }

  // Step 1: Detect grouping row (From ... To)
  const groupingRowIndex = detectGroupingRow(rawData);

  // Step 2: Detect header row (start after grouping row if found)
  const searchStart = groupingRowIndex !== null ? groupingRowIndex + 1 : 0;
  const headerDetection = detectWireListHeader(rawData, searchStart);
  const { headerRowIndex, confidence } = headerDetection;

  if (confidence === "low") {
    warnings.push(`Low confidence header detection at row ${headerRowIndex + 1}`);
  }

  // Step 3: Extract intro rows (everything before header)
  const introRows = rawData.slice(0, headerRowIndex);

  // Step 4: Get header row
  const headerRow = rawData[headerRowIndex] || [];
  const rawHeaders = headerRow.map(c => String(c ?? ""));
  const normalizedHeaders = headerRow.map(normalizeHeaderCell);

  // Step 5: Get data rows (after header)
  const afterHeaderRows = rawData.slice(headerRowIndex + 1);

  // Step 6: Trim footer rows
  const { dataRows, footerRows } = trimTrailingNonDataRows(afterHeaderRows);

  // Step 7: Build column map
  const columnMap = buildColumnMap(headerRow);

  // Step 8: Build semantic rows
  const semanticRows = buildSemanticWireListRows(dataRows, columnMap, headerRowIndex);

  // Step 9: Build raw rows (for backwards compatibility)
  const rawRows = dataRows.map((row, idx) => {
    const obj: Record<string, string | number | boolean | Date | null> = {};
    for (let i = 0; i < normalizedHeaders.length; i++) {
      const key = normalizedHeaders[i] || `Column_${i + 1}`;
      obj[key] = row[i] ?? null;
    }
    obj.__rowId = `raw-${idx}`;
    return obj;
  });

  // Step 10: Extract metadata
  const metadata = extractWireListMetadata(introRows);

  // Step 11: Build diagnostics
  const diagnostics: WireListParserDiagnostics = {
    groupingRowIndex,
    headerRowIndex,
    footerStartIndex: footerRows.length > 0 ? headerRowIndex + 1 + dataRows.length : null,
    rawHeaders,
    normalizedHeaders,
    confidence,
    columnMap,
    warnings,
  };

  return {
    semanticRows,
    rawRows,
    introRows,
    footerRows,
    metadata,
    diagnostics,
  };
}

// ============================================================================
// Display Helpers
// ============================================================================

/**
 * Extended semantic column key type including computed fields.
 */
export type SemanticColumnKey = keyof SemanticWireListRow | "estimatedLength" | "estTime";

/**
 * Display column definition for the semantic wire list table.
 */
export interface SemanticDisplayColumn {
  key: SemanticColumnKey;
  header: string;
  group: "from" | "to" | "meta" | "length";
  visible: boolean;
  sortOrder: number;
}

/**
 * Get display columns for the semantic wire list table.
 * 
 * Column order follows the Brand List format:
 * From: Wire No | Gauge/Size | Wire ID | Device ID | Location
 * To: Device ID | Location
 * With Length column between From and To groups
 */
export function getSemanticDisplayColumns(): SemanticDisplayColumn[] {
  // Default columns: From (Checkmark, Device ID), Wire No., Wire ID, Length, 
  // To (Checkmark, Device ID, Location), IPV Checkmark, Comments
  return [
    // From columns
    { key: "fromDeviceId", header: "Device ID", group: "from", visible: true, sortOrder: 1 },
    { key: "wireNo", header: "Wire No.", group: "from", visible: true, sortOrder: 2 },
    { key: "wireId", header: "Wire ID", group: "from", visible: true, sortOrder: 3 },
    { key: "gaugeSize", header: "Gauge/Size", group: "from", visible: true, sortOrder: 4 },
    { key: "fromPartNumber", header: "Part No.", group: "from", visible: true, sortOrder: 5 },
    { key: "fromLocation", header: "Location", group: "from", visible: false, sortOrder: 6 },
    { key: "wireType", header: "Type", group: "from", visible: false, sortOrder: 7 },
    { key: "fromPageZone", header: "Page/Zone", group: "from", visible: false, sortOrder: 8 },
    // Length column - positioned between From and To groups
    { key: "estimatedLength", header: "Length", group: "length", visible: false, sortOrder: 9 },
    // Estimated time column — computed from gauge + section kind
    { key: "estTime", header: "Est. Time", group: "length", visible: false, sortOrder: 9.5 },
    // To columns
    { key: "toDeviceId", header: "Device ID", group: "to", visible: true, sortOrder: 10 },
    { key: "toPartNumber", header: "Part No.", group: "to", visible: true, sortOrder: 11 },
    { key: "toLocation", header: "Location", group: "to", visible: true, sortOrder: 12 },
    { key: "toPageZone", header: "Page/Zone", group: "to", visible: false, sortOrder: 13 },
  ];
}

/**
 * Format a gauge/size value for display.
 */
export function formatGaugeSizeDisplay(value: string): string {
  if (!value || value === "") return "---";
  return value;
}

/**
 * Format a wire type for display.
 * Returns the raw value (SC, JC, WC1242, etc.)
 */
export function formatWireTypeDisplay(value: string): string {
  if (!value || value === "") return "-";
  return value;
}
