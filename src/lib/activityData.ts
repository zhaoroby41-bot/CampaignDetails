export type CellValue = string | number | boolean | Date | null | undefined;
export type ActivityDataRow = CellValue[];

export interface DealerActivityFile {
  dealerName: string;
  rowCount: number;
  rows: ActivityDataRow[];
}

export interface ActivityDataSplitResult {
  header: ActivityDataRow;
  totalRows: number;
  invalidRows: number;
  dealerFiles: DealerActivityFile[];
}

export function splitActivityDataRows(rows: ActivityDataRow[]): ActivityDataSplitResult {
  const header = rows[0] ?? [];
  const dealerGroups = new Map<string, ActivityDataRow[]>();
  let invalidRows = 0;

  for (const row of rows.slice(1)) {
    const dealerName = String(row[0] ?? "").trim();

    if (!dealerName) {
      invalidRows += 1;
      continue;
    }

    const currentRows = dealerGroups.get(dealerName) ?? [];
    currentRows.push(row);
    dealerGroups.set(dealerName, currentRows);
  }

  const dealerFiles = Array.from(dealerGroups.entries()).map(([dealerName, dealerRows]) => ({
    dealerName,
    rowCount: dealerRows.length,
    rows: [header, ...dealerRows],
  }));

  return {
    header,
    totalRows: Math.max(rows.length - 1, 0),
    invalidRows,
    dealerFiles,
  };
}

export function getDealerActivityData(result: ActivityDataSplitResult | null, dealerName: string) {
  return result?.dealerFiles.find((file) => file.dealerName === dealerName);
}

export function sanitizeExcelFileName(name: string) {
  return name.replace(/[\\/:*?"<>|]/g, "_").trim() || "activity-data";
}
