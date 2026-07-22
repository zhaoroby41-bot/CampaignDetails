import { describe, expect, it } from "vitest";
import { getDealerActivityData, sanitizeExcelFileName, splitActivityDataRows } from "./activityData";

const rows = [
  ["VAD/ T1", "POS Apple ID", "POS Name", "2026/07/17", "2026/07/18"],
  ["北京中恒驿站数码信息技术有限公司", "1001", "门店 A", 12, 16],
  ["上海仲璇电子科技有限公司", "2001", "门店 B", 9, 11],
  ["北京中恒驿站数码信息技术有限公司", "1002", "门店 C", 3, 4],
  ["", "3001", "缺失经销商门店", 1, 1],
];

describe("activity data splitting", () => {
  it("splits activity rows by the first VAD/T1 column", () => {
    const result = splitActivityDataRows(rows);

    expect(result.totalRows).toBe(4);
    expect(result.invalidRows).toBe(1);
    expect(result.dealerFiles).toHaveLength(2);
    expect(result.dealerFiles.find((file) => file.dealerName === "北京中恒驿站数码信息技术有限公司")?.rowCount).toBe(2);
  });

  it("keeps the header and dynamic columns in each dealer file", () => {
    const result = splitActivityDataRows(rows);
    const beijing = result.dealerFiles.find((file) => file.dealerName === "北京中恒驿站数码信息技术有限公司");

    expect(beijing?.rows[0]).toEqual(["VAD/ T1", "POS Apple ID", "POS Name", "2026/07/17", "2026/07/18"]);
    expect(beijing?.rows[1]).toEqual(["北京中恒驿站数码信息技术有限公司", "1001", "门店 A", 12, 16]);
    expect(beijing?.rows[2]).toEqual(["北京中恒驿站数码信息技术有限公司", "1002", "门店 C", 3, 4]);
  });

  it("returns only the current dealer activity file", () => {
    const result = splitActivityDataRows(rows);

    expect(getDealerActivityData(result, "上海仲璇电子科技有限公司")?.rowCount).toBe(1);
    expect(getDealerActivityData(result, "不存在的经销商")).toBeUndefined();
  });

  it("sanitizes dealer names for Excel file names", () => {
    expect(sanitizeExcelFileName("A/B:C*D?E")).toBe("A_B_C_D_E");
  });
});
