import { describe, expect, it } from "vitest";
import type { Dealer, PendingFile } from "../types";
import { confirmAllDealerMatch, matchFilesToDealers } from "./dealerMatcher";

const dealers: Dealer[] = [
  { id: "dealer-shanghai", name: "上海仲璇电子科技有限公司", region: "东区", storeCount: 24 },
  { id: "dealer-shenzhen", name: "深圳市酷果星创数码有限公司", region: "南区", storeCount: 31 },
  { id: "dealer-tibet", name: "西藏酷爱通信有限公司", region: "西区", storeCount: 18 },
];

const files: PendingFile[] = [
  { id: "1", fileName: "LED 素材 for 上海仲璇电子科技有限公司.zip", fileType: "zip", size: "20 MB" },
  { id: "2", fileName: "LED-深圳酷果.zip", fileType: "zip", size: "18 MB" },
  { id: "3", fileName: "LED-上海仲璇.zip", fileType: "zip", size: "19 MB" },
  { id: "4", fileName: "LED 素材包.zip", fileType: "zip", size: "16 MB" },
];

describe("dealer file name matching", () => {
  it("exactly matches files containing the full dealer name", () => {
    const rows = matchFilesToDealers([files[0]], dealers);

    expect(rows[0]).toMatchObject({
      dealerId: "dealer-shanghai",
      status: "exact",
      confidence: 100,
      needsReview: false,
    });
  });

  it("matches known aliases", () => {
    const rows = matchFilesToDealers([files[1]], dealers, { "深圳酷果": "dealer-shenzhen" });

    expect(rows[0]).toMatchObject({
      dealerId: "dealer-shenzhen",
      status: "alias",
      needsReview: false,
    });
    expect(rows[0].confidence).toBeGreaterThanOrEqual(90);
  });

  it("marks partial matches as fuzzy", () => {
    const rows = matchFilesToDealers([files[2]], dealers);

    expect(rows[0]).toMatchObject({
      dealerId: "dealer-shanghai",
      status: "fuzzy",
      needsReview: false,
    });
    expect(rows[0].confidence).toBeGreaterThanOrEqual(70);
    expect(rows[0].confidence).toBeLessThan(90);
  });

  it("defaults unmatched files to all dealers and requires review", () => {
    const rows = matchFilesToDealers([files[3]], dealers);

    expect(rows[0]).toMatchObject({
      dealerId: "",
      dealerName: "全部参与活动经销商",
      status: "unmatched",
      confidence: 0,
      needsReview: true,
    });
  });

  it("confirms an unmatched file as all dealers", () => {
    const rows = matchFilesToDealers([files[3]], dealers);
    const confirmed = confirmAllDealerMatch(rows[0]);

    expect(confirmed).toMatchObject({
      dealerId: "",
      dealerName: "全部参与活动经销商",
      status: "manualAll",
      needsReview: false,
    });
  });
});
