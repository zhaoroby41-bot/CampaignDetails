import { describe, expect, it } from "vitest";
import { participatingDealers } from "../data";
import type { ActivityMaterial, PendingFile } from "../types";
import { createMaterialRecords, getDealerNames, getVisibleMaterials, inferFileType } from "./materials";

const materials: ActivityMaterial[] = [
  {
    id: "1",
    fileName: "all.zip",
    fileType: "zip",
    size: "10 MB",
    uploadedAt: "2026-07-22 10:00",
    customTag: "LED",
    dealerIds: [],
  },
  {
    id: "2",
    fileName: "shenzhen.zip",
    fileType: "zip",
    size: "12 MB",
    uploadedAt: "2026-07-22 10:10",
    customTag: "LED",
    dealerIds: ["dealer-shenzhen"],
  },
];

describe("material distribution", () => {
  it("shows all materials to Apple Lead", () => {
    expect(getVisibleMaterials(materials, "appleLead", "dealer-tibet")).toHaveLength(2);
  });

  it("shows all-dealer materials and directly assigned materials to dealer users", () => {
    expect(getVisibleMaterials(materials, "dealer", "dealer-shenzhen").map((material) => material.id)).toEqual(["1", "2"]);
  });

  it("creates one material record per uploaded file with the same dealer assignment", () => {
    const files: PendingFile[] = [
      { id: "file-1", fileName: "LED-1.zip", fileType: "zip", size: "20 MB" },
      { id: "file-2", fileName: "LED-2.jpeg", fileType: "jpeg", size: "8 MB" },
    ];

    const records = createMaterialRecords(files, ["dealer-tibet", "dealer-shanghai"], "LED");

    expect(records).toHaveLength(2);
    expect(records.every((record) => record.dealerIds.join(",") === "dealer-tibet,dealer-shanghai")).toBe(true);
  });

  it("maps empty or full dealer assignment to all participating dealers", () => {
    expect(getDealerNames([], participatingDealers)).toBe("全部参与活动经销商");
    expect(getDealerNames(participatingDealers.map((dealer) => dealer.id), participatingDealers)).toBe("全部参与活动经销商");
  });

  it("maps selected dealer ids to readable names", () => {
    expect(getDealerNames(["dealer-tibet"], participatingDealers)).toBe("西藏酷爱通信有限公司");
  });

  it("infers file type from file names", () => {
    expect(inferFileType("poster.jpeg")).toBe("jpeg");
    expect(inferFileType("briefing")).toBe("file");
  });
});
