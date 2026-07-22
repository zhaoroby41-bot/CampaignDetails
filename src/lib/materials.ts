import type { ActivityMaterial, Dealer, PendingFile, UserRole } from "../types";

export function getVisibleMaterials(materials: ActivityMaterial[], role: UserRole, dealerId: string) {
  if (role === "appleLead") {
    return materials;
  }

  return materials.filter((material) => material.dealerIds.length === 0 || material.dealerIds.includes(dealerId));
}

export function getDealerNames(dealerIds: string[], dealers: Dealer[]) {
  if (dealerIds.length === 0 || dealerIds.length === dealers.length) {
    return "全部参与活动经销商";
  }

  const dealerMap = new Map(dealers.map((dealer) => [dealer.id, dealer.name]));
  return dealerIds.map((dealerId) => dealerMap.get(dealerId)).filter(Boolean).join("、");
}

export function inferFileType(fileName: string) {
  if (!fileName.includes(".")) {
    return "file";
  }

  const extension = fileName.split(".").pop()?.toLowerCase();
  return extension || "file";
}

export function createMaterialRecords(files: PendingFile[], dealerIds: string[], customTag: string): ActivityMaterial[] {
  const timestamp = new Date().toLocaleString("zh-CN", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });

  return files.map((file, index) => ({
    id: `upload-${Date.now()}-${index + 1}`,
    fileName: file.fileName,
    fileType: file.fileType,
    size: file.size,
    uploadedAt: timestamp,
    customTag,
    dealerIds,
  }));
}
