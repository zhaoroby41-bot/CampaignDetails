export type UserRole = "appleLead" | "dealer";

export interface Dealer {
  id: string;
  name: string;
  region: string;
  storeCount: number;
}

export interface ActivityMaterial {
  id: string;
  fileName: string;
  fileType: string;
  size: string;
  uploadedAt: string;
  customTag: string;
  dealerIds: string[];
}

export interface PendingFile {
  id: string;
  fileName: string;
  fileType: string;
  size: string;
}
