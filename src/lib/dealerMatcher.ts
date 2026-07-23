import type { Dealer, PendingFile } from "../types";

export type DealerMatchStatus = "exact" | "alias" | "fuzzy" | "multiple" | "unmatched" | "manual" | "manualAll";

export interface DealerMatchRow extends PendingFile {
  dealerId: string;
  dealerName: string;
  status: DealerMatchStatus;
  confidence: number;
  needsReview: boolean;
  candidates: Array<{ dealerId: string; dealerName: string; confidence: number }>;
}

export type DealerAliasMap = Record<string, string>;

function normalizeText(value: string) {
  return value
    .toLowerCase()
    .replace(/\.[a-z0-9]+$/i, "")
    .replace(/有限公司|有限责任公司|科技|电子|通信|通讯|数码|信息|技术|公司/g, "")
    .replace(/[\s_\-—–·.,，。()（）【】\[\]：:]+/g, "");
}

function scorePartialMatch(fileName: string, dealerName: string) {
  const normalizedFile = normalizeText(fileName);
  const normalizedDealer = normalizeText(dealerName);

  if (!normalizedFile || !normalizedDealer) {
    return 0;
  }

  if (normalizedFile.includes(normalizedDealer)) {
    return 86;
  }

  const bigrams = new Set<string>();
  for (let index = 0; index < normalizedDealer.length - 1; index += 1) {
    bigrams.add(normalizedDealer.slice(index, index + 2));
  }

  if (bigrams.size === 0) {
    return 0;
  }

  const hits = Array.from(bigrams).filter((part) => normalizedFile.includes(part)).length;
  return Math.round((hits / bigrams.size) * 86);
}

export function confirmAllDealerMatch(row: DealerMatchRow): DealerMatchRow {
  return {
    ...row,
    dealerId: "",
    dealerName: "全部参与活动经销商",
    status: "manualAll",
    confidence: 100,
    needsReview: false,
    candidates: [],
  };
}

export function matchFilesToDealers(files: PendingFile[], dealers: Dealer[], aliases: DealerAliasMap = {}): DealerMatchRow[] {
  return files.map((file) => {
    const exactDealer = dealers.find((dealer) => file.fileName.includes(dealer.name));

    if (exactDealer) {
      return {
        ...file,
        dealerId: exactDealer.id,
        dealerName: exactDealer.name,
        status: "exact",
        confidence: 100,
        needsReview: false,
        candidates: [{ dealerId: exactDealer.id, dealerName: exactDealer.name, confidence: 100 }],
      };
    }

    const alias = Object.entries(aliases).find(([aliasName]) => normalizeText(file.fileName).includes(normalizeText(aliasName)));
    const aliasDealer = alias ? dealers.find((dealer) => dealer.id === alias[1]) : undefined;

    if (aliasDealer) {
      return {
        ...file,
        dealerId: aliasDealer.id,
        dealerName: aliasDealer.name,
        status: "alias",
        confidence: 94,
        needsReview: false,
        candidates: [{ dealerId: aliasDealer.id, dealerName: aliasDealer.name, confidence: 94 }],
      };
    }

    const candidates = dealers
      .map((dealer) => ({
        dealerId: dealer.id,
        dealerName: dealer.name,
        confidence: scorePartialMatch(file.fileName, dealer.name),
      }))
      .filter((candidate) => candidate.confidence >= 70)
      .sort((left, right) => right.confidence - left.confidence);

    if (candidates.length > 1 && candidates[0].confidence === candidates[1].confidence) {
      return {
        ...file,
        dealerId: "",
        dealerName: "全部参与活动经销商",
        status: "multiple",
        confidence: candidates[0].confidence,
        needsReview: true,
        candidates,
      };
    }

    if (candidates[0]) {
      return {
        ...file,
        dealerId: candidates[0].dealerId,
        dealerName: candidates[0].dealerName,
        status: "fuzzy",
        confidence: candidates[0].confidence,
        needsReview: false,
        candidates,
      };
    }

    return {
      ...file,
      dealerId: "",
      dealerName: "全部参与活动经销商",
      status: "unmatched",
      confidence: 0,
      needsReview: true,
      candidates: [],
    };
  });
}
