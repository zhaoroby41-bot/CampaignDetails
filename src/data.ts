import type { ActivityMaterial, Dealer } from "./types";

export const participatingDealers: Dealer[] = [
  { id: "dealer-tibet", name: "西藏酷爱通信有限公司", region: "西区", storeCount: 18 },
  { id: "dealer-shanghai", name: "上海仲璇电子科技有限公司", region: "东区", storeCount: 24 },
  { id: "dealer-shenzhen", name: "深圳市酷果星创数码有限公司", region: "南区", storeCount: 31 },
  { id: "dealer-beijing", name: "北京恒洲盈科贸有限公司", region: "北区", storeCount: 16 },
  { id: "dealer-chengdu", name: "成都新一方数码科技有限公司", region: "西区", storeCount: 12 },
  { id: "dealer-guangzhou", name: "广州苹果优选体验有限公司", region: "南区", storeCount: 22 },
];

export const initialMaterials: ActivityMaterial[] = [
  {
    id: "b6ac2f0d-cdaa-4429-b32b-6ba3ab519834",
    fileName: "A1优质-王者荣耀.pdf",
    fileType: "pdf",
    size: "18.6 MB",
    uploadedAt: "2026-07-20 11:14",
    customTag: "授权文件",
    dealerIds: participatingDealers.map((dealer) => dealer.id),
  },
  {
    id: "7002511a-0548-42e8-89a1-b3491154ceda",
    fileName: "A4台卡-王者荣耀.jpeg",
    fileType: "jpeg",
    size: "7.2 MB",
    uploadedAt: "2026-07-20 11:22",
    customTag: "门店台卡",
    dealerIds: ["dealer-shanghai", "dealer-shenzhen", "dealer-guangzhou"],
  },
  {
    id: "b8dfee58-75df-4b59-a245-458753bd4188",
    fileName: "LED 素材 for 西藏酷爱通信有限公司.zip",
    fileType: "zip",
    size: "426 MB",
    uploadedAt: "2026-07-21 15:40",
    customTag: "LED",
    dealerIds: ["dealer-tibet"],
  },
  {
    id: "78e9ef83-f2ad-4044-8f0d-10e2239256ec",
    fileName: "LED 素材 for 上海仲璇电子科技有限公司.zip",
    fileType: "zip",
    size: "398 MB",
    uploadedAt: "2026-07-21 15:43",
    customTag: "LED",
    dealerIds: ["dealer-shanghai"],
  },
  {
    id: "173cab4f-87ae-4bb8-a2cb-7ba8d4b87a41",
    fileName: "LED 素材 for 深圳市酷果星创数码有限公司.zip",
    fileType: "zip",
    size: "412 MB",
    uploadedAt: "2026-07-21 15:45",
    customTag: "LED",
    dealerIds: ["dealer-shenzhen"],
  },
];
