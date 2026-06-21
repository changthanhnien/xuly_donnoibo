import React, { useState, useCallback, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
// @ts-ignore
import pdfWorker from 'pdfjs-dist/build/pdf.worker.mjs?url';
import { PDFDocument, rgb } from 'pdf-lib';
import ExcelJS from 'exceljs';
import { 
  FileSpreadsheet, 
  FileText, 
  CheckCircle2, 
  AlertCircle, 
  Download, 
  RefreshCcw, 
  LayoutList,
  AlertTriangle,
  Loader2,
  Table as TableIcon,
  HelpCircle,
  Search,
  Eye,
  EyeOff,
  ArrowRight,
  ArrowUp,
  ArrowDown,
  Layers,
  MapPin,
  SlidersHorizontal,
  ShoppingBag,
  Boxes
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

import { OrderInfo, PDFPageInfo, MatchResult, ProcessingStats, ExtractedOrderItem, ExtractedOrder } from './types';
import SkuConfig from './components/SkuConfig';

// Set up PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

function getNormalizedSkuKeys(sku: string): string[] {
  const clean = sku.trim().toUpperCase();
  const candidates: string[] = [clean];
  
  // Remove trailing -CD or CD
  let withoutCD = clean;
  if (clean.endsWith('-CD')) {
    withoutCD = clean.slice(0, -3);
    candidates.push(withoutCD);
  } else if (clean.endsWith('CD')) {
    withoutCD = clean.slice(0, -2);
    candidates.push(withoutCD);
  }
  
  // Check for suffix like -[number] on withoutCD
  const parts = withoutCD.split('-');
  if (parts.length > 1) {
    const lastPart = parts[parts.length - 1];
    if (/^\d+$/.test(lastPart)) {
      const base = parts.slice(0, -1).join('-');
      candidates.push(base);
    }
  }
  
  if (parts[0] && parts[0] !== clean && !candidates.includes(parts[0])) {
    candidates.push(parts[0]);
  }
  
  return candidates;
}

function findSkuName(sku: string, masterSkus: Record<string, string>): string | null {
  const candidates = getNormalizedSkuKeys(sku);
  
  // 1. Try candidates directly
  for (const cand of candidates) {
    if (masterSkus[cand]) {
      return masterSkus[cand];
    }
  }
  
  // 2. Try reverse matching: if any candidate matches the base SKU of a key in masterSkus
  const baseOfSku = candidates[candidates.length - 1];
  
  for (const masterKey of Object.keys(masterSkus)) {
    const mCandidates = getNormalizedSkuKeys(masterKey);
    const baseOfMaster = mCandidates[mCandidates.length - 1];
    if (baseOfMaster === baseOfSku) {
      return masterSkus[masterKey];
    }
  }
  
  return null;
}

function getComboMultiplier(sku: string): number {
  const clean = sku.trim().toUpperCase();
  const parts = clean.split('-');
  if (parts.length <= 1) {
    return 1;
  }
  
  const firstAfterDash = parts[1].trim();
  if (/^\d+/.test(firstAfterDash)) {
    const numMatch = firstAfterDash.match(/^\d+/);
    if (numMatch) {
      return parseInt(numMatch[0], 10);
    }
  }
  
  return 1;
}

function isLikelySku(token: string, orderId: string, masterSkus: Record<string, string>): boolean {
  const clean = token.toUpperCase().trim();
  
  // Exclude length extremes
  if (clean.length < 4 || clean.length > 50) return false;
  
  // Exclude if it is the Order ID of this page
  if (orderId && clean.includes(orderId.toUpperCase())) return false;
  
  // Exclude common noise or descriptive terms
  const badKeywords = [
    "PRODUCT", "NAME", "SKU", "QTY", "TOTAL", "SẢN", "PHẨM", "SỐ", "LƯỢNG", "ĐƠN", "HÀNG", "ĐƠN HÀNG", 
    "MÃ ĐƠN", "MÃ SKU", "THẠCH", "COMBO", "BOX", "GHI CHÚ", "NOTE", "LH", "L.H", "I.B.O", "IBO", "HƯƠNG", "HUONG", "VỊ", "VI", "PAGE", "TOTAL", "SENDER", "RECEIVER", "SHIPPED", "DELIVERED", "STATUS", "INVOICE", "ORDER", "SELLER", "CUSTOMER", "RECIPIENT", "PHONE", "DATE", "PRICE"
  ];
  if (badKeywords.includes(clean)) return false;

  // Exclude units of volume and mass (e.g. 200ML, 500G, 1KG, 2L)
  if (/^\d+(ML|G|KG|L|CHAI|HŨ|LON|GÓI)$/i.test(clean)) return false;

  // Exclude combo descriptors (e.g. COMBO1, COMBO3)
  if (/^COMBO\d+$/i.test(clean)) return false;

  // Exclude if it's purely non-alphanumeric or starts/ends with illegal symbols
  // E.g. purely letters with accents ("HƯƠNG")
  const isPureLetters = /^[A-ZẮẰẲẴẶÁÀẢÃẠÂẤẦẨẪẬĂEẾỀỂỄỆÉÈẺẼẸÊIÍÌỈĨỊOỐỒỔỖỘỚỜỞỠỢÓÒỎÕỌÔUỨỪỬỮỰÚÙỦŨỤƯYÝỲỶỸỴ\s\u0300-\u036f]+$/i.test(clean);
  if (isPureLetters) return false;

  // Exclude phone number formats
  if (/^0\d{8,11}$/.test(clean) || /^\+?84\d{8,11}$/.test(clean)) return false;

  // Exclude dates
  if (/^\d{2,4}[\-\/\.]\d{2}[\-\/\.]\d{2,4}$/.test(clean)) return false;

  // Compound SKU check: "SKU1+SKU2+SKU3"
  if (clean.includes('+')) {
    const parts = clean.split('+');
    return parts.every(part => {
      const cleanPart = part.trim().replace(/-\d+$/, '').replace(/-CD$/, '');
      return cleanPart.length >= 4 && !/^[A-Z\s]+$/i.test(cleanPart);
    });
  }

  // Check explicit lookup in masterSkus
  if (findSkuName(clean, masterSkus)) {
    return true;
  }

  // Fallback to checking structured SKU pattern: digits (4-18) optionally with numbers/CD
  const dashSegments = clean.split('-');
  if (dashSegments.length > 1) {
    const base = dashSegments[0];
    if (base.length >= 4 && !/^[A-Z\s]+$/i.test(base)) {
      return true;
    }
  }

  // Pure digits of length >= 4
  if (/^\d{4,20}$/.test(clean)) {
    return true;
  }

  // Alphanumeric with numbers and letters
  if (/^[A-Z0-9\-]{5,30}$/.test(clean)) {
    if (/\d/.test(clean)) {
      return true;
    }
  }

  return false;
}

function parseQuantity(line: string, lines: string[], lineIdx: number): number {
  let match = line.match(/(Qty|SL|Số\s*lượng)[:\s]+([1-9]\d*)\b/i);
  if (match) return parseInt(match[2]);

  match = line.match(/\b([1-9]\d*)\s*(cái|hộp|gói|chai|lon|pcs|SL|Qty|cặp)/i);
  if (match) return parseInt(match[1]);

  match = line.match(/[\s(x\*Xx]([1-9]\d*)\b/);
  if (match) return parseInt(match[1]);
              
  const endMatch = line.match(/\s+([1-9]\d*)$/);
  if (endMatch) {
    const val = parseInt(endMatch[1]);
    if (val < 100) return val;
  }
  
  for (let offset = -1; offset <= 1; offset += 2) {
    const adjIdx = lineIdx + offset;
    if (lines[adjIdx]) {
      const adjLine = lines[adjIdx];
      const adjMatch = adjLine.match(/(Qty|SL|Số\s*lượng)[:\s]+([1-9]\d*)\b/i);
      if (adjMatch) return parseInt(adjMatch[2]);
      
      const adjMatchX = adjLine.match(/\bx\s*([1-9]\d*)\b/i);
      if (adjMatchX) return parseInt(adjMatchX[1]);
    }
  }
  
  return 1;
}

export default function App() {
  // Master SKU State
  const [masterSkus, setMasterSkus] = useState<Record<string, string>>(() => {
    try {
      const saved = localStorage.getItem('MASTER_SKU_DB');
      if (saved) return JSON.parse(saved);
    } catch (e) {
      console.error(e);
    }
    return {};
  });

  const [masterSkusMeta, setMasterSkusMeta] = useState<{
    totalCount: number;
    lastUpdated: string;
    isMasterMarked: boolean;
  } | null>(() => {
    try {
      const saved = localStorage.getItem('MASTER_SKU_META');
      if (saved) return JSON.parse(saved);
    } catch (e) {}
    return null;
  });

  // Files
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [pdfFile, setPdfFile] = useState<File | null>(null);
  
  // Data
  const [excelOrders, setExcelOrders] = useState<OrderInfo[]>([]);
  const [pdfPages, setPdfPages] = useState<PDFPageInfo[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [statusMessage, setStatusMessage] = useState<string>('');

  // States
  const [activeTab, setActiveTab] = useState<'unified' | 'sku_config'>('unified');
  const [step, setStep] = useState(1); // 1: Upload, 2: Preview/Processing

  // Excel Refiner Files/State
  const [rawExcelFile, setRawExcelFile] = useState<File | null>(null);
  const [excelMapping, setExcelMapping] = useState<Record<string, string>>({});
  const [rawExcelData, setRawExcelData] = useState<any[]>([]);

  // Unified PDF-to-Excel/Sort State
  const [unifiedPdfFile, setUnifiedPdfFile] = useState<File | null>(null);
  const [extractedOrders, setExtractedOrders] = useState<ExtractedOrder[]>([]);
  const [sortOption, setSortOption] = useState<'default' | 'province' | 'sku' | 'qty'>('default');
  const [provincePriority, setProvincePriority] = useState<string[]>(["Hồ Chí Minh", "Hà Nội", "Đà Nẵng", "Bình Dương", "Đồng Nai", "Khác"]);
  const [isProcessingUnified, setIsProcessingUnified] = useState(false);
  const [unifiedProgress, setUnifiedProgress] = useState(0);
  const [unifiedStatus, setUnifiedStatus] = useState('');
  const [stepUnified, setStepUnified] = useState(1); // 1: Upload, 2: Preview/Export
  const maskProductInfo = true;
  const maskHeightPercent = 30.5;

  const excelRefinerInputRef = useRef<HTMLInputElement>(null);
  const excelInputRef = useRef<HTMLInputElement>(null);
  const pdfInputRef = useRef<HTMLInputElement>(null);
  const unifiedPdfInputRef = useRef<HTMLInputElement>(null);

  const resetAll = () => {
    setExcelFile(null);
    setPdfFile(null);
    setExcelOrders([]);
    setPdfPages([]);
    setIsProcessing(false);
    setProgress(0);
    setError(null);
    setStep(1);
    setStatusMessage('');
    setRawExcelFile(null);
    setExcelMapping({});
    setRawExcelData([]);
    
    // Reset Unified
    setUnifiedPdfFile(null);
    setExtractedOrders([]);
    setSortOption('province');
    setIsProcessingUnified(false);
    setUnifiedProgress(0);
    setUnifiedStatus('');
    setStepUnified(1);
    
    if (excelInputRef.current) excelInputRef.current.value = '';
    if (pdfInputRef.current) pdfInputRef.current.value = '';
    if (excelRefinerInputRef.current) excelRefinerInputRef.current.value = '';
    if (unifiedPdfInputRef.current) unifiedPdfInputRef.current.value = '';
  };

  // ============================================
  // UNIFIED WORKFLOW: 1-CLICK PDF-TO-EXCEL & SORT
  // ============================================

  const extractDataFromPdfFile = async (
    file: File,
    onProgress: (progress: number, message: string) => void
  ): Promise<ExtractedOrder[]> => {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const totalPages = pdf.numPages;
    const orders: ExtractedOrder[] = [];

    for (let i = 1; i <= totalPages; i++) {
      onProgress(Math.round((i / totalPages) * 100), `Đang trích xuất: trang ${i}/${totalPages}...`);

      const page = await pdf.getPage(i);
      const textContent = await page.getTextContent();
      const textItems = (textContent.items as any[]).filter(item => item.transform);

      // Reconstruct logical text lines using Y-coordinates with 4px tolerance
      const lineMap: Record<number, any[]> = {};
      textItems.forEach(item => {
        const y = Math.round(item.transform[5]);
        let foundY = Object.keys(lineMap).find(key => Math.abs(Number(key) - y) < 4);
        if (foundY) {
          lineMap[Number(foundY)].push(item);
        } else {
          lineMap[y] = [item];
        }
      });

      const sortedYKeys = Object.keys(lineMap).map(Number).sort((a, b) => b - a);
      const lines = sortedYKeys.map(y => {
        const sortedItems = lineMap[y].sort((a, b) => a.transform[4] - b.transform[4]);
        return sortedItems.map(item => item.str).join(' ');
      });

      const fullText = lines.join('\n');

      // 1. Parse Order ID
      let orderId = "";
      // Match "Order ID: 123" or similar keywords
      const orderIdRegex = /(Order\s*ID|Mã\s*đơn\s*hàng|Mã\s*đơn|OrderID)[:\s]+([A-Za-z0-9_\-]+)/i;
      const labelMatch = fullText.match(orderIdRegex);
      if (labelMatch) {
         orderId = labelMatch[2].trim();
      } else {
         // Standalone 10-30 digit fallback
         const standaloneMatch = fullText.match(/\b\d{10,32}\b/);
         if (standaloneMatch) {
           orderId = standaloneMatch[0];
         } else {
           orderId = `PAGE-${i}`;
         }
      }
      orderId = orderId.replace(/['"\s]/g, '');

      // 2. Parse Sender Province
      const provinceConfigs = [
        {
          name: "Hồ Chí Minh",
          phrases: [
            "hồ chí minh", "tp.hcm", "tp. hcm", "sài gòn", "sai gon", "thủ đức", "thu duc", 
            "tân bình", "tan binh", "bình thạnh", "binh thanh", "gò vấp", "go vap", 
            "phú nhuận", "phu nhuan", "bình tân", "binh tan", "tân phú", "tan phu", 
            "củ chi", "cu chi", "hóc môn", "hoc mon", "bình chánh", "binh chanh", 
            "nhà bè", "nha be", "cần giờ", "can gio", "hồ chí", "ho chi", "tp hcm", "tphcm"
          ],
          compressed: [
            "hồchíminh", "tphcm", "sàigòn", "saigon", "thủđức", "thuduc", 
            "tânbình", "tanbinh", "bìnhthạnh", "binhthanh", "gòvấp", "govap", 
            "phúnhuận", "phunhuan", "bìnhtân", "binhtan", "tânphú", "tanphu", 
            "củchi", "cuchi", "hócmôn", "hocmon", "bìnhchánh", "binhchanh", 
            "nhàbè", "nhabe", "cầngiờ", "cangio", "hồchí", "hochi"
          ],
          shorthands: ["hồ", "ho", "hcm", "sg"]
        },
        {
          name: "Hà Nội",
          phrases: [
            "hà nội", "ha noi", "hoàn kiếm", "hoan kiem", "ba đình", "ba dinh", 
            "hai bà trưng", "hai ba trung", "đống đa", "dong da", "cầu giấy", "cau giay", 
            "thanh xuân", "thanh xuan", "tây hồ", "tay ho", "long biên", "long bien", 
            "hà đông", "ha dong", "nam từ liêm", "nam tu liem", "bắc từ liêm", "bac tu liem", 
            "hoàng mai", "hoang mai", "thanh trì", "thanh tri", "gia lâm", "gia lam", 
            "đông anh", "dong anh", "sóc sơn", "soc son"
          ],
          compressed: [
            "hànội", "hanoi", "hoànkiếm", "hoankiem", "bađình", "badinh", 
            "haibàtrưng", "haibatrung", "đốngđa", "dongda", "cầuấy", "caugiay", 
            "thanhxuân", "thanhxuan", "tâyhồ", "tayho", "longbiên", "longbien", 
            "hàđông", "hadong", "namtừliêm", "namtuliem", "bắctừliêm", "bactuliem", 
            "hoàngmai", "hoangmai", "thanhtrì", "thanhtri", "gialâm", "gialam", 
            "đônganh", "donganh", "sócsơn", "socson"
          ],
          shorthands: ["hà", "ha", "hn"]
        },
        {
          name: "Đà Nẵng",
          phrases: [
            "đà nẵng", "da nang", "ngũ hành sơn", "ngu hanh son", "liên chiểu", "lien chieu", 
            "cẩm lệ", "cam le", "hòa vang", "hoa vang", "hải châu", "hai chau", 
            "thanh khê", "thanh khe", "sơn trà", "son tra", "ngũhành"
          ],
          compressed: [
            "đànẵng", "danang", "ngũhànhsơn", "nguhanhson", "liênchiểu", "lienchieu", 
            "cẩmlệ", "camle", "hoàvang", "hoavang", "hảichâu", "haichau", 
            "thanhkhê", "thanhkhe", "sơntrà", "sontra"
          ],
          shorthands: ["đà", "da", "dn"]
        },
        {
          name: "Bình Dương",
          phrases: [
            "bình dương", "binh duong", "thủ dầu một", "thu dau mot", "thuận an", "thuan an", 
            "dĩ an", "di an", "bến cát", "ben cat", "tân uyên", "tan uyen"
          ],
          compressed: [
            "bìnhdương", "binhduong", "thủdầumột", "thudaumot", "thuậnan", "thuanan", 
            "dĩan", "dian", "bếncát", "bencat", "tânuyên", "tanuyen"
          ],
          shorthands: ["bd", "bìnhdương", "binhduong"]
        },
        {
          name: "Đồng Nai",
          phrases: [
            "đồng nai", "dong nai", "biên hòa", "bien hoa", "long thành", "long thanh", 
            "nhơn trạch", "nhon trach"
          ],
          compressed: [
            "đồngnai", "dongnai", "biênhòa", "bienhoa", "longthành", "longthanh", 
            "nhơntrạch", "nhontrach"
          ],
          shorthands: ["đn", "đồngnai", "dongnai"]
        }
      ];

      // Helper to match whole words safely with Vietnamese Unicode characters
      const matchWholeWordVN = (text: string, keyword: string): boolean => {
        const vnCharClass = "[a-z0-9àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ_]";
        const regex = new RegExp(`(?:^|[^${vnCharClass}])${keyword.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&')}(?:$|[^${vnCharClass}])`, 'i');
        return regex.test(text);
      };

      const detectProvince = (zoneText: string): string | null => {
        const normZone = zoneText.normalize('NFC').toLowerCase().replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ');
        // Remove spaces, newlines and common punctuation completely to test compressed strings
        const compZone = zoneText.normalize('NFC').toLowerCase().replace(/[^a-z0-9àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]/g, '');

        // Tier A: Check exact phrases in normalized text
        for (const config of provinceConfigs) {
          for (const phrase of config.phrases) {
            if (normZone.includes(phrase)) {
              return config.name;
            }
          }
        }

        // Tier B: Check exact compressed phrases in compressed text
        for (const config of provinceConfigs) {
          for (const comp of config.compressed) {
            if (compZone.includes(comp)) {
              return config.name;
            }
          }
        }

        // Tier C: Check shorthands as whole words in normalized text
        for (const config of provinceConfigs) {
          for (const sh of config.shorthands) {
            if (matchWholeWordVN(normZone, sh)) {
              return config.name;
            }
          }
        }

        return null;
      };

      // Helper function to check for the custom prefix matches
      const checkSpecialSender = (block: string): string | null => {
        if (!block) return null;
        
        // Normalize the text (lowercase, collapse multiple spaces, remove newlines)
        const normalized = block.normalize('NFC').toLowerCase().replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
        // Also create a version with absolutely NO spaces to catch PDF extraction artifacts (e.g. "C 2  L ô  2 0")
        const stripped = normalized.replace(/\s+/g, '');
        
        const rules = [
          { keyword: "32 trần thị báo", strippedKeyword: "32trầnthịbáo", province: "Hồ Chí Minh" },
          { keyword: "phú thạnh, hồ chí minh", strippedKeyword: "phúthạnh,hồchíminh", province: "Hồ Chí Minh" },
          { keyword: "131 đường nguyễn văn tạo", strippedKeyword: "131đườngnguyễnvăntạo", province: "Đà Nẵng" },
          { keyword: "131 nguyễn văn tạo", strippedKeyword: "131nguyễnvăntạo", province: "Đà Nẵng" },
          { keyword: "an khê, đà nẵng", strippedKeyword: "ankhê,đànẵng", province: "Đà Nẵng" },
          { keyword: "c2 lô 20", strippedKeyword: "c2lô20", province: "Hà Nội" },
          { keyword: "c2 lo 20", strippedKeyword: "c2lo20", province: "Hà Nội" },
          { keyword: "đt mới định công", strippedKeyword: "đtmớiđịnhcông", province: "Hà Nội" },
          { keyword: "định công, hoàng mai", strippedKeyword: "địnhcông,hoàngmai", province: "Hà Nội" }
        ];
        
        for (const rule of rules) {
          if (normalized.includes(rule.keyword) || stripped.includes(rule.strippedKeyword)) {
            return rule.province;
          }
        }
        
        return null;
      };

      let senderProvince = "Khác";
      let senderBlock = "";
      const senderIdx = fullText.search(/Người\s*gửi|Sender|Gửi\s*từ|From/i);
      const recipientIdx = fullText.search(/Người\s*nhận|Recipient|To:|Ship\s*to/i);
      if (senderIdx !== -1) {
        if (recipientIdx !== -1 && recipientIdx > senderIdx) {
          senderBlock = fullText.substring(senderIdx, recipientIdx);
        } else {
          senderBlock = fullText.substring(senderIdx, senderIdx + 400);
        }
      }

      // Run robust tiered detection with priority on user's custom rules
      let detectedProvince = checkSpecialSender(senderBlock) || checkSpecialSender(fullText);
      
      if (!detectedProvince && senderBlock) {
        detectedProvince = detectProvince(senderBlock);
      }
      if (!detectedProvince) {
        detectedProvince = detectProvince(fullText);
      }
      senderProvince = detectedProvince || "Khác";

      // 3. Parse Item List (Product Name, SKU, Qty) using high-precision columnar bounding-box analysis
      const extractedItems: ExtractedOrderItem[] = [];

      // Detect table coordinates using header labels
      const nameHeader = textItems.find(item => /Product\s*Name|Tên\s*sản\s*phẩm/i.test(item.str));
      const skuHeader = textItems.find(item => /\bSKU\b/i.test(item.str) && !/Seller/i.test(item.str) || /Phân\s*loại/i.test(item.str));
      const sellerSkuHeader = textItems.find(item => /Seller\s*SKU|Mã\s*SKU|Mã\s*phân\s*loại/i.test(item.str) && !item.str.includes("Seller SKU") && item !== skuHeader ? true : /Seller\s*SKU|Mã\s*SKU/i.test(item.str));
      const qtyHeader = textItems.find(item => /\bQty\b|Số\s*lượng/i.test(item.str));

      let colNameX = 30;
      let colSkuX = 230;
      let colSellerSkuX = 345;
      let colQtyX = 495;

      if (nameHeader) colNameX = nameHeader.transform[4];
      if (skuHeader) colSkuX = skuHeader.transform[4];
      if (sellerSkuHeader) colSellerSkuX = sellerSkuHeader.transform[4];
      if (qtyHeader) colQtyX = qtyHeader.transform[4];

      // Sanity checks on relative order coordinates
      if (colNameX < 0 || colSkuX <= colNameX || colSellerSkuX <= colSkuX || colQtyX <= colSellerSkuX) {
        colNameX = 30;
        colSkuX = 230;
        colSellerSkuX = 345;
        colQtyX = 495;
      }

      // Determine table vertical bounds to isolate rows
      let headerY = -1;
      if (sellerSkuHeader) headerY = sellerSkuHeader.transform[5];
      else if (nameHeader) headerY = nameHeader.transform[5];
      else if (skuHeader) headerY = skuHeader.transform[5];
      else if (qtyHeader) headerY = qtyHeader.transform[5];

      if (headerY === -1) {
        const fallbackHeaderItem = textItems.find(item => /Product\s*Name|Seller\s*SKU|Tên\s*sản\s*phẩm|Mã\s*SKU|Mã\s*phân\s*loại/i.test(item.str));
        if (fallbackHeaderItem) {
          headerY = fallbackHeaderItem.transform[5];
        }
      }

      let footerY = 0;
      // Search for items below headerY that contain keywords pointing to a footer or out-of-table content
      const lowerFooters = textItems.filter(item => {
        const y = item.transform[5];
        if (headerY !== -1 && y >= headerY - 5) return false;
        const s = item.str.toUpperCase().trim();
        return (
          s.includes("ORDER ID") ||
          s.includes("ORDERID") ||
          s.includes("MÃ ĐƠN") ||
          s.includes("TIKTOK SHOP") ||
          s.includes("PAGE") ||
          s.includes("TRANG") ||
          s.includes("TOTAL QTY") ||
          s.includes("QTY TOTAL") ||
          s.includes("TỔNG SỐ LƯỢNG") ||
          s.includes("TỔNG CỘNG") ||
          s.includes("THỜI GIAN") ||
          s.includes("NGÀY") ||
          s.includes("IN TRANSIT") ||
          (orderId && s.includes(orderId.toUpperCase()))
        );
      });

      if (lowerFooters.length > 0) {
        const footerYCoords = lowerFooters.map(item => item.transform[5]);
        footerY = Math.max(...footerYCoords);
      } else {
        const footerItem = textItems.find(item => 
          /Qty\s*Total|Total\s*Qty|Tổng\s*số\s*lượng|Tổng\s*cộng|TikTok\s*Shop/i.test(item.str) && 
          (headerY !== -1 ? item.transform[5] < headerY : true)
        );
        if (footerItem) {
          footerY = footerItem.transform[5];
        }
      }

      // Filter and isolate elements within the table's graphical bounds
      const tableItems = textItems.filter(item => {
        const y = item.transform[5];
        if (headerY !== -1 && y >= headerY - 3) return false; // Exclude header and text above
        if (footerY !== 0 && y <= footerY + 3) return false;   // Exclude footer and text below
        
        const x = item.transform[4];
        if (x < colNameX - 25 || x > colQtyX + 80) return false; // Exclude side margin texts (vertical barcode labels, etc.)
        
        // Strictly exclude items that themselves contain general out-of-table strings
        const s = item.str.toUpperCase().trim();
        if (
          s.includes("ORDER ID") || 
          s.includes("ORDERID") || 
          s.includes("TIKTOK SHOP") || 
          s.includes("MÃ ĐƠN") ||
          s.includes("PAGE") ||
          s.includes("TRANG") ||
          s.includes("IN TRANSIT") ||
          s.includes("TỔNG SỐ LƯỢNG") ||
          s.includes("TỔNG CỘNG") ||
          s.includes("QTY TOTAL") ||
          s.includes("TOTAL QTY") ||
          s.includes("THỜI GIAN") ||
          s.includes("ĐƠN HÀNG LÀ") ||
          s.includes("J&T EXPRESS")
        ) {
          return false;
        }
        if (orderId && s.includes(orderId.toUpperCase())) {
          return false;
        }
        
        return true;
      });

      // Group table elements into horizontal row slices
      const yMap: Record<number, any[]> = {};
      tableItems.forEach(item => {
        const y = Math.round(item.transform[5]);
        const foundY = Object.keys(yMap).find(key => Math.abs(Number(key) - y) < 4);
        if (foundY) {
          yMap[Number(foundY)].push(item);
        } else {
          yMap[y] = [item];
        }
      });

      const sortedY = Object.keys(yMap).map(Number).sort((a, b) => b - a);

      /*********************************************************
       * Determine column index of a bounding-box element's horizontal start
       *********************************************************/
      function getColumnIndex(x: number): number {
        if (x >= colNameX - 25 && x < colSkuX - 12) return 0;       // Product Name
        if (x >= colSkuX - 12 && x < colSellerSkuX - 12) return 1;  // SKU / Attributes
        if (x >= colSellerSkuX - 12 && x < colQtyX - 12) return 2;  // Seller SKU
        if (x >= colQtyX - 12 && x < colQtyX + 75) return 3;        // Qty
        return -1;
      }

      interface ExtractedTableRow {
        nameParts: string[];
        skuParts: string[];
        sellerSkuParts: string[];
        qtyParts: string[];
      }

      const rows: ExtractedTableRow[] = [];
      let currentIdx = -1;
      let lastLineY = -1;

      for (const y of sortedY) {
        const lineItems = yMap[y];
        const col0Items = lineItems.filter(i => getColumnIndex(i.transform[4]) === 0).sort((a, b) => a.transform[4] - b.transform[4]);
        const col1Items = lineItems.filter(i => getColumnIndex(i.transform[4]) === 1).sort((a, b) => a.transform[4] - b.transform[4]);
        const col2Items = lineItems.filter(i => getColumnIndex(i.transform[4]) === 2).sort((a, b) => a.transform[4] - b.transform[4]);
        const col3Items = lineItems.filter(i => getColumnIndex(i.transform[4]) === 3).sort((a, b) => a.transform[4] - b.transform[4]);

        const str0 = col0Items.map(i => i.str).join(' ').trim();
        const str1 = col1Items.map(i => i.str).join(' ').trim();
        const str2 = col2Items.map(i => i.str).join('').trim(); // Dense join for fragmented SKUs
        const str3 = col3Items.map(i => i.str).join(' ').trim();

        // A new row always begins if there is a positive digit present in the Qty column
        const startsNewRow = str3 !== "" && /\d+/.test(str3);
        const yDiff = lastLineY !== -1 ? Math.abs(lastLineY - y) : 0;

        if (startsNewRow || currentIdx === -1 || yDiff > 28) {
          rows.push({
            nameParts: str0 ? [str0] : [],
            skuParts: str1 ? [str1] : [],
            sellerSkuParts: str2 ? [str2] : [],
            qtyParts: str3 ? [str3] : []
          });
          currentIdx = rows.length - 1;
          lastLineY = y;
        } else {
          // Continuity line: append text segments within the exact horizontal scope
          if (str0) rows[currentIdx].nameParts.push(str0);
          if (str1) rows[currentIdx].skuParts.push(str1);
          if (str2) rows[currentIdx].sellerSkuParts.push(str2);
          if (str3) rows[currentIdx].qtyParts.push(str3);
          lastLineY = y;
        }
      }

      // Convert rows to the structured order items with multiplication applied
      for (const r of rows) {
        const rawSku = r.sellerSkuParts.join("").trim();
        if (!rawSku) continue;

        let skuCandidate = rawSku.toUpperCase().replace(/\s/g, '');
        
        // Healing logic for CD
        const combinedRowText = (r.nameParts.join(" ") + " " + r.skuParts.join(" ") + " " + r.sellerSkuParts.join(" ") + " " + r.qtyParts.join(" ")).toUpperCase();
        const hasCDInRow = /\bCD\b/i.test(combinedRowText) || combinedRowText.includes("-CD") || combinedRowText.includes(" CD");
        if (hasCDInRow && !skuCandidate.includes("CD")) {
          // If sku candidate ends with a hyphen, we append CD
          if (skuCandidate.endsWith('-')) {
            skuCandidate += 'CD';
          } else {
            const endsWithNumber = /-\d+$/.test(skuCandidate);
            if (endsWithNumber) {
              skuCandidate += '-CD';
            } else {
              skuCandidate += '-CD';
            }
          }
        }

        // Clean leading/trailing hyphens as per instructions
        skuCandidate = skuCandidate.replace(/^-+/, '').replace(/-+$/, '');

        if (skuCandidate.length >= 4) {
          const qtyStr = r.qtyParts.join(" ").trim();
          let lineQty = 1;
          const numMatch = qtyStr.match(/\b([1-9]\d*)\b/);
          if (numMatch) {
            lineQty = parseInt(numMatch[1], 10);
          }

          // Handle composite combo SKU (separated by +)
          if (skuCandidate.includes('+')) {
            let isCompoundCD = skuCandidate.endsWith('-CD') || skuCandidate.includes('-CD-') || skuCandidate.endsWith('CD') || hasCDInRow;
            let cleanCompound = skuCandidate;
            
            // Check and clean CD
            if (cleanCompound.endsWith('-CD') || cleanCompound.endsWith('CD') || cleanCompound.includes('-CD-')) {
              isCompoundCD = true;
              cleanCompound = cleanCompound.replace(/-CD$/, '').replace(/CD$/, '').replace(/-+$/, '');
            }
            
            // Extract compound multiplier from the end, e.g. -2 or -3
            let compoundMultiplier = 1;
            const multMatch = cleanCompound.match(/-(\d+)$/);
            if (multMatch) {
              compoundMultiplier = parseInt(multMatch[1], 10);
              cleanCompound = cleanCompound.replace(/-\d+$/, '').replace(/-+$/, '');
            }

            // Secondary check if removing the multiplier revealed a hyphenated CD at the end
            if (cleanCompound.endsWith('-CD') || cleanCompound.endsWith('CD')) {
              isCompoundCD = true;
              cleanCompound = cleanCompound.replace(/-CD$/, '').replace(/CD$/, '').replace(/-+$/, '');
            }

            // Split and map each sub-sku
            const splitSkus = cleanCompound.split('+');
            splitSkus.forEach(part => {
              let subSku = part.trim().toUpperCase();
              if (!subSku) return;

              if (isCompoundCD && !subSku.endsWith('CD')) {
                subSku += '-CD';
              }

              subSku = subSku.replace(/^-+/, '').replace(/-+$/, '');
              const finalQty = lineQty * compoundMultiplier;
              const name = findSkuName(subSku, masterSkus) || "SKU_CHUA_KHAI_BAO";

              extractedItems.push({
                name,
                sku: subSku,
                qty: finalQty
              });
            });
          } else {
            // Standard single SKU
            const upperSku = skuCandidate.toUpperCase();
            const mult = getComboMultiplier(upperSku);
            const computedQty = lineQty * mult;
            const name = findSkuName(upperSku, masterSkus) || "SKU_CHUA_KHAI_BAO";

            extractedItems.push({
              name,
              sku: skuCandidate,
              qty: computedQty
            });
          }
        }
      }

      // Consolidate duplicate SKUs inside the same order page
      const consolidatedMap = new Map<string, ExtractedOrderItem>();
      extractedItems.forEach(item => {
        const key = item.sku.trim().toUpperCase();
        if (consolidatedMap.has(key)) {
          consolidatedMap.get(key)!.qty += item.qty;
        } else {
          consolidatedMap.set(key, { ...item });
        }
      });
      const uniqueItems = Array.from(consolidatedMap.values());

      // No default elements added, we keep uniqueItems completely empty if no products are found on this page

      orders.push({
        stt: i,
        orderId,
        province: senderProvince,
        originalPageIndex: i - 1,
        items: uniqueItems
      });
    }

    // Merge multi-page orders having the exact same Order ID
    const consolidatedOrdersMap = new Map<string, ExtractedOrder>();
    orders.forEach(order => {
      const idKey = order.orderId.trim().toUpperCase();
      if (consolidatedOrdersMap.has(idKey)) {
        const existingOrder = consolidatedOrdersMap.get(idKey)!;
        order.items.forEach(newItem => {
          const existingItem = existingOrder.items.find(it => it.sku.trim().toUpperCase() === newItem.sku.trim().toUpperCase());
          if (existingItem) {
            existingItem.qty += newItem.qty;
          } else {
            existingOrder.items.push({ ...newItem });
          }
        });
      } else {
        consolidatedOrdersMap.set(idKey, { ...order });
      }
    });

    const finalConsolidatedOrders = Array.from(consolidatedOrdersMap.values()).filter(order => order.items.length > 0);
    finalConsolidatedOrders.forEach((order, idx) => {
      order.stt = idx + 1;
    });

    return finalConsolidatedOrders;
  };

  const handleUnifiedPdfUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (file.size > 100 * 1024 * 1024) {
      setError("File PDF không được vượt quá 100MB.");
      return;
    }

    setUnifiedPdfFile(file);
    setIsProcessingUnified(true);
    setUnifiedProgress(0);
    setUnifiedStatus("Đang nạp file PDF và bắt đầu trích xuất hệ thống...");

    try {
      const orders = await extractDataFromPdfFile(file, (prog, msg) => {
        setUnifiedProgress(prog);
        setUnifiedStatus(msg);
      });

      if (orders.length === 0) {
        throw new Error("Không thể tìm thấy thông tin đơn hàng nào trong file PDF này.");
      }

      setExtractedOrders(orders);
      setStepUnified(2);
      setUnifiedStatus("Trích xuất dữ liệu hoàn tất! Xin vui lòng xem trước báo cáo và xuất file.");
    } catch (err: any) {
      setError(`Lỗi phân tích PDF: ${err.message}`);
      setUnifiedPdfFile(null);
    } finally {
      setIsProcessingUnified(false);
    }
  };

  // List of missing SKUs computed from extracted orders
  const missingSKUs = useMemo(() => {
    const missing = new Set<string>();
    extractedOrders.forEach(order => {
      order.items.forEach(item => {
        const cleanSku = item.sku.trim().toUpperCase();
        if (!findSkuName(cleanSku, masterSkus)) {
          missing.add(item.sku);
        }
      });
    });
    return Array.from(missing);
  }, [extractedOrders, masterSkus]);

  // Download logs of missing SKUs as requested in Section 6
  const downloadMissingSkusLog = (missing: string[], orders: ExtractedOrder[]) => {
    let content = `=========================================================\n`;
    content += `   DANH SÁCH SKU CHƯA KHAI BÁO TRONG HỆ THỐNG MASTER SKU\n`;
    content += `   Thời gian xuất log: ${new Date().toLocaleString('vi-VN')}\n`;
    content += `   Phát hiện tổng cộng: ${missing.length} SKU lỗi\n`;
    content += `=========================================================\n\n`;

    missing.forEach((sku, idx) => {
      const matchingOrders = orders.filter(o => o.items.some(it => it.sku.trim().toUpperCase() === sku.trim().toUpperCase()));
      const orderIds = matchingOrders.map(o => o.orderId).join(', ');
      
      content += `${idx + 1}. Mã SKU: "${sku}"\n`;
      content += `   - Số lượng đơn bị ảnh hưởng: ${matchingOrders.length}\n`;
      content += `   - Danh sách Mã đơn hàng mẫu: [${orderIds}]\n\n`;
    });

    content += `---------------------------------------------------------\n`;
    content += `HƯỚNG DẪN KHẮC PHỤC:\n`;
    content += `1. Vào mục "Cấu hình dữ liệu SKU"\n`;
    content += `2. Cập nhật thêm các SKU thiếu trên vào file Master SKU Excel mới\n`;
    content += `3. Tải lên file mới và lưu làm dữ liệu mặc định để hệ thống tự động ghi nhận.\n`;
    content += `=========================================================`;

    const blob = new Blob([content], { type: 'text/plain;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `log_sku_chua_khai_bao_${new Date().toISOString().slice(0,10)}.txt`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const getCustomFileName = (province: string, count: number) => {
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = now.getFullYear();
    const hour = now.getHours();
    const min = String(now.getMinutes()).padStart(2, '0');
    const sec = String(now.getSeconds()).padStart(2, '0');
    const buoi = hour < 12 ? "Sáng" : "Chiều";
    const timeStr = `${String(hour).padStart(2, '0')}-${min}-${sec}`;
    return `${province}-${count}_đơn-${day}_${month}_${year}-${buoi}-${timeStr}`;
  };

  const exportCleanedExcelFromExtracted = async (sortedOrders: ExtractedOrder[], customFilename?: string) => {
    if (sortedOrders.length === 0) return;

    setIsProcessingUnified(true);
    setUnifiedStatus("Đang tạo và định dạng bảng dữ liệu Excel...");

    try {
      const workbook = new ExcelJS.Workbook();

      // --- SHEET 1: ĐƠN HÀNG ---
      const worksheet1 = workbook.addWorksheet('Đơn hàng');
      const headers1 = ["STT", "Mã đơn hàng", "Tên sản phẩm", "Mã SKU", "Số lượng sản phẩm", "Ghi chú"];

      const headerRow1 = worksheet1.addRow(headers1);
      headerRow1.eachCell((cell, colNumber) => {
        cell.font = { bold: true };
        if (colNumber === 5) {
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'left' };
        }
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
      });

      // Flatten matching list
      const flatRows: { stt: number; orderId: string; name: string; sku: string; qty: number; note: string }[] = [];
      sortedOrders.forEach((order, ordIdx) => {
        order.items.forEach(item => {
          flatRows.push({
            stt: ordIdx + 1,
            orderId: order.orderId,
            name: item.name,
            sku: item.sku,
            qty: item.qty,
            note: ""
          });
        });
      });

      let lastSttValue: any = null;
      let mergeStartRow = 2;

      flatRows.forEach((row, idx) => {
        const currentStt = row.stt;

        const rowData = [
          currentStt,
          row.orderId,
          row.name,
          row.sku,
          row.qty,
          row.note
        ];

        const newRow = worksheet1.addRow(rowData);
        const currentRowNum = newRow.number;

        const qty = row.qty;
        newRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          cell.border = {
            top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
          };
          
          if (qty > 1) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFA6A6A6' }
            };
            cell.font = { color: { argb: 'FFFFFFFF' }, size: 12 };
          }
          
          if (colNumber === 1) {
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
          } else if (colNumber === 5) {
            if (qty > 1) {
              cell.alignment = { vertical: 'middle', horizontal: 'center' };
            } else {
              cell.alignment = { vertical: 'middle', horizontal: 'right' };
            }
          } else {
            cell.alignment = { vertical: 'middle', horizontal: 'left' };
          }
        });

        // Merge STT A cell
        if (idx > 0 && currentStt === lastSttValue && currentStt !== null) {
          // Merge active segment
        } else {
          if (idx > 0 && currentRowNum - 1 > mergeStartRow) {
            worksheet1.mergeCells(`A${mergeStartRow}:A${currentRowNum - 1}`);
          }
          mergeStartRow = currentRowNum;
        }

        lastSttValue = currentStt;

        if (idx === flatRows.length - 1 && currentRowNum > mergeStartRow) {
          worksheet1.mergeCells(`A${mergeStartRow}:A${currentRowNum}`);
        }
      });

      worksheet1.columns.forEach((column, index) => {
        if (index === 0) column.width = 8; // STT
        else if (index === 1) column.width = 25; // Order ID
        else if (index === 2) column.width = 40; // Product Name
        else column.width = 20;
      });

      // --- SHEET 2: DANH SÁCH SẢN PHẨM ---
      const worksheet2 = workbook.addWorksheet('Danh sách sản phẩm');
      const summaryMap = new Map<string, { name: string; sku: string; totalQty: number }>();

      flatRows.forEach(row => {
        const key = `${row.sku}|${row.name}`;
        if (summaryMap.has(key)) {
          summaryMap.get(key)!.totalQty += row.qty;
        } else {
          summaryMap.set(key, { name: row.name, sku: row.sku, totalQty: row.qty });
        }
      });

      const headers2 = ["Tên sản phẩm", "Mã SKU", "Sum of Số lượng sản phẩm"];
      const headerRow2 = worksheet2.addRow(headers2);
      headerRow2.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
      });

      const sortedSummary = Array.from(summaryMap.values()).sort((a, b) => a.name.localeCompare(b.name));

      sortedSummary.forEach((data) => {
        const newRow = worksheet2.addRow([data.name, data.sku, data.totalQty]);
        newRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          cell.border = {
            top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
          };
          if (colNumber === 1) {
            cell.font = { bold: true };
            cell.alignment = { vertical: 'middle', horizontal: 'left' };
          } else if (colNumber === 3) {
            cell.alignment = { vertical: 'middle', horizontal: 'right' };
          } else {
            cell.alignment = { vertical: 'middle', horizontal: 'left' };
          }
        });
      });

      worksheet2.autoFilter = {
        from: 'A1',
        to: 'B' + (sortedSummary.length + 1)
      };

      const totalRow = worksheet2.addRow(['Grand Total', '', sortedSummary.reduce((acc, curr) => acc + curr.totalQty, 0)]);
      totalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
        if (colNumber === 3) {
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'left' };
        }
      });

      worksheet2.columns.forEach((column, index) => {
        column.width = index === 0 ? 50 : 25;
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = customFilename ? `${customFilename}.xlsx` : `Orders_Report_${new Date().getTime()}.xlsx`;
      link.click();

      setUnifiedStatus("Xuất file báo cáo Excel hoàn tất!");
    } catch (err: any) {
      setError(`Lỗi sinh Excel: ${err.message}`);
    } finally {
      setIsProcessingUnified(false);
    }
  };

  const generateSortedPdfFromExtracted = async (originalFile: File, sortedOrders: ExtractedOrder[], customFilename?: string) => {
    if (!originalFile || sortedOrders.length === 0) return;
    setIsProcessingUnified(true);
    setUnifiedProgress(0);
    setUnifiedStatus('Đang nạp dữ liệu PDF gốc...');

    try {
      const existingPdfBytes = await originalFile.arrayBuffer();
      const originalPdfDoc = await PDFDocument.load(new Uint8Array(existingPdfBytes), { 
        ignoreEncryption: true 
      });
      const newPdfDoc = await PDFDocument.create();

      for (let i = 0; i < sortedOrders.length; i++) {
        const order = sortedOrders[i];
        const pageIdx = order.originalPageIndex;
        const [copiedPage] = await newPdfDoc.copyPages(originalPdfDoc, [pageIdx]);
        const addedPage = newPdfDoc.addPage(copiedPage);

        if (maskProductInfo) {
          const { width, height } = addedPage.getSize();
          addedPage.drawRectangle({
            x: 0,
            y: 0,
            width: width,
            height: height * (maskHeightPercent / 100),
            color: rgb(1, 1, 1),
          });
        }

        setUnifiedStatus(`Đang tiến hành gom sắp xếp trang ${i + 1}/${sortedOrders.length}...`);
        setUnifiedProgress(Math.round(((i + 1) / sortedOrders.length) * 100));
      }

      setUnifiedStatus('Đang hoàn tất và đóng gói PDF...');
      const pdfBytes = await newPdfDoc.save();
      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      const originalName = originalFile.name.replace('.pdf', '');
      link.href = url;
      link.download = customFilename ? `${customFilename}.pdf` : `${originalName}_sorted.pdf`;
      link.click();
      setUnifiedStatus('Xuất file PDF sắp xếp hoàn tất!');
    } catch (err: any) {
      setError(`Lỗi sinh PDF sắp xếp: ${err.message}`);
    } finally {
      setIsProcessingUnified(false);
    }
  };

  const handleRawExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {

    const file = e.target.files?.[0];
    if (!file) return;

    if (file.size > 100 * 1024 * 1024) {
      setError("File Excel không được vượt quá 100MB.");
      return;
    }

    setRawExcelFile(file);
    setError(null);
    setStatusMessage("Đang phân tích file...");

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json<any>(sheet, { header: 'A', raw: false });

      if (jsonData.length === 0) throw new Error("File trống.");

      // Find header row
      let headerIdx = -1;
      let targetMapping: Record<string, string> = {};
      
      const targets = [
        { label: "STT", keywords: ["stt", "số thứ tự", "no.", "index"] },
        { label: "Mã đơn hàng", keywords: ["đơn hàng", "order id", "mã đơn"] },
        { label: "Tên sản phẩm", keywords: ["tên sản phẩm", "product name", "tên hàng"] },
        { label: "Mã SKU", keywords: ["sku", "mã hàng", "phân loại", "variant"] },
        { 
          label: "Số lượng sản phẩm", 
          keywords: ["số lượng", "quantity", "qty"],
          exclude: ["tổng", "total"] 
        },
        { label: "Ghi chú", keywords: ["ghi chú", "note"] }
      ];

      for (let i = 0; i < Math.min(jsonData.length, 20); i++) {
        const row = jsonData[i];
        let foundCount = 0;
        const currentMapping: Record<string, string> = {};

        for (const target of targets) {
          for (const [key, val] of Object.entries(row)) {
            const lowVal = String(val || '').toLowerCase().trim();
            
            const matches = target.keywords.some(k => lowVal.includes(k));
            const excluded = ('exclude' in target) && (target.exclude as string[]).some(e => lowVal.includes(e));

            if (matches && !excluded) {
              currentMapping[target.label] = key;
              foundCount++;
              break;
            }
          }
        }

        if (foundCount >= 2) {
          headerIdx = i;
          targetMapping = currentMapping;
          break;
        }
      }

      if (headerIdx === -1) {
        throw new Error("Không thể tự động nhận diện các cột. Vui lòng kiểm tra lại file Excel.");
      }

      setExcelMapping(targetMapping);
      setRawExcelData(jsonData.slice(headerIdx + 1));
      setStatusMessage("Phân tích xong. Sẵn sàng tải xuống.");

    } catch (err: any) {
      setError(`Lỗi: ${err.message}`);
      setRawExcelFile(null);
    }
  };

  const exportCleanedExcel = async () => {
    if (!rawExcelData.length) return;
    
    setIsProcessing(true);
    setStatusMessage("Đang tạo file excel với 2 trang...");
    try {
      const workbook = new ExcelJS.Workbook();
      
      // --- SHEET 1: CHI TIẾT ---
      const worksheet1 = workbook.addWorksheet('Đơn hàng');
      const headers1 = ["STT", "Mã đơn hàng", "Tên sản phẩm", "Mã SKU", "Số lượng sản phẩm", "Ghi chú"];
      
      const headerRow1 = worksheet1.addRow(headers1);
      headerRow1.eachCell((cell, colNumber) => {
        cell.font = { bold: true };
        // Headers left by default, right for specified numeric columns
        if (colNumber === 5) {
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'left' };
        }
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
      });

      let lastSttValue: any = null;
      let mergeStartRow = 2;

      rawExcelData.forEach((row, idx) => {
        // Carry forward STT if empty (common for merged cells in source)
        const rawStt = row[excelMapping["STT"]];
        const currentStt = (rawStt !== undefined && rawStt !== null && String(rawStt).trim() !== "") 
          ? String(rawStt).trim() 
          : lastSttValue;

        const rowData = [
          currentStt,
          row[excelMapping["Mã đơn hàng"]] || "",
          row[excelMapping["Tên sản phẩm"]] || "",
          row[excelMapping["Mã SKU"]] || "",
          Number(row[excelMapping["Số lượng sản phẩm"]] || 0),
          row[excelMapping["Ghi chú"]] || ""
        ];
        
        const newRow = worksheet1.addRow(rowData);
        const currentRowNum = newRow.number;

        const qty = Number(row[excelMapping["Số lượng sản phẩm"]] || 0);

        // Apply borders and alignment
        newRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          cell.border = {
            top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
          };
          
          if (qty > 1) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFA6A6A6' }
            };
            cell.font = { color: { argb: 'FFFFFFFF' }, size: 12 };
          }
          
          // Numeric columns right, STT center, rest left
          if (colNumber === 1) {
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
          } else if (colNumber === 5) {
            if (qty > 1) {
              cell.alignment = { vertical: 'middle', horizontal: 'center' };
            } else {
              cell.alignment = { vertical: 'middle', horizontal: 'right' };
            }
          } else {
            cell.alignment = { vertical: 'middle', horizontal: 'left' };
          }
        });

        // Merge STT Column (Col A) logic
        if (idx > 0 && currentStt === lastSttValue && currentStt !== null) {
          // Still in the same STT group
        } else {
          if (idx > 0 && currentRowNum - 1 > mergeStartRow) {
            worksheet1.mergeCells(`A${mergeStartRow}:A${currentRowNum - 1}`);
          }
          mergeStartRow = currentRowNum;
        }
        
        lastSttValue = currentStt;

        // Final merge for the last group
        if (idx === rawExcelData.length - 1 && currentRowNum > mergeStartRow) {
          worksheet1.mergeCells(`A${mergeStartRow}:A${currentRowNum}`);
        }
      });

      worksheet1.columns.forEach((column, index) => {
        if (index === 0) column.width = 8; // STT
        else if (index === 1) column.width = 25; // Order ID
        else if (index === 2) column.width = 40; // Product Name
        else column.width = 20;
      });

      // --- SHEET 2: TỔNG HỢP ---
      const worksheet2 = workbook.addWorksheet('Danh sách sản phẩm');
      
      const summaryMap = new Map<string, { name: string, sku: string, totalQty: number }>();
      
      rawExcelData.forEach(row => {
        const sku = String(row[excelMapping["Mã SKU"]] || 'N/A').trim();
        const name = String(row[excelMapping["Tên sản phẩm"]] || 'N/A').trim();
        const qty = Number(row[excelMapping["Số lượng sản phẩm"]] || 0);
        
        const key = `${sku}|${name}`;
        if (summaryMap.has(key)) {
          summaryMap.get(key)!.totalQty += qty;
        } else {
          summaryMap.set(key, { name, sku, totalQty: qty });
        }
      });

      const headers2 = ["Tên sản phẩm", "Mã SKU", "Sum of Số lượng sản phẩm"];
      const headerRow2 = worksheet2.addRow(headers2);
      headerRow2.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
        // Headers left as requested (all headers left, except specific ones in sheet 1)
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
      });

      const sortedSummary = Array.from(summaryMap.values()).sort((a, b) => a.name.localeCompare(b.name));

      sortedSummary.forEach((data) => {
        const newRow = worksheet2.addRow([data.name, data.sku, data.totalQty]);
        newRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          cell.border = {
            top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
          };
          // Column 1 bold, Column 3 right-aligned
          if (colNumber === 1) {
            cell.font = { bold: true };
            cell.alignment = { vertical: 'middle', horizontal: 'left' };
          } else if (colNumber === 3) {
            cell.alignment = { vertical: 'middle', horizontal: 'right' };
          } else {
            cell.alignment = { vertical: 'middle', horizontal: 'left' };
          }
        });
      });

      // Add AutoFilter for Columns 1 and 2
      worksheet2.autoFilter = {
        from: 'A1',
        to: 'B' + (sortedSummary.length + 1)
      };

      const totalRow = worksheet2.addRow(['Grand Total', '', sortedSummary.reduce((acc, curr) => acc + curr.totalQty, 0)]);
      totalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
        if (colNumber === 3) {
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'left' };
        }
      });

      worksheet2.columns.forEach((column, index) => {
        column.width = index === 0 ? 50 : 25;
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `orders_refined_${new Date().getTime()}.xlsx`;
      link.click();
      
      setStatusMessage("Đã xuất hoàn tất file gồm 2 trang.");
    } catch (err: any) {
      setError(`Lỗi xuất file: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const cleanOrderId = (val: any): string => {
    return String(val || '').trim().replace(/['"\s]/g, '');
  };

  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // 100MB check
    if (file.size > 100 * 1024 * 1024) {
      setError("File Excel/CSV không được vượt quá 100MB.");
      return;
    }

    setExcelFile(file);
    setError(null);
    
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      
      let sheetName = workbook.SheetNames.find(name => name.toLowerCase() === 'orders');
      if (!sheetName && workbook.SheetNames.length > 0) {
        sheetName = workbook.SheetNames[0];
      }

      if (!sheetName) throw new Error("Không thể tìm thấy sheet \"Orders\".");

      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json<any>(worksheet, { header: 'A', raw: false });

      let headerRowIndex = -1;
      let orderIdColumnKey = 'B';

      // Advanced header detection
      for (let i = 0; i < Math.min(jsonData.length, 20); i++) {
        const row = jsonData[i];
        for (const key in row) {
          const val = String(row[key] || '').toLowerCase().trim();
          if (val === 'mã đơn hàng' || val === 'order id' || val === 'mã đơn' || val === 'orderid') {
            headerRowIndex = i;
            orderIdColumnKey = key;
            break;
          }
        }
        if (headerRowIndex !== -1) break;
      }

      const orders: OrderInfo[] = [];
      const seenOrders = new Set<string>();

      for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        let orderId = cleanOrderId(row[orderIdColumnKey]);
        
        if (!orderId || orderId === 'undefined' || orderId === '') continue;

        if (!seenOrders.has(orderId)) {
          seenOrders.add(orderId);
          orders.push({
            stt: orders.length + 1,
            orderId,
            sourceRow: i + 1
          });
        }
      }

      if (orders.length === 0) {
        throw new Error("Không tìm thấy mã đơn hàng nào trong file Excel.");
      }

      setExcelOrders(orders);
    } catch (err: any) {
      setError(`Lỗi Excel: ${err.message}`);
      setExcelFile(null);
      setExcelOrders([]);
    }
  };

  const handlePdfUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // 100MB check
    if (file.size > 100 * 1024 * 1024) {
      setError("File PDF không được vượt quá 100MB.");
      return;
    }

    setPdfFile(file);
    setError(null);
  };

  const processFiles = async () => {
    if (!excelFile || !pdfFile || excelOrders.length === 0) return;
    
    setIsProcessing(true);
    setProgress(0);
    setError(null);
    setStatusMessage('Đang trích xuất dữ liệu từ PDF...');
    
    try {
      const arrayBuffer = await pdfFile.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      const totalPages = pdf.numPages;
      const pagesInfo: PDFPageInfo[] = [];

      const orderIdRegex = /Order\s*ID\s*:\s*(\d{10,35})/i;
      const excelOrderIds = excelOrders.map(o => o.orderId);

      for (let i = 1; i <= totalPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        
        const strItems = textContent.items.map((item: any) => item.str);
        const fullText = strItems.join(' ');
        const compressedText = strItems.join('').replace(/\s/g, '');
        
        let foundOrderId: string | null = null;
        
        const match = fullText.match(orderIdRegex);
        if (match && match[1]) {
          foundOrderId = match[1];
        } else {
          for (const id of excelOrderIds) {
            if (fullText.includes(id) || compressedText.includes(id)) {
              foundOrderId = id;
              break;
            }
          }
        }

        pagesInfo.push({
          pageIndex: i - 1,
          orderId: foundOrderId,
          status: foundOrderId ? 'matched' : 'not_found'
        });

        setProgress(Math.round((i / totalPages) * 100));
      }

      setPdfPages(pagesInfo);
      setStep(2);
      setStatusMessage('Phân tích hoàn tất');
    } catch (err: any) {
      setError(`Lỗi PDF: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const matchResults = useMemo((): MatchResult[] => {
    const results: MatchResult[] = [];
    const pdfMap = new Map<string, number[]>();
    
    pdfPages.forEach((page) => {
      if (page.orderId) {
        const existing = pdfMap.get(page.orderId) || [];
        existing.push(page.pageIndex);
        pdfMap.set(page.orderId, existing);
      }
    });

    excelOrders.forEach((excelOrder, idx) => {
      const pageIndices = pdfMap.get(excelOrder.orderId);
      
      if (!pageIndices || pageIndices.length === 0) {
        results.push({
          excelIndex: idx,
          orderId: excelOrder.orderId,
          pdfPageIndex: null,
          status: 'missing_in_pdf'
        });
      } else if (pageIndices.length > 1) {
        results.push({
          excelIndex: idx,
          orderId: excelOrder.orderId,
          pdfPageIndex: pageIndices[0],
          status: 'duplicate_in_pdf'
        });
      } else {
        results.push({
          excelIndex: idx,
          orderId: excelOrder.orderId,
          pdfPageIndex: pageIndices[0],
          status: 'matched'
        });
      }
    });

    return results;
  }, [excelOrders, pdfPages]);

  const stats = useMemo((): ProcessingStats => {
    const matchedCount = matchResults.filter(r => r.status === 'matched').length;
    const missingInPdfCount = matchResults.filter(r => r.status === 'missing_in_pdf').length;
    const excelIdSet = new Set(excelOrders.map(o => o.orderId));
    const extraInPdfCount = pdfPages.filter(p => p.orderId && !excelIdSet.has(p.orderId)).length;
    const errorPdfCount = pdfPages.filter(p => !p.orderId).length;

    return {
      totalExcelRows: excelOrders.length,
      uniqueOrders: excelOrders.length,
      totalPdfPages: pdfPages.length,
      matchedCount,
      missingInPdfCount,
      extraInPdfCount,
      errorPdfCount
    };
  }, [excelOrders, pdfPages, matchResults]);

  const generateSortedPdf = async () => {
    if (!pdfFile || matchResults.length === 0) return;
    setIsProcessing(true);
    setProgress(0);
    setStatusMessage('Đang khởi tạo file PDF mới...');

    try {
      const existingPdfBytes = await pdfFile.arrayBuffer();
      const originalPdfDoc = await PDFDocument.load(new Uint8Array(existingPdfBytes), { 
        ignoreEncryption: true 
      });
      const newPdfDoc = await PDFDocument.create();

      const exportItems = matchResults.filter(r => (r.status === 'matched' || r.status === 'duplicate_in_pdf') && r.pdfPageIndex !== null);

      for (let i = 0; i < exportItems.length; i++) {
        const result = exportItems[i];
        if (result.pdfPageIndex !== null) {
          const [copiedPage] = await newPdfDoc.copyPages(originalPdfDoc, [result.pdfPageIndex]);
          const addedPage = newPdfDoc.addPage(copiedPage);

          if (maskProductInfo) {
            const { width, height } = addedPage.getSize();
            addedPage.drawRectangle({
              x: 0,
              y: 0,
              width: width,
              height: height * (maskHeightPercent / 100),
              color: rgb(1, 1, 1),
            });
          }
        }
        setStatusMessage(`Đang thêm trang ${i + 1}/${exportItems.length}...`);
        setProgress(Math.round(((i + 1) / exportItems.length) * 100));
      }

      setStatusMessage('Gói file dữ liệu...');
      const pdfBytes = await newPdfDoc.save();
      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      const originalName = pdfFile.name.replace('.pdf', '');
      link.href = url;
      link.download = `${originalName}_sorted.pdf`;
      link.click();
      setStatusMessage('Hoàn tất. File đã được tải về.');
    } catch (err: any) {
      setError(`Lỗi tạo PDF: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="bg-slate-50 w-full min-h-screen flex flex-col font-sans overflow-x-hidden text-slate-900 selection:bg-indigo-100">
      <header className="bg-white/80 backdrop-blur-md border-b border-slate-200 px-6 py-4 flex flex-col sm:flex-row justify-between items-center shrink-0 shadow-sm sticky top-0 z-50">
        <div className="flex items-center mb-4 sm:mb-0">
          <div className="size-[50px] bg-slate-100 rounded-2xl flex items-center justify-center shadow-inner mr-4 shrink-0 overflow-hidden border border-slate-200">
            <img 
              src="https://i.postimg.cc/7ZLc3GnZ/anh-logot.png" 
              alt="New Era Food Logo" 
              className="h-full w-full object-contain scale-[1.2]" 
              referrerPolicy="no-referrer" 
            />
          </div>
          <div>
            <h1 className="text-xl font-black tracking-tight text-slate-800 italic leading-tight">XỬ LÝ ĐƠN NỘI BỘ</h1>
            <p className="text-[10px] text-slate-400 uppercase font-black tracking-widest leading-none">v2.6 • High Density Mode</p>
          </div>
        </div>
      </header>

      {/* Tab Navigation */}
      <nav className="bg-slate-50/50 px-6 pt-2 flex items-center space-x-2 shrink-0 z-40">
        <button 
          onClick={() => { setActiveTab('unified'); }}
          className={cn(
            "px-8 py-3 text-[11px] font-black uppercase tracking-widest transition-all duration-300 rounded-t-xl border-t border-x flex items-center gap-2",
            activeTab === 'unified' 
              ? "bg-white border-slate-200 text-indigo-600 shadow-[0_-4px_12px_-4px_rgba(0,0,0,0.05)] font-extrabold" 
              : "bg-transparent border-transparent text-slate-400 hover:text-slate-600 hover:bg-slate-100"
          )}
        >
          <span className="relative flex h-2 w-2">
            <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-indigo-400 opacity-75"></span>
            <span className="relative inline-flex rounded-full h-2 w-2 bg-indigo-600"></span>
          </span>
          Quy trình tinh gọn (MỚI)
        </button>
        <button 
          onClick={() => { setActiveTab('sku_config'); }}
          className={cn(
            "px-8 py-3 text-[11px] font-black uppercase tracking-widest transition-all duration-300 rounded-t-xl border-t border-x flex items-center gap-2",
            activeTab === 'sku_config' 
              ? "bg-white border-slate-200 text-indigo-600 shadow-[0_-4px_12px_-4px_rgba(0,0,0,0.05)] font-extrabold" 
              : "bg-transparent border-transparent text-slate-400 hover:text-slate-600 hover:bg-slate-100"
          )}
        >
          <SlidersHorizontal className="w-3.5 h-3.5" />
          Cấu hình dữ liệu SKU
        </button>
      </nav>

      <main className="flex-1 flex flex-col p-6 space-y-6 max-w-7xl mx-auto w-full overflow-hidden bg-white shadow-[0_0_50px_-12px_rgba(0,0,0,0.05)] border-x border-slate-100">
        
        {/* Computing Unified Data States */}
        {(() => {
          // Calculate province counts
          const provinceCounts: Record<string, number> = {};
          extractedOrders.forEach(o => {
            provinceCounts[o.province] = (provinceCounts[o.province] || 0) + 1;
          });

          const dnOrders = extractedOrders.filter(o => o.province === "Đà Nẵng");
          const hnOrders = extractedOrders.filter(o => o.province === "Hà Nội");
          const hcmOrders = extractedOrders.filter(o => o.province === "Hồ Chí Minh");
          const otherOrders = extractedOrders.filter(o => o.province !== "Đà Nẵng" && o.province !== "Hà Nội" && o.province !== "Hồ Chí Minh");



          // Sort extracted orders list based on settings
          const sortedExtractedOrders = [...extractedOrders];
          if (sortOption === 'province') {
            sortedExtractedOrders.sort((a, b) => {
              let idxA = provincePriority.indexOf(a.province);
              let idxB = provincePriority.indexOf(b.province);
              if (idxA === -1) idxA = provincePriority.length;
              if (idxB === -1) idxB = provincePriority.length;

              if (idxA !== idxB) {
                return idxA - idxB;
              }
              return a.orderId.localeCompare(b.orderId);
            });
          } else if (sortOption === 'sku') {
            sortedExtractedOrders.sort((a, b) => {
              const skuA = a.items[0]?.sku || '';
              const skuB = b.items[0]?.sku || '';
              return skuA.localeCompare(skuB);
            });
          } else if (sortOption === 'qty') {
            sortedExtractedOrders.sort((a, b) => {
              const qtyA = a.items.reduce((sum, i) => sum + i.qty, 0);
              const qtyB = b.items.reduce((sum, i) => sum + i.qty, 0);
              return qtyA - qtyB;
            });
          }

          // Aggregate item preparado count list
          const listMap = new Map<string, { name: string; sku: string; totalQty: number }>();
          extractedOrders.forEach(o => {
            o.items.forEach(item => {
              const key = `${item.sku}|${item.name}`;
              if (listMap.has(key)) {
                listMap.get(key)!.totalQty += item.qty;
              } else {
                listMap.set(key, { name: item.name, sku: item.sku, totalQty: item.qty });
              }
            });
          });
          const aggregateSummaryList = Array.from(listMap.values()).sort((a, b) => b.totalQty - a.totalQty);

          return (
            <>
              {/* === TAB: UNIFIED (Quy trình tinh gọn (MỚI)) === */}
              <div className={cn("flex-1 flex flex-col space-y-6", activeTab !== "unified" && "hidden")}>
                {stepUnified === 1 ? (
                  <motion.div 
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0 }}
                    className="bg-white rounded-[2rem] border border-slate-200 p-8 sm:p-12 flex flex-col items-center justify-center text-center space-y-10 shadow-sm flex-1 min-h-[450px]"
                  >
                    <div className={cn(
                      "group relative border-2 border-dashed rounded-[2.5rem] p-12 sm:p-16 w-full max-w-2xl transition-all duration-500 shadow-inner",
                      unifiedPdfFile ? "border-indigo-600 bg-indigo-50/10" : "border-slate-200 hover:border-indigo-400 bg-slate-50/30"
                    )}>
                      <input 
                        type="file" 
                        accept=".pdf" 
                        onChange={handleUnifiedPdfUpload} 
                        className="absolute inset-0 opacity-0 cursor-pointer z-10" 
                        ref={unifiedPdfInputRef} 
                        disabled={isProcessingUnified} 
                      />
                      
                      <div className={cn(
                        "size-24 rounded-3xl mx-auto mb-8 flex items-center justify-center shadow-lg transition-all duration-500", 
                        unifiedPdfFile ? "bg-indigo-600 text-white shadow-indigo-200 scale-110 rotate-6" : "bg-white text-slate-400 group-hover:scale-110"
                      )}>
                        <FileText className="w-12 h-12" />
                      </div>

                      <h2 className="text-xl sm:text-2xl font-black text-slate-800 uppercase tracking-tight">
                        QUY TRÌNH TINH GỌN (1-CLICK)
                      </h2>
                      <p className="text-xs sm:text-sm text-slate-400 mt-2 font-black uppercase tracking-wider">
                        BƯỚC 1: TẢI SẢN PHẨM / ĐƠN HÀNG FILE PDF GỐC
                      </p>
                      <p className="text-xs text-slate-400 mt-2 font-medium max-w-sm mx-auto leading-relaxed">
                        Hệ thống sẽ tự động trích xuất Order ID, Seller SKU, Số lượng, Địa chỉ Hồ Chí Minh / Hà Nội / Đà Nẵng từ File PDF gốc. Sau đó hỗ trợ xuất cả Excel dữ liệu & PDF gom sắp xếp.
                      </p>
                      
                      {isProcessingUnified && (
                        <div className="mt-8 space-y-4 max-w-md mx-auto">
                          <div className="flex justify-between items-center text-xs text-indigo-600 font-bold uppercase tracking-wider">
                            <span>{unifiedStatus}</span>
                            <span>{unifiedProgress}%</span>
                          </div>
                          <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden p-0.5 border border-slate-200 shadow-sm">
                            <motion.div 
                              className="bg-indigo-600 h-full rounded-full" 
                              initial={{ width: 0 }} 
                              animate={{ width: `${unifiedProgress}%` }} 
                            />
                          </div>
                        </div>
                      )}

                      {!isProcessingUnified && unifiedPdfFile && (
                        <motion.div 
                          initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}
                          className="mt-8 px-6 py-3 bg-white rounded-2xl shadow-sm border border-indigo-100 inline-flex items-center gap-3"
                        >
                          <div className="size-2 rounded-full bg-indigo-500 animate-pulse"></div>
                          <span className="text-[11px] font-black text-slate-700 uppercase tracking-widest truncate max-w-[200px]">{unifiedPdfFile.name}</span>
                        </motion.div>
                      )}
                    </div>
                  </motion.div>
                ) : (
                  // Step 2: Custom Dashboard & Action Center
                  <div className="flex-1 flex flex-col space-y-6">
                    
                    {/* Visual Widgets Grid */}
                    <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                      {/* Total Orders Widget */}
                      <div className="bg-slate-50 border border-slate-200/60 rounded-3xl p-5 flex items-center space-x-4 shadow-sm">
                        <div className="p-3.5 bg-indigo-100/80 rounded-2xl text-indigo-600">
                          <ShoppingBag className="w-6 h-6" />
                        </div>
                        <div>
                          <p className="text-[10px] uppercase font-black text-slate-400 tracking-wider">Tổng Đơn Hàng</p>
                          <p className="text-2xl font-black text-slate-800 tracking-tight font-mono">{extractedOrders.length}</p>
                        </div>
                      </div>

                      {/* Total Items SKU Widget */}
                      <div className="bg-slate-50 border border-slate-200/60 rounded-3xl p-5 flex items-center space-x-4 shadow-sm">
                        <div className="p-3.5 bg-emerald-100/80 rounded-2xl text-emerald-600">
                          <Boxes className="w-6 h-6" />
                        </div>
                        <div>
                          <p className="text-[10px] uppercase font-black text-slate-400 tracking-wider">Tổng sản phẩm SKU</p>
                          <p className="text-2xl font-black text-slate-800 tracking-tight font-mono">
                            {extractedOrders.reduce((sum, ord) => sum + ord.items.reduce((acc, i) => acc + i.qty, 0), 0)}
                          </p>
                        </div>
                      </div>

                      {/* Single Orders Widget */}
                      <div className="bg-slate-50 border border-slate-200/60 rounded-3xl p-5 flex items-center space-x-4 shadow-sm">
                        <div className="p-3.5 bg-sky-100/80 rounded-2xl text-sky-600">
                          <ArrowRight className="w-6 h-6 animate-pulse" />
                        </div>
                        <div>
                          <p className="text-[10px] uppercase font-black text-slate-400 tracking-wider">Đơn Lẻ (Qty = 1)</p>
                          <p className="text-2xl font-black text-slate-800 tracking-tight font-mono">
                            {extractedOrders.filter(ord => ord.items.reduce((sum, i) => sum + i.qty, 0) === 1).length}
                          </p>
                        </div>
                      </div>

                      {/* Combos Combo Widget */}
                      <div className="bg-slate-50 border border-slate-200/60 rounded-3xl p-5 flex items-center space-x-4 shadow-sm">
                        <div className="p-3.5 bg-amber-100/80 rounded-2xl text-amber-600">
                          <Layers className="w-6 h-6" />
                        </div>
                        <div>
                          <p className="text-[10px] uppercase font-black text-slate-400 tracking-wider">Đơn Sỉ / Combo (&gt;1)</p>
                          <p className="text-2xl font-black text-slate-800 tracking-tight font-mono">
                            {extractedOrders.filter(ord => ord.items.reduce((sum, i) => sum + i.qty, 0) > 1).length}
                          </p>
                        </div>
                      </div>
                    </div>

                    {/* Dashboard Core Columns */}
                    <div className="max-w-4xl mx-auto w-full flex flex-col space-y-6">
                      
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 items-stretch">
                        
                        {/* Primary Export Actions Center */}
                        <div className="bg-[#fcfcff] border border-indigo-100 rounded-[2rem] p-6 space-y-4 shadow-sm flex flex-col justify-between h-full">
                          <div className="flex justify-between items-center">
                            <h4 className="text-xs font-black text-indigo-950 uppercase tracking-widest flex items-center gap-1.5">
                              <CheckCircle2 className="w-4 h-4 text-indigo-600" />
                              TRÌNH XUẤT 2-TRONG-1
                            </h4>
                            <button
                              onClick={() => window.location.reload()}
                              className="p-1 px-2.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-black text-[10px] uppercase tracking-wider rounded-xl border border-indigo-100/60 flex items-center gap-1 cursor-pointer transition-all active:scale-95 shadow-sm"
                              title="Tải lại trang"
                            >
                              <RefreshCcw className="w-3.5 h-3.5" /> Reload
                            </button>
                          </div>
                          <p className="text-[11px] text-slate-400 leading-relaxed font-semibold">
                            Sau khi điều chỉnh mức độ ưu tiên và tiêu chí sắp xếp bên dưới, bạn có thể tải ngay 2 file báo cáo tương ứng:
                          </p>

                          {missingSKUs.length > 0 && (
                            <div className="bg-amber-50 border border-amber-200 rounded-2xl p-4 space-y-3 shadow-inner">
                              <div className="flex items-start gap-2">
                                <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5 animate-bounce" />
                                <div>
                                  <p className="text-[10px] font-black text-amber-950 uppercase tracking-wider">SKU CHƯA KHAI BÁO!</p>
                                  <p className="text-[9px] text-amber-700 font-bold uppercase leading-tight mt-0.5">
                                    Có {missingSKUs.length} SKU từ PDF chưa có tên trong Master SKU.
                                  </p>
                                </div>
                              </div>
                              <button 
                                onClick={() => downloadMissingSkusLog(missingSKUs, extractedOrders)}
                                className="w-full bg-amber-600 hover:bg-amber-700 text-white text-[9px] font-black uppercase py-2 px-3 rounded-xl transition-all flex items-center justify-center gap-1.5 cursor-pointer shadow-sm"
                              >
                                <Download className="w-3.5 h-3.5" /> TẢI FILE LOG SKU THIẾU
                              </button>
                            </div>
                          )}

                          <div className="space-y-3 pt-2">
                            {/* 1. EXCEL EXPORTER BUTTON */}
                            <motion.button
                              whileHover={{ scale: 1.02, y: -2 }}
                              whileTap={{ scale: 0.98 }}
                              onClick={() => exportCleanedExcelFromExtracted(sortedExtractedOrders, getCustomFileName("Tổng-Hợp", sortedExtractedOrders.length))}
                              disabled={isProcessingUnified}
                              className="w-full bg-emerald-600 hover:bg-emerald-700 text-white rounded-2xl py-4 px-4 font-black text-[12px] uppercase tracking-widest shadow-lg shadow-emerald-100 transition-all flex items-center justify-center gap-3 disabled:opacity-30 pointer-events-auto cursor-pointer"
                            >
                              <FileSpreadsheet className="w-5 h-5 shrink-0" />
                              TẢI EXCEL ĐÃ LỌC
                            </motion.button>

                            {/* 2. PDF EXPORTER BUTTON */}
                            <motion.button
                              whileHover={{ scale: 1.02, y: -2 }}
                              whileTap={{ scale: 0.98 }}
                              onClick={() => generateSortedPdfFromExtracted(unifiedPdfFile!, sortedExtractedOrders, getCustomFileName("Tổng-Hợp", sortedExtractedOrders.length))}
                              disabled={isProcessingUnified || !unifiedPdfFile}
                              className="w-full bg-indigo-600 hover:bg-indigo-700 text-white rounded-2xl py-4 px-4 font-black text-[12px] uppercase tracking-widest shadow-lg shadow-indigo-100 transition-all flex items-center justify-center gap-3 disabled:opacity-30 pointer-events-auto cursor-pointer"
                            >
                              <FileText className="w-5 h-5 shrink-0" />
                              TẢI PDF SẮP XẾP
                            </motion.button>
                          </div>

                          {unifiedStatus && (
                            <div className={`p-3 rounded-xl border flex items-center gap-3 ${
                              isProcessingUnified 
                              ? "bg-indigo-50/50 border-indigo-100/50" 
                              : "bg-emerald-50/50 border-emerald-100/50"
                            }`}>
                              {isProcessingUnified ? (
                                <Loader2 className="w-3.5 h-3.5 animate-spin text-indigo-600 shrink-0" />
                              ) : (
                                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                              )}
                              <span className={`text-[10px] font-black uppercase tracking-wider ${
                                isProcessingUnified ? "text-indigo-700" : "text-emerald-700"
                              }`}>{unifiedStatus}</span>
                            </div>
                          )}
                        </div>

                        {/* TÁCH QUẢN LÝ THEO TỈNH THÀNH */}
                        <div className="bg-[#fcfcff] border border-indigo-100/50 rounded-[2rem] p-6 space-y-4 shadow-sm flex flex-col justify-between h-full">
                          <h4 className="text-xs font-black text-indigo-950 uppercase tracking-widest flex items-center gap-1.5">
                            <MapPin className="w-4 h-4 text-emerald-600" />
                            TÁCH QUẢN LÝ THEO TỈNH THÀNH
                          </h4>
                          <p className="text-[11px] text-slate-400 leading-relaxed font-semibold">
                            Phân rã nhanh danh sách Excel lọc sạch & nhãn PDF được sắp xếp thành các tập tin độc lập theo 3 thành phố lớn:
                          </p>

                          <div className="space-y-3 pt-1">
                            {/* Đà Nẵng Column */}
                            <div className="p-3 bg-white border border-slate-100 rounded-2xl flex flex-col gap-2">
                              <div className="flex justify-between items-center">
                                <span className="text-[11px] font-black uppercase text-indigo-900 tracking-wider flex items-center gap-1">
                                  <span className="size-1.5 rounded-full bg-emerald-500 animate-pulse"></span> Đà Nẵng
                                </span>
                                <span className="bg-slate-100 font-mono text-[10px] font-bold px-2 py-0.5 rounded-full text-slate-600">
                                  {dnOrders.length} đơn
                                </span>
                              </div>
                              <div className="grid grid-cols-2 gap-2">
                                <button
                                  onClick={() => {
                                    const list = sortedExtractedOrders.filter(o => o.province === "Đà Nẵng");
                                    exportCleanedExcelFromExtracted(list, getCustomFileName("Đà-Nẵng", list.length));
                                  }}
                                  disabled={dnOrders.length === 0}
                                  className="bg-emerald-50 hover:bg-emerald-100/80 disabled:opacity-30 text-emerald-800 text-[10px] uppercase font-black tracking-wider py-2 px-2 rounded-xl border border-emerald-100/50 transition-all flex items-center justify-center gap-1 cursor-pointer"
                                >
                                  <FileSpreadsheet className="w-3.5 h-3.5" /> File Excel
                                </button>
                                <button
                                  onClick={() => {
                                    const list = sortedExtractedOrders.filter(o => o.province === "Đà Nẵng");
                                    generateSortedPdfFromExtracted(unifiedPdfFile!, list, getCustomFileName("Đà-Nẵng", list.length));
                                  }}
                                  disabled={dnOrders.length === 0 || !unifiedPdfFile}
                                  className="bg-indigo-50 hover:bg-indigo-100/80 disabled:opacity-30 text-indigo-800 text-[10px] uppercase font-black tracking-wider py-2 px-2 rounded-xl border border-indigo-100/50 transition-all flex items-center justify-center gap-1 cursor-pointer"
                                >
                                  <FileText className="w-3.5 h-3.5" /> File PDF
                                </button>
                              </div>
                            </div>

                            {/* Hà Nội Column */}
                            <div className="p-3 bg-white border border-slate-100 rounded-2xl flex flex-col gap-2">
                              <div className="flex justify-between items-center">
                                <span className="text-[11px] font-black uppercase text-indigo-900 tracking-wider flex items-center gap-1">
                                  <span className="size-1.5 rounded-full bg-orange-500 animate-pulse"></span> Hà Nội
                                </span>
                                <span className="bg-slate-100 font-mono text-[10px] font-bold px-2 py-0.5 rounded-full text-slate-600">
                                  {hnOrders.length} đơn
                                </span>
                              </div>
                              <div className="grid grid-cols-2 gap-2">
                                <button
                                  onClick={() => {
                                    const list = sortedExtractedOrders.filter(o => o.province === "Hà Nội");
                                    exportCleanedExcelFromExtracted(list, getCustomFileName("Hà-Nội", list.length));
                                  }}
                                  disabled={hnOrders.length === 0}
                                  className="bg-emerald-50 hover:bg-emerald-100/80 disabled:opacity-30 text-emerald-800 text-[10px] uppercase font-black tracking-wider py-2 px-2 rounded-xl border border-emerald-100/50 transition-all flex items-center justify-center gap-1 cursor-pointer"
                                >
                                  <FileSpreadsheet className="w-3.5 h-3.5" /> File Excel
                                </button>
                                <button
                                  onClick={() => {
                                    const list = sortedExtractedOrders.filter(o => o.province === "Hà Nội");
                                    generateSortedPdfFromExtracted(unifiedPdfFile!, list, getCustomFileName("Hà-Nội", list.length));
                                  }}
                                  disabled={hnOrders.length === 0 || !unifiedPdfFile}
                                  className="bg-indigo-50 hover:bg-indigo-100/80 disabled:opacity-30 text-indigo-800 text-[10px] uppercase font-black tracking-wider py-2 px-2 rounded-xl border border-indigo-100/50 transition-all flex items-center justify-center gap-1 cursor-pointer"
                                >
                                  <FileText className="w-3.5 h-3.5" /> File PDF
                                </button>
                              </div>
                            </div>

                            {/* Hồ Chí Minh Column */}
                            <div className="p-3 bg-white border border-slate-100 rounded-2xl flex flex-col gap-2">
                              <div className="flex justify-between items-center">
                                <span className="text-[11px] font-black uppercase text-indigo-900 tracking-wider flex items-center gap-1">
                                  <span className="size-1.5 rounded-full bg-blue-500 animate-pulse"></span> TP. Hồ Chí Minh
                                </span>
                                <span className="bg-slate-100 font-mono text-[10px] font-bold px-2 py-0.5 rounded-full text-slate-600">
                                  {hcmOrders.length} đơn
                                </span>
                              </div>
                              <div className="grid grid-cols-2 gap-2">
                                <button
                                  onClick={() => {
                                    const list = sortedExtractedOrders.filter(o => o.province === "Hồ Chí Minh");
                                    exportCleanedExcelFromExtracted(list, getCustomFileName("Hồ-Chí-Minh", list.length));
                                  }}
                                  disabled={hcmOrders.length === 0}
                                  className="bg-emerald-50 hover:bg-emerald-100/80 disabled:opacity-30 text-emerald-800 text-[10px] uppercase font-black tracking-wider py-2 px-2 rounded-xl border border-emerald-100/50 transition-all flex items-center justify-center gap-1 cursor-pointer"
                                >
                                  <FileSpreadsheet className="w-3.5 h-3.5" /> File Excel
                                </button>
                                <button
                                  onClick={() => {
                                    const list = sortedExtractedOrders.filter(o => o.province === "Hồ Chí Minh");
                                    generateSortedPdfFromExtracted(unifiedPdfFile!, list, getCustomFileName("Hồ-Chí-Minh", list.length));
                                  }}
                                  disabled={hcmOrders.length === 0 || !unifiedPdfFile}
                                  className="bg-indigo-50 hover:bg-indigo-100/80 disabled:opacity-30 text-indigo-800 text-[10px] uppercase font-black tracking-wider py-2 px-2 rounded-xl border border-indigo-100/50 transition-all flex items-center justify-center gap-1 cursor-pointer"
                                >
                                  <FileText className="w-3.5 h-3.5" /> File PDF
                                </button>
                              </div>
                            </div>
                          </div>

                          {/* Sequential Multi Download Button */}
                          <button
                            onClick={async () => {
                              setIsProcessingUnified(true);
                              setUnifiedProgress(10);
                              setUnifiedStatus("Vui lòng đợi bưu cục: Đang xuất & tải lần lượt các bưu cục...");
                              
                              try {
                                // Đà Nẵng
                                if (dnOrders.length > 0) {
                                  const list = sortedExtractedOrders.filter(o => o.province === "Đà Nẵng");
                                  await exportCleanedExcelFromExtracted(list, getCustomFileName("Đà-Nẵng", list.length));
                                  await new Promise(r => setTimeout(r, 600));
                                  await generateSortedPdfFromExtracted(unifiedPdfFile!, list, getCustomFileName("Đà-Nẵng", list.length));
                                  await new Promise(r => setTimeout(r, 600));
                                }
                                setUnifiedProgress(40);

                                // Hà Nội
                                if (hnOrders.length > 0) {
                                  const list = sortedExtractedOrders.filter(o => o.province === "Hà Nội");
                                  await exportCleanedExcelFromExtracted(list, getCustomFileName("Hà-Nội", list.length));
                                  await new Promise(r => setTimeout(r, 600));
                                  await generateSortedPdfFromExtracted(unifiedPdfFile!, list, getCustomFileName("Hà-Nội", list.length));
                                  await new Promise(r => setTimeout(r, 600));
                                }
                                setUnifiedProgress(70);

                                // Hồ Chí Minh
                                if (hcmOrders.length > 0) {
                                  const list = sortedExtractedOrders.filter(o => o.province === "Hồ Chí Minh");
                                  await exportCleanedExcelFromExtracted(list, getCustomFileName("Hồ-Chí-Minh", list.length));
                                  await new Promise(r => setTimeout(r, 600));
                                  await generateSortedPdfFromExtracted(unifiedPdfFile!, list, getCustomFileName("Hồ-Chí-Minh", list.length));
                                }

                                setUnifiedProgress(100);
                                setUnifiedStatus("Xuất toàn bộ bưu cục tỉnh thành hoàn tất!");
                              } catch (err: any) {
                                setError(`Lỗi xuất bưu cục lẻ: ${err.message}`);
                              } finally {
                                setIsProcessingUnified(false);
                              }
                            }}
                              disabled={isProcessingUnified || sortedExtractedOrders.length === 0}
                              className="w-full bg-emerald-600 hover:bg-emerald-700 text-white rounded-2xl py-3 px-4 font-black text-[10px] uppercase tracking-widest shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer disabled:opacity-30"
                            >
                              <Download className="w-4 h-4" /> TẢI ĐỒNG LOẠT TOÀN BỘ FILE LẺ
                            </button>
                          </div>
                        </div>

                      </div>

                      {/* Reset State Action */}
                      <button
                        onClick={resetAll}
                        className="w-full bg-slate-100 hover:bg-slate-200 text-slate-500 hover:text-slate-700 rounded-2xl py-4 font-black text-[11px] uppercase tracking-widest border border-slate-200/60 transition-all flex items-center justify-center gap-2 cursor-pointer"
                      >
                        <RefreshCcw className="w-4 h-4" /> TRỞ LẠI & BẮT ĐẦU LẠI
                      </button>

                    </div>
                )}
              </div>

              {/* === TAB: SKU_CONFIG (Cấu hình dữ liệu SKU) === */}
              <div className={cn("flex-1 flex flex-col", activeTab !== "sku_config" && "hidden")}>
                <SkuConfig 
                  masterSkus={masterSkus}
                  masterSkusMeta={masterSkusMeta}
                  onUpdate={(newDb, newMeta) => {
                    setMasterSkus(newDb);
                    setMasterSkusMeta(newMeta);
                  }}
                  onClear={() => {
                    setMasterSkus({});
                    setMasterSkusMeta(null);
                  }}
                />
              </div>
            </>
          );

          // Legacy helper container for historical views
          const renderLegacyPreviewBlocks = () => {
            return (
              <div className="hidden">

                      {/* Right Column: Previews Tables */}
                      <div className="lg:col-span-8 flex flex-col space-y-6">
                        
                        {/* Active Live Grid Table */}
                        <div className="bg-white rounded-[2rem] border border-slate-200 overflow-hidden flex flex-col min-h-[400px] shadow-sm flex-1">
                          <div className="bg-slate-50 border-b border-slate-200 px-6 py-4 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                            <div className="flex items-center gap-2.5">
                              <div className="bg-indigo-600 p-2 rounded-xl text-white shadow-md shadow-indigo-100">
                                <TableIcon className="w-4 h-4" />
                              </div>
                              <div>
                                <h3 className="text-xs font-black text-slate-800 uppercase tracking-widest">BẢNG XEM TRƯỚC ĐƠN HÀNG ({sortedExtractedOrders.length})</h3>
                                <p className="text-[9px] text-slate-400 font-bold uppercase tracking-wider">Trình bày chi tiết dữ liệu sắp kết nối từ file PDF gốc</p>
                              </div>
                            </div>

                            <div className="bg-white px-3 py-1.5 rounded-lg border border-slate-200 text-[10px] font-black text-indigo-600 uppercase tracking-wider shrink-0 flex items-center gap-1.5 shadow-sm">
                              <SlidersHorizontal className="w-3.5 h-3.5 text-indigo-500" />
                              Sắp xếp: {sortOption === 'province' ? "Theo Tỉnh thành" : sortOption === 'sku' ? "Theo SKU" : sortOption === 'qty' ? "Đơn lẻ trước" : "File gốc"}
                            </div>
                          </div>

                          <div className="flex-1 overflow-auto bg-[radial-gradient(#f1f5f9_1.2px,transparent_1.2px)] [background-size:24px_24px] max-h-[500px]">
                            <table className="w-full text-left border-collapse table-fixed">
                              <thead className="sticky top-0 bg-white shadow-[0_2px_10px_-2px_rgba(0,0,0,0.05)] z-10">
                                <tr>
                                  <th className="p-3.5 text-[10px] font-black text-slate-500 uppercase tracking-wider w-16 text-center border-b border-slate-200">STT</th>
                                  <th className="p-3.5 text-[10px] font-black text-slate-500 uppercase tracking-wider border-b border-slate-200 w-44">Mã Đơn Hàng</th>
                                  <th className="p-3.5 text-[10px] font-black text-slate-500 uppercase tracking-wider border-b border-slate-200">Sản phẩm (SKU)</th>
                                  <th className="p-3.5 text-[10px] font-black text-slate-500 uppercase tracking-wider border-b border-slate-200 w-32 text-center">Tỉnh thành gửi</th>
                                  <th className="p-3.5 text-[10px] font-black text-slate-500 uppercase tracking-wider border-b border-slate-200 w-24 text-center">Trang PDF</th>
                                </tr>
                              </thead>
                              <tbody className="text-xs">
                                {sortedExtractedOrders.map((order, indexNo) => (
                                  <tr 
                                    key={order.orderId + "-" + indexNo} 
                                    className={cn(
                                      "transition-all duration-200 hover:bg-indigo-50/20 border-b border-slate-100",
                                      indexNo % 2 === 0 ? "bg-white" : "bg-slate-50/20"
                                    )}
                                  >
                                    <td className="p-3.5 text-center font-mono font-black text-slate-400">
                                      {indexNo + 1}
                                    </td>
                                    <td className="p-3.5 font-mono font-bold text-slate-500 truncate select-all">
                                      {order.orderId}
                                    </td>
                                    <td className="p-3.5 space-y-1">
                                      {order.items.map((item, keyIdx) => (
                                        <div key={keyIdx} className="flex items-center justify-between text-[11px] leading-snug">
                                          <span className="text-slate-800 font-bold truncate max-w-[200px]" title={item.name}>{item.name}</span>
                                          <div className="flex items-center gap-1.5 ml-2">
                                            <span className="text-[10px] bg-slate-100 font-mono text-slate-500 px-1 border rounded">{item.sku}</span>
                                            <span className="font-extrabold text-indigo-600 font-mono">x{item.qty}</span>
                                          </div>
                                        </div>
                                      ))}
                                    </td>
                                    <td className="p-3.5 text-center">
                                      <span className={cn(
                                        "px-2.5 py-1 rounded-full text-[10px] font-black uppercase tracking-wider border",
                                        order.province === "Hồ Chí Minh" ? "bg-emerald-50 text-emerald-700 border-emerald-100" :
                                        order.province === "Hà Nội" ? "bg-sky-50 text-sky-700 border-sky-100" :
                                        order.province === "Đà Nẵng" ? "bg-purple-50 text-purple-700 border-purple-100" :
                                        "bg-slate-100 text-slate-600 border-slate-200"
                                      )}>
                                        {order.province}
                                      </span>
                                    </td>
                                    <td className="p-3.5 text-center">
                                      <span className="bg-slate-100 text-slate-700 text-[11px] font-mono font-black px-2.5 py-1 rounded-lg border border-slate-200">
                                        {String(order.originalPageIndex + 1).padStart(2, '0')}
                                      </span>
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        {/* Summary statistics layout block */}
                        <div className="bg-slate-50 border border-slate-200 rounded-[2rem] p-6 space-y-4 shadow-sm">
                          <div className="flex items-center gap-2">
                            <Boxes className="w-5 h-5 text-emerald-600" />
                            <h4 className="text-xs font-black text-slate-700 uppercase tracking-widest">
                              TỔNG HỢP NHANH SẢN PHẨM CẦN CHUẨN BỊ ({aggregateSummaryList.length} SKUs)
                            </h4>
                          </div>

                          <div className="grid grid-cols-1 md:grid-cols-2 gap-3 pt-1">
                            {aggregateSummaryList.slice(0, 6).map((item, keyID) => (
                              <div key={keyID} className="bg-white p-3.5 rounded-2xl border border-slate-200/60 flex items-center justify-between shadow-sm">
                                <div className="space-y-1 block truncate pr-3">
                                  <p className="text-[11px] font-bold text-slate-800 truncate" title={item.name}>{item.name}</p>
                                  <p className="text-[9px] font-mono text-slate-400">SKU: <span className="font-bold text-slate-600">{item.sku}</span></p>
                                </div>
                                <span className="text-sm font-black font-mono text-emerald-600 shrink-0 bg-emerald-50 px-3 py-1 rounded-lg border border-emerald-100">
                                  SL: {item.totalQty}
                                </span>
                              </div>
                            ))}
                          </div>

                          {aggregateSummaryList.length > 6 && (
                            <p className="text-[10px] text-slate-400 font-bold uppercase tracking-wide text-center pt-1">
                              ... và {aggregateSummaryList.length - 6} mã sản phẩm sỉ lẻ khác (đầy đủ trong File Excel đã xuất)
                            </p>
                          )}
                        </div>

                      </div>
              </div>
            );
          };

          // Legacy active tabs backward compatible selections
          const renderLegacyFallback = () => activeTab === 'pdf' ? (
            <>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 h-auto shrink-0">
                <motion.div 
                  whileHover={{ y: -4, boxShadow: "0 20px 25px -5px rgb(0 0 0 / 0.05), 0 8px 10px -6px rgb(0 0 0 / 0.05)" }}
                  className={cn(
                    "group bg-white border-2 border-dashed rounded-3xl p-8 flex flex-col items-center justify-center space-y-4 transition-all duration-300 relative",
                    excelFile ? "border-emerald-600 bg-emerald-50" : "border-slate-200 hover:border-indigo-400 hover:bg-indigo-50/5"
                  )}
                >
                  <input type="file" accept=".xlsx, .xls, .csv" onChange={handleExcelUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" ref={excelInputRef} />
                  <div className={cn("p-5 rounded-2xl transition-all duration-500 shadow-sm", excelFile ? "bg-emerald-100 text-emerald-700 scale-110 rotate-3 shadow-emerald-100" : "bg-slate-50 text-slate-400 group-hover:scale-110 group-hover:rotate-3 shadow-slate-100")}>
                    <FileSpreadsheet className="w-10 h-10" />
                  </div>
                  <div className="text-center">
                    <p className="text-sm font-black text-slate-800 uppercase tracking-wide">1. TẢI LÊN EXCEL / CSV</p>
                    <p className="text-[11px] text-slate-400 font-bold max-w-[240px] mt-2 leading-relaxed">
                      {excelFile ? <span className="text-emerald-700 italic font-black truncate block">{excelFile.name}</span> : "Kéo thả file .xlsx, .csv (Tối đa 100MB)"}
                    </p>
                  </div>
                  {excelOrders.length > 0 && (
                    <motion.span 
                      initial={{ scale: 0.8 }} animate={{ scale: 1 }}
                      className="text-[10px] bg-emerald-500 text-white px-4 py-1.5 rounded-full font-black uppercase tracking-widest shadow-lg shadow-emerald-200"
                    >
                      ĐÃ NHẬN {excelOrders.length} Dòng
                    </motion.span>
                  )}
                </motion.div>

                <motion.div 
                  whileHover={{ y: -4, boxShadow: "0 20px 25px -5px rgb(0 0 0 / 0.05), 0 8px 10px -6px rgb(0 0 0 / 0.05)" }}
                  className={cn(
                    "group bg-white border-2 border-dashed rounded-3xl p-8 flex flex-col items-center justify-center space-y-4 transition-all duration-300 relative",
                    pdfFile ? "border-rose-500 bg-rose-50" : "border-slate-200 hover:border-indigo-400 hover:bg-indigo-50/5"
                  )}
                >
                  <input type="file" accept=".pdf" onChange={handlePdfUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" ref={pdfInputRef} />
                  <div className={cn("p-5 rounded-2xl transition-all duration-500 shadow-sm", pdfFile ? "bg-rose-100 text-rose-600 scale-110 -rotate-3 shadow-rose-100" : "bg-slate-50 text-slate-400 group-hover:scale-110 group-hover:-rotate-3 shadow-slate-100")}>
                    <FileText className="w-10 h-10" />
                  </div>
                  <div className="text-center">
                    <p className="text-sm font-black text-slate-800 uppercase tracking-wide">2. TẢI LÊN PDF GỐC</p>
                    <p className="text-[11px] text-slate-400 font-bold max-w-[240px] mt-2 leading-relaxed">
                      {pdfFile ? <span className="text-rose-600 italic font-black truncate block">{pdfFile.name}</span> : "Kéo thả file .pdf (Tối đa 100MB)"}
                    </p>
                  </div>
                  {pdfPages.length > 0 && (
                    <motion.span 
                      initial={{ scale: 0.8 }} animate={{ scale: 1 }}
                      className="text-[10px] bg-indigo-500 text-white px-4 py-1.5 rounded-full font-black uppercase tracking-widest shadow-lg shadow-indigo-200"
                    >
                      ĐÃ NHẬN {pdfPages.length} TRANG GỐC
                    </motion.span>
                  )}
                </motion.div>
              </div>

              {(excelOrders.length > 0 && pdfFile && step === 1) && (
                <motion.div initial={{ opacity: 0, scale: 0.9 }} animate={{ opacity: 1, scale: 1 }} className="flex justify-center py-4">
                  <motion.button 
                    whileHover={{ scale: 1.02, y: -2, boxShadow: "0 25px 50px -12px rgba(99, 102, 241, 0.25)" }}
                    whileTap={{ scale: 0.98 }}
                    onClick={processFiles}
                    disabled={isProcessing}
                    className="bg-indigo-600 text-white px-10 py-5 rounded-2xl font-black text-sm transition-all shadow-xl flex items-center gap-4 uppercase tracking-widest hover:bg-indigo-700 disabled:opacity-50 disabled:grayscale"
                  >
                    {isProcessing ? <Loader2 className="w-5 h-5 animate-spin" /> : <Search className="w-5 h-5" />}
                    BẮT ĐẦU SẮP XẾP FILE PDF
                  </motion.button>
                </motion.div>
              )}

              <div className="bg-slate-50/50 rounded-2xl p-6 shrink-0 border border-slate-100 space-y-4">
                <div className="flex flex-col lg:flex-row justify-between items-stretch lg:items-center gap-6">
                  <div className="flex-1 space-y-2">
                    <div className="flex justify-between items-end">
                      <div className="flex items-center gap-3">
                        <div className="p-2 bg-indigo-100 rounded-lg">
                          {isProcessing ? <Loader2 className="w-4 h-4 animate-spin text-indigo-600" /> : <div className="w-4 h-4 rounded-full bg-indigo-600" />}
                        </div>
                        <p className="text-[11px] font-black text-slate-400 uppercase tracking-[0.1em]">
                          {statusMessage || (step === 2 ? "Hoàn tất đối chiếu" : "Chờ bắt đầu")}
                        </p>
                      </div>
                      <span className="text-xl font-black text-indigo-600 font-mono tracking-tighter">{progress}%</span>
                    </div>
                    <div className="w-full bg-white h-3.5 rounded-full overflow-hidden border border-slate-200 shadow-sm p-0.5">
                      <motion.div 
                        className="bg-gradient-to-r from-indigo-500 to-indigo-700 h-full rounded-full shadow-[0_0_10px_rgba(99,102,241,0.5)]"
                        initial={{ width: 0 }}
                        animate={{ width: `${progress}%` }}
                        transition={{ duration: 0.5, ease: "easeOut" }}
                      />
                    </div>
                  </div>
                  
                  {step === 2 && (
                    <div className="grid grid-cols-2 lg:flex items-center gap-4 lg:gap-8 lg:border-l lg:border-slate-200 lg:pl-8">
                      <div className="bg-white p-3 rounded-2xl border border-slate-100 shadow-sm min-w-[100px]">
                        <p className="text-[9px] uppercase font-black text-slate-400 tracking-wider mb-1">Số dòng Excel</p>
                        <p className="text-xl font-black text-slate-800 tracking-tighter">{stats.uniqueOrders}</p>
                      </div>
                      <div className="bg-emerald-50 p-3 rounded-2xl border border-emerald-100 shadow-sm min-w-[100px]">
                        <p className="text-[9px] uppercase font-black text-emerald-600 tracking-wider mb-1">Đã khớp</p>
                        <p className="text-xl font-black text-emerald-600 tracking-tighter">{stats.matchedCount}</p>
                      </div>
                      <div className="bg-rose-50 p-3 rounded-2xl border border-rose-100 shadow-sm min-w-[100px]">
                        <p className="text-[9px] uppercase font-black text-rose-500 tracking-wider mb-1">Thiếu PDF</p>
                        <p className="text-xl font-black text-rose-500 tracking-tighter">{stats.missingInPdfCount}</p>
                      </div>
                      <div className="bg-amber-50 p-3 rounded-2xl border border-amber-100 shadow-sm min-w-[100px]">
                        <p className="text-[9px] uppercase font-black text-amber-500 tracking-wider mb-1">Lỗi / Dư</p>
                        <p className="text-xl font-black text-amber-500 tracking-tighter">{stats.errorPdfCount + stats.extraInPdfCount}</p>
                      </div>
                    </div>
                  )}
                </div>
              </div>

              <div className="flex-1 bg-white rounded-3xl border border-slate-200 overflow-hidden flex flex-col min-h-[400px] shadow-sm">
                <div className="bg-slate-50 border-b border-slate-200 px-6 py-4 flex justify-between items-center shrink-0">
                  <div className="flex items-center gap-3">
                    <div className="bg-indigo-600 p-1.5 rounded-lg">
                      <TableIcon className="w-4 h-4 text-white" />
                    </div>
                    <h3 className="text-[11px] font-black text-slate-700 uppercase tracking-[0.1em]">
                      Bảng Đối Chiếu Dữ Liệu
                    </h3>
                  </div>
                </div>
                <div className="flex-1 overflow-auto bg-[radial-gradient(#f1f5f9_1px,transparent_1px)] [background-size:24px_24px]">
                  {step === 1 ? (
                    <div className="h-full flex flex-col items-center justify-center p-12 text-center space-y-6">
                      <div className="w-20 h-20 bg-slate-50 rounded-3xl flex items-center justify-center border border-slate-100">
                        <Search className="w-10 h-10 text-slate-200" />
                      </div>
                      <div className="space-y-2">
                        <p className="text-xs font-black text-slate-400 uppercase tracking-widest">Đang chờ cấu trúc file...</p>
                        <p className="text-[10px] text-slate-300 font-bold max-w-[200px] mx-auto uppercase">Vui lòng tải lên cả file Excel và PDF để bắt đầu</p>
                      </div>
                    </div>
                  ) : (
                    <table className="w-full text-left border-collapse table-fixed">
                      <thead className="sticky top-0 bg-white shadow-[0_2px_10px_-2px_rgba(0,0,0,0.05)] z-20">
                        <tr className="bg-slate-50/50 backdrop-blur-md">
                          <th className="p-4 text-[10px] font-black text-slate-500 uppercase w-20 text-center border-b border-slate-200">#</th>
                          <th className="p-4 text-[10px] font-black text-slate-500 uppercase border-b border-slate-200">Mã Đơn Hàng (Data Node)</th>
                          <th className="p-4 text-[10px] font-black text-slate-500 uppercase w-32 text-center border-b border-slate-200">Trang PDF</th>
                          <th className="p-4 text-[10px] font-black text-slate-500 uppercase border-b border-slate-200">Trạng Thái Kết Nối</th>
                        </tr>
                      </thead>
                      <tbody className="text-xs">
                        {matchResults.map((result, i) => (
                          <tr key={i} className={cn(
                            "transition-all duration-200 group",
                            i % 2 === 0 ? "bg-white" : "bg-slate-50/30",
                            "hover:bg-indigo-50/30"
                          )}>
                            <td className="p-4 text-center font-mono font-bold text-slate-400 group-hover:text-indigo-400">
                              {String(i + 1).padStart(2, '0')}
                            </td>
                            <td className="p-4 font-mono font-bold text-slate-700 group-hover:translate-x-1 transition-transform">
                              {result.orderId}
                            </td>
                            <td className="p-4 text-center">
                              {result.pdfPageIndex !== null ? (
                                <span className="bg-slate-100 text-slate-700 px-3 py-1 rounded-lg font-mono font-bold border border-slate-200">
                                  {String(result.pdfPageIndex + 1).padStart(2, '0')}
                                </span>
                              ) : (
                                <span className="text-slate-300 font-mono font-bold">--</span>
                              )}
                            </td>
                            <td className="p-4">
                              {result.status === 'matched' ? (
                                <div className="inline-flex items-center gap-2">
                                  <div className="w-1.5 h-1.5 rounded-full bg-emerald-500"></div>
                                  <span className="text-[10px] font-black text-emerald-600 uppercase tracking-widest">Đã khớp</span>
                                </div>
                              ) : result.status === 'duplicate_in_pdf' ? (
                                <div className="inline-flex items-center gap-2">
                                  <div className="w-1.5 h-1.5 rounded-full bg-amber-500 animate-pulse"></div>
                                  <span className="text-[10px] font-black text-amber-600 uppercase tracking-widest">Data Trùng</span>
                                </div>
                              ) : (
                                <div className="inline-flex items-center gap-2">
                                  <div className="w-1.5 h-1.5 rounded-full bg-rose-500"></div>
                                  <span className="text-[10px] font-black text-rose-600 uppercase tracking-widest">Không tìm thấy</span>
                                </div>
                              )}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>

              <div className="flex flex-col sm:flex-row justify-between items-center bg-slate-50 p-6 rounded-3xl border border-slate-100 gap-6 mt-auto">
                <div className="flex gap-4 w-full sm:w-auto">
                  <motion.button 
                    whileHover={{ scale: 1.02, backgroundColor: "#f8fafc" }} 
                    whileTap={{ scale: 0.98 }}
                    onClick={resetAll} 
                    className="flex-1 sm:flex-none flex items-center justify-center gap-3 bg-white text-slate-500 px-8 py-4 rounded-2xl font-black text-[11px] border border-slate-200 transition-all uppercase tracking-widest shadow-sm hover:text-slate-700"
                  >
                    <RefreshCcw className="w-4 h-4" /> Làm lại
                  </motion.button>
                </div>
                
                <div className="flex items-center gap-6 w-full sm:w-auto">
                  {step === 2 && (
                    <>
                      <div className="text-right hidden md:block">
                        <p className={cn("text-[10px] font-black uppercase tracking-widest mb-1", stats.missingInPdfCount > 0 ? "text-rose-500" : "text-emerald-500")}>
                          {stats.missingInPdfCount > 0 ? `${stats.missingInPdfCount} ĐƠN KHÔNG TÌM THẤY` : "ĐỒNG BỘ DỮ LIỆU HOÀN TẤT"}
                        </p>
                        <p className="text-[9px] text-slate-400 uppercase font-bold tracking-tight">Sắp xếp PDF tự động theo trình tự Excel</p>
                      </div>
                      <motion.button 
                        whileHover={{ scale: 1.02, y: -2, boxShadow: "0 25px 50px -12px rgba(79, 70, 229, 0.25)" }}
                        whileTap={{ scale: 0.98 }}
                        onClick={generateSortedPdf}
                        disabled={stats.matchedCount === 0 || isProcessing}
                        className="flex-1 sm:flex-none bg-indigo-600 text-white px-12 py-4 rounded-2xl font-black text-xs shadow-xl shadow-indigo-100 flex items-center justify-center gap-4 uppercase tracking-widest hover:bg-indigo-700 transition-all disabled:opacity-30"
                      >
                        {isProcessing ? <Loader2 className="w-5 h-5 animate-spin" /> : <Download className="w-5 h-5" />}
                        XUẤT PDF
                      </motion.button>
                    </>
                  )}
                </div>
              </div>
            </>
          ) : (
            <div className="flex-1 flex flex-col space-y-6">
               {/* Excel Refiner Tool UI */}
               <motion.div 
                  initial={{ opacity: 0, scale: 0.98 }}
                  animate={{ opacity: 1, scale: 1 }}
                  className="bg-white rounded-[2rem] border border-slate-200 p-12 flex flex-col items-center justify-center text-center space-y-10 shadow-sm"
                >
                  <div className={cn(
                    "group relative border-2 border-dashed rounded-[2.5rem] p-16 w-full max-w-2xl transition-all duration-500 shadow-inner",
                    rawExcelFile ? "border-emerald-600 bg-emerald-50" : "border-slate-100 hover:border-indigo-200 bg-slate-50/30"
                  )}>
                    <input type="file" accept=".xlsx, .xls, .csv" onChange={handleRawExcelUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" ref={excelRefinerInputRef} />
                    <div className={cn("size-24 rounded-3xl mx-auto mb-8 flex items-center justify-center shadow-lg transition-all duration-500", rawExcelFile ? "bg-emerald-600 text-white shadow-emerald-200 scale-110 rotate-6" : "bg-white text-slate-200 group-hover:scale-110")}>
                      <FileSpreadsheet className="w-12 h-12" />
                    </div>
                    <h2 className="text-2xl font-black text-slate-800 uppercase tracking-tight">
                      TẢI LÊN FILE EXCEL THÔ
                    </h2>
                    <p className="text-sm text-slate-400 mt-3 font-medium max-w-sm mx-auto leading-relaxed">
                      Kéo thả file .xlsx, .csv (Tối đa 100MB)
                    </p>
                    
                    {rawExcelFile && (
                      <motion.div 
                        initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}
                        className="mt-8 px-6 py-3 bg-white rounded-2xl shadow-sm border border-indigo-100 inline-flex items-center gap-3"
                      >
                        <div className="size-2 rounded-full bg-emerald-500 animate-pulse"></div>
                        <span className="text-[11px] font-black text-slate-700 uppercase tracking-widest truncate max-w-[200px]">{rawExcelFile.name}</span>
                      </motion.div>
                    )}
                  </div>

                  {rawExcelData.length > 0 && (
                    <motion.div 
                      initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }}
                      className="w-full max-w-md bg-emerald-50/50 p-8 rounded-[2rem] border border-emerald-100 flex items-center gap-6 text-left"
                    >
                      <div className="bg-emerald-100 p-4 rounded-2xl shrink-0">
                        <CheckCircle2 className="w-8 h-8 text-emerald-600" />
                      </div>
                      <div>
                        <h3 className="text-[11px] font-black uppercase text-emerald-700 tracking-widest mb-1">XỬ LÝ FILE EXCEL THÀNH CÔNG</h3>
                        <p className="text-lg font-black text-emerald-950 tracking-tighter leading-tight">{rawExcelData.length} Bản ghi dữ liệu</p>
                        <p className="text-[10px] text-emerald-600 font-bold uppercase mt-1">Sẵn sàng xuất file Standard</p>
                      </div>
                    </motion.div>
                  )}

                  <div className="flex flex-wrap justify-center gap-4">
                    {rawExcelFile && (
                      <motion.button 
                        whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                        onClick={resetAll} 
                        className="px-10 py-5 rounded-2xl bg-slate-50 text-slate-400 font-black text-[11px] uppercase tracking-widest border border-slate-100 hover:bg-white hover:text-slate-600 transition-all"
                      >
                        Làm lại
                      </motion.button>
                    )}
                    <motion.button 
                      whileHover={{ scale: 1.02, y: -4, boxShadow: "0 25px 50px -12px rgba(79, 70, 229, 0.25)" }}
                      whileTap={{ scale: 0.98 }}
                      onClick={exportCleanedExcel}
                      disabled={!rawExcelFile || isProcessing}
                      className="px-12 py-5 rounded-2xl bg-indigo-600 text-white font-black text-sm uppercase tracking-widest shadow-xl shadow-indigo-50 hover:bg-indigo-700 disabled:opacity-30 disabled:grayscale transition-all flex items-center gap-4"
                    >
                      {isProcessing ? <Loader2 className="w-5 h-5 animate-spin" /> : <Download className="w-5 h-5" />}
                      Tải file chuẩn
                    </motion.button>
                  </div>
                </motion.div>
            </div>
          );
        ;
        })()}

        <AnimatePresence>
          {error && (
            <motion.div initial={{ y: 50, opacity: 0 }} animate={{ y: 0, opacity: 1 }} exit={{ y: 50, opacity: 0 }} className="fixed bottom-6 left-6 right-6 z-50 p-4 bg-rose-600 text-white rounded-2xl shadow-xl flex items-center justify-between border border-white/20">
              <div className="flex items-center gap-3">
                <AlertTriangle className="w-6 h-6 flex-shrink-0" />
                <p className="text-xs font-bold">{error}</p>
              </div>
              <button onClick={() => setError(null)} className="p-1 hover:bg-white/10 rounded-full">×</button>
            </motion.div>
          )}
        </AnimatePresence>
      </main>
    </div>
  );
}
