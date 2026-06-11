
export interface OrderInfo {
  stt: number;
  orderId: string;
  sourceRow: number;
}

export interface PDFPageInfo {
  pageIndex: number; // 0-indexed
  orderId: string | null;
  status: 'matched' | 'not_found' | 'duplicate' | 'extra' | 'error';
  errorMessage?: string;
}

export interface MatchResult {
  excelIndex: number;
  orderId: string;
  pdfPageIndex: number | null;
  status: 'matched' | 'missing_in_pdf' | 'duplicate_in_pdf' | 'error_pdf';
}

export interface ProcessingStats {
  totalExcelRows: number;
  uniqueOrders: number;
  totalPdfPages: number;
  matchedCount: number;
  missingInPdfCount: number;
  extraInPdfCount: number;
  errorPdfCount: number;
}

export interface ExtractedOrderItem {
  name: string;
  sku: string;
  qty: number;
}

export interface ExtractedOrder {
  stt: number;
  orderId: string;
  province: string;
  originalPageIndex: number;
  items: ExtractedOrderItem[];
}
