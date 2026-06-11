import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  FileSpreadsheet, 
  Upload, 
  Database, 
  Search, 
  Trash2, 
  CheckCircle2, 
  AlertCircle, 
  RefreshCcw, 
  FileText, 
  X,
  HelpCircle
} from 'lucide-react';
import { motion } from 'motion/react';

interface SkuConfigProps {
  masterSkus: Record<string, string>;
  masterSkusMeta: {
    totalCount: number;
    lastUpdated: string;
    isMasterMarked: boolean;
  } | null;
  onUpdate: (newDb: Record<string, string>, newMeta: any) => void;
  onClear: () => void;
}

export default function SkuConfig({ masterSkus, masterSkusMeta, onUpdate, onClear }: SkuConfigProps) {
  const [file, setFile] = useState<File | null>(null);
  const [isParsing, setIsParsing] = useState(false);
  const [parseError, setParseError] = useState<string | null>(null);
  const [parsedData, setParsedData] = useState<{ sku: string; name: string }[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 8;

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Simulation of merging
  const mergeStats = useMemo(() => {
    if (parsedData.length === 0) return null;

    let added = 0;
    let updated = 0;
    let unchanged = 0;

    parsedData.forEach(item => {
      const skuKey = item.sku.trim().toUpperCase();
      const existingName = masterSkus[skuKey];

      if (existingName === undefined) {
        added++;
      } else if (existingName !== item.name.trim()) {
        updated++;
      } else {
        unchanged++;
      }
    });

    return { added, updated, unchanged, total: parsedData.length };
  }, [parsedData, masterSkus]);

  // Handle uploading and parsing the Excel/CSV file
  const handleUploadClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (!selectedFile) return;

    if (selectedFile.size > 100 * 1024 * 1024) {
      setParseError("File SKU không được vượt quá 100MB.");
      return;
    }

    setFile(selectedFile);
    setParseError(null);
    setIsParsing(true);
    setParsedData([]);

    try {
      const data = await selectedFile.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      
      if (workbook.SheetNames.length === 0) {
        throw new Error("File Excel không có sheet nào.");
      }

      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json<any>(worksheet, { header: 'A', raw: false });

      if (jsonData.length === 0) {
        throw new Error("File Excel trống hoặc không đọc được dữ liệu.");
      }

      // Advanced detection of columns for SKU and Name
      let skuCol: string | null = null;
      let nameCol: string | null = null;
      let headerRowIdx = -1;

      const skuKeywords = ["mã sku", "sku", "mã hàng", "mã sp", "mã sản phẩm", "seller sku", "mã_sku"];
      const nameKeywords = ["tên sản phẩm", "product name", "tên hàng", "tên sp", "sản phẩm", "tên_sản_phẩm"];

      // Scan first 30 rows for headers
      for (let i = 0; i < Math.min(jsonData.length, 30); i++) {
        const row = jsonData[i];
        for (const key in row) {
          const val = String(row[key] || '').toLowerCase().trim();
          if (!skuCol && skuKeywords.some(kw => val === kw || val.includes(kw))) {
            skuCol = key;
            headerRowIdx = i;
          }
          if (!nameCol && nameKeywords.some(kw => val === kw || val.includes(kw))) {
            nameCol = key;
            headerRowIdx = i;
          }
        }
        if (skuCol && nameCol) {
          headerRowIdx = i;
          break;
        }
      }

      // If one of them is missing, fallback to smart automated column detection based on text size characteristics
      if (!skuCol || !nameCol) {
        const firstRow = jsonData[0] || {};
        const keys = Object.keys(firstRow).filter(k => k !== '__rowNum__');
        
        if (keys.length >= 2) {
          if (skuCol && !nameCol) {
            nameCol = keys.find(k => k !== skuCol) || keys[1];
          } else if (nameCol && !skuCol) {
            skuCol = keys.find(k => k !== nameCol) || keys[0];
          } else {
            // Both are null. Analyze the first few rows to see which one has longer average strings (Name) and shorter (SKU/barcode)
            const colLengths: Record<string, number> = {};
            keys.forEach(k => { colLengths[k] = 0; });
            
            let sampleRows = Math.min(jsonData.length, 10);
            for (let i = 0; i < sampleRows; i++) {
              const r = jsonData[i];
              if (r) {
                keys.forEach(k => {
                  colLengths[k] += String(r[k] || '').trim().length;
                });
              }
            }
            
            const sortedKeys = [...keys].sort((a, b) => colLengths[a] - colLengths[b]);
            skuCol = sortedKeys[0];
            nameCol = sortedKeys[1];
            headerRowIdx = -1; // process from row 0
          }
        } else if (keys.length === 1) {
          skuCol = keys[0];
          nameCol = keys[0];
          headerRowIdx = -1;
        }
      }

      const tempRows: { sku: string; name: string }[] = [];
      const startIdx = headerRowIdx + 1;

      for (let i = startIdx; i < jsonData.length; i++) {
        const row = jsonData[i];
        const rawSku = row[skuCol];
        const rawName = row[nameCol];

        if (rawSku !== undefined && rawName !== undefined) {
          const skuStr = String(rawSku).trim();
          const nameStr = String(rawName).trim();
          if (skuStr && nameStr && !skuKeywords.includes(skuStr.toLowerCase()) && !nameKeywords.includes(nameStr.toLowerCase())) {
            tempRows.push({
              sku: skuStr,
              name: nameStr
            });
          }
        }
      }

      if (tempRows.length === 0) {
        // Fallback guess other way around
        throw new Error("Không thể trích xuất được dòng dữ liệu SKU hợp lệ. Vui lòng kiểm tra lại cấu trúc file!");
      }

      setParsedData(tempRows);
    } catch (err: any) {
      setParseError(err.message || "Lỗi đọc file Excel.");
      setFile(null);
    } finally {
      setIsParsing(false);
    }
  };

  // Merge simulated stats to Master SKU DB
  const handleSaveAsDefault = () => {
    if (parsedData.length === 0) return;

    const mergedDb = { ...masterSkus };
    
    parsedData.forEach(item => {
      const skuKey = item.sku.trim().toUpperCase();
      const nameVal = item.name.trim();
      mergedDb[skuKey] = nameVal; // add or update
    });

    const newSize = Object.keys(mergedDb).length;
    const meta = {
      totalCount: newSize,
      lastUpdated: new Date().toLocaleString('vi-VN'),
      isMasterMarked: true
    };

    // Save to LocalStorage
    localStorage.setItem('MASTER_SKU_DB', JSON.stringify(mergedDb));
    localStorage.setItem('MASTER_SKU_META', JSON.stringify(meta));

    // Callback
    onUpdate(mergedDb, meta);

    // Reset UI File Upload state
    setFile(null);
    setParsedData([]);
  };

  const handleDeleteMasterData = () => {
    const isConfirmed = window.confirm("Bạn có chắc chắn muốn xóa toàn bộ dữ liệu MASTER SKU hiện có? Thao tác này không thể hoàn tác.");
    if (!isConfirmed) return;

    localStorage.removeItem('MASTER_SKU_DB');
    localStorage.removeItem('MASTER_SKU_META');
    onClear();
    setFile(null);
    setParsedData([]);
  };

  // Filter current master lists with query
  const filteredMasterItems = useMemo(() => {
    const query = searchQuery.trim().toLowerCase();
    const items = Object.entries(masterSkus).map(([sku, name]) => ({ sku, name }));
    if (!query) return items;
    return items.filter(item => 
      item.sku.toLowerCase().includes(query) || 
      item.name.toLowerCase().includes(query)
    );
  }, [masterSkus, searchQuery]);

  // Paginated Master items
  const paginatedItems = useMemo(() => {
    const startIdx = (currentPage - 1) * itemsPerPage;
    return filteredMasterItems.slice(startIdx, startIdx + itemsPerPage);
  }, [filteredMasterItems, currentPage]);

  const totalPages = Math.ceil(filteredMasterItems.length / itemsPerPage);

  const resetUploadState = () => {
    setFile(null);
    setParsedData([]);
    setParseError(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  return (
    <div className="space-y-6">
      {/* Title */}
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 border-b border-slate-200 pb-5">
        <div>
          <h2 className="text-xl font-black text-slate-800 uppercase tracking-tight flex items-center gap-2">
            <Database className="w-5 h-5 text-indigo-600 animate-pulse" />
            CẤU HÌNH DỮ LIỆU SKU MASTER
          </h2>
          <p className="text-xs text-slate-400 font-bold uppercase mt-1 leading-normal tracking-wide">
            Lưu trữ danh mục SKU vĩnh viễn trên trình duyệt. Không cần upload lại mỗi khi trích xuất PDF.
          </p>
        </div>

        {masterSkusMeta && (
          <div className="flex items-center gap-3">
            <span className="text-[10px] uppercase font-black tracking-widest bg-emerald-50 text-emerald-700 px-3 py-1 rounded-full border border-emerald-100 flex items-center gap-1.5 shadow-sm">
              <CheckCircle2 className="w-3.5 h-3.5" />
              MASTER SKU
            </span>
            <button 
              onClick={handleDeleteMasterData}
              className="text-rose-600 hover:text-rose-700 text-xs font-black uppercase flex items-center gap-1 border border-rose-200 hover:bg-rose-50/50 px-3 py-1.5 rounded-xl transition-all cursor-pointer shadow-sm"
            >
              <Trash2 className="w-3.5 h-3.5" /> Xóa dữ liệu
            </button>
          </div>
        )}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-stretch">
        
        {/* Left Card - Upload & Merge Config */}
        <div className="lg:col-span-5 flex flex-col space-y-6">
          <div className="bg-white border border-slate-200 rounded-[2rem] p-6 space-y-6 shadow-sm flex flex-col">
            <h3 className="text-xs font-black text-slate-700 uppercase tracking-widest flex items-center gap-2">
              <Upload className="w-4 h-4 text-indigo-600" />
              CẬP NHẬT / TẢI LÊN MỚI
            </h3>

            {/* Drag and drop input */}
            <div 
              onClick={handleUploadClick}
              className={`border-2 border-dashed rounded-[1.5rem] p-8 text-center cursor-pointer transition-all duration-300 relative group flex flex-col items-center justify-center space-y-3 shadow-inner ${
                file ? "border-emerald-500 bg-emerald-50/10" : "border-slate-200 hover:border-indigo-400 bg-slate-50/30"
              }`}
            >
              <input 
                type="file" 
                ref={fileInputRef} 
                onChange={handleFileChange} 
                accept=".xlsx, .xls, .csv" 
                className="hidden" 
              />
              <div className={`p-4 rounded-2xl transition-all duration-500 shadow-sm ${
                file ? "bg-emerald-100 text-emerald-700 scale-110" : "bg-white text-slate-400 group-hover:scale-110"
              }`}>
                <FileSpreadsheet className="w-8 h-8" />
              </div>
              <div>
                <p className="text-xs font-black text-slate-800 uppercase tracking-wide">CHỌN FILE DỮ LIỆU SKU</p>
                <p className="text-[10px] text-slate-400 font-bold mt-1 uppercase max-w-[200px] leading-tight">
                  {file ? file.name : "Kéo thả file .xlsx, .csv (Tối đa 100MB)"}
                </p>
              </div>
              
              {file && (
                <button 
                  onClick={(e) => { e.stopPropagation(); resetUploadState(); }}
                  className="absolute top-2 right-2 p-1 bg-slate-200 hover:bg-slate-300 text-slate-500 hover:text-slate-700 rounded-full cursor-pointer"
                >
                  <X className="w-3 h-3" />
                </button>
              )}
            </div>

            {isParsing && (
              <div className="bg-slate-50 p-4 rounded-2xl flex items-center justify-center gap-3 text-xs text-indigo-600 font-bold uppercase tracking-wider">
                <RefreshCcw className="w-5 h-5 animate-spin" />
                Đang xử lý phân tích file...
              </div>
            )}

            {parseError && (
              <div className="bg-rose-50 border border-rose-100 p-4 rounded-2xl text-xs text-rose-600 font-bold flex items-center gap-2.5">
                <AlertCircle className="w-5 h-5 shrink-0" />
                <span>{parseError}</span>
              </div>
            )}

            {/* Merge Simulator Stats */}
            {mergeStats && (
              <motion.div 
                initial={{ opacity: 0, y: 15 }} 
                animate={{ opacity: 1, y: 0 }}
                className="bg-[#fcfcff] border border-indigo-100 rounded-2xl p-5 space-y-4 shadow-sm"
              >
                <div>
                  <h4 className="text-[11px] font-black uppercase text-indigo-950 tracking-wider">KẾT QUẢ PHÂN TÍCH SO SÁNH</h4>
                  <p className="text-[10px] text-slate-400 font-bold">Quy tắc thông minh: Tránh ghi đè toàn bộ dữ liệu mẫu cũ</p>
                </div>

                <div className="grid grid-cols-3 gap-2">
                  <div className="bg-emerald-50 border border-emerald-100/55 p-3 rounded-xl text-center">
                    <p className="text-[9px] uppercase font-black text-emerald-600 tracking-wide mb-1">ĐÃ THÊM MỚI</p>
                    <p className="text-lg font-black text-emerald-700 tracking-tighter">+{mergeStats.added}</p>
                    <p className="text-[8px] text-emerald-400 font-bold uppercase tracking-widest">SKU mới</p>
                  </div>
                  <div className="bg-sky-50 border border-sky-100/55 p-3 rounded-xl text-center">
                    <p className="text-[9px] uppercase font-black text-sky-600 tracking-wide mb-1">CẬP NHẬT TÊN</p>
                    <p className="text-lg font-black text-sky-700 tracking-tighter">{mergeStats.updated}</p>
                    <p className="text-[8px] text-sky-400 font-bold uppercase tracking-widest">SKU trùng</p>
                  </div>
                  <div className="bg-slate-50 border border-slate-100/55 p-3 rounded-xl text-center">
                    <p className="text-[9px] uppercase font-black text-slate-500 tracking-wide mb-1">GIỮ NGUYÊN</p>
                    <p className="text-lg font-black text-slate-600 tracking-tighter">{mergeStats.unchanged}</p>
                    <p className="text-[8px] text-slate-400 font-bold uppercase tracking-widest">Không đổi</p>
                  </div>
                </div>

                <button 
                  onClick={handleSaveAsDefault}
                  className="w-full bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl py-3 px-4 font-black text-[11px] uppercase tracking-wider transition-all flex items-center justify-center gap-2 cursor-pointer shadow-md shadow-indigo-100"
                >
                  <CheckCircle2 className="w-4 h-4" /> LƯU LÀM DỮ LIỆU MẶC ĐỊNH
                </button>
              </motion.div>
            )}

            {/* Instruction block */}
            <div className="mt-auto bg-slate-50 border border-slate-200/60 p-4 rounded-2xl flex items-start gap-3">
              <HelpCircle className="w-4 h-5 text-indigo-500 shrink-0 mt-0.5" />
              <div className="space-y-1">
                <p className="text-[10px] font-black text-slate-700 uppercase tracking-wider">HƯỚNG DẪN ĐỊNH DẠNG FILE EXCEL</p>
                <p className="text-[9px] text-slate-400 font-bold leading-normal uppercase">
                  File của bạn chỉ cần chứa 2 cột chính:<br />
                  <span className="text-slate-700 font-black">1. Mã SKU</span> (ví dụ A001)<br />
                  <span className="text-slate-700 font-black">2. Tên sản phẩm</span> (ví dụ Sản phẩm A)<br />
                  Cột không phân biệt vị trí, hệ thống sẽ tự quét tên bảng thông minh.
                </p>
              </div>
            </div>
          </div>
        </div>

        {/* Right Card - Existing SKU Database View */}
        <div className="lg:col-span-7 flex flex-col space-y-6">
          <div className="bg-white border border-slate-200 rounded-[2rem] p-6 shadow-sm flex-1 flex flex-col min-h-[450px]">
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-4 shrink-0">
              <div>
                <h3 className="text-xs font-black text-slate-700 uppercase tracking-widest flex items-center gap-2">
                  <Database className="w-4 h-4 text-emerald-600" />
                  XEM DANH SÁCH DỮ LIỆU SẴN CÓ
                </h3>
                <p className="text-[10px] text-slate-400 font-bold uppercase mt-1 leading-tight">
                  {masterSkusMeta 
                    ? `Tổng: ${masterSkusMeta.totalCount} SKU • Cập nhật lúc: ${masterSkusMeta.lastUpdated}` 
                    : "Chưa có dữ liệu nào được lưu trữ"}
                </p>
              </div>

              {/* Search Bar */}
              <div className="relative w-full sm:w-60">
                <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
                <input 
                  type="text" 
                  value={searchQuery}
                  onChange={(e) => { setSearchQuery(e.target.value); setCurrentPage(1); }}
                  placeholder="Tìm mã SKU hoặc Tên..."
                  className="w-full text-xs bg-slate-50 border border-slate-200 rounded-xl pl-9 pr-4 py-2 focus:bg-white focus:ring-1 focus:ring-indigo-500 outline-none placeholder:text-slate-400"
                />
              </div>
            </div>

            {/* Database Table or Empty Block */}
            <div className="flex-1 overflow-auto border border-slate-100 rounded-xl scrollbar-thin max-h-[300px]">
              {filteredMasterItems.length === 0 ? (
                <div className="h-full flex flex-col items-center justify-center p-12 text-center space-y-3">
                  <div className="size-14 bg-slate-50 border border-slate-100 rounded-2xl flex items-center justify-center text-slate-200">
                    <Database className="w-6 h-6" />
                  </div>
                  <div>
                    <p className="text-xs font-black text-slate-400 uppercase tracking-widest">Danh mục rỗng</p>
                    <p className="text-[9px] text-slate-300 font-bold uppercase mt-1">Xin vui lòng chọn file Excel bên trái để bắt đầu nạp SKU</p>
                  </div>
                </div>
              ) : (
                <table className="w-full text-left border-collapse table-fixed">
                  <thead className="sticky top-0 bg-white shadow-[0_1px_4px_rgba(0,0,0,0.02)] z-10">
                    <tr className="border-b border-slate-100 bg-slate-50/50">
                      <th className="p-3 text-[10px] font-black text-slate-500 uppercase tracking-wider w-12 text-center">STT</th>
                      <th className="p-3 text-[10px] font-black text-slate-500 uppercase tracking-wider w-36">Mã SKU</th>
                      <th className="p-3 text-[10px] font-black text-slate-500 uppercase tracking-wider">Tên sản phẩm tương ứng</th>
                    </tr>
                  </thead>
                  <tbody className="text-xs">
                    {paginatedItems.map((item, index) => {
                      const absoluteIdx = (currentPage - 1) * itemsPerPage + index + 1;
                      return (
                        <tr 
                          key={item.sku} 
                          className="hover:bg-slate-50/40 border-b border-slate-100 transition-colors"
                        >
                          <td className="p-3 text-center font-mono font-bold text-slate-400">{absoluteIdx}</td>
                          <td className="p-3 font-mono font-extrabold text-slate-700 truncate select-all">{item.sku}</td>
                          <td className="p-3 font-semibold text-slate-600 truncate" title={item.name}>{item.name}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>

            {/* Pagination Controls */}
            {totalPages > 1 && (
              <div className="flex justify-between items-center pt-4 border-t border-slate-100 mt-4 shrink-0">
                <span className="text-[10px] font-bold text-slate-400 uppercase">
                  Trang {currentPage} / {totalPages} ({filteredMasterItems.length} kết quả)
                </span>
                <div className="flex items-center space-x-1">
                  <button 
                    onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                    disabled={currentPage === 1}
                    className="px-2.5 py-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-200 text-slate-500 hover:text-slate-700 disabled:opacity-30 disabled:hover:bg-slate-50 rounded-lg text-[10px] font-black uppercase cursor-pointer transition-all"
                  >
                    Trước
                  </button>
                  <button 
                    onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
                    disabled={currentPage === totalPages}
                    className="px-2.5 py-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-200 text-slate-500 hover:text-slate-700 disabled:opacity-30 disabled:hover:bg-slate-50 rounded-lg text-[10px] font-black uppercase cursor-pointer transition-all"
                  >
                    Sau
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>

      </div>
    </div>
  );
}
