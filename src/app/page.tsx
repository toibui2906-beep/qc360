'use client';

import React, { useState, useRef } from 'react';
import ExcelJS from 'exceljs';
import Papa from 'papaparse';
import { processSurveyData, type SurveyRow, type UserSummary } from '@/lib/calculator';
import { 
  UploadCloud, 
  FileSpreadsheet, 
  Download, 
  RefreshCw, 
  AlertCircle, 
  Activity, 
  Users, 
  Award, 
  BarChart3, 
  Sparkles,
  ShieldCheck,
  Building2,
  TrendingUp,
  RotateCcw,
  SlidersHorizontal
} from 'lucide-react';

export default function Home() {
  const [data, setData] = useState<UserSummary[]>([]);
  const [rawSurveyData, setRawSurveyData] = useState<SurveyRow[]>([]); // Lưu dữ liệu thô để re-sort nhanh
  const [sortBy, setSortBy] = useState<'trust' | 'avg'>('trust'); // 'trust': % Tín nhiệm, 'avg': Điểm TB KPI
  const [fileName, setFileName] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Xử lý khi thay đổi Option sắp xếp
  const handleSortChange = (newSortBy: 'trust' | 'avg') => {
    setSortBy(newSortBy);
    if (rawSurveyData.length > 0) {
      const reSorted = processSurveyData(rawSurveyData, newSortBy);
      setData(reSorted);
    }
  };

  // Xử lý upload file Excel
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    setLoading(true);
    setErrorMessage(null);

    try {
      const extension = file.name.split('.').pop()?.toLowerCase();
      let rawData: SurveyRow[];

      if (extension === 'csv') {
        const result = Papa.parse<SurveyRow>(await file.text(), {
          header: true,
          skipEmptyLines: true,
        });
        if (result.errors.length > 0) {
          throw new Error('File CSV có lỗi định dạng và không thể đọc.');
        }
        rawData = result.data;
      } else if (extension === 'xlsx') {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(await file.arrayBuffer());
        const worksheet = workbook.worksheets[0];
        if (!worksheet) {
          throw new Error('File Excel trống hoặc không có dữ liệu!');
        }

        const headerRow = worksheet.getRow(1);
        const headers = Array.from({ length: headerRow.cellCount }, (_, index) =>
          String(headerRow.getCell(index + 1).value ?? '').trim()
        );
        rawData = [];
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return;
          const record: SurveyRow = {};
          headers.forEach((header, index) => {
            if (header) record[header] = row.getCell(index + 1).value;
          });
          rawData.push(record);
        });
      } else {
        throw new Error('Chỉ hỗ trợ file .xlsx và .csv.');
      }

      if (rawData.length === 0) {
        throw new Error('Không tìm thấy dữ liệu trong file!');
      }

      setRawSurveyData(rawData);
      setData(processSurveyData(rawData, sortBy));
    } catch (error) {
      console.error('File Processing Error:', error);
      setErrorMessage(error instanceof Error ? error.message : 'File không hợp lệ hoặc lỗi định dạng!');
      setData([]);
    } finally {
      setLoading(false);
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
    }
  };

  // Xuất Báo Cáo Excel đẹp với ExcelJS (Định dạng Times New Roman 12, Màu sắc chuẩn)
  const handleExportExcel = async () => {
    if (data.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('BaoCaoTinNhiem360');

    // Khai báo cột
    worksheet.columns = [
      { header: 'Hạng (STT)', key: 'stt', width: 12 },
      { header: 'Họ tên cán bộ y tế', key: 'fullName', width: 28 },
      { header: 'Số phiếu', key: 'totalVotes', width: 12 },
      { header: 'ĐTB Tổng', key: 'avgTotalScore', width: 14 },
      { header: 'TN Cao', key: 'highTrustCount', width: 12 },
      { header: 'Tín Nhiệm', key: 'trustCount', width: 12 },
      { header: 'Trung Bình', key: 'mediumTrustCount', width: 12 },
      { header: 'TN Thấp', key: 'lowTrustCount', width: 12 },
      { header: 'Không TN', key: 'noTrustCount', width: 12 },
      { header: 'Mức Độ TN (%)', key: 'trustScorePercent', width: 18 },
      { header: 'Xếp Loại', key: 'classification', width: 20 },
    ];

    // Thêm dữ liệu
    data.forEach((item) => {
      worksheet.addRow({
        stt: item.stt,
        fullName: item.fullName,
        totalVotes: item.totalVotes,
        avgTotalScore: item.avgTotalScore,
        highTrustCount: item.highTrustCount,
        trustCount: item.trustCount,
        mediumTrustCount: item.mediumTrustCount,
        lowTrustCount: item.lowTrustCount,
        noTrustCount: item.noTrustCount,
        trustScorePercent: `${item.trustScorePercent}%`,
        classification: item.classification,
      });
    });

    // Định dạng Header
    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: 'FFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '065F46' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'medium' },
        right: { style: 'thin' },
      };
    });

    // Định dạng Dữ liệu & Màu sắc
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;

      row.height = 24;
      row.eachCell((cell, colNumber) => {
        cell.font = { name: 'Times New Roman', size: 12 };
        cell.border = {
          top: { style: 'thin', color: { argb: 'E2E8F0' } },
          left: { style: 'thin', color: { argb: 'E2E8F0' } },
          bottom: { style: 'thin', color: { argb: 'E2E8F0' } },
          right: { style: 'thin', color: { argb: 'E2E8F0' } },
        };
        cell.alignment = { vertical: 'middle', horizontal: colNumber === 2 ? 'left' : 'center' };
      });

      // Tô màu theo Mức độ Tín nhiệm
      const trustPercent = data[rowNumber - 2]?.trustScorePercent || 0;
      const trustCell = row.getCell(10);
      const classCell = row.getCell(11);

      if (trustPercent >= 80) {
        trustCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '065F46' } };
        trustCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D1FAE5' } };
        classCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '065F46' } };
        classCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D1FAE5' } };
      } else if (trustPercent >= 50) {
        trustCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '1E40AF' } };
        trustCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'DBEAFE' } };
        classCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '1E40AF' } };
        classCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'DBEAFE' } };
      } else if (trustPercent >= 0) {
        trustCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '334155' } };
        trustCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F1F5F9' } };
        classCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '334155' } };
        classCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F1F5F9' } };
      } else {
        trustCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '991B1B' } };
        trustCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FEE2E2' } };
        classCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '991B1B' } };
        classCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FEE2E2' } };
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Bao_Cao_Tin_Nhiem_CapCuu_${sortBy === 'trust' ? 'KhenThuong' : 'KPI'}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
  };

  const totalStaff = data.length;
  const topExcellence = data.filter(d => d.trustScorePercent >= 80).length;
  const avgDepartmentTrust = totalStaff > 0 
    ? Math.round((data.reduce((acc, curr) => acc + curr.trustScorePercent, 0) / totalStaff) * 10) / 10 
    : 0;

  return (
    <main className="min-h-screen bg-slate-50/80 text-slate-900 pb-16">
      
      {/* SECTION 1: HERO HEADER */}
      <header className="bg-gradient-to-r from-emerald-800 via-teal-800 to-slate-900 text-white pt-10 pb-20 px-6 shadow-md relative overflow-hidden">
        <div className="absolute -right-10 -bottom-10 opacity-10 pointer-events-none">
          <Activity className="w-96 h-96 text-white" />
        </div>
        
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row justify-between items-start md:items-center gap-6 relative z-10">
          <div className="space-y-3 max-w-3xl">
            <div className="inline-flex items-center space-x-2 bg-emerald-500/20 border border-emerald-400/30 backdrop-blur-md px-3.5 py-1.5 rounded-full text-emerald-200 text-xs font-semibold tracking-wide uppercase">
              <Sparkles className="w-3.5 h-3.5 text-emerald-300" />
              <span>Hệ Thống Đánh Giá Năng Lực Chuẩn Y Khoa</span>
            </div>
            
            <h1 className="text-3xl md:text-4xl font-extrabold tracking-tight leading-tight">
              Đánh Giá 360° Cán Bộ Y Tế <span className="text-emerald-400">Khoa Cấp Cứu</span>
            </h1>
            
            <p className="text-slate-300 text-sm md:text-base leading-relaxed">
              Giải pháp chuẩn hóa, phân tích đa chiều chỉ số tín nhiệm và hiệu suất làm việc của đội ngũ y bác sĩ, điều dưỡng Cấp cứu.
            </p>
          </div>

          <div className="flex items-center space-x-3 bg-white/10 backdrop-blur-md border border-white/10 p-4 rounded-2xl shrink-0">
            <Building2 className="w-10 h-10 text-emerald-400" />
            <div>
              <p className="text-xs text-slate-300 font-medium">Khoa phòng</p>
              <p className="text-sm font-bold text-white">Khoa Cấp Cứu - Bệnh Viện ĐKQT Vinmec TimesCity</p>
            </div>
          </div>
        </div>
      </header>

      {/* Main Container */}
      <div className="max-w-7xl mx-auto px-6 -mt-10 space-y-8 relative z-20">

        {/* SECTION 2: IMPORT DATA SECTION */}
        <section className="bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200/80 p-6 md:p-8">
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center space-x-3">
              <div className="p-2.5 bg-emerald-50 text-emerald-600 rounded-xl">
                <UploadCloud className="w-6 h-6" />
              </div>
              <div>
                <h2 className="text-lg font-bold text-slate-800">Nhập Dữ Liệu Khảo Sát</h2>
                <p className="text-xs text-slate-500">Tải lên file Excel kết quả đánh giá 360 độ từ biểu mẫu khảo sát</p>
              </div>
            </div>

            {data.length > 0 && (
              <button
                onClick={() => { setData([]); setRawSurveyData([]); setFileName(''); }}
                className="inline-flex items-center space-x-1.5 text-xs text-slate-500 hover:text-red-600 transition"
              >
                <RotateCcw className="w-3.5 h-3.5" />
                <span>Làm mới</span>
              </button>
            )}
          </div>

          <label 
            className={`group relative flex flex-col items-center justify-center border-2 border-dashed rounded-2xl p-8 text-center transition-all duration-200 ${
              loading 
                ? 'bg-slate-50 border-slate-300 cursor-not-allowed opacity-75' 
                : 'border-slate-300 hover:border-emerald-500 hover:bg-emerald-50/30 cursor-pointer'
            }`}
          >
            {loading ? (
              <RefreshCw className="w-12 h-12 text-emerald-600 animate-spin mb-3" />
            ) : (
              <div className="p-4 bg-emerald-50 group-hover:bg-emerald-100/80 text-emerald-600 rounded-full mb-3 transition">
                <FileSpreadsheet className="w-8 h-8" />
              </div>
            )}

            <span className="text-base font-semibold text-slate-700 group-hover:text-emerald-700 transition">
              {loading 
                ? 'Đang phân tích & xử lý dữ liệu y tế...' 
                : fileName 
                  ? `File hiện tại: ${fileName}` 
                  : 'Kéo thả file Excel kết quả vào đây hoặc nhấp để chọn'}
            </span>
            
            <p className="text-xs text-slate-400 mt-1">
              Định dạng hỗ trợ: .XLSX, .CSV
            </p>

            <input 
              ref={fileInputRef}
              type="file" 
              accept=".xlsx, .csv" 
              className="hidden" 
              disabled={loading}
              onChange={handleFileUpload} 
            />
          </label>

          {errorMessage && (
            <div className="mt-4 bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded-xl flex items-center space-x-3 text-sm">
              <AlertCircle className="w-5 h-5 text-red-500 shrink-0" />
              <span>{errorMessage}</span>
            </div>
          )}
        </section>

        {/* SECTION 3: DASHBOARD STATS */}
        {data.length > 0 && !loading && (
          <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200/80 flex items-center space-x-4">
              <div className="p-3.5 bg-blue-50 text-blue-600 rounded-2xl">
                <Users className="w-6 h-6" />
              </div>
              <div>
                <p className="text-xs font-medium text-slate-500">Tổng Cán Bộ Đánh Giá</p>
                <p className="text-2xl font-bold text-slate-800">{totalStaff} <span className="text-xs font-normal text-slate-400">nhân sự</span></p>
              </div>
            </div>

            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200/80 flex items-center space-x-4">
              <div className="p-3.5 bg-emerald-50 text-emerald-600 rounded-2xl">
                <Award className="w-6 h-6" />
              </div>
              <div>
                <p className="text-xs font-medium text-slate-500">Tín Nhiệm Cao (&ge;80%)</p>
                <p className="text-2xl font-bold text-emerald-600">{topExcellence} <span className="text-xs font-normal text-slate-400">nhân sự</span></p>
              </div>
            </div>

            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200/80 flex items-center space-x-4">
              <div className="p-3.5 bg-teal-50 text-teal-600 rounded-2xl">
                <TrendingUp className="w-6 h-6" />
              </div>
              <div>
                <p className="text-xs font-medium text-slate-500">Chỉ Số Tín Nhiệm TB Khoa</p>
                <p className="text-2xl font-bold text-teal-700">{avgDepartmentTrust}%</p>
              </div>
            </div>

            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200/80 flex items-center space-x-4">
              <div className="p-3.5 bg-indigo-50 text-indigo-600 rounded-2xl">
                <ShieldCheck className="w-6 h-6" />
              </div>
              <div>
                <p className="text-xs font-medium text-slate-500">Trạng Thái Chế Độ</p>
                <p className="text-sm font-bold text-indigo-900 mt-1">
                  {sortBy === 'trust' ? '🏆 Khen Thưởng Quý' : '📊 Đánh Giá KPI'}
                </p>
              </div>
            </div>
          </section>
        )}

        {/* SECTION 4: TABLE REPORT SECTION WITH 2 SORT OPTIONS */}
        {data.length > 0 && !loading && (
          <section className="bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200/80 p-6 space-y-6">
            
            {/* Header Bảng & Nút Xuất Excel */}
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 border-b border-slate-100 pb-4">
              <div className="flex items-center space-x-3">
                <div className="p-2 bg-emerald-50 text-emerald-600 rounded-lg">
                  <BarChart3 className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-lg font-bold text-slate-800">
                    Bảng Xếp Hạng Tín Nhiệm Cán Bộ
                  </h3>
                  <p className="text-xs text-slate-500">
                    {sortBy === 'trust' 
                      ? 'Đang ưu tiên sắp xếp từ cao xuống thấp theo % Mức Độ Tín Nhiệm' 
                      : 'Đang ưu tiên sắp xếp từ cao xuống thấp theo Điểm Trung Bình Tổng (KPI)'}
                  </p>
                </div>
              </div>

              <button
                onClick={handleExportExcel}
                className="inline-flex items-center space-x-2 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white px-4 py-2.5 rounded-xl text-sm font-semibold shadow-sm shadow-emerald-200 transition shrink-0"
              >
                <Download className="w-4 h-4" />
                <span>Xuất Báo Cáo Excel</span>
              </button>
            </div>

            {/* BỘ CHỌN 2 OPTION SẮP XẾP */}
            <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 bg-slate-100/80 p-3.5 rounded-xl border border-slate-200">
              <div className="flex items-center space-x-2 text-slate-700 font-bold text-xs uppercase tracking-wider">
                <SlidersHorizontal className="w-4 h-4 text-emerald-600" />
                <span>Chọn Tiêu Chí Xếp Hạng:</span>
              </div>
              
              <div className="flex items-center space-x-2 bg-white p-1 rounded-xl border border-slate-200 shadow-sm w-full sm:w-auto">
                <button
                  onClick={() => handleSortChange('trust')}
                  className={`flex-1 sm:flex-none px-4 py-2 rounded-lg text-xs font-semibold transition-all ${
                    sortBy === 'trust'
                      ? 'bg-emerald-600 text-white shadow-sm'
                      : 'text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                  }`}
                >
                  🏆 Option 1: % Tín Nhiệm (Khen Thưởng Quý)
                </button>

                <button
                  onClick={() => handleSortChange('avg')}
                  className={`flex-1 sm:flex-none px-4 py-2 rounded-lg text-xs font-semibold transition-all ${
                    sortBy === 'avg'
                      ? 'bg-emerald-600 text-white shadow-sm'
                      : 'text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                  }`}
                >
                  📊 Option 2: Điểm TB (Đánh Giá KPI)
                </button>
              </div>
            </div>

            {/* Bảng Dữ Liệu */}
            <div className="overflow-x-auto rounded-xl border border-slate-200">
              <table className="w-full text-sm text-left text-slate-600">
                <thead className="bg-slate-50 text-slate-700 font-bold text-xs uppercase tracking-wider border-b border-slate-200">
                  <tr>
                    <th className="p-3.5 text-center w-16">Hạng</th>
                    <th className="p-3.5">Họ và tên cán bộ</th>
                    <th className="p-3.5 text-center">Tổng phiếu</th>
                    <th className={`p-3.5 text-center ${sortBy === 'avg' ? 'bg-emerald-100/70 text-emerald-900 font-extrabold' : ''}`}>
                      ĐTB Tổng {sortBy === 'avg' && '⭐'}
                    </th>
                    <th className="p-3.5 text-center text-emerald-700">TN Cao</th>
                    <th className="p-3.5 text-center text-blue-700">Tín Nhiệm</th>
                    <th className="p-3.5 text-center text-slate-600">Trung Bình</th>
                    <th className="p-3.5 text-center text-amber-700">TN Thấp</th>
                    <th className="p-3.5 text-center text-red-700">Không TN</th>
                    <th className={`p-3.5 text-center ${sortBy === 'trust' ? 'bg-emerald-100/70 text-emerald-900 font-extrabold' : ''}`}>
                      Mức Độ TN (%) {sortBy === 'trust' && '⭐'}
                    </th>
                    <th className="p-3.5 text-center">Xếp Loại</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {data.map((row) => (
                    <tr key={row.stt} className="hover:bg-slate-50/80 transition-colors">
                      <td className="p-3.5 text-center font-bold">
                        {row.stt === 1 ? (
                          <span className="inline-flex items-center justify-center w-7 h-7 bg-amber-100 text-amber-700 rounded-full text-xs">🥇 1</span>
                        ) : row.stt === 2 ? (
                          <span className="inline-flex items-center justify-center w-7 h-7 bg-slate-200 text-slate-700 rounded-full text-xs">🥈 2</span>
                        ) : row.stt === 3 ? (
                          <span className="inline-flex items-center justify-center w-7 h-7 bg-amber-800/10 text-amber-800 rounded-full text-xs">🥉 3</span>
                        ) : (
                          <span className="text-slate-500">{row.stt}</span>
                        )}
                      </td>

                      <td className="p-3.5 font-semibold text-slate-800">{row.fullName}</td>
                      <td className="p-3.5 text-center font-medium">{row.totalVotes}</td>
                      
                      {/* Highlight ĐTB khi chọn Option 2 */}
                      <td className={`p-3.5 text-center font-medium ${sortBy === 'avg' ? 'bg-emerald-50/60 font-bold text-emerald-800' : 'text-slate-700'}`}>
                        {row.avgTotalScore}
                      </td>

                      <td className="p-3.5 text-center font-semibold text-emerald-600">{row.highTrustCount}</td>
                      <td className="p-3.5 text-center font-semibold text-blue-600">{row.trustCount}</td>
                      <td className="p-3.5 text-center text-slate-500">{row.mediumTrustCount}</td>
                      <td className="p-3.5 text-center font-semibold text-amber-600">{row.lowTrustCount}</td>
                      <td className="p-3.5 text-center font-semibold text-red-600">{row.noTrustCount}</td>
                      
                      {/* Highlight % Tín nhiệm khi chọn Option 1 */}
                      <td className={`p-3.5 text-center font-extrabold text-base ${sortBy === 'trust' ? 'bg-emerald-50/60 text-emerald-700' : 'text-slate-700'}`}>
                        {row.trustScorePercent}%
                      </td>

                      <td className="p-3.5 text-center">
                        <span className={`inline-flex items-center px-2.5 py-1 rounded-full text-xs font-semibold ${
                          row.trustScorePercent >= 80 ? 'bg-emerald-100 text-emerald-800 border border-emerald-200' :
                          row.trustScorePercent >= 50 ? 'bg-blue-100 text-blue-800 border border-blue-200' :
                          row.trustScorePercent >= 0 ? 'bg-slate-100 text-slate-700 border border-slate-200' : 'bg-red-100 text-red-800 border border-red-200'
                        }`}>
                          {row.classification}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        )}

      </div>
    </main>
  );
}