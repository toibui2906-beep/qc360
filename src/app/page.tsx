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
  RotateCcw
} from 'lucide-react';

export default function Home() {
  const [data, setData] = useState<UserSummary[]>([]);
  const [fileName, setFileName] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setErrorMessage(null);
    setFileName(file.name);

    try {
      const extension = file.name.split('.').pop()?.toLowerCase();
      let rows: SurveyRow[];

      if (extension === 'csv') {
        const result = Papa.parse<SurveyRow>(await file.text(), {
          header: true,
          skipEmptyLines: true,
        });
        if (result.errors.length > 0) {
          throw new Error('File CSV có lỗi định dạng và không thể đọc.');
        }
        rows = result.data;
      } else if (extension === 'xlsx') {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(await file.arrayBuffer());
        const worksheet = workbook.worksheets[0];
        if (!worksheet) {
          throw new Error('File không có trang tính để đọc.');
        }

        const headerValues = worksheet.getRow(1).values as unknown[];
        const headers = headerValues.slice(1).map((value) => String(value ?? '').trim());
        rows = [];
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return;
          const record: SurveyRow = {};
          headers.forEach((header, index) => {
            if (header) record[header] = row.getCell(index + 1).text;
          });
          rows.push(record);
        });
      } else {
        throw new Error('Chỉ hỗ trợ file .xlsx và .csv.');
      }

      const summaries = processSurveyData(rows);
      if (summaries.length === 0) {
        throw new Error('Không tìm thấy dữ liệu hợp lệ hoặc cột họ tên trong file.');
      }

      setData(summaries);
    } catch (error) {
      setData([]);
      setErrorMessage(error instanceof Error ? error.message : 'Không thể đọc file khảo sát.');
    } finally {
      setLoading(false);
    }
  };

  const handleExportExcel = async () => {
  if (data.length === 0) return;

  // 1. Khởi tạo Workbook & Worksheet
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('BaoCaoTinNhiem360');

  // 2. Định nghĩa danh sách Cột
  worksheet.columns = [
    { header: 'STT', key: 'stt', width: 8 },
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

  // 3. Thêm dữ liệu từng dòng
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

  // 4. Định dạng Dòng HEADER (In đậm, Nền xanh Y tế, Chữ trắng, Font Times New Roman 12)
  const headerRow = worksheet.getRow(1);
  headerRow.height = 30;

  headerRow.eachCell((cell) => {
    cell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: 'FFFFFF' } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '065F46' }, // Màu xanh lá Y tế thẫm (Emerald-800)
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'medium' },
      right: { style: 'thin' },
    };
  });

  // 5. Định dạng các dòng DỮ LIỆU (Times New Roman 12, Căn lề & Tô màu Mức độ TN)
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // Bỏ qua Header

    row.height = 24;

    row.eachCell((cell, colNumber) => {
      // Font cơ bản cho tất cả ô dữ liệu
      cell.font = { name: 'Times New Roman', size: 12 };
      
      // Đường viền ô
      cell.border = {
        top: { style: 'thin', color: { argb: 'E2E8F0' } },
        left: { style: 'thin', color: { argb: 'E2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'E2E8F0' } },
        right: { style: 'thin', color: { argb: 'E2E8F0' } },
      };

      // Căn lề: Cột Tên căn trái, các cột khác căn giữa
      if (colNumber === 2) {
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    });

    // --- TÔ MÀU VÀ ĐỊNH DẠNG THEO MỨC ĐỘ TÍN NHIỆM ---
    const trustPercent = data[rowNumber - 2]?.trustScorePercent || 0;
    const trustCell = row.getCell(10); // Cột 'Mức Độ TN (%)'
    const classCell = row.getCell(11); // Cột 'Xếp Loại'

    if (trustPercent >= 80) {
      // Tín nhiệm cao: Nền xanh lá nhạt, chữ in đậm
      trustCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '065F46' } };
      trustCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D1FAE5' } };
      
      classCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '065F46' } };
      classCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D1FAE5' } };
    } else if (trustPercent >= 50) {
      // Tín nhiệm: Nền xanh dương nhạt
      trustCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '1E40AF' } };
      trustCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'DBEAFE' } };

      classCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '1E40AF' } };
      classCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'DBEAFE' } };
    } else if (trustPercent >= 0) {
      // Trung bình / Thấp: Nền xám nhạt
      trustCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '334155' } };
      trustCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F1F5F9' } };

      classCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '334155' } };
      classCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F1F5F9' } };
    } else {
      // Không tín nhiệm (< 0%): Nền đỏ nhạt
      trustCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '991B1B' } };
      trustCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FEE2E2' } };

      classCell.font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: '991B1B' } };
      classCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FEE2E2' } };
    }
  });

  // 6. Ghi file Excel và Tải về
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Bao_Cao_Tin_Nhiem_CapCuu_${new Date().toISOString().slice(0, 10)}.xlsx`;
  a.click();
  window.URL.revokeObjectURL(url);
};

  // Tính toán nhanh số liệu thống kê chung cho KPI Cards
  const totalStaff = data.length;
  const topExcellence = data.filter(d => d.trustScorePercent >= 80).length;
  const avgDepartmentTrust = totalStaff > 0 
    ? Math.round((data.reduce((acc, curr) => acc + curr.trustScorePercent, 0) / totalStaff) * 10) / 10 
    : 0;

  return (
    <main className="min-h-screen bg-slate-50/80 text-slate-900 pb-16">
      
      {/* SECTION 1: HERO HEADER (Branding & Marketing) */}
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
              Giải pháp chuẩn hóa, phân tích đa chiều chỉ số tín nhiệm và hiệu suất làm việc của đội ngũ y bác sĩ, điều dưỡng Cấp cứu. Tự động hóa báo cáo minh bạch & chính xác.
            </p>
          </div>

          <div className="flex items-center space-x-3 bg-white/10 backdrop-blur-md border border-white/10 p-4 rounded-2xl shrink-0">
            <Building2 className="w-10 h-10 text-emerald-400" />
            <div>
              <p className="text-xs text-slate-300 font-medium">Đơn vị ứng dụng</p>
              <p className="text-sm font-bold text-white">Khoa Cấp Cứu - Bệnh Viện ĐKQT Vinmec TimesCity</p>
            </div>
          </div>
        </div>
      </header>

      {/* Container chính đè lên Header */}
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
                onClick={() => { setData([]); setFileName(''); }}
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
              Định dạng hỗ trợ: .XLSX, .CSV (Hệ thống tự động lọc & xếp hạng)
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

          {/* Alert Lỗi */}
          {errorMessage && (
            <div className="mt-4 bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded-xl flex items-center space-x-3 text-sm">
              <AlertCircle className="w-5 h-5 text-red-500 shrink-0" />
              <span>{errorMessage}</span>
            </div>
          )}
        </section>

        {/* SECTION 3: DASHBOARD STATS (KPI Cards - Chỉ hiện khi có dữ liệu) */}
        {data.length > 0 && !loading && (
          <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
            
            {/* Card 1 */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200/80 flex items-center space-x-4">
              <div className="p-3.5 bg-blue-50 text-blue-600 rounded-2xl">
                <Users className="w-6 h-6" />
              </div>
              <div>
                <p className="text-xs font-medium text-slate-500">Tổng Cán Bộ Đánh Giá</p>
                <p className="text-2xl font-bold text-slate-800">{totalStaff} <span className="text-xs font-normal text-slate-400">nhân sự</span></p>
              </div>
            </div>

            {/* Card 2 */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200/80 flex items-center space-x-4">
              <div className="p-3.5 bg-emerald-50 text-emerald-600 rounded-2xl">
                <Award className="w-6 h-6" />
              </div>
              <div>
                <p className="text-xs font-medium text-slate-500">Tín Nhiệm Cao (&ge;80%)</p>
                <p className="text-2xl font-bold text-emerald-600">{topExcellence} <span className="text-xs font-normal text-slate-400">nhân sự</span></p>
              </div>
            </div>

            {/* Card 3 */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200/80 flex items-center space-x-4">
              <div className="p-3.5 bg-teal-50 text-teal-600 rounded-2xl">
                <TrendingUp className="w-6 h-6" />
              </div>
              <div>
                <p className="text-xs font-medium text-slate-500">Chỉ Số Tín Nhiệm Trung Bình</p>
                <p className="text-2xl font-bold text-teal-700">{avgDepartmentTrust}%</p>
              </div>
            </div>

            {/* Card 4 */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200/80 flex items-center space-x-4">
              <div className="p-3.5 bg-indigo-50 text-indigo-600 rounded-2xl">
                <ShieldCheck className="w-6 h-6" />
              </div>
              <div>
                <p className="text-xs font-medium text-slate-500">Trạng Thái Xếp Hạng</p>
                <p className="text-sm font-bold text-indigo-900 mt-1">Đã Chuẩn Hóa 100%</p>
              </div>
            </div>

          </section>
        )}

        {/* SECTION 4: TABLE REPORT SECTION */}
        {data.length > 0 && !loading && (
          <section className="bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200/80 p-6 space-y-5">
            
            {/* Header Bảng */}
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 border-b border-slate-100 pb-4">
              <div className="flex items-center space-x-3">
                <div className="p-2 bg-emerald-50 text-emerald-600 rounded-lg">
                  <BarChart3 className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-lg font-bold text-slate-800">
                    Bảng Xếp Hạng Tín Nhiệm Cán Bộ
                  </h3>
                  <p className="text-xs text-slate-500">Sắp xếp tự động theo thứ tự Mức độ tín nhiệm giảm dần</p>
                </div>
              </div>

              <button
                onClick={handleExportExcel}
                className="inline-flex items-center space-x-2 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white px-4 py-2.5 rounded-xl text-sm font-semibold shadow-sm shadow-emerald-200 transition"
              >
                <Download className="w-4 h-4" />
                <span>Xuất Báo Cáo Excel</span>
              </button>
            </div>

            {/* Bảng Dữ Liệu */}
            <div className="overflow-x-auto rounded-xl border border-slate-200">
              <table className="w-full text-sm text-left text-slate-600">
                <thead className="bg-slate-50 text-slate-700 font-bold text-xs uppercase tracking-wider border-b border-slate-200">
                  <tr>
                    <th className="p-3.5 text-center w-16">Hạng</th>
                    <th className="p-3.5">Họ và tên cán bộ</th>
                    <th className="p-3.5 text-center">Tổng phiếu</th>
                    <th className="p-3.5 text-center">ĐTB Tổng</th>
                    <th className="p-3.5 text-center text-emerald-700">TN Cao</th>
                    <th className="p-3.5 text-center text-blue-700">Tín Nhiệm</th>
                    <th className="p-3.5 text-center text-slate-600">Trung Bình</th>
                    <th className="p-3.5 text-center text-amber-700">TN Thấp</th>
                    <th className="p-3.5 text-center text-red-700">Không TN</th>
                    <th className="p-3.5 text-center font-extrabold text-emerald-900 bg-emerald-50/50">Mức Độ TN (%)</th>
                    <th className="p-3.5 text-center">Xếp Loại</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {data.map((row) => (
                    <tr key={row.stt} className="hover:bg-slate-50/80 transition-colors">
                      
                      {/* Cột Hạng / STT */}
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
                      <td className="p-3.5 text-center font-medium text-slate-700">{row.avgTotalScore}</td>
                      <td className="p-3.5 text-center font-semibold text-emerald-600">{row.highTrustCount}</td>
                      <td className="p-3.5 text-center font-semibold text-blue-600">{row.trustCount}</td>
                      <td className="p-3.5 text-center text-slate-500">{row.mediumTrustCount}</td>
                      <td className="p-3.5 text-center font-semibold text-amber-600">{row.lowTrustCount}</td>
                      <td className="p-3.5 text-center font-semibold text-red-600">{row.noTrustCount}</td>
                      
                      {/* Cột % Tín nhiệm nổi bật */}
                      <td className="p-3.5 text-center font-extrabold text-emerald-700 bg-emerald-50/30 text-base">
                        {row.trustScorePercent}%
                      </td>

                      {/* Cột Xếp loại */}
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