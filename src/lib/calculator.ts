export interface SurveyRow {
  ID?: number;
  [key: string]: unknown;
}

export interface UserSummary {
  stt: number;
  fullName: string;
  totalVotes: number;
  avgTotalScore: number;
  highTrustCount: number;
  trustCount: number;
  mediumTrustCount: number;
  lowTrustCount: number;
  noTrustCount: number;
  trustScorePercent: number;
  classification: string;
}

export const TRUST_LEVELS = {
  HIGH: 'Tín nhiệm cao',
  TRUSTED: 'Tín nhiệm',
  MEDIUM: 'Trung bình',
  LOW: 'Tín nhiệm thấp',
  NO_TRUST: 'Không tín nhiệm',
} as const;

/**
 * Xử lý và phân tích dữ liệu khảo sát 360 độ
 * @param rawData Dữ liệu thô đọc từ Excel
 * @param sortBy 'trust' (Sắp xếp theo % Tín nhiệm) | 'avg' (Sắp xếp theo Điểm TB KPI)
 */
export function processSurveyData(
  rawData: SurveyRow[],
  sortBy: 'trust' | 'avg' = 'trust'
): UserSummary[] {
  if (!rawData || rawData.length === 0) return [];

  // 1. Tự động nhận diện tên cột Họ tên và Mức độ tín nhiệm
  const firstRow = rawData[0];
  const keys = Object.keys(firstRow);

  const nameKey =
    keys.find((k) => k.trim().toLowerCase().includes('họ tên')) ||
    'Họ tên người được khảo sát';
  const trustKey = keys.find((k) =>
    k.trim().toLowerCase().includes('mức độ tín nhiệm')
  );

  // 2. Nhóm dữ liệu theo Tên cán bộ y tế
  const grouped = new Map<string, SurveyRow[]>();

  rawData.forEach((row) => {
    const name = row[nameKey]?.toString().trim();
    if (!name) return;

    if (!grouped.has(name)) {
      grouped.set(name, []);
    }
    grouped.get(name)!.push(row);
  });

  const summaryList: UserSummary[] = [];

  // 3. Tính toán tổng hợp cho từng cá nhân
  grouped.forEach((votes, fullName) => {
    const totalVotes = votes.length;
    let sumTotalScores = 0;
    let highTrustCount = 0;
    let trustCount = 0;
    let mediumTrustCount = 0;
    let lowTrustCount = 0;
    let noTrustCount = 0;

    votes.forEach((vote) => {
      // Tính tổng điểm các tiêu chí TC
      Object.keys(vote).forEach((key) => {
        if (key.toUpperCase().startsWith('TC')) {
          const score = parseFloat(String(vote[key]));
          if (!isNaN(score)) {
            sumTotalScores += score;
          }
        }
      });

      // Lấy mức độ tín nhiệm
      const rawLevel = trustKey ? vote[trustKey]?.toString().trim() : '';

      if (rawLevel === TRUST_LEVELS.HIGH) highTrustCount++;
      else if (rawLevel === TRUST_LEVELS.TRUSTED) trustCount++;
      else if (rawLevel === TRUST_LEVELS.MEDIUM) mediumTrustCount++;
      else if (rawLevel === TRUST_LEVELS.LOW) lowTrustCount++;
      else if (rawLevel === TRUST_LEVELS.NO_TRUST) noTrustCount++;
    });

    const avgTotalScore = totalVotes > 0 ? sumTotalScores / totalVotes : 0;

    // Công thức tính Mức độ tín nhiệm (%)
    const scoreNumerator =
      highTrustCount + trustCount - (lowTrustCount + 2 * noTrustCount);
    const trustScorePercent =
      totalVotes > 0 ? (scoreNumerator / totalVotes) * 100 : 0;

    // Phân loại xếp loại năng lực
    let classification = 'Trung bình';
    if (trustScorePercent >= 80) classification = 'Tín nhiệm cao';
    else if (trustScorePercent >= 50) classification = 'Tín nhiệm';
    else if (trustScorePercent < 0) classification = 'Không tín nhiệm';

    summaryList.push({
      stt: 0, // Sẽ gán lại sau khi xếp hạng
      fullName,
      totalVotes,
      avgTotalScore: Math.round(avgTotalScore * 100) / 100,
      highTrustCount,
      trustCount,
      mediumTrustCount,
      lowTrustCount,
      noTrustCount,
      trustScorePercent: Math.round(trustScorePercent * 100) / 100,
      classification,
    });
  });

  // 4. Sắp xếp danh sách theo Option được chọn
  summaryList.sort((a, b) => {
    if (sortBy === 'trust') {
      // Option 1: Ưu tiên % Tín nhiệm -> Trùng thì xét Điểm TB
      if (b.trustScorePercent !== a.trustScorePercent) {
        return b.trustScorePercent - a.trustScorePercent;
      }
      return b.avgTotalScore - a.avgTotalScore;
    } else {
      // Option 2: Ưu tiên Điểm TB (KPI) -> Trùng thì xét % Tín nhiệm
      if (b.avgTotalScore !== a.avgTotalScore) {
        return b.avgTotalScore - a.avgTotalScore;
      }
      return b.trustScorePercent - a.trustScorePercent;
    }
  });

  // 5. Đánh lại STT/Hạng chuẩn từ 1 đến hết theo danh sách đã xếp hạng
  return summaryList.map((item, index) => ({
    ...item,
    stt: index + 1,
  }));
}