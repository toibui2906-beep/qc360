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

export function processSurveyData(rawData: SurveyRow[]): UserSummary[] {
  if (!rawData || rawData.length === 0) return [];

  // 1. Tự động nhận diện tên cột Họ tên và Mức độ tín nhiệm
  const firstRow = rawData[0];
  const keys = Object.keys(firstRow);

  const nameKey = keys.find(k => k.trim().toLowerCase().includes('họ tên')) || 'Họ tên người được khảo sát';
  const trustKey = keys.find(k => k.trim().toLowerCase().includes('mức độ tín nhiệm'));

  // 2. Nhóm dữ liệu theo Tên nhân sự
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

  // 3. Tính toán số liệu cho từng cá nhân
  grouped.forEach((votes, fullName) => {
    const totalVotes = votes.length;
    let sumTotalScores = 0;
    let highTrustCount = 0;
    let trustCount = 0;
    let mediumTrustCount = 0;
    let lowTrustCount = 0;
    let noTrustCount = 0;

    votes.forEach((vote) => {
      // Tính tổng điểm tiêu chí TC
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
    
    // Tính phần trăm tín nhiệm
    const scoreNumerator = (highTrustCount + trustCount) - (lowTrustCount + 2 * noTrustCount);
    const trustScorePercent = totalVotes > 0 ? (scoreNumerator / totalVotes) * 100 : 0;

    // Phân loại xếp loại
    let classification = 'Trung bình';
    if (trustScorePercent >= 80) classification = 'Tín nhiệm cao';
    else if (trustScorePercent >= 50) classification = 'Tín nhiệm';
    else if (trustScorePercent < 0) classification = 'Không tín nhiệm';

    summaryList.push({
      stt: 0, // Sẽ gán lại sau khi sắp xếp
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

  // 4. Sắp xếp danh sách giảm dần theo Mức độ tín nhiệm (%), trùng % thì xét theo ĐTB tổng
  summaryList.sort((a, b) => {
    if (b.trustScorePercent !== a.trustScorePercent) {
      return b.trustScorePercent - a.trustScorePercent;
    }
    return b.avgTotalScore - a.avgTotalScore;
  });

  // 5. Đánh lại STT theo thứ tự hạng từ 1 đến hết
  return summaryList.map((item, index) => ({
    ...item,
    stt: index + 1,
  }));
}