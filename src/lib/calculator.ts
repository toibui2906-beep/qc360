import criteriaConfig from '@/config/criteria.json'; // 1. Import file config hệ số

export interface SurveyRow {
  ID?: number;
  [key: string]: any;
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

export function processSurveyData(
  rawData: SurveyRow[],
  sortBy: 'trust' | 'avg' = 'trust'
): UserSummary[] {
  if (!rawData || rawData.length === 0) return [];

  const firstRow = rawData[0];
  const keys = Object.keys(firstRow);

  const nameKey =
    keys.find((k) => k.trim().toLowerCase().includes('họ tên')) ||
    'Họ tên người được khảo sát';
  const trustKey = keys.find((k) =>
    k.trim().toLowerCase().includes('mức độ tín nhiệm')
  );

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

  // Ép kiểu config hệ số
  const weights = criteriaConfig as Record<string, number>;

  grouped.forEach((votes, fullName) => {
    const totalVotes = votes.length;
    let sumWeightedScores = 0; // Tổng điểm đã nhân hệ số
    let highTrustCount = 0;
    let trustCount = 0;
    let mediumTrustCount = 0;
    let lowTrustCount = 0;
    let noTrustCount = 0;

    votes.forEach((vote) => {
      let voteWeightedSum = 0;
      let voteTotalWeight = 0;

      // Duyệt qua các tiêu chí (TC1, TC2,...)
      Object.keys(vote).forEach((key) => {
        const trimmedKey = key.trim().toUpperCase();
        if (trimmedKey.startsWith('TC')) {
          const score = parseFloat(vote[key]);
          if (!isNaN(score)) {
            // Lấy hệ số từ file config (mặc định là 1 nếu không có trong config)
            const weight = weights[trimmedKey] ?? 1;
            
            voteWeightedSum += score * weight;
            voteTotalWeight += weight;
          }
        }
      });

      // Điểm trung bình có trọng số cho 1 phiếu khảo sát
      if (voteTotalWeight > 0) {
        sumWeightedScores += voteWeightedSum / voteTotalWeight;
      }

      // Đếm số lượng phiếu tín nhiệm
      const rawLevel = trustKey ? vote[trustKey]?.toString().trim() : '';
      if (rawLevel === TRUST_LEVELS.HIGH) highTrustCount++;
      else if (rawLevel === TRUST_LEVELS.TRUSTED) trustCount++;
      else if (rawLevel === TRUST_LEVELS.MEDIUM) mediumTrustCount++;
      else if (rawLevel === TRUST_LEVELS.LOW) lowTrustCount++;
      else if (rawLevel === TRUST_LEVELS.NO_TRUST) noTrustCount++;
    });

    // Điểm Trung Bình Tổng của tất cả các phiếu
    const avgTotalScore = totalVotes > 0 ? sumWeightedScores / totalVotes : 0;

    // Tính % Tín nhiệm
    const scoreNumerator =
      highTrustCount + trustCount - (lowTrustCount + 2 * noTrustCount);
    const trustScorePercent =
      totalVotes > 0 ? (scoreNumerator / totalVotes) * 100 : 0;

    // Phân loại
    let classification = 'Trung bình';
    if (trustScorePercent >= 80) classification = 'Tín nhiệm cao';
    else if (trustScorePercent >= 50) classification = 'Tín nhiệm';
    else if (trustScorePercent < 0) classification = 'Không tín nhiệm';

    summaryList.push({
      stt: 0,
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

  // Sắp xếp theo Option
  summaryList.sort((a, b) => {
    if (sortBy === 'trust') {
      if (b.trustScorePercent !== a.trustScorePercent) {
        return b.trustScorePercent - a.trustScorePercent;
      }
      return b.avgTotalScore - a.avgTotalScore;
    } else {
      if (b.avgTotalScore !== a.avgTotalScore) {
        return b.avgTotalScore - a.avgTotalScore;
      }
      return b.trustScorePercent - a.trustScorePercent;
    }
  });

  return summaryList.map((item, index) => ({
    ...item,
    stt: index + 1,
  }));
}