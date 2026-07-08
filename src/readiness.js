const riskWeights = {
  low: 15,
  medium: 45,
  high: 75
};

export function summarizeReadiness(releases) {
  return releases.map((release) => {
    const changeCount = release.changes.length;
    const baseRisk = riskWeights[release.risk] ?? riskWeights.medium;
    const readinessScore = Math.max(0, 100 - baseRisk - Math.max(0, changeCount - 2) * 5);

    return {
      id: release.id,
      title: release.title,
      status: release.status,
      owner: release.owner,
      readinessScore,
      risk: release.risk,
      changeCount
    };
  });
}

