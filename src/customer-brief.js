export function buildCustomerBrief(release) {
  const impact =
    release.risk === "high"
      ? "visible operational changes"
      : "routine product improvements";

  return {
    title: `${release.title} is ready for customer review`,
    audience: "customer_success",
    summary: `This release includes ${release.changes.length} changes focused on ${release.owner}.`,
    impact,
    citations: release.changes.map((change, index) => ({
      id: `change-${index + 1}`,
      label: change
    }))
  };
}

