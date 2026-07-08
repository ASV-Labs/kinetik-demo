export const releases = [
  {
    id: "rel-2026-07",
    title: "July Readiness Report",
    owner: "operations",
    status: "draft",
    risk: "medium",
    changes: ["customer export cleanup", "status summary endpoint"]
  }
];

export function listReleases() {
  return releases.map((release) => ({ ...release }));
}

