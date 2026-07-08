const demoSessions = new Map([
  ["demo-owner-token", { userId: "user-owner", role: "owner" }],
  ["demo-reviewer-token", { userId: "user-reviewer", role: "reviewer" }]
]);

export function readSession(request) {
  const header = request.headers.authorization ?? "";
  const match = header.match(/^Bearer (?<token>[A-Za-z0-9._-]+)$/);
  if (!match?.groups?.token) {
    return undefined;
  }

  return demoSessions.get(match.groups.token);
}

export function requireReleaseApprover(request) {
  const session = readSession(request);
  if (!session) {
    return { allowed: false, status: 401, error: "unauthorized" };
  }

  if (session.role !== "owner" && session.role !== "reviewer") {
    return { allowed: false, status: 403, error: "forbidden" };
  }

  return { allowed: true, session };
}

