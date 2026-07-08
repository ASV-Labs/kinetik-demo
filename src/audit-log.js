export function createAuditEvent({ actorId, action, targetType, targetId, metadata = {} }) {
  if (!actorId || !action || !targetType || !targetId) {
    throw new Error("audit event requires actor, action, and target");
  }

  return {
    id: `aud-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    actorId,
    action,
    targetType,
    targetId,
    metadata,
    createdAt: new Date().toISOString()
  };
}

export function recordReleaseAudit(release, action, actorId = "system") {
  return createAuditEvent({
    actorId,
    action,
    targetType: "release",
    targetId: release.id,
    metadata: {
      owner: release.owner,
      status: release.status,
      risk: release.risk
    }
  });
}

