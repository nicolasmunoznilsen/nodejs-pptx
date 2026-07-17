export function partDir(partName: string): string {
  const normalized = normalizePackagePath(partName);
  const i = normalized.lastIndexOf('/');
  return i < 0 ? '' : normalized.slice(0, i);
}

export function relsPath(partName: string): string {
  const normalized = normalizePackagePath(partName);
  const i = normalized.lastIndexOf('/');
  return `${normalized.slice(0, i)}/_rels/${normalized.slice(i + 1)}.rels`;
}

export function normalizePackagePath(path: string): string {
  const out: string[] = [];
  for (const segment of path.replaceAll('\\', '/').split('/')) {
    if (!segment || segment === '.') continue;
    if (segment === '..') out.pop(); else out.push(segment);
  }
  return out.join('/');
}

export function resolveRelationshipTarget(sourcePart: string, target: string): string {
  if (target.startsWith('/')) return normalizePackagePath(target.slice(1));
  return normalizePackagePath(`${partDir(sourcePart)}/${target}`);
}

export function resolveRootRelationshipTarget(target: string): string {
  if (target.startsWith('/')) return normalizePackagePath(target.slice(1));
  return normalizePackagePath(target);
}
