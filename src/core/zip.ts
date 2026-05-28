import JSZip from 'jszip';
import type { ZipEntry } from './types.js';

export async function unzip(buffer: ArrayBuffer): Promise<ZipEntry[]> {
  const archive = await JSZip.loadAsync(buffer);
  const entries: ZipEntry[] = [];

  const promises: Promise<void>[] = [];
  archive.forEach((path, file) => {
    if (!file.dir) {
      const p = file.async('arraybuffer').then(data => {
        entries.push({ path, data });
      });
      promises.push(p);
    }
  });

  await Promise.all(promises);

  // Sort by path for deterministic order
  entries.sort((a, b) => a.path.localeCompare(b.path));

  return entries;
}

export async function zip(entries: ZipEntry[]): Promise<ArrayBuffer> {
  const archive = new JSZip();

  for (const entry of entries) {
    archive.file(entry.path, entry.data);
  }

  return archive.generateAsync({ type: 'arraybuffer' });
}
