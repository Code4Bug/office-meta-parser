import { readFile, writeFile } from 'fs/promises';
import type { Writable } from 'stream';

/**
 * 从文件读取为 ArrayBuffer
 */
export async function loadFromFile(path: string): Promise<ArrayBuffer> {
  const buf = await readFile(path);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

/**
 * 将 ArrayBuffer 写入文件
 */
export async function saveToFile(buffer: ArrayBuffer, path: string): Promise<void> {
  await writeFile(path, Buffer.from(buffer));
}

/**
 * 语义模型 → JSON 对象（剥离 rawXmlParts / extraParts 等不可序列化字段）
 */
export function toJSON<T>(semantic: T): T {
  return JSON.parse(JSON.stringify(semantic, (_key, value) => {
    if (value instanceof Map) return undefined;
    if (value instanceof ArrayBuffer) return undefined;
    if (value instanceof Buffer) return undefined;
    if (_key === 'rawXmlParts' || _key === 'extraParts' || _key === 'extraRels' || _key === 'extraEntries') return undefined;
    return value;
  }));
}

/**
 * 语义模型 → JSON 字符串
 */
export function toJSONString<T>(semantic: T, space?: number): string {
  return JSON.stringify(toJSON(semantic), null, space);
}

/**
 * 语义模型 → JSON 文件
 */
export async function saveToJSON<T>(semantic: T, path: string, space = 2): Promise<void> {
  await writeFile(path, toJSONString(semantic, space), 'utf-8');
}

/**
 * ArrayBuffer 转 Node.js Buffer
 */
export function toBuffer(buffer: ArrayBuffer): Buffer {
  return Buffer.from(buffer);
}

/**
 * 将 ArrayBuffer 写入可写流
 */
export async function writeToStream(buffer: ArrayBuffer, stream: Writable): Promise<void> {
  return new Promise((resolve, reject) => {
    stream.write(Buffer.from(buffer), (err) => {
      if (err) reject(err);
      else resolve();
    });
  });
}
