declare class Buffer extends Uint8Array {
  static from(input: string | ArrayBuffer | Uint8Array): Buffer;
}
declare module 'node:fs/promises' {
  export function access(path: string): Promise<void>;
  export function readFile(path: string): Promise<Buffer>;
  export function writeFile(path: string, data: string | Uint8Array): Promise<void>;
  export function mkdtemp(prefix: string): Promise<string>;
}
declare module 'node:path' {
  export function basename(path: string): string;
  export function join(...parts: string[]): string;
}
declare module 'node:os' { export function tmpdir(): string; }
declare module 'jszip' {
  export default class JSZip {
    files: Record<string, unknown>;
    static loadAsync(data: Buffer | Uint8Array | ArrayBuffer): Promise<JSZip>;
    file(name: string): { async(type: 'string'): Promise<string>; async(type: 'nodebuffer'): Promise<Buffer> } | null;
    file(name: string, data: string | Uint8Array): this;
    generateAsync(options: { type: 'nodebuffer'; compression?: string }): Promise<Buffer>;
  }
}
declare module 'fast-xml-parser' {
  export class XMLParser { constructor(options?: Record<string, unknown>); parse(xml: string): unknown; }
  export class XMLBuilder { constructor(options?: Record<string, unknown>); build(value: unknown): string; }
}
declare module 'vitest' {
  export function describe(name: string, fn: () => void): void;
  export function it(name: string, fn: () => unknown | Promise<unknown>): void;
  export const expect: any;
}
