// Minimal ambient types for qrcode-terminal (no @types package published).
declare module 'qrcode-terminal' {
  export function generate(input: string, opts?: { small?: boolean }): void;
  const qrcodeTerminal: { generate: typeof generate };
  export default qrcodeTerminal;
}
