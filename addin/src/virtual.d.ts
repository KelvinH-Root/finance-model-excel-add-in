// Modules the build fills in: the library bundle and the entity logos.
declare module 'virtual:library' {
  import type { LibraryBundle } from '../../engine/src/library.ts';
  const bundle: LibraryBundle;
  export default bundle;
}
declare module 'virtual:logos' {
  const logos: Record<string, { base64: string; width: number; height: number }>;
  export default logos;
}
