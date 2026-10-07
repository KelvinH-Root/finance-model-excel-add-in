// Modules the build fills in: the libraries, the whole-model recipes and the entity logos.
declare module 'virtual:library' {
  import type { LibraryBundle, Recipe } from '../../engine/src/index.ts';
  const bundle: {
    libraries: Record<string, LibraryBundle>;
    recipes: { id: string; label: string; note: string; library: string; model: Recipe }[];
  };
  export default bundle;
}
declare module 'virtual:logos' {
  const logos: Record<string, { base64: string; width: number; height: number }>;
  export default logos;
}
