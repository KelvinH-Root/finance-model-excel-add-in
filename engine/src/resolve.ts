// Link resolution (steps 3, 5 and 6 of an insert): which blocks the model has and who sends each link.

import { AssemblyError } from './frame.ts';
import type { ModuleDef } from './library.ts';
import type { Instance, Model } from './model.ts';

export interface Block {
  id: string;
  inst: Instance;
  mod: ModuleDef;
  title: string;
  /** [producer block id, row key] for a mirror block. */
  src: [string, string] | null;
}

/** Link -> [sending block, row key], in the order the links first appear. */
export type Producers = Map<string, [Block, string][]>;

function producersOf(blocks: Block[]): Producers {
  const out: Producers = new Map();
  const add = (link: string, b: Block, row: string) => {
    const list = out.get(link);
    if (list) list.push([b, row]);
    else out.set(link, [[b, row]]);
  };
  for (const b of blocks) {
    for (const o of b.mod.outputs || []) add(o.link, b, o.row);
    for (const r of b.mod.rows || []) {
      if (r.check) add(`check.${r.check}`, b, r.key as string);
    }
  }
  return out;
}

/** Blocks in sheet order, repeated until mirror blocks stop producing new ones. */
export function resolve(model: Model): { blocks: Block[]; producers: Producers } {
  const lib = model.lib;
  const order = new Map(model.instances.map((i, k) => [i.uid, k]));
  const areaOf = (i: Instance) => lib.areas.indexOf(lib.module(i.module).area);
  const ordered = [...model.instances].sort((a, b) => areaOf(a) - areaOf(b) || order.get(a.uid)! - order.get(b.uid)!);
  let blocks: Block[] = [];
  for (let pass = 0; pass < 25; pass++) {   // mirror blocks can produce links other mirrors take
    const producers = producersOf(blocks);
    const next: Block[] = [];
    for (const inst of ordered) {
      const mod = lib.module(inst.module);
      const title = model.title(inst);
      if (mod.mirror) {
        for (const [pb, row] of producers.get(mod.mirror) || []) {
          if (pb.inst.uid === inst.uid) continue;
          next.push({ id: `${inst.uid}/${pb.id}`, inst, mod, title: `${title}: ${pb.title}`, src: [pb.id, row] });
        }
      } else {
        next.push({ id: inst.uid, inst, mod, title, src: null });
      }
    }
    if (next.length === blocks.length && next.every((b, k) => b.id === blocks[k].id)) return { blocks, producers };
    blocks = next;
  }
  throw new AssemblyError('links did not settle after 25 passes (a mirror feeds itself)');
}
