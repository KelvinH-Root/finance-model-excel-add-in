// A model: the module instances a workbook is built from, plus its assurance state.

import { AssemblyError } from './frame.ts';
import type { Library } from './library.ts';

export type Settings = Record<string, unknown>;

export class Instance {
  module: string;
  number: number;
  settings: Settings;

  constructor(module: string, number: number, settings: Settings) {
    this.module = module;
    this.number = number;
    this.settings = settings;
  }

  get uid(): string {
    return `${this.module}#${this.number}`;
  }
}

export interface ModelDict {
  periods: number;
  counters: Record<string, number>;
  instances: { module: string; number: number; settings: Settings }[];
  assurance?: Record<string, unknown>;
}

const clone = <T>(v: T): T => structuredClone(v);

export class Model {
  lib: Library;
  periods: number;
  instances: Instance[] = [];
  counters: Record<string, number> = {};
  /** Model assurance state: the group set in use, the latest version known, input records, the change log. */
  assurance: Record<string, any> = {};

  constructor(lib: Library, periods = 12) {
    this.lib = lib;
    this.periods = periods;
  }

  /** Steps 1 (compatibility) and 2 (instance naming) of an insert. */
  insert(moduleId: string, settings: Settings = {}): Instance {
    if (!this.lib.modules.has(moduleId)) throw new AssemblyError(`no module '${moduleId}' in the library`);
    const mod = this.lib.module(moduleId);
    if (this.lib.kind(moduleId) !== 'category' && this.instances.some(i => i.module === moduleId)) {
      throw new AssemblyError(`${mod.title} is already in the model and can only be inserted once`);
    }
    const known = new Map((mod.settings || []).map(s => [s.key, s]));
    const unknown = Object.keys(settings).filter(k => !known.has(k)).sort();
    if (unknown.length) throw new AssemblyError(`${mod.title} has no setting ${unknown.join(', ')}`);
    const values: Settings = {};
    for (const [k, s] of known) values[k] = s.default === undefined ? null : s.default;
    Object.assign(values, settings);
    const n = (this.counters[moduleId] || 0) + 1;   // numbers are never reused
    this.counters[moduleId] = n;
    const inst = new Instance(moduleId, n, values);
    this.instances.push(inst);
    return inst;
  }

  instance(uid: string): Instance {
    const i = this.instances.find(x => x.uid === uid);
    if (!i) throw new AssemblyError(`no instance ${uid} in the model`);
    return i;
  }

  /** Point a setting at the group assumption its module names for it. */
  bind(uid: string, key: string): void {
    const inst = this.instance(uid);
    const spec = (this.lib.module(inst.module).settings || []).find(s => s.key === key);
    if (!spec) throw new AssemblyError(`${this.title(inst)} has no setting ${key}`);
    if (!spec.group) throw new AssemblyError(`${this.title(inst)}: ${spec.label} has no group assumption`);
    inst.settings[key] = { group: spec.group.item };
  }

  /** Use a local value in place of the group assumption. */
  unbind(uid: string, key: string, value: number): void {
    this.instance(uid).settings[key] = value;
  }

  remove(uid: string): void {
    const before = this.instances.length;
    this.instances = this.instances.filter(i => i.uid !== uid);
    if (this.instances.length === before) throw new AssemblyError(`no instance ${uid} in the model`);
  }

  copy(): Model {
    const m = new Model(this.lib, this.periods);
    m.instances = this.instances.map(i => new Instance(i.module, i.number, clone(i.settings)));
    m.counters = { ...this.counters };
    m.assurance = clone(this.assurance);
    return m;
  }

  title(inst: Instance): string {
    const mod = this.lib.module(inst.module);
    return this.lib.kind(inst.module) === 'category' ? `${mod.title} ${inst.number}` : mod.title;
  }

  toDict(): ModelDict {
    const d: ModelDict = {
      periods: this.periods,
      counters: { ...this.counters },
      instances: this.instances.map(i => ({ module: i.module, number: i.number, settings: clone(i.settings) })),
    };
    if (Object.keys(this.assurance).length) d.assurance = clone(this.assurance);
    return d;
  }

  static fromDict(lib: Library, d: ModelDict): Model {
    const m = new Model(lib, d.periods);
    m.counters = { ...d.counters };
    m.instances = d.instances.map(i => new Instance(i.module, i.number, clone(i.settings)));
    m.assurance = clone(d.assurance || {});
    return m;
  }

  /** True when a module that brings the assurance sheets (framework: assurance) is in the model. */
  assured(): boolean {
    return this.instances.some(i => this.lib.module(i.module).framework === 'assurance');
  }
}
