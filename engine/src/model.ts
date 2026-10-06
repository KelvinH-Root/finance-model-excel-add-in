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

/** The HFG entities a model can belong to; each brings its theme and logo. */
export type Brand = 'HF' | 'HCP' | 'HCL' | 'KM' | 'TWK';

/**
 * What the New model wizard collects. A model with it is built in the standard frame; a model
 * without it is laid out as the Phase 0 proof laid it out.
 */
export interface ModelInfo {
  title: string;
  entity: { name: string; brand: Brand };
  /** The contents header's third line, such as "Prepared by Group Finance". */
  preparedBy: string;
  notes: string[];
  timeline: {
    /** First month of the model, "yyyy-mm". */
    start: string;
    /** The month the financial year ends in, 1 to 12 (3 for a March year end). */
    fyEndMonth: number;
    /** The last month of actuals, as a period number; 0 when the model has none. */
    lastActual: number;
    denomination: '$' | '$000' | '$m';
  };
  /** Which status phrases the model name line shows. */
  display: { errors: boolean; alerts: boolean };
}

export interface ModelDict {
  periods: number;
  counters: Record<string, number>;
  instances: { module: string; number: number; settings: Settings }[];
  assurance?: Record<string, unknown>;
  info?: ModelInfo;
}

const clone = <T>(v: T): T => structuredClone(v);

export class Model {
  lib: Library;
  periods: number;
  instances: Instance[] = [];
  counters: Record<string, number> = {};
  /** Model assurance state: the group set in use, the latest version known, input records, the change log. */
  assurance: Record<string, any> = {};
  /** Set by the New model wizard; puts the model in the standard frame. */
  info: ModelInfo | null = null;

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
    m.info = this.info ? clone(this.info) : null;
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
    if (this.info) d.info = clone(this.info);
    return d;
  }

  static fromDict(lib: Library, d: ModelDict): Model {
    const m = new Model(lib, d.periods);
    m.counters = { ...d.counters };
    m.instances = d.instances.map(i => new Instance(i.module, i.number, clone(i.settings)));
    m.assurance = clone(d.assurance || {});
    m.info = d.info ? clone(d.info) : null;
    return m;
  }

  /** True when a module that brings the assurance sheets (framework: assurance) is in the model. */
  assured(): boolean {
    return this.instances.some(i => this.lib.module(i.module).framework === 'assurance');
  }
}
