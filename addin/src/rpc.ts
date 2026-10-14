// Dialogs and the runtime behind them. Office dialogs (DialogApi 1.2) cannot reach the workbook,
// so a dialog's view calls the model service through messages: the dialog sends
// {t: 'call', id, m, a}; the runtime that opened it runs the method and answers with
// {t: 'ret', id, ok, v | e}. A dialog can also ask to close, or to be replaced by another
// command's dialog (Group structure opening Add entity, say). One dialog is open at a time.

/* eslint-disable @typescript-eslint/no-explicit-any */
declare const Office: any;

type Call = { t: 'call'; id: number; m: string; a: unknown[] };
type Ret = { t: 'ret'; id: number; ok: boolean; v?: unknown; e?: string };
type Msg = Call | Ret | { t: 'close' } | { t: 'open'; key: string; arg?: string } | { t: 'done'; key: string; message: string };

/** Width and height in percent of the screen, by the view a command opens. */
export const DIALOG_SIZES: Record<string, [number, number]> = {
  explorer: [72, 82], group: [80, 84], insert: [58, 72], versions: [46, 74], wizard: [42, 84], impacts: [60, 84],
  entity: [40, 84], remove: [40, 52], delete: [40, 52], rename: [36, 44], properties: [56, 72], message: [34, 34], notbuilt: [36, 46],
};

export const dialogsSupported = () => typeof Office !== 'undefined' && !!Office?.context?.ui?.displayDialogAsync
  && !!Office?.context?.requirements?.isSetSupported('DialogApi', '1.2');

// ------------------------------------------------------------------ the runtime side

let current: any = null;

export interface Host {
  /** Runs a call from a dialog. */
  call(method: string, args: unknown[]): Promise<unknown>;
  /** A dialog asked for another command (key) to open in its place. */
  open(key: string, arg?: string): void;
  /** A dialog finished a command and reported what it did. */
  done?(key: string, message: string): void;
}

/** Open the dialog page for a command; resolves once it is showing. */
export function openDialog(url: string, size: [number, number], host: Host): Promise<void> {
  return new Promise((resolve, reject) => {
    const show = () => Office.context.ui.displayDialogAsync(url, { width: size[0], height: size[1], displayInIframe: true }, (res: any) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(`The dialog did not open (${res.error?.code}: ${res.error?.message}).`));
        return;
      }
      const d = res.value;
      current = d;
      d.addEventHandler(Office.EventType.DialogMessageReceived, async (arg: any) => {
        let msg: Msg;
        try { msg = JSON.parse(arg.message); } catch { return; }
        if (msg.t === 'call') {
          let reply: Ret;
          try {
            reply = { t: 'ret', id: msg.id, ok: true, v: await host.call(msg.m, msg.a) };
          } catch (e) {
            reply = { t: 'ret', id: msg.id, ok: false, e: (e as Error)?.message ?? String(e) };
          }
          try { d.messageChild(JSON.stringify(reply)); } catch { /* the dialog closed meanwhile */ }
        } else if (msg.t === 'close') {
          closeDialog();
        } else if (msg.t === 'open') {
          closeDialog();
          setTimeout(() => host.open(msg.key, msg.arg), 400);   // Office needs a moment before the next dialog
        } else if (msg.t === 'done') {
          host.done?.(msg.key, msg.message);
        }
      });
      d.addEventHandler(Office.EventType.DialogEventReceived, () => { if (current === d) current = null; });
      resolve();
    });
    if (current) {
      closeDialog();
      setTimeout(show, 400);
    } else {
      show();
    }
  });
}

export function closeDialog(): void {
  if (!current) return;
  try { current.close(); } catch { /* already closed */ }
  current = null;
}

// ------------------------------------------------------------------ the dialog side

/** A stand-in for the model service inside a dialog: each method call goes to the runtime that opened it. */
export function remote<T extends object>(): T {
  const waiting = new Map<number, { res: (v: unknown) => void; rej: (e: Error) => void }>();
  let next = 1;
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg: any) => {
    let msg: Msg;
    try { msg = JSON.parse(arg.message); } catch { return; }
    if (msg.t !== 'ret') return;
    const w = waiting.get(msg.id);
    if (!w) return;
    waiting.delete(msg.id);
    if (msg.ok) w.res(msg.v);
    else w.rej(new Error(msg.e));
  });
  return new Proxy({}, {
    get: (_t, m) => (...a: unknown[]) => new Promise((res, rej) => {
      const id = next++;
      waiting.set(id, { res, rej });
      Office.context.ui.messageParent(JSON.stringify({ t: 'call', id, m: String(m), a } satisfies Call));
    }),
  }) as T;
}

/** Ask the runtime to close this dialog (outside Office, the page just says so). */
export function askClose(): void {
  if (typeof Office !== 'undefined' && Office?.context?.ui?.messageParent) Office.context.ui.messageParent(JSON.stringify({ t: 'close' }));
}

/** Ask the runtime to open another command's dialog in place of this one. */
export function askOpen(key: string, arg?: string): void {
  if (typeof Office !== 'undefined' && Office?.context?.ui?.messageParent) Office.context.ui.messageParent(JSON.stringify({ t: 'open', key, arg }));
}

export function tellDone(key: string, message: string): void {
  if (typeof Office !== 'undefined' && Office?.context?.ui?.messageParent) Office.context.ui.messageParent(JSON.stringify({ t: 'done', key, message }));
}
