// Autosave for big workbooks without freezing: instead of writing the whole
// .xlsx after every edit (seconds for a million cells), each call that changes
// the document is appended to a journal in IndexedDB — a few milliseconds.
// Opening replays the journal on top of the last saved file. The full file is
// rewritten ("compacted") only now and then; see SheetsApp's autosave.
//
// Which calls change the document? Everything except the read-only methods
// listed below, so a newly added method is journaled by default (safe: at
// worst a read-only call is replayed for nothing).

import type { SheetsDocument } from "@office/engine";
import { readJournal, writeJournal } from "../documents";

const READ_ONLY = new Set([
  "autoFilter",
  "canSort",
  "cfBlock",
  "cfRules",
  "chartSourceRange",
  "chartViewJson",
  "checkInput",
  "currentRegion",
  "filterColumn",
  "filterCounts",
  "filterItems",
  "invalidCells",
  "listItems",
  "validationAt",
  "columnWidths",
  "displayBlock",
  "drawingsJson",
  "extent",
  "find",
  "freezePanes",
  "hasCopied",
  "hiddenColumns",
  "input",
  "jump",
  "kindBlock",
  "mergeDiscardsData",
  "pictureData",
  "merges",
  "rangeStats",
  "rowLayout",
  "rowsToMeasure",
  "sheetNames",
  "spillRange",
  "styleBlock",
  "toCsv",
  "toXlsx",
  "free",
  "constructor",
]);

/** One recorded call: [method name, arguments]. Arguments are structured-cloneable. */
export type JournalEntry = [string, unknown[]];

export class Journal {
  private entries: JournalEntry[] = [];
  private writing: Promise<boolean> = Promise.resolve(true);

  constructor(
    readonly docId: string,
    existing: JournalEntry[] = [],
  ) {
    this.entries = existing;
  }

  get length(): number {
    return this.entries.length;
  }

  record(method: string, args: unknown[]) {
    this.entries.push([method, args]);
  }

  /** Store the journal; resolves to false if the browser refused. */
  flush(): Promise<boolean> {
    const snapshot = this.entries.slice();
    this.writing = this.writing.then(() => writeJournal(this.docId, snapshot));
    return this.writing;
  }

  /**
   * After the full file was saved: drop the first `count` entries, which the
   * saved file already contains. Entries recorded meanwhile are kept.
   */
  dropFirst(count: number): Promise<boolean> {
    this.entries = this.entries.slice(count);
    return this.flush();
  }

  static async load(docId: string): Promise<Journal> {
    const stored = (await readJournal(docId)) as JournalEntry[];
    return new Journal(docId, Array.isArray(stored) ? stored : []);
  }

  /** Re-apply recorded calls to a freshly opened document. */
  replay(doc: SheetsDocument) {
    const target = doc as unknown as Record<string, (...args: unknown[]) => unknown>;
    for (const [method, args] of this.entries) {
      try {
        target[method]?.(...args);
      } catch (err) {
        console.warn(`Skipped a journal entry (${method})`, err);
      }
    }
  }
}

/** Wrap a document so every call that may change it is recorded in `journal`. */
export function journaled(doc: SheetsDocument, journal: Journal): SheetsDocument {
  return new Proxy(doc, {
    get(target, prop) {
      const value = Reflect.get(target, prop, target);
      if (typeof value !== "function" || typeof prop !== "string") return value;
      if (READ_ONLY.has(prop)) return value.bind(target);
      return (...args: unknown[]) => {
        const result = value.apply(target, args);
        journal.record(prop, args);
        return result;
      };
    },
  });
}
