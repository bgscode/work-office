import { describe, expect, it } from "vitest";
import { Journal, journaled } from "./journal";

class FakeDoc {
  calls: string[] = [];
  value = 0;
  setInputs(n: number) {
    this.value = n;
    this.calls.push(`set ${n}`);
  }
  displayBlock() {
    return [String(this.value)];
  }
}

describe("journal", () => {
  it("records changing calls, not reads, and replays them", () => {
    const journal = new Journal("doc");
    const doc = journaled(new FakeDoc() as never, journal) as unknown as FakeDoc;
    doc.setInputs(5);
    doc.displayBlock();
    doc.setInputs(7);
    expect(journal.length).toBe(2);
    const copy = new FakeDoc();
    journal.replay(copy as never);
    expect(copy.value).toBe(7);
    expect(copy.calls).toEqual(["set 5", "set 7"]);
  });
});
