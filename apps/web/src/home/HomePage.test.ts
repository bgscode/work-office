import { describe, expect, it } from "vitest";
import { relativeTime } from "./HomePage";

describe("relativeTime", () => {
  const now = Date.UTC(2026, 9, 21, 12, 0, 0);
  it("describes recent times", () => {
    expect(relativeTime(now - 10_000, now)).toBe("Just now");
    expect(relativeTime(now - 5 * 60_000, now)).toBe("5 minutes ago");
    expect(relativeTime(now - 3 * 3_600_000, now)).toBe("3 hours ago");
    expect(relativeTime(now - 26 * 3_600_000, now)).toBe("Yesterday");
    expect(relativeTime(now - 3 * 86_400_000, now)).toBe("3 days ago");
  });
});
