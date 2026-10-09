import { runWithConcurrency } from "../pool";

describe("runWithConcurrency", () => {
  it("should run every task once with at most `limit` at a time", async () => {
    let active = 0;
    let maxActive = 0;
    const seen: number[] = [];

    await runWithConcurrency(8, 3, async (index) => {
      active++;
      maxActive = Math.max(maxActive, active);
      await new Promise((resolve) => setTimeout(resolve, 2));
      seen.push(index);
      active--;
    });

    expect(seen.sort((a, b) => a - b)).toEqual([0, 1, 2, 3, 4, 5, 6, 7]);
    expect(maxActive).toBe(3);
  });

  it("should do nothing for zero tasks", async () => {
    const task = jest.fn();
    await runWithConcurrency(0, 3, task);
    expect(task).not.toHaveBeenCalled();
  });
});
