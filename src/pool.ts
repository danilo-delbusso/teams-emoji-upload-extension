// Run task(0..count-1) with at most `limit` running at once
export async function runWithConcurrency(
  count: number,
  limit: number,
  task: (index: number) => Promise<void>,
): Promise<void> {
  // A few workers pull indexes off a shared counter until none are left
  let next = 0;
  const worker = async () => {
    while (next < count) {
      await task(next++);
    }
  };
  await Promise.all(
    Array.from({ length: Math.min(Math.max(limit, 1), count) }, worker),
  );
}
