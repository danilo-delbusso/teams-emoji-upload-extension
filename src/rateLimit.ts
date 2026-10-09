// Thrown when Teams answers 429; carries how long to wait before retrying
export class RateLimitError extends Error {
  constructor(
    message: string,
    public readonly retryAfterMs: number,
    public readonly limit?: number,
    public readonly windowMs?: number,
  ) {
    super(message);
    this.name = "RateLimitError";
  }
}

const DEFAULT_RETRY_MS = 60_000;

// Teams answers 429 with a body like
// {"Throttled":true,"Dimension":"User","Limit":10,"WindowMinutes":1,"RetryAfterMinutes":1}
export async function rateLimitError(
  response: Response,
  action: string,
): Promise<RateLimitError> {
  let limit: number | undefined;
  let windowMinutes: number | undefined;
  let retryAfterMs = DEFAULT_RETRY_MS;

  try {
    const body = JSON.parse(await response.text());
    if (typeof body.Limit === "number") limit = body.Limit;
    if (typeof body.WindowMinutes === "number")
      windowMinutes = body.WindowMinutes;
    if (typeof body.RetryAfterMinutes === "number") {
      retryAfterMs = body.RetryAfterMinutes * 60_000;
    }
  } catch {
    // Not JSON; fall back to the header or the default
  }
  const header = Number(response.headers?.get?.("retry-after"));
  if (header > 0) {
    retryAfterMs = Math.max(retryAfterMs, header * 1000);
  }

  const rule =
    limit && windowMinutes
      ? `Teams allows ${limit} ${action} per ${windowMinutes === 1 ? "minute" : `${windowMinutes} minutes`}`
      : "Teams is limiting how fast requests can be made";
  return new RateLimitError(
    `${rule}. Try again in ${Math.ceil(retryAfterMs / 1000)}s.`,
    retryAfterMs,
    limit,
    windowMinutes ? windowMinutes * 60_000 : undefined,
  );
}

// Keeps requests under `limit` per rolling `windowMs`, and honours server pauses
export class RateLimiter {
  private starts: number[] = [];
  private pausedUntil = 0;

  constructor(
    private limit: number,
    private windowMs: number,
    private now: () => number = () => Date.now(),
    private sleep: (ms: number) => Promise<void> = (ms) =>
      new Promise((resolve) => setTimeout(resolve, ms)),
  ) {}

  // Resolves once another request may start; onWait gets the time left in ms
  async acquire(onWait?: (msLeft: number) => void): Promise<void> {
    for (;;) {
      const now = this.now();
      this.starts = this.starts.filter((start) => now - start < this.windowMs);
      const until = Math.max(
        this.pausedUntil,
        this.starts.length >= this.limit ? this.starts[0] + this.windowMs : 0,
      );
      if (until <= now) {
        this.starts.push(now);
        return;
      }
      onWait?.(until - now);
      // Re-check every second so callers can show a countdown
      await this.sleep(Math.min(until - now, 1000));
    }
  }

  // Adopt the limits Teams reported and wait as long as it asked
  backOff(error: RateLimitError) {
    if (error.limit) this.limit = error.limit;
    if (error.windowMs) this.windowMs = error.windowMs;
    this.pausedUntil = Math.max(
      this.pausedUntil,
      this.now() + error.retryAfterMs,
    );
  }
}
