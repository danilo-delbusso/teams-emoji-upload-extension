import { RateLimiter, RateLimitError, rateLimitError } from "../rateLimit";

const response429 = (body: string, retryAfter?: string) =>
  ({
    status: 429,
    text: async () => body,
    headers: {
      get: (name: string) =>
        name === "retry-after" ? (retryAfter ?? null) : null,
    },
  }) as unknown as Response;

describe("rateLimitError", () => {
  it("should read Teams' throttling details", async () => {
    const error = await rateLimitError(
      response429(
        '{"Throttled":true,"Dimension":"User","Limit":10,"WindowMinutes":1,"RetryAfterMinutes":1}',
      ),
      "deletions",
    );

    expect(error).toBeInstanceOf(RateLimitError);
    expect(error.message).toBe(
      "Teams allows 10 deletions per minute. Try again in 60s.",
    );
    expect(error.retryAfterMs).toBe(60_000);
    expect(error.limit).toBe(10);
    expect(error.windowMs).toBe(60_000);
  });

  it("should fall back to the Retry-After header or a default", async () => {
    expect(
      (await rateLimitError(response429("", "90"), "x")).retryAfterMs,
    ).toBe(90_000);
    const error = await rateLimitError(response429("not json"), "x");
    expect(error.retryAfterMs).toBe(60_000);
    expect(error.message).toBe(
      "Teams is limiting how fast requests can be made. Try again in 60s.",
    );
  });
});

describe("RateLimiter", () => {
  let now: number;
  let waits: number[];
  const limiter = (limit: number, windowMs: number) =>
    new RateLimiter(
      limit,
      windowMs,
      () => now,
      async (ms) => {
        waits.push(ms);
        now += ms;
      },
    );

  beforeEach(() => {
    now = 0;
    waits = [];
  });

  it("should let `limit` requests through, then wait for the window", async () => {
    const rl = limiter(2, 1000);
    await rl.acquire();
    now += 100;
    await rl.acquire();
    expect(waits).toEqual([]);

    const onWait = jest.fn();
    await rl.acquire(onWait);
    // The first request started at 0, so the third may start at 1000
    expect(now).toBe(1000);
    expect(onWait).toHaveBeenCalledWith(900);
  });

  it("should pause for as long as Teams asked and adopt its limits", async () => {
    const rl = limiter(10, 60_000);
    rl.backOff(new RateLimitError("slow down", 5000, 1, 2000));

    await rl.acquire();
    expect(now).toBe(5000);
    // New limit of 1 per 2s applies from now on
    await rl.acquire();
    expect(now).toBe(7000);
  });

  it("should count down in steps of at most a second", async () => {
    const rl = limiter(10, 60_000);
    rl.backOff(new RateLimitError("slow down", 2500));
    const onWait = jest.fn();
    await rl.acquire(onWait);
    expect(onWait.mock.calls.map(([ms]) => ms)).toEqual([2500, 1500, 500]);
  });
});
