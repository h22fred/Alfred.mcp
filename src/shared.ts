import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Validate a Dynamics GUID. Throws if invalid — prevents path manipulation in API URLs. */
export function requireGuid(value: string, label: string): string {
  if (!GUID_RE.test(value)) throw new Error(`Invalid ${label}: expected a Dynamics GUID.`);
  return value;
}

/** Create a progress callback that logs to stderr and sends MCP logging messages. */
export function makeProgress(srv: McpServer) {
  return (msg: string) => {
    console.error(`[progress] ${msg}`);
    srv.server.sendLoggingMessage({ level: "info", data: msg });
  };
}

/** Simple in-process write rate limiter — max N writes per rolling window. */
export class WriteRateLimiter {
  private timestamps: number[] = [];
  constructor(private readonly max: number, private readonly windowMs: number) {}
  check(action: string): void {
    const now = Date.now();
    this.timestamps = this.timestamps.filter(t => now - t < this.windowMs);
    if (this.timestamps.length >= this.max) {
      throw new Error(
        `Rate limit: no more than ${this.max} ${action} operations per ${this.windowMs / 60000} minutes. ` +
        `Please review what you are doing before continuing.`
      );
    }
    this.timestamps.push(now);
  }
}

export const FORECAST_NAMES: Record<number, string> = {
  100000001: "Pipeline",
  100000002: "Best Case",
  100000003: "Committed",
  100000004: "Omitted",
};
