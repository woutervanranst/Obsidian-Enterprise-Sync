export type LogLevel = "INFO" | "WARN" | "ERROR";

export interface LogEntry {
  timestamp: string;
  level: LogLevel;
  message: string;
  details?: string;
}

export class SyncLogger {
  private entries: LogEntry[] = [];
  private maxEntries = 250;

  log(level: LogLevel, message: string, details?: unknown): void {
    const timestamp = new Date().toISOString();
    let detailsStr: string | undefined;
    if (details !== undefined && details !== null) {
      if (typeof details === "string") {
        detailsStr = details;
      } else {
        try {
          detailsStr = JSON.stringify(details);
        } catch {
          detailsStr = String(details);
        }
      }
    }

    const entry: LogEntry = { timestamp, level, message, details: detailsStr };
    this.entries.push(entry);
    if (this.entries.length > this.maxEntries) {
      this.entries.shift();
    }

    const consoleLine = `[EntraSync] [${timestamp}] [${level}] ${message}${
      detailsStr ? ` | ${detailsStr}` : ""
    }`;
    if (level === "ERROR") {
      console.error(consoleLine);
    } else if (level === "WARN") {
      console.warn(consoleLine);
    } else {
      console.log(consoleLine);
    }
  }

  info(message: string, details?: unknown): void {
    this.log("INFO", message, details);
  }

  warn(message: string, details?: unknown): void {
    this.log("WARN", message, details);
  }

  error(message: string, details?: unknown): void {
    this.log("ERROR", message, details);
  }

  getEntries(): LogEntry[] {
    return [...this.entries];
  }

  formatText(): string {
    if (this.entries.length === 0) {
      return "No diagnostic log entries recorded yet.";
    }
    return this.entries
      .map(
        (e) =>
          `[${e.timestamp}] [${e.level}] ${e.message}${
            e.details ? `\n  -> ${e.details}` : ""
          }`
      )
      .join("\n");
  }

  clear(): void {
    this.entries = [];
  }
}

export const defaultLogger = new SyncLogger();
