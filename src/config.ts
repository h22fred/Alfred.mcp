import { readFileSync, existsSync } from "fs";
import { homedir } from "os";
import { join } from "path";

interface AlfredConfig {
  dynamicsUrl?: string;
  teamsWebhook?: string;
  role?: "sc" | "ssc" | "manager";
  engagementTypes?: string[];
  installedVersion?: string;
}

const configPath = join(homedir(), ".alfred-config.json");
const raw: AlfredConfig = existsSync(configPath)
  ? JSON.parse(readFileSync(configPath, "utf8")) as AlfredConfig
  : {};

/** Base URL of the customer's Dynamics 365 instance, e.g. https://acme.crm.dynamics.com */
export const DYNAMICS_HOST: string =
  raw.dynamicsUrl ?? "https://servicenow.crm.dynamics.com";

export const alfredConfig = raw;
