import { Page } from "playwright";
import { BotConfig } from "../../types";
import { log, callStartupCallback } from "../../utils";
import { hasStopSignalReceived } from "../../index";
import { callStatusChangeCallback } from "../../services/unified-callback";

export type AdmissionDecision = {
  admitted: boolean;
  rejected?: boolean;
  reason?: string;
};

export type AdmissionResult = boolean | AdmissionDecision;

export type LeaveReason =
  | "admission_rejected_by_admin"
  | "admission_timeout"
  | "removed_by_admin"
  | "left_alone_timeout"
  | "startup_alone_timeout"
  | "normal_completion"
  | string;

function generateReasonTokens(platform: string): {
  removedToken: string;
  leftAloneToken: string;
  startupAloneToken: string;
} {
  const platformUpper = platform.toUpperCase();
  return {
    removedToken: `${platformUpper}_BOT_REMOVED_BY_ADMIN`,
    leftAloneToken: `${platformUpper}_BOT_LEFT_ALONE_TIMEOUT`,
    startupAloneToken: `${platformUpper}_BOT_STARTUP_ALONE_TIMEOUT`
  };
}

export type PlatformStrategies = {
  join: (page: Page, botConfig: BotConfig) => Promise<void>;
  waitForAdmission: (page: Page, timeoutMs: number, botConfig: BotConfig) => Promise<AdmissionResult>;
  checkAdmissionSilent: (page: Page) => Promise<boolean>; // Silent check without callbacks
  prepare: (page: Page, botConfig: BotConfig) => Promise<void>;
  startRecording: (page: Page, botConfig: BotConfig) => Promise<void>;
  startRemovalMonitor: (page: Page, onRemoval?: () => void | Promise<void>) => () => void;
  leave: (page: Page | null, botConfig?: BotConfig, reason?: LeaveReason) => Promise<boolean>;
};

function normalizeMeetingTitleCandidate(value: unknown): string | null {
  if (typeof value !== "string") return null;
  let v = value.replace(/\s+/g, " ").trim();
  if (!v) return null;

  // Strip common Teams noise / suffixes
  v = v.replace(/\s*(\||-)\s*Microsoft Teams\s*$/i, "").trim();

  // Reject obvious generic titles
  const lower = v.toLowerCase();
  if (!v || v.length < 3) return null;
  if (lower === "microsoft teams" || lower === "microsoft teams meeting") return null;
  if (lower === "join conversation" || lower === "join meeting") return null;
  if (lower.includes("meetup-join")) return null;

  if (v.length > 180) v = v.slice(0, 180).trim();
  return v || null;
}

async function bestEffortExtractMeetingTitle(page: Page): Promise<string | null> {
  // Best-effort. Never throw.
  try {
    const t = normalizeMeetingTitleCandidate(await page.title().catch(() => ""));
    if (t) return t;
  } catch {}

  try {
    const domTitle = await page.evaluate(() => {
      const pick = (s: string) => (typeof s === "string" ? s.replace(/\s+/g, " ").trim() : "");

      const metaOg = document.querySelector('meta[property="og:title"]') as HTMLMetaElement | null;
      const metaTitle = document.querySelector('meta[name="title"]') as HTMLMetaElement | null;
      const metaTwitter = document.querySelector('meta[name="twitter:title"]') as HTMLMetaElement | null;
      const titleTag = document.querySelector("title");

      const headings = Array.from(document.querySelectorAll('h1,h2,[role="heading"]'));
      const headingText = headings
        .map((el) => pick((el as HTMLElement).innerText || (el as HTMLElement).textContent || ""))
        .filter(Boolean);

      const candidates = [
        pick(metaOg?.content || ""),
        pick(metaTwitter?.content || ""),
        pick(metaTitle?.content || ""),
        pick(titleTag?.textContent || ""),
        ...headingText,
      ].filter(Boolean);

      return candidates[0] || "";
    });

    const t = normalizeMeetingTitleCandidate(domTitle);
    if (t) return t;
  } catch {}

  return null;
}

export async function runMeetingFlow(
  platform: string,
  botConfig: BotConfig,
  page: Page,
  gracefulLeaveFunction: (page: Page | null, exitCode: number, reason: string, errorDetails?: any) => Promise<void>,
  strategies: PlatformStrategies
): Promise<void> {
  const tokens = generateReasonTokens(platform);
  if (!botConfig.meetingUrl) {
    log(`Error: Meeting URL is required for ${platform} but is null.`);
    await gracefulLeaveFunction(page, 1, "missing_meeting_url");
    return;
  }

  // Join
  try {
    await strategies.join(page, botConfig);
  } catch (error: any) {
    const errorDetails = {
      error_message: error?.message,
      error_stack: error?.stack,
      error_name: error?.name,
      context: "join_meeting_error",
      platform,
      timestamp: new Date().toISOString()
    };
    await gracefulLeaveFunction(page, 1, "join_meeting_error", errorDetails);
    return;
  }

  // Stop-signal guard
  if (hasStopSignalReceived()) {
    log("⛔ Stop signal detected before admission wait. Exiting without joining.");
    await gracefulLeaveFunction(page, 0, "stop_requested_pre_admission");
    return;
  }

  // Admission + prepare in parallel
  try {
    const [admissionResult] = await Promise.all([
      strategies
        .waitForAdmission(page, botConfig.automaticLeave.waitingRoomTimeout, botConfig)
        .catch((error: any) => {
          const msg: string = error?.message || String(error);
          if (msg.includes("rejected by meeting admin")) {
            return { admitted: false, rejected: true, reason: "admission_rejected_by_admin" } as AdmissionDecision;
          }
          return { admitted: false, rejected: false, reason: "admission_timeout" } as AdmissionDecision;
        }),
      strategies.prepare(page, botConfig),
    ]);

    const isAdmitted = admissionResult === true || (typeof admissionResult === "object" && !!(admissionResult as AdmissionDecision).admitted);
    if (!isAdmitted) {
      const decision: AdmissionDecision = typeof admissionResult === "object"
        ? (admissionResult as AdmissionDecision)
        : { admitted: false, reason: "admission_timeout" };

      if (decision.rejected) {
        await gracefulLeaveFunction(page, 0, decision.reason || "admission_rejected_by_admin");
        return;
      }

      // Attempt stateless leave before graceful exit
      try {
        const result = await page.evaluate(async () => {
          if (typeof (window as any).performLeaveAction === "function") {
            return await (window as any).performLeaveAction();
          }
          return false;
        });
        if (result) log("✅ Successfully performed graceful leave during admission timeout");
      } catch {}

      await gracefulLeaveFunction(page, 0, decision.reason || "admission_timeout");
      return;
    }

    // CRITICAL: If bot was immediately admitted, ensure AWAITING_ADMISSION state is processed before ACTIVE
    // The waitForAdmission function sends AWAITING_ADMISSION callback when immediately admitted,
    // but we need to wait a moment for the state machine to process that transition before sending ACTIVE
    log("Bot admitted - ensuring AWAITING_ADMISSION state is processed before sending ACTIVE...");
    await new Promise(resolve => setTimeout(resolve, 1000)); // Wait 1 second for state transition
    
    // Startup callback (sends ACTIVE status)
    try {
      // Try to capture meeting title early (may still be unavailable at this moment)
      const earlyTitle = platform === "teams" ? await bestEffortExtractMeetingTitle(page) : null;
      if (earlyTitle) {
        log(`[Meeting Title] Early capture: "${earlyTitle}"`);
      }

      await callStartupCallback(botConfig, earlyTitle);
      
      // CRITICAL: Verify bot is still in meeting after callback (prevent false positives)
      // Use silent check to avoid sending AWAITING_ADMISSION callback again
      log("Verifying bot is still in meeting after ACTIVE callback...");
      const stillAdmitted = await strategies.checkAdmissionSilent(page);
      if (!stillAdmitted) {
        log("🚨 Bot is NOT in meeting after ACTIVE callback - false positive detected!");
        await gracefulLeaveFunction(page, 0, "admission_false_positive");
        return;
      }
      log("✅ Bot verified to be in meeting after ACTIVE callback");

      // Post-admission title capture with retries: Teams often renders subject only inside the call UI.
      if (platform === "teams") {
        const MAX_MS = 60000;
        const INTERVAL_MS = 2000;
        const start = Date.now();
        let lastSeen: string | null = earlyTitle || null;

        while ((Date.now() - start) < MAX_MS) {
          const t = await bestEffortExtractMeetingTitle(page);
          if (t && t !== lastSeen) {
            lastSeen = t;
            log(`[Meeting Title] Captured after admission: "${t}" — sending update callback`);
            try {
              // Send a duplicate ACTIVE callback with meeting_title so bot-manager can persist it.
              await callStatusChangeCallback(botConfig, "active", "meeting_title_discovered", undefined, undefined, undefined, undefined, t);
            } catch (e: any) {
              log(`[Meeting Title] Failed to send title callback: ${e?.message || String(e)}`);
            }
            break;
          }
          await new Promise(resolve => setTimeout(resolve, INTERVAL_MS));
        }

        if (!lastSeen) {
          log("[Meeting Title] Not captured within retry window (best-effort).");
        }
      }
    } catch (error: any) {
      log(`Error during startup callback or verification: ${error?.message || String(error)}`);
      // Continue to recording phase even if callback/verification fails
    }

    // Removal monitoring + recording race
    let signalRemoval: (() => void) | null = null;
    const removalPromise = new Promise<never>((_, reject) => {
      signalRemoval = () => reject(new Error(tokens.removedToken));
    });
    const stopRemoval = strategies.startRemovalMonitor(page, () => { if (signalRemoval) signalRemoval(); });

    try {
      await Promise.race([
        strategies.startRecording(page, botConfig),
        removalPromise
      ]);

      // Normal completion
      await gracefulLeaveFunction(page, 0, "normal_completion");
    } catch (error: any) {
      const msg: string = error?.message || String(error);
      if (msg === tokens.removedToken || msg.includes(tokens.removedToken)) {
        await gracefulLeaveFunction(page, 0, "removed_by_admin");
        return;
      }
      if (msg === tokens.leftAloneToken || msg.includes(tokens.leftAloneToken)) {
        await gracefulLeaveFunction(page, 0, "left_alone_timeout");
        return;
      }
      if (msg === tokens.startupAloneToken || msg.includes(tokens.startupAloneToken)) {
        await gracefulLeaveFunction(page, 0, "startup_alone_timeout");
        return;
      }

      const errorDetails = {
        error_message: error?.message,
        error_stack: error?.stack,
        error_name: error?.name,
        context: "post_join_setup_error",
        platform,
        timestamp: new Date().toISOString()
      };
      await gracefulLeaveFunction(page, 1, "post_join_setup_error", errorDetails);
      return;
    } finally {
      stopRemoval();
    }
  } catch (error: any) {
    const msg: string = error?.message || String(error);
    if (msg.includes(tokens.removedToken)) {
      await gracefulLeaveFunction(page, 0, "removed_by_admin");
      return;
    }
    if (msg.includes(tokens.leftAloneToken)) {
      await gracefulLeaveFunction(page, 0, "left_alone_timeout");
      return;
    }
    if (msg.includes(tokens.startupAloneToken)) {
      await gracefulLeaveFunction(page, 0, "startup_alone_timeout");
      return;
    }

    const errorDetails = {
      error_message: error?.message,
      error_stack: error?.stack,
      error_name: error?.name,
      context: "post_join_setup_error",
      platform,
      timestamp: new Date().toISOString()
    };
    await gracefulLeaveFunction(page, 1, "post_join_setup_error", errorDetails);
  }
}


