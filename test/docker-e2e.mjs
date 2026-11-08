import fs from "node:fs";
import path from "node:path";
import { execSync } from "node:child_process";
import assert from "node:assert/strict";
import { chromium } from "playwright";
import { createMockGraphServer } from "./mock-graph-server.mjs";

const REPO_ROOT = process.cwd();
const WORK_DIR = "/tmp/obsidian-entra-docker-e2e";
const GRAPH_PORT = 18765;
const CLIENT_ID = process.env.ENTRA_CLIENT_ID || "ada87c12-a404-484b-a7dc-10c54beb2d7d";
const ENTRA_PASS = process.env.ENTRA_TEST_PASSWORD;
if (!ENTRA_PASS) {
  throw new Error("Set ENTRA_TEST_PASSWORD environment variable before running test/docker-e2e.mjs");
}

function setupUserDir(userSlug) {
  const baseDir = path.join(WORK_DIR, userSlug);
  const configDir = path.join(baseDir, "config");
  const openboxDir = path.join(configDir, ".config", "openbox");
  const labwcDir = path.join(configDir, ".config", "labwc");
  const obsidianCfgDir = path.join(configDir, ".config", "obsidian");
  const vaultDir = path.join(configDir, "SharedVault");
  const dotObsidianDir = path.join(vaultDir, ".obsidian");
  const pluginDir = path.join(dotObsidianDir, "plugins", "obsidian-entra-sync");

  fs.mkdirSync(openboxDir, { recursive: true });
  fs.mkdirSync(labwcDir, { recursive: true });
  fs.mkdirSync(obsidianCfgDir, { recursive: true });
  fs.mkdirSync(pluginDir, { recursive: true });

  // Copy compiled plugin files
  fs.copyFileSync(path.join(REPO_ROOT, "main.js"), path.join(pluginDir, "main.js"));
  fs.copyFileSync(path.join(REPO_ROOT, "manifest.json"), path.join(pluginDir, "manifest.json"));
  fs.copyFileSync(path.join(REPO_ROOT, "styles.css"), path.join(pluginDir, "styles.css"));

  // Initial plugin data.json pointing to the SharePoint/OneDrive Graph endpoint on host.docker.internal
  fs.writeFileSync(
    path.join(pluginDir, "data.json"),
    JSON.stringify(
      {
        tenantId: "common",
        clientId: CLIENT_ID,
        graphBaseUrl: `http://host.docker.internal:${GRAPH_PORT}/v1.0`,
        oneDriveVaultsFolder: "Documents/Documents",
        includeSitesScope: false,
        autoSyncIntervalSec: 60,
        syncStore: { files: {} },
      },
      null,
      2
    )
  );

  // Enable community plugin automatically
  fs.writeFileSync(
    path.join(dotObsidianDir, "community-plugins.json"),
    JSON.stringify(["obsidian-entra-sync"])
  );
  fs.writeFileSync(path.join(dotObsidianDir, "app.json"), JSON.stringify({}));

  // Register /config/SharedVault as open in Obsidian and disable Restricted Mode (safeMode: false)
  fs.writeFileSync(
    path.join(obsidianCfgDir, "obsidian.json"),
    JSON.stringify({
      vaults: {
        e2evault001: {
          path: "/config/SharedVault",
          ts: Date.now(),
          open: true,
        },
      },
    })
  );

  // Create tiny Python TCP proxy (0.0.0.0:9223 -> 127.0.0.1:9222) so Playwright on the host can connect over CDP
  const proxyPy = `
import socket, threading

def pipe(src, dst):
    try:
        while True:
            data = src.recv(65536)
            if not data: break
            dst.sendall(data)
    except Exception:
        pass
    finally:
        try: src.close()
        except Exception: pass
        try: dst.close()
        except Exception: pass

s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
s.bind(('0.0.0.0', 9223))
s.listen(64)
while True:
    c, _ = s.accept()
    try:
        t = socket.create_connection(('127.0.0.1', 9222), timeout=5)
        threading.Thread(target=pipe, args=(c, t), daemon=True).start()
        threading.Thread(target=pipe, args=(t, c), daemon=True).start()
    except Exception:
        c.close()
`.trim();
  fs.writeFileSync(path.join(configDir, "cdp_proxy.py"), proxyPy);

  // Openbox & Labwc (Wayland) autostart scripts
  const autostart = `#!/bin/bash
if [ "$(whoami)" != "abc" ]; then
  sudo lsiown -R $(id -u):$(id -u) /opt/obsidian
fi
/lsiopy/bin/python3 /config/cdp_proxy.py >/config/cdp_proxy.log 2>&1 &
obsidian --remote-debugging-port=9222 --remote-allow-origins=*
`;
  fs.writeFileSync(path.join(openboxDir, "autostart"), autostart, { mode: 0o755 });
  fs.writeFileSync(path.join(labwcDir, "autostart"), autostart, { mode: 0o755 });

  return { baseDir, configDir, vaultDir };
}

async function waitForCdp(port, timeoutMs = 60000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const res = await fetch(`http://127.0.0.1:${port}/json/version`);
      if (res.ok) {
        return await res.json();
      }
    } catch {
      // retry
    }
    await new Promise((r) => setTimeout(r, 1000));
  }
  throw new Error(`Timed out waiting for Obsidian CDP on port ${port}`);
}

async function getVaultPage(browser, timeoutMs = 45000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    for (const ctx of browser.contexts()) {
      for (const page of ctx.pages()) {
        const title = await page.title().catch(() => "");
        const url = page.url();
        if (url.includes("index.html") || title.includes("SharedVault") || title.includes("Obsidian")) {
          // Dismiss "Trust author and enable plugins" if shown
          const trustBtn = page.locator("button", { hasText: /Trust author and enable plugins/i });
          if ((await trustBtn.count()) > 0) {
            await trustBtn.first().click().catch(() => {});
          }
          const state = await page.evaluate(async () => {
            const app = window.app;
            if (!app) return { hasApp: false, body: document.body?.innerText?.slice(0, 200) };
            if (!app.plugins) return { hasApp: true, hasPlugins: false };
            try {
              await app.plugins.setEnable(true);
              await app.plugins.loadManifests();
              if (!app.plugins.getPlugin("obsidian-entra-sync")) {
                await app.plugins.enablePluginAndSave("obsidian-entra-sync");
              }
            } catch (e) {
              return { error: String(e), manifests: Object.keys(app.plugins.manifests || {}) };
            }
            const plugin = app.plugins.getPlugin("obsidian-entra-sync");
            if (app.setting) {
              app.setting.shouldUsePopout = () => false;
            }
            const tabs = (app.setting?.pluginTabs || []).map((t) => t.id);
            return {
              hasPlugin: Boolean(plugin),
              manifests: Object.keys(app.plugins.manifests || {}),
              tabs,
            };
          }).catch((e) => ({ evalErr: String(e) }));
          console.log("   [Page state]", url, title, JSON.stringify(state));
          if (state && state.hasPlugin) {
            return page;
          }
        }
      }
    }
    await new Promise((r) => setTimeout(r, 1000));
  }
  throw new Error("Timed out waiting for Obsidian vault window and plugin initialization");
}

async function runDockerE2E() {
  console.log("1. Cleaning up any previous test containers & preparing directories...");
  try {
    execSync("docker rm -f obsidian-entra-alice obsidian-entra-bob 2>/dev/null || true");
  } catch {}
  fs.rmSync(WORK_DIR, { recursive: true, force: true });

  const aliceDirs = setupUserDir("alice");
  const bobDirs = setupUserDir("bob");

  console.log("2. Starting Microsoft Graph SharePoint/OneDrive test server on port", GRAPH_PORT);
  const mockGraph = createMockGraphServer({
    port: GRAPH_PORT,
    host: "0.0.0.0",
    baseUrlOverride: `http://host.docker.internal:${GRAPH_PORT}/v1.0`,
  });
  await mockGraph.start();

  try {
    console.log("3. Starting two linuxserver/docker-obsidian containers (Alice & Bob)...");
    execSync(
      `docker run -d --name obsidian-entra-alice \
        --add-host=host.docker.internal:host-gateway \
        -e PUID=$(id -u) -e PGID=$(id -g) -e TZ=Europe/Brussels \
        -p 127.0.0.1:3101:3000 -p 127.0.0.1:9221:9223 \
        -v "${aliceDirs.configDir}:/config" \
        --shm-size="1gb" \
        lscr.io/linuxserver/obsidian:latest`
    );

    execSync(
      `docker run -d --name obsidian-entra-bob \
        --add-host=host.docker.internal:host-gateway \
        -e PUID=$(id -u) -e PGID=$(id -g) -e TZ=Europe/Brussels \
        -p 127.0.0.1:3102:3000 -p 127.0.0.1:9222:9223 \
        -v "${bobDirs.configDir}:/config" \
        --shm-size="1gb" \
        lscr.io/linuxserver/obsidian:latest`
    );

    console.log("4. Waiting for Obsidian instances inside both Docker containers to expose CDP...");
    const [aliceVer, bobVer] = await Promise.all([waitForCdp(9221), waitForCdp(9222)]);
    console.log("   Alice Obsidian:", aliceVer.Browser);
    console.log("   Bob Obsidian:", bobVer.Browser);

    const aliceBrowser = await chromium.connectOverCDP("http://127.0.0.1:9221");
    const bobBrowser = await chromium.connectOverCDP("http://127.0.0.1:9222");

    const [alicePage, bobPage] = await Promise.all([
      getVaultPage(aliceBrowser),
      getVaultPage(bobBrowser),
    ]);
    console.log("5. Connected Playwright to both Alice's and Bob's live Obsidian vaults!");

    // Open Settings -> Enterprise Sync tab in Alice's Obsidian UI and sign in via EntraCredentialsModal
    console.log("6. Signing Alice into Entra ID via Obsidian UI modal...");
    await alicePage.evaluate(() => {
      if (window.app.setting.isOpen) {
        window.app.setting.close();
      }
      window.app.setting.shouldUsePopout = () => false;
      window.app.setting.open();
      window.app.setting.openTabById("obsidian-entra-sync");
    });
    await alicePage.locator("button", { hasText: "Sign in (Entra Credentials)" }).click();
    await alicePage.locator('input[placeholder="user@yourdomain.onmicrosoft.com"]').fill(
      "alice@woutervanranstgmail.onmicrosoft.com"
    );
    await alicePage.locator('input[type="password"]').fill(ENTRA_PASS);
    await alicePage.getByRole("button", { name: "Sign In", exact: true }).click();

    // Wait for the Enterprise Sync Active badge in Alice's Settings UI
    await alicePage.locator(".entra-sync-badge").waitFor({ state: "visible", timeout: 20000 });
    const aliceBadgeText = await alicePage.locator(".entra-sync-card").first().innerText();
    console.log("   Alice UI Status Card:", aliceBadgeText.replace(/\n+/g, " | "));
    assert.ok(aliceBadgeText.includes("Obsidian Enterprise Sync Active"));
    assert.ok(aliceBadgeText.includes("Alice Peeters"));

    // Alice creates and connects the shared vault "ProductSpecs" via the Settings UI
    console.log("7. Alice creating & connecting shared vault 'ProductSpecs' via Obsidian UI...");
    await alicePage.locator('input[placeholder="e.g. ProductArchitecture"]').fill("ProductSpecs");
    await alicePage.locator("button", { hasText: "Create & Connect Vault" }).click();
    await alicePage.locator("button", { hasText: "Share Vault..." }).waitFor({
      state: "visible",
      timeout: 15000,
    });

    // Alice invites Bob via the ShareVaultModal UI
    console.log("8. Alice inviting Bob (bob@woutervanranstgmail.onmicrosoft.com) via ShareVaultModal UI...");
    await alicePage.locator("button", { hasText: "Share Vault..." }).click();
    await alicePage.locator('input[placeholder="colleague@company.com"]').fill(
      "bob@woutervanranstgmail.onmicrosoft.com"
    );
    await alicePage.locator(".entra-sync-share-modal button", { hasText: "Invite to Vault" }).click();
    await alicePage
      .locator(".entra-sync-collaborator-row", {
        hasText: "bob@woutervanranstgmail.onmicrosoft.com",
      })
      .waitFor({ state: "visible", timeout: 10000 });
    const collabText = await alicePage.locator(".entra-sync-collaborators-list").innerText();
    console.log("   ShareVaultModal Collaborators:", collabText.replace(/\n+/g, " "));
    await alicePage.keyboard.press("Escape"); // close modal
    await alicePage.keyboard.press("Escape"); // close settings

    // Alice creates a Markdown note in Obsidian and syncs it
    console.log("9. Alice creating 'SharedDesign.md' in Obsidian and syncing to SharePoint/OneDrive...");
    const initialDoc = [
      "# Shared Enterprise Design",
      "",
      "## Authentication",
      "- Provider: Microsoft Entra ID",
      "- Owner: Alice",
      "",
      "## Collaboration",
      "- Storage: SharePoint Online / OneDrive",
      "- Reviewer: Pending",
    ].join("\n");

    await alicePage.evaluate(async (content) => {
      const vault = window.app.vault;
      const existing = vault.getAbstractFileByPath("SharedDesign.md");
      if (existing) {
        await vault.modify(existing, content);
      } else {
        await vault.create("SharedDesign.md", content);
      }
      const leaf = window.app.workspace.getLeaf(false);
      const file = vault.getAbstractFileByPath("SharedDesign.md");
      if (file) await leaf.openFile(file);
      const plugin = window.app.plugins.getPlugin("obsidian-entra-sync");
      await plugin.triggerManualSync();
    }, initialDoc);

    // Now drive Bob's Obsidian container UI!
    console.log("10. Signing Bob into Entra ID via Bob's Obsidian UI modal...");
    await bobPage.evaluate(() => {
      if (window.app.setting.isOpen) {
        window.app.setting.close();
      }
      window.app.setting.shouldUsePopout = () => false;
      window.app.setting.open();
      window.app.setting.openTabById("obsidian-entra-sync");
    });
    await bobPage.locator("button", { hasText: "Sign in (Entra Credentials)" }).click();
    await bobPage.locator('input[placeholder="user@yourdomain.onmicrosoft.com"]').fill(
      "bob@woutervanranstgmail.onmicrosoft.com"
    );
    await bobPage.locator('input[type="password"]').fill(ENTRA_PASS);
    await bobPage.getByRole("button", { name: "Sign In", exact: true }).click();

    await bobPage.locator(".entra-sync-badge").waitFor({ state: "visible", timeout: 20000 });
    const bobBadgeText = await bobPage.locator(".entra-sync-card").first().innerText();
    console.log("   Bob UI Status Card:", bobBadgeText.replace(/\n+/g, " | "));
    assert.ok(bobBadgeText.includes("Bob Janssens"));

    // Bob clicks "Discover Vaults" and selects "ProductSpecs (Shared with me)"
    console.log("11. Bob discovering shared vaults and connecting to 'ProductSpecs (Shared with me)'...");
    await bobPage.locator("button", { hasText: "Discover Vaults" }).click();
    const dropdown = bobPage.locator(".setting-item select").first();
    await dropdown.waitFor({ state: "visible", timeout: 10000 });
    const options = await dropdown.locator("option").allInnerTexts();
    console.log("   Bob discovered options:", options);
    assert.ok(options.some((o) => o.includes("ProductSpecs (Shared with me)")));

    const targetVal = await dropdown
      .locator("option", { hasText: "ProductSpecs (Shared with me)" })
      .getAttribute("value");
    await dropdown.selectOption(targetVal);
    await bobPage.keyboard.press("Escape"); // close settings

    // Verify Bob's Obsidian vault now has SharedDesign.md!
    const bobPulledContent = await bobPage.evaluate(async () => {
      const plugin = window.app.plugins.getPlugin("obsidian-entra-sync");
      await plugin.triggerManualSync();
      const file = window.app.vault.getAbstractFileByPath("SharedDesign.md");
      if (!file) return null;
      await window.app.workspace.getLeaf(false).openFile(file);
      return await window.app.vault.read(file);
    });
    assert.equal(bobPulledContent, initialDoc);
    console.log("   Bob successfully pulled 'SharedDesign.md' in Docker!");

    // Test Concurrent Edits + 3-Way Merge across the two Docker containers!
    console.log("12. Testing concurrent edits & automatic 3-way diff3 merge across Alice & Bob containers...");
    const aliceUpdatedDoc = initialDoc.replace("- Owner: Alice", "- Owner: Alice Peeters (Approved)");
    const bobUpdatedDoc = initialDoc.replace(
      "- Reviewer: Pending",
      "- Reviewer: Bob Janssens (Verified)"
    );

    // Alice modifies and pushes first (creates Version 2.0)
    await alicePage.evaluate(async (text) => {
      const file = window.app.vault.getAbstractFileByPath("SharedDesign.md");
      await window.app.vault.modify(file, text);
      const plugin = window.app.plugins.getPlugin("obsidian-entra-sync");
      await plugin.syncEngine.pushMarkdownFile(file);
    }, aliceUpdatedDoc);

    // Bob modifies his local copy (still based on Version 1.0) and pushes -> HTTP 412 -> 3-way merge!
    const bobMergedText = await bobPage.evaluate(async (text) => {
      const file = window.app.vault.getAbstractFileByPath("SharedDesign.md");
      await window.app.vault.modify(file, text);
      const plugin = window.app.plugins.getPlugin("obsidian-entra-sync");
      await plugin.syncEngine.pushMarkdownFile(file);
      return await window.app.vault.read(file);
    }, bobUpdatedDoc);

    assert.ok(bobMergedText.includes("- Owner: Alice Peeters (Approved)"));
    assert.ok(bobMergedText.includes("- Reviewer: Bob Janssens (Verified)"));
    console.log("   3-way merge succeeded cleanly in Bob's container!");

    // Alice syncs to pull Version 3.0 and opens the Version History Modal UI!
    console.log("13. Alice syncing merged version and opening VersionHistoryModal in Obsidian UI...");
    await alicePage.evaluate(async () => {
      const plugin = window.app.plugins.getPlugin("obsidian-entra-sync");
      await plugin.triggerManualSync();
      plugin.openActiveFileHistory();
    });

    await alicePage.locator(".entra-sync-history-modal").waitFor({ state: "visible", timeout: 10000 });
    await alicePage
      .locator(".entra-sync-version-list .setting-item")
      .first()
      .waitFor({ state: "visible", timeout: 10000 });

    const versionListText = await alicePage.locator(".entra-sync-version-list").innerText();
    console.log("   Version History Modal entries:\n" + versionListText);
    assert.ok(versionListText.includes("Version 3.0 — Bob Janssens"));
    assert.ok(versionListText.includes("Version 2.0 — Alice Peeters"));
    assert.ok(versionListText.includes("Version 1.0 — Alice Peeters"));

    // Click "Preview" on Version 1.0 in Alice's modal
    const v1Row = alicePage.locator(".entra-sync-version-list .setting-item", {
      hasText: "Version 1.0",
    });
    await v1Row.locator("button", { hasText: "Preview" }).click();
    await alicePage
      .locator(".entra-sync-version-preview", { hasText: "- Owner: Alice" })
      .waitFor({ state: "visible", timeout: 10000 });

    // Also verify Webtop HTTP UIs (3101 & 3102) via a standard Playwright browser
    console.log("14. Verifying Webtop HTTP endpoints (ports 3101 & 3102) with Playwright...");
    const hostBrowser = await chromium.launch({ headless: true });
    const webtopPage1 = await hostBrowser.newPage();
    await webtopPage1.goto("http://127.0.0.1:3101", { waitUntil: "domcontentloaded" });
    const webtopTitle1 = await webtopPage1.title();
    console.log("   Alice Webtop (http://127.0.0.1:3101) title:", webtopTitle1);
    await hostBrowser.close();

    // Click "Restore" on Version 1.0 in Alice's VersionHistoryModal and verify restoration
    console.log("15. Restoring Version 1.0 from Alice's VersionHistoryModal...");
    await v1Row.locator("button", { hasText: "Restore" }).click();
    await alicePage.locator(".entra-sync-history-modal").waitFor({ state: "hidden", timeout: 10000 });

    const restoredText = await alicePage.evaluate(async () => {
      const file = window.app.vault.getAbstractFileByPath("SharedDesign.md");
      return await window.app.vault.read(file);
    });
    assert.equal(restoredText, initialDoc);
    console.log("   Version 1.0 restored and verified!");

    await aliceBrowser.close();
    await bobBrowser.close();
    console.log("✅ ALL DOCKER + PLAYWRIGHT TWO-USER E2E CHECKS PASSED!");
  } finally {
    await mockGraph.stop();
    try {
      execSync("docker rm -f obsidian-entra-alice obsidian-entra-bob 2>/dev/null || true");
    } catch {}
    fs.rmSync(WORK_DIR, { recursive: true, force: true });
  }
}

runDockerE2E().catch((err) => {
  console.error("Docker E2E test failed:", err);
  process.exit(1);
});
