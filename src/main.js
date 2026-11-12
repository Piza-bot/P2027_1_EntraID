import {
  InteractionRequiredAuthError,
  PublicClientApplication,
} from "@azure/msal-browser";
import {
  driveMetricRequest,
  isAuthConfigured,
  loginRequest,
  mailApiRequest,
  msalConfig,
} from "./authConfig.js";
import { getStorageMetricsUrl } from "./storageMetricsUrl.js";

const elements = {
  signIn: document.querySelector("#sign-in"),
  signOut: document.querySelector("#sign-out"),
  nav: document.querySelector("#app-nav"),
  testsToggle: document.querySelector("#tests-toggle"),
  testsMenu: document.querySelector("#tests-menu"),
  analyticNav: document.querySelector("#onedrive-analytic-nav"),
  authBadge: document.querySelector("#auth-badge"),
  signedOutLayout: document.querySelector("#signed-out-layout"),
  signedOutPanel: document.querySelector("#signed-out-panel"),
  signedInPanel: document.querySelector("#signed-in-panel"),
  setupMessage: document.querySelector("#setup-message"),
  mailSubject: document.querySelector("#mail-subject"),
  profileMessage: document.querySelector("#profile-message"),
  profileName: document.querySelector("#profile-name"),
  profileAccount: document.querySelector("#profile-account"),
  profileEmail: document.querySelector("#profile-email"),
  profileJob: document.querySelector("#profile-job"),
  mailForm: document.querySelector("#mail-form"),
  mailTo: document.querySelector("#mail-to"),
  mailBody: document.querySelector("#mail-body"),
  sendMail: document.querySelector("#send-mail"),
  mailMessage: document.querySelector("#mail-message"),
  loadMyDriveMetric: document.querySelector("#load-my-drive-metric"),
  myDriveAccount: document.querySelector("#my-drive-account"),
  myDriveMetricCards: document.querySelector("#my-drive-metric-cards"),
  myDriveTotal: document.querySelector("#my-drive-total"),
  myDriveUsed: document.querySelector("#my-drive-used"),
  myDriveRemaining: document.querySelector("#my-drive-remaining"),
  myDriveDeleted: document.querySelector("#my-drive-deleted"),
  myDriveMetricMessage: document.querySelector("#my-drive-metric-message"),
  myDriveMetricMeta: document.querySelector("#my-drive-metric-meta"),
  myDriveQuotaChart: document.querySelector("#my-drive-quota-chart"),
  myDriveStorageLink: document.querySelector("#my-drive-storage-link"),
  driveAnalyticForm: document.querySelector("#drive-analytic-form"),
  driveAnalyticAccount: document.querySelector("#drive-analytic-account"),
  loadDriveAnalytic: document.querySelector("#load-drive-analytic"),
  driveAnalyticMetricCards: document.querySelector("#drive-analytic-metric-cards"),
  analyticDriveTotal: document.querySelector("#analytic-drive-total"),
  analyticDriveUsed: document.querySelector("#analytic-drive-used"),
  analyticDriveRemaining: document.querySelector("#analytic-drive-remaining"),
  analyticDriveDeleted: document.querySelector("#analytic-drive-deleted"),
  driveAnalyticMessage: document.querySelector("#drive-analytic-message"),
  driveAnalyticMeta: document.querySelector("#drive-analytic-meta"),
  analyticQuotaChart: document.querySelector("#analytic-quota-chart"),
  analyticStorageLink: document.querySelector("#analytic-storage-link"),
  appStatus: document.querySelector("#app-status"),
  result: document.querySelector("#result"),
};

const msalInstance = new PublicClientApplication(msalConfig);
let canUseDriveAnalytics = false;

function setStatus(message, state = "ready") {
  elements.appStatus.textContent = message;
  elements.result.dataset.state = state;
}

function describeError(error) {
  return error?.errorMessage || error?.message || "เกิดข้อผิดพลาดที่ไม่ทราบสาเหตุ";
}

async function readApiJson(response, action) {
  const body = await response.text();
  let result;
  try {
    result = body ? JSON.parse(body) : null;
  } catch {
    const isHtml = /text\/html/i.test(response.headers.get("content-type") || "")
      || /^\s*<(?:!doctype html|html)\b/i.test(body);
    if (isHtml) {
      if (/<pre>Cannot (?:POST|GET) \/api\/[^<]*<\/pre>/i.test(body)) {
        throw new Error(
          `Node API ที่ทำงานอยู่ยังไม่มี endpoint นี้ (HTTP ${response.status}) ต้อง deploy หรือ restart Node API ด้วย server/index.js เวอร์ชันล่าสุด`,
        );
      }
      throw new Error(
        `API ส่งหน้าเว็บ HTML กลับมาแทนข้อมูล (HTTP ${response.status}) ตรวจสอบว่า Node API เวอร์ชันล่าสุดทำงานอยู่และ Reverse Proxy ส่ง /api/ ไปยัง Node API`,
      );
    }
    throw new Error(`API ส่งข้อมูลที่ไม่ใช่ JSON กลับมา (HTTP ${response.status})`);
  }

  if (!response.ok) {
    throw new Error(result?.error || `${action} ไม่สำเร็จ (HTTP ${response.status})`);
  }
  if (!result || typeof result !== "object") {
    throw new Error("API ส่งข้อมูลกลับมาไม่ครบถ้วน");
  }
  return result;
}

function showSignedOut(message, isError = false) {
  canUseDriveAnalytics = false;
  elements.analyticNav.hidden = true;
  elements.signedOutLayout.hidden = false;
  elements.signedOutPanel.hidden = false;
  elements.signedInPanel.hidden = true;
  elements.nav.hidden = true;
  elements.signIn.hidden = false;
  elements.signOut.hidden = true;
  elements.testsMenu.hidden = true;
  elements.testsToggle.setAttribute("aria-expanded", "false");
  elements.authBadge.textContent = isError ? "เชื่อมต่อไม่สำเร็จ" : "ยังไม่ลงชื่อเข้าใช้";
  elements.authBadge.dataset.state = isError ? "error" : "setup";
  if (message) elements.setupMessage.textContent = message;
}

function showSignedIn(account) {
  elements.signedOutLayout.hidden = true;
  elements.signedOutPanel.hidden = true;
  elements.signedInPanel.hidden = false;
  elements.nav.hidden = false;
  elements.signIn.hidden = true;
  elements.signOut.hidden = false;
  setActiveView("home");
  elements.authBadge.textContent = "ลงชื่อเข้าใช้แล้ว";
  elements.authBadge.dataset.state = "signed-in";
  elements.profileName.textContent = account.name || "บัญชีองค์กร";
  elements.profileAccount.textContent = account.username;
  elements.profileEmail.textContent = "กำลังโหลด...";
  elements.profileJob.textContent = "กำลังโหลด...";
  elements.myDriveAccount.textContent = `บัญชีที่ลงชื่อเข้าใช้: ${account.username}`;
}

function setActiveView(viewName) {
  if (viewName === "onedrive-metric-analytic" && !canUseDriveAnalytics) {
    viewName = "home";
  }
  document.querySelectorAll(".workspace-view").forEach((view) => {
    view.hidden = view.id !== `${viewName}-view`;
  });
  document.querySelectorAll("[data-view]").forEach((button) => {
    if (button.dataset.view === viewName) button.setAttribute("aria-current", "page");
    else button.removeAttribute("aria-current");
  });
  const viewLabels = {
    home: "หน้าแรก",
    "my-onedrive-metric": "MyOneDrive Metric",
    "onedrive-metric-analytic": "OneDrive Metric Analytic",
  };
  elements.appStatus.textContent = viewLabels[viewName] || "พร้อมใช้งาน";
  elements.result.dataset.state = "signed-in";
}

elements.nav.addEventListener("click", (event) => {
  const viewButton = event.target.closest("[data-view]");
  if (viewButton) {
    setActiveView(viewButton.dataset.view);
    elements.testsMenu.hidden = true;
    elements.testsToggle.setAttribute("aria-expanded", "false");
  }
});

elements.testsToggle.addEventListener("click", () => {
  const isExpanded = elements.testsToggle.getAttribute("aria-expanded") === "true";
  elements.testsToggle.setAttribute("aria-expanded", String(!isExpanded));
  elements.testsMenu.hidden = isExpanded;
});

document.addEventListener("click", (event) => {
  if (!event.target.closest(".nav-dropdown")) {
    elements.testsMenu.hidden = true;
    elements.testsToggle.setAttribute("aria-expanded", "false");
  }
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    elements.testsMenu.hidden = true;
    elements.testsToggle.setAttribute("aria-expanded", "false");
    elements.testsToggle.focus();
  }
});

async function getAccessToken(account, request = loginRequest) {
  try {
    return await msalInstance.acquireTokenSilent({ ...request, account });
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError)) throw error;
    return msalInstance.acquireTokenPopup({ ...request, account });
  }
}

async function loadDriveAnalyticsAccess(account) {
  canUseDriveAnalytics = false;
  elements.analyticNav.hidden = true;
  try {
    const tokenResponse = await getAccessToken(account, mailApiRequest);
    const response = await fetch("/api/drive-metrics/access", {
      headers: { Authorization: `Bearer ${tokenResponse.accessToken}` },
    });
    const result = await readApiJson(response, "ตรวจสอบสิทธิ์ผู้ดูแล");
    canUseDriveAnalytics = result.allowed === true;
    elements.analyticNav.hidden = !canUseDriveAnalytics;
  } catch (error) {
    elements.analyticNav.hidden = true;
    console.error("OneDrive analytics access check failed:", describeError(error));
  }
}

function formatQuotaValue(bytes) {
  if (bytes === null || bytes === undefined || bytes === "") return "ไม่ทราบ";
  const value = Number(bytes);
  if (!Number.isFinite(value) || value < 0) return "ไม่ทราบ";
  return `${new Intl.NumberFormat("th-TH", { maximumFractionDigits: 2 }).format(value / 1_000_000_000)} GB`;
}

function renderQuota(quota, fields) {
  fields.total.textContent = formatQuotaValue(quota?.total);
  fields.used.textContent = formatQuotaValue(quota?.used);
  fields.remaining.textContent = formatQuotaValue(quota?.remaining);
  fields.deleted.textContent = formatQuotaValue(quota?.deleted);
}

function renderQuotaChart(quota, prefix, container) {
  const total = Number(quota?.total);
  const used = Number(quota?.used);
  const remaining = Number(quota?.remaining);
  if (![total, used, remaining].every(Number.isFinite) || total <= 0) {
    container.hidden = true;
    return;
  }

  const usedPercent = Math.min(100, Math.max(0, (used / total) * 100));
  const remainingPercent = 100 - usedPercent;
  const bar = document.querySelector(`#${prefix}-quota-bar`);
  bar.setAttribute(
    "aria-label",
    `ใช้พื้นที่ ${usedPercent.toFixed(1)} เปอร์เซ็นต์ ใช้งาน ${formatQuotaValue(used)} จาก ${formatQuotaValue(total)}`,
  );
  bar.querySelector(".quota-used").style.width = `${usedPercent}%`;
  bar.querySelector(".quota-remaining").style.width = `${remainingPercent}%`;
  document.querySelector(`#${prefix}-chart-total`).textContent = formatQuotaValue(total);
  document.querySelector(`#${prefix}-chart-percent`).textContent = `ใช้งาน ${usedPercent.toFixed(1)}%`;
  document.querySelector(`#${prefix}-chart-used`).textContent = formatQuotaValue(used);
  document.querySelector(`#${prefix}-chart-remaining`).textContent = formatQuotaValue(remaining);
  document.querySelector(`#${prefix}-chart-deleted`).textContent = formatQuotaValue(Number(quota.deleted));
  container.hidden = false;
}

function setStorageMetricsLink(link, webUrl) {
  const storageUrl = getStorageMetricsUrl(webUrl);
  link.hidden = !storageUrl;
  if (!storageUrl) return;
  link.href = storageUrl;
}

function formatFetchedAt(date = new Date()) {
  return new Intl.DateTimeFormat("th-TH", {
    dateStyle: "medium",
    timeStyle: "medium",
  }).format(date);
}

elements.loadMyDriveMetric.addEventListener("click", async () => {
  const account = msalInstance.getActiveAccount();
  if (!account) return;

  elements.loadMyDriveMetric.disabled = true;
  elements.myDriveMetricCards.hidden = true;
  elements.myDriveQuotaChart.hidden = true;
  elements.myDriveStorageLink.hidden = true;
  elements.myDriveMetricMeta.textContent = "";
  elements.myDriveMetricMessage.dataset.state = "pending";
  elements.myDriveMetricMessage.textContent = "กำลังอ่านโควตาจาก Microsoft Graph...";
  try {
    const tokenResponse = await getAccessToken(account, driveMetricRequest);
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me/drive?$select=quota,driveType,webUrl",
      { headers: { Authorization: `Bearer ${tokenResponse.accessToken}` } },
    );
    if (!response.ok) {
      const permissionHint = response.status === 403
        ? " ตรวจสอบว่า App Registration ได้รับ delegated permission Files.Read และ consent แล้ว"
        : "";
      throw new Error(`Microsoft Graph ตอบกลับ HTTP ${response.status}.${permissionHint}`);
    }

    const drive = await response.json();
    if (!drive.quota) throw new Error("Microsoft Graph ไม่ส่งข้อมูลโควตาของ OneDrive กลับมา");
    renderQuota(drive.quota, {
      total: elements.myDriveTotal,
      used: elements.myDriveUsed,
      remaining: elements.myDriveRemaining,
      deleted: elements.myDriveDeleted,
    });
    elements.myDriveMetricCards.hidden = false;
    renderQuotaChart(drive.quota, "my-drive", elements.myDriveQuotaChart);
    elements.myDriveMetricMessage.dataset.state = "success";
    elements.myDriveMetricMessage.textContent = `อ่านข้อมูลสำเร็จ${drive.quota.state ? ` · สถานะ ${drive.quota.state}` : ""}`;
    elements.myDriveMetricMeta.textContent = `บัญชี: ${account.username} · อ่านเมื่อ ${formatFetchedAt()}`;
    setStorageMetricsLink(elements.myDriveStorageLink, drive.webUrl);
  } catch (error) {
    elements.myDriveMetricMessage.dataset.state = "error";
    elements.myDriveMetricMessage.textContent = `อ่านข้อมูล OneDrive ไม่สำเร็จ: ${describeError(error)}`;
  } finally {
    elements.loadMyDriveMetric.disabled = false;
  }
});

elements.driveAnalyticForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const account = msalInstance.getActiveAccount();
  if (!account) return;

  elements.loadDriveAnalytic.disabled = true;
  elements.driveAnalyticMetricCards.hidden = true;
  elements.analyticQuotaChart.hidden = true;
  elements.analyticStorageLink.hidden = true;
  elements.driveAnalyticMeta.textContent = "";
  elements.driveAnalyticMessage.dataset.state = "pending";
  elements.driveAnalyticMessage.textContent = "กำลังตรวจสอบสิทธิ์และค้นหาข้อมูล...";
  try {
    const tokenResponse = await getAccessToken(account, mailApiRequest);
    const response = await fetch("/api/drive-metrics", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${tokenResponse.accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ account: elements.driveAnalyticAccount.value.trim() }),
    });
    const result = await readApiJson(response, "ค้นหาข้อมูล");

    renderQuota(result.quota, {
      total: elements.analyticDriveTotal,
      used: elements.analyticDriveUsed,
      remaining: elements.analyticDriveRemaining,
      deleted: elements.analyticDriveDeleted,
    });
    elements.driveAnalyticMetricCards.hidden = false;
    renderQuotaChart(result.quota, "analytic", elements.analyticQuotaChart);
    elements.driveAnalyticMessage.dataset.state = "success";
    elements.driveAnalyticMessage.textContent = `อ่านข้อมูลของ ${result.user.displayName || result.user.userPrincipalName} สำเร็จ${result.quota.state ? ` · สถานะ ${result.quota.state}` : ""}`;
    elements.driveAnalyticMeta.textContent = `UPN: ${result.user.userPrincipalName || "-"} · อีเมล: ${result.user.mail || "-"} · อ่านเมื่อ ${formatFetchedAt()}`;
    setStorageMetricsLink(elements.analyticStorageLink, result.storageMetricsUrl);
  } catch (error) {
    elements.driveAnalyticMessage.dataset.state = "error";
    elements.driveAnalyticMessage.textContent = `ค้นหาข้อมูล OneDrive ไม่สำเร็จ: ${describeError(error)}`;
  } finally {
    elements.loadDriveAnalytic.disabled = false;
  }
});

elements.mailForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const account = msalInstance.getActiveAccount();
  if (!account) return;

  elements.sendMail.disabled = true;
  elements.mailMessage.dataset.state = "pending";
  elements.mailMessage.textContent = "กำลังส่งอีเมล...";
  try {
    const tokenResponse = await getAccessToken(account, mailApiRequest);
    const response = await fetch("/api/send-mail", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${tokenResponse.accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        to: elements.mailTo.value,
        subject: elements.mailSubject.value,
        message: elements.mailBody.value,
      }),
    });
    const responseBody = await response.text();
    let result = null;
    if (responseBody) {
      try {
        result = JSON.parse(responseBody);
      } catch {
        if (response.ok) throw new Error(`API ตอบกลับข้อมูลไม่ถูกต้อง (HTTP ${response.status})`);
        result = { error: responseBody };
      }
    }
    if (!response.ok) {
      const accountHint = result?.account ? ` บัญชีที่ตรวจพบ: ${result.account}` : "";
      throw new Error(`${result?.error || `ส่งอีเมลไม่สำเร็จ (HTTP ${response.status})`}${accountHint}`);
    }

    elements.mailForm.reset();
    elements.mailMessage.dataset.state = "success";
    elements.mailMessage.textContent = "ส่งอีเมลเรียบร้อยแล้ว";
  } catch (error) {
    elements.mailMessage.dataset.state = "error";
    elements.mailMessage.textContent = `ส่งอีเมลไม่สำเร็จ: ${describeError(error)}`;
  } finally {
    elements.sendMail.disabled = false;
  }
});

async function loadGraphProfile(account) {
  elements.profileMessage.textContent = "กำลังอ่านโปรไฟล์จาก Microsoft Graph...";
  try {
    const tokenResponse = await getAccessToken(account);
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle",
      { headers: { Authorization: `Bearer ${tokenResponse.accessToken}` } },
    );

    if (!response.ok) {
      throw new Error(`Microsoft Graph ตอบกลับ HTTP ${response.status}`);
    }

    const profile = await response.json();
    elements.profileName.textContent = profile.displayName || account.name || "บัญชีองค์กร";
    elements.profileAccount.textContent = profile.userPrincipalName || account.username;
    elements.profileEmail.textContent = profile.mail || profile.userPrincipalName || "ไม่มีข้อมูลอีเมล";
    elements.profileJob.textContent = profile.jobTitle || "ไม่มีข้อมูลตำแหน่ง";
    elements.profileMessage.textContent = "โหลดข้อมูลจาก Microsoft Graph แล้ว";
    setStatus("เชื่อมต่อ Graph แล้ว", "signed-in");
  } catch (error) {
    elements.profileMessage.textContent = `อ่านโปรไฟล์ไม่สำเร็จ: ${describeError(error)}`;
    setStatus("เรียก Graph ไม่สำเร็จ", "error");
  }
}

elements.signIn.addEventListener("click", async () => {
  elements.signIn.disabled = true;
  setStatus("กำลังเปิดหน้าลงชื่อเข้าใช้...");
  try {
    await msalInstance.loginRedirect(loginRequest);
  } catch (error) {
    showSignedOut(`ลงชื่อเข้าใช้ไม่สำเร็จ: ${describeError(error)}`, true);
    setStatus("ลงชื่อเข้าใช้ไม่สำเร็จ", "error");
  } finally {
    elements.signIn.disabled = false;
  }
});

elements.signOut.addEventListener("click", async () => {
  elements.signOut.disabled = true;
  try {
    await msalInstance.logoutPopup({
      account: msalInstance.getActiveAccount(),
      postLogoutRedirectUri: msalConfig.auth.postLogoutRedirectUri,
    });
    showSignedOut("ออกจากระบบแล้ว");
    setStatus("ออกจากระบบแล้ว");
  } catch (error) {
    elements.profileMessage.textContent = `ออกจากระบบไม่สำเร็จ: ${describeError(error)}`;
    setStatus("ออกจากระบบไม่สำเร็จ", "error");
  } finally {
    elements.signOut.disabled = false;
  }
});

async function start() {
  if (!isAuthConfigured) {
    showSignedOut("ตั้งค่า Client ID และ Tenant ID ใน src/authConfig.js ก่อนเริ่มใช้งาน");
    elements.signIn.disabled = true;
    setStatus("รอตั้งค่า Entra ID");
    return;
  }

  try {
    await msalInstance.initialize();
    const redirectResponse = await msalInstance.handleRedirectPromise();
    if (redirectResponse?.account) msalInstance.setActiveAccount(redirectResponse.account);

    const accounts = msalInstance.getAllAccounts();
    const account = msalInstance.getActiveAccount() || accounts[0];

    if (!account) {
      showSignedOut("ลงชื่อเข้าใช้ด้วยบัญชีองค์กรเพื่อดูโปรไฟล์ของคุณ");
      elements.signIn.disabled = false;
      setStatus("พร้อมลงชื่อเข้าใช้");
      return;
    }

    msalInstance.setActiveAccount(account);
    showSignedIn(account);
    await Promise.all([
      loadGraphProfile(account),
      loadDriveAnalyticsAccess(account),
    ]);
  } catch (error) {
    showSignedOut(`เริ่มระบบยืนยันตัวตนไม่สำเร็จ: ${describeError(error)}`, true);
    setStatus("เริ่มระบบไม่สำเร็จ", "error");
  }
}

start();