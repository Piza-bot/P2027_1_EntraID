import {
  InteractionRequiredAuthError,
  PublicClientApplication,
} from "@azure/msal-browser";
import { isAuthConfigured, loginRequest, mailApiRequest, msalConfig } from "./authConfig.js";

const elements = {
  signIn: document.querySelector("#sign-in"),
  signOut: document.querySelector("#sign-out"),
  authBadge: document.querySelector("#auth-badge"),
  signedOutPanel: document.querySelector("#signed-out-panel"),
  signedInPanel: document.querySelector("#signed-in-panel"),
  setupMessage: document.querySelector("#setup-message"),
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
  appStatus: document.querySelector("#app-status"),
  result: document.querySelector("#result"),
};

const msalInstance = new PublicClientApplication(msalConfig);

function setStatus(message, state = "ready") {
  elements.appStatus.textContent = message;
  elements.result.dataset.state = state;
}

function describeError(error) {
  return error?.errorMessage || error?.message || "เกิดข้อผิดพลาดที่ไม่ทราบสาเหตุ";
}

function showSignedOut(message, isError = false) {
  elements.signedOutPanel.hidden = false;
  elements.signedInPanel.hidden = true;
  elements.signIn.hidden = false;
  elements.signOut.hidden = true;
  elements.authBadge.textContent = isError ? "เชื่อมต่อไม่สำเร็จ" : "ยังไม่ลงชื่อเข้าใช้";
  elements.authBadge.dataset.state = isError ? "error" : "setup";
  if (message) elements.setupMessage.textContent = message;
}

function showSignedIn(account) {
  elements.signedOutPanel.hidden = true;
  elements.signedInPanel.hidden = false;
  elements.signIn.hidden = true;
  elements.signOut.hidden = false;
  elements.authBadge.textContent = "ลงชื่อเข้าใช้แล้ว";
  elements.authBadge.dataset.state = "signed-in";
  elements.profileName.textContent = account.name || "บัญชีองค์กร";
  elements.profileAccount.textContent = account.username;
  elements.profileEmail.textContent = "กำลังโหลด...";
  elements.profileJob.textContent = "กำลังโหลด...";
}

async function getAccessToken(account, request = loginRequest) {
  try {
    return await msalInstance.acquireTokenSilent({ ...request, account });
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError)) throw error;
    return msalInstance.acquireTokenPopup({ ...request, account });
  }
}

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
    const response = await msalInstance.loginPopup(loginRequest);
    msalInstance.setActiveAccount(response.account);
    showSignedIn(response.account);
    await loadGraphProfile(response.account);
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
    await loadGraphProfile(account);
  } catch (error) {
    showSignedOut(`เริ่มระบบยืนยันตัวตนไม่สำเร็จ: ${describeError(error)}`, true);
    setStatus("เริ่มระบบไม่สำเร็จ", "error");
  }
}

start();