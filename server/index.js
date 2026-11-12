import express from "express";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { createRemoteJWKSet, decodeJwt, jwtVerify } from "jose";
import { existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { getStorageMetricsUrl } from "../src/storageMetricsUrl.js";

const envFile = resolve(dirname(fileURLToPath(import.meta.url)), "../.env");
if (existsSync(envFile)) process.loadEnvFile(envFile);

const {
  ENTRA_TENANT_ID: tenantId,
  ENTRA_CLIENT_ID: clientId,
  ENTRA_CLIENT_SECRET: clientSecret,
  MAIL_FROM_ADDRESS: mailFromAddress,
} = process.env;
const oneDriveAdmins = new Set(
  (process.env.ONEDRIVE_ADMIN_UPNS || "")
    .split(",")
    .map((upn) => upn.trim().toLowerCase())
    .filter(Boolean),
);

const requiredConfig = {
  ENTRA_TENANT_ID: tenantId,
  ENTRA_CLIENT_ID: clientId,
  ENTRA_CLIENT_SECRET: clientSecret,
  MAIL_FROM_ADDRESS: mailFromAddress,
};
const missingConfig = Object.entries(requiredConfig)
  .filter(([, value]) => !value)
  .map(([key]) => key);

if (missingConfig.length) {
  throw new Error(`Missing required environment variables: ${missingConfig.join(", ")}`);
}

const authority = `https://login.microsoftonline.com/${tenantId}`;
const expectedAudience = [`api://${clientId}`, clientId];
const v1Issuer = `https://sts.windows.net/${tenantId}/`;
const v2Issuer = `${authority}/v2.0`;
const jwksByIssuer = new Map([
  [v1Issuer, createRemoteJWKSet(new URL(`${authority}/discovery/keys`))],
  [v2Issuer, createRemoteJWKSet(new URL(`${authority}/discovery/v2.0/keys`))],
]);
const confidentialClient = new ConfidentialClientApplication({
  auth: {
    clientId,
    authority,
    clientSecret,
  },
});
const app = express();

app.disable("x-powered-by");
app.use(express.json({ limit: "16kb" }));

async function authenticateUser(request, response, next) {
  const authorization = request.get("authorization") || "";
  const token = authorization.match(/^Bearer\s+(.+)$/i)?.[1];
  if (!token) {
    response.status(401).json({ error: "ต้องลงชื่อเข้าใช้ก่อนเรียก API" });
    return;
  }

  try {
    const tokenIssuer = decodeJwt(token).iss;
    const signingKeys = jwksByIssuer.get(tokenIssuer);
    if (!signingKeys) {
      response.status(401).json({ error: "เซสชันไม่ถูกต้องหรือหมดอายุ กรุณาลงชื่อเข้าใช้อีกครั้ง" });
      return;
    }

    const { payload } = await jwtVerify(token, signingKeys, {
      issuer: tokenIssuer,
      audience: expectedAudience,
    });
    const scopes = typeof payload.scp === "string" ? payload.scp.split(" ") : [];
    const account = String(payload.preferred_username || payload.upn || payload.email || "").toLowerCase();
    const callingClientId = payload.azp || payload.appid;

    if (payload.tid !== tenantId || callingClientId !== clientId || !scopes.includes("access_as_user")) {
      response.status(403).json({ error: "ไม่มีสิทธิ์เรียก API นี้" });
      return;
    }

    request.authenticatedUser = account;
    next();
  } catch (error) {
    console.warn(`Access token validation failed: ${error.code || error.name}`);
    response.status(401).json({ error: "เซสชันไม่ถูกต้องหรือหมดอายุ กรุณาลงชื่อเข้าใช้อีกครั้ง" });
  }
}

async function getGraphJson(accessToken, url) {
  const graphResponse = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!graphResponse.ok) {
    const error = new Error(`Microsoft Graph returned HTTP ${graphResponse.status}`);
    error.status = graphResponse.status;
    throw error;
  }
  return graphResponse.json();
}

app.get("/api/drive-metrics/access", authenticateUser, (request, response) => {
  response.json({ allowed: oneDriveAdmins.has(request.authenticatedUser) });
});

app.post("/api/drive-metrics", authenticateUser, async (request, response) => {
  if (!oneDriveAdmins.has(request.authenticatedUser)) {
    response.status(403).json({ error: "บัญชีนี้ไม่ได้รับอนุญาตให้ตรวจสอบ OneDrive ของผู้ใช้อื่น" });
    return;
  }

  const account = typeof request.body?.account === "string" ? request.body.account.trim() : "";
  if (account.length > 254 || !/^[^\s@]+@[^\s@]+$/.test(account)) {
    response.status(400).json({ error: "กรุณาระบุ UPN หรืออีเมลให้ถูกต้อง" });
    return;
  }

  try {
    const token = await confidentialClient.acquireTokenByClientCredential({
      scopes: ["https://graph.microsoft.com/.default"],
    });
    if (!token?.accessToken) throw new Error("Microsoft identity platform returned no access token");

    const escapedAccount = account.replaceAll("'", "''");
    const usersUrl = new URL("https://graph.microsoft.com/v1.0/users");
    usersUrl.searchParams.set(
      "$filter",
      `userPrincipalName eq '${escapedAccount}' or mail eq '${escapedAccount}'`,
    );
    usersUrl.searchParams.set("$select", "id,displayName,userPrincipalName,mail");
    const usersResult = await getGraphJson(token.accessToken, usersUrl);
    const users = Array.isArray(usersResult.value) ? usersResult.value : [];

    if (users.length === 0) {
      response.status(404).json({ error: "ไม่พบบัญชีที่ตรงกับ UPN หรืออีเมลนี้" });
      return;
    }
    if (users.length > 1) {
      response.status(409).json({ error: "พบหลายบัญชีที่ตรงกัน กรุณาระบุ UPN ของบัญชีให้ชัดเจน" });
      return;
    }

    const user = users[0];
    const driveUrl = new URL(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(user.id)}/drive`,
    );
    driveUrl.searchParams.set("$select", "quota,driveType,webUrl");
    const drive = await getGraphJson(token.accessToken, driveUrl);
    if (!drive.quota) {
      response.status(502).json({ error: "Microsoft Graph ไม่ส่งข้อมูลโควตาของ OneDrive กลับมา" });
      return;
    }
    const storageMetricsUrl = getStorageMetricsUrl(drive.webUrl);

    response.json({
      user: {
        displayName: user.displayName || null,
        userPrincipalName: user.userPrincipalName || null,
        mail: user.mail || null,
      },
      quota: drive.quota,
      storageMetricsUrl,
    });
  } catch (error) {
    const graphStatus = Number.isInteger(error.status) ? error.status : null;
    console.error(`OneDrive metric lookup failed: ${graphStatus || error.code || error.name || "unknown error"}`);
    if (graphStatus === 404) {
      response.status(404).json({ error: "ไม่พบ OneDrive ของบัญชีนี้ หรือบัญชียังไม่มี OneDrive" });
      return;
    }
    if (graphStatus === 403) {
      response.status(502).json({
        error: "Microsoft Graph ปฏิเสธการอ่านข้อมูล ตรวจสอบ Application permissions, Admin consent และการรองรับ endpoint ของ Tenant",
      });
      return;
    }
    response.status(502).json({ error: "อ่านข้อมูล OneDrive ไม่สำเร็จ ตรวจสอบการตั้งค่า Microsoft Graph และสิทธิ์ของแอป" });
  }
});

app.post("/api/send-mail", authenticateUser, async (request, response) => {
  const { to, subject, message } = request.body || {};
  if (typeof to !== "string" || to.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) {
    response.status(400).json({ error: "กรุณาระบุอีเมลผู้รับให้ถูกต้อง" });
    return;
  }
  if (typeof subject !== "string" || !subject.trim() || subject.length > 255) {
    response.status(400).json({ error: "กรุณาระบุเรื่องอีเมลไม่เกิน 255 ตัวอักษร" });
    return;
  }
  if (typeof message !== "string" || !message.trim() || message.length > 10000) {
    response.status(400).json({ error: "กรุณาระบุข้อความไม่เกิน 10,000 ตัวอักษร" });
    return;
  }

  try {
    const token = await confidentialClient.acquireTokenByClientCredential({
      scopes: ["https://graph.microsoft.com/.default"],
    });
    if (!token?.accessToken) throw new Error("Microsoft identity platform returned no access token");

    const graphResponse = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailFromAddress)}/sendMail`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token.accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          message: {
            body: { contentType: "Text", content: message },
            subject,
            toRecipients: [{ emailAddress: { address: to } }],
          },
          saveToSentItems: true,
        }),
      },
    );

    if (!graphResponse.ok) {
      console.error(`Graph sendMail failed with HTTP ${graphResponse.status} for ${request.authenticatedUser}`);
      response.status(502).json({ error: `Microsoft Graph ปฏิเสธการส่งอีเมล (HTTP ${graphResponse.status})` });
      return;
    }

    response.status(202).json({ sent: true });
  } catch (error) {
    console.error("Mail send failed:", error.message);
    response.status(502).json({ error: "ส่งอีเมลไม่สำเร็จ ตรวจสอบการตั้งค่า Node API และ Exchange RBAC" });
  }
});

app.use("/api", (request, response) => {
  response.status(404).json({ error: "ไม่พบ API endpoint ที่ร้องขอ" });
});

const port = Number(process.env.PORT || 3000);
app.listen(port, "0.0.0.0", () => {
  console.log(`Mail API listening on port ${port}`);
});