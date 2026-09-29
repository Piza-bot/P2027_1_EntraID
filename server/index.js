import express from "express";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { createRemoteJWKSet, decodeJwt, jwtVerify } from "jose";

const {
  ENTRA_TENANT_ID: tenantId,
  ENTRA_CLIENT_ID: clientId,
  ENTRA_CLIENT_SECRET: clientSecret,
  MAIL_FROM_ADDRESS: mailFromAddress,
} = process.env;

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
    response.status(401).json({ error: "ต้องลงชื่อเข้าใช้ก่อนส่งอีเมล" });
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
      response.status(403).json({ error: "ไม่มีสิทธิ์เรียก API ส่งอีเมล" });
      return;
    }

    request.authenticatedUser = account;
    next();
  } catch (error) {
    console.warn(`Access token validation failed: ${error.code || error.name}`);
    response.status(401).json({ error: "เซสชันไม่ถูกต้องหรือหมดอายุ กรุณาลงชื่อเข้าใช้อีกครั้ง" });
  }
}

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

const port = Number(process.env.PORT || 3000);
app.listen(port, "0.0.0.0", () => {
  console.log(`Mail API listening on port ${port}`);
});