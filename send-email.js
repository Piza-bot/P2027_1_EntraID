import { ConfidentialClientApplication } from "@azure/msal-node";

// ส่วนที่หนึ่ง: ข้อมูลอีเมล
const recipientEmail = "recipient@example.com";
const emailSubject = "ตัวอย่างส่งอีเมล";
const emailMessage = "ข้อความตัวอย่าง";

// ส่วนที่สอง: ค่าจาก Microsoft Entra ID และ Microsoft Graph
const tenantId = "YOUR_TENANT_ID";
const clientId = "YOUR_CLIENT_ID";
const clientSecret = "YOUR_CLIENT_SECRET";
const servicePrincipalObjectId = "YOUR_SERVICE_PRINCIPAL_OBJECT_ID"; // ใช้ตอนตั้งค่า Exchange RBAC
const senderEmail = "sender@example.com";
const authorityUri = `https://login.microsoftonline.com/${tenantId}`;
const graphScope = "https://graph.microsoft.com/.default";
const sendMailUri = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(senderEmail)}/sendMail`;

const msalClient = new ConfidentialClientApplication({
  auth: {
    clientId,
    authority: authorityUri,
    clientSecret,
  },
});

// ส่วนที่สาม: ขอ access token และส่งอีเมล
async function sendEmail() {
  const tokenResult = await msalClient.acquireTokenByClientCredential({
    scopes: [graphScope],
  });

  if (!tokenResult?.accessToken) {
    throw new Error("ไม่สามารถขอ Microsoft Graph access token ได้");
  }

  const response = await fetch(sendMailUri, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${tokenResult.accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      message: {
        subject: emailSubject,
        body: {
          contentType: "Text",
          content: emailMessage,
        },
        toRecipients: [
          {
            emailAddress: {
              address: recipientEmail,
            },
          },
        ],
      },
      saveToSentItems: true,
    }),
  });

  if (!response.ok) {
    throw new Error(`ส่งอีเมลไม่สำเร็จ (HTTP ${response.status}): ${await response.text()}`);
  }

  console.log("ส่งอีเมลเรียบร้อยแล้ว");
}

await sendEmail();
