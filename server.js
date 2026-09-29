import express from "express";
import session from "express-session";
import { google } from "googleapis";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { ImapFlow } from "imapflow";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use(
  session({
    secret: process.env.SESSION_SECRET || "change-this-secret",
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      secure: process.env.NODE_ENV === "production",
      sameSite: "lax"
    }
  })
);

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

app.get("/api/status", (req, res) => {
  res.json({
    loggedIn: !!req.session.account,
    account: req.session.account
      ? {
          provider: req.session.account.provider,
          email: req.session.account.email
        }
      : null
  });
});

app.post("/api/logout", (req, res) => {
  req.session.destroy(() => {
    res.json({ ok: true });
  });
});

/* =========================
   GMAIL
========================= */

app.get("/auth/gmail", (req, res) => {
  if (
    !process.env.GOOGLE_CLIENT_ID ||
    !process.env.GOOGLE_CLIENT_SECRET
  ) {
    return res.status(500).send("Gmail OAuth ist nicht konfiguriert.");
  }

  const redirect =
    process.env.APP_URL
      ? `${process.env.APP_URL}/auth/gmail/callback`
      : `${req.protocol}://${req.get("host")}/auth/gmail/callback`;

  const oauth = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    redirect
  );

  const url = oauth.generateAuthUrl({
    access_type: "offline",
    prompt: "consent",
    scope: [
      "openid",
      "email",
      "profile",
      "https://www.googleapis.com/auth/gmail.readonly",
      "https://www.googleapis.com/auth/gmail.modify"
    ]
  });

  res.redirect(url);
});

app.get("/auth/gmail/callback", async (req, res) => {
  try {
    const redirect =
      process.env.APP_URL
        ? `${process.env.APP_URL}/auth/gmail/callback`
        : `${req.protocol}://${req.get("host")}/auth/gmail/callback`;

    const oauth = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET,
      redirect
    );

    const { tokens } = await oauth.getToken(req.query.code);

    oauth.setCredentials(tokens);

    const gmail = google.gmail({
      version: "v1",
      auth: oauth
    });

    const profile = await gmail.users.getProfile({
      userId: "me"
    });

    req.session.account = {
      provider: "Gmail",
      email: profile.data.emailAddress,
      tokens
    };

    res.redirect("/");
  } catch (error) {
    console.error(error);
    res.status(500).send("Gmail-Anmeldung fehlgeschlagen.");
  }
});

async function gmailClient(account) {
  const oauth = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET
  );

  oauth.setCredentials(account.tokens);

  return google.gmail({
    version: "v1",
    auth: oauth
  });
}

app.get("/api/mail", async (req, res) => {
  if (!req.session.account) {
    return res.status(401).json({
      error: "Nicht angemeldet"
    });
  }

  try {
    if (req.session.account.provider === "Gmail") {
      const gmail = await gmailClient(req.session.account);

      const result = await gmail.users.messages.list({
        userId: "me",
        maxResults: 30,
        q: "in:inbox"
      });

      const messages = result.data.messages || [];

      const output = [];

      for (const item of messages) {
        const mail = await gmail.users.messages.get({
          userId: "me",
          id: item.id,
          format: "metadata",
          metadataHeaders: [
            "From",
            "To",
            "Subject",
            "Date"
          ]
        });

        const headers =
          mail.data.payload?.headers || [];

        const get = name =>
          headers.find(
            h => h.name.toLowerCase() === name.toLowerCase()
          )?.value || "";

        output.push({
          id: item.id,
          from: get("From"),
          to: get("To"),
          subject: get("Subject"),
          date: get("Date"),
          unread:
            mail.data.labelIds?.includes("UNREAD") || false
        });
      }

      return res.json(output);
    }

    return res.json([]);
  } catch (error) {
    console.error(error);
    res.status(500).json({
      error: "E-Mails konnten nicht geladen werden."
    });
  }
});

/* =========================
   GMAIL DELETE
========================= */

app.post("/api/mail/delete", async (req, res) => {
  if (!req.session.account) {
    return res.status(401).json({
      error: "Nicht angemeldet"
    });
  }

  if (req.session.account.provider !== "Gmail") {
    return res.status(400).json({
      error: "Dieser Anbieter wird noch nicht unterstützt."
    });
  }

  try {
    const gmail = await gmailClient(req.session.account);

    const ids = Array.isArray(req.body.ids)
      ? req.body.ids
      : [];

    for (const id of ids) {
      await gmail.users.messages.trash({
        userId: "me",
        id
      });
    }

    res.json({
      ok: true,
      deleted: ids.length
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({
      error: "E-Mails konnten nicht gelöscht werden."
    });
  }
});

/* =========================
   OUTLOOK
========================= */

function msalClient() {
  return new ConfidentialClientApplication({
    auth: {
      clientId: process.env.MS_CLIENT_ID,
      clientSecret: process.env.MS_CLIENT_SECRET,
      authority:
        `https://login.microsoftonline.com/` +
        `${process.env.MS_TENANT_ID || "common"}`
    }
  });
}

app.get("/auth/outlook", async (req, res) => {
  if (
    !process.env.MS_CLIENT_ID ||
    !process.env.MS_CLIENT_SECRET
  ) {
    return res.status(500).send(
      "Outlook OAuth ist nicht konfiguriert."
    );
  }

  const redirect =
    process.env.APP_URL
      ? `${process.env.APP_URL}/auth/outlook/callback`
      : `${req.protocol}://${req.get("host")}/auth/outlook/callback`;

  const url = await msalClient().getAuthCodeUrl({
    scopes: [
      "openid",
      "profile",
      "email",
      "offline_access",
      "User.Read",
      "Mail.Read",
      "Mail.ReadWrite"
    ],
    redirectUri: redirect
  });

  res.redirect(url);
});

app.get("/auth/outlook/callback", async (req, res) => {
  try {
    const redirect =
      process.env.APP_URL
        ? `${process.env.APP_URL}/auth/outlook/callback`
        : `${req.protocol}://${req.get("host")}/auth/outlook/callback`;

    const result =
      await msalClient().acquireTokenByCode({
        code: req.query.code,
        scopes: [
          "openid",
          "profile",
          "email",
          "offline_access",
          "User.Read",
          "Mail.Read",
          "Mail.ReadWrite"
        ],
        redirectUri: redirect
      });

    req.session.account = {
      provider: "Outlook",
      email:
        result.account?.username || "Outlook",
      accessToken: result.accessToken
    };

    res.redirect("/");
  } catch (error) {
    console.error(error);
    res.status(500).send(
      "Outlook-Anmeldung fehlgeschlagen."
    );
  }
});

async function graph(pathname, token, options = {}) {
  const response = await fetch(
    `https://graph.microsoft.com/v1.0${pathname}`,
    {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...(options.headers || {})
      }
    }
  );

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph ${response.status}`
    );
  }

  if (response.status === 204) {
    return null;
  }

  return response.json();
}

app.get("/api/outlook/mail", async (req, res) => {
  if (
    !req.session.account ||
    req.session.account.provider !== "Outlook"
  ) {
    return res.status(401).json({
      error: "Nicht angemeldet"
    });
  }

  try {
    const data = await graph(
      "/me/mailFolders/inbox/messages?$top=30&$select=id,subject,from,toRecipients,receivedDateTime,isRead",
      req.session.account.accessToken
    );

    const mails = (data.value || []).map(m => ({
      id: m.id,
      from: m.from?.emailAddress?.address || "",
      subject: m.subject || "",
      date: m.receivedDateTime || "",
      unread: !m.isRead
    }));

    res.json(mails);
  } catch (error) {
    console.error(error);
    res.status(500).json({
      error: "Outlook-Mails konnten nicht geladen werden."
    });
  }
});

app.post("/api/outlook/delete", async (req, res) => {
  if (
    !req.session.account ||
    req.session.account.provider !== "Outlook"
  ) {
    return res.status(401).json({
      error: "Nicht angemeldet"
    });
  }

  try {
    const ids = Array.isArray(req.body.ids)
      ? req.body.ids
      : [];

    for (const id of ids) {
      await graph(
        `/me/messages/${encodeURIComponent(id)}`,
        req.session.account.accessToken,
        {
          method: "DELETE"
        }
      );
    }

    res.json({
      ok: true,
      deleted: ids.length
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({
      error: "Outlook-Mails konnten nicht gelöscht werden."
    });
  }
});

/* =========================
   PROTON MAIL
========================= */

app.post("/api/proton/connect", async (req, res) => {
  const {
    host,
    port,
    user
  } = req.body;

  if (!host || !user || !process.env.PROTON_IMAP_PASSWORD) {
    return res.status(400).json({
      error:
        "Proton-IMAP-Konfiguration fehlt."
    });
  }

  try {
    const client = new ImapFlow({
      host,
      port: Number(port) || 993,
      secure: true,
      auth: {
        user,
        pass: process.env.PROTON_IMAP_PASSWORD
      }
    });

    await client.connect();

    await client.logout();

    req.session.account = {
      provider: "Proton",
      email: user,
      proton: {
        host,
        port: Number(port) || 993,
        user
      }
    };

    res.json({
      ok: true,
      email: user
    });
  } catch (error) {
    console.error(error);

    res.status(500).json({
      error:
        "Proton-Verbindung konnte nicht hergestellt werden."
    });
  }
});

async function protonClient(account) {
  return new ImapFlow({
    host: account.proton.host,
    port: account.proton.port,
    secure: true,
    auth: {
      user: account.proton.user,
      pass: process.env.PROTON_IMAP_PASSWORD
    }
  });
}

app.get("/api/proton/mail", async (req, res) => {
  if (
    !req.session.account ||
    req.session.account.provider !== "Proton"
  ) {
    return res.status(401).json({
      error: "Nicht angemeldet"
    });
  }

  let client;

  try {
    client = await protonClient(
      req.session.account
    );

    await client.connect();

    const lock =
      await client.getMailboxLock("INBOX");

    try {
      const mails = [];

      for await (
        const message of client.fetch(
          "1:*",
          {
            envelope: true,
            flags: true,
            uid: true
          },
          {
            uid: true
          }
        )
      ) {
        mails.push({
          id: String(message.uid),
          from:
            message.envelope?.from?.[0]?.address ||
            "",
          subject:
            message.envelope?.subject ||
            "",
          date:
            message.envelope?.date ||
            "",
          unread:
            !message.flags?.has("\\Seen")
        });

        if (mails.length >= 30) break;
      }

      res.json(mails);
    } finally {
      lock.release();
    }

    await client.logout();
  } catch (error) {
    console.error(error);
    if (client) {
      try {
        await client.logout();
      } catch {}
    }

    res.status(500).json({
      error: "Proton-Mails konnten nicht geladen werden."
    });
  }
});

app.post("/api/proton/delete", async (req, res) => {
  if (
    !req.session.account ||
    req.session.account.provider !== "Proton"
  ) {
    return res.status(401).json({
      error: "Nicht angemeldet"
    });
  }

  const ids = Array.isArray(req.body.ids)
    ? req.body.ids
    : [];

  let client;

  try {
    client = await protonClient(
      req.session.account
    );

    await client.connect();

    const lock =
      await client.getMailboxLock("INBOX");

    try {
      await client.messageMove(
        ids.map(Number),
        "Trash",
        {
          uid: true
        }
      );
    } finally {
      lock.release();
    }

    await client.logout();

    res.json({
      ok: true,
      deleted: ids.length
    });
  } catch (error) {
    console.error(error);

    res.status(500).json({
      error: "Proton-Mails konnten nicht gelöscht werden."
    });
  }
});

/* =========================
   DEMO AGENT
========================= */

app.post("/api/agent/analyze", (req, res) => {
  const mails = Array.isArray(req.body.mails)
    ? req.body.mails
    : [];

  const importantWords = [
    "schule",
    "schul",
    "eltern",
    "familie",
    "rechnung",
    "bestellung",
    "termin",
    "wichtig",
    "sicherheit",
    "konto",
    "vertrag",
    "bewerbung",
    "arzt"
  ];

  const important = [];
  const unimportant = [];

  for (const mail of mails) {
    const text =
      `${mail.subject} ${mail.from}`.toLowerCase();

    const isImportant =
      importantWords.some(word =>
        text.includes(word)
      );

    if (isImportant) {
      important.push(mail);
    } else {
      unimportant.push(mail);
    }
  }

  res.json({
    important,
    unimportant,
    total: mails.length
  });
});

app.listen(PORT, () => {
  console.log(`NEXUS AI läuft auf Port ${PORT}`);
});
