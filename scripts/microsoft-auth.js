import http from "node:http";
import { URLSearchParams } from "node:url";

const PORT = 53682;
const REDIRECT_URI = `http://localhost:${PORT}/callback`;
const AUTH_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
const TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
const SCOPES = "offline_access Files.ReadWrite User.Read Mail.Send";

const clientId = process.env.MS_CLIENT_ID;
const clientSecret = process.env.MS_CLIENT_SECRET;

if (!clientId || !clientSecret) {
  console.error("Set MS_CLIENT_ID and MS_CLIENT_SECRET before running this script.");
  console.error("PowerShell example:");
  console.error("$env:MS_CLIENT_ID='your-client-id'; $env:MS_CLIENT_SECRET='your-client-secret'; node scripts/microsoft-auth.js");
  process.exit(1);
}

const authParams = new URLSearchParams({
  client_id: clientId,
  response_type: "code",
  redirect_uri: REDIRECT_URI,
  response_mode: "query",
  scope: SCOPES,
});

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, REDIRECT_URI);

  if (url.pathname !== "/callback") {
    res.writeHead(404);
    res.end("Not found");
    return;
  }

  const code = url.searchParams.get("code");
  const error = url.searchParams.get("error_description") || url.searchParams.get("error");

  if (!code) {
    res.writeHead(400, { "Content-Type": "text/plain" });
    res.end(error || "No authorization code received.");
    return;
  }

  try {
    const tokenBody = new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      code,
      redirect_uri: REDIRECT_URI,
      grant_type: "authorization_code",
      scope: SCOPES,
    });

    const tokenResponse = await fetch(TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: tokenBody,
    });
    const tokenPayload = await tokenResponse.json();

    if (!tokenResponse.ok) {
      throw new Error(tokenPayload.error_description || tokenPayload.error || "Token exchange failed.");
    }

    res.writeHead(200, { "Content-Type": "text/plain" });
    res.end("Success. You can close this tab and copy the refresh token from the terminal.");

    console.log("");
    console.log("Add this Vercel environment variable:");
    console.log("");
    console.log(`MS_REFRESH_TOKEN=${tokenPayload.refresh_token}`);
    console.log("");

    server.close();
  } catch (exchangeError) {
    res.writeHead(500, { "Content-Type": "text/plain" });
    res.end(exchangeError.message);
  }
});

server.listen(PORT, () => {
  console.log("");
  console.log("1. Add this redirect URI to your Microsoft app registration:");
  console.log(REDIRECT_URI);
  console.log("");
  console.log("2. Open this URL in your browser and sign in to the Microsoft account that owns the Excel file:");
  console.log("");
  console.log(`${AUTH_URL}?${authParams.toString()}`);
  console.log("");
});
