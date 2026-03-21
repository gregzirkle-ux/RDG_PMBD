// Parses cookies from the request header
export function parseCookies(req) {
  const cookies = {};
  const cookieHeader = req.headers.cookie || '';
  cookieHeader.split(';').forEach((cookie) => {
    const [name, ...rest] = cookie.trim().split('=');
    if (name) {
      cookies[name.trim()] = rest.join('=').trim();
    }
  });
  return cookies;
}

// Refreshes the access token using the refresh token
export async function refreshAccessToken(refreshToken) {
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;
  const tenantId = process.env.AZURE_TENANT_ID;

  const response = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
        scope: 'openid profile offline_access Calendars.ReadWrite',
      }),
    }
  );

  const data = await response.json();

  if (data.error) {
    throw new Error(data.error_description || data.error);
  }

  return data;
}
