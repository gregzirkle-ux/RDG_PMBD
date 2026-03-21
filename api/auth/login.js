export default function handler(req, res) {
  const clientId = process.env.AZURE_CLIENT_ID;
  const tenantId = process.env.AZURE_TENANT_ID;
  const redirectUri = `https://rdg-pmbd.vercel.app/api/auth/callback`;
  const scope = 'openid profile offline_access Calendars.ReadWrite';

  const authUrl =
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize` +
    `?client_id=${clientId}` +
    `&response_type=code` +
    `&redirect_uri=${encodeURIComponent(redirectUri)}` +
    `&scope=${encodeURIComponent(scope)}` +
    `&response_mode=query`;

  res.redirect(302, authUrl);
}
