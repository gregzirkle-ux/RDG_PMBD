export default async function handler(req, res) {
  const { code, error, error_description } = req.query;

  if (error) {
    return res.redirect(
      `/?auth_error=${encodeURIComponent(error_description || error)}`
    );
  }

  if (!code) {
    return res.redirect('/?auth_error=No%20authorization%20code%20received');
  }

  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;
  const tenantId = process.env.AZURE_TENANT_ID;
  const redirectUri = `https://rdg-pmbd.vercel.app/api/auth/callback`;

  try {
    const tokenResponse = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          code: code,
          redirect_uri: redirectUri,
          grant_type: 'authorization_code',
          scope: 'openid profile offline_access Calendars.ReadWrite',
        }),
      }
    );

    const tokenData = await tokenResponse.json();

    if (tokenData.error) {
      return res.redirect(
        `/?auth_error=${encodeURIComponent(tokenData.error_description || tokenData.error)}`
      );
    }

    const { access_token, refresh_token, expires_in } = tokenData;

    // Store tokens in HTTP-only cookies (secure, not accessible by JS)
    const cookieOptions = [
      'HttpOnly',
      'Secure',
      'SameSite=Lax',
      'Path=/',
    ];

    // Access token cookie - expires when the token does
    const accessCookie = [
      `outlook_access_token=${access_token}`,
      `Max-Age=${expires_in}`,
      ...cookieOptions,
    ].join('; ');

    // Refresh token cookie - long lived (90 days)
    const refreshCookie = [
      `outlook_refresh_token=${refresh_token}`,
      `Max-Age=${90 * 24 * 60 * 60}`,
      ...cookieOptions,
    ].join('; ');

    // Flag cookie - readable by frontend JS to know auth status
    const statusCookie = [
      'outlook_connected=true',
      `Max-Age=${90 * 24 * 60 * 60}`,
      'Secure',
      'SameSite=Lax',
      'Path=/',
    ].join('; ');

    res.setHeader('Set-Cookie', [accessCookie, refreshCookie, statusCookie]);
    res.redirect('/?auth_success=true');
  } catch (err) {
    console.error('Token exchange error:', err);
    res.redirect('/?auth_error=Token%20exchange%20failed');
  }
}
