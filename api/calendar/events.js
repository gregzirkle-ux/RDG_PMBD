import { parseCookies, refreshAccessToken } from '../auth/utils.js';

export default async function handler(req, res) {
  // CORS headers for frontend fetch calls
  res.setHeader('Access-Control-Allow-Origin', 'https://rdg-pmbd.vercel.app');
  res.setHeader('Access-Control-Allow-Credentials', 'true');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  const cookies = parseCookies(req);
  let accessToken = cookies.outlook_access_token;
  const refreshToken = cookies.outlook_refresh_token;

  if (!accessToken && !refreshToken) {
    return res.status(401).json({
      error: 'Not authenticated',
      message: 'Please connect your Outlook account first.',
    });
  }

  // If no access token but we have a refresh token, try to refresh
  if (!accessToken && refreshToken) {
    try {
      const tokenData = await refreshAccessToken(refreshToken);
      accessToken = tokenData.access_token;

      // Update cookies with new tokens
      const cookieOptions = ['HttpOnly', 'Secure', 'SameSite=Lax', 'Path=/'];

      const newAccessCookie = [
        `outlook_access_token=${tokenData.access_token}`,
        `Max-Age=${tokenData.expires_in}`,
        ...cookieOptions,
      ].join('; ');

      const newRefreshCookie = [
        `outlook_refresh_token=${tokenData.refresh_token}`,
        `Max-Age=${90 * 24 * 60 * 60}`,
        ...cookieOptions,
      ].join('; ');

      const statusCookie = [
        'outlook_connected=true',
        `Max-Age=${90 * 24 * 60 * 60}`,
        'Secure',
        'SameSite=Lax',
        'Path=/',
      ].join('; ');

      res.setHeader('Set-Cookie', [newAccessCookie, newRefreshCookie, statusCookie]);
    } catch (err) {
      console.error('Token refresh failed:', err);
      // Clear cookies on refresh failure
      res.setHeader('Set-Cookie', [
        'outlook_access_token=; Max-Age=0; Path=/',
        'outlook_refresh_token=; Max-Age=0; Path=/',
        'outlook_connected=; Max-Age=0; Path=/',
      ]);
      return res.status(401).json({
        error: 'Session expired',
        message: 'Please reconnect your Outlook account.',
      });
    }
  }

  // Build the Graph API query
  const now = new Date();
  const startDate = req.query.start || now.toISOString();

  // Default to 30 days ahead if no end date specified
  const defaultEnd = new Date(now);
  defaultEnd.setDate(defaultEnd.getDate() + 30);
  const endDate = req.query.end || defaultEnd.toISOString();

  const top = req.query.top || 50;

  const graphUrl =
    `https://graph.microsoft.com/v1.0/me/calendarView` +
    `?startDateTime=${encodeURIComponent(startDate)}` +
    `&endDateTime=${encodeURIComponent(endDate)}` +
    `&$top=${top}` +
    `&$orderby=start/dateTime` +
    `&$select=subject,start,end,location,organizer,isAllDay,bodyPreview,webLink,categories,showAs`;

  try {
    let graphResponse = await fetch(graphUrl, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
    });

    // If 401, try refreshing the token once
    if (graphResponse.status === 401 && refreshToken) {
      try {
        const tokenData = await refreshAccessToken(refreshToken);
        accessToken = tokenData.access_token;

        const cookieOptions = ['HttpOnly', 'Secure', 'SameSite=Lax', 'Path=/'];
        const newAccessCookie = [
          `outlook_access_token=${tokenData.access_token}`,
          `Max-Age=${tokenData.expires_in}`,
          ...cookieOptions,
        ].join('; ');
        const newRefreshCookie = [
          `outlook_refresh_token=${tokenData.refresh_token}`,
          `Max-Age=${90 * 24 * 60 * 60}`,
          ...cookieOptions,
        ].join('; ');
        const statusCookie = [
          'outlook_connected=true',
          `Max-Age=${90 * 24 * 60 * 60}`,
          'Secure',
          'SameSite=Lax',
          'Path=/',
        ].join('; ');
        res.setHeader('Set-Cookie', [newAccessCookie, newRefreshCookie, statusCookie]);

        graphResponse = await fetch(graphUrl, {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
          },
        });
      } catch (refreshErr) {
        console.error('Token refresh on 401 failed:', refreshErr);
        res.setHeader('Set-Cookie', [
          'outlook_access_token=; Max-Age=0; Path=/',
          'outlook_refresh_token=; Max-Age=0; Path=/',
          'outlook_connected=; Max-Age=0; Path=/',
        ]);
        return res.status(401).json({
          error: 'Session expired',
          message: 'Please reconnect your Outlook account.',
        });
      }
    }

    const data = await graphResponse.json();

    if (data.error) {
      console.error('Graph API error:', data.error);
      return res.status(graphResponse.status).json({
        error: data.error.code,
        message: data.error.message,
      });
    }

    // Return the events
    return res.status(200).json({
      events: data.value || [],
      count: (data.value || []).length,
    });
  } catch (err) {
    console.error('Calendar fetch error:', err);
    return res.status(500).json({
      error: 'Server error',
      message: 'Failed to fetch calendar events.',
    });
  }
}
