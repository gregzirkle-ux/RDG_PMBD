import { parseCookies, refreshAccessToken } from '../auth/utils.js';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', 'https://rdg-pmbd.vercel.app');
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
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

  // Refresh token if needed
  if (!accessToken && refreshToken) {
    try {
      const tokenData = await refreshAccessToken(refreshToken);
      accessToken = tokenData.access_token;

      const cookieOptions = ['HttpOnly', 'Secure', 'SameSite=Lax', 'Path=/'];
      res.setHeader('Set-Cookie', [
        [`outlook_access_token=${tokenData.access_token}`, `Max-Age=${tokenData.expires_in}`, ...cookieOptions].join('; '),
        [`outlook_refresh_token=${tokenData.refresh_token}`, `Max-Age=${90 * 24 * 60 * 60}`, ...cookieOptions].join('; '),
        [`outlook_connected=true`, `Max-Age=${90 * 24 * 60 * 60}`, 'Secure', 'SameSite=Lax', 'Path=/'].join('; '),
      ]);
    } catch (err) {
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

  const { subject, start, end, location, body, isAllDay } = req.body;

  if (!subject || !start || !end) {
    return res.status(400).json({
      error: 'Missing required fields',
      message: 'subject, start, and end are required.',
    });
  }

  const eventPayload = {
    subject,
    start: {
      dateTime: start,
      timeZone: 'Eastern Standard Time',
    },
    end: {
      dateTime: end,
      timeZone: 'Eastern Standard Time',
    },
    isAllDay: isAllDay || false,
  };

  if (location) {
    eventPayload.location = { displayName: location };
  }
  if (body) {
    eventPayload.body = { contentType: 'text', content: body };
  }

  try {
    const graphResponse = await fetch(
      'https://graph.microsoft.com/v1.0/me/events',
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(eventPayload),
      }
    );

    const data = await graphResponse.json();

    if (data.error) {
      return res.status(graphResponse.status).json({
        error: data.error.code,
        message: data.error.message,
      });
    }

    return res.status(201).json({
      success: true,
      event: data,
    });
  } catch (err) {
    console.error('Create event error:', err);
    return res.status(500).json({
      error: 'Server error',
      message: 'Failed to create calendar event.',
    });
  }
}
