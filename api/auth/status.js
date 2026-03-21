import { parseCookies } from './utils.js';

export default function handler(req, res) {
  const cookies = parseCookies(req);
  const isConnected = !!(cookies.outlook_access_token || cookies.outlook_refresh_token);

  return res.status(200).json({
    connected: isConnected,
  });
}
