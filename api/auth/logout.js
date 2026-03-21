export default function handler(req, res) {
  res.setHeader('Set-Cookie', [
    'outlook_access_token=; Max-Age=0; Path=/; HttpOnly; Secure; SameSite=Lax',
    'outlook_refresh_token=; Max-Age=0; Path=/; HttpOnly; Secure; SameSite=Lax',
    'outlook_connected=; Max-Age=0; Path=/; Secure; SameSite=Lax',
  ]);

  res.redirect('/?logout=true');
}
