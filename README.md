// /api/book.js — Vercel Serverless: tạo event Outlook + Teams, Outlook tự gửi thư mời
export default async function handler(req, res) {
  try {
    if (req.method !== 'POST') return res.status(405).json({ ok:false, error:'Method Not Allowed' });
    if (req.headers['x-api-key'] !== process.env.API_KEY) return res.status(401).json({ ok:false, error:'Unauthorized' });

    const { title, startTime, endTime, description, location, guests, organizerUpn } = req.body || {};
    if (!title || !startTime || !endTime) {
      return res.status(400).json({ ok:false, error:'Missing title/startTime/endTime' });
    }

    const attendees = String(guests || '')
      .split(',')
      .map(s => s.trim())
      .filter(Boolean)
      .map(address => ({ emailAddress: { address }, type: 'required' }));

    // Organizer mặc định: tungtt1@vitadairy.com.vn
    const organizer = organizerUpn || process.env.ORGANIZER_UPN;
    const token = await getAppToken();

    const payload = {
      subject: title,
      body: { contentType: 'HTML', content: description || '' },
      start: { dateTime: toLocal(startTime), timeZone: 'SE Asia Standard Time' },
      end:   { dateTime: toLocal(endTime),   timeZone: 'SE Asia Standard Time' },
      attendees,
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness'
    };
    if (location) payload.location = { displayName: location };

    const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(organizer)}/events`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        Prefer: 'outlook.timezone="SE Asia Standard Time"'
      },
      body: JSON.stringify(payload)
    });

    const text = await r.text();
    if (!r.ok) {
      console.error('GraphError', r.status, text);
      return res.status(r.status).json({ ok:false, error:'GraphError', details:text });
    }

    const data = JSON.parse(text);
    return res.status(201).json({
      ok: true,
      message: 'Event created (Outlook sent invites).',
      eventId: data.id,
      webLink: data.webLink || null,
      meetUrl: data?.onlineMeeting?.joinUrl || null
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok:false, error:String(e?.message || e) });
  }
}

async function getAppToken() {
  const url = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    grant_type: 'client_credentials',
    scope: 'https://graph.microsoft.com/.default'
  });
  const r = await fetch(url, { method: 'POST', body });
  if (!r.ok) throw new Error(`Token error: ${r.status} ${await r.text()}`);
  const j = await r.json();
  return j.access_token;
}

function toLocal(input) {
  return String(input).replace(/([+-]\d{2}:\d{2}|Z)$/,'');
}

