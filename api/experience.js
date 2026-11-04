module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ success: false, message: 'POST 요청만 허용됩니다.' });
  }
  const scriptUrl = process.env.APPS_SCRIPT_URL;
  const secret = process.env.APP_SECRET;
  if (!scriptUrl || !secret) {
    return res.status(500).json({ success: false, message: 'Vercel 환경변수 APPS_SCRIPT_URL 또는 APP_SECRET이 설정되지 않았습니다.' });
  }
  try {
    const payload = Object.assign({}, req.body || {}, { secret });
    const response = await fetch(scriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
      redirect: 'follow'
    });
    const result = await response.json();
    return res.status(200).json(result);
  } catch (error) {
    return res.status(502).json({ success: false, message: error.message || 'Apps Script 연결에 실패했습니다.' });
  }
};