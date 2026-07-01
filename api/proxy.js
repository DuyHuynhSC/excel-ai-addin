const { HttpsProxyAgent } = require('https-proxy-agent');

module.exports = async (req, res) => {
  // Thiết lập CORS headers để cho phép các yêu cầu từ Excel Add-in (WebView2)
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version, Authorization, X-Target-Url');

  // Xử lý tiền kiểm tra CORS (Preflight request)
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  const targetUrl = req.headers['x-target-url'];
  if (!targetUrl) {
    res.status(400).send('Lỗi: Thiếu header X-Target-Url');
    return;
  }

  // Phân tích đường dẫn tương đối (ví dụ /chat/completions) từ URL yêu cầu
  const urlPath = req.url.replace(/^\/api\/proxy/, '').replace(/^\/api-proxy/, '');
  const destination = targetUrl.replace(/\/$/, '') + urlPath;

  try {
    const headers = {};
    // Sao chép các headers từ client, lọc bỏ host và x-target-url
    Object.keys(req.headers).forEach(key => {
      if (key !== 'host' && key !== 'x-target-url') {
        headers[key] = req.headers[key];
      }
    });

    const bodyContent = req.method !== 'GET' && req.method !== 'HEAD' 
      ? (typeof req.body === 'object' ? JSON.stringify(req.body) : req.body) 
      : undefined;

    // Cấu hình proxy doanh nghiệp nếu có
    const proxyUrl = process.env.HTTPS_PROXY || process.env.HTTP_PROXY;
    const agent = proxyUrl ? new HttpsProxyAgent(proxyUrl) : undefined;

    console.log(`[Vercel Proxy] Forwarding: ${req.method} ${req.url} -> ${destination}`);

    // Sử dụng fetch toàn cục của Node.js v18+
    const response = await fetch(destination, {
      method: req.method,
      headers: headers,
      body: bodyContent,
      agent: agent
    } as any);

    // Truyền tải status code và headers trả về từ server đích
    res.status(response.status);
    response.headers.forEach((value, key) => {
      const lowerKey = key.toLowerCase();
      if (lowerKey !== 'access-control-allow-origin' && lowerKey !== 'content-encoding' && lowerKey !== 'transfer-encoding') {
        res.setHeader(key, value);
      }
    });

    const responseText = await response.text();
    res.send(responseText);
  } catch (error) {
    console.error('[Vercel Proxy Error]:', error);
    res.status(500).send(`PROXY_ERROR: ${error.message}`);
  }
};
