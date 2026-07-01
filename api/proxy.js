const { HttpsProxyAgent } = require('https-proxy-agent');

// Tắt xác thực chứng chỉ SSL để cho phép kết nối tới Gateway doanh nghiệp dùng chứng chỉ nội bộ/tự ký
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

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
    // Chỉ chuyển tiếp các headers cần thiết cho API từ client, tránh chuyển các header trình duyệt tự sinh
    // như Origin, Referer, Accept-Encoding hay Cookie có thể làm lỗi Gateway hoặc vi phạm chính sách bảo mật
    const allowedRequestHeaders = ['authorization', 'content-type', 'accept'];
    Object.keys(req.headers).forEach(key => {
      const lowerKey = key.toLowerCase();
      if (allowedRequestHeaders.includes(lowerKey)) {
        headers[lowerKey] = req.headers[key];
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
    });

    // Truyền tải status code và CHỈ chuyển tiếp header Content-Type từ server đích
    res.status(response.status);
    const contentType = response.headers.get('content-type');
    if (contentType) {
      res.setHeader('Content-Type', contentType);
    }

    const responseText = await response.text();
    res.send(responseText);
  } catch (error) {
    console.error('[Vercel Proxy Error]:', error);
    const causeStr = error.cause ? (error.cause.message || String(error.cause)) : 'none';
    res.status(500).send(`PROXY_ERROR: ${error.message} (Cause: ${causeStr}) - Stack: ${error.stack}`);
  }
};
