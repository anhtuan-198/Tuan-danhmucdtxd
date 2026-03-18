const https = require('https');

function get(url) {
  return new Promise((resolve, reject) => {
    https.get(url, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        get(res.headers.location).then(resolve).catch(reject);
      } else {
        const data = [];
        res.on('data', chunk => data.push(chunk));
        res.on('end', () => {
          const buffer = Buffer.concat(data);
          const base64 = buffer.toString('base64');
          const mimeType = res.headers['content-type'] || 'image/png';
          resolve(`data:${mimeType};base64,${base64}`);
        });
      }
    }).on('error', reject);
  });
}

const fileId = '1GAoEg3co_ypn1fpAz95tv8A3Zgg4GCfF';
const url = `https://drive.google.com/uc?export=view&id=${fileId}`;

get(url).then(console.log).catch(console.error);
