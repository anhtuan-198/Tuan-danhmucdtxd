import https from 'https';
import { JSDOM } from 'jsdom';

https.get('https://docs.google.com/spreadsheets/d/1B237SBdWeaQvc0GWH7hwcJI9ztiSxdBxXFbN4nBnxzU/htmlview/sheet?headers=true&gid=0', (res) => {
  let data = '';
  res.on('data', (chunk) => {
    data += chunk;
  });
  res.on('end', () => {
    const linkMap = new Map();
    const dom = new JSDOM(data);
    const doc = dom.window.document;
    const rows = doc.querySelectorAll('table tbody tr');
    
    rows.forEach(row => {
      const cells = row.querySelectorAll('td');
      if (cells.length > 1) {
        const textCell = cells[1];
        const textContent = textCell.textContent?.replace(/\s+/g, ' ').trim();
        
        const linkElement = row.querySelector('a');
        
        if (textContent && linkElement && linkElement.href) {
          let href = linkElement.href;
          if (href.includes('google.com/url?')) {
            try {
              const urlParams = new URLSearchParams(href.split('?')[1]);
              const q = urlParams.get('q');
              if (q) href = q;
            } catch (e) {}
          }
          linkMap.set(textContent, href);
        }
      }
    });
    console.log(`Extracted ${linkMap.size} links`);
    if (linkMap.size > 0) {
      console.log(Array.from(linkMap.entries()).slice(0, 3));
    }
  });
});
