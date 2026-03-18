import https from 'https';
import Papa from 'papaparse';

https.get('https://docs.google.com/spreadsheets/d/1B237SBdWeaQvc0GWH7hwcJI9ztiSxdBxXFbN4nBnxzU/export?format=csv&gid=0', (res) => {
  let data = '';
  res.on('data', (chunk) => {
    data += chunk;
  });
  res.on('end', () => {
    Papa.parse(data, {
      complete: (results) => {
        const allRows = results.data;
        const col5 = allRows.map(r => r[5]).filter(Boolean);
        console.log('Column 5 values:', col5.slice(0, 20));
      }
    });
  });
});
