const http = require('http')
const fs = require('fs')

const server = http.createServer((req, res) => {
  res.writeHead(200, { 'content-type': 'text/html' })
  fs.createReadStream('index.html').pipe(res)
})

const portNo = process.env.PORT || 3000;
server.listen(portNo)
console.log(`Server started successfully on port: ${portNo}`);