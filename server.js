const express = require("express");
const https = require("https");
const fs = require("fs");
const path = require("path");

const app = express();

// Load SSL certificates
const options = {
    pfx: fs.readFileSync("localhost.pfx"),
    passphrase: "yourpassword"
};

// Serve static files from the public directory
app.use(express.static(path.join(__dirname, "public")));

// Start the HTTPS server
const PORT = 3000;
https.createServer(options, app).listen(PORT, () => {
    console.log(`Office Add-in Server running at https://127.0.0.1:${PORT}`);
});
