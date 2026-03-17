const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

const OUT = "D365SpeedUp-v1.1.zip";

// Files and folders to include in the store zip
const include = [
  "manifest.json",
  "speedup.html",
  "sidepanel-empty.html",
  "sidepanel-empty.js",
  "background.js",
  "style.css",
  "config.json",
  "assets",
  "libs",
  "dist",
];

// Delete old zip if exists
if (fs.existsSync(OUT)) fs.unlinkSync(OUT);

const items = include.join(" ");
execSync(`powershell Compress-Archive -Path ${include.map(i => `"./${i}"`).join(",")} -DestinationPath "${OUT}"`, { stdio: "inherit" });

console.log(`\nCreated: ${OUT}`);
