// Run this script with Node.js to generate simple PNG icon placeholders.
// For production, replace these with proper icons.
// Usage: node src/assets/generate-icons.js

const fs = require("fs");
const path = require("path");

const sizes = [16, 32, 64, 80, 128];

// Simple 1x1 blue pixel PNG as a minimal placeholder
// In production, replace with actual designed icons
const svgTemplate = (size) => `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">
  <rect width="${size}" height="${size}" rx="${size * 0.15}" fill="#2563eb"/>
  <text x="50%" y="55%" dominant-baseline="middle" text-anchor="middle" fill="white" font-family="Arial, sans-serif" font-size="${size * 0.45}" font-weight="bold">EA</text>
</svg>`;

for (const size of sizes) {
  const svg = svgTemplate(size);
  fs.writeFileSync(path.join(__dirname, `icon-${size}.svg`), svg);
  console.log(`Created icon-${size}.svg`);
}

console.log("\nNote: The manifest references .png files. For development, update the manifest");
console.log("to use .svg or convert these to .png using a tool like Inkscape or sharp.");
