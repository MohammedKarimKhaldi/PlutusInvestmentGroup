const fs = require("node:fs");
const path = require("node:path");

const repoRoot = path.resolve(__dirname, "..");
const packageFile = path.join(repoRoot, "ios", "App", "CapApp-SPM", "Package.swift");

if (!fs.existsSync(packageFile)) {
  console.log("No iOS SPM package file found, skipping path patch.");
  process.exit(0);
}

const replacements = [
  {
    from: '../../../node_modules/@capacitor/filesystem',
    to: '../../capacitor-plugin-packages/filesystem',
  },
  {
    from: '../../../node_modules/@capacitor/share',
    to: '../../capacitor-plugin-packages/share',
  },
];

let content = fs.readFileSync(packageFile, "utf8");
let changed = false;

for (const replacement of replacements) {
  if (content.includes(replacement.from)) {
    content = content.split(replacement.from).join(replacement.to);
    changed = true;
  }
}

if (changed) {
  fs.writeFileSync(packageFile, content);
  console.log("Patched ios/App/CapApp-SPM/Package.swift to use vendored Capacitor plugins.");
} else {
  console.log("ios/App/CapApp-SPM/Package.swift already points at vendored Capacitor plugins.");
}
