const fs = require("node:fs");
const path = require("node:path");

const repoRoot = path.resolve(__dirname, "..");
const podfilePath = path.join(repoRoot, "ios", "App", "Podfile");

if (!fs.existsSync(podfilePath)) {
  console.error(`Podfile not found: ${podfilePath}`);
  process.exit(1);
}

const original = "pod 'CapacitorFilesystem', :path => '../../node_modules/@capacitor/filesystem'";
const replacement = "pod 'CapacitorFilesystem', :path => '../local-pods/CapacitorFilesystem'";

let content = fs.readFileSync(podfilePath, "utf8");

if (content.includes(original)) {
  content = content.replace(original, replacement);
  fs.writeFileSync(podfilePath, content);
  console.log("Patched ios/App/Podfile to use the local CapacitorFilesystem pod.");
} else if (content.includes(replacement)) {
  console.log("ios/App/Podfile already uses the local CapacitorFilesystem pod.");
} else {
  console.warn("CapacitorFilesystem pod entry was not found in ios/App/Podfile.");
}
