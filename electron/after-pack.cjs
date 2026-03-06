const fs = require("fs");
const path = require("path");

function ensureExecutable(filePath) {
  try {
    const stat = fs.statSync(filePath);
    if (!stat.isFile()) return;
    fs.chmodSync(filePath, 0o755);
  } catch {
    // Ignore missing paths or chmod failures for non-critical files.
  }
}

function chmodAllFiles(dirPath) {
  if (!fs.existsSync(dirPath)) return;
  const entries = fs.readdirSync(dirPath, { withFileTypes: true });
  for (const entry of entries) {
    const fullPath = path.join(dirPath, entry.name);
    if (entry.isDirectory()) {
      chmodAllFiles(fullPath);
    } else if (entry.isFile()) {
      ensureExecutable(fullPath);
    }
  }
}

module.exports = async function afterPack(context) {
  if (context.electronPlatformName !== "darwin") return;

  const productFilename = context.packager.appInfo.productFilename;
  const appBundlePath = path.join(context.appOutDir, `${productFilename}.app`);
  const macOSBinDir = path.join(appBundlePath, "Contents", "MacOS");
  const frameworksDir = path.join(appBundlePath, "Contents", "Frameworks");

  chmodAllFiles(macOSBinDir);

  if (fs.existsSync(frameworksDir)) {
    const frameworkEntries = fs.readdirSync(frameworksDir, { withFileTypes: true });
    for (const entry of frameworkEntries) {
      if (!entry.isDirectory() || !entry.name.endsWith(".app")) continue;
      const helperBinDir = path.join(frameworksDir, entry.name, "Contents", "MacOS");
      chmodAllFiles(helperBinDir);
    }
  }
};
