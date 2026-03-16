const fs = require("fs");
const path = require("path");

const {
  sourceAppDir,
  webBuildDir,
} = require("../config/project-paths.cjs");

function copyRecursive(src, dest) {
  const stat = fs.statSync(src);

  if (stat.isDirectory()) {
    fs.mkdirSync(dest, { recursive: true });
    for (const entry of fs.readdirSync(src)) {
      if (entry === ".DS_Store") continue;
      copyRecursive(path.join(src, entry), path.join(dest, entry));
    }
    return;
  }

  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.copyFileSync(src, dest);
}

fs.rmSync(webBuildDir, { recursive: true, force: true });
fs.mkdirSync(webBuildDir, { recursive: true });
copyRecursive(sourceAppDir, webBuildDir);

console.log("Prepared web assets in ./build/web");
