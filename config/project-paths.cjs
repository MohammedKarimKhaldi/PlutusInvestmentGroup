const path = require("path");

const rootDir = path.resolve(__dirname, "..");
const sourceAppDir = path.join(rootDir, "app");
const buildDir = path.join(rootDir, "build");
const webBuildDir = path.join(buildDir, "web");

const webSubdirs = {
  pages: "pages",
  scripts: "scripts",
  styles: "styles",
  data: "data",
};

const dataFiles = {
  config: "config.json",
  deals: "deals.json",
  tasks: "tasks.json",
  sharedTasks: "sharedrive-tasks.json",
  teamStorePath: "team-store-path.json",
};

module.exports = {
  rootDir,
  sourceAppDir,
  buildDir,
  webBuildDir,
  webSubdirs,
  dataFiles,
};
