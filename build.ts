// 构建脚本：bun build.ts
import { copyFileSync } from 'fs';

await Bun.build({
  entrypoints: ["./src/html-FileNest.js"],
  compile: {
    outfile: "./dist/html-FileNest.exe",
    windows: {
      title: "html-FileNest",
      publisher: "https://github.com/FatFatYoung",
      version: "1.0.0.0",
      copyright: "© 2026 FatFatYoung. All rights reserved.",
    },
  },
});

// 把 index.html 复制到 dist/（exe 首次部署时需要它在旁边）
copyFileSync("./src/index.html", "./dist/index.html");

console.log("OK -> dist/html-FileNest.exe + dist/index.html");
