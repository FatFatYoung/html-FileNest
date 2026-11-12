// 构建脚本：bun build.ts
// 构建时自动把 src/index.html 注入到 JS 的 __EMBEDDED_UI_HTML__ 占位符，
// 确保 exe 单独运行时也能加载 UI（磁盘有 index.html 则优先用磁盘版）。
import { readFileSync, writeFileSync, unlinkSync } from 'fs';

const HTML = readFileSync('./src/index.html', 'utf8');
const JS_SRC = readFileSync('./src/html-FileNest.js', 'utf8');
const NOTICES = readFileSync('./THIRD-PARTY-NOTICES.txt', 'utf8');

// 转义模板字符串中的特殊字符
const escaped = HTML
  .replace(/\\/g, '\\\\')
  .replace(/`/g, '\\`')
  .replace(/\$\{/g, '\\${');

let JS_BUNDLED = JS_SRC.replace(
  '__EMBEDDED_UI_HTML__',
  '`' + escaped + '`'
);

// 注入第三方许可全文（JSON 字符串安全转义）
JS_BUNDLED = JS_BUNDLED.replace(
  '__EMBEDDED_NOTICES__',
  JSON.stringify(NOTICES)
);

// 写临时文件编译
const TMP = './src/_build_tmp.js';
writeFileSync(TMP, JS_BUNDLED, 'utf8');

await Bun.build({
  entrypoints: [TMP],
  compile: {
    outfile: "./dist/html-FileNest.exe",
    windows: {
      title: "html-FileNest",
      description: "html-FileNest - Local File Manager",
      publisher: "https://github.com/FatFatYoung",
      version: "1.4.2.2",
      copyright: "© 2026 FatFatYoung. All rights reserved.",
    },
  },
});

unlinkSync(TMP); // 清理临时文件

console.log("OK -> dist/html-FileNest.exe");
