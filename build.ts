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

// 注入 Ghostscript 最小集 ZIP（base64；vendor/gs-min.zip 缺失时给空串，gs 渲染退化为本机探测）
// vendor/gs-min.zip 不入库（.git/info/exclude），内容=gs/bin/{gswin64c.exe,gsdll64.dll}+gs/Resource/**
let gsB64 = '';
try {
  gsB64 = readFileSync('./vendor/gs-min.zip').toString('base64');
  console.log('GS payload: vendor/gs-min.zip ' + Math.round(gsB64.length / 1048576) + 'MB(base64)');
} catch (e) {
  console.log('GS payload: vendor/gs-min.zip 缺失，构建产物不带 Ghostscript');
}
JS_BUNDLED = JS_BUNDLED.replace(
  '__EMBEDDED_GS_ZIP_B64__',
  JSON.stringify(gsB64)
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
      version: "2.4.0.0",
      copyright: "© 2026 FatFatYoung. All rights reserved.",
    },
  },
});

unlinkSync(TMP); // 清理临时文件

console.log("OK -> dist/html-FileNest.exe");
