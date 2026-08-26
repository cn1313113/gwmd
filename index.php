<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>GWMD编辑器</title>
<link rel="stylesheet" href="style.css?v=<?= filemtime(__DIR__ . '/style.css') ?>">
</head>
<body>
<header class="toolbar">
  <div class="brand">
    <span class="logo">文</span>
    <span class="title">GWMD编辑器</span>
    <span class="sub">纯文本 → 标准公文 docx</span>
  </div>
  <div class="actions">
    <button id="btn-example" title="载入示例文档">示例</button>
    <button id="btn-syntax" title="查看语法说明">语法</button>
    <button id="btn-import-gwmd" class="primary-outline" title="导入 .gwmd 文件">导入 .gwmd</button>
    <button id="btn-save-gwmd" class="primary-outline" title="保存为 .gwmd 文本文件">保存 .gwmd</button>
    <button id="btn-export-docx" class="primary" title="导出为标准公文格式 Word 文档">导出 .docx</button>
  </div>
</header>

<main class="workspace">
  <section class="editor-pane">
    <div class="pane-head">
      <span>编辑区 <small>（GWMD 语法）</small></span>
      <div class="font-controls">
        <button id="font-dec" title="减小字号">A－</button>
        <button id="font-inc" title="增大字号">A＋</button>
      </div>
      <span class="pane-meta" id="edit-meta"></span>
    </div>
    <textarea id="editor" spellcheck="false" placeholder="@title 关于XXX的通知&#10;&#10;@to 各单位：&#10;&#10;正文内容……"></textarea>
  </section>
  <section class="preview-pane">
    <div class="pane-head">
      <span>预览 <small>（A4 公文效果）</small></span>
      <span class="pane-meta" id="preview-meta"></span>
    </div>
    <div class="preview-scroll">
      <div id="preview" class="page"></div>
    </div>
  </section>
</main>

<footer class="statusbar">
  <span id="status-text">就绪</span>
  <span id="status-right"></span>
</footer>

<div id="syntax-modal" class="modal hidden">
  <div class="modal-box">
    <div class="modal-head">
      <span>GWMD 语法速查</span>
      <button class="modal-close" id="syntax-close">×</button>
    </div>
    <div class="modal-body" id="syntax-body"></div>
  </div>
</div>

<div id="toast" class="toast hidden"></div>

<script src="app.js?v=<?= filemtime(__DIR__ . '/app.js') ?>"></script>
</body>
</html>
