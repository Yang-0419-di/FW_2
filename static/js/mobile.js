/* =====================================================================
   mobile.js — 全站手機版輔助（搭配 static/css/mobile.css）
   1. 手機選單按鈕：展開／收合 #siteNav
   2. 欄位多的表格：標記 .m-table（不換行），外面沒有可左右捲動的容器時自動包一層 .m-scroll
   只在螢幕寬度 ≤ 768px 時處理表格；桌機完全不動。
   ===================================================================== */
(function () {
    'use strict';
    var MQ = window.matchMedia ? window.matchMedia('(max-width: 768px)') : { matches: false };
    var WIDE_COLS = 5;          // 欄位數 ≥ 5 視為寬表格
    var WRAP_CHARS = 18;        // 儲存格文字超過此長度允許換行（避免單欄過寬）
    // 這些表格已經有自己的手機寬度設定，或是外掛元件（FullCalendar）產生的
    var SKIP = '.fc table, [data-m-skip], [data-m-skip] table';
    var OWN = '.log-table, .mfp-table, .personal-main-table';   // 自己有手機寬度設定：只加提示

    // ---------- 1. 手機選單 ----------
    function initNav() {
        var btn = document.querySelector('.m-nav-toggle');
        var nav = document.getElementById('siteNav');
        if (!btn || !nav) return;
        btn.addEventListener('click', function () {
            var open = nav.classList.toggle('open');
            btn.setAttribute('aria-expanded', open ? 'true' : 'false');
            var arrow = btn.querySelector('.m-nav-arrow');
            if (arrow) arrow.textContent = open ? '▲' : '▼';
        });
    }

    // ---------- 2. 寬表格 ----------
    function colCount(table) {
        var max = 0;
        for (var i = 0; i < table.rows.length && i < 30; i++) {
            var n = 0, cells = table.rows[i].cells;
            for (var j = 0; j < cells.length; j++) n += cells[j].colSpan || 1;
            if (n > max) max = n;
        }
        return max;
    }

    function hasScrollParent(el) {
        for (var p = el.parentElement; p && p !== document.body; p = p.parentElement) {
            var ox = window.getComputedStyle(p).overflowX;
            if (ox === 'auto' || ox === 'scroll') return true;
        }
        return false;
    }

    function markWrapCells(table) {
        var tds = table.querySelectorAll('td');
        for (var i = 0; i < tds.length; i++) {
            var td = tds[i];
            if (td.querySelector('input, select, textarea, button')) continue;
            if ((td.textContent || '').trim().length > WRAP_CHARS) td.classList.add('m-wrap');
        }
    }

    function processTable(table) {
        if (table.dataset.mDone) return;
        if (table.matches(SKIP)) return;
        table.dataset.mDone = '1';
        if (table.matches(OWN)) { addHint(table); return; }
        if (window.getComputedStyle(table).tableLayout === 'fixed') return;   // 固定欄寬的表格本來就會塞進畫面
        var cols = colCount(table);
        if (cols < WIDE_COLS) return;
        table.classList.add('m-table');
        markWrapCells(table);
        if (!hasScrollParent(table)) {
            var wrap = document.createElement('div');
            wrap.className = 'm-scroll';
            table.parentNode.insertBefore(wrap, table);
            wrap.appendChild(table);
        }
        addHint(table);
    }

    function addHint(table) {
        // 「可左右滑動」提示：先插入，是否顯示由 refreshHints 依實際寬度決定
        var box = table.closest('.m-scroll, .table-container, .table-container-NO, .table-responsive') || table;
        if (!(box.previousElementSibling && box.previousElementSibling.classList.contains('m-scroll-hint'))) {
            var hint = document.createElement('div');
            hint.className = 'm-scroll-hint';
            hint.hidden = true;
            hint.textContent = '↔ 表格可左右滑動';
            box.parentNode.insertBefore(hint, box);
        }
    }

    // 只有表格真的超出畫面、而且區塊是展開的，才顯示提示
    function refreshHints() {
        var hints = document.querySelectorAll('.m-scroll-hint');
        for (var i = 0; i < hints.length; i++) {
            var box = hints[i].nextElementSibling;
            var show = MQ.matches && box && box.offsetParent !== null && box.scrollWidth > box.clientWidth + 8;   // 只差幾個像素不算
            hints[i].hidden = !show;
        }
    }
    var hintTimer = null;
    function refreshHintsSoon() { clearTimeout(hintTimer); hintTimer = setTimeout(refreshHints, 80); }

    function processTables(root) {
        if (!MQ.matches) return;
        var tables = (root || document).querySelectorAll('table');
        for (var i = 0; i < tables.length; i++) processTable(tables[i]);
        refreshHints();
    }

    // 動態新增的資料列（例如首頁硬碟檢測以 AJAX 載入）也要標記長文字
    function watchDynamic() {
        if (!window.MutationObserver) return;
        var mo = new MutationObserver(function (list) {
            if (!MQ.matches) return;
            for (var i = 0; i < list.length; i++) {
                var t = list[i].target;
                var table = t.closest ? t.closest('table.m-table') : null;
                if (table) markWrapCells(table);
            }
            refreshHintsSoon();
        });
        var bodies = document.querySelectorAll('tbody');
        for (var i = 0; i < bodies.length; i++) mo.observe(bodies[i], { childList: true });
    }

    function init() {
        initNav();
        processTables();
        watchDynamic();
        // 收合／展開區塊、旋轉螢幕後重新判斷提示
        document.addEventListener('click', refreshHintsSoon, true);
        window.addEventListener('resize', refreshHintsSoon);
        // 由桌機寬度轉成手機寬度（例如旋轉螢幕）時再處理一次
        if (MQ.addEventListener) MQ.addEventListener('change', function () { processTables(); });
        else if (MQ.addListener) MQ.addListener(function () { processTables(); });
    }

    // 讓其他頁面（例如之後才產生的表格）可以手動呼叫
    window.MobileUI = { processTables: processTables, refreshHints: refreshHints };

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
})();
