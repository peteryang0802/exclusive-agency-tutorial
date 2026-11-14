/* 公共小工具：章节导航高亮 + 风险提醒渲染 */

(function (global) {
  'use strict';

  // 根据当前文件名高亮顶栏 / 首页目录
  function highlightNav() {
    var path = location.pathname;
    var m = path.match(/(\d{1,2})-/);           // 章节页 1-xxx.html ~ 10-xxx.html
    if (m) {
      var n = parseInt(m[1], 10);
      document.querySelectorAll('.toc a[data-ch]').forEach(function (a) {
        if (parseInt(a.getAttribute('data-ch'), 10) === n) {
          a.style.borderColor = 'var(--brand)';
          a.style.boxShadow = 'var(--shadow)';
        }
      });
    }
  }

  // 渲染实时风险提醒区
  function renderRisks(el, risks) {
    if (!el) return;
    if (!risks || risks.length === 0) {
      el.innerHTML = '<div class="risk-empty">当前组合未触发风险提醒——但关税、合同约束仍要签约前单独核准。</div>';
      return;
    }
    var html = '<div class="risks-title">风险提醒 · 调数据时别忽略</div>';
    risks.forEach(function (rk) {
      html += '<div class="risk-item ' + rk.level + '"><span class="t">' + rk.title + '</span>' + rk.text + '</div>';
    });
    el.innerHTML = html;
  }

  // 简单 SVG 漏斗 / 堆叠条不在此实现，各页按需内联

  global.Site = {
    highlightNav: highlightNav,
    renderRisks: renderRisks,
  };

  document.addEventListener('DOMContentLoaded', highlightNav);
})(window);
