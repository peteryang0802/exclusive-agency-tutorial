/* ============================================================
   独家代理业务预算 · 简化计算引擎
   ------------------------------------------------------------
   单位：人民币万元（除非另行注明）
   口径说明：这是「整年简化版」，用于建立业务直觉与方向感。
   与 Excel 模板（SS/FW 分季、跨年、到岸附加逐笔、付款节点峰值）
   对照，默认参数下 2027 年：
   · 净利润、毛利、已售成本三项与模板完全对齐；
   · 库存 561 万（模板 805 万）、净现金流 −713 万（模板 −956 万），
     均低估约 243 万——整年一次反推漏掉了 SS/FW 季节滚动；
   · 资金需求 813 万（模板 1550 万），模板另含付款节点峰值。
   精确数字请以 Excel 模板为准，不要把本引擎结果当作申报口径。
   ============================================================ */

(function (global) {
  'use strict';

  /* —— 默认输入（对齐 Excel「核心输入」2027 年默认值）—— */
  const DEFAULTS = {
    revenue: 2000,         // 公司年度含税收入目标（万元）
    // 五类渠道比例（合计应 = 100%）
    chOnline: 0.80,        // 直营网店
    chOffline: 0.00,       // 直营线下单品牌店（含 POP）
    chDistSingle: 0.00,    // 分销线下单品牌店
    chDistMulti: 0.15,     // 分销线下多品牌店
    chBuyer: 0.05,         // 分销线下买手店
    sellThrough: 0.50,     // DTC 当季新货售罄率
    oldPct: 0.00,          // 老货计划贡献 DTC 收入比例
    dtcMargin: 0.65,       // DTC 货值毛利率
    distMargin: 0.30,      // 分销货值毛利率
    landingRate: 0.1327,   // 到岸附加成本率（运费/关税/清关，加到货款与已售成本）
    promoRate: 0.15,       // 品牌推广费率（× 含税收入）
    adRate: 0.20,          // 线上投流费率（× DTC 含税）
    storeOpexRate: 0.09,   // 线上店铺运营费率（× DTC 含税：佣金/物流/客服）
    backofficeRate: 0.015, // 后台其他费率（× 不含税收入）
    staffCost: 280,        // 人员公司成本（万元，整十估算值）
    distService: 6,        // 分销服务成本（万元）
    vat: 0.13,             // 增值税率（货物销售一般纳税人基本税率）
    cashBuffer: 100,       // 最低现金安全余额（万元）
    depreciation: 0,       // 装修/设备折旧（万元，默认 0）
    interest: 0,           // 付息融资利息（万元，默认 0）
    taxRate: 0.25,         // 企业所得税基本税率
    currency: 'USD',       // 采购报价币种（美元采购）
    fxRate: 7.2,           // 定价基准汇率（元/美元）
  };

  /* —— 核心计算 —— */
  function compute(input) {
    const p = Object.assign({}, DEFAULTS, input || {});

    // 1. 渠道占比：直营（网店+线下） vs 分销（三类）
    const dtcPct = p.chOnline + p.chOffline;
    const distPct = p.chDistSingle + p.chDistMulti + p.chBuyer;

    // 2. 收入
    const revTaxIncl = p.revenue;                       // 含税收入
    const revExcl = revTaxIncl / (1 + p.vat);           // 不含税收入
    const dtcRevExcl = revExcl * dtcPct;                // DTC 不含税
    const distRevExcl = revExcl * distPct;              // 分销不含税

    // 3. 新货 / 老货（DTC）
    const dtcNewExcl = dtcRevExcl * (1 - p.oldPct);     // DTC 新货不含税销售
    const dtcOldExcl = dtcRevExcl * p.oldPct;           // DTC 老货不含税销售

    // 4. 采购（货款，不含到岸附加的「基准」）
    //    关键公式：采购 = 新货不含税销售 × (1−毛利率) ÷ 售罄率
    const dtcBuyBase = dtcNewExcl * (1 - p.dtcMargin) / p.sellThrough;
    const distBuyBase = distRevExcl * (1 - p.distMargin);   // 分销现款提货、不压库存
    const buyBase = dtcBuyBase + distBuyBase;

    // 5. 已售成本（基准，不含到岸附加）
    const dtcCogsBase = dtcNewExcl * (1 - p.dtcMargin);
    const dtcOldCogsBase = dtcOldExcl;                  // 老货按成本结转（简化）
    const distCogsBase = distRevExcl * (1 - p.distMargin);
    const cogsBase = dtcCogsBase + dtcOldCogsBase + distCogsBase;

    // 6. 到岸附加（运费/关税/清关，加到采购与已售成本上）
    const landing = 1 + p.landingRate;
    const buy = buyBase * landing;          // 支付给品牌方的货款（含到岸附加）
    const cogs = cogsBase * landing;        // 已售货品总成本（含到岸附加）

    // 7. 库存（年末 DTC 库存 ≈ 采购基准 − 已售基准，再加到岸附加）
    const inventory = (dtcBuyBase - dtcCogsBase) * landing;

    // 8. 毛利
    const grossProfit = revExcl - cogs;
    const grossMargin = revExcl > 0 ? grossProfit / revExcl : 0;

    // 9. 费用
    const dtcTaxIncl = dtcRevExcl * (1 + p.vat);        // DTC 含税（投流/店铺运营的基数）
    const promo = revTaxIncl * p.promoRate;             // 品牌推广
    const ad = dtcTaxIncl * p.adRate;                   // 线上投流
    const storeOpex = dtcTaxIncl * p.storeOpexRate;     // 线上店铺运营
    const backoffice = revExcl * p.backofficeRate;      // 后台其他
    const staff = p.staffCost;
    const distService = p.distService;
    const opex = promo + ad + storeOpex + backoffice + staff + distService;

    // 10. 利润（区分 EBITDA / 税前利润 / 净利润）
    const ebitda = grossProfit - opex;        // 息税折旧摊销前利润
    const depreciation = p.depreciation;      // 装修/设备折旧
    const interest = p.interest;              // 付息融资利息
    const pretaxProfit = ebitda - depreciation - interest; // 税前利润
    const tax = pretaxProfit > 0 ? pretaxProfit * p.taxRate : 0; // 所得税（正利润才计提）
    const netProfit = pretaxProfit - tax;     // 净利润

    // 11. 现金流（整年简化口径）
    //     净现金流 ≈ 净利润 − 年末库存占用（采购已付、货未卖完）
    const netCash = netProfit - inventory;
    //     资金需求（简化，年末口径）：补足净流出 + 最低安全余额
    //     真实模板还要覆盖「提货前 25 天一次性付款」的峰值缺口，因此更大。
    const fundNeed = Math.max(0, -netCash) + p.cashBuffer;

    return {
      p, dtcPct, distPct,
      revTaxIncl, revExcl, dtcRevExcl, distRevExcl,
      dtcNewExcl, dtcOldExcl,
      buy, cogs, inventory,
      grossProfit, grossMargin,
      promo, ad, storeOpex, backoffice, staff, distService, opex,
      ebitda, pretaxProfit, tax, netProfit,
      netCash, fundNeed,
    };
  }

  /* —— 数字格式化 —— */
  function fmtWan(x, d) {
    if (x === null || x === undefined || isNaN(x)) return '—';
    d = (d === undefined) ? 1 : d;
    const s = Number(x).toFixed(d);
    // 千分位
    const neg = s.startsWith('-');
    const abs = neg ? s.slice(1) : s;
    const parts = abs.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return (neg ? '-' : '') + parts.join('.');
  }
  function fmtPct(x, d) {
    if (x === null || x === undefined || isNaN(x)) return '—';
    d = (d === undefined) ? 1 : d;
    return (x * 100).toFixed(d) + '%';
  }
  function fmtInt(x) {
    if (x === null || x === undefined || isNaN(x)) return '—';
    return Math.round(x).toLocaleString('zh-CN');
  }
  // 万元 → 元（用于显示具体货币）
  function fmtYuan(x) {
    if (x === null || x === undefined || isNaN(x)) return '—';
    return (x * 10000).toLocaleString('zh-CN', { maximumFractionDigits: 0 });
  }

  /* —— 风险检测：调数据时实时触发 —— */
  function assessRisk(r) {
    var risks = [];
    var p = r.p;

    if (p.currency === 'USD') {
      risks.push({
        level: 'warn',
        title: '汇率风险',
        text: '采购按美元计价，基准汇率 ' + p.fxRate + '。人民币一旦贬到 7.5，货款就会多 ' + fmtPct(7.5 / p.fxRate - 1) + '。要不要锁汇、留出汇兑缓冲？'
      });
    }

    if (p.dtcMargin >= 0.65) {
      risks.push({
        level: 'warn',
        title: '平行进口 / 海淘风险',
        text: 'DTC 毛利率 ' + fmtPct(p.dtcMargin) + ' 偏高。若采购成本过高、国内零售价明显高于品牌海外同款，消费者会海淘、代购绕开官方渠道。请核对海外定价、控制价差。'
      });
    }

    if (p.sellThrough < 0.5) {
      risks.push({
        level: 'warn',
        title: '库存积压',
        text: '售罄率 ' + fmtPct(p.sellThrough) + ' 偏低，意味着要多备货，年末约 ' + fmtWan(r.inventory) + ' 万压在仓库，占现金、占仓储。'
      });
    }

    var chTotal = p.chOnline + p.chOffline + p.chDistSingle + p.chDistMulti + p.chBuyer;
    if (Math.abs(chTotal - 1) > 0.005) {
      risks.push({
        level: 'note',
        title: '渠道结构错配',
        text: '五类渠道合计 ' + fmtPct(chTotal) + '，不是 100%。蛋糕切多了或切少了。'
      });
    }

    if (r.netCash < 0) {
      risks.push({
        level: 'warn',
        title: '现金流风险',
        text: '净现金流 ' + fmtWan(r.netCash) + ' 万，全年资金需求约 ' + fmtWan(r.fundNeed) + ' 万。8 月 FW 全款前必须备足，否则会断供。'
      });
    }

    if (r.netProfit < 0) {
      risks.push({
        level: 'note',
        title: '持续亏损',
        text: '当前净利润 ' + fmtWan(r.netProfit) + ' 万。要能回答：这是战略投入期，还是模式没跑通？'
      });
    }

    var mktTotal = r.promo + r.ad;
    var mktRate = r.revTaxIncl > 0 ? mktTotal / r.revTaxIncl : 0;
    if (mktRate > 0.30) {
      risks.push({
        level: 'warn',
        title: '费用吃毛利',
        text: '推广 + 投流合计 ' + fmtWan(mktTotal, 0) + ' 万，占含税收入 ' + fmtPct(mktRate) + '。两笔钱都在吞毛利，别只盯一个费率。'
      });
    }

    risks.push({
      level: 'note',
      title: '关税需核准',
      text: '到岸附加率 ' + fmtPct(p.landingRate) + ' 是预算估值，不是法定税率。真实关税按 HS 编码、原产地、享惠凭证核定。'
    });

    if (r.revTaxIncl > 0 && r.staff / r.revTaxIncl > 0.2) {
      risks.push({
        level: 'note',
        title: '人力超前',
        text: '人员成本占收入 ' + fmtPct(r.staff / r.revTaxIncl) + '，偏高。人要跟业务量走，不是提前配满。'
      });
    }

    return risks;
  }

  global.CalcModel = {
    DEFAULTS: DEFAULTS,
    compute: compute,
    assessRisk: assessRisk,
    fmtWan: fmtWan,
    fmtPct: fmtPct,
    fmtInt: fmtInt,
    fmtYuan: fmtYuan,
  };
})(window);
