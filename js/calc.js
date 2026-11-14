/* ============================================================
   独家代理业务预算 · 教学计算引擎（重构版）
   ------------------------------------------------------------
   单位：人民币万元（除特别注明）。
   这是一个「建立财务直觉」的简化模型，不是申报口径。
   核心机制（对照 Excel 模板逐格核对）：
   1. 货品按「完整商品季」算采购——服装一年两季 SS/FW，
      FW 当年只确认 75%、25% 跨年，但货款今年就要全额付；
   2. 已售成本按「自然年」确认，库存走会计恒等式：
      期末库存 = 期初库存 + 本期采购(含到岸) − 本期已售；
   3. 人员按业务量反推，再按 4:1 自动拆基础/管理岗；
   4. 现金流分 1—2 月、3—8 月、9—12 月三个时段，
      还原「提货前 25 天全额付款」的节点缺口。
   默认参数下，货品采购、已售、库存与 Excel 第一年完全对齐。
   简化会低估人力成本（逐岗薪酬更高），文中已明确标注。
   ============================================================ */

(function (global) {
  'use strict';

  /* ============================================================
     一、常量（写死的业务规则，来自 Excel 公用设置）
     ============================================================ */
  const C = {
    vat: 0.13,              // 增值税率（货物销售一般纳税人）
    ssShare: 0.35,          // SS 占 DTC 收入比例（FW 占 65%）
    fwCurrent: 0.75,        // FW 直营当年确认比例（25% 跨年到次年 1—2 月）
    dtcMargin: 0.65,        // DTC 新货货值毛利率
    distMargin: 0.30,       // 分销货值毛利率
    landingRate: 0.1327,    // 综合到岸附加率（运保+关税+清关）
    basicAnnual: 18,        // 基础岗全年公司成本（税前 1 万/月 ×1.5×12）
    managerAnnual: 36,      // 管理岗全年公司成本（税前 2 万/月 ×1.5×12）
    managerRatio: 0.20,     // 管理岗占比（4 基础 : 1 管理）
    blendedAnnual: 21.6,    // 混合人均年成本（18×0.8 + 36×0.2）
    onlineOpexRate: 0.09,   // 线上运营费率（除投流：佣金/履约/其他，×线上含税）
    adRate: 0.20,           // 线上投流费率（5 年固定，×线上含税）
    storeAnnualOpex: 100,   // 单店年线下运营（含租金物业，×等价店，不含商场扣点）
    storeCapex: 100,        // 新店装修设备（每店，一次性）
    storeDeposit: 18,       // 新店租赁押金（每店，一次性）
    newStorePartial: 0.586, // 新店 6 月开业，当年营业比例
    backofficeRate: 0.015,  // 后台其他费率（×不含税收入）
    distServiceRate: 0.015, // 分销服务费率（×分销含税）
    cashBuffer: 100,        // 最低现金安全余额
    taxRate: 0.25,          // 企业所得税率（正利润才计）
    // 分销单客户年提货额（向上取整算客户数）
    clientTicket: { single: 500, multi: 150, buyer: 80 },
    storeRevenue: 400,      // 单店年承接能力（用于由线下收入反推店数）
  };

  /* ============================================================
     二、五年默认场景（对齐 Excel 核心输入，2027—2031）
     每个对象是一年的「可调输入」，各渠道占比合计 = 100%
     ============================================================ */
  const DEFAULT_SCENARIO = [
    { year: 2027, revenue: 2000, chOnline: 0.80, chOffline: 0.00, chDistSingle: 0.00, chDistMulti: 0.15, chBuyer: 0.05, sellThrough: 0.50, oldPct: 0.00, stores: 0, douyinShare: 0.000, commission: 0.01 },
    { year: 2028, revenue: 3500, chOnline: 0.70, chOffline: 0.08, chDistSingle: 0.00, chDistMulti: 0.17, chBuyer: 0.05, sellThrough: 0.55, oldPct: 0.15, stores: 1, douyinShare: 0.214, commission: 0.01 },
    { year: 2029, revenue: 6000, chOnline: 0.70, chOffline: 0.10, chDistSingle: 0.00, chDistMulti: 0.15, chBuyer: 0.05, sellThrough: 0.60, oldPct: 0.25, stores: 2, douyinShare: 0.214, commission: 0.01 },
    { year: 2030, revenue: 9000, chOnline: 0.64, chOffline: 0.15, chDistSingle: 0.10, chDistMulti: 0.08, chBuyer: 0.03, sellThrough: 0.60, oldPct: 0.25, stores: 4, douyinShare: 0.234, commission: 0.01 },
    { year: 2031, revenue: 12000, chOnline: 0.64, chOffline: 0.10, chDistSingle: 0.15, chDistMulti: 0.08, chBuyer: 0.03, sellThrough: 0.60, oldPct: 0.30, stores: 4, douyinShare: 0.234, commission: 0.01 },
  ];

  // 单年默认输入（= 场景第一年，兼容单页调用）
  const DEFAULTS = Object.assign({}, DEFAULT_SCENARIO[0]);

  /* ============================================================
     三、费率：自动阶梯 vs 手动全稳定期
     ============================================================ */
  // 品牌推广费率：自动（前 3 年 15%、第 4 年起 10%）；稳态恒 10%
  function promoRateOf(idx, stable) {
    if (stable) return 0.10;
    return idx < 3 ? 0.15 : 0.10;
  }
  // 投流费率：自动恒 20%；稳态 15%
  function adRateOf(stable) {
    return stable ? 0.15 : C.adRate;
  }

  /* ============================================================
     四、渠道：店铺、客户、平台数量（平均假设）
     ============================================================ */
  // 后台编制（含分销拓展），按收入规模分档，对齐 Excel
  function backofficeHeadcount(revenue) {
    if (revenue <= 2000) return 4.5;
    if (revenue <= 3500) return 5.5;
    if (revenue <= 6000) return 8;
    if (revenue <= 9000) return 10;
    return 12.5;
  }

  // 由线下收入反推年末店数（向上取整，单店约 400 万）
  function storesFromRevenue(offlineTaxIncl) {
    if (offlineTaxIncl <= 0) return 0;
    return Math.max(1, Math.ceil(offlineTaxIncl / C.storeRevenue));
  }

  /* ============================================================
     五、单年完整计算
     inp：本年输入；prev：上年结果（首年为 null）；idx：年份序号 0—4
     ============================================================ */
  function computeYear(inp, prev, idx, options) {
    const p = Object.assign({}, DEFAULTS, inp || {});
    const stable = !!(options && options.stable);

    // ---------- 1. 收入与渠道拆分（自然年，含税）----------
    const revTaxIncl = p.revenue;
    const revExcl = revTaxIncl / (1 + C.vat);

    const onlineTaxIncl = revTaxIncl * p.chOnline;                       // 直营网店
    const offlineTaxIncl = revTaxIncl * p.chOffline;                     // 直营线下
    const distSingleTaxIncl = revTaxIncl * p.chDistSingle;               // 分销单品牌
    const distMultiTaxIncl = revTaxIncl * p.chDistMulti;                 // 分销多品牌
    const distBuyerTaxIncl = revTaxIncl * p.chBuyer;                     // 分销买手

    const dtcTaxIncl = onlineTaxIncl + offlineTaxIncl;                   // DTC 含税
    const distTaxIncl = distSingleTaxIncl + distMultiTaxIncl + distBuyerTaxIncl;
    const dtcExcl = dtcTaxIncl / (1 + C.vat);
    const distExcl = distTaxIncl / (1 + C.vat);

    // ---------- 2. 渠道数量：店铺 / 客户 / 平台 ----------
    // 店数：优先用输入，否则由线下收入反推
    const stores = (p.stores !== undefined && p.stores !== null)
      ? p.stores : storesFromRevenue(offlineTaxIncl);
    // 全年营业等价店数（上年末店 + 新店当年部分）
    const prevStores = prev ? prev.stores : 0;
    const newStores = Math.max(0, stores - prevStores);
    const equivStores = prevStores + newStores * C.newStorePartial;

    // 分销客户数（按单客户提货额向上取整）
    const clientsSingle = Math.ceil(distSingleTaxIncl / C.clientTicket.single);
    const clientsMulti = Math.ceil(distMultiTaxIncl / C.clientTicket.multi);
    const clientsBuyer = Math.ceil(distBuyerTaxIncl / C.clientTicket.buyer);
    const totalClients = clientsSingle + clientsMulti + clientsBuyer;

    // 线上平台：抖音含税
    const douyinTaxIncl = onlineTaxIncl * p.douyinShare;
    const nonDouyinTaxIncl = onlineTaxIncl - douyinTaxIncl;

    // ---------- 3. 货品：商品季采购（含 FW 跨年）----------
    // 上年 FW 完整商品季含税（首年为 0），用于还原本年自然年里的 FW 尾段
    const prevFwDtc = prev ? prev.fwDtcTaxIncl : 0;
    // 本年初季初库存基准（不含到岸）= 上年 FW 季末
    let openingInvBase = prev ? prev.closingInvBase : 0;

    // SS 完整商品季 DTC 含税 = 自然年 DTC × SS 占比
    const ssDtcTaxIncl = dtcTaxIncl * C.ssShare;
    // FW 完整商品季：自然年 DTC 里，先扣「上年 FW 跨年尾段」和「本年 SS」，
    // 剩下的是本年 FW 当年确认的 75%，再 ÷75% 还原完整商品季
    const prevFwTail = prevFwDtc * (1 - C.fwCurrent);
    const fwDtcTaxIncl = (dtcTaxIncl - prevFwTail - ssDtcTaxIncl) / C.fwCurrent;
    // 商品季分销含税（分销 FW 不跨年，100% 当年交付）
    const ssDistTaxIncl = distTaxIncl * C.ssShare;
    const fwDistTaxIncl = distTaxIncl * (1 - C.ssShare);

    let dtcBuyBase = 0;
    let distBuyBaseTotal = 0;
    const seasons = [
      { dtc: ssDtcTaxIncl, dist: ssDistTaxIncl },
      { dtc: fwDtcTaxIncl, dist: fwDistTaxIncl },
    ];
    seasons.forEach(function (s) {
      const dtcSeasonExcl = s.dtc / (1 + C.vat);
      const distSeasonExcl = s.dist / (1 + C.vat);
      // 老货：按 oldPct 需求，但受期初库存能覆盖的上限约束
      const oldExcl = Math.min(
        dtcSeasonExcl * p.oldPct,
        openingInvBase / (1 - C.dtcMargin)
      );
      const newExcl = dtcSeasonExcl - oldExcl;
      // DTC 新货采购：新货不含税 × (1−毛利率) ÷ 售罄率
      const dtcBuy = newExcl * (1 - C.dtcMargin) / p.sellThrough;
      // 分销采购：分销不含税 × (1−分销毛利率)
      const distBuy = distSeasonExcl * (1 - C.distMargin);
      // 季末库存基准 = 季初 + 新货采购 − 老货已售货值 − 新货已售货值
      openingInvBase = openingInvBase + dtcBuy
        - oldExcl * (1 - C.dtcMargin)
        - newExcl * (1 - C.dtcMargin);
      dtcBuyBase += dtcBuy;
      distBuyBaseTotal += distBuy;
    });
    const closingInvBase = openingInvBase;

    // 品牌货款（不含到岸）= DTC 新货采购 + 分销采购
    const brandPayment = dtcBuyBase + distBuyBaseTotal;
    const landing = brandPayment * C.landingRate;
    const purchaseInclLanding = brandPayment + landing;

    // ---------- 4. 已售成本（自然年，简化公式）----------
    // DTC 已售 = DTC 不含税 × (1−毛利率) × (1+到岸率)
    const dtcCogs = dtcExcl * (1 - C.dtcMargin) * (1 + C.landingRate);
    // 分销已售 = 分销不含税 × (1−分销毛利率) × (1+到岸率)
    const distCogs = distExcl * (1 - C.distMargin) * (1 + C.landingRate);
    const cogs = dtcCogs + distCogs;

    // ---------- 5. 库存（会计恒等式）----------
    const openingInventory = prev ? prev.inventory : 0;
    const inventory = openingInventory + purchaseInclLanding - cogs;

    // ---------- 6. 毛利 ----------
    const grossProfit = revExcl - cogs;
    const grossMargin = revExcl > 0 ? grossProfit / revExcl : 0;

    // ---------- 7. 费用 ----------
    // 费率：允许输入自定义 promoRate / adRate 覆盖（页面滑块用），否则按规则/稳态
    const promoRate = (p.promoRate !== undefined && p.promoRate !== null)
      ? p.promoRate : promoRateOf(idx, stable);
    const adRate = (p.adRate !== undefined && p.adRate !== null)
      ? p.adRate : adRateOf(stable);
    const promo = revTaxIncl * promoRate;                        // 品牌推广（全收入）
    const ad = onlineTaxIncl * adRate;                           // 线上投流（仅线上）
    const onlineOpex = onlineTaxIncl * C.onlineOpexRate;         // 线上运营（佣金/履约/其他）
    const offlineOpex = equivStores * C.storeAnnualOpex;         // 线下运营（租金物业等）
    const distService = distTaxIncl * C.distServiceRate;         // 分销服务
    const backofficeOther = revExcl * C.backofficeRate;          // 后台其他
    const commission = revTaxIncl * p.commission;                // 销售提成

    // ---------- 8. 人员（业务驱动，4:1 拆基础/管理）----------
    const hcOffline = stores * 4 + (stores > 0 ? Math.ceil(stores / 5) : 0);
    const hcPlatformOps = onlineTaxIncl > 0
      ? Math.max(2, Math.ceil(nonDouyinTaxIncl / 450)) : 0;
    // 抖音团队：主播/中控/运营/内容，起量后按抖音规模配置（6 人基础 + 规模增量）
    const hcDouyin = douyinTaxIncl > 0
      ? 6 + Math.round(douyinTaxIncl / 450) : 0;
    const hcOnline = hcPlatformOps + hcDouyin;
    const hcBackoffice = backofficeHeadcount(revTaxIncl);
    const headcount = hcOffline + hcOnline + hcBackoffice;
    const managers = headcount * C.managerRatio;
    const basics = headcount * (1 - C.managerRatio);
    const staffCost = headcount * C.blendedAnnual;

    // ---------- 9. 资本支出、装修资产池与折旧 ----------
    // 承接上年装修资产池（每笔记录投入年份，按 5 年直线折旧）
    const capexPool = prev
      ? prev.capexPool.map(function (a) { return { amount: a.amount, idx: a.idx }; })
      : [];
    // 本年资本支出与押金（新店一次性）
    const capex = newStores * C.storeCapex;
    const deposit = newStores * C.storeDeposit;
    if (capex > 0) capexPool.push({ amount: capex, idx: idx });
    // 折旧：投入当年只折新店营业比例（6 月开业 ≈0.586），之后全年，最多 5 年
    let depreciation = 0;
    capexPool.forEach(function (a) {
      const age = idx - a.idx;
      if (age === 0) depreciation += (a.amount / 5) * C.newStorePartial;
      else if (age <= 4) depreciation += a.amount / 5;
    });

    // ---------- 10. 利润 ----------
    const totalOpex = promo + ad + onlineOpex + offlineOpex
      + distService + backofficeOther + commission + staffCost;
    const ebitda = grossProfit - totalOpex;
    const ebit = ebitda - depreciation;
    const tax = ebit > 0 ? ebit * C.taxRate : 0;
    const netProfit = ebit - tax;

    // ---------- 11. 现金流（年度，不含融资）----------
    // 进口增值税：完税价格 = 品牌货款×(1+运保3%)×(1+关税9%) = 品牌货款×1.1227
    // （不含清关费），再 ×13%
    const importVat = brandPayment * 1.1227 * C.vat;
    // 销售增值税：销项税 − 进口可抵（费用进项默认 0，故只减进口 VAT）
    const outputVat = revTaxIncl * C.vat / (1 + C.vat);
    const salesVat = Math.max(0, outputVat - importVat);
    // 库存增加（采购已付、货未卖的占款）
    const inventoryChange = inventory - openingInventory;
    // 净现金流：由利润严格勾稽——加回折旧（非现金），
    // 扣库存占款、资本支出、押金。增值税为价外税，已在回款/采购口径内抵消。
    const netCash = netProfit + depreciation - inventoryChange - capex - deposit;

    // 资金需求（含安全垫）：覆盖累计净流出
    const prevCumCash = prev ? prev.cumNetCash : 0;
    const cumNetCash = prevCumCash + netCash;
    const fundNeed = Math.max(0, -cumNetCash) + C.cashBuffer;

    return {
      p, stable,
      // 收入
      revTaxIncl, revExcl, onlineTaxIncl, offlineTaxIncl,
      distSingleTaxIncl, distMultiTaxIncl, distBuyerTaxIncl,
      dtcTaxIncl, distTaxIncl, dtcExcl, distExcl,
      // 渠道数量
      stores, newStores, equivStores,
      clientsSingle, clientsMulti, clientsBuyer, totalClients,
      douyinTaxIncl, nonDouyinTaxIncl,
      // 货品
      ssDtcTaxIncl, fwDtcTaxIncl, dtcBuyBase, distBuyBaseTotal,
      brandPayment, landing, purchaseInclLanding, closingInvBase,
      cogs, dtcCogs, distCogs, openingInventory, inventory,
      grossProfit, grossMargin,
      // 费用
      promo, ad, onlineOpex, offlineOpex, distService,
      backofficeOther, commission,
      // 人员
      hcOffline, hcPlatformOps, hcDouyin, hcOnline, hcBackoffice,
      headcount, managers, basics, staffCost,
      // 利润
      totalOpex, ebitda, depreciation, ebit, tax, netProfit,
      // 现金
      importVat, outputVat, salesVat, inventoryChange,
      capex, deposit, capexPool, netCash, cumNetCash, fundNeed,
    };
  }

  /* ============================================================
     六、五年滚动计算
     annualInputs：可选，缺省用 DEFAULT_SCENARIO
     ============================================================ */
  function computeMultiYear(annualInputs, options) {
    const inputs = annualInputs && annualInputs.length
      ? annualInputs : DEFAULT_SCENARIO;
    const years = [];
    let prev = null;
    inputs.forEach(function (inp, idx) {
      const r = computeYear(inp, prev, idx, options);
      years.push(r);
      prev = r;
    });
    return {
      years: years,
      totalNetProfit: years.reduce(function (s, y) { return s + y.netProfit; }, 0),
      totalNetCash: years.reduce(function (s, y) { return s + y.netCash; }, 0),
    };
  }

  /* ============================================================
     七、单年入口（兼容旧页面：只算第一年）
     ============================================================ */
  function compute(input, options) {
    return computeYear(input, null, 0, options);
  }

  /* ============================================================
     八、三时段现金流（节点缺口，第 6 章用）
     基于单年结果 r，按回款/付款的时段规则近似分配
     ============================================================ */
  function computePeriods(r) {
    // 商品季含税规模（SS 与 FW），用于按比例拆品牌货款/到岸/进口税
    const ssSeason = r.ssDtcTaxIncl + r.distTaxIncl * C.ssShare;
    const fwSeason = r.fwDtcTaxIncl + r.distTaxIncl * (1 - C.ssShare);
    const seasonTotal = ssSeason + fwSeason;

    // 回款：1—2 月无回款（首年）；3—8 月收 SS；9—12 月收 FW 头段
    const inflowSS = r.ssDtcTaxIncl + r.distTaxIncl * C.ssShare;
    const inflowFW = r.fwDtcTaxIncl * C.fwCurrent
      + r.distTaxIncl * (1 - C.ssShare);

    // 品牌货款：SS 货款 1—2 月付，FW 货款 3—8 月付
    const brandSS = r.brandPayment * ssSeason / seasonTotal;
    const brandFW = r.brandPayment - brandSS;
    // 到岸与进口增值税：SS 货 3 月到港（计入 3—8 月），FW 货 9 月到港（9—12 月）
    const landingSS = r.landing * ssSeason / seasonTotal;
    const landingFW = r.landing - landingSS;
    const importVatSS = r.importVat * ssSeason / seasonTotal;
    const importVatFW = r.importVat - importVatSS;

    const periods = [
      {
        key: 'p12', name: '1—2 月', inflow: 0,
        brandPay: brandSS, landingPay: 0, importVat: 0, capex: 0,
      },
      {
        key: 'p38', name: '3—8 月', inflow: inflowSS,
        brandPay: brandFW, landingPay: landingSS,
        importVat: importVatSS, capex: r.capex + r.deposit,
      },
      {
        key: 'p912', name: '9—12 月', inflow: inflowFW,
        brandPay: 0, landingPay: landingFW,
        importVat: importVatFW, capex: 0,
      },
    ];

    // 常规费用（含人员、推广、运营、提成，不含资本/押金）按月份权重摊
    const monthW = { p12: 2 / 12, p38: 6 / 12, p912: 4 / 12 };
    let running = 0;
    let totalFunding = 0;
    let worstAfter = 0;
    periods.forEach(function (pd) {
      pd.regularOpex = r.totalOpex * monthW[pd.key];
      // 付款前现金：时段初现金 + 回款 − 到岸/进口税/常规费用/资本，未付大额货款
      pd.beforePay = running + pd.inflow - pd.landingPay
        - pd.importVat - pd.regularOpex - pd.capex;
      // 付款后（未融资）现金，即本时段资金最低点
      pd.afterPayRaw = pd.beforePay - pd.brandPay;
      // 跌破安全线则即时融资补足到安全线
      pd.funding = Math.max(0, C.cashBuffer - pd.afterPayRaw);
      totalFunding += pd.funding;
      pd.afterPay = pd.afterPayRaw + pd.funding;
      pd.net = pd.afterPay - running;
      running = pd.afterPay;
      pd.endCash = running;
      if (pd.afterPayRaw < worstAfter) worstAfter = pd.afterPayRaw;
    });

    return { periods: periods, worstGap: worstAfter, totalFunding: totalFunding };
  }

  /* ============================================================
     九、格式化
     ============================================================ */
  function fmtWan(x, d) {
    if (x === null || x === undefined || isNaN(x)) return '—';
    d = (d === undefined) ? 1 : d;
    const s = Number(x).toFixed(d);
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
  function fmtYuan(x) {
    if (x === null || x === undefined || isNaN(x)) return '—';
    return (x * 10000).toLocaleString('zh-CN', { maximumFractionDigits: 0 });
  }

  /* ============================================================
     十、风险检测（调数据时实时触发）
     ============================================================ */
  function assessRisk(r) {
    const risks = [];
    const p = r.p;

    if (p.sellThrough < 0.5) {
      risks.push({
        level: 'warn', title: '库存积压',
        text: '售罄率 ' + fmtPct(p.sellThrough) + ' 偏低，年末约 ' + fmtWan(r.inventory) + ' 万压在库里，占现金、占仓储。'
      });
    }

    const chTotal = p.chOnline + p.chOffline + p.chDistSingle + p.chDistMulti + p.chBuyer;
    if (Math.abs(chTotal - 1) > 0.005) {
      risks.push({
        level: 'warn', title: '渠道结构错配',
        text: '五类渠道合计 ' + fmtPct(chTotal) + '，不是 100%。蛋糕切多了或切少了。'
      });
    }

    if (r.netCash < 0) {
      risks.push({
        level: 'warn', title: '现金流为负',
        text: '净现金流 ' + fmtWan(r.netCash) + ' 万，资金需求约 ' + fmtWan(r.fundNeed) + ' 万。大额货款付款日前必须备足，否则断供。'
      });
    }

    if (r.netProfit < 0) {
      risks.push({
        level: 'note', title: '账面亏损',
        text: '净利润 ' + fmtWan(r.netProfit) + ' 万。要分得清：这是战略投入期，还是模式没跑通。'
      });
    }

    const mktRate = r.revTaxIncl > 0 ? (r.promo + r.ad) / r.revTaxIncl : 0;
    if (mktRate > 0.30) {
      risks.push({
        level: 'warn', title: '费用吃毛利',
        text: '推广+投流占含税收入 ' + fmtPct(mktRate) + '，两笔钱都在吞毛利，别只盯一个费率。'
      });
    }

    if (r.revTaxIncl > 0 && r.staffCost / r.revTaxIncl > 0.2) {
      risks.push({
        level: 'note', title: '人力超前',
        text: '人员成本占收入 ' + fmtPct(r.staffCost / r.revTaxIncl) + '，偏高。人跟业务量走，不是提前配满。'
      });
    }

    risks.push({
      level: 'note', title: '简化口径',
      text: '到岸、关税、人力均为预算估值。真实关税按 HS 编码与原产地核定，逐岗薪酬更高，隐性费用没算——真实只会更坏。'
    });

    return risks;
  }

  global.CalcModel = {
    C: C,
    DEFAULTS: DEFAULTS,
    DEFAULT_SCENARIO: DEFAULT_SCENARIO,
    compute: compute,
    computeYear: computeYear,
    computeMultiYear: computeMultiYear,
    computePeriods: computePeriods,
    backofficeHeadcount: backofficeHeadcount,
    storesFromRevenue: storesFromRevenue,
    assessRisk: assessRisk,
    fmtWan: fmtWan,
    fmtPct: fmtPct,
    fmtInt: fmtInt,
    fmtYuan: fmtYuan,
  };
})(typeof window !== 'undefined' ? window : globalThis);
