/**
 * 税金・社会保険料最適化プログラム（完全版 v15）
 * 【最適化目的】総資産の最大化（3人の手取り + Linhコスト + P&I税引後利益の実質価値 + 企業型DC・中退共の実質価値）
 *   ※内部留保は将来役員退職金として払出時の実効税率を考慮
 *   ※企業型DC・中退共は将来退職所得として受取時の実効税率10%を想定
 * 【制約条件】
 *   1. P&Iの税引後利益 ≥ 0
 *   2. 林の事業所得 ≥ 0
 *   3. 土井の月次給与 ≥ 35万円
 * 【社会保険料】健康保険 + 介護保険 + 厚生年金 + 雇用保険
 * 【退職金制度】中小企業退職金共済（土井郁子・Linh）
 * ・居住地：林佑樹・Linh→神戸市、土井郁子→西宮市
 * ・Veltra業務委託990万円（固定・事業所得）
 * ・P&Iから林への役員報酬（可変・給与所得・定期同額給与）
 * ・青色申告控除65万円
 * ・専従者給与96万円（手取り計算に含む）
 * ・個人事業経費約300万円
 * ・土井郁子：中退共3万円/月
 * ・Linh：中退共5千円/月、給与は兵庫県最低賃金で固定
 * ・林佑樹：住宅ローン控除
 * ・Linh：配偶者控除（無収入配偶者）
 * ・Linh：手取りではなく給与+社保会社負担+中退共で評価
 * ・ふるさと納税限度額算出
 */

// ============================================================
// メイン関数：スプレッドシートにメニューを追加
// ============================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('税金最適化')
    .addItem('初期設定', 'setupSheet')
    .addItem('最適化実行（総資産の最大化）', 'runOptimization')
    .addToUi();
}

// ============================================================
// 初期設定：シートのセットアップ
// ============================================================
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存の設定値を保存
  let existingValues = {};
  let inputSheet = ss.getSheetByName('optimization');
  if (inputSheet) {
    // 既存のシートがある場合、現在の値を読み取る
    const lastRow = inputSheet.getLastRow();
    if (lastRow > 1) {
      const existingData = inputSheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (let i = 0; i < existingData.length; i++) {
        const label = existingData[i][0];
        const value = existingData[i][1];
        if (label && value !== '' && value !== null) {
          existingValues[label] = value;
        }
      }
    }
  } else {
    inputSheet = ss.insertSheet('optimization');
  }
  inputSheet.clear();

  // ヘッダー設定
  inputSheet.getRange('A1:B1').setValues([['項目', '値']]);
  inputSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');

  // 入力項目
  const inputs = [
    ['【林佑樹の事業収入（固定）】', ''],
    ['Veltra業務委託（円/年）', 9900000],
    ['個人コンサル収入（円/年）', 5000000],
    ['', ''],
    ['【林佑樹の個人事業経費】', ''],
    ['個人事業の固定経費（円/年）', 3000000],
    ['専従者給与（円/年）', 960000],
    ['青色申告控除（円/年）', 650000],
    ['', ''],
    ['【林佑樹の所得控除】', ''],
    ['住宅ローン控除額（円/年）', 200000],
    ['', ''],
    ['【中小企業退職金共済（中退共）】', ''],
    ['土井郁子・掛け金（円/月）', 30000],
    ['Linh・掛け金（円/月）', 5000],
    ['', ''],
    ['【P&Iの収入】', ''],
    ['コンサルティング売上（円/年）', 30000000],
    ['', ''],
    ['【P&Iの固定費】', ''],
    ['オフィス賃料（円/年）', 6000000],
    ['その他固定経費（円/年）', 2000000],
    ['', ''],
    ['【最適化変数の初期値】', ''],
    ['林・事務委託費（円/年）', 3000000],
    ['林・役員報酬（円/月）※給与所得・定期同額', 500000],
    ['土井郁子・給与（円/月）', 400000],
    ['Linh・給与（円/月）', 350000],
    ['', ''],
    ['【探索範囲：事務委託費】', ''],
    ['最小値（円/年）', 0],
    ['最大値（円/年）', 10000000],
    ['刻み幅（円）', 500000],
    ['', ''],
    ['【探索範囲：給与】', ''],
    ['最小値（円/月）', 150000],
    ['最大値（円/月）', 800000],
    ['刻み幅（円）', 50000],
    ['', ''],
    ['【売上バリエーション】', ''],
    ['最大増加額（円/年）', 30000000],
    ['刻み幅（円/年）', 3000000],
    ['', ''],
    ['【その他設定】', ''],
    ['協会けんぽ料率・兵庫県（%）', 10.29],
    ['介護保険料率（%）', 1.60],
    ['厚生年金料率（%）', 18.3],
    ['雇用保険料率・労働者（%）', 0.6],
    ['雇用保険料率・事業主（%）', 0.95],
    ['', ''],
    ['【内部留保の将来コスト】', ''],
    ['退職金の想定税率（%）※最適化用・理論上限27.5%', 27.5],
    ['', ''],
    ['【払い出し最適化設定】', ''],
    ['役員退職金・想定勤続年数（年）', 30],
    ['小規模企業共済・林掛金（円/月）※最大7万円', 0],
    ['企業型DC・林掛金（円/月）※最大5.5万円', 55000],
    ['企業型DC・土井掛金（円/月）※最大5.5万円', 55000],
    ['企業型DC・Linh掛金（円/月）', 27500],
    ['', ''],
    ['【土井郁子・役員設定】', ''],
    ['役員退職金・想定勤続年数（年）', 30]
  ];
  
  // 既存の値がある場合はそれを使用、ない場合はデフォルト値を使用
  const finalInputs = inputs.map(row => {
    const label = row[0];
    const defaultValue = row[1];
    // ラベルが存在し、既存の値がある場合はそれを使用
    if (label && existingValues.hasOwnProperty(label)) {
      return [label, existingValues[label]];
    }
    return row;
  });

  inputSheet.getRange(2, 1, finalInputs.length, 2).setValues(finalInputs);
  inputSheet.setColumnWidth(1, 300);
  inputSheet.setColumnWidth(2, 150);
  
  // B列を右揃えに
  inputSheet.getRange('B:B').setHorizontalAlignment('right');
  
  // カテゴリ行に色付け
  const categoryRows = [2, 6, 11, 14, 18, 21, 25, 31, 36, 41, 45, 52, 55, 62];
  categoryRows.forEach(row => {
    inputSheet.getRange(row, 1).setBackground('#e8f0fe').setFontWeight('bold');
  });

  // 重要な固定値を強調
  inputSheet.getRange('A3:B4').setBackground('#fff9c4');
  inputSheet.getRange('A8:B9').setBackground('#fce4ec');
  
  // 結果シートの作成
  let resultSheet = ss.getSheetByName('最適化結果');
  if (!resultSheet) {
    resultSheet = ss.insertSheet('最適化結果');
  }

  const preservedCount = Object.keys(existingValues).length;
  if (preservedCount > 0) {
    Logger.log('初期設定が完了しました（既存の設定値 ' + preservedCount + ' 項目を保持）');
  } else {
    Logger.log('初期設定が完了しました');
  }
}

// ============================================================
// 社会保険料の計算（健保+介護+年金+雇用保険）
// ============================================================
function calcShakaihoken(monthlySalary, kenpoRate = 0.1029, kaigoRate = 0.016, nenkinRate = 0.183, koyoRateWorker = 0.006, koyoRateEmployer = 0.0095) {
  const stdSalary = getStandardSalary(monthlySalary);

  // 標準報酬月額ベースの保険料（健保+介護+年金）
  const shahoMonthly = stdSalary * (kenpoRate + kaigoRate + nenkinRate);

  // 雇用保険料（実際の給与ベース）
  const koyoMonthlyWorker = monthlySalary * koyoRateWorker;
  const koyoMonthlyEmployer = monthlySalary * koyoRateEmployer;

  // 年額計算
  const shahoYearly = shahoMonthly * 12;
  const koyoYearlyWorker = koyoMonthlyWorker * 12;
  const koyoYearlyEmployer = koyoMonthlyEmployer * 12;

  return {
    total: shahoYearly + koyoYearlyWorker + koyoYearlyEmployer,
    worker: (shahoYearly / 2) + koyoYearlyWorker,  // 本人負担
    employer: (shahoYearly / 2) + koyoYearlyEmployer  // 会社負担
  };
}

// 標準報酬月額の取得
function getStandardSalary(salary) {
  const grades = [
    [63000, 58000], [73000, 68000], [83000, 78000], [93000, 88000],
    [101000, 98000], [107000, 104000], [114000, 110000], [122000, 118000],
    [130000, 126000], [138000, 134000], [146000, 142000], [155000, 150000],
    [165000, 160000], [175000, 170000], [185000, 180000], [195000, 190000],
    [210000, 200000], [230000, 220000], [250000, 240000], [270000, 260000],
    [290000, 280000], [310000, 300000], [330000, 320000], [350000, 340000],
    [370000, 360000], [395000, 380000], [425000, 410000], [455000, 440000],
    [485000, 470000], [515000, 500000], [545000, 530000], [575000, 560000],
    [605000, 590000], [635000, 620000], [665000, 650000]
  ];
  
  for (let i = 0; i < grades.length; i++) {
    if (salary <= grades[i][0]) {
      return grades[i][1];
    }
  }
  return 650000; // 上限
}

// ============================================================
// 所得税の計算（累進課税）
// ============================================================
function calcIncomeTax(taxableIncome) {
  let tax = 0;
  
  if (taxableIncome <= 1950000) {
    tax = taxableIncome * 0.05;
  } else if (taxableIncome <= 3300000) {
    tax = 97500 + (taxableIncome - 1950000) * 0.1;
  } else if (taxableIncome <= 6950000) {
    tax = 427500 + (taxableIncome - 3300000) * 0.2;
  } else if (taxableIncome <= 9000000) {
    tax = 1357500 + (taxableIncome - 6950000) * 0.23;
  } else if (taxableIncome <= 18000000) {
    tax = 1828000 + (taxableIncome - 9000000) * 0.33;
  } else if (taxableIncome <= 40000000) {
    tax = 4798500 + (taxableIncome - 18000000) * 0.4;
  } else {
    tax = 13598500 + (taxableIncome - 40000000) * 0.45;
  }
  
  return tax * 1.021; // 復興特別所得税 2.1%
}

// ============================================================
// 所得税率の取得
// ============================================================
function getIncomeTaxRate(taxableIncome) {
  if (taxableIncome <= 1950000) return 0.05;
  if (taxableIncome <= 3300000) return 0.10;
  if (taxableIncome <= 6950000) return 0.20;
  if (taxableIncome <= 9000000) return 0.23;
  if (taxableIncome <= 18000000) return 0.33;
  if (taxableIncome <= 40000000) return 0.40;
  return 0.45;
}

// ============================================================
// 住民税の計算（神戸市・西宮市）
// ============================================================
function calcResidentTax(taxableIncome, city = 'kobe') {
  if (taxableIncome <= 0) {
    // 均等割のみ
    return city === 'nishinomiya' ? 5300 : 5300; // 神戸市・西宮市とも5300円
  }
  // 所得割10% + 均等割
  const shotokuwari = taxableIncome * 0.1;
  const kintowari = city === 'nishinomiya' ? 5300 : 5300;
  return shotokuwari + kintowari;
}

// ============================================================
// 法人税の計算
// ============================================================
function calcCorporateTax(corporateIncome) {
  if (corporateIncome <= 0) return 0;
  
  // 法人税
  let houjinzei = corporateIncome <= 8000000 
    ? corporateIncome * 0.15 
    : 8000000 * 0.15 + (corporateIncome - 8000000) * 0.232;
  
  // 地方法人税
  const chihouHoujinzei = houjinzei * 0.103;
  
  // 法人住民税
  const houjinJuminzei = houjinzei * 0.07 + 70000;
  
  // 法人事業税
  const jigyouzei = corporateIncome * 0.07;
  
  // 特別地方法人税
  const tokubetsuChihou = jigyouzei * 0.37;
  
  return houjinzei + chihouHoujinzei + houjinJuminzei + jigyouzei + tokubetsuChihou;
}

// ============================================================
// 役員退職金の実効税率計算
// ============================================================
function calcRetirementTaxRate(retirementPay, yearsOfService) {
  if (retirementPay <= 0 || yearsOfService <= 0) return 0;

  // 退職所得控除の計算
  let deduction = 0;
  if (yearsOfService <= 20) {
    deduction = Math.max(800000, 400000 * yearsOfService);
  } else {
    deduction = 8000000 + 700000 * (yearsOfService - 20);
  }

  // 退職所得の計算（1/2課税）
  const retirementIncome = Math.max(0, (retirementPay - deduction) / 2);

  // 所得税の計算（分離課税）
  const incomeTax = calcIncomeTax(retirementIncome);

  // 住民税の計算（10%）
  const residentTax = retirementIncome * 0.1;

  // 実効税率
  const effectiveTaxRate = (incomeTax + residentTax) / retirementPay;

  return effectiveTaxRate;
}

// ============================================================
// 給与所得控除の計算
// ============================================================
function calcSalaryDeduction(grossIncome) {
  if (grossIncome <= 1625000) return 550000;
  if (grossIncome <= 1800000) return grossIncome * 0.4 - 100000;
  if (grossIncome <= 3600000) return grossIncome * 0.3 + 80000;
  if (grossIncome <= 6600000) return grossIncome * 0.2 + 440000;
  if (grossIncome <= 8500000) return grossIncome * 0.1 + 1100000;
  return 1950000;
}

// ============================================================
// 個人事業税の計算
// ============================================================
function calcKojinJigyoTax(jigyoShotoku) {
  // 事業主控除：290万円
  const deduction = 2900000;
  const taxableIncome = Math.max(0, jigyoShotoku - deduction);
  // 税率：5%
  return taxableIncome * 0.05;
}

// ============================================================
// 配偶者控除額の計算（無収入配偶者）
// ============================================================
function calcHaigushaKoujo(totalIncome, isIncomeTax = true) {
  // 合計所得金額に応じた控除額
  if (totalIncome > 10000000) return 0;

  if (isIncomeTax) {
    // 所得税の配偶者控除
    if (totalIncome <= 9000000) return 380000;
    if (totalIncome <= 9500000) return 260000;
    if (totalIncome <= 10000000) return 130000;
  } else {
    // 住民税の配偶者控除
    if (totalIncome <= 9000000) return 330000;
    if (totalIncome <= 9500000) return 220000;
    if (totalIncome <= 10000000) return 110000;
  }

  return 0;
}

// ============================================================
// ふるさと納税上限額の計算（金銭的メリットがある額）
// ============================================================
function calcFurusatoLimit(taxableIncome, incomeTaxRate, incomeTax, residentTax) {
  // 所得税・住民税の合計が2000円以下ならふるさと納税のメリットなし
  const totalTax = incomeTax + residentTax;
  if (totalTax <= 2000) return 0;
  
  if (taxableIncome <= 0) return 0;
  
  const residentTaxShotokuwari = taxableIncome * 0.1;
  
  // 控除上限額 = 住民税所得割額 × 20% / (90% - 所得税率 × 1.021) + 2,000円
  const limit = (residentTaxShotokuwari * 0.2) / (0.9 - incomeTaxRate * 1.021) + 2000;
  
  // 実際に控除できる税額を超えないように制限
  // ふるさと納税で控除できるのは (寄付額 - 2000円) まで
  // 実際の控除額 = 所得税控除 + 住民税控除
  const maxBenefit = totalTax - 2000; // 2000円は自己負担として残る
  const calculatedLimit = Math.floor(limit);
  
  // メリットがある上限額（2000円の自己負担を考慮）
  return Math.min(calculatedLimit, maxBenefit + 2000);
}

// ============================================================
// 総コストの計算
// ============================================================
function calcTotalCost(params) {
  const {
    hayashiYakuin,    // 林の報酬（月額）
    doiSalary,        // 土井の給与（月額）
    linhSalary,       // Linhの給与（月額）
    jimuItakuhi,      // P&Iへの事務委託費（年額）
    veltraSalary,     // Veltraからの給与（年額・固定）
    kojinRevenue,     // 個人事業のコンサル収入（年額）
    kojinExpense,     // 個人事業の固定経費（年額）
    haigusha,         // 専従者給与（年額）
    aoiroDeduction,   // 青色申告控除（年額）
    housingLoanDeduction, // 住宅ローン控除（年額）
    doiChutaikyo,     // 土井の中退共（月額）
    linhChutaikyo,    // Linhの中退共（月額）
    consultingRevenue, // P&Iのコンサル売上（年額）
    officeExpense,    // P&Iのオフィス経費（年額）
    otherExpense,     // P&Iのその他経費（年額）
    kenpoRate,        // 健保料率
    kaigoRate,        // 介護保険料率
    nenkinRate,       // 年金料率
    koyoRateWorker,   // 雇用保険料率・労働者
    koyoRateEmployer, // 雇用保険料率・事業主
    shoukiboKyousai,  // 小規模企業共済掛金（月額・林）
    hayashiKigyouDC,  // 企業型DC掛金（月額・林）
    doiKigyouDC,      // 企業型DC掛金（月額・土井）
    linhKigyouDC,     // 企業型DC掛金（月額・Linh）
    doiIsYakuin,      // 土井が役員かどうか（1=役員、0=従業員）
    doiRetirementYears // 土井の想定勤続年数（役員の場合）
  } = params;
  
  // 年額
  const hayashiYakuinYearly = hayashiYakuin * 12;
  const doiSalaryYearly = doiSalary * 12;
  const linhSalaryYearly = linhSalary * 12;

  // 小規模企業共済（林・年額）
  const shoukiboKyousaiYearly = shoukiboKyousai * 12;

  // 企業型DC（年額）
  const hayashiKigyouDCYearly = hayashiKigyouDC * 12;
  const doiKigyouDCYearly = doiKigyouDC * 12;
  const linhKigyouDCYearly = linhKigyouDC * 12;
  
  // 社会保険料
  const hayashiInsuranceObj = calcShakaihoken(hayashiYakuin, kenpoRate, kaigoRate, nenkinRate, koyoRateWorker, koyoRateEmployer);
  // 土井が役員の場合は雇用保険料なし
  const doiKoyoWorker = doiIsYakuin ? 0 : koyoRateWorker;
  const doiKoyoEmployer = doiIsYakuin ? 0 : koyoRateEmployer;
  const doiInsuranceObj = calcShakaihoken(doiSalary, kenpoRate, kaigoRate, nenkinRate, doiKoyoWorker, doiKoyoEmployer);
  const linhInsuranceObj = calcShakaihoken(linhSalary, kenpoRate, kaigoRate, nenkinRate, koyoRateWorker, koyoRateEmployer);

  const hayashiInsurance = hayashiInsuranceObj.total;
  const doiInsurance = doiInsuranceObj.total;
  const linhInsurance = linhInsuranceObj.total;

  const companyInsuranceBurden = hayashiInsuranceObj.employer + doiInsuranceObj.employer + linhInsuranceObj.employer;
  
  // ============================================================
  // 専従者の手取り計算
  // ============================================================
  const haigushaKyuyoShotoku = haigusha - calcSalaryDeduction(haigusha);
  const haigushaDeduction = 480000; // 基礎控除のみ（社保なし）
  const haigushaTaxableIncome = Math.max(0, haigushaKyuyoShotoku - haigushaDeduction);
  const haigushaIncomeTax = calcIncomeTax(haigushaTaxableIncome);
  const haigushaIncomeTaxRate = getIncomeTaxRate(haigushaTaxableIncome);
  const haigushaResidentTax = calcResidentTax(haigushaTaxableIncome, 'kobe');
  const haigushaTedori = haigusha - haigushaIncomeTax - haigushaResidentTax;
  const haigushaFurusatoLimit = calcFurusatoLimit(haigushaTaxableIncome, haigushaIncomeTaxRate,
                                                   haigushaIncomeTax, haigushaResidentTax);
  
  // ============================================================
  // 合同会社P&Iの損益
  // ============================================================
  const piRevenue = consultingRevenue + jimuItakuhi;

  // 中退共（年額）※役員は中退共に加入できない
  const doiChutaikyoYearly = doiIsYakuin ? 0 : doiChutaikyo * 12;
  const linhChutaikyoYearly = linhChutaikyo * 12;

  // 林への役員報酬（定期同額給与）、土井・Linhの給与、中退共、企業型DCを経費計上
  const piExpense = hayashiYakuinYearly + doiSalaryYearly + linhSalaryYearly +
                    companyInsuranceBurden + doiChutaikyoYearly + linhChutaikyoYearly +
                    hayashiKigyouDCYearly + doiKigyouDCYearly + linhKigyouDCYearly +
                    officeExpense + otherExpense;
  const piIncome = piRevenue - piExpense;
  const piTax = calcCorporateTax(piIncome);
  const piTaxRate = piIncome > 0 ? piTax / piIncome : 0; // 実効税率
  
  // ============================================================
  // 林佑樹の個人事業（事業所得）
  // ============================================================
  // 事業収入：Veltra業務委託 + 個人コンサル収入
  const kojinTotalRevenue = veltraSalary + kojinRevenue;
  const kojinTotalExpense = kojinExpense + haigusha + jimuItakuhi + aoiroDeduction;
  const kojinJigyoShotoku = kojinTotalRevenue - kojinTotalExpense;

  // ============================================================
  // 林佑樹の給与所得
  // ============================================================
  // P&Iからの役員報酬（定期同額給与）
  const piKyuyoShotoku = hayashiYakuinYearly - calcSalaryDeduction(hayashiYakuinYearly);

  // ============================================================
  // 林佑樹の総所得と課税所得
  // ============================================================
  const hayashiTotalIncome = piKyuyoShotoku + kojinJigyoShotoku;
  
  // 所得控除
  const hayashiInsurancePersonal = hayashiInsuranceObj.worker;
  const hayashiDeduction = 480000 + hayashiInsurancePersonal + shoukiboKyousaiYearly + hayashiKigyouDCYearly; // 基礎控除+社会保険料控除+小規模企業共済+企業型DC
  
  // 課税所得
  const hayashiTaxableIncome = Math.max(0, hayashiTotalIncome - hayashiDeduction);
  
  // 所得税（住宅ローン控除前）
  const hayashiIncomeTaxBeforeHousing = calcIncomeTax(hayashiTaxableIncome);
  const hayashiIncomeTax = Math.max(0, hayashiIncomeTaxBeforeHousing - housingLoanDeduction);
  
  // 住宅ローン控除の住民税への適用（所得税で引ききれない分）
  const housingLoanToResident = Math.max(0, housingLoanDeduction - hayashiIncomeTaxBeforeHousing);
  const housingLoanResidentLimit = Math.min(housingLoanToResident, 
    Math.min(hayashiTaxableIncome * 0.05, 97500)); // 住民税からの控除上限
  
  // 所得税率
  const hayashiIncomeTaxRate = getIncomeTaxRate(hayashiTaxableIncome);
  
  // 住民税
  const hayashiResidentTaxBeforeHousing = calcResidentTax(hayashiTaxableIncome, 'kobe');
  const hayashiResidentTax = Math.max(0, hayashiResidentTaxBeforeHousing - housingLoanResidentLimit);
  
  // 個人事業税
  const hayashiJigyoTax = calcKojinJigyoTax(kojinJigyoShotoku);
  
  // ふるさと納税上限
  const hayashiFurusatoLimit = calcFurusatoLimit(hayashiTaxableIncome, hayashiIncomeTaxRate, 
                                                   hayashiIncomeTax, hayashiResidentTax);
  
  // ============================================================
  // 土井郁子の所得税・住民税
  // ============================================================
  const doiKyuyoShotoku = doiSalaryYearly - calcSalaryDeduction(doiSalaryYearly);
  const doiInsurancePersonal = doiInsuranceObj.worker;
  const doiDeduction = 480000 + doiInsurancePersonal; // 基礎控除+社保
  const doiTaxableIncome = Math.max(0, doiKyuyoShotoku - doiDeduction);
  const doiIncomeTax = calcIncomeTax(doiTaxableIncome);
  const doiIncomeTaxRate = getIncomeTaxRate(doiTaxableIncome);
  const doiResidentTax = calcResidentTax(doiTaxableIncome, 'nishinomiya');
  const doiFurusatoLimit = calcFurusatoLimit(doiTaxableIncome, doiIncomeTaxRate,
                                              doiIncomeTax, doiResidentTax);
  
  // ============================================================
  // Linhの所得税・住民税
  // ============================================================
  const linhKyuyoShotoku = linhSalaryYearly - calcSalaryDeduction(linhSalaryYearly);
  const linhInsurancePersonal = linhInsuranceObj.worker;

  // 配偶者控除（無収入配偶者）
  const linhHaigushaKoujoIncomeTax = calcHaigushaKoujo(linhKyuyoShotoku, true);
  const linhHaigushaKoujoResidentTax = calcHaigushaKoujo(linhKyuyoShotoku, false);

  // 所得税の課税所得
  const linhDeductionIncomeTax = 480000 + linhInsurancePersonal + linhHaigushaKoujoIncomeTax;
  const linhTaxableIncomeForIncomeTax = Math.max(0, linhKyuyoShotoku - linhDeductionIncomeTax);
  const linhIncomeTax = calcIncomeTax(linhTaxableIncomeForIncomeTax);
  const linhIncomeTaxRate = getIncomeTaxRate(linhTaxableIncomeForIncomeTax);

  // 住民税の課税所得
  const linhDeductionResidentTax = 480000 + linhInsurancePersonal + linhHaigushaKoujoResidentTax;
  const linhTaxableIncomeForResidentTax = Math.max(0, linhKyuyoShotoku - linhDeductionResidentTax);
  const linhResidentTax = calcResidentTax(linhTaxableIncomeForResidentTax, 'kobe');

  // ふるさと納税上限額（所得税の課税所得で計算）
  const linhFurusatoLimit = calcFurusatoLimit(linhTaxableIncomeForIncomeTax, linhIncomeTaxRate,
                                               linhIncomeTax, linhResidentTax);
  
  // ============================================================
  // 手取り計算
  // ============================================================
  // 林佑樹の手取り
  // 収入：P&I役員報酬 + Veltra業務委託 + 個人コンサル
  // 支出：税金 + 社保（本人負担）+ 実費経費（固定経費 + 専従者給与 + 事務委託費）
  // 注意：青色申告控除は税制上の控除であり実支出ではないので引かない
  const hayashiTedori = hayashiYakuinYearly + kojinTotalRevenue -
                        hayashiIncomeTax - hayashiResidentTax - hayashiJigyoTax -
                        hayashiInsurancePersonal -
                        (kojinExpense + haigusha + jimuItakuhi); // 青色控除を除外
  
  const doiTedori = doiSalaryYearly - doiIncomeTax - doiResidentTax -
                    doiInsurancePersonal;
  
  const linhTedori = linhSalaryYearly - linhIncomeTax - linhResidentTax - 
                     linhInsurancePersonal;
  
  // 専従者の手取りを追加
  const totalTedori = hayashiTedori + doiTedori + linhTedori + haigushaTedori;
  
  // ============================================================
  // 総コスト
  // ============================================================
  const totalInsurance = hayashiInsurance + doiInsurance + linhInsurance;
  const totalTax = piTax + hayashiIncomeTax + hayashiResidentTax + hayashiJigyoTax +
                   doiIncomeTax + doiResidentTax + linhIncomeTax + linhResidentTax +
                   haigushaIncomeTax + haigushaResidentTax;
  const totalCost = totalInsurance + totalTax;
  
  return {
    totalCost: totalCost,
    totalInsurance: totalInsurance,
    totalTax: totalTax,
    totalTedori: totalTedori,
    
    // P&I関連
    piRevenue: piRevenue,
    piExpense: piExpense,
    piIncome: piIncome,
    piTax: piTax,
    piTaxRate: piTaxRate,
    
    // 林佑樹関連
    veltraSalary: veltraSalary,
    piYakuin: hayashiYakuinYearly,
    piKyuyoShotoku: piKyuyoShotoku,
    kojinRevenue: kojinRevenue,
    kojinExpense: kojinExpense,
    haigusha: haigusha,
    aoiroDeduction: aoiroDeduction,
    kojinTotalRevenue: kojinTotalRevenue,
    kojinTotalExpense: kojinTotalExpense,
    kojinJigyoShotoku: kojinJigyoShotoku,
    hayashiTotalIncome: hayashiTotalIncome,
    hayashiTaxableIncome: hayashiTaxableIncome,
    hayashiIncomeTax: hayashiIncomeTax,
    hayashiIncomeTaxRate: hayashiIncomeTaxRate,
    hayashiResidentTax: hayashiResidentTax,
    hayashiJigyoTax: hayashiJigyoTax,
    hayashiInsurance: hayashiInsurance,
    hayashiInsuranceWorker: hayashiInsuranceObj.worker,
    hayashiInsuranceEmployer: hayashiInsuranceObj.employer,
    hayashiTedori: hayashiTedori,
    hayashiFurusatoLimit: hayashiFurusatoLimit,
    housingLoanDeduction: housingLoanDeduction,
    officeExpense: officeExpense,
    otherExpense: otherExpense,
    shoukiboKyousai: shoukiboKyousaiYearly,
    hayashiKigyouDC: hayashiKigyouDCYearly,

    // 土井郁子関連
    doiIncomeTax: doiIncomeTax,
    doiIncomeTaxRate: doiIncomeTaxRate,
    doiResidentTax: doiResidentTax,
    doiInsurance: doiInsurance,
    doiInsuranceWorker: doiInsuranceObj.worker,
    doiInsuranceEmployer: doiInsuranceObj.employer,
    doiTedori: doiTedori,
    doiFurusatoLimit: doiFurusatoLimit,
    doiChutaikyo: doiChutaikyoYearly,
    doiKigyouDC: doiKigyouDCYearly,
    doiIsYakuin: doiIsYakuin,
    doiRetirementYears: doiRetirementYears,

    // Linh関連
    linhSalaryYearly: linhSalaryYearly,
    linhHaigushaKoujoIncomeTax: linhHaigushaKoujoIncomeTax,
    linhHaigushaKoujoResidentTax: linhHaigushaKoujoResidentTax,
    linhTaxableIncomeForIncomeTax: linhTaxableIncomeForIncomeTax,
    linhTaxableIncomeForResidentTax: linhTaxableIncomeForResidentTax,
    linhIncomeTax: linhIncomeTax,
    linhIncomeTaxRate: linhIncomeTaxRate,
    linhResidentTax: linhResidentTax,
    linhInsurance: linhInsurance,
    linhInsuranceWorker: linhInsuranceObj.worker,
    linhInsuranceEmployer: linhInsuranceObj.employer,
    linhTedori: linhTedori,
    linhCost: linhSalaryYearly + linhInsuranceObj.employer + linhChutaikyoYearly + linhKigyouDCYearly,
    linhFurusatoLimit: linhFurusatoLimit,
    linhChutaikyo: linhChutaikyoYearly,
    linhKigyouDC: linhKigyouDCYearly,
    
    // 専従者関連
    haigushaIncomeTax: haigushaIncomeTax,
    haigushaIncomeTaxRate: haigushaIncomeTaxRate,
    haigushaResidentTax: haigushaResidentTax,
    haigushaTedori: haigushaTedori,
    haigushaFurusatoLimit: haigushaFurusatoLimit
  };
}

// ============================================================
// 最適化実行（総資産の最大化 + 制約：P&I利益≥0 & 林事業所得≥0 & 土井給与≥35万円、Linh給与は最低賃金で固定）
// ※内部留保は将来の払出時課税を考慮した実質価値で評価
// ============================================================
function runOptimization() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('optimization');
  
  if (!inputSheet) {
    Logger.log('エラー：先に「初期設定」を実行してください');
    return;
  }
  
  // 入力値の取得
  const veltraSalary = inputSheet.getRange('B3').getValue();
  const kojinRevenue = inputSheet.getRange('B4').getValue();
  const kojinExpense = inputSheet.getRange('B7').getValue();
  const haigusha = inputSheet.getRange('B8').getValue();
  const aoiroDeduction = inputSheet.getRange('B9').getValue();
  const housingLoanDeduction = inputSheet.getRange('B12').getValue();
  const doiChutaikyo = inputSheet.getRange('B15').getValue();
  const linhChutaikyo = inputSheet.getRange('B16').getValue();
  const consultingRevenueBase = inputSheet.getRange('B19').getValue();
  const officeExpense = inputSheet.getRange('B22').getValue();
  const otherExpense = inputSheet.getRange('B23').getValue();
  const kenpoRate = inputSheet.getRange('B46').getValue() / 100;
  const kaigoRate = inputSheet.getRange('B47').getValue() / 100;
  const nenkinRate = inputSheet.getRange('B48').getValue() / 100;
  const koyoRateWorker = inputSheet.getRange('B49').getValue() / 100;
  const koyoRateEmployer = inputSheet.getRange('B50').getValue() / 100;
  const futurePayoutTaxRate = inputSheet.getRange('B53').getValue() / 100;

  // 払い出し最適化設定
  const retirementYears = inputSheet.getRange('B56').getValue();
  const shoukiboKyousai = inputSheet.getRange('B57').getValue();
  const hayashiKigyouDC = inputSheet.getRange('B58').getValue();
  const doiKigyouDC = inputSheet.getRange('B59').getValue();
  const linhKigyouDC = inputSheet.getRange('B60').getValue();

  // 土井郁子・役員設定
  const doiRetirementYears = inputSheet.getRange('B63').getValue();

  // 探索範囲
  const itakuhiMin = inputSheet.getRange('B32').getValue();
  const itakuhiMax = inputSheet.getRange('B33').getValue();
  const itakuhiStep = inputSheet.getRange('B34').getValue();

  const salaryMin = inputSheet.getRange('B37').getValue();
  const salaryMax = inputSheet.getRange('B38').getValue();
  const salaryStep = inputSheet.getRange('B39').getValue();

  // 売上バリエーション設定
  const revenueIncreaseMax = inputSheet.getRange('B42').getValue();
  const revenueIncreaseStep = inputSheet.getRange('B43').getValue();

  // Linhの給与を兵庫県最低賃金で固定（1,150円/時）
  // フルタイム想定：1,150円 × 8時間 × 22日 = 202,400円
  const linhSalaryFixed = 202400;

  const startTime = new Date();
  Logger.log('最適化を開始します - 目的：総資産の最大化、制約：P&I利益≥0 & 林事業所得≥0 & 土井給与≥35万円、Linh給与：兵庫県最低賃金で固定');
  Logger.log('内部留保の将来払出時想定税率：' + (futurePayoutTaxRate * 100) + '%');
  Logger.log('企業型DC・中退共の受取時想定税率：10%（退職所得控除適用）');
  Logger.log('売上バリエーション：+3000万円まで' + (revenueIncreaseStep / 10000).toLocaleString() + '万円刻み、+5000万円まで500万円刻み、+1億円まで1000万円刻み（最大+' + (revenueIncreaseMax / 10000).toLocaleString() + '万円）');
  Logger.log('土井郁子: 役員版と従業員版の両方を計算します');

  // 売上バリエーション（可変刻み幅）
  const revenueVariations = [];
  revenueVariations.push({label: 'ベース', amount: consultingRevenueBase});

  const increases = [];

  // +3000万円まで: パラメータ指定の刻み幅
  for (let i = revenueIncreaseStep; i <= Math.min(30000000, revenueIncreaseMax); i += revenueIncreaseStep) {
    increases.push(i);
  }

  // +3000万円～+5000万円: 500万円刻み
  if (revenueIncreaseMax > 30000000) {
    for (let i = 35000000; i <= Math.min(50000000, revenueIncreaseMax); i += 5000000) {
      increases.push(i);
    }
  }

  // +5000万円～+1億円: 1000万円刻み
  if (revenueIncreaseMax > 50000000) {
    for (let i = 60000000; i <= Math.min(100000000, revenueIncreaseMax); i += 10000000) {
      increases.push(i);
    }
  }

  for (let increase of increases) {
    revenueVariations.push({
      label: '+' + (increase / 10000).toLocaleString() + '万円',
      amount: consultingRevenueBase + increase
    });
  }

  // 役員版と従業員版の両方を計算
  const modesResults = {}; // {yakuin: [...], jugyoin: [...]}

  for (let doiIsYakuin of [0, 1]) {
    const modeName = doiIsYakuin ? '役員' : '従業員';
    Logger.log('');
    Logger.log('【' + modeName + 'モード】最適化開始');

    const allResults = [];
    let totalSearchCount = 0;

  // 各売上パターンで最適化（制約：P&I利益≥0 & 林事業所得≥0 & 土井給与≥35万円、Linh給与は最低賃金で固定）
  // 内部留保は将来払出時の税負担を考慮した実質価値で評価
  for (let variation of revenueVariations) {
    let bestWealth = -Infinity; // 総資産を最大化
    let bestParams = null;
    let searchCount = 0;
    
    // グリッドサーチ（Linhの給与は最低賃金で固定、土井は役員でも最適な報酬額を探索）
    for (let itakuhi = itakuhiMin; itakuhi <= itakuhiMax; itakuhi += itakuhiStep) {
      for (let h = salaryMin; h <= salaryMax; h += salaryStep) {
        for (let d = salaryMin; d <= salaryMax; d += salaryStep) {
          const l = linhSalaryFixed; // Linhの給与は固定
          searchCount++;
          totalSearchCount++;

          const result = calcTotalCost({
            hayashiYakuin: h,
            doiSalary: d,
            linhSalary: l,
            jimuItakuhi: itakuhi,
            veltraSalary: veltraSalary,
            kojinRevenue: kojinRevenue,
            kojinExpense: kojinExpense,
            haigusha: haigusha,
            aoiroDeduction: aoiroDeduction,
            housingLoanDeduction: housingLoanDeduction,
            doiChutaikyo: doiChutaikyo,
            linhChutaikyo: linhChutaikyo,
            consultingRevenue: variation.amount,
            officeExpense: officeExpense,
            otherExpense: otherExpense,
            kenpoRate: kenpoRate,
            kaigoRate: kaigoRate,
            nenkinRate: nenkinRate,
            koyoRateWorker: koyoRateWorker,
            koyoRateEmployer: koyoRateEmployer,
            shoukiboKyousai: shoukiboKyousai,
            hayashiKigyouDC: hayashiKigyouDC,
            doiKigyouDC: doiKigyouDC,
            linhKigyouDC: linhKigyouDC,
            doiIsYakuin: doiIsYakuin,
            doiRetirementYears: doiRetirementYears
          });

          // 制約条件1：P&Iの税引後利益がマイナスの場合はスキップ
          const piAfterTaxProfit = result.piIncome - result.piTax;
          if (piAfterTaxProfit < 0) {
            continue;
          }

          // 制約条件2：林の事業所得がマイナスの場合はスキップ
          if (result.kojinJigyoShotoku < 0) {
            continue;
          }

          // 制約条件3：土井の月次給与が35万円未満の場合はスキップ
          if (d < 350000) {
            continue;
          }

          // 総資産 = 3人の手取り + Linhのコスト + P&I課税所得の実質価値 + 企業型DC・中退共の実質価値
          const linhCost = result.linhCost; // Linhの企業型DC・中退共はlinhCostに含まれている
          // 退職金ルートでは法人税は損金算入により実質0円
          // P&I課税所得がそのまま退職金原資となり、個人税（退職所得税）のみ負担
          const piAfterTaxProfitRealValue = result.piIncome * (1 - futurePayoutTaxRate);
          // 実際の金額に基づく実効税率（出力用）
          const actualRetirementTaxRate = calcRetirementTaxRate(result.piIncome, retirementYears);
          // 企業型DC・中退共は将来退職所得として受け取る（退職所得控除適用、実効税率10%と想定）
          const dcChutaikyoTaxRate = 0.10;
          const dcChutaikyoValue = (result.hayashiKigyouDC + result.doiKigyouDC + result.doiChutaikyo) * (1 - dcChutaikyoTaxRate);
          const totalWealth = result.hayashiTedori + result.doiTedori + linhCost + result.haigushaTedori
                            + piAfterTaxProfitRealValue + dcChutaikyoValue;

          if (totalWealth > bestWealth) {
            bestWealth = totalWealth;
            bestParams = {
              hayashi: h,
              doi: d,
              linh: l,
              itakuhi: itakuhi,
              result: result,
              totalWealth: totalWealth,
              retirementTaxRate: actualRetirementTaxRate, // 実際の実効税率（出力用）
              piAfterTaxProfitRealValue: piAfterTaxProfitRealValue
            };
          }
        }
      }
    }

    // 有効な解が見つからなかった場合の処理
    if (bestParams === null) {
      Logger.log(variation.label + ' - 有効な解が見つかりませんでした（制約条件を満たす組み合わせなし）');
      continue; // この売上パターンはスキップ
    }

    allResults.push({
      label: variation.label,
      revenue: variation.amount,
      params: bestParams,
      searchCount: searchCount
    });

    Logger.log(variation.label + ' - 検索回数: ' + searchCount);
  }

  Logger.log('【' + modeName + 'モード】総検索回数: ' + totalSearchCount);

  // 有効な結果がない場合
  if (allResults.length === 0) {
    Logger.log('【' + modeName + 'モード】エラー: すべての売上パターンで有効な解が見つかりませんでした');
  } else {
    // 結果を保存
    modesResults[doiIsYakuin ? 'yakuin' : 'jugyoin'] = allResults;
    Logger.log('【' + modeName + 'モード】完了 - 有効な売上パターン数: ' + allResults.length);
  }
  } // doiIsYakuinループの終わり

  const endTime = new Date();
  const elapsedSeconds = Math.round((endTime - startTime) / 1000);

  Logger.log('');
  Logger.log('全体計算時間: ' + elapsedSeconds + '秒');

  // 両方のモードで有効な結果がない場合
  if (!modesResults.jugyoin && !modesResults.yakuin) {
    SpreadsheetApp.getUi().alert(
      '最適化エラー',
      'すべての売上パターンで制約条件を満たす解が見つかりませんでした。\n\n' +
      '以下を確認してください：\n' +
      '1. P&Iの売上が経費に対して十分か\n' +
      '2. 林の事業収入が経費に対して十分か\n' +
      '3. 土井の給与探索範囲が35万円以上を含んでいるか',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // 結果シートに出力
  if (modesResults.jugyoin) {
    outputResultsWithVariations(modesResults.jugyoin, futurePayoutTaxRate, retirementYears, '最適化結果（従業員）');
    Logger.log('従業員版を「最適化結果（従業員）」シートに出力しました');
  }

  if (modesResults.yakuin) {
    outputResultsWithVariations(modesResults.yakuin, futurePayoutTaxRate, retirementYears, '最適化結果（役員）');
    Logger.log('役員版を「最適化結果（役員）」シートに出力しました');
  }

  Logger.log('最適化が完了しました');
}

// ============================================================
// 結果の出力（売上バリエーション対応・横並び）
// ============================================================
function outputResultsWithVariations(allResults, futurePayoutTaxRate, retirementYears, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheetName = sheetName || '最適化結果'; // デフォルト値
  let resultSheet = ss.getSheetByName(sheetName);

  if (!resultSheet) {
    resultSheet = ss.insertSheet(sheetName);
  }
  resultSheet.clear();

  const numPatterns = allResults.length;
  const numCols = numPatterns + 1;

  // タイトル（シート名に応じて変更）
  const modeLabel = sheetName.includes('役員') ? '【土井：役員】' :
                    sheetName.includes('従業員') ? '【土井：従業員】' : '';
  const titleText = modeLabel + '最適化結果（総資産の最大化：3人の手取り+Linhコスト+P&I利益実質価値+DC・中退共実質価値）';
  const isDoiYakuin = sheetName.includes('役員');
  const doiLabel = isDoiYakuin ? '土井・役員報酬' : '土井・給与';

  // すべてのデータとスタイル情報を事前構築
  const allData = [];
  const allBackgrounds = [];
  const allFontWeights = [];
  const allNumberFormats = [];
  const allFontColors = [];

  // ヘッダー行（1行目）
  const headerRow = ['項目'];
  for (let pattern of allResults) {
    headerRow.push(pattern.label + '\n（' + (pattern.revenue / 10000).toLocaleString() + '万円）');
  }
  allData.push(headerRow);
  allBackgrounds.push(Array(numCols).fill('#e8f0fe'));
  allFontWeights.push(Array(numCols).fill('bold'));
  allNumberFormats.push(Array(numCols).fill('@'));
  allFontColors.push(Array(numCols).fill('#000000'));
  
  // データ行の構築
  const sections = [
    {
      title: '【最適な配分（月額）】',
      rows: [
        {label: '林・役員報酬', getValue: (p) => p.params.hayashi},
        {label: doiLabel, getValue: (p) => p.params.doi},
        {label: 'Linh・給与', getValue: (p) => p.params.linh}
      ]
    },
    {
      title: '【最適な配分（年額）】',
      rows: [
        {label: '林・事務委託費', getValue: (p) => p.params.itakuhi},
        {label: '林・役員報酬', getValue: (p) => p.params.hayashi * 12},
        {label: doiLabel, getValue: (p) => p.params.doi * 12},
        {label: 'Linh・給与', getValue: (p) => p.params.linh * 12}
      ]
    },
    {
      title: '【最適化結果】',
      rows: [
        {label: '★最大化された総資産', getValue: (p) => p.params.totalWealth, highlight: true},
        {label: '　林・土井・専従者の手取り', getValue: (p) => p.params.result.hayashiTedori + p.params.result.doiTedori + p.params.result.haigushaTedori},
        {label: '　Linhのコスト（給与+社保+中退共+DC）', getValue: (p) => p.params.result.linhCost},
        {label: '　P&I課税所得の実質価値（退職金ルート）', getValue: (p) => p.params.piAfterTaxProfitRealValue},
        {label: '　　将来期待される役員退職金（額面）', getValue: (p) => p.params.result.piIncome},
        {label: '　　退職金・実効税率', getValue: (p) => p.params.retirementTaxRate, isPercent: true},
        {label: '　　退職金・手取り（実質価値）', getValue: (p) => p.params.piAfterTaxProfitRealValue},
        {label: '　企業型DC・中退共（実質価値）', getValue: (p) => (p.params.result.hayashiKigyouDC + p.params.result.doiKigyouDC + p.params.result.doiChutaikyo) * 0.9},
        {label: '　　内訳：林・企業型DC', getValue: (p) => p.params.result.hayashiKigyouDC},
        {label: '　　内訳：土井・企業型DC', getValue: (p) => p.params.result.doiKigyouDC},
        {label: '　　内訳：土井・中退共', getValue: (p) => p.params.result.doiChutaikyo},
        {label: '　（参考）1人あたり手取り（当期）', getValue: (p) => (p.params.result.hayashiTedori + p.params.result.doiTedori + p.params.result.haigushaTedori + p.params.result.linhCost) / 2},
        {label: '　（参考）1人あたり手取り（DC・中退共含む）', getValue: (p) => (p.params.result.hayashiTedori + p.params.result.doiTedori + p.params.result.haigushaTedori + p.params.result.linhCost + (p.params.result.hayashiKigyouDC + p.params.result.doiKigyouDC + p.params.result.doiChutaikyo) * 0.9) / 2},
        {label: '　（参考）1人あたり手取り（全資産・林100%退職金）', getValue: (p) => (p.params.result.hayashiTedori + p.params.result.doiTedori + p.params.result.haigushaTedori + p.params.result.linhCost + (p.params.result.hayashiKigyouDC + p.params.result.doiKigyouDC + p.params.result.doiChutaikyo) * 0.9 + p.params.piAfterTaxProfitRealValue) / 2},
        {label: '　（参考）1人あたり手取り（全資産・土井100%退職金）', getValue: (p) => {
          if (!p.params.result.doiIsYakuin) return 'N/A';
          const piProfit = p.params.result.piIncome - p.params.result.piTax;
          const doiRetirementRate = calcRetirementTaxRate(piProfit, p.params.result.doiRetirementYears);
          const doiRetirementValue = piProfit * (1 - doiRetirementRate);
          return (p.params.result.hayashiTedori + p.params.result.doiTedori + p.params.result.haigushaTedori + p.params.result.linhCost + (p.params.result.hayashiKigyouDC + p.params.result.doiKigyouDC + p.params.result.doiChutaikyo) * 0.9 + doiRetirementValue) / 2;
        }},
        {label: '　（参考）1人あたり手取り（全資産・5:5分散退職金）', getValue: (p) => {
          if (!p.params.result.doiIsYakuin) return 'N/A';
          const piProfit = p.params.result.piIncome - p.params.result.piTax;
          const hayashiAmount = piProfit * 0.5;
          const doiAmount = piProfit * 0.5;
          const hayashiRate = calcRetirementTaxRate(hayashiAmount, retirementYears);
          const doiRate = calcRetirementTaxRate(doiAmount, p.params.result.doiRetirementYears);
          const dispersedValue = hayashiAmount * (1 - hayashiRate) + doiAmount * (1 - doiRate);
          return (p.params.result.hayashiTedori + p.params.result.doiTedori + p.params.result.haigushaTedori + p.params.result.linhCost + (p.params.result.hayashiKigyouDC + p.params.result.doiKigyouDC + p.params.result.doiChutaikyo) * 0.9 + dispersedValue) / 2;
        }}
      ]
    },
    {
      title: '【参考：コスト】',
      rows: [
        {label: '総コスト（社保+税金）', getValue: (p) => p.params.result.totalCost},
        {label: '　社会保険料合計', getValue: (p) => p.params.result.totalInsurance},
        {label: '　税金合計', getValue: (p) => p.params.result.totalTax}
      ]
    },
    {
      title: '【手取り・利益の内訳】',
      rows: [
        {label: '4人の手取り合計', getValue: (p) => p.params.result.totalTedori},
        {label: '　林佑樹・手取り', getValue: (p) => p.params.result.hayashiTedori},
        {label: '　土井郁子・手取り', getValue: (p) => p.params.result.doiTedori},
        {label: '　Linh・手取り', getValue: (p) => p.params.result.linhTedori},
        {label: '　専従者・手取り', getValue: (p) => p.params.result.haigushaTedori},
        {label: 'P&I課税所得（将来の退職金原資）', getValue: (p) => p.params.result.piIncome},
        {label: 'P&I課税所得の実質価値（退職金ルート）', getValue: (p) => p.params.piAfterTaxProfitRealValue},
        {label: '　林佑樹・等分配差分', getValue: (p) => (p.params.result.hayashiTedori + p.params.result.haigushaTedori) - (p.params.result.totalTedori / 2)},
        {label: '　土井郁子・等分配差分', getValue: (p) => p.params.result.doiTedori - (p.params.result.totalTedori / 2)},
        {label: 'Linhからの払い戻し後の残金', getValue: (p) => ((p.params.result.hayashiTedori + p.params.result.haigushaTedori) - (p.params.result.totalTedori / 2)) + (p.params.result.doiTedori - (p.params.result.totalTedori / 2)) + p.params.result.linhCost}
      ]
    },
    {
      title: '【P&I法人・収入と経費】',
      rows: [
        {label: '売上合計', getValue: (p) => p.params.result.piRevenue},
        {label: '　コンサル売上', getValue: (p) => p.revenue},
        {label: '　林・事務委託費', getValue: (p) => p.params.itakuhi},
        {label: '経費合計', getValue: (p) => p.params.result.piExpense},
        {label: '　林・役員報酬', getValue: (p) => p.params.hayashi * 12},
        {label: '　' + doiLabel, getValue: (p) => p.params.doi * 12},
        {label: '　Linh・給与', getValue: (p) => p.params.linh * 12},
        {label: '　社会保険料・会社負担', getValue: (p) => p.params.result.hayashiInsuranceEmployer + p.params.result.doiInsuranceEmployer + p.params.result.linhInsuranceEmployer},
        {label: '　中退共・土井郁子', getValue: (p) => p.params.result.doiChutaikyo},
        {label: '　中退共・Linh', getValue: (p) => p.params.result.linhChutaikyo},
        {label: '　企業型DC・林佑樹', getValue: (p) => p.params.result.hayashiKigyouDC},
        {label: '　企業型DC・土井郁子', getValue: (p) => p.params.result.doiKigyouDC},
        {label: '　企業型DC・Linh', getValue: (p) => p.params.result.linhKigyouDC},
        {label: '　オフィス賃料', getValue: (p) => p.params.result.officeExpense},
        {label: '　その他固定経費', getValue: (p) => p.params.result.otherExpense}
      ]
    },
    {
      title: '【P&I法人・利益と税金】',
      rows: [
        {label: '課税所得', getValue: (p) => p.params.result.piIncome},
        {label: '法人税率（実効）', getValue: (p) => p.params.result.piTaxRate, isPercent: true},
        {label: '法人税等', getValue: (p) => p.params.result.piTax},
        {label: '税引後利益（額面）', getValue: (p) => p.params.result.piIncome - p.params.result.piTax}
      ]
    },
    {
      title: '【払い出し最適化・退職金ルート vs 配当ルート】',
      rows: [
        {label: '想定勤続年数', getValue: (p) => retirementYears.toLocaleString() + '年'},
        {label: '', getValue: (p) => ''},
        {label: '【退職金ルート】', getValue: (p) => ''},
        {label: 'P&I課税所得（退職金原資）', getValue: (p) => p.params.result.piIncome},
        {label: '法人税', getValue: (p) => '損金算入により0円'},
        {label: '退職金・実効税率', getValue: (p) => p.params.retirementTaxRate, isPercent: true},
        {label: '退職金・個人税額', getValue: (p) => p.params.result.piIncome * p.params.retirementTaxRate},
        {label: '退職金・手取り', getValue: (p) => p.params.piAfterTaxProfitRealValue},
        {label: '総合実効税率', getValue: (p) => p.params.retirementTaxRate, isPercent: true},
        {label: '', getValue: (p) => ''},
        {label: '【配当ルート（参考）】', getValue: (p) => ''},
        {label: 'P&I課税所得', getValue: (p) => p.params.result.piIncome},
        {label: '法人税（23.2%）', getValue: (p) => p.params.result.piTax},
        {label: '税引後利益', getValue: (p) => p.params.result.piIncome - p.params.result.piTax},
        {label: '配当・個人税率', getValue: (p) => 0.20315, isPercent: true},
        {label: '配当・個人税額', getValue: (p) => (p.params.result.piIncome - p.params.result.piTax) * 0.20315},
        {label: '配当・手取り', getValue: (p) => (p.params.result.piIncome - p.params.result.piTax) * (1 - 0.20315)},
        {label: '総合実効税率', getValue: (p) => (p.params.result.piTax + (p.params.result.piIncome - p.params.result.piTax) * 0.20315) / p.params.result.piIncome, isPercent: true},
        {label: '', getValue: (p) => ''},
        {label: '退職金の優位性（手取り差額）', getValue: (p) => p.params.piAfterTaxProfitRealValue - (p.params.result.piIncome - p.params.result.piTax) * (1 - 0.20315)},
        {label: '（参考）最適化用想定税率', getValue: (p) => futurePayoutTaxRate, isPercent: true},
        {label: '', getValue: (p) => ''},
        {label: '小規模企業共済・林掛金（年額）', getValue: (p) => p.params.result.shoukiboKyousai},
        {label: '小規模企業共済・節税効果', getValue: (p) => p.params.result.shoukiboKyousai * (p.params.result.hayashiIncomeTaxRate * 1.021 + 0.1)},
        {label: '企業型DC・林掛金（年額）', getValue: (p) => p.params.result.hayashiKigyouDC},
        {label: '企業型DC・土井掛金（年額）', getValue: (p) => p.params.result.doiKigyouDC},
        {label: '企業型DC・Linh掛金（年額）', getValue: (p) => p.params.result.linhKigyouDC}
      ]
    },
    {
      title: '【林佑樹・給与所得】',
      rows: [
        {label: 'P&I役員報酬（年額）', getValue: (p) => p.params.hayashi * 12},
        {label: '給与所得控除後', getValue: (p) => p.params.result.piKyuyoShotoku}
      ]
    },
    {
      title: '【林佑樹・事業所得】',
      rows: [
        {label: '事業収入合計', getValue: (p) => p.params.result.kojinTotalRevenue},
        {label: '　Veltra業務委託', getValue: (p) => p.params.result.veltraSalary},
        {label: '　個人コンサル収入', getValue: (p) => p.params.result.kojinRevenue},
        {label: '事業経費合計', getValue: (p) => p.params.result.kojinTotalExpense},
        {label: '　固定経費', getValue: (p) => p.params.result.kojinExpense},
        {label: '　専従者給与', getValue: (p) => p.params.result.haigusha},
        {label: '　林・事務委託費', getValue: (p) => p.params.itakuhi},
        {label: '　青色申告控除', getValue: (p) => p.params.result.aoiroDeduction},
        {label: '事業所得', getValue: (p) => p.params.result.kojinJigyoShotoku}
      ]
    },
    {
      title: '【林佑樹・税金と手取り】',
      rows: [
        {label: '総所得金額', getValue: (p) => p.params.result.hayashiTotalIncome},
        {label: '課税所得', getValue: (p) => p.params.result.hayashiTaxableIncome},
        {label: '所得税率', getValue: (p) => p.params.result.hayashiIncomeTaxRate, isPercent: true},
        {label: '所得税', getValue: (p) => p.params.result.hayashiIncomeTax},
        {label: '住民税（神戸市）', getValue: (p) => p.params.result.hayashiResidentTax},
        {label: '個人事業税', getValue: (p) => p.params.result.hayashiJigyoTax},
        {label: '標準報酬月額', getValue: (p) => getStandardSalary(p.params.hayashi).toLocaleString() + '円'},
        {label: '社会保険料（年額）', getValue: (p) => p.params.result.hayashiInsurance},
        {label: '　本人負担', getValue: (p) => p.params.result.hayashiInsuranceWorker},
        {label: '　会社負担', getValue: (p) => p.params.result.hayashiInsuranceEmployer},
        {label: '　実質負担率', getValue: (p) => p.params.result.hayashiInsurance / (p.params.hayashi * 12), isPercent: true},
        {label: '小規模企業共済（年額・所得控除）', getValue: (p) => p.params.result.shoukiboKyousai},
        {label: '企業型DC（年額・所得控除）', getValue: (p) => p.params.result.hayashiKigyouDC},
        {label: 'ふるさと納税上限額', getValue: (p) => p.params.result.hayashiFurusatoLimit},
        {label: '手取り', getValue: (p) => p.params.result.hayashiTedori}
      ]
    },
    {
      title: '【土井郁子】',
      rows: [
        {label: '役員/従業員', getValue: (p) => p.params.result.doiIsYakuin ? '役員' : '従業員'},
        {label: (p) => p.params.result.doiIsYakuin ? 'P&I役員報酬（年額）' : '給与収入（年額）', getValue: (p) => p.params.doi * 12},
        {label: '課税所得', getValue: (p) => p.params.doi * 12 - calcSalaryDeduction(p.params.doi * 12) - 480000 - p.params.result.doiInsuranceWorker},
        {label: '所得税率', getValue: (p) => p.params.result.doiIncomeTaxRate, isPercent: true},
        {label: '所得税', getValue: (p) => p.params.result.doiIncomeTax},
        {label: '住民税（西宮市）', getValue: (p) => p.params.result.doiResidentTax},
        {label: '標準報酬月額', getValue: (p) => getStandardSalary(p.params.doi).toLocaleString() + '円'},
        {label: '社会保険料（年額）', getValue: (p) => p.params.result.doiInsurance},
        {label: '　本人負担', getValue: (p) => p.params.result.doiInsuranceWorker},
        {label: '　会社負担', getValue: (p) => p.params.result.doiInsuranceEmployer},
        {label: '　実質負担率', getValue: (p) => p.params.result.doiInsurance / (p.params.doi * 12), isPercent: true},
        {label: '中退共・掛け金（年額・会社負担）', getValue: (p) => p.params.result.doiIsYakuin ? 0 : p.params.result.doiChutaikyo},
        {label: '企業型DC・掛け金（年額・会社負担）', getValue: (p) => p.params.result.doiKigyouDC},
        {label: 'ふるさと納税上限額', getValue: (p) => p.params.result.doiFurusatoLimit},
        {label: '手取り', getValue: (p) => p.params.result.doiTedori}
      ]
    },
    {
      title: '【Linh】',
      rows: [
        {label: '給与収入', getValue: (p) => p.params.linh * 12},
        {label: '配偶者控除（所得税）', getValue: (p) => p.params.result.linhHaigushaKoujoIncomeTax},
        {label: '配偶者控除（住民税）', getValue: (p) => p.params.result.linhHaigushaKoujoResidentTax},
        {label: '課税所得（所得税）', getValue: (p) => p.params.result.linhTaxableIncomeForIncomeTax},
        {label: '課税所得（住民税）', getValue: (p) => p.params.result.linhTaxableIncomeForResidentTax},
        {label: '所得税率', getValue: (p) => p.params.result.linhIncomeTaxRate, isPercent: true},
        {label: '所得税', getValue: (p) => p.params.result.linhIncomeTax},
        {label: '住民税（神戸市）', getValue: (p) => p.params.result.linhResidentTax},
        {label: '標準報酬月額', getValue: (p) => getStandardSalary(p.params.linh).toLocaleString() + '円'},
        {label: '社会保険料（年額）', getValue: (p) => p.params.result.linhInsurance},
        {label: '　本人負担', getValue: (p) => p.params.result.linhInsuranceWorker},
        {label: '　会社負担', getValue: (p) => p.params.result.linhInsuranceEmployer},
        {label: '　実質負担率', getValue: (p) => p.params.result.linhInsurance / (p.params.linh * 12), isPercent: true},
        {label: '中退共・掛け金（年額・会社負担）', getValue: (p) => p.params.result.linhChutaikyo},
        {label: '企業型DC・掛け金（年額・会社負担）', getValue: (p) => p.params.result.linhKigyouDC},
        {label: 'ふるさと納税上限額', getValue: (p) => p.params.result.linhFurusatoLimit},
        {label: '手取り', getValue: (p) => p.params.result.linhTedori}
      ]
    },
    {
      title: '【専従者】',
      rows: [
        {label: '給与収入', getValue: (p) => p.params.result.haigusha},
        {label: '課税所得', getValue: (p) => Math.max(0, p.params.result.haigusha - calcSalaryDeduction(p.params.result.haigusha) - 480000)},
        {label: '所得税率', getValue: (p) => p.params.result.haigushaIncomeTaxRate, isPercent: true},
        {label: '所得税', getValue: (p) => p.params.result.haigushaIncomeTax},
        {label: '住民税（神戸市）', getValue: (p) => p.params.result.haigushaResidentTax},
        {label: 'ふるさと納税上限額', getValue: (p) => p.params.result.haigushaFurusatoLimit},
        {label: '手取り', getValue: (p) => p.params.result.haigushaTedori}
      ]
    },
    {
      title: '【役員退職金シミュレーション（P&I課税所得の分配）】',
      rows: [
        {label: 'P&I課税所得（退職金原資）', getValue: (p) => p.params.result.piIncome},
        {label: '林・想定勤続年数', getValue: (p) => retirementYears.toLocaleString() + '年'},
        {label: '土井・想定勤続年数', getValue: (p) => p.params.result.doiIsYakuin ? p.params.result.doiRetirementYears.toLocaleString() + '年' : 'N/A（従業員）'},
        {label: '', getValue: (p) => ''},

        {label: '【分配パターン1：林100%】', getValue: (p) => ''},
        {label: '林・退職金額', getValue: (p) => p.params.result.piIncome},
        {label: '林・実効税率', getValue: (p) => p.params.retirementTaxRate, isPercent: true},
        {label: '林・手取り', getValue: (p) => p.params.piAfterTaxProfitRealValue},
        {label: '（参考）配当での手取り', getValue: (p) => (p.params.result.piIncome - p.params.result.piTax) * (1 - 0.20315)},
        {label: '退職金の優位性（配当比）', getValue: (p) => p.params.piAfterTaxProfitRealValue - (p.params.result.piIncome - p.params.result.piTax) * (1 - 0.20315)},
        {label: '', getValue: (p) => ''},

        {label: '【分配パターン2：土井100%】', getValue: (p) => p.params.result.doiIsYakuin ? '' : 'N/A'},
        {label: '土井・退職金額', getValue: (p) => p.params.result.doiIsYakuin ? p.params.result.piIncome : 'N/A'},
        {label: '土井・実効税率', getValue: (p) => {
          if (!p.params.result.doiIsYakuin) return 'N/A';
          const amount = p.params.result.piIncome;
          return calcRetirementTaxRate(amount, p.params.result.doiRetirementYears);
        }, isPercent: true},
        {label: '土井・手取り', getValue: (p) => {
          if (!p.params.result.doiIsYakuin) return 'N/A';
          const amount = p.params.result.piIncome;
          const rate = calcRetirementTaxRate(amount, p.params.result.doiRetirementYears);
          return amount * (1 - rate);
        }},
        {label: '（参考）配当での手取り', getValue: (p) => p.params.result.doiIsYakuin ? (p.params.result.piIncome - p.params.result.piTax) * (1 - 0.20315) : 'N/A'},
        {label: '退職金の優位性（配当比）', getValue: (p) => {
          if (!p.params.result.doiIsYakuin) return 'N/A';
          const amount = p.params.result.piIncome;
          const rate = calcRetirementTaxRate(amount, p.params.result.doiRetirementYears);
          const retirementTedori = amount * (1 - rate);
          const dividendTedori = (p.params.result.piIncome - p.params.result.piTax) * (1 - 0.20315);
          return retirementTedori - dividendTedori;
        }},
        {label: '雇用保険料削減（年額）', getValue: (p) => {
          if (!p.params.result.doiIsYakuin) return 'N/A';
          return p.params.doi * 12 * (0.006 + 0.0095);
        }},
        {label: '雇用保険料削減（30年累計）', getValue: (p) => {
          if (!p.params.result.doiIsYakuin) return 'N/A';
          return p.params.doi * 12 * (0.006 + 0.0095) * 30;
        }},
        {label: '', getValue: (p) => ''},

        {label: '【分配パターン3：5:5均等分散】', getValue: (p) => p.params.result.doiIsYakuin ? '' : 'N/A'},
        {label: '林・退職金額', getValue: (p) => p.params.result.piIncome * 0.5},
        {label: '林・実効税率', getValue: (p) => {
          const amount = p.params.result.piIncome * 0.5;
          return calcRetirementTaxRate(amount, retirementYears);
        }, isPercent: true},
        {label: '林・手取り', getValue: (p) => {
          const amount = p.params.result.piIncome * 0.5;
          const rate = calcRetirementTaxRate(amount, retirementYears);
          return amount * (1 - rate);
        }},
        {label: '土井・退職金額', getValue: (p) => p.params.result.doiIsYakuin ? p.params.result.piIncome * 0.5 : 'N/A'},
        {label: '土井・実効税率', getValue: (p) => {
          if (!p.params.result.doiIsYakuin) return 'N/A';
          const amount = p.params.result.piIncome * 0.5;
          return calcRetirementTaxRate(amount, p.params.result.doiRetirementYears);
        }, isPercent: true},
        {label: '土井・手取り', getValue: (p) => {
          if (!p.params.result.doiIsYakuin) return 'N/A';
          const amount = p.params.result.piIncome * 0.5;
          const rate = calcRetirementTaxRate(amount, p.params.result.doiRetirementYears);
          return amount * (1 - rate);
        }},
        {label: '合計手取り', getValue: (p) => {
          if (!p.params.result.doiIsYakuin) return 'N/A';
          const amount = p.params.result.piIncome * 0.5;
          const hayashiRate = calcRetirementTaxRate(amount, retirementYears);
          const doiRate = calcRetirementTaxRate(amount, p.params.result.doiRetirementYears);
          return amount * (2 - hayashiRate - doiRate);
        }},
        {label: '分散による追加節税額', getValue: (p) => {
          if (!p.params.result.doiIsYakuin) return 'N/A';
          const amount = p.params.result.piIncome * 0.5;
          const hayashiRate = calcRetirementTaxRate(amount, retirementYears);
          const doiRate = calcRetirementTaxRate(amount, p.params.result.doiRetirementYears);
          const totalTedori = amount * (2 - hayashiRate - doiRate);
          return totalTedori - p.params.piAfterTaxProfitRealValue;
        }},
      ]
    }
  ];

  // データを構築（配列として）
  for (let section of sections) {
    // セクションタイトル行
    const sectionTitleRow = [section.title];
    for (let i = 0; i < numPatterns; i++) sectionTitleRow.push('');
    allData.push(sectionTitleRow);
    allBackgrounds.push(Array(numCols).fill('#e8f0fe'));
    allFontWeights.push(Array(numCols).fill('bold'));
    allNumberFormats.push(Array(numCols).fill('@'));
    allFontColors.push(Array(numCols).fill('#000000'));

    // データ行
    for (let rowDef of section.rows) {
      const label = typeof rowDef.label === 'function' ? rowDef.label(allResults[0]) : rowDef.label;
      const dataRow = [label];
      for (let pattern of allResults) {
        const value = rowDef.getValue(pattern);
        dataRow.push(value);
      }
      allData.push(dataRow);

      // 背景色・太字の設定
      const rowBg = rowDef.highlight ? Array(numCols).fill('#fff9c4') : Array(numCols).fill('#ffffff');
      allBackgrounds.push(rowBg);
      const rowWeight = rowDef.highlight ? Array(numCols).fill('bold') : Array(numCols).fill('normal');
      allFontWeights.push(rowWeight);

      // 数値フォーマットと色
      const rowFormats = ['@']; // 最初の列（ラベル）はテキスト
      const rowColors = ['#000000'];

      for (let i = 0; i < numPatterns; i++) {
        const value = dataRow[i + 1];
        if (typeof value === 'number') {
          if (rowDef.isPercent) {
            rowFormats.push('0.00%');
          } else {
            rowFormats.push('¥#,##0;[Red]-¥#,##0');
          }
          rowColors.push('#000000');
        } else if (typeof value === 'string') {
          const numValue = parseFloat(value);
          if (!isNaN(numValue) && numValue < 0) {
            rowColors.push('#ff0000');
          } else {
            rowColors.push('#000000');
          }
          rowFormats.push('@');
        } else {
          rowFormats.push('@');
          rowColors.push('#000000');
        }
      }
      allNumberFormats.push(rowFormats);
      allFontColors.push(rowColors);
    }

    // セクション間の空白行
    allData.push(Array(numCols).fill(''));
    allBackgrounds.push(Array(numCols).fill('#ffffff'));
    allFontWeights.push(Array(numCols).fill('normal'));
    allNumberFormats.push(Array(numCols).fill('@'));
    allFontColors.push(Array(numCols).fill('#000000'));
  }

  // ============ 一括書き込み（高速化） ============
  const numRows = allData.length;

  // 1. データを一括書き込み
  resultSheet.getRange(1, 1, numRows, numCols).setValues(allData);

  // 2. 背景色を一括適用
  resultSheet.getRange(1, 1, numRows, numCols).setBackgrounds(allBackgrounds);

  // 3. フォントウェイトを一括適用
  resultSheet.getRange(1, 1, numRows, numCols).setFontWeights(allFontWeights);

  // 4. 数値フォーマットを一括適用
  resultSheet.getRange(1, 1, numRows, numCols).setNumberFormats(allNumberFormats);

  // 5. フォントカラーを一括適用
  resultSheet.getRange(1, 1, numRows, numCols).setFontColors(allFontColors);

  // 6. ヘッダー行のスタイル
  resultSheet.getRange(1, 1, 1, numCols)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // 7. 列幅設定
  resultSheet.setColumnWidth(1, 250);
  for (let col = 2; col <= numCols; col++) {
    resultSheet.setColumnWidth(col, 130);
  }

  // 8. 数値列の右揃え
  if (numRows > 1) {
    resultSheet.getRange(2, 2, numRows - 1, numPatterns).setHorizontalAlignment('right');
  }

  // 9. ヘッダー行の高さ調整
  resultSheet.setRowHeight(1, 50);

  // 10. 1行目と1列目をフリーズ（スクロール時に固定表示）
  resultSheet.setFrozenRows(1);
  resultSheet.setFrozenColumns(1);
}