<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>TOYOTA 顧客滿意度數據分析與主管月報平台</title>
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- SheetJS CDN (Excel 解析與匯出) -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- Chart.js CDN (圖表繪製) -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <!-- Lucide Icons -->
  <script src="https://unpkg.com/lucide@latest"></script>
  <style>
    @media print {
      body * { visibility: hidden; }
      #printable-report, #printable-report * { visibility: visible; }
      #printable-report { position: absolute; left: 0; top: 0; width: 100%; margin: 0; padding: 20px; }
      .no-print { display: none !important; }
      @page { size: A4 landscape; margin: 10mm; }
    }
  </style>
</head>
<body class="bg-slate-50 text-slate-800 font-sans antialiased min-h-screen flex flex-col">

  <!-- 頂部導覽列 -->
  <header class="bg-[#0f2444] text-white border-b border-slate-700 shadow-md sticky top-0 z-40 no-print">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
      <div class="flex items-center space-x-3">
        <div class="w-9 h-9 rounded-lg bg-[#EB0A1E] flex items-center justify-center font-black text-white text-base shadow">
          T
        </div>
        <div>
          <h1 class="text-base font-bold tracking-tight text-white flex items-center gap-2">
            TOYOTA 顧客滿意度分析與月報平台
            <span class="text-xs font-normal text-slate-300 bg-slate-800/80 px-2 py-0.5 rounded border border-slate-600">電服中心專用</span>
          </h1>
          <p class="text-xs text-slate-400">售後保修 (CSI) ＆ 購車新車 (SSI) 跨月／去年同期分析系統</p>
        </div>
      </div>

      <!-- 雙軌切換器 (SSI / CSI) -->
      <div class="flex items-center bg-slate-800/90 p-1 rounded-lg border border-slate-600">
        <button id="btn-mode-csi" onclick="switchTrackMode('CSI')" class="flex items-center space-x-1.5 px-3 py-1.5 rounded-md text-xs font-bold transition-all text-slate-300 hover:text-white">
          <i data-lucide="wrench" class="w-3.5 h-3.5"></i>
          <span>保修滿意度 (CSI)</span>
        </button>
        <button id="btn-mode-ssi" onclick="switchTrackMode('SSI')" class="flex items-center space-x-1.5 px-3 py-1.5 rounded-md text-xs font-bold transition-all shadow bg-[#EB0A1E] text-white">
          <i data-lucide="car" class="w-3.5 h-3.5"></i>
          <span>購車滿意度 (SSI)</span>
        </button>
      </div>
    </div>
  </header>

  <!-- 主版面結構 -->
  <div class="flex-1 flex max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 gap-6">
    
    <!-- 側邊選單 -->
    <aside class="w-56 shrink-0 no-print">
      <nav class="space-y-1 bg-white p-3 rounded-xl border border-slate-200 shadow-sm sticky top-24">
        <div class="text-[10px] font-bold text-slate-400 uppercase tracking-wider px-3 py-1.5">分析功能</div>
        <button onclick="navigateView('dashboard')" id="nav-dashboard" class="w-full flex items-center space-x-2.5 px-3 py-2 text-xs font-semibold rounded-lg bg-red-50 text-[#EB0A1E] transition">
          <i data-lucide="layout-dashboard" class="w-4 h-4"></i>
          <span>首頁儀表板</span>
        </button>
        <button onclick="navigateView('checklist-trends')" id="nav-checklist-trends" class="w-full flex items-center space-x-2.5 px-3 py-2 text-xs font-semibold rounded-lg text-slate-600 hover:bg-slate-100 transition">
          <i data-lucide="check-square" class="w-4 h-4"></i>
          <span>檢核題推移分析</span>
        </button>
        <button onclick="navigateView('import')" id="nav-import" class="w-full flex items-center space-x-2.5 px-3 py-2 text-xs font-semibold rounded-lg text-slate-600 hover:bg-slate-100 transition">
          <i data-lucide="file-up" class="w-4 h-4"></i>
          <span>資料匯入與清除</span>
        </button>
        <button onclick="navigateView('questions')" id="nav-questions" class="w-full flex items-center space-x-2.5 px-3 py-2 text-xs font-semibold rounded-lg text-slate-600 hover:bg-slate-100 transition">
          <i data-lucide="list-checks" class="w-4 h-4"></i>
          <span>題目管理與設定</span>
        </button>
        <button onclick="navigateView('item-analysis')" id="nav-item-analysis" class="w-full flex items-center space-x-2.5 px-3 py-2 text-xs font-semibold rounded-lg text-slate-600 hover:bg-slate-100 transition">
          <i data-lucide="bar-chart-2" class="w-4 h-4"></i>
          <span>個別題目分析</span>
        </button>
        <button onclick="navigateView('branch-compare')" id="nav-branch-compare" class="w-full flex items-center space-x-2.5 px-3 py-2 text-xs font-semibold rounded-lg text-slate-600 hover:bg-slate-100 transition">
          <i data-lucide="git-compare" class="w-4 h-4"></i>
          <span id="label-nav-branch">營業所比較</span>
        </button>
        <button onclick="navigateView('dimensions')" id="nav-dimensions" class="w-full flex items-center justify-between px-3 py-2 text-xs font-semibold rounded-lg text-slate-400 hover:bg-slate-50 transition">
          <div class="flex items-center space-x-2.5">
            <i data-lucide="layers" class="w-4 h-4"></i>
            <span>構面分析</span>
          </div>
          <span class="text-[9px] bg-slate-100 text-slate-500 px-1.5 py-0.5 rounded border border-slate-200">擴充保留</span>
        </button>

        <div class="text-[10px] font-bold text-slate-400 uppercase tracking-wider px-3 pt-3 py-1.5">報告與管理</div>
        <button onclick="navigateView('report')" id="nav-report" class="w-full flex items-center space-x-2.5 px-3 py-2 text-xs font-semibold rounded-lg text-slate-600 hover:bg-slate-100 transition">
          <i data-lucide="file-text" class="w-4 h-4"></i>
          <span>主管月報產生器</span>
        </button>
        <button onclick="navigateView('settings')" id="nav-settings" class="w-full flex items-center space-x-2.5 px-3 py-2 text-xs font-semibold rounded-lg text-slate-600 hover:bg-slate-100 transition">
          <i data-lucide="settings" class="w-4 h-4"></i>
          <span>系統備份與資安</span>
        </button>
      </nav>
    </aside>

    <!-- 主要內容區塊 -->
    <main class="flex-1 min-w-0">
      
      <!-- 全域篩選與比較基準列 -->
      <div id="filter-bar" class="bg-white p-3.5 rounded-xl border border-slate-200 shadow-sm mb-5 flex flex-wrap items-center justify-between gap-3 text-xs no-print">
        <div class="flex flex-wrap items-center gap-3">
          <div class="flex items-center space-x-1.5">
            <span class="font-bold text-slate-500">檢視月份：</span>
            <select id="filter-month" onchange="renderActiveView()" class="border border-slate-300 rounded px-2.5 py-1.5 bg-slate-50 font-medium text-slate-700 outline-none focus:border-red-500">
            </select>
          </div>

          <div class="flex items-center space-x-1.5">
            <span class="font-bold text-slate-500" id="label-filter-branch">據點代碼：</span>
            <select id="filter-branch" onchange="renderActiveView()" class="border border-slate-300 rounded px-2.5 py-1.5 bg-slate-50 font-medium text-slate-700 outline-none focus:border-red-500">
            </select>
          </div>

          <div class="flex items-center space-x-1.5">
            <span class="font-bold text-slate-500">調查來源：</span>
            <select id="filter-channel" onchange="renderActiveView()" class="border border-slate-300 rounded px-2.5 py-1.5 bg-slate-50 font-medium text-slate-700 outline-none focus:border-red-500">
              <option value="ALL">全部 (問卷 ＋ 電訪)</option>
              <option value="問卷">問卷</option>
              <option value="電訪">電訪</option>
            </select>
          </div>

          <div class="flex items-center space-x-1.5 pl-2 border-l border-slate-200">
            <span class="font-bold text-slate-500">差異比較基準：</span>
            <select id="filter-compare-mode" onchange="renderActiveView()" class="border border-slate-300 rounded px-2.5 py-1.5 bg-slate-50 font-semibold text-slate-700 outline-none focus:border-red-500">
              <option value="MoM">較上月 (MoM)</option>
              <option value="YoY">較去年同期 (YoY)</option>
            </select>
          </div>
        </div>

        <div class="text-[11px] text-slate-500 flex items-center gap-1.5">
          <span class="inline-block w-2 h-2 rounded-full bg-emerald-500"></span>
          本機資料庫有效成功：<strong id="stat-record-count" class="text-slate-800">0</strong> 筆
        </div>
      </div>

      <!-- VIEW 1: 首頁儀表板 -->
      <section id="view-dashboard" class="space-y-5">
        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3.5">
          <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
            <div class="text-xs font-bold text-slate-500">整體加權滿意度</div>
            <div class="text-2xl font-black text-slate-900 mt-1" id="kpi-avg-score">--</div>
            <div class="text-[11px] mt-1 font-semibold" id="kpi-avg-diff">--</div>
          </div>
          <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
            <div class="text-xs font-bold text-slate-500">有效成功填答數</div>
            <div class="text-2xl font-black text-slate-900 mt-1" id="kpi-total-resp">--</div>
            <div class="text-[11px] text-slate-400 mt-1">已排除失敗與無效跳題</div>
          </div>
          <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
            <div class="text-xs font-bold text-slate-500">高滿意度比例 (9-10分)</div>
            <div class="text-2xl font-black text-emerald-600 mt-1" id="kpi-high-rate">--</div>
            <div class="text-[11px] mt-1 font-semibold" id="kpi-high-diff">--</div>
          </div>
          <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
            <div class="text-xs font-bold text-slate-500">正向滿意度比例 (8-10分)</div>
            <div class="text-2xl font-black text-blue-600 mt-1" id="kpi-pos-rate">--</div>
            <div class="text-[11px] mt-1 font-semibold" id="kpi-pos-diff">--</div>
          </div>
          <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
            <div class="text-xs font-bold text-slate-500">低分警戒率 (1-6分)</div>
            <div class="text-2xl font-black text-rose-600 mt-1" id="kpi-low-rate">--</div>
            <div class="text-[11px] mt-1 font-semibold" id="kpi-low-diff">--</div>
          </div>
        </div>

        <!-- 首頁檢核題達成率 KPI 矩陣卡片 (S2, S3, S5, S6) -->
        <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
          <div class="flex items-center justify-between mb-3">
            <h3 class="text-xs font-bold text-slate-800 flex items-center gap-1.5 uppercase tracking-wider">
              <i data-lucide="check-check" class="w-4 h-4 text-emerald-600"></i>
              <span>標準服務流程落實率 (檢核題達成率 KPI 概況)</span>
            </h3>
            <span class="text-[11px] text-slate-400">分母為純有效填答，已排除「不清楚/不知道/不記得」及 0</span>
          </div>
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 text-center" id="dashboard-checklist-grid">
          </div>
        </div>

        <!-- 首頁主要圖表區 -->
        <div class="grid grid-cols-1 lg:grid-cols-2 gap-5">
          <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
            <div class="flex items-center justify-between mb-4">
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <i data-lucide="award" class="w-4 h-4 text-amber-500"></i>
                <span id="title-branch-ranking">據點滿意度綜合排行 (當月)</span>
              </h3>
              <span class="text-[11px] text-slate-400">採純有效填答加權計算</span>
            </div>
            <div class="h-64 relative">
              <canvas id="chart-branch-ranking"></canvas>
            </div>
          </div>

          <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
            <div class="flex items-center justify-between mb-4">
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <i data-lucide="sliders" class="w-4 h-4 text-blue-500"></i>
                <span>各評分題分佈比例 (100% 堆疊)</span>
              </h3>
              <div class="flex items-center space-x-2 text-[10px]">
                <span class="flex items-center gap-1"><span class="w-2.5 h-2.5 bg-emerald-500 rounded-sm"></span>9-10分</span>
                <span class="flex items-center gap-1"><span class="w-2.5 h-2.5 bg-sky-400 rounded-sm"></span>7-8分</span>
                <span class="flex items-center gap-1"><span class="w-2.5 h-2.5 bg-rose-500 rounded-sm"></span>1-6分</span>
              </div>
            </div>
            <div class="h-64 relative">
              <canvas id="chart-question-stack"></canvas>
            </div>
          </div>
        </div>

        <!-- 連續月份走勢圖 -->
        <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
          <div class="flex items-center justify-between mb-3">
            <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
              <i data-lucide="trending-up" class="w-4 h-4 text-emerald-600"></i>
              <span>歷史滿意度連續月度走勢與樣本數</span>
            </h3>
            <span class="text-[11px] text-slate-400">平均得分 (左軸) / 成功樣本數 (右軸)</span>
          </div>
          <div class="h-56 relative">
            <canvas id="chart-trend"></canvas>
          </div>
        </div>
      </section>

      <!-- VIEW: 檢核題推移分析 -->
      <section id="view-checklist-trends" class="space-y-5 hidden">
        <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-4">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <i data-lucide="line-chart" class="w-4 h-4 text-indigo-600"></i>
                <span>檢核題達成率跨月份時間推移分析</span>
              </h3>
              <p class="text-xs text-slate-500 mt-0.5">追蹤各項服務流程（保險說明、大禮包、服務廠介紹、主管介紹）在各月份的執行落實率變化。</p>
            </div>
          </div>
          <div class="h-64 relative">
            <canvas id="chart-checklist-month-trend"></canvas>
          </div>
        </div>

        <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-4">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <i data-lucide="bar-chart-3" class="w-4 h-4 text-emerald-600"></i>
                <span id="title-checklist-branch-header">各據點檢核題落實達成率比較 (當月)</span>
              </h3>
              <p class="text-xs text-slate-500 mt-0.5">橫向比對所有據點（自然排序），快速找出未落實特定流程之單位。</p>
            </div>
            <button onclick="exportChecklistExcel()" class="px-3 py-1.5 border border-slate-300 hover:bg-slate-50 rounded text-xs font-semibold text-slate-700 flex items-center gap-1.5">
              <i data-lucide="sheet" class="w-3.5 h-3.5"></i>
              匯出檢核題分析表 (Excel)
            </button>
          </div>
          <div class="overflow-x-auto">
            <table class="w-full text-left text-xs border-collapse" id="checklist-branch-table">
              <thead>
                <tr class="bg-slate-100 text-slate-700 border-b border-slate-200 font-bold" id="checklist-table-header">
                </tr>
              </thead>
              <tbody id="checklist-table-body" class="divide-y divide-slate-100">
              </tbody>
            </table>
          </div>
        </div>
      </section>

      <!-- VIEW 2: 資料匯入與清除管理 -->
      <section id="view-import" class="space-y-5 hidden">
        <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <i data-lucide="upload-cloud" class="w-4 h-4 text-[#EB0A1E]"></i>
                <span>滿意度原始調查 Excel／CSV 檔案匯入</span>
              </h3>
              <p class="text-xs text-slate-500 mt-0.5">
                當前目標：<strong id="import-target-badge" class="text-[#EB0A1E]">購車滿意度 (SSI)</strong>（首欄支援「編號」，「不清楚/不知道/不記得」自動列入無效）
              </p>
            </div>
            <button onclick="downloadSampleTemplate()" class="px-3 py-1.5 border border-slate-300 hover:bg-slate-50 rounded text-xs font-semibold text-slate-700 flex items-center gap-1.5">
              <i data-lucide="download" class="w-3.5 h-3.5"></i>
              下載標準範例檔 (含編號第一欄)
            </button>
          </div>

          <div id="drop-zone" class="mt-4 border-2 border-dashed border-slate-300 hover:border-red-400 bg-slate-50/60 rounded-xl p-8 text-center cursor-pointer transition">
            <input type="file" id="file-uploader" accept=".xlsx, .xls, .csv" class="hidden" onchange="handleFileSelect(event)">
            <i data-lucide="file-spreadsheet" class="w-10 h-10 text-slate-400 mx-auto mb-2"></i>
            <p class="text-xs font-bold text-slate-700">點擊此處或將 Excel / CSV 檔案拖曳至此</p>
            <p class="text-[11px] text-slate-400 mt-1">支援跨月份、跨年度資料批次匯入</p>
          </div>

          <div id="import-preview-box" class="mt-5 space-y-4 hidden">
            <div class="flex items-center justify-between bg-slate-100 p-3 rounded-lg text-xs">
              <div class="space-x-3">
                <span>檔案名稱：<strong id="preview-filename">--</strong></span>
                <span>總讀取筆數：<strong id="preview-rowcount">0</strong> 筆</span>
              </div>
              <div class="flex items-center space-x-2">
                <span>工作表：</span>
                <select id="sheet-selector" onchange="changeSheet()" class="border border-slate-300 rounded px-2 py-1 bg-white text-xs"></select>
              </div>
            </div>

            <div class="border border-slate-200 rounded-lg p-3.5 bg-slate-50">
              <h4 class="text-xs font-bold text-slate-700 mb-2.5">系統核心欄位對應（已依題庫自動配對）：</h4>
              <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 text-xs" id="column-mapper-grid"></div>
            </div>

            <div class="overflow-x-auto max-h-60 border border-slate-200 rounded-lg">
              <table class="w-full text-left text-xs border-collapse bg-white">
                <thead class="bg-slate-100 text-slate-600 font-bold sticky top-0" id="preview-table-header"></thead>
                <tbody class="divide-y divide-slate-200" id="preview-table-body"></tbody>
              </table>
            </div>

            <div class="flex justify-end gap-2 pt-2">
              <button onclick="executeDataImport('replace')" class="px-4 py-2 bg-slate-800 hover:bg-slate-900 text-white font-semibold rounded-lg text-xs">
                覆蓋當月資料
              </button>
              <button onclick="executeDataImport('append')" class="px-4 py-2 bg-[#EB0A1E] hover:bg-red-700 text-white font-semibold rounded-lg text-xs flex items-center gap-1.5 shadow">
                <i data-lucide="check" class="w-3.5 h-3.5"></i>
                確認匯入資料庫
              </button>
            </div>
          </div>
        </div>

        <div class="bg-white p-5 rounded-xl border border-rose-200 shadow-sm">
          <div class="flex items-center justify-between pb-3 border-b border-rose-100">
            <div>
              <h3 class="text-sm font-bold text-rose-800 flex items-center gap-2">
                <i data-lucide="trash-2" class="w-4 h-4 text-rose-600"></i>
                <span>手動清除指定條件資料</span>
              </h3>
              <p class="text-xs text-slate-500 mt-0.5">
                可精準選擇要刪除的特定月份、據點代碼、管道或無效筆數，避免誤刪完整歷史資料。
              </p>
            </div>
            <span class="text-[11px] font-semibold text-rose-600 bg-rose-50 px-2 py-1 rounded border border-rose-200">
              危險操作區
            </span>
          </div>

          <div class="mt-4 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
            <div>
              <label class="block font-semibold text-slate-700 mb-1">指定月份：</label>
              <select id="purge-month" onchange="updatePurgeCountPreview()" class="w-full border border-slate-300 rounded px-2.5 py-1.5 bg-slate-50">
                <option value="ALL">全部月份 (不限)</option>
              </select>
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">指定據點：</label>
              <select id="purge-branch" onchange="updatePurgeCountPreview()" class="w-full border border-slate-300 rounded px-2.5 py-1.5 bg-slate-50">
                <option value="ALL">全部據點 (不限)</option>
              </select>
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">指定調查來源：</label>
              <select id="purge-channel" onchange="updatePurgeCountPreview()" class="w-full border border-slate-300 rounded px-2.5 py-1.5 bg-slate-50">
                <option value="ALL">全部管道 (不限)</option>
                <option value="問卷">問卷</option>
                <option value="電訪">電訪</option>
              </select>
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">指定調查狀態：</label>
              <select id="purge-status" onchange="updatePurgeCountPreview()" class="w-full border border-slate-300 rounded px-2.5 py-1.5 bg-slate-50">
                <option value="ALL">全部狀態 (成功 ＋ 失敗)</option>
                <option value="成功">僅成功件</option>
                <option value="失敗">僅失敗件 (釋放空間)</option>
              </select>
            </div>
          </div>

          <div class="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between">
            <div class="text-xs text-slate-600">
              符合條件預計清除：<strong id="purge-preview-count" class="text-rose-600 text-sm">0</strong> 筆資料
            </div>
            <button onclick="executeConditionPurge()" class="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white font-semibold rounded-lg text-xs flex items-center gap-1.5 shadow">
              <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
              確認執行清除
            </button>
          </div>
        </div>
      </section>

      <!-- VIEW 3: 題目管理與設定 -->
      <section id="view-questions" class="space-y-5 hidden">
        <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <i data-lucide="list-checks" class="w-4 h-4 text-blue-600"></i>
                <span>調查題目維護管理</span>
              </h3>
              <p class="text-xs text-slate-500 mt-0.5">自訂題項與 Excel 原始欄位對應。空白與 0 自動排除於分母之外。</p>
            </div>
            <button onclick="openQuestionModal()" class="px-3 py-1.5 bg-[#EB0A1E] hover:bg-red-700 text-white rounded text-xs font-semibold flex items-center gap-1.5 shadow">
              <i data-lucide="plus" class="w-3.5 h-3.5"></i>
              新增問卷題目
            </button>
          </div>

          <div class="mt-4 overflow-x-auto">
            <table class="w-full text-left text-xs border-collapse">
              <thead>
                <tr class="bg-slate-100 text-slate-600 border-b border-slate-200">
                  <th class="p-2.5 font-bold">題號代碼</th>
                  <th class="p-2.5 font-bold">題目名稱說明</th>
                  <th class="p-2.5 font-bold">題型機制</th>
                  <th class="p-2.5 font-bold">Excel 原始對應欄</th>
                  <th class="p-2.5 font-bold text-center">計入加權滿意度</th>
                  <th class="p-2.5 font-bold text-right">操作</th>
                </tr>
              </thead>
              <tbody id="question-table-body" class="divide-y divide-slate-100"></tbody>
            </table>
          </div>
        </div>
      </section>

      <!-- VIEW 4: 個別題目統計分析 -->
      <section id="view-item-analysis" class="space-y-5 hidden">
        <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
          <div class="flex items-center justify-between mb-4">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <i data-lucide="table" class="w-4 h-4 text-emerald-600"></i>
                <span>個別題目統計彙整表</span>
              </h3>
              <p class="text-xs text-slate-500 mt-0.5">分母為純有效填答數（排除跳題、0 與無效回覆）。統計量包含平均數、中位數、眾數與高低分比率。</p>
            </div>
            <button onclick="exportItemAnalysisExcel()" class="px-3 py-1.5 border border-slate-300 hover:bg-slate-50 rounded text-xs font-semibold text-slate-700 flex items-center gap-1.5">
              <i data-lucide="sheet" class="w-3.5 h-3.5"></i>
              匯出題目統計表 (Excel)
            </button>
          </div>

          <div class="overflow-x-auto">
            <table class="w-full text-left text-xs border-collapse">
              <thead>
                <tr class="bg-slate-100 text-slate-700 border-b border-slate-200 font-bold">
                  <th class="p-2.5">代碼</th>
                  <th class="p-2.5">題目名稱</th>
                  <th class="p-2.5 text-right">有效樣本</th>
                  <th class="p-2.5 text-right">平均得分 / 達成率</th>
                  <th class="p-2.5 text-right">中位數</th>
                  <th class="p-2.5 text-right">眾數</th>
                  <th class="p-2.5 text-right text-emerald-700">9-10分率</th>
                  <th class="p-2.5 text-right text-blue-700">8-10分率</th>
                  <th class="p-2.5 text-right text-rose-700">1-6分率</th>
                </tr>
              </thead>
              <tbody id="item-analysis-table-body" class="divide-y divide-slate-100"></tbody>
            </table>
          </div>
        </div>
      </section>

      <!-- VIEW 5: 據點比較 -->
      <section id="view-branch-compare" class="space-y-5 hidden">
        <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
          <div class="flex items-center justify-between mb-4">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <i data-lucide="git-compare" class="w-4 h-4 text-indigo-600"></i>
                <span id="title-branch-compare-header">據點績效表現綜合表</span>
              </h3>
              <p class="text-xs text-slate-500 mt-0.5">點擊各欄位標題可進行正反向排序，支援快速定位低分警戒單位。</p>
            </div>
            <button onclick="exportBranchCompareExcel()" class="px-3 py-1.5 border border-slate-300 hover:bg-slate-50 rounded text-xs font-semibold text-slate-700 flex items-center gap-1.5">
              <i data-lucide="sheet" class="w-3.5 h-3.5"></i>
              匯出據點報表 (Excel)
            </button>
          </div>

          <div class="overflow-x-auto">
            <table class="w-full text-left text-xs border-collapse" id="branch-compare-table">
              <thead>
                <tr class="bg-slate-100 text-slate-700 border-b border-slate-200 font-bold">
                  <th class="p-2.5 cursor-pointer hover:bg-slate-200" onclick="sortBranchTable('name')">據點代碼/名稱 ↕</th>
                  <th class="p-2.5 text-right cursor-pointer hover:bg-slate-200" onclick="sortBranchTable('resp')">有效成功填答 ↕</th>
                  <th class="p-2.5 text-right cursor-pointer hover:bg-slate-200" onclick="sortBranchTable('avg')">平均滿意度 ↕</th>
                  <th class="p-2.5 text-right cursor-pointer hover:bg-slate-200 text-emerald-700" onclick="sortBranchTable('high')">高分率 (9-10) ↕</th>
                  <th class="p-2.5 text-right cursor-pointer hover:bg-slate-200 text-blue-700" onclick="sortBranchTable('pos')">正向率 (8-10) ↕</th>
                  <th class="p-2.5 text-right cursor-pointer hover:bg-slate-200 text-rose-700" onclick="sortBranchTable('low')">低分率 (1-6) ↕</th>
                  <th class="p-2.5 text-right">差異 (<span id="compare-mode-col-header">上月</span>)</th>
                  <th class="p-2.5 text-center">排名</th>
                </tr>
              </thead>
              <tbody id="branch-compare-table-body" class="divide-y divide-slate-100"></tbody>
            </table>
          </div>
        </div>
      </section>

      <!-- VIEW 6: 構面分析 (擴充保留) -->
      <section id="view-dimensions" class="space-y-5 hidden">
        <div class="bg-white p-8 rounded-xl border border-slate-200 shadow-sm text-center">
          <div class="w-12 h-12 bg-slate-100 rounded-full flex items-center justify-center mx-auto text-slate-500 mb-3">
            <i data-lucide="layers" class="w-6 h-6"></i>
          </div>
          <h3 class="text-base font-bold text-slate-800">構面分析功能（架構保留中）</h3>
          <p class="text-xs text-slate-500 max-w-md mx-auto mt-1.5 leading-relaxed">
            系統目前無綁定任何預設構面，直接以個別題目進行分析。未來如需依自訂構面彙總，可隨時至題目設定進行指派。
          </p>
          <div class="mt-4">
            <button onclick="navigateView('questions')" class="px-4 py-2 bg-slate-800 hover:bg-slate-900 text-white rounded-lg text-xs font-semibold">
              前往題目管理介面
            </button>
          </div>
        </div>
      </section>

      <!-- VIEW 7: 主管月報產生器 -->
      <section id="view-report" class="space-y-5 hidden">
        <div class="flex items-center justify-between no-print bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
          <div>
            <h3 class="text-sm font-bold text-slate-800">主管月會營運決策報表</h3>
            <p class="text-xs text-slate-500 mt-0.5">整合公司加權績效、題目警示排行、檢核落實度與據點總評，適應 A4 橫式列印排版。</p>
          </div>
          <div class="flex items-center space-x-2">
            <button onclick="window.print()" class="px-3.5 py-1.5 bg-[#EB0A1E] hover:bg-red-700 text-white font-semibold rounded-lg text-xs flex items-center gap-1.5 shadow">
              <i data-lucide="printer" class="w-3.5 h-3.5"></i>
              列印／另存 PDF (A4 橫向)
            </button>
          </div>
        </div>

        <div id="printable-report" class="bg-white p-8 rounded-xl border border-slate-200 shadow-sm space-y-6">
          <div class="border-b-2 border-slate-900 pb-4 flex justify-between items-end">
            <div>
              <div class="text-xs font-bold text-[#EB0A1E] uppercase tracking-wider" id="report-track-title">TOYOTA 顧客滿意度月報 (SSI 購車調查體系)</div>
              <h2 class="text-xl font-black text-slate-900 mt-1" id="report-header-title">顧客滿意度整體績效報告</h2>
            </div>
            <div class="text-right text-xs text-slate-500">
              <div>產出日期：<span id="report-print-date">--</span></div>
              <div>成功評估樣本：<strong id="report-total-sample" class="text-slate-800">--</strong> 筆</div>
            </div>
          </div>

          <div>
            <h4 class="text-xs font-bold text-slate-700 uppercase tracking-wider mb-2.5">一、全公司整體滿意度總評</h4>
            <div class="grid grid-cols-4 gap-3 text-center border border-slate-200 rounded-lg p-3 bg-slate-50">
              <div>
                <div class="text-[11px] text-slate-500">整體加權平均分</div>
                <div class="text-xl font-black text-slate-900 mt-0.5" id="report-kpi-avg">--</div>
              </div>
              <div>
                <div class="text-[11px] text-slate-500">高滿意率 (9-10分)</div>
                <div class="text-xl font-black text-emerald-600 mt-0.5" id="report-kpi-high">--</div>
              </div>
              <div>
                <div class="text-[11px] text-slate-500">正向滿意率 (8-10分)</div>
                <div class="text-xl font-black text-blue-600 mt-0.5" id="report-kpi-pos">--</div>
              </div>
              <div>
                <div class="text-[11px] text-slate-500">低分警戒率 (1-6分)</div>
                <div class="text-xl font-black text-rose-600 mt-0.5" id="report-kpi-low">--</div>
              </div>
            </div>
          </div>

          <div>
            <h4 class="text-xs font-bold text-slate-700 uppercase tracking-wider mb-2">二、關鍵檢核流程落實達成率</h4>
            <div class="grid grid-cols-4 gap-2 text-center text-xs" id="report-checklist-boxes"></div>
          </div>

          <div class="grid grid-cols-2 gap-4">
            <div class="border border-slate-200 rounded-lg p-3">
              <h4 class="text-xs font-bold text-emerald-800 flex items-center gap-1 mb-2">
                <i data-lucide="thumbs-up" class="w-3.5 h-3.5"></i> 表現最優異評分題 TOP 2
              </h4>
              <ul class="space-y-1.5 text-xs" id="report-top-questions"></ul>
            </div>
            <div class="border border-slate-200 rounded-lg p-3">
              <h4 class="text-xs font-bold text-rose-800 flex items-center gap-1 mb-2">
                <i data-lucide="alert-triangle" class="w-3.5 h-3.5"></i> 最需聚焦改善題項 TOP 2
              </h4>
              <ul class="space-y-1.5 text-xs" id="report-worst-questions"></ul>
            </div>
          </div>

          <div>
            <h4 class="text-xs font-bold text-slate-700 uppercase tracking-wider mb-2.5">三、各據點綜合表現評比</h4>
            <table class="w-full text-xs text-left border-collapse border border-slate-200">
              <thead class="bg-slate-100 text-slate-700">
                <tr>
                  <th class="p-2 border border-slate-200">排名</th>
                  <th class="p-2 border border-slate-200">據點名稱/代碼</th>
                  <th class="p-2 border border-slate-200 text-right">有效成功樣本</th>
                  <th class="p-2 border border-slate-200 text-right">平均得分</th>
                  <th class="p-2 border border-slate-200 text-right">高分率 (9-10)</th>
                  <th class="p-2 border border-slate-200 text-right">低分率 (1-6)</th>
                  <th class="p-2 border border-slate-200 text-center">警戒燈號</th>
                </tr>
              </thead>
              <tbody id="report-branch-rows"></tbody>
            </table>
          </div>
        </div>
      </section>

      <!-- VIEW 8: 系統設定與資安 -->
      <section id="view-settings" class="space-y-5 hidden">
        <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-4">
          <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
            <i data-lucide="shield-check" class="w-4 h-4 text-emerald-600"></i>
            <span>資料安全性與本機備份機制</span>
          </h3>
          <p class="text-xs text-slate-600 leading-relaxed">
            本平台採用純前端本機運算（Client-side Memory Architecture）。所有 Excel 資料在瀏覽器內部直接以 SheetJS 解析，<strong>絕不傳送至外部伺服器或第三方雲端</strong>。車主姓名、電話與車牌等個人識別欄位均在匯入時自動隔離，且已整合「編號」以支援後續原檔稽核。
          </p>

          <div class="pt-3 border-t border-slate-200 flex flex-wrap gap-3">
            <button onclick="backupDatabaseToFile()" class="px-3.5 py-2 bg-slate-800 hover:bg-slate-900 text-white font-semibold rounded-lg text-xs flex items-center gap-1.5 shadow">
              <i data-lucide="download" class="w-3.5 h-3.5"></i>
              匯出完整資料備份 (.json)
            </button>
            <label class="px-3.5 py-2 border border-slate-300 hover:bg-slate-50 text-slate-700 font-semibold rounded-lg text-xs flex items-center gap-1.5 cursor-pointer">
              <i data-lucide="upload" class="w-3.5 h-3.5"></i>
              還原資料庫
              <input type="file" accept=".json" class="hidden" onchange="restoreDatabaseFromFile(event)">
            </label>
            <button onclick="resetToInitialData()" class="px-3.5 py-2 text-rose-600 hover:bg-rose-50 border border-rose-200 font-semibold rounded-lg text-xs ml-auto">
              重設回系統原廠示範資料
            </button>
          </div>
        </div>
      </section>

    </main>
  </div>

  <!-- 新增/編輯題目 Modal -->
  <div id="modal-question" class="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center hidden">
    <div class="bg-white rounded-xl shadow-xl border border-slate-200 max-w-md w-full p-6 text-xs space-y-4">
      <div class="flex justify-between items-center pb-2 border-b border-slate-200">
        <h3 class="font-bold text-slate-800 text-sm" id="modal-question-title">新增題目</h3>
        <button onclick="closeQuestionModal()" class="text-slate-400 hover:text-slate-600 font-bold text-base">&times;</button>
      </div>

      <div class="space-y-3">
        <div>
          <label class="block font-semibold text-slate-600 mb-1">題目代碼 (如 S1 / Q1)：</label>
          <input type="text" id="input-q-code" class="w-full border border-slate-300 rounded px-2.5 py-1.5 bg-slate-50 outline-none focus:border-red-500">
        </div>
        <div>
          <label class="block font-semibold text-slate-600 mb-1">題目名稱說明：</label>
          <input type="text" id="input-q-name" class="w-full border border-slate-300 rounded px-2.5 py-1.5 bg-slate-50 outline-none focus:border-red-500">
        </div>
        <div>
          <label class="block font-semibold text-slate-600 mb-1">Excel 原始欄位標題：</label>
          <input type="text" id="input-q-col" class="w-full border border-slate-300 rounded px-2.5 py-1.5 bg-slate-50 outline-none focus:border-red-500">
        </div>
        <div>
          <label class="block font-semibold text-slate-600 mb-1">題目類型：</label>
          <select id="input-q-type" class="w-full border border-slate-300 rounded px-2 py-1.5 bg-slate-50">
            <option value="rating">1～10 分評分題</option>
            <option value="yesno">二分檢核題 (全部皆有 / 有 / 是)</option>
          </select>
        </div>
        <div class="flex items-center space-x-2 pt-2">
          <input type="checkbox" id="input-q-calc" checked class="rounded text-red-600">
          <label for="input-q-calc" class="font-semibold text-slate-700">納入整體平均滿意度統計</label>
        </div>
      </div>

      <div class="flex justify-end gap-2 pt-3 border-t border-slate-200">
        <button onclick="closeQuestionModal()" class="px-3 py-1.5 border border-slate-300 rounded text-slate-600 font-semibold">取消</button>
        <button onclick="saveQuestionModal()" class="px-4 py-1.5 bg-[#EB0A1E] hover:bg-red-700 text-white rounded font-semibold shadow">儲存</button>
      </div>
    </div>
  </div>

  <script>
    function naturalSort(a, b) {
      return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
    }

    const DEFAULT_APP_STATE = {
      activeMode: 'SSI',
      currentMonth: '2026/07',

      ssi_questions: [
        { id: 'S1', code: 'S1', name: '整體滿意度', type: 'rating', excelCol: 'S1.整體滿意度)', inCalc: true },
        { id: 'S2', code: 'S2', name: '保險說明', type: 'yesno', excelCol: 'S2.保險說明', inCalc: false },
        { id: 'S3', code: 'S3', name: '新車大禮包', type: 'yesno', excelCol: 'S3.新車大禮包', inCalc: false },
        { id: 'S4', code: 'S4', name: '業代迎賓接待', type: 'rating', excelCol: 'S4.業代迎賓接待', inCalc: true },
        { id: 'S5', code: 'S5', name: '廠人員介紹', type: 'yesno', excelCol: 'S5.廠人員介紹', inCalc: false },
        { id: 'S6', code: 'S6', name: '所主管介紹', type: 'yesno', excelCol: 'S6.所主管介紹', inCalc: false }
      ],

      csi_questions: [
        { id: 'Q1', code: 'Q1', name: '整體服務滿意度', type: 'rating', excelCol: '整體滿意度', inCalc: true },
        { id: 'Q2', code: 'Q2', name: '預約服務與接待', type: 'rating', excelCol: '接待服務', inCalc: true },
        { id: 'Q3', code: 'Q3', name: '專員態度與解說', type: 'rating', excelCol: '服務態度', inCalc: true },
        { id: 'Q4', code: 'Q4', name: '保修費用透明度', type: 'rating', excelCol: '費用透明度', inCalc: true },
        { id: 'Q5', code: 'Q5', name: '維修技術與品質', type: 'rating', excelCol: '維修品質', inCalc: true },
        { id: 'Q6', code: 'Q6', name: '完工車輛清潔', type: 'rating', excelCol: '交車清潔', inCalc: true },
        { id: 'Q7', code: 'Q7', name: '是否做到一次修復', type: 'yesno', excelCol: '一次修復', inCalc: false }
      ],

      ssi_branches: ['E01', 'E02', 'E06', 'E15', 'E16', 'E20'],
      csi_branches: ['永康服務廠', '東台南服務廠', '新營服務廠', '北台南服務廠', '南台南服務廠'],

      ssi_records: [],
      csi_records: []
    };

    function generateMockData() {
      const months = ['2025/10', '2026/07', '2026/08', '2026/09', '2026/10'];
      const ssiData = [];
      const csiData = [];

      months.forEach(m => {
        DEFAULT_APP_STATE.ssi_branches.forEach((b, bIdx) => {
          const count = 18 + Math.floor(Math.random() * 8);
          for (let i = 0; i < count; i++) {
            const isFailed = Math.random() < 0.04;
            const dayStr = String(Math.floor(Math.random() * 26) + 1).padStart(2, '0');
            const seqNo = `SSI-${m.replace('/', '')}-${b}-${String(i + 1).padStart(3, '0')}`;
            
            if (isFailed) {
              ssiData.push({
                no: seqNo, id: seqNo, month: m, date: `${m}/${dayStr}`, branch: b, channel: '電訪', status: '失敗', scores: {}
              });
              return;
            }

            const isLowScore = Math.random() < 0.05;
            const s1Score = isLowScore ? 6 : (Math.random() > 0.15 ? 10 : 9);
            const s4Score = Math.random() > 0.1 ? (isLowScore ? 7 : 10) : '';

            ssiData.push({
              no: seqNo,
              id: seqNo,
              month: m,
              date: `${m}/${dayStr}`,
              branch: b,
              channel: Math.random() > 0.4 ? '問卷' : '電訪',
              status: '成功',
              scores: {
                S1: s1Score,
                S2: Math.random() > 0.1 ? '全部皆有' : (Math.random() > 0.5 ? '沒有讓您選擇保險公司' : '不清楚/不知道/不記得'),
                S3: Math.random() > 0.08 ? '有' : (Math.random() > 0.5 ? '沒有' : '不清楚/不知道/不記得'),
                S4: s4Score,
                S5: Math.random() > 0.1 ? '有' : '沒有',
                S6: Math.random() > 0.2 ? '有' : '沒有'
              }
            });
          }
        });

        DEFAULT_APP_STATE.csi_branches.forEach((b, bIdx) => {
          for (let i = 0; i < 20; i++) {
            const seqNo = `CSI-${m.replace('/', '')}-${bIdx}-${String(i + 1).padStart(3, '0')}`;
            csiData.push({
              no: seqNo, id: seqNo, month: m, date: `${m}/15`, branch: b, channel: '電訪', status: '成功',
              scores: { Q1: 10, Q2: 10, Q3: 10, Q4: 9, Q5: 10, Q6: 10, Q7: '是' }
            });
          }
        });
      });

      DEFAULT_APP_STATE.ssi_records = ssiData;
      DEFAULT_APP_STATE.csi_records = csiData;
    }

    let appState = null;
    function loadAppState() {
      const stored = localStorage.getItem('toyota_csi_ssi_app_data_v7');
      if (stored) {
        try {
          appState = JSON.parse(stored);
        } catch (e) {
          console.error(e);
        }
      }
      if (!appState || !appState.ssi_records || appState.ssi_records.length === 0) {
        generateMockData();
        appState = JSON.parse(JSON.stringify(DEFAULT_APP_STATE));
        saveAppState();
      }
    }

    function saveAppState() {
      localStorage.setItem('toyota_csi_ssi_app_data_v7', JSON.stringify(appState));
    }

    function resetToInitialData() {
      if (confirm('確定要清除所有資料並重設回原廠示範資料庫嗎？')) {
        localStorage.removeItem('toyota_csi_ssi_app_data_v7');
        generateMockData();
        appState = JSON.parse(JSON.stringify(DEFAULT_APP_STATE));
        saveAppState();
        location.reload();
      }
    }

    let activeView = 'dashboard';
    let chartRanking = null;
    let chartQuestionStack = null;
    let chartTrend = null;
    let chartChecklistTrend = null;
    let currentSortField = 'avg';
    let currentSortAsc = false;

    function switchTrackMode(mode) {
      appState.activeMode = mode;
      saveAppState();

      const btnCsi = document.getElementById('btn-mode-csi');
      const btnSsi = document.getElementById('btn-mode-ssi');

      if (mode === 'SSI') {
        btnSsi.className = 'flex items-center space-x-1.5 px-3 py-1.5 rounded-md text-xs font-bold transition-all shadow bg-[#EB0A1E] text-white';
        btnCsi.className = 'flex items-center space-x-1.5 px-3 py-1.5 rounded-md text-xs font-bold transition-all text-slate-300 hover:text-white';
        document.getElementById('label-nav-branch').innerText = '營業所比較';
        document.getElementById('label-filter-branch').innerText = '據點代碼：';
        document.getElementById('title-branch-ranking').innerText = '營業所滿意度綜合排行 (當月)';
        document.getElementById('title-branch-compare-header').innerText = '營業所績效表現綜合表';
        document.getElementById('title-checklist-branch-header').innerText = '各據點檢核題落實達成率比較 (當月)';
        document.getElementById('import-target-badge').innerText = '購車滿意度 (SSI)';
        document.getElementById('report-track-title').innerText = 'TOYOTA 顧客滿意度月報 (SSI 購車調查體系)';
      } else {
        btnCsi.className = 'flex items-center space-x-1.5 px-3 py-1.5 rounded-md text-xs font-bold transition-all shadow bg-[#EB0A1E] text-white';
        btnSsi.className = 'flex items-center space-x-1.5 px-3 py-1.5 rounded-md text-xs font-bold transition-all text-slate-300 hover:text-white';
        document.getElementById('label-nav-branch').innerText = '服務廠比較';
        document.getElementById('label-filter-branch').innerText = '服務廠：';
        document.getElementById('title-branch-ranking').innerText = '服務廠滿意度綜合排行 (當月)';
        document.getElementById('title-branch-compare-header').innerText = '服務廠績效表現綜合表';
        document.getElementById('title-checklist-branch-header').innerText = '各服務廠檢核題落實達成率比較 (當月)';
        document.getElementById('import-target-badge').innerText = '保修滿意度 (CSI)';
        document.getElementById('report-track-title').innerText = 'TOYOTA 顧客滿意度月報 (CSI 保修體系)';
      }

      initMonthFilterDropdown();
      initBranchFilterDropdown();
      initPurgeOptions();
      renderActiveView();
    }

    function getActiveRecords() {
      return appState.activeMode === 'CSI' ? appState.csi_records : appState.ssi_records;
    }

    function getActiveQuestions() {
      return appState.activeMode === 'CSI' ? appState.csi_questions : appState.ssi_questions;
    }

    function getActiveBranches() {
      return (appState.activeMode === 'CSI' ? appState.csi_branches : appState.ssi_branches).slice().sort(naturalSort);
    }

    function initMonthFilterDropdown() {
      const records = getActiveRecords();
      const select = document.getElementById('filter-month');
      const months = Array.from(new Set(records.map(r => r.month))).filter(Boolean).sort().reverse();

      const currentVal = select.value;
      select.innerHTML = '';

      if (months.length === 0) {
        select.innerHTML = '<option value="2026/07">2026/07</option>';
        return;
      }

      months.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m;
        opt.textContent = m;
        select.appendChild(opt);
      });

      if (months.includes(currentVal)) {
        select.value = currentVal;
      } else {
        select.value = months[0];
      }
    }

    function initBranchFilterDropdown() {
      const select = document.getElementById('filter-branch');
      const branches = getActiveBranches();
      const currentVal = select.value;

      select.innerHTML = '<option value="ALL">公司全體 (所有據點)</option>';
      branches.forEach(b => {
        const opt = document.createElement('option');
        opt.value = b;
        opt.textContent = b;
        select.appendChild(opt);
      });

      if (branches.includes(currentVal)) select.value = currentVal;
      else select.value = 'ALL';
    }

    function initPurgeOptions() {
      const records = getActiveRecords();
      const monthSelect = document.getElementById('purge-month');
      const branchSelect = document.getElementById('purge-branch');

      const months = Array.from(new Set(records.map(r => r.month))).filter(Boolean).sort().reverse();
      monthSelect.innerHTML = '<option value="ALL">全部月份 (不限)</option>';
      months.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m;
        opt.textContent = m;
        monthSelect.appendChild(opt);
      });

      const branches = getActiveBranches();
      branchSelect.innerHTML = '<option value="ALL">全部據點 (不限)</option>';
      branches.forEach(b => {
        const opt = document.createElement('option');
        opt.value = b;
        opt.textContent = b;
        branchSelect.appendChild(opt);
      });

      updatePurgeCountPreview();
    }

    function getMatchedPurgeRecords() {
      const monthVal = document.getElementById('purge-month').value;
      const branchVal = document.getElementById('purge-branch').value;
      const channelVal = document.getElementById('purge-channel').value;
      const statusVal = document.getElementById('purge-status').value;
      const records = getActiveRecords();

      return records.filter(r => {
        if (monthVal !== 'ALL' && r.month !== monthVal) return false;
        if (branchVal !== 'ALL' && r.branch !== branchVal) return false;
        if (channelVal !== 'ALL' && r.channel !== channelVal) return false;
        if (statusVal !== 'ALL' && (r.status || '成功') !== statusVal) return false;
        return true;
      });
    }

    function updatePurgeCountPreview() {
      const matched = getMatchedPurgeRecords();
      document.getElementById('purge-preview-count').innerText = matched.length;
    }

    function executeConditionPurge() {
      const matched = getMatchedPurgeRecords();
      if (matched.length === 0) {
        alert('目前條件下無符合之資料，無需清除。');
        return;
      }

      const monthVal = document.getElementById('purge-month').value;
      const branchVal = document.getElementById('purge-branch').value;
      const channelVal = document.getElementById('purge-channel').value;
      const statusVal = document.getElementById('purge-status').value;

      const confirmText = `警告：您即將從【${appState.activeMode === 'SSI' ? '購車調查' : '保修調查'}】中刪除 ${matched.length} 筆資料！\n` +
                          `條件：\n- 月份：${monthVal}\n- 據點：${branchVal}\n- 調查來源：${channelVal}\n- 調查狀態：${statusVal}\n\n此操作無法復原，是否確定刪除？`;

      if (confirm(confirmText)) {
        const matchedIds = new Set(matched.map(r => r.id));
        if (appState.activeMode === 'SSI') {
          appState.ssi_records = appState.ssi_records.filter(r => !matchedIds.has(r.id));
        } else {
          appState.csi_records = appState.csi_records.filter(r => !matchedIds.has(r.id));
        }

        saveAppState();
        initMonthFilterDropdown();
        initBranchFilterDropdown();
        initPurgeOptions();
        renderActiveView();
        alert(`已成功清除 ${matched.length} 筆指定條件之資料！`);
      }
    }

    function navigateView(viewName) {
      activeView = viewName;
      ['dashboard', 'checklist-trends', 'import', 'questions', 'item-analysis', 'branch-compare', 'dimensions', 'report', 'settings'].forEach(v => {
        const el = document.getElementById(`view-${v}`);
        const nav = document.getElementById(`nav-${v}`);
        if (el) el.classList.toggle('hidden', v !== viewName);
        if (nav) {
          if (v === viewName) nav.className = 'w-full flex items-center space-x-2.5 px-3 py-2 text-xs font-semibold rounded-lg bg-red-50 text-[#EB0A1E] transition';
          else nav.className = 'w-full flex items-center space-x-2.5 px-3 py-2 text-xs font-semibold rounded-lg text-slate-600 hover:bg-slate-100 transition';
        }
      });

      document.getElementById('filter-bar').classList.toggle('hidden', ['import', 'questions', 'settings', 'dimensions'].includes(viewName));
      if (viewName === 'import') initPurgeOptions();
      renderActiveView();
    }

    function getFilteredData(targetMonth = null) {
      const month = targetMonth || document.getElementById('filter-month').value;
      const branch = document.getElementById('filter-branch').value;
      const channel = document.getElementById('filter-channel').value;
      const records = getActiveRecords();

      return records.filter(r => {
        if (r.status && r.status === '失敗') return false;
        if (r.month !== month) return false;
        if (branch !== 'ALL' && r.branch !== branch) return false;
        if (channel !== 'ALL' && r.channel !== channel) return false;
        return true;
      });
    }

    // 核心統計函數：將「不清楚、不知道、不記得」列為無效，嚴格辨識「沒有」為未達成
    function calculateStatistics(dataset) {
      const questions = getActiveQuestions();
      let totalScoresSum = 0;
      let totalValidScoresCount = 0;
      let countHigh = 0;
      let countPos = 0;
      let countLow = 0;

      const questionStats = {};
      questions.forEach(q => {
        questionStats[q.code] = {
          validCount: 0,
          sum: 0,
          scores: [],
          scoreFreq: { 1:0, 2:0, 3:0, 4:0, 5:0, 6:0, 7:0, 8:0, 9:0, 10:0 },
          highCount: 0,
          posCount: 0,
          lowCount: 0,
          yesCount: 0,
          noCount: 0
        };
      });

      dataset.forEach(row => {
        questions.forEach(q => {
          const val = row.scores ? row.scores[q.code] : undefined;
          if (val === undefined || val === null || val === '') return;

          const qStat = questionStats[q.code];

          if (q.type === 'rating') {
            const num = Number(val);
            if (!isNaN(num) && num >= 1 && num <= 10) {
              qStat.validCount++;
              qStat.sum += num;
              qStat.scores.push(num);
              qStat.scoreFreq[num] = (qStat.scoreFreq[num] || 0) + 1;

              if (num >= 9) qStat.highCount++;
              if (num >= 8) qStat.posCount++;
              if (num <= 6) qStat.lowCount++;

              if (q.inCalc) {
                totalScoresSum += num;
                totalValidScoresCount++;
                if (num >= 9) countHigh++;
                if (num >= 8) countPos++;
                if (num <= 6) countLow++;
              }
            }
          } else if (q.type === 'yesno') {
            const strVal = String(val).trim();
            
            // 規則 1：0 或空白視為跳題／未填答，直接排除
            if (strVal === '0' || strVal === '') return;

            // 規則 2：「不清楚」、「不知道」、「不記得」列入無效資料，直接排除不計入分母
            if (strVal.includes('不清楚') || strVal.includes('不知道') || strVal.includes('不記得')) {
              return;
            }

            // 規則 3：正面達成（包含「全部皆有」且未勾選任何沒有，或完全等於「有」、「是」、「1」）
            const isPositive = (strVal === '有' || strVal === '是' || strVal === '1' || (strVal.includes('全部皆有') && !strVal.includes('沒有')));
            
            // 規則 4：負向未達成（包含「沒有」、「無」、「否」）
            const isNegative = (strVal.includes('沒有') || strVal.includes('無') || strVal.includes('否'));

            if (isPositive) {
              qStat.validCount++;
              qStat.yesCount++;
            } else if (isNegative) {
              qStat.validCount++;
              qStat.noCount++;
            }
          }
        });
      });

      questions.forEach(q => {
        const qStat = questionStats[q.code];
        if (q.type === 'rating' && qStat.validCount > 0) {
          qStat.avg = qStat.sum / qStat.validCount;
          qStat.scores.sort((a, b) => a - b);
          const mid = Math.floor(qStat.scores.length / 2);
          qStat.median = qStat.scores.length % 2 !== 0 ? qStat.scores[mid] : (qStat.scores[mid - 1] + qStat.scores[mid]) / 2;

          let maxFreq = 0;
          let modeVal = qStat.scores[0];
          for (let s = 1; s <= 10; s++) {
            if (qStat.scoreFreq[s] > maxFreq) {
              maxFreq = qStat.scoreFreq[s];
              modeVal = s;
            }
          }
          qStat.mode = modeVal;
          qStat.highRate = (qStat.highCount / qStat.validCount) * 100;
          qStat.posRate = (qStat.posCount / qStat.validCount) * 100;
          qStat.lowRate = (qStat.lowCount / qStat.validCount) * 100;
        } else if (q.type === 'yesno' && qStat.validCount > 0) {
          qStat.yesRate = (qStat.yesCount / qStat.validCount) * 100;
        } else {
          qStat.avg = 0;
          qStat.median = 0;
          qStat.mode = 0;
          qStat.highRate = 0;
          qStat.posRate = 0;
          qStat.lowRate = 0;
          qStat.yesRate = 0;
        }
      });

      const overallAvg = totalValidScoresCount > 0 ? (totalScoresSum / totalValidScoresCount) : 0;
      const overallHighRate = totalValidScoresCount > 0 ? (countHigh / totalValidScoresCount) * 100 : 0;
      const overallPosRate = totalValidScoresCount > 0 ? (countPos / totalValidScoresCount) * 100 : 0;
      const overallLowRate = totalValidScoresCount > 0 ? (countLow / totalValidScoresCount) * 100 : 0;

      return {
        totalResponses: dataset.length,
        overallAvg,
        overallHighRate,
        overallPosRate,
        overallLowRate,
        questionStats
      };
    }

    function getCompareTargetMonth(currMonth, mode = 'MoM') {
      const parts = currMonth.split('/');
      let y = parseInt(parts[0], 10);
      let m = parseInt(parts[1], 10);

      if (mode === 'YoY') {
        return `${y - 1}/${String(m).padStart(2, '0')}`;
      } else {
        m -= 1;
        if (m === 0) { m = 12; y -= 1; }
        return `${y}/${String(m).padStart(2, '0')}`;
      }
    }

    function renderActiveView() {
      const records = getActiveRecords().filter(r => (!r.status || r.status === '成功'));
      document.getElementById('stat-record-count').innerText = records.length;

      if (activeView === 'dashboard') renderDashboard();
      else if (activeView === 'checklist-trends') renderChecklistTrendsView();
      else if (activeView === 'item-analysis') renderItemAnalysis();
      else if (activeView === 'branch-compare') renderBranchCompare();
      else if (activeView === 'questions') renderQuestionsView();
      else if (activeView === 'report') renderExecutiveReport();

      lucide.createIcons();
    }

    function renderDashboard() {
      const currMonth = document.getElementById('filter-month').value;
      const compareMode = document.getElementById('filter-compare-mode').value;
      const compareTargetMonth = getCompareTargetMonth(currMonth, compareMode);

      const currData = getFilteredData(currMonth);
      const prevData = getFilteredData(compareTargetMonth);

      const currStat = calculateStatistics(currData);
      const prevStat = calculateStatistics(prevData);

      document.getElementById('kpi-avg-score').innerText = currStat.overallAvg > 0 ? currStat.overallAvg.toFixed(2) : '--';
      document.getElementById('kpi-total-resp').innerText = `${currStat.totalResponses} 份`;
      document.getElementById('kpi-high-rate').innerText = `${currStat.overallHighRate.toFixed(1)}%`;
      document.getElementById('kpi-pos-rate').innerText = `${currStat.overallPosRate.toFixed(1)}%`;
      document.getElementById('kpi-low-rate').innerText = `${currStat.overallLowRate.toFixed(1)}%`;

      const compareSuffix = compareMode === 'YoY' ? ' (去年同期)' : ' (上月)';
      renderDiffBadge('kpi-avg-diff', currStat.overallAvg - prevStat.overallAvg, '分', false, compareSuffix);
      renderDiffBadge('kpi-high-diff', currStat.overallHighRate - prevStat.overallHighRate, '%', false, compareSuffix);
      renderDiffBadge('kpi-pos-diff', currStat.overallPosRate - prevStat.overallPosRate, '%', false, compareSuffix);
      renderDiffBadge('kpi-low-diff', currStat.overallLowRate - prevStat.overallLowRate, '%', true, compareSuffix);

      renderDashboardChecklistCards(currStat, prevStat, compareSuffix);
      renderBranchRankingChart(currMonth);
      renderQuestionStackChart(currStat);
      renderTrendChart();
    }

    function renderDiffBadge(elementId, diffVal, unit, invertColor = false, suffix = '') {
      const el = document.getElementById(elementId);
      if (isNaN(diffVal) || Math.abs(diffVal) < 0.001) {
        el.className = 'text-[11px] text-slate-400 mt-1 font-semibold';
        el.innerText = `持平 ${suffix}`;
        return;
      }
      const isPositive = diffVal > 0;
      const isGood = invertColor ? !isPositive : isPositive;
      el.className = `text-[11px] mt-1 font-semibold ${isGood ? 'text-emerald-600' : 'text-rose-600'}`;
      el.innerText = `${isPositive ? '▲ 上升' : '▼ 下降'} ${Math.abs(diffVal).toFixed(2)}${unit}${suffix}`;
    }

    function renderDashboardChecklistCards(currStat, prevStat, compareSuffix) {
      const container = document.getElementById('dashboard-checklist-grid');
      container.innerHTML = '';
      const yesNoQuestions = getActiveQuestions().filter(q => q.type === 'yesno');

      if (yesNoQuestions.length === 0) {
        container.innerHTML = '<div class="col-span-4 text-xs text-slate-400 py-2">目前題庫無二分檢核題目</div>';
        return;
      }

      yesNoQuestions.forEach(q => {
        const cRate = currStat.questionStats[q.code]?.yesRate || 0;
        const pRate = prevStat.questionStats[q.code]?.yesRate || 0;
        const diff = cRate - pRate;
        const valid = currStat.questionStats[q.code]?.validCount || 0;
        const yesCnt = currStat.questionStats[q.code]?.yesCount || 0;
        const noCnt = currStat.questionStats[q.code]?.noCount || 0;

        const card = document.createElement('div');
        card.className = 'p-3 bg-slate-50 border border-slate-200 rounded-lg text-center';
        card.innerHTML = `
          <div class="text-[11px] font-bold text-slate-600 truncate">${q.code} ${q.name}</div>
          <div class="text-xl font-black text-emerald-600 mt-1">${cRate.toFixed(1)}%</div>
          <div class="text-[10px] text-slate-400 mt-0.5">達成：${yesCnt} / 未達成：${noCnt}</div>
          <div class="text-[10px] font-semibold mt-0.5 ${diff >= 0 ? 'text-emerald-600' : 'text-rose-600'}">
            ${diff >= 0 ? '▲ +' : '▼ '}${diff.toFixed(1)}%
          </div>
        `;
        container.appendChild(card);
      });
    }

    function renderBranchRankingChart(month) {
      const branches = getActiveBranches();
      const records = getActiveRecords().filter(r => r.month === month && (!r.status || r.status === '成功'));

      const branchScores = branches.map(b => {
        const bData = records.filter(r => r.branch === b);
        const stat = calculateStatistics(bData);
        return { name: b, avg: stat.overallAvg, count: bData.length };
      }).sort((a, b) => b.avg - a.avg);

      const ctx = document.getElementById('chart-branch-ranking').getContext('2d');
      if (chartRanking) chartRanking.destroy();

      chartRanking = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: branchScores.map(x => x.name),
          datasets: [{
            data: branchScores.map(x => Number(x.avg.toFixed(2))),
            backgroundColor: branchScores.map((x, i) => i === 0 ? '#0f2444' : (x.avg < 8.5 ? '#f43f5e' : '#3b82f6')),
            borderRadius: 6
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: {
            y: { min: 7, max: 10, ticks: { stepSize: 1 } },
            x: { ticks: { font: { size: 11 } } }
          }
        }
      });
    }

    function renderQuestionStackChart(stat) {
      const questions = getActiveQuestions().filter(q => q.type === 'rating');
      const labels = questions.map(q => q.name);

      const highRates = questions.map(q => Number((stat.questionStats[q.code]?.highRate || 0).toFixed(1)));
      const posRates = questions.map(q => {
        const pos = stat.questionStats[q.code]?.posRate || 0;
        const high = stat.questionStats[q.code]?.highRate || 0;
        return Number((Math.max(0, pos - high)).toFixed(1));
      });
      const lowRates = questions.map(q => Number((stat.questionStats[q.code]?.lowRate || 0).toFixed(1)));

      const ctx = document.getElementById('chart-question-stack').getContext('2d');
      if (chartQuestionStack) chartQuestionStack.destroy();

      chartQuestionStack = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [
            { label: '9-10分 (高分)', data: highRates, backgroundColor: '#10b981' },
            { label: '7-8分 (正向)', data: posRates, backgroundColor: '#38bdf8' },
            { label: '1-6分 (警戒)', data: lowRates, backgroundColor: '#f43f5e' }
          ]
        },
        options: {
          indexAxis: 'y',
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            x: { stacked: true, max: 100, ticks: { callback: v => v + '%' } },
            y: { stacked: true }
          },
          plugins: { legend: { display: false } }
        }
      });
    }

    function renderTrendChart() {
      const records = getActiveRecords().filter(r => (!r.status || r.status === '成功'));
      const allMonths = Array.from(new Set(records.map(r => r.month))).filter(Boolean).sort();
      const avgScores = [];
      const sampleCounts = [];

      allMonths.forEach(m => {
        const d = getFilteredData(m);
        const stat = calculateStatistics(d);
        avgScores.push(Number(stat.overallAvg.toFixed(2)));
        sampleCounts.push(stat.totalResponses);
      });

      const ctx = document.getElementById('chart-trend').getContext('2d');
      if (chartTrend) chartTrend.destroy();

      chartTrend = new Chart(ctx, {
        type: 'line',
        data: {
          labels: allMonths,
          datasets: [
            {
              label: '平均滿意度得分',
              data: avgScores,
              borderColor: '#EB0A1E',
              backgroundColor: 'rgba(235, 10, 30, 0.1)',
              yAxisID: 'y',
              fill: true,
              tension: 0.3
            },
            {
              label: '成功樣本數',
              data: sampleCounts,
              type: 'bar',
              backgroundColor: 'rgba(15, 36, 68, 0.25)',
              yAxisID: 'y1',
              borderRadius: 4
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            y: { min: 7.5, max: 10, title: { display: true, text: '平均分數' } },
            y1: { position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: '樣本數' } }
          }
        }
      });
    }

    function renderChecklistTrendsView() {
      const records = getActiveRecords().filter(r => (!r.status || r.status === '成功'));
      const allMonths = Array.from(new Set(records.map(r => r.month))).filter(Boolean).sort();
      const yesNoQuestions = getActiveQuestions().filter(q => q.type === 'yesno');
      const colors = ['#10b981', '#3b82f6', '#f59e0b', '#8b5cf6', '#ec4899'];

      const datasets = yesNoQuestions.map((q, idx) => {
        const rates = allMonths.map(m => {
          const d = records.filter(r => r.month === m);
          const stat = calculateStatistics(d);
          return Number((stat.questionStats[q.code]?.yesRate || 0).toFixed(1));
        });

        return {
          label: `${q.code} ${q.name}`,
          data: rates,
          borderColor: colors[idx % colors.length],
          backgroundColor: colors[idx % colors.length],
          tension: 0.25,
          fill: false
        };
      });

      const ctx = document.getElementById('chart-checklist-month-trend').getContext('2d');
      if (chartChecklistTrend) chartChecklistTrend.destroy();

      chartChecklistTrend = new Chart(ctx, {
        type: 'line',
        data: { labels: allMonths, datasets: datasets },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            y: { min: 60, max: 100, ticks: { callback: v => v + '%' }, title: { display: true, text: '達成率 (%)' } }
          }
        }
      });

      const currMonth = document.getElementById('filter-month').value;
      const branches = getActiveBranches();
      const monthRecords = records.filter(r => r.month === currMonth);

      const thead = document.getElementById('checklist-table-header');
      thead.innerHTML = `
        <th class="p-2.5">據點代碼</th>
        <th class="p-2.5 text-right">有效成功填答</th>
        ${yesNoQuestions.map(q => `<th class="p-2.5 text-right">${q.code} ${q.name}</th>`).join('')}
      `;

      const tbody = document.getElementById('checklist-table-body');
      tbody.innerHTML = '';

      branches.forEach(b => {
        const bData = monthRecords.filter(r => r.branch === b);
        const bStat = calculateStatistics(bData);

        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';
        tr.innerHTML = `
          <td class="p-2.5 font-bold text-slate-800">${b}</td>
          <td class="p-2.5 text-right font-semibold text-slate-600">${bData.length}</td>
          ${yesNoQuestions.map(q => {
            const rate = bStat.questionStats[q.code]?.yesRate || 0;
            const isWarn = rate < 85;
            return `<td class="p-2.5 text-right font-bold ${isWarn ? 'text-rose-600' : 'text-emerald-600'}">${rate.toFixed(1)}%</td>`;
          }).join('')}
        `;
        tbody.appendChild(tr);
      });
    }

    function renderItemAnalysis() {
      const currMonth = document.getElementById('filter-month').value;
      const data = getFilteredData(currMonth);
      const stat = calculateStatistics(data);
      const questions = getActiveQuestions();

      const tbody = document.getElementById('item-analysis-table-body');
      tbody.innerHTML = '';

      questions.forEach(q => {
        const qStat = stat.questionStats[q.code] || {};
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';

        if (q.type === 'rating') {
          tr.innerHTML = `
            <td class="p-2.5 font-bold text-slate-700">${q.code}</td>
            <td class="p-2.5 font-medium text-slate-800">${q.name}</td>
            <td class="p-2.5 text-right font-semibold">${qStat.validCount || 0}</td>
            <td class="p-2.5 text-right font-black text-slate-900">${(qStat.avg || 0).toFixed(2)}</td>
            <td class="p-2.5 text-right text-slate-600">${(qStat.median || 0).toFixed(1)}</td>
            <td class="p-2.5 text-right text-slate-600">${qStat.mode || '--'}</td>
            <td class="p-2.5 text-right font-bold text-emerald-600">${(qStat.highRate || 0).toFixed(1)}%</td>
            <td class="p-2.5 text-right font-bold text-blue-600">${(qStat.posRate || 0).toFixed(1)}%</td>
            <td class="p-2.5 text-right font-bold text-rose-600">${(qStat.lowRate || 0).toFixed(1)}%</td>
          `;
        } else {
          const yesRate = qStat.yesRate || 0;
          tr.innerHTML = `
            <td class="p-2.5 font-bold text-slate-700">${q.code}</td>
            <td class="p-2.5 font-medium text-slate-800">${q.name} <span class="text-[10px] bg-slate-100 text-slate-500 px-1 py-0.5 rounded">檢核題</span></td>
            <td class="p-2.5 text-right font-semibold">${qStat.validCount || 0}</td>
            <td class="p-2.5 text-right font-black text-emerald-600">${yesRate.toFixed(1)}%</td>
            <td class="p-2.5 text-right text-slate-500" colspan="2">達成筆數：${qStat.yesCount || 0} 筆</td>
            <td class="p-2.5 text-right text-rose-600" colspan="3">未達成/沒有：${qStat.noCount || 0} 筆</td>
          `;
        }
        tbody.appendChild(tr);
      });
    }

    function renderBranchCompare() {
      const currMonth = document.getElementById('filter-month').value;
      const compareMode = document.getElementById('filter-compare-mode').value;
      const compareTargetMonth = getCompareTargetMonth(currMonth, compareMode);

      document.getElementById('compare-mode-col-header').innerText = compareMode === 'YoY' ? '去年同期' : '上月';

      const branches = getActiveBranches();
      const records = getActiveRecords().filter(r => (!r.status || r.status === '成功'));

      const branchList = branches.map(b => {
        const cData = records.filter(r => r.month === currMonth && r.branch === b);
        const pData = records.filter(r => r.month === compareTargetMonth && r.branch === b);
        const cStat = calculateStatistics(cData);
        const pStat = calculateStatistics(pData);

        return {
          name: b,
          resp: cStat.totalResponses,
          avg: cStat.overallAvg,
          high: cStat.overallHighRate,
          pos: cStat.overallPosRate,
          low: cStat.overallLowRate,
          diff: cStat.overallAvg - pStat.overallAvg
        };
      });

      branchList.sort((a, b) => {
        const factor = currentSortAsc ? 1 : -1;
        return (a[currentSortField] - b[currentSortField]) * factor;
      });

      const tbody = document.getElementById('branch-compare-table-body');
      tbody.innerHTML = '';

      branchList.forEach((item, index) => {
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';
        tr.innerHTML = `
          <td class="p-2.5 font-bold text-slate-800">${item.name}</td>
          <td class="p-2.5 text-right font-semibold text-slate-600">${item.resp}</td>
          <td class="p-2.5 text-right font-black text-slate-900">${item.avg.toFixed(2)}</td>
          <td class="p-2.5 text-right font-bold text-emerald-600">${item.high.toFixed(1)}%</td>
          <td class="p-2.5 text-right font-bold text-blue-600">${item.pos.toFixed(1)}%</td>
          <td class="p-2.5 text-right font-bold text-rose-600">${item.low.toFixed(1)}%</td>
          <td class="p-2.5 text-right text-xs ${item.diff >= 0 ? 'text-emerald-600' : 'text-rose-600'} font-semibold">
            ${item.diff >= 0 ? '▲' : '▼'} ${Math.abs(item.diff).toFixed(2)}
          </td>
          <td class="p-2.5 text-center">
            <span class="inline-block w-5 h-5 rounded-full ${index === 0 ? 'bg-amber-100 text-amber-800' : 'bg-slate-100 text-slate-600'} font-bold leading-5 text-[11px]">
              ${index + 1}
            </span>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function sortBranchTable(field) {
      if (currentSortField === field) currentSortAsc = !currentSortAsc;
      else { currentSortField = field; currentSortAsc = false; }
      renderBranchCompare();
    }

    function renderQuestionsView() {
      const questions = getActiveQuestions();
      const tbody = document.getElementById('question-table-body');
      tbody.innerHTML = '';

      questions.forEach(q => {
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';
        tr.innerHTML = `
          <td class="p-2.5 font-bold text-slate-700">${q.code}</td>
          <td class="p-2.5 font-medium text-slate-900">${q.name}</td>
          <td class="p-2.5 text-slate-600">${q.type === 'rating' ? '1～10 分評分題' : '二分檢核題'}</td>
          <td class="p-2.5 font-mono text-slate-500">${q.excelCol}</td>
          <td class="p-2.5 text-center">${q.inCalc ? '<span class="text-emerald-600 font-bold">✔ 是</span>' : '<span class="text-slate-400">否</span>'}</td>
          <td class="p-2.5 text-right space-x-1">
            <button onclick="editQuestion('${q.id}')" class="text-blue-600 hover:text-blue-800 font-semibold px-1.5 py-0.5">編輯</button>
            <button onclick="deleteQuestion('${q.id}')" class="text-rose-600 hover:text-rose-800 font-semibold px-1.5 py-0.5">刪除</button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    let editingQuestionId = null;
    function openQuestionModal(qId = null) {
      editingQuestionId = qId;
      const modal = document.getElementById('modal-question');
      modal.classList.remove('hidden');

      if (qId) {
        const q = getActiveQuestions().find(x => x.id === qId);
        document.getElementById('modal-question-title').innerText = '編輯調查題目';
        document.getElementById('input-q-code').value = q.code;
        document.getElementById('input-q-name').value = q.name;
        document.getElementById('input-q-col').value = q.excelCol;
        document.getElementById('input-q-type').value = q.type;
        document.getElementById('input-q-calc').checked = q.inCalc;
      } else {
        document.getElementById('modal-question-title').innerText = '新增調查題目';
        document.getElementById('input-q-code').value = '';
        document.getElementById('input-q-name').value = '';
        document.getElementById('input-q-col').value = '';
        document.getElementById('input-q-type').value = 'rating';
        document.getElementById('input-q-calc').checked = true;
      }
    }

    function closeQuestionModal() {
      document.getElementById('modal-question').classList.add('hidden');
      editingQuestionId = null;
    }

    function saveQuestionModal() {
      const code = document.getElementById('input-q-code').value.trim();
      const name = document.getElementById('input-q-name').value.trim();
      const excelCol = document.getElementById('input-q-col').value.trim();
      const type = document.getElementById('input-q-type').value;
      const inCalc = document.getElementById('input-q-calc').checked;

      if (!code || !name) {
        alert('請填寫題目代碼與題目名稱！');
        return;
      }

      const questions = getActiveQuestions();
      if (editingQuestionId) {
        const idx = questions.findIndex(x => x.id === editingQuestionId);
        if (idx !== -1) {
          questions[idx] = { id: editingQuestionId, code, name, excelCol: excelCol || name, type, inCalc };
        }
      } else {
        questions.push({
          id: 'Q_' + Date.now(),
          code,
          name,
          excelCol: excelCol || name,
          type,
          inCalc
        });
      }

      saveAppState();
      closeQuestionModal();
      renderQuestionsView();
    }

    function editQuestion(id) { openQuestionModal(id); }
    function deleteQuestion(id) {
      if (confirm('確定要刪除此題目嗎？該題後續將不再納入統計。')) {
        const questions = getActiveQuestions().filter(x => x.id !== id);
        if (appState.activeMode === 'CSI') appState.csi_questions = questions;
        else appState.ssi_questions = questions;
        saveAppState();
        renderQuestionsView();
      }
    }

    function renderExecutiveReport() {
      const month = document.getElementById('filter-month').value;
      const data = getFilteredData(month);
      const stat = calculateStatistics(data);

      document.getElementById('report-header-title').innerText = `${month} 顧客滿意度整體績效報告`;
      document.getElementById('report-print-date').innerText = new Date().toISOString().slice(0, 10).replace(/-/g, '/');
      document.getElementById('report-total-sample').innerText = stat.totalResponses;
      document.getElementById('report-kpi-avg').innerText = stat.overallAvg.toFixed(2);
      document.getElementById('report-kpi-high').innerText = `${stat.overallHighRate.toFixed(1)}%`;
      document.getElementById('report-kpi-pos').innerText = `${stat.overallPosRate.toFixed(1)}%`;
      document.getElementById('report-kpi-low').innerText = `${stat.overallLowRate.toFixed(1)}%`;

      const chkBoxes = document.getElementById('report-checklist-boxes');
      chkBoxes.innerHTML = '';
      getActiveQuestions().filter(q => q.type === 'yesno').forEach(q => {
        const rate = stat.questionStats[q.code]?.yesRate || 0;
        const div = document.createElement('div');
        div.className = 'border border-slate-200 rounded p-2 bg-slate-50';
        div.innerHTML = `<div class="text-[11px] text-slate-500">${q.code} ${q.name}</div><div class="text-base font-bold text-emerald-600 mt-0.5">${rate.toFixed(1)}%</div>`;
        chkBoxes.appendChild(div);
      });

      const ratingQuestions = getActiveQuestions().filter(q => q.type === 'rating');
      const rankedQuestions = ratingQuestions.map(q => {
        return { name: q.name, avg: stat.questionStats[q.code]?.avg || 0, low: stat.questionStats[q.code]?.lowRate || 0 };
      }).sort((a, b) => b.avg - a.avg);

      const topUl = document.getElementById('report-top-questions');
      const worstUl = document.getElementById('report-worst-questions');
      topUl.innerHTML = '';
      worstUl.innerHTML = '';

      rankedQuestions.slice(0, 2).forEach(q => {
        const li = document.createElement('li');
        li.className = 'flex justify-between items-center text-slate-700';
        li.innerHTML = `<span>✔ ${q.name}</span><strong class="text-emerald-600">${q.avg.toFixed(2)} 分</strong>`;
        topUl.appendChild(li);
      });

      rankedQuestions.slice(-2).reverse().forEach(q => {
        const li = document.createElement('li');
        li.className = 'flex justify-between items-center text-slate-700';
        li.innerHTML = `<span>▲ ${q.name}</span><strong class="text-rose-600">${q.avg.toFixed(2)} 分 (低分率: ${q.low.toFixed(1)}%)</strong>`;
        worstUl.appendChild(li);
      });

      const branches = getActiveBranches();
      const records = getActiveRecords().filter(r => r.month === month && (!r.status || r.status === '成功'));
      const bList = branches.map(b => {
        const bData = records.filter(r => r.branch === b);
        const bStat = calculateStatistics(bData);
        return { name: b, resp: bData.length, avg: bStat.overallAvg, high: bStat.overallHighRate, low: bStat.overallLowRate };
      }).sort((a, b) => b.avg - a.avg);

      const reportBranchBody = document.getElementById('report-branch-rows');
      reportBranchBody.innerHTML = '';

      bList.forEach((b, idx) => {
        const tr = document.createElement('tr');
        const isAlert = b.low > 10 || b.avg < 8.5;
        tr.innerHTML = `
          <td class="p-2 border border-slate-200 text-center font-bold">${idx + 1}</td>
          <td class="p-2 border border-slate-200 font-semibold">${b.name}</td>
          <td class="p-2 border border-slate-200 text-right">${b.resp}</td>
          <td class="p-2 border border-slate-200 text-right font-black">${b.avg.toFixed(2)}</td>
          <td class="p-2 border border-slate-200 text-right text-emerald-600 font-semibold">${b.high.toFixed(1)}%</td>
          <td class="p-2 border border-slate-200 text-right ${b.low > 10 ? 'text-rose-600 font-bold' : ''}">${b.low.toFixed(1)}%</td>
          <td class="p-2 border border-slate-200 text-center font-bold ${isAlert ? 'text-rose-600' : 'text-emerald-600'}">
            ${isAlert ? '● 需改善' : '○ 良好'}
          </td>
        `;
        reportBranchBody.appendChild(tr);
      });
    }

    function exportItemAnalysisExcel() {
      const month = document.getElementById('filter-month').value;
      const data = getFilteredData(month);
      const stat = calculateStatistics(data);
      const questions = getActiveQuestions();

      const exportRows = questions.map(q => {
        const qStat = stat.questionStats[q.code] || {};
        return {
          '月份': month,
          '題目代碼': q.code,
          '題目名稱': q.name,
          '題目類型': q.type === 'rating' ? '1-10分題' : '檢核題',
          '有效樣本': qStat.validCount || 0,
          '平均得分/達成率': q.type === 'rating' ? Number((qStat.avg || 0).toFixed(2)) : `${(qStat.yesRate || 0).toFixed(1)}%`,
          '中位數': q.type === 'rating' ? qStat.median : 'N/A',
          '眾數': q.type === 'rating' ? qStat.mode : 'N/A',
          '9-10分率': q.type === 'rating' ? `${qStat.highRate.toFixed(1)}%` : 'N/A',
          '1-6分率': q.type === 'rating' ? `${qStat.lowRate.toFixed(1)}%` : 'N/A'
        };
      });

      const ws = XLSX.utils.json_to_sheet(exportRows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, '個別題目分析');
      XLSX.writeFile(wb, `TOYOTA_${appState.activeMode}_題目統計_${month.replace('/', '')}.xlsx`);
    }

    function exportBranchCompareExcel() {
      const month = document.getElementById('filter-month').value;
      const branches = getActiveBranches();
      const records = getActiveRecords().filter(r => r.month === month && (!r.status || r.status === '成功'));

      const exportRows = branches.map(b => {
        const bData = records.filter(r => r.branch === b);
        const bStat = calculateStatistics(bData);
        return {
          '月份': month,
          '據點代碼/名稱': b,
          '有效樣本': bData.length,
          '平均滿意度': Number(bStat.overallAvg.toFixed(2)),
          '高分率(9-10)': `${bStat.overallHighRate.toFixed(1)}%`,
          '低分率(1-6)': `${bStat.overallLowRate.toFixed(1)}%`
        };
      });

      const ws = XLSX.utils.json_to_sheet(exportRows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, '據點比較績效表');
      XLSX.writeFile(wb, `TOYOTA_${appState.activeMode}_據點統計_${month.replace('/', '')}.xlsx`);
    }

    function exportChecklistExcel() {
      const month = document.getElementById('filter-month').value;
      const branches = getActiveBranches();
      const yesNoQuestions = getActiveQuestions().filter(q => q.type === 'yesno');
      const records = getActiveRecords().filter(r => r.month === month && (!r.status || r.status === '成功'));

      const exportRows = branches.map(b => {
        const bData = records.filter(r => r.branch === b);
        const bStat = calculateStatistics(bData);
        const row = {
          '月份': month,
          '據點代碼': b,
          '成功填答數': bData.length
        };
        yesNoQuestions.forEach(q => {
          row[`${q.code}_${q.name}_達成率`] = `${(bStat.questionStats[q.code]?.yesRate || 0).toFixed(1)}%`;
        });
        return row;
      });

      const ws = XLSX.utils.json_to_sheet(exportRows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, '檢核題達成率報表');
      XLSX.writeFile(wb, `TOYOTA_${appState.activeMode}_檢核題推移統計_${month.replace('/', '')}.xlsx`);
    }

    function backupDatabaseToFile() {
      const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(appState, null, 2));
      const a = document.createElement('a');
      a.href = dataStr;
      a.download = `TOYOTA_滿意度系統備份_${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      a.remove();
    }

    function restoreDatabaseFromFile(e) {
      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = function(evt) {
        try {
          const restored = JSON.parse(evt.target.result);
          if (restored && restored.ssi_questions) {
            appState = restored;
            saveAppState();
            alert('資料庫已成功還原！');
            location.reload();
          } else {
            alert('檔案格式不符合系統備份標準！');
          }
        } catch (err) {
          alert('讀取失敗，請確認檔案格式！');
        }
      };
      reader.readAsText(file);
    }

    function downloadSampleTemplate() {
      const isSsi = appState.activeMode === 'SSI';
      let sampleData = [];

      if (isSsi) {
        sampleData = [
          { '編號': 'NO-202607-001', '調查來源': '問卷', '狀態': '成功', '據點代碼': 'E01', '調查日期': '2026/07/01', 'S1.整體滿意度)': 10, 'S2.保險說明': '全部皆有', 'S3.新車大禮包': '有', 'S4.業代迎賓接待': 10, 'S5.廠人員介紹': '有', 'S6.所主管介紹': '有' },
          { '編號': 'NO-202607-002', '調查來源': '電訪', '狀態': '成功', '據點代碼': 'E01', '調查日期': '2026/07/02', 'S1.整體滿意度)': 10, 'S2.保險說明': '沒有讓您選擇保險公司', 'S3.新車大禮包': '沒有', 'S4.業代迎賓接待': 10, 'S5.廠人員介紹': '沒有', 'S6.所主管介紹': '沒有' }
        ];
      } else {
        sampleData = [
          { '編號': 'CSI-202607-001', '調查來源': '電訪', '狀態': '成功', '據點代碼': '永康服務廠', '調查日期': '2026/07/15', '整體滿意度': 10, '接待服務': 10, '服務態度': 10, '費用透明度': 9, '維修品質': 10, '交車清潔': 10, '一次修復': '是' }
        ];
      }

      const ws = XLSX.utils.json_to_sheet(sampleData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, '標準調查範本');
      XLSX.writeFile(wb, `TOYOTA_${appState.activeMode}_標準匯入範本.xlsx`);
    }

    let parsedWorkbook = null;
    let parsedRawRows = [];
    const dropZone = document.getElementById('drop-zone');
    const fileUploader = document.getElementById('file-uploader');

    dropZone.addEventListener('click', () => fileUploader.click());
    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('border-red-500'); });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('border-red-500'));
    dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropZone.classList.remove('border-red-500');
      if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
    });

    function handleFileSelect(e) {
      if (e.target.files.length) handleFile(e.target.files[0]);
    }

    function handleFile(file) {
      const reader = new FileReader();
      reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        parsedWorkbook = XLSX.read(data, { type: 'array' });

        document.getElementById('preview-filename').innerText = file.name;
        const sheetSelector = document.getElementById('sheet-selector');
        sheetSelector.innerHTML = '';
        parsedWorkbook.SheetNames.forEach(sName => {
          const opt = document.createElement('option');
          opt.value = sName;
          opt.textContent = sName;
          sheetSelector.appendChild(opt);
        });

        changeSheet();
      };
      reader.readAsArrayBuffer(file);
    }

    function changeSheet() {
      const sName = document.getElementById('sheet-selector').value;
      const ws = parsedWorkbook.Sheets[sName];
      parsedRawRows = XLSX.utils.sheet_to_json(ws, { defval: '' });

      document.getElementById('preview-rowcount').innerText = parsedRawRows.length;
      document.getElementById('import-preview-box').classList.remove('hidden');

      renderColumnMapper();
      renderPreviewTable();
    }

    function renderColumnMapper() {
      if (parsedRawRows.length === 0) return;
      const headers = Object.keys(parsedRawRows[0]);
      const questions = getActiveQuestions();
      const grid = document.getElementById('column-mapper-grid');
      grid.innerHTML = '';

      const coreFields = [
        { key: 'no', label: '調查編號 (第一欄)', match: ['編號', '序號', 'ID', 'no'] },
        { key: 'date', label: '調查日期欄位', match: ['日期', '調查日期'] },
        { key: 'branch', label: '據點代碼/名稱', match: ['據點', '據點代碼'] },
        { key: 'channel', label: '調查來源', match: ['來源', '調查來源'] },
        { key: 'status', label: '調查狀態', match: ['狀態'] }
      ];

      coreFields.forEach(f => {
        const div = document.createElement('div');
        div.innerHTML = `
          <label class="block font-bold text-slate-700 mb-1">${f.label}：</label>
          <select id="map-col-${f.key}" class="w-full border border-slate-300 rounded px-2 py-1.5 bg-white">
            ${headers.map(h => `<option value="${h}" ${f.match.some(m => h.toLowerCase().includes(m.toLowerCase())) ? 'selected' : ''}>${h}</option>`).join('')}
          </select>
        `;
        grid.appendChild(div);
      });

      questions.forEach(q => {
        const div = document.createElement('div');
        div.innerHTML = `
          <label class="block font-bold text-slate-700 mb-1">${q.code} (${q.name})：</label>
          <select id="map-col-${q.code}" class="w-full border border-slate-300 rounded px-2 py-1.5 bg-white">
            ${headers.map(h => `<option value="${h}" ${h === q.excelCol || h.includes(q.code) || h.includes(q.name) ? 'selected' : ''}>${h}</option>`).join('')}
          </select>
        `;
        grid.appendChild(div);
      });
    }

    function renderPreviewTable() {
      if (parsedRawRows.length === 0) return;
      const headers = Object.keys(parsedRawRows[0]);
      const thead = document.getElementById('preview-table-header');
      const tbody = document.getElementById('preview-table-body');

      thead.innerHTML = `<tr>${headers.map(h => `<th class="p-2 border-b border-slate-200">${h}</th>`).join('')}</tr>`;
      tbody.innerHTML = '';

      parsedRawRows.slice(0, 15).forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = headers.map(h => `<td class="p-2 border-b border-slate-100">${row[h]}</td>`).join('');
        tbody.appendChild(tr);
      });
    }

    function executeDataImport(mode) {
      if (parsedRawRows.length === 0) return;

      const noCol = document.getElementById('map-col-no')?.value;
      const dateCol = document.getElementById('map-col-date').value;
      const branchCol = document.getElementById('map-col-branch').value;
      const channelCol = document.getElementById('map-col-channel').value;
      const statusCol = document.getElementById('map-col-status')?.value;
      const questions = getActiveQuestions();

      const newRecords = [];
      const newBranchesSet = new Set(getActiveBranches());

      parsedRawRows.forEach((row, idx) => {
        const rawNo = noCol ? String(row[noCol] || '').trim() : '';
        const rawDate = String(row[dateCol] || '').trim();
        const branchVal = String(row[branchCol] || '').trim();
        const channelVal = String(row[channelCol] || '問卷').trim();
        const statusVal = statusCol ? String(row[statusCol] || '成功').trim() : '成功';

        if (!rawDate || !branchVal) return;

        newBranchesSet.add(branchVal);

        const monthMatch = rawDate.match(/(\d{4})[\/\-](\d{1,2})/);
        const monthVal = monthMatch ? `${monthMatch[1]}/${monthMatch[2].padStart(2, '0')}` : '2026/07';

        const scores = {};
        questions.forEach(q => {
          const mappedCol = document.getElementById(`map-col-${q.code}`)?.value;
          scores[q.code] = mappedCol ? row[mappedCol] : '';
        });

        const recordId = rawNo || `IMP-${Date.now()}-${idx}`;
        newRecords.push({
          no: rawNo || recordId,
          id: recordId,
          month: monthVal,
          date: rawDate,
          branch: branchVal,
          channel: channelVal,
          status: statusVal,
          scores: scores
        });
      });

      if (newRecords.length === 0) {
        alert('匯入失敗：未能成功解析符合標準的調查紀錄！');
        return;
      }

      if (appState.activeMode === 'SSI') {
        appState.ssi_branches = Array.from(newBranchesSet);
        if (mode === 'replace') {
          const importMonth = newRecords[0].month;
          appState.ssi_records = appState.ssi_records.filter(r => r.month !== importMonth).concat(newRecords);
        } else {
          appState.ssi_records = appState.ssi_records.concat(newRecords);
        }
      } else {
        appState.csi_branches = Array.from(newBranchesSet);
        if (mode === 'replace') {
          const importMonth = newRecords[0].month;
          appState.csi_records = appState.csi_records.filter(r => r.month !== importMonth).concat(newRecords);
        } else {
          appState.csi_records = appState.csi_records.concat(newRecords);
        }
      }

      saveAppState();
      initMonthFilterDropdown();
      initBranchFilterDropdown();
      initPurgeOptions();
      
      document.getElementById('filter-month').value = newRecords[0].month;

      alert(`成功匯入 ${newRecords.length} 筆資料至 ${appState.activeMode === 'SSI' ? '購車調查' : '保修調查'} 資料庫！已自動切換至 ${newRecords[0].month} 月份。`);
      navigateView('dashboard');
    }

    window.addEventListener('DOMContentLoaded', () => {
      loadAppState();
      switchTrackMode(appState.activeMode || 'SSI');
    });
  </script>
</body>
</html>
