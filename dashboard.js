// dashboard.js - 10가지 통계 분석 알고리즘 & Chart.js 시각화 엔진

let rawData = [];
let charts = {};

// 1. 초기 실행 & 센서 데이터 로딩 (임베드 데이터 우선, 없으면 fetch 또는 파일 업로드 대기)
window.addEventListener("DOMContentLoaded", async () => {
  if (window.EMBEDDED_SENSOR_DATA && Array.isArray(window.EMBEDDED_SENSOR_DATA) && window.EMBEDDED_SENSOR_DATA.length > 0) {
    rawData = window.EMBEDDED_SENSOR_DATA;
    initApp();
    showToast(`센서 데이터 ${rawData.length}건 즉시 로드 완료!`);
    return;
  }

  showToast("데이터 분석 로드 중...");
  try {
    const res = await fetch("sensor_data.json");
    if (!res.ok) throw new Error("Fetch failed");
    rawData = await res.json();
    initApp();
  } catch (e) {
    console.warn("로컬 json fetch 불가, 내장 샘플 데이터 또는 엑셀 업로드 모드로 동작", e);
    document.getElementById("dataSummaryMeta").innerText = "상단 '엑셀 파일 열기'를 통해 센서 스무딩.xlsx를 로드해 주세요.";
  }
});

// 파일 업로드 핸들러 (사용자가 파일 드래그/선택 시)
document.getElementById("filePicker").addEventListener("change", function(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
      
      // 파싱
      const header = json[0];
      const parsed = [];
      for (let i = 1; i < json.length; i++) {
        const row = json[i];
        if (!row || row.length < 5) continue;
        parsed.push({
          time: row[0] ? formatExcelTime(row[0]) : `${i}분`,
          minute: Number(row[1]) || (i - 1),
          hour: Number(row[2]) || Math.round((i - 1)/60*1000)/1000,
          tempRaw: Number(row[3]) || 0,
          tempTrue: Number(row[4]) || 0,
          vibRms: Number(row[5]) || 0,
          pressure: Number(row[6]) || 0,
          event: row[7] || "정상"
        });
      }
      if (parsed.length > 0) {
        rawData = parsed;
        initApp();
        showToast(`엑셀 파일 로드 완료: ${parsed.length}개 관측치`);
      }
    } catch (err) {
      alert("엑셀 파일 파싱 오류: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
});

function formatExcelTime(serial) {
  if (typeof serial === "string" && serial.includes("-")) return serial;
  if (typeof serial === "number") {
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    const fractional_day = serial - Math.floor(serial) + 0.0000001;
    let total_seconds = Math.floor(86400 * fractional_day);
    const seconds = total_seconds % 60;
    total_seconds -= seconds;
    const hours = Math.floor(total_seconds / (60 * 60));
    const minutes = Math.floor(total_seconds / 60) % 60;
    return `${date_info.getFullYear()}-${String(date_info.getMonth()+1).padStart(2,'0')}-${String(date_info.getDate()).padStart(2,'0')} ${String(hours).padStart(2,'0')}:${String(minutes).padStart(2,'0')}`;
  }
  return String(serial);
}

// ==========================================
// 10가지 통계 분석 엔진 (Statistical Methods)
// ==========================================

// Method 1: 기술 통계량 (Descriptive Statistics: Mean, Var, SD, Median, IQR, Skew, Kurt)
function calcDescriptive(arr) {
  const n = arr.length;
  if (n === 0) return null;
  const sum = arr.reduce((a, b) => a + b, 0);
  const mean = sum / n;
  
  const sorted = [...arr].sort((a, b) => a - b);
  const median = n % 2 === 0 ? (sorted[n/2 - 1] + sorted[n/2]) / 2 : sorted[Math.floor(n/2)];
  
  const q1 = sorted[Math.floor(n * 0.25)];
  const q3 = sorted[Math.floor(n * 0.75)];
  const iqr = q3 - q1;
  const min = sorted[0];
  const max = sorted[n - 1];

  let sqDiffSum = 0;
  let cubeDiffSum = 0;
  let fourthDiffSum = 0;

  for (let x of arr) {
    const diff = x - mean;
    sqDiffSum += diff * diff;
    cubeDiffSum += Math.pow(diff, 3);
    fourthDiffSum += Math.pow(diff, 4);
  }

  const variance = sqDiffSum / (n - 1);
  const sd = Math.sqrt(variance);

  // Skewness (왜도)
  const skew = (cubeDiffSum / n) / Math.pow(variance, 1.5);
  // Kurtosis (첨도 - 초과첨도: 정규분포=0)
  const kurt = (fourthDiffSum / n) / Math.pow(variance, 2) - 3;

  return { n, mean, sd, variance, median, q1, q3, iqr, min, max, skew, kurt };
}

// Method 2: 확률 밀도 히스토그램 (Histogram & Gaussian Fit)
function calcHistogram(arr, binCount = 30) {
  const min = Math.min(...arr);
  const max = Math.max(...arr);
  const binWidth = (max - min) / binCount;
  const bins = Array(binCount).fill(0);
  const labels = [];

  for (let i = 0; i < binCount; i++) {
    const start = min + i * binWidth;
    const mid = start + binWidth / 2;
    labels.push(mid.toFixed(2));
  }

  for (let x of arr) {
    let idx = Math.floor((x - min) / binWidth);
    if (idx >= binCount) idx = binCount - 1;
    bins[idx]++;
  }

  // Gaussian Curve Fit
  const stats = calcDescriptive(arr);
  const gaussian = labels.map(mid => {
    const x = parseFloat(mid);
    const exponent = -0.5 * Math.pow((x - stats.mean) / stats.sd, 2);
    const pdf = (1 / (stats.sd * Math.sqrt(2 * Math.PI))) * Math.exp(exponent);
    return pdf * arr.length * binWidth;
  });

  return { labels, counts: bins, gaussian };
}

// Method 3: 이동평균 (Moving Average) & Savitzky-Golay 다항식 필터
function calcMovingAverage(arr, windowSize) {
  const half = Math.floor(windowSize / 2);
  const result = [];
  for (let i = 0; i < arr.length; i++) {
    let start = Math.max(0, i - half);
    let end = Math.min(arr.length, i + half + 1);
    let sum = 0;
    for (let j = start; j < end; j++) sum += arr[j];
    result.push(Number((sum / (end - start)).toFixed(3)));
  }
  return result;
}

function calcEMA(arr, span = 15) {
  const alpha = 2 / (span + 1);
  const res = [arr[0]];
  for (let i = 1; i < arr.length; i++) {
    res.push(Number((alpha * arr[i] + (1 - alpha) * res[i - 1]).toFixed(3)));
  }
  return res;
}

// Savitzky-Golay Filter (Window: 21, Order: 2) Convolved Weights
function calcSavitzkyGolay(arr, windowSize = 21, polyOrder = 2) {
  // Precomputed SG convolution coefficients for m=10 (window 21), order 2/3
  // c_i = (3m^2 + 3m - 1 - 5i^2) / norm
  const m = Math.floor(windowSize / 2);
  const norm = (2 * m - 1) * (2 * m + 1) * (2 * m + 3) / 3;
  const weights = [];
  let sumW = 0;
  for (let i = -m; i <= m; i++) {
    const w = (3 * m * m + 3 * m - 1 - 5 * i * i);
    weights.push(w);
    sumW += w;
  }

  const result = [];
  for (let i = 0; i < arr.length; i++) {
    let valSum = 0;
    let actualWeightSum = 0;
    for (let k = -m; k <= m; k++) {
      const idx = i + k;
      if (idx >= 0 && idx < arr.length) {
        valSum += arr[idx] * weights[k + m];
        actualWeightSum += weights[k + m];
      }
    }
    result.push(Number((valSum / actualWeightSum).toFixed(3)));
  }
  return result;
}

// Method 4: Peak 포락선 & 볼린저 신뢰대 (Envelope & Bollinger Bands)
function calcEnvelopes(arr, windowSize = 30) {
  const upper = [];
  const lower = [];
  const half = Math.floor(windowSize / 2);
  for (let i = 0; i < arr.length; i++) {
    let start = Math.max(0, i - half);
    let end = Math.min(arr.length, i + half + 1);
    let slice = arr.slice(start, end);
    upper.push(Math.max(...slice));
    lower.push(Math.min(...slice));
  }
  return { upper, lower };
}

function calcBollingerBands(arr, maArr, sdMul = 2.5) {
  const upper = [];
  const lower = [];
  const win = 20;
  for (let i = 0; i < arr.length; i++) {
    let start = Math.max(0, i - win);
    let slice = arr.slice(start, i + 1);
    let m = maArr[i];
    let sq = slice.reduce((acc, v) => acc + Math.pow(v - m, 2), 0) / slice.length;
    let localSd = Math.sqrt(sq);
    upper.push(Number((m + sdMul * localSd).toFixed(3)));
    lower.push(Number((m - sdMul * localSd).toFixed(3)));
  }
  return { upper, lower };
}

// Method 5: 이상치 탐지 (Hampel Filter - Median Absolute Deviation MAD)
function detectAnomalies(arr, trueArr, windowSize = 15, threshold = 3.0) {
  const half = Math.floor(windowSize / 2);
  const anomalies = [];
  
  for (let i = 0; i < arr.length; i++) {
    let start = Math.max(0, i - half);
    let end = Math.min(arr.length, i + half + 1);
    let slice = arr.slice(start, end).sort((a, b) => a - b);
    let med = slice[Math.floor(slice.length / 2)];
    
    let diffs = slice.map(x => Math.abs(x - med)).sort((a, b) => a - b);
    let mad = diffs[Math.floor(diffs.length / 2)];
    let scale = 1.4826 * mad; // Normal distribution consistency factor
    
    let score = scale > 0.001 ? Math.abs(arr[i] - med) / scale : 0;
    if (score > threshold) {
      anomalies.push({
        index: i,
        score: Number(score.toFixed(2)),
        val: arr[i],
        diff: Number(Math.abs(arr[i] - (trueArr[i] || med)).toFixed(2))
      });
    }
  }
  return anomalies;
}

// Method 6: 잔차 및 오차 평가 (RMSE, MAE, R²)
function calcResidualMetrics(predArr, trueArr) {
  let sqSum = 0;
  let absSum = 0;
  let maxAbs = 0;
  let trueMean = trueArr.reduce((a, b) => a + b, 0) / trueArr.length;
  let totSum = 0;

  for (let i = 0; i < predArr.length; i++) {
    let err = predArr[i] - trueArr[i];
    sqSum += err * err;
    let absErr = Math.abs(err);
    absSum += absErr;
    if (absErr > maxAbs) maxAbs = absErr;
    totSum += Math.pow(trueArr[i] - trueMean, 2);
  }

  const rmse = Math.sqrt(sqSum / predArr.length);
  const mae = absSum / predArr.length;
  const r2 = 1 - (sqSum / totSum);

  return {
    rmse: Number(rmse.toFixed(3)),
    mae: Number(mae.toFixed(3)),
    r2: Number(r2.toFixed(4)),
    maxAbs: Number(maxAbs.toFixed(3))
  };
}

// Method 7: 피어슨 상관계수 행렬 (Pearson Correlation)
function calcPearson(x, y) {
  const n = x.length;
  const mx = x.reduce((a, b) => a + b, 0) / n;
  const my = y.reduce((a, b) => a + b, 0) / n;
  let num = 0;
  let den1 = 0;
  let den2 = 0;
  for (let i = 0; i < n; i++) {
    let dx = x[i] - mx;
    let dy = y[i] - my;
    num += dx * dy;
    den1 += dx * dx;
    den2 += dy * dy;
  }
  return Number((num / (Math.sqrt(den1) * Math.sqrt(den2))).toFixed(3));
}

// Method 8: 시계열 자기상관함수 (ACF: Autocorrelation Function)
function calcACF(arr, maxLag = 60) {
  const n = arr.length;
  const mean = arr.reduce((a, b) => a + b, 0) / n;
  let c0 = 0;
  for (let i = 0; i < n; i++) {
    c0 += Math.pow(arr[i] - mean, 2);
  }
  
  const acf = [];
  const lags = [];
  for (let k = 1; k <= maxLag; k++) {
    let ck = 0;
    for (let i = 0; i < n - k; i++) {
      ck += (arr[i] - mean) * (arr[i + k] - mean);
    }
    acf.push(Number((ck / c0).toFixed(3)));
    lags.push(`${k}분`);
  }
  return { lags, acf };
}

// Method 9: 고속 푸리에 변환 에너지 스펙트럼 (FFT Periodogram 근사)
function calcFFTSpectrum(arr, maxFreq = 64) {
  const N = arr.length;
  const mean = arr.reduce((a, b) => a + b, 0) / N;
  const detrended = arr.map(x => x - mean);
  const spectrum = [];
  const freqs = [];

  for (let k = 1; k <= maxFreq; k++) {
    let re = 0;
    let im = 0;
    for (let t = 0; t < N; t++) {
      let angle = (2 * Math.PI * k * t) / N;
      re += detrended[t] * Math.cos(angle);
      im -= detrended[t] * Math.sin(angle);
    }
    let power = (re * re + im * im) / N;
    spectrum.push(Number(power.toFixed(2)));
    let periodMinutes = Math.round(N / k);
    freqs.push(`${periodMinutes}분 주기`);
  }
  return { freqs, spectrum };
}

// Method 10: 이벤트 그룹별 ANOVA 및 평균/분산 통계
function calcEventGroupStats(data) {
  const groups = {};
  for (let d of data) {
    const ev = d.event || "정상";
    if (!groups[ev]) {
      groups[ev] = { count: 0, temps: [], vibs: [], press: [] };
    }
    groups[ev].count++;
    groups[ev].temps.push(d.tempRaw);
    groups[ev].vibs.push(d.vibRms);
    groups[ev].press.push(d.pressure);
  }

  const result = [];
  for (let key in groups) {
    const g = groups[key];
    const avg = arr => arr.reduce((a, b) => a + b, 0) / arr.length;
    result.push({
      event: key,
      count: g.count,
      tempMean: Number(avg(g.temps).toFixed(2)),
      vibMean: Number(avg(g.vibs).toFixed(3)),
      pressMean: Number(avg(g.press).toFixed(2))
    });
  }
  return result.sort((a, b) => b.count - a.count);
}

// ==========================================
// 대시보드 렌더링 & UI 제어
// ==========================================

function initApp() {
  if (!rawData || rawData.length === 0) return;

  const count = rawData.length;
  document.getElementById("kpiTotalSamples").innerText = count.toLocaleString() + " 건";
  document.getElementById("dataSummaryMeta").innerText = `측정기간: 24시간 | 관측주기: 1분 | 총 ${count.toLocaleString()} 데이터 로드됨`;

  const tempsRaw = rawData.map(d => d.tempRaw);
  const tempsTrue = rawData.map(d => d.tempTrue);
  const vibs = rawData.map(d => d.vibRms);
  const pressures = rawData.map(d => d.pressure);

  // 1. Descriptive KPIs
  const tempStats = calcDescriptive(tempsRaw);
  document.getElementById("kpiTempMean").innerText = `${tempStats.mean.toFixed(2)} °C`;
  document.getElementById("kpiTempSd").innerText = `± ${tempStats.sd.toFixed(2)}°C (최대: ${tempStats.max.toFixed(1)}°C, 최소: ${tempStats.min.toFixed(1)}°C)`;

  // 2. Render Descriptive Tables
  renderDescriptiveTable(tempStats, calcDescriptive(tempsTrue), calcDescriptive(vibs), calcDescriptive(pressures));

  // 3. Render Smoothing Chart
  updateSmoothingChart();

  // 4. Render Histogram
  renderHistogramChart();

  // 5. Render Envelopes
  renderEnvelopeChart();

  // 6. Render Anomalies
  renderAnomalySection();

  // 7. Render Error & Residuals
  renderResidualsSection();

  // 8. Render Correlation Matrix & Scatter
  renderCorrelationSection();

  // 9. Render ACF & FFT
  renderACFChart();
  renderFFTChart();

  // 10. Render Event Groups
  renderEventGroupSection();

  // Raw Table & Detail tabs
  renderRawTable();
  renderMultivariateChart();
  renderDetailedTables();
}

// Method 1: 테이블 렌더링
function renderDescriptiveTable(tRaw, tTrue, vib, pres) {
  const tbody = document.querySelector("#tableDescriptive tbody");
  tbody.innerHTML = `
    <tr>
      <td><strong>온도 센서 원시값(°C)</strong></td>
      <td>${tRaw.mean.toFixed(2)}</td>
      <td>${tRaw.sd.toFixed(2)}</td>
      <td>${tRaw.median.toFixed(2)}</td>
      <td>${tRaw.iqr.toFixed(2)}</td>
      <td><span class="${tRaw.skew > 0.5 ? 'tag tag-overheat' : ''}">${tRaw.skew.toFixed(2)}</span></td>
      <td>${tRaw.kurt.toFixed(2)}</td>
    </tr>
    <tr>
      <td><strong>온도 참값·비교용(°C)</strong></td>
      <td>${tTrue.mean.toFixed(2)}</td>
      <td>${tTrue.sd.toFixed(2)}</td>
      <td>${tTrue.median.toFixed(2)}</td>
      <td>${tTrue.iqr.toFixed(2)}</td>
      <td>${tTrue.skew.toFixed(2)}</td>
      <td>${tTrue.kurt.toFixed(2)}</td>
    </tr>
    <tr>
      <td><strong>진동 RMS (mm/s)</strong></td>
      <td>${vib.mean.toFixed(3)}</td>
      <td>${vib.sd.toFixed(3)}</td>
      <td>${vib.median.toFixed(3)}</td>
      <td>${vib.iqr.toFixed(3)}</td>
      <td>${vib.skew.toFixed(2)}</td>
      <td>${vib.kurt.toFixed(2)}</td>
    </tr>
    <tr>
      <td><strong>압력 (kPa)</strong></td>
      <td>${pres.mean.toFixed(2)}</td>
      <td>${pres.sd.toFixed(2)}</td>
      <td>${pres.median.toFixed(2)}</td>
      <td>${pres.iqr.toFixed(2)}</td>
      <td>${pres.skew.toFixed(2)}</td>
      <td>${pres.kurt.toFixed(2)}</td>
    </tr>
  `;
}

// Method 3: 스무딩 메인 차트 렌더링
function updateSmoothingChart() {
  const maWin = parseInt(document.getElementById("selMaWindow").value);
  const sgOrder = parseInt(document.getElementById("selSgOrder").value);

  const raw = rawData.map(d => d.tempRaw);
  const truth = rawData.map(d => d.tempTrue);
  const labels = rawData.map(d => d.time.split(" ")[1] || `${d.minute}m`);

  const ma = calcMovingAverage(raw, maWin);
  const sg = calcSavitzkyGolay(raw, sgOrder === 2 ? 21 : 31, sgOrder);

  // Update KPI noise reduction & RMSE
  const errRaw = calcResidualMetrics(raw, truth);
  const errSg = calcResidualMetrics(sg, truth);
  const reduction = ((1 - errSg.rmse / errRaw.rmse) * 100).toFixed(1);
  document.getElementById("kpiNoiseReduction").innerText = `${reduction}%`;
  document.getElementById("kpiRmseVal").innerText = `SG-RMSE: ${errSg.rmse}°C (원시: ${errRaw.rmse}°C)`;

  renderChart("chartSmoothing", {
    type: "line",
    data: {
      labels: labels,
      datasets: [
        {
          label: "센서 원시값 (Raw + Noise)",
          data: raw,
          borderColor: "rgba(148, 163, 184, 0.4)",
          backgroundColor: "transparent",
          borderWidth: 1,
          pointRadius: 0,
        },
        {
          label: `단순 이동평균 (MA ${maWin}분 창)`,
          data: ma,
          borderColor: "#38bdf8",
          backgroundColor: "transparent",
          borderWidth: 2,
          pointRadius: 0,
        },
        {
          label: `Savitzky-Golay 다항식 필터`,
          data: sg,
          borderColor: "#f59e0b",
          backgroundColor: "transparent",
          borderWidth: 2.2,
          pointRadius: 0,
        },
        {
          label: "온도 참값 기준 (True)",
          data: truth,
          borderColor: "#10b981",
          borderDash: [5, 5],
          backgroundColor: "transparent",
          borderWidth: 2,
          pointRadius: 0,
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { labels: { color: "#cbd5e1", font: { size: 11 } } },
        tooltip: {
          backgroundColor: "#0f172a",
          borderColor: "#334155",
          borderWidth: 1
        }
      },
      scales: {
        x: { ticks: { color: "#64748b", maxTicksLimit: 16 }, grid: { color: "#1e293b" } },
        y: { title: { display: true, text: "온도 (°C)", color: "#94a3b8" }, ticks: { color: "#64748b" }, grid: { color: "#1e293b" } }
      }
    }
  });
}

// Method 2: 히스토그램 차트
function renderHistogramChart() {
  const sel = document.getElementById("selHistVar").value;
  const arr = rawData.map(d => d[sel]);
  const hist = calcHistogram(arr, 25);

  renderChart("chartHistogram", {
    type: "bar",
    data: {
      labels: hist.labels,
      datasets: [
        {
          type: "bar",
          label: "실제 빈도수 (Counts)",
          data: hist.counts,
          backgroundColor: "rgba(56, 189, 248, 0.45)",
          borderColor: "#38bdf8",
          borderWidth: 1,
          borderRadius: 4
        },
        {
          type: "line",
          label: "정규분포 이론곡선 (Gaussian Fit)",
          data: hist.gaussian,
          borderColor: "#f43f5e",
          borderWidth: 2,
          pointRadius: 0,
          fill: false
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: "#cbd5e1" } }
      },
      scales: {
        x: { ticks: { color: "#64748b" }, grid: { color: "#1e293b" } },
        y: { ticks: { color: "#64748b" }, grid: { color: "#1e293b" } }
      }
    }
  });
}

// Method 4: 포락선 및 신뢰대 차트
function renderEnvelopeChart() {
  const raw = rawData.map(d => d.tempRaw);
  const labels = rawData.map(d => d.time.split(" ")[1] || `${d.minute}m`);
  const env = calcEnvelopes(raw, 30);
  const ma = calcMovingAverage(raw, 20);
  const bands = calcBollingerBands(raw, ma, 2.5);

  renderChart("chartEnvelope", {
    type: "line",
    data: {
      labels: labels,
      datasets: [
        {
          label: "상단 포락선 (Peak Max)",
          data: env.upper,
          borderColor: "rgba(244, 63, 94, 0.8)",
          backgroundColor: "transparent",
          borderWidth: 1.5,
          pointRadius: 0
        },
        {
          label: "상한 신뢰대 (μ + 2.5σ)",
          data: bands.upper,
          borderColor: "rgba(168, 85, 247, 0.6)",
          borderDash: [4, 4],
          backgroundColor: "transparent",
          borderWidth: 1.5,
          pointRadius: 0
        },
        {
          label: "센서 원시값",
          data: raw,
          borderColor: "rgba(56, 189, 248, 0.35)",
          backgroundColor: "transparent",
          borderWidth: 1,
          pointRadius: 0
        },
        {
          label: "하단 포락선 (Trough Min)",
          data: env.lower,
          borderColor: "rgba(16, 185, 129, 0.8)",
          backgroundColor: "transparent",
          borderWidth: 1.5,
          pointRadius: 0
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { labels: { color: "#cbd5e1" } } },
      scales: {
        x: { ticks: { color: "#64748b", maxTicksLimit: 12 }, grid: { color: "#1e293b" } },
        y: { ticks: { color: "#64748b" }, grid: { color: "#1e293b" } }
      }
    }
  });
}

// Method 5: 이상치 탐지 차트
function renderAnomalySection() {
  const raw = rawData.map(d => d.tempRaw);
  const truth = rawData.map(d => d.tempTrue);
  const labels = rawData.map(d => d.time.split(" ")[1] || `${d.minute}m`);

  const anomalies = detectAnomalies(raw, truth, 15, 3.2);
  const anomalyPoints = Array(raw.length).fill(null);
  
  for (let a of anomalies) {
    anomalyPoints[a.index] = a.val;
  }

  const spikeRows = rawData.filter(d => d.event === "스파이크 이상치");
  document.getElementById("badgeSpikeCount").innerText = `Hampel 검출: ${anomalies.length}건 / 라벨 스파이크: ${spikeRows.length}건`;

  renderChart("chartAnomaly", {
    type: "line",
    data: {
      labels: labels,
      datasets: [
        {
          label: "센서 시계열 원시 신호",
          data: raw,
          borderColor: "rgba(148, 163, 184, 0.5)",
          borderWidth: 1,
          pointRadius: 0
        },
        {
          label: "탐지된 이상치 스파이크",
          data: anomalyPoints,
          borderColor: "#f43f5e",
          backgroundColor: "#f43f5e",
          pointRadius: 5,
          pointHoverRadius: 7,
          showLine: false
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { labels: { color: "#cbd5e1" } } },
      scales: {
        x: { ticks: { color: "#64748b", maxTicksLimit: 12 }, grid: { color: "#1e293b" } },
        y: { ticks: { color: "#64748b" }, grid: { color: "#1e293b" } }
      }
    }
  });

  // 세부 이상치 테이블 목록
  const tb = document.querySelector("#tableAnomaliesList tbody");
  if (tb) {
    let html = "";
    for (let a of anomalies) {
      const row = rawData[a.index];
      html += `
        <tr>
          <td>${row.time}</td>
          <td>${row.minute}분</td>
          <td><strong style="color:#fb7185;">${a.val.toFixed(2)}</strong></td>
          <td>${row.tempTrue.toFixed(2)}</td>
          <td>+${a.diff.toFixed(2)}°C</td>
          <td><span class="tag tag-spike">${a.score.toFixed(2)} σ</span></td>
          <td>${row.event}</td>
          <td>Hampel MAD 필터 이탈</td>
        </tr>
      `;
    }
    tb.innerHTML = html;
  }
}

// Method 6: 잔차 및 오차 평가
function renderResidualsSection() {
  const raw = rawData.map(d => d.tempRaw);
  const truth = rawData.map(d => d.tempTrue);
  const ma15 = calcMovingAverage(raw, 15);
  const ema15 = calcEMA(raw, 15);
  const sg = calcSavitzkyGolay(raw, 21, 2);

  const mRaw = calcResidualMetrics(raw, truth);
  const mMa = calcResidualMetrics(ma15, truth);
  const mEma = calcResidualMetrics(ema15, truth);
  const mSg = calcResidualMetrics(sg, truth);

  const tbody = document.querySelector("#tableErrorMetrics tbody");
  tbody.innerHTML = `
    <tr>
      <td>원시 신호 (Raw + Noise)</td>
      <td>${mRaw.rmse}</td>
      <td>${mRaw.mae}</td>
      <td>${mRaw.r2}</td>
      <td>${mRaw.maxAbs}</td>
    </tr>
    <tr>
      <td>단순 이동평균 (SMA 15분)</td>
      <td>${mMa.rmse}</td>
      <td>${mMa.mae}</td>
      <td>${mMa.r2}</td>
      <td>${mMa.maxAbs}</td>
    </tr>
    <tr>
      <td>지수 이동평균 (EMA 15분)</td>
      <td>${mEma.rmse}</td>
      <td>${mEma.mae}</td>
      <td>${mEma.r2}</td>
      <td>${mEma.maxAbs}</td>
    </tr>
    <tr style="background: rgba(56, 189, 248, 0.1);">
      <td><strong>Savitzky-Golay (최적)</strong></td>
      <td><strong style="color:#38bdf8;">${mSg.rmse}</strong></td>
      <td><strong>${mSg.mae}</strong></td>
      <td><strong>${mSg.r2}</strong></td>
      <td><strong>${mSg.maxAbs}</strong></td>
    </tr>
  `;

  // 잔차 분포 시계열 차트
  const residualsSG = sg.map((v, i) => Number((v - truth[i]).toFixed(3)));
  const labels = rawData.map(d => d.time.split(" ")[1] || `${d.minute}m`);

  renderChart("chartResiduals", {
    type: "line",
    data: {
      labels: labels,
      datasets: [
        {
          label: "SG 잔차 (Residual = Filtered - True)",
          data: residualsSG,
          borderColor: "#a855f7",
          backgroundColor: "rgba(168, 85, 247, 0.1)",
          fill: true,
          borderWidth: 1.2,
          pointRadius: 0
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { labels: { color: "#cbd5e1" } } },
      scales: {
        x: { ticks: { color: "#64748b", maxTicksLimit: 12 }, grid: { color: "#1e293b" } },
        y: { title: { display: true, text: "오차 (°C)", color: "#94a3b8" }, ticks: { color: "#64748b" }, grid: { color: "#1e293b" } }
      }
    }
  });
}

// Method 7: 상관분석
function renderCorrelationSection() {
  const tRaw = rawData.map(d => d.tempRaw);
  const tTrue = rawData.map(d => d.tempTrue);
  const vib = rawData.map(d => d.vibRms);
  const pres = rawData.map(d => d.pressure);

  const rRawTrue = calcPearson(tRaw, tTrue);
  const rRawVib = calcPearson(tRaw, vib);
  const rRawPres = calcPearson(tRaw, pres);
  const rTrueVib = calcPearson(tTrue, vib);
  const rTruePres = calcPearson(tTrue, pres);
  const rVibPres = calcPearson(vib, pres);

  const tbody = document.querySelector("#tableCorrelation tbody");
  tbody.innerHTML = `
    <tr>
      <td><strong>온도 원시값</strong></td>
      <td>1.000</td>
      <td>${rRawTrue}</td>
      <td><span class="${rRawVib > 0.4 ? 'tag tag-overheat' : ''}">${rRawVib}</span></td>
      <td>${rRawPres}</td>
    </tr>
    <tr>
      <td><strong>온도 참값</strong></td>
      <td>${rRawTrue}</td>
      <td>1.000</td>
      <td>${rTrueVib}</td>
      <td>${rTruePres}</td>
    </tr>
    <tr>
      <td><strong>진동 RMS</strong></td>
      <td>${rRawVib}</td>
      <td>${rTrueVib}</td>
      <td>1.000</td>
      <td>${rVibPres}</td>
    </tr>
    <tr>
      <td><strong>압력 (kPa)</strong></td>
      <td>${rRawPres}</td>
      <td>${rTruePres}</td>
      <td>${rVibPres}</td>
      <td>1.000</td>
    </tr>
  `;

  // 산점도 (온도 vs 진동)
  const scatterData = rawData.map(d => ({ x: d.tempRaw, y: d.vibRms }));
  renderChart("chartScatterCorr", {
    type: "scatter",
    data: {
      datasets: [{
        label: "온도 (°C) vs 진동 RMS (mm/s)",
        data: scatterData,
        backgroundColor: "rgba(56, 189, 248, 0.5)",
        pointRadius: 2.5
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { labels: { color: "#cbd5e1" } } },
      scales: {
        x: { title: { display: true, text: "온도 (°C)", color: "#94a3b8" }, ticks: { color: "#64748b" }, grid: { color: "#1e293b" } },
        y: { title: { display: true, text: "진동 (mm/s)", color: "#94a3b8" }, ticks: { color: "#64748b" }, grid: { color: "#1e293b" } }
      }
    }
  });

  // 상세 탭 산점도 2개
  const scatterPres = rawData.map(d => ({ x: d.tempRaw, y: d.pressure }));
  renderChart("chartScatterDetail1", {
    type: "scatter",
    data: {
      datasets: [{
        label: "온도 vs 진동 (r = " + rRawVib + ")",
        data: scatterData,
        backgroundColor: "rgba(168, 85, 247, 0.6)",
        pointRadius: 2.5
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { title: { display: true, text: "온도 (°C)" } },
        y: { title: { display: true, text: "진동 (mm/s)" } }
      }
    }
  });

  renderChart("chartScatterDetail2", {
    type: "scatter",
    data: {
      datasets: [{
        label: "온도 vs 압력 (r = " + rRawPres + ")",
        data: scatterPres,
        backgroundColor: "rgba(245, 158, 11, 0.6)",
        pointRadius: 2.5
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { title: { display: true, text: "온도 (°C)" } },
        y: { title: { display: true, text: "압력 (kPa)" } }
      }
    }
  });
}

// Method 8: ACF 차트
function renderACFChart() {
  const raw = rawData.map(d => d.tempRaw);
  const { lags, acf } = calcACF(raw, 50);

  // 95% 신뢰구간선 (± 1.96 / sqrt(N))
  const ci = 1.96 / Math.sqrt(raw.length);
  const ciUpper = Array(lags.length).fill(Number(ci.toFixed(3)));
  const ciLower = Array(lags.length).fill(Number((-ci).toFixed(3)));

  renderChart("chartACF", {
    type: "bar",
    data: {
      labels: lags,
      datasets: [
        {
          label: "자기상관계수 (ACF)",
          data: acf,
          backgroundColor: "#38bdf8",
          borderRadius: 3
        },
        {
          type: "line",
          label: "+95% 신뢰한계",
          data: ciUpper,
          borderColor: "#f43f5e",
          borderDash: [4, 4],
          pointRadius: 0
        },
        {
          type: "line",
          label: "-95% 신뢰한계",
          data: ciLower,
          borderColor: "#f43f5e",
          borderDash: [4, 4],
          pointRadius: 0
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { labels: { color: "#cbd5e1" } } },
      scales: {
        x: { ticks: { color: "#64748b" }, grid: { color: "#1e293b" } },
        y: { min: -0.2, max: 1.0, ticks: { color: "#64748b" }, grid: { color: "#1e293b" } }
      }
    }
  });
}

// Method 9: FFT 스펙트럼
function renderFFTChart() {
  const raw = rawData.map(d => d.tempRaw);
  const { freqs, spectrum } = calcFFTSpectrum(raw, 40);

  renderChart("chartFFT", {
    type: "bar",
    data: {
      labels: freqs,
      datasets: [{
        label: "주파수 파워 스펙트럼 밀도 (Power Spectral Density)",
        data: spectrum,
        backgroundColor: "rgba(129, 140, 248, 0.7)",
        borderRadius: 4
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { labels: { color: "#cbd5e1" } } },
      scales: {
        x: { ticks: { color: "#64748b", maxTicksLimit: 12 }, grid: { color: "#1e293b" } },
        y: { ticks: { color: "#64748b" }, grid: { color: "#1e293b" } }
      }
    }
  });
}

// Method 10: 이벤트 그룹 통계
function renderEventGroupSection() {
  const groups = calcEventGroupStats(rawData);
  const tbody = document.querySelector("#tableEventGroup tbody");
  let html = "";
  for (let g of groups) {
    let tagClass = "tag-normal";
    if (g.event.includes("과열")) tagClass = "tag-overheat";
    else if (g.event.includes("스파이크")) tagClass = "tag-spike";
    else if (g.event.includes("냉각")) tagClass = "tag-cool";
    else if (g.event.includes("부하")) tagClass = "tag-load";

    html += `
      <tr>
        <td><span class="tag ${tagClass}">${g.event}</span></td>
        <td>${g.count}건</td>
        <td><strong>${g.tempMean}°C</strong></td>
        <td>${g.vibMean}</td>
        <td>${g.pressMean}</td>
        <td>${g.event === '정상' ? '정상 작동' : '조치/모니터링 요망'}</td>
      </tr>
    `;
  }
  tbody.innerHTML = html;

  // 그룹 평균 온도 비교 차트
  renderChart("chartEventGroup", {
    type: "bar",
    data: {
      labels: groups.map(g => g.event),
      datasets: [
        {
          label: "이벤트별 평균 온도 (°C)",
          data: groups.map(g => g.tempMean),
          backgroundColor: groups.map(g => {
            if (g.event.includes("과열")) return "#fb7185";
            if (g.event.includes("냉각")) return "#38bdf8";
            if (g.event.includes("부하")) return "#c084fc";
            return "#34d399";
          }),
          borderRadius: 6
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: "#94a3b8" }, grid: { display: false } },
        y: { ticks: { color: "#64748b" }, grid: { color: "#1e293b" } }
      }
    }
  });
}

// 세부 탭들 시각화
function renderMultivariateChart() {
  const temps = rawData.map(d => d.tempRaw);
  const vibs = rawData.map(d => d.vibRms);
  const pres = rawData.map(d => d.pressure);

  const z = (arr) => {
    const s = calcDescriptive(arr);
    return arr.map(x => Number(((x - s.mean) / s.sd).toFixed(2)));
  };

  const zTemp = z(temps);
  const zVib = z(vibs);
  const zPres = z(pres);
  const labels = rawData.map(d => d.time.split(" ")[1] || `${d.minute}m`);

  renderChart("chartStandardized", {
    type: "line",
    data: {
      labels: labels,
      datasets: [
        {
          label: "온도 (표준화 Z)",
          data: zTemp,
          borderColor: "#f43f5e",
          borderWidth: 1.5,
          pointRadius: 0
        },
        {
          label: "진동 RMS (표준화 Z)",
          data: zVib,
          borderColor: "#38bdf8",
          borderWidth: 1.5,
          pointRadius: 0
        },
        {
          label: "압력 (표준화 Z)",
          data: zPres,
          borderColor: "#f59e0b",
          borderWidth: 1.5,
          pointRadius: 0
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { labels: { color: "#cbd5e1" } } },
      scales: {
        x: { ticks: { color: "#64748b", maxTicksLimit: 14 }, grid: { color: "#1e293b" } },
        y: { title: { display: true, text: "표준편차 배수 (Z-Score)", color: "#94a3b8" }, ticks: { color: "#64748b" }, grid: { color: "#1e293b" } }
      }
    }
  });
}

function renderDetailedTables() {
  const keys = [
    { name: "온도 원시값(°C)", prop: "tempRaw" },
    { name: "온도 참값(°C)", prop: "tempTrue" },
    { name: "진동 RMS(mm/s)", prop: "vibRms" },
    { name: "압력(kPa)", prop: "pressure" }
  ];

  const tbody = document.querySelector("#tableDescriptiveFull tbody");
  if (!tbody) return;

  let html = "";
  for (let k of keys) {
    const arr = rawData.map(d => d[k.prop]);
    const s = calcDescriptive(arr);
    html += `
      <tr>
        <td><strong>${k.name}</strong></td>
        <td>${s.n.toLocaleString()}</td>
        <td>${s.mean.toFixed(3)}</td>
        <td>${s.sd.toFixed(3)}</td>
        <td>${s.median.toFixed(3)}</td>
        <td>${s.min.toFixed(2)}</td>
        <td>${s.q1.toFixed(2)}</td>
        <td>${s.q3.toFixed(2)}</td>
        <td>${s.max.toFixed(2)}</td>
        <td>${s.skew.toFixed(3)}</td>
        <td>${s.kurt.toFixed(3)}</td>
      </tr>
    `;
  }
  tbody.innerHTML = html;
}

// 원시 데이터 테이블 렌더링
function renderRawTable() {
  const tbody = document.getElementById("rawTableBody");
  filterRawTable();
}

function filterRawTable() {
  const search = (document.getElementById("filterInput")?.value || "").toLowerCase();
  const evFilter = document.getElementById("eventSelect")?.value || "ALL";

  const tbody = document.getElementById("rawTableBody");
  if (!tbody) return;

  const filtered = rawData.filter(d => {
    const matchEv = evFilter === "ALL" || d.event === evFilter;
    const matchSearch = !search || d.event.toLowerCase().includes(search) || d.time.includes(search);
    return matchEv && matchSearch;
  });

  // Limit display to first 300 rows for smooth DOM performance
  const displaySlice = filtered.slice(0, 300);
  let html = "";
  displaySlice.forEach((d, idx) => {
    let tagClass = "tag-normal";
    if (d.event.includes("과열")) tagClass = "tag-overheat";
    else if (d.event.includes("스파이크")) tagClass = "tag-spike";
    else if (d.event.includes("냉각")) tagClass = "tag-cool";
    else if (d.event.includes("부하")) tagClass = "tag-load";

    html += `
      <tr>
        <td>${idx + 1}</td>
        <td>${d.time}</td>
        <td>${d.minute}</td>
        <td>${d.hour}</td>
        <td>${d.tempRaw.toFixed(2)}</td>
        <td>${d.tempTrue.toFixed(2)}</td>
        <td>${d.vibRms.toFixed(3)}</td>
        <td>${d.pressure.toFixed(2)}</td>
        <td><span class="tag ${tagClass}">${d.event}</span></td>
      </tr>
    `;
  });

  if (filtered.length > 300) {
    html += `<tr><td colspan="9" style="text-align:center; color:#94a3b8; padding:12px;">... 외 ${filtered.length - 300}개 행이 더 있습니다 (상단 필터로 검색 가능) ...</td></tr>`;
  }

  tbody.innerHTML = html;
}

// Chart Helper
function renderChart(canvasId, config) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  if (charts[canvasId]) {
    charts[canvasId].destroy();
  }
  charts[canvasId] = new Chart(canvas, config);
}

// Tab Switching
function switchTab(tabId) {
  document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
  document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));

  const target = document.getElementById(tabId);
  if (target) target.classList.add("active");

  const clickedBtn = Array.from(document.querySelectorAll(".tab-btn")).find(b => b.getAttribute("onclick")?.includes(tabId));
  if (clickedBtn) clickedBtn.classList.add("active");

  // Resize charts
  Object.values(charts).forEach(c => c && c.resize());
}

// CSV Export
function exportAnalysisReport() {
  if (!rawData || rawData.length === 0) return;
  const groups = calcEventGroupStats(rawData);
  let csv = "이벤트유형,데이터수,온도평균,진동평균,압력평균\n";
  for (let g of groups) {
    csv += `"${g.event}",${g.count},${g.tempMean},${g.vibMean},${g.pressMean}\n`;
  }
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "센서_통계분석요약.csv";
  a.click();
  showToast("분석 요약 CSV 다운로드 완료");
}

function showToast(msg) {
  const toast = document.getElementById("statusToast");
  toast.innerText = msg;
  toast.style.display = "block";
  setTimeout(() => { toast.style.display = "none"; }, 3000);
}
