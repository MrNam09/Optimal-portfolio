/**
 * Main Application Controller
 * Handles UI interactions, step navigation, and orchestrates all modules.
 */
(function () {
    'use strict';

    // ===== State =====
    let currentStep = 1;
    let stockFilesLoaded = false;
    let indexFileLoaded = false;
    let sectorFileLoaded = false;
    let selectedTickers = [];
    let currentWeights = [];
    let selectedPeriod = '5y';
    let riskFreeRate = 0.045;
    let riskAversion = 3;
    let marketIndex = 'VNINDEX';
    let simCount = 1000;

    // Calculated data stored for export
    let calcResults = null;

    // ===== Utility =====
    function showToast(msg, type = 'info') {
        const container = document.getElementById('toast-container');
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.innerHTML = `<i class="fas fa-${type === 'success' ? 'check-circle' : type === 'error' ? 'times-circle' : 'info-circle'}"></i> ${msg}`;
        container.appendChild(toast);
        setTimeout(() => toast.remove(), 4000);
    }

    function showLoading(text = 'Đang xử lý...') {
        document.getElementById('loading-text').textContent = text;
        document.getElementById('loading-overlay').classList.add('show');
    }

    function hideLoading() {
        document.getElementById('loading-overlay').classList.remove('show');
    }

    function formatNum(n, decimals = 2) {
        return Number(n).toFixed(decimals);
    }

    // ===== Step Navigation =====
    function goToStep(step) {
        document.querySelectorAll('.step-panel').forEach(p => p.classList.remove('active'));
        document.querySelectorAll('.step').forEach(s => {
            const sStep = parseInt(s.dataset.step);
            s.classList.remove('active', 'completed');
            if (sStep < step) s.classList.add('completed');
            if (sStep === step) s.classList.add('active');
        });
        document.getElementById(`step${step}-panel`).classList.add('active');
        currentStep = step;
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    // ===== Step 1: File Upload =====
    function setupDropzones() {
        const zones = [
            { dropId: 'stock-dropzone', inputId: 'stock-file', listId: 'stock-file-list', type: 'stock' },
            { dropId: 'index-dropzone', inputId: 'index-file', listId: 'index-file-list', type: 'index' },
            { dropId: 'sector-dropzone', inputId: 'sector-file', listId: 'sector-file-list', type: 'sector' }
        ];

        for (const zone of zones) {
            const dropEl = document.getElementById(zone.dropId);
            const inputEl = document.getElementById(zone.inputId);

            // Drag events
            dropEl.addEventListener('dragover', (e) => { e.preventDefault(); dropEl.classList.add('dragover'); });
            dropEl.addEventListener('dragleave', () => dropEl.classList.remove('dragover'));
            dropEl.addEventListener('drop', (e) => {
                e.preventDefault();
                dropEl.classList.remove('dragover');
                handleFiles(e.dataTransfer.files, zone);
            });

            // Click upload
            inputEl.addEventListener('change', (e) => {
                handleFiles(e.target.files, zone);
            });
        }
    }

    async function handleFiles(files, zone) {
        if (!files || files.length === 0) return;

        const listEl = document.getElementById(zone.listId);

        try {
            showLoading('Đang đọc dữ liệu...');

            if (zone.type === 'stock') {
                const results = await DataProcessor.loadStockFiles(files);
                stockFilesLoaded = true;
                listEl.innerHTML = '';
                for (const r of results) {
                    listEl.innerHTML += `
                        <div class="file-item">
                            <span class="file-name"><i class="fas fa-file-csv"></i> ${r.name}</span>
                            <span class="file-info">${r.records.toLocaleString()} bản ghi</span>
                        </div>`;
                }
                const allTickers = DataProcessor.getAllTickers();
                showToast(`Đã tải ${allTickers.length} mã cổ phiếu`, 'success');

                // Auto-create default sectors if no sector file
                if (!sectorFileLoaded) {
                    autoCreateSectors();
                }
            } else if (zone.type === 'index') {
                const result = await DataProcessor.loadIndexFile(files[0]);
                indexFileLoaded = true;
                listEl.innerHTML = `
                    <div class="file-item">
                        <span class="file-name"><i class="fas fa-file-csv"></i> ${result.name}</span>
                        <span class="file-info">Chỉ số: ${result.indices.join(', ')}</span>
                    </div>`;
                showToast(`Đã tải chỉ số: ${result.indices.join(', ')}`, 'success');

                // Update market index dropdown
                const select = document.getElementById('market-index');
                select.innerHTML = '';
                for (const idx of result.indices) {
                    select.innerHTML += `<option value="${idx}">${idx}</option>`;
                }
            } else if (zone.type === 'sector') {
                const result = await DataProcessor.loadSectorFile(files[0]);
                sectorFileLoaded = true;
                listEl.innerHTML = `
                    <div class="file-item">
                        <span class="file-name"><i class="fas fa-file-excel"></i> ${result.name}</span>
                        <span class="file-info">${result.sectors} ngành, ${result.stocks} cổ phiếu</span>
                    </div>`;
                showToast(`Đã tải ${result.sectors} ngành`, 'success');
            }

            checkStep1Ready();
        } catch (err) {
            showToast('Lỗi đọc file: ' + err.message, 'error');
            console.error(err);
        } finally {
            hideLoading();
        }
    }

    function autoCreateSectors() {
        // Group tickers by first letter as simple grouping
        // This will be replaced when sector file is loaded
    }

    function checkStep1Ready() {
        const ready = stockFilesLoaded && indexFileLoaded;
        document.getElementById('btn-next-1').disabled = !ready;
    }

    // ===== Step 2: Data Selection =====
    function setupStep2() {
        // Period selection
        const periodInputs = document.querySelectorAll('input[name="period"]');
        periodInputs.forEach(input => {
            input.addEventListener('change', () => {
                selectedPeriod = input.value;
                updatePeriodInfo();
            });
        });

        // Read parameters
        riskFreeRate = parseFloat(document.getElementById('risk-free-rate').value) / 100;
        riskAversion = parseFloat(document.getElementById('risk-aversion').value);
        marketIndex = document.getElementById('market-index').value;

        // Build sector and stock lists
        buildSectorUI();
        updatePeriodInfo();
    }

    function updatePeriodInfo() {
        const dr = DataProcessor.getDateRange();
        const info = document.getElementById('period-info');
        if (!dr.min || !dr.max) {
            info.textContent = '';
            return;
        }
        const labels = {
            '5y': `Tối đa 5 năm gần nhất (đến ${formatDate(dr.max)}). Dùng dữ liệu có sẵn nếu chưa đủ.`,
            '3y': `Tối đa 3 năm gần nhất (đến ${formatDate(dr.max)}). Dùng dữ liệu có sẵn nếu chưa đủ.`,
            '1y': `Năm ${dr.max.getFullYear() - 1} (01/01 - 31/12). Dùng dữ liệu có sẵn nếu chưa đủ.`,
            '250d': `Tối đa 250 phiên giao dịch gần nhất (đến ${formatDate(dr.max)})`
        };
        info.textContent = labels[selectedPeriod] || '';
    }

    function yearsBefore(date, n) {
        const d = new Date(date);
        d.setFullYear(d.getFullYear() - n);
        return d;
    }

    function formatDate(d) {
        return d.toLocaleDateString('vi-VN');
    }

    function buildSectorUI() {
        const sectors = DataProcessor.buildSectorList();
        const sectorList = document.getElementById('sector-list');
        sectorList.innerHTML = '';

        const sectorNames = Object.keys(sectors).sort();
        for (const name of sectorNames) {
            const div = document.createElement('div');
            div.className = 'sector-item';
            div.textContent = `${name} (${sectors[name].length})`;
            div.dataset.sector = name;
            div.addEventListener('click', () => {
                document.querySelectorAll('.sector-item').forEach(s => s.classList.remove('active'));
                div.classList.add('active');
                showStocksInSector(name, sectors[name]);
            });
            sectorList.appendChild(div);
        }

        // Sector search
        document.getElementById('sector-search').addEventListener('input', (e) => {
            const q = e.target.value.toLowerCase();
            sectorList.querySelectorAll('.sector-item').forEach(item => {
                item.style.display = item.textContent.toLowerCase().includes(q) ? '' : 'none';
            });
        });

        // Stock search (across all)
        document.getElementById('stock-search').addEventListener('input', (e) => {
            const q = e.target.value.toUpperCase().trim();
            if (q.length < 1) return;

            // Search across all sectors
            const allTickers = DataProcessor.getAllTickers();
            const matches = allTickers.filter(t => t.includes(q)).slice(0, 50);
            showStocksList(matches);
        });
    }

    function showStocksInSector(sectorName, tickers) {
        showStocksList(tickers);
    }

    function showStocksList(tickers) {
        const stockList = document.getElementById('stock-list');
        stockList.innerHTML = '';

        for (const ticker of tickers) {
            const div = document.createElement('div');
            div.className = 'stock-item' + (selectedTickers.includes(ticker) ? ' selected' : '');
            div.innerHTML = `<span class="stock-ticker">${ticker}</span>`;
            div.addEventListener('click', () => toggleStock(ticker));
            stockList.appendChild(div);
        }
    }

    function toggleStock(ticker) {
        const idx = selectedTickers.indexOf(ticker);
        if (idx >= 0) {
            selectedTickers.splice(idx, 1);
        } else {
            if (selectedTickers.length >= 10) {
                showToast('Tối đa 10 cổ phiếu', 'error');
                return;
            }
            selectedTickers.push(ticker);
        }
        updateSelectedUI();
        updateWeightSection();
        checkStep2Ready();
    }

    function updateSelectedUI() {
        // Update selected panel
        const container = document.getElementById('selected-stocks');
        container.innerHTML = '';
        for (const t of selectedTickers) {
            const div = document.createElement('div');
            div.className = 'selected-stock-tag';
            div.innerHTML = `<span>${t}</span> <span class="remove-stock" data-ticker="${t}"><i class="fas fa-times"></i></span>`;
            div.querySelector('.remove-stock').addEventListener('click', () => toggleStock(t));
            container.appendChild(div);
        }
        document.getElementById('selected-count').textContent = selectedTickers.length;

        // Update stock list selection state
        document.querySelectorAll('.stock-item').forEach(item => {
            const ticker = item.querySelector('.stock-ticker').textContent;
            item.classList.toggle('selected', selectedTickers.includes(ticker));
        });
    }

    function updateWeightSection() {
        const section = document.getElementById('weight-section');
        if (selectedTickers.length > 0) {
            section.style.display = '';
            updateWeightInputs();
        } else {
            section.style.display = 'none';
        }
    }

    function updateWeightInputs() {
        const mode = document.querySelector('input[name="weight-mode"]:checked').value;
        const container = document.getElementById('weight-inputs');
        container.innerHTML = '';

        if (mode === 'equal') {
            const equalWeight = (100 / selectedTickers.length).toFixed(2);
            currentWeights = selectedTickers.map(() => 1 / selectedTickers.length);
            for (const t of selectedTickers) {
                container.innerHTML += `
                    <div class="weight-input-group">
                        <label>${t}</label>
                        <input type="number" value="${equalWeight}" disabled>
                        <span>%</span>
                    </div>`;
            }
            updateWeightTotal();
        } else {
            currentWeights = selectedTickers.map(() => 100 / selectedTickers.length);
            for (let i = 0; i < selectedTickers.length; i++) {
                const div = document.createElement('div');
                div.className = 'weight-input-group';
                div.innerHTML = `
                    <label>${selectedTickers[i]}</label>
                    <input type="number" value="${(100 / selectedTickers.length).toFixed(2)}" min="0" max="100" step="0.1" data-idx="${i}">
                    <span>%</span>`;
                div.querySelector('input').addEventListener('input', (e) => {
                    currentWeights[parseInt(e.target.dataset.idx)] = parseFloat(e.target.value) || 0;
                    updateWeightTotal();
                    checkStep2Ready();
                });
                container.appendChild(div);
            }
            updateWeightTotal();
        }
    }

    function updateWeightTotal() {
        const mode = document.querySelector('input[name="weight-mode"]:checked').value;
        let total;
        if (mode === 'equal') {
            total = 100;
            currentWeights = selectedTickers.map(() => 1 / selectedTickers.length);
        } else {
            total = currentWeights.reduce((s, v) => s + v, 0);
        }
        const el = document.getElementById('weight-total');
        el.textContent = formatNum(total) + '%';
        el.className = 'weight-value' + (Math.abs(total - 100) > 0.1 ? ' invalid' : '');
    }

    function checkStep2Ready() {
        const hasStocks = selectedTickers.length >= 1;
        const mode = document.querySelector('input[name="weight-mode"]:checked').value;
        let weightsOk = true;
        if (mode === 'custom') {
            const total = currentWeights.reduce((s, v) => s + v, 0);
            weightsOk = Math.abs(total - 100) <= 0.1;
        }
        document.getElementById('btn-next-2').disabled = !(hasStocks && weightsOk);
    }

    // ===== Step 3: Calculate Results =====
    async function calculateResults() {
        showLoading('Đang tính toán chỉ số...');

        try {
            // Read parameters
            riskFreeRate = parseFloat(document.getElementById('risk-free-rate').value) / 100;
            riskAversion = parseFloat(document.getElementById('risk-aversion').value);
            marketIndex = document.getElementById('market-index').value;

            // Normalize weights
            const mode = document.querySelector('input[name="weight-mode"]:checked').value;
            let weights;
            if (mode === 'equal') {
                weights = selectedTickers.map(() => 1 / selectedTickers.length);
            } else {
                const total = currentWeights.reduce((s, v) => s + v, 0);
                weights = currentWeights.map(w => w / total);
            }

            // Get returns for each stock and market
            const returnsByTicker = {};
            const priceData = {};
            for (const ticker of selectedTickers) {
                const prices = DataProcessor.getClosingPrices(ticker, selectedPeriod);
                priceData[ticker] = prices;
                returnsByTicker[ticker] = DataProcessor.computeReturns(prices);
            }

            const indexPrices = DataProcessor.getIndexByPeriod(marketIndex, selectedPeriod);
            const marketReturns = DataProcessor.computeReturns(indexPrices);

            // Align returns
            const { aligned, dates, marketAligned } = PortfolioEngine.alignReturns(returnsByTicker, marketReturns);

            if (dates.length < 2) {
                showToast('Không có dữ liệu chung giữa các cổ phiếu. Hãy kiểm tra lại dữ liệu hoặc chọn cổ phiếu khác.', 'error');
                hideLoading();
                return false;
            }

            // Notify user about actual data range used
            const actualStart = dates[0];
            const actualEnd = dates[dates.length - 1];
            showToast(`Sử dụng ${dates.length} phiên giao dịch (${actualStart} → ${actualEnd})`, 'info');

            // Calculate individual metrics
            const stockMetrics = {};
            const returns = [];
            const betas = [];
            for (const ticker of selectedTickers) {
                const metrics = PortfolioEngine.calcStockMetrics(aligned[ticker], marketAligned, riskFreeRate);
                stockMetrics[ticker] = metrics;
                returns.push(metrics.annualReturn);
                betas.push(metrics.beta);
            }

            // Correlation & Covariance matrices
            const corrMatrix = PortfolioEngine.calcCorrelationMatrix(aligned);
            const covMatrix = PortfolioEngine.calcCovarianceMatrix(aligned);

            // Portfolio metrics
            const portfolioMetrics = PortfolioEngine.calcPortfolioMetrics(
                weights, returns, betas, covMatrix.matrix, riskFreeRate, riskAversion
            );

            // Store results
            calcResults = {
                tickers: selectedTickers,
                weights,
                stockMetrics,
                corrMatrix,
                covMatrix,
                portfolioMetrics,
                priceData,
                aligned,
                marketAligned,
                returns,
                betas,
                period: selectedPeriod,
                riskFreeRate,
                riskAversion,
                marketIndex,
                dates
            };

            // Render Step 3
            renderStep3();
            return true;
        } catch (err) {
            showToast('Lỗi tính toán: ' + err.message, 'error');
            console.error(err);
            return false;
        } finally {
            hideLoading();
        }
    }

    function renderStep3() {
        const { tickers, stockMetrics, corrMatrix, portfolioMetrics, weights, priceData } = calcResults;

        // Stock metrics table
        const tbody = document.querySelector('#stock-metrics-table tbody');
        tbody.innerHTML = '';
        tickers.forEach((t, i) => {
            const m = stockMetrics[t];
            const retClass = m.annualReturn >= 0 ? 'style="color:#0f9d58"' : 'style="color:#ea4335"';
            tbody.innerHTML += `
                <tr>
                    <td><strong>${t}</strong></td>
                    <td ${retClass}>${formatNum(m.annualReturn * 100, 2)}</td>
                    <td>${formatNum(m.annualRisk * 100, 2)}</td>
                    <td>${formatNum(m.beta, 3)}</td>
                    <td>${formatNum(weights[i] * 100, 1)}</td>
                    <td>${formatNum(m.sharpe, 3)}</td>
                </tr>`;
        });

        // Correlation matrix table
        const corrTable = document.getElementById('correlation-table');
        let html = '<thead><tr><th></th>';
        tickers.forEach(t => html += `<th>${t}</th>`);
        html += '</tr></thead><tbody>';
        tickers.forEach((t, i) => {
            html += `<tr><td><strong>${t}</strong></td>`;
            corrMatrix.matrix[i].forEach((v, j) => {
                const color = v > 0.5 ? '#1a73e8' : v < -0.5 ? '#ea4335' : '#5f6368';
                html += `<td style="color:${color}">${formatNum(v, 3)}</td>`;
            });
            html += '</tr>';
        });
        html += '</tbody>';
        corrTable.innerHTML = html;

        // Correlation heatmap
        ChartRenderer.renderCorrelationHeatmap('correlation-heatmap', corrMatrix.matrix, tickers);

        // Portfolio metrics
        const pm = portfolioMetrics;
        document.getElementById('port-return').textContent = formatNum(pm.portReturn * 100, 2);
        document.getElementById('port-risk').textContent = formatNum(pm.portRisk * 100, 2);
        document.getElementById('port-beta').textContent = formatNum(pm.portBeta, 3);
        document.getElementById('port-sharpe').textContent = formatNum(pm.sharpe, 3);
        document.getElementById('port-utility').textContent = formatNum(pm.utility * 100, 2);

        // Color code return
        document.getElementById('port-return').style.color = pm.portReturn >= 0 ? '#0f9d58' : '#ea4335';

        // Price chart
        ChartRenderer.renderPriceChart('price-chart', priceData, tickers);
    }

    // ===== Step 4: Optimization =====
    async function runOptimization() {
        if (!calcResults) return;

        showLoading('Đang chạy mô phỏng Monte Carlo...');
        const progressBar = document.getElementById('sim-progress');
        progressBar.style.display = 'block';

        try {
            const { returns, betas, covMatrix, riskFreeRate, riskAversion, tickers, weights } = calcResults;

            // Read max weight constraint
            const maxWeightPct = parseFloat(document.getElementById('max-weight-pct').value) || 100;
            const maxWeight = Math.max(0.1, Math.min(1.0, maxWeightPct / 100));
            calcResults.maxWeightPct = maxWeightPct;

            const simResults = await PortfolioEngine.simulatePortfolios(
                simCount, returns, betas, covMatrix.matrix, riskFreeRate, riskAversion,
                (progress) => {
                    const pct = Math.round(progress * 100);
                    progressBar.querySelector('.progress-fill').style.width = pct + '%';
                    progressBar.querySelector('.progress-text').textContent = pct + '%';
                },
                maxWeight
            );

            // Extract efficient frontier from simulated portfolios
            const frontier = PortfolioEngine.extractEfficientFrontier(simResults.portfolios);

            // CML
            const maxRisk = simResults.portfolios.length > 0 ?
                Math.max(...simResults.portfolios.map(p => p.portRisk)) : 0.5;
            const cml = PortfolioEngine.calcCML(riskFreeRate, simResults.maxSharpePort, maxRisk);

            // Utility curves
            const utilityCurves = PortfolioEngine.calcUtilityCurves(
                riskAversion, simResults.maxUtilityPort, maxRisk
            );

            // Store sim results
            calcResults.simResults = simResults;
            calcResults.frontier = frontier;
            calcResults.cml = cml;
            calcResults.utilityCurves = utilityCurves;

            // Individual stock points for frontier chart
            const individualStocks = tickers.map(t => ({
                ticker: t,
                ret: calcResults.stockMetrics[t].annualReturn,
                risk: calcResults.stockMetrics[t].annualRisk
            }));

            // Render frontier chart
            ChartRenderer.renderFrontierChart('frontier-chart', {
                portfolios: simResults.portfolios,
                frontier,
                cml,
                maxReturnPort: simResults.maxReturnPort,
                minRiskPort: simResults.minRiskPort,
                maxSharpePort: simResults.maxSharpePort,
                optimalPort: simResults.maxUtilityPort,
                riskFreeRate,
                individualStocks
            });

            // Render utility chart
            ChartRenderer.renderUtilityChart('utility-chart', {
                utilityCurves,
                frontier,
                optimalPort: simResults.maxUtilityPort,
                maxSharpePort: simResults.maxSharpePort,
                riskFreeRate
            });

            // Render optimal portfolio details
            renderOptimalDetails(simResults, tickers);

            // Weight comparison chart
            ChartRenderer.renderWeightComparison('weight-comparison-chart', {
                tickers,
                currentWeights: weights,
                maxReturnWeights: simResults.maxReturnPort?.weights,
                minRiskWeights: simResults.minRiskPort?.weights,
                maxSharpeWeights: simResults.maxSharpePort?.weights,
                optimalWeights: simResults.maxUtilityPort?.weights
            });

            showToast(`Mô phỏng ${simCount.toLocaleString()} danh mục hoàn tất`, 'success');
        } catch (err) {
            showToast('Lỗi mô phỏng: ' + err.message, 'error');
            console.error(err);
        } finally {
            hideLoading();
        }
    }

    function renderOptimalDetails(simResults, tickers) {
        const configs = [
            { id: 'max-return-port', port: simResults.maxReturnPort },
            { id: 'min-risk-port', port: simResults.minRiskPort },
            { id: 'tangency-port', port: simResults.maxSharpePort },
            { id: 'optimal-port', port: simResults.maxUtilityPort }
        ];

        for (const cfg of configs) {
            const container = document.querySelector(`#${cfg.id} .opt-details`);
            if (!cfg.port) {
                container.innerHTML = '<p class="empty-state">N/A</p>';
                continue;
            }
            const p = cfg.port;
            let html = `
                <div class="opt-row"><span class="opt-label">Lợi nhuận:</span><span class="opt-value" style="color:${p.portReturn >= 0 ? '#0f9d58' : '#ea4335'}">${formatNum(p.portReturn * 100, 2)}%</span></div>
                <div class="opt-row"><span class="opt-label">Rủi ro (σ):</span><span class="opt-value">${formatNum(p.portRisk * 100, 2)}%</span></div>
                <div class="opt-row"><span class="opt-label">Beta:</span><span class="opt-value">${formatNum(p.portBeta, 3)}</span></div>
                <div class="opt-row"><span class="opt-label">Sharpe:</span><span class="opt-value">${formatNum(p.sharpe, 3)}</span></div>
                <div class="opt-row"><span class="opt-label">Hữu dụng:</span><span class="opt-value">${formatNum(p.utility * 100, 2)}%</span></div>
                <hr style="margin:6px 0;border-color:#eee">
                <div style="font-size:12px;color:#5f6368;margin-bottom:4px"><strong>Tỷ trọng:</strong></div>`;
            tickers.forEach((t, i) => {
                const w = p.weights[i] * 100;
                html += `<div class="opt-row"><span class="opt-label">${t}:</span><span class="opt-value">${formatNum(w, 1)}%</span></div>`;
            });
            container.innerHTML = html;
        }
    }

    // ===== History Management =====
    function saveResult() {
        if (!calcResults) return;

        const history = JSON.parse(localStorage.getItem('portfolio_history') || '[]');
        const entry = {
            id: Date.now(),
            date: new Date().toLocaleString('vi-VN'),
            tickers: calcResults.tickers,
            weights: calcResults.weights,
            period: calcResults.period,
            riskFreeRate: calcResults.riskFreeRate,
            riskAversion: calcResults.riskAversion,
            portfolioMetrics: calcResults.portfolioMetrics,
            simCount: calcResults.simResults ? calcResults.simResults.portfolios.length : 0,
            maxSharpe: calcResults.simResults?.maxSharpePort?.sharpe || null,
            optimalReturn: calcResults.simResults?.maxUtilityPort?.portReturn || null,
            optimalRisk: calcResults.simResults?.maxUtilityPort?.portRisk || null
        };

        history.unshift(entry);
        // Keep last 50
        if (history.length > 50) history.length = 50;
        localStorage.setItem('portfolio_history', JSON.stringify(history));
        loadHistory();
        showToast('Đã lưu kết quả', 'success');
    }

    function loadHistory() {
        const history = JSON.parse(localStorage.getItem('portfolio_history') || '[]');
        const container = document.getElementById('history-list');

        if (history.length === 0) {
            container.innerHTML = '<p class="empty-state">Chưa có kết quả nào được lưu</p>';
            return;
        }

        container.innerHTML = '';
        for (const entry of history) {
            const div = document.createElement('div');
            div.className = 'history-item';
            div.innerHTML = `
                <div class="hist-date">${entry.date}</div>
                <div class="hist-stocks">${entry.tickers.join(', ')}</div>
                <div class="hist-metrics">
                    R: ${formatNum((entry.portfolioMetrics?.portReturn || 0) * 100)}% |
                    σ: ${formatNum((entry.portfolioMetrics?.portRisk || 0) * 100)}% |
                    Sharpe: ${formatNum(entry.portfolioMetrics?.sharpe || 0, 3)}
                </div>
                <div class="hist-actions">
                    <button class="btn btn-outline" onclick="deleteHistory(${entry.id})" style="font-size:11px;padding:2px 8px;">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>`;
            container.appendChild(div);
        }
    }

    window.deleteHistory = function (id) {
        let history = JSON.parse(localStorage.getItem('portfolio_history') || '[]');
        history = history.filter(h => h.id !== id);
        localStorage.setItem('portfolio_history', JSON.stringify(history));
        loadHistory();
        showToast('Đã xoá kết quả', 'info');
    };

    // ===== Event Binding =====
    function bindEvents() {
        // Step 1 -> Step 2
        document.getElementById('btn-next-1').addEventListener('click', () => {
            setupStep2();
            goToStep(2);
        });

        // Step 2 -> Step 3
        document.getElementById('btn-next-2').addEventListener('click', async () => {
            const ok = await calculateResults();
            if (ok) goToStep(3);
        });

        // Step 3 -> Step 4
        document.getElementById('btn-next-3').addEventListener('click', () => {
            goToStep(4);
            // Auto-run simulation
            runOptimization();
        });

        // Back buttons
        document.getElementById('btn-back-2').addEventListener('click', () => goToStep(1));
        document.getElementById('btn-back-3').addEventListener('click', () => goToStep(2));
        document.getElementById('btn-back-4').addEventListener('click', () => goToStep(3));

        // Weight mode change
        document.querySelectorAll('input[name="weight-mode"]').forEach(input => {
            input.addEventListener('change', () => {
                updateWeightInputs();
                checkStep2Ready();
            });
        });

        // Simulation count
        document.querySelectorAll('.btn-sim').forEach(btn => {
            btn.addEventListener('click', () => {
                document.querySelectorAll('.btn-sim').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                simCount = parseInt(btn.dataset.n);
            });
        });

        // Run simulation
        document.getElementById('btn-run-sim').addEventListener('click', runOptimization);

        // Export
        document.getElementById('btn-export-excel').addEventListener('click', () => {
            if (!calcResults) return;
            try {
                const filename = Exporter.exportToExcel(calcResults);
                showToast(`Đã xuất: ${filename}`, 'success');
            } catch (err) {
                showToast('Lỗi xuất Excel: ' + err.message, 'error');
            }
        });

        document.getElementById('btn-export-word').addEventListener('click', async () => {
            if (!calcResults) return;
            try {
                showLoading('Đang tạo báo cáo Word...');
                const filename = await Exporter.exportToWord(calcResults);
                showToast(`Đã xuất: ${filename}`, 'success');
            } catch (err) {
                showToast('Lỗi xuất Word: ' + err.message, 'error');
            } finally {
                hideLoading();
            }
        });

        document.getElementById('btn-save-result').addEventListener('click', saveResult);

        // History sidebar
        document.getElementById('history-toggle').addEventListener('click', () => {
            document.getElementById('history-sidebar').classList.toggle('open');
        });

        // Close sidebar when clicking outside
        document.addEventListener('click', (e) => {
            const sidebar = document.getElementById('history-sidebar');
            const toggle = document.getElementById('history-toggle');
            if (sidebar.classList.contains('open') && !sidebar.contains(e.target) && !toggle.contains(e.target)) {
                sidebar.classList.remove('open');
            }
        });

        // Step clicks (for navigation)
        document.querySelectorAll('.step').forEach(s => {
            s.addEventListener('click', () => {
                const step = parseInt(s.dataset.step);
                if (step < currentStep) goToStep(step);
            });
        });
    }

    // ===== Init =====
    function init() {
        setupDropzones();
        bindEvents();
        loadHistory();
        checkStep1Ready();
    }

    document.addEventListener('DOMContentLoaded', init);
})();
