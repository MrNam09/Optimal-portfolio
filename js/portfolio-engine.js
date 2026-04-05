/**
 * Portfolio Engine Module
 * Handles all financial calculations: returns, risk, beta, correlation,
 * portfolio metrics, optimization, efficient frontier, CML, utility.
 */
const PortfolioEngine = (() => {
    const TRADING_DAYS = 250;

    // ===== Basic Statistics =====
    function mean(arr) {
        if (arr.length === 0) return 0;
        return arr.reduce((s, v) => s + v, 0) / arr.length;
    }

    function variance(arr) {
        const m = mean(arr);
        return arr.reduce((s, v) => s + (v - m) ** 2, 0) / (arr.length - 1);
    }

    function stddev(arr) {
        return Math.sqrt(variance(arr));
    }

    function covariance(arr1, arr2) {
        const n = Math.min(arr1.length, arr2.length);
        const m1 = mean(arr1.slice(0, n));
        const m2 = mean(arr2.slice(0, n));
        let sum = 0;
        for (let i = 0; i < n; i++) {
            sum += (arr1[i] - m1) * (arr2[i] - m2);
        }
        return sum / (n - 1);
    }

    function correlation(arr1, arr2) {
        const cov = covariance(arr1, arr2);
        const s1 = stddev(arr1);
        const s2 = stddev(arr2);
        if (s1 === 0 || s2 === 0) return 0;
        return cov / (s1 * s2);
    }

    // ===== Stock Metrics =====

    /**
     * Calculate annualized return, risk, beta for a single stock
     * @param {Array} stockReturns - daily returns of stock
     * @param {Array} marketReturns - daily returns of market index (aligned)
     * @param {number} riskFreeRate - annual risk-free rate (decimal)
     * @returns {Object} {annualReturn, annualRisk, beta, sharpe, dailyReturns}
     */
    function calcStockMetrics(stockReturns, marketReturns, riskFreeRate) {
        const dailyRf = riskFreeRate / TRADING_DAYS;
        const avgDaily = mean(stockReturns);
        const dailyStd = stddev(stockReturns);

        // Annualize
        const annualReturn = avgDaily * TRADING_DAYS;
        const annualRisk = dailyStd * Math.sqrt(TRADING_DAYS);

        // Beta
        let beta = 1;
        if (marketReturns && marketReturns.length > 0) {
            const n = Math.min(stockReturns.length, marketReturns.length);
            const sRet = stockReturns.slice(stockReturns.length - n);
            const mRet = marketReturns.slice(marketReturns.length - n);
            const covSM = covariance(sRet, mRet);
            const varM = variance(mRet);
            beta = varM > 0 ? covSM / varM : 1;
        }

        // Sharpe Ratio
        const sharpe = annualRisk > 0 ? (annualReturn - riskFreeRate) / annualRisk : 0;

        return { annualReturn, annualRisk, beta, sharpe };
    }

    /**
     * Align returns by date for multiple stocks
     * Returns object with aligned arrays (same dates across all)
     */
    function alignReturns(returnsByTicker, marketReturnsArr) {
        // Find common dates
        const tickers = Object.keys(returnsByTicker);
        if (tickers.length === 0) return { aligned: {}, dates: [], marketAligned: [] };

        // Build date maps
        const dateMaps = {};
        for (const t of tickers) {
            dateMaps[t] = {};
            for (const r of returnsByTicker[t]) {
                const key = r.date.toISOString().slice(0, 10);
                dateMaps[t][key] = r.ret;
            }
        }

        // Market date map
        const marketMap = {};
        if (marketReturnsArr) {
            for (const r of marketReturnsArr) {
                const key = r.date.toISOString().slice(0, 10);
                marketMap[key] = r.ret;
            }
        }

        // Find common dates across all tickers and market
        let commonDates = Object.keys(dateMaps[tickers[0]]);
        for (let i = 1; i < tickers.length; i++) {
            const set = new Set(Object.keys(dateMaps[tickers[i]]));
            commonDates = commonDates.filter(d => set.has(d));
        }
        if (marketReturnsArr && marketReturnsArr.length > 0) {
            const marketSet = new Set(Object.keys(marketMap));
            commonDates = commonDates.filter(d => marketSet.has(d));
        }
        commonDates.sort();

        // Build aligned arrays
        const aligned = {};
        for (const t of tickers) {
            aligned[t] = commonDates.map(d => dateMaps[t][d]);
        }
        const marketAligned = commonDates.map(d => marketMap[d] || 0);

        return { aligned, dates: commonDates, marketAligned };
    }

    // ===== Correlation Matrix =====
    function calcCorrelationMatrix(alignedReturns) {
        const tickers = Object.keys(alignedReturns);
        const n = tickers.length;
        const matrix = Array.from({ length: n }, () => Array(n).fill(0));

        for (let i = 0; i < n; i++) {
            for (let j = i; j < n; j++) {
                if (i === j) {
                    matrix[i][j] = 1;
                } else {
                    const corr = correlation(alignedReturns[tickers[i]], alignedReturns[tickers[j]]);
                    matrix[i][j] = corr;
                    matrix[j][i] = corr;
                }
            }
        }
        return { tickers, matrix };
    }

    // ===== Covariance Matrix =====
    function calcCovarianceMatrix(alignedReturns) {
        const tickers = Object.keys(alignedReturns);
        const n = tickers.length;
        const matrix = Array.from({ length: n }, () => Array(n).fill(0));

        for (let i = 0; i < n; i++) {
            for (let j = i; j < n; j++) {
                const cov = covariance(alignedReturns[tickers[i]], alignedReturns[tickers[j]]);
                matrix[i][j] = cov;
                matrix[j][i] = cov;
            }
        }
        return { tickers, matrix };
    }

    // ===== Portfolio Metrics =====

    /**
     * Calculate portfolio return, risk, beta given weights and individual metrics
     * @param {Array} weights - array of weights (summing to 1)
     * @param {Array} returns - array of annualized returns
     * @param {Array} betas - array of betas
     * @param {Array<Array>} covMatrix - covariance matrix (daily)
     * @param {number} riskFreeRate - annual risk-free rate
     * @param {number} riskAversion - risk aversion coefficient
     */
    function calcPortfolioMetrics(weights, returns, betas, covMatrix, riskFreeRate, riskAversion) {
        const n = weights.length;

        // Portfolio return = sum(wi * ri)
        let portReturn = 0;
        for (let i = 0; i < n; i++) {
            portReturn += weights[i] * returns[i];
        }

        // Portfolio variance = w' * Sigma * w (annualized)
        let portVariance = 0;
        for (let i = 0; i < n; i++) {
            for (let j = 0; j < n; j++) {
                portVariance += weights[i] * weights[j] * covMatrix[i][j] * TRADING_DAYS;
            }
        }
        const portRisk = Math.sqrt(Math.max(0, portVariance));

        // Portfolio beta = sum(wi * beta_i)
        let portBeta = 0;
        for (let i = 0; i < n; i++) {
            portBeta += weights[i] * betas[i];
        }

        // Sharpe Ratio
        const sharpe = portRisk > 0 ? (portReturn - riskFreeRate) / portRisk : 0;

        // Utility = E(r) - 0.5 * A * sigma^2
        const utility = portReturn - 0.5 * riskAversion * portVariance;

        return { portReturn, portRisk, portBeta, sharpe, utility };
    }

    // ===== Random Portfolio Generation =====
    function generateRandomWeights(n, maxWeight = 1.0) {
        // Generate random weights respecting max weight constraint
        // Uses Dirichlet-like sampling with rejection for max weight
        let weights;
        let attempts = 0;
        const maxAttempts = 100;

        do {
            const raw = Array.from({ length: n }, () => Math.random());
            const sum = raw.reduce((s, v) => s + v, 0);
            weights = raw.map(v => v / sum);
            attempts++;

            // Check if all weights are within limit
            if (weights.every(w => w <= maxWeight + 0.001)) {
                return weights;
            }

            // If max weight is restrictive, use redistribution approach
            if (attempts >= maxAttempts) break;
        } while (true);

        // Fallback: clamp and redistribute
        weights = Array.from({ length: n }, () => Math.random());
        let total = weights.reduce((s, v) => s + v, 0);
        weights = weights.map(v => v / total);

        // Iteratively clamp and redistribute excess
        for (let iter = 0; iter < 50; iter++) {
            let excess = 0;
            let freeCount = 0;
            for (let i = 0; i < n; i++) {
                if (weights[i] > maxWeight) {
                    excess += weights[i] - maxWeight;
                    weights[i] = maxWeight;
                } else {
                    freeCount++;
                }
            }
            if (excess < 0.0001) break;
            // Distribute excess to non-capped weights
            for (let i = 0; i < n; i++) {
                if (weights[i] < maxWeight) {
                    weights[i] += excess / freeCount;
                }
            }
        }
        // Final normalize
        total = weights.reduce((s, v) => s + v, 0);
        weights = weights.map(v => v / total);
        return weights;
    }

    /**
     * Generate N random portfolios and find key portfolios
     * @param {number} maxWeight - max weight per asset (0-1), default 1.0 (no limit)
     */
    function simulatePortfolios(n, returns, betas, covMatrix, riskFreeRate, riskAversion, progressCallback, maxWeight = 1.0) {
        const numAssets = returns.length;
        const portfolios = [];
        let maxReturnPort = null;
        let minRiskPort = null;
        let maxSharpePort = null;
        let maxUtilityPort = null;

        // Adaptive batch size: bigger for large n to avoid UI lag
        const batchSize = n >= 50000 ? 2000 : 500;
        let processed = 0;

        // For large simulations, don't store all portfolios (memory)
        const storeAll = n <= 50000;

        function processBatch() {
            const end = Math.min(processed + batchSize, n);
            for (let i = processed; i < end; i++) {
                const weights = generateRandomWeights(numAssets, maxWeight);
                const metrics = calcPortfolioMetrics(weights, returns, betas, covMatrix, riskFreeRate, riskAversion);
                const port = { weights: [...weights], ...metrics };

                if (storeAll) {
                    portfolios.push(port);
                } else {
                    // For 100K+, only keep every 5th portfolio for charting + always keep top ones
                    if (i % 5 === 0) portfolios.push(port);
                }

                if (!maxReturnPort || metrics.portReturn > maxReturnPort.portReturn) {
                    maxReturnPort = port;
                }
                if (!minRiskPort || metrics.portRisk < minRiskPort.portRisk) {
                    minRiskPort = port;
                }
                if (!maxSharpePort || metrics.sharpe > maxSharpePort.sharpe) {
                    maxSharpePort = port;
                }
                if (!maxUtilityPort || metrics.utility > maxUtilityPort.utility) {
                    maxUtilityPort = port;
                }
            }
            processed = end;
            if (progressCallback) progressCallback(processed / n);
        }

        return new Promise((resolve) => {
            function runBatch() {
                processBatch();
                if (processed < n) {
                    setTimeout(runBatch, 0);
                } else {
                    resolve({
                        portfolios,
                        maxReturnPort,
                        minRiskPort,
                        maxSharpePort,
                        maxUtilityPort
                    });
                }
            }
            runBatch();
        });
    }

    // ===== Efficient Frontier (analytical for 2+ assets) =====

    /**
     * Compute efficient frontier points using optimization
     * Uses quadratic programming approximation via random sampling + grid
     */
    function computeEfficientFrontier(returns, betas, covMatrix, riskFreeRate, riskAversion, numPoints = 100) {
        const n = returns.length;
        const minRet = Math.min(...returns);
        const maxRet = Math.max(...returns);
        const retRange = maxRet - minRet;
        const frontierPoints = [];

        // For each target return, find minimum variance portfolio
        for (let p = 0; p <= numPoints; p++) {
            const targetReturn = minRet - retRange * 0.1 + (retRange * 1.3) * p / numPoints;

            // Use iterative optimization (gradient descent approximation)
            let bestWeights = null;
            let bestRisk = Infinity;

            // Try many random starts
            for (let trial = 0; trial < 200; trial++) {
                let weights = generateRandomWeights(n);

                // Project to target return constraint using simple iterative adjustment
                for (let iter = 0; iter < 50; iter++) {
                    // Calculate current return
                    let currRet = 0;
                    for (let i = 0; i < n; i++) currRet += weights[i] * returns[i];

                    // Adjust weights to get closer to target return
                    const retDiff = targetReturn - currRet;
                    if (Math.abs(retDiff) < 0.0001) break;

                    // Find the asset with highest/lowest return
                    let maxRetIdx = 0, minRetIdx = 0;
                    for (let i = 1; i < n; i++) {
                        if (returns[i] > returns[maxRetIdx]) maxRetIdx = i;
                        if (returns[i] < returns[minRetIdx]) minRetIdx = i;
                    }

                    // Shift weight
                    const shift = Math.min(0.05, Math.abs(retDiff) / Math.max(0.01, Math.abs(returns[maxRetIdx] - returns[minRetIdx])));
                    if (retDiff > 0) {
                        const transfer = Math.min(shift, weights[minRetIdx]);
                        weights[maxRetIdx] += transfer;
                        weights[minRetIdx] -= transfer;
                    } else {
                        const transfer = Math.min(shift, weights[maxRetIdx]);
                        weights[minRetIdx] += transfer;
                        weights[maxRetIdx] -= transfer;
                    }

                    // Ensure non-negative and normalize
                    for (let i = 0; i < n; i++) weights[i] = Math.max(0, weights[i]);
                    const sum = weights.reduce((s, v) => s + v, 0);
                    for (let i = 0; i < n; i++) weights[i] /= sum;
                }

                // Calculate risk for these weights
                let portVar = 0;
                for (let i = 0; i < n; i++) {
                    for (let j = 0; j < n; j++) {
                        portVar += weights[i] * weights[j] * covMatrix[i][j] * TRADING_DAYS;
                    }
                }
                const risk = Math.sqrt(Math.max(0, portVar));

                // Check if return is close to target
                let portRet = 0;
                for (let i = 0; i < n; i++) portRet += weights[i] * returns[i];

                if (Math.abs(portRet - targetReturn) < retRange * 0.05 && risk < bestRisk) {
                    bestRisk = risk;
                    bestWeights = [...weights];
                }
            }

            if (bestWeights) {
                let portRet = 0;
                for (let i = 0; i < n; i++) portRet += bestWeights[i] * returns[i];
                frontierPoints.push({ ret: portRet, risk: bestRisk, weights: bestWeights });
            }
        }

        // Sort by risk
        frontierPoints.sort((a, b) => a.risk - b.risk);

        // Keep only the efficient part (upper envelope)
        const efficient = [];
        let maxRetSoFar = -Infinity;
        // Find the minimum variance point
        let minVarIdx = 0;
        for (let i = 1; i < frontierPoints.length; i++) {
            if (frontierPoints[i].risk < frontierPoints[minVarIdx].risk) {
                minVarIdx = i;
            }
        }
        // Only keep points from min variance upward
        for (let i = minVarIdx; i < frontierPoints.length; i++) {
            if (frontierPoints[i].ret >= maxRetSoFar) {
                maxRetSoFar = frontierPoints[i].ret;
                efficient.push(frontierPoints[i]);
            }
        }

        return efficient;
    }

    /**
     * Build efficient frontier from simulated portfolios (extracting the envelope)
     */
    function extractEfficientFrontier(portfolios) {
        // Sort by risk
        const sorted = [...portfolios].sort((a, b) => a.portRisk - b.portRisk);

        // Bin by risk and keep max return in each bin
        const numBins = 200;
        if (sorted.length === 0) return [];

        const minRisk = sorted[0].portRisk;
        const maxRisk = sorted[sorted.length - 1].portRisk;
        const binWidth = (maxRisk - minRisk) / numBins;

        const bins = Array.from({ length: numBins + 1 }, () => null);
        for (const p of sorted) {
            const binIdx = Math.min(numBins, Math.floor((p.portRisk - minRisk) / (binWidth || 1)));
            if (!bins[binIdx] || p.portReturn > bins[binIdx].portReturn) {
                bins[binIdx] = p;
            }
        }

        // Find min variance portfolio
        let minVarPort = sorted[0];
        for (const p of sorted) {
            if (p.portRisk < minVarPort.portRisk) minVarPort = p;
        }

        // Filter to efficient frontier (above min variance return)
        const frontier = bins.filter(b => b !== null && b.portReturn >= minVarPort.portReturn);
        frontier.sort((a, b) => a.portRisk - b.portRisk);

        return frontier;
    }

    /**
     * Calculate CML (Capital Market Line) points
     */
    function calcCML(riskFreeRate, tangencyPort, maxRisk) {
        if (!tangencyPort || tangencyPort.portRisk === 0) return [];
        const slope = (tangencyPort.portReturn - riskFreeRate) / tangencyPort.portRisk;
        const points = [];
        for (let i = 0; i <= 50; i++) {
            const risk = (maxRisk * 1.2) * i / 50;
            const ret = riskFreeRate + slope * risk;
            points.push({ risk, ret });
        }
        return points;
    }

    /**
     * Calculate utility indifference curves
     */
    function calcUtilityCurves(riskAversion, optimalPort, maxRisk, numCurves = 5) {
        const curves = [];
        if (!optimalPort) return curves;

        const optUtility = optimalPort.utility;
        const spread = Math.abs(optUtility) * 0.5 || 0.1;

        for (let c = 0; c < numCurves; c++) {
            const U = optUtility - spread + (2 * spread) * c / (numCurves - 1);
            const points = [];
            for (let i = 0; i <= 100; i++) {
                const sigma = (maxRisk * 1.2) * i / 100;
                const ret = U + 0.5 * riskAversion * sigma * sigma;
                points.push({ risk: sigma, ret });
            }
            curves.push({ utility: U, points, isOptimal: c === Math.floor(numCurves / 2) });
        }
        return curves;
    }

    // Public API
    return {
        mean,
        variance,
        stddev,
        covariance,
        correlation,
        calcStockMetrics,
        alignReturns,
        calcCorrelationMatrix,
        calcCovarianceMatrix,
        calcPortfolioMetrics,
        generateRandomWeights,
        simulatePortfolios,
        computeEfficientFrontier,
        extractEfficientFrontier,
        calcCML,
        calcUtilityCurves,
        TRADING_DAYS
    };
})();
