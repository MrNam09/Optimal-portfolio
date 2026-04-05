/**
 * Data Processor Module
 * Handles loading and parsing CSV/Excel files for stock data, index data, and sector classification.
 */
const DataProcessor = (() => {
    // Store parsed data
    let stockData = {};      // { ticker: [ {date, open, high, low, close, volume}, ...] }
    let indexData = {};       // { indexName: [ {date, close}, ...] }
    let sectorData = {};     // { sectorName: [ticker1, ticker2, ...] }
    let tickerSector = {};   // { ticker: sectorName }
    let allTickers = [];
    let dateRange = { min: null, max: null };

    /**
     * Parse CafeF CSV text into structured data
     */
    function parseCSV(text) {
        const lines = text.trim().split('\n');
        const records = [];
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;
            const parts = line.split(',');
            if (parts.length < 7) continue;
            const ticker = parts[0].replace(/[<>"]/g, '').trim();
            const dateStr = parts[1].replace(/[<>"]/g, '').trim();
            if (!ticker || !dateStr || dateStr.length !== 8) continue;
            const year = parseInt(dateStr.substring(0, 4));
            const month = parseInt(dateStr.substring(4, 6)) - 1;
            const day = parseInt(dateStr.substring(6, 8));
            const date = new Date(year, month, day);
            records.push({
                ticker,
                date,
                dateStr,
                open: parseFloat(parts[2]) * 1000,
                high: parseFloat(parts[3]) * 1000,
                low: parseFloat(parts[4]) * 1000,
                close: parseFloat(parts[5]) * 1000,
                volume: parseInt(parts[6])
            });
        }
        return records;
    }

    /**
     * Load stock data from CSV files
     */
    function loadStockFiles(files) {
        return new Promise((resolve, reject) => {
            const promises = [];
            for (const file of files) {
                promises.push(new Promise((res, rej) => {
                    const reader = new FileReader();
                    reader.onload = (e) => {
                        try {
                            const text = e.target.result;
                            const records = parseCSV(text);
                            let count = 0;
                            for (const rec of records) {
                                if (!stockData[rec.ticker]) {
                                    stockData[rec.ticker] = [];
                                }
                                stockData[rec.ticker].push(rec);
                                count++;
                            }
                            res({ name: file.name, records: count });
                        } catch (err) {
                            rej(err);
                        }
                    };
                    reader.onerror = rej;
                    reader.readAsText(file, 'utf-8');
                }));
            }
            Promise.all(promises).then(results => {
                // Sort each ticker by date
                for (const ticker of Object.keys(stockData)) {
                    stockData[ticker].sort((a, b) => a.date - b.date);
                }
                allTickers = Object.keys(stockData).sort();
                updateDateRange();
                resolve(results);
            }).catch(reject);
        });
    }

    /**
     * Load index data from CSV file
     */
    function loadIndexFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const text = e.target.result;
                    const records = parseCSV(text);
                    for (const rec of records) {
                        if (!indexData[rec.ticker]) {
                            indexData[rec.ticker] = [];
                        }
                        indexData[rec.ticker].push(rec);
                    }
                    // Sort by date
                    for (const idx of Object.keys(indexData)) {
                        indexData[idx].sort((a, b) => a.date - b.date);
                    }
                    resolve({ name: file.name, indices: Object.keys(indexData) });
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = reject;
            reader.readAsText(file, 'utf-8');
        });
    }

    /**
     * Load sector classification from Excel file
     */
    function loadSectorFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheetName = workbook.SheetNames[0];
                    const sheet = workbook.Sheets[sheetName];
                    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                    // Try to detect sector and ticker columns
                    sectorData = {};
                    tickerSector = {};

                    // Find header row
                    let headerRow = 0;
                    let tickerCol = -1;
                    let sectorCol = -1;

                    for (let r = 0; r < Math.min(5, json.length); r++) {
                        const row = json[r];
                        if (!row) continue;
                        for (let c = 0; c < row.length; c++) {
                            const val = String(row[c] || '').toLowerCase().trim();
                            if (val.includes('mã') || val.includes('ticker') || val.includes('symbol') || val === 'mã ck' || val === 'mã cp') {
                                tickerCol = c;
                                headerRow = r;
                            }
                            if (val.includes('ngành') || val.includes('sector') || val.includes('industry') || val.includes('phân ngành')) {
                                sectorCol = c;
                                headerRow = r;
                            }
                        }
                    }

                    // If not found by header, try column patterns
                    if (tickerCol === -1 || sectorCol === -1) {
                        // Assume first col is number/stt, second is ticker, and look for sector
                        for (let r = 0; r < Math.min(10, json.length); r++) {
                            const row = json[r];
                            if (!row) continue;
                            for (let c = 0; c < row.length; c++) {
                                const val = String(row[c] || '').trim();
                                // Ticker-like: 3 uppercase letters
                                if (tickerCol === -1 && /^[A-Z]{3,4}$/.test(val)) {
                                    tickerCol = c;
                                    headerRow = Math.max(0, r - 1);
                                }
                            }
                        }
                        // Sector is usually a few columns after ticker
                        if (tickerCol >= 0 && sectorCol === -1) {
                            for (let c = tickerCol + 1; c < (json[headerRow + 1] || []).length; c++) {
                                const val = String(json[headerRow + 1][c] || '').trim();
                                if (val.length > 5 && !/^\d+$/.test(val)) {
                                    sectorCol = c;
                                    break;
                                }
                            }
                        }
                    }

                    if (tickerCol >= 0 && sectorCol >= 0) {
                        for (let r = headerRow + 1; r < json.length; r++) {
                            const row = json[r];
                            if (!row) continue;
                            const ticker = String(row[tickerCol] || '').trim().toUpperCase();
                            const sector = String(row[sectorCol] || '').trim();
                            if (ticker && sector && ticker.length <= 5) {
                                if (!sectorData[sector]) sectorData[sector] = [];
                                if (!sectorData[sector].includes(ticker)) {
                                    sectorData[sector].push(ticker);
                                }
                                tickerSector[ticker] = sector;
                            }
                        }
                    }

                    // If no sectors found, create a default "All" sector
                    if (Object.keys(sectorData).length === 0) {
                        sectorData['Tất cả'] = allTickers.slice();
                    }

                    resolve({
                        name: file.name,
                        sectors: Object.keys(sectorData).length,
                        stocks: Object.keys(tickerSector).length
                    });
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    function updateDateRange() {
        let min = null, max = null;
        for (const ticker of Object.keys(stockData)) {
            const prices = stockData[ticker];
            if (prices.length > 0) {
                const first = prices[0].date;
                const last = prices[prices.length - 1].date;
                if (!min || first < min) min = first;
                if (!max || last > max) max = last;
            }
        }
        dateRange = { min, max };
    }

    /**
     * Filter stock data by time period.
     * If the stock doesn't have enough data for the full period,
     * returns all available data instead of an empty array.
     */
    function filterByPeriod(ticker, period) {
        const prices = stockData[ticker];
        if (!prices || prices.length === 0) return [];

        const lastDate = dateRange.max;
        let startDate;
        let filtered;

        switch (period) {
            case '5y':
                startDate = new Date(lastDate);
                startDate.setFullYear(startDate.getFullYear() - 5);
                filtered = prices.filter(p => p.date >= startDate);
                break;
            case '3y':
                startDate = new Date(lastDate);
                startDate.setFullYear(startDate.getFullYear() - 3);
                filtered = prices.filter(p => p.date >= startDate);
                break;
            case '1y':
                // Year: Jan 1 to Dec 31 of the most recent full year
                const lastYear = lastDate.getFullYear() - 1;
                startDate = new Date(lastYear, 0, 1);
                const endDate = new Date(lastYear, 11, 31);
                filtered = prices.filter(p => p.date >= startDate && p.date <= endDate);
                // If no data in that exact year, fall back to last 250 sessions
                if (filtered.length === 0) {
                    filtered = prices.slice(-250);
                }
                break;
            case '250d':
                // Last 250 trading sessions (or fewer if not enough data)
                filtered = prices.slice(-250);
                break;
            default:
                startDate = new Date(lastDate);
                startDate.setFullYear(startDate.getFullYear() - 5);
                filtered = prices.filter(p => p.date >= startDate);
        }

        // If filtered result is too small, use all available data
        if (filtered.length < 2 && prices.length >= 2) {
            return prices;
        }
        return filtered;
    }

    /**
     * Get index data filtered by period.
     * Falls back to all available data if not enough for the selected period.
     */
    function getIndexByPeriod(indexName, period) {
        const data = indexData[indexName];
        if (!data || data.length === 0) return [];

        const lastDate = dateRange.max;
        let startDate;
        let filtered;

        switch (period) {
            case '5y':
                startDate = new Date(lastDate);
                startDate.setFullYear(startDate.getFullYear() - 5);
                filtered = data.filter(p => p.date >= startDate);
                break;
            case '3y':
                startDate = new Date(lastDate);
                startDate.setFullYear(startDate.getFullYear() - 3);
                filtered = data.filter(p => p.date >= startDate);
                break;
            case '1y':
                const lastYear = lastDate.getFullYear() - 1;
                startDate = new Date(lastYear, 0, 1);
                const endDate = new Date(lastYear, 11, 31);
                filtered = data.filter(p => p.date >= startDate && p.date <= endDate);
                if (filtered.length === 0) {
                    filtered = data.slice(-250);
                }
                break;
            case '250d':
                filtered = data.slice(-250);
                break;
            default:
                startDate = new Date(lastDate);
                startDate.setFullYear(startDate.getFullYear() - 5);
                filtered = data.filter(p => p.date >= startDate);
        }

        // Fall back to all available data if filtered is too small
        if (filtered.length < 2 && data.length >= 2) {
            return data;
        }
        return filtered;
    }

    /**
     * Build sector list - merge file sectors with available tickers
     */
    function buildSectorList() {
        // If we have sector data from file, filter to only tickers we have data for
        const result = {};
        if (Object.keys(sectorData).length > 0 && !sectorData['Tất cả']) {
            for (const [sector, tickers] of Object.entries(sectorData)) {
                const available = tickers.filter(t => stockData[t] && stockData[t].length > 0);
                if (available.length > 0) {
                    result[sector] = available.sort();
                }
            }
            // Add uncategorized tickers
            const categorized = new Set(Object.values(result).flat());
            const uncategorized = allTickers.filter(t => !categorized.has(t) && stockData[t].length > 10);
            if (uncategorized.length > 0) {
                result['Chưa phân ngành'] = uncategorized;
            }
        } else {
            // No sector file: group alphabetically
            result['Tất cả cổ phiếu'] = allTickers.filter(t => stockData[t].length > 10);
        }
        return result;
    }

    /**
     * Get closing prices array for a ticker in a period
     */
    function getClosingPrices(ticker, period) {
        const filtered = filterByPeriod(ticker, period);
        return filtered.map(p => ({ date: p.date, close: p.close }));
    }

    /**
     * Compute daily returns (simple returns)
     */
    function computeReturns(prices) {
        const returns = [];
        for (let i = 1; i < prices.length; i++) {
            if (prices[i - 1].close > 0) {
                returns.push({
                    date: prices[i].date,
                    ret: (prices[i].close - prices[i - 1].close) / prices[i - 1].close
                });
            }
        }
        return returns;
    }

    // Public API
    return {
        loadStockFiles,
        loadIndexFile,
        loadSectorFile,
        filterByPeriod,
        getIndexByPeriod,
        getClosingPrices,
        computeReturns,
        buildSectorList,
        getStockData: () => stockData,
        getIndexData: () => indexData,
        getSectorData: () => sectorData,
        getTickerSector: () => tickerSector,
        getAllTickers: () => allTickers,
        getDateRange: () => dateRange,
        getAvailableIndices: () => Object.keys(indexData),
        reset: () => {
            stockData = {};
            indexData = {};
            sectorData = {};
            tickerSector = {};
            allTickers = [];
            dateRange = { min: null, max: null };
        }
    };
})();
