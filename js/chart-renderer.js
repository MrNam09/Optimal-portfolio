/**
 * Chart Renderer Module
 * Renders all charts using Plotly.js
 */
const ChartRenderer = (() => {
    const defaultLayout = {
        font: { family: 'Inter, sans-serif', size: 12 },
        paper_bgcolor: 'transparent',
        plot_bgcolor: '#fafafa',
        margin: { t: 40, r: 30, b: 50, l: 60 },
        xaxis: { gridcolor: '#e8e8e8', zerolinecolor: '#ccc' },
        yaxis: { gridcolor: '#e8e8e8', zerolinecolor: '#ccc' },
        legend: { orientation: 'h', y: -0.15, x: 0.5, xanchor: 'center' },
        hoverlabel: { font: { family: 'Inter, sans-serif' } }
    };

    const colors = [
        '#1a73e8', '#ea4335', '#34a853', '#fbbc04', '#ff6d01',
        '#46bdc6', '#7baaf7', '#ee675c', '#57bb8a', '#e8710a'
    ];

    /**
     * Render normalized price chart for selected stocks
     */
    function renderPriceChart(containerId, priceData, tickers) {
        const traces = tickers.map((ticker, i) => {
            const prices = priceData[ticker];
            if (!prices || prices.length === 0) return null;
            const basePrice = prices[0].close;
            return {
                x: prices.map(p => p.date),
                y: prices.map(p => (p.close / basePrice - 1) * 100),
                type: 'scatter',
                mode: 'lines',
                name: ticker,
                line: { color: colors[i % colors.length], width: 2 }
            };
        }).filter(Boolean);

        const layout = {
            ...defaultLayout,
            title: { text: 'Biến động giá chuẩn hoá (%)', font: { size: 14 } },
            xaxis: { ...defaultLayout.xaxis, title: 'Thời gian' },
            yaxis: { ...defaultLayout.yaxis, title: '% thay đổi' },
            hovermode: 'x unified'
        };

        Plotly.newPlot(containerId, traces, layout, { responsive: true });
    }

    /**
     * Render correlation heatmap
     */
    function renderCorrelationHeatmap(containerId, corrMatrix, tickers) {
        const trace = {
            z: corrMatrix,
            x: tickers,
            y: tickers,
            type: 'heatmap',
            colorscale: [
                [0, '#ea4335'],
                [0.5, '#ffffff'],
                [1, '#1a73e8']
            ],
            zmin: -1,
            zmax: 1,
            text: corrMatrix.map(row => row.map(v => v.toFixed(3))),
            texttemplate: '%{text}',
            textfont: { size: 11 },
            hovertemplate: '%{x} - %{y}: %{z:.4f}<extra></extra>'
        };

        const layout = {
            ...defaultLayout,
            title: { text: 'Ma trận tương quan', font: { size: 14 } },
            xaxis: { ...defaultLayout.xaxis, side: 'bottom' },
            yaxis: { ...defaultLayout.yaxis, autorange: 'reversed' },
            margin: { t: 40, r: 80, b: 80, l: 80 }
        };

        Plotly.newPlot(containerId, [trace], layout, { responsive: true });
    }

    /**
     * Render efficient frontier, CML, simulated portfolios, and key portfolios
     */
    function renderFrontierChart(containerId, data) {
        const { portfolios, frontier, cml, maxReturnPort, minRiskPort, maxSharpePort, optimalPort, riskFreeRate, individualStocks } = data;
        const traces = [];

        // Simulated portfolios (scatter)
        if (portfolios && portfolios.length > 0) {
            // Limit displayed points for performance
            const displayPorts = portfolios.length > 5000 ?
                portfolios.filter((_, i) => i % Math.ceil(portfolios.length / 5000) === 0) : portfolios;
            traces.push({
                x: displayPorts.map(p => p.portRisk * 100),
                y: displayPorts.map(p => p.portReturn * 100),
                mode: 'markers',
                type: 'scatter',
                name: 'Danh mục ngẫu nhiên',
                marker: {
                    size: 3,
                    color: displayPorts.map(p => p.sharpe),
                    colorscale: 'Viridis',
                    colorbar: { title: 'Sharpe', thickness: 12, len: 0.5 },
                    opacity: 0.5
                },
                hovertemplate: 'Rủi ro: %{x:.2f}%<br>Lợi nhuận: %{y:.2f}%<br>Sharpe: %{marker.color:.3f}<extra></extra>'
            });
        }

        // Efficient frontier
        if (frontier && frontier.length > 0) {
            traces.push({
                x: frontier.map(p => (p.portRisk || p.risk) * 100),
                y: frontier.map(p => (p.portReturn || p.ret) * 100),
                mode: 'lines',
                type: 'scatter',
                name: 'Đường biên hiệu quả',
                line: { color: '#e91e63', width: 3 },
                hovertemplate: 'Rủi ro: %{x:.2f}%<br>Lợi nhuận: %{y:.2f}%<extra>Biên hiệu quả</extra>'
            });
        }

        // CML
        if (cml && cml.length > 0) {
            traces.push({
                x: cml.map(p => p.risk * 100),
                y: cml.map(p => p.ret * 100),
                mode: 'lines',
                type: 'scatter',
                name: 'CML',
                line: { color: '#ff9800', width: 2, dash: 'dash' },
                hovertemplate: 'Rủi ro: %{x:.2f}%<br>Lợi nhuận: %{y:.2f}%<extra>CML</extra>'
            });
        }

        // Risk-free rate point
        if (riskFreeRate !== undefined) {
            traces.push({
                x: [0],
                y: [riskFreeRate * 100],
                mode: 'markers+text',
                type: 'scatter',
                name: 'Rf',
                marker: { size: 12, color: '#ff9800', symbol: 'diamond' },
                text: ['Rf'],
                textposition: 'top right',
                textfont: { size: 11, color: '#ff9800' },
                showlegend: true
            });
        }

        // Individual stocks
        if (individualStocks) {
            traces.push({
                x: individualStocks.map(s => s.risk * 100),
                y: individualStocks.map(s => s.ret * 100),
                mode: 'markers+text',
                type: 'scatter',
                name: 'Cổ phiếu riêng lẻ',
                marker: { size: 10, color: colors.slice(0, individualStocks.length), symbol: 'circle' },
                text: individualStocks.map(s => s.ticker),
                textposition: 'top center',
                textfont: { size: 10 },
                hovertemplate: '%{text}<br>Rủi ro: %{x:.2f}%<br>Lợi nhuận: %{y:.2f}%<extra></extra>'
            });
        }

        // Key portfolios
        const keyPorts = [
            { port: maxReturnPort, name: 'Max Return', color: '#00c853', symbol: 'triangle-up' },
            { port: minRiskPort, name: 'Min Risk', color: '#2196f3', symbol: 'square' },
            { port: maxSharpePort, name: 'Max Sharpe (Tiếp tuyến)', color: '#ff9800', symbol: 'star' },
            { port: optimalPort, name: 'Tối ưu (Max Utility)', color: '#e91e63', symbol: 'hexagram' }
        ];

        for (const kp of keyPorts) {
            if (kp.port) {
                traces.push({
                    x: [kp.port.portRisk * 100],
                    y: [kp.port.portReturn * 100],
                    mode: 'markers+text',
                    type: 'scatter',
                    name: kp.name,
                    marker: { size: 16, color: kp.color, symbol: kp.symbol, line: { color: 'white', width: 2 } },
                    text: [kp.name],
                    textposition: 'top center',
                    textfont: { size: 10, color: kp.color },
                    hovertemplate: `${kp.name}<br>Rủi ro: %{x:.2f}%<br>Lợi nhuận: %{y:.2f}%<extra></extra>`
                });
            }
        }

        const layout = {
            ...defaultLayout,
            title: { text: 'Đường Biên Hiệu Quả & Đường CML', font: { size: 14 } },
            xaxis: { ...defaultLayout.xaxis, title: 'Rủi ro - σ (%/năm)' },
            yaxis: { ...defaultLayout.yaxis, title: 'Lợi nhuận kỳ vọng (%/năm)' },
            hovermode: 'closest',
            legend: { ...defaultLayout.legend, y: -0.2 },
            margin: { t: 50, r: 30, b: 80, l: 70 }
        };

        Plotly.newPlot(containerId, traces, layout, { responsive: true });
    }

    /**
     * Render utility indifference curves with efficient frontier
     */
    function renderUtilityChart(containerId, data) {
        const { utilityCurves, frontier, optimalPort, maxSharpePort, riskFreeRate } = data;
        const traces = [];

        // Efficient frontier
        if (frontier && frontier.length > 0) {
            traces.push({
                x: frontier.map(p => (p.portRisk || p.risk) * 100),
                y: frontier.map(p => (p.portReturn || p.ret) * 100),
                mode: 'lines',
                type: 'scatter',
                name: 'Đường biên hiệu quả',
                line: { color: '#e91e63', width: 3 }
            });
        }

        // Utility curves
        if (utilityCurves) {
            const curveColors = ['#ccc', '#aaa', '#1a73e8', '#aaa', '#ccc'];
            utilityCurves.forEach((curve, i) => {
                traces.push({
                    x: curve.points.map(p => p.risk * 100),
                    y: curve.points.map(p => p.ret * 100),
                    mode: 'lines',
                    type: 'scatter',
                    name: curve.isOptimal ? `U* = ${(curve.utility * 100).toFixed(2)}%` : `U = ${(curve.utility * 100).toFixed(2)}%`,
                    line: {
                        color: curveColors[i] || '#ccc',
                        width: curve.isOptimal ? 3 : 1.5,
                        dash: curve.isOptimal ? 'solid' : 'dot'
                    },
                    showlegend: true
                });
            });
        }

        // Optimal portfolio point
        if (optimalPort) {
            traces.push({
                x: [optimalPort.portRisk * 100],
                y: [optimalPort.portReturn * 100],
                mode: 'markers+text',
                type: 'scatter',
                name: 'Danh mục tối ưu',
                marker: { size: 16, color: '#e91e63', symbol: 'hexagram', line: { color: 'white', width: 2 } },
                text: ['Tối ưu'],
                textposition: 'top right',
                textfont: { size: 11, color: '#e91e63' }
            });
        }

        if (maxSharpePort) {
            traces.push({
                x: [maxSharpePort.portRisk * 100],
                y: [maxSharpePort.portReturn * 100],
                mode: 'markers+text',
                type: 'scatter',
                name: 'Tiếp tuyến',
                marker: { size: 14, color: '#ff9800', symbol: 'star', line: { color: 'white', width: 2 } },
                text: ['Tiếp tuyến'],
                textposition: 'top left',
                textfont: { size: 11, color: '#ff9800' }
            });
        }

        const maxRisk = frontier && frontier.length > 0 ?
            Math.max(...frontier.map(p => (p.portRisk || p.risk))) * 100 * 1.3 : 50;
        const maxRet = frontier && frontier.length > 0 ?
            Math.max(...frontier.map(p => (p.portReturn || p.ret))) * 100 * 1.5 : 50;

        const layout = {
            ...defaultLayout,
            title: { text: 'Đường Cong Hữu Dụng & Đường Biên Hiệu Quả', font: { size: 14 } },
            xaxis: { ...defaultLayout.xaxis, title: 'Rủi ro - σ (%/năm)', range: [0, maxRisk] },
            yaxis: { ...defaultLayout.yaxis, title: 'Lợi nhuận kỳ vọng (%/năm)', range: [-maxRet * 0.3, maxRet] },
            hovermode: 'closest',
            legend: { ...defaultLayout.legend, y: -0.2 },
            margin: { t: 50, r: 30, b: 80, l: 70 }
        };

        Plotly.newPlot(containerId, traces, layout, { responsive: true });
    }

    /**
     * Render weight comparison bar chart for key portfolios
     */
    function renderWeightComparison(containerId, data) {
        const { tickers, currentWeights, maxReturnWeights, minRiskWeights, maxSharpeWeights, optimalWeights } = data;
        const traces = [];

        const configs = [
            { weights: currentWeights, name: 'Hiện tại', color: '#5f6368' },
            { weights: maxReturnWeights, name: 'Max Return', color: '#00c853' },
            { weights: minRiskWeights, name: 'Min Risk', color: '#2196f3' },
            { weights: maxSharpeWeights, name: 'Max Sharpe', color: '#ff9800' },
            { weights: optimalWeights, name: 'Tối ưu', color: '#e91e63' }
        ];

        for (const cfg of configs) {
            if (cfg.weights) {
                traces.push({
                    x: tickers,
                    y: cfg.weights.map(w => w * 100),
                    type: 'bar',
                    name: cfg.name,
                    marker: { color: cfg.color }
                });
            }
        }

        const layout = {
            ...defaultLayout,
            title: { text: 'So sánh tỷ trọng các danh mục (%)', font: { size: 14 } },
            barmode: 'group',
            xaxis: { ...defaultLayout.xaxis, title: 'Mã cổ phiếu' },
            yaxis: { ...defaultLayout.yaxis, title: 'Tỷ trọng (%)' },
            legend: { ...defaultLayout.legend, y: -0.2 }
        };

        Plotly.newPlot(containerId, traces, layout, { responsive: true });
    }

    return {
        renderPriceChart,
        renderCorrelationHeatmap,
        renderFrontierChart,
        renderUtilityChart,
        renderWeightComparison
    };
})();
