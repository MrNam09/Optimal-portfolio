/**
 * Exporter Module
 * Exports results to Excel and Word files with charts.
 */
const Exporter = (() => {

    /**
     * Capture a Plotly chart as PNG base64
     */
    async function captureChart(containerId, width = 700, height = 450) {
        const el = document.getElementById(containerId);
        if (!el || !el.data) return null;
        try {
            const result = await Plotly.toImage(el, { format: 'png', width, height, scale: 2 });
            // result is a data URL "data:image/png;base64,..."
            const base64 = result.split(',')[1];
            return base64;
        } catch (e) {
            console.warn('Failed to capture chart:', containerId, e);
            return null;
        }
    }

    /**
     * Convert base64 to Uint8Array for docx
     */
    function base64ToUint8Array(base64) {
        const binary = atob(base64);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) {
            bytes[i] = binary.charCodeAt(i);
        }
        return bytes;
    }

    /**
     * Export all results to Excel
     */
    function exportToExcel(data) {
        const wb = XLSX.utils.book_new();
        const { tickers, stockMetrics, corrMatrix, portfolioMetrics, weights, simResults, period, riskFreeRate, riskAversion } = data;

        const periodLabels = { '5y': '5 năm gần nhất', '3y': '3 năm gần nhất', '1y': 'Năm gần nhất', '250d': '250 phiên gần nhất' };

        // Sheet 1: Parameters
        const paramsData = [
            ['THÔNG SỐ ĐẦU VÀO'],
            ['Khoảng thời gian', periodLabels[period] || period],
            ['Lãi suất phi rủi ro (%/năm)', (riskFreeRate * 100).toFixed(2)],
            ['Hệ số ngại rủi ro (A)', riskAversion],
            ['Chỉ số thị trường', data.marketIndex || 'VNINDEX'],
            ['Giới hạn tỷ trọng tối đa (%)', data.maxWeightPct || 100],
            ['Số phiên giao dịch', data.dates ? data.dates.length : 'N/A'],
            ['Ngày xuất kết quả', new Date().toLocaleDateString('vi-VN')],
            [],
            ['CỔ PHIẾU TRONG DANH MỤC'],
            ['Mã CP', 'Tỷ trọng (%)']
        ];
        tickers.forEach((t, i) => {
            paramsData.push([t, (weights[i] * 100).toFixed(2)]);
        });
        const wsParams = XLSX.utils.aoa_to_sheet(paramsData);
        wsParams['!cols'] = [{ wch: 30 }, { wch: 20 }];
        XLSX.utils.book_append_sheet(wb, wsParams, 'Thông số');

        // Sheet 2: Stock Metrics
        const metricsHeader = ['Mã CP', 'Lợi nhuận KV (%/năm)', 'Rủi ro (%/năm)', 'Beta', 'Sharpe Ratio', 'Tỷ trọng (%)'];
        const metricsRows = tickers.map((t, i) => [
            t,
            (stockMetrics[t].annualReturn * 100).toFixed(4),
            (stockMetrics[t].annualRisk * 100).toFixed(4),
            stockMetrics[t].beta.toFixed(4),
            stockMetrics[t].sharpe.toFixed(4),
            (weights[i] * 100).toFixed(2)
        ]);
        const wsMetrics = XLSX.utils.aoa_to_sheet([
            ['CHỈ SỐ TỪNG CỔ PHIẾU'],
            metricsHeader,
            ...metricsRows,
            [],
            ['CHỈ SỐ DANH MỤC'],
            ['Lợi nhuận kỳ vọng (%/năm)', (portfolioMetrics.portReturn * 100).toFixed(4)],
            ['Rủi ro (%/năm)', (portfolioMetrics.portRisk * 100).toFixed(4)],
            ['Beta', portfolioMetrics.portBeta.toFixed(4)],
            ['Sharpe Ratio', portfolioMetrics.sharpe.toFixed(4)],
            ['Hữu dụng (%)', (portfolioMetrics.utility * 100).toFixed(4)]
        ]);
        wsMetrics['!cols'] = [{ wch: 20 }, { wch: 22 }, { wch: 18 }, { wch: 12 }, { wch: 14 }, { wch: 14 }];
        XLSX.utils.book_append_sheet(wb, wsMetrics, 'Chỉ số CP');

        // Sheet 3: Correlation Matrix
        const corrHeader = ['', ...tickers];
        const corrRows = tickers.map((t, i) => [t, ...corrMatrix.matrix[i].map(v => v.toFixed(4))]);
        const wsCorr = XLSX.utils.aoa_to_sheet([
            ['MA TRẬN HỆ SỐ TƯƠNG QUAN'],
            corrHeader,
            ...corrRows
        ]);
        XLSX.utils.book_append_sheet(wb, wsCorr, 'Tương quan');

        // Sheet 4: Covariance Matrix
        if (data.covMatrix) {
            const covHeader = ['', ...tickers];
            const covRows = tickers.map((t, i) => [t, ...data.covMatrix.matrix[i].map(v => v.toFixed(8))]);
            const wsCov = XLSX.utils.aoa_to_sheet([
                ['MA TRẬN HIỆP PHƯƠNG SAI (NGÀY)'],
                covHeader,
                ...covRows
            ]);
            XLSX.utils.book_append_sheet(wb, wsCov, 'Hiệp PS');
        }

        // Sheet 5: Simulation Results
        if (simResults) {
            const keyPorts = [
                { label: 'DANH MỤC LỢI NHUẬN CAO NHẤT', port: simResults.maxReturnPort },
                { label: 'DANH MỤC RỦI RO THẤP NHẤT', port: simResults.minRiskPort },
                { label: 'DANH MỤC SHARPE TỐI ĐA (TIẾP TUYẾN)', port: simResults.maxSharpePort },
                { label: 'DANH MỤC TỐI ƯU (HỮU DỤNG TỐI ĐA)', port: simResults.maxUtilityPort }
            ];

            const simData = [['KẾT QUẢ TỐI ƯU HOÁ'], ['Số danh mục mô phỏng', simResults.portfolios.length], ['Giới hạn tỷ trọng tối đa (%)', data.maxWeightPct || 100], []];

            for (const kp of keyPorts) {
                if (!kp.port) continue;
                simData.push([kp.label]);
                simData.push(['Lợi nhuận (%/năm)', (kp.port.portReturn * 100).toFixed(4)]);
                simData.push(['Rủi ro (%/năm)', (kp.port.portRisk * 100).toFixed(4)]);
                simData.push(['Beta', kp.port.portBeta.toFixed(4)]);
                simData.push(['Sharpe Ratio', kp.port.sharpe.toFixed(4)]);
                simData.push(['Hữu dụng (%)', (kp.port.utility * 100).toFixed(4)]);
                simData.push(['Tỷ trọng:']);
                tickers.forEach((t, i) => {
                    simData.push(['  ' + t, (kp.port.weights[i] * 100).toFixed(2) + '%']);
                });
                simData.push([]);
            }

            const wsSim = XLSX.utils.aoa_to_sheet(simData);
            wsSim['!cols'] = [{ wch: 40 }, { wch: 20 }];
            XLSX.utils.book_append_sheet(wb, wsSim, 'Tối ưu hoá');

            // Sheet 6: Top portfolios
            const topPorts = [...simResults.portfolios]
                .sort((a, b) => b.sharpe - a.sharpe)
                .slice(0, 1000);
            const allPortHeader = ['STT', 'Lợi nhuận (%)', 'Rủi ro (%)', 'Sharpe', 'Hữu dụng (%)', ...tickers.map(t => t + ' (%)')];
            const allPortRows = topPorts.map((p, idx) => [
                idx + 1,
                (p.portReturn * 100).toFixed(4),
                (p.portRisk * 100).toFixed(4),
                p.sharpe.toFixed(4),
                (p.utility * 100).toFixed(4),
                ...p.weights.map(w => (w * 100).toFixed(2))
            ]);
            const wsAll = XLSX.utils.aoa_to_sheet([
                ['TOP 1000 DANH MỤC (THEO SHARPE RATIO)'],
                allPortHeader,
                ...allPortRows
            ]);
            XLSX.utils.book_append_sheet(wb, wsAll, 'Top DM');
        }

        // Save
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const filename = `Portfolio_Result_${new Date().toISOString().slice(0, 10)}.xlsx`;
        saveAs(blob, filename);
        return filename;
    }

    /**
     * Export results to Word document with tables and chart images
     */
    async function exportToWord(data) {
        const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
            WidthType, AlignmentType, HeadingLevel, BorderStyle, ImageRun,
            ShadingType, TableLayoutType, PageBreak } = docx;

        const { tickers, stockMetrics, corrMatrix, portfolioMetrics, weights, simResults, period, riskFreeRate, riskAversion } = data;

        const periodLabels = { '5y': '5 năm gần nhất', '3y': '3 năm gần nhất', '1y': 'Năm gần nhất', '250d': '250 phiên gần nhất' };

        // ===== Capture all chart images =====
        const chartImages = {};
        const chartIds = ['price-chart', 'correlation-heatmap', 'frontier-chart', 'utility-chart', 'weight-comparison-chart'];
        for (const id of chartIds) {
            chartImages[id] = await captureChart(id, 680, 420);
        }

        const children = [];

        // ===== Helper functions =====
        const cellBorder = {
            top: { style: BorderStyle.SINGLE, size: 1, color: '999999' },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: '999999' },
            left: { style: BorderStyle.SINGLE, size: 1, color: '999999' },
            right: { style: BorderStyle.SINGLE, size: 1, color: '999999' }
        };

        const makeCell = (text, options = {}) => {
            const { bold = false, header = false, width, color } = options;
            const shading = header ? { type: ShadingType.SOLID, color: '1a73e8', fill: '1a73e8' } : undefined;
            const fontColor = header ? 'FFFFFF' : (color || '333333');
            return new TableCell({
                children: [new Paragraph({
                    children: [new TextRun({
                        text: String(text),
                        bold: bold || header,
                        size: header ? 18 : 19,
                        font: 'Arial',
                        color: fontColor
                    })],
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 30, after: 30 }
                })],
                borders: cellBorder,
                shading,
                width: width ? { size: width, type: WidthType.DXA } : undefined,
                verticalAlign: 'center'
            });
        };

        const addImage = (base64, w = 600, h = 370) => {
            if (!base64) return new Paragraph({ text: '[Biểu đồ không khả dụng]', spacing: { after: 200 } });
            return new Paragraph({
                children: [new ImageRun({
                    data: base64ToUint8Array(base64),
                    transformation: { width: w, height: h },
                    type: 'png'
                })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 100, after: 200 }
            });
        };

        const addCaption = (text) => new Paragraph({
            children: [new TextRun({ text, italics: true, size: 18, color: '666666', font: 'Arial' })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 }
        });

        const addSpacer = () => new Paragraph({ spacing: { after: 200 } });

        // ============================================================
        // TITLE PAGE
        // ============================================================
        children.push(new Paragraph({ spacing: { after: 600 } }));
        children.push(new Paragraph({
            children: [new TextRun({ text: 'BÁO CÁO', bold: true, size: 40, font: 'Arial', color: '1a73e8' })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 100 }
        }));
        children.push(new Paragraph({
            children: [new TextRun({ text: 'TỐI ƯU HOÁ DANH MỤC ĐẦU TƯ', bold: true, size: 36, font: 'Arial', color: '333333' })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 }
        }));
        children.push(new Paragraph({
            children: [new TextRun({ text: `Ngày lập: ${new Date().toLocaleDateString('vi-VN')}`, size: 22, font: 'Arial', color: '666666' })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 100 }
        }));
        children.push(new Paragraph({
            children: [new TextRun({ text: `Danh mục: ${tickers.join(', ')}`, size: 22, font: 'Arial', color: '666666' })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 100 }
        }));
        children.push(new Paragraph({
            children: [new TextRun({ text: `Khoảng thời gian: ${periodLabels[period] || period}`, size: 22, font: 'Arial', color: '666666' })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 100 }
        }));
        children.push(new Paragraph({
            children: [new TextRun({ text: `Số phiên giao dịch: ${data.dates ? data.dates.length : 'N/A'}`, size: 22, font: 'Arial', color: '666666' })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 600 }
        }));

        // Page break
        children.push(new Paragraph({ children: [new PageBreak()] }));

        // ============================================================
        // SECTION 1: INPUT PARAMETERS
        // ============================================================
        children.push(new Paragraph({
            children: [new TextRun({ text: 'BƯỚC 1: THÔNG SỐ ĐẦU VÀO', bold: true, size: 26, font: 'Arial', color: '1a73e8' })],
            spacing: { before: 200, after: 200 }
        }));

        // Parameters table
        const paramRows = [
            ['Khoảng thời gian', periodLabels[period] || period],
            ['Lãi suất phi rủi ro', `${(riskFreeRate * 100).toFixed(2)}%/năm`],
            ['Hệ số ngại rủi ro (A)', `${riskAversion}`],
            ['Chỉ số thị trường', data.marketIndex || 'VNINDEX'],
            ['Giới hạn tỷ trọng tối đa', `${data.maxWeightPct || 100}%`],
            ['Số phiên giao dịch', `${data.dates ? data.dates.length : 'N/A'}`]
        ];

        children.push(new Table({
            rows: paramRows.map(([label, value]) => new TableRow({
                children: [
                    makeCell(label, { bold: true, width: 4000 }),
                    makeCell(value, { width: 4000 })
                ]
            })),
            width: { size: 100, type: WidthType.PERCENTAGE }
        }));
        children.push(addSpacer());

        // Stocks and weights table
        children.push(new Paragraph({
            children: [new TextRun({ text: 'Danh sách cổ phiếu và tỷ trọng ban đầu:', bold: true, size: 22, font: 'Arial' })],
            spacing: { before: 200, after: 100 }
        }));

        const stockWeightRows = [
            new TableRow({ children: [makeCell('Mã CP', { header: true }), makeCell('Tỷ trọng (%)', { header: true })] }),
            ...tickers.map((t, i) => new TableRow({
                children: [makeCell(t, { bold: true }), makeCell((weights[i] * 100).toFixed(2))]
            }))
        ];
        children.push(new Table({
            rows: stockWeightRows,
            width: { size: 100, type: WidthType.PERCENTAGE }
        }));

        children.push(new Paragraph({ children: [new PageBreak()] }));

        // ============================================================
        // SECTION 2: INDIVIDUAL STOCK METRICS
        // ============================================================
        children.push(new Paragraph({
            children: [new TextRun({ text: 'BƯỚC 2: LỢI NHUẬN, RỦI RO VÀ BETA TỪNG CỔ PHIẾU', bold: true, size: 26, font: 'Arial', color: '1a73e8' })],
            spacing: { before: 200, after: 200 }
        }));

        // Metrics table
        const metricsTableRows = [
            new TableRow({
                children: ['Mã CP', 'Lợi nhuận (%/năm)', 'Rủi ro (%/năm)', 'Beta (β)', 'Sharpe Ratio', 'Tỷ trọng (%)'].map(h => makeCell(h, { header: true }))
            }),
            ...tickers.map((t, i) => {
                const m = stockMetrics[t];
                const retColor = m.annualReturn >= 0 ? '0f9d58' : 'ea4335';
                return new TableRow({
                    children: [
                        makeCell(t, { bold: true }),
                        makeCell((m.annualReturn * 100).toFixed(2), { color: retColor }),
                        makeCell((m.annualRisk * 100).toFixed(2)),
                        makeCell(m.beta.toFixed(3)),
                        makeCell(m.sharpe.toFixed(3)),
                        makeCell((weights[i] * 100).toFixed(1))
                    ]
                });
            })
        ];
        children.push(new Table({
            rows: metricsTableRows,
            width: { size: 100, type: WidthType.PERCENTAGE }
        }));
        children.push(addCaption('Bảng 1: Chỉ số từng cổ phiếu trong danh mục'));

        // Price chart
        children.push(new Paragraph({
            children: [new TextRun({ text: 'Biểu đồ giá cổ phiếu chuẩn hoá:', bold: true, size: 22, font: 'Arial' })],
            spacing: { before: 200, after: 100 }
        }));
        children.push(addImage(chartImages['price-chart']));
        children.push(addCaption('Hình 1: Biến động giá cổ phiếu chuẩn hoá theo thời gian'));

        // ============================================================
        // SECTION 3: CORRELATION MATRIX
        // ============================================================
        children.push(new Paragraph({
            children: [new TextRun({ text: 'MA TRẬN HỆ SỐ TƯƠNG QUAN', bold: true, size: 24, font: 'Arial', color: '1a73e8' })],
            spacing: { before: 300, after: 200 }
        }));

        const corrTableRows = [
            new TableRow({
                children: [makeCell('', { header: true }), ...tickers.map(t => makeCell(t, { header: true }))]
            }),
            ...tickers.map((t, i) => new TableRow({
                children: [
                    makeCell(t, { bold: true }),
                    ...corrMatrix.matrix[i].map(v => {
                        const color = v > 0.7 ? '1a73e8' : v < -0.3 ? 'ea4335' : '333333';
                        return makeCell(v.toFixed(3), { color });
                    })
                ]
            }))
        ];
        children.push(new Table({
            rows: corrTableRows,
            width: { size: 100, type: WidthType.PERCENTAGE }
        }));
        children.push(addCaption('Bảng 2: Ma trận hệ số tương quan giữa các cổ phiếu'));

        // Heatmap
        children.push(addImage(chartImages['correlation-heatmap']));
        children.push(addCaption('Hình 2: Bản đồ nhiệt tương quan'));

        children.push(new Paragraph({ children: [new PageBreak()] }));

        // ============================================================
        // SECTION 4: PORTFOLIO METRICS
        // ============================================================
        children.push(new Paragraph({
            children: [new TextRun({ text: 'BƯỚC 3: CHỈ SỐ DANH MỤC ĐẦU TƯ', bold: true, size: 26, font: 'Arial', color: '1a73e8' })],
            spacing: { before: 200, after: 200 }
        }));

        const pm = portfolioMetrics;
        const portMetricRows = [
            ['Lợi nhuận kỳ vọng (%/năm)', (pm.portReturn * 100).toFixed(4)],
            ['Rủi ro - σ (%/năm)', (pm.portRisk * 100).toFixed(4)],
            ['Beta (β)', pm.portBeta.toFixed(4)],
            ['Sharpe Ratio', pm.sharpe.toFixed(4)],
            ['Hữu dụng - U (%)', (pm.utility * 100).toFixed(4)]
        ];
        children.push(new Table({
            rows: [
                new TableRow({ children: [makeCell('Chỉ số', { header: true }), makeCell('Giá trị', { header: true })] }),
                ...portMetricRows.map(([label, value]) => new TableRow({
                    children: [makeCell(label, { bold: true }), makeCell(value)]
                }))
            ],
            width: { size: 100, type: WidthType.PERCENTAGE }
        }));
        children.push(addCaption('Bảng 3: Chỉ số tổng hợp của danh mục đầu tư'));

        // Covariance matrix
        if (data.covMatrix) {
            children.push(new Paragraph({
                children: [new TextRun({ text: 'MA TRẬN HIỆP PHƯƠNG SAI (NGÀY)', bold: true, size: 24, font: 'Arial', color: '1a73e8' })],
                spacing: { before: 300, after: 200 }
            }));

            const covTableRows = [
                new TableRow({
                    children: [makeCell('', { header: true }), ...tickers.map(t => makeCell(t, { header: true }))]
                }),
                ...tickers.map((t, i) => new TableRow({
                    children: [
                        makeCell(t, { bold: true }),
                        ...data.covMatrix.matrix[i].map(v => makeCell(v.toFixed(6)))
                    ]
                }))
            ];
            children.push(new Table({
                rows: covTableRows,
                width: { size: 100, type: WidthType.PERCENTAGE }
            }));
            children.push(addCaption('Bảng 4: Ma trận hiệp phương sai'));
        }

        children.push(new Paragraph({ children: [new PageBreak()] }));

        // ============================================================
        // SECTION 5: OPTIMIZATION
        // ============================================================
        if (simResults) {
            children.push(new Paragraph({
                children: [new TextRun({ text: 'BƯỚC 4: TỐI ƯU HOÁ DANH MỤC', bold: true, size: 26, font: 'Arial', color: '1a73e8' })],
                spacing: { before: 200, after: 200 }
            }));

            children.push(new Paragraph({
                children: [
                    new TextRun({ text: `Phương pháp: `, bold: true, size: 22, font: 'Arial' }),
                    new TextRun({ text: `Mô phỏng Monte Carlo với ${simResults.portfolios.length.toLocaleString()} danh mục ngẫu nhiên`, size: 22, font: 'Arial' })
                ],
                spacing: { after: 100 }
            }));
            children.push(new Paragraph({
                children: [
                    new TextRun({ text: `Giới hạn tỷ trọng: `, bold: true, size: 22, font: 'Arial' }),
                    new TextRun({ text: `Tối đa ${data.maxWeightPct || 100}% cho mỗi cổ phiếu`, size: 22, font: 'Arial' })
                ],
                spacing: { after: 200 }
            }));

            // Key portfolios comparison table
            const keyPorts = [
                { label: 'Lợi nhuận cao nhất', port: simResults.maxReturnPort },
                { label: 'Rủi ro thấp nhất', port: simResults.minRiskPort },
                { label: 'Sharpe tối đa (Tiếp tuyến)', port: simResults.maxSharpePort },
                { label: 'Tối ưu (Hữu dụng tối đa)', port: simResults.maxUtilityPort }
            ];

            // Summary comparison table
            const compHeaderCells = [makeCell('Chỉ số', { header: true })];
            keyPorts.forEach(kp => compHeaderCells.push(makeCell(kp.label, { header: true })));

            const compMetrics = [
                { label: 'Lợi nhuận (%/năm)', fn: p => (p.portReturn * 100).toFixed(2) },
                { label: 'Rủi ro (%/năm)', fn: p => (p.portRisk * 100).toFixed(2) },
                { label: 'Beta (β)', fn: p => p.portBeta.toFixed(3) },
                { label: 'Sharpe Ratio', fn: p => p.sharpe.toFixed(3) },
                { label: 'Hữu dụng (%)', fn: p => (p.utility * 100).toFixed(2) }
            ];

            const compRows = [
                new TableRow({ children: compHeaderCells }),
                ...compMetrics.map(metric => new TableRow({
                    children: [
                        makeCell(metric.label, { bold: true }),
                        ...keyPorts.map(kp => makeCell(kp.port ? metric.fn(kp.port) : 'N/A'))
                    ]
                }))
            ];

            // Add weight rows
            tickers.forEach((t, idx) => {
                compRows.push(new TableRow({
                    children: [
                        makeCell(`Tỷ trọng ${t} (%)`, { bold: true }),
                        ...keyPorts.map(kp => makeCell(kp.port ? (kp.port.weights[idx] * 100).toFixed(1) : 'N/A'))
                    ]
                }));
            });

            children.push(new Table({
                rows: compRows,
                width: { size: 100, type: WidthType.PERCENTAGE }
            }));
            children.push(addCaption('Bảng 5: So sánh các danh mục tối ưu'));

            // Efficient Frontier chart
            children.push(new Paragraph({
                children: [new TextRun({ text: 'Đường biên hiệu quả và đường CML:', bold: true, size: 22, font: 'Arial' })],
                spacing: { before: 300, after: 100 }
            }));
            children.push(addImage(chartImages['frontier-chart']));
            children.push(addCaption('Hình 3: Đường biên hiệu quả, CML và các danh mục tối ưu'));

            // Utility curve chart
            children.push(new Paragraph({
                children: [new TextRun({ text: 'Đường cong hữu dụng:', bold: true, size: 22, font: 'Arial' })],
                spacing: { before: 200, after: 100 }
            }));
            children.push(addImage(chartImages['utility-chart']));
            children.push(addCaption('Hình 4: Đường cong hữu dụng và danh mục tối ưu'));

            // Weight comparison chart
            children.push(new Paragraph({
                children: [new TextRun({ text: 'So sánh tỷ trọng các danh mục:', bold: true, size: 22, font: 'Arial' })],
                spacing: { before: 200, after: 100 }
            }));
            children.push(addImage(chartImages['weight-comparison-chart']));
            children.push(addCaption('Hình 5: So sánh tỷ trọng giữa các danh mục'));

            // Detailed optimal portfolio section
            children.push(new Paragraph({ children: [new PageBreak()] }));
            children.push(new Paragraph({
                children: [new TextRun({ text: 'CHI TIẾT TỪNG DANH MỤC TỐI ƯU', bold: true, size: 26, font: 'Arial', color: '1a73e8' })],
                spacing: { before: 200, after: 200 }
            }));

            const portLabels = [
                { label: '1. Danh mục lợi nhuận cao nhất', desc: 'Danh mục có lợi nhuận kỳ vọng cao nhất trong tất cả các danh mục mô phỏng.', port: simResults.maxReturnPort },
                { label: '2. Danh mục rủi ro thấp nhất (MVP)', desc: 'Danh mục có phương sai nhỏ nhất - nằm ở đỉnh trái của đường biên hiệu quả.', port: simResults.minRiskPort },
                { label: '3. Danh mục tiếp tuyến (Sharpe tối đa)', desc: 'Danh mục có tỷ lệ Sharpe cao nhất - điểm tiếp xúc giữa CML và đường biên hiệu quả.', port: simResults.maxSharpePort },
                { label: '4. Danh mục tối ưu (Hữu dụng tối đa)', desc: 'Danh mục tối đa hoá hàm hữu dụng U = E(r) - 0.5×A×σ² với A = ' + riskAversion + '.', port: simResults.maxUtilityPort }
            ];

            for (const pl of portLabels) {
                if (!pl.port) continue;

                children.push(new Paragraph({
                    children: [new TextRun({ text: pl.label, bold: true, size: 24, font: 'Arial', color: '333333' })],
                    spacing: { before: 200, after: 80 }
                }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: pl.desc, italics: true, size: 20, font: 'Arial', color: '666666' })],
                    spacing: { after: 100 }
                }));

                const p = pl.port;
                const detailRows = [
                    new TableRow({ children: [makeCell('Chỉ số', { header: true }), makeCell('Giá trị', { header: true })] }),
                    new TableRow({ children: [makeCell('Lợi nhuận kỳ vọng', { bold: true }), makeCell((p.portReturn * 100).toFixed(4) + '%/năm')] }),
                    new TableRow({ children: [makeCell('Rủi ro (σ)', { bold: true }), makeCell((p.portRisk * 100).toFixed(4) + '%/năm')] }),
                    new TableRow({ children: [makeCell('Beta (β)', { bold: true }), makeCell(p.portBeta.toFixed(4))] }),
                    new TableRow({ children: [makeCell('Sharpe Ratio', { bold: true }), makeCell(p.sharpe.toFixed(4))] }),
                    new TableRow({ children: [makeCell('Hữu dụng (U)', { bold: true }), makeCell((p.utility * 100).toFixed(4) + '%')] })
                ];

                // Weight rows
                tickers.forEach((t, i) => {
                    detailRows.push(new TableRow({
                        children: [makeCell(`Tỷ trọng ${t}`, { bold: true }), makeCell((p.weights[i] * 100).toFixed(2) + '%')]
                    }));
                });

                children.push(new Table({
                    rows: detailRows,
                    width: { size: 100, type: WidthType.PERCENTAGE }
                }));
                children.push(addSpacer());
            }
        }

        // ============================================================
        // FOOTER
        // ============================================================
        children.push(addSpacer());
        children.push(new Paragraph({
            children: [new TextRun({
                text: '─────────────────────────────────────────',
                size: 16, color: 'cccccc', font: 'Arial'
            })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 300, after: 100 }
        }));
        children.push(new Paragraph({
            children: [new TextRun({
                text: 'Báo cáo được tạo bởi Portfolio Optimizer',
                size: 18, italics: true, color: '999999', font: 'Arial'
            })],
            alignment: AlignmentType.CENTER
        }));

        // Build document
        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: { top: 720, right: 720, bottom: 720, left: 720 }
                    }
                },
                children
            }],
            creator: 'Portfolio Optimizer',
            title: 'Báo Cáo Tối Ưu Hoá Danh Mục Đầu Tư'
        });

        const blob = await Packer.toBlob(doc);
        const filename = `Portfolio_Report_${new Date().toISOString().slice(0, 10)}.docx`;
        saveAs(blob, filename);
        return filename;
    }

    return { exportToExcel, exportToWord };
})();
