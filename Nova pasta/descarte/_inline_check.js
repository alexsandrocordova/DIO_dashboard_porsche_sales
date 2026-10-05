
        // ============================================================
        // INGESTAO DE DADOS — base de producao sanitizada
        // Prioridade: window.PORSCHE_SANITIZED_DATA (dashboard_data.js).
        // Upload manual .xlsx/.csv permanece disponivel como contingencia.
        // ============================================================
        let rawSalesData = [];
        let filteredData = [];
        let currentPage = 1;
        const rowsPerPage = 7;

        function parsePrice(value) {
            if (typeof value === 'number') return value;
            const cleaned = String(value).replace(/[^0-9.,-]/g, '').replace(',', '.');
            const n = parseFloat(cleaned);
            return isNaN(n) ? 0 : n;
        }

        // Normaliza uma linha (ativo JS de producao OU upload de planilha)
        // para o schema interno do dashboard, com deteccao flexivel de colunas.
        function normalizeRow(row, index) {
            const keys = Object.keys(row);
            const findValue = (candidates) => {
                const key = keys.find(k => {
                    const cleanKey = k.toLowerCase().trim().replace(/[_\s]/g, '');
                    return candidates.some(c => cleanKey.includes(c));
                });
                return key ? row[key] : null;
            };

            let modelo = String(findValue(['modelo', 'porschemodelsanitized', 'porschemodel', 'model']) || '').trim();
            if (modelo && !modelo.toLowerCase().includes('porsche')) {
                modelo = `Porsche ${modelo}`;
            }

            const rawYear = findValue(['anomodelo', 'modelyearsanitized', 'modelyear', 'ano']);
            let anoModelo = parseInt(String(rawYear).replace(/[^0-9]/g, ''), 10);
            if (isNaN(anoModelo) || anoModelo < 1900 || anoModelo > 2035) anoModelo = null;

            const cidade = String(findValue(['cidade', 'citysanitized', 'city']) || '').trim();
            const estado = String(findValue(['estado', 'statesanitized', 'state']) || '').trim();
            const formaPagamento = String(findValue(['formapagamento', 'paymethodsanitized', 'paymethod', 'paymentmethod', 'pagamento']) || '').trim();
            const entrega = String(findValue(['entrega', 'deliverystatussanitized', 'deliverystatus', 'deliverystatus']) || '').trim();
            const dataVenda = String(findValue(['data', 'saledatesanitized', 'saledate']) || '').trim();
            const quilometragem = parseInt(String(findValue(['quilometragem', 'vehiclemileagesanitized', 'vehiclemileage', 'mileage']) || '').replace(/[^0-9]/g, ''), 10);
            const preco = parsePrice(findValue(['preco', 'salespricesanitized', 'saleprice', 'price']));
            const id = findValue(['saleid', 'id']) || `POR-${1000 + index}`;

            return {
                id: String(id),
                data: dataVenda,
                modelo: modelo,
                anoModelo: anoModelo,
                cidade: cidade,
                estado: estado,
                formaPagamento: formaPagamento,
                entrega: entrega,
                quilometragem: isNaN(quilometragem) ? null : quilometragem,
                preco: Math.round(preco)
            };
        }

        // Aplica um lote de linhas ao estado global do dashboard
        function processRows(rows) {
            if (!rows || rows.length === 0) return false;
            const normalizedData = rows.map(normalizeRow).filter(r => r.modelo);
            if (normalizedData.length === 0) return false;

            rawSalesData = normalizedData;
            populateFilterOptions();
            applyFilters();
            return true;
        }

        // Atualiza badges de status e rotulo da base ativa
        function updateDataStatus(success, statusText, labelHtml) {
            const badge = document.getElementById('data-status-badge');
            const text = document.getElementById('data-status-text');
            const label = document.getElementById('dropzone-label');
            badge.className = success
                ? 'hidden md:flex items-center space-x-2 bg-emerald-950/80 border border-emerald-500/50 px-3 py-1.5 rounded-full'
                : 'hidden md:flex items-center space-x-2 bg-porsche-card border border-porsche-border px-3 py-1.5 rounded-full';
            text.innerText = statusText;
            text.className = success ? 'text-emerald-400 font-bold' : 'text-porsche-gray font-medium';
            if (labelHtml) {
                label.innerHTML = labelHtml;
            } else {
                label.innerHTML = success
                    ? `<span class="text-emerald-400 font-bold"><i class="fa-solid fa-circle-check"></i> Base ativa:</span> <strong>${rawSalesData.length} registros</strong> sanitizados carregados automaticamente.`
                    : 'Aguardando base de dados...';
            }
        }

        // Read File Handler
        function handleExcelFile(file) {
            if (!file) return;
            const reader = new FileReader();

            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);

                    if (processRows(jsonData)) {
                        updateDataStatus(true, `Planilha XLSX (${rawSalesData.length} Vendas)`);
                    } else {
                        updateDataStatus(false, 'Nenhum registro valido', `<span class="text-porsche-red font-bold"><i class="fa-solid fa-triangle-exclamation"></i> Nenhum registro valido encontrado no arquivo.</span>`);
                    }
                } catch (err) {
                    console.error("Erro ao ler a planilha Excel:", err);
                    document.getElementById('dropzone-label').innerHTML = 
                        `<span class="text-porsche-red font-bold"><i class="fa-solid fa-triangle-exclamation"></i> Erro ao processar o arquivo.</span> Certifique-se de que é uma planilha Excel válida (.xlsx, .xls ou .csv).`;
                }
            };

            reader.readAsArrayBuffer(file);
        }

        // Initialize Filter Dropdowns
        function populateFilterOptions() {
            const modeloSelect = document.getElementById('filter-modelo');
            const yearSelect = document.getElementById('filter-year');
            const citySelect = document.getElementById('filter-city');
            const paySelect = document.getElementById('filter-pay');

            // Preserve current selections if possible
            const selMod = modeloSelect.value;
            const selYr = yearSelect.value;
            const selCt = citySelect.value;
            const selPy = paySelect.value;

            modeloSelect.innerHTML = '<option value="ALL">Todos os Modelos</option>';
            yearSelect.innerHTML = '<option value="ALL">Todos os Anos</option>';
            citySelect.innerHTML = '<option value="ALL">Todas as Cidades</option>';
            paySelect.innerHTML = '<option value="ALL">Todas as Formas</option>';

            // Unique Models
            const models = [...new Set(rawSalesData.map(d => d.modelo))].filter(Boolean).sort();
            models.forEach(m => modeloSelect.appendChild(new Option(m, m)));

            // Unique Years
            const years = [...new Set(rawSalesData.map(d => d.anoModelo))].filter(Boolean).sort((a,b) => b - a);
            years.forEach(y => yearSelect.appendChild(new Option(`Model Year ${y}`, y)));

            // Unique Cities
            const cities = [...new Set(rawSalesData.map(d => d.cidade))].filter(Boolean).sort();
            cities.forEach(c => citySelect.appendChild(new Option(c, c)));

            // Unique Payment Methods
            const payMethods = [...new Set(rawSalesData.map(d => d.formaPagamento))].filter(Boolean).sort();
            payMethods.forEach(p => paySelect.appendChild(new Option(p, p)));

            if ([...modeloSelect.options].some(o => o.value === selMod)) modeloSelect.value = selMod;
            if ([...yearSelect.options].some(o => o.value === selYr)) yearSelect.value = selYr;
            if ([...citySelect.options].some(o => o.value === selCt)) citySelect.value = selCt;
            if ([...paySelect.options].some(o => o.value === selPy)) paySelect.value = selPy;
        }

        // Apply Active Filters
        function applyFilters() {
            const selectedModelo = document.getElementById('filter-modelo').value;
            const selectedYear = document.getElementById('filter-year').value;
            const selectedCity = document.getElementById('filter-city').value;
            const selectedPay = document.getElementById('filter-pay').value;

            filteredData = rawSalesData.filter(item => {
                const matchModelo = (selectedModelo === 'ALL' || item.modelo === selectedModelo);
                const matchYear = (selectedYear === 'ALL' || item.anoModelo == selectedYear);
                const matchCity = (selectedCity === 'ALL' || item.cidade === selectedCity);
                const matchPay = (selectedPay === 'ALL' || item.formaPagamento === selectedPay);
                return matchModelo && matchYear && matchCity && matchPay;
            });

            currentPage = 1;
            updateDashboard();
        }

        // Format Currency USD (base de producao em dolares)
        function formatUSD(value) {
            return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(value);
        }

        function updateDashboard() {
            document.getElementById('active-filter-count').innerText = `Exibindo ${filteredData.length} de ${rawSalesData.length} vendas`;

            // Calculate Metrics
            const totalRevenue = filteredData.reduce((acc, curr) => acc + curr.preco, 0);
            const totalUnits = filteredData.length;
            const avgTicket = totalUnits > 0 ? totalRevenue / totalUnits : 0;

            // Model frequency
            const modelCounts = {};
            filteredData.forEach(d => {
                modelCounts[d.modelo] = (modelCounts[d.modelo] || 0) + 1;
            });

            let topModel = "--";
            let topModelCount = 0;
            Object.keys(modelCounts).forEach(m => {
                if (modelCounts[m] > topModelCount) {
                    topModelCount = modelCounts[m];
                    topModel = m;
                }
            });

            const topModelShare = totalUnits > 0 ? ((topModelCount / totalUnits) * 100).toFixed(1) : 0;

            // Render KPIs
            document.getElementById('kpi-total-revenue').innerText = formatUSD(totalRevenue);
            document.getElementById('kpi-total-units').innerText = `${totalUnits} unid.`;
            document.getElementById('kpi-avg-ticket').innerText = formatUSD(avgTicket);
            document.getElementById('kpi-top-model').innerText = topModel;
            document.getElementById('kpi-top-model-share').innerText = `${topModelShare}% do total selecionado (${topModelCount} vds)`;

            // Business Insights Calculation
            generateBusinessInsights(modelCounts, totalUnits, totalRevenue);

            // Redraw All Canvas Charts
            renderCanvasModelsByCity();
            renderCanvasModelYearTrend();
            renderCanvasPayMethod();
            renderCanvasTopPopularModels();

            // Render Data Table
            renderTable();
        }

        function generateBusinessInsights(modelCounts, totalUnits, totalRevenue) {
            if (totalUnits === 0) {
                document.getElementById('insight-q1').innerText = "Nenhum registro encontrado para os filtros selecionados.";
                document.getElementById('insight-q2').innerText = "Nenhum registro encontrado para os filtros selecionados.";
                document.getElementById('insight-q3').innerText = "Nenhum registro encontrado para os filtros selecionados.";
                return;
            }

            // Insight Q1: Principais modelos por cidade
            const cityModelMap = {};
            filteredData.forEach(d => {
                if (!cityModelMap[d.cidade]) cityModelMap[d.cidade] = {};
                cityModelMap[d.cidade][d.modelo] = (cityModelMap[d.cidade][d.modelo] || 0) + 1;
            });

            let citySummary = [];
            const sortedCities = Object.keys(cityModelMap).sort((a,b) => {
                const totalA = Object.values(cityModelMap[a]).reduce((s, x) => s + x, 0);
                const totalB = Object.values(cityModelMap[b]).reduce((s, x) => s + x, 0);
                return totalB - totalA;
            });

            sortedCities.slice(0, 3).forEach(c => {
                let bestM = "";
                let maxC = 0;
                Object.keys(cityModelMap[c]).forEach(m => {
                    if (cityModelMap[c][m] > maxC) {
                        maxC = cityModelMap[c][m];
                        bestM = m.replace("Porsche ", "");
                    }
                });
                citySummary.push(`<strong>${c}</strong> (${bestM})`);
            });

            document.getElementById('insight-q1').innerHTML = 
                `Em destaque regional: ${citySummary.join(', ')} registraram as maiores densidades de entregas na planilha. Concessionárias de grande porte concentram a procura por esportivos de topo.`;
            document.getElementById('insight-q1-badge').innerText = `Liderança em ${sortedCities[0] || 'São Paulo'}`;

            // Insight Q2: Model Year mais demandado
            const yearCounts = {};
            filteredData.forEach(d => {
                if (d.anoModelo) yearCounts[d.anoModelo] = (yearCounts[d.anoModelo] || 0) + 1;
            });
            let topYear = "";
            let topYearVal = 0;
            Object.keys(yearCounts).forEach(y => {
                if (yearCounts[y] > topYearVal) {
                    topYearVal = yearCounts[y];
                    topYear = y;
                }
            });

            const yearShare = ((topYearVal / totalUnits) * 100).toFixed(1);
            document.getElementById('insight-q2').innerHTML = 
                `O <strong>Model Year ${topYear}</strong> representa o maior volume na base, responsável por <strong>${yearShare}%</strong> (${topYearVal} unidades). Demonstra rápida absorção de mercado e renovação de inventário.`;
            document.getElementById('insight-q2-badge').innerText = `Pico em Model Year ${topYear}`;

            // Insight Q3: Carros populares / preferência por vendas
            const suvCount = filteredData.filter(d => d.modelo.toLowerCase().includes("macan") || d.modelo.toLowerCase().includes("cayenne")).length;
            const suvShare = ((suvCount / totalUnits) * 100).toFixed(1);

            document.getElementById('insight-q3').innerHTML = 
                `Os modelos SUV (<strong>Macan e Cayenne</strong>) respondem por <strong>${suvShare}%</strong> das entregas totais na planilha. O modelo <strong>${document.getElementById('kpi-top-model').innerText}</strong> é a preferência líder do cliente Porsche no recorte.`;
            document.getElementById('insight-q3-badge').innerText = `SUVs somam ${suvShare}% do volume`;
        }

        function setupCanvasPixelRatio(canvasId) {
            const canvas = document.getElementById(canvasId);
            if (!canvas) return null;
            const ctx = canvas.getContext('2d');
            const dpr = window.devicePixelRatio || 1;
            const rect = canvas.getBoundingClientRect();

            if (rect.width <= 0 || rect.height <= 0) return null;

            canvas.width = rect.width * dpr;
            canvas.height = rect.height * dpr;
            ctx.scale(dpr, dpr);
            return { ctx, width: rect.width, height: rect.height };
        }

        // Chart 1: Models by City
        function renderCanvasModelsByCity() {
            const setup = setupCanvasPixelRatio('canvas-models-by-city');
            if (!setup) return;
            const { ctx, width, height } = setup;

            ctx.clearRect(0, 0, width, height);

            const cityData = {};
            filteredData.forEach(d => {
                cityData[d.cidade] = (cityData[d.cidade] || 0) + 1;
            });

            const labels = Object.keys(cityData).sort((a,b) => cityData[b] - cityData[a]).slice(0, 7);
            const values = labels.map(l => cityData[l]);
            const maxVal = Math.max(...values, 5);

            const padding = { top: 25, bottom: 45, left: 35, right: 20 };
            const chartW = width - padding.left - padding.right;
            const chartH = height - padding.top - padding.bottom;

            ctx.lineWidth = 1;
            ctx.strokeStyle = '#2a3038';
            ctx.fillStyle = '#8a94a0';
            ctx.font = '10px Inter';

            for (let i = 0; i <= 4; i++) {
                const yVal = Math.round((maxVal / 4) * i);
                const yPos = padding.top + chartH - (i * (chartH / 4));

                ctx.beginPath();
                ctx.moveTo(padding.left, yPos);
                ctx.lineTo(width - padding.right, yPos);
                ctx.stroke();

                ctx.fillText(yVal, 5, yPos + 3);
            }

            if (labels.length === 0) return;

            const barWidth = Math.min(36, (chartW / labels.length) * 0.55);
            const step = chartW / labels.length;

            labels.forEach((city, index) => {
                const val = values[index];
                const barH = (val / maxVal) * chartH;
                const x = padding.left + (index * step) + (step - barWidth) / 2;
                const y = padding.top + chartH - barH;

                const grad = ctx.createLinearGradient(0, y, 0, y + barH);
                grad.addColorStop(0, '#d5001c');
                grad.addColorStop(1, '#66000e');

                ctx.fillStyle = grad;
                ctx.beginPath();
                if (barH > 5) {
                    ctx.roundRect(x, y, barWidth, barH, [4, 4, 0, 0]);
                } else {
                    ctx.rect(x, y, barWidth, barH);
                }
                ctx.fill();

                if (val > 0) {
                    ctx.fillStyle = '#ffffff';
                    ctx.font = 'bold 11px Montserrat';
                    ctx.textAlign = 'center';
                    ctx.fillText(val, x + barWidth / 2, y - 6);
                }

                ctx.fillStyle = '#8a94a0';
                ctx.font = '10px Inter';
                ctx.textAlign = 'center';
                const labelCity = city.length > 9 ? city.substring(0, 8) + '...' : city;
                ctx.fillText(labelCity, x + barWidth / 2, height - 15);
            });
        }

        // Chart 2: Model Year Trend
        function renderCanvasModelYearTrend() {
            const setup = setupCanvasPixelRatio('canvas-model-year-trend');
            if (!setup) return;
            const { ctx, width, height } = setup;

            ctx.clearRect(0, 0, width, height);

            const yearMap = {};
            filteredData.forEach(d => {
                if (d.anoModelo) yearMap[d.anoModelo] = (yearMap[d.anoModelo] || 0) + 1;
            });

            const labels = Object.keys(yearMap).map(Number).sort((a,b) => a - b);
            const values = labels.map(y => yearMap[y]);
            const maxVal = Math.max(...values, 5);

            const padding = { top: 25, bottom: 40, left: 35, right: 25 };
            const chartW = width - padding.left - padding.right;
            const chartH = height - padding.top - padding.bottom;

            ctx.strokeStyle = '#2a3038';
            ctx.lineWidth = 1;
            ctx.fillStyle = '#8a94a0';
            ctx.font = '10px Inter';
            ctx.textAlign = 'left';

            for (let i = 0; i <= 4; i++) {
                const yVal = Math.round((maxVal / 4) * i);
                const yPos = padding.top + chartH - (i * (chartH / 4));

                ctx.beginPath();
                ctx.moveTo(padding.left, yPos);
                ctx.lineTo(width - padding.right, yPos);
                ctx.stroke();

                ctx.fillText(yVal, 5, yPos + 3);
            }

            if (labels.length === 0) return;

            const step = labels.length > 1 ? chartW / (labels.length - 1) : chartW;
            const points = labels.map((year, i) => {
                const val = values[i];
                const x = labels.length > 1 ? padding.left + (i * step) : padding.left + chartW / 2;
                const y = padding.top + chartH - ((val / maxVal) * chartH);
                return { x, y, val, year };
            });

            if (points.length > 1) {
                const areaGrad = ctx.createLinearGradient(0, padding.top, 0, padding.top + chartH);
                areaGrad.addColorStop(0, 'rgba(194, 161, 95, 0.35)');
                areaGrad.addColorStop(1, 'rgba(194, 161, 95, 0.0)');

                ctx.beginPath();
                ctx.moveTo(points[0].x, padding.top + chartH);
                ctx.lineTo(points[0].x, points[0].y);

                for (let i = 1; i < points.length; i++) {
                    ctx.lineTo(points[i].x, points[i].y);
                }

                ctx.lineTo(points[points.length - 1].x, padding.top + chartH);
                ctx.closePath();
                ctx.fillStyle = areaGrad;
                ctx.fill();

                ctx.beginPath();
                ctx.moveTo(points[0].x, points[0].y);
                for (let i = 1; i < points.length; i++) {
                    ctx.lineTo(points[i].x, points[i].y);
                }
                ctx.strokeStyle = '#c2a15f';
                ctx.lineWidth = 3;
                ctx.stroke();
            }

            points.forEach(pt => {
                ctx.beginPath();
                ctx.arc(pt.x, pt.y, 5, 0, Math.PI * 2);
                ctx.fillStyle = '#1b1f24';
                ctx.fill();
                ctx.strokeStyle = '#c2a15f';
                ctx.lineWidth = 2.5;
                ctx.stroke();

                ctx.fillStyle = '#ffffff';
                ctx.font = 'bold 11px Montserrat';
                ctx.textAlign = 'center';
                ctx.fillText(pt.val, pt.x, pt.y - 10);

                ctx.fillStyle = '#8a94a0';
                ctx.font = '10px Inter';
                ctx.fillText(`MY '${String(pt.year).slice(-2)}`, pt.x, height - 12);
            });
        }

        // Chart 3: Payment Method Donut
        function renderCanvasPayMethod() {
            const setup = setupCanvasPixelRatio('canvas-pay-method');
            if (!setup) return;
            const { ctx, width, height } = setup;

            ctx.clearRect(0, 0, width, height);

            const payMap = {};
            filteredData.forEach(d => {
                payMap[d.formaPagamento] = (payMap[d.formaPagamento] || 0) + 1;
            });

            const keys = Object.keys(payMap);
            const total = filteredData.length;
            const colors = ['#d5001c', '#c2a15f', '#3b82f6', '#10b981', '#a855f7', '#f97316'];

            const centerX = width / 2;
            const centerY = height / 2;

            const radius = Math.max(10, Math.min(width, height) / 2.3 - 10);
            const innerRadius = Math.max(5, radius * 0.62);

            let startAngle = -Math.PI / 2;

            keys.forEach((pay, idx) => {
                const count = payMap[pay];
                const sliceAngle = total > 0 ? (count / total) * Math.PI * 2 : 0;
                const endAngle = startAngle + sliceAngle;

                if (count > 0 && radius > 0) {
                    ctx.beginPath();
                    ctx.arc(centerX, centerY, radius, startAngle, endAngle);
                    ctx.arc(centerX, centerY, innerRadius, endAngle, startAngle, true);
                    ctx.closePath();

                    ctx.fillStyle = colors[idx % colors.length];
                    ctx.fill();

                    ctx.strokeStyle = '#1b1f24';
                    ctx.lineWidth = 2;
                    ctx.stroke();
                }

                startAngle = endAngle;
            });

            if (innerRadius > 1) {
                ctx.beginPath();
                ctx.arc(centerX, centerY, innerRadius - 1, 0, Math.PI * 2);
                ctx.fillStyle = '#14171a';
                ctx.fill();
            }

            ctx.fillStyle = '#ffffff';
            ctx.font = 'bold 20px Montserrat';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.fillText(total, centerX, centerY - 6);

            ctx.fillStyle = '#8a94a0';
            ctx.font = '10px Inter';
            ctx.fillText('Contratos', centerX, centerY + 14);

            const legendContainer = document.getElementById('pay-legend');
            legendContainer.innerHTML = '';
            keys.slice(0, 4).forEach((pay, idx) => {
                const cnt = payMap[pay];
                const pct = total > 0 ? ((cnt / total) * 100).toFixed(0) : 0;
                
                const itemDiv = document.createElement('div');
                itemDiv.className = 'flex items-center space-x-2';
                itemDiv.innerHTML = `
                    <span class="w-2.5 h-2.5 rounded-full inline-block" style="background-color: ${colors[idx % colors.length]}"></span>
                    <span class="text-porsche-gray font-medium truncate flex-1" title="${pay}">${pay}</span>
                    <span class="text-white font-bold font-mono">${pct}%</span>
                `;
                legendContainer.appendChild(itemDiv);
            });
        }

        // Chart 4: Top Models Horizontal Revenue
        function renderCanvasTopPopularModels() {
            const setup = setupCanvasPixelRatio('canvas-top-popular-models');
            if (!setup) return;
            const { ctx, width, height } = setup;

            ctx.clearRect(0, 0, width, height);

            const modelMetrics = {};
            filteredData.forEach(d => {
                if (!modelMetrics[d.modelo]) {
                    modelMetrics[d.modelo] = { units: 0, revenue: 0 };
                }
                modelMetrics[d.modelo].units += 1;
                modelMetrics[d.modelo].revenue += d.preco;
            });

            const sortedModels = Object.keys(modelMetrics).sort((a,b) => modelMetrics[b].revenue - modelMetrics[a].revenue);
            const topModels = sortedModels.slice(0, 6);

            const maxRev = topModels.length > 0 ? Math.max(...topModels.map(m => modelMetrics[m].revenue)) : 1;

            const padding = { top: 15, bottom: 20, left: 150, right: 90 };
            const chartW = width - padding.left - padding.right;
            const chartH = height - padding.top - padding.bottom;

            const rowHeight = chartH / Math.max(topModels.length, 1);

            topModels.forEach((modelName, idx) => {
                const metric = modelMetrics[modelName];
                const barWidth = (metric.revenue / maxRev) * chartW;
                const y = padding.top + (idx * rowHeight) + (rowHeight * 0.2);
                const h = Math.max(rowHeight * 0.6, 12);

                ctx.fillStyle = '#e2e8f0';
                ctx.font = '11px Montserrat';
                ctx.textAlign = 'right';
                ctx.textBaseline = 'middle';
                const shortName = modelName.replace("Porsche ", "");
                const truncatedName = shortName.length > 18 ? shortName.substring(0, 16) + '...' : shortName;
                ctx.fillText(truncatedName, padding.left - 12, y + h / 2);

                ctx.fillStyle = '#14171a';
                ctx.beginPath();
                ctx.roundRect(padding.left, y, chartW, h, 4);
                ctx.fill();

                const barGrad = ctx.createLinearGradient(padding.left, 0, padding.left + barWidth, 0);
                if (idx === 0) {
                    barGrad.addColorStop(0, '#d5001c');
                    barGrad.addColorStop(1, '#ff3b53');
                } else {
                    barGrad.addColorStop(0, '#c2a15f');
                    barGrad.addColorStop(1, '#e5ca93');
                }

                ctx.fillStyle = barGrad;
                ctx.beginPath();
                ctx.roundRect(padding.left, y, Math.max(barWidth, 6), h, 4);
                ctx.fill();

                ctx.fillStyle = '#ffffff';
                ctx.font = 'bold 10px Inter';
                ctx.textAlign = 'left';
                ctx.textBaseline = 'middle';
                const revFormatted = (metric.revenue / 1000000).toFixed(2) + 'M';
                ctx.fillText(`US$ ${revFormatted} (${metric.units} u.)`, padding.left + barWidth + 8, y + h / 2);
            });
        }

        function renderTable() {
            const tbody = document.getElementById('sales-table-body');
            tbody.innerHTML = '';

            const totalRows = filteredData.length;
            const totalPages = Math.ceil(totalRows / rowsPerPage) || 1;
            document.getElementById('current-page').innerText = currentPage;
            document.getElementById('total-pages').innerText = totalPages;

            document.getElementById('table-row-counter').innerText = `${totalRows} registros encontrados`;

            document.getElementById('prev-page').disabled = (currentPage === 1);
            document.getElementById('next-page').disabled = (currentPage === totalPages);

            const startIdx = (currentPage - 1) * rowsPerPage;
            const pageData = filteredData.slice(startIdx, startIdx + rowsPerPage);

            if (pageData.length === 0) {
                tbody.innerHTML = `
                    <tr>
                        <td colspan="7" class="p-8 text-center text-porsche-gray italic">
                            Nenhum veículo encontrado com os parâmetros informados.
                        </td>
                    </tr>
                `;
                return;
            }

            pageData.forEach(row => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-porsche-dark/60 transition-colors border-b border-porsche-border/20';

                tr.innerHTML = `
                    <td class="p-3 font-mono text-porsche-gray text-[11px]">${row.id}</td>
                    <td class="p-3 font-semibold text-white">${row.modelo}</td>
                    <td class="p-3 font-mono text-porsche-gold">${row.anoModelo ?? 'N/D'}</td>
                    <td class="p-3">${row.cidade || 'N/D'}</td>
                    <td class="p-3 font-mono font-bold text-porsche-red">${row.estado || 'N/D'}</td>
                    <td class="p-3"><span class="bg-porsche-dark px-2 py-0.5 rounded border border-porsche-border text-[11px]">${row.formaPagamento || 'N/D'}</span></td>
                    <td class="p-3 text-right font-bold text-emerald-400 font-mono">${formatUSD(row.preco)}</td>
                `;
                tbody.appendChild(tr);
            });
        }

        function setupEventListeners() {
            // Drag and Drop Excel Upload Handlers
            const dropzone = document.getElementById('excel-dropzone');
            const fileInput = document.getElementById('excel-file-input');
            const uploadBtn = document.getElementById('btn-trigger-upload');

            uploadBtn.addEventListener('click', () => fileInput.click());

            fileInput.addEventListener('change', (e) => {
                if (e.target.files && e.target.files[0]) {
                    handleExcelFile(e.target.files[0]);
                }
            });

            ['dragenter', 'dragover'].forEach(eventName => {
                dropzone.addEventListener(eventName, (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    dropzone.classList.add('dropzone-active');
                }, false);
            });

            ['dragleave', 'drop'].forEach(eventName => {
                dropzone.addEventListener(eventName, (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    dropzone.classList.remove('dropzone-active');
                }, false);
            });

            dropzone.addEventListener('drop', (e) => {
                const dt = e.dataTransfer;
                const files = dt.files;
                if (files && files[0]) {
                    handleExcelFile(files[0]);
                }
            }, false);

            // Filter Change Listeners
            document.getElementById('filter-modelo').addEventListener('change', applyFilters);
            document.getElementById('filter-year').addEventListener('change', applyFilters);
            document.getElementById('filter-city').addEventListener('change', applyFilters);
            document.getElementById('filter-pay').addEventListener('change', applyFilters);

            // Reset Filters Button
            document.getElementById('reset-filters-btn').addEventListener('click', () => {
                document.getElementById('filter-modelo').value = 'ALL';
                document.getElementById('filter-year').value = 'ALL';
                document.getElementById('filter-city').value = 'ALL';
                document.getElementById('filter-pay').value = 'ALL';
                applyFilters();
            });

            // Pagination Buttons
            document.getElementById('prev-page').addEventListener('click', () => {
                if (currentPage > 1) {
                    currentPage--;
                    renderTable();
                }
            });

            document.getElementById('next-page').addEventListener('click', () => {
                const totalPages = Math.ceil(filteredData.length / rowsPerPage);
                if (currentPage < totalPages) {
                    currentPage++;
                    renderTable();
                }
            });

            // Export CSV
            document.getElementById('export-btn').addEventListener('click', () => {
                let csv = 'ID,Modelo,Model Year,Cidade,Estado,Forma Pagamento,Entrega,Preco USD\n';
                filteredData.forEach(r => {
                    csv += `"${r.id}","${r.modelo}",${r.anoModelo ?? ''},"${r.cidade}","${r.estado}","${r.formaPagamento}","${r.entrega}",${r.preco}\n`;
                });
                
                const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.setAttribute('href', url);
                link.setAttribute('download', `porsche_sales_report_${new Date().toISOString().slice(0,10)}.csv`);
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            });

            // Resize Canvas Handle
            window.addEventListener('resize', () => {
                renderCanvasModelsByCity();
                renderCanvasModelYearTrend();
                renderCanvasPayMethod();
                renderCanvasTopPopularModels();
            });
        }

        // Ingestao automatica da base de producao sanitizada (dashboard_data.js)
        function loadProductionData() {
            const ok = processRows(window.PORSCHE_SANITIZED_DATA || []);
            if (ok) {
                updateDataStatus(true, `Base de Produ\u00e7\u00e3o (${rawSalesData.length} Vendas)`);
            } else {
                updateDataStatus(false, 'Sem dados ativos');
            }
        }

        window.onload = function() {
            loadProductionData();
            setupEventListeners();
        };
    