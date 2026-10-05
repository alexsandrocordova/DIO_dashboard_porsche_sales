# -*- coding: utf-8 -*-
"""Patch cirurgico do dashboard HTML (executado uma unica vez)."""

FILE = "porsche_intelligence_dashboard.html"
src = open(FILE, encoding="utf-8").read()
n_subs = 0


def replace_once(old, new, label):
    global src, n_subs
    assert src.count(old) == 1, f"Marcador nao unico/encontrado: {label} ({src.count(old)})"
    src = src.replace(old, new)
    n_subs += 1
    print("OK  ", label)


# 1) Remove processExcelRows legado -> updateDataStatus reutilizavel
start_marker = "        // Excel Parser & Column Auto-Detector\n"
end_marker = "        // Read File Handler\n"
i = src.index(start_marker)
j = src.index(end_marker)
assert i < j, "Ordem de marcadores inesperada"
new_block = '''        // Atualiza badges de status e rotulo da base ativa
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

'''
src = src[:i] + new_block + src[j:]
n_subs += 1
print("OK   processExcelRows -> updateDataStatus")

# 2) handleExcelFile passa a usar processRows
replace_once(
    "                    processExcelRows(jsonData);",
    """                    if (processRows(jsonData)) {
                        updateDataStatus(true, `Planilha XLSX (${rawSalesData.length} Vendas)`);
                    } else {
                        updateDataStatus(false, 'Nenhum registro valido', `<span class="text-porsche-red font-bold"><i class="fa-solid fa-triangle-exclamation"></i> Nenhum registro valido encontrado no arquivo.</span>`);
                    }""",
    "handleExcelFile usa processRows",
)

# 3) Moeda USD (dados da base em dolares)
replace_once(
    """        // Format Currency BRL
        function formatBRL(value) {
            return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 0 }).format(value);
        }""",
    """        // Format Currency USD (base de producao em dolares)
        function formatUSD(value) {
            return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(value);
        }""",
    "formatBRL -> formatUSD",
)
src = src.replace("formatBRL(totalRevenue)", "formatUSD(totalRevenue)")
src = src.replace("formatBRL(avgTicket)", "formatUSD(avgTicket)")
src = src.replace("${formatBRL(row.preco)}", "${formatUSD(row.preco)}")
n_subs += 3
print("OK   usages formatUSD x3")

# 4) Chart 4: rotulo monetario
replace_once(
    "ctx.fillText(`R$ ${revFormatted} (${metric.units} u.)`",
    "ctx.fillText(`US$ ${revFormatted} (${metric.units} u.)`",
    "chart4 US$",
)


# 5) Tabela: linha com Estado (USPS) e valores nulos protegidos
replace_once(
    """                tr.innerHTML = `
                    <td class="p-3 font-mono text-porsche-gray text-[11px]">${row.id}</td>
                    <td class="p-3 font-semibold text-white">${row.modelo}</td>
                    <td class="p-3 font-mono text-porsche-gold">${row.anoModelo}</td>
                    <td class="p-3">${row.cidade}</td>
                    <td class="p-3"><span class="bg-porsche-dark px-2 py-0.5 rounded border border-porsche-border text-[11px]">${row.formaPagamento}</span></td>
                    <td class="p-3 text-porsche-gray">${row.cor}</td>
                    <td class="p-3 text-right font-bold text-emerald-400 font-mono">${formatUSD(row.preco)}</td>
                `;""",
    """                tr.innerHTML = `
                    <td class="p-3 font-mono text-porsche-gray text-[11px]">${row.id}</td>
                    <td class="p-3 font-semibold text-white">${row.modelo}</td>
                    <td class="p-3 font-mono text-porsche-gold">${row.anoModelo ?? 'N/D'}</td>
                    <td class="p-3">${row.cidade || 'N/D'}</td>
                    <td class="p-3 font-mono font-bold text-porsche-red">${row.estado || 'N/D'}</td>
                    <td class="p-3"><span class="bg-porsche-dark px-2 py-0.5 rounded border border-porsche-border text-[11px]">${row.formaPagamento || 'N/D'}</span></td>
                    <td class="p-3 text-right font-bold text-emerald-400 font-mono">${formatUSD(row.preco)}</td>
                `;""",
    "renderTable linha (estado/US$)",
)

# 6) Chart 2 e Insight Q2: protecao contra anos INVALID (null)
replace_once(
    """            const yearMap = {};
            filteredData.forEach(d => {
                yearMap[d.anoModelo] = (yearMap[d.anoModelo] || 0) + 1;
            });""",
    """            const yearMap = {};
            filteredData.forEach(d => {
                if (d.anoModelo) yearMap[d.anoModelo] = (yearMap[d.anoModelo] || 0) + 1;
            });""",
    "chart2 anos nulos",
)
replace_once(
    """            const yearCounts = {};
            filteredData.forEach(d => {
                yearCounts[d.anoModelo] = (yearCounts[d.anoModelo] || 0) + 1;
            });""",
    """            const yearCounts = {};
            filteredData.forEach(d => {
                if (d.anoModelo) yearCounts[d.anoModelo] = (yearCounts[d.anoModelo] || 0) + 1;
            });""",
    "insight-q2 anos nulos",
)

# 7) Export CSV: Estado/Entrega no lugar de Cor
replace_once(
    """                let csv = 'ID,Modelo,Model Year,Cidade,Forma Pagamento,Cor,Preco\\n';
                filteredData.forEach(r => {
                    csv += `"${r.id}","${r.modelo}",${r.anoModelo},"${r.cidade}","${r.formaPagamento}","${r.cor}",${r.preco}\\n`;
                });""",
    """                let csv = 'ID,Modelo,Model Year,Cidade,Estado,Forma Pagamento,Entrega,Preco USD\\n';
                filteredData.forEach(r => {
                    csv += `"${r.id}","${r.modelo}",${r.anoModelo ?? ''},"${r.cidade}","${r.estado}","${r.formaPagamento}","${r.entrega}",${r.preco}\\n`;
                });""",
    "export CSV",
)

# 8) Boot: carrega automaticamente a base de producao sanitizada
replace_once(
    """        window.onload = function() {
            populateFilterOptions();
            setupEventListeners();
            applyFilters();
        };""",
    """        // Ingestao automatica da base de producao sanitizada (dashboard_data.js)
        function loadProductionData() {
            const ok = processRows(window.PORSCHE_SANITIZED_DATA || []);
            if (ok) {
                updateDataStatus(true, `Base de Produ\\u00e7\\u00e3o (${rawSalesData.length} Vendas)`);
            } else {
                updateDataStatus(false, 'Sem dados ativos');
            }
        }

        window.onload = function() {
            loadProductionData();
            setupEventListeners();
        };""",
    "window.onload auto-load",
)

open(FILE, "w", encoding="utf-8").write(src)
print(f"\nPatch concluido: {n_subs} substituicoes.")
