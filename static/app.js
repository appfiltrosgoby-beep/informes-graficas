"use strict";

const STORAGE_KEY = "goby_report_history_v1";
const BASE_TOTAL_CLIENTES_DEFAULT = 55;
const BASE_TOTAL_CLIENTES_STORAGE_KEY = "goby_report_base_total_clientes_v1";

const uploadForm = document.getElementById("uploadForm");
const monthInput = document.getElementById("reportMonth");
const baseClientesInput = document.getElementById("baseClientes");
const excelInput = document.getElementById("excelFile");
const downloadPdfBtn = document.getElementById("downloadPdfBtn");
const clearStorageBtn = document.getElementById("clearStorageBtn");
const messages = document.getElementById("messages");

const chart1El = document.getElementById("grafica1");
const chart2El = document.getElementById("grafica2");
const chart3El = document.getElementById("grafica3");
const tableContainer = document.getElementById("tablaContainer");

let currentActiveMonth = "";

const MONTH_NAMES = {
	"01": "Enero",
	"02": "Febrero",
	"03": "Marzo",
	"04": "Abril",
	"05": "Mayo",
	"06": "Junio",
	"07": "Julio",
	"08": "Agosto",
	"09": "Septiembre",
	"10": "Octubre",
	"11": "Noviembre",
	"12": "Diciembre",
};

function setCurrentMonthDefault() {
	if (monthInput.value) {
		return;
	}
	const now = new Date();
	monthInput.value = String(now.getMonth() + 1).padStart(2, "0");
}

function showMessage(text, type = "info") {
	messages.innerHTML = `<p class="msg ${type}">${text}</p>`;
}

function normalizeText(value) {
	return String(value || "").trim();
}

function normalizeKey(value) {
	return normalizeText(value).toLowerCase();
}

function isTotalRowName(value) {
	const raw = normalizeText(value);
	if (!raw) {
		return false;
	}
	return /^total\b/i.test(raw);
}

function parseNumber(value) {
	if (value === null || value === undefined || value === "") {
		return 0;
	}
	if (typeof value === "number") {
		return Number.isFinite(value) ? value : 0;
	}

	let text = String(value).trim();
	if (!text) {
		return 0;
	}

	const hasComma = text.includes(",");
	const hasDot = text.includes(".");
	if (hasComma && hasDot) {
		text = text.replace(/\./g, "").replace(/,/g, ".");
	} else if (hasComma) {
		text = text.replace(/,/g, ".");
	}

	const num = Number(text);
	return Number.isFinite(num) ? num : 0;
}

function normalizeMonthKey(value) {
	const raw = String(value || "").trim();
	if (/^\d{2}$/.test(raw)) {
		return raw;
	}
	const fromYearMonth = raw.match(/^\d{4}-(\d{2})$/);
	if (fromYearMonth) {
		return fromYearMonth[1];
	}
	return "";
}

function monthLabel(monthKey) {
	const normalized = normalizeMonthKey(monthKey);
	return MONTH_NAMES[normalized] || String(monthKey || "");
}

function getThemeColor(name, fallback) {
	const value = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
	return value || fallback;
}

function loadHistory() {
	try {
		const raw = localStorage.getItem(STORAGE_KEY);
		if (!raw) {
			return { months: {} };
		}
		const parsed = JSON.parse(raw);
		if (!parsed || typeof parsed !== "object" || !parsed.months) {
			return { months: {} };
		}

		const normalizedMonths = {};
		Object.keys(parsed.months).forEach((key) => {
			const normalized = normalizeMonthKey(key);
			if (!normalized) {
				return;
			}
			normalizedMonths[normalized] = parsed.months[key];
		});

		return { months: normalizedMonths };
	} catch (error) {
		return { months: {} };
	}
}

function saveHistory(history) {
	localStorage.setItem(STORAGE_KEY, JSON.stringify(history));
}

function sanitizeBaseTotal(value) {
	const parsed = Number(value);
	if (!Number.isFinite(parsed)) {
		return BASE_TOTAL_CLIENTES_DEFAULT;
	}
	return Math.max(Math.round(parsed), 0);
}

function loadBaseTotalClientes() {
	const raw = localStorage.getItem(BASE_TOTAL_CLIENTES_STORAGE_KEY);
	return sanitizeBaseTotal(raw);
}

function saveBaseTotalClientes(value) {
	const base = sanitizeBaseTotal(value);
	localStorage.setItem(BASE_TOTAL_CLIENTES_STORAGE_KEY, String(base));
}

function getSelectedBaseTotalClientes() {
	return sanitizeBaseTotal(baseClientesInput?.value);
}

function recalculateClientesTotales(history, baseTotalClientes) {
	const months = sortedMonthKeys(history);
	const fallbackBaseTotal = sanitizeBaseTotal(baseTotalClientes);

	months.forEach((month) => {
		const monthData = history.months[month] || {};
		const clientesNuevos = Math.max(
			Number(monthData.clientesNuevos ?? monthData.clientesTotalesMes ?? monthData.clientesTotales ?? 0),
			0
		);
		const baseMes = sanitizeBaseTotal(monthData.baseClientesMes ?? fallbackBaseTotal);

		monthData.clientesTotalesMes = clientesNuevos;
		monthData.baseClientesMes = baseMes;
		monthData.clientesTotales = baseMes;

		history.months[month] = monthData;
	});
}

function extractReferenceFromProduct(productValue) {
	const text = normalizeText(productValue).replace(/\s+/g, " ");
	if (!text) {
		return "";
	}

	const upper = text.toUpperCase();
	const filtroIndex = upper.indexOf("FILTRO");
	if (filtroIndex <= 0) {
		return text;
	}

	return text.slice(0, filtroIndex).trim();
}

function detectHeaderRowIndex(rows) {
	const maxScan = Math.min(rows.length, 35);
	for (let i = 0; i < maxScan; i += 1) {
		const row = rows[i] || [];
		const a = normalizeKey(row[0]);
		const b = normalizeKey(row[1]);
		const c = normalizeKey(row[2]);
		const hasExpectedHeaders =
			a.includes("nombre") && b.includes("producto") && (c.includes("cantidad") || c.includes("cant"));
		if (hasExpectedHeaders) {
			return i;
		}
	}
	return -1;
}

function extractCompanyFromTotalRow(totalText, currentCustomer) {
	const cleaned = normalizeText(totalText).replace(/^total\s*/i, "").trim();
	if (cleaned) {
		return cleaned;
	}
	return normalizeText(currentCustomer) || "Sin nombre";
}

function extractFromSheetRows(rows, aggregated) {
	if (!Array.isArray(rows) || rows.length === 0) {
		return;
	}

	const headerIndex = detectHeaderRowIndex(rows);
	const startIndex = headerIndex >= 0 ? headerIndex + 1 : 0;
	let currentCustomer = "";

	for (let idx = startIndex; idx < rows.length; idx += 1) {
		const row = rows[idx] || [];
		const nameValue = normalizeText(row[0]);
		const productValue = normalizeText(row[1]);
		const qtyValue = parseNumber(row[2]);

		const emptyRow = !nameValue && !productValue && !qtyValue;
		if (emptyRow) {
			continue;
		}

		const isTotalRow = isTotalRowName(nameValue);
		if (nameValue && !isTotalRow) {
			currentCustomer = nameValue;
			aggregated.customerSet.add(normalizeKey(nameValue));
		}

		if (isTotalRow) {
			const empresa = extractCompanyFromTotalRow(nameValue, currentCustomer);
			const companyKey = normalizeKey(empresa);
			if (!aggregated.companyTotalsMap[companyKey]) {
				aggregated.companyTotalsMap[companyKey] = { empresa, total: 0 };
			}
			aggregated.companyTotalsMap[companyKey].total += Math.max(qtyValue, 0);
			currentCustomer = "";
			continue;
		}

		const activeCustomer = normalizeText(nameValue || currentCustomer);
		if (activeCustomer && (productValue || qtyValue > 0)) {
			aggregated.customersWithOrder.add(normalizeKey(activeCustomer));
		}

		if (productValue) {
			aggregated.productRowCount += 1;
			const reference = extractReferenceFromProduct(productValue);
			if (reference) {
				aggregated.soldByReference[reference] =
					(aggregated.soldByReference[reference] || 0) + Math.max(qtyValue, 0);
			}
		}
	}
}

function extractFromWorkbook(workbook) {
	const aggregated = {
		customerSet: new Set(),
		customersWithOrder: new Set(),
		productRowCount: 0,
		soldByReference: {},
		companyTotalsMap: {},
	};

	(workbook.SheetNames || []).forEach((sheetName) => {
		const sheet = workbook.Sheets[sheetName];
		if (!sheet) {
			return;
		}

		const rows = XLSX.utils.sheet_to_json(sheet, {
			header: 1,
			defval: "",
			blankrows: false,
			raw: false,
		});

		extractFromSheetRows(rows, aggregated);
	});

	return {
		clientesTotales: aggregated.customerSet.size,
		clientesNuevos: aggregated.customerSet.size,
		clientesTotalesMes: aggregated.customerSet.size,
		clientesConPedido: aggregated.customersWithOrder.size,
		totalProductos: aggregated.productRowCount,
		vendidosPorReferencia: aggregated.soldByReference,
		tablaEmpresas: Object.values(aggregated.companyTotalsMap),
		hojasProcesadas: (workbook.SheetNames || []).length,
	};
}

function parseWorkbook(fileData) {
	const workbook = XLSX.read(fileData, { type: "array" });
	if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
		throw new Error("El archivo no contiene hojas.");
	}

	return extractFromWorkbook(workbook);
}

function sortedMonthKeys(history) {
	return Object.keys(history.months || {}).sort((a, b) => Number(a) - Number(b));
}

function updateDownloadButtonState(history) {
	if (!downloadPdfBtn) {
		return;
	}
	const hasHistory = sortedMonthKeys(history).length > 0;
	downloadPdfBtn.disabled = !hasHistory;
}

async function generateDashboardPdf() {
	if (!downloadPdfBtn) {
		return;
	}

	if (!window.html2canvas || !window.jspdf?.jsPDF) {
		showMessage("No se pudo iniciar la descarga PDF: faltan librerias de exportacion.", "error");
		return;
	}

	const panels = [
		{ title: "Clientes totales vs clientes con pedido", element: chart1El?.closest(".panel") },
		{ title: "Total de productos por mes", element: chart2El?.closest(".panel") },
		{ title: "Cantidad vendida por referencia", element: chart3El?.closest(".panel") },
		{ title: "Total de unidades por cliente del mes", element: tableContainer?.closest(".panel") },
		{ title: "Comparativa de empresas por mes", element: document.getElementById("tablaComparativaContainer")?.closest(".panel") },
	].filter((item) => item.element);

	if (!panels.length) {
		showMessage("No hay contenido para exportar a PDF.", "warn");
		return;
	}

	downloadPdfBtn.disabled = true;
	showMessage("Generando PDF...", "info");

	try {
		const { jsPDF } = window.jspdf;
		const doc = new jsPDF({ orientation: "l", unit: "mm", format: "a4" });
		const pageWidth = doc.internal.pageSize.getWidth();
		const pageHeight = doc.internal.pageSize.getHeight();
		const sideMargin = 10;
		const topMargin = 14;
		const bottomMargin = 10;
		const slotGap = 6;
		const availableWidth = pageWidth - sideMargin * 2;
		const slotHeight = (pageHeight - topMargin - bottomMargin - slotGap) / 2;
		const monthText = monthLabel(currentActiveMonth || normalizeMonthKey(monthInput.value));

		for (let i = 0; i < panels.length; i += 1) {
			const panel = panels[i];
			if (i > 0 && i % 2 === 0) {
				doc.addPage();
			}

			const slotIndex = i % 2;
			const slotStartY = topMargin + slotIndex * (slotHeight + slotGap);

			doc.setFontSize(11);
			doc.text(`Reporte ${monthText}`, sideMargin, 8);
			doc.setFontSize(14);
			doc.text(panel.title, sideMargin, slotStartY + 5);

			const canvas = await window.html2canvas(panel.element, {
				scale: 2,
				useCORS: true,
				backgroundColor: "#ffffff",
				scrollY: -window.scrollY,
			});

			const imageData = canvas.toDataURL("image/png", 1.0);
			const maxImgWidth = availableWidth;
			const maxImgHeight = slotHeight - 10;
			const widthRatio = maxImgWidth / canvas.width;
			const heightRatio = maxImgHeight / canvas.height;
			const ratio = Math.min(widthRatio, heightRatio);
			const finalWidth = canvas.width * ratio;
			const finalHeight = canvas.height * ratio;
			const imgX = sideMargin + (maxImgWidth - finalWidth) / 2;
			const imgY = slotStartY + 8;

			doc.addImage(imageData, "PNG", imgX, imgY, finalWidth, finalHeight, undefined, "FAST");
		}

		const fileMonth = normalizeMonthKey(currentActiveMonth || monthInput.value) || "mes";
		doc.save(`reporte-graficas-${fileMonth}.pdf`);
		showMessage("PDF generado correctamente.", "ok");
	} catch (error) {
		const detail = error instanceof Error ? error.message : "Error desconocido.";
		showMessage(`No fue posible generar el PDF: ${detail}`, "error");
	} finally {
		const history = loadHistory();
		updateDownloadButtonState(history);
	}
}

function renderChart1(history) {
	const months = sortedMonthKeys(history);
	const labels = months.map(monthLabel);
	const clientesTotales = months.map((key) => history.months[key].clientesTotales || 0);
	const clientesConPedido = months.map((key) => history.months[key].clientesConPedido || 0);

	if (!months.length) {
		chart1El.innerHTML = '<p class="empty">Sin datos para la grafica 1.</p>';
		return;
	}

	Plotly.newPlot(
		chart1El,
		[
			{
				x: labels,
				y: clientesTotales,
				name: "Clientes totales",
				type: "bar",
				marker: { color: getThemeColor("--chart-1", "#2d6cdf") },
			},
			{
				x: labels,
				y: clientesConPedido,
				name: "Clientes con pedido",
				type: "bar",
				marker: { color: getThemeColor("--chart-2", "#f08c2e") },
			},
		],
		{
			barmode: "group",
			margin: { t: 20, r: 20, l: 50, b: 45 },
			paper_bgcolor: "rgba(0,0,0,0)",
			plot_bgcolor: "rgba(255,255,255,0.7)",
			yaxis: { title: "Cantidad" },
			xaxis: { title: "Mes" },
		},
		{ responsive: true, displayModeBar: false }
	);
}

function renderChart2(history) {
	const months = sortedMonthKeys(history);
	const labels = months.map(monthLabel);
	const totalProductos = months.map((key) => history.months[key].totalProductos || 0);

	if (!months.length) {
		chart2El.innerHTML = '<p class="empty">Sin datos para la grafica 2.</p>';
		return;
	}

	Plotly.newPlot(
		chart2El,
		[
			{
				x: labels,
				y: totalProductos,
				type: "bar",
				marker: { color: getThemeColor("--chart-3", "#22a574") },
				name: "Total productos",
			},
		],
		{
			margin: { t: 20, r: 20, l: 50, b: 45 },
			paper_bgcolor: "rgba(0,0,0,0)",
			plot_bgcolor: "rgba(255,255,255,0.7)",
			yaxis: { title: "Cantidad de productos" },
			xaxis: { title: "Mes" },
		},
		{ responsive: true, displayModeBar: false }
	);
}

function renderChart3(currentMonthData) {
	const entries = Object.entries(currentMonthData.vendidosPorReferencia || {})
		.map(([referencia, total]) => ({ referencia, total }))
		.sort((a, b) => b.total - a.total);

	if (!entries.length) {
		chart3El.innerHTML = '<p class="empty">Sin datos para la grafica 3.</p>';
		return;
	}

	Plotly.newPlot(
		chart3El,
		[
			{
				x: entries.map((item) => item.referencia),
				y: entries.map((item) => item.total),
				type: "bar",
				marker: {
					color: entries.map((_, idx) => {
						const palette = [
							getThemeColor("--chart-1", "#2d6cdf"),
							getThemeColor("--chart-2", "#f08c2e"),
							getThemeColor("--chart-3", "#22a574"),
							getThemeColor("--chart-4", "#d64550"),
							getThemeColor("--chart-5", "#7f56d9"),
							getThemeColor("--chart-6", "#009fb7"),
							getThemeColor("--chart-7", "#f2c94c"),
							getThemeColor("--chart-8", "#5c677d"),
						];
						return palette[idx % palette.length];
					}),
				},
				name: "Cantidad vendida",
			},
		],
		{
			margin: { t: 20, r: 20, l: 50, b: 90 },
			paper_bgcolor: "rgba(0,0,0,0)",
			plot_bgcolor: "rgba(255,255,255,0.7)",
			yaxis: { title: "Unidades" },
			xaxis: { title: "Referencia", tickangle: -30 },
		},
		{ responsive: true, displayModeBar: false }
	);
}

function renderTable4(currentMonthData) {
	const rows = (currentMonthData.tablaEmpresas || []).slice().sort((a, b) => b.total - a.total);
	if (!rows.length) {
		tableContainer.innerHTML = '<p class="empty">Sin filas total para la grafica 4.</p>';
		return;
	}

	const body = rows
		.map(
			(item) =>
				`<tr><td>${item.empresa}</td><td class="num">${Number(item.total || 0).toLocaleString("es-CO")}</td></tr>`
		)
		.join("");

	tableContainer.innerHTML = `
		<table>
			<thead>
				<tr>
					<th>Empresa</th>
					<th>Total unidades</th>
				</tr>
			</thead>
			<tbody>
				${body}
			</tbody>
		</table>
	`;
}

function renderTableComparativa(history) {
	const months = sortedMonthKeys(history);
	if (!months.length) {
		const comparativaContainer = document.getElementById("tablaComparativaContainer");
		if (comparativaContainer) {
			comparativaContainer.innerHTML = '<p class="empty">Sin datos para la tabla comparativa.</p>';
		}
		return;
	}

	// Recolectar todas las empresas de todos los meses
	const empresasMap = {};
	months.forEach((month) => {
		const monthData = history.months[month] || {};
		const tablaEmpresas = monthData.tablaEmpresas || [];
		tablaEmpresas.forEach((item) => {
			const empresaKey = normalizeKey(item.empresa || "");
			if (!empresaKey) return;
			if (!empresasMap[empresaKey]) {
				empresasMap[empresaKey] = { displayName: item.empresa, monthData: {} };
			}
			empresasMap[empresaKey].monthData[month] = item.total || 0;
		});
	});

	const empresas = Object.values(empresasMap).sort((a, b) => {
		const totalA = Object.values(a.monthData).reduce((s, v) => s + v, 0);
		const totalB = Object.values(b.monthData).reduce((s, v) => s + v, 0);
		return totalB - totalA;
	});

	if (!empresas.length) {
		const comparativaContainer = document.getElementById("tablaComparativaContainer");
		if (comparativaContainer) {
			comparativaContainer.innerHTML = '<p class="empty">Sin empresas para mostrar la comparativa.</p>';
		}
		return;
	}

	// Construir headers con meses
	const monthHeaders = months.map((m) => `<th>${monthLabel(m)}</th>`).join("");

	// Construir filas con empresas
	const filas = empresas
		.map((empresa) => {
			const celdas = months.map((month) => {
				const valor = empresa.monthData[month] || 0;
				return `<td class="num">${Number(valor).toLocaleString("es-CO")}</td>`;
			}).join("");

			const totalEmpresa = Object.values(empresa.monthData).reduce((s, v) => s + v, 0);
			return `<tr>
				<td><strong>${empresa.displayName}</strong></td>
				${celdas}
				<td class="num"><strong>${Number(totalEmpresa).toLocaleString("es-CO")}</strong></td>
			</tr>`;
		})
		.join("");

	// Fila de totales por mes
	const totalesMes = months
		.map((month) => {
			const total = empresas.reduce((sum, emp) => sum + (emp.monthData[month] || 0), 0);
			return `<td class="num"><strong>${Number(total).toLocaleString("es-CO")}</strong></td>`;
		})
		.join("");

	const totalGeneral = empresas.reduce((sum, emp) => sum + Object.values(emp.monthData).reduce((s, v) => s + v, 0), 0);

	const html = `
		<table>
			<thead>
				<tr>
					<th>Empresa</th>
					${monthHeaders}
					<th>Total</th>
				</tr>
			</thead>
			<tbody>
				${filas}
				<tr style="background: rgba(31, 111, 144, 0.1);">
					<td><strong>TOTAL</strong></td>
					${totalesMes}
					<td class="num"><strong>${Number(totalGeneral).toLocaleString("es-CO")}</strong></td>
				</tr>
			</tbody>
		</table>
	`;

	const comparativaContainer = document.getElementById("tablaComparativaContainer");
	if (comparativaContainer) {
		comparativaContainer.innerHTML = html;
	}
}

function renderAll(history, activeMonth) {
	currentActiveMonth = activeMonth;
	const activeMonthData = history.months[activeMonth];
	renderChart1(history);
	renderChart2(history);
	if (activeMonthData) {
		renderChart3(activeMonthData);
		renderTable4(activeMonthData);
	}
	renderTableComparativa(history);
}

uploadForm?.addEventListener("submit", async (event) => {
	event.preventDefault();

	const file = excelInput.files?.[0];
	const month = normalizeMonthKey(monthInput.value);

	if (!month) {
		showMessage("Selecciona el mes del reporte.", "warn");
		return;
	}
	if (!file) {
		showMessage("Selecciona un archivo Excel.", "warn");
		return;
	}

	try {
		showMessage("Procesando archivo...", "info");
		const arrayBuffer = await file.arrayBuffer();
		const monthData = parseWorkbook(arrayBuffer);
		const baseTotal = getSelectedBaseTotalClientes();
		monthData.baseClientesMes = baseTotal;
		monthData.clientesTotales = baseTotal;

		const history = loadHistory();
		history.months[month] = monthData;
		recalculateClientesTotales(history, baseTotal);
		saveHistory(history);
		saveBaseTotalClientes(baseTotal);
		updateDownloadButtonState(history);

		renderAll(history, month);
		showMessage(
			`Reporte de ${monthLabel(month)} actualizado correctamente. Hojas leidas: ${monthData.hojasProcesadas || 0}.`,
			"ok"
		);
	} catch (error) {
		const detail = error instanceof Error ? error.message : "Error desconocido.";
		showMessage(`No fue posible procesar el archivo: ${detail}`, "error");
	}
});

setCurrentMonthDefault();
const baseTotalInicial = loadBaseTotalClientes();
if (baseClientesInput) {
	baseClientesInput.value = String(baseTotalInicial);
}
const savedHistory = loadHistory();
recalculateClientesTotales(savedHistory, baseTotalInicial);
saveHistory(savedHistory);
updateDownloadButtonState(savedHistory);
const months = sortedMonthKeys(savedHistory);
if (months.length) {
	const lastMonth = months[months.length - 1];
	monthInput.value = normalizeMonthKey(lastMonth);
	renderAll(savedHistory, lastMonth);
	showMessage(`Historico cargado. Mes activo: ${monthLabel(lastMonth)}.`, "info");
} else {
	showMessage("Carga el Excel del mes para iniciar el historico.", "info");
}

baseClientesInput?.addEventListener("change", () => {
	const baseTotal = getSelectedBaseTotalClientes();
	saveBaseTotalClientes(baseTotal);

	const availableMonths = sortedMonthKeys(savedHistory);
	if (!availableMonths.length) {
		showMessage(`Base de clientes actualizada a ${baseTotal}.`, "info");
		return;
	}

	const selectedMonth = normalizeMonthKey(monthInput.value);
	const activeMonth = availableMonths.includes(selectedMonth)
		? selectedMonth
		: availableMonths[availableMonths.length - 1];

	if (savedHistory.months[activeMonth]) {
		savedHistory.months[activeMonth].baseClientesMes = baseTotal;
		savedHistory.months[activeMonth].clientesTotales = baseTotal;
	}
	recalculateClientesTotales(savedHistory, baseTotal);
	saveHistory(savedHistory);

	renderAll(savedHistory, activeMonth);
	showMessage(`Base de clientes actualizada a ${baseTotal}.`, "ok");
});

downloadPdfBtn?.addEventListener("click", () => {
	void generateDashboardPdf();
});

clearStorageBtn?.addEventListener("click", () => {
	if (confirm("¿Estás seguro de que deseas eliminar todo el almacenamiento? Esta acción no se puede deshacer.")) {
		localStorage.clear();
		showMessage("Almacenamiento eliminado correctamente. Recargando...", "ok");
		setTimeout(() => location.reload(), 1500);
	}
});
