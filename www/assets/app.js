(() => {
    "use strict";

    const unitInput = document.getElementById("glucose_unit");
    const glucoseInput = document.getElementById("glucose_value");
    if (unitInput && glucoseInput) {
        const updateUnits = () => {
            glucoseInput.max = unitInput.value === "mg" ? "2000" : "111";
            glucoseInput.placeholder = unitInput.value === "mg" ? "e.g. 101" : "e.g. 5.6";
        };
        unitInput.addEventListener("change", updateUnits);
        updateUnits();
    }

    document.querySelectorAll("[data-saving-form]").forEach(form => {
        form.addEventListener("submit", () => {
            const button = form.querySelector('button[type="submit"]');
            button.dataset.originalText = button.textContent;
            button.disabled = true;
            button.textContent = "Saving...";
        });
    });
    addEventListener("pageshow", () => {
        document.querySelectorAll("button[data-original-text]").forEach(button => {
            button.disabled = false;
            button.textContent = button.dataset.originalText;
        });
    });
    const errorSummary = document.querySelector("[data-error-summary]");
    if (errorSummary) errorSummary.focus();

    const chartData = document.getElementById("trend-data");
    if (!chartData) return;
    const data = JSON.parse(chartData.textContent);
    const ns = "http://www.w3.org/2000/svg";
    const format = new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 });
    const shortDate = new Intl.DateTimeFormat(undefined, { day: "numeric", month: "short" });
    const fullDate = new Intl.DateTimeFormat(undefined, {
        day: "numeric", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit"
    });

    function element(name, attributes, text) {
        const node = document.createElementNS(ns, name);
        Object.entries(attributes).forEach(([key, value]) => node.setAttribute(key, value));
        if (text !== undefined) node.textContent = text;
        return node;
    }

    function niceStep(value) {
        const power = 10 ** Math.floor(Math.log10(value));
        const fraction = value / power;
        const nice = fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 2.5 ? 2.5 : fraction <= 5 ? 5 : 10;
        return nice * power;
    }

    function drawChart(id, entries, series, unit, title) {
        if (entries.length === 0) return;
        const points = entries.slice().reverse().map(entry => ({
            ...entry, timestamp: new Date(entry.date).getTime()
        }));
        const container = document.getElementById(id);
        const width = 620;
        const height = 235;
        const margin = { left: 47, right: 18, top: 17, bottom: 34 };
        const plotWidth = width - margin.left - margin.right;
        const plotHeight = height - margin.top - margin.bottom;
        const values = points.flatMap(point => series.map(item => point[item.key]));
        const lowest = Math.min(...values);
        const highest = Math.max(...values);
        const padding = highest === lowest ? Math.max(highest * 0.1, 0.5) : (highest - lowest) * 0.18;
        const step = niceStep((highest - lowest + padding * 2) / 4);
        const bottom = Math.max(0, Math.floor((lowest - padding) / step) * step);
        const top = Math.ceil((highest + padding) / step) * step;
        const first = points[0].timestamp;
        const last = points[points.length - 1].timestamp;
        const x = time => margin.left + (last === first ? plotWidth / 2 : (time - first) / (last - first) * plotWidth);
        const y = value => margin.top + (top - value) / (top - bottom) * plotHeight;
        const svg = element("svg", {
            viewBox: `0 0 ${width} ${height}`, class: "trend-chart", role: "img",
            "aria-labelledby": `${id}-title ${id}-description`
        });
        svg.append(element("title", { id: `${id}-title` }, title));
        svg.append(element("desc", { id: `${id}-description` },
            `${points.length} readings from ${fullDate.format(first)} to ${fullDate.format(last)}. ` +
            `Values range from ${format.format(lowest)} to ${format.format(highest)} ${unit}. ` +
            "All individual readings are available in reading history."));

        const tickCount = Math.round((top - bottom) / step);
        for (let index = 0; index <= tickCount; index++) {
            const value = bottom + step * index;
            const position = y(value);
            svg.append(element("line", {
                x1: margin.left, x2: width - margin.right, y1: position, y2: position, class: "chart-grid"
            }));
            svg.append(element("text", {
                x: margin.left - 10, y: position + 4, "text-anchor": "end", class: "chart-label"
            }, format.format(value)));
        }

        const times = first === last ? [first] : [first, first + (last - first) / 2, last];
        times.forEach((time, index) => {
            const label = last - first < 86400000
                ? new Intl.DateTimeFormat(undefined, { hour: "2-digit", minute: "2-digit" }).format(time)
                : shortDate.format(time);
            svg.append(element("text", {
                x: x(time), y: height - 8,
                "text-anchor": times.length === 1 ? "middle" : index === 0 ? "start" : index === times.length - 1 ? "end" : "middle",
                class: "chart-label"
            }, label));
        });

        series.forEach(item => {
            const path = points.map((point, index) =>
                `${index === 0 ? "M" : "L"}${x(point.timestamp).toFixed(2)},${y(point[item.key]).toFixed(2)}`
            ).join(" ");
            svg.append(element("path", { d: path, class: `chart-line line-${item.className}` }));
            points.forEach(point => {
                const dot = element("circle", {
                    cx: x(point.timestamp), cy: y(point[item.key]), r: points.length > 60 ? 2.5 : 3.7,
                    class: `chart-point point-${item.className}`
                });
                dot.append(element("title", {},
                    `${fullDate.format(point.timestamp)} - ${item.label}: ${format.format(point[item.key])} ${unit}`));
                svg.append(dot);
            });
        });
        container.replaceChildren(svg);
    }

    drawChart("pressure-chart", data.pressure, [
        { key: "systolic", label: "Systolic", className: "systolic" },
        { key: "diastolic", label: "Diastolic", className: "diastolic" }
    ], "mmHg", "Blood pressure trend");
    drawChart("glucose-chart", data.glucose, [
        { key: "value", label: "Glucose", className: "glucose" }
    ], data.unit, "Blood glucose trend");
})();
