/* ============================================================
   miniChart.js - Graficos con Canvas puro, sin dependencias
   Soporta: bar, doughnut
   ============================================================ */

class Chart {
    constructor(canvas, config) {
        this.canvas = canvas;
        this.config = config;
        this.ctx = canvas.getContext('2d');
        this._resize();
        this.render();
        this._onResize = () => { this._resize(); this.render(); };
        window.addEventListener('resize', this._onResize);
    }

    _resize() {
        const parent = this.canvas.parentElement;
        this.canvas.width = parent.clientWidth - 48 || 300;
        this.canvas.height = 220;
    }

    destroy() {
        window.removeEventListener('resize', this._onResize);
        this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
    }

    render() {
        const type = this.config.type;
        if (type === 'bar') this._renderBar();
        else if (type === 'doughnut' || type === 'pie') this._renderDoughnut();
    }

    _renderBar() {
        const ctx = this.ctx;
        const W = this.canvas.width;
        const H = this.canvas.height;
        ctx.clearRect(0, 0, W, H);

        const data = this.config.data;
        const labels = data.labels || [];
        const dataset = data.datasets[0];
        const values = dataset.data || [];
        const colors = Array.isArray(dataset.backgroundColor)
            ? dataset.backgroundColor
            : values.map(() => dataset.backgroundColor || '#3b82f6');

        if (values.length === 0) return;

        const padL = 60, padR = 16, padT = 16, padB = 40;
        const chartW = W - padL - padR;
        const chartH = H - padT - padB;

        const maxVal = Math.max(...values.map(Math.abs), 1);
        const minVal = Math.min(0, ...values);
        const range = maxVal - minVal;

        // Zero line Y position
        const zeroY = padT + chartH * (1 - (0 - minVal) / (range || 1));

        // Grid lines
        ctx.strokeStyle = '#f3f4f6';
        ctx.lineWidth = 1;
        const steps = 4;
        for (let i = 0; i <= steps; i++) {
            const y = padT + (chartH / steps) * i;
            ctx.beginPath();
            ctx.moveTo(padL, y);
            ctx.lineTo(W - padR, y);
            ctx.stroke();

            // Y axis labels
            const val = maxVal - (maxVal - minVal) * (i / steps);
            ctx.fillStyle = '#9ca3af';
            ctx.font = '10px -apple-system, sans-serif';
            ctx.textAlign = 'right';
            ctx.fillText(this._fmtNum(val), padL - 6, y + 4);
        }

        // Bars
        const barW = Math.min((chartW / values.length) * 0.6, 80);
        const gap = chartW / values.length;

        values.forEach((val, i) => {
            const x = padL + gap * i + gap / 2 - barW / 2;
            const barH = Math.abs(val) / (range || 1) * chartH;
            const y = val >= 0 ? zeroY - barH : zeroY;

            // Bar with rounded top
            ctx.fillStyle = colors[i] || '#3b82f6';
            ctx.beginPath();
            const r = Math.min(4, barH / 2);
            if (val >= 0) {
                ctx.moveTo(x + r, y);
                ctx.lineTo(x + barW - r, y);
                ctx.quadraticCurveTo(x + barW, y, x + barW, y + r);
                ctx.lineTo(x + barW, y + barH);
                ctx.lineTo(x, y + barH);
                ctx.lineTo(x, y + r);
                ctx.quadraticCurveTo(x, y, x + r, y);
            } else {
                ctx.moveTo(x, y);
                ctx.lineTo(x + barW, y);
                ctx.lineTo(x + barW, y + barH - r);
                ctx.quadraticCurveTo(x + barW, y + barH, x + barW - r, y + barH);
                ctx.lineTo(x + r, y + barH);
                ctx.quadraticCurveTo(x, y + barH, x, y + barH - r);
                ctx.lineTo(x, y);
            }
            ctx.fill();

            // Value label on bar
            ctx.fillStyle = '#374151';
            ctx.font = 'bold 10px -apple-system, sans-serif';
            ctx.textAlign = 'center';
            const labelY = val >= 0 ? y - 4 : y + barH + 12;
            ctx.fillText(this._fmtNum(val), x + barW / 2, labelY);

            // X axis label
            ctx.fillStyle = '#6b7280';
            ctx.font = '11px -apple-system, sans-serif';
            ctx.textAlign = 'center';
            const label = (labels[i] || '').toString();
            // Wrap long labels
            const maxLabelW = gap - 4;
            if (ctx.measureText(label).width > maxLabelW) {
                const words = label.split(' ');
                let line = '';
                let lineY = H - padB + 14;
                for (const word of words) {
                    const test = line ? line + ' ' + word : word;
                    if (ctx.measureText(test).width > maxLabelW && line) {
                        ctx.fillText(line, x + barW / 2, lineY);
                        line = word;
                        lineY += 11;
                    } else { line = test; }
                }
                if (line) ctx.fillText(line, x + barW / 2, lineY);
            } else {
                ctx.fillText(label, x + barW / 2, H - padB + 14);
            }
        });

        // Zero axis line (if mix of pos/neg)
        if (minVal < 0) {
            ctx.strokeStyle = '#9ca3af';
            ctx.lineWidth = 1.5;
            ctx.beginPath();
            ctx.moveTo(padL, zeroY);
            ctx.lineTo(W - padR, zeroY);
            ctx.stroke();
        }

        // Tooltip on hover
        this._setupBarTooltip(values, labels, colors, padL, padR, padT, padB, chartW, chartH, gap, barW);
    }

    _setupBarTooltip(values, labels, colors, padL, padR, padT, padB, chartW, chartH, gap, barW) {
        // Remove old listener
        if (this._mousemoveHandler) this.canvas.removeEventListener('mousemove', this._mousemoveHandler);
        if (this._mouseleaveHandler) this.canvas.removeEventListener('mouseleave', this._mouseleaveHandler);

        this._mousemoveHandler = (e) => {
            const rect = this.canvas.getBoundingClientRect();
            const mx = e.clientX - rect.left;
            this.render(); // Redraw clean

            values.forEach((val, i) => {
                const barCenterX = padL + gap * i + gap / 2;
                if (Math.abs(mx - barCenterX) < gap / 2) {
                    // Highlight
                    const x = padL + gap * i + gap / 2 - barW / 2;
                    const minVal = Math.min(0, ...values);
                    const maxVal = Math.max(...values.map(Math.abs));
                    const range = maxVal - minVal;
                    const zeroY = padT + chartH * (1 - (0 - minVal) / (range || 1));
                    const barH = Math.abs(val) / (range || 1) * chartH;
                    const y = val >= 0 ? zeroY - barH : zeroY;

                    this.ctx.fillStyle = 'rgba(0,0,0,0.08)';
                    this.ctx.fillRect(x - 4, padT, barW + 8, chartH);

                    // Tooltip box
                    const tip = (labels[i] || '') + ': ' + this._fmtNum(val);
                    this.ctx.font = 'bold 11px -apple-system, sans-serif';
                    const tw = this.ctx.measureText(tip).width + 16;
                    const tx = Math.min(Math.max(barCenterX - tw / 2, 4), this.canvas.width - tw - 4);
                    const ty = Math.max(y - 36, 4);
                    this.ctx.fillStyle = 'rgba(17, 24, 39, 0.9)';
                    this.ctx.beginPath();
                    this.ctx.roundRect(tx, ty, tw, 24, 4);
                    this.ctx.fill();
                    this.ctx.fillStyle = '#fff';
                    this.ctx.textAlign = 'left';
                    this.ctx.fillText(tip, tx + 8, ty + 16);
                }
            });
        };
        this._mouseleaveHandler = () => this.render();
        this.canvas.addEventListener('mousemove', this._mousemoveHandler);
        this.canvas.addEventListener('mouseleave', this._mouseleaveHandler);
    }

    _renderDoughnut() {
        const ctx = this.ctx;
        const W = this.canvas.width;
        const H = this.canvas.height;
        ctx.clearRect(0, 0, W, H);

        const data = this.config.data;
        const labels = data.labels || [];
        const dataset = data.datasets[0];
        const values = dataset.data || [];
        const colors = dataset.backgroundColor || ['#3b82f6'];
        const total = values.reduce((a, b) => a + b, 0);
        const cutout = this.config.options?.cutout === '55%' ? 0.55 : 0;

        const legendH = Math.ceil(labels.length / 2) * 20 + 8;
        const drawH = H - legendH;
        const cx = W / 2;
        const cy = drawH / 2;
        const outerR = Math.min(W / 2, drawH / 2) - 12;
        const innerR = outerR * cutout;

        if (total === 0) {
            ctx.fillStyle = '#e5e7eb';
            ctx.beginPath();
            ctx.arc(cx, cy, outerR, 0, Math.PI * 2);
            ctx.fill();
            if (innerR > 0) {
                ctx.fillStyle = '#fff';
                ctx.beginPath();
                ctx.arc(cx, cy, innerR, 0, Math.PI * 2);
                ctx.fill();
            }
            return;
        }

        let startAngle = -Math.PI / 2;
        values.forEach((val, i) => {
            const slice = (val / total) * Math.PI * 2;
            ctx.fillStyle = colors[i % colors.length];
            ctx.beginPath();
            ctx.moveTo(cx, cy);
            ctx.arc(cx, cy, outerR, startAngle, startAngle + slice);
            if (innerR > 0) {
                ctx.arc(cx, cy, innerR, startAngle + slice, startAngle, true);
            }
            ctx.closePath();
            ctx.fill();
            ctx.strokeStyle = '#fff';
            ctx.lineWidth = 2;
            ctx.stroke();
            startAngle += slice;
        });

        // Center text (for doughnut)
        if (innerR > 0) {
            ctx.fillStyle = '#1f2937';
            ctx.font = 'bold 16px -apple-system, sans-serif';
            ctx.textAlign = 'center';
            ctx.fillText(total, cx, cy + 4);
            ctx.fillStyle = '#9ca3af';
            ctx.font = '10px -apple-system, sans-serif';
            ctx.fillText('total', cx, cy + 17);
        }

        // Legend
        const legY0 = drawH + 8;
        const colW = W / 2;
        ctx.font = '11px -apple-system, sans-serif';
        labels.forEach((label, i) => {
            const col = i % 2;
            const row = Math.floor(i / 2);
            const lx = col * colW + 12;
            const ly = legY0 + row * 20;
            ctx.fillStyle = colors[i % colors.length];
            ctx.fillRect(lx, ly, 10, 10);
            ctx.fillStyle = '#6b7280';
            ctx.textAlign = 'left';
            const pct = total > 0 ? Math.round(values[i] / total * 100) : 0;
            ctx.fillText(label + ' (' + pct + '%)', lx + 14, ly + 9);
        });
    }

    _fmtNum(n) {
        const abs = Math.abs(n);
        const sign = n < 0 ? '-' : '';
        if (abs >= 1000000) return sign + '$' + (abs / 1000000).toFixed(1) + 'M';
        if (abs >= 1000) return sign + '$' + (abs / 1000).toFixed(0) + 'k';
        if (Number.isInteger(n) && this.config.type === 'bar' && this.config.data?.datasets?.[0]?.label === 'ROI %')
            return n + '%';
        return sign + '$' + Math.round(abs).toLocaleString('es-AR');
    }
}

// Polyfill roundRect for older browsers
if (!CanvasRenderingContext2D.prototype.roundRect) {
    CanvasRenderingContext2D.prototype.roundRect = function(x, y, w, h, r) {
        this.beginPath();
        this.moveTo(x + r, y);
        this.lineTo(x + w - r, y);
        this.quadraticCurveTo(x + w, y, x + w, y + r);
        this.lineTo(x + w, y + h - r);
        this.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
        this.lineTo(x + r, y + h);
        this.quadraticCurveTo(x, y + h, x, y + h - r);
        this.lineTo(x, y + r);
        this.quadraticCurveTo(x, y, x + r, y);
        this.closePath();
    };
}
