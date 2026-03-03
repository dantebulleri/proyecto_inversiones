/* ============================================================
   CHARTS - Graficos del Dashboard con Chart.js
   ============================================================ */

const Charts = (() => {
    let instances = {};

    function destroy(key) {
        if (instances[key]) {
            instances[key].destroy();
            delete instances[key];
        }
    }

    function destroyAll() {
        Object.keys(instances).forEach(destroy);
    }

    function fmt(n) {
        return '$' + Math.round(n).toLocaleString('es-AR');
    }

    function renderInversionSocios(kpis) {
        destroy('inversionSocios');
        const ctx = document.getElementById('chart-inversion-socios');
        if (!ctx) return;

        instances.inversionSocios = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: [kpis.config.socioA, kpis.config.socioB],
                datasets: [{
                    label: 'Inversion en Gastos',
                    data: [kpis.invSocioA, kpis.invSocioB],
                    backgroundColor: ['#3b82f6', '#f97316'],
                    borderRadius: 6,
                    borderSkipped: false
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: ctx => fmt(ctx.parsed.y)
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            callback: v => fmt(v),
                            font: { size: 11 }
                        },
                        grid: { color: '#f3f4f6' }
                    },
                    x: {
                        grid: { display: false },
                        ticks: { font: { size: 12, weight: 'bold' } }
                    }
                }
            }
        });
    }

    function renderROI(kpis) {
        destroy('roi');
        const ctx = document.getElementById('chart-roi');
        if (!ctx) return;

        const motos = DataStore.getMotos().filter(m => m.precioVenta);
        const labels = [];
        const datos = [];
        const colors = [];

        motos.forEach(m => {
            const calc = DataStore.calcularGanancia(m.id);
            if (calc) {
                labels.push(m.marca + ' ' + m.modelo);
                datos.push(Math.round(calc.roi * 100));
                colors.push(calc.roi >= 0 ? '#10b981' : '#ef4444');
            }
        });

        if (labels.length === 0) {
            labels.push('Sin ventas');
            datos.push(0);
            colors.push('#d1d5db');
        }

        instances.roi = new Chart(ctx, {
            type: 'bar',
            data: {
                labels,
                datasets: [{
                    label: 'ROI %',
                    data: datos,
                    backgroundColor: colors,
                    borderRadius: 6,
                    borderSkipped: false
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: { label: ctx => ctx.parsed.y + '%' }
                    }
                },
                scales: {
                    y: {
                        ticks: {
                            callback: v => v + '%',
                            font: { size: 11 }
                        },
                        grid: { color: '#f3f4f6' }
                    },
                    x: {
                        grid: { display: false },
                        ticks: { font: { size: 10 } }
                    }
                }
            }
        });
    }

    function renderEstados(kpis) {
        destroy('estados');
        const ctx = document.getElementById('chart-estados');
        if (!ctx) return;

        const labels = Object.keys(kpis.estados);
        const datos = Object.values(kpis.estados);
        const total = datos.reduce((a, b) => a + b, 0);

        if (total === 0) {
            // Empty state
            instances.estados = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: ['Sin datos'],
                    datasets: [{ data: [1], backgroundColor: ['#e5e7eb'] }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: true,
                    plugins: { legend: { display: false } }
                }
            });
            return;
        }

        instances.estados = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels,
                datasets: [{
                    data: datos,
                    backgroundColor: [
                        '#fca5a5', // Desarmado
                        '#fcd34d', // Esperando
                        '#93c5fd', // Taller
                        '#6ee7b7', // Lista
                        '#d1d5db'  // Vendida
                    ],
                    borderWidth: 2,
                    borderColor: '#fff'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                cutout: '55%',
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: { font: { size: 11 }, padding: 12, usePointStyle: true }
                    }
                }
            }
        });
    }

    function renderFlujoCaja(kpis) {
        destroy('flujoCaja');
        const ctx = document.getElementById('chart-flujo');
        if (!ctx) return;

        instances.flujoCaja = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['Ventas', 'Compras', 'Gastos', 'Flujo Neto'],
                datasets: [{
                    data: [kpis.totalVentas, -kpis.totalCompras, -kpis.totalGastos, kpis.flujoCaja],
                    backgroundColor: [
                        '#10b981',
                        '#ef4444',
                        '#f59e0b',
                        kpis.flujoCaja >= 0 ? '#3b82f6' : '#ef4444'
                    ],
                    borderRadius: 6,
                    borderSkipped: false
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: { label: ctx => fmt(ctx.parsed.y) }
                    }
                },
                scales: {
                    y: {
                        ticks: {
                            callback: v => fmt(v),
                            font: { size: 11 }
                        },
                        grid: { color: '#f3f4f6' }
                    },
                    x: {
                        grid: { display: false },
                        ticks: { font: { size: 11, weight: 'bold' } }
                    }
                }
            }
        });
    }

    return {
        renderAll(kpis) {
            renderInversionSocios(kpis);
            renderROI(kpis);
            renderEstados(kpis);
            renderFlujoCaja(kpis);
        },
        destroyAll
    };
})();
