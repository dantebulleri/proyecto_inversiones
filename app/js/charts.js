/* ============================================================
   CHARTS - Graficos del Dashboard con miniChart (sin dependencias)
   ============================================================ */

const Charts = (() => {
    let instances = {};

    function destroy(key) {
        if (instances[key]) {
            instances[key].destroy();
            delete instances[key];
        }
    }

    function renderInversionSocios(kpis) {
        destroy('inv');
        const canvas = document.getElementById('chart-inversion-socios');
        if (!canvas) return;
        instances.inv = new Chart(canvas, {
            type: 'bar',
            data: {
                labels: [kpis.config.socioA, kpis.config.socioB],
                datasets: [{
                    data: [kpis.invSocioA, kpis.invSocioB],
                    backgroundColor: ['#3b82f6', '#f97316']
                }]
            }
        });
    }

    function renderROI(kpis) {
        destroy('roi');
        const canvas = document.getElementById('chart-roi');
        if (!canvas) return;

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

        instances.roi = new Chart(canvas, {
            type: 'bar',
            data: {
                labels,
                datasets: [{
                    label: 'ROI %',
                    data: datos,
                    backgroundColor: colors
                }]
            }
        });
    }

    function renderEstados(kpis) {
        destroy('estados');
        const canvas = document.getElementById('chart-estados');
        if (!canvas) return;

        const labels = Object.keys(kpis.estados);
        const datos = Object.values(kpis.estados);

        instances.estados = new Chart(canvas, {
            type: 'doughnut',
            data: {
                labels,
                datasets: [{
                    data: datos,
                    backgroundColor: ['#93c5fd', '#6ee7b7', '#d1d5db']
                }]
            },
            options: { cutout: '55%' }
        });
    }

    function renderFlujoCaja(kpis) {
        destroy('flujo');
        const canvas = document.getElementById('chart-flujo');
        if (!canvas) return;

        instances.flujo = new Chart(canvas, {
            type: 'bar',
            data: {
                labels: ['Ventas', 'Compras', 'Gastos', 'Flujo Neto'],
                datasets: [{
                    data: [kpis.totalVentas, -kpis.totalCompras, -kpis.totalGastos, kpis.flujoCaja],
                    backgroundColor: [
                        '#10b981', '#ef4444', '#f59e0b',
                        kpis.flujoCaja >= 0 ? '#3b82f6' : '#ef4444'
                    ]
                }]
            }
        });
    }

    function renderAportePorMoto(kpis) {
        destroy('aporteMoto');
        const canvas = document.getElementById('chart-aporte-moto');
        if (!canvas) return;

        const motos = DataStore.getMotos();
        const config = kpis.config;
        const labels = [];
        const datosA = [];
        const datosB = [];

        motos.forEach(m => {
            labels.push(m.marca + ' ' + m.modelo);
            const pctCA = m.pctAporteA !== undefined ? m.pctAporteA : 50;
            const aporteCompraA = m.precioCompra * pctCA / 100;
            const aporteCompraB = m.precioCompra - aporteCompraA;
            const gastosA = DataStore.getGastosSocio(m.id, config.socioA);
            const gastosB = DataStore.getGastosSocio(m.id, config.socioB);
            datosA.push(aporteCompraA + gastosA);
            datosB.push(aporteCompraB + gastosB);
        });

        if (labels.length === 0) {
            labels.push('Sin motos');
            datosA.push(0);
            datosB.push(0);
        }

        instances.aporteMoto = new Chart(canvas, {
            type: 'stackedBar',
            data: {
                labels,
                datasets: [
                    { label: config.socioA, data: datosA, backgroundColor: '#3b82f6' },
                    { label: config.socioB, data: datosB, backgroundColor: '#f97316' }
                ]
            }
        });
    }

    return {
        renderAll(kpis) {
            renderInversionSocios(kpis);
            renderROI(kpis);
            renderEstados(kpis);
            renderFlujoCaja(kpis);
            renderAportePorMoto(kpis);
        }
    };
})();
