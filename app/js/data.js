/* ============================================================
   DATA LAYER - Persistencia con localStorage
   ============================================================ */

const DataStore = (() => {
    const STORAGE_KEY = 'motoflip_data';

    const defaultConfig = {
        socioA: 'Socio A',
        socioB: 'Socio B',
        pctA: 50,
        pctB: 50,
        categorias: ['Repuestos', 'Mano de obra', 'Pintura', 'Gestoria', 'Transporte', 'Seguro', 'Otros']
    };

    const defaultData = {
        config: { ...defaultConfig },
        motos: [],
        gastos: [],
        nextMotoId: 1,
        nextGastoId: 1
    };

    function load() {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return { ...defaultData, config: { ...defaultConfig } };
        try {
            const parsed = JSON.parse(raw);
            // Merge with defaults for forward-compatibility
            return {
                ...defaultData,
                ...parsed,
                config: { ...defaultConfig, ...parsed.config }
            };
        } catch {
            return { ...defaultData, config: { ...defaultConfig } };
        }
    }

    function save(data) {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    }

    let data = load();

    return {
        // --- Config ---
        getConfig() { return { ...data.config }; },
        setConfig(config) {
            data.config = { ...config };
            save(data);
        },

        // --- Motos ---
        getMotos() { return [...data.motos]; },

        getMoto(id) { return data.motos.find(m => m.id === id) || null; },

        addMoto(moto) {
            const id = 'MOTO-' + String(data.nextMotoId).padStart(3, '0');
            data.nextMotoId++;
            const newMoto = {
                id,
                marca: moto.marca,
                modelo: moto.modelo,
                anio: moto.anio,
                cilindrada: moto.cilindrada || '',
                precioCompra: Number(moto.precioCompra),
                fechaCompra: moto.fechaCompra,
                precioVenta: null,
                fechaVenta: null,
                estado: 'Desarmado',
                observaciones: moto.observaciones || '',
                historialEstados: [{ estado: 'Desarmado', fecha: moto.fechaCompra }],
                createdAt: new Date().toISOString()
            };
            data.motos.push(newMoto);
            save(data);
            return newMoto;
        },

        updateMoto(id, updates) {
            const idx = data.motos.findIndex(m => m.id === id);
            if (idx === -1) return null;
            data.motos[idx] = { ...data.motos[idx], ...updates };
            save(data);
            return data.motos[idx];
        },

        deleteMoto(id) {
            data.motos = data.motos.filter(m => m.id !== id);
            // Also delete associated expenses
            data.gastos = data.gastos.filter(g => g.motoId !== id);
            save(data);
        },

        cambiarEstado(id, nuevoEstado) {
            const moto = data.motos.find(m => m.id === id);
            if (!moto) return null;
            moto.estado = nuevoEstado;
            moto.historialEstados.push({ estado: nuevoEstado, fecha: new Date().toISOString().split('T')[0] });
            save(data);
            return moto;
        },

        registrarVenta(id, precioVenta, fechaVenta) {
            const moto = data.motos.find(m => m.id === id);
            if (!moto) return null;
            moto.precioVenta = Number(precioVenta);
            moto.fechaVenta = fechaVenta;
            moto.estado = 'Vendida';
            moto.historialEstados.push({ estado: 'Vendida', fecha: fechaVenta });
            save(data);
            return moto;
        },

        // --- Gastos ---
        getGastos() { return [...data.gastos]; },

        getGastosByMoto(motoId) { return data.gastos.filter(g => g.motoId === motoId); },

        addGasto(gasto) {
            const id = 'GASTO-' + String(data.nextGastoId).padStart(3, '0');
            data.nextGastoId++;
            const newGasto = {
                id,
                motoId: gasto.motoId,
                fecha: gasto.fecha,
                categoria: gasto.categoria,
                descripcion: gasto.descripcion || '',
                monto: Number(gasto.monto),
                pagadoPor: gasto.pagadoPor,
                comprobante: gasto.comprobante || '',
                createdAt: new Date().toISOString()
            };
            data.gastos.push(newGasto);
            save(data);
            return newGasto;
        },

        deleteGasto(id) {
            data.gastos = data.gastos.filter(g => g.id !== id);
            save(data);
        },

        // --- Calculos ---
        getGastosSocio(motoId, socio) {
            return data.gastos
                .filter(g => g.motoId === motoId && g.pagadoPor === socio)
                .reduce((sum, g) => sum + g.monto, 0);
        },

        getTotalGastosMoto(motoId) {
            return data.gastos
                .filter(g => g.motoId === motoId)
                .reduce((sum, g) => sum + g.monto, 0);
        },

        getInversionTotal(motoId) {
            const moto = this.getMoto(motoId);
            if (!moto) return 0;
            return moto.precioCompra + this.getTotalGastosMoto(motoId);
        },

        calcularGanancia(motoId) {
            const moto = this.getMoto(motoId);
            if (!moto || !moto.precioVenta) return null;

            const config = this.getConfig();
            const gastosA = this.getGastosSocio(motoId, config.socioA);
            const gastosB = this.getGastosSocio(motoId, config.socioB);
            const totalGastos = gastosA + gastosB;
            const inversionTotal = moto.precioCompra + totalGastos;
            const utilidadNeta = moto.precioVenta - inversionTotal;
            const pctA = config.pctA / 100;
            const pctB = config.pctB / 100;

            // Devolucion: mitad de compra + gastos propios + % de ganancia
            const devolucionA = (moto.precioCompra / 2) + gastosA + (utilidadNeta * pctA);
            const devolucionB = (moto.precioCompra / 2) + gastosB + (utilidadNeta * pctB);
            const roi = inversionTotal > 0 ? utilidadNeta / inversionTotal : 0;

            return {
                motoId,
                motoDesc: moto.marca + ' ' + moto.modelo,
                precioCompra: moto.precioCompra,
                gastosA,
                gastosB,
                totalGastos,
                inversionTotal,
                precioVenta: moto.precioVenta,
                utilidadNeta,
                devolucionA,
                devolucionB,
                roi
            };
        },

        getDiasStock(motoId) {
            const moto = this.getMoto(motoId);
            if (!moto) return 0;
            const inicio = new Date(moto.fechaCompra);
            const fin = moto.fechaVenta ? new Date(moto.fechaVenta) : new Date();
            return Math.floor((fin - inicio) / (1000 * 60 * 60 * 24));
        },

        // --- KPIs ---
        getKPIs() {
            const motos = this.getMotos();
            const config = this.getConfig();
            const totalMotos = motos.length;
            const vendidas = motos.filter(m => m.estado === 'Vendida');
            const activas = totalMotos - vendidas.length;

            let inversionTotal = 0;
            let totalCompras = 0;
            let totalVentas = 0;
            let totalGastos = 0;

            motos.forEach(m => {
                totalCompras += m.precioCompra;
                if (m.precioVenta) totalVentas += m.precioVenta;
            });
            totalGastos = data.gastos.reduce((sum, g) => sum + g.monto, 0);
            inversionTotal = totalCompras + totalGastos;

            // ROI promedio de vendidas
            let rois = [];
            let gananciaNeta = 0;
            vendidas.forEach(m => {
                const calc = this.calcularGanancia(m.id);
                if (calc) {
                    rois.push(calc.roi);
                    gananciaNeta += calc.utilidadNeta;
                }
            });
            const roiPromedio = rois.length > 0 ? rois.reduce((a, b) => a + b, 0) / rois.length : 0;

            // Dias promedio de stock
            let diasStock = motos.map(m => this.getDiasStock(m.id));
            const diasPromedio = diasStock.length > 0 ? diasStock.reduce((a, b) => a + b, 0) / diasStock.length : 0;

            // Flujo de caja
            const flujoCaja = totalVentas - totalCompras - totalGastos;

            // Inversion por socio
            const invSocioA = data.gastos
                .filter(g => g.pagadoPor === config.socioA)
                .reduce((sum, g) => sum + g.monto, 0);
            const invSocioB = data.gastos
                .filter(g => g.pagadoPor === config.socioB)
                .reduce((sum, g) => sum + g.monto, 0);

            // Conteo por estado
            const estados = {
                'Desarmado': 0,
                'Esperando repuestos': 0,
                'En taller': 0,
                'Lista para venta': 0,
                'Vendida': 0
            };
            motos.forEach(m => { if (estados.hasOwnProperty(m.estado)) estados[m.estado]++; });

            return {
                totalMotos, activas, vendidas: vendidas.length,
                inversionTotal, gananciaNeta, roiPromedio, diasPromedio, flujoCaja,
                invSocioA, invSocioB, totalCompras, totalVentas, totalGastos,
                estados, config
            };
        },

        // --- Export / Import ---
        exportJSON() {
            return JSON.stringify(data, null, 2);
        },

        importJSON(jsonString) {
            const imported = JSON.parse(jsonString);
            data = { ...defaultData, ...imported, config: { ...defaultConfig, ...imported.config } };
            save(data);
            return true;
        },

        // --- Mottos activos (no vendidas) para selects ---
        getMotosActivas() {
            return data.motos.filter(m => m.estado !== 'Vendida');
        },

        getMotosParaSelect() {
            return data.motos.map(m => ({
                id: m.id,
                label: m.id + ' - ' + m.marca + ' ' + m.modelo
            }));
        }
    };
})();
