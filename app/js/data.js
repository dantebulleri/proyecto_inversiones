/* ============================================================
   DATA LAYER - Persistencia con Firebase Firestore
   ============================================================ */

const DataStore = (() => {

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

    let data = { ...defaultData, config: { ...defaultConfig } };
    let negocioId = null;
    let unsubscribe = null;

    function save() {
        if (negocioId && typeof db !== 'undefined') {
            db.collection('negocios').doc(negocioId).set({
                config: data.config,
                motos: data.motos,
                gastos: data.gastos,
                nextMotoId: data.nextMotoId,
                nextGastoId: data.nextGastoId
            }, { merge: true }).catch(err => console.error('Error guardando:', err));
        }
    }

    return {
        // --- Connection to Firestore ---
        connect(id) {
            negocioId = id;
            if (unsubscribe) unsubscribe();

            let isFirst = true;
            unsubscribe = db.collection('negocios').doc(id).onSnapshot(doc => {
                if (doc.exists) {
                    const remote = doc.data();
                    data = {
                        ...defaultData,
                        config: { ...defaultConfig },
                        motos: remote.motos || [],
                        gastos: remote.gastos || [],
                        nextMotoId: remote.nextMotoId || 1,
                        nextGastoId: remote.nextGastoId || 1
                    };
                    if (remote.config) {
                        data.config = { ...defaultConfig, ...remote.config };
                    }

                    if (isFirst) {
                        isFirst = false;
                        if (window.App && window.App.init) window.App.init();
                    } else {
                        if (window.App && window.App.refreshCurrent) window.App.refreshCurrent();
                    }
                }
            }, err => {
                console.error('Error en listener Firestore:', err);
            });
        },

        disconnect() {
            if (unsubscribe) { unsubscribe(); unsubscribe = null; }
            negocioId = null;
            data = { ...defaultData, config: { ...defaultConfig } };
        },

        getDefaults() {
            return {
                ...defaultData,
                config: { ...defaultConfig }
            };
        },

        // --- Config ---
        getConfig() { return { ...data.config }; },
        setConfig(config) {
            data.config = { ...config };
            save();
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
                pctAporteA: moto.pctAporteA !== undefined ? Number(moto.pctAporteA) : 50,
                fechaCompra: moto.fechaCompra,
                precioVenta: null,
                fechaVenta: null,
                estado: 'Reparacion',
                observaciones: moto.observaciones || '',
                historialEstados: [{ estado: 'Reparacion', fecha: moto.fechaCompra }],
                createdAt: new Date().toISOString()
            };
            data.motos.push(newMoto);
            save();
            return newMoto;
        },

        updateMoto(id, updates) {
            const idx = data.motos.findIndex(m => m.id === id);
            if (idx === -1) return null;
            data.motos[idx] = { ...data.motos[idx], ...updates };
            save();
            return data.motos[idx];
        },

        deleteMoto(id) {
            data.motos = data.motos.filter(m => m.id !== id);
            data.gastos = data.gastos.filter(g => g.motoId !== id);
            save();
        },

        cambiarEstado(id, nuevoEstado) {
            const moto = data.motos.find(m => m.id === id);
            if (!moto) return null;
            moto.estado = nuevoEstado;
            moto.historialEstados.push({ estado: nuevoEstado, fecha: new Date().toISOString().split('T')[0] });
            save();
            return moto;
        },

        registrarVenta(id, precioVenta, fechaVenta) {
            const moto = data.motos.find(m => m.id === id);
            if (!moto) return null;
            moto.precioVenta = Number(precioVenta);
            moto.fechaVenta = fechaVenta;
            moto.estado = 'Vendida';
            moto.historialEstados.push({ estado: 'Vendida', fecha: fechaVenta });
            save();
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
            save();
            return newGasto;
        },

        deleteGasto(id) {
            data.gastos = data.gastos.filter(g => g.id !== id);
            save();
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

            const pctCompraA = moto.pctAporteA !== undefined ? moto.pctAporteA : 50;
            const aporteCompraA = moto.precioCompra * pctCompraA / 100;
            const aporteCompraB = moto.precioCompra - aporteCompraA;
            const devolucionA = aporteCompraA + gastosA + (utilidadNeta * pctA);
            const devolucionB = aporteCompraB + gastosB + (utilidadNeta * pctB);
            const roi = inversionTotal > 0 ? utilidadNeta / inversionTotal : 0;

            return {
                motoId,
                motoDesc: moto.marca + ' ' + moto.modelo,
                precioCompra: moto.precioCompra,
                aporteCompraA,
                aporteCompraB,
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

            let diasStock = motos.map(m => this.getDiasStock(m.id));
            const diasPromedio = diasStock.length > 0 ? diasStock.reduce((a, b) => a + b, 0) / diasStock.length : 0;

            const flujoCaja = totalVentas - totalCompras - totalGastos;

            // Inversion por socio: aporte a compras + gastos pagados
            let invSocioA = data.gastos
                .filter(g => g.pagadoPor === config.socioA)
                .reduce((sum, g) => sum + g.monto, 0);
            let invSocioB = data.gastos
                .filter(g => g.pagadoPor === config.socioB)
                .reduce((sum, g) => sum + g.monto, 0);

            motos.forEach(m => {
                const pctCA = m.pctAporteA !== undefined ? m.pctAporteA : 50;
                invSocioA += m.precioCompra * pctCA / 100;
                invSocioB += m.precioCompra * (100 - pctCA) / 100;
            });

            // Estados: mapear legacy a los 3 nuevos
            const estados = { 'Reparacion': 0, 'Lista para vender': 0, 'Vendida': 0 };
            const stateMap = {
                'Desarmado': 'Reparacion', 'Esperando repuestos': 'Reparacion',
                'En taller': 'Reparacion', 'Reparacion': 'Reparacion',
                'Lista para venta': 'Lista para vender', 'Lista para vender': 'Lista para vender',
                'Vendida': 'Vendida'
            };
            motos.forEach(m => {
                const mapped = stateMap[m.estado] || 'Reparacion';
                estados[mapped]++;
            });

            return {
                totalMotos, activas, vendidas: vendidas.length,
                inversionTotal, gananciaNeta, roiPromedio, diasPromedio, flujoCaja,
                invSocioA, invSocioB, totalCompras, totalVentas, totalGastos,
                estados, config
            };
        },

        // --- Export / Import ---
        exportJSON() {
            return JSON.stringify({
                config: data.config,
                motos: data.motos,
                gastos: data.gastos,
                nextMotoId: data.nextMotoId,
                nextGastoId: data.nextGastoId
            }, null, 2);
        },

        importJSON(jsonString) {
            const imported = JSON.parse(jsonString);
            data = {
                ...defaultData,
                ...imported,
                config: { ...defaultConfig, ...(imported.config || {}) }
            };
            save();
            return true;
        },

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
