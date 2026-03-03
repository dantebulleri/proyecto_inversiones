/* ============================================================
   APP.JS - Logica principal de la aplicacion MotoFlip
   ============================================================ */

document.addEventListener('DOMContentLoaded', () => {

    // ============================================================
    // NAVIGATION
    // ============================================================
    const navLinks = document.querySelectorAll('.nav-link');
    const sections = document.querySelectorAll('.section');
    let currentSection = 'dashboard';

    function navigateTo(sectionId) {
        currentSection = sectionId;
        sections.forEach(s => s.classList.remove('active'));
        navLinks.forEach(l => l.classList.remove('active'));

        const section = document.getElementById('section-' + sectionId);
        const link = document.querySelector('[data-section="' + sectionId + '"]');
        if (section) section.classList.add('active');
        if (link) link.classList.add('active');

        // Refresh section data
        switch (sectionId) {
            case 'dashboard': refreshDashboard(); break;
            case 'motos': refreshMotos(); break;
            case 'gastos': refreshGastos(); break;
            case 'kanban': refreshKanban(); break;
            case 'ganancias': refreshGanancias(); break;
            case 'config': refreshConfig(); break;
        }
    }

    navLinks.forEach(link => {
        link.addEventListener('click', e => {
            e.preventDefault();
            navigateTo(link.dataset.section);
        });
    });

    // ============================================================
    // UTILITIES
    // ============================================================
    function fmt(n) {
        return '$' + Math.round(n).toLocaleString('es-AR');
    }

    function fmtDate(dateStr) {
        if (!dateStr) return '-';
        const d = new Date(dateStr + 'T12:00:00');
        return d.toLocaleDateString('es-AR');
    }

    function todayStr() {
        return new Date().toISOString().split('T')[0];
    }

    function toast(message, type = 'success') {
        const container = document.getElementById('toast-container');
        const t = document.createElement('div');
        t.className = 'toast toast-' + type;
        t.textContent = message;
        container.appendChild(t);
        setTimeout(() => t.remove(), 3000);
    }

    function estadoBadge(estado) {
        const map = {
            'Desarmado': 'badge-desarmado',
            'Esperando repuestos': 'badge-esperando',
            'En taller': 'badge-taller',
            'Lista para venta': 'badge-lista',
            'Vendida': 'badge-vendida'
        };
        return '<span class="badge ' + (map[estado] || '') + '">' + estado + '</span>';
    }

    // ============================================================
    // MODALS
    // ============================================================
    function openModal(id) {
        document.getElementById(id).classList.add('show');
    }

    function closeModal(id) {
        document.getElementById(id).classList.remove('show');
    }

    // Close buttons
    document.querySelectorAll('[data-close]').forEach(btn => {
        btn.addEventListener('click', () => closeModal(btn.dataset.close));
    });

    // Close on overlay click
    document.querySelectorAll('.modal-overlay').forEach(overlay => {
        overlay.addEventListener('click', e => {
            if (e.target === overlay) overlay.classList.remove('show');
        });
    });

    // Close on Escape
    document.addEventListener('keydown', e => {
        if (e.key === 'Escape') {
            document.querySelectorAll('.modal-overlay.show').forEach(m => m.classList.remove('show'));
        }
    });

    // ============================================================
    // DASHBOARD
    // ============================================================
    function refreshDashboard() {
        const kpis = DataStore.getKPIs();

        document.getElementById('kpi-total-motos').textContent = kpis.totalMotos;
        document.getElementById('kpi-activas').textContent = kpis.activas;
        document.getElementById('kpi-vendidas').textContent = kpis.vendidas;
        document.getElementById('kpi-inversion').textContent = fmt(kpis.inversionTotal);
        document.getElementById('kpi-ganancia').textContent = fmt(kpis.gananciaNeta);
        document.getElementById('kpi-roi').textContent = Math.round(kpis.roiPromedio * 100) + '%';
        document.getElementById('kpi-dias-stock').textContent = Math.round(kpis.diasPromedio);
        document.getElementById('kpi-flujo').textContent = fmt(kpis.flujoCaja);

        // Flujo de caja color
        const flujoCard = document.getElementById('kpi-flujo-card');
        flujoCard.classList.remove('kpi-success', 'kpi-warning');
        flujoCard.classList.add(kpis.flujoCaja >= 0 ? 'kpi-success' : 'kpi-warning');

        Charts.renderAll(kpis);
    }

    // ============================================================
    // MOTOS
    // ============================================================
    function refreshMotos() {
        const motos = DataStore.getMotos();
        const tbody = document.getElementById('tbody-motos');
        const empty = document.getElementById('empty-motos');

        if (motos.length === 0) {
            tbody.innerHTML = '';
            empty.style.display = 'block';
            return;
        }

        empty.style.display = 'none';
        tbody.innerHTML = motos.map(m => {
            const gastoTotal = DataStore.getTotalGastosMoto(m.id);
            return '<tr>' +
                '<td><strong>' + m.id + '</strong></td>' +
                '<td>' + m.marca + '</td>' +
                '<td>' + m.modelo + '</td>' +
                '<td>' + m.anio + '</td>' +
                '<td>' + (m.cilindrada || '-') + '</td>' +
                '<td>' + fmt(m.precioCompra) + '</td>' +
                '<td>' + fmtDate(m.fechaCompra) + '</td>' +
                '<td>' + estadoBadge(m.estado) + '</td>' +
                '<td><div class="action-buttons">' +
                    (m.estado !== 'Vendida' ?
                        '<button class="btn-icon success" onclick="App.abrirVenta(\'' + m.id + '\')" title="Registrar venta">$</button>' +
                        '<button class="btn-icon" onclick="App.abrirEstado(\'' + m.id + '\')" title="Cambiar estado">&#9654;</button>' +
                        '<button class="btn-icon" onclick="App.abrirGastoParaMoto(\'' + m.id + '\')" title="Agregar gasto">+</button>'
                    : '') +
                    '<button class="btn-icon danger" onclick="App.eliminarMoto(\'' + m.id + '\')" title="Eliminar">&#10005;</button>' +
                '</div></td>' +
            '</tr>';
        }).join('');
    }

    // Nueva moto
    document.getElementById('btn-nueva-moto').addEventListener('click', () => {
        document.getElementById('moto-edit-id').value = '';
        document.getElementById('modal-moto-title').textContent = 'Registrar Nueva Moto';
        document.getElementById('moto-marca').value = '';
        document.getElementById('moto-modelo').value = '';
        document.getElementById('moto-anio').value = '';
        document.getElementById('moto-cilindrada').value = '';
        document.getElementById('moto-precio').value = '';
        document.getElementById('moto-fecha').value = todayStr();
        document.getElementById('moto-obs').value = '';
        document.getElementById('btn-guardar-moto').textContent = 'Registrar';
        openModal('modal-moto');
    });

    document.getElementById('btn-guardar-moto').addEventListener('click', () => {
        const marca = document.getElementById('moto-marca').value.trim();
        const modelo = document.getElementById('moto-modelo').value.trim();
        const anio = document.getElementById('moto-anio').value.trim();
        const cilindrada = document.getElementById('moto-cilindrada').value.trim();
        const precio = document.getElementById('moto-precio').value;
        const fecha = document.getElementById('moto-fecha').value;
        const obs = document.getElementById('moto-obs').value.trim();

        if (!marca || !modelo || !anio || !precio || !fecha) {
            toast('Completa todos los campos obligatorios (*)', 'error');
            return;
        }

        const moto = DataStore.addMoto({
            marca, modelo, anio, cilindrada,
            precioCompra: precio, fechaCompra: fecha,
            observaciones: obs
        });

        toast('Moto ' + moto.id + ' registrada: ' + marca + ' ' + modelo);
        closeModal('modal-moto');
        refreshMotos();
    });

    // ============================================================
    // GASTOS
    // ============================================================
    function refreshGastos() {
        const config = DataStore.getConfig();
        let gastos = DataStore.getGastos();
        const motos = DataStore.getMotos();

        // Populate filters
        const filtroMoto = document.getElementById('filtro-moto-gastos');
        const filtroSocio = document.getElementById('filtro-socio-gastos');
        const filtroCat = document.getElementById('filtro-categoria-gastos');

        // Save current filter values
        const fmVal = filtroMoto.value;
        const fsVal = filtroSocio.value;
        const fcVal = filtroCat.value;

        // Rebuild filter options
        filtroMoto.innerHTML = '<option value="">Todas las motos</option>';
        motos.forEach(m => {
            filtroMoto.innerHTML += '<option value="' + m.id + '">' + m.id + ' - ' + m.marca + ' ' + m.modelo + '</option>';
        });
        filtroMoto.value = fmVal;

        filtroSocio.innerHTML = '<option value="">Todos los socios</option>';
        filtroSocio.innerHTML += '<option value="' + config.socioA + '">' + config.socioA + '</option>';
        filtroSocio.innerHTML += '<option value="' + config.socioB + '">' + config.socioB + '</option>';
        filtroSocio.value = fsVal;

        filtroCat.innerHTML = '<option value="">Todas las categorias</option>';
        config.categorias.forEach(c => {
            filtroCat.innerHTML += '<option value="' + c + '">' + c + '</option>';
        });
        filtroCat.value = fcVal;

        // Apply filters
        if (fmVal) gastos = gastos.filter(g => g.motoId === fmVal);
        if (fsVal) gastos = gastos.filter(g => g.pagadoPor === fsVal);
        if (fcVal) gastos = gastos.filter(g => g.categoria === fcVal);

        // Summary
        const totalGastos = gastos.reduce((s, g) => s + g.monto, 0);
        const gastosA = gastos.filter(g => g.pagadoPor === config.socioA).reduce((s, g) => s + g.monto, 0);
        const gastosB = gastos.filter(g => g.pagadoPor === config.socioB).reduce((s, g) => s + g.monto, 0);

        document.getElementById('gastos-resumen').innerHTML =
            '<div class="gasto-resumen-item"><div class="value">' + fmt(totalGastos) + '</div><div class="label">Total Gastos</div></div>' +
            '<div class="gasto-resumen-item"><div class="value">' + fmt(gastosA) + '</div><div class="label">' + config.socioA + '</div></div>' +
            '<div class="gasto-resumen-item"><div class="value">' + fmt(gastosB) + '</div><div class="label">' + config.socioB + '</div></div>' +
            '<div class="gasto-resumen-item"><div class="value">' + gastos.length + '</div><div class="label">Registros</div></div>';

        // Table
        const tbody = document.getElementById('tbody-gastos');
        const empty = document.getElementById('empty-gastos');

        if (gastos.length === 0) {
            tbody.innerHTML = '';
            empty.style.display = 'block';
            return;
        }

        empty.style.display = 'none';
        // Sort by date desc
        gastos.sort((a, b) => b.fecha.localeCompare(a.fecha));

        tbody.innerHTML = gastos.map(g => {
            const moto = DataStore.getMoto(g.motoId);
            const motoLabel = moto ? moto.marca + ' ' + moto.modelo : g.motoId;
            return '<tr>' +
                '<td>' + fmtDate(g.fecha) + '</td>' +
                '<td>' + motoLabel + '</td>' +
                '<td>' + g.categoria + '</td>' +
                '<td>' + (g.descripcion || '-') + '</td>' +
                '<td><strong>' + fmt(g.monto) + '</strong></td>' +
                '<td>' + g.pagadoPor + '</td>' +
                '<td><div class="action-buttons">' +
                    '<button class="btn-icon danger" onclick="App.eliminarGasto(\'' + g.id + '\')" title="Eliminar">&#10005;</button>' +
                '</div></td>' +
            '</tr>';
        }).join('');
    }

    // Filters
    ['filtro-moto-gastos', 'filtro-socio-gastos', 'filtro-categoria-gastos'].forEach(id => {
        document.getElementById(id).addEventListener('change', refreshGastos);
    });

    // Nuevo gasto
    function openGastoModal(preselectedMotoId) {
        const config = DataStore.getConfig();
        const motos = DataStore.getMotos();

        const selMoto = document.getElementById('gasto-moto');
        selMoto.innerHTML = '<option value="">Seleccionar moto...</option>';
        motos.forEach(m => {
            selMoto.innerHTML += '<option value="' + m.id + '">' + m.id + ' - ' + m.marca + ' ' + m.modelo + '</option>';
        });
        if (preselectedMotoId) selMoto.value = preselectedMotoId;

        const selCat = document.getElementById('gasto-categoria');
        selCat.innerHTML = '<option value="">Seleccionar...</option>';
        config.categorias.forEach(c => {
            selCat.innerHTML += '<option value="' + c + '">' + c + '</option>';
        });

        const selPagado = document.getElementById('gasto-pagado-por');
        selPagado.innerHTML = '<option value="">Seleccionar...</option>';
        selPagado.innerHTML += '<option value="' + config.socioA + '">' + config.socioA + '</option>';
        selPagado.innerHTML += '<option value="' + config.socioB + '">' + config.socioB + '</option>';

        document.getElementById('gasto-fecha').value = todayStr();
        document.getElementById('gasto-descripcion').value = '';
        document.getElementById('gasto-monto').value = '';
        document.getElementById('gasto-comprobante').value = '';
        document.getElementById('gasto-edit-id').value = '';

        openModal('modal-gasto');
    }

    document.getElementById('btn-nuevo-gasto').addEventListener('click', () => openGastoModal(null));

    document.getElementById('btn-guardar-gasto').addEventListener('click', () => {
        const motoId = document.getElementById('gasto-moto').value;
        const fecha = document.getElementById('gasto-fecha').value;
        const categoria = document.getElementById('gasto-categoria').value;
        const descripcion = document.getElementById('gasto-descripcion').value.trim();
        const monto = document.getElementById('gasto-monto').value;
        const pagadoPor = document.getElementById('gasto-pagado-por').value;
        const comprobante = document.getElementById('gasto-comprobante').value.trim();

        if (!motoId || !fecha || !categoria || !monto || !pagadoPor) {
            toast('Completa todos los campos obligatorios (*)', 'error');
            return;
        }

        const gasto = DataStore.addGasto({
            motoId, fecha, categoria, descripcion,
            monto, pagadoPor, comprobante
        });

        toast('Gasto de ' + fmt(Number(monto)) + ' registrado');
        closeModal('modal-gasto');
        refreshGastos();
    });

    // ============================================================
    // KANBAN
    // ============================================================
    function refreshKanban() {
        const motos = DataStore.getMotos();
        const config = DataStore.getConfig();

        const columns = {
            'Desarmado': document.getElementById('kanban-desarmado'),
            'Esperando repuestos': document.getElementById('kanban-esperando'),
            'En taller': document.getElementById('kanban-taller'),
            'Lista para venta': document.getElementById('kanban-lista'),
            'Vendida': document.getElementById('kanban-vendida')
        };

        const counts = {
            'Desarmado': 0, 'Esperando repuestos': 0,
            'En taller': 0, 'Lista para venta': 0, 'Vendida': 0
        };

        // Clear columns
        Object.values(columns).forEach(col => col.innerHTML = '');

        motos.forEach(m => {
            const col = columns[m.estado];
            if (!col) return;

            counts[m.estado]++;
            const dias = DataStore.getDiasStock(m.id);
            const inversion = DataStore.getInversionTotal(m.id);

            let actionsHtml = '';
            if (m.estado !== 'Vendida') {
                actionsHtml =
                    '<div class="kanban-card-actions">' +
                        '<button class="btn btn-sm btn-outline" onclick="App.abrirEstado(\'' + m.id + '\')">Mover</button>' +
                        '<button class="btn btn-sm btn-success" onclick="App.abrirVenta(\'' + m.id + '\')">Vender</button>' +
                    '</div>';
            }

            col.innerHTML +=
                '<div class="kanban-card">' +
                    '<div class="kanban-card-title">' + m.marca + ' ' + m.modelo + '</div>' +
                    '<div class="kanban-card-meta">' +
                        '<span>' + m.id + '</span>' +
                        '<span>' + dias + ' dias</span>' +
                    '</div>' +
                    '<div class="kanban-card-investment">Inv: ' + fmt(inversion) + '</div>' +
                    actionsHtml +
                '</div>';
        });

        // Update counts
        document.getElementById('count-desarmado').textContent = counts['Desarmado'];
        document.getElementById('count-esperando').textContent = counts['Esperando repuestos'];
        document.getElementById('count-taller').textContent = counts['En taller'];
        document.getElementById('count-lista').textContent = counts['Lista para venta'];
        document.getElementById('count-vendida').textContent = counts['Vendida'];
    }

    // Cambiar estado
    document.getElementById('btn-confirmar-estado').addEventListener('click', () => {
        const motoId = document.getElementById('estado-moto-id').value;
        const nuevoEstado = document.getElementById('estado-nuevo').value;

        if (!motoId || !nuevoEstado) return;

        DataStore.cambiarEstado(motoId, nuevoEstado);
        toast('Estado cambiado a: ' + nuevoEstado);
        closeModal('modal-estado');
        refreshKanban();
        refreshMotos();
    });

    // Registrar venta
    document.getElementById('btn-confirmar-venta').addEventListener('click', () => {
        const motoId = document.getElementById('venta-moto-id').value;
        const precio = document.getElementById('venta-precio').value;
        const fecha = document.getElementById('venta-fecha').value;

        if (!motoId || !precio || !fecha) {
            toast('Completa todos los campos', 'error');
            return;
        }

        const moto = DataStore.registrarVenta(motoId, precio, fecha);
        const calc = DataStore.calcularGanancia(motoId);

        if (calc) {
            const msg = calc.utilidadNeta >= 0
                ? 'Venta registrada. Ganancia: ' + fmt(calc.utilidadNeta)
                : 'Venta registrada. Perdida: ' + fmt(calc.utilidadNeta);
            toast(msg, calc.utilidadNeta >= 0 ? 'success' : 'error');
        } else {
            toast('Venta registrada');
        }

        closeModal('modal-venta');
        refreshMotos();
        refreshKanban();
    });

    // ============================================================
    // GANANCIAS
    // ============================================================
    function refreshGanancias() {
        const motos = DataStore.getMotos();
        const config = DataStore.getConfig();

        // Update table headers with partner names
        const thead = document.querySelector('#tabla-ganancias thead tr');
        const ths = thead.querySelectorAll('th');
        ths[2].textContent = 'Gastos ' + config.socioA;
        ths[3].textContent = 'Gastos ' + config.socioB;
        ths[7].textContent = 'Devol. ' + config.socioA;
        ths[8].textContent = 'Devol. ' + config.socioB;

        // Table data - all motos
        const tbody = document.getElementById('tbody-ganancias');
        const empty = document.getElementById('empty-ganancias');

        const rows = [];
        let totalUtilidad = 0;
        let totalInvA = 0;
        let totalInvB = 0;
        let totalDevA = 0;
        let totalDevB = 0;

        motos.forEach(m => {
            const gastosA = DataStore.getGastosSocio(m.id, config.socioA);
            const gastosB = DataStore.getGastosSocio(m.id, config.socioB);
            const invTotal = m.precioCompra + gastosA + gastosB;

            totalInvA += gastosA;
            totalInvB += gastosB;

            if (m.precioVenta) {
                const calc = DataStore.calcularGanancia(m.id);
                if (calc) {
                    totalUtilidad += calc.utilidadNeta;
                    totalDevA += calc.devolucionA;
                    totalDevB += calc.devolucionB;

                    const utilClass = calc.utilidadNeta >= 0 ? 'amount-positive' : 'amount-negative';
                    const roiStr = Math.round(calc.roi * 100) + '%';

                    rows.push('<tr>' +
                        '<td><strong>' + m.marca + ' ' + m.modelo + '</strong><br><small>' + m.id + '</small></td>' +
                        '<td>' + fmt(m.precioCompra) + '</td>' +
                        '<td>' + fmt(gastosA) + '</td>' +
                        '<td>' + fmt(gastosB) + '</td>' +
                        '<td>' + fmt(invTotal) + '</td>' +
                        '<td>' + fmt(m.precioVenta) + '</td>' +
                        '<td class="' + utilClass + '">' + fmt(calc.utilidadNeta) + '</td>' +
                        '<td>' + fmt(calc.devolucionA) + '</td>' +
                        '<td>' + fmt(calc.devolucionB) + '</td>' +
                        '<td>' + roiStr + '</td>' +
                    '</tr>');
                }
            } else {
                rows.push('<tr style="opacity:0.6">' +
                    '<td><strong>' + m.marca + ' ' + m.modelo + '</strong><br><small>' + m.id + ' - En proceso</small></td>' +
                    '<td>' + fmt(m.precioCompra) + '</td>' +
                    '<td>' + fmt(gastosA) + '</td>' +
                    '<td>' + fmt(gastosB) + '</td>' +
                    '<td>' + fmt(invTotal) + '</td>' +
                    '<td>-</td><td>-</td><td>-</td><td>-</td><td>-</td>' +
                '</tr>');
            }
        });

        if (rows.length === 0) {
            tbody.innerHTML = '';
            empty.style.display = 'block';
        } else {
            empty.style.display = 'none';
            tbody.innerHTML = rows.join('');
        }

        // Reparto global
        const pctA = config.pctA / 100;
        const pctB = config.pctB / 100;
        const gananciaA = totalUtilidad * pctA;
        const gananciaB = totalUtilidad * pctB;

        document.getElementById('reparto-global').innerHTML =
            '<div class="reparto-card">' +
                '<h3>' + config.socioA + '</h3>' +
                '<div class="reparto-row"><span class="reparto-label">Gastos realizados</span><span class="reparto-value">' + fmt(totalInvA) + '</span></div>' +
                '<div class="reparto-row"><span class="reparto-label">Devolucion inversion</span><span class="reparto-value">' + fmt(totalInvA) + '</span></div>' +
                '<div class="reparto-row"><span class="reparto-label">Ganancia (' + config.pctA + '%)</span><span class="reparto-value">' + fmt(gananciaA) + '</span></div>' +
                '<div class="reparto-row total"><span class="reparto-label">Total a recibir</span><span class="reparto-value">' + fmt(totalDevA) + '</span></div>' +
            '</div>' +
            '<div class="reparto-card">' +
                '<h3>' + config.socioB + '</h3>' +
                '<div class="reparto-row"><span class="reparto-label">Gastos realizados</span><span class="reparto-value">' + fmt(totalInvB) + '</span></div>' +
                '<div class="reparto-row"><span class="reparto-label">Devolucion inversion</span><span class="reparto-value">' + fmt(totalInvB) + '</span></div>' +
                '<div class="reparto-row"><span class="reparto-label">Ganancia (' + config.pctB + '%)</span><span class="reparto-value">' + fmt(gananciaB) + '</span></div>' +
                '<div class="reparto-row total"><span class="reparto-label">Total a recibir</span><span class="reparto-value">' + fmt(totalDevB) + '</span></div>' +
            '</div>' +
            '<div class="reparto-card">' +
                '<h3>Resumen General</h3>' +
                '<div class="reparto-row"><span class="reparto-label">Motos vendidas</span><span class="reparto-value">' + DataStore.getMotos().filter(m => m.precioVenta).length + '</span></div>' +
                '<div class="reparto-row"><span class="reparto-label">Utilidad neta total</span><span class="reparto-value ' + (totalUtilidad >= 0 ? 'amount-positive' : 'amount-negative') + '">' + fmt(totalUtilidad) + '</span></div>' +
                '<div class="reparto-row"><span class="reparto-label">ROI promedio</span><span class="reparto-value">' + Math.round(DataStore.getKPIs().roiPromedio * 100) + '%</span></div>' +
                '<div class="reparto-row"><span class="reparto-label">Dias prom. stock</span><span class="reparto-value">' + Math.round(DataStore.getKPIs().diasPromedio) + '</span></div>' +
            '</div>';
    }

    // ============================================================
    // CONFIG
    // ============================================================
    function refreshConfig() {
        const config = DataStore.getConfig();
        document.getElementById('config-socio-a').value = config.socioA;
        document.getElementById('config-socio-b').value = config.socioB;
        document.getElementById('config-pct-a').value = config.pctA;
        document.getElementById('config-pct-b').value = config.pctB;
        renderCategorias(config.categorias);
        validateConfig();
    }

    function renderCategorias(categorias) {
        const container = document.getElementById('config-categorias-list');
        container.innerHTML = categorias.map(c =>
            '<span class="cat-tag">' + c + ' <span class="remove-cat" data-cat="' + c + '">&times;</span></span>'
        ).join('');

        container.querySelectorAll('.remove-cat').forEach(btn => {
            btn.addEventListener('click', () => {
                const config = DataStore.getConfig();
                config.categorias = config.categorias.filter(c => c !== btn.dataset.cat);
                DataStore.setConfig(config);
                renderCategorias(config.categorias);
            });
        });
    }

    function validateConfig() {
        const pctA = Number(document.getElementById('config-pct-a').value);
        const pctB = Number(document.getElementById('config-pct-b').value);
        const validation = document.getElementById('config-validation');
        if (pctA + pctB === 100) {
            validation.textContent = 'OK - Los porcentajes suman 100%';
            validation.className = 'config-validation ok';
        } else {
            validation.textContent = 'Los porcentajes deben sumar 100% (actual: ' + (pctA + pctB) + '%)';
            validation.className = 'config-validation error';
        }
    }

    document.getElementById('config-pct-a').addEventListener('input', validateConfig);
    document.getElementById('config-pct-b').addEventListener('input', validateConfig);

    document.getElementById('btn-agregar-cat').addEventListener('click', () => {
        const input = document.getElementById('config-nueva-cat');
        const val = input.value.trim();
        if (!val) return;
        const config = DataStore.getConfig();
        if (config.categorias.includes(val)) {
            toast('La categoria ya existe', 'error');
            return;
        }
        config.categorias.push(val);
        DataStore.setConfig(config);
        renderCategorias(config.categorias);
        input.value = '';
        toast('Categoria agregada');
    });

    document.getElementById('btn-guardar-config').addEventListener('click', () => {
        const pctA = Number(document.getElementById('config-pct-a').value);
        const pctB = Number(document.getElementById('config-pct-b').value);

        if (pctA + pctB !== 100) {
            toast('Los porcentajes deben sumar 100%', 'error');
            return;
        }

        const config = DataStore.getConfig();
        config.socioA = document.getElementById('config-socio-a').value.trim() || 'Socio A';
        config.socioB = document.getElementById('config-socio-b').value.trim() || 'Socio B';
        config.pctA = pctA;
        config.pctB = pctB;

        DataStore.setConfig(config);
        toast('Configuracion guardada');
    });

    // ============================================================
    // EXPORT / IMPORT
    // ============================================================
    document.getElementById('btn-export').addEventListener('click', () => {
        const json = DataStore.exportJSON();
        const blob = new Blob([json], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'motoflip_backup_' + todayStr() + '.json';
        a.click();
        URL.revokeObjectURL(url);
        toast('Datos exportados', 'info');
    });

    document.getElementById('btn-import').addEventListener('click', () => {
        document.getElementById('import-file').click();
    });

    document.getElementById('import-file').addEventListener('change', e => {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = ev => {
            try {
                DataStore.importJSON(ev.target.result);
                toast('Datos importados correctamente');
                navigateTo('dashboard');
            } catch {
                toast('Error al importar el archivo', 'error');
            }
        };
        reader.readAsText(file);
        e.target.value = '';
    });

    // ============================================================
    // PUBLIC API (for inline onclick handlers)
    // ============================================================
    window.App = {
        abrirVenta(motoId) {
            const moto = DataStore.getMoto(motoId);
            if (!moto) return;

            const inversion = DataStore.getInversionTotal(motoId);
            const config = DataStore.getConfig();
            const gastosA = DataStore.getGastosSocio(motoId, config.socioA);
            const gastosB = DataStore.getGastosSocio(motoId, config.socioB);

            document.getElementById('venta-resumen').innerHTML =
                '<strong>' + moto.marca + ' ' + moto.modelo + ' (' + moto.id + ')</strong><br><br>' +
                'Precio compra: ' + fmt(moto.precioCompra) + '<br>' +
                'Gastos ' + config.socioA + ': ' + fmt(gastosA) + '<br>' +
                'Gastos ' + config.socioB + ': ' + fmt(gastosB) + '<br>' +
                '<strong>Inversion total: ' + fmt(inversion) + '</strong>';

            document.getElementById('venta-moto-id').value = motoId;
            document.getElementById('venta-precio').value = '';
            document.getElementById('venta-fecha').value = todayStr();
            openModal('modal-venta');
        },

        abrirEstado(motoId) {
            const moto = DataStore.getMoto(motoId);
            if (!moto) return;
            document.getElementById('estado-moto-id').value = motoId;
            document.getElementById('estado-nuevo').value = moto.estado;
            openModal('modal-estado');
        },

        abrirGastoParaMoto(motoId) {
            openGastoModal(motoId);
        },

        eliminarMoto(motoId) {
            if (!confirm('Eliminar moto ' + motoId + ' y todos sus gastos asociados?')) return;
            DataStore.deleteMoto(motoId);
            toast('Moto eliminada');
            refreshMotos();
        },

        eliminarGasto(gastoId) {
            if (!confirm('Eliminar este gasto?')) return;
            DataStore.deleteGasto(gastoId);
            toast('Gasto eliminado');
            refreshGastos();
        },

        // Called by DataStore on first data load
        init() {
            navigateTo('dashboard');
        },

        // Called by DataStore on real-time updates from other users
        refreshCurrent() {
            switch (currentSection) {
                case 'dashboard': refreshDashboard(); break;
                case 'motos': refreshMotos(); break;
                case 'gastos': refreshGastos(); break;
                case 'kanban': refreshKanban(); break;
                case 'ganancias': refreshGanancias(); break;
                case 'config': refreshConfig(); break;
            }
        }
    };

    // ============================================================
    // INIT - Start auth flow (app renders when data is ready)
    // ============================================================
    Auth.init();
});
