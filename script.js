// ==========================================
// 1. CONFIGURACIÓN Y BASE DE DATOS REAL
// ==========================================

// Valores base posibles de la cuota según hermanos/beneficio
const CUOTA_BASE_DEFAULT = 28000;

// Base de datos de afiliados (Generada desde tus archivos Excel/CSV)
const afiliados = {
    "46657886": { "nombre": "Basualdo, Lautaro", "dni": "46657886", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "46973053": { "nombre": "Bussoli, Maitena", "dni": "46973053", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05"] },
    "46222783": { "nombre": "Contreras, Alma Martina", "dni": "46222783", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "99000000": { "nombre": "CONTRERAS, Morena", "dni": "99000000", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "46885541": { "nombre": "Fisogni, Esteban Alejandro", "dni": "46885541", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04"] },
    "46032812": { "nombre": "GUZMAN, Facundo", "dni": "46032812", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05"] },
    "47174886": { "nombre": "Lazarte Castro, Micaela Agostina", "dni": "47174886", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "47662249": { "nombre": "Lunati, Sabrina Natalí", "dni": "47662249", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "46588508": { "nombre": "Mansilla, Gonzalo Ezequiel", "dni": "46588508", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05"] },
    "47001643": { "nombre": "Mansilla, Valentina Abril", "dni": "47001643", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "47273550": { "nombre": "Nuñez, Bárbara", "dni": "47273550", "ingreso": "2024-04-01", "cuotasPagas": ["2024-05", "2024-06"] },
    "99000001": { "nombre": "OLMEDO, Ivalu", "dni": "99000001", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "99000002": { "nombre": "Peirano, Lautaro", "dni": "99000002", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "99000003": { "nombre": "Pereyra, Hernán", "dni": "99000003", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "46973109": { "nombre": "Romero, Gonzalo  Agustín", "dni": "46973109", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05"] },
    "47175030": { "nombre": "Torreblanca, Lara Natalí", "dni": "47175030", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04"] },
    "46845188": { "nombre": "Uncal Meyer, Lautaro", "dni": "46845188", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "46506832": { "nombre": "Urtiaga, Abril", "dni": "46506832", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05"] },
    "45934634": { "nombre": "Urtiaga, Tomás", "dni": "45934634", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05"] },
    "48070608": { "nombre": "VALDIVIEZO BERNARDI, Emma", "dni": "48070608", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "47174851": { "nombre": "Vidal, Vittorio Máximo", "dni": "47174851", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "50938303": { "nombre": "ANTONELLO, Lara", "dni": "50938303", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "51276925": { "nombre": "BRUNO, Camilo", "dni": "51276925", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "49998113": { "nombre": "FLORES CEBOLLADA, Juliana Milena", "dni": "49998113", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "52158358": { "nombre": "GOMEZ PELLIZA, Malena", "dni": "52158358", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05"] },
    "48534049": { "nombre": "GUIÑAZÚ, Patricio Joaquin", "dni": "48534049", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "50629267": { "nombre": "GUZMAN, Joaquin", "dni": "50629267", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05"] },
    "51196628": { "nombre": "KORIAKOS, Jeremias", "dni": "51196628", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "50082001": { "nombre": "LEIVA Nayeli Luana", "dni": "50082001", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "50719356": { "nombre": "LESCANO, Natanael", "dni": "50719356", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "48722583": { "nombre": "MALICHO, Santino Tomas", "dni": "48722583", "ingreso": "2024-04-01", "cuotasPagas": ["2024-05", "2024-06"] },
    "52156281": { "nombre": "QUAGLIA, Marino", "dni": "52156281", "ingreso": "2024-04-01", "cuotasPagas": ["2024-05", "2024-06", "2024-07"] },
    "99000004": { "nombre": "RODRIGUEZ, Nawel", "dni": "99000004", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "52158860": { "nombre": "SCALTRITTI, Bianca", "dni": "52158860", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04"] },
    "50631328": { "nombre": "SEGURA Genaro", "dni": "50631328", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04"] },
    "50222330": { "nombre": "URTIAGA, Francisco gaspar", "dni": "50222330", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "99000005": { "nombre": "ZURCHSMITTEN, Santiago", "dni": "99000005", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06", "2024-07"] },
    "55007783": { "nombre": "ARGUELLO, Luisana", "dni": "55007783", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "53077829": { "nombre": "Arguello, Ramiro Benjamín", "dni": "53077829", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "52867806": { "nombre": "BORDA, Aurora", "dni": "52867806", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "54475873": { "nombre": "BRANE IRAZBAL, Catalina", "dni": "54475873", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04"] },
    "99000006": { "nombre": "CORDOBA, Guillermina", "dni": "99000006", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "52464435": { "nombre": "CORREA Alma", "dni": "52464435", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "55289982": { "nombre": "DE LA O'PRADOZ MARTOS, Joaquin", "dni": "55289982", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "52867860": { "nombre": "FLEITA, Pilar", "dni": "52867860", "ingreso": "2024-04-01", "cuotasPagas": ["2024-05", "2024-06"] },
    "55599595": { "nombre": "GUIÑAZU, Malena", "dni": "55599595", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "55290487": { "nombre": "HAUMAN LEIVA, Lisandro Hazziel", "dni": "55290487", "ingreso": "2024-04-01", "cuotasPagas": ["2024-05", "2024-06"] },
    "54759543": { "nombre": "HEREDIA, Ambar", "dni": "54759543", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "54019229": { "nombre": "KORIAKOS, Taiana", "dni": "54019229", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "54308719": { "nombre": "LUDUEÑA, Malena", "dni": "54308719", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "99000007": { "nombre": "MONITTO, Tobias", "dni": "99000007", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "55006627": { "nombre": "MORETTO, Leon", "dni": "55006627", "ingreso": "2024-04-01", "cuotasPagas": ["2024-05"] },
    "53507865": { "nombre": "PARASKEVOPULOS, Agustin", "dni": "53507865", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05"] },
    "53076137": { "nombre": "Peinado, Octavio Valentín", "dni": "53076137", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "52466914": { "nombre": "Pereyra, Sofía", "dni": "52466914", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "54472614": { "nombre": "PEREYRA, Victoria", "dni": "54472614", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "52896130": { "nombre": "SEQUEIRA, Lionel", "dni": "52896130", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05"] },
    "52866822": { "nombre": "Urtiaga, Joaquín", "dni": "52866822", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05"] },
    "52866821": { "nombre": "Urtiaga, Juan", "dni": "52866821", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04"] },
    "53265870": { "nombre": "Urtiaga, Pedro Valentin", "dni": "53265870", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "54761028": { "nombre": "Valdiviezo Berardi, Noah", "dni": "54761028", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "53263881": { "nombre": "Valdiviezo Bernardi, Liz", "dni": "53263881", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "53731909": { "nombre": "Vitale Masso, Martina", "dni": "53731909", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "56764563": { "nombre": "ALFONSO, Bianca", "dni": "56764563", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05"] },
    "55836993": { "nombre": "CORDOBA, Martina", "dni": "55836993", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "55523457": { "nombre": "CORREA Matheo", "dni": "55523457", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "57349836": { "nombre": "DIAZ GUZMAN, Jeremias", "dni": "57349836", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "57569042": { "nombre": "DVORKIN, Julian", "dni": "57569042", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "56305998": { "nombre": "DVORKIN, Zoe", "dni": "56305998", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "99000008": { "nombre": "GOMEZ, Francesca", "dni": "99000008", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "56978763": { "nombre": "GUERRICO, Pedro", "dni": "56978763", "ingreso": "2024-04-01", "cuotasPagas": ["2024-05"] },
    "55790635": { "nombre": "GUIGNARD, Lucia", "dni": "55790635", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "56537469": { "nombre": "HAUMAN LEIVA, Regino", "dni": "56537469", "ingreso": "2024-04-01", "cuotasPagas": ["2024-05", "2024-06", "2024-07"] },
    "56059756": { "nombre": "HEREDIA, India", "dni": "56059756", "ingreso": "2024-04-01", "cuotasPagas": [] },
    "56538935": { "nombre": "HIDALGO, Benicio", "dni": "56538935", "ingreso": "2024-04-01", "cuotasPagas": ["2024-05"] },
    "56976679": { "nombre": "JORNET LAVAYEN, Bautista", "dni": "56976679", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05"] },
    "56061055": { "nombre": "LUDUEÑA FASSI, Aldana", "dni": "56061055", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06"] },
    "99000009": { "nombre": "MALDONA MIRANDEI Alexander", "dni": "57152569", "ingreso": "2024-04-01", "cuotasPagas": ["2024-05"] },
    "56761267": { "nombre": "NAVARRETE, Alba", "dni": "56761267", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05"] },
    "57527420": { "nombre": "RECHE, Isabella", "dni": "57527420", "ingreso": "2024-04-01", "cuotasPagas": ["2024-05", "2024-06", "2024-07", "2024-08"] },
    "56535352": { "nombre": "SCARFO, Nazareno", "dni": "56535352", "ingreso": "2024-04-01", "cuotasPagas": ["2024-05", "2024-06", "2024-07"] },
    "55125830": { "nombre": "URTIAGA, Julian", "dni": "55125830", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05"] },
    "56763170": { "nombre": "VAIRA, Genaro", "dni": "56763170", "ingreso": "2024-04-01", "cuotasPagas": ["2024-04", "2024-05", "2024-06", "2024-07"] }
};

// ==========================================
// 2. CAPTURA DE ELEMENTOS Y EVENTOS
// ==========================================

const dniInput = document.getElementById('dniInput');
const btnConsultar = document.getElementById('btnConsultar');
const errorCard = document.getElementById('errorCard');
const dashboardLayout = document.getElementById('dashboardLayout');
const tableCard = document.getElementById('tableCard');

let cuotasAbrJulioData = [];
let cuotasPorNombre = new Map();

btnConsultar.addEventListener('click', manejarConsulta);
dniInput.addEventListener('keypress', function(e) {
    if (e.key === 'Enter') manejarConsulta();
});
dniInput.addEventListener('input', limpiarVistas);

cargarCuotasJson();

function normalizarNombre(nombre) {
    return nombre
        .trim()
        .toUpperCase()
        .normalize('NFD')
        .replace(/\p{Diacritic}/gu, '')
        .replace(/\s+/g, ' ');
}

async function cargarCuotasJson() {
    try {
        const respuesta = await fetch('cuotas_abril_julio.json');
        if (!respuesta.ok) throw new Error(`HTTP ${respuesta.status}`);

        cuotasAbrJulioData = await respuesta.json();
        cuotasPorNombre = new Map(
            cuotasAbrJulioData.map(item => [normalizarNombre(item.Nombre), item])
        );
    } catch (error) {
        console.error('No se pudo cargar cuotas_abril_julio.json:', error);
    }
}

// ==========================================
// 3. FUNCIONES DE LÓGICA PRINCIPAL
// ==========================================

function manejarConsulta() {
    const dniBuscado = dniInput.value.trim();
    if (!dniBuscado) return;

    limpiarVistas();

    const afiliadoEncontrado = buscarAfiliado(dniBuscado);
    if (!afiliadoEncontrado) {
        errorCard.style.display = 'flex';
        return;
    }

    const listadoCuotasCalculadas = calcularCuotas(afiliadoEncontrado);
    mostrarResumen(afiliadoEncontrado, listadoCuotasCalculadas);
    generarHistorial(listadoCuotasCalculadas);
}

function buscarAfiliado(dni) {
    const afiliado = afiliados[dni] ? { ...afiliados[dni] } : null;
    if (!afiliado) return null;

    const registroJson = cuotasPorNombre.get(normalizarNombre(afiliado.nombre));
    afiliado.cuotasAbrJulio = registroJson || null;
    return afiliado;
}

function parseImporteJson(valor) {
    if (typeof valor !== 'string') return 0;
    const numeric = valor.replace(/[^0-9.,-]/g, '').replace(/,/g, '');
    return parseFloat(numeric) || 0;
}

function determinarCuotaBase(registroJson) {
    if (!registroJson) return CUOTA_BASE_DEFAULT;

    const pagos = ['Abril', 'Mayo', 'Junio', 'Julio']
        .map(mes => parseImporteJson(registroJson[mes]))
        .filter(valor => valor > 0);

    if (pagos.length === 0) {
        return CUOTA_BASE_DEFAULT;
    }

    const mayorPago = Math.max(...pagos);
    if (mayorPago === 21000) return 21000;
    if (mayorPago === 23000) return 23000;
    return CUOTA_BASE_DEFAULT;
}

function calcularCuotas(afiliado) {
    const periodosFijos = [
        { codigoPeriodo: '2026-04', mesTexto: 'Abril 2026', key: 'Abril' },
        { codigoPeriodo: '2026-05', mesTexto: 'Mayo 2026', key: 'Mayo' },
        { codigoPeriodo: '2026-06', mesTexto: 'Junio 2026', key: 'Junio' },
        { codigoPeriodo: '2026-07', mesTexto: 'Julio 2026', key: 'Julio' }
    ];

    const registroJson = afiliado.cuotasAbrJulio;
    const cuotaBase = determinarCuotaBase(registroJson);
    const listadoCuotas = [];

    periodosFijos.forEach(periodo => {
        const importePagado = registroJson
            ? parseImporteJson(registroJson[periodo.key])
            : 0;
        const importeAdeudado = Math.max(0, cuotaBase - importePagado);
        const estaPagado = importeAdeudado === 0;
        const importePagadoVisual = estaPagado ? cuotaBase : importePagado;

        listadoCuotas.push({
            codigoPeriodo: periodo.codigoPeriodo,
            mesTexto: periodo.mesTexto,
            importePagado,
            importePagadoVisual,
            importeAdeudado,
            pagado: estaPagado,
            pagoParcial: importePagado > 0 && !estaPagado
        });
    });

    afiliado.cuotaBase = cuotaBase;
    return listadoCuotas;
}

function mostrarResumen(afiliado, listadoCuotas) {
    const totalCuotas = listadoCuotas.length;
    const cuotasPagas = listadoCuotas.filter(item => item.pagado).length;
    const cuotasPendientes = listadoCuotas.filter(item => item.importeAdeudado > 0).length;
    const totalAdeudado = listadoCuotas.reduce((sum, item) => sum + item.importeAdeudado, 0);

    document.getElementById('lblNombre').textContent = afiliado.nombre;
    document.getElementById('lblDni').textContent = afiliado.dni;
    document.getElementById('lblIngreso').textContent = formatearFechaARG(afiliado.ingreso);
    document.getElementById('lblCuotaBase').textContent = formatearPesos(afiliado.cuotaBase || CUOTA_BASE_DEFAULT);

    document.getElementById('metTotal').textContent = totalCuotas;
    document.getElementById('metPagas').textContent = cuotasPagas;
    document.getElementById('metPendientes').textContent = cuotasPendientes;
    document.getElementById('metDeuda').textContent = formatearPesos(totalAdeudado);

    const badgeEstado = document.getElementById('badgeEstado');
    const boxTotalAdeudado = document.getElementById('boxTotalAdeudado');
    const participationNotice = document.getElementById('participationNotice');

    if (badgeEstado && boxTotalAdeudado) {
        if (totalAdeudado > 0) {
            badgeEstado.className = "status-pill debt";
            badgeEstado.innerHTML = `<i class="bi bi-x-circle-fill"></i> Posee deuda`;
            boxTotalAdeudado.className = "metric-box metric-highlight debt";
            if (participationNotice) {
                participationNotice.style.display = 'block';
                participationNotice.className = 'participation-note debt';
                participationNotice.innerHTML = `<strong>Salida 11/07 al Camping San Martín:</strong> Por favor regularice su deuda para poder participar.`;
            }
        } else {
            badgeEstado.className = "status-pill clean";
            badgeEstado.innerHTML = `<i class="bi bi-check-circle-fill"></i> Sin deuda`;
            boxTotalAdeudado.className = "metric-box metric-highlight";
            if (participationNotice) {
                participationNotice.style.display = 'block';
                participationNotice.className = 'participation-note clear';
                participationNotice.innerHTML = `<strong>Salida 11/07 al Campign San Martín:</strong> El protgaonista puede participar. Mantenga sus cuotas al día.`;
            }
        }
    }

    if (dashboardLayout) {
        dashboardLayout.style.display = 'grid';
    }
}

function generarHistorial(listadoCuotas) {
    const tbody = document.getElementById('tablaHistorialBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    const listadoInvertido = [...listadoCuotas].reverse();

    listadoInvertido.forEach(cuota => {
        const fila = document.createElement('tr');
        fila.className = cuota.pagado ? 'row-paid' : 'row-debt';

        const cellMes = document.createElement('td');
        cellMes.style.fontWeight = "600";
        cellMes.textContent = cuota.mesTexto;
        fila.appendChild(cellMes);

        const cellImporte = document.createElement('td');
        if (cuota.pagado) {
            cellImporte.textContent = formatearPesos(cuota.importePagadoVisual);
        } else if (cuota.pagoParcial) {
            cellImporte.textContent = `${formatearPesos(cuota.importePagado)} pagado / ${formatearPesos(cuota.importeAdeudado)} falta`;
        } else {
            cellImporte.textContent = `${formatearPesos(0)} pagado / ${formatearPesos(cuota.importeAdeudado)} falta`;
        }
        fila.appendChild(cellImporte);

        const cellEstado = document.createElement('td');
        const spanEstado = document.createElement('span');

        if (cuota.pagado) {
            spanEstado.className = "cell-status paid";
            spanEstado.innerHTML = `✔ Pagado`;
        } else if (cuota.pagoParcial) {
            spanEstado.className = "cell-status pending";
            spanEstado.innerHTML = `⚠ Pago parcial`;
        } else {
            spanEstado.className = "cell-status pending";
            spanEstado.innerHTML = `✖ Pendiente`;
        }

        cellEstado.appendChild(spanEstado);
        fila.appendChild(cellEstado);
        tbody.appendChild(fila);
    });

    if (tableCard) {
        tableCard.style.display = 'block';
    }
}

function formatearPesos(valor) {
    return '$' + valor.toLocaleString('es-AR', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    });
}

function formatearFechaARG(fechaString) {
    const fragmentos = fechaString.split('-');
    if (fragmentos.length !== 3) return fechaString;
    return `${fragmentos[2]}/${fragmentos[1]}/${fragmentos[0]}`;
}

function limpiarVistas() {
    errorCard.style.display = 'none';
    dashboardLayout.style.display = 'none';
    if (tableCard) {
        tableCard.style.display = 'none';
    }
    const participationNotice = document.getElementById('participationNotice');
    if (participationNotice) {
        participationNotice.style.display = 'none';
    }
}
