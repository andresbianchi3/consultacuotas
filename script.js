// ==========================================
// 1. CONFIGURACIÓN Y BASE DE DATOS REAL
// ==========================================

// Valores base posibles de la cuota según hermanos/beneficio
const CUOTA_BASE_DEFAULT = 28000;

// Base de datos de afiliados (Generada desde tus archivos Excel/CSV)
const afiliados = {
    "46657886": { "nombre": "Basualdo, Lautaro", "dni": "46657886", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "46973053": { "nombre": "Bussoli, Maitena", "dni": "46973053", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05"] },
    "46222783": { "nombre": "Contreras, Alma Martina", "dni": "46222783", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "99000000": { "nombre": "CONTRERAS, Morena", "dni": "99000000", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "46885541": { "nombre": "Fisogni, Esteban Alejandro", "dni": "46885541", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04"] },
    "46032812": { "nombre": "GUZMAN, Facundo", "dni": "46032812", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05"] },
    "47174886": { "nombre": "Lazarte Castro, Micaela Agostina", "dni": "47174886", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "47662249": { "nombre": "Lunati, Sabrina Natalí", "dni": "47662249", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "46588508": { "nombre": "Mansilla, Gonzalo Ezequiel", "dni": "46588508", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05"] },
    "47001643": { "nombre": "Mansilla, Valentina Abril", "dni": "47001643", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "47273550": { "nombre": "Nuñez, Bárbara", "dni": "47273550", "ingreso": "2026-04-01", "cuotasPagas": ["2026-05", "2026-06"] },
    "99000001": { "nombre": "OLMEDO, Ivalu", "dni": "99000001", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "99000002": { "nombre": "Peirano, Lautaro", "dni": "99000002", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "99000003": { "nombre": "Pereyra, Hernán", "dni": "99000003", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "46973109": { "nombre": "Romero, Gonzalo  Agustín", "dni": "46973109", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05"] },
    "47175030": { "nombre": "Torreblanca, Lara Natalí", "dni": "47175030", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04"] },
    "46845188": { "nombre": "Uncal Meyer, Lautaro", "dni": "46845188", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "46506832": { "nombre": "Urtiaga, Abril", "dni": "46506832", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05"] },
    "45934634": { "nombre": "Urtiaga, Tomás", "dni": "45934634", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05"] },
    "48070608": { "nombre": "VALDIVIEZO BERNARDI, Emma", "dni": "48070608", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "47174851": { "nombre": "Vidal, Vittorio Máximo", "dni": "47174851", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "50938303": { "nombre": "ANTONELLO, Lara", "dni": "50938303", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "51276925": { "nombre": "BRUNO, Camilo", "dni": "51276925", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "49998113": { "nombre": "FLORES CEBOLLADA, Juliana Milena", "dni": "49998113", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "52158358": { "nombre": "GOMEZ PELLIZA, Malena", "dni": "52158358", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05"] },
    "48534049": { "nombre": "GUIÑAZÚ, Patricio Joaquin", "dni": "48534049", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "50629267": { "nombre": "GUZMAN, Joaquin", "dni": "50629267", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05"] },
    "51196628": { "nombre": "KORIAKOS, Jeremias", "dni": "51196628", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "50082001": { "nombre": "LEIVA Nayeli Luana", "dni": "50082001", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "50719356": { "nombre": "LESCANO, Natanael", "dni": "50719356", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "48722583": { "nombre": "MALICHO, Santino Tomas", "dni": "48722583", "ingreso": "2026-04-01", "cuotasPagas": ["2026-05", "2026-06"] },
    "52156281": { "nombre": "QUAGLIA, Marino", "dni": "52156281", "ingreso": "2026-04-01", "cuotasPagas": ["2026-05", "2026-06", "2026-07"] },
    "99000004": { "nombre": "RODRIGUEZ, Nawel", "dni": "99000004", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "52158860": { "nombre": "SCALTRITTI, Bianca", "dni": "52158860", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04"] },
    "50631328": { "nombre": "SEGURA Genaro", "dni": "50631328", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04"] },
    "50222330": { "nombre": "URTIAGA, Francisco gaspar", "dni": "50222330", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "99000005": { "nombre": "ZURCHSMITTEN, Santiago", "dni": "99000005", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06", "2026-07"] },
    "55007783": { "nombre": "ARGUELLO, Luisana", "dni": "55007783", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "53077829": { "nombre": "Arguello, Ramiro Benjamín", "dni": "53077829", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "52867806": { "nombre": "BORDA, Aurora", "dni": "52867806", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "54475873": { "nombre": "BRANE IRAZBAL, Catalina", "dni": "54475873", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04"] },
    "99000006": { "nombre": "CORDOBA, Guillermina", "dni": "99000006", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "52464435": { "nombre": "CORREA Alma", "dni": "52464435", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "55289982": { "nombre": "DE LA O'PRADOZ MARTOS, Joaquin", "dni": "55289982", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "52867860": { "nombre": "FLEITA, Pilar", "dni": "52867860", "ingreso": "2026-04-01", "cuotasPagas": ["2026-05", "2026-06"] },
    "55599595": { "nombre": "GUIÑAZU, Malena", "dni": "55599595", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "55290487": { "nombre": "HAUMAN LEIVA, Lisandro Hazziel", "dni": "55290487", "ingreso": "2026-04-01", "cuotasPagas": ["2026-05", "2026-06"] },
    "54759543": { "nombre": "HEREDIA, Ambar", "dni": "54759543", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "54019229": { "nombre": "KORIAKOS, Taiana", "dni": "54019229", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "54308719": { "nombre": "LUDUEÑA, Malena", "dni": "54308719", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "99000007": { "nombre": "MONITTO, Tobias", "dni": "99000007", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "55006627": { "nombre": "MORETTO, Leon", "dni": "55006627", "ingreso": "2026-04-01", "cuotasPagas": ["2026-05"] },
    "53507865": { "nombre": "PARASKEVOPULOS, Agustin", "dni": "53507865", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05"] },
    "53076137": { "nombre": "Peinado, Octavio Valentín", "dni": "53076137", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "52466914": { "nombre": "Pereyra, Sofía", "dni": "52466914", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "54472614": { "nombre": "PEREYRA, Victoria", "dni": "54472614", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "52896130": { "nombre": "SEQUEIRA, Lionel", "dni": "52896130", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05"] },
    "52866822": { "nombre": "Urtiaga, Joaquín", "dni": "52866822", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05"] },
    "52866821": { "nombre": "Urtiaga, Juan", "dni": "52866821", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04"] },
    "53265870": { "nombre": "Urtiaga, Pedro Valentin", "dni": "53265870", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "54761028": { "nombre": "Valdiviezo Berardi, Noah", "dni": "54761028", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "53263881": { "nombre": "Valdiviezo Bernardi, Liz", "dni": "53263881", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "53731909": { "nombre": "Vitale Masso, Martina", "dni": "53731909", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "56764563": { "nombre": "ALFONSO, Bianca", "dni": "56764563", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05"] },
    "55836993": { "nombre": "CORDOBA, Martina", "dni": "55836993", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "55523457": { "nombre": "CORREA Matheo", "dni": "55523457", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "57349836": { "nombre": "DIAZ GUZMAN, Jeremias", "dni": "57349836", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "57569042": { "nombre": "DVORKIN, Julian", "dni": "57569042", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "56305998": { "nombre": "DVORKIN, Zoe", "dni": "56305998", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "99000008": { "nombre": "GOMEZ, Francesca", "dni": "99000008", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "56978763": { "nombre": "GUERRICO, Pedro", "dni": "56978763", "ingreso": "2026-04-01", "cuotasPagas": ["2026-05"] },
    "55790635": { "nombre": "GUIGNARD, Lucia", "dni": "55790635", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "56537469": { "nombre": "HAUMAN LEIVA, Regino", "dni": "56537469", "ingreso": "2026-04-01", "cuotasPagas": ["2026-05", "2026-06", "2026-07"] },
    "56059756": { "nombre": "HEREDIA, India", "dni": "56059756", "ingreso": "2026-04-01", "cuotasPagas": [] },
    "56538935": { "nombre": "HIDALGO, Benicio", "dni": "56538935", "ingreso": "2026-04-01", "cuotasPagas": ["2026-05"] },
    "56976679": { "nombre": "JORNET LAVAYEN, Bautista", "dni": "56976679", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05"] },
    "56061055": { "nombre": "LUDUEÑA FASSI, Aldana", "dni": "56061055", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06"] },
    "99000009": { "nombre": "MALDONADO Alexander", "dni": "57152569", "ingreso": "2026-04-01", "cuotasPagas": ["2026-05"] },
    "56761267": { "nombre": "NAVARRETE, Alba", "dni": "56761267", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05"] },
    "57527420": { "nombre": "RECHE, Isabella", "dni": "57527420", "ingreso": "2026-04-01", "cuotasPagas": ["2026-05", "2026-06", "2026-07", "2026-08"] },
    "56535352": { "nombre": "SCARFO, Nazareno", "dni": "56535352", "ingreso": "2026-04-01", "cuotasPagas": ["2026-05", "2026-06", "2026-07"] },
    "55125830": { "nombre": "URTIAGA, Julian", "dni": "55125830", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05"] },
    "56763170": { "nombre": "VAIRA, Genaro", "dni": "56763170", "ingreso": "2026-04-01", "cuotasPagas": ["2026-04", "2026-05", "2026-06", "2026-07"] }
};

// Datos suministrados por el usuario (resumen + afiliados con cuotas abr-26..sep-26)
const datosExternos = {
    "resumen": {
        "total_afiliados": 88,
        "total_pagado": 8718800,
        "fecha_extraccion": "2026-08-10",
        "periodo": "Abril - Septiembre 2026"
    },
    "afiliados": [
        { "nombre": "ALDAY, Zoe", "grupo": "Caminantes", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 106000 },
        { "nombre": "ALFONSO, Bianca", "grupo": "Manada", "afiliacion": 60000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 28000, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 144000 },
        { "nombre": "ANTONELLO, Lara", "grupo": "Caminantes", "afiliacion": 50000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 50000 },
        { "nombre": "ARGUELLO, Luisana", "grupo": "Unidad", "afiliacion": 50000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 50000 },
        { "nombre": "Arguello, Ramiro Benjamín", "grupo": "Unidad", "afiliacion": 50000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 50000 },
        { "nombre": "BORDA, Aurora", "grupo": "Unidad", "afiliacion": 60000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 28000, "jul-26": 28000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 172000 },
        { "nombre": "BRANE IRAZBAL, Catalina", "grupo": "Unidad", "afiliacion": 60000, "cuotas": { "abr-26": 20000, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 80000 },
        { "nombre": "BRUNO, Camilo", "grupo": "Caminantes", "afiliacion": 50000, "cuotas": { "abr-26": 23000, "may-26": 23000, "jun-26": 17000, "jul-26": 23000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 136000 },
        { "nombre": "Basualdo, Lautaro", "grupo": "Rovers", "afiliacion": 50000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 50000 },
        { "nombre": "Bussoli, Maitena", "grupo": "Rovers", "afiliacion": 60000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 28000, "jul-26": 28000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 172000 },
        { "nombre": "CONTRERAS, Morena", "grupo": "Rovers", "afiliacion": 0, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 0 },
        { "nombre": "CORDOBA, Guillermina", "grupo": "Unidad", "afiliacion": 0, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 0 },
        { "nombre": "CORDOBA, Martina", "grupo": "Manada", "afiliacion": 60000, "cuotas": { "abr-26": 28000, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 88000 },
        { "nombre": "CORREA Alma", "grupo": "Unidad", "afiliacion": 0, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 0 },
        { "nombre": "CORREA Matheo", "grupo": "Manada", "afiliacion": 0, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 0 },
        { "nombre": "Contreras, Alma Martina", "grupo": "Rovers", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 28000, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 134000 },
        { "nombre": "DE LA O'PRADOZ MARTOS, Joaquin", "grupo": "Unidad", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 28000, "jul-26": 28000, "ago-26": 28000, "sep-26": 0 }, "total_pagado": 190000 },
        { "nombre": "DIAZ GUZMAN, Jeremias", "grupo": "Manada", "afiliacion": 50000, "cuotas": { "abr-26": 23000, "may-26": 23000, "jun-26": 23000, "jul-26": 23000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 142000 },
        { "nombre": "DVORKIN, Julian", "grupo": "Manada", "afiliacion": 60000, "cuotas": { "abr-26": 23000, "may-26": 23000, "jun-26": 23000, "jul-26": 28000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 157000 },
        { "nombre": "DVORKIN, Zoe", "grupo": "Manada", "afiliacion": 60000, "cuotas": { "abr-26": 23000, "may-26": 23000, "jun-26": 23000, "jul-26": 28000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 157000 },
        { "nombre": "FLEITA, Pilar", "grupo": "Unidad", "afiliacion": 60000, "cuotas": { "abr-26": 0, "may-26": 28000, "jun-26": 28000, "jul-26": 28000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 144000 },
        { "nombre": "FLORES CEBOLLADA, Juliana Milena", "grupo": "Caminantes", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 28000, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 134000 },
        { "nombre": "Fisogni, Esteban Alejandro", "grupo": "Rovers", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 28000, "jul-26": 28000, "ago-26": 28000, "sep-26": 0 }, "total_pagado": 190000 },
        { "nombre": "GALVAN, Bianca", "grupo": "Unidad", "afiliacion": 60000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 28000, "ago-26": 28000, "sep-26": 0 }, "total_pagado": 116000 },
        { "nombre": "GOMEZ PELLIZA, Malena", "grupo": "Caminantes", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 106000 },
        { "nombre": "GOMEZ, Francesca", "grupo": "Manada", "afiliacion": 60000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 28000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 88000 },
        { "nombre": "GUERRICO, Pedro", "grupo": "Manada", "afiliacion": 60000, "cuotas": { "abr-26": 0, "may-26": 28000, "jun-26": 28000, "jul-26": 28000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 144000 },
        { "nombre": "GUIGNARD, Lucia", "grupo": "Manada", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 28000, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 134000 },
        { "nombre": "GUIÑAZU, Malena", "grupo": "Unidad", "afiliacion": 50000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 50000 },
        { "nombre": "GUIÑAZÚ, Patricio Joaquin", "grupo": "Caminantes", "afiliacion": 50000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 50000 },
        { "nombre": "GUZMAN, Facundo", "grupo": "Rovers", "afiliacion": 60000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 116000 },
        { "nombre": "GUZMAN, Joaquin", "grupo": "Caminantes", "afiliacion": 60000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 116000 },
        { "nombre": "HAUMAN LEIVA, Lisandro Hazziel", "grupo": "Unidad", "afiliacion": 60000, "cuotas": { "abr-26": 0, "may-26": 21000, "jun-26": 21000, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 102000 },
        { "nombre": "HAUMAN LEIVA, Regino", "grupo": "Manada", "afiliacion": 60000, "cuotas": { "abr-26": 0, "may-26": 21000, "jun-26": 21000, "jul-26": 3000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 105000 },
        { "nombre": "HEREDIA, Ambar", "grupo": "Unidad", "afiliacion": 50000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 50000 },
        { "nombre": "HEREDIA, India", "grupo": "Manada", "afiliacion": 50000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 50000 },
        { "nombre": "HIDALGO, Benicio", "grupo": "Manada", "afiliacion": 60000, "cuotas": { "abr-26": 0, "may-26": 28000, "jun-26": 28000, "jul-26": 28000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 144000 },
        { "nombre": "JORNET LAVAYEN, Bautista", "grupo": "Manada", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 4000, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 82000 },
        { "nombre": "KORIAKOS, Jeremias", "grupo": "Caminantes", "afiliacion": 50000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 50000 },
        { "nombre": "KORIAKOS, Taiana", "grupo": "Unidad", "afiliacion": 50000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 50000 },
        { "nombre": "LEIVA Nayeli Luana", "grupo": "Caminantes", "afiliacion": 60000, "cuotas": { "abr-26": 21000, "may-26": 21000, "jun-26": 21000, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 123000 },
        { "nombre": "LESCANO, Natanael", "grupo": "Caminantes", "afiliacion": 0, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 0 },
        { "nombre": "LUDUEÑA FASSI, Aldana", "grupo": "Manada", "afiliacion": 50000, "cuotas": { "abr-26": 23000, "may-26": 23000, "jun-26": 20900, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 116900 },
        { "nombre": "LUDUEÑA, Malena", "grupo": "Unidad", "afiliacion": 50000, "cuotas": { "abr-26": 23000, "may-26": 23000, "jun-26": 20900, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 116900 },
        { "nombre": "Lazarte Castro, Micaela Agostina", "grupo": "Rovers", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 28000, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 134000 },
        { "nombre": "Lunati, Sabrina Natalí", "grupo": "Rovers", "afiliacion": 60000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 60000 },
        { "nombre": "MALDONA MIRANDEI Alexander", "grupo": "Manada", "afiliacion": 60000, "cuotas": { "abr-26": 0, "may-26": 28000, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 88000 },
        { "nombre": "MALICHO, Santino Tomas", "grupo": "Caminantes", "afiliacion": 60000, "cuotas": { "abr-26": 0, "may-26": 28000, "jun-26": 28000, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 116000 },
        { "nombre": "MONITTO, Tobias", "grupo": "Unidad", "afiliacion": 0, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 0 },
        { "nombre": "MORETTO, Leon", "grupo": "Unidad", "afiliacion": 60000, "cuotas": { "abr-26": 0, "may-26": 28000, "jun-26": 28000, "jul-26": 28000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 144000 },
        { "nombre": "Mansilla, Gonzalo Ezequiel", "grupo": "Rovers", "afiliacion": 60000, "cuotas": { "abr-26": 23000, "may-26": 23000, "jun-26": 23000, "jul-26": 23000, "ago-26": 23000, "sep-26": 0 }, "total_pagado": 175000 },
        { "nombre": "Mansilla, Valentina Abril", "grupo": "Rovers", "afiliacion": 60000, "cuotas": { "abr-26": 23000, "may-26": 23000, "jun-26": 23000, "jul-26": 23000, "ago-26": 23000, "sep-26": 0 }, "total_pagado": 175000 },
        { "nombre": "NAVARRETE, Alba", "grupo": "Manada", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 28000, "jul-26": 28000, "ago-26": 28000, "sep-26": 0 }, "total_pagado": 190000 },
        { "nombre": "Nuñez, Bárbara", "grupo": "Rovers", "afiliacion": 50000, "cuotas": { "abr-26": 0, "may-26": 28000, "jun-26": 28000, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 106000 },
        { "nombre": "OLMEDO, Ivalu", "grupo": "Rovers", "afiliacion": 0, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 0 },
        { "nombre": "PARASKEVOPULOS, Agustin", "grupo": "Unidad", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 106000 },
        { "nombre": "PEREYRA, Victoria", "grupo": "Unidad", "afiliacion": 60000, "cuotas": { "abr-26": 23000, "may-26": 23000, "jun-26": 23000, "jul-26": 23000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 152000 },
        { "nombre": "Peinado, Octavio Valentín", "grupo": "Unidad", "afiliacion": 50000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 50000 },
        { "nombre": "Peirano, Lautaro", "grupo": "Rovers", "afiliacion": 0, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 0 },
        { "nombre": "Pereyra, Hernán", "grupo": "Rovers", "afiliacion": 0, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 0 },
        { "nombre": "Pereyra, Sofía", "grupo": "Unidad", "afiliacion": 60000, "cuotas": { "abr-26": 23000, "may-26": 23000, "jun-26": 23000, "jul-26": 23000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 152000 },
        { "nombre": "QUAGLIA, Marino", "grupo": "Caminantes", "afiliacion": 50000, "cuotas": { "abr-26": 0, "may-26": 28000, "jun-26": 28000, "jul-26": 28000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 134000 },
        { "nombre": "RECHE, Isabella", "grupo": "Manada", "afiliacion": 60000, "cuotas": { "abr-26": 0, "may-26": 28000, "jun-26": 28000, "jul-26": 28000, "ago-26": 28000, "sep-26": 0 }, "total_pagado": 172000 },
        { "nombre": "RODRIGUEZ, Nawel", "grupo": "Caminantes", "afiliacion": 0, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 0 },
        { "nombre": "ROMERO, Gael Bautista", "grupo": "Manada", "afiliacion": 0, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 0 },
        { "nombre": "ROMERO, Noah Benjamin", "grupo": "Unidad", "afiliacion": 0, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 0 },
        { "nombre": "Romero, Gonzalo  Agustín", "grupo": "Rovers", "afiliacion": 60000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 116000 },
        { "nombre": "SCALTRITTI, Bianca", "grupo": "Caminantes", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 78000 },
        { "nombre": "SCARFO, Nazareno", "grupo": "Manada", "afiliacion": 60000, "cuotas": { "abr-26": 0, "may-26": 28000, "jun-26": 28000, "jul-26": 14000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 130000 },
        { "nombre": "SEGURA Genaro", "grupo": "Caminantes", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 78000 },
        { "nombre": "SEQUEIRA, Lionel", "grupo": "Unidad", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 0, "jun-26": 28000, "jul-26": 28000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 134000 },
        { "nombre": "TISSERA Nicole", "grupo": "Unidad", "afiliacion": 60000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 28000, "sep-26": 0 }, "total_pagado": 88000 },
        { "nombre": "Torreblanca, Lara Natalí", "grupo": "Rovers", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 78000 },
        { "nombre": "URTIAGA, Francisco gaspar", "grupo": "Caminantes", "afiliacion": 50000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 50000 },
        { "nombre": "URTIAGA, Julian", "grupo": "Manada", "afiliacion": 50000, "cuotas": { "abr-26": 21000, "may-26": 21000, "jun-26": 21000, "jul-26": 21000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 134000 },
        { "nombre": "Uncal Meyer, Lautaro", "grupo": "Rovers", "afiliacion": 50000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 50000 },
        { "nombre": "Urtiaga, Abril", "grupo": "Rovers", "afiliacion": 50000, "cuotas": { "abr-26": 21000, "may-26": 21000, "jun-26": 21000, "jul-26": 21000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 134000 },
        { "nombre": "Urtiaga, Joaquín", "grupo": "Unidad", "afiliacion": 50000, "cuotas": { "abr-26": 21000, "may-26": 21000, "jun-26": 21000, "jul-26": 21000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 134000 },
        { "nombre": "Urtiaga, Juan", "grupo": "Unidad", "afiliacion": 50000, "cuotas": { "abr-26": 21000, "may-26": 21000, "jun-26": 21000, "jul-26": 21000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 134000 },
        { "nombre": "Urtiaga, Pedro Valentin", "grupo": "Unidad", "afiliacion": 50000, "cuotas": { "abr-26": 23000, "may-26": 23000, "jun-26": 23000, "jul-26": 23000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 142000 },
        { "nombre": "Urtiaga, Tomás", "grupo": "Rovers", "afiliacion": 50000, "cuotas": { "abr-26": 21000, "may-26": 21000, "jun-26": 21000, "jul-26": 21000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 134000 },
        { "nombre": "VAIRA, Genaro", "grupo": "Manada", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 28000, "jul-26": 28000, "ago-26": 0, "sep-26": 0 }, "total_pagado": 162000 },
        { "nombre": "VALDIVIEZO BERNARDI, Emma", "grupo": "Rovers", "afiliacion": 50000, "cuotas": { "abr-26": 21000, "may-26": 21000, "jun-26": 21000, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 113000 },
        { "nombre": "Valdiviezo Berardi, Noah", "grupo": "Unidad", "afiliacion": 50000, "cuotas": { "abr-26": 21000, "may-26": 21000, "jun-26": 21000, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 113000 },
        { "nombre": "Valdiviezo Bernardi, Liz", "grupo": "Unidad", "afiliacion": 50000, "cuotas": { "abr-26": 21000, "may-26": 21000, "jun-26": 21000, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 113000 },
        { "nombre": "Vidal, Vittorio Máximo", "grupo": "Rovers", "afiliacion": 60000, "cuotas": { "abr-26": 0, "may-26": 0, "jun-26": 0, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 60000 },
        { "nombre": "Vitale Masso, Martina", "grupo": "Unidad", "afiliacion": 60000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 28000, "jul-26": 0, "ago-26": 0, "sep-26": 0 }, "total_pagado": 144000 },
        { "nombre": "ZURCHSMITTEN, Santiago", "grupo": "Caminantes", "afiliacion": 50000, "cuotas": { "abr-26": 28000, "may-26": 28000, "jun-26": 28000, "jul-26": 28000, "ago-26": 28000, "sep-26": 0 }, "total_pagado": 190000 }
    ]
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
        .replace(/[^A-Z0-9\s]/g, '')
        .replace(/\s+/g, ' ');
}

function buscarRegistroPorNombre(nombreAfiliado) {
    const nombreNormalizado = normalizarNombre(nombreAfiliado);
    const tokensAfiliado = nombreNormalizado.split(' ');
    const apellidoAfiliado = tokensAfiliado[0] || '';
    const nombrePropioAfiliado = tokensAfiliado[tokensAfiliado.length - 1] || '';

    for (const item of cuotasAbrJulioData) {
        const nombreJson = normalizarNombre(item.Nombre);
        if (nombreJson === nombreNormalizado) return item;
    }

    for (const item of cuotasAbrJulioData) {
        const nombreJson = normalizarNombre(item.Nombre);
        const tokensJson = nombreJson.split(' ');
        const nombrePropioJson = tokensJson[tokensJson.length - 1] || '';
        const apellidoJson = tokensJson[0] || '';

        if (nombrePropioJson !== nombrePropioAfiliado) continue;
        if (apellidoJson === apellidoAfiliado) return item;
        if (apellidoJson.startsWith(apellidoAfiliado.slice(0, 6)) || apellidoAfiliado.startsWith(apellidoJson.slice(0, 6))) {
            return item;
        }
    }

    return null;
}

async function cargarCuotasJson() {
    try {
        // Si hay datos proporcionados localmente, convertirlos al formato esperado
        if (typeof datosExternos !== 'undefined' && Array.isArray(datosExternos.afiliados)) {
            cuotasAbrJulioData = datosExternos.afiliados.map(item => {
                return {
                    Nombre: item.nombre,
                    Abril: (item.cuotas && item.cuotas['abr-26']) ? String(item.cuotas['abr-26']) : '0',
                    Mayo: (item.cuotas && item.cuotas['may-26']) ? String(item.cuotas['may-26']) : '0',
                    Junio: (item.cuotas && item.cuotas['jun-26']) ? String(item.cuotas['jun-26']) : '0',
                    Julio: (item.cuotas && item.cuotas['jul-26']) ? String(item.cuotas['jul-26']) : '0',
                    Agosto: (item.cuotas && item.cuotas['ago-26']) ? String(item.cuotas['ago-26']) : '0',
                    Septiembre: (item.cuotas && item.cuotas['sep-26']) ? String(item.cuotas['sep-26']) : '0'
                };
            });
            cuotasPorNombre = new Map(
                cuotasAbrJulioData.map(item => [normalizarNombre(item.Nombre), item])
            );
            return;
        }

        const archivos = ['cuotas.json', 'cuotas_abril_julio.json'];
        let datosCargados = null;

        for (const archivo of archivos) {
            try {
                const respuesta = await fetch(archivo);
                if (!respuesta.ok) continue;
                datosCargados = await respuesta.json();
                break;
            } catch (error) {
                continue;
            }
        }

        if (!datosCargados) {
            throw new Error('No se encontró un JSON válido de cuotas');
        }

        cuotasAbrJulioData = datosCargados.map(item => {
            return {
                Nombre: item.Nombre || item.nombre || '',
                Abril: item.Abril ?? item['abr-26'] ?? '0',
                Mayo: item.Mayo ?? item['may-26'] ?? '0',
                Junio: item.Junio ?? item['jun-26'] ?? '0',
                Julio: item.Julio ?? item['jul-26'] ?? '0',
                Agosto: item.Agosto ?? item['ago-26'] ?? '0',
                Septiembre: item.Septiembre ?? item['sep-26'] ?? '0'
            };
        });

        cuotasPorNombre = new Map(
            cuotasAbrJulioData.map(item => [normalizarNombre(item.Nombre), item])
        );
    } catch (error) {
        console.error('No se pudo cargar ningún JSON de cuotas:', error);
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
    let afiliado = afiliados[dni] ? { ...afiliados[dni] } : null;
    if (!afiliado) {
        const clave = Object.keys(afiliados).find(key => afiliados[key].dni === dni);
        if (clave) {
            afiliado = { ...afiliados[clave] };
        }
    }
    if (!afiliado) return null;

    // Asegurar que siempre exista el arreglo de cuotas pagas en el objeto afiliado
    afiliado.cuotasPagas = afiliado.cuotasPagas || [];

    let registroJson = cuotasPorNombre.get(normalizarNombre(afiliado.nombre));
    if (!registroJson) {
        registroJson = buscarRegistroPorNombre(afiliado.nombre);
    }
    afiliado.cuotasAbrJulio = registroJson || null;
    return afiliado;
}

function parseImporteJson(valor) {
    if (valor === null || valor === undefined || valor === '') return 0;

    const texto = String(valor).trim();
    if (!texto || texto === 'null' || texto === 'undefined' || texto.startsWith('=')) return 0;

    let limpio = texto.replace(/[$\s]/g, '');

    if (limpio.includes(',') && limpio.includes('.')) {
        limpio = limpio.replace(/\./g, '').replace(',', '.');
    } else if (limpio.includes(',') && !limpio.includes('.')) {
        limpio = limpio.replace(',', '.');
    } else if (limpio.includes('.') && limpio.split('.').length > 2) {
        const partes = limpio.split('.');
        const ultimo = partes.pop();
        limpio = partes.join('') + '.' + ultimo;
    }

    const numeric = limpio.replace(/[^0-9.\-]/g, '');
    if (!numeric || numeric === '-' || numeric === '.' || numeric === '-.') return 0;

    const numero = Number.parseFloat(numeric);
    return Number.isFinite(numero) ? numero : 0;
}

function determinarCuotaBase(registroJson) {
    if (!registroJson) return CUOTA_BASE_DEFAULT;

    const pagos = ['Abril', 'Mayo', 'Junio', 'Julio', 'Agosto']
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
        { codigoPeriodo: '2026-07', mesTexto: 'Julio 2026', key: 'Julio' },
        { codigoPeriodo: '2026-08', mesTexto: 'Agosto 2026', key: 'Agosto' },
        { codigoPeriodo: '2026-09', mesTexto: 'Septiembre 2026', key: 'Septiembre' }
    ];

    const hoy = new Date();
    const codigoPeriodoActual = `${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, '0')}`;
    const registroJson = afiliado.cuotasAbrJulio;
    const cuotaBase = determinarCuotaBase(registroJson);
    const listadoCuotas = [];

    // Determinar índice de ingreso dentro de los periodos para marcar meses anteriores como "no aplica"
    let ingresoIndex = 0;
    if (afiliado.ingreso) {
        const ingresoPeriodo = afiliado.ingreso.slice(0, 7); // 'YYYY-MM'
        ingresoIndex = periodosFijos.findIndex(p => p.codigoPeriodo === ingresoPeriodo);
        if (ingresoIndex === -1) {
            const ingresoDate = new Date(afiliado.ingreso);
            const firstDate = new Date(periodosFijos[0].codigoPeriodo + '-01');
            const lastDate = new Date(periodosFijos[periodosFijos.length - 1].codigoPeriodo + '-01');
            if (ingresoDate <= firstDate) {
                ingresoIndex = 0;
            } else if (ingresoDate > lastDate) {
                ingresoIndex = periodosFijos.length;
            } else {
                ingresoIndex = periodosFijos.findIndex(p => new Date(p.codigoPeriodo + '-01') >= ingresoDate);
            }
        }
    }

    const pagos = periodosFijos.map((periodo, idx) => {
        // No computar pagos para meses anteriores al ingreso
        if (idx < ingresoIndex) return 0;
        const importeRegistrado = registroJson ? parseImporteJson(registroJson[periodo.key]) : 0;
        if ((importeRegistrado === 0 || isNaN(importeRegistrado)) && Array.isArray(afiliado.cuotasPagas) && afiliado.cuotasPagas.includes(periodo.codigoPeriodo)) {
            return cuotaBase;
        }
        return importeRegistrado;
    });

    const pagosConsecutivos = [...pagos];
    for (let i = pagosConsecutivos.length - 1; i > 0; i--) {
        if (pagosConsecutivos[i] > 0 && pagosConsecutivos[i - 1] === 0 && (i - 1) >= ingresoIndex) {
            pagosConsecutivos[i - 1] = cuotaBase;
        }
    }

    periodosFijos.forEach((periodo, index) => {
        const importePagado = (index < ingresoIndex)
            ? 0
            : (registroJson ? parseImporteJson(registroJson[periodo.key]) : (Array.isArray(afiliado.cuotasPagas) && afiliado.cuotasPagas.includes(periodo.codigoPeriodo) ? cuotaBase : 0));
        const importePagadoEfectivo = pagosConsecutivos[index] || 0;
        const esMesAnteriorAlIngreso = index < ingresoIndex;
        const esMesPosteriorAlActual = periodo.codigoPeriodo > codigoPeriodoActual;
        const importeAdeudado = (esMesAnteriorAlIngreso || esMesPosteriorAlActual) ? 0 : Math.max(0, cuotaBase - importePagadoEfectivo);
        const estaPagado = !esMesAnteriorAlIngreso && !esMesPosteriorAlActual && importePagadoEfectivo >= cuotaBase && importeAdeudado === 0;
        const importePagadoVisual = (esMesAnteriorAlIngreso || esMesPosteriorAlActual) ? 0 : (estaPagado ? cuotaBase : importePagado);

        listadoCuotas.push({
            codigoPeriodo: periodo.codigoPeriodo,
            mesTexto: periodo.mesTexto,
            importePagado,
            importePagadoVisual,
            importeAdeudado,
            pagado: estaPagado,
            pagoParcial: !esMesAnteriorAlIngreso && !esMesPosteriorAlActual && importePagado > 0 && !estaPagado,
            noAplica: esMesAnteriorAlIngreso || esMesPosteriorAlActual
        });
    });

    afiliado.cuotaBase = cuotaBase;
    return listadoCuotas;
}

function mostrarResumen(afiliado, listadoCuotas) {
    // Excluir meses marcados como noAplica de los totales
    const cuotasAplicables = listadoCuotas.filter(item => !item.noAplica);
    const totalCuotas = cuotasAplicables.length;
    const cuotasPagas = cuotasAplicables.filter(item => item.pagado).length;
    const cuotasPendientes = cuotasAplicables.filter(item => item.importeAdeudado > 0).length;
    const totalAdeudado = cuotasAplicables.reduce((sum, item) => sum + item.importeAdeudado, 0);

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

    // Determinar si puede participar con tolerancia de X meses (por defecto 1)
    function puedeParticiparConTolerancia(listadoCuotas, toleranciaMeses = 1) {
        if (!Array.isArray(listadoCuotas) || listadoCuotas.length === 0) return false;
        const hoy = new Date();
        const year = hoy.getFullYear();
        const month = String(hoy.getMonth() + 1).padStart(2, '0');
        const codigoActual = `${year}-${month}`;

        let indiceActual = listadoCuotas.findIndex(item => item.codigoPeriodo === codigoActual);
        if (indiceActual === -1) {
            // Si el periodo actual no está en la lista, usar el último periodo conocido
            indiceActual = listadoCuotas.length - 1;
        }

        let ultimoPagoIndex = -1;
        listadoCuotas.forEach((item, idx) => {
            if (item.pagado) ultimoPagoIndex = idx;
        });

        return ultimoPagoIndex >= (indiceActual - toleranciaMeses);
    }

    const puedeParticipar = puedeParticiparConTolerancia(listadoCuotas, 1);
    if (badgeEstado && boxTotalAdeudado) {
        if (totalAdeudado > 0) {
            badgeEstado.className = "status-pill debt";
            badgeEstado.innerHTML = `<i class="bi bi-exclamation-circle-fill"></i> ⚠️ DEBE PLATA`;
            boxTotalAdeudado.className = "metric-box metric-highlight debt";
            if (participationNotice) {
                participationNotice.style.display = 'block';
                if (puedeParticipar) {
                    participationNotice.className = 'participation-note clear';
                    participationNotice.innerHTML = `<strong>⚠️ DEUDA REGISTRADA:</strong> Debe ${formatearPesos(totalAdeudado)} en ${cuotasPendientes} cuota(s). Podrá participar con tolerancia, pero debe regularizar.`;
                } else {
                    participationNotice.className = 'participation-note debt';
                    participationNotice.innerHTML = `<strong>❌ DEUDA IMPORTANTE:</strong> Debe ${formatearPesos(totalAdeudado)} en ${cuotasPendientes} cuota(s). No podrá participar hasta regularizar.`;
                }
            }
        } else {
            badgeEstado.className = "status-pill clean";
            badgeEstado.innerHTML = `<i class="bi bi-check-circle-fill"></i> Al día`;
            boxTotalAdeudado.className = "metric-box metric-highlight";
            if (participationNotice) {
                participationNotice.style.display = 'block';
                participationNotice.className = 'participation-note clear';
                participationNotice.innerHTML = `<strong>Eventos fuera del grupo:</strong> Podrá participar.`;
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
        if (cuota.noAplica) {
            cellImporte.textContent = '-';
        } else if (cuota.pagado) {
            cellImporte.textContent = formatearPesos(cuota.importePagadoVisual);
        } else if (cuota.pagoParcial) {
            cellImporte.textContent = `${formatearPesos(cuota.importePagado)} pagado / ${formatearPesos(cuota.importeAdeudado)} falta`;
        } else {
            cellImporte.textContent = `${formatearPesos(0)} pagado / ${formatearPesos(cuota.importeAdeudado)} falta`;
        }
        fila.appendChild(cellImporte);

        const cellEstado = document.createElement('td');
        const spanEstado = document.createElement('span');

        if (cuota.noAplica) {
            spanEstado.className = "cell-status clear";
            spanEstado.innerHTML = `- No aplica`;
        } else if (cuota.pagado) {
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
