// Configuración de la app de canjistas (app.html). El dashboard tiene sus datos dentro de dashboard.html.
window.AJE = {
  // URL de la app web de Apps Script (termina en /exec).
  // Vacío = modo demo: la app guarda todo solo en el celular y no envía nada.
  API_URL: '',

  CAMPANA: 'Canje Mayorista AJE',
  MERCADOS: ['Productores', 'Caquetá', 'Unicachi Norte', 'Huamantanga'],

  // Mecánica (igual que la hoja "IMPULSO Y CANJES MAYORISTA LIMA 2026").
  // paq = paquetes mínimos de cada drop (D1..D5); FAMILIAR / PERSONAL = unidades de premio por drop.
  // El backend manda la versión vigente; esta copia se usa sin conexión.
  DROPS: {
    Productores: { paq: [24, 50, 150, 500, 800], FAMILIAR: [3, 6, 9, 15, 20], PERSONAL: [8, 15, 30, 45, 60] },
    otros:       { paq: [12, 25, 50, 80, 150],  FAMILIAR: [3, 6, 9, 15, 20], PERSONAL: [6, 12, 18, 24, 30] },
  },

  // Bebidas de premio que se usaron en la campaña, por formato. "Otra" permite escribir una nueva.
  BEBIDAS: {
    FAMILIAR: ['CIFRUT 3.03L', 'KR 3.03L', 'ORO 3.03L', 'CIFRUT 1.5L', 'FREE TEA'],
    PERSONAL: ['KOLITA 400 ML', 'CIELO MANZANA 600ML', 'AMAYU 600ML', 'CIFRUT 350ML', 'FREE TEA'],
  },

  INCIDENCIAS: ['No hay mercadería', 'Se acabaron los premios', 'Feriado o mercado cerrado', 'Problema con un cliente', 'Otro'],

  // Solo para el modo demo (sin backend). Con backend, los accesos están en la hoja Equipo.
  DEMO_USUARIOS: {
    productores: { clave: 'aje2026', nombre: 'Canjista Productores', mercado: 'Productores', rol: 'canjista' },
    caqueta:     { clave: 'aje2026', nombre: 'Canjista Caquetá', mercado: 'Caquetá', rol: 'canjista' },
    unicachi:    { clave: 'aje2026', nombre: 'Canjista Unicachi Norte', mercado: 'Unicachi Norte', rol: 'canjista' },
    huamantanga: { clave: 'aje2026', nombre: 'Canjista Huamantanga', mercado: 'Huamantanga', rol: 'canjista' },
    supervisor:  { clave: 'aje2026', nombre: 'Supervisor', mercado: '', rol: 'supervisor' },
  },
};
