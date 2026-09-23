function doGet(e) {
  let view = (e && e.parameter && e.parameter.view) ? e.parameter.view : 'portada';
  
  let templateName = 'portada';
  if (view === 'captura') templateName = 'index';
  if (view === 'reporte') templateName = 'reporte';
  if (view === 'itinerario') templateName = 'itinerario';
  if (view === 'dashboard') templateName = 'dashboard'; // 👈 Nueva vista de Dashboard General

  let template = HtmlService.createTemplateFromFile(templateName);
  template.scriptUrl = ScriptApp.getService().getUrl();

  return template.evaluate()
      .setTitle('Rally Universitario - Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Obtener datos consolidados para el Dashboard General
function getDatosDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetFixture = ss.getSheetByName("Fixture");
  
  if (!sheetFixture) return { estaciones: [], rondas: [], matriz: {}, resumen: {} };

  const dataFixture = sheetFixture.getDataRange().getValues();
  const mapaEquipos = getMapaEquipos();
  
  // Nombres de Estaciones
  let sheetEstaciones = ss.getSheetByName("Estaciones") || ss.getSheetByName("Estacion");
  let mapaEstaciones = {};
  if (sheetEstaciones) {
    let dataEst = sheetEstaciones.getDataRange().getValues();
    for (let i = 1; i < dataEst.length; i++) {
      if (dataEst[i][0] !== "") {
        mapaEstaciones[dataEst[i][0]] = dataEst[i][1];
      }
    }
  }

  let matrizEnfrentamientos = {};
  let totalPartidos = 0;
  let finalizados = 0;

  for (let i = 1; i < dataFixture.length; i++) {
    let numEstacion = dataFixture[i][1];
    let ronda = dataFixture[i][2];
    let eq1Val = String(dataFixture[i][3]).trim();
    let eq2Val = String(dataFixture[i][4]).trim();
    let ganadorVal = String(dataFixture[i][5]).trim();
    let estatus = String(dataFixture[i][6]).trim();

    if (ronda && numEstacion) {
      totalPartidos++;
      if (estatus === "FINALIZADO") finalizados++;

      let infoEq1 = mapaEquipos[eq1Val] || { nombre: eq1Val, color: '#6c757d' };
      let infoEq2 = mapaEquipos[eq2Val] || { nombre: eq2Val, color: '#6c757d' };
      let infoGanador = mapaEquipos[ganadorVal] || null;

      let claveMatriz = `E${numEstacion}-R${ronda}`;
      matrizEnfrentamientos[claveMatriz] = {
        estacion: numEstacion,
        ronda: ronda,
        eq1Nombre: infoEq1.nombre,
        eq1Color: infoEq1.color,
        eq2Nombre: infoEq2.nombre,
        eq2Color: infoEq2.color,
        estatus: estatus,
        ganador: ganadorVal,
        ganadorNombre: infoGanador ? infoGanador.nombre : (ganadorVal === 'EMPATE' ? '🤝 Empate' : '')
      };
    }
  }

  return {
    mapaEstaciones: mapaEstaciones,
    matriz: matrizEnfrentamientos,
    resumen: {
      total: totalPartidos,
      finalizados: finalizados,
      pendientes: totalPartidos - finalizados,
      porcentaje: totalPartidos > 0 ? Math.round((finalizados / totalPartidos) * 100) : 0
    }
  };
}

// Obtener Itinerario Completo para itinerario.html
function getItinerarioCompleto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetFixture = ss.getSheetByName("Fixture");
  
  if (!sheetFixture) return { estaciones: {}, mapaEquipos: {}, fixture: [] };

  const dataFixture = sheetFixture.getDataRange().getValues();
  const mapaEquipos = getMapaEquipos();
  
  // Mapeo de Estaciones
  let sheetEstaciones = ss.getSheetByName("Estaciones") || ss.getSheetByName("Estacion");
  let mapaEstaciones = {};
  if (sheetEstaciones) {
    let dataEst = sheetEstaciones.getDataRange().getValues();
    for (let i = 1; i < dataEst.length; i++) {
      if (dataEst[i][0] !== "") {
        mapaEstaciones[dataEst[i][0]] = dataEst[i][1];
      }
    }
  }

  let listaFixture = [];
  
  for (let i = 1; i < dataFixture.length; i++) {
    let numEstacion = dataFixture[i][1];
    let ronda = dataFixture[i][2];
    let eq1Val = String(dataFixture[i][3]).trim(); // "Equipo 10" o "10" o "J"
    let eq2Val = String(dataFixture[i][4]).trim(); // "Equipo 17" o "17" o "Q"
    
    if (ronda && eq1Val !== "") {
      // Buscar información del Equipo 1
      let infoEq1 = mapaEquipos[eq1Val] || { nombre: eq1Val, color: '#6c757d' };
      // Buscar información del Equipo 2
      let infoEq2 = mapaEquipos[eq2Val] || { nombre: eq2Val, color: '#6c757d' };

      listaFixture.push({
        numEstacion: numEstacion,
        nombreEstacion: mapaEstaciones[numEstacion] || `Estación ${numEstacion}`,
        ronda: ronda,
        eq1Clave: eq1Val,
        eq1Nombre: infoEq1.nombre,
        eq1Color: infoEq1.color,
        eq2Clave: eq2Val,
        eq2Nombre: infoEq2.nombre,
        eq2Color: infoEq2.color
      });
    }
  }

  return {
    mapaEquipos: mapaEquipos,
    fixture: listaFixture
  };
}

// Obtener la lista dinámica de estaciones desde la hoja "Estaciones" o "Estacion"
function getListaEstaciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Busca la hoja 'Estaciones' o 'Estacion'
  let sheet = ss.getSheetByName("Estaciones");
  
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  let estaciones = [];
  
  // Asumiendo que la fila 1 son los encabezados
  // Columna A (data[i][0]) = No_Estacion
  // Columna B (data[i][1]) = Nombre_Actividad
  for (let i = 1; i < data.length; i++) {
    let numEstacion = data[i][0];
    let nombreActividad = data[i][1];
    
    if (numEstacion !== "" && numEstacion !== null) {
      estaciones.push({
        id: numEstacion,
        nombre: `Estación ${numEstacion} - ${nombreActividad}`
      });
    }
  }
  
  return estaciones;
}

// Validar PIN de Organizadores (Puedes cambiar '1234' por el PIN que tú elijas)
function validarPinOrganizador(pinIngresado) {
  const PIN_CORRECTO = "1234"; // 👈 CAMBIA TU PIN AQUÍ
  return pinIngresado === PIN_CORRECTO;
}

// Obtener mapa flexible de Equipos (Acepta Clave, Número o Nombre)
function getMapaEquipos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEquipos = ss.getSheetByName("Equipos") || ss.getSheetByName("Equipo");
  
  if (!sheetEquipos) return {};

  const data = sheetEquipos.getDataRange().getValues();
  let mapaEquipos = {};

  for (let i = 1; i < data.length; i++) {
    let numEquipo = String(data[i][0]).trim();  // Columna A: ID (1, 2, 3...)
    let letraClave = String(data[i][1]).trim(); // Columna B: Letra (A, B, C...)
    let color = String(data[i][3]).trim();      // Columna D: Hexadecimal (#FF0000)
    let nombre = String(data[i][4]).trim();     // Columna E: Nombre asignado

    if (numEquipo !== "" && numEquipo !== "undefined") {
      let infoEquipo = {
        id: numEquipo,
        nombre: nombre || `Equipo ${numEquipo}`,
        color: color || '#6c757d'
      };

      // Guardamos la referencia principal por ID único
      mapaEquipos[numEquipo] = infoEquipo;
      
      // Alias para búsqueda rápida en otras funciones
      mapaEquipos[`Equipo ${numEquipo}`] = infoEquipo;
      mapaEquipos[letraClave] = infoEquipo;
      if (nombre) mapaEquipos[nombre] = infoEquipo;
    }
  }
  return mapaEquipos;
}

// Obtener la ronda actual/pendiente de una estación específica
function getEnfrentamientoActual(numEstacion) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Fixture");
  const data = sheet.getDataRange().getValues();
  const mapaEquipos = getMapaEquipos();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == numEstacion && data[i][6] === "PENDIENTE") {
      let eq1Clave = data[i][3];
      let eq2Clave = data[i][4];
      
      return {
        row: i + 1,
        id: data[i][0],
        estacion: data[i][1],
        ronda: data[i][2],
        equipo1: {
          clave: eq1Clave,
          nombre: mapaEquipos[eq1Clave] ? mapaEquipos[eq1Clave].nombre : eq1Clave,
          color: mapaEquipos[eq1Clave] ? mapaEquipos[eq1Clave].color : '#6c757d'
        },
        equipo2: {
          clave: eq2Clave,
          nombre: mapaEquipos[eq2Clave] ? mapaEquipos[eq2Clave].nombre : eq2Clave,
          color: mapaEquipos[eq2Clave] ? mapaEquipos[eq2Clave].color : '#6c757d'
        }
      };
    }
  }
  return null;
}

// Registrar resultado (3 Pts ganador / 1 Pt empate)
function registrarResultado(row, ganadorClave, eq1Clave, eq2Clave) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Fixture");
  
  let ptsEq1 = 0;
  let ptsEq2 = 0;
  
  if (ganadorClave === "EMPATE") {
    ptsEq1 = 1;
    ptsEq2 = 1;
  } else if (ganadorClave === eq1Clave) {
    ptsEq1 = 3;
    ptsEq2 = 0;
  } else if (ganadorClave === eq2Clave) {
    ptsEq1 = 0;
    ptsEq2 = 3;
  }
  
  sheet.getRange(row, 6).setValue(ganadorClave); // Columna F: Ganador
  sheet.getRange(row, 7).setValue("FINALIZADO"); // Columna G: Estatus
  sheet.getRange(row, 8).setValue(ptsEq1);       // Columna H: Puntos Equipo 1
  sheet.getRange(row, 9).setValue(ptsEq2);       // Columna I: Puntos Equipo 2
  
  return { success: true };
}

// Obtener tabla general de puntuaciones para el Reporte
// Obtener Tabla General de Posiciones sin duplicados (Exactamente 20 equipos)
function getTablaPosiciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetFixture = ss.getSheetByName("Fixture");
  const dataFixture = sheetFixture.getDataRange().getValues();
  const mapaEquipos = getMapaEquipos();
  
  // 1. Usar un Map/Diccionario único filtrando solo por ID numérico del equipo
  let tablaPuntos = {};
  
  Object.keys(mapaEquipos).forEach(clave => {
    let eq = mapaEquipos[clave];
    // Solo registrar si la clave es el ID único (ej. "1", "2", "20")
    if (clave === eq.id && !tablaPuntos[eq.id]) {
      tablaPuntos[eq.id] = {
        id: eq.id,
        nombre: eq.nombre,
        color: eq.color,
        pts: 0
      };
    }
  });

  // 2. Función auxiliar para encontrar el ID único a partir de cualquier texto ("Equipo 1", "A", "1", etc.)
  function obtenerIdEquipo(valor) {
    let valStr = String(valor).trim();
    if (mapaEquipos[valStr]) return mapaEquipos[valStr].id;
    return null;
  }

  // 3. Sumar puntos de partidos finalizados
  for (let i = 1; i < dataFixture.length; i++) {
    let eq1Val = dataFixture[i][3];
    let eq2Val = dataFixture[i][4];
    let ganadorVal = dataFixture[i][5];
    let estatus = dataFixture[i][6];
    
    let ptsEq1 = Number(dataFixture[i][7]) || 0;
    let ptsEq2 = Number(dataFixture[i][8]) || 0;
    
    if (estatus === "FINALIZADO") {
      let id1 = obtenerIdEquipo(eq1Val);
      let id2 = obtenerIdEquipo(eq2Val);
      
      if (id1 && tablaPuntos[id1]) tablaPuntos[id1].pts += ptsEq1;
      if (id2 && tablaPuntos[id2]) tablaPuntos[id2].pts += ptsEq2;
    }
  }

  // 4. Convertir a Arreglo y ordenar descendentemente por Puntos
  let ranking = Object.values(tablaPuntos);
  ranking.sort((a, b) => b.pts - a.pts);
  
  return ranking;
}