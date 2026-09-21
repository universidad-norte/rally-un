function doGet(e) {
  let view = (e && e.parameter && e.parameter.view) ? e.parameter.view : 'portada';
  
  let templateName = 'portada';
  if (view === 'captura') templateName = 'index';
  if (view === 'reporte') templateName = 'reporte';
  if (view === 'itinerario') templateName = 'itinerario'; // 👈 Nueva vista de Itinerario / Impresión

  let template = HtmlService.createTemplateFromFile(templateName);
  template.scriptUrl = ScriptApp.getService().getUrl();

  return template.evaluate()
      .setTitle('Rally Universitario')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Obtener el Itinerario Completo (Rondas, Estaciones y Enfrentamientos)
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
        mapaEstaciones[dataEst[i][0]] = dataEst[i][1]; // ID -> Nombre de Actividad
      }
    }
  }

  let listaFixture = [];
  
  for (let i = 1; i < dataFixture.length; i++) {
    let numEstacion = dataFixture[i][1];
    let ronda = dataFixture[i][2];
    let eq1Clave = dataFixture[i][3];
    let eq2Clave = dataFixture[i][4];
    
    if (ronda) {
      listaFixture.push({
        numEstacion: numEstacion,
        nombreEstacion: mapaEstaciones[numEstacion] || `Estación ${numEstacion}`,
        ronda: ronda,
        eq1Clave: eq1Clave,
        eq1Nombre: mapaEquipos[eq1Clave] ? mapaEquipos[eq1Clave].nombre : eq1Clave,
        eq1Color: mapaEquipos[eq1Clave] ? mapaEquipos[eq1Clave].color : '#6c757d',
        eq2Clave: eq2Clave,
        eq2Nombre: mapaEquipos[eq2Clave] ? mapaEquipos[eq2Clave].nombre : eq2Clave,
        eq2Color: mapaEquipos[eq2Clave] ? mapaEquipos[eq2Clave].color : '#6c757d'
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

// Función para obtener el mapeo de Equipos (Nombre y Color Hexadecimal)
function getMapaEquipos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEquipos = ss.getSheetByName("Equipos");
  const data = sheetEquipos.getDataRange().getValues();
  
  let mapaEquipos = {};
  
  for (let i = 1; i < data.length; i++) {
    let letra = data[i][1];     // Columna B: Letra / Clave (A, B, C...)
    let color = data[i][3];    // Columna D: Código Hexadecimal (ej. #FF0000)
    let nombre = data[i][4];   // Columna E: Nombre del Equipo
    
    if (letra) {
      mapaEquipos[letra] = {
        nombre: nombre || `Equipo ${letra}`,
        color: color || '#6c757d'
      };
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
function getTablaPosiciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetFixture = ss.getSheetByName("Fixture");
  const data = sheetFixture.getDataRange().getValues();
  const mapaEquipos = getMapaEquipos();
  
  let puntos = {};
  
  Object.keys(mapaEquipos).forEach(clave => {
    puntos[clave] = {
      nombre: mapaEquipos[clave].nombre,
      color: mapaEquipos[clave].color,
      pts: 0
    };
  });
  
  for (let i = 1; i < data.length; i++) {
    let eq1 = data[i][3];
    let eq2 = data[i][4];
    let ganador = data[i][5];
    let estatus = data[i][6];
    
    if (estatus === "FINALIZADO") {
      if (ganador === eq1) {
        if (puntos[eq1]) puntos[eq1].pts += 3;
      } else if (ganador === eq2) {
        if (puntos[eq2]) puntos[eq2].pts += 3;
      } else if (ganador === "EMPATE") {
        if (puntos[eq1]) puntos[eq1].pts += 1;
        if (puntos[eq2]) puntos[eq2].pts += 1;
      }
    }
  }
  
  let ranking = Object.values(puntos);
  ranking.sort((a, b) => b.pts - a.pts);
  
  return ranking;
}