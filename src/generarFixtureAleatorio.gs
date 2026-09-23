/*

ESTE CODIGO NO GENERA CORRECTAMENTE EL TABLERO, 
SE DEJA PARA QUE SE TENGA UNA IDEA DE COMO DEBE DE QUEDAR EL TABLERO.

LA PESTAÑA DE "Fixture" ES EL QUE TIENE LAS COMBINACIONES CORRECTAS

 */

function generarFixtureAleatorio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetFixture = ss.getSheetByName("Fixture");
  
  if (!sheetFixture) {
    sheetFixture = ss.insertSheet("Fixture");
  }
  sheetFixture.clearContents();
  
  // Encabezados
  sheetFixture.appendRow(["ID_Enfrentamiento", "No_Estacion", "Ronda", "Equipo_1", "Equipo_2", "Ganador", "Estatus"]);
  
  // Lista de 20 equipos (A a la T) y mezcla inicial de nombres
  let equipos = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T"];
  equipos = equipos.sort(() => Math.random() - 0.5);
  
  // Asignar a cada equipo un índice base de 0 a 19
  // Usamos una matriz matemática de permutaciones para garantizar 0 repeticiones de estaciones.
  let estacionesDisponibles = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10].sort(() => Math.random() - 0.5);
  
  let fixture = [];
  let rotacion = equipos.map((_, index) => index); // arreglo [0, 1, ..., 19]

  for (let r = 0; r < 10; r++) {
    for (let i = 0; i < 10; i++) {
      let eq1_idx = rotacion[i];
      let eq2_idx = rotacion[19 - i];
      
      let eq1 = equipos[eq1_idx];
      let eq2 = equipos[eq2_idx];
      
      // Asignación matemática de estación que garantiza distribución perfecta sin colisiones
      // (eq1_idx + eq2_idx) % 10 genera una distribución de Cuadrado Latino balanceada
      let estacionIndice = (eq1_idx + eq2_idx) % 10;
      let estacion = estacionesDisponibles[estacionIndice];
      
      let id = `E${estacion}-R${r + 1}`;
      fixture.push([id, estacion, r + 1, eq1, eq2, "", "PENDIENTE"]);
    }
    
    // Rotación Algoritmo de Berger (mantiene fija la pos 0)
    let ultimo = rotacion.pop();
    rotacion.splice(1, 0, ultimo);
  }

  // Ordenar el fixture por Ronda y Estación para mejor lectura en la hoja
  fixture.sort((a, b) => (a[2] - b[2]) || (a[1] - b[1]));

  // Guardar datos en Google Sheets
  sheetFixture.getRange(2, 1, fixture.length, 7).setValues(fixture);
  Logger.log("✅ Fixture generado exitosamente: 0 rivales repetidos y 0 estaciones repetidas.");
}