function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('CANIA - Consumos Clínicos')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function obtenerDatosIniciales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const pacientesData = ss.getSheetByName('Pacientes').getDataRange().getValues().filter(r => String(r[0]).trim() !== "");
  const inventarioData = ss.getSheetByName('Inventario').getDataRange().getValues().filter(r => String(r[0]).trim() !== "");
  const registrosData = ss.getSheetByName('Registros').getDataRange().getValues().filter(r => String(r[0]).trim() !== "");
  
  const cronicosSheet = ss.getSheetByName('Resumen_Cronicos');
  let cronicosData = [];
  if (cronicosSheet) {
      cronicosData = cronicosSheet.getDataRange().getValues().filter(r => String(r[0]).trim() !== "");
  }

  // NUEVO: Lectura de la hoja de Diagnósticos (Col A: Nx, Col B: Abr_Nx, Col C: As, Col D: Abr_As)
  const diagSheet = ss.getSheetByName('Data_Diagnosticos');
  let diagData = [];
  if (diagSheet) {
      // Filtramos filas que tengan al menos algún dato en Nx o As
      diagData = diagSheet.getDataRange().getValues().filter(r => String(r[0]).trim() !== "" || String(r[2]).trim() !== "");
  }
  
  const ssTz = ss.getSpreadsheetTimeZone();
  const sanitizar = (datos) => {
    return datos.map(fila => fila.map(celda => {
      if (celda instanceof Date) {
        return Utilities.formatDate(celda, ssTz, "yyyy-MM-dd'T'HH:mm:ss");
      }
      return celda;
    }));
  };
  
  return {
    pacientes: pacientesData.length > 1 ? sanitizar(pacientesData.slice(1)) : [],
    inventario: inventarioData.length > 1 ? sanitizar(inventarioData.slice(1)) : [],
    registros: registrosData.length > 1 ? sanitizar(registrosData.slice(1)).reverse() : [],
    cronicos: cronicosData.length > 1 ? sanitizar(cronicosData.slice(1)) : [],
    diagnosticos: diagData.length > 1 ? sanitizar(diagData.slice(1)) : []
  };
}

function agregarPaciente(datos) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pacientes');
    const data = sheet.getDataRange().getValues();
    
    const triajeNuevo = String(datos.id).trim();
    const nombreCompletoNuevo = (datos.apellidos + " " + datos.nombres).toLowerCase().trim();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === triajeNuevo) {
         return { success: false, message: `El número de triaje ${triajeNuevo} ya existe en el sistema.` };
      }
      const nombreCompletoExistente = (String(data[i][1]) + " " + String(data[i][2])).toLowerCase().trim();
      if (nombreCompletoExistente === nombreCompletoNuevo) {
        return { success: false, message: "Ya existe un paciente con estos Nombres y Apellidos." };
      }
    }
    
    sheet.appendRow([datos.id, datos.apellidos, datos.nombres, datos.tipo, datos.diagnosticoNx, datos.diagnosticoAs]);
    SpreadsheetApp.flush();
    return { success: true, data: [datos.id, datos.apellidos, datos.nombres, datos.tipo, datos.diagnosticoNx, datos.diagnosticoAs] };
  } catch (error) {
    return { success: false, message: "Error al guardar paciente: " + error.toString() };
  }
}

// Comparador Global Inteligente
const normalizarTexto = (str) => {
  if (!str) return "";
  return String(str).normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase().replace(/\s+/g, ' ');
};

const isLoteEqual = (l1, l2) => {
  let a = normalizarTexto(l1).replace(/[^A-Z0-9]/g, ''); 
  let b = normalizarTexto(l2).replace(/[^A-Z0-9]/g, '');
  const vacios = ["", "SL", "SINLOTE", "NA", "0", "N", "S"];
  
  if (vacios.includes(a)) a = "";
  if (vacios.includes(b)) b = "";
  
  a = a.replace(/^0+/, '');
  b = b.replace(/^0+/, '');

  return a === b;
};

function guardarConsumo(paciente, modalidad, fechaEntrega, carrito) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let emailUsuario = 'Usuario Web';
    try { const email = Session.getActiveUser().getEmail(); if (email) emailUsuario = email; } catch(e) {}
    
    const fechaActual = new Date();
    const mes = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
    const anio = fechaActual.getFullYear();
    const mesAnio = `${mes}/${anio}`;

    // --- 1. GUARDAR EN HOJA 'Registros' ---
    const sheetReg = ss.getSheetByName('Registros');
    const filasAInsertar = carrito.map((item, index) => {
      const idRegistro = "REG-" + fechaActual.getTime() + "-" + index;
      return [
        idRegistro, fechaActual, mesAnio, paciente.nombreCompleto, modalidad, paciente.tipo, 
        item.nombre, item.cantidad, emailUsuario, "Consumido", "", fechaEntrega
      ];
    });

    if (filasAInsertar.length > 0) {
      const dataColA = sheetReg.getRange("A:A").getValues();
      let ultimaFila = 0;
      for (let i = dataColA.length - 1; i >= 0; i--) { if (dataColA[i][0] !== "") { ultimaFila = i + 1; break; } }
      if (ultimaFila === 0) ultimaFila = 1; 
      sheetReg.getRange(ultimaFila + 1, 1, filasAInsertar.length, filasAInsertar[0].length).setValues(filasAInsertar);
    }

    // --- 2. DESCONTAR DEL INVENTARIO ---
    const sheetInv = ss.getSheetByName('Inventario');
    if(sheetInv) {
      const invData = sheetInv.getDataRange().getValues();
      carrito.forEach(item => {
        const iNombre = normalizarTexto(item.nombre);
        for(let i=1; i<invData.length; i++) {
          const nombreSAP = normalizarTexto(invData[i][2]);
          const nombreReal = normalizarTexto(invData[i][11]); 
          const loteInv = invData[i][1];

          if((nombreSAP === iNombre || nombreReal === iNombre) && isLoteEqual(loteInv, item.lote)) {
            let cantActual = Number(invData[i][3]) || 0;
            sheetInv.getRange(i+1, 4).setValue(cantActual - Number(item.cantidad));
            invData[i][3] = cantActual - Number(item.cantidad); 
            break; 
          }
        }
      });
    }

    // --- 3. ACTUALIZAR HOJAS DIARIAS ---
    const tipoPac = normalizarTexto(paciente.tipo);
    const esCronico = (tipoPac === 'CRONICO');
    const sheetDiariaName = esCronico ? 'Cronicos' : 'Ambulatorio';
    const sheetDiaria = ss.getSheetByName(sheetDiariaName);
    
    if (sheetDiaria) {
      const dataDiaria = sheetDiaria.getDataRange().getValues();
      
      carrito.forEach(item => {
        let found = false;
        let unidadMedida = item.unidad || "UND";
        const iNombre = normalizarTexto(item.nombre);

        if (esCronico) {
            for(let i=1; i<dataDiaria.length; i++) {
              const nombreCol = normalizarTexto(dataDiaria[i][2]); 
              const loteCol = dataDiaria[i][5];                    

              if(nombreCol !== "" && nombreCol === iNombre && isLoteEqual(loteCol, item.lote)) {
                 let cantActual = Number(dataDiaria[i][4]) || 0; 
                 sheetDiaria.getRange(i+1, 5).setValue(cantActual + Number(item.cantidad)); 
                 dataDiaria[i][4] = cantActual + Number(item.cantidad); 
                 found = true; 
                 break;
              }
            }
            if(!found) {
              sheetDiaria.appendRow([
                "CAB000000001", "", item.nombre, unidadMedida, item.cantidad, "'" + item.lote, "", "", 
                item.nombre, "Farmacia/Fórmula", fechaEntrega, modalidad
              ]);
            }
        } else {
            for(let i=1; i<dataDiaria.length; i++) {
              const nombreCol = normalizarTexto(dataDiaria[i][0]); 
              const loteCol = dataDiaria[i][3];                    

              if(nombreCol !== "" && nombreCol === iNombre && isLoteEqual(loteCol, item.lote)) {
                 let cantActual = Number(dataDiaria[i][2]) || 0; 
                 sheetDiaria.getRange(i+1, 3).setValue(cantActual + Number(item.cantidad)); 
                 dataDiaria[i][2] = cantActual + Number(item.cantidad);
                 found = true; 
                 break;
              }
            }
            if(!found) {
              sheetDiaria.appendRow([
                item.nombre, unidadMedida, item.cantidad, "'" + item.lote, "", "", 
                item.nombre, "Farmacia/Fórmula", fechaEntrega, modalidad
              ]);
            }
        }
      });
    }

    // --- 4. ACTUALIZAR "Resumen_Cronicos" ---
    if (esCronico) {
      const sheetResumen = ss.getSheetByName('Resumen_Cronicos');
      if (sheetResumen) {
        const lastRow = sheetResumen.getLastRow();
        const startRow = lastRow === 0 ? 1 : lastRow;
        
        const arrNaranja = [["", "", "", "", "", "", "", "", ""]];
        sheetResumen.getRange(startRow + 1, 1, 1, 9).setValues(arrNaranja).setBackground('#fce5cd');

        const filasResumen = carrito.map(item => [
          paciente.id,                    
          paciente.apellidos,             
          paciente.nombres,               
          paciente.diagnosticoAs,         
          paciente.diagnosticoNx,         
          fechaEntrega,                   
          item.nombre,                    
          item.cantidad,                  
          ""                              
        ]);

        sheetResumen.getRange(startRow + 2, 1, filasResumen.length, 9)
          .setValues(filasResumen)
          .setBackground(null)
          .setFontWeight("normal");
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: "¡Consumo registrado exitosamente!" };
  } catch (error) {
    return { success: false, message: "Error al guardar: " + error.toString() };
  }
}

function ingresarInventario(datosIngreso) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetInv = ss.getSheetByName('Inventario');
    
    if(!sheetInv) return { success: false, message: "No se encontró la hoja 'Inventario'." };
    
    const invData = sheetInv.getDataRange().getValues();
    let actualizados = 0;
    let nuevos = 0;

    datosIngreso.forEach(item => {
      const iNombre = normalizarTexto(item.nombre);
      let found = false;

      for(let i=1; i<invData.length; i++) {
        const nombreSAP = normalizarTexto(invData[i][2]);
        const nombreReal = normalizarTexto(invData[i][11]); 
        const loteInv = invData[i][1];

        if((nombreSAP === iNombre || nombreReal === iNombre) && isLoteEqual(loteInv, item.lote)) {
          let cantActual = Number(invData[i][3]) || 0;
          let cantNueva = Number(item.cantidad) || 0;
          sheetInv.getRange(i+1, 4).setValue(cantActual + cantNueva);
          invData[i][3] = cantActual + cantNueva; 
          actualizados++;
          found = true;
          break; 
        }
      }

      if(!found) {
        // Soporte para Columna L (DescReal) y Columna M (Familia de producto)
        sheetInv.appendRow([
          item.codigo,
          "'" + item.lote, 
          item.nombre,
          item.cantidad,
          item.unidad,
          item.vencimiento,
          item.estado,
          "", "", "", "", "", "" // H, I, J, K, L, M quedan limpios
        ]);
        nuevos++;
      }
    });

    SpreadsheetApp.flush();
    return { success: true, message: `Se actualizaron ${actualizados} lotes y se crearon ${nuevos} nuevos.` };
  } catch (error) {
    return { success: false, message: "Error: " + error.toString() };
  }
}
