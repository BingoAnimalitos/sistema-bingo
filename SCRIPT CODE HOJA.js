// ==========================================
// BACKEND COMPLETO Y UNIFICADO - BINGO ANIMALITOS (codigo.gs)
// ==========================================

const HOJA_TICKETS = "Tickets";
const HOJA_SORTEOS = "RESULTADOS";
const HOJA_CONFIG = "CONFIGURACION";
const HOJA_CONTROL = "CONTROL_JUEGO"; 
const HOJA_VENDEDORES = "VENDEDORES";
const HOJA_HISTORIAL = "HISTORIAL_SORTEOS";
const HOJA_PAGOS_MOVIL = "PagosMovil";
const HOJA_ADMINISTRADORES = "ADMINISTRADORES";
const HOJA_HISTORIAL_CAJAS = "HISTORIAL_CAJAS"; // <-- Pestaña agregada
const PRECIO_TICKET_DEFECTO = 100;
const DOMINIO_PERMITIDO = "https://bingoanimalitos.github.io";

function validarOrigen(e) {
  try {
    if (!e || !e.parameter) return true;
    return true; 
  } catch (err) {
    return false;
  }
}

function doPost(e) {
  try {
    // --- CAPA DE SEGURIDAD POR ORIGEN ---
    if (!validarOrigen(e)) {
      return responderJSON({ exito: false, mensaje: "Acceso no autorizado desde este origen." });
    }

    const data = JSON.parse(e.postData.contents);
    const accion = data.accion;

    // --- ACCIÓN DE CONSULTA PÚBLICA DE TICKETS ---
    if (accion === "consultarTicketPublico") return responderJSON(consultarTicketPublico(data));

    // --- ACCIONES DE ADMINISTRADORES ---
    if (accion === "loginAdmin") return responderJSON(loginAdminSeguro(data));
    if (accion === "crearAdmin") return responderJSON(crearAdminSeguro(data));
    if (accion === "obtenerAdministradores") return responderJSON(obtenerAdministradoresAdmin());
    // -----------------------------------

    if (accion === "iniciarSesion") return responderJSON(iniciarSesion(data));
    if (accion === "obtenerEstadisticas") return responderJSON(obtenerEstadisticasVendedor(data.idVendedor));
    if (accion === "vender") return responderJSON(procesarVenta(data));
    if (accion === "guardarResultadoManual") return responderJSON(guardarResultadoManual(data));
    if (accion === "obtenerEstadoJuego") return responderJSON(obtenerEstadoJuego());
    if (accion === "obtenerResultados") return responderJSON(obtenerResultados());
    if (accion === "corregirResultado") return responderJSON(corregirResultado(data));
    if (accion === "corregirSorteo") return responderJSON(corregirResultado(data));
    if (accion === "reiniciarSorteo") return responderJSON(reiniciarSorteo(data));
    if (accion === "obtenerTickets") return responderJSON(obtenerTicketsAdmin());
    if (accion === "obtenerVendedores") return responderJSON(obtenerVendedoresAdmin());
    if (accion === "crearVendedor") return responderJSON(crearVendedorAdmin(data));
    if (accion === "actualizarEstadoVendedor") return responderJSON(actualizarEstadoVendedorAdmin(data));
    if (accion === "actualizarComisionVendedor") return responderJSON(actualizarComisionVendedorAdmin(data));
    if (accion === "actualizarParametros") return responderJSON(actualizarParametrosAdmin(data));
    if (accion === "obtenerHistorialVentasFechas") return responderJSON(obtenerHistorialVentasFechas(data));
    if (accion === "obtenerMetricasGlobales") return responderJSON(obtenerMetricasGlobalesAdmin());
    if (accion === "obtenerParametros") return responderJSON(obtenerParametrosAdminPersonalizados());
    if (accion === "obtenerHistorialSorteos") return responderJSON(obtenerHistorialSorteos());
    if (accion === "registrarPagoMovil") return responderJSON(registrarPagoMovilEnHoja(data));
    if (accion === "obtenerPagosMoviles") return responderJSON(obtenerPagosMovilesAdmin());
    if (accion === "aprobarPagoMovil") return responderJSON(aprobarPagoMovilAdmin(data));

    // --- ACCIÓN: REGISTRAR CUADRE DE CAJA ---
    if (accion === "registrarCuadre") {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var hojaCuadres = ss.getSheetByName(HOJA_HISTORIAL_CAJAS);
        
        if (!hojaCuadres) {
          hojaCuadres = ss.insertSheet(HOJA_HISTORIAL_CAJAS);
          // Creamos las cabeceras incluyendo "Referencia" antes de "Estatus"
          hojaCuadres.appendRow(["Fecha_Hora", "ID_Vendedor", "Nombre", "Total_Vendido", "Comision", "Neto_Entregado", "Referencia", "Estatus"]);
        }

        var fechaHora = new Date().toLocaleString();
        var idVendedor = data.idVendedor || "";
        var nombre = data.nombre || "";
        var totalVendido = parseFloat(data.totalVendido) || 0;
        var comision = parseFloat(data.comision) || 0;
        var netoEntregado = parseFloat(data.netoEntregado) || 0;
        var referencia = data.referencia || data.ref || "N/A"; // Captura de la referencia de pago/entrega
        var estatus = "Pendiente";

        hojaCuadres.appendRow([
            fechaHora,
            idVendedor,
            nombre,
            totalVendido,
            comision,
            netoEntregado,
            referencia,
            estatus
        ]);

        return responderJSON({
            exito: true,
            mensaje: "Cuadre registrado con éxito en el historial."
        });
    }

    return responderJSON({ exito: false, mensaje: "Acción no válida" });
  } catch (err) {
    return responderJSON({ exito: false, mensaje: "Error: " + err.toString() });
  }
}

function responderJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// MÓDULO DE CONSULTA PÚBLICA DE TICKET (OPTIMIZADO Y MULTIPLE)
// ==========================================
function consultarTicketPublico(data) {
  try {
    let ticketsParam = data.tickets || data.ticket || data.idTicket || "";
    let idsArray = [];
    
    if (Array.isArray(ticketsParam)) {
      idsArray = ticketsParam.map(s => String(s).trim()).filter(Boolean);
    } else if (typeof ticketsParam === 'string') {
      idsArray = ticketsParam.split(',').map(s => s.trim()).filter(Boolean);
    }

    if (idsArray.length === 0) {
      return { exito: false, mensaje: "Debe proporcionar al menos un ID de ticket válido." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetTickets = ss.getSheetByName(HOJA_TICKETS);
    if (!sheetTickets) {
      return { exito: false, mensaje: "La hoja de Tickets no existe." };
    }

    const datosTickets = sheetTickets.getDataRange().getDisplayValues();
    let ticketsEncontrados = [];

    for (let i = 1; i < datosTickets.length; i++) {
      let idFila = String(datosTickets[i][0] || "").trim();
      let coincide = idsArray.some(idBuscado => idFila.toLowerCase() === idBuscado.toLowerCase());
      
      if (coincide) {
        let estadoTicket = String(datosTickets[i][5] || "ACTIVO").trim();
        let modalidadTicket = String(datosTickets[i][7] || "");
        let montoTicket = parseFloat(datosTickets[i][8]) || PRECIO_TICKET_DEFECTO;
        
        let premioAsignado = 0;
        
        const sheetSorteos = ss.getSheetByName(HOJA_SORTEOS);
        if (sheetSorteos) {
          let datosSorteos = sheetSorteos.getDataRange().getDisplayValues();
          for (let s = 1; s < datosSorteos.length; s++) {
            if (String(datosSorteos[s][4]).trim() === idFila) {
              premioAsignado = parseFloat(datosSorteos[s][5]) || 0;
              break;
            }
          }
        }

        if (premioAsignado === 0) {
          const sheetPagos = ss.getSheetByName(HOJA_PAGOS_MOVIL);
          if (sheetPagos && sheetPagos.getLastRow() > 1) {
            let datosPagos = sheetPagos.getDataRange().getDisplayValues();
            for (let p = 1; p < datosPagos.length; p++) {
              if (String(datosPagos[p][2]).trim() === idFila) {
                premioAsignado = parseFloat(datosPagos[p][6]) || 0;
                break;
              }
            }
          }
        }

        if (estadoTicket.includes("GANADOR") && premioAsignado === 0) {
          const sheetControl = ss.getSheetByName(HOJA_CONTROL);
          if (sheetControl && sheetControl.getLastRow() > 1) {
            premioAsignado = parseFloat(sheetControl.getRange(2, 4).getValue()) || 0;
          }
        }

        ticketsEncontrados.push({
          idTicket: idFila,
          fecha: String(datosTickets[i][1] || ""),
          vendedor: String(datosTickets[i][2] || ""),
          loteria: String(datosTickets[i][3] || "Lotto Activo y La Granjita"),
          numeros: String(datosTickets[i][4] || "").split(", "),
          estado: estadoTicket,
          sorteo: String(datosTickets[i][6] || ""),
          modalidad: modalidadTicket,
          monto: montoTicket,
          premio: premioAsignado
        });
      }
    }

    if (ticketsEncontrados.length > 0) {
      const animalesSorteo = obtenerAnimalitosSorteoActual();
      const estadoJuego = obtenerEstadoJuego();

      return {
        exito: true,
        tickets: ticketsEncontrados,
        animalesSorteo: animalesSorteo,
        poteActual: estadoJuego.pote,
        nroSorteo: estadoJuego.nroSorteo
      };
    } else {
      return { exito: false, mensaje: "Los tickets consultados no existen o expiraron." };
    }
  } catch (err) {
    return { exito: false, mensaje: "Error al consultar el ticket: " + err.toString() };
  }
}

// ==========================================
// MÓDULO DE GESTIÓN Y LOGIN DE ADMINISTRADORES
// ==========================================
function loginAdminSeguro(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(HOJA_ADMINISTRADORES);
    
    if (!hoja) {
      if (data.usuario === "admin" && data.password === "admin2026") {
        return { exito: true, mensaje: "Acceso concedido (Emergencia)" };
      }
      return { exito: false, mensaje: "La hoja ADMINISTRADORES no existe en el Google Sheets." };
    }

    const datos = hoja.getDataRange().getDisplayValues();
    let accesoConcedido = false;

    for (let i = 1; i < datos.length; i++) {
      let idAdmin = String(datos[i][0] || "").trim();
      let usuarioSheet = String(datos[i][2] || "").trim();
      let claveSheet = String(datos[i][3] || "").trim();
      let estadoSheet = String(datos[i][4] || "ACTIVO").trim();

      if (usuarioSheet === String(data.usuario).trim() && 
          claveSheet === String(data.password).trim() && 
          estadoSheet === "ACTIVO" && idAdmin) {
        accesoConcedido = true;
        break;
      }
    }

    if (!accesoConcedido && data.usuario === "admin" && data.password === "admin2026") {
      accesoConcedido = true;
    }

    if (accesoConcedido) {
      return { exito: true, mensaje: "Acceso concedido" };
    } else {
      return { exito: false, mensaje: "Usuario o contraseña de administrador incorrectos." };
    }
  } catch (err) {
    return { exito: false, mensaje: "Error en login de admin: " + err.toString() };
  }
}

function crearAdminSeguro(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(HOJA_ADMINISTRADORES);
    
    if (!hoja) {
      hoja = ss.insertSheet(HOJA_ADMINISTRADORES);
      hoja.appendRow(["ID_Admin", "Nombre", "Usuario", "Clave", "Estado"]);
    }

    const datos = hoja.getDataRange().getDisplayValues();
    let idNuevo = "ADM-" + Math.floor(1000 + Math.random() * 9000);

    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][2]).trim() === String(data.usuario).trim()) {
        return { exito: false, mensaje: "El nombre de usuario ya está registrado." };
      }
    }

    hoja.appendRow([
      idNuevo,
      data.nombre || "Admin",
      data.usuario,
      data.clave,
      "ACTIVO"
    ]);

    return { exito: true, mensaje: "Administrador creado exitosamente." };
  } catch (err) {
    return { exito: false, mensaje: "Error al crear admin: " + err.toString() };
  }
}

function obtenerAdministradoresAdmin() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(HOJA_ADMINISTRADORES);
    
    if (!hoja) {
      hoja = ss.insertSheet(HOJA_ADMINISTRADORES);
      hoja.appendRow(["ID_Admin", "Nombre", "Usuario", "Clave", "Estado"]);
      hoja.appendRow(["ADM-1001", "Administrador Principal", "admin", "admin2026", "ACTIVO"]);
    }

    const datos = hoja.getDataRange().getDisplayValues();
    let administradores = [];

    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0]) {
        administradores.push({
          idAdmin: datos[i][0],
          nombre: datos[i][1],
          usuario: datos[i][2],
          estado: datos[i][4] || "ACTIVO"
        });
      }
    }
    return { exito: true, administradores: administradores };
  } catch (err) {
    return { exito: false, mensaje: err.toString() };
  }
}

function obtenerComisionAdmin() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetConfig = ss.getSheetByName(HOJA_CONFIG);
    if (sheetConfig) {
      let valDecimal = parseFloat(sheetConfig.getRange("B3").getValue());
      if (!isNaN(valDecimal) && valDecimal > 0) {
        return valDecimal > 1 ? valDecimal / 100 : valDecimal;
      }
    }
    return 0.10; 
  } catch (err) {
    return 0.10;
  }
}

function iniciarSesion(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const datos = ss.getSheetByName(HOJA_VENDEDORES).getDataRange().getDisplayValues();
  for (let i = 1; i < datos.length; i++) {
    let estadoVend = String(datos[i][7] || "ACTIVO").trim();
    if (estadoVend === "SUSPENDIDO") continue; 

    if (String(datos[i][5]).trim() === data.usuario && String(datos[i][6]).trim() === data.clave) {
      return { exito: true, idVendedor: datos[i][0], nombre: datos[i][1], comision: datos[i][2] };
    }
  }
  return { exito: false, mensaje: "Usuario o contraseña incorrectos, o cuenta suspendida." };
}

function obtenerAnimalitosSorteoActual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSorteos = ss.getSheetByName(HOJA_SORTEOS);
  if (!sheetSorteos) return [];
  
  const datos = sheetSorteos.getDataRange().getDisplayValues();
  let animalitos = [];
  for (let i = 1; i < datos.length; i++) {
    let anim = String(datos[i][3]).trim();
    if (anim && !animalitos.includes(anim)) {
      animalitos.push(anim);
    }
  }
  return animalitos;
}

function obtenerEstadisticasVendedor(idVendedor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTickets = ss.getSheetByName(HOJA_TICKETS);
  const nombreVendedor = obtenerNombreVendedor(idVendedor);
  
  let comisionPorc = 18;
  const datosVend = ss.getSheetByName(HOJA_VENDEDORES).getDataRange().getDisplayValues();
  for(let i = 1; i < datosVend.length; i++){
    if(String(datosVend[i][0]).trim() === String(idVendedor).trim()) { 
      comisionPorc = parseFloat(datosVend[i][2]) || 18; 
      break; 
    }
  }

  const estadoJuego = obtenerEstadoJuego(); 
  let textoSorteoActual = "Sorteo #" + String(estadoJuego.nroSorteo).padStart(2, '0');
  
  let totalVendido = 0;
  let misTickets = [];
  let todosLosTicketsSorteo = [];
  
  if (sheetTickets) {
    const datosTickets = sheetTickets.getDataRange().getDisplayValues();
    for (let i = 1; i < datosTickets.length; i++) {
      let vendedorTicket = String(datosTickets[i][2]).trim();
      let estadoTicket = String(datosTickets[i][5] || "ACTIVO").trim();
      let referenciaTicket = String(datosTickets[i][6] || "").trim();
      let modalidadTicket = String(datosTickets[i][7] || "").trim();
      let montoTicket = parseFloat(datosTickets[i][8]) || PRECIO_TICKET_DEFECTO;

      if (estadoTicket.startsWith("GANADOR - ")) {
        let partes = estadoTicket.split(" - ");
        estadoTicket = "GANADOR";
        if (!modalidadTicket) modalidadTicket = partes[1];
      }

      let modalidadLimpia = modalidadTicket || (estadoTicket === "ACTIVO" ? "Activo" : (estadoTicket === "GANADOR" ? "Ganador" : "No ganó"));

      let objetoTicket = {
        idTicket: String(datosTickets[i][0]),
        fecha: String(datosTickets[i][1]),
        vendedor: vendedorTicket,
        loteria: String(datosTickets[i][3] || "Lotto Activo y La Granjita"),
        estado: estadoTicket,
        modalidad: modalidadLimpia,
        numeros: String(datosTickets[i][4] || "").split(", "),
        sorteo: referenciaTicket,
        premio: (estadoTicket === "GANADOR" && referenciaTicket === textoSorteoActual) ? (estadoJuego.premioPorGanador || 0) : 0
      };

      if (referenciaTicket === textoSorteoActual) {
        todosLosTicketsSorteo.push(objetoTicket);
      }

      if (vendedorTicket.toLowerCase() === String(nombreVendedor).trim().toLowerCase()) {
        misTickets.push(objetoTicket);
        
        if (referenciaTicket === textoSorteoActual && estadoTicket === "ACTIVO") {
          totalVendido += montoTicket;
        }
      }
    }
  }

  let gananciaComision = 0;
  let netoEntregar = 0;

  if (estadoJuego.estado === "ACTIVO") {
    gananciaComision = totalVendido * (comisionPorc / 100);
    netoEntregar = totalVendido - gananciaComision;
    actualizarMetricasVendedorEnHoja(idVendedor, totalVendido, gananciaComision);
  } else {
    for(let i = 1; i < datosVend.length; i++){
      if(String(datosVend[i][0]).trim() === String(idVendedor).trim()) { 
        totalVendido = parseFloat(datosVend[i][3]) || 0;
        gananciaComision = parseFloat(datosVend[i][4]) || 0;
        netoEntregar = totalVendido - gananciaComision;
        break; 
      }
    }
  }

  let animalesSorteoActual = obtenerAnimalitosSorteoActual();

  return {
    exito: true,
    totalVendido: totalVendido,
    comisionPorc: comisionPorc,
    ganancia: gananciaComision,
    netoEntregar: netoEntregar,
    poteActual: estadoJuego.pote,
    nroSorteo: estadoJuego.nroSorteo,
    tickets: misTickets.reverse(),
    ticketsGlobalSorteo: todosLosTicketsSorteo,
    animalesSorteo: animalesSorteoActual
  };
}

function actualizarMetricasVendedorEnHoja(idVendedor, totalVendido, gananciaComision) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetVend = ss.getSheetByName(HOJA_VENDEDORES);
  if (!sheetVend) return;
  const datos = sheetVend.getDataRange().getDisplayValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === String(idVendedor).trim()) {
      sheetVend.getRange(i + 1, 4).setValue(totalVendido);       
      sheetVend.getRange(i + 1, 5).setValue(gananciaComision);   
      break;
    }
  }
}

function procesarVenta(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheetControl = ss.getSheetByName(HOJA_CONTROL);
  let estadoJuegoActual = "ACTIVO";
  let nroSorteoActual = 1;

  if (sheetControl) {
    const datosCtrl = sheetControl.getDataRange().getDisplayValues();
    if (datosCtrl.length > 1) {
      estadoJuegoActual = String(datosCtrl[1][0] || "ACTIVO").trim();
      let valNroSorteo = datosCtrl[1][4];
      nroSorteoActual = (valNroSorteo !== undefined && valNroSorteo !== null && String(valNroSorteo).trim() !== "") 
        ? parseInt(valNroSorteo, 10) 
        : 1;
    }
  }

  if (estadoJuegoActual === "FINALIZADO" || estadoJuegoActual === "FINALIZADO_SIN_GANADOR") {
    return { 
      exito: false, 
      mensaje: "El sorteo ya ha finalizado. No se pueden generar nuevos tickets hasta reiniciar el juego." 
    };
  }

  let precioActualCarton = PRECIO_TICKET_DEFECTO;
  const sheetConfig = ss.getSheetByName(HOJA_CONFIG);
  if (sheetConfig) {
    let valPrecio = parseFloat(sheetConfig.getRange("B2").getValue());
    if (!isNaN(valPrecio) && valPrecio > 0) {
      precioActualCarton = valPrecio;
    }
  }

  const nombreVendedor = data.vendedor || obtenerNombreVendedor(data.idVendedor);
  const cantidadCartones = parseInt(data.cantidadCartones) || 1;

  let comisionVendedor = 18;
  const datosVend = ss.getSheetByName(HOJA_VENDEDORES).getDataRange().getDisplayValues();
  for(let i = 1; i < datosVend.length; i++){
    if(String(datosVend[i][0]).trim() === String(data.idVendedor).trim()) { 
      comisionVendedor = parseFloat(datosVend[i][2]) || 18; 
      break; 
    }
  }

  let sheetTickets = ss.getSheetByName(HOJA_TICKETS);
  if (!sheetTickets) {
    sheetTickets = ss.insertSheet(HOJA_TICKETS);
    sheetTickets.appendRow(["ID Ticket", "Fecha", "ID Vendedor", "Lotería", "Números", "Estado", "Referencia", "Modalidad", "Monto"]);
  }

  const todosAnimalitos = [
    "0 Delfin", "00 Ballena", "1 Carnero", "2 Toro", "3 Ciempies", "4 Alacran",
    "5 Leon", "6 Rana", "7 Perico", "8 Raton", "9 Aguila", "10 Tigre",
    "11 Gato", "12 Caballo", "13 Mono", "14 Paloma", "15 Zorro", "16 Oso",
    "17 Pavo", "18 Burro", "19 Chivo", "20 Cochino", "21 Gallo", "22 Camello",
    "23 Cebra", "24 Iguana", "25 Gallina", "26 Vaca", "27 Perro", "28 Zamuro",
    "29 Elefante", "30 Caiman", "31 Lapa", "32 Ardilla", "33 Pescado", "34 Venado",
    "35 Jirafa", "36 Culebra"
  ];

  const ticketsGenerados = [];
  const fechaStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy hh:mm a");
  const textoSorteo = "Sorteo #" + String(nroSorteoActual).padStart(2, '0');

  for (let i = 0; i < cantidadCartones; i++) {
    const idTicket = "TK-" + Math.floor(100000 + Math.random() * 900000);
    const cartonesMezclados = [...todosAnimalitos].sort(() => 0.5 - Math.random());
    const seleccionados = cartonesMezclados.slice(0, 16);
    const numerosCadena = seleccionados.join(", ");

    sheetTickets.appendRow([idTicket, fechaStr, nombreVendedor, "Lotto Activo y La Granjita", numerosCadena, "ACTIVO", textoSorteo, "", precioActualCarton]);

    registrarPagoMovilEnHoja({
      vendedor: nombreVendedor,
      ticketId: idTicket,
      banco: data.banco || data.bancoCliente || data.bancoOrigen || data.bancoDestino || "",
      telefono: data.telefono || data.telefonoCliente || "",
      cedula: data.cedula || data.cedulaCliente || "",
      montoPremio: 0,
      estatus: "PENDIENTE",
      referencia: data.referencia || data.referenciaPago || data.numeroReferencia || ""
    });

    ticketsGenerados.push({
      idTicket: idTicket,
      fecha: fechaStr,
      vendedor: nombreVendedor,
      loteria: "Lotto Activo y La Granjita",
      modalidad: "Activo",
      numeros: seleccionados,
      sorteo: textoSorteo,
      monto: precioActualCarton
    });
  }

  const poteCalculadoGlobal = recalcularPoteYVentasSorteoActual(textoSorteo, precioActualCarton, comisionVendedor);

  let totalVendidoActualizado = 0;
  const datosT = sheetTickets.getDataRange().getDisplayValues();
  for(let t = 1; t < datosT.length; t++) {
    if(String(datosT[t][2]).trim() === String(nombreVendedor).trim() && String(datosT[t][5]).trim() === "ACTIVO" && String(datosT[t][6]).trim() === textoSorteo) {
      let mFila = parseFloat(datosT[t][8]) || precioActualCarton;
      totalVendidoActualizado += mFila;
    }
  }
  let comisionAcumuladaActualizada = totalVendidoActualizado * (comisionVendedor / 100);
  actualizarMetricasVendedorEnHoja(data.idVendedor, totalVendidoActualizado, comisionAcumuladaActualizada);

  return { 
    exito: true, 
    vendedor: nombreVendedor,
    precioTicket: precioActualCarton,
    tickets: ticketsGenerados, 
    poteActual: poteCalculadoGlobal.poteTotal 
  };
}

function recalcularPoteYVentasSorteoActual(textoSorteo, precioCartonDefecto, comisionVendedorPorc) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTickets = ss.getSheetByName(HOJA_TICKETS);
  const sheetControl = ss.getSheetByName(HOJA_CONTROL);
  
  if (!sheetTickets) return { poteTotal: 0 };

  const comisionAdminPorc = obtenerComisionAdmin();
  const datosTickets = sheetTickets.getDataRange().getDisplayValues();
  
  let totalVendidoSorteo = 0;

  for (let t = 1; t < datosTickets.length; t++) {
    let referenciaTicket = String(datosTickets[t][6] || "").trim();
    let estadoTicket = String(datosTickets[t][5] || "ACTIVO").trim();

    if (referenciaTicket === textoSorteo && estadoTicket === "ACTIVO") {
      let mFila = parseFloat(datosTickets[t][8]) || precioCartonDefecto;
      totalVendidoSorteo += mFila;
    }
  }

  const montoComisionVendedorTotal = totalVendidoSorteo * (comisionVendedorPorc / 100);
  const montoComisionAdminTotal = totalVendidoSorteo * comisionAdminPorc;
  const potePremioTotal = totalVendidoSorteo - montoComisionVendedorTotal - montoComisionAdminTotal;

  if (sheetControl && sheetControl.getLastRow() > 1) {
    sheetControl.getRange(2, 2).setValue(potePremioTotal);
  }

  return { poteTotal: potePremioTotal };
}

function registrarPagoMovilEnHoja(datos) {
  if (!datos) {
    return { exito: false, mensaje: "No se recibieron datos para registrar el pago." };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(HOJA_PAGOS_MOVIL);
    
    if (!hoja) {
      hoja = ss.insertSheet(HOJA_PAGOS_MOVIL);
      hoja.appendRow([
        "Fecha/Hora", 
        "Vendedor", 
        "Ticket ID", 
        "Banco", 
        "Teléfono", 
        "Cédula", 
        "Monto Premio", 
        "Estatus",
        "Referencia"
      ]);
    }

    const fechaHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy hh:mm a");
    
    let vendedorNombre = datos.vendedor || datos.nombreVendedor || "";
    if (!vendedorNombre && datos.idVendedor) {
      vendedorNombre = obtenerNombreVendedor(datos.idVendedor);
    }
    if (!vendedorNombre) vendedorNombre = "SIN VENDEDOR";

    const ticketId = datos.ticketId || datos.idTicket || "";
    const banco = datos.banco || datos.bancoCliente || datos.bancoOrigen || datos.bancoDestino || "";
    const telefono = datos.telefono || datos.telefonoCliente || "";
    const cedula = datos.cedula || datos.cedulaCliente || "";
    const montoPremio = parseFloat(datos.montoPremio || datos.monto) || 0;
    const estatus = datos.estatus || "PENDIENTE";
    const referencia = datos.referencia || datos.numeroReferencia || datos.ref || datos.referenciaPago || "";

    hoja.appendRow([
      fechaHora,
      vendedorNombre,
      ticketId,
      banco,
      telefono,
      cedula,
      montoPremio,
      estatus,
      referencia
    ]);

    return { exito: true, mensaje: "Pago Móvil registrado exitosamente." };
  } catch (err) {
    return { exito: false, mensaje: "Error al guardar el pago: " + err.message };
  }
}

function registrarPagoMovilSheet(datos) {
  return registrarPagoMovilEnHoja(datos);
}

function obtenerPagosMovilesAdmin() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(HOJA_PAGOS_MOVIL);
    if (!hoja) return { exito: true, pagos: [] };

    const datos = hoja.getDataRange().getDisplayValues();
    let pagos = [];

    for (let i = 1; i < datos.length; i++) {
      if (datos[i][2]) { 
        pagos.push({
          id: datos[i][2],                                    
          fecha: datos[i][0],                                 
          usuario: datos[i][1] + " (CI: " + datos[i][5] + ")", 
          banco: datos[i][3],                                 
          telefono: datos[i][4],                              
          monto: parseFloat(datos[i][6]) || 0,                
          estatus: datos[i][7] || "PENDIENTE",                
          referencia: datos[i][8] || ""                       
        });
      }
    }
    return { exito: true, pagos: pagos.reverse() };
  } catch (err) {
    return { exito: false, mensaje: err.toString() };
  }
}

function aprobarPagoMovilAdmin(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(HOJA_PAGOS_MOVIL);
    if (!hoja) return { exito: false, mensaje: "Hoja no encontrada." };

    const datos = hoja.getDataRange().getDisplayValues();
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][2]).trim() === String(data.idPago || data.ticketId).trim()) {
        hoja.getRange(i + 1, 8).setValue("PAGADO");           
        hoja.getRange(i + 1, 9).setValue(data.referencia);    
        return { exito: true, mensaje: "Pago aprobado con éxito." };
      }
    }
    return { exito: false, mensaje: "Ticket no encontrado." };
  } catch (err) {
    return { exito: false, mensaje: err.toString() };
  }
}

function obtenerNombreVendedor(idOTexto) {
  if (!idOTexto) return "VENDEDOR 01";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetVendedores = ss.getSheetByName(HOJA_VENDEDORES);
  if (!sheetVendedores) return idOTexto;

  const datos = sheetVendedores.getDataRange().getDisplayValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0].toString().trim() === idOTexto.toString().trim() || datos[i][1].toString().trim().toLowerCase() === idOTexto.toString().trim().toLowerCase()) {
      return datos[i][1];
    }
  }
  return idOTexto;
}

function evaluarModalidadesCarton(numerosTicketArray, cantadosGlobales) {
  let matriz = [];
  for (let r = 0; r < 4; r++) {
    matriz.push(numerosTicketArray.slice(r * 4, (r + 1) * 4));
  }

  let aciertosMatriz = [];
  for (let r = 0; r < 4; r++) {
    aciertosMatriz[r] = [];
    for (let c = 0; c < 4; c++) {
      aciertosMatriz[r][c] = cantadosGlobales.includes(matriz[r][c]);
    }
  }

  let modalidadesCumplidas = [];

  for (let r = 0; r < 4; r++) {
    if (aciertosMatriz[r][0] && aciertosMatriz[r][1] && aciertosMatriz[r][2] && aciertosMatriz[r][3]) {
      modalidadesCumplidas.push("Horizontal");
      break;
    }
  }

  for (let c = 0; c < 4; c++) {
    if (aciertosMatriz[0][c] && aciertosMatriz[1][c] && aciertosMatriz[2][c] && aciertosMatriz[3][c]) {
      modalidadesCumplidas.push("Columna");
      break;
    }
  }

  if (aciertosMatriz[0][0] && aciertosMatriz[0][3] && aciertosMatriz[3][0] && aciertosMatriz[3][3]) {
    modalidadesCumplidas.push("4 esquinas");
  }

  let diagonalPrincipal = aciertosMatriz[0][0] && aciertosMatriz[1][1] && aciertosMatriz[2][2] && aciertosMatriz[3][3];
  let diagonalSecundaria = aciertosMatriz[0][3] && aciertosMatriz[1][2] && aciertosMatriz[2][1] && aciertosMatriz[3][0];
  if (diagonalPrincipal || diagonalSecundaria) {
    modalidadesCumplidas.push("Línea diagonal");
  }

  return modalidadesCumplidas;
}

function guardarResultadoManual(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetSorteos = ss.getSheetByName(HOJA_SORTEOS);
  if (!sheetSorteos) {
    sheetSorteos = ss.insertSheet(HOJA_SORTEOS);
    sheetSorteos.appendRow(["Fecha", "Lotería", "Hora_Sorteo", "Animalito", "ID_Ticket", "Premio"]);
  }
  
  const loteria = data.loteria || "Lotto Activo";
  const horaSorteo = data.horaSorteo || "09:00 AM";
  const cantados = data.cantados || [];
  const fechaStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  const todosResultadosRows = sheetSorteos.getDataRange().getDisplayValues();
  let registrosMismaHora = [];
  let loteriasEnEstaHora = [];
  let totalHorariosUnicosDelDia = new Set();

  for (let i = 1; i < todosResultadosRows.length; i++) {
    let fRow = String(todosResultadosRows[i][0]).trim();
    let lRow = String(todosResultadosRows[i][1]).trim();
    let hRow = String(todosResultadosRows[i][2]).trim();

    if (fRow === fechaStr) {
      totalHorariosUnicosDelDia.add(hRow.toLowerCase());
      if (hRow.toLowerCase() === horaSorteo.toLowerCase()) {
        loteriasEnEstaHora.push(lRow.toLowerCase());
        registrosMismaHora.push({ loteria: lRow });
      }
    }
  }

  if (loteriasEnEstaHora.includes(loteria.toLowerCase())) {
    return { exito: false, mensaje: "Esta lotería ya tiene un resultado registrado para la hora " + horaSorteo };
  }

  if (registrosMismaHora.length >= 2) {
    return { exito: false, mensaje: "Ya se registraron los dos sorteos (Lotto Activo y La Granjita) para este horario." };
  }

  cantados.forEach(animal => {
    sheetSorteos.appendRow([fechaStr, loteria, horaSorteo, animal, "", ""]);
  });

  registrosMismaHora.push({ loteria: loteria });
  if (!totalHorariosUnicosDelDia.has(horaSorteo.toLowerCase()) && registrosMismaHora.length === 2) {
    totalHorariosUnicosDelDia.add(horaSorteo.toLowerCase());
  }

  let ganadoresEncontrados = [];
  let premioPorGanador = 0;

  if (registrosMismaHora.length === 2) {
    let todosAnimalitosCantadosGlobal = [];
    let refreshRows = sheetSorteos.getDataRange().getDisplayValues();
    for (let i = 1; i < refreshRows.length; i++) {
      let anim = String(refreshRows[i][3]).trim();
      if (anim && !todosAnimalitosCantadosGlobal.includes(anim)) {
        todosAnimalitosCantadosGlobal.push(anim);
      }
    }

    let sheetControl = ss.getSheetByName(HOJA_CONTROL);
    let nroSorteoActual = 1;
    if (sheetControl) {
      let datosCtrl = sheetControl.getDataRange().getDisplayValues();
      if (datosCtrl.length > 1) {
        let valNroSorteo = datosCtrl[1][4];
        nroSorteoActual = (valNroSorteo !== undefined && valNroSorteo !== null && String(valNroSorteo).trim() !== "") 
          ? parseInt(valNroSorteo, 10) 
          : 1;
      }
    }
    let textoSorteoActual = "Sorteo #" + String(nroSorteoActual).padStart(2, '0');

    let sheetTickets = ss.getSheetByName(HOJA_TICKETS);
    if (sheetTickets) {
      let datosTickets = sheetTickets.getDataRange().getDisplayValues();
      let cantGanadores = 0;
      let ticketsValidosParaGanar = [];
      let filasParaActualizarEstado = [];

      for (let i = 1; i < datosTickets.length; i++) {
        let idTicket = String(datosTickets[i][0]);
        let estadoTicket = String(datosTickets[i][5] || "ACTIVO").trim();
        let referenciaTicket = String(datosTickets[i][6] || "").trim();

        if (referenciaTicket !== textoSorteoActual || estadoTicket !== "ACTIVO") continue;

        let numerosTicket = String(datosTickets[i][4]).split(",").map(n => n.trim());
        let modalidades = evaluarModalidadesCarton(numerosTicket, todosAnimalitosCantadosGlobal);

        filasParaActualizarEstado.push({ fila: i + 1, modalidades: modalidades });

        if (modalidades.length > 0) {
          cantGanadores++;
          ticketsValidosParaGanar.push({
            row: i + 1,
            idTicket: idTicket,
            vendedor: String(datosTickets[i][2]),
            modalidadTexto: modalidades.join(", ")
          });
        }
      }

      if (cantGanadores > 0) {
        let poteActual = 0;
        if (sheetControl) {
          poteActual = parseFloat(sheetControl.getRange(2, 2).getValue()) || 0;
          sheetControl.getRange(2, 1).setValue("FINALIZADO");
          sheetControl.getRange(2, 3).setValue(cantGanadores);
          premioPorGanador = poteActual / cantGanadores;
          sheetControl.getRange(2, 4).setValue(premioPorGanador);
        } else {
          premioPorGanador = poteActual / cantGanadores;
        }

        filasParaActualizarEstado.forEach(item => {
          let ganadorEncontrado = ticketsValidosParaGanar.find(g => g.row === item.fila);
          if (ganadorEncontrado) {
            sheetTickets.getRange(item.fila, 6).setValue("GANADOR");
            sheetTickets.getRange(item.fila, 8).setValue(ganadorEncontrado.modalidadTexto);
          } else {
            sheetTickets.getRange(item.fila, 6).setValue("NO GANÓ");
            sheetTickets.getRange(item.fila, 8).setValue("");
          }
        });

        let matchingRows = [];
        for (let r = 1; r <= sheetSorteos.getLastRow(); r++) {
          let rowHora = String(sheetSorteos.getRange(r, 3).getValue()).trim();
          if (rowHora.toLowerCase() === horaSorteo.toLowerCase()) {
            matchingRows.push(r);
          }
        }

        matchingRows.forEach((r, index) => {
          let winnerIndex = Math.min(Math.floor(index * ticketsValidosParaGanar.length / matchingRows.length), ticketsValidosParaGanar.length - 1);
          let g = ticketsValidosParaGanar[winnerIndex];
          sheetSorteos.getRange(r, 5).setValue(g.idTicket);
          sheetSorteos.getRange(r, 6).setValue(premioPorGanador);
        });

        let hojaPagos = ss.getSheetByName(HOJA_PAGOS_MOVIL);
        if (hojaPagos && hojaPagos.getLastRow() > 1) {
          let datosPagos = hojaPagos.getDataRange().getDisplayValues();
          for (let p = 1; p < datosPagos.length; p++) {
            let idTicketEnPagos = String(datosPagos[p][2]).trim(); 
            let gMatch = ticketsValidosParaGanar.find(function(g) { return g.idTicket === idTicketEnPagos; });
            if (gMatch) {
              hojaPagos.getRange(p + 1, 7).setValue(premioPorGanador); 
              hojaPagos.getRange(p + 1, 8).setValue("GANADOR - POR PAGAR"); 
            }
          }
        }

        ticketsValidosParaGanar.forEach(g => {
          ganadoresEncontrados.push({
            idTicket: g.idTicket,
            vendedor: g.vendedor,
            modalidad: `(4 aciertos modalidad: ${g.modalidadTexto})`
          });
        });
      } 
      else {
        if (totalHorariosUnicosDelDia.size >= 11) {
          filasParaActualizarEstado.forEach(item => {
            sheetTickets.getRange(item.fila, 6).setValue("NO GANÓ");
            sheetTickets.getRange(item.fila, 8).setValue("");
          });
          if (sheetControl) {
            sheetControl.getRange(2, 1).setValue("FINALIZADO_SIN_GANADOR");
          }
        }
      }
    }
  }

  return { 
    exito: true, 
    loteria: loteria, 
    segundoSorteoCompletado: (registrosMismaHora.length === 2),
    ganadores: ganadoresEncontrados, 
    premioPorGanador: premioPorGanador 
  };
}

function obtenerEstadoJuego() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetControl = ss.getSheetByName(HOJA_CONTROL);
  let pote = 0;
  let estado = "ACTIVO";
  let nroSorteo = 1;
  let premioPorGanador = 0;

  if (!sheetControl) {
    sheetControl = ss.insertSheet(HOJA_CONTROL);
    sheetControl.appendRow(["Estado del Juego", "Pote Acumulado (Bs)", "Cantidad Ganadores", "Premio por Ganador", "Numero Sorteo"]);
    sheetControl.appendRow(["ACTIVO", 0, 0, 0, 1]);
  } else {
    const datos = sheetControl.getDataRange().getDisplayValues();
    if (datos.length > 1) {
      estado = String(datos[1][0] || "ACTIVO").trim();
      pote = parseFloat(datos[1][1]) || 0;
      premioPorGanador = parseFloat(datos[1][3]) || 0;
      let valNroSorteo = datos[1][4];
      nroSorteo = (valNroSorteo !== undefined && valNroSorteo !== null && String(valNroSorteo).trim() !== "") 
        ? parseInt(valNroSorteo, 10) 
        : 1;
    }
  }
  return { exito: true, estado: estado, pote: pote, nroSorteo: nroSorteo, premioPorGanador: premioPorGanador };
}

function obtenerResultados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSorteos = ss.getSheetByName(HOJA_SORTEOS);
  if (!sheetSorteos) return { exito: true, resultados: [] };

  const datos = sheetSorteos.getDataRange().getDisplayValues();
  const resultados = [];
  for (let i = 1; i < datos.length; i++) {
    resultados.push({
      fecha: String(datos[i][0]),
      loteria: String(datos[i][1]),
      horaSorteo: String(datos[i][2]),
      animalito: String(datos[i][3]),
      idTicket: String(datos[i][4] || ""),
      premio: String(datos[i][5] || "")
    });
  }
  return { exito: true, resultados: resultados.reverse() };
}

function corregirResultado(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSorteos = ss.getSheetByName(HOJA_SORTEOS);
  if (!sheetSorteos) return { exito: false, mensaje: "Hoja no encontrada." };

  const loteria = (data.loteria || "").toLowerCase();
  const horaSorteo = (data.horaSorteo || "").toLowerCase();
  const datos = sheetSorteos.getDataRange().getDisplayValues();
  let eliminados = 0;

  for (let i = datos.length - 1; i >= 1; i--) {
    if (String(datos[i][1]).toLowerCase() === loteria && String(datos[i][2]).toLowerCase() === horaSorteo) {
      sheetSorteos.deleteRow(i + 1);
      eliminados++;
    }
  }
  return { exito: true, mensaje: `Se eliminaron ${eliminados} registros.` };
}

function reiniciarSorteo(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetControl = ss.getSheetByName(HOJA_CONTROL);
  const nuevoPote = parseFloat(data.nuevoPote) || 0;

  let nuevoSorteoNum = 1;
  let sorteoActualEnCelda = 1;
  
  if (sheetControl) {
    const datosCtrl = sheetControl.getDataRange().getDisplayValues();
    if (datosCtrl.length > 1) {
      let valNroSorteo = datosCtrl[1][4];
      sorteoActualEnCelda = (valNroSorteo !== undefined && valNroSorteo !== null && String(valNroSorteo).trim() !== "") 
        ? parseInt(valNroSorteo, 10) 
        : 1;
      nuevoSorteoNum = sorteoActualEnCelda + 1; 
    }
    sheetControl.getRange(2, 1).setValue("ACTIVO");
    sheetControl.getRange(2, 2).setValue(nuevoPote);
    sheetControl.getRange(2, 3).setValue(0);
    sheetControl.getRange(2, 4).setValue(0);
    sheetControl.getRange(2, 5).setValue(nuevoSorteoNum); 
  }

  let sheetSorteos = ss.getSheetByName(HOJA_SORTEOS);
  if (sheetSorteos && sheetSorteos.getLastRow() > 1) {
    let sheetHistorial = ss.getSheetByName(HOJA_HISTORIAL);
    if (!sheetHistorial) {
      sheetHistorial = ss.insertSheet(HOJA_HISTORIAL);
    }
    
    if (sheetHistorial.getLastRow() === 0) {
      sheetHistorial.appendRow(["Fecha", "Lotería", "Hora_Sorteo", "Animalito", "ID_Ticket", "Premio", "Nro_Sorteo"]);
    }

    let resRows = sheetSorteos.getDataRange().getDisplayValues();

    for (let r = 1; r < resRows.length; r++) {
      let filaRes = resRows[r];
      if (filaRes[0] || filaRes[1] || filaRes[3]) {
        sheetHistorial.appendRow([
          String(filaRes[0] || ""),      
          String(filaRes[1] || ""),      
          String(filaRes[2] || ""),      
          String(filaRes[3] || ""),      
          String(filaRes[4] || ""),      
          String(filaRes[5] || ""),      
          sorteoActualEnCelda            
        ]);
      }
    }

    const ultimaFila = sheetSorteos.getLastRow();
    if (ultimaFila > 1) {
      sheetSorteos.getRange(2, 1, ultimaFila - 1, sheetSorteos.getLastColumn()).clearContent();
    }
  }

  let hojaPagosMovil = ss.getSheetByName(HOJA_PAGOS_MOVIL);
  if (hojaPagosMovil && hojaPagosMovil.getLastRow() > 1) {
    hojaPagosMovil.getRange(2, 1, hojaPagosMovil.getLastRow() - 1, hojaPagosMovil.getLastColumn()).clearContent();
  }

  let sheetVend = ss.getSheetByName(HOJA_VENDEDORES);
  if (sheetVend) {
    let datosVend = sheetVend.getDataRange().getDisplayValues();
    for (let v = 1; v < datosVend.length; v++)  {
      sheetVend.getRange(v + 1, 4).setValue(0); 
      sheetVend.getRange(v + 1, 5).setValue(0); 
    }
  }

  return { exito: true, nuevoSorteo: nuevoSorteoNum, mensaje: "Juego reiniciado con éxito. Sorteo Nro " + nuevoSorteoNum };
}

function obtenerHistorialSorteos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetHistorial = ss.getSheetByName(HOJA_HISTORIAL);
    if (!sheetHistorial) return { exito: true, historial: [] };

    const datos = sheetHistorial.getDataRange().getDisplayValues();
    let historial = [];

    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0]) {
        historial.push({
          fecha: String(datos[i][0]),
          loteria: String(datos[i][1]),
          horaSorteo: String(datos[i][2]),
          animalito: String(datos[i][3]),
          idTicket: String(datos[i][4] || ""),
          premio: String(datos[i][5] || ""),
          nroSorteo: String(datos[i][6] || "") 
        });
      }
    }

    return { exito: true, historial: historial.reverse() };
  } catch (err) {
    return { exito: false, mensaje: err.toString() };
  }
}

// ==========================================
// FUNCIÓN PARA GESTIONAR HISTORIAL_CAJAS (Panel de Administración)
// ==========================================
function obtenerHistorialCajasAdmin() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hojaCajas = ss.getSheetByName(HOJA_HISTORIAL_CAJAS);
    
    if (!hojaCajas) {
      hojaCajas = ss.insertSheet(HOJA_HISTORIAL_CAJAS);
      hojaCajas.appendRow(["Fecha_Hora", "ID_Vendedor", "Nombre", "Total_Vendido", "Comision", "Neto_Entregado", "Referencia", "Estatus"]);
    }

    const datos = hojaCajas.getDataRange().getDisplayValues();
    let registrosCajas = [];

    // Empezamos desde i = 1 para saltar la fila de títulos (cabeceras)
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0]) {
        registrosCajas.push({
          fechaHora: String(datos[i][0]),
          idVendedor: String(datos[i][1]),
          nombre: String(datos[i][2]),
          totalVendido: parseFloat(datos[i][3]) || 0,
          comision: parseFloat(datos[i][4]) || 0,
          netoEntregado: parseFloat(datos[i][5]) || 0,
          referencia: String(datos[i][6] || "N/A"), // Lectura de la columna de Referencia
          estatus: String(datos[i][7] || "Pendiente") // Lectura de la columna de Estatus
        });
      }
    }

    // Retorna los registros invirtiendo el orden para ver los más recientes primero
    return { exito: true, historialCajas: registrosCajas.reverse() };
  } catch (err) {
    return { exito: false, mensaje: err.toString() };
  }
}

function asegurarConfiguracionGlobal(hojaConfig) {
  try {
    if (!hojaConfig.getRange("A2").getValue()) {
      hojaConfig.getRange("A2").setValue("Precio_Carton");
      hojaConfig.getRange("B2").setValue(PRECIO_TICKET_DEFECTO);
    }
    if (!hojaConfig.getRange("A3").getValue()) {
      hojaConfig.getRange("A3").setValue("Comision_Admin_%");
      hojaConfig.getRange("B3").setValue(0.10);
      hojaConfig.getRange("B3").setNumberFormat("0%");
    }

    hojaConfig.getRange("A4").setValue("Total_Ventas");
    hojaConfig.getRange("B4").setFormula("=SUM(Tickets!I2:I)");

    hojaConfig.getRange("A5").setValue("Total_Comision_Vendedores");
    hojaConfig.getRange("B5").setFormula("=B4*D2");

    hojaConfig.getRange("A6").setValue("Total_Comision_Admin");
    hojaConfig.getRange("B6").setFormula("=B4*B3");

    hojaConfig.getRange("A7").setValue("POTE_PREMIO");
    hojaConfig.getRange("B7").setFormula("=B4-B5-B6");
    
    SpreadsheetApp.flush(); 
  } catch (e) {}
}

function obtenerMetricasGlobalesAdmin() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hojaConfig = ss.getSheetByName(HOJA_CONFIG);
    if (!hojaConfig) {
      hojaConfig = ss.insertSheet(HOJA_CONFIG);
      hojaConfig.appendRow(["Parametro", "Valor"]);
      hojaConfig.appendRow(["Precio_Carton", PRECIO_TICKET_DEFECTO]);
      hojaConfig.appendRow(["Comision_Admin_%", 0.10]);
    }

    asegurarConfiguracionGlobal(hojaConfig);

    let totalVentas = hojaConfig.getRange("B4").getValue();
    let totalComisionVendedores = hojaConfig.getRange("B5").getValue();
    let totalComisionAdmin = hojaConfig.getRange("B6").getValue();
    let potePremio = hojaConfig.getRange("B7").getValue();
    let precioCarton = hojaConfig.getRange("B2").getValue();
    let comisionAdminPorc = hojaConfig.getRange("B3").getValue();

    return {
      exito: true,
      totalVentas: parseFloat(totalVentas) || 0,
      totalComisionVendedores: parseFloat(totalComisionVendedores) || 0,
      totalComisionAdmin: parseFloat(totalComisionAdmin) || 0,
      potePremio: parseFloat(potePremio) || 0,
      precioCarton: parseFloat(precioCarton) || PRECIO_TICKET_DEFECTO,
      comisionAdminPorc: parseFloat(comisionAdminPorc) || 0.10
    };
  } catch (err) {
    return { exito: false, mensaje: err.toString() };
  }
}

function obtenerParametrosAdminPersonalizados() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hojaConfig = ss.getSheetByName(HOJA_CONFIG);
    if (!hojaConfig) {
      return { exito: true, precioTicket: PRECIO_TICKET_DEFECTO, comisionAdmin: 0.10, pote_premio: 0 };
    }
    
    asegurarConfiguracionGlobal(hojaConfig);
    
    let precioCarton = parseFloat(hojaConfig.getRange("B2").getValue()) || PRECIO_TICKET_DEFECTO;
    let comisionAdmin = parseFloat(hojaConfig.getRange("B3").getValue()) || 0.10;
    let potePremio = parseFloat(hojaConfig.getRange("B7").getValue()) || 0;

    return {
      exito: true,
      precioTicket: precioCarton,
      comisionAdmin: comisionAdmin,
      pote_premio: potePremio,
      poteGlobal: potePremio
    };
  } catch (err) {
    return { exito: false, mensaje: err.toString() };
  }
}

function obtenerTicketsAdmin() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_TICKETS);
    if (!hoja) return { exito: true, tickets: [] };
    
    var datos = hoja.getDataRange().getDisplayValues();
    var tickets = [];
    
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0]) {
        let montoTicket = parseFloat(datos[i][8]) || PRECIO_TICKET_DEFECTO;
        tickets.push({
          idTicket: datos[i][0],                 
          fecha: datos[i][1],                    
          vendedor: datos[i][2],                 
          monto: montoTicket,          
          numeros: datos[i][4] || "Sin números", 
          estado: datos[i][5] || "ACTIVO",       
          referencia: datos[i][6] || "N/D",      
          modalidad: datos[i][7] || "N/D"        
        });
      }
    }
    return { exito: true, tickets: tickets.reverse() };
  } catch (error) {
    return { exito: false, error: error.toString() };
  }
}

function obtenerVendedoresAdmin() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_VENDEDORES);
    if (!hoja) return { exito: true, vendedores: [] };
    
    var datos = hoja.getDataRange().getDisplayValues();
    var vendedores = [];
    
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0]) {
        vendedores.push({
          idVendedor: datos[i][0],
          nombre: datos[i][1],
          comision: datos[i][2],
          totalVendido: datos[i][3] || 0,
          comisionAcumulada: datos[i][4] || 0,
          usuario: datos[i][5],
          estado: datos[i][7] || "ACTIVO"
        });
      }
    }
    return { exito: true, vendedores: vendedores };
  } catch (error) {
    return { exito: false, error: error.toString() };
  }
}

function crearVendedorAdmin(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_VENDEDORES);
    if (!hoja) return { exito: false, mensaje: "Hoja VENDEDORES não encontrada" };

    var datos = hoja.getDataRange().getDisplayValues();
    var idNuevo = data.idVendedor;

    if (!idNuevo) {
      var ultimoId = "MAFRAS-00";
      if (datos.length > 1) {
        ultimoId = datos[datos.length - 1][0] || "MAFRAS-00";
      }

      let match = ultimoId.match(/^(.*-)(\d+)$/);
      if (match) {
        let prefijo = match[1]; 
        let numero = parseInt(match[2], 10) + 1;
        idNuevo = prefijo + String(numero).padStart(2, '0'); 
      } else {
        idNuevo = "MAFRAS-" + String(datos.length).padStart(2, '0');
      }
    }

    var nombre = data.nombre;
    var comision = data.comision || 18;
    var usuario = data.usuario;
    var clave = data.clave;

    hoja.appendRow([idNuevo, nombre, comision, 0, 0, usuario, clave, "ACTIVO"]);
    return { exito: true };
  } catch (err) {
    return { exito: false, mensaje: err.toString() };
  }
}

function actualizarEstadoVendedorAdmin(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_VENDEDORES);
    var datos = hoja.getDataRange().getDisplayValues();

    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim() === String(data.idVendedor).trim()) {
        hoja.getRange(i + 1, 8).setValue(data.estado); 
        return { exito: true };
      }
    }
    return { exito: false, mensaje: "Vendedor no encontrado" };
  } catch (err) {
    return { exito: false, mensaje: err.toString() };
  }
}

function actualizarComisionVendedorAdmin(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_VENDEDORES);
    if (!hoja) return { exito: false, mensaje: "Hoja VENDEDORES no encontrada" };

    var datos = hoja.getDataRange().getDisplayValues();
    var idTarget = String(data.idVendedor || "").trim();
    var nuevaComision = parseFloat(data.comision || data.porcentajeComision);

    if (isNaN(nuevaComision)) return { exito: false, mensaje: "Comisión inválida" };

    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim() === idTarget) {
        hoja.getRange(i + 1, 3).setValue(nuevaComision);
        return { exito: true };
      }
    }
    return { exito: false, mensaje: "Vendedor no encontrado" };
  } catch (err) {
    return { exito: false, mensaje: err.toString() };
  }
}

function actualizarParametrosAdmin(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaConfig = ss.getSheetByName(HOJA_CONFIG);
    if (!hojaConfig) {
      hojaConfig = ss.insertSheet(HOJA_CONFIG);
      hojaConfig.appendRow(["Parametro", "Valor"]);
      hojaConfig.appendRow(["Precio_Carton", PRECIO_TICKET_DEFECTO]);
      hojaConfig.appendRow(["Comision_Admin_%", 0.10]);
    }

    var precioNew = parseFloat(data.precioTicket);
    var comisionAdminNew = parseFloat(data.comisionAdmin);

    if (!isNaN(precioNew) && precioNew > 0) {
      hojaConfig.getRange("B2").setValue(precioNew);
    }

    if (!isNaN(comisionAdminNew)) {
      var valorDecimal = comisionAdminNew > 1 ? comisionAdminNew / 100 : comisionAdminNew;
      var celdaB3 = hojaConfig.getRange("B3");
      celdaB3.setValue(valorDecimal);
      celdaB3.setNumberFormat("0%");
    }

    asegurarConfiguracionGlobal(hojaConfig);

    return { exito: true };
  } catch (err) {
    return { exito: false, mensaje: err.toString() };
  }
}

function obtenerHistorialVentasFechas(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetTickets = ss.getSheetByName(HOJA_TICKETS);
    if (!sheetTickets) return { exito: true, historial: [], totalVendido: 0, totalComision: 0, neto: 0 };
    
    const idVendedor = data.idVendedor || "";
    const nombreVendedorFiltro = data.nombreVendedorFiltro || ""; 
    
    let nombreVendedorAsignado = "";
    if (idVendedor) {
      nombreVendedorAsignado = obtenerNombreVendedor(idVendedor);
    } else if (nombreVendedorFiltro) {
      nombreVendedorAsignado = nombreVendedorFiltro;
    }

    const tipoFiltro = data.tipoFiltro || "todos"; 
    const fechaInicioStr = data.fechaInicio; 
    const fechaFinStr = data.fechaFin;       

    const datosTickets = sheetTickets.getDataRange().getDisplayValues();
    let historial = [];
    let totalVendido = 0;
    let totalComisionVendedor = 0;

    let comisionPorc = 18;
    if (idVendedor) {
      const datosVend = ss.getSheetByName(HOJA_VENDEDORES).getDataRange().getDisplayValues();
      for(let i = 1; i < datosVend.length; i++){
        if(String(datosVend[i][0]).trim() === String(idVendedor).trim()) { 
          comisionPorc = parseFloat(datosVend[i][2]) || 18; 
          break; 
        }
      }
    }

    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    for (let i = 1; i < datosTickets.length; i++) {
      let idTicket = String(datosTickets[i][0] || "").trim();
      if (!idTicket) continue;
      
      let fechaStr = String(datosTickets[i][1] || "").trim(); 
      let vendedorTicket = String(datosTickets[i][2] || "").trim();
      let loteria = String(datosTickets[i][3] || "").trim();
      let numeros = String(datosTickets[i][4] || "").trim();
      let estado = String(datosTickets[i][5] || "ACTIVO").trim();
      let referencia = String(datosTickets[i][6] || "").trim();
      let modalidad = String(datosTickets[i][7] || "").trim();
      let montoTicket = parseFloat(datosTickets[i][8]) || PRECIO_TICKET_DEFECTO;

      if (nombreVendedorAsignado && vendedorTicket.toLowerCase() !== String(nombreVendedorAsignado).toLowerCase()) {
        continue;
      }

      let partesFechaHora = fechaStr.split(" ");
      let partesFecha = partesFechaHora[0].split("/");
      if (partesFecha.length === 3) {
        let d = parseInt(partesFecha[0], 10);
        let m = parseInt(partesFecha[1], 10) - 1;
        let y = parseInt(partesFecha[2], 10);
        let fechaTicketDate = new Date(y, m, d);
        fechaTicketDate.setHours(0, 0, 0, 0);

        if (tipoFiltro === "hoy") {
          if (fechaTicketDate.getTime() !== hoy.getTime()) continue;
        } else if (tipoFiltro === "semana") {
          let primerDiaSemana = new Date(hoy);
          primerDiaSemana.setDate(hoy.getDate() - hoy.getDay());
          if (fechaTicketDate < primerDiaSemana) continue;
        } else if (tipoFiltro === "mes") {
          if (fechaTicketDate.getMonth() !== hoy.getMonth() || fechaTicketDate.getFullYear() !== hoy.getFullYear()) continue;
        } else if (tipoFiltro === "personalizado" && fechaInicioStr && fechaFinStr) {
          let fI = new Date(fechaInicioStr);
          fI.setHours(0, 0, 0, 0);
          let fF = new Date(fechaFinStr);
          fF.setHours(23, 59, 59, 999);
          if (fechaTicketDate < fI || fechaTicketDate > fF) continue;
        }
      }

      totalVendido += montoTicket;

      historial.push({
        idTicket: idTicket,
        fecha: fechaStr,
        vendedor: vendedorTicket,
        loteria: loteria,
        numeros: numeros.split(", "),
        estado: estado,
        referencia: referencia,
        modalidad: modalidad,
        monto: montoTicket
      });
    }

    totalComisionVendedor = totalVendido * (comisionPorc / 100);
    const comisionAdminPorc = obtenerComisionAdmin();
    let comisionAdminCalculada = totalVendido * comisionAdminPorc; 
    let neto = totalVendido - totalComisionVendedor - comisionAdminCalculada;

    let animalesSorteoActual = obtenerAnimalitosSorteoActual();

    return {
      exito: true,
      historial: historial.reverse(),
      totalVendido: totalVendido,
      totalComision: totalComisionVendedor,
      comisionAdmin: comisionAdminCalculada,
      neto: neto,
      comisionPorc: comisionPorc
    };
  } catch (err) {
    return { exito: false, mensaje: err.toString() };
  }
}