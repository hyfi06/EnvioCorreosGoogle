function Envio_Correo (){
  // CONFIGURACIÓN //
  var Hoja = "Hoja 1";
  var colInicial = 2;
  var NumColEscribir = 10; // # de la columna donde se escribirá la salida, con A=0,B=1, etc.
  var LetColEscribir = "K"; // Letra de la columna donde se escribirá, entre comillas y en mayúscula ejs. "A", "B",etc.
  var NombreHTML = "correo.html"; //Nombre del archivo html del mensaje (Archivo -> Nuevo -> Archivo html)
  var NombreDelRemitente = "Name"; //Nombre 
  var NumColCorreo = 2; // # de la columna donde se encuentra la información del correo, con A=0,B=1, etc.
  var Asunto = "Aquí va el asuntos";
  
  
  // LECTURA DE DATOS //
  var Rango = Leer(Hoja,colInicial); //cargamos los datos con la CONFIGURACIÓN GENERAL
  for ( var i in Rango){ //recorremos todos los datos
    var DatosR = Rango[i]; //cargamos los datos del rengón i
    
    // CONDICIONES PARA OMITIR RENGLÓN //
    if( DatosR[NumColEscribir] != "" ) continue; // Omitimos este reglón si hay algo escrito en la columna de salida (Si ya se envió el correo)
    
    // PERSONALIZACIÓN DEL CORREO // <-------------------------------------------
    // En este lugar se colocan las cosas a rempalzar con el formato "A remplazar":"por lo que rempezar", separados por comas
    var remplazo = {"-NOMBRES-":DatosR[1],
                    "-A_remplazar-":"Por lo que remplazar",
                    };
    
    // MENSAJE //
    var html = HtmlService.createHtmlOutputFromFile(NombreHTML).getContent(); //Cargamos el mensaje del archivo html
    //var html ="<p> </p>"  //escribir el mensaje directamente aquí comentando la linea anterior. se puede usar html += "" para ir agregando texto.
    
        
    // Ciclo de remplazo //
    html = replaceString(html,remplazo);
    
    
    // CORREO //
    //Logger.log(html); //previsualización del correo Ctrl + Intro. Ojo: comentar el con /* */ el bloque de try{}catch(e){}
    
    var DatosEscribir = {};
    try{
      MailApp.sendEmail({
        name: NombreDelRemitente, //Nombre del remitente
        to: DatosR[NumColCorreo], // Destinatario
        subject: Asunto, //Asunto
        htmlBody: html //texto del mensaje
      });
      DatosEscribir[LetColEscribir] = Fecha(null,2);
    } catch(e) {
      DatosEscribir[LetColEscribir] = "Error: " +e;
    } 
    
    Escribir (DatosEscribir,i,Hoja,colInicial);
  }
}



//  FUNCIONES  //

/**
* Lee los datos de un Google Sheet
* @param {string} NHoja Nombre de la hoja
* @param {number} RowIni Renglón donde empiezan los datos
* @param {string} idLibro id del libro
* @return {Object[][]} Datos de la tabla.
*/
function Leer (NHoja,RowIni,idLibro) {
  // Toma el rango de datos registrados
  
  var idLibro = idLibro || SpreadsheetApp.getActiveSpreadsheet().getId();
  try{
    var Libro =  SpreadsheetApp.openById(idLibro);
  } catch(e){
    return null;
  }
  
  var HojaRegistro_Datos =  Libro.getSheetByName(NHoja) || Libro.getActiveSheet();
  if ( HojaRegistro_Datos == null ) return null;
  
  if ( typeof RowIni !== 'number' ) {
    var RowIni = 3;
  }
  
  if ( HojaRegistro_Datos.getLastRow() < RowIni) return null;
  
  var Rango = HojaRegistro_Datos.getRange(RowIni,1,HojaRegistro_Datos.getLastRow() - RowIni + 1,HojaRegistro_Datos.getLastColumn()).getA1Notation();
  return HojaRegistro_Datos.getRange(Rango).getValues();
}

/**
* Escribe en datos en un celda de una hoja de Google Sheet
* @param {Object} datos Datos a escribir. Ejemplo "A":dato
* @param {number} renglon Número de renglón a escribir (empezado de 0)
* @param {string} NHoja Nombre de la hoja
* @param {number} RowIni Renglón donde empiezan los datos
* @param {string} idLibro id del libro
*/
function Escribir (datos,renglon,NHoja,RowIni,idLibro){
  var idLibro = idLibro || SpreadsheetApp.getActiveSpreadsheet().getId();
  try{
    var Libro =  SpreadsheetApp.openById(idLibro);
  } catch(e){
    return null;
  }
  var HojaRegistro_Datos =  Libro.getSheetByName(NHoja) || Libro.getActiveSheet();
  if ( HojaRegistro_Datos == null ) return null;
  
  if ( typeof RowIni !== 'number' ) {
    var RowIni = 3;
  }
  
  var regRenglon =/[0-9]+/;
  if (regRenglon.test(renglon)){
    var IndiceHoja = parseInt(renglon) + RowIni;
  } else {
    var IndiceHoja = HojaRegistro_Datos.getLastRow()+1;
  }

  for ( var i in datos ){
    var RangoIndice = i + IndiceHoja;
    HojaRegistro_Datos.getRange(RangoIndice).setValue(datos[i]);
  }
}


/**
* Coloca la fecha en un formato específico.
* @param {date} Dfecha Fecha especifica, si se omita se toma la fecha actual.
* @param {number} formato 0 a 10.
* @return {string} Cadena con la fecha con el formato indicado.
*/
function Fecha(Dfecha,formato) {
  var meses = new Array ("enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre");
  var formato = (formato % 13) || 0;
  
  var f = Dfecha || new Date();
  var fechaA = "";
  
  switch ( formato ) {
    case 0 : //0 : 06 de septiembre de 2018
      fechaA += Ceros(f.getDate()) + " de " + meses[f.getMonth()] + " de " + f.getFullYear();
      break;
    case 1 : //1 : 14:01:01
      fechaA += Ceros(f.getHours()) +":"+Ceros(f.getMinutes())+":"+Ceros(f.getSeconds());
      break;
    case 2 : //2 : 2018-09-06 14:01:01
      fechaA += f.getFullYear() +"-"+ Ceros(f.getMonth()+1) + "-"+ Ceros(f.getDate()) +" "+ Ceros(f.getHours()) +":"+Ceros(f.getMinutes())+":"+Ceros(f.getSeconds());
      break;
    case 3 : //3 : 2018-09-06
      fechaA +=  f.getFullYear() + "-" + Ceros(f.getMonth()+1) + "-"+ Ceros(f.getDate());
      break;
    case 4 : //4 : 2018/09/06
      fechaA += f.getFullYear() + "/" + Ceros(f.getMonth()+1) + "/"+ Ceros(f.getDate());
      break;
    case 5 : //5 : 20180906
      fechaA += f.getFullYear() +""+ Ceros(f.getMonth()+1) + "" + Ceros(f.getDate());
      break;
    case 6 : //6 : 06-09-2018 14:01:01
      fechaA += Ceros(f.getDate()) + "-" + Ceros(f.getMonth()+1) + "-"+ f.getFullYear() +" "+ Ceros(f.getHours()) +":"+Ceros(f.getMinutes())+":"+Ceros(f.getSeconds());
      break;
    case 7 : //7 : 06-09-2018
      fechaA += Ceros(f.getDate()) + "-" + Ceros(f.getMonth()+1) + "-"+ f.getFullYear();
      break;
    case 8 : //8 : 06/09/2018
      fechaA += Ceros(f.getDate()) + "/" + Ceros(f.getMonth()+1) + "/"+ f.getFullYear();
      break;
    case 9 : //9 : 06092018
      fechaA += Ceros(f.getDate()) +""+ Ceros(f.getMonth()+1) + "" + f.getFullYear();
      break;
    case 10 : //10 : 2018
      fechaA += f.getFullYear();
      break;
    case 11 : //11 : 09
      fechaA += Ceros(f.getMonth()+1);
      break;
    case 12 : //12 : 06
      fechaA += Ceros(f.getDate());
      break;
  }
  return fechaA;
}

/**
* Genera números con ceros a la izquierda.
* @param {number} numero Número a colocar Ceros a la izquierda.
* @param {number} grado Longitud mínima de la cadena final.
* @return {string} Cadena de 'numero' con ceros a la izquierda para alcanzar una longitud mínima grado.
*/
function Ceros (numero,grado) {
  var grado = grado || 2;
  if ( typeof grado !== "number"){
    var grado = grado.length;
    if ( typeof grado !== "number") return numero;
  }
  var num_s = numero.toString();
  while (num_s.length < grado) {
        num_s = '0' + num_s;
  }
  return num_s;
}


/**
* Remplazar String
* @param {string} texto Intrada del html
* @param {Object} remplazo Objeto con key lo que se va a remplazar, y dato por lo que se remplazará
* @return {string} Regresa el string con remplazo
*/
function replaceString (texto,remplazo){
  if ( typeof texto !== 'string') return null;
  if ( typeof remplazo === 'undefined' ) return texto;
  
  for (var r in remplazo) texto = texto.replace( new RegExp(r,"g") , remplazo[r] );
  
  return texto;
}
