// CONFIGURACIÓN GENERAL //
var HojaPrincipal= "Hoja 1"; //Nombre de la hoja de donde se van a leer los datos
var RowInicial = 3; //Renglón donde inician los datos (omitir los encabezados)

// Obtiene el id de la hoja de calculo activa
var Libro = SpreadsheetApp.getActiveSpreadsheet();
var HojaRegistro = Libro.getSheetByName(HojaPrincipal);



function Envio_Correo (){
  // CONFIGURACIÓN DE ESCRITURA DE LA SALIDA //
  var NumColEscribir = 10; // # de la columna donde se escribirá la salida, con A=0,B=1, etc.
  var LetColEscribir = "K"; // Letra de la columna donde se escribirá, entre comillas y en mayúscula ejs. "A", "B",etc.
  
  // LECTURA DE DATOS //
  var Rango = Datos(); //cargamos los datos con la CONFIGURACIÓN GENERAL
  for ( var i in Rango){ //recorremos todos los datos
    var DatosR = Rango[i]; //cargamos los datos del rengón i
    
    // CONDICIONES PARA OMITIR RENGLÓN //
    if( DatosR[NumColEscribir] != "" ) continue; // Omitimos este reglón si hay algo escrito en la columna de salida (Si ya se envió el correo)
    
    // MENSAJE //
    var html = HtmlService.createHtmlOutputFromFile("correo.html").getContent(); //Cargamos el mensaje del archivo html (Archivo -> Nuevo -> Archivo html)
    //var html ="<p> </p>"  //escribir el mensaje directamente aquí comentando la linea anterior. se puede usar html += "" para ir agregando texto.
    
    // REMPLAZAR //
    var nombre = DatosR[0]+" "+ DatosR[1]+ " "+DatosR[2]; //variable creada para facilitar el remplazo
    
    // En este lugar se colocan las cosas a rempalzar con el formato "A remplazar":"por lo que rempezar", separados por comas
    var remplazo = {"|*NOMBRES*|":nombre , "|*NOMBREPILA*|":DatosR[2] , "|*UNIVERSIDAD*|":DatosR[4]};
    
    // Ciclo de remplazo //
    for (var j in remplazo){
      var html = html.replace(j,remplazo[j]);
    }
    
    
    // CORREO //
    //Logger.log(html); //previsualización del correo Ctrl + Intro. Ojo: comentar el con /* */ el bloque de try{}catch(e){}
    
    try{
      MailApp.sendEmail({
        name: "Secretaría de Asuntos Estudiantiles", //Nombre del remitente
        to: DatosR[6] + "," + DatosR[7], // Destinatario
        subject:"Invitación a compartir tu experiencia en la Feria de Movilidad", //Asunto
        htmlBody: html //texto del mensaje
      });
      Escribir(LetColEscribir,i,"Enviado");
    } catch(e){
      Escribir(LetColEscribir,i,"No enviado, error: "+ e);
    }
    
    
    
  }
}



//  FUNCIONES  //

function Datos(NHoja,RowIni){ // Lee los datos de la hoja indicada o por defecto los de la CONFIGURACIÓN GENERAL
  var HojaRegistro_Datos =  Libro.getSheetByName(NHoja) || HojaRegistro; //cargamos la hoja con nombre NHoja o la hoja de la CONFIGURACIÓN GENERAL
  var RowIni = RowIni || RowInicial; // Renglón inicial a leer RowIni o el de la CONFIGURACIÓN GENERAL
  var Rango = HojaRegistro_Datos.getRange(RowIni,1,HojaRegistro_Datos.getLastRow() - RowIni + 1,HojaRegistro_Datos.getLastColumn()).getA1Notation(); // Rango de datos
  return HojaRegistro_Datos.getRange(Rango).getValues(); //Regresamos los datos leidos [][]
}

function Escribir(Columna,renglon,mensaje,NHoja,RowIni){
  // Esta función escribe el texto mensaje en la NHoja contando desde el rengón RowIni (si se omiten se toman de CONFIGURACIÓN GENERAL) 
  // en la columna Columna (A,B,C,..) en el rengón de datos reglon
  var RowIni = RowIni || RowInicial; //Reglón inical o el de CONFIGURACIÓN GENERAL
  var IndiceHoja = parseInt(renglon) + RowIni; //recalculamos el reglón de los datos al de la hoja
  var RangoIndice = Columna + IndiceHoja; //Generamos el codigo A1 de la celda donde escribir
  var HojaRegistro_Datos =  Libro.getSheetByName(NHoja) || HojaRegistro; // cargamos la hoja
  HojaRegistro_Datos.getRange(RangoIndice).setValue(mensaje); //Escribimos en la hoja en la celda indicada
}

function GCURP (CURP, formato) {
  // Ingresa el curp y devuelve el el genero en diferentes formatos
  // (default)0: a/o, 1: la/el, 2: de la / del 3: a la / al, 4:M/H, 5: F/M
  var genero = new Array ("Mujer","Hombre");
  var g = "";
  var formato = (formato % 6) || 0; //Default
  
  switch ( formato ) {
    case 0:
      if ( CURP.substr(10,1) == "M") {
        g += "a";
      } else if ( CURP.substr(10,1) == "H"){
        g += "o";
      } else {
        g += "x";
      }
      break;
    case 1:
      if ( CURP.substr(10,1) == "M") {
        g += "la";
      } else {
        g += "el";
      }
      break;
    case 2:
      if ( CURP.substr(10,1) == "M") {
        g += "de la";
      } else {
        g += "del";
      }
      break;
    case 3:
      if ( CURP.substr(10,1) == "M") {
        g += "a la";
      } else {
        g += "al";
      }
      break;
	case 4:
      if ( CURP.substr(10,1) == "M") {
        g += "M";
      } else {
        g += "H";
      }
      break;
	case 5:
      if ( CURP.substr(10,1) == "M") {
        g += "F";
      } else {
        g += "M";
      }
      break;
  }
  return g;
}


function Ceros (numero,grado) { // coloca el número "numero" en formato de dígitos enteros "grado", colocando ceros al incio si su longitud es menor a "grado"
  var grado = grado || 2;
  var num_s = numero.toString()
  while (num_s.length < grado) {
        num_s = '0' + num_s;
  }
  return num_s;
}

function FechaActual(Dfecha,formato) {
  var meses = new Array ("enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre");
  var formato = (formato % 6) || 0;
  
  var f = Dfecha || new Date();
  var fechaA = "";
  
  switch ( formato ) {
    case 0 : //0 : 06 de septiembre de 2018
      fechaA += Ceros(f.getDate()) + " de " + meses[f.getMonth()] + " de " + f.getFullYear();
      break;
    case 1 : //1 : 20180906
      fechaA += f.getFullYear() +""+ Ceros(f.getMonth()+1) + "" + Ceros(f.getDate());
      break;
    case 2 : //2 : 2018-09-06 14:01:01
      fechaA += f.getFullYear() +"-"+ Ceros(f.getMonth()+1) + "-"+ Ceros(f.getDate()) +" "+ Ceros(f.getHours()) +":"+Ceros(f.getMinutes())+":"+Ceros(f.getSeconds());
      break;
    case 3 : //3 : 06-09-2018 14:01:01
      fechaA += Ceros(f.getDate()) + "-" + Ceros(f.getMonth()+1) + "-"+ f.getFullYear() +" "+ Ceros(f.getHours()) +":"+Ceros(f.getMinutes())+":"+Ceros(f.getSeconds());
      break;
    case 4 : //4 : 06-09-2018
      fechaA += Ceros(f.getDate()) + "-" + Ceros(f.getMonth()+1) + "-"+ f.getFullYear();
      break;
    case 5 : //5 : 14:01:01
      fechaA += Ceros(f.getHours()) +":"+Ceros(f.getMinutes())+":"+Ceros(f.getSeconds());
      break;
  }
  return fechaA;
}
