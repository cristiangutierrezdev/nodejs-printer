class AccionesJs {
  /**
   * El símbolo 'Fin De Línea' correcto de la plataforma en uso.
   *
   * @return {string}
   */
  static finLinea() {
    return "\n";
  }

  /**
   * Crea una línea de comandos EPL
   *
   * Concatena el array separados por comas en una sola línea
   *
   * @param {string} comando - El comando de línea
   * @param {Array} parametros - Los parámetros del comando EPL
   * @return {string} - Devuelve una línea de código EPL
   */
  static escribirComando(comando, parametros) {
    return comando + parametros.join(",") + AccionesJs.finLinea();
  }
}

class ComandosEpl {
  // Configuración
  static LIMPIAR_BUFFER = "N";
  static INICIAR_ETIQUETA = "";
  static NUMERO_IMPRESION = "P1";
  static COMA = ",";
  static COMILLA = '"';

  // Letras
  static TEXTO_N = "N";
  static TEXTO_R = "R";

  // Estilo de impresión
  static CADENA_A = "A";
  static BARRAS_B = "B";

  // Variables para líneas
  static HORIZONTAL = 1;
  static VERTICAL = 0;

  // Color de líneas
  static LINEA_NEGRA_LO = "LO";
  static LINEA_INVERTIA_LE = "LE";
  static LINEA_BLANCA_LW = "LW";
  static LINEA_NEGRA_DIAGONAL_LS = "LS";

  // Códigos de barra
  // Human Readable = B (Yes) or N (No)
  static BARRA_LEGIBLE_B = "B";
  static BARRA_ILEGIBLE_N = "N";
}

class ExisteImpresoraWindows {
  // Configuración
  static RUTA_POWERSHELL =
    "c:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe";
  static OPCION_PARA_EJECUTAR_COMANDOS = "-c";
  static ESPACIO = " ";
  static COMILLA = '"';
  static COMANDO = "get-WmiObject -class Win32_printer | ft name";

  constructor() {
    //...
  }

  verificarImpresora(dispositivo, informacion = false) {
    let listaDeImpresoras = [];
    let bandera = false;
    let info = "";
    const exec = require("child_process").exec;

    exec(
      ExisteImpresoraWindows.RUTA_POWERSHELL +
        ExisteImpresoraWindows.ESPACIO +
        ExisteImpresoraWindows.OPCION_PARA_EJECUTAR_COMANDOS +
        ExisteImpresoraWindows.ESPACIO +
        ExisteImpresoraWindows.COMILLA +
        ExisteImpresoraWindows.COMANDO +
        ExisteImpresoraWindows.COMILLA,
      (error, stdout, stderr) => {
        if (error) {
          console.error("Error al ejecutar el comando:", stderr);
          return;
        }

        const resultado = stdout
          .split("\n")
          .map((item) => item.trim().replace(/\s+/g, " "))
          .filter((item) => item !== "");
        if (Array.isArray(resultado)) {
          // Verificamos si existe el dispositivo
          bandera = resultado.includes(dispositivo);

          if (informacion) {
            if (bandera) {
              info = `Dispositivo ${dispositivo} encontrado.`;
              console.log(info);
            } else {
              info = `Dispositivo ${dispositivo} NO encontrado.`;
              console.log(info);
            }

            // Omitir los primeros 3 datos del arreglo, porque son el encabezado
            console.log(`Lista de impresoras`);
            for (let i = 3; i < resultado.length; i++) {
              console.log(resultado[i].trim());
            }
          }
        }
      }
    );

    return bandera;
  }
}

class ImprimirEpl {
  /**
   * Método constructor sin parámetros
   */
  constructor() {
    //...
  }

  /**
   * Envía comandos EPL al dispositivo indicado
   *
   * @param {string} etiqueta - Comandos a imprimir
   * @param {string} impresora - El nombre de la impresora
   * @param {boolean} impresion - Si activa la impresión en el dispositivo
   * @param {boolean} informacion - Si desea imprimir información de impresión
   */
  imprimir(etiqueta, impresora, impresion = true, informacion = false) {
    const fs = require("fs");
    const os = require("os");
    const exec = require("child_process").exec;

    const archivoTemporal = fs.mkdtempSync(`${os.tmpdir()}/Eti`);
    fs.writeFileSync(archivoTemporal, etiqueta);

    if (impresion) {
      exec(
        `print /d:"\\%COMPUTERNAME%\\${impresora}" "${archivoTemporal}"`,
        (error, stdout, stderr) => {
          if (error) {
            console.error("Error al imprimir:", stderr);
            return;
          }
        }
      );
    }

    fs.unlinkSync(archivoTemporal);

    if (informacion) {
      let informe = `**************************************************************************`;
      informe +=
        `*Información de impresión` +
        `**************************************************************************` +
        `Impresora: ${impresora}` +
        `Sistema operativo: ${os.type()}` +
        `Host: ${os.hostname()}` +
        `Fecha: ${new Date().toLocaleDateString()}` +
        `Hora: ${new Date().toLocaleTimeString()}` +
        `Comandos EPL: ${etiqueta}` +
        `Archivo temporal creado: ${archivoTemporal}`;

      informe += `Archivo eliminado: se ha eliminado el archivo temporal.`;

      if (impresion) {
        informe += `Impresión: activada.`;
      } else {
        informe += `Impresión: desactivada.`;
        informe += `Nota: La impresión está desactivada, para habilitarla, cambiar a true en la función: -> imprimir().`;
      }
      informe +=
        `Nota: Para desactivar la información, cambiar a false en la llamada de la función: -> imprimir().*` +
        `**************************************************************************`;

      console.log(informe);
    }
  }

  /**
   * Se agrega un encabezado y un pie a la etiqueta.
   *
   * @param {string} etiqueta - El contenido de la etiqueta
   * @param {number} cantidad - La cantidad de impresión
   * @return {string} - Devuelve la etiqueta lista para imprimir
   */
  construirEtiqueta(etiqueta, cantidad = 1) {
    let nuevaEtiqueta = ComandosEpl.INICIAR_ETIQUETA + AccionesJs.finLinea();
    nuevaEtiqueta += ComandosEpl.LIMPIAR_BUFFER + AccionesJs.finLinea();
    nuevaEtiqueta += etiqueta;
    nuevaEtiqueta +=
      ComandosEpl.NUMERO_IMPRESION +
      ComandosEpl.COMA +
      cantidad +
      AccionesJs.finLinea();
    return nuevaEtiqueta;
  }

  /**
   * Escribe texto de caracteres ascii
   *
   * @param {string} texto - La cadena de texto
   * @param {number} ejeX - Posición en el eje x (horizontal)
   * @param {number} ejeY - Posición en el eje y (vertical)
   * @param {number} fuente - Tamaño de la fuente [1-5]
   * @param {boolean} constraste - Invierte el contraste del texto
   * @param {number} rotacion - Rotación del texto (0, 90, 180, 270 grados)
   * @param {number} expandirX - Multiplicador horizontal
   * @param {number} expandirY - Multiplicador vertical
   * @return {string} - Devuelve el comando completo
   */
  escribirTexto(
    texto,
    ejeX,
    ejeY,
    fuente,
    constraste = false,
    rotacion = 0,
    expandirX = 1,
    expandirY = 1
  ) {
    const style = constraste ? ComandosEpl.TEXTO_R : ComandosEpl.TEXTO_N;
    return AccionesJs.escribirComando(ComandosEpl.CADENA_A, [
      ejeX,
      ejeY,
      rotacion,
      fuente,
      expandirX,
      expandirY,
      style,
      `${ComandosEpl.COMILLA}${texto}${ComandosEpl.COMILLA}`,
    ]);
  }

  /**
   * Pinta una línea específica
   *
   * @param {number} ejeX - Posición inicial en el eje x
   * @param {number} ejeY - Posición inicial en el eje y
   * @param {number} tamanho - Tamaño de la línea en puntos
   * @param {number} grosor - Grosor de la línea en puntos
   * @param {number} orientacion - Orientación de la línea (0: vertical, 1: horizontal)
   * @param {number} colorLinea - Color de la línea (0: negro, 1: invertido, 2: blanco)
   * @return {string}
   */
  pintarLinea(
    ejeX,
    ejeY,
    tamanho,
    grosor = 10,
    orientacion = 1,
    colorLinea = 0
  ) {
    let xTamanho, yTamanho;
    if (orientacion === ComandosEpl.HORIZONTAL) {
      xTamanho = tamanho;
      yTamanho = grosor;
    } else {
      yTamanho = tamanho;
      xTamanho = grosor;
    }

    let comando = ComandosEpl.LINEA_NEGRA_LO;
    switch (colorLinea) {
      case 1:
        comando = ComandosEpl.LINEA_INVERTIA_LE;
        break;
      case 2:
        comando = ComandosEpl.LINEA_BLANCA_LW;
        break;
    }

    return AccionesJs.escribirComando(comando, [
      ejeX,
      ejeY,
      xTamanho,
      yTamanho,
    ]);
  }

  /**
   * Escribe un código de barra
   *
   * @param {string} dato - Datos del código de barra
   * @param {number} ejeX - Posición en el eje x
   * @param {number} ejeY - Posición en el eje y
   * @param {number} altura - Altura de puntos de código de barras
   * @param {number} tipo - Selección de código de barras
   * @param {boolean} legibilidad - Código legible para humanos
   * @param {number} rotacion - Rotación del código de barra (0, 90, 180, 270 grados)
   * @param {number} estrecho - Estrecho de la barra
   * @param {number} anchuraBarras - Anchura
   * @return {string}
   */
  pintarCodigoBarra(
    dato,
    ejeX,
    ejeY,
    altura,
    tipo = 3,
    legibilidad = true,
    rotacion = 0,
    estrecho = 2,
    anchuraBarras = 3
  ) {
    const legibilidadH = legibilidad
      ? ComandosEpl.BARRA_LEGIBLE_B
      : ComandosEpl.BARRA_ILEGIBLE_N;
    return AccionesJs.escribirComando(ComandosEpl.BARRAS_B, [
      ejeX,
      ejeY,
      rotacion,
      tipo,
      estrecho,
      anchuraBarras,
      altura,
      legibilidadH,
      `${ComandosEpl.COMILLA}${dato}${ComandosEpl.COMILLA}`,
    ]);
  }
}

// Ejemplo de uso:
const linea = AccionesJs.escribirComando("A", ["param1", "param2", "param3"]);
console.log(linea);

const impresora = new ExisteImpresoraWindows();
impresora.verificarImpresora("Microsoft Print to PDF", true);

const imprimirEpl = new ImprimirEpl();
const etiqueta = imprimirEpl.construirEtiqueta("Test Label", 2);
console.log(etiqueta);

imprimirEpl.imprimir(etiqueta, "Microsoft Print to PDF", false, true);

// // Definir y construir una etiqueta con información específica
// const fecha = new Date().toLocaleDateString();
// const hora = new Date().toLocaleTimeString();

// let etiquetaCompleta = imprimirEpl.escribirTexto(
//   "TABASCO ES MAS",
//   170,
//   10,
//   1,
//   false,
//   0,
//   3,
//   3
// );
// etiquetaCompleta += imprimirEpl.pintarLinea(5, 75, 765, 10, 1, 0);
// etiquetaCompleta += imprimirEpl.escribirTexto(
//   "410",
//   350,
//   95,
//   5,
//   false,
//   0,
//   2,
//   2
// );
// etiquetaCompleta += imprimirEpl.escribirTexto(
//   "43 - 10",
//   200,
//   210,
//   4,
//   false,
//   0,
//   3,
//   3
// );
// etiquetaCompleta += imprimirEpl.escribirTexto(
//   "#800",
//   310,
//   310,
//   4,
//   false,
//   0,
//   3,
//   3
// );
// etiquetaCompleta += imprimirEpl.escribirTexto(
//   "BS2002142400",
//   240,
//   405,
//   2,
//   false,
//   0,
//   2,
//   2
// );
// etiquetaCompleta += imprimirEpl.escribirTexto(
//   `${fecha} Cunduacan Tabasco ${hora}`,
//   140,
//   450,
//   3,
//   false,
//   0,
//   1,
//   1
// );

// // Imprimir la etiqueta construida
// imprimirEpl.imprimir(
//   imprimirEpl.construirEtiqueta(etiquetaCompleta, 1),
//   "Microsoft Print to PDF",
//   false,
//   true
// );
