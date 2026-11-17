/* Configuración del Reto 2 · Estilos, Imágenes y Elementos Especiales */
window.EXAMEN_CFG = Object.assign({
 "key": "tbox_reto2_v1",
 "titulo": "Reto 2 · Estilos, Imágenes y Elementos Especiales",
 "etiqueta": "Área",
 "etiquetaPlural": "áreas",
 "total": 100,
 "aprobado": 70,
 "pdf": false,
 "docTrabajo": "Reto_2.docx",
 "docBase": "Reto2_Base.docx",
 "carpeta": "Documentos › TBox_Word › Reto_2",
 "descripcionDoc": "Trabajas sobre el documento base «Bono – Inversiones Inteligentes».",
 "tareas": [
  {
   "id": 1,
   "titulo": "Estilos y diseño",
   "resumen": "Aplica estilos, un diseño de documento y limpia formatos.",
   "pasos": [
    {
     "item": 1,
     "t": "A los <b>cuatro subtítulos numerados</b> del documento (1. a 4.) aplícales el estilo <b>Título 2</b>.",
     "ayuda": "Selecciona el subtítulo · Inicio › Estilos › Título 2 · repite en los otros tres."
    },
    {
     "item": 2,
     "t": "A todo el documento aplícale el diseño <b>Sombreado</b>.",
     "ayuda": "Diseño › Formato del documento › elige el conjunto de estilos «Sombreado»."
    },
    {
     "item": 3,
     "t": "Elimina el formato de la frase <b>«desde Investment Grade hasta High Yield»</b> (sin borrar el texto).",
     "ayuda": "Selecciona la frase · Inicio › Borrar todo el formato (A con borrador)."
    }
   ]
  },
  {
   "id": 2,
   "titulo": "Texto y símbolos",
   "resumen": "Resaltado y caracteres especiales.",
   "pasos": [
    {
     "item": 4,
     "t": "Antes de la palabra <b>Moody’s</b> inserta el símbolo de <b>Copyright (©)</b>.",
     "ayuda": "Cursor justo a la izquierda de «Moody’s» · Insertar › Símbolo › Más símbolos › ©."
    },
    {
     "item": 5,
     "t": "A la <b>fórmula</b> de la sección 2 aplícale un resaltado <b>Gris 25 %</b>.",
     "ayuda": "Selecciona toda la fórmula · Inicio › Color de resaltado de texto › Gris 25 %."
    }
   ]
  },
  {
   "id": 3,
   "titulo": "Imágenes",
   "resumen": "Inserción y ajustes de imagen.",
   "pasos": [
    {
     "item": 6,
     "t": "Inserta la imagen <b>bono.jpg</b> <b>después del primer párrafo</b> y <b>aliníala al centro</b>.",
     "ayuda": "Cursor al final del primer párrafo · Enter · Insertar › Imágenes › Este dispositivo › bono.jpg · Inicio › Centrar."
    },
    {
     "item": 7,
     "t": "Aplica a la imagen una <b>saturación de color del 66 %</b>.",
     "ayuda": "Selecciona la imagen · Formato de imagen › Color › Saturación de color › Saturación: 66 %."
    }
   ]
  },
  {
   "id": 4,
   "titulo": "Tablas y listas",
   "resumen": "Estilo de tabla y viñetas personalizadas.",
   "pasos": [
    {
     "item": 8,
     "t": "A la tabla aplícale el estilo <b>Tabla de cuadrícula 2</b> (la variante sin color de énfasis).",
     "ayuda": "Clic dentro de la tabla · Diseño de tabla › Estilos de tabla › Tabla de cuadrícula 2."
    },
    {
     "item": 9,
     "t": "A los <b>tres últimos párrafos</b> del documento aplícales una <b>viñeta con la imagen dinero.png</b>.",
     "ayuda": "Selecciona solo esos tres párrafos · Inicio › Viñetas › Definir nueva viñeta › Imagen › Desde un archivo › dinero.png."
    }
   ]
  },
  {
   "id": 5,
   "titulo": "Revisión",
   "resumen": "Comentarios de auditoría.",
   "pasos": [
    {
     "item": 10,
     "t": "En la palabra <b>«Inteligentes»</b> del título coloca un <b>comentario</b> con tu <b>nombre y apellido completos</b> (los mismos que escribiste al iniciar el reto).",
     "ayuda": "Selecciona la palabra «Inteligentes» · Revisar › Nuevo comentario · escribe tu nombre completo."
    }
   ]
  }
 ],
 "recursos": [
  {
   "archivo": "assets/docs/reto2/Reto2_Base.docx",
   "nombre": "Reto2_Base.docx",
   "desc": "Documento base del reto",
   "icono": "fluent:document-24-regular"
  },
  {
   "archivo": "assets/images/bono.jpg",
   "nombre": "bono.jpg",
   "desc": "Imagen para insertar",
   "icono": "fluent:image-24-regular"
  },
  {
   "archivo": "assets/images/dinero.png",
   "nombre": "dinero.png",
   "desc": "Imagen para la viñeta",
   "icono": "fluent:image-24-regular"
  }
 ]
}, {
  recursosB64: RECURSOS_B64_RETO2,
  puntos: CHECKER_RETO2.PUNTOS, nombreNivel: CHECKER_RETO2.NOMBRE_NIVEL,
  evaluar: CHECKER_RETO2.evaluar
});
