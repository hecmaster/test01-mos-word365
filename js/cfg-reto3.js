/* Configuración del Reto 3 · Formatos Avanzados y Elementos Visuales */
window.EXAMEN_CFG = Object.assign({
 "key": "tbox_reto3_v1",
 "titulo": "Reto 3 · Formatos Avanzados y Elementos Visuales",
 "etiqueta": "Área",
 "etiquetaPlural": "áreas",
 "total": 100,
 "aprobado": 70,
 "pdf": false,
 "docTrabajo": "Reto_3.docx",
 "docBase": "Reto3_Base.docx",
 "carpeta": "Documentos › TBox_Word › Reto_3",
 "descripcionDoc": "Trabajas sobre el documento base «Liquidez y rendimiento».",
 "tareas": [
  {
   "id": 1,
   "titulo": "Encabezado y pie de página",
   "resumen": "Encabezados y numeración de página con diseños de la galería.",
   "pasos": [
    {
     "item": 1,
     "t": "Inserta el encabezado llamado <b>Con bandas</b> y escribe el texto <b>INVERSIONES INTELIGENTES</b>.",
     "ayuda": "Insertar › Encabezado › Con bandas › en el campo Título escribe INVERSIONES INTELIGENTES."
    },
    {
     "item": 2,
     "t": "Agrega la numeración de página llamada <b>Círculo</b> en la <b>parte inferior</b> del documento.",
     "ayuda": "Insertar › Número de página › Final de página › Círculo."
    }
   ]
  },
  {
   "id": 2,
   "titulo": "Diseño de página",
   "resumen": "Márgenes, color de página y columnas.",
   "pasos": [
    {
     "item": 3,
     "t": "Aplica al documento el margen <b>Moderado</b>.",
     "ayuda": "Disposición › Márgenes › Moderado."
    },
    {
     "item": 4,
     "t": "Cambia el color de la página a <b>Gris claro, Fondo 2</b>.",
     "ayuda": "Diseño › Color de página › Colores del tema › Gris claro, Fondo 2."
    },
    {
     "item": 5,
     "t": "Configura en <b>dos columnas</b> todos los párrafos del <b>Tema 3</b> («3. Estrategias operativas para optimizar el par liquidez-rendimiento»).",
     "ayuda": "Selecciona los párrafos debajo del título del Tema 3 (sin el título) · Disposición › Columnas › Dos."
    }
   ]
  },
  {
   "id": 3,
   "titulo": "Referencias y propiedades",
   "resumen": "Notas al pie y metadatos.",
   "pasos": [
    {
     "item": 6,
     "t": "Al <b>final del primer párrafo</b> agrega una <b>nota al pie</b> con el texto: <i>«Es importante diferenciar entre el rendimiento nominal y el rendimiento real.»</i>",
     "ayuda": "Cursor al final del primer párrafo del documento · Referencias › Insertar nota al pie."
    },
    {
     "item": 7,
     "t": "En las propiedades del documento, escribe tu <b>nombre completo</b> en <b>Comentarios</b> (el mismo que escribiste al iniciar el reto).",
     "ayuda": "Archivo › Información › Mostrar todas las propiedades › Comentarios."
    }
   ]
  },
  {
   "id": 4,
   "titulo": "Ilustraciones y elementos",
   "resumen": "Cuadros de texto, íconos y formas.",
   "pasos": [
    {
     "item": 8,
     "t": "Al <b>final del último párrafo del Tema 1</b> inserta el cuadro de texto <b>Cotización de filigrana</b> con: <i>«A mayor liquidez, menor rendimiento esperado; a mayor inmovilización del capital, mayor compensación exigida por el mercado.»</i> y la atribución <b>Principio fundamental de las finanzas</b>.",
     "ayuda": "Cursor al final del último párrafo del Tema 1 · Insertar › Cuadro de texto › Cotización de filigrana · reemplaza el texto de ejemplo."
    },
    {
     "item": 9,
     "t": "Al <b>inicio del primer párrafo</b> inserta un <b>ícono relacionado con dinero o finanzas</b>.",
     "ayuda": "Cursor antes de la primera palabra del primer párrafo · Insertar › Íconos › busca «dinero» o «finanzas»."
    },
    {
     "item": 10,
     "t": "Al <b>final del documento</b> inserta la forma <b>Pergamino: Horizontal</b> con el texto <b>FINANZAS</b>.",
     "ayuda": "Ctrl+Fin · Insertar › Formas › Cintas y estrellas › Pergamino: Horizontal · clic derecho › Agregar texto."
    }
   ]
  }
 ],
 "recursos": [
  {
   "archivo": "assets/docs/reto3/Reto3_Base.docx",
   "nombre": "Reto3_Base.docx",
   "desc": "Documento base del reto",
   "icono": "fluent:document-24-regular"
  }
 ]
}, {
  recursosB64: RECURSOS_B64_RETO3,
  puntos: CHECKER_RETO3.PUNTOS, nombreNivel: CHECKER_RETO3.NOMBRE_NIVEL,
  evaluar: CHECKER_RETO3.evaluar
});
