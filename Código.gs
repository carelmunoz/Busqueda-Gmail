// Archivo: Code.gs


/**
* Función que se ejecuta cuando se abre la hoja de cálculo.
* Crea un menú personalizado para abrir el panel lateral.
*/
function onOpen() {
 // Obtiene la interfaz de usuario de la hoja de cálculo activa.
 const ui = SpreadsheetApp.getUi();
 // Crea un menú personalizado llamado "Gmail App".
 ui.createMenu('Gmail App')
     .addItem('Abrir panel lateral', 'showSidebar') // Añade un elemento para abrir el panel lateral.
     .addToUi(); // Añade el menú a la interfaz de usuario.
}


/**
* Muestra el panel lateral personalizado.
*/
function showSidebar() {
 // Carga el contenido del archivo HTML llamado 'Sidebar'.
 const html = HtmlService.createHtmlOutputFromFile('Sidebar')
     .setTitle('Gmail Extractor') // Establece el título del panel lateral.
     .setWidth(300); // Establece el ancho del panel lateral.
 // Muestra el panel lateral en la interfaz de usuario de la hoja de cálculo.
 SpreadsheetApp.getUi().showSidebar(html);
}


/**
* Función que busca correos de Gmail basándose en los criterios proporcionados.
*
* @param {Object} options - Objeto que contiene las opciones de búsqueda.
* @param {string} options.labels - Etiquetas de Gmail a buscar (separadas por coma).
* @param {string} options.keywords - Palabras clave a buscar en los correos.
* @param {string} options.startDate - Fecha de inicio para la búsqueda (formato YYYY-MM-DD).
* @param {string} options.endDate - Fecha de fin para la búsqueda (formato YYYY-MM-DD).
* @param {boolean} options.includeAttachments - Si se deben incluir correos con archivos adjuntos.
*/
function searchEmails(options) {
 const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 // Limpia cualquier contenido anterior en la hoja activa (excepto la primera fila para encabezados).
 sheet.clearContents();
 // Escribe los encabezados en la primera fila.
 sheet.appendRow(['Fecha', 'De', 'Asunto', 'Cuerpo del Correo', 'ID del Mensaje']);


 // Obtiene las propiedades de usuario para almacenar el estado de detención.
 const userProperties = PropertiesService.getUserProperties();
 // Restablece la bandera de detención al inicio de una nueva búsqueda.
 userProperties.setProperty('stopSearch', 'false');


 let queryString = '';


 // Construye la cadena de consulta de Gmail.
 if (options.labels) {
   // Si hay etiquetas, las añade a la consulta.
   options.labels.split(',').forEach(label => {
     queryString += `label:${label.trim()} `;
   });
 }
 if (options.keywords) {
   // Si hay palabras clave, las añade a la consulta.
   queryString += `${options.keywords.trim()} `;
 }
 if (options.startDate) {
   // Si hay una fecha de inicio, la añade a la consulta.
   queryString += `after:${options.startDate} `;
 }
 if (options.endDate) {
   // Si hay una fecha de fin, la añade a la consulta.
   queryString += `before:${options.endDate} `;
 }
 if (options.includeAttachments) {
   // Si se incluyen adjuntos, añade el filtro.
   queryString += `has:attachment `;
 }


 // Quita los espacios extra del final de la cadena de consulta.
 queryString = queryString.trim();


 Logger.log('Query String: ' + queryString);


 let start = 0; // Inicio para paginación de correos.
 const batchSize = 100; // Número de correos a procesar por lote.
 let totalProcessed = 0; // Contador de correos procesados.


 while (true) {
   // Comprueba si la bandera de detención ha sido activada por el usuario.
   if (userProperties.getProperty('stopSearch') === 'true') {
     Logger.log('Búsqueda detenida por el usuario.');
     // Devuelve el total procesado si se detiene.
     return { status: 'stopped', count: totalProcessed };
   }


   // Busca hilos de Gmail con la cadena de consulta, limitando por lote.
   const threads = GmailApp.search(queryString, start, batchSize);


   if (threads.length === 0) {
     // Si no hay más hilos, sale del bucle.
     break;
   }


   // Itera sobre cada hilo encontrado.
   threads.forEach(thread => {
     // Obtiene todos los mensajes dentro del hilo.
     const messages = thread.getMessages();
     // Itera sobre cada mensaje.
     messages.forEach(message => {
       // Incrementa el contador de correos procesados.
       totalProcessed++;
       // Obtiene los datos del mensaje.
       const date = message.getDate();
       const sender = message.getFrom();
       const subject = message.getSubject();
       // Obtiene un fragmento del cuerpo del mensaje para evitar procesar cuerpos muy largos.
       const bodySnippet = message.getPlainBody().substring(0, 500); // Primeros 500 caracteres.
       const messageId = message.getId();


       // Añade una nueva fila a la hoja de cálculo con los datos del correo.
       sheet.appendRow([date, sender, subject, bodySnippet, messageId]);


       // Envía el recuento actual al panel lateral para actualizar la UI.
       // Esto solo funciona si el script se ejecuta de forma asíncrona (con google.script.run).
       // Para actualizaciones en tiempo real durante un bucle largo, se necesitan funciones separadas o más complejas.
       // Por ahora, se envía el contador al final del bucle de mensajes.
       if (totalProcessed % 10 === 0) { // Actualiza cada 10 correos procesados para no sobrecargar.
           HtmlService.createHtmlOutput(`<script>document.getElementById('resultsCounter').textContent = '${totalProcessed} correos extraídos.';</script>`);
       }
     });
   });


   start += batchSize; // Avanza al siguiente lote de correos.
   // Pequeña pausa para evitar exceder las cuotas de la API de Gmail.
   Utilities.sleep(1000); // Espera 1 segundo.
 }


 // Devuelve el estado y el recuento final una vez que la búsqueda ha terminado.
 return { status: 'completed', count: totalProcessed };
}


/**
* Función para detener la búsqueda de correos.
* Establece una propiedad de usuario 'stopSearch' a 'true'.
*/
function stopEmailSearch() {
 // Obtiene las propiedades de usuario.
 const userProperties = PropertiesService.getUserProperties();
 // Establece la bandera de detención a 'true'.
 userProperties.setProperty('stopSearch', 'true');
 Logger.log('Bandera de detención establecida a true.');
}


/**
* Función para obtener las etiquetas de Gmail del usuario.
* @returns {string[]} Un array de nombres de etiquetas.
*/
function getGmailLabels() {
 // Obtiene todas las etiquetas del usuario de Gmail.
 const labels = GmailApp.getUserLabels();
 // Mapea el array de objetos Label a un array de nombres de etiquetas.
 return labels.map(label => label.getName());
}





