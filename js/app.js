(function(){
  const inputExcel = document.getElementById('inputExcel');
  const inputWord = document.getElementById('inputWord');
  const btnAnalizar = document.getElementById('btnAnalizar');
  const btnGenerar = document.getElementById('btnGenerar');
  const btnLimpiar = document.getElementById('btnLimpiar');
  const seccionDiagnostico = document.getElementById('seccionDiagnostico');
  const seccionGenerar = document.getElementById('seccionGenerar');
  const tablaColumnas = document.getElementById('tablaColumnas');
  const estadoGeneral = document.getElementById('estadoGeneral');
  const infoFilas = document.getElementById('infoFilas');
  const barraProgreso = document.getElementById('barraProgreso');
  const estadoTexto = document.getElementById('estadoTexto');
  const logEl = document.getElementById('log');

  let plantillaBuffer = null;
  let filasExcel = [];

  function sanitizeFilename(nombre){
    return nombre.replace(/[\\/:*?"<>|]/g,'_').trim();
  }

  async function leerExcel(file){
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, {type:'array'});
    const hoja = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(hoja, {defval:''});
  }

  function resetTodo(){
    inputExcel.value = '';
    inputWord.value = '';
    filasExcel = [];
    plantillaBuffer = null;

    tablaColumnas.innerHTML = '';
    estadoGeneral.textContent = '';
    infoFilas.textContent = '';
    barraProgreso.style.width = '0%';
    estadoTexto.textContent = '';
    logEl.innerHTML = '';

    seccionDiagnostico.hidden = true;
    seccionGenerar.hidden = true;
    btnGenerar.disabled = false;
  }

  btnLimpiar.addEventListener('click', resetTodo);

  btnAnalizar.addEventListener('click', async () => {
    if (!inputExcel.files[0] || !inputWord.files[0]) {
      alert('Por favor, selecciona ambos archivos (.xlsx y .docx)');
      return;
    }

    try {
      filasExcel = await leerExcel(inputExcel.files[0]);
      plantillaBuffer = await inputWord.files[0].arrayBuffer();

      if (filasExcel.length === 0) throw new Error("El Excel no contiene registros.");

      const columnas = Object.keys(filasExcel[0]);
      tablaColumnas.innerHTML = `<strong>Columnas encontradas listas para inyectar al Word:</strong><br>` + 
        columnas.map(c => `• ${c}`).join('<br>');

      infoFilas.textContent = `Se procesará un lote de ${filasExcel.length} certificados.`;
      estadoGeneral.textContent = "✓ Motores listos. Estructura cargada correctamente en la memoria de la PC.";
      
      seccionDiagnostico.hidden = false;
      seccionGenerar.hidden = false;
    } catch (e) {
      alert("Error leyendo insumos: " + e.message);
    }
  });

  btnGenerar.addEventListener('click', async () => {
    // Verificación explícita de existencia del objeto en la ventana del navegador
    const PizZipConstructor = window.PizZip || (typeof PizZip !== 'undefined' ? PizZip : null);
    const DocxConstructor = window.docxtemplater || (typeof docxtemplater !== 'undefined' ? docxtemplater : null);

    if (!PizZipConstructor || !DocxConstructor) {
      alert("Error crítico: Las librerías de empaquetado no se han inicializado en el navegador. Usa la extensión Live Server de VSC para saltar restricciones locales.");
      return;
    }

    btnGenerar.disabled = true;
    logEl.textContent = '';
    barraProgreso.style.width = '0%';
    
    const zipSalida = new JSZip();
    let generados = 0;
    const errores = [];

    for (let i = 0; i < filasExcel.length; i++) {
      const filaOriginal = filasExcel[i];
      const filaProcesada = { fecha: new Date().toLocaleDateString('es-PE') };
      
      for (const key in filaOriginal) {
        // Mapea tanto claves originales como en minúsculas por seguridad
        filaProcesada[key.trim().toLowerCase()] = String(filaOriginal[key]).trim();
        filaProcesada[key] = String(filaOriginal[key]).trim();
      }

      try {
        const zipInterno = new PizZipConstructor(plantillaBuffer);
        const doc = new DocxConstructor(zipInterno, {
          paragraphLoop: true,
          linebreaks: true,
          delimiters: { start: '{{', end: '}}' },
          nullGetter() {
            return '';
          }
        });

        doc.render(filaProcesada);
        console.log("Fila procesada:", filaProcesada);

        const blobDoc = doc.getZip().generate({
          type: 'blob',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });

        const nro = filaProcesada.nro || '';
        const asegurado = filaProcesada.asegurado || '';
        const poliza = filaProcesada.poliza || '';
        let nombreArchivo = nro ? `${nro}_${asegurado}_${poliza}` : `${asegurado}_${poliza}`;
        nombreArchivo = sanitizeFilename(nombreArchivo) || `Certificado_${i+1}`;

        zipSalida.file(`${nombreArchivo}.docx`, blobDoc);
        generados++;
      } catch (err) {
        console.log("ERROR COMPLETO", err);

        let detalle = err.message;

        if (err.properties) {
          if (err.properties.errors) {
            err.properties.errors.forEach((e, idx) => {
              console.log("ERROR", idx + 1);
              console.log(e.properties);
            });

            detalle = err.properties.errors
              .map(e => e.properties?.explanation || e.properties?.id || e.name)
              .join(" | ");

          } else if (err.properties.explanation) {
            detalle = err.properties.explanation;
          }
        }

        errores.push(`Fila ${i + 2}: ${detalle}`);
      }

      barraProgreso.style.width = Math.round(((i + 1) / filasExcel.length) * 100) + '%';
      estadoTexto.textContent = `Procesando: ${i + 1} de ${filasExcel.length}...`;
      await new Promise(r => setTimeout(r, 1));
    }

    if (generados > 0) {
      estadoTexto.textContent = 'Empaquetando lote final...';
      const zipBlob = await zipSalida.generateAsync({type: 'blob'});
      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `DOCUMENTOS.zip`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      estadoTexto.textContent = `✓ ¡Éxito total! ${generados} archivos descargados localmente en tu ZIP.`;
    }

    if (errores.length) {
      logEl.innerHTML = '<strong>Alertas:</strong><br>' + errores.join('<br>');
    }
    btnGenerar.disabled = false;
  });
})();