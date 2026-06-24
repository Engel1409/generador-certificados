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

  // ✅ NORMALIZACIÓN ROBUSTA
  function normalizar(texto) {
    return texto
      .toString()
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/\s+/g, '');
  }

  // ✅ LECTURA CORREGIDA DEL EXCEL
  async function leerExcel(file){
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, {type:'array'});
    const hoja = wb.Sheets[wb.SheetNames[0]];

    const data = XLSX.utils.sheet_to_json(hoja, {
      range: 1, // 🔥 SALTA FILA INCORRECTA
      defval: '',
      raw: false
    });

    return data;
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

  // ✅ EXTRAER VARIABLES DEL WORD
  function extraerVariablesWord(buffer) {
    const zip = new PizZip(buffer);
    const xml = zip.files["word/document.xml"].asText();

    const regex = /{{(.*?)}}/g;
    const variables = new Set();
    let match;

    while ((match = regex.exec(xml)) !== null) {
      variables.add(normalizar(match[1]));
    }

    return [...variables];
  }

  btnAnalizar.addEventListener('click', async () => {
    if (!inputExcel.files[0] || !inputWord.files[0]) {
      alert('Selecciona Excel y Word');
      return;
    }

    try {
      filasExcel = await leerExcel(inputExcel.files[0]);
      plantillaBuffer = await inputWord.files[0].arrayBuffer();

      if (filasExcel.length === 0) throw new Error("Excel vacío");

      console.log("Columnas detectadas:", Object.keys(filasExcel[0]));

      const columnasOriginales = Object.keys(filasExcel[0]);
      const columnasExcel = columnasOriginales.map(c => normalizar(c));
      const variablesWord = extraerVariablesWord(plantillaBuffer);

      let html = `<strong>Validación Word vs Excel:</strong><br><br>`;

      let errores = 0;
      let correctos = 0;

      variablesWord.forEach(v => {
        const existe = columnasExcel.includes(v);

        if (existe) {
          html += `✅ ${v} (OK)<br>`;
          correctos++;
        } else {
          html += `❌ ${v} (NO EXISTE en Excel)<br>`;
          errores++;
        }
      });

      html += `<br><strong>Columnas Excel sin uso:</strong><br>`;

      columnasOriginales.forEach((col, i) => {
        if (!variablesWord.includes(columnasExcel[i])) {
          html += `⚠ ${col}<br>`;
        }
      });

      tablaColumnas.innerHTML = html;

      estadoGeneral.textContent = errores === 0
        ? `✅ Perfecto (${correctos} variables OK)`
        : `⚠ ${errores} campos faltantes`;

      infoFilas.textContent = `Registros: ${filasExcel.length}`;

      btnGenerar.disabled = errores > 0;

      seccionDiagnostico.hidden = false;
      seccionGenerar.hidden = false;

    } catch (e) {
      alert("Error: " + e.message);
    }
  });

  btnGenerar.addEventListener('click', async () => {

    const PizZipConstructor = window.PizZip || PizZip;
    const DocxConstructor = window.docxtemplater || docxtemplater;

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
        filaProcesada[normalizar(key)] = String(filaOriginal[key]).trim();
      }

      try {
        const zipInterno = new PizZipConstructor(plantillaBuffer);
        const doc = new DocxConstructor(zipInterno, {
          paragraphLoop: true,
          linebreaks: true,
          delimiters: { start: '{{', end: '}}' },
          nullGetter() { return ''; }
        });

        doc.render(filaProcesada);

        const blobDoc = doc.getZip().generate({
          type: 'blob',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });

        const nombre = sanitizeFilename(
          `${filaProcesada.nro || ''}_${filaProcesada.asegurado || ''}_${filaProcesada.poliza || ''}`
        );

        zipSalida.file(`${nombre || 'Certificado_'+i}.docx`, blobDoc);
        generados++;

      } catch (err) {
        errores.push(`Fila ${i+2}: ${err.message}`);
      }

      barraProgreso.style.width = Math.round(((i + 1) / filasExcel.length) * 100) + '%';
      estadoTexto.textContent = `Procesando ${i+1}/${filasExcel.length}`;
      await new Promise(r => setTimeout(r, 1));
    }

    if (generados > 0) {
      estadoTexto.textContent = 'Empaquetando…';

      const zipBlob = await zipSalida.generateAsync({type:'blob'});
      const url = URL.createObjectURL(zipBlob);

      const a = document.createElement('a');
      a.href = url;
      a.download = "DOCUMENTOS.zip";
      a.click();

      URL.revokeObjectURL(url);
      estadoTexto.textContent = `✅ ${generados} generados`;
    }

    if (errores.length) {
      logEl.innerHTML = errores.join('<br>');
    }

    btnGenerar.disabled = false;
  });

})(); 