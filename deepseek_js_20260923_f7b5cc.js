// ============================================================
// IMPORTAR ENTRADAS DESDE IMÁGENES DE QR
// ============================================================
function renderImportar() {
  const eventos = DB.eventos();
  if (eventos.length === 0) {
    return `
      <div class="card p-8 text-center fade-in">
        <p class="text-white/60">Primero crea un evento.</p>
        <button onclick="ir('eventos')" class="btn-primary mt-4">Ir a Eventos</button>
      </div>`;
  }

  return `
    <div class="space-y-6 fade-in">
      <div>
        <h1 class="text-2xl font-bold">Importar entradas</h1>
        <p class="text-sm text-white/50 mt-1">
          Sube los PNGs de las entradas ya generadas. El sistema leerá el código QR
          de cada una y te pedirá completar el nombre y el tipo para registrarlas
          en el evento seleccionado.
        </p>
      </div>

      <div class="card p-5">
        <label class="label">Evento destino</label>
        <select class="input" id="import-evento">
          ${eventos.map(e => `<option value="${e.id}">${esc(e.nombre)} — ${fecha(e.fecha)}</option>`).join('')}
        </select>
      </div>

      <div class="card p-5">
        <label class="label">Imágenes de las entradas (PNG / JPG)</label>
        <input class="input" type="file" id="import-files" accept="image/*" multiple />
        <p class="mt-1 text-xs text-white/40">Puedes seleccionar varias a la vez. Espera mientras las procesa.</p>
        <div id="import-mensaje" class="mt-3"></div>
      </div>

      <div id="import-tabla"></div>
    </div>
  `;
}

function postImportar() {
  const input = $('#import-files');
  const eventoSelect = $('#import-evento');
  if (!input) return;

  let resultados = []; // [{ nombreArchivo, codigo, valido, nombre, tipo, miniatura }]

  const renderTabla = () => {
    const evento = DB.eventos().find(e => e.id === eventoSelect.value);
    const cont = $('#import-tabla');
    if (!cont) return;
    if (resultados.length === 0) { cont.innerHTML = ''; return; }

    const tipos = Object.keys(evento.precios);
    const validos = resultados.filter(r => r.valido).length;

    cont.innerHTML = `
      <div class="card p-5">
        <h2 class="mb-4 text-sm font-semibold uppercase tracking-wide text-white/70">
          Entradas detectadas (${resultados.length})
        </h2>
        <div class="space-y-3">
          ${resultados.map((r, i) => `
            <div class="entrada-edit" data-idx="${i}">
              <div class="flex gap-3 items-start">
                <img src="${r.miniatura}" class="w-20 h-20 object-cover rounded-lg border border-white/10 shrink-0" />
                <div class="flex-1 min-w-0">
                  <div class="text-xs text-white/50 mb-1">${esc(r.nombreArchivo)}</div>
                  <div class="text-[10px] font-mono text-white/40 truncate" title="${esc(r.codigo)}">${esc(r.codigo)}</div>
                  ${r.valido
                    ? `<div class="text-[10px] text-emerald-400 mt-1">✓ Firma válida</div>`
                    : `<div class="text-[10px] text-red-400 mt-1">✗ No se pudo leer o firma inválida — no se importará</div>`}
                </div>
              </div>
              ${r.valido ? `
                <div class="grid grid-cols-2 gap-2 mt-3">
                  <input class="input import-nombre" placeholder="Nombre del comprador" value="${esc(r.nombre || '')}" />
                  <select class="input import-tipo">
                    <option value="">— Tipo de entrada —</option>
                    ${tipos.map(t => `<option value="${esc(t)}" ${r.tipo === t ? 'selected' : ''}>${esc(t)}</option>`).join('')}
                  </select>
                </div>
              ` : ''}
            </div>
          `).join('')}
        </div>

        <div class="flex gap-3 mt-5">
          <button class="btn-primary flex-1" id="import-confirmar" ${validos === 0 ? 'disabled' : ''}>
            ✓ Importar ${validos} entrada(s)
          </button>
          <button class="btn-secondary" id="import-cancelar">Cancelar</button>
        </div>

        <p class="mt-3 text-xs text-white/40">
          💡 Las entradas sin nombre o sin tipo no se van a importar. Completa ambas antes de confirmar.
        </p>
      </div>
    `;

    // Actualizar valores al escribir / cambiar
    $$('[data-idx]').forEach(el => {
      const i = Number(el.dataset.idx);
      el.querySelector('.import-nombre')?.addEventListener('input', (e) => {
        resultados[i].nombre = e.target.value;
      });
      el.querySelector('.import-tipo')?.addEventListener('change', (e) => {
        resultados[i].tipo = e.target.value;
      });
    });

    $('#import-cancelar')?.addEventListener('click', () => {
      resultados = [];
      input.value = '';
      renderTabla();
    });

    $('#import-confirmar')?.addEventListener('click', () => {
      const r = confirmarImportacion(resultados, evento);
      resultados = [];
      input.value = '';
      renderTabla();
      const msg = document.createElement('div');
      msg.className = 'card p-5';
      msg.innerHTML = `
        <div class="ok-box">
          ✓ <b>${r.creadas}</b> entrada(s) importadas.
          ${r.duplicadas > 0 ? `<br>⚠ ${r.duplicadas} ya existían en la base y no se duplicaron.` : ''}
          ${r.sinDatos > 0 ? `<br>⚠ ${r.sinDatos} quedaron sin nombre o tipo y no se importaron.` : ''}
        </div>
      `;
      $('#import-tabla').appendChild(msg);
      setTimeout(() => msg.remove(), 8000);
    });
  };

  input.addEventListener('change', async (ev) => {
    const files = [...(ev.target.files || [])];
    if (files.length === 0) return;

    const msg = $('#import-mensaje');
    resultados = [];

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      msg.innerHTML = `<div class="ok-box">Procesando ${i + 1} de ${files.length}…</div>`;

      // Miniatura (independiente de si el QR se lee o no)
      const miniatura = await new Promise((res) => {
        const r = new FileReader();
        r.onload = () => res(r.result);
        r.readAsDataURL(file);
      });

      try {
        const texto = await escanearArchivo(file);
        const codigo = await verificar(texto);
        resultados.push({
          nombreArchivo: file.name,
          codigo: codigo || texto.slice(0, 60),
          valido: !!codigo,
          nombre: '',
          tipo: '',
          miniatura,
        });
      } catch (e) {
        console.warn('[import] no se pudo leer', file.name, e.message);
        resultados.push({
          nombreArchivo: file.name,
          codigo: '(no se pudo leer)',
          valido: false,
          miniatura,
        });
      }
    }

    msg.innerHTML = `<div class="ok-box">✓ Procesadas ${files.length} imagen(es). Completa nombre y tipo abajo.</div>`;
    setTimeout(() => { msg.innerHTML = ''; }, 4000);
    renderTabla();
  });

  eventoSelect.addEventListener('change', renderTabla);
}

function confirmarImportacion(resultados, evento) {
  const entradas = DB.entradas();
  const transacciones = DB.transacciones();
  const ahora = new Date().toISOString();

  let creadas = 0, duplicadas = 0, sinDatos = 0;

  for (const r of resultados) {
    if (!r.valido) continue;
    if (!r.nombre || !r.tipo) { sinDatos++; continue; }

    // Evitar duplicados por código
    if (entradas.find(e => e.codigo === r.codigo)) { duplicadas++; continue; }

    const precio = Number(evento.precios[r.tipo]) || 0;
    const txId = uid();

    transacciones.push({
      id: txId,
      eventoId: evento.id,
      compradorNombre: r.nombre,
      compradorCi: null,
      telefono: null,
      metodoPago: 'importado',
      notas: `Importado desde ${r.nombreArchivo}`,
      total: precio,
      anulado: false,
      createdAt: ahora,
    });

    entradas.push({
      id: uid(),
      transaccionId: txId,
      eventoId: evento.id,
      codigo: r.codigo,
      tipo: r.tipo,
      descripcion: null,
      precio,
      estado: 'valido',
      escaneadoAt: null,
      createdAt: ahora,
    });

    creadas++;
  }

  DB.guardarTransacciones(transacciones);
  DB.guardarEntradas(entradas);

  return { creadas, duplicadas, sinDatos };
}