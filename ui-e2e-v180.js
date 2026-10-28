/* Prueba end-to-end de la interfaz v1.9.0 (Chrome real por CDP, sin instalar nada).
   Uso: node ui-e2e-v180.js <puerto-cdp> <url-base> <carpeta-capturas> <archivo-credenciales>
   Requiere el servidor y Chrome con --remote-debugging-port ya iniciados. */
const fs = require("fs");
const path = require("path");

const CDP = Number(process.argv[2] || 9222);
const BASE = process.argv[3] || "http://127.0.0.1:8390";
const CAPTURAS = process.argv[4] || "";
const CREDENCIALES = process.argv[5] || "";

const resultados = [];
const erroresNavegador = [];
const RUINOSO = /Failed to load resource.*(400|401|403|404|409)/;

const esperar = (ms) => new Promise((r) => setTimeout(r, ms));

async function jsonDe(url) {
  const r = await fetch(url);
  return r.json();
}

async function conectar() {
  let pagina = null;
  for (let i = 0; i < 60 && !pagina; i++) {
    try {
      const lista = await jsonDe(`http://127.0.0.1:${CDP}/json/list`);
      pagina = lista.find((t) => t.type === "page");
    } catch (e) { /* chrome aún no abre el puerto */ }
    if (!pagina) await esperar(250);
  }
  if (!pagina) throw new Error("no apareció la pestaña del navegador");
  const ws = new WebSocket(pagina.webSocketDebuggerUrl);
  await new Promise((resolver, rechazar) => {
    ws.addEventListener("open", resolver, {once: true});
    ws.addEventListener("error", rechazar, {once: true});
  });
  let siguiente = 0;
  const pendientes = new Map();
  ws.addEventListener("message", (evento) => {
    const mensaje = JSON.parse(evento.data);
    if (mensaje.id && pendientes.has(mensaje.id)) {
      const {resolver, rechazar} = pendientes.get(mensaje.id);
      pendientes.delete(mensaje.id);
      if (mensaje.error) rechazar(new Error(JSON.stringify(mensaje.error)));
      else resolver(mensaje.result);
      return;
    }
    if (mensaje.method === "Runtime.exceptionThrown") {
      const d = mensaje.params.exceptionDetails;
      erroresNavegador.push("EXCEPCIÓN: " + (d.exception && d.exception.description
        ? d.exception.description : d.text));
    }
    if (mensaje.method === "Runtime.consoleAPICalled" && mensaje.params.type === "error") {
      erroresNavegador.push("CONSOLE: "
        + mensaje.params.args.map((a) => a.value ?? a.description).join(" "));
    }
    if (mensaje.method === "Log.entryAdded" && mensaje.params.entry.level === "error") {
      const e = mensaje.params.entry;
      const texto = e.text + " " + (e.url || "");
      if (!RUINOSO.test(texto)) erroresNavegador.push("LOG: " + texto);
    }
  });
  const enviar = (method, params = {}) => new Promise((resolver, rechazar) => {
    const id = ++siguiente;
    pendientes.set(id, {resolver, rechazar});
    ws.send(JSON.stringify({id, method, params}));
  });
  const js = async (cuerpo) => {
    const respuesta = await enviar("Runtime.evaluate", {
      expression: `(async () => { ${cuerpo} })()`,
      awaitPromise: true,
      returnByValue: true,
      userGesture: true,
    });
    if (respuesta.exceptionDetails) {
      const d = respuesta.exceptionDetails;
      throw new Error("JS: " + (d.exception && d.exception.description ? d.exception.description
        : d.text));
    }
    return respuesta.result.value;
  };
  return {enviar, js, ws};
}

async function esperarEnPagina(js, condicion, descripcion, ms = 12000) {
  const limite = Date.now() + ms;
  while (Date.now() < limite) {
    if (await js(`return !!(${condicion});`)) return true;
    await esperar(150);
  }
  throw new Error("tiempo agotado esperando: " + descripcion);
}

function anotar(nombre, ok, detalle) {
  resultados.push({prueba: nombre, ok: !!ok, detalle: detalle === undefined ? "" : detalle});
  console.log(`${ok ? "OK  " : "FALLA"}  ${nombre}${detalle ? "  · " + detalle : ""}`);
}

async function captura(enviar, nombre) {
  if (!CAPTURAS) return;
  const {data} = await enviar("Page.captureScreenshot", {format: "png"});
  fs.mkdirSync(CAPTURAS, {recursive: true});
  fs.writeFileSync(path.join(CAPTURAS, nombre), Buffer.from(data, "base64"));
}

function credenciales() {
  const texto = fs.readFileSync(CREDENCIALES, "utf-8");
  const usuario = texto.match(/usuario\s*:\s*(\S+)/)[1];
  const clave = texto.match(/contrase[ñn]a\s*:\s*(\S+)/)[1];
  return {usuario, clave};
}

const ESC = {type: "keyDown", key: "Escape", code: "Escape",
             windowsVirtualKeyCode: 27, nativeVirtualKeyCode: 27};

(async () => {
  const cliente = await conectar();
  const {enviar, js} = cliente;
  try {
    await enviar("Runtime.enable");
    await enviar("Log.enable");
    await enviar("Page.enable");
    await enviar("Network.enable");
    await enviar("Network.clearBrowserCookies");
    await enviar("Emulation.setDeviceMetricsOverride",
                 {width: 1440, height: 900, deviceScaleFactor: 1, mobile: false});
    await enviar("Page.navigate", {url: BASE + "/"});
    await esperar(400);
    await esperarEnPagina(js,
      `document.readyState === 'complete' && !!document.getElementById('form-login')`,
      "carga de la página");
    await js(`localStorage.clear(); return true;`);     // primera visita de verdad
    await enviar("Page.reload", {ignoreCache: true});
    await esperarEnPagina(js, `document.readyState === 'complete'`, "recarga limpia");
    anotar("la página carga", true, BASE);

    // ---------------------------------------------------------------- login
    anotar("pantalla de acceso visible",
      await js(`return !document.getElementById('login').hidden;`));
    anotar("el acceso explica dónde está la contraseña inicial",
      await js(`return document.querySelector('#login')
        .textContent.includes('CREDENCIALES_INICIALES');`));
    await captura(enviar, "pantalla_login.png");

    const {usuario, clave} = credenciales();
    await js(`
      const $ = (id) => document.getElementById(id);
      $('l-usuario').value = ${JSON.stringify(usuario)};
      $('l-password').value = ${JSON.stringify(clave)};
      $('form-login').dispatchEvent(new Event('submit', {bubbles: true, cancelable: true}));
      return true;`);
    await esperarEnPagina(js, `!document.getElementById('app').hidden`, "entrada al sistema");
    anotar("el formulario de acceso funciona", true, usuario);

    // ---------------------------------------------------------------- primeros pasos
    await esperarEnPagina(js, `document.querySelectorAll('#lista .lote').length > 0`,
                          "lista de lotes");
    anotar("los primeros pasos aparecen la primera vez",
      await js(`return !document.getElementById('primeros-pasos').hidden;`));
    await captura(enviar, "pantalla_primeros_pasos.png");
    await js(`document.getElementById('btn-entendido').click(); return true;`);
    anotar("«Entendido» los oculta",
      await js(`return document.getElementById('primeros-pasos').hidden;`));

    const kpi = await js(`return {total: document.getElementById('k-total').textContent,
                          cob: document.getElementById('k-cob').textContent,
                          lotes: document.querySelectorAll('#lista .lote').length};`);
    anotar("la lista y los indicadores se pintan",
      Number(kpi.total) > 100 && kpi.lotes > 50, JSON.stringify(kpi));

    await enviar("Page.reload", {ignoreCache: true});
    await esperarEnPagina(js, `!document.getElementById('app').hidden`, "sesión recordada");
    anotar("al recargar sigue la sesión (sin volver al login)", true);
    anotar("los primeros pasos no vuelven a molestar",
      await js(`return document.getElementById('primeros-pasos').hidden;`));
    await captura(enviar, "pantalla_inventario.png");

    // ---------------------------------------------------------------- ayuda
    await js(`document.getElementById('btn-ayuda').click(); return true;`);
    await esperarEnPagina(js, `!document.getElementById('modal-ayuda').hidden`, "ayuda abierta");
    await esperarEnPagina(js,
      `document.getElementById('ayuda-version').textContent.trim().length > 3`,
      "versión en la ayuda", 6000);
    const ayuda = await js(`return {secciones:
        document.querySelectorAll('#modal-ayuda .ayuda-seccion').length,
        version: document.getElementById('ayuda-version').textContent.trim()};`);
    anotar("la ayuda tiene sus secciones y la versión",
      ayuda.secciones >= 7 && ayuda.version.length > 3, JSON.stringify(ayuda));
    await captura(enviar, "pantalla_ayuda.png");
    await enviar("Input.dispatchKeyEvent", ESC);
    await esperar(250);
    anotar("Esc cierra la ayuda",
      await js(`return document.getElementById('modal-ayuda').hidden;`));

    // ---------------------------------------------------------------- depuración
    await js(`setTimeout(() => { throw new Error("prueba de depuracion E2E"); }, 0);
              return true;`);
    await esperar(1500);
    const depuracion = await js(`
      const r = await fetch('/api/eventos?limite=80');
      const j = await r.json();
      const lista = j.eventos || [];
      return {hay: lista.some((e) => e.evento === 'navegador.error'
                                     && String(e.mensaje || '').includes('prueba de depuracion E2E')),
              total: lista.length};`);
    anotar("un error de JavaScript de la pantalla llega al registro del servidor",
      depuracion.hay, JSON.stringify(depuracion));

    // ---------------------------------------------------------------- búsqueda aproximada
    const caso = await js(`
      const lista = await (await fetch('/api/inventario?tipo=mp')).json();
      for (const f of lista.filas.slice(0, 40)) {
        for (const palabra of String(f.descripcion || '').split(' ')) {
          if (palabra.length < 6) continue;
          const letras = palabra.split('');
          let i = 0;
          while (i < letras.length - 2 && letras[i] === letras[i + 1]) i++;
          if (i >= letras.length - 2) continue;
          [letras[i], letras[i + 1]] = [letras[i + 1], letras[i]];
          const typo = letras.join('');
          if (typo === palabra) continue;
          const j = await (await fetch('/api/inventario?tipo=mp&q=' + encodeURIComponent(typo))).json();
          if (j.difuso && j.filas.length && (j.sugerencias || []).length) {
            return {typo, palabra, filas: j.filas.length, sugerencias: j.sugerencias.length};
          }
        }
      }
      return {fallo: true};`);
    anotar("el servidor encuentra parecidos cuando no hay coincidencia exacta",
      !caso.fallo, caso.fallo ? "sin caso de prueba" : `${caso.typo} ~ ${caso.palabra}`);

    await js(`
      const q = document.getElementById('q');
      q.value = ${JSON.stringify(caso.typo || "")};
      q.dispatchEvent(new Event('input', {bubbles: true}));
      return true;`);
    const difusa = await js(`
      const info = document.getElementById('q-info');
      const limite = Date.now() + 8000;
      while (Date.now() < limite && !info.textContent.includes('aproximados')) {
        await new Promise((r) => setTimeout(r, 150));
      }
      return {info: info.textContent.trim(),
              chips: [...info.querySelectorAll('[data-sugerencia]')].map((b) => b.textContent),
              filas: document.querySelectorAll('#lista .lote').length};`);
    anotar("la pantalla dice que los resultados son aproximados y sugiere",
      difusa.filas > 0 && difusa.info.includes("aproximados") && difusa.chips.length > 0,
      `${difusa.filas} fila(s) · ${difusa.chips.length} sugerencia(s)`);
    await captura(enviar, "pantalla_busqueda_aproximada.png");
    await js(`const b = document.querySelector('#q-info [data-sugerencia]');
              if (b) b.click(); return true;`);
    await esperar(600);
    anotar("al tocar la sugerencia se busca el nombre exacto",
      await js(`return document.getElementById('q').value.length > 0;`));
    await js(`const q = document.getElementById('q'); q.value = '';
              q.dispatchEvent(new Event('input', {bubbles: true})); return true;`);
    await esperarEnPagina(js, `document.querySelectorAll('#lista .lote').length > 50`,
                          "volver a la lista completa");

    // ---------------------------------------------------------------- ubicación
    const contar = () => js(`
      const r = await fetch('/api/bitacora?tipo=mp&limite=200');
      return (await r.json()).movimientos.length;`);
    const antes = await contar();
    const noValida = await js(`
      const input = document.querySelector('#lista input.u');
      window.__fila = Number(input.dataset.fila);
      input.value = 'RACK-INVENTADO';
      input.dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter', bubbles: true}));
      const limite = Date.now() + 6000;
      while (Date.now() < limite && !document.querySelector('.toast.err')) {
        await new Promise((r) => setTimeout(r, 100));
      }
      const t = document.querySelector('.toast.err');
      return {toast: t ? t.textContent : '', marcado: input.classList.contains('err')};`);
    anotar("una ubicación fuera del catálogo se rechaza con explicación",
      noValida.toast.includes("catálogo") && noValida.marcado, noValida.toast);
    anotar("el intento inválido no ensucia la bitácora", (await contar()) === antes);

    const destino = await js(`
      const lote = await (await fetch('/api/lote?fila=' + window.__fila + '&tipo=mp')).json();
      const r = await (await fetch('/api/ubicaciones?tipo=mp&limite=30')).json();
      return (r.sugerencias || []).map((s) => s.codigo)
        .filter((c) => c !== lote.lote.ubicacion)[0] || null;`);
    anotar("el servidor ofrece ubicaciones válidas para sugerir", !!destino, String(destino));
    const guardado = await js(`
      const input = document.querySelector('#lista input.u');
      input.value = ${JSON.stringify(destino)};
      input.dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter', bubbles: true}));
      const limite = Date.now() + 8000;
      while (Date.now() < limite && !document.querySelector('.toast.ok')) {
        await new Promise((r) => setTimeout(r, 100));
      }
      return {toast: (document.querySelector('.toast.ok') || {}).textContent || '',
              valor: input.value, porOmision: input.defaultValue};`);
    anotar("una ubicación válida se guarda",
      guardado.valor === destino && guardado.toast.includes("Ubicación"), guardado.toast);
    anotar("el campo queda sincronizado (no se guarda dos veces)",
      guardado.porOmision === guardado.valor, `${guardado.porOmision} = ${guardado.valor}`);
    const trasGuardar = await contar();
    await js(`document.querySelector('#lista input.u')
      .dispatchEvent(new FocusEvent('blur', {bubbles: false})); return true;`);
    await esperar(700);
    anotar("salir del campo otra vez no duplica el movimiento", (await contar()) === trasGuardar);

    // ---------------------------------------------------------------- parcialidad con «5 kg»
    const grande = await js(`
      const j = await (await fetch('/api/inventario?tipo=mp')).json();
      const f = j.filas.filter((x) => typeof x.cantidad === 'number' && x.cantidad >= 100)[0];
      return f ? {fila: f.fila, cantidad: f.cantidad, ubicacion: f.ubicacion} : null;`);
    if (!grande) throw new Error("no hay ningún lote con cantidad >= 100 para la parcialidad");
    const destino2 = await js(`
      const r = await (await fetch('/api/ubicaciones?tipo=mp&limite=30')).json();
      return (r.sugerencias || []).map((s) => s.codigo)
        .filter((c) => c !== ${JSON.stringify(grande && grande.ubicacion)})[0] || null;`);
    await js(`document.querySelector('#lista .lote[data-fila="' +  ${JSON.stringify(grande.fila)}
      + '"] [data-mover]').click(); return true;`);
    await esperarEnPagina(js, `!document.getElementById('m-form').hidden`, "modal de movimiento");
    const parcial = await js(`
      const c = document.getElementById('m-cantidad');
      c.value = '5 kg';
      c.dispatchEvent(new Event('input', {bubbles: true}));
      const d = document.getElementById('m-destino');
      d.value = ${JSON.stringify(destino2)};
      d.dispatchEvent(new Event('input', {bubbles: true}));
      await new Promise((r) => setTimeout(r, 200));
      return {nota: document.getElementById('m-nota-parcial').textContent,
              oculta: document.getElementById('m-nota-parcial').hidden};`);
    anotar("la parcialidad se explica y acepta «5 kg»",
      !parcial.oculta && parcial.nota.includes("Parcialidad"), parcial.nota.slice(0, 100));
    const antesDeMover = await js(`
      return {cantidad: document.getElementById('m-cantidad').value,
              destino: document.getElementById('m-destino').value,
              nota: document.getElementById('m-nota-parcial').textContent.slice(0, 60)};`);
    await js(`document.getElementById('btn-registrar').click(); return true;`);
    await esperarEnPagina(js,
      `[...document.querySelectorAll('.toast')].some(
        (t) => t.textContent.includes('Parcialidad registrada')
            || t.textContent.includes('Movimiento registrado')
            || t.textContent.includes('no puedes mover'))`,
      "movimiento registrado", 10000);
    const registro = await js(`
      const j = await (await fetch('/api/bitacora?tipo=mp&limite=1')).json();
      return j.movimientos[0];`);
    anotar("el movimiento queda como parcialidad de 5 kg",
      registro.accion === "PARCIAL" && Number(registro.cantidad) === 5,
      `${registro.accion} ${registro.cantidad} ${registro.unidad} · antes: ${JSON.stringify(antesDeMover)}`);
    await captura(enviar, "pantalla_mover_parcial.png");

    // ---------------------------------------------------------------- salida a producción
    const otra = await js(`
      const j = await (await fetch('/api/inventario?tipo=mp')).json();
      const f = j.filas.filter((x) => typeof x.cantidad === 'number' && x.cantidad >= 10)[0];
      return f ? {fila: f.fila, descripcion: f.descripcion} : null;`);
    if (!otra) throw new Error("no hay un segundo lote con cantidad >= 10 para la salida");
    await js(`document.querySelector('#lista .lote[data-fila="' + ${JSON.stringify(otra.fila)}
      + '"] [data-mover]').click(); return true;`);
    await esperarEnPagina(js, `!document.getElementById('m-form').hidden`, "modal de movimiento");
    const produccion = await js(`
      const chip = [...document.querySelectorAll('#modal-mover .chips .chip')]
        .find((b) => b.textContent.includes('Producción'));
      chip.click();
      await new Promise((r) => setTimeout(r, 150));
      return {destino: document.getElementById('m-destino').value,
              aviso: document.getElementById('m-aviso-prod').textContent,
              boton: document.getElementById('btn-registrar').textContent};`);
    anotar("el atajo de producción explica que el material sale del almacén",
      produccion.destino === "PRODUCCION" && produccion.aviso.includes("SALE"),
      produccion.boton);
    await js(`document.getElementById('m-cantidad').value = '';
              document.getElementById('btn-registrar').click(); return true;`);
    await esperarEnPagina(js,
      `[...document.querySelectorAll('.toast.ok')].some(
        (t) => t.textContent.includes('Salida a producción'))`, "salida registrada");
    const salida = await js(`
      const j = await (await fetch('/api/bitacora?tipo=mp&limite=1')).json();
      return j.movimientos[0];`);
    anotar("la salida a producción queda en la bitácora",
      salida.accion === "SALIDA", `${salida.accion} ${salida.cantidad} ${salida.unidad}`);
    await captura(enviar, "pantalla_produccion.png");

    // ---------------------------------------------------------------- alta tolerante
    await js(`document.getElementById('btn-alta').click(); return true;`);
    await esperarEnPagina(js, `!document.getElementById('modal-alta').hidden`, "modal de alta");
    await js(`const e = document.getElementById('a-erp');
              e.value = 'MIGH-XX999';
              e.dispatchEvent(new Event('input', {bubbles: true})); return true;`);
    await esperarEnPagina(js,
      `document.getElementById('a-catalogo').textContent.includes('no está en el catálogo')`,
      "aviso de clave inexistente");
    anotar("una clave que no existe en el ERP se avisa al escribir", true);

    const aceptada = await js(`
      const set = (id, v) => { document.getElementById(id).value = v; };
      set('a-desc', 'PRODUCTO INVENTADO');
      set('a-lote', 'L-XX-1');
      set('a-cant', '5');
      set('a-prov', 'NADIE');
      set('a-cad', '2028-01-01');
      document.getElementById('btn-guardar-alta').click();
      const limite = Date.now() + 10000;
      while (Date.now() < limite && !document.getElementById('modal-alta').hidden) {
        await new Promise((r) => setTimeout(r, 120));
      }
      await new Promise((r) => setTimeout(r, 1200));
      const fila = await (await fetch('/api/inventario?tipo=mp&q=L-XX-1')).json();
      return {cerrado: document.getElementById('modal-alta').hidden,
              filas: fila.filas.length,
              descripcion: (fila.filas[0] || {}).descripcion,
              error: document.getElementById('a-error').textContent,
              avisos: [...document.querySelectorAll('.toast')].map((t) => t.textContent)};`);
    anotar("se acepta un insumo que no está en el catálogo (con aviso, sin bloquear)",
      aceptada.cerrado && aceptada.filas === 1 && aceptada.descripcion === "PRODUCTO INVENTADO"
      && aceptada.avisos.some((x) => x.includes("no está en el catálogo")),
      JSON.stringify({cerrado: aceptada.cerrado, filas: aceptada.filas,
                      error: aceptada.error, avisos: aceptada.avisos.slice(0, 3)}));
    await captura(enviar, "pantalla_alta_aceptada.png");

    // ---------------------------------------------------------------- descripción → clave
    await js(`document.getElementById('btn-alta').click(); return true;`);
    await esperarEnPagina(js, `!document.getElementById('modal-alta').hidden`, "modal de alta");
    const porDesc = await js(`
      const d = document.getElementById('a-desc');
      d.value = 'ACAIVERRY EXT';
      d.dispatchEvent(new Event('input', {bubbles: true}));
      const limite = Date.now() + 9000;
      while (Date.now() < limite && !document.getElementById('a-erp').value) {
        await new Promise((r) => setTimeout(r, 150));
      }
      return {erp: document.getElementById('a-erp').value,
              unidad: document.getElementById('a-unidad').value,
              prov: document.getElementById('a-prov').value,
              cat: document.getElementById('a-cat').value,
              destino: document.getElementById('a-destino-natur').textContent,
              pista: document.getElementById('a-desc-estado').textContent,
              ficha: document.getElementById('a-ficha').textContent};`);
    anotar("escribir la descripción llena sola la clave del ERP y sus datos",
      porDesc.erp === "MIGH-MP001" && porDesc.unidad === "kg"
      && porDesc.prov === "ACAIVERRY EXT" && porDesc.cat === "Excipiente"
      && porDesc.destino.includes("Inventario_MP.xlsx")
      && porDesc.ficha.includes("ACAIVERRY EXT"),
      JSON.stringify({erp: porDesc.erp, unidad: porDesc.unidad, cat: porDesc.cat}));
    anotar("la pantalla dice de dónde salieron los datos",
      porDesc.pista.includes("catálogo del ERP"), porDesc.pista.slice(0, 90));

    // ---------------------------------------------------------------- clave → datos
    const mep = await js(`
      const e = document.getElementById('a-erp');
      e.value = 'MIGH-MT001';
      e.dispatchEvent(new Event('input', {bubbles: true}));
      const limite = Date.now() + 9000;
      while (Date.now() < limite
             && !document.getElementById('a-desc').value.includes('LATA 401x411')) {
        await new Promise((r) => setTimeout(r, 150));
      }
      return {erp: document.getElementById('a-erp').value,
              desc: document.getElementById('a-desc').value,
              unidad: document.getElementById('a-unidad').value,
              prov: document.getElementById('a-prov').value,
              cat: document.getElementById('a-cat').value,
              destino: document.getElementById('a-destino-natur').textContent,
              ficha: document.getElementById('a-ficha').textContent,
              fichaVisible: !document.getElementById('a-ficha').hidden};`);
    anotar("escribir la clave del ERP llena descripción, unidad, categoría y proveedor",
      mep.desc.includes("LATA 401x411") && mep.unidad === "pza"
      && mep.prov === "CROW LA VILLA" && mep.cat === "Empaque"
      && mep.destino.includes("Inventario_MEP.xlsx"),
      JSON.stringify({desc: mep.desc, unidad: mep.unidad, prov: mep.prov, cat: mep.cat}));
    anotar("la ficha muestra todos los datos del insumo (clave, máximo, mínimo, fabricante)",
      mep.fichaVisible && ["Clave", "Máximo", "Mínimo", "Fabricante"].every((x) => mep.ficha.includes(x)),
      mep.ficha.replace(/\s+/g, " ").slice(0, 140));

    const sinTipo = await js(`
      const e = document.getElementById('a-erp');
      e.value = 'MIGH-MH-001';
      e.dispatchEvent(new Event('input', {bubbles: true}));
      const limite = Date.now() + 9000;
      while (Date.now() < limite
             && !document.getElementById('a-desc').value.includes('GUANTE')) {
        await new Promise((r) => setTimeout(r, 150));
      }
      const chip = document.querySelector('#a-naturaleza .chip.activo');
      return {desc: document.getElementById('a-desc').value,
              ficha: document.getElementById('a-ficha').textContent,
              chip: chip && chip.dataset ? chip.dataset.natur : null};`);
    anotar("un insumo sin tipo en el ERP se llena pero NO cambia de inventario solo",
      sinTipo.desc.includes("GUANTE") && sinTipo.ficha.includes("GUANTE")
      && sinTipo.chip === "mep", JSON.stringify({desc: sinTipo.desc, chip: sinTipo.chip}));
    await captura(enviar, "pantalla_alta.png");
    await enviar("Input.dispatchKeyEvent", ESC);
    await esperar(400);

    // ---------------------------------------------------------------- selección y borrado
    await js(`document.querySelector('#lista .sel').click(); return true;`);
    await esperar(300);
    const sel = await js(`return {boton: document.getElementById('btn-borrar-sel').textContent,
                                   marcados: document.querySelectorAll('#lista .sel:checked').length};`);
    anotar("marcar un renglón habilita el borrado con su cuenta",
      sel.marcados === 1 && sel.boton.includes("Borrar 1 renglón"), sel.boton);
    await js(`document.querySelector('.selector-inv [data-inv="mep"]').click(); return true;`);
    await esperarEnPagina(js,
      `document.querySelector('.selector-inv [data-inv="mep"]').classList.contains('activo')
       && document.getElementById('archivo').textContent.includes('empaque')`,
      "inventario de empaque");
    await esperar(700);
    const trasCambio = await js(`return {
      marcados: document.querySelectorAll('#lista .sel:checked').length,
      seleccion: (typeof estado !== 'undefined' && estado.seleccion) ? estado.seleccion.size : -1,
      boton: document.getElementById('btn-borrar-sel').textContent,
      filas: document.querySelectorAll('#lista .lote').length,
      archivo: document.getElementById('archivo').textContent};`);
    anotar("cambiar de inventario limpia la selección (no borra en el libro equivocado)",
      trasCambio.marcados === 0 && trasCambio.seleccion <= 0
      && trasCambio.boton.includes("Borrar seleccionados") && trasCambio.filas > 5,
      JSON.stringify(trasCambio));
    await captura(enviar, "pantalla_mep.png");
    await js(`document.querySelector('.selector-inv [data-inv="mp"]').click(); return true;`);
    await esperarEnPagina(js, `document.querySelectorAll('#lista .lote').length > 50`, "volver a MP");

    // ---------------------------------------------------------------- bitácora
    await js(`document.querySelector('nav.tabs [data-vista="bitacora"]').click(); return true;`);
    await esperarEnPagina(js, `document.querySelectorAll('#movimientos .mov').length > 0`,
                          "movimientos");
    anotar("la bitácora muestra los movimientos",
      await js(`return document.querySelectorAll('#movimientos .mov').length > 0;`));
    await captura(enviar, "pantalla_bitacora.png");
    await js(`document.querySelector('nav.tabs [data-vista="inventario"]').click(); return true;`);

    // ---------------------------------------------------------------- QR / Barcode E2E (v1.9.0)
    // 1. Verificar #modal-escaner, #btn-esc-manual y el fallback a prompt sin cámara/HTTPS
    const pruebaModalEscaner = await js(`
      const promptOriginal = window.prompt;
      const getUserMediaOriginal = navigator.mediaDevices && navigator.mediaDevices.getUserMedia;
      let capturadoManual = null;
      window.prompt = () => 'QR-MANUAL-PROMPT';
      if (navigator.mediaDevices) {
        navigator.mediaDevices.getUserMedia = () => new Promise(() => {});
      }
      abrirEscaner((codigo) => { capturadoManual = codigo; });
      const modalVisible = !document.getElementById('modal-escaner').hidden;
      const titulo = document.querySelector('#modal-escaner h2').textContent;
      document.getElementById('btn-esc-manual').click();
      const modalCerrado = document.getElementById('modal-escaner').hidden;
      if (navigator.mediaDevices && getUserMediaOriginal) {
        navigator.mediaDevices.getUserMedia = getUserMediaOriginal;
      }
      window.prompt = promptOriginal;
      return {modalVisible, titulo, capturadoManual, modalCerrado};`);
    anotar("el modal de cámara abre y permite captura manual (#btn-esc-manual)",
      pruebaModalEscaner.modalVisible && pruebaModalEscaner.modalCerrado
      && pruebaModalEscaner.capturadoManual === "QR-MANUAL-PROMPT",
      JSON.stringify(pruebaModalEscaner));

    // 2. Stub de window.abrirEscaner = (cb) => cb(...) para probar alta, escaneo de material y rack
    const altaQr = await js(`
      window.__abrirEscanerOriginal = window.abrirEscaner;
      window.__proximoQr = 'QR-E2E-1';
      window.abrirEscaner = (cb) => cb(window.__proximoQr);

      document.getElementById('btn-alta').click();
      document.getElementById('btn-esc-qr').click();
      const qrEnCampo = document.getElementById('a-qr').value;

      window.__proximoQr = 'MPA-A-010';
      document.getElementById('btn-esc-ubi').click();
      const ubiEnCampo = document.getElementById('a-ubi').value;

      document.getElementById('a-desc').value = 'INSUMO QR E2E';
      document.getElementById('a-lote').value = 'L-QR-E2E-1';
      document.getElementById('a-cant').value = '25';
      document.getElementById('a-prov').value = 'PROVEEDOR QR';
      document.getElementById('a-cad').value = '2028-06-01';
      document.getElementById('btn-guardar-alta').click();

      const limite = Date.now() + 10000;
      while (Date.now() < limite && !document.getElementById('modal-alta').hidden) {
        await new Promise((r) => setTimeout(r, 120));
      }
      await new Promise((r) => setTimeout(r, 600));
      const resQr = await (await fetch('/api/qr?codigo=QR-E2E-1')).json();
      return {qrEnCampo, ubiEnCampo, encontrado: resQr.encontrado,
              fila: resQr.fila, lote: resQr.renglon && resQr.renglon.lote};`);
    anotar("alta con #btn-esc-qr y #btn-esc-ubi guarda el qr_id y /api/qr lo resuelve",
      altaQr.qrEnCampo === "QR-E2E-1" && altaQr.ubiEnCampo === "MPA-A-010"
      && altaQr.encontrado && altaQr.lote === "L-QR-E2E-1",
      JSON.stringify(altaQr));

    // 3. Escanear material desde la barra (#btn-escanear-material) -> abre #modal-mover -> #btn-esc-destino -> mueve
    const escaneoMover = await js(`
      window.__proximoQr = 'QR-E2E-1';
      document.getElementById('btn-escanear-material').click();
      const limite = Date.now() + 8000;
      while (Date.now() < limite
             && (document.getElementById('modal-mover').hidden
                 || document.getElementById('m-form').hidden
                 || !document.getElementById('m-detalle').textContent.includes('INSUMO QR E2E'))) {
        await new Promise((r) => setTimeout(r, 100));
      }
      const descModal = document.getElementById('m-detalle').textContent;
      window.__proximoQr = 'MPA-A-015';
      document.getElementById('btn-esc-destino').click();
      const destinoModal = document.getElementById('m-destino').value;
      document.getElementById('btn-registrar').click();
      const limite2 = Date.now() + 8000;
      while (Date.now() < limite2 && !document.getElementById('modal-mover').hidden) {
        await new Promise((r) => setTimeout(r, 100));
      }
      const postQr = await (await fetch('/api/qr?codigo=QR-E2E-1')).json();
      return {descModal, destinoModal, nuevaUbi: postQr.renglon && postQr.renglon.ubicacion};`);
    anotar("escanear material abre #modal-mover, #btn-esc-destino llena el rack y registra el traslado",
      escaneoMover.descModal.includes("INSUMO QR E2E")
      && escaneoMover.destinoModal === "MPA-A-015"
      && escaneoMover.nuevaUbi === "MPA-A-015",
      JSON.stringify(escaneoMover));

    // 4. Código inexistente con confirm=true abre #modal-alta con #a-qr precargado
    const escaneoInexistente = await js(`
      const confirmOriginal = window.confirm;
      window.confirm = () => true;
      window.__proximoQr = 'QR-NO-REGISTRADO-99';
      document.getElementById('btn-escanear-material').click();
      const limite = Date.now() + 8000;
      while (Date.now() < limite && document.getElementById('modal-alta').hidden) {
        await new Promise((r) => setTimeout(r, 100));
      }
      const abierto = !document.getElementById('modal-alta').hidden;
      const qrPrecargado = document.getElementById('a-qr').value;
      window.confirm = confirmOriginal;
      window.abrirEscaner = window.__abrirEscanerOriginal;
      return {abierto, qrPrecargado};`);
    anotar("escanear un código inexistente ofrece darlo de alta con #a-qr precargado",
      escaneoInexistente.abierto && escaneoInexistente.qrPrecargado === "QR-NO-REGISTRADO-99",
      JSON.stringify(escaneoInexistente));
    await enviar("Input.dispatchKeyEvent", ESC);
    await js(`const q = document.getElementById('q'); q.value = '';
              q.dispatchEvent(new Event('input', {bubbles: true})); return true;`);
    await esperarEnPagina(js, `document.querySelectorAll('#lista .lote').length > 50`,
                          "lista completa tras pruebas QR");

    // ---------------------------------------------------------------- móvil y escritorio
    await enviar("Emulation.setDeviceMetricsOverride",
                 {width: 390, height: 844, deviceScaleFactor: 2, mobile: true});
    await esperar(500);
    const movil = await js(`return {
      desborde: Math.max(document.documentElement.scrollWidth, document.body.scrollWidth)
                - window.innerWidth,
      cabecera: getComputedStyle(document.querySelector('.cabecera-lista')).display,
      ayuda: !!document.getElementById('btn-ayuda')};`);
    anotar("en teléfono no hay desborde horizontal y la ayuda está a mano",
      movil.desborde <= 1 && movil.cabecera === "none" && movil.ayuda, JSON.stringify(movil));
    await captura(enviar, "pantalla_movil.png");
    await enviar("Emulation.setDeviceMetricsOverride",
                 {width: 1280, height: 900, deviceScaleFactor: 1, mobile: false});
    await esperar(500);
    const escritorio = await js(`return {
      desborde: Math.max(document.documentElement.scrollWidth, document.body.scrollWidth)
                - window.innerWidth,
      cabecera: getComputedStyle(document.querySelector('.cabecera-lista')).display,
      filas: document.querySelectorAll('#lista .lote').length};`);
    anotar("en escritorio la tabla usa la cabecera y no desborda",
      escritorio.desborde <= 1 && escritorio.cabecera === "grid" && escritorio.filas > 50,
      JSON.stringify(escritorio));
    await captura(enviar, "pantalla_principal.png");

    // El error provocado a propósito para probar el reporte no cuenta como falla real.
    const erroresReales = erroresNavegador.filter(
      (e) => !e.includes("prueba de depuracion E2E"));
    anotar("no hubo errores de JavaScript en el navegador", erroresReales.length === 0,
      erroresReales.slice(0, 3).join(" | "));
  } catch (e) {
    anotar("la prueba terminó sin excepciones", false, e.message);
    console.error(e.stack || e);
  } finally {
    const fallas = resultados.filter((r) => !r.ok);
    const resumen = {total: resultados.length, fallas: fallas.length, resultados, erroresNavegador};
    if (CAPTURAS) {
      fs.mkdirSync(CAPTURAS, {recursive: true});
      fs.writeFileSync(path.join(CAPTURAS, "e2e-v180.json"),
                       JSON.stringify(resumen, null, 2));
    }
    console.log(`\n${resultados.length - fallas.length}/${resultados.length} pruebas OK`);
    for (const falla of fallas) console.log("FALLÓ: " + falla.prueba + " · " + falla.detalle);
    process.exit(fallas.length ? 1 : 0);
  }
})();
