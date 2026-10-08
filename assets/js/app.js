/*
 * Interface do Gerador de Declarações: navegação, formulários, pré-visualização e download.
 */
(function (NS) {
  'use strict';

  const F = NS.formatacao;
  const D = NS.docx;
  const P = NS.previa;
  const TIPOS = NS.tipos;
  const CHAVE_SESSAO = 'gerador-declaracoes:formularios';

  const ICONES = {
    adicionar: '<path d="M12 5v14M5 12h14"/>',
    remover: '<path d="M4 7h16M10 11v6M14 11v6M6 7l1 12a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-12M9 7V4h6v3"/>',
  };

  const $ = (id) => document.getElementById(id);
  const el = {
    inicio: $('tela-inicio'),
    formulario: $('tela-formulario'),
    form: $('formulario'),
    campos: $('campos'),
    sobretitulo: $('form-sobretitulo'),
    titulo: $('form-titulo'),
    alternador: $('alternador'),
    nota: $('nota-data'),
    resumo: $('resumo-erros'),
    gerar: $('botao-gerar'),
    limpar: $('botao-limpar'),
    previa: $('previa'),
    statusPrevia: $('status-previa'),
    destacar: $('destacar'),
    aviso: $('aviso'),
  };

  const estados = lerSessao();
  let tipo = null; // declaração aberta
  let estado = null; // dados do formulário aberto
  let validarAoDigitar = false;
  let temporizadorPrevia = 0;
  let temporizadorSessao = 0;
  let temporizadorAviso = 0;
  let geracaoPrevia = 0;
  let primeiraRota = true;

  // ------------------------------------------------------------------ utilitários

  function h(tag, atributos = {}, ...filhos) {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(atributos)) {
      if (v == null || v === false) continue;
      if (k === 'class') e.className = v;
      else if (k === 'text') e.textContent = v;
      else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
      else e.setAttribute(k, v === true ? '' : v);
    }
    for (const f of filhos.flat()) if (f != null && f !== false) e.append(f);
    return e;
  }

  function icone(nome) {
    const span = document.createElement('span');
    span.className = 'icone';
    span.innerHTML = `<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">${ICONES[nome]}</svg>`;
    return span;
  }

  function lerSessao() {
    try {
      const v = JSON.parse(sessionStorage.getItem(CHAVE_SESSAO) || '{}');
      return v && typeof v === 'object' ? v : {};
    } catch (e) {
      return {};
    }
  }

  function salvarSessao() {
    clearTimeout(temporizadorSessao);
    temporizadorSessao = setTimeout(() => {
      try { sessionStorage.setItem(CHAVE_SESSAO, JSON.stringify(estados)); } catch (e) { /* armazenamento indisponível */ }
    }, 300);
  }

  /** Estado do formulário do tipo, recuperando o que foi digitado nesta aba (se houver). */
  function estadoDe(t) {
    const inicial = t.estadoInicial();
    const salvo = estados[t.id];
    const e = { ...inicial };
    if (salvo && typeof salvo === 'object') {
      for (const k of Object.keys(inicial)) if (k !== 'aulas' && typeof salvo[k] === 'string') e[k] = salvo[k];
      if (t.tabelaAulas && Array.isArray(salvo.aulas) && salvo.aulas.length) {
        e.aulas = salvo.aulas.map((a) => {
          const aula = t.aulaVazia();
          for (const k of Object.keys(aula)) if (a && typeof a[k] === 'string') aula[k] = a[k];
          return aula;
        });
      }
    }
    estados[t.id] = e;
    return e;
  }

  // ------------------------------------------------------------------ navegação

  function rotear() {
    const rota = location.hash || '#/';
    const t = Object.values(TIPOS).find((x) => x.rota === rota);
    if (t) abrirFormulario(t);
    else abrirInicio();
    primeiraRota = false;
  }

  function focarTitulo(h1) {
    if (primeiraRota) return;
    h1.setAttribute('tabindex', '-1');
    h1.focus({ preventScroll: true });
  }

  function abrirInicio() {
    tipo = null;
    el.formulario.hidden = true;
    el.inicio.hidden = false;
    document.title = 'Gerador de declarações · Coren-SP Educação';
    window.scrollTo(0, 0);
    focarTitulo($('titulo-inicio'));
  }

  function abrirFormulario(t) {
    tipo = t;
    estado = estadoDe(t);
    validarAoDigitar = false;
    el.inicio.hidden = true;
    el.formulario.hidden = false;
    el.sobretitulo.textContent = t.sobretitulo;
    el.titulo.textContent = t.titulo;
    document.title = `${t.titulo} · Gerador de declarações`;
    el.alternador.hidden = t.perfil !== 'docente';
    for (const a of el.alternador.querySelectorAll('a')) {
      const ativo = a.dataset.tipo === t.id;
      a.classList.toggle('ativo', ativo);
      if (ativo) a.setAttribute('aria-current', 'page');
      else a.removeAttribute('aria-current');
    }
    el.nota.textContent = t.notaData;
    el.resumo.hidden = true;
    montarCampos();
    window.scrollTo(0, 0);
    focarTitulo(el.titulo);
    atualizarPrevia(true);
  }

  // ------------------------------------------------------------------ formulário

  function montarCampos() {
    el.campos.replaceChildren();
    for (const secao of tipo.secoes) {
      el.campos.append(h('fieldset', { class: 'secao' },
        h('legend', { text: secao.titulo }),
        h('div', { class: 'grade-campos' }, secao.campos.map(criarCampo))));
    }
    if (tipo.tabelaAulas) el.campos.append(criarSecaoAulas());
    atualizarDicas();
  }

  function criarCampo(def) {
    const id = `campo-${def.id}`;
    const tipoInput = { data: 'date', hora: 'time' }[def.tipo] || 'text';
    const input = h('input', {
      id,
      name: def.id,
      class: 'entrada',
      type: tipoInput,
      'data-chave': def.id,
      autocomplete: def.autocomplete || 'off',
      inputmode: def.tipo === 'cpf' ? 'numeric' : null,
      maxlength: def.tipo === 'cpf' ? '14' : tipoInput === 'text' ? '200' : null,
      placeholder: def.placeholder || null,
      spellcheck: tipoInput === 'text' ? 'false' : null,
      min: tipoInput === 'date' ? '1900-01-01' : null,
      max: tipoInput === 'date' ? '2100-12-31' : null,
      'aria-describedby': `${id}-dica ${id}-erro`,
      required: true,
    });
    input.value = estado[def.id] || '';
    input.addEventListener('input', () => {
      if (def.tipo === 'cpf') {
        const mascarado = F.mascararCpf(input.value);
        if (mascarado !== input.value) input.value = mascarado;
      }
      const anterior = estado[def.id];
      estado[def.id] = input.value;
      if (def.id === 'data' && tipo.tabelaAulas) sincronizarDatas(anterior, input.value);
      aoAlterar();
    });
    return h('div', { class: 'campo' + (def.metade ? ' campo-metade' : '') },
      h('label', { for: id }, def.rotulo, h('span', { class: 'obrigatorio', 'aria-hidden': 'true', text: ' *' })),
      input,
      h('p', { class: 'dica', id: `${id}-dica`, 'data-dica': def.id }),
      h('p', { class: 'mensagem-erro', id: `${id}-erro`, 'data-erro': def.id, hidden: true }));
  }

  function atualizarDicas() {
    for (const secao of tipo.secoes) {
      for (const def of secao.campos) {
        const dica = el.campos.querySelector(`[data-dica="${def.id}"]`);
        if (!dica) continue;
        const aviso = def.aviso ? def.aviso(estado) : null;
        const texto = aviso || (def.eco ? def.eco(estado) : def.dica) || '';
        dica.textContent = texto;
        dica.classList.toggle('dica-aviso', Boolean(aviso));
        dica.hidden = !texto;
      }
    }
  }

  function criarSecaoAulas() {
    return h('fieldset', { class: 'secao' },
      h('legend', { text: 'Aulas (tabela da declaração)' }),
      h('p', { class: 'secao-ajuda', text: 'Cada aula vira uma linha da tabela. Adicione quantas forem necessárias.' }),
      h('ol', { class: 'aulas', id: 'lista-aulas' }, estado.aulas.map(criarAula)),
      h('p', { class: 'mensagem-erro', 'data-erro': 'aulas', hidden: true }),
      h('button', { type: 'button', class: 'botao botao-secundario', id: 'botao-adicionar-aula', onclick: adicionarAula },
        icone('adicionar'), 'Adicionar aula'));
  }

  function criarAula(aula, i) {
    const campo = (chave, rotulo, tipoInput, extra = {}) => {
      const id = `aula-${i}-${chave}`;
      const input = h('input', {
        id,
        class: 'entrada',
        type: tipoInput,
        'data-chave': `aulas.${i}.${chave}`,
        'aria-describedby': `${id}-erro`,
        required: true,
        ...extra,
      });
      input.value = aula[chave] || '';
      input.addEventListener('input', () => {
        aula[chave] = input.value;
        aoAlterar();
      });
      return h('div', { class: `campo campo-${chave}` },
        h('label', { for: id, text: rotulo }),
        input,
        h('p', { class: 'mensagem-erro', id: `${id}-erro`, 'data-erro': `aulas.${i}.${chave}`, hidden: true }));
    };
    const varias = estado.aulas.length > 1;
    return h('li', { class: 'aula', 'aria-label': `Aula ${i + 1}` },
      h('div', { class: 'aula-topo' },
        h('span', { class: 'aula-numero', text: `Aula ${i + 1}` }),
        varias ? h('button', {
          type: 'button', class: 'botao-icone', title: 'Remover esta aula',
          'aria-label': `Remover aula ${i + 1}`, onclick: () => removerAula(i),
        }, icone('remover')) : null),
      h('div', { class: 'aula-grade' },
        campo('data', 'Data', 'date', { min: '1900-01-01', max: '2100-12-31' }),
        campo('inicio', 'Início', 'time'),
        campo('fim', 'Término', 'time'),
        campo('aula', 'Nome da aula', 'text', { maxlength: '300', autocomplete: 'off', spellcheck: 'true' })));
  }

  function redesenharAulas() {
    $('lista-aulas').replaceChildren(...estado.aulas.map(criarAula));
  }

  function adicionarAula() {
    const anterior = estado.aulas[estado.aulas.length - 1];
    const nova = tipo.aulaVazia((anterior && anterior.data) || estado.data || '');
    estado.aulas.push(nova);
    redesenharAulas();
    aoAlterar();
    const i = estado.aulas.length - 1;
    const foco = $(`aula-${i}-${nova.data ? 'inicio' : 'data'}`);
    if (foco) {
      foco.focus();
      foco.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }
  }

  function removerAula(i) {
    estado.aulas.splice(i, 1);
    redesenharAulas();
    aoAlterar();
    const foco = $(`aula-${Math.min(i, estado.aulas.length - 1)}-data`) || $('botao-adicionar-aula');
    if (foco) foco.focus();
  }

  /** Aulas sem data (ou com a data anterior da atividade) acompanham a data da atividade. */
  function sincronizarDatas(anterior, nova) {
    estado.aulas.forEach((a, i) => {
      if (!nova || (a.data && a.data !== anterior)) return;
      a.data = nova;
      const input = $(`aula-${i}-data`);
      if (input) input.value = nova;
    });
  }

  function aoAlterar() {
    salvarSessao();
    atualizarDicas();
    if (validarAoDigitar) exibirErros(tipo.validar(estado), false);
    clearTimeout(temporizadorPrevia);
    temporizadorPrevia = setTimeout(() => atualizarPrevia(false), 150);
  }

  function exibirErros(erros, focar) {
    for (const msg of el.form.querySelectorAll('[data-erro]')) {
      const texto = erros[msg.dataset.erro];
      msg.textContent = texto || '';
      msg.hidden = !texto;
    }
    for (const input of el.form.querySelectorAll('[data-chave]')) {
      if (erros[input.dataset.chave]) input.setAttribute('aria-invalid', 'true');
      else input.removeAttribute('aria-invalid');
    }
    const total = Object.keys(erros).length;
    el.resumo.hidden = total === 0;
    el.resumo.textContent = total === 1 ? 'Há 1 campo para corrigir.' : `Há ${total} campos para corrigir.`;
    if (focar && total) {
      const primeiro = el.form.querySelector('[aria-invalid="true"]');
      if (primeiro) {
        primeiro.focus({ preventScroll: true });
        primeiro.scrollIntoView({ block: 'center', behavior: 'smooth' });
      }
    }
  }

  // ------------------------------------------------------------------ geração do documento

  async function gerar(evento) {
    evento.preventDefault();
    if (!tipo || el.gerar.disabled) return;
    const t = tipo; // a geração continua mesmo se o usuário trocar de tela
    const erros = t.validar(estado);
    validarAoDigitar = true;
    exibirErros(erros, true);
    if (Object.keys(erros).length) return;

    const conteudo = Array.from(el.gerar.childNodes);
    el.gerar.disabled = true;
    el.gerar.textContent = 'Gerando…';
    try {
      const valores = t.valores(estado);
      const modelo = await D.carregar(t.modelo);
      const doc = D.lerXml(modelo.documentoXml);
      t.preencher(doc, valores, false);
      const blob = await D.montar(modelo, doc);
      const nome = `${F.nomeDeArquivo(t.nomeArquivo(valores))}.docx`;
      baixar(blob, nome);
      avisar(`Declaração gerada: ${nome}`, 'sucesso');
    } catch (erro) {
      console.error(erro);
      avisar(`Não foi possível gerar a declaração. ${erro.message}`, 'erro');
    } finally {
      el.gerar.disabled = false;
      el.gerar.replaceChildren(...conteudo);
    }
  }

  function baixar(blob, nome) {
    const url = URL.createObjectURL(blob);
    const a = h('a', { href: url, download: nome, hidden: true });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  function avisar(texto, tipoAviso) {
    el.aviso.textContent = texto;
    el.aviso.className = `aviso aviso-${tipoAviso}`;
    el.aviso.hidden = false;
    clearTimeout(temporizadorAviso);
    temporizadorAviso = setTimeout(() => { el.aviso.hidden = true; }, tipoAviso === 'erro' ? 12000 : 6000);
  }

  // ------------------------------------------------------------------ pré-visualização

  async function atualizarPrevia(inicial) {
    if (!tipo) return;
    const t = tipo;
    const numero = ++geracaoPrevia;
    if (inicial) el.previa.replaceChildren(h('p', { class: 'previa-mensagem', text: 'Carregando pré-visualização…' }));
    try {
      const modelo = await D.carregar(t.modelo);
      const recursos = await P.preparar(modelo);
      if (numero !== geracaoPrevia || tipo !== t) return;
      const doc = D.lerXml(modelo.documentoXml);
      t.preencher(doc, t.valores(estado), true);
      const paginas = P.renderizar(recursos, doc, el.previa);
      el.statusPrevia.textContent = paginas > 1 ? `(${paginas} páginas)` : '';
    } catch (erro) {
      console.error(erro);
      if (numero === geracaoPrevia) {
        el.previa.replaceChildren(h('p', { class: 'previa-mensagem previa-erro', text: `Não foi possível montar a pré-visualização. ${erro.message}` }));
      }
    }
  }

  // ------------------------------------------------------------------ início

  function iniciar() {
    if (!window.JSZip || !window.MODELOS_EMBUTIDOS) {
      document.querySelector('main').replaceChildren(h('p', { class: 'alerta-erro', text: 'Arquivos da aplicação ausentes (assets/vendor/jszip.min.js ou assets/js/modelos-embutidos.js).' }));
      return;
    }
    el.form.addEventListener('submit', gerar);
    el.limpar.addEventListener('click', () => {
      if (!tipo || !window.confirm('Limpar todos os campos deste formulário?')) return;
      estados[tipo.id] = tipo.estadoInicial();
      salvarSessao();
      abrirFormulario(tipo);
      avisar('Formulário limpo.', 'info');
    });
    el.destacar.addEventListener('change', () => {
      el.previa.classList.toggle('sem-destaque', !el.destacar.checked);
      $('legenda').hidden = !el.destacar.checked;
    });

    // Tamanho da pré-visualização: largura do painel ou página inteira (lembrado neste navegador).
    const botoesZoom = document.querySelectorAll('[data-zoom]');
    const aplicarZoom = (modo) => {
      el.previa.dataset.zoom = modo;
      botoesZoom.forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.zoom === modo)));
      P.ajustarEscala(el.previa);
    };
    let zoomSalvo = 'largura';
    try { zoomSalvo = localStorage.getItem('gerador-declaracoes:zoom') || 'largura'; } catch (e) { /* indisponível */ }
    aplicarZoom(zoomSalvo === 'pagina' ? 'pagina' : 'largura');
    botoesZoom.forEach((b) => b.addEventListener('click', () => {
      aplicarZoom(b.dataset.zoom);
      try { localStorage.setItem('gerador-declaracoes:zoom', b.dataset.zoom); } catch (e) { /* indisponível */ }
    }));

    let larguraPrevia = 0;
    new ResizeObserver(() => {
      if (el.previa.clientWidth !== larguraPrevia) {
        larguraPrevia = el.previa.clientWidth;
        P.ajustarEscala(el.previa);
      }
    }).observe(el.previa);
    window.addEventListener('resize', () => {
      if (el.previa.dataset.zoom === 'pagina') P.ajustarEscala(el.previa);
    });
    window.addEventListener('hashchange', rotear);
    rotear();
    // Carrega os modelos em segundo plano para agilizar a primeira pré-visualização.
    for (const t of Object.values(TIPOS)) D.carregar(t.modelo).then(P.preparar).catch(() => {});
  }

  iniciar();
})(window.Declaracoes = window.Declaracoes || {});
