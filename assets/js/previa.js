/*
 * Pré-visualização: desenha em HTML o documento Word já preenchido (página A4, cabeçalho,
 * rodapé, texto e tabela) a partir do XML do próprio modelo.
 * É uma aproximação para conferência; o arquivo .docx gerado é que preserva o modelo original.
 */
(function (NS) {
  'use strict';

  const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const NS_V = 'urn:schemas-microsoft-com:vml';
  const NS_W10 = 'urn:schemas-microsoft-com:office:word';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const NS_WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  const PX = 96 / 72; // pixels CSS por ponto
  const EMU = 12700; // EMUs por ponto

  // Altura da linha simples no Word, como fração do tamanho da fonte.
  const ALTURA_LINHA = { 'arial': 1.149, 'arial black': 1.41, 'times new roman': 1.149, 'calibri': 1.2207, 'cambria': 1.172 };
  const PILHAS = {
    'arial': "Arial, 'Liberation Sans', Helvetica, sans-serif",
    'arial black': "'Arial Black', Arial, sans-serif",
    'times new roman': "'Times New Roman', 'Liberation Serif', Times, serif",
    'calibri': "Calibri, Carlito, 'Segoe UI', Arial, sans-serif",
  };
  const MARCAS = { '\uE000': 'campo-preenchido', '\uE002': 'campo-pendente', '\uE004': 'campo-manual' };
  const TITULOS = { '\uE000': 'Preenchido pelo formulário', '\uE002': 'Ainda não preenchido', '\uE004': 'Preenchimento manual posterior' };
  const RE_MARCAS = /([\uE000\uE002\uE004])([^\uE000-\uE005]*)[\uE001\uE003\uE005]/g;
  const REALCES = { yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', lightGray: '#c0c0c0', darkGray: '#808080' };

  // ------------------------------------------------------------------ utilitários XML

  function filho(el, nome, ns = NS_W) {
    if (!el) return null;
    for (const c of el.children) if (c.localName === nome && c.namespaceURI === ns) return c;
    return null;
  }

  function filhos(el, nome, ns = NS_W) {
    return el ? Array.from(el.children).filter((c) => c.localName === nome && c.namespaceURI === ns) : [];
  }

  function atributo(el, nome, ns = NS_W) {
    if (!el) return null;
    const v = ns ? el.getAttributeNS(ns, nome) : el.getAttribute(nome);
    return v === '' ? null : v;
  }

  const numero = (v, padrao = null) => (v == null || v === '' || isNaN(Number(v)) ? padrao : Number(v));

  function twips(el, nome) {
    const v = numero(atributo(el, nome));
    return v == null ? null : v / 20;
  }

  function ligado(el) {
    const v = atributo(el, 'val');
    return v == null || !['0', 'false', 'off', 'none'].includes(v);
  }

  const pilhaDaFonte = (f) => PILHAS[(f || '').toLowerCase()] || `"${String(f || 'Times New Roman').replace(/"/g, '')}", serif`;
  const alturaLinha = (f) => ALTURA_LINHA[(f || '').toLowerCase()] || 1.17;

  // ------------------------------------------------------------------ recursos do modelo

  async function relacoes(zip, parte) {
    const pasta = parte.replace(/[^/]+$/, '');
    const arquivo = zip.file(`${pasta}_rels/${parte.slice(pasta.length)}.rels`);
    const mapa = new Map();
    if (!arquivo) return mapa;
    const doc = NS.docx.lerXml(await arquivo.async('string'));
    for (const r of doc.getElementsByTagNameNS(NS_REL, 'Relationship')) {
      if (r.getAttribute('TargetMode') === 'External') continue;
      mapa.set(r.getAttribute('Id'), { tipo: r.getAttribute('Type') || '', alvo: resolverCaminho(pasta, r.getAttribute('Target') || '') });
    }
    return mapa;
  }

  function resolverCaminho(pasta, alvo) {
    if (alvo.startsWith('/')) return alvo.slice(1);
    const saida = [];
    for (const p of (pasta + alvo).split('/')) {
      if (p === '..') saida.pop();
      else if (p && p !== '.') saida.push(p);
    }
    return saida.join('/');
  }

  function tipoDeImagem(caminho) {
    const ext = caminho.split('.').pop().toLowerCase();
    return { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml' }[ext] || '';
  }

  async function imagens(zip, rels) {
    const mapa = new Map();
    for (const [id, rel] of rels) {
      const tipo = tipoDeImagem(rel.alvo);
      const arquivo = zip.file(rel.alvo);
      if (!/\/image$/.test(rel.tipo) || !tipo || !arquivo) continue;
      mapa.set(id, URL.createObjectURL(new Blob([await arquivo.async('uint8array')], { type: tipo })));
    }
    return mapa;
  }

  function lerEstilos(doc) {
    const est = { rPadrao: null, pPadrao: null, paragrafo: new Map(), padrao: null };
    if (!doc) return est;
    const raiz = doc.documentElement;
    const padroes = filho(raiz, 'docDefaults');
    est.rPadrao = filho(filho(padroes, 'rPrDefault'), 'rPr');
    est.pPadrao = filho(filho(padroes, 'pPrDefault'), 'pPr');
    for (const s of filhos(raiz, 'style')) {
      if (atributo(s, 'type') !== 'paragraph') continue;
      const id = atributo(s, 'styleId');
      est.paragrafo.set(id, { base: atributo(filho(s, 'basedOn'), 'val'), pPr: filho(s, 'pPr'), rPr: filho(s, 'rPr') });
      if (atributo(s, 'default') === '1') est.padrao = id;
    }
    return est;
  }

  const preparados = new WeakMap();

  /** Lê estilos, cabeçalhos, rodapés e imagens do modelo (uma vez por modelo). */
  function preparar(modelo) {
    if (!preparados.has(modelo)) {
      preparados.set(modelo, (async () => {
        const zip = modelo.zip;
        const xml = async (caminho) => {
          const f = zip.file(caminho);
          return f ? NS.docx.lerXml(await f.async('string')) : null;
        };
        const relsDocumento = await relacoes(zip, 'word/document.xml');
        const partes = new Map();
        for (const [id, rel] of relsDocumento) {
          if (!/\/(header|footer)$/.test(rel.tipo)) continue;
          const doc = await xml(rel.alvo);
          if (doc) partes.set(id, { doc, imagens: await imagens(zip, await relacoes(zip, rel.alvo)) });
        }
        return {
          estilos: lerEstilos(await xml('word/styles.xml')),
          partes,
          imagensDocumento: await imagens(zip, relsDocumento),
        };
      })());
    }
    return preparados.get(modelo);
  }

  // ------------------------------------------------------------------ propriedades

  function cadeiaDeEstilos(est, id) {
    const lista = [];
    const vistos = new Set();
    let atual = id && est.paragrafo.has(id) ? id : est.padrao;
    while (atual && est.paragrafo.has(atual) && !vistos.has(atual)) {
      vistos.add(atual);
      const s = est.paragrafo.get(atual);
      lista.unshift(s);
      atual = s.base;
    }
    return lista;
  }

  function lerRun(rPr, r) {
    if (!rPr) return r;
    for (const c of rPr.children) {
      if (c.namespaceURI !== NS_W) continue;
      switch (c.localName) {
        case 'rFonts': { const f = atributo(c, 'ascii') || atributo(c, 'hAnsi'); if (f) r.fonte = f; break; }
        case 'sz': { const v = numero(atributo(c, 'val')); if (v) r.tamanho = v / 2; break; }
        case 'b': r.negrito = ligado(c); break;
        case 'i': r.italico = ligado(c); break;
        case 'u': r.sublinhado = ligado(c); break;
        case 'strike': r.tachado = ligado(c); break;
        case 'caps': r.maiusculas = ligado(c); break;
        case 'smallCaps': r.versalete = ligado(c); break;
        case 'vanish': r.oculto = ligado(c); break;
        case 'color': { const v = atributo(c, 'val'); r.cor = v && v !== 'auto' ? '#' + v : null; break; }
        case 'spacing': r.espacamento = numero(atributo(c, 'val'), 0) / 20; break;
        case 'vertAlign': r.posicao = atributo(c, 'val'); break;
        case 'highlight': { const v = atributo(c, 'val'); r.realce = v && v !== 'none' ? v : null; break; }
        default: break;
      }
    }
    return r;
  }

  function lerParagrafo(pPr, p) {
    if (!pPr) return p;
    for (const c of pPr.children) {
      if (c.namespaceURI !== NS_W) continue;
      if (c.localName === 'jc') {
        p.alinhamento = atributo(c, 'val');
      } else if (c.localName === 'ind') {
        const esq = twips(c, 'left') ?? twips(c, 'start');
        const dir = twips(c, 'right') ?? twips(c, 'end');
        const primeira = twips(c, 'firstLine');
        const deslocada = twips(c, 'hanging');
        if (esq != null) p.recuoEsq = esq;
        if (dir != null) p.recuoDir = dir;
        if (primeira != null) p.primeiraLinha = primeira;
        if (deslocada != null) p.primeiraLinha = -deslocada;
      } else if (c.localName === 'spacing') {
        const antes = twips(c, 'before');
        const depois = twips(c, 'after');
        const linha = numero(atributo(c, 'line'));
        if (antes != null) p.antes = antes;
        if (depois != null) p.depois = depois;
        if (linha != null) { p.linha = linha; p.regra = atributo(c, 'lineRule') || 'auto'; }
      }
    }
    return p;
  }

  function propriedades(est, pPr) {
    const p = { alinhamento: 'left', recuoEsq: 0, recuoDir: 0, primeiraLinha: 0, antes: 0, depois: 0, linha: 240, regra: 'auto' };
    const r = { fonte: 'Times New Roman', tamanho: 10, negrito: false, italico: false, cor: null, espacamento: 0 };
    lerParagrafo(est.pPadrao, p);
    lerRun(est.rPadrao, r);
    for (const s of cadeiaDeEstilos(est, atributo(filho(pPr, 'pStyle'), 'val'))) {
      lerParagrafo(s.pPr, p);
      lerRun(s.rPr, r);
    }
    lerParagrafo(pPr, p);
    return { p, r };
  }

  // ------------------------------------------------------------------ texto

  function estiloDoTexto(el, r) {
    const s = el.style;
    s.fontFamily = pilhaDaFonte(r.fonte);
    s.fontSize = r.tamanho + 'pt';
    s.fontWeight = r.negrito ? '700' : '400';
    s.fontStyle = r.italico ? 'italic' : 'normal';
    const decoracao = [r.sublinhado && 'underline', r.tachado && 'line-through'].filter(Boolean).join(' ');
    if (decoracao) s.textDecoration = decoracao;
    if (r.cor) s.color = r.cor;
    if (r.espacamento) s.letterSpacing = r.espacamento + 'pt';
    if (r.maiusculas) s.textTransform = 'uppercase';
    if (r.versalete) s.fontVariant = 'small-caps';
    if (r.realce) s.backgroundColor = REALCES[r.realce] || r.realce;
    if (r.posicao === 'superscript' || r.posicao === 'subscript') {
      s.verticalAlign = r.posicao === 'superscript' ? 'super' : 'sub';
      s.fontSize = r.tamanho * 0.65 + 'pt';
    }
  }

  function trecho(texto, r) {
    const s = document.createElement('span');
    s.textContent = texto;
    estiloDoTexto(s, r);
    return s;
  }

  /** Acrescenta o texto, transformando os delimitadores da pré-visualização em destaques. */
  function adicionarTexto(destino, texto, r) {
    if (!texto) return;
    let ultimo = 0;
    let m;
    RE_MARCAS.lastIndex = 0;
    while ((m = RE_MARCAS.exec(texto))) {
      if (m.index > ultimo) destino.appendChild(trecho(texto.slice(ultimo, m.index), r));
      const t = trecho(m[2], r);
      t.classList.add(MARCAS[m[1]]);
      t.title = TITULOS[m[1]];
      destino.appendChild(t);
      ultimo = RE_MARCAS.lastIndex;
    }
    if (ultimo < texto.length) destino.appendChild(trecho(texto.slice(ultimo), r));
  }

  function removerEspacosFinais(el) {
    const textos = [];
    const percorrer = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    for (let n = percorrer.nextNode(); n; n = percorrer.nextNode()) {
      if (!n.parentElement.closest('.w-flutuante')) textos.push(n);
    }
    for (let i = textos.length - 1; i >= 0; i--) {
      textos[i].data = textos[i].data.replace(/[  ]+$/, '');
      if (textos[i].data.replace(/\u200B/g, '')) break;
    }
  }

  function runsDoParagrafo(p) {
    const lista = [];
    const conteiners = ['hyperlink', 'smartTag', 'ins', 'fldSimple', 'customXml', 'sdt', 'sdtContent', 'moveTo', 'dir', 'bdo'];
    (function percorrer(el) {
      for (const c of el.children) {
        if (c.namespaceURI === NS_W) {
          if (c.localName === 'r') lista.push(c);
          else if (conteiners.includes(c.localName)) percorrer(c);
        } else if (c.namespaceURI === NS_MC && c.localName === 'AlternateContent') {
          const alt = filho(c, 'Fallback', NS_MC) || filho(c, 'Choice', NS_MC);
          if (alt) percorrer(alt);
        }
      }
    })(p);
    return lista;
  }

  function textoDoRun(run) {
    let t = '';
    for (const c of run.children) if (c.namespaceURI === NS_W && c.localName === 't') t += c.textContent;
    return t;
  }

  function renderizarRun(run, r, ctx, destino) {
    if (r.oculto) return;
    for (const c of run.children) {
      if (c.namespaceURI === NS_W) {
        switch (c.localName) {
          case 't': adicionarTexto(destino, c.textContent, r); break;
          case 'tab': {
            const tab = document.createElement('span');
            tab.style.display = 'inline-block';
            tab.style.width = '36pt';
            destino.appendChild(tab);
            break;
          }
          case 'br':
          case 'cr': destino.appendChild(document.createElement('br')); break;
          case 'noBreakHyphen': adicionarTexto(destino, '\u2011', r); break;
          case 'pict': renderizarVml(c, ctx, destino); break;
          case 'drawing': renderizarDesenho(c, ctx, destino); break;
          default: break;
        }
      } else if (c.namespaceURI === NS_MC && c.localName === 'AlternateContent') {
        const alt = filho(c, 'Fallback', NS_MC) || filho(c, 'Choice', NS_MC);
        if (alt) renderizarRun(alt, r, ctx, destino);
      }
    }
  }

  // ------------------------------------------------------------------ parágrafos e tabelas

  function renderizarParagrafo(pXml, ctx) {
    const pPr = filho(pXml, 'pPr');
    const { p, r: rBase } = propriedades(ctx.estilos, pPr);
    const rMarca = lerRun(filho(pPr, 'rPr'), { ...rBase });
    const el = document.createElement('div');
    el.className = 'w-p';
    const ctxP = { ...ctx, recuoEsq: p.recuoEsq };

    // A altura da linha segue a maior fonte com texto (ou a marca de parágrafo, se vazio).
    let dominante = rMarca;
    for (const run of runsDoParagrafo(pXml)) {
      const r = lerRun(filho(run, 'rPr'), { ...rBase });
      if (textoDoRun(run) && !r.oculto && r.tamanho > dominante.tamanho) dominante = r;
      renderizarRun(run, r, ctxP, el);
    }
    const alinhamento = { both: 'justify', distribute: 'justify', center: 'center', right: 'right', end: 'right' }[p.alinhamento] || 'left';
    // O Word não considera os espaços do fim da linha ao centralizar/alinhar à direita.
    if (alinhamento === 'center' || alinhamento === 'right') removerEspacosFinais(el);

    // Parágrafo vazio ou terminado em quebra de linha: um caractere invisível cria a linha final,
    // como o Word faz. (Nos demais, ele não é usado para não afetar o alinhamento da última linha.)
    const conteudo = Array.from(el.childNodes).filter((n) => n.nodeName === 'BR'
      || (n.classList && n.classList.contains('w-em-linha'))
      || (n.textContent && !(n.classList && n.classList.contains('w-flutuante'))));
    if (!conteudo.length || conteudo[conteudo.length - 1].nodeName === 'BR') {
      const escora = document.createElement('span');
      escora.textContent = '\u200B';
      el.appendChild(escora);
    }

    const s = el.style;
    s.textAlign = alinhamento;
    if (p.recuoEsq) s.marginLeft = p.recuoEsq + 'pt';
    if (p.recuoDir) s.marginRight = p.recuoDir + 'pt';
    if (p.primeiraLinha) s.textIndent = p.primeiraLinha + 'pt';
    if (p.antes) s.paddingTop = p.antes + 'pt';
    if (p.depois) s.paddingBottom = p.depois + 'pt';
    s.fontFamily = pilhaDaFonte(dominante.fonte);
    s.fontSize = dominante.tamanho + 'pt';

    const simples = alturaLinha(dominante.fonte) * dominante.tamanho;
    let altura = simples * (p.linha / 240);
    if (p.regra === 'exact') altura = p.linha / 20;
    else if (p.regra === 'atLeast') altura = Math.max(p.linha / 20, simples);
    s.lineHeight = altura + 'pt';
    // O Word coloca o espaço extra do entrelinhas abaixo do texto; o navegador o divide acima e abaixo.
    if (p.regra === 'auto' && altura > simples) s.top = -(altura - simples) / 2 + 'pt';
    return el;
  }

  function lerBorda(b) {
    if (!b) return undefined;
    const v = atributo(b, 'val');
    if (!v || v === 'nil' || v === 'none') return 'none';
    const largura = Math.max(0.25, numero(atributo(b, 'sz'), 4) / 8);
    const cor = atributo(b, 'color');
    const estilo = v === 'double' ? 'double' : v === 'dotted' ? 'dotted' : v.startsWith('dash') ? 'dashed' : 'solid';
    return `${v === 'double' ? Math.max(largura * 3, 1.5) : largura}pt ${estilo} ${cor && cor !== 'auto' ? '#' + cor : '#000'}`;
  }

  function lerMargens(el, m) {
    if (!el) return m;
    const lados = { top: 'top', bottom: 'bottom', left: 'left', right: 'right', start: 'left', end: 'right' };
    for (const c of el.children) {
      const lado = lados[c.localName];
      if (lado && atributo(c, 'type') !== 'pct') m[lado] = numero(atributo(c, 'w'), 0) / 20;
    }
    return m;
  }

  function renderizarTabela(tbl, ctx) {
    const tblPr = filho(tbl, 'tblPr');
    const grade = filhos(filho(tbl, 'tblGrid'), 'gridCol').map((g) => numero(atributo(g, 'w'), 0) / 20);
    const margens = lerMargens(filho(tblPr, 'tblCellMar'), { top: 0, bottom: 0, left: 5.4, right: 5.4 });
    const ind = filho(tblPr, 'tblInd');
    const recuo = ind && atributo(ind, 'type') !== 'pct' ? numero(atributo(ind, 'w'), 0) / 20 : 0;
    const bt = filho(tblPr, 'tblBorders');
    const bordas = {};
    for (const lado of ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']) bordas[lado] = lerBorda(filho(bt, lado));

    const tabela = document.createElement('table');
    tabela.className = 'w-tbl';
    tabela.style.width = grade.reduce((a, b) => a + b, 0) + 'pt';
    if (atributo(filho(tblPr, 'jc'), 'val') === 'center') tabela.style.margin = '0 auto';
    else tabela.style.marginLeft = recuo - margens.left + 'pt';
    const colunas = document.createElement('colgroup');
    for (const largura of grade) {
      const col = document.createElement('col');
      col.style.width = largura + 'pt';
      colunas.appendChild(col);
    }
    tabela.appendChild(colunas);
    const corpo = document.createElement('tbody');
    tabela.appendChild(corpo);

    const linhas = filhos(tbl, 'tr');
    linhas.forEach((tr, i) => {
      const linha = document.createElement('tr');
      const celulas = filhos(tr, 'tc');
      const altura = numero(atributo(filho(filho(tr, 'trPr'), 'trHeight'), 'val'));
      if (altura) {
        // No Word, a borda horizontal soma-se à altura mínima da linha.
        const borda = Math.max(0, ...celulas.map((tc) => {
          const b = filho(filho(filho(tc, 'tcPr'), 'tcBorders'), 'top');
          return b && !['nil', 'none'].includes(atributo(b, 'val')) ? numero(atributo(b, 'sz'), 4) / 8 : 0;
        }));
        linha.style.height = altura / 20 + borda + 'pt';
      }
      celulas.forEach((tc, j) => {
        const tcPr = filho(tc, 'tcPr');
        const td = document.createElement('td');
        const span = numero(atributo(filho(tcPr, 'gridSpan'), 'val'), 1);
        if (span > 1) td.colSpan = span;
        const bc = filho(tcPr, 'tcBorders');
        const borda = (lado, alternativo, interno, externo) => {
          const v = lerBorda(filho(bc, lado)) ?? lerBorda(filho(bc, alternativo));
          if (v !== undefined) return v;
          return (externo ? bordas[lado] : bordas[interno]) || 'none';
        };
        td.style.borderTop = borda('top', 'top', 'insideH', i === 0);
        td.style.borderBottom = borda('bottom', 'bottom', 'insideH', i === linhas.length - 1);
        td.style.borderLeft = borda('left', 'start', 'insideV', j === 0);
        td.style.borderRight = borda('right', 'end', 'insideV', j === celulas.length - 1);
        const m = lerMargens(filho(tcPr, 'tcMar'), { ...margens });
        td.style.padding = `${m.top}pt ${m.right}pt ${m.bottom}pt ${m.left}pt`;
        const va = atributo(filho(tcPr, 'vAlign'), 'val');
        td.style.verticalAlign = va === 'center' ? 'middle' : va === 'bottom' ? 'bottom' : 'top';
        const fundo = atributo(filho(tcPr, 'shd'), 'fill');
        if (fundo && fundo !== 'auto') td.style.background = '#' + fundo;
        renderizarBlocos(tc, ctx, td);
        linha.appendChild(td);
      });
      corpo.appendChild(linha);
    });
    return tabela;
  }

  function renderizarBlocos(pai, ctx, destino) {
    for (const c of pai.children) {
      if (c.namespaceURI !== NS_W) continue;
      if (c.localName === 'p') destino.appendChild(renderizarParagrafo(c, ctx));
      else if (c.localName === 'tbl') destino.appendChild(renderizarTabela(c, ctx));
      else if (c.localName === 'sdt') renderizarBlocos(filho(c, 'sdtContent') || c, ctx, destino);
      else if (c.localName === 'customXml') renderizarBlocos(c, ctx, destino);
    }
  }

  // ------------------------------------------------------------------ imagens e caixas de texto

  function estiloVml(texto) {
    const css = {};
    for (const parte of (texto || '').split(';')) {
      const i = parte.indexOf(':');
      if (i > 0) css[parte.slice(0, i).trim().toLowerCase()] = parte.slice(i + 1).trim();
    }
    return css;
  }

  function medida(v) {
    if (v == null || v === '') return null;
    const m = /^(-?[\d.]+)\s*(pt|in|cm|mm|px|pc|emu)?$/i.exec(String(v).trim());
    if (!m) return null;
    const n = parseFloat(m[1]);
    switch ((m[2] || 'px').toLowerCase()) {
      case 'pt': return n;
      case 'in': return n * 72;
      case 'cm': return (n * 72) / 2.54;
      case 'mm': return (n * 72) / 25.4;
      case 'pc': return n * 12;
      case 'emu': return n / EMU;
      default: return n * 0.75;
    }
  }

  function corVml(v) {
    const c = (v || '').trim().split(/\s+/)[0];
    return /^#[0-9a-f]{3,8}$/i.test(c) || /^[a-z]+$/i.test(c) ? c : null;
  }

  /** Posiciona um objeto flutuante (imagem ou caixa de texto) como o Word faz. */
  function posicionar(el, d, ctx, destino) {
    const g = ctx.geo;
    el.classList.add('w-flutuante');
    el.style.zIndex = d.atras ? '-1' : '1';
    el.dataset.quebra = d.quebra;

    let base;
    let faixa;
    switch (d.relH) {
      case 'page': base = 0; faixa = g.largura; break;
      case 'left-margin-area': base = 0; faixa = g.esquerda; break;
      case 'right-margin-area': base = g.largura - g.direita; faixa = g.direita; break;
      default: base = g.esquerda; faixa = g.texto; break; // margem, coluna ou caractere
    }
    let x = base + d.dx;
    if (d.alinH === 'center') x = base + (faixa - d.largura) / 2;
    else if (d.alinH === 'right' || d.alinH === 'outside') x = base + faixa - d.largura;
    else if (d.alinH === 'left' || d.alinH === 'inside') x = base;

    if (['page', 'margin', 'top-margin-area', 'bottom-margin-area'].includes(d.relV)) {
      let baseY = g.topo;
      let faixaY = g.altura - g.topo - g.base;
      if (d.relV === 'page') { baseY = 0; faixaY = g.altura; }
      else if (d.relV === 'top-margin-area') { baseY = 0; faixaY = g.topo; }
      else if (d.relV === 'bottom-margin-area') { baseY = g.altura - g.base; faixaY = g.base; }
      let y = baseY + d.dy;
      if (d.alinV === 'center') y = baseY + (faixaY - d.altura) / 2;
      else if (d.alinV === 'bottom' || d.alinV === 'outside') y = baseY + faixaY - d.altura;
      else if (d.alinV === 'top' || d.alinV === 'inside') y = baseY;
      el.style.left = x + 'pt';
      el.style.top = y + 'pt';
      ctx.flutuantesPagina.push(el);
    } else {
      // Relativo ao parágrafo (padrão do Word): o objeto acompanha o parágrafo em que está ancorado.
      el.style.left = x - ctx.esquerda - (ctx.recuoEsq || 0) + 'pt';
      el.style.top = d.dy + 'pt';
      destino.appendChild(el);
    }
  }

  function renderizarVml(pict, ctx, destino) {
    for (const forma of pict.children) {
      if (forma.namespaceURI !== NS_V || forma.localName === 'shapetype') continue;
      const css = estiloVml(forma.getAttribute('style'));
      if (css.visibility === 'hidden' || css.display === 'none') continue;
      const largura = medida(css.width) || 0;
      const altura = medida(css.height) || 0;
      const imagem = filho(forma, 'imagedata', NS_V);
      const caixa = filho(forma, 'textbox', NS_V);
      let el;
      let alturaAutomatica = false;
      if (imagem) {
        const url = ctx.imagens.get(imagem.getAttributeNS(NS_R, 'id') || imagem.getAttributeNS(NS_R, 'pict'));
        if (!url) continue;
        el = document.createElement('img');
        el.src = url;
        el.alt = '';
        el.className = 'w-img';
      } else if (caixa) {
        el = document.createElement('div');
        el.className = 'w-caixa';
        const inset = (caixa.getAttribute('inset') || '').split(',').map(medida);
        el.style.padding = `${inset[1] ?? 3.6}pt ${inset[2] ?? 7.2}pt ${inset[3] ?? 3.6}pt ${inset[0] ?? 7.2}pt`;
        const conteudo = filho(caixa, 'txbxContent');
        if (conteudo) renderizarBlocos(conteudo, { ...ctx, esquerda: 0, recuoEsq: 0, flutuantesPagina: [] }, el);
        alturaAutomatica = /mso-fit-shape-to-text\s*:\s*t/.test(caixa.getAttribute('style') || '');
        const preenchida = !['f', 'false'].includes(forma.getAttribute('filled'));
        const contornada = !['f', 'false'].includes(forma.getAttribute('stroked'));
        if (preenchida) el.style.background = corVml(forma.getAttribute('fillcolor')) || '#fff';
        if (contornada) el.style.border = `${medida(forma.getAttribute('strokeweight')) || 0.75}pt solid ${corVml(forma.getAttribute('strokecolor')) || '#000'}`;
        if (filho(forma, 'shadow', NS_V) && !preenchida) el.classList.add('w-sombra-texto');
      } else {
        continue;
      }
      el.style.width = largura + 'pt';
      if (!alturaAutomatica) el.style.height = altura + 'pt';

      if (css.position === 'absolute') {
        const quebra = filho(forma, 'wrap', NS_W10);
        posicionar(el, {
          relH: css['mso-position-horizontal-relative'] || 'text',
          alinH: css['mso-position-horizontal'] || 'absolute',
          dx: (medida(css['margin-left']) || 0) + (medida(css.left) || 0),
          relV: css['mso-position-vertical-relative'] || 'text',
          alinV: css['mso-position-vertical'] || 'absolute',
          dy: (medida(css['margin-top']) || 0) + (medida(css.top) || 0),
          largura,
          altura,
          atras: numero(css['z-index'], 0) < 0,
          quebra: quebra ? quebra.getAttribute('type') || 'square' : 'none',
        }, ctx, destino);
      } else {
        el.classList.add('w-em-linha');
        destino.appendChild(el);
      }
    }
  }

  function renderizarDesenho(desenho, ctx, destino) {
    const RELH = { page: 'page', margin: 'margin', column: 'text', character: 'char', leftMargin: 'left-margin-area', rightMargin: 'right-margin-area' };
    const RELV = { page: 'page', margin: 'margin', paragraph: 'text', line: 'text', topMargin: 'top-margin-area', bottomMargin: 'bottom-margin-area' };
    for (const c of desenho.children) {
      if (c.namespaceURI !== NS_WP) continue;
      const extensao = filho(c, 'extent', NS_WP);
      const largura = numero(atributo(extensao, 'cx', null), 0) / EMU;
      const altura = numero(atributo(extensao, 'cy', null), 0) / EMU;
      const blip = c.getElementsByTagNameNS(NS_A, 'blip')[0];
      const url = blip && ctx.imagens.get(blip.getAttributeNS(NS_R, 'embed'));
      if (!url) continue;
      const img = document.createElement('img');
      img.src = url;
      img.alt = '';
      img.className = 'w-img';
      img.style.width = largura + 'pt';
      img.style.height = altura + 'pt';
      if (c.localName !== 'anchor') {
        img.classList.add('w-em-linha');
        destino.appendChild(img);
        continue;
      }
      const posicao = (el) => ({
        rel: atributo(el, 'relativeFrom', null),
        alin: filho(el, 'align', NS_WP)?.textContent || null,
        desloc: numero(filho(el, 'posOffset', NS_WP)?.textContent, 0) / EMU,
      });
      const h = posicao(filho(c, 'positionH', NS_WP));
      const v = posicao(filho(c, 'positionV', NS_WP));
      const quebra = ['wrapSquare', 'wrapTight', 'wrapThrough', 'wrapTopAndBottom'].find((n) => filho(c, n, NS_WP));
      posicionar(img, {
        relH: RELH[h.rel] || 'text', alinH: h.alin || 'absolute', dx: h.desloc,
        relV: RELV[v.rel] || 'text', alinV: v.alin || 'absolute', dy: v.desloc,
        largura, altura,
        atras: c.getAttribute('behindDoc') === '1',
        quebra: quebra ? quebra.replace('wrap', '').toLowerCase() : 'none',
      }, ctx, destino);
    }
  }

  // ------------------------------------------------------------------ páginas

  function geometria(secao) {
    const pg = filho(secao, 'pgSz');
    const mg = filho(secao, 'pgMar');
    const v = (el, nome, padrao) => Math.abs(numero(atributo(el, nome), padrao)) / 20;
    const g = {
      largura: v(pg, 'w', 11906), altura: v(pg, 'h', 16838),
      topo: v(mg, 'top', 1440), base: v(mg, 'bottom', 1440),
      esquerda: v(mg, 'left', 1800), direita: v(mg, 'right', 1800),
      cabecalho: v(mg, 'header', 720), rodape: v(mg, 'footer', 720),
    };
    g.texto = g.largura - g.esquerda - g.direita;
    return g;
  }

  function parteReferenciada(recursos, secao, nome) {
    const refs = filhos(secao, nome);
    const ref = refs.find((r) => atributo(r, 'type') === 'default') || refs[0];
    return ref ? recursos.partes.get(ref.getAttributeNS(NS_R, 'id')) || null : null;
  }

  function contexto(recursos, imgs, g) {
    return { estilos: recursos.estilos, imagens: imgs, geo: g, esquerda: g.esquerda, recuoEsq: 0, flutuantesPagina: [] };
  }

  function criarArea(classe, g) {
    const area = document.createElement('div');
    area.className = classe;
    area.style.left = g.esquerda + 'pt';
    area.style.width = g.texto + 'pt';
    return area;
  }

  function posicaoNaPagina(el, pagina) {
    let x = 0;
    let y = 0;
    for (let atual = el; atual && atual !== pagina; atual = atual.offsetParent) {
      x += atual.offsetLeft;
      y += atual.offsetTop;
    }
    return { x: x / PX, y: y / PX };
  }

  /** Move para novas páginas os blocos (e linhas de tabela) que não cabem na página. */
  function paginar(primeira, novaPagina) {
    let atual = primeira;
    for (let guarda = 0; guarda < 100; guarda++) {
      const limite = atual.disponivel * PX + 1;
      const blocos = Array.from(atual.area.children);
      let i = blocos.findIndex((b) => b.offsetTop + b.offsetHeight > limite);
      if (i < 0) return;
      const proxima = novaPagina();
      const bloco = blocos[i];
      if (bloco.tagName === 'TABLE') {
        const linhas = Array.from(bloco.rows);
        const k = linhas.findIndex((tr) => bloco.offsetTop + tr.offsetTop + tr.offsetHeight > limite);
        if (k > 0) {
          const resto = bloco.cloneNode(false);
          const colunas = bloco.querySelector('colgroup');
          if (colunas) resto.appendChild(colunas.cloneNode(true));
          const corpo = document.createElement('tbody');
          linhas.slice(k).forEach((tr) => corpo.appendChild(tr));
          resto.appendChild(corpo);
          proxima.area.appendChild(resto);
          i += 1;
        }
      }
      if (i === 0 && !proxima.area.children.length) i = 1; // bloco maior que a página
      blocos.slice(i).forEach((b) => proxima.area.appendChild(b));
      if (!proxima.area.children.length) {
        proxima.pagina.parentNode.remove();
        return;
      }
      atual = proxima;
    }
  }

  /**
   * Ajusta o tamanho das páginas ao painel: à largura disponível ou, com data-zoom="pagina",
   * para que a página inteira caiba na altura do painel.
   */
  function ajustarEscala(alvo) {
    const folhas = alvo.querySelector('.folhas');
    if (!folhas) return;
    const estilo = getComputedStyle(alvo);
    const disponivel = alvo.clientWidth - parseFloat(estilo.paddingLeft) - parseFloat(estilo.paddingRight);
    let alturaDisponivel = Infinity;
    if (alvo.dataset.zoom === 'pagina') {
      const maxima = parseFloat(estilo.maxHeight);
      alturaDisponivel = (isNaN(maxima) ? window.innerHeight * 0.85 : maxima)
        - parseFloat(estilo.paddingTop) - parseFloat(estilo.paddingBottom) - 2;
    }
    for (const folha of folhas.children) {
      const pagina = folha.firstElementChild;
      const largura = pagina.offsetWidth;
      const altura = pagina.offsetHeight;
      const escala = Math.max(0.2, Math.min(1, disponivel / largura, alturaDisponivel / altura));
      pagina.style.transform = `scale(${escala})`;
      folha.style.width = largura * escala + 'px';
      folha.style.height = altura * escala + 'px';
    }
  }

  /** Desenha o documento em `alvo`; devolve o número de páginas. */
  function renderizar(recursos, doc, alvo) {
    const corpo = doc.getElementsByTagNameNS(NS_W, 'body')[0];
    const secao = filho(corpo, 'sectPr') || Array.from(corpo.getElementsByTagNameNS(NS_W, 'sectPr')).pop();
    const g = geometria(secao);
    const cabecalho = parteReferenciada(recursos, secao, 'headerReference');
    const rodape = parteReferenciada(recursos, secao, 'footerReference');

    const folhas = document.createElement('div');
    folhas.className = 'folhas';
    alvo.replaceChildren(folhas);

    const desenharParte = (parte, area, pagina) => {
      const ctx = contexto(recursos, parte.imagens, g);
      renderizarBlocos(parte.doc.documentElement, ctx, area);
      ctx.flutuantesPagina.forEach((el) => pagina.appendChild(el));
    };

    const novaPagina = () => {
      const folha = document.createElement('div');
      folha.className = 'folha';
      const pagina = document.createElement('div');
      pagina.className = 'pagina';
      pagina.style.width = g.largura + 'pt';
      pagina.style.height = g.altura + 'pt';
      folha.appendChild(pagina);
      folhas.appendChild(folha);

      let topo = g.topo;
      let base = g.altura - g.base;
      if (cabecalho) {
        const area = criarArea('w-cabecalho', g);
        area.style.top = g.cabecalho + 'pt';
        pagina.appendChild(area);
        desenharParte(cabecalho, area, pagina);
        topo = Math.max(topo, g.cabecalho + area.offsetHeight / PX);
      }
      if (rodape) {
        const area = criarArea('w-rodape', g);
        area.style.bottom = g.rodape + 'pt';
        pagina.appendChild(area);
        desenharParte(rodape, area, pagina);
        base = Math.min(base, g.altura - g.rodape - area.offsetHeight / PX);
      }
      // Imagens do cabeçalho/rodapé com quebra de texto empurram o corpo do documento.
      for (const el of pagina.querySelectorAll('.w-flutuante')) {
        if (!el.dataset.quebra || el.dataset.quebra === 'none') continue;
        const largura = el.offsetWidth / PX;
        const altura = el.offsetHeight / PX;
        if (largura < g.texto / 2) continue;
        const { y } = posicaoNaPagina(el, pagina);
        if (y + altura / 2 < g.altura / 2) topo = Math.max(topo, y + altura);
        else base = Math.min(base, y);
      }
      const area = criarArea('w-corpo', g);
      area.style.top = topo + 'pt';
      pagina.appendChild(area);
      return { pagina, area, disponivel: base - topo };
    };

    const primeira = novaPagina();
    const ctx = contexto(recursos, recursos.imagensDocumento, g);
    renderizarBlocos(corpo, ctx, primeira.area);
    ctx.flutuantesPagina.forEach((el) => primeira.pagina.appendChild(el));
    paginar(primeira, novaPagina);
    ajustarEscala(alvo);
    return folhas.children.length;
  }

  NS.previa = { preparar, renderizar, ajustarEscala };
})(window.Declaracoes = window.Declaracoes || {});
