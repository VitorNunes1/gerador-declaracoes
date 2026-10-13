/*
 * Preenchimento dos modelos Word (.docx).
 *
 * O arquivo gerado é o próprio modelo original: somente o texto dos marcadores (ex.: "[NOME]")
 * é trocado e, na tabela de aulas, a linha modelo é repetida. Layout, fontes, cabeçalho,
 * rodapé e demais elementos permanecem exatamente como no modelo.
 */
(function (NS) {
  'use strict';

  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
  const XML_NS = 'http://www.w3.org/XML/1998/namespace';
  const MIME_DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const DECLARACAO_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

  // Delimitadores usados apenas na pré-visualização para destacar os campos
  // (caracteres Unicode de uso privado; nunca vão para o arquivo .docx).
  const MARCAS = {
    preenchido: ['\uE000', '\uE001'],
    pendente: ['\uE002', '\uE003'],
    manual: ['\uE004', '\uE005'],
  };

  const modelos = new Map();

  function bytesDeBase64(b64) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  /** Carrega (uma única vez) o modelo embutido em assets/js/modelos-embutidos.js. */
  function carregar(id) {
    if (!modelos.has(id)) {
      const promessa = (async () => {
        const b64 = (window.MODELOS_EMBUTIDOS || {})[id];
        if (!b64) throw new Error(`O modelo "${id}" não foi encontrado (assets/js/modelos-embutidos.js).`);
        const bytes = bytesDeBase64(b64);
        const zip = await JSZip.loadAsync(bytes);
        const arquivo = zip.file('word/document.xml');
        if (!arquivo) throw new Error(`O modelo "${id}" não é um documento Word válido.`);
        return { id, bytes, zip, documentoXml: await arquivo.async('string') };
      })();
      promessa.catch(() => modelos.delete(id));
      modelos.set(id, promessa);
    }
    return modelos.get(id);
  }

  function lerXml(texto) {
    const doc = new DOMParser().parseFromString(texto, 'application/xml');
    const erro = doc.getElementsByTagName('parsererror')[0];
    if (erro) throw new Error('O XML do modelo é inválido: ' + erro.textContent.slice(0, 200));
    return doc;
  }

  function escreverXml(doc) {
    return DECLARACAO_XML + new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^?]*\?>\s*/, '');
  }

  const ehW = (el, nome) => el.namespaceURI === W && el.localName === nome;

  /** Todos os parágrafos (w:p) do documento, inclusive os de tabelas e caixas de texto. */
  function paragrafos(raiz) {
    return Array.from(raiz.getElementsByTagNameNS(W, 'p'));
  }

  /** Nós w:t que pertencem ao parágrafo (sem os de parágrafos aninhados em caixas de texto). */
  function nosDeTexto(p) {
    const lista = [];
    (function percorrer(el) {
      for (const filho of el.children) {
        if (ehW(filho, 'p')) continue;
        if (ehW(filho, 't')) lista.push(filho);
        else percorrer(filho);
      }
    })(p);
    return lista;
  }

  function textoDoParagrafo(p) {
    return nosDeTexto(p).map((t) => t.textContent).join('');
  }

  function definirTexto(t, texto) {
    t.textContent = texto;
    if (/^\s|\s$|\s\s/.test(texto)) t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  }

  /**
   * Substitui, em ordem, as ocorrências de `regex` (com a flag "g") nos parágrafos informados.
   * O Word costuma dividir um mesmo marcador em vários trechos (runs); a busca é feita sobre o
   * texto completo do parágrafo e o novo texto herda a formatação do trecho onde começa o
   * conteúdo do marcador (colchetes são ignorados nessa escolha).
   *
   * `valor(indice, encontrado)` devolve o texto da ocorrência; vazio/null mantém o marcador
   * (com `opcoes.apagar`, texto vazio apaga a ocorrência). Pode devolver também { texto }: um texto
   * provisório, ainda incompleto, que a pré-visualização destaca como pendente.
   * `opcoes.marcar` (só na pré-visualização): true, ou função (vazio) → 'preenchido' | 'pendente' | 'manual'.
   * Retorna o número de ocorrências encontradas.
   */
  function substituir(listaParagrafos, regex, valor, opcoes = {}) {
    let ocorrencia = 0;
    for (const p of listaParagrafos) {
      let inicioBusca = 0;
      for (;;) {
        const nos = nosDeTexto(p);
        const textos = nos.map((n) => n.textContent);
        const completo = textos.join('');
        regex.lastIndex = inicioBusca;
        const m = regex.exec(completo);
        if (!m || !m[0]) break;
        const ini = m.index;
        const fim = ini + m[0].length;

        let novo = valor(ocorrencia, m[0]);
        const provisorio = novo != null && typeof novo === 'object';
        if (provisorio) novo = novo.texto;
        const vazio = novo == null || (novo === '' && !opcoes.apagar);
        if (vazio) novo = m[0];
        if (opcoes.marcar) {
          const tipo = typeof opcoes.marcar === 'function' ? opcoes.marcar(vazio) : (vazio || provisorio ? 'pendente' : 'preenchido');
          if (tipo) novo = MARCAS[tipo][0] + novo + MARCAS[tipo][1];
        }

        // Trecho que recebe o texto novo: o que contém o primeiro caractere "real" do marcador.
        let alvo = -1;
        let reserva = -1;
        let pos = 0;
        for (let i = 0; i < nos.length && alvo < 0; i++) {
          const s = pos;
          const e = pos + textos[i].length;
          pos = e;
          for (let k = Math.max(s, ini); k < Math.min(e, fim); k++) {
            if (reserva < 0) reserva = i;
            if (!'[] '.includes(completo[k])) { alvo = i; break; }
          }
        }
        if (alvo < 0) alvo = reserva;

        pos = 0;
        for (let i = 0; i < nos.length; i++) {
          const s = pos;
          const e = pos + textos[i].length;
          pos = e;
          if (e <= ini || s >= fim) continue;
          const a = Math.max(ini, s) - s;
          const b = Math.min(fim, e) - s;
          const t = textos[i];
          definirTexto(nos[i], i === alvo ? t.slice(0, a) + novo + t.slice(b) : t.slice(0, a) + t.slice(b));
        }
        inicioBusca = ini + novo.length;
        ocorrencia++;
      }
    }
    return ocorrencia;
  }

  /** Apaga do texto dos parágrafos as ocorrências de `regex`; retorna quantas foram apagadas. */
  function apagar(listaParagrafos, regex) {
    return substituir(listaParagrafos, regex, () => '', { apagar: true });
  }

  /** Linha de tabela (w:tr) cujo texto contém `marcador`. */
  function linhaComMarcador(doc, marcador) {
    return Array.from(doc.getElementsByTagNameNS(W, 'tr'))
      .find((tr) => paragrafos(tr).some((p) => textoDoParagrafo(p).includes(marcador))) || null;
  }

  function idsExistentes(doc) {
    const usados = new Set();
    for (const el of doc.getElementsByTagName('*')) {
      for (const nome of ['paraId', 'textId']) {
        const v = el.getAttributeNS(W14, nome);
        if (v) usados.add(v.toUpperCase());
      }
    }
    return usados;
  }

  function novoId(usados) {
    let id;
    do {
      id = Math.floor(Math.random() * 0x7FFFFFFF).toString(16).toUpperCase().padStart(8, '0');
    } while (usados.has(id));
    usados.add(id);
    return id;
  }

  /** Mantém únicos os identificadores de parágrafo e remove marcadores (bookmarks) duplicados. */
  function renovarIdentificadores(el, usados) {
    for (const no of [el, ...el.getElementsByTagName('*')]) {
      for (const nome of ['paraId', 'textId']) {
        if (no.hasAttributeNS(W14, nome)) no.setAttributeNS(W14, 'w14:' + nome, novoId(usados));
      }
    }
    for (const nome of ['bookmarkStart', 'bookmarkEnd']) {
      for (const b of Array.from(el.getElementsByTagNameNS(W, nome))) b.parentNode.removeChild(b);
    }
  }

  /**
   * Repete a linha modelo da tabela, uma vez para cada item, e remove a linha original.
   * `preencher(paragrafosDaLinha, item, indice)` substitui os marcadores de cada cópia.
   */
  function repetirLinha(linhaModelo, itens, preencher) {
    const doc = linhaModelo.ownerDocument;
    const usados = idsExistentes(doc);
    const pai = linhaModelo.parentNode;
    itens.forEach((item, i) => {
      const copia = linhaModelo.cloneNode(true);
      renovarIdentificadores(copia, usados);
      preencher(paragrafos(copia), item, i);
      pai.insertBefore(copia, linhaModelo);
    });
    if (itens.length) pai.removeChild(linhaModelo);
  }

  /**
   * O Word representa o hífen inseparável com um elemento próprio (w:noBreakHyphen), com a mesma
   * aparência do hífen comum; o caractere U+2011 é convertido nesse elemento.
   */
  function converterHifensInseparaveis(doc) {
    for (const t of Array.from(doc.getElementsByTagNameNS(W, 't'))) {
      if (!t.textContent.includes('\u2011')) continue;
      const pai = t.parentNode;
      t.textContent.split('\u2011').forEach((parte, i) => {
        if (i > 0) pai.insertBefore(doc.createElementNS(W, 'w:noBreakHyphen'), t);
        if (parte) {
          const novo = doc.createElementNS(W, 'w:t');
          definirTexto(novo, parte);
          pai.insertBefore(novo, t);
        }
      });
      pai.removeChild(t);
    }
  }

  /** Monta o .docx: o pacote original com apenas o word/document.xml substituído. */
  async function montar(modelo, doc) {
    converterHifensInseparaveis(doc);
    const zip = await JSZip.loadAsync(modelo.bytes);
    zip.file('word/document.xml', escreverXml(doc));
    return zip.generateAsync({
      type: 'blob',
      mimeType: MIME_DOCX,
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
  }

  NS.docx = {
    W, MARCAS, carregar, lerXml, escreverXml, paragrafos, textoDoParagrafo,
    substituir, apagar, linhaComMarcador, repetirLinha, montar,
  };
})(window.Declaracoes = window.Declaracoes || {});
