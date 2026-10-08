/*
 * Formatação e validação dos dados digitados nos formulários.
 */
(function (NS) {
  'use strict';

  const MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
    'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];

  const doisDigitos = (n) => String(n).padStart(2, '0');

  /** "AAAA-MM-DD" (valor de <input type="date">) → { dia, mes, ano }, ou null se inválida. */
  function lerData(valor) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(valor || '');
    if (!m) return null;
    const ano = Number(m[1]), mes = Number(m[2]), dia = Number(m[3]);
    const d = new Date(ano, mes - 1, dia);
    if (d.getFullYear() !== ano || d.getMonth() !== mes - 1 || d.getDate() !== dia) return null;
    if (ano < 1900 || ano > 2100) return null;
    return { dia, mes, ano };
  }

  /** 08/10/2026 */
  function dataNumerica(valor) {
    const d = lerData(valor);
    return d ? `${doisDigitos(d.dia)}/${doisDigitos(d.mes)}/${d.ano}` : '';
  }

  /** Partes da data para "[DIA] de [MÊS] de [ANO]": { dia: '8', mes: 'outubro', ano: '2026' } (dia 1 → "1º"). */
  function dataPorExtenso(valor) {
    const d = lerData(valor);
    if (!d) return null;
    return { dia: d.dia === 1 ? '1º' : String(d.dia), mes: MESES[d.mes - 1], ano: String(d.ano) };
  }

  /** "HH:MM" (valor de <input type="time">) → minutos desde 00:00, ou null se inválido. */
  function lerHora(valor) {
    const m = /^(\d{1,2}):(\d{2})/.exec(valor || '');
    if (!m) return null;
    const h = Number(m[1]), min = Number(m[2]);
    if (h > 23 || min > 59) return null;
    return h * 60 + min;
  }

  /** Formato da tabela de aulas dos docentes: 08h00 */
  function horaComH(valor) {
    const t = lerHora(valor);
    return t == null ? '' : `${doisDigitos(Math.floor(t / 60))}h${doisDigitos(t % 60)}`;
  }

  /** Formato da declaração de aluno: 08:00 */
  function horaComDoisPontos(valor) {
    const t = lerHora(valor);
    return t == null ? '' : `${doisDigitos(Math.floor(t / 60))}:${doisDigitos(t % 60)}`;
  }

  const somenteDigitos = (s) => String(s || '').replace(/\D/g, '');

  /** Aplica a máscara 000.000.000-00 à medida que o CPF é digitado. */
  function mascararCpf(valor) {
    const d = somenteDigitos(valor).slice(0, 11);
    let r = d.slice(0, 3);
    if (d.length > 3) r += '.' + d.slice(3, 6);
    if (d.length > 6) r += '.' + d.slice(6, 9);
    if (d.length > 9) r += '-' + d.slice(9, 11);
    return r;
  }

  /** Confere os dígitos verificadores do CPF. */
  function cpfValido(valor) {
    const d = somenteDigitos(valor);
    if (d.length !== 11 || /^(\d)\1{10}$/.test(d)) return false;
    const digito = (n) => {
      let soma = 0;
      for (let i = 0; i < n; i++) soma += Number(d[i]) * (n + 1 - i);
      const resto = (soma * 10) % 11;
      return resto === 10 ? 0 : resto;
    };
    return digito(9) === Number(d[9]) && digito(10) === Number(d[10]);
  }

  /**
   * Número do COREN-SP no formato do modelo ("000.000-ENF").
   * Somente números (ex.: 123456) → 123.456-ENF, mantendo o sufixo do modelo.
   * Com categoria (ex.: 123456-TE) → 123.456-TE. Outros formatos ficam como digitados.
   */
  function formatarCoren(valor) {
    const original = String(valor || '').trim();
    const s = original.toUpperCase()
      .replace(/^COREN\s*-?\s*SP\s*:?\s*/, '')
      .replace(/^COREN\s*:?\s*/, '')
      .replace(/^N[º°O.]\s*/, '')
      .replace(/\s+/g, '');
    const m = /^(\d[\d.]*)(?:-?([A-Z]{1,6}))?$/.exec(s);
    if (!m) return original;
    const numero = m[1].replace(/\./g, '').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return `${numero}-${m[2] || 'ENF'}`;
  }

  const normalizarEspacos = (s) => String(s || '').replace(/\s+/g, ' ').trim();

  /** Remove caracteres que o Windows não aceita em nomes de arquivo. */
  function nomeDeArquivo(s) {
    return normalizarEspacos(String(s || '').replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' '))
      .replace(/[. ]+$/, '')
      .slice(0, 150);
  }

  NS.formatacao = {
    MESES, lerData, dataNumerica, dataPorExtenso, lerHora, horaComH, horaComDoisPontos,
    somenteDigitos, mascararCpf, cpfValido, formatarCoren, normalizarEspacos, nomeDeArquivo,
  };
})(window.Declaracoes = window.Declaracoes || {});
