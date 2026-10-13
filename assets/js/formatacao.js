/*
 * Formatação e validação dos dados digitados nos formulários.
 */
(function (NS) {
  'use strict';

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

  const normalizarEspacos = (s) => String(s || '').replace(/\s+/g, ' ').trim();

  /**
   * Número do COREN-SP exatamente como digitado (ex.: 123456 ou 123.456), sem acréscimos.
   * Se a pessoa digitar também "COREN-SP" ou "nº" no início, eles são retirados, porque o
   * texto "Coren-SP" já é escrito antes do número na declaração.
   */
  function numeroCoren(valor) {
    return normalizarEspacos(valor)
      .replace(/^coren\s*(?:-\s*)?(?:sp)?\s*:?\s*/i, '')
      .replace(/^n\s*[º°o.]\s*/i, '')
      .trim();
  }

  /** Impede quebra de linha dentro do texto (ex.: "CPF 529.982.247-25"): espaços e hífens inseparáveis. */
  function inseparavel(texto) {
    return String(texto || '').replace(/ /g, '\u00A0').replace(/-/g, '\u2011');
  }

  /** Remove caracteres que o Windows não aceita em nomes de arquivo. */
  function nomeDeArquivo(s) {
    return normalizarEspacos(String(s || '').replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' '))
      .replace(/[. ]+$/, '')
      .slice(0, 150);
  }

  NS.formatacao = {
    lerData, dataNumerica, lerHora, horaComH, horaComDoisPontos,
    somenteDigitos, mascararCpf, cpfValido, numeroCoren, normalizarEspacos, inseparavel, nomeDeArquivo,
  };
})(window.Declaracoes = window.Declaracoes || {});
