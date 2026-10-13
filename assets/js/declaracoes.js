/*
 * Tipos de declaração: campos do formulário, validação e preenchimento de cada modelo.
 *
 * Marcadores substituídos (exatamente como estão nos modelos):
 *   Docente (presencial e online): [NOME], "Coren-SP 000.000-ENF" (texto do modelo, trocado pelo
 *     documento escolhido: "Coren-SP <número>" ou "CPF <número>"),
 *     [PROGRAMA] e, na tabela de aulas: 00/00/0000, 00h00 (início), 00h00 (término), [NOME DA AULA].
 *     O trecho "no dia [DIA] de [MÊS] de [ANO]" (ou "[DIA] de [MÊS] de [ANO]", no modelo online)
 *     é retirado, e a frase termina em "conforme abaixo."
 *   Aluno: [NOME], [000.000.000-00], [00/00/0000], [00:00] (início), [00:00] (término)
 * A data da assinatura/declaração não é preenchida: fica no documento para preenchimento manual.
 */
(function (NS) {
  'use strict';

  const F = NS.formatacao;
  const D = NS.docx;

  const MANUAIS_DOCENTE = /\[DIA ASSINATURA\]|\[?MÊS ASSINATURA\]|\[ANO ASSINATURA\]/g;
  const MANUAIS_ALUNO = /\[(?:DIA|MÊS|ANO) DECLARAÇÃO\]/g;

  function exigir(quantidade, marcador) {
    if (quantidade === 0) {
      throw new Error(`O marcador "${marcador}" não foi encontrado no modelo. Verifique se o arquivo da pasta "modelos" foi alterado.`);
    }
  }

  function erroDeData(valor, obrigatorio) {
    if (!valor) return obrigatorio;
    return F.lerData(valor) ? null : 'Data inválida.';
  }

  // ---------------------------------------------------------------- Docente

  const aulaVazia = (data = '') => ({ data, inicio: '', fim: '', aula: '' });
  const usaCpf = (e) => e.tipoDocumento === 'cpf';

  /**
   * Documento do docente como aparece na declaração ("Coren-SP 123456" ou "CPF 123.456.789-00",
   * sem quebra de linha no meio). Enquanto não está completo, `texto` traz o modelo do tipo
   * escolhido ou o que já foi digitado, para a pré-visualização.
   */
  function documentoDocente(e) {
    if (usaCpf(e)) {
      const digitos = F.somenteDigitos(e.documento);
      const completo = digitos.length === 11 && F.cpfValido(digitos);
      return { completo, texto: F.inseparavel(`CPF ${digitos ? F.mascararCpf(digitos) : '000.000.000-00'}`) };
    }
    const numero = F.numeroCoren(e.documento);
    return { completo: Boolean(numero), texto: F.inseparavel(`Coren-SP ${numero || '000.000'}`) };
  }

  function validarDocente(e) {
    const erros = {};
    if (!F.normalizarEspacos(e.nome)) erros.nome = 'Informe o nome completo.';
    if (usaCpf(e)) {
      if (!F.somenteDigitos(e.documento)) erros.documento = 'Informe o CPF.';
      else if (!F.cpfValido(e.documento)) erros.documento = 'CPF inválido. Confira os números.';
    } else if (!F.numeroCoren(e.documento)) {
      erros.documento = 'Informe o número do COREN-SP.';
    }
    if (!F.normalizarEspacos(e.programa)) erros.programa = 'Informe o programa.';
    if (!e.aulas.length) erros.aulas = 'Adicione ao menos uma aula.';
    e.aulas.forEach((a, i) => {
      const chave = (campo) => `aulas.${i}.${campo}`;
      const erroDataAula = erroDeData(a.data, 'Informe a data.');
      if (erroDataAula) erros[chave('data')] = erroDataAula;
      if (!a.inicio) erros[chave('inicio')] = 'Informe o início.';
      if (!a.fim) erros[chave('fim')] = 'Informe o término.';
      else if (a.inicio && F.lerHora(a.fim) <= F.lerHora(a.inicio)) erros[chave('fim')] = 'Deve ser depois do início.';
      if (!F.normalizarEspacos(a.aula)) erros[chave('aula')] = 'Informe o nome da aula.';
    });
    return erros;
  }

  function valoresDocente(e) {
    const documento = documentoDocente(e);
    return {
      nome: F.normalizarEspacos(e.nome),
      documento: documento.completo ? documento.texto : '',
      documentoProvisorio: documento.texto,
      programa: F.normalizarEspacos(e.programa),
      aulas: e.aulas.map((a) => ({
        data: F.dataNumerica(a.data),
        inicio: F.horaComH(a.inicio),
        fim: F.horaComH(a.fim),
        aula: F.normalizarEspacos(a.aula),
      })),
    };
  }

  function preencherDocente(doc, v, previa) {
    const op = { marcar: previa };
    const linha = D.linhaComMarcador(doc, '[NOME DA AULA]');
    if (!linha) throw new Error('A tabela de aulas (linha com "[NOME DA AULA]") não foi encontrada no modelo.');
    const tabela = linha.parentNode;
    const texto = D.paragrafos(doc).filter((p) => !tabela.contains(p));

    exigir(D.substituir(texto, /\[NOME\]/g, () => v.nome, op), '[NOME]');
    // Na pré-visualização, enquanto o número não está completo, aparece o modelo do documento escolhido.
    const documento = () => v.documento || (previa ? { texto: v.documentoProvisorio } : null);
    exigir(D.substituir(texto, /Coren-SP 000\.000-ENF/g, documento, op), 'Coren-SP 000.000-ENF');
    exigir(D.substituir(texto, /\[PROGRAMA\]/g, () => v.programa, op), '[PROGRAMA]');
    // "..., no dia [DIA] de [MÊS] de [ANO] conforme abaixo." → "..., conforme abaixo."
    exigir(D.apagar(texto, /(?:no dia )?\[DIA\] de \[MÊS\] de \[ANO\] /g), '[DIA] de [MÊS] de [ANO]');

    D.repetirLinha(linha, v.aulas, (celulas, aula) => {
      D.substituir(celulas, /00\/00\/0000/g, () => aula.data, op);
      D.substituir(celulas, /00h00/g, (i) => (i === 0 ? aula.inicio : aula.fim), op);
      D.substituir(celulas, /\[NOME DA AULA\]/g, () => aula.aula, op);
    });

    if (previa) D.substituir(texto, MANUAIS_DOCENTE, () => null, { marcar: () => 'manual' });
  }

  function camposDocente(modalidade) {
    const trecho = modalidade === 'online' ? 'no programa' : 'para o Programa';
    return [
      {
        titulo: 'Dados do docente',
        campos: [
          {
            id: 'nome', rotulo: 'Nome completo', tipo: 'texto', autocomplete: 'name',
            dica: 'Aparece após "Drª" na declaração.',
          },
          {
            id: 'tipoDocumento', rotulo: 'Documento', tipo: 'escolha', metade: true,
            opcoes: [{ valor: 'coren', rotulo: 'COREN-SP' }, { valor: 'cpf', rotulo: 'CPF' }],
            limpa: 'documento', // COREN e CPF são números diferentes: ao trocar, o número é apagado
          },
          {
            id: 'documento', rotulo: 'Número do documento', tipo: 'texto', metade: true,
            atributos: (e) => (usaCpf(e)
              ? { placeholder: '000.000.000-00', inputmode: 'numeric', maxlength: '14' }
              : { placeholder: '000.000', inputmode: null, maxlength: '60' }),
            mascara: (valor, e) => (usaCpf(e) ? F.mascararCpf(valor) : valor),
          },
        ],
        // Dica da linha do documento, exibida abaixo dos dois campos.
        eco: (e) => {
          const doc = documentoDocente(e);
          if (doc.completo) return `Na declaração: “Drª ${F.normalizarEspacos(e.nome) || '…'}, ${doc.texto}, ministrou…”`;
          if (!usaCpf(e)) return 'Digite o número do COREN-SP como deve aparecer na declaração.';
          const faltam = 11 - F.somenteDigitos(e.documento).length;
          if (faltam === 11) return 'Digite o número do CPF.';
          return faltam > 0 ? `Faltam ${faltam} ${faltam === 1 ? 'dígito' : 'dígitos'} do CPF.` : 'CPF inválido. Confira os números.';
        },
      },
      {
        titulo: 'Atividade',
        campos: [
          {
            id: 'programa', rotulo: 'Programa', tipo: 'texto',
            eco: (e) => `Na declaração: “${trecho} ${F.normalizarEspacos(e.programa) || '…'} Mais”`,
            aviso: (e) => (/\bmais$/i.test(F.normalizarEspacos(e.programa))
              ? 'O modelo já tem a palavra “Mais” depois do programa; ela ficaria repetida.' : null),
          },
        ],
      },
    ];
  }

  function tipoDocente(modalidade) {
    const id = `docente-${modalidade}`;
    return {
      id,
      modelo: id,
      perfil: 'docente',
      modalidade,
      rota: `#/docente/${modalidade}`,
      sobretitulo: 'Docente',
      titulo: modalidade === 'online' ? 'Declaração de aula online' : 'Declaração de aula presencial',
      secoes: camposDocente(modalidade),
      tabelaAulas: true,
      notaData: 'A data da assinatura (“São Paulo, [DIA ASSINATURA] de [MÊS ASSINATURA] de [ANO ASSINATURA]”) não é preenchida aqui: ela permanece no documento para ser preenchida manualmente depois.',
      estadoInicial: () => ({ nome: '', tipoDocumento: 'coren', documento: '', programa: '', aulas: [aulaVazia()] }),
      aulaVazia,
      validar: validarDocente,
      valores: valoresDocente,
      preencher: preencherDocente,
      nomeArquivo: (v) => `Declaração docente - aula ${modalidade} - ${v.nome}`,
    };
  }

  // ---------------------------------------------------------------- Aluno

  function validarAluno(e) {
    const erros = {};
    if (!F.normalizarEspacos(e.nome)) erros.nome = 'Informe o nome completo.';
    if (!e.cpf) erros.cpf = 'Informe o CPF.';
    else if (!F.cpfValido(e.cpf)) erros.cpf = 'CPF inválido. Confira os números.';
    const erroData = erroDeData(e.data, 'Informe a data.');
    if (erroData) erros.data = erroData;
    if (!e.inicio) erros.inicio = 'Informe o horário de início.';
    if (!e.fim) erros.fim = 'Informe o horário de término.';
    else if (e.inicio && F.lerHora(e.fim) <= F.lerHora(e.inicio)) erros.fim = 'O término deve ser depois do início.';
    return erros;
  }

  function valoresAluno(e) {
    const digitos = F.somenteDigitos(e.cpf);
    const cpf = digitos ? F.inseparavel(F.mascararCpf(digitos)) : '';
    const completo = digitos.length === 11 && F.cpfValido(digitos);
    return {
      nome: F.normalizarEspacos(e.nome),
      cpf: completo ? cpf : '',
      cpfProvisorio: completo ? '' : cpf,
      data: F.dataNumerica(e.data),
      inicio: F.horaComDoisPontos(e.inicio),
      fim: F.horaComDoisPontos(e.fim),
    };
  }

  function preencherAluno(doc, v, previa) {
    const op = { marcar: previa };
    const texto = D.paragrafos(doc);
    exigir(D.substituir(texto, /\[NOME\]/g, () => v.nome, op), '[NOME]');
    // Na pré-visualização, o CPF ainda incompleto aparece como digitado (destacado como pendente).
    const cpf = () => v.cpf || (previa && v.cpfProvisorio ? { texto: v.cpfProvisorio } : null);
    exigir(D.substituir(texto, /\[000\.000\.000-00\]/g, cpf, op), '[000.000.000-00]');
    exigir(D.substituir(texto, /\[00\/00\/0000\]/g, () => v.data, op), '[00/00/0000]');
    exigir(D.substituir(texto, /\[00:00\]/g, (i) => (i === 0 ? v.inicio : v.fim), op), '[00:00]');
    if (previa) D.substituir(texto, MANUAIS_ALUNO, () => null, { marcar: () => 'manual' });
  }

  const tipoAluno = {
    id: 'aluno',
    modelo: 'aluno',
    perfil: 'aluno',
    rota: '#/aluno',
    sobretitulo: 'Sou aluno',
    titulo: 'Declaração de comparecimento',
    secoes: [
      {
        titulo: 'Dados do aluno',
        campos: [
          { id: 'nome', rotulo: 'Nome completo', tipo: 'texto', autocomplete: 'name' },
          {
            id: 'cpf', rotulo: 'CPF', tipo: 'texto',
            atributos: () => ({ placeholder: '000.000.000-00', inputmode: 'numeric', maxlength: '14' }),
            mascara: (valor) => F.mascararCpf(valor),
          },
        ],
      },
      {
        titulo: 'Comparecimento',
        campos: [
          {
            id: 'data', rotulo: 'Data da atividade', tipo: 'data',
            eco: (e) => (F.dataNumerica(e.data) ? `Na declaração: no dia ${F.dataNumerica(e.data)}` : null),
          },
          { id: 'inicio', rotulo: 'Horário de início', tipo: 'hora', metade: true },
          { id: 'fim', rotulo: 'Horário de término', tipo: 'hora', metade: true },
        ],
      },
    ],
    tabelaAulas: false,
    notaData: 'A data da declaração (“São Paulo, [DIA DECLARAÇÃO] de [MÊS DECLARAÇÃO] de [ANO DECLARAÇÃO]”) não é preenchida aqui: ela permanece no documento para ser preenchida manualmente depois.',
    estadoInicial: () => ({ nome: '', cpf: '', data: '', inicio: '', fim: '' }),
    validar: validarAluno,
    valores: valoresAluno,
    preencher: preencherAluno,
    nomeArquivo: (v) => `Declaração de comparecimento - ${v.nome}`,
  };

  NS.tipos = {
    'docente-presencial': tipoDocente('presencial'),
    'docente-online': tipoDocente('online'),
    aluno: tipoAluno,
  };
})(window.Declaracoes = window.Declaracoes || {});
