/*
 * Tipos de declaração: campos do formulário, validação e preenchimento de cada modelo.
 *
 * Marcadores substituídos (exatamente como estão nos modelos):
 *   Docente (presencial e online): [NOME], 000.000-ENF, [PROGRAMA], [DIA], [MÊS], [ANO]
 *     e, na tabela de aulas: 00/00/0000, 00h00 (início), 00h00 (término), [NOME DA AULA]
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

  function validarDocente(e) {
    const erros = {};
    if (!F.normalizarEspacos(e.nome)) erros.nome = 'Informe o nome completo.';
    if (!F.normalizarEspacos(e.coren)) erros.coren = 'Informe o número do COREN-SP.';
    if (!F.normalizarEspacos(e.programa)) erros.programa = 'Informe o programa.';
    const erroData = erroDeData(e.data, 'Informe a data da atividade.');
    if (erroData) erros.data = erroData;
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
    const extenso = F.dataPorExtenso(e.data) || {};
    return {
      nome: F.normalizarEspacos(e.nome),
      coren: F.formatarCoren(e.coren),
      programa: F.normalizarEspacos(e.programa),
      dia: extenso.dia || '',
      mes: extenso.mes || '',
      ano: extenso.ano || '',
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
    exigir(D.substituir(texto, /000\.000-ENF/g, () => v.coren, op), '000.000-ENF');
    exigir(D.substituir(texto, /\[PROGRAMA\]/g, () => v.programa, op), '[PROGRAMA]');
    exigir(D.substituir(texto, /\[DIA\]/g, () => v.dia, op), '[DIA]');
    exigir(D.substituir(texto, /\[MÊS\]/g, () => v.mes, op), '[MÊS]');
    exigir(D.substituir(texto, /\[ANO\]/g, () => v.ano, op), '[ANO]');

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
            id: 'coren', rotulo: 'COREN-SP', tipo: 'coren', placeholder: 'Ex.: 123456 ou 123.456-ENF',
            eco: (e) => (F.normalizarEspacos(e.coren) ? `Na declaração: Coren-SP ${F.formatarCoren(e.coren)}` : 'Somente números: o sufixo -ENF do modelo é mantido.'),
          },
        ],
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
          {
            id: 'data', rotulo: 'Data da atividade', tipo: 'data',
            eco: (e) => {
              const x = F.dataPorExtenso(e.data);
              return x ? `Na declaração: ${x.dia} de ${x.mes} de ${x.ano}` : 'Escrita por extenso: dia, mês e ano.';
            },
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
      estadoInicial: () => ({ nome: '', coren: '', programa: '', data: '', aulas: [aulaVazia()] }),
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
    return {
      nome: F.normalizarEspacos(e.nome),
      cpf: F.somenteDigitos(e.cpf).length === 11 ? F.mascararCpf(e.cpf) : '',
      data: F.dataNumerica(e.data),
      inicio: F.horaComDoisPontos(e.inicio),
      fim: F.horaComDoisPontos(e.fim),
    };
  }

  function preencherAluno(doc, v, previa) {
    const op = { marcar: previa };
    const texto = D.paragrafos(doc);
    exigir(D.substituir(texto, /\[NOME\]/g, () => v.nome, op), '[NOME]');
    exigir(D.substituir(texto, /\[000\.000\.000-00\]/g, () => v.cpf, op), '[000.000.000-00]');
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
          { id: 'cpf', rotulo: 'CPF', tipo: 'cpf', placeholder: '000.000.000-00' },
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
