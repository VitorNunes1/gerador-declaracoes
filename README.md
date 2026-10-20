# Gerador de declarações · Coren-SP Educação

Aplicação web para preencher e gerar automaticamente as declarações de presença de **docentes**
(aula presencial e aula online) e de **alunos** (comparecimento), a partir dos modelos oficiais.

**Acesse:** https://vitornunes1.github.io/gerador-declaracoes/

## Como usar

1. Acesse o endereço acima (ou abra o arquivo `index.html` desta pasta) no Chrome, Edge ou Firefox.
2. Escolha a declaração:
   - **Sou docente** → *Declaração de aula presencial* ou *Declaração de aula online*;
   - **Sou aluno** → *Declaração de comparecimento*.
3. Preencha o formulário. A pré-visualização ao lado mostra a declaração sendo montada.
4. Clique em **Gerar declaração (.docx)**. O arquivo é salvo na pasta *Downloads*.
5. Abra o arquivo no Word para preencher a data da assinatura (ou da declaração) e assinar.

Não é preciso instalar nada: tudo funciona dentro do navegador, inclusive sem internet
quando aberto a partir do `index.html`.

## Campos preenchidos

| Declaração | Campo do formulário | No modelo | Como aparece |
|---|---|---|---|
| Docente (presencial e online) | Nome completo | `[NOME]` | como digitado (mantém o negrito do modelo) |
| | Documento (COREN-SP ou CPF) e número | `Coren-SP 000.000-ENF` | `Coren-SP 123456` (número como digitado, sem acréscimos) ou `CPF 529.982.247-25` (CPF conferido) |
| | Programa | `[PROGRAMA]` | como digitado (a palavra "Mais" depois do programa já faz parte do modelo) |
| | Aulas (tabela): data, início, término, nome | `00/00/0000`, `00h00`, `00h00`, `[NOME DA AULA]` | uma linha da tabela para cada aula: `08/10/2026`, `08h00`, `12h00`, nome |
| Aluno | Nome completo | `[NOME]` | como digitado (negrito) |
| | CPF | `[000.000.000-00]` | `529.982.247-25` (negrito; os dígitos verificadores são conferidos) |
| | Nome do curso | (novo trecho) | `Sutura`, em "…da aula do curso de Sutura para Enfermeiros…" |
| | Data da atividade | `[00/00/0000]` | `08/10/2026` |
| | Horário de início e de término | `[00:00]` às `[00:00]` | `14:00` às `15:30` |

Nas declarações de docente, o trecho `no dia [DIA] de [MÊS] de [ANO]` (no modelo online,
`[DIA] de [MÊS] de [ANO]`) é retirado: a frase termina em "…, conforme abaixo.", e as datas
ficam apenas na tabela de aulas.

Na declaração de aluno, a frase do modelo é ajustada na hora de gerar, mantendo a formatação:
"Declaramos, para os fins, que **[nome]**, inscrito(a) no CPF sob nº **[CPF]**, participou da aula
do curso de [curso] para Enfermeiros realizada neste Conselho no dia [data], no período das
[início] às [término]."

A **data da assinatura** (docentes) e a **data da declaração** (aluno) não têm campo no formulário:
ficam no documento como no modelo, para preenchimento manual.

## Fidelidade aos modelos

- O arquivo gerado é o **próprio modelo original**: os `.doc` foram convertidos para `.docx` sem
  alteração de layout (o documento continua em modo de compatibilidade, como o original).
  Ao gerar, apenas o texto dos marcadores é trocado, herdando a formatação do modelo
  (fonte, tamanho, negrito, alinhamento). Cabeçalho, rodapé, imagens e textos fixos não são tocados.
- As linhas adicionais da tabela de aulas são cópias exatas da linha modelo.
- A pré-visualização é uma aproximação para conferência; o documento final é sempre o `.docx`.

## Privacidade

Os dados digitados não são enviados a nenhum servidor: a declaração é montada dentro do próprio
navegador, mesmo quando a aplicação é acessada pelo site. O que é digitado fica guardado apenas na
aba aberta, para não se perder ao atualizar a página, e é descartado quando a aba é fechada.
Em computadores compartilhados, use **Limpar formulário** ao terminar.

## Site público

O site é publicado pelo GitHub Pages a partir da branch `main` deste repositório: cada alteração
enviada ao GitHub (push) aparece no site em cerca de um minuto. Links diretos para os formulários:

- Docente, aula presencial: https://vitornunes1.github.io/gerador-declaracoes/#/docente/presencial
- Docente, aula online: https://vitornunes1.github.io/gerador-declaracoes/#/docente/online
- Aluno: https://vitornunes1.github.io/gerador-declaracoes/#/aluno

Por ser um site estático, a pasta também pode ser copiada para qualquer outro servidor web.

O navegador guarda os arquivos do site por alguns minutos. Para que uma atualização chegue a
todos, os endereços dos arquivos no `index.html` levam um número de versão (`?v=...`): ao alterar
o código, troque esse número (a ferramenta de modelos faz isso sozinha). Quem estiver com o site
aberto vê a novidade ao recarregar a página (Ctrl + F5).

## Atualizar os modelos

Os modelos usados ficam em `modelos/` e são embutidos em `assets/js/modelos-embutidos.js`
(o que permite abrir a aplicação direto do disco).

- **Alterou os `.doc` originais** (na pasta acima de `gerador-declaracoes`):
  execute `ferramentas\atualizar-modelos.bat -ConverterOriginais` (requer o Microsoft Word;
  ele é usado só para ler uma cópia dos arquivos, nada é salvo pelo Word).
- **Editou diretamente um `.docx` da pasta `modelos`**: execute `ferramentas\atualizar-modelos.bat`.

Os marcadores da tabela acima precisam continuar no modelo, escritos exatamente da mesma forma.
A ferramenta avisa se algum deles estiver faltando e também remove dos arquivos os metadados
internos (nomes de autores e propriedades do SharePoint). Depois, envie as alterações ao GitHub
para atualizar o site.

## Estrutura

```
gerador-declaracoes/
├── index.html                    página da aplicação
├── assets/
│   ├── css/app.css               aparência
│   ├── js/formatacao.js          datas, horários, CPF e COREN
│   ├── js/docx.js                preenchimento do modelo Word
│   ├── js/declaracoes.js         campos, validação e marcadores de cada declaração
│   ├── js/previa.js              pré-visualização da página
│   ├── js/app.js                 telas e formulários
│   ├── js/modelos-embutidos.js   modelos .docx embutidos (gerado pela ferramenta)
│   ├── img/logo-coren-sp.png     logotipo (extraído do modelo)
│   └── vendor/jszip.min.js       biblioteca JSZip 3.10.1 (licença MIT)
├── modelos/                      modelos .docx convertidos dos originais
└── ferramentas/                  atualização dos modelos
```

## Observações sobre os modelos originais

Os itens abaixo estão nos modelos e foram mantidos como estão, porque a regra é não alterar nada
além dos campos:

- Modelo online: a data da assinatura está escrita `MÊS ASSINATURA]`, sem o colchete inicial.
- Modelo presencial: há um ponto (`.`) sozinho na linha logo abaixo da tabela.
- Modelos de docente: `Drª` e a palavra `Mais` depois do programa são texto fixo.
- Com muitas aulas (cerca de 20 ou mais), a declaração passa para duas páginas e a linha
  "São Paulo, [DIA ASSINATURA]…" pode encostar na faixa azul do rodapé, como acontece no próprio
  modelo do Word. Nesses casos, ajuste o espaçamento no Word antes de imprimir.
