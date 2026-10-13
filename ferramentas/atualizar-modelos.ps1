<#
.SYNOPSIS
  Atualiza os modelos usados pelo Gerador de Declarações.

.DESCRIPTION
  Sem parâmetros: embute os arquivos .docx da pasta "modelos" em
  "assets\js\modelos-embutidos.js". Execute sempre que um modelo .docx for alterado.
  Antes de embutir, remove dos arquivos os metadados internos (nomes de autores e
  propriedades do SharePoint); o conteúdo e o layout dos documentos não mudam.

  Com -ConverterOriginais: antes disso, converte os modelos originais (.doc) para .docx
  (pasta "modelos"), sem alterar o layout. O Microsoft Word é usado apenas para LEITURA
  de uma cópia temporária; nenhum arquivo é salvo pelo Word e os originais não são modificados.

.PARAMETER ConverterOriginais
  Converte os modelos .doc originais em .docx antes de embuti-los.

.PARAMETER PastaOriginais
  Pasta onde estão os modelos .doc originais. Padrão: a pasta que contém "gerador-declaracoes".

.EXAMPLE
  .\atualizar-modelos.ps1

.EXAMPLE
  .\atualizar-modelos.ps1 -ConverterOriginais
#>
[CmdletBinding()]
param(
  [switch]$ConverterOriginais,
  [string]$PastaOriginais
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$raiz = Split-Path -Parent $PSScriptRoot
$pastaModelos = Join-Path $raiz 'modelos'
$arquivoJs = Join-Path $raiz 'assets\js\modelos-embutidos.js'
if (-not $PastaOriginais) { $PastaOriginais = Split-Path -Parent $raiz }
$utf8 = New-Object System.Text.UTF8Encoding($false)

# Modelo da aplicação -> arquivo .doc original
$originais = [ordered]@{
  'docente-presencial' = 'Declaração - Modelo para automatização (docentes - presencial).doc'
  'docente-online'     = 'Declaração - Modelo para automatização (docentes - online).doc'
  'aluno'              = 'Declaração de comparecimento para aluno (Modelo automatização).doc'
}

# Marcadores que a aplicação substitui em cada modelo (usados para conferência).
$marcadores = @{
  'docente-presencial' = @('[NOME]', 'Coren-SP 000.000-ENF', '[PROGRAMA]', '[DIA] de [MÊS] de [ANO]', '00/00/0000', '00h00', '[NOME DA AULA]')
  'docente-online'     = @('[NOME]', 'Coren-SP 000.000-ENF', '[PROGRAMA]', '[DIA] de [MÊS] de [ANO]', '00/00/0000', '00h00', '[NOME DA AULA]')
  'aluno'              = @('[NOME]', '[000.000.000-00]', '[00/00/0000]', '[00:00]')
}

function Add-EntradaZip($zip, [string]$nome, [byte[]]$bytes) {
  $entrada = $zip.CreateEntry($nome.TrimStart('/'), [System.IO.Compression.CompressionLevel]::Optimal)
  $fluxo = $entrada.Open()
  try { $fluxo.Write($bytes, 0, $bytes.Length) } finally { $fluxo.Dispose() }
}

# Converte o "Flat OPC" (Document.WordOpenXML) em um pacote .docx, copiando o XML de cada parte sem alterações.
function ConvertTo-Docx([string]$FlatOpc, [string]$Destino) {
  $padrao = '<pkg:part pkg:name="([^"]+)" pkg:contentType="([^"]+)"[^>]*>(?:<pkg:xmlData((?: [^>]*)?)>(.*?)</pkg:xmlData>|<pkg:binaryData>(.*?)</pkg:binaryData>)</pkg:part>'
  $encontradas = [regex]::Matches($FlatOpc, $padrao, [System.Text.RegularExpressions.RegexOptions]::Singleline)
  $total = ([regex]::Matches($FlatOpc, '<pkg:part ')).Count
  if ($encontradas.Count -eq 0 -or $encontradas.Count -ne $total) {
    throw "Não foi possível interpretar o pacote do Word ($($encontradas.Count) de $total partes)."
  }
  $declSim = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`r`n"
  $declNao = '<?xml version="1.0" encoding="UTF-8" standalone="no"?>' + "`r`n"
  $extImagem = @('png', 'jpeg', 'jpg', 'gif', 'emf', 'wmf', 'bmp', 'tif', 'tiff', 'svg')
  $padroes = [ordered]@{ 'rels' = 'application/vnd.openxmlformats-package.relationships+xml'; 'xml' = 'application/xml' }
  $sobrescritas = [ordered]@{}
  $partes = New-Object System.Collections.Generic.List[object]
  foreach ($m in $encontradas) {
    $nome = $m.Groups[1].Value
    $tipo = $m.Groups[2].Value
    if ($m.Groups[4].Success) {
      $decl = $declSim
      if ($m.Groups[3].Value -match 'originalXmlStandalone="no"') { $decl = $declNao }
      $bytes = $utf8.GetBytes($decl + $m.Groups[4].Value)
    } else {
      $bytes = [Convert]::FromBase64String(($m.Groups[5].Value -replace '\s', ''))
    }
    $ext = [System.IO.Path]::GetExtension($nome).TrimStart('.').ToLowerInvariant()
    if ($extImagem -contains $ext) {
      if (-not $padroes.Contains($ext)) { $padroes[$ext] = $tipo }
    } elseif (-not ($ext -eq 'rels' -or ($ext -eq 'xml' -and $tipo -eq 'application/xml'))) {
      $sobrescritas[$nome] = $tipo
    }
    $partes.Add([pscustomobject]@{ Nome = $nome; Bytes = $bytes })
  }
  $ct = New-Object System.Text.StringBuilder
  [void]$ct.Append($declSim).Append('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
  foreach ($k in $padroes.Keys) {
    [void]$ct.Append('<Default Extension="' + [System.Security.SecurityElement]::Escape($k) + '" ContentType="' + [System.Security.SecurityElement]::Escape($padroes[$k]) + '"/>')
  }
  foreach ($k in $sobrescritas.Keys) {
    [void]$ct.Append('<Override PartName="' + [System.Security.SecurityElement]::Escape($k) + '" ContentType="' + [System.Security.SecurityElement]::Escape($sobrescritas[$k]) + '"/>')
  }
  [void]$ct.Append('</Types>')

  $fluxo = [System.IO.File]::Open($Destino, [System.IO.FileMode]::Create)
  try {
    $zip = New-Object System.IO.Compression.ZipArchive($fluxo, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
      Add-EntradaZip $zip '[Content_Types].xml' ($utf8.GetBytes($ct.ToString()))
      foreach ($primeira in @('/_rels/.rels', '/word/document.xml')) {
        foreach ($p in $partes) { if ($p.Nome -eq $primeira) { Add-EntradaZip $zip $p.Nome $p.Bytes } }
      }
      foreach ($p in $partes) {
        if ($p.Nome -ne '/_rels/.rels' -and $p.Nome -ne '/word/document.xml') { Add-EntradaZip $zip $p.Nome $p.Bytes }
      }
    } finally { $zip.Dispose() }
  } finally { $fluxo.Dispose() }
}

function Convert-Originais {
  $temp = Join-Path ([System.IO.Path]::GetTempPath()) ('gerador-declaracoes-' + [guid]::NewGuid().ToString('N'))
  New-Item -ItemType Directory -Force -Path $temp | Out-Null
  $word = $null
  try {
    $copias = [ordered]@{}
    foreach ($id in $originais.Keys) {
      $origem = Join-Path $PastaOriginais $originais[$id]
      if (-not (Test-Path -LiteralPath $origem)) { throw "Modelo original não encontrado: $origem" }
      $copia = Join-Path $temp ($id + '.doc')
      Copy-Item -LiteralPath $origem -Destination $copia
      $copias[$id] = $copia
    }
    Write-Host 'Lendo os modelos no Word (nada será salvo)...'
    $antes = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $word = New-Object -ComObject Word.Application
    $pidWord = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $antes -notcontains $_.Id } | ForEach-Object { $_.Id })
    $word.Visible = $false
    $word.DisplayAlerts = 0
    foreach ($id in $copias.Keys) {
      # A cópia temporária é aberta em modo normal (não "somente leitura") para que as propriedades
      # do documento não sejam marcadas como somente leitura; ela é fechada sem salvar e depois apagada.
      $doc = $word.Documents.Open($copias[$id], $false, $false, $false)
      try { $flat = $doc.WordOpenXML } finally { $doc.Close(0) }
      $destino = Join-Path $pastaModelos ($id + '.docx')
      ConvertTo-Docx $flat $destino
      Write-Host ("  Convertido: {0} -> modelos\{1}.docx" -f $originais[$id], $id)
    }
  } finally {
    if ($word) {
      # Se outro documento foi aberto nesta instância enquanto o script rodava (por exemplo, um
      # duplo clique em um .doc), o Word não é fechado: ele fica visível para o usuário.
      $outros = 0
      try { $outros = $word.Documents.Count } catch { }
      if ($outros -gt 0) {
        try { $word.Visible = $true } catch { }
        Write-Warning 'O Word foi mantido aberto porque há outro documento aberto nele.'
        $pidWord = @()
      } else {
        try { $word.Quit() } catch { }
      }
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()
      # Garante que a instância do Word aberta por este script foi encerrada.
      foreach ($p in $pidWord) {
        $proc = Get-Process -Id $p -ErrorAction SilentlyContinue
        if ($proc -and -not $proc.WaitForExit(5000)) { Stop-Process -Id $p -Force -ErrorAction SilentlyContinue }
      }
    }
    Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
  }
}

# Texto do documento (concatena os trechos de cada parágrafo) para conferir os marcadores.
function Get-TextoDocx([string]$Caminho) {
  $zip = [System.IO.Compression.ZipFile]::OpenRead($Caminho)
  try {
    $entrada = $zip.GetEntry('word/document.xml')
    if (-not $entrada) { throw "Arquivo inválido (sem word/document.xml): $Caminho" }
    $leitor = New-Object System.IO.StreamReader($entrada.Open(), $utf8)
    try { $xml = $leitor.ReadToEnd() } finally { $leitor.Dispose() }
  } finally { $zip.Dispose() }
  $sb = New-Object System.Text.StringBuilder
  foreach ($m in [regex]::Matches($xml, '<w:t(?: [^>]*)?>([^<]*)</w:t>|</w:p>')) {
    if ($m.Value -eq '</w:p>') { [void]$sb.Append("`n") } else { [void]$sb.Append($m.Groups[1].Value) }
  }
  return [System.Net.WebUtility]::HtmlDecode($sb.ToString())
}

# Remove metadados internos do arquivo (nomes de autores e propriedades do SharePoint),
# sem alterar o conteúdo nem o layout do documento.
function Clear-Metadados([string]$Caminho) {
  $partes = New-Object System.Collections.Generic.List[object]
  $zip = [System.IO.Compression.ZipFile]::OpenRead($Caminho)
  try {
    foreach ($e in $zip.Entries) {
      $ms = New-Object System.IO.MemoryStream
      $s = $e.Open()
      try { $s.CopyTo($ms) } finally { $s.Dispose() }
      $partes.Add([pscustomobject]@{ Nome = $e.FullName; Bytes = $ms.ToArray() })
    }
  } finally { $zip.Dispose() }

  # Já está limpo? Então não reescreve o arquivo.
  $sujo = $false
  foreach ($p in $partes) {
    if ($p.Nome -like 'customXml/*' -or $p.Nome -eq 'docProps/custom.xml') { $sujo = $true }
    if ($p.Nome -eq 'docProps/core.xml' -and $utf8.GetString($p.Bytes) -match '<dc:creator>[^<]+</dc:creator>|<cp:lastModifiedBy>[^<]+</cp:lastModifiedBy>') { $sujo = $true }
  }
  if (-not $sujo) { return }

  $fluxo = [System.IO.File]::Open($Caminho, [System.IO.FileMode]::Create)
  try {
    $novo = New-Object System.IO.Compression.ZipArchive($fluxo, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
      foreach ($p in $partes) {
        if ($p.Nome -like 'customXml/*' -or $p.Nome -eq 'docProps/custom.xml') { continue }
        $bytes = $p.Bytes
        $texto = $null
        switch ($p.Nome) {
          '[Content_Types].xml' { $texto = $utf8.GetString($bytes) -replace '<Override PartName="/(customXml/[^"]*|docProps/custom\.xml)"[^>]*/>', '' }
          '_rels/.rels' { $texto = $utf8.GetString($bytes) -replace '<Relationship [^>]*Target="/?docProps/custom\.xml"[^>]*/>', '' }
          'word/_rels/document.xml.rels' { $texto = $utf8.GetString($bytes) -replace '<Relationship [^>]*Target="[^"]*customXml/[^"]*"[^>]*/>', '' }
          'docProps/core.xml' {
            $texto = ($utf8.GetString($bytes) -replace '<dc:creator>[^<]*</dc:creator>', '<dc:creator></dc:creator>') -replace '<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>', '<cp:lastModifiedBy></cp:lastModifiedBy>'
          }
        }
        if ($null -ne $texto) { $bytes = $utf8.GetBytes($texto) }
        Add-EntradaZip $novo $p.Nome $bytes
      }
    } finally { $novo.Dispose() }
  } finally { $fluxo.Dispose() }
}

function Write-ModelosEmbutidos {
  $arquivos = @(Get-ChildItem -LiteralPath $pastaModelos -Filter '*.docx' | Where-Object { -not $_.Name.StartsWith('~$') } | Sort-Object Name)
  if ($arquivos.Count -eq 0) { throw "Nenhum arquivo .docx encontrado em $pastaModelos" }
  foreach ($a in $arquivos) { Clear-Metadados $a.FullName }
  $arquivos = @(Get-ChildItem -LiteralPath $pastaModelos -Filter '*.docx' | Where-Object { -not $_.Name.StartsWith('~$') } | Sort-Object Name)
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine('/* Gerado automaticamente por ferramentas/atualizar-modelos.ps1 em ' + (Get-Date -Format 'dd/MM/yyyy HH:mm') + '.')
  [void]$sb.AppendLine('   Não edite este arquivo: altere os modelos da pasta "modelos" e execute a ferramenta novamente. */')
  [void]$sb.AppendLine('window.MODELOS_EMBUTIDOS = {')
  foreach ($a in $arquivos) {
    $id = [System.IO.Path]::GetFileNameWithoutExtension($a.Name)
    $texto = Get-TextoDocx $a.FullName
    if ($marcadores.ContainsKey($id)) {
      $faltando = @($marcadores[$id] | Where-Object { -not $texto.Contains($_) })
      if ($faltando.Count -gt 0) {
        Write-Warning ("O modelo {0}.docx não contém: {1}. Esses campos não serão preenchidos." -f $id, ($faltando -join ', '))
      }
    }
    $b64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($a.FullName))
    [void]$sb.AppendLine('  "' + $id + '": "' + $b64 + '",')
    Write-Host ("  Embutido: modelos\{0} ({1:N0} KB)" -f $a.Name, ($a.Length / 1KB))
  }
  [void]$sb.AppendLine('};')
  [System.IO.File]::WriteAllText($arquivoJs, $sb.ToString(), $utf8)
  Write-Host "Arquivo atualizado: assets\js\modelos-embutidos.js"
}

New-Item -ItemType Directory -Force -Path $pastaModelos | Out-Null
if ($ConverterOriginais) { Convert-Originais }
Write-ModelosEmbutidos
Write-Host 'Concluído.'
