/*
 * Gera assets/js/assinatura-protegida.js (assinatura criptografada, publicada no site) a partir de
 * assets/js/assinatura-local.js (fora do Git), com uma senha nova, mostrada uma única vez.
 *
 * Uso, na pasta gerador-declaracoes:  node ferramentas/proteger-assinatura.js
 *
 * O arquivo criptografado é público, então a senha é longa e aleatória: quem o baixar
 * não consegue descobri-la por tentativa e erro.
 */
'use strict';

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const JS = path.join(__dirname, '..', 'assets', 'js');
const ITERACOES = 600000;
const ALFABETO = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'; // sem letras e números que se confundem

function novaSenha() {
  let s = '';
  for (let i = 0; i < 20; i++) s += ALFABETO[crypto.randomInt(ALFABETO.length)];
  return s.match(/.{5}/g).join('-');
}

const fonte = fs.readFileSync(path.join(JS, 'assinatura-local.js'), 'utf8');
const inicio = fonte.indexOf('window.ASSINATURA_LOCAL = ') + 'window.ASSINATURA_LOCAL = '.length;
const dados = JSON.stringify(JSON.parse(fonte.slice(inicio, fonte.lastIndexOf('}') + 1)));

const senha = novaSenha();
const sal = crypto.randomBytes(16);
const iv = crypto.randomBytes(12);
const chave = crypto.pbkdf2Sync(senha, sal, ITERACOES, 32, 'sha256');
const cifra = crypto.createCipheriv('aes-256-gcm', chave, iv);
const cifrado = Buffer.concat([cifra.update(dados, 'utf8'), cifra.final(), cifra.getAuthTag()]);

const conteudo = {
  iteracoes: ITERACOES,
  sal: sal.toString('base64'),
  iv: iv.toString('base64'),
  dados: cifrado.toString('base64'),
};
fs.writeFileSync(
  path.join(JS, 'assinatura-protegida.js'),
  '// Assinatura criptografada: só pode ser usada com a senha. Gerado por ferramentas/proteger-assinatura.js.\n'
    + `window.ASSINATURA_PROTEGIDA = ${JSON.stringify(conteudo)};\n`,
);
console.log(`Senha da assinatura: ${senha}`);
