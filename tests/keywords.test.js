const test = require("node:test");
const assert = require("node:assert");
const { encontrarPalavrasChave } = require("../src/commands/commands.js");

test("detecta palavras no assunto e no corpo, sem acento e sem caixa", () => {
  assert.deepStrictEqual(encontrarPalavrasChave("CADASTRO de cliente", ""), ["cadastro"]);
  assert.deepStrictEqual(encontrarPalavrasChave("", "Favor inserir o CPF."), ["inserir"]);
  assert.deepStrictEqual(encontrarPalavrasChave("Inserção", "insercao e Cadastramento"), ["insercao", "cadastramento"]);
  assert.deepStrictEqual(encontrarPalavrasChave("Cadastrar fornecedor", ""), ["cadastrar"]);
  assert.deepStrictEqual(encontrarPalavrasChave("Novos cadastros", "duas inserções"), ["cadastros", "insercoes"]);
});

test("não dispara sem palavras-chave nem em palavras derivadas", () => {
  assert.deepStrictEqual(encontrarPalavrasChave("Reunião amanhã", "Segue a pauta."), []);
  assert.deepStrictEqual(encontrarPalavrasChave("Usuário cadastrado", "reinserir depois"), []);
  assert.deepStrictEqual(encontrarPalavrasChave(null, undefined), []);
});
