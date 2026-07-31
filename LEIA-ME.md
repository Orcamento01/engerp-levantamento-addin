# ENGERP Levantamento — Suplemento do Excel

## O que isso é

Um suplemento (Office Add-in) que coloca um botão "Levantamento" na aba Página Inicial do Excel.
Ao clicar, abre a mesma ferramenta de leitura de plantas (PDF/DXF/imagem) que vocês já usam, numa
janela grande. Ao terminar o levantamento, um botão "📤 Enviar para o Excel" escreve os dados
direto nas abas da planilha aberta (Memorial de Cálculo, Paredes, Esquadrias, Louças-Metais-Luminárias)
— sem precisar baixar e abrir um arquivo separado.

A ferramenta continua funcionando normalmente também fora do Excel (abrindo dialog.html direto
no navegador) — o botão "Baixar planilha (.xlsx)" que já existia continua lá.

## Arquivos deste repositório

- manifest.xml — o "cartão de identidade" do suplemento (já preenchido com a URL real deste repositório)
- taskpane.html / taskpane.js — o painel lateral que abre dentro do Excel
- dialog.html — a ferramenta completa (idêntica à que vocês já usam)
- icone-16.png / icone-32.png / icone-80.png — ícones do suplemento

## Status atual

- Repositório criado e a maior parte dos arquivos já commitada por aqui mesmo, direto do navegador.
- Falta ativar o GitHub Pages (Settings → Pages → Branch: main) para os arquivos ficarem acessíveis em HTTPS.
- Os 3 ícones PNG podem precisar ser enviados manualmente (arraste os arquivos na página de upload do repositório) caso não tenham sido enviados automaticamente.

