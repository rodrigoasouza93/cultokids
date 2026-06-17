@/Users/rodrigosouza/.codex/RTK.md

# Instrucoes Para Agentes - Culto Kids

Este projeto organiza a criacao de apresentacoes infantis para aulas biblicas do Culto Kids.

## Antes de criar ou editar apresentacoes

1. Leia `docs/especificacoes-apresentacoes.md`.
2. Leia `docs/fluxo-desenvolvimento-apresentacoes.md`.
3. Leia `regras/regras-de-slides.md`.
4. Use os arquivos em `referencias/` apenas como modelo e referencia; nao altere os originais.
5. Use os assets aprovados em `assets/`, especialmente `assets/personagens/vovo-docura.png` quando a aula pedir a personagem.

## Regras obrigatorias

- Publico: criancas de 4 a 13 anos.
- Ambiente: apresentacao projetada em tela, possivelmente vista de longe.
- Formato padrao: widescreen 16:9.
- Fundo padrao para novos testes: azul estilo tela do Windows (`#0078D7`).
- Fonte minima: 24 pt em todo texto visivel.
- Cor do texto principal: branco.
- Destaques principais: branco em negrito.
- Evite blocos longos; cada slide deve ter uma ideia central.
- Use linguagem simples, respeitosa e biblicamente fiel.
- Nao use imagens assustadoras, violentas, confusas ou com excesso de detalhes.
- Versiculos devem indicar livro, capitulo e versiculo.
- Use cores e imagens em todos ou quase todos os slides para manter a apresentacao atrativa para criancas.
- Cores devem aparecer em fundos, formas, molduras, etiquetas e detalhes; nao use cor clara em texto essencial.

## Estilo visual recomendado

- Visual infantil, alegre e limpo.
- Fundo azul padrao com alto contraste; use sombra discreta, faixa azul mais escura ou ajuste de composicao quando o texto branco perder legibilidade.
- Paleta inspirada nos modelos: amarelo quente, azul, azul claro, rosa, coral, verde, lilas suave e branco.
- Imagens grandes e relevantes para ajudar a memoria da crianca.
- Elementos decorativos podem existir, mas nao devem competir com o texto.

## Fluxo recomendado para novas apresentacoes

1. Preencher `templates/briefing-nova-apresentacao.md`.
2. Criar roteiro com abertura, desenvolvimento, revisao e encerramento.
3. Planejar os slides antes de gerar o PPTX.
4. Criar o PPTX editavel usando o fluxo principal de `docs/fluxo-desenvolvimento-apresentacoes.md`.
5. Renderizar todos os slides e revisar legibilidade: fonte >= 24 pt, branco, negrito nos destaques.
6. Revisar adequacao infantil, fidelidade biblica, fontes, assets e QA visual, usando `templates/qa-apresentacao.md` como base.
7. Importar para Google Slides somente quando isso for pedido ou necessario para colaboracao.

## Criterios de revisao

Uma apresentacao so deve ser considerada pronta quando:

- Todas as criancas conseguem ler os textos principais.
- Cada slide tem pouco texto e mensagem clara.
- A historia ou ensino biblico esta correto.
- As imagens ajudam a explicar ou memorizar.
- Ha momentos de pergunta, repeticao ou participacao.
- O PPTX final e editavel e passou por renderizacao/QA visual.
