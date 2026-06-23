# Culto Kids - Apresentacoes

Projeto para centralizar diretrizes, referencias e materiais reutilizaveis para criacao de apresentacoes do Culto Kids.

O publico principal sao criancas de 4 a 13 anos em aulas sobre a Biblia e a Igreja Evangelica. As apresentacoes sao usadas em projetor, por isso a prioridade e leitura facil, texto branco, destaques em negrito e fonte grande.

Para os proximos testes, o fundo padrao dos slides deve ser azul estilo tela do Windows (`#0078D7`), usando texto branco e negrito nos destaques para aumentar o contraste.

## Estrutura

- `docs/`: especificacoes, guias editoriais e analises de modelos.
- `regras/`: regras objetivas para criacao e revisao de slides.
- `templates/`: modelos de briefing, QA e roteiros para novas apresentacoes.
- `assets/`: imagens aprovadas para uso nas aulas.
- `referencias/`: apresentacoes e materiais usados como exemplo visual ou pedagogico.

## Referencias iniciais

- Apresentacao modelo: `referencias/apresentacoes/01-historia-da-biblia.pptx`
- Apresentacao modelo: `referencias/apresentacoes/02-versiculos-pentateuco.pptx`
- Apresentacao modelo: `referencias/apresentacoes/03-versiculos-livros-historicos.pptx`
- Personagem Vovo Docura: `assets/personagens/vovo-docura.png`

## Principios do Culto Kids

1. Ensinar a Biblia com linguagem simples, fiel e acolhedora.
2. Priorizar criancas que estao vendo de longe em ambiente com projetor.
3. Usar imagens como apoio de memoria, nao como enfeite vazio.
4. Usar cores alegres e infantis para tornar a aula atrativa.
5. Manter cada slide com uma ideia principal.
6. Criar ritmo infantil: visual alegre, texto curto, repeticao didatica e convite a participacao.

## Fluxo padrao de desenvolvimento

O fluxo oficial esta em `docs/fluxo-desenvolvimento-apresentacoes.md`.

Pasta oficial para guardar apresentacoes criadas no Google Drive:

- https://drive.google.com/drive/folders/1SWf0gtLBg7CgS0hJ2XX0BhUjRCZyFWUo?usp=drive_link

Resumo:

1. Preencher o briefing.
2. Criar roteiro e plano de slides.
3. Gerar PPTX editavel com o fluxo principal de apresentacoes.
4. Usar assets aprovados ou gerar novos assets em `assets/gerados/`.
5. Renderizar todos os slides e fazer QA visual.
6. Registrar fontes, assets usados e pendencias.
7. Importar para Google Slides na pasta oficial do Drive apenas quando houver necessidade de colaboracao ou entrega nesse formato.

## Trabalho pelo telefone e publicacao

O GitHub e a fonte oficial das apresentacoes; o Google Drive e o ponto de acesso e colaboracao pelo telefone. Cada apresentacao nova deve incluir um `publicacao.json`, baseado em `templates/publicacao.json`, que registra a versao publicada no Drive e seu commit de origem.

Leia `docs/operacao-remota-e-drive.md` antes de publicar, editar pelo Google Slides ou reconciliar uma alteracao feita no celular.

## Regra rapida de legibilidade

Todo texto deve ter no minimo 24 pt. Textos principais devem ser brancos e em negrito quando forem titulo, versiculo, palavra-chave ou chamada de participacao.
