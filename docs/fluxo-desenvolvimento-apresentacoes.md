# Fluxo de Desenvolvimento de Apresentacoes

Este fluxo padroniza como criar, revisar e entregar apresentacoes do Culto Kids. Ele complementa `docs/especificacoes-apresentacoes.md` e `regras/regras-de-slides.md`.

## Objetivo do fluxo

- Produzir apresentacoes infantis em PPTX editavel.
- Manter textos legiveis em projetor, com fonte minima de 24 pt.
- Usar imagens e cores como apoio de memoria para as criancas.
- Registrar fontes, assets e revisoes para facilitar manutencao.
- Permitir publicacao em Google Slides quando necessario, sem perder a qualidade do PPTX original.

## Ferramentas recomendadas

### Fluxo principal: PPTX editavel

Use a skill `presentations` para criar, editar, renderizar, revisar e exportar arquivos `.pptx`.

Regras para esse caminho:

- O entregavel final deve ser PPTX editavel, com texto, formas e imagens como elementos editaveis.
- Nao use slides finais como uma imagem unica de tela inteira, salvo pedido explicito.
- Renderize todos os slides antes de considerar a apresentacao pronta.
- Corrija sobreposicoes, texto cortado e contraste fraco antes da entrega.

### Assets visuais

Use `imagegen` quando for necessario criar imagens novas, por exemplo:

- cenas biblicas em estilo infantil;
- variacoes visuais da Vovo Docura;
- objetos simbolicos da aula;
- fundos ou ilustracoes especificas que nao existam em `assets/`.

Regras para imagens geradas:

- Registrar no roteiro ou no registro de QA que a imagem foi gerada.
- Guardar o arquivo final em `assets/gerados/<data-ou-slug-da-aula>/`.
- Evitar imagens assustadoras, violentas, escuras ou detalhadas demais.
- Preferir composicoes simples, grandes e claras.

### QA visual e PDF

Use renderizacao de slides e, quando util, a skill `pdf` para gerar uma visualizacao revisavel.

O QA deve verificar:

- fonte minima de 24 pt;
- texto principal branco;
- destaques em negrito;
- uma ideia principal por slide;
- contraste bom no fundo azul ou em faixa de apoio;
- nenhuma sobreposicao indevida;
- imagens apropriadas para criancas;
- versiculos com referencia completa.

### Google Slides e Google Drive

Use `google-drive:google-slides` e o MCP do Google Drive apenas quando o destino precisar ser Google Slides, revisao colaborativa ou armazenamento no Drive.

Pasta oficial de apresentacoes criadas:

- Google Drive: `https://drive.google.com/drive/folders/1SWf0gtLBg7CgS0hJ2XX0BhUjRCZyFWUo?usp=drive_link`
- ID da pasta: `1SWf0gtLBg7CgS0hJ2XX0BhUjRCZyFWUo`

Fluxo recomendado:

1. Criar primeiro o PPTX local com `presentations`.
2. Verificar visualmente o PPTX local.
3. Importar para Google Slides quando necessario, preferencialmente dentro da pasta oficial do Drive.
4. Fazer leitura de volta pelo conector para confirmar titulo, quantidade de slides e conteudo.
5. Verificar thumbnails dos slides alterados quando houver edicao direta no Google Slides.

Nao crie um deck novo diretamente no Google Slides se o objetivo for alta qualidade visual; prefira PPTX local e importacao.

### Planilhas de planejamento

Use planilhas somente quando ajudarem a organizar aulas maiores ou series. Uma planilha pode conter:

- numero do slide;
- objetivo do slide;
- texto principal;
- versiculo;
- imagem planejada;
- asset usado;
- ponto de participacao;
- status de QA.

## Etapas obrigatorias

1. Ler as especificacoes e regras do projeto.
2. Preencher o briefing em `templates/briefing-nova-apresentacao.md`.
3. Criar ou atualizar o roteiro da aula.
4. Planejar os slides antes de gerar o PPTX.
5. Separar ou gerar assets visuais.
6. Gerar o PPTX editavel.
7. Renderizar e revisar todos os slides.
8. Corrigir problemas visuais ou pedagogicos.
9. Registrar QA e fontes.
10. Entregar o PPTX final e, se solicitado, publicar ou importar a versao Google Slides na pasta oficial do Drive.
11. Quando houver publicacao no Drive, criar ou atualizar `publicacao.json` e registrar o commit Git correspondente.

## Registro minimo por apresentacao

Cada pasta em `apresentacoes/<data-slug>/` deve ter:

- `roteiro.md`: roteiro e plano de slides.
- `gerar-pptx.js` ou script equivalente: geracao do PPTX.
- `package.json`: dependencias locais quando houver.
- arquivo `.pptx` final.
- registro de QA, preferencialmente em `qa.md` usando `templates/qa-apresentacao.md` ou secao equivalente no `roteiro.md`.

Quando houver assets gerados, a pasta correspondente em `assets/gerados/` deve ter nomes claros e reutilizaveis.

Quando a apresentacao for publicada no Google Drive, a mesma pasta tambem deve conter `publicacao.json`, baseado em `templates/publicacao.json`. O manifesto vincula a fonte oficial no GitHub ao artefato remoto e sinaliza edicoes mobile pendentes de reconciliacao. Consulte `docs/operacao-remota-e-drive.md`.

## Fontes confiaveis

Para fatos biblicos, versiculos e referencias:

- usar a Biblia e materiais confiaveis da igreja como fonte principal;
- registrar livro, capitulo e versiculo;
- nao inventar citacoes, datas, mapas ou detalhes historicos;
- quando usar fonte externa, preferir editoras biblicas, sociedades biblicas, documentacao oficial ou material fornecido pelo professor.

Para ferramentas e APIs:

- preferir documentacao oficial do MCP, Google Slides API, Google Drive e ferramentas usadas no projeto;
- evitar depender de skills pouco instaladas ou fontes sem manutencao clara;
- registrar quando uma ferramenta externa foi usada para gerar, importar ou revisar.

## Quando usar Slidev

Slidev pode ser usado para rascunhos interativos, aulas web ou prototipos em Markdown. Ele nao deve ser o fluxo final padrao para PPTX do Culto Kids, porque a exportacao PPTX do Slidev pode transformar os slides em imagens e reduzir a editabilidade do texto.

## Definicao de pronto

Uma apresentacao so esta pronta quando:

- o PPTX final existe e e editavel;
- todos os slides foram renderizados e revisados;
- os textos visiveis respeitam 24 pt ou mais;
- o conteudo e adequado para criancas de 4 a 13 anos;
- os versiculos estao corretos e com referencia;
- as imagens ajudam a memoria e nao competem com o texto;
- o registro de QA nao tem pendencias criticas.
