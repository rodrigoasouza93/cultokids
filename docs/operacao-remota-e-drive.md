# Operação remota: GitHub, Google Drive e celular

Este guia permite continuar o trabalho das apresentações pelo telefone sem perder a fonte oficial do projeto.

## Fonte oficial

Para cada aula, a fonte oficial é sempre o conteúdo versionado no repositório:

- `roteiro.md` para conteúdo, plano de slides e referências;
- `gerar-pptx.js` para a composição editável;
- assets em `assets/`;
- `publicacao.json` para o vínculo com a cópia publicada no Google Drive.

O arquivo `.pptx` e o Google Slides são entregáveis publicados. Eles não substituem o roteiro e o gerador como fonte oficial.

## Publicar uma apresentação

1. Criar ou atualizar `publicacao.json` a partir de `templates/publicacao.json` dentro da pasta da apresentação.
2. Gerar o PPTX e concluir a renderização e o QA visual.
3. Publicar o PPTX na pasta oficial do Google Drive.
4. Preencher no manifesto o ID, a URL, o formato, a data de publicação e o commit Git da versão publicada.
5. Fazer commit do manifesto junto com as fontes alteradas.

Use Google Slides somente quando a colaboração ou a edição mobile realmente for necessária. Para a melhor fidelidade visual, mantenha o PPTX validado como artefato de entrega.

## Estrutura no Google Drive

A pasta oficial continua sendo o ponto de entrada do projeto. As pastas abaixo organizam apenas as novas publicações; arquivos antigos não devem ser movidos sem uma revisão deliberada.

- [Em edição (Google Slides)](https://drive.google.com/drive/folders/1JZ3Fq6QsWrJtEQwM-ADguljxm7t3cNYM): cópias nativas do Google Slides que podem ser alteradas pelo telefone.
- [Prontas para apresentar (PPTX e PDF)](https://drive.google.com/drive/folders/16zZ2EWZ0e0hFdN_L5JZzZ5NkVYsbetD6): PPTX validado e, quando solicitado, PDF de apoio para projetor ou leitura.

Para uma nova publicação, use a pasta correspondente ao formato e registre o link final no manifesto.

## Pelo telefone

### Alteração planejada

Descreva a alteração em um chat ligado ao projeto ou abra uma issue no GitHub. Informe o slug da aula e o resultado esperado, por exemplo:

> Em `2026-06-29-jogando-limpo`, substitua a pergunta do slide 7 por uma pergunta sobre honestidade. Mantenha o estilo e a fonte mínima de 24 pt.

A alteração deve ser feita no `roteiro.md` e/ou no `gerar-pptx.js`, depois publicada novamente no Drive.

### Alteração urgente em Google Slides

1. Edite a versão publicada no aplicativo Google Slides.
2. Registre no `publicacao.json` ou em uma nota no `roteiro.md` quais slides foram alterados e o que mudou.
3. Na próxima sessão no computador, reconcilie a alteração no `roteiro.md` e no `gerar-pptx.js` antes de gerar uma nova versão.

Uma edição direta no Slides não retorna automaticamente para o código. Nunca gere e publique novamente sem antes reconciliar uma alteração mobile conhecida.

## Estados do manifesto

- `rascunho`: ainda não publicado;
- `publicado`: Drive contém a versão correspondente ao commit registrado;
- `edicao-mobile-pendente`: houve alteração direta no Google Slides e ela precisa voltar para a fonte oficial;
- `arquivado`: aula encerrada; manter o histórico, sem editar diretamente o artefato.
- `verificacao-pendente`: vínculo inicial registrado; ainda é preciso confirmar qual commit gerou o arquivo já existente no Drive.

## Checklist de reconciliação

- Compare a versão do Drive com o PPTX local e com o roteiro.
- Atualize texto, estrutura e assets no código, não apenas no artefato remoto.
- Gere novamente o PPTX e faça QA visual.
- Publique a nova versão no Drive e atualize o manifesto.
