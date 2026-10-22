# Controle de PO's

Painel web que lê as planilhas exportadas do SAP e guarda tudo no PostgreSQL
(`db_po.controle_po`).

## Como usar

1. Inicie o servidor manualmente:

   ```
   python server.py
   ```

2. Abra no navegador: http://localhost:8000

Existem dois botões de importação, com funções diferentes:

- **"Importar dados do SAP"** -> importação **automática**: pega o `.xlsx` mais
  recente que estiver na subpasta `Files/` dentro da pasta do projeto e grava
  tudo no banco.

- **"Exportar CSV"** -> abre o **seletor de arquivos** do Windows para você
  escolher um arquivo **.csv** ou **.xlsx**; o arquivo escolhido é lido, salvo
  na subpasta `Files/` dentro da pasta do projeto, as colunas são mapeadas e
  tudo é gravado no banco.

> Para usar outra porta: `set PORT=9000` (Windows) antes de rodar.

## Mapeamento de colunas (arquivo -> banco)

| Arquivo (cabeçalho)         | Banco (`db_po.controle_po`)   | Obs.                       |
|-----------------------------|-------------------------------|----------------------------|
| Documento de compras        | documento_de_compras          |                            |
| Item                        | item                          |                            |
| Contrato básico             | contrato_basico               |                            |
| Data do documento           | data_documento                |                            |
| Preço líquido               | preco_liquido                 |                            |
| a ser fornecido (valor)     | a_ser_fornecido_valor         | nome sem fechar parêntese  |
| Ainda a faturar (valor)     | ainda_a_faturar_valor         |                            |
| (manual)                    | nome_do_servico               | entra como "-"             |
| (manual)                    | competencia                   | entra como "-"             |

> **Completo:** ao importar, o banco grava **todas as 66 colunas** do arquivo
> (como `TEXT`), não só as 7 acima. A página web, porém, exibe apenas as 7
> principais (Documento, Item, Contrato, Data, Preço líquido, A ser fornecido,
> Ainda a faturar) + Nome do serviço/Competência (fixos como `-`).

As colunas são mapeadas pela ordem do export do SAP. O CSV é lido em UTF-8 ou
Latin-1/CP1252 — então variações de encoding entre arquivos funcionam igual.

> **Histórico de vínculo PO ↔ arquivo (base):** o dropdown de arquivos foi
> **removido do layout** para reduzir a poluição visual. A aplicação agora
> exibe sempre o conjunto completo de POs — na tabela, nos KPIs, na aba
> **Projetos** e no **Dashboard** — sem filtro por base importada.
>
> O histórico continua sendo gravado no banco pela tabela auxiliar
> `db_po.controle_po_arquivo` (chave `documento_de_compras + item +
> contrato_basico + arquivo_origem`), disponível para auditoria. Uma PO presente
> em dois arquivos aparece nos dois. O vínculo é gravado com `ON CONFLICT DO
> NOTHING`, ou seja, **nada é sobrescrito ou apagado** — inclusive as edições
> manuais de `nome_do_servico` e `competencia`, que continuam na tabela
> `controle_po`.
>
> A coluna `arquivo_origem` da tabela `controle_po` continua sendo atualizada
> a cada importação (guarda o **último** import).
>
> **Estrutura da tabela de vínculo** (todas as colunas do tipo `TEXT`):
>
> ```sql
> CREATE TABLE db_po.controle_po_arquivo (
>     documento_de_compras TEXT,
>     item                 TEXT,
>     contrato_basico      TEXT,
>     arquivo_origem       TEXT,
>     importado_em         TEXT DEFAULT to_char(now(), 'YYYY-MM-DD HH24:MI:SS'),
>     CONSTRAINT pk_controle_po_arquivo
>         UNIQUE (documento_de_compras, item, contrato_basico, arquivo_origem)
> );
> ```
>
> O `importado_em` é `TEXT` no formato `YYYY-MM-DD HH24:MI:SS`, que ordena
> corretamente como texto (mais recente primeiro) e é gravado automaticamente
> em toda importação.
>
> A tabela é criada automaticamente ao iniciar o servidor e é retroalimentada
> com os registros que já existiam em `controle_po` (via `arquivo_origem`).
> Se você criar a tabela manualmente, use o `CREATE TABLE` acima — a aplicação
> detecta o formato e converte automaticamente caso a coluna já exista como
> `TIMESTAMP`.

> **Filtro por mês de 2026:** na aba **POs**, o dropdown de meses (Janeiro/2026 a
> Dezembro/2026) filtra as POs pela coluna `data_documento` (data do documento).
> Funciona com o formato ISO `AAAA-MM-DD` (como está gravado no banco) e também
> com `DD/MM/AAAA`. Pode ser combinado com a busca, o filtro de competência e o
> checkbox de saldo. Não existe a opção "Todos os meses": o filtro é sempre
> ativo e inicia em **Janeiro/2026**.

> **Edição manual:** na aba **POs**, dê um **duplo clique** em qualquer linha
> para abrir um modal e editar o **projeto** (contrato básico) e a **competência**
> (formato `MM/AAAA`) daquela PO. O valor é gravado no banco
> (`POST /api/update`) e sobrevive a F5/reinício do servidor.
>
> No modal, o campo **"Contrato do projeto"** lista apenas o **contrato básico**
> (ex.: `5700023196`) — identificador único do projeto. O nome do projeto fica
> visível na caixa *"Sugestão do projeto"* e no tooltip (hover) da opção. Ao
> salvar, a coluna `nome_do_servico` grava o par **`Nome do projeto (Contrato)`**
> (ex.: `Centro Tático (5700023196)`), que identifica o projeto mesmo quando
> vários projetos compartilham o mesmo nome. Registros gravados no formato
> antigo (só o nome) continuam funcionando.

## Endpoints do servidor

- `GET  /`           -> página (index.html)
- `GET  /api/pos`    -> devolve as POs do banco (JSON). Aceita o parâmetro
  opcional `?origem=NOME_DO_ARQUIVO` para devolver apenas as POs daquela base.
  Cada PO vem com o campo `arquivos` (lista de bases que contêm a PO).
- `GET  /api/arquivos` -> devolve a lista de arquivos (bases) importados com a
  contagem de POs de cada um (JSON `{ok, arquivos:[{arquivo, count}]}`),
  ordenada do mais recente para o mais antigo
- `GET  /api/info`    -> destino real da conexão (host, porta, banco, schema,
  tabela) + total de linhas — usado no rodapé “Destino:” e no botão
  “Verificar conexão”
- `POST /api/import` -> recebe o arquivo enviado (JSON `{filename, data_base64}`),
  salva-o na subpasta `Files/` dentro da pasta do projeto, mapeia e grava no
  banco (UPSERT por documento+item+contrato; não duplica ao reimportar)
- `POST /api/clear`  -> esvazia a tabela `controle_po` **e** a de vínculo
  `controle_po_arquivo` (apaga também as edições manuais e o histórico de
  arquivos)
- `POST /api/update` -> atualiza **nome do serviço** e **competência (range)** de
  uma PO (JSON `{doc, item, contrato, nome_servico, data_inicio, data_fim}`) —
  acionado pelo **duplo clique** na linha da PO. A coluna `competencia` grava
  o texto "início a fim" (ex.: `01/10/2026 a 31/10/2026`) ou `-` quando limpo.

> Os dados vão para o banco `db_gerencial` (padrão), schema `db_po`, tabela
> `controle_po` — sobrescreva com a variável de ambiente `DB_PGC_DATABASE` se
> necessário. O rodapé da página mostra o destino em uso.

## Requisitos

Já vêm instalados neste ambiente: `psycopg2`, `openpyxl` (via Python).
Nada precisa ser instalado.

