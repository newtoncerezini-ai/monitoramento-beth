# Monitoramento de Planos de Acao

Sistema web em Python + SQLite para cadastrar planos de acao, subatividades, responsaveis e acompanhamentos.

## Rodar localmente

```bash
python -m pip install -r requirements.txt
python app.py
```

Depois acesse:

```text
http://127.0.0.1:5000
```

## Primeiro acesso

O sistema cria automaticamente um usuario administrador inicial:

```text
Usuario: admin
Senha: admin123
```

Depois do primeiro acesso, entre em `Administracao` para criar os usuarios reais do sistema e alterar a senha do administrador.

## Perfis de acesso

- `admin`: gerencia usuarios, importa planilha e tambem pode cadastrar/alterar dados.
- `gestor`: pode cadastrar e alterar acoes, prazos, responsaveis e acompanhamentos.
- `usuario`: pode consultar dashboards, acoes, responsaveis e historicos, mas nao pode incluir ou alterar dados.

## Cadastros

- `Nova acao`: cadastra planos macro, atividades e subatividades.
- `Responsaveis`: cadastra pessoas responsaveis pelas acoes.
- `Administracao`: cadastra usuarios que podem acessar o sistema.

Responsaveis pelas acoes e usuarios do sistema sao cadastros separados. Uma pessoa pode aparecer como responsavel sem necessariamente ter login no sistema.

Na interface, os cadastros operacionais ficam concentrados em `Administracao`.
O menu lateral mantem foco em consulta: `Dashboard`, `Monitoramento` e `Administracao` para perfis com permissao de escrita.

## Telas de monitoramento

- `Monitoramento`: lista geral de acoes com filtros.
- `Alertas e atrasos`: acoes vencidas, proximas do vencimento e sem responsavel.
- `Calendario`: prazos agrupados por mes.
- `Responsaveis`: carga, atrasos e conclusoes por pessoa.
- `Planos macro`: progresso agregado e subatividades criticas por plano.
- `Kanban`: acoes organizadas por status.

## Deploy de demonstracao no Vercel

O projeto inclui `vercel.json` e `api/index.py` para rodar como uma funcao Python no Vercel.

Passos:

```bash
npm i -g vercel
vercel login
vercel
```

Na primeira publicacao, confirme o diretorio do projeto e deixe o Vercel detectar a configuracao.

Para producao/demo, configure uma variavel de ambiente:

```text
SECRET_KEY=uma-chave-segura
```

Observacao importante: no Vercel, o SQLite roda como banco temporario em `/tmp`. O sistema copia `data/monitoramento.db` como base inicial, mas alteracoes feitas no ambiente publicado podem se perder quando a funcao serverless reiniciar. Para uso real com persistencia, migre o banco para Postgres, Supabase ou Neon.

## Deploy no Render

O projeto inclui `render.yaml` para publicar como Web Service Python.

Passos recomendados:

1. Suba este projeto para um repositorio no GitHub.
2. Acesse `https://render.com`.
3. Clique em `New +`.
4. Escolha `Blueprint`.
5. Conecte o repositorio do GitHub.
6. Selecione o arquivo `render.yaml`.
7. Confirme a criacao do servico.

O Render usara:

```text
Build Command: pip install -r requirements.txt
Start Command: gunicorn app:app
```

Observacao: com SQLite simples, o banco pode nao ser persistente entre deploys/restarts. Para demonstracao, funciona. Para uso real, use Postgres ou configure um Persistent Disk no Render.

## Importar a planilha inicial

Com o servidor parado ou rodando, execute:

```bash
python app.py --import "C:\Users\newton.cerezini\Downloads\COMFEM - BD PE26271.xlsx"
```

O banco SQLite fica em `data/monitoramento.db`.

## Modelo

- `actions`: guarda planos, atividades e subatividades ate o terceiro nivel.
- `people`: cadastro unico de responsaveis.
- `action_people`: vinculo muitos-para-muitos entre acoes e responsaveis.
- `updates`: historico de monitoramento por acao.

Os codigos vindos da planilha sao normalizados para hierarquia. Por exemplo, `2.` vira `2`.
