# AutoMDFText

Automacao corporativa para preenchimento de MDF-e via navegador, com selecao de perfis, validacao de campos e controles de seguranca. O projeto combina um motor de automacao (teclado/clipboard) com um editor de perfis em GUI, permitindo padronizar rotinas e reduzir erros operacionais.

## Visao geral

- Objetivo: acelerar o preenchimento de MDF-e com dados padronizados por perfil (rotas, unidades, contratantes).
- Abordagem: scripts de configuracao em texto + automacao com teclado e clipboard.
- Operacao: menu central via `run.bat` com instalacao de dependencias e execucao dos modulos.

## Principais componentes

- Motor de automacao: `modular_mdfe.py`
  - Executa o fluxo de preenchimento com base no perfil selecionado.
  - Possui failsafe (F8), pausa (F9), verificacoes de campo e logs por sessao.
- Editor de perfis: `script_editor.py`
  - Interface grafica para criar, editar e salvar perfis `.txt`.
  - Assistente para gerar perfil a partir do template.
- Perfis e templates: `scripts/*.txt` e `scripts/template_config.txt`
  - Estrutura padronizada de chaves por secao (ex: `[MDFE]`, `[MODAL_RODOVIARIO]`).
- Logs de execucao: `logs/automation_YYYYMMDD_HHMMSS.log`
  - Registro de passos e mensagens para auditoria e troubleshooting.

## Como executar

1. Abra o menu principal:
   - Execute `run.bat`.
2. Escolha uma opcao:
   - `1` Executar preenchimento do MDF-e
   - `2` Abrir o editor de templates de script
   - `3` Instalar/atualizar dependencias
3. Siga as instrucoes exibidas na tela e selecione o perfil desejado.

## Passo a passo (menu completo)

1. Instalar ou atualizar dependencias (opcao 3)
  - Execute `run.bat` e selecione `3` na primeira utilizacao ou quando houver atualizacao.
  - Aguarde a confirmacao de instalacao concluida.
2. Criar ou revisar perfis (opcao 2)
  - Selecione `2` para abrir o editor.
  - Use "Novo Perfil" ou "Novo do Template" para iniciar um arquivo.
  - Opcional: utilize o "Assistente" para preencher campos com validacoes basicas.
  - Clique em "Salvar" para manter o perfil atualizado.
3. Executar a automacao (opcao 1)
  - Selecione `1` no menu principal.
  - Escolha o perfil desejado quando solicitado.
  - Informe dados dinamicos (ex.: numero do DT) quando o prompt aparecer.
  - Acompanhe a execucao; use F9 para pausar e F8 para encerrar em emergencia.
4. Auditoria e revisao
  - Consulte os arquivos de log em `logs/` para confirmar o fluxo e diagnosticos.

## Capturas sugeridas

- Tela inicial do menu `run.bat`.
- Editor de perfis com um arquivo carregado.
- Assistente de criacao de script em uso.
- Execucao da automacao com o navegador aberto.
- Exemplo de log gerado em `logs/`.

## Requisitos

- Windows com Python disponivel no PATH (ou a TI pode instalar localmente).
- Dependencias Python instaladas pelo menu:
  - `pyautogui`, `pyperclip`, `Pillow`, `pywin32`, `pynput`

## Controles de seguranca (operacao)

- F8: encerra a automacao imediatamente (failsafe).
- F9: pausa a automacao no proximo ponto seguro; janela de retomar/cancelar.
- Bloqueio de duplicidade:
  - `run.bat` impede multiplas instancias e automacoes paralelas.

## Fluxo de automacao (detalhado)

1. Inicializacao
  - Ao executar `modular_mdfe.py`, o sistema cria um log de sessao com timestamp.
  - A automacao entra no modo de seguranca (failsafe e pausa).
2. Selecao de perfil
  - O operador escolhe um arquivo em `scripts/` com os dados padronizados.
  - O perfil e validado e carregado na memoria para uso no preenchimento.
3. Coleta de dados dinamicos
  - Campos variaveis (ex.: numero do DT) sao solicitados via prompt.
  - O tempo de automacao considera pausas enquanto o operador responde.
4. Preenchimento guiado
  - O robô navega pelos campos com teclas (tab/enter) e cola valores via clipboard.
  - Sempre que possivel, ocorre verificacao do valor colado para reduzir divergencias.
  - Pontos seguros permitem pausa (F9) sem interromper sequencias criticas.
5. Validacoes de tela
  - Antes de etapas sensiveis, ha validacoes basicas para evitar preenchimento fora de contexto.
6. Encerramento
  - Ao concluir, o log registra o status final e o tempo total.
  - Caso o operador cancele na pausa ou acione F8, a sessao e encerrada imediatamente.

## Estrutura dos perfis

Os perfis sao arquivos `.txt` organizados por secoes. Exemplo (valores ilustrativos):

```ini
[GENERAL]
dt_prompt_text = Digite o numero do DT:

[MDFE]
prestador_tipo = PRESTADOR
emitente_codigo = 0000-00
uf_carregamento = UF
uf_descarga = UF
municipio_carregamento = CIDADE
codigo_produto_descricao = PRODUTO
ncm_primary = 00000000
ncm_secondary = 00000000
ncm_tertiary = 00000000
cep_origem = 00000000
cep_destino = 00000000
unidade_medida = 1
carga_tipo = 00

[MODAL_RODOVIARIO]
rntrc = 00000000
contratante_nome = EMPRESA
contratante_cnpj = 00000000000000

[INFORMACOES_ADICIONAIS]
contribuinte_cnpj = 00000000000000
seguradora_nome = SEGURADORA
seguradora_cnpj = 00000000000000
numero_apolice = 0000000000
frete_valor = 0.00
frete_tipo = FRETE
frete_identificador = FRETE
numero_banco = 000
agencia = 0000/0
forma_pagamento = 0
numero_parcelas = 0
```

## Editor de perfis (GUI)

No editor, voce pode:

- Carregar e atualizar perfis existentes.
- Criar um perfil novo em branco.
- Criar um perfil novo a partir do template (limpando valores).
- Usar o assistente para preencher campos com validacoes basicas.

## Logs e auditoria

Cada sessao gera um arquivo em `logs/` com:

- timestamp da sessao
- perfil utilizado
- passos e eventos (pausa, cancelamento, validacoes)

Isso facilita auditoria e diagnostico de falhas de preenchimento.

## Boas praticas

- Mantenha os perfis atualizados e revisados por unidade.
- Teste um perfil novo em ambiente controlado antes do uso produtivo.
- Nao execute outras automacoes em paralelo durante o uso.
- Use F9 para pausar se precisar validar dados antes de continuar.

## Suporte

Para ajustes de template ou novas rotas, use o editor de perfis ou edite os arquivos em `scripts/` diretamente, sempre mantendo o padrao de secoes.
