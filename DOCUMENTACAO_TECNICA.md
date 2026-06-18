# AutoMDFText — Documentação Técnica do Projeto

> **Versão do documento:** 1.1  
> **Data:** Junho 2026  
> **Escopo:** Estado atual do projeto (pós-refatoração modular) + schema para registrar cada melhoria futura  

---

## Sumário

1. [Visão Geral do Projeto](#1-visão-geral-do-projeto)
2. [Restrições do Ambiente Corporativo](#2-restrições-do-ambiente-corporativo)
3. [Arquitetura do Sistema](#3-arquitetura-do-sistema)
4. [Documentação dos Arquivos e Funções](#4-documentação-dos-arquivos-e-funções)
5. [Perfis de Configuração](#5-perfis-de-configuração)
6. [Fluxo de Execução Detalhado](#6-fluxo-de-execução-detalhado)
7. [Log de Decisões de Engenharia](#7-log-de-decisões-de-engenharia)
8. [Template de Change Record](#8-template-de-change-record)

---

## 1. Visão Geral do Projeto

**AutoMDFText** é uma ferramenta de automação desktop desenvolvida para o ambiente
corporativo que elimina o preenchimento manual de formulários de **Manifesto de Documentos
Fiscais Eletrônicos (MDF-e)** no sistema web **InvoiSys**.

### Problema resolvido

O operador precisava preencher manualmente dezenas de campos em sequência exata no
InvoiSys a cada emissão de MDF-e. O processo era sujeito a erros de digitação,
omissão de campos e inconsistência entre rotas. A ferramenta elimina esse esforço
repetitivo e garante que os valores enviados ao sistema sejam sempre os mesmos para
cada rota configurada.

### Como funciona, em resumo

1. O operador abre o menu via `run.bat` e escolhe um **perfil de rota** (ex.: ITU × DHL).
2. Informa os dados dinâmicos da operação (número do DT, CT-e, NFs).
3. A automação assume o controle do teclado e clipboard, navega pelos campos do InvoiSys
   e preenche cada valor na ordem exata esperada pelo formulário.
4. **(WIP — próxima tarefa)** O preenchimento do CIOT será integrado ao pipeline como etapa obrigatória após a conclusão do MDF-e.

### Tecnologias utilizadas

| Tecnologia | Versão mínima | Papel |
|---|---|---|
| Python | 3.11 | Runtime da automação |
| pyautogui | última | Envio de teclas e clipboard |
| pyperclip | última | Leitura/escrita de clipboard |
| pynput | última | Listener global de teclado (F8/F9) |
| tkinter | stdlib | Janelas de diálogo |
| pywin32 | última | APIs Win32 (foco de janela, mutex) |
| Pillow | última | Dependência indireta do pyautogui |

---

## 2. Restrições do Ambiente Corporativo

> [!IMPORTANT]
> Estas restrições determinam **todas** as decisões de arquitetura de execução.
> Qualquer mudança futura deve ser avaliada contra esta lista.

### 2.1 Python não está no PATH do sistema

Nos computadores corporativos gerenciados pela TI, o Python **nunca** está registrado
na variável de ambiente `PATH`. Isso significa que comandos como `python script.py`
falham no terminal padrão.

**Solução adotada:** o `run.bat` detecta o Python disponível (testando `py`, `py.exe`
e `python` com `where`) e cria um ambiente virtual local (`.venv`) no diretório do projeto.
Toda invocação usa o caminho absoluto `%VENV_DIR%\Scripts\python.exe`.

Trecho do `run.bat` que implementa isso:
```bat
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
if exist "%VENV_PY%" (
    set "PYTHON=%VENV_PY%"
) else (
    "%BASE_PY%" -m venv "%VENV_DIR%"
    set "PYTHON=%VENV_PY%"
)
```

### 2.2 Sem permissão para instalar globalmente

Dependências são instaladas **dentro do `.venv` local** via opção `3` do menu do
`run.bat`. O operador não precisa de privilégios de administrador para isso.

### 2.3 Ambiente Windows-only

Toda a detecção de janela (Win32 API), controle de mutex (prevenção de instância
duplicada) e interação com o console (mostrar/ocultar janela) usa APIs exclusivas
do Windows. O código contém guards `if os.name == "nt"` em todos esses pontos para
não quebrar em outros sistemas, mas a operação real é sempre no Windows.

### 2.4 Navegador sem controle de driver

O InvoiSys roda no navegador (Edge/Chrome) e a automação **não usa Selenium nem
WebDriver**. Todo preenchimento é feito via simulação de teclado e clipboard porque:

- O acesso à instância do navegador via driver não é viável no ambiente corporativo.
- A abordagem por teclado/clipboard é mais resiliente a mudanças de versão do navegador.

### 2.5 Único ponto de entrada para o operador

O operador não deve executar arquivos Python diretamente. O único ponto de entrada
suportado e documentado é o `run.bat`, que gerencia Python, dependências e guards
de duplicidade automaticamente.

---

## 3. Arquitetura do Sistema

### 3.1 Diagrama de componentes (estado atual)

```markdown
        ─────────────────────────────────────────────────────────────────
      │  Ponto de entrada do operador                                    │
      │  run.bat                                                         │
      │  • Detecta / cria .venv                                          │
      │  • Guard de instância duplicada                                  │
      │  • Menu: [1] MDF-e  [2] Editor  [3] Deps  [4] Sair               │
        ─────────────────────────────────────────────────────────────────
            │                      │
            ▼                      ▼
   modular_mdfe.py          script_editor.py
   (wrapper legado)         (editor de perfis GUI)
            │
            ▼
   mdfe.runner.main()
   (motor de automação)
            │
            ├── mdfe/logger.py       ← logging e console
            ├── mdfe/timing.py       ← temporização e métricas
            ├── mdfe/failsafe.py     ← listener F8/F9
            ├── mdfe/pause.py        ← diálogo de pausa
            ├── mdfe/console.py      ← controle de janela
            ├── mdfe/instance.py     ← mutex single-instance
            ├── mdfe/dialogs.py      ← prompts e alertas
            ├── mdfe/keyboard.py     ← teclado e clipboard
            ├── mdfe/browser.py      ← detecção/foco de janela
            ├── mdfe/profile.py      ← parsing de perfis
            ├── mdfe/steps/          ← etapas de preenchimento
            │   ├── navigate.py
            │   ├── fill_mdfe.py
            │   ├── fill_modal_rodo.py
            │   ├── fill_additional_info.py
            │   └── averbacao.py
            ├── constants.py         ← constantes centralizadas
            ├── scripts/*.txt        ← perfis de rota
            ├── logs/                ← logs de sessão
            └── test_suite.py        ← testes unitários
```

### 3.2 Fluxo de chamada entre módulos

```
run.bat
  └─► modular_mdfe.py (wrapper, 6 linhas)
        └─► mdfe.runner.main()
              ├── lê constants.py
              ├── mdfe.profile.ConfigProfile → scripts/<perfil>.txt
              ├── mdfe.instance.ensure_single_instance()
              ├── mdfe.dialogs.prompt_dt_blocking()   ← DT
              ├── mdfe.dialogs.prompt_batch_info()    ← CT-e, NF, NCM
              ├── mdfe.browser.verify_cte_on_page()
              ├── mdfe.steps.navigate_to_mdfe()
              ├── mdfe.steps.fill_mdfe(profile, ncm)
              ├── mdfe.steps.fill_modal_rodo(profile)
              ├── mdfe.steps.fill_additional_info(profile)
              └── mdfe.steps.perform_averbacao(cte, dt, nf)

```

---

## 4. Documentação dos Arquivos e Funções

> **Nota:** O projeto passou por refatoração completa (Partes 1-5) que converteu o monólito
> `modular_mdfe.py` (originalmente 2228 linhas) em um pacote modular `mdfe/`. As funções
> estão agora distribuídas por módulos especializados conforme documentado abaixo.

---

### `run.bat` — Ponto de entrada e menu principal

**Responsabilidade:** único arquivo que o operador executa. Gerencia o ambiente Python
e serve como menu central da ferramenta.

| Seção | O que faz |
|---|---|
| Detecção de Python | Testa `py`, `py.exe`, `python` com `where`; cria `.venv` se necessário |
| Guard de duplicidade | PowerShell verifica se já existe janela do toolkit ou automação rodando |
| Menu principal | `choice /C 1234` para seleção sem prompt interativo do Python |
| `:modular` | Chama `python modular_mdfe.py`; detecta exit code 99 para voltar ao menu |
| `:editor` | Chama `python script_editor.py` |
| `:install` | Instala pacotes no `.venv` via `pip install` |

**Decisão de design:** o exit code `99` é uma convenção interna. Quando o usuário
escolhe "0 - Voltar" no seletor de perfil, o Python termina com `raise SystemExit(99)`,
e o bat detecta esse código com `IF ERRORLEVEL 99 goto prompt` para retornar ao menu
sem encerrar o terminal.

---

### `constants.py` — Constantes centralizadas

**Responsabilidade:** único ponto de definição de todos os delays, timeouts e
configurações numéricas usados na automação. Centralizar aqui evita que um ajuste de
timing precise ser feito em múltiplos arquivos.

#### Constantes documentadas

| Constante | Valor | Uso |
|---|---|---|
| `TAB_DELAY` | `0.10 s` | Delay padrão entre pressionamentos de Tab |
| `TAB_DELAY_LONG` | `0.25 s` | Tab em campos mais lentos ou sensíveis |
| `CTRL_F_DELAY` | `0.20 s` | Após abrir Ctrl+F (busca no navegador) |
| `DROPDOWN_SETTLE_DELAY` | `0.40 s` | Aguarda dropdown renderizar após seleção |
| `SLEEP_SHORT` | `0.15 s` | Pausa leve entre ações não-críticas |
| `SLEEP_MEDIUM` | `0.25 s` | Pausa padrão de navegação entre campos |
| `SLEEP_LONG` | `0.40 s` | Aguarda carregamento de seção ou dropdown |
| `SLEEP_LONGER` | `0.70 s` | Transições pesadas (mudança de aba/seção) |
| `SLEEP_ONE` | `1.00 s` | Operações críticas que dependem de resposta da página |
| `SLEEP_ONE_HALF` | `1.50 s` | Menu de seleção, operações longas |
| `FORM_WAIT_TIMEOUT` | `15.0 s` | Timeout máximo para o formulário abrir |
| `FORM_WAIT_INTERVAL` | `1.0 s` | Intervalo entre tentativas de detecção |
| `FORM_COPY_ATTEMPTS` | `2` | Tentativas de Ctrl+A / Ctrl+C por ciclo |
| `VERIFY_FIELD_TIMEOUT` | `6.0 s` | Timeout para verificar CT-e na página |
| `EDGE_SEARCHBAR_HEIGHT` | `70 px` | Altura da barra de endereço do Edge |
| `EDGE_CLICK_OFFSET` | `200 px` | Distância abaixo da barra para clicar no conteúdo |
| `MIN_PASTE_LENGTH` | `4 chars` | Textos ≥ 4 chars usam clipboard em vez de digitação |
| `WRITE_INTERVAL` | `0.10 s` | Intervalo entre teclas ao digitar caractere a caractere |
| `LOG_TIMESTAMP_FORMAT` | `%Y%m%d_%H%M%S` | Formato do timestamp no nome do arquivo de log |
| `LOG_MESSAGE_FORMAT` | `%Y-%m-%dT%H:%M:%S` | Formato do timestamp nas mensagens de log |
| `MUTEX_NAME` | `Global\AutoMDFText_Mutex` | Nome do mutex Windows para single-instance |

---

### `modular_mdfe.py` — Wrapper legado (ponto de entrada)

**Responsabilidade:** arquivo de transição que mantém compatibilidade com o `run.bat`
e com scripts existentes que chamam `python modular_mdfe.py`. Atualmente é um wrapper
de 6 linhas que importa e delega para `mdfe.runner.main()`.

```python
"""Ponto de entrada legado — delega para o pacote mdfe."""
from mdfe.runner import main
if __name__ == "__main__":
    main()
```

> Nenhuma lógica de automação reside mais neste arquivo.

---

### `mdfe/logger.py` — Logging e console

**Responsabilidade:** registrar mensagens em arquivo de log de sessão e formatar
saída colorida no console.

| Função | Descrição |
|---|---|
| `log(msg: str)` | Registra mensagem em `logs/automation_<ts>.log`. Não imprime no console. Silencia exceções de I/O. |
| `ui_print(msg: str, style: str = "info")` | Imprime no console com formatação ANSI. Estilos: `header` (azul `═`), `step` (azul `▸`), `success` (verde `✓`), `error` (vermelho `✗`), `warning` (amarelo `⚠`), `info` (padrão, indentado). |
| `start_automation_session(selected: str, profile_path: Path) → float` | Inicializa sessão de log com timestamp, reinicia contadores de tempo em `mdfe.timing`, retorna instante monotônico do início real. |

---

### `mdfe/timing.py` — Temporização e métricas

**Responsabilidade:** controlar os contadores de tempo de automação, descontando
pausas de interação do operador (prompts, diálogos).

| Função | Descrição |
|---|---|
| `pause_automation_timer()` | Registra instante de início de uma pausa de usuário. |
| `resume_automation_timer()` | Calcula tempo decorrido desde a última pausa e acumula em `_automation_time_paused`. |
| `format_duration(seconds: float) → str` | Formata duração: `< 60s → "42s"`, `≥ 60s → "2m 15s"`. |

**Globals:** `_automation_start_time` e `_automation_time_paused` (lidos por `runner.py`).

---

### `mdfe/failsafe.py` — Failsafe de teclado (F8/F9)

**Responsabilidade:** listener global de teclado via `pynput` para F8 (encerrar
imediatamente) e F9 (sinalizar pausa). Mantém estado de pausa compartilhado.

| Função | Descrição |
|---|---|
| `start_failsafe_f8()` | Inicia thread listener. F8 → `os._exit(1)` + alerta Win32 MessageBox. F9 → `request_pause()`. Usa `win32_event_filter` para ignorar teclas injetadas. |
| `stop_failsafe_f8()` | Para o listener. Chamado no `finally` do `main()`. |
| `request_pause()` | Sinaliza `_pause_requested = True` de forma thread-safe. |

**Globals:** `_pause_requested`, `_pause_active`, `_pause_lock` (lidos por `mdfe/pause.py`).

---

### `mdfe/pause.py` — Diálogo de pausa

**Responsabilidade:** exibir janela de diálogo topmost quando o operador solicita
pausa (F9), com watchdog para manter-se visível.

| Função | Descrição |
|---|---|
| `show_pause_dialog() → str` | Janela tkinter topmost com botões "Retomar"/"Cancelar". Watchdog reaplica TOPMOST a cada 600 ms. Retorna `"resume"` ou `"cancel"`. |
| `check_pause()` | Verifica `_pause_requested`; se True, pausa timer, exibe diálogo, aguarda decisão. Se cancelar, `SystemExit(1)`. |
| `pause_point()` | Ponto seguro chamado entre etapas. Delega para `check_pause()`. |
| `_verify_last_write_before_pause()` | Revalida último valor enviado via `smart_write` usando Ctrl+A/Ctrl+C. Se divergir, reaplica via `paste_text`. |

---

### `mdfe/console.py` — Gerenciamento de janela do console

**Responsabilidade:** ocultar/restaurar a janela do terminal e emitir beep.

| Função | Descrição |
|---|---|
| `hide_console_window()` | Oculta console via `SW_HIDE`. |
| `restore_console_popup()` | Restaura console como popup (ShowWindow → SetForegroundWindow → HWND_TOPMOST → HWND_NOTOPMOST). |
| `play_low_beep()` | Beep 400 Hz / 180 ms via `winsound.Beep`. |

---

### `mdfe/instance.py` — Single-instance guard

**Responsabilidade:** impedir execução duplicada via mutex Win32.

| Função | Descrição |
|---|---|
| `ensure_single_instance()` | Cria mutex `Global\AutoMDFText_Mutex` via `CreateMutexW`. Se já existir, exibe alerta e `SystemExit(0)`. Handle mantido para evitar GC prematuro. |

---

### `mdfe/dialogs.py` — Diálogos e prompts

**Responsabilidade:** wrappers de diálogos pyautogui/tkinter com controle de timer.

| Função | Descrição |
|---|---|
| `focused_alert(text, title, button)` | Wrapper de `pyautogui.alert` com pausa/retomada de timer. |
| `focused_confirm(text, title, buttons)` | Wrapper de `pyautogui.confirm`. Retorna **string do botão** (não índice). |
| `focused_prompt(text, title, default)` | Wrapper de `pyautogui.prompt` com controle de timer. |
| `prompt_dt_blocking(text, title)` | Diálogo tkinter personalizado para DT. Validação: não permite OK com campo vazio. |
| `prompt_batch_info(ncm_options)` | Diálogo tkinter unificado para CT-e, NF1, NF2, NCM (radiobutton + "Outro"). Retorna `dict` ou `None`. |

---

### `mdfe/keyboard.py` — Teclado e clipboard

**Responsabilidade:** simulação inteligente de digitação e manipulação de clipboard.

| Função | Descrição |
|---|---|
| `press_tab(count=1, delay=TAB_DELAY)` | Pressiona Tab N vezes com delay entre cada. |
| `skip_tabs(count, log_msg="")` | Wrapper de `press_tab` com log opcional. |
| `paste_text(text, verify, retries, delay, restore_clipboard)` | Cola via clipboard com verificação opcional (Ctrl+A/Ctrl+C). CPF/CNPJ têm verify desativado (máscara). |
| `smart_write(value, interval, min_paste_len, verify)` | Escolhe entre clipboard (textos ≥ 4 chars ou com chars especiais) e digitação (`pyautogui.write`). |
| `ensure_caps_off()` | Desativa Caps Lock via `GetKeyState` se estiver ligado. |
| `upload_latest_xml()` | Localiza arquivo mais recente em Downloads, escreve caminho no campo de upload, confirma com Enter. |

---

### `mdfe/browser.py` — Detecção e foco de navegador

**Responsabilidade:** identificar janelas do navegador, focá-las, aguardar formulários
e verificar conteúdo de página via clipboard.

| Função | Descrição |
|---|---|
| `_get_foreground_title()` | Título da janela em foco via `GetForegroundWindow` + `GetWindowTextW`. |
| `_get_foreground_class()` | Classe da janela em foco via `GetClassNameW`. |
| `_get_window_process_name(hwnd)` | Nome do executável do processo via `GetModuleBaseNameW`. |
| `_is_cloaked_window(hwnd)` | Verifica se janela está oculta pela DWM. |
| `_is_top_level_app_window(hwnd)` | Exclui owned windows e toolwindows. |
| `_is_standard_window(hwnd)` | Verifica estilo `WS_OVERLAPPEDWINDOW`. |
| `_is_browser_window(title, cls, process_name)` | Combina título/classe/processo para identificar navegador. |
| `_find_browser_windows()` | Enumera janelas visíveis e filtra usando os predicados acima. |
| `focus_browser_if_needed()` | Foca navegador via Win+1 (primeiro app fixado na barra). Só age se necessário. |
| `_click_below_edge_searchbar(offset)` | Clica no conteúdo da página abaixo da barra de endereço do Edge. |
| `_focus_page_for_copy()` | ESC + `_click_below_edge_searchbar` para garantir foco no corpo da página. |
| `wait_for_form(target_text, tempo_maximo, intervalo, copy_attempts)` | Aguarda formulário abrir verificando conteúdo via clipboard. Timeout → `SystemExit(1)`. |
| `verify_cte_on_page(numero_cte, tempo_maximo, intervalo)` | Verifica CT-e na página com 3 estratégias de match. |

---

### `mdfe/profile.py` — Perfis de configuração

**Responsabilidade:** ler arquivos `.txt` de perfil, representá-los em memória com
hot-reload e oferecer seleção interativa.

| Função / Classe | Descrição |
|---|---|
| `parse_profile(path) → dict` | Lê perfil INI-like em dicionário de seções. Ignora `#` e linhas vazias. |
| `class ConfigProfile` | Perfil em memória com detecção automática de mudança (mtime). Métodos: `reload()`, `ensure_current()`, `get_value(section, key, default)`. |
| `list_profiles() → list[str]` | Lista arquivos `.txt` em `scripts/`, ordenados. |
| `choose_profile(interactive_list) → str` | Menu interativo no terminal (índice ou nome, case-insensitive). `0` → `SystemExit(99)` (volta ao menu). Oculta console após seleção. |

---

### `mdfe/steps/` — Etapas de preenchimento

**Responsabilidade:** cada etapa da automação em um módulo dedicado.

---

##### `mdfe/steps/navigate.py` — `navigate_to_mdfe()`

Navega ao formulário MDF-e:
1. Ctrl+3 → aba 3 (InvoiSys).
2. Ctrl+F → "EMITIR NOTA" → Esc → Enter.
3. Ctrl+F → "MDF-E" → Esc → Enter.

---

##### `mdfe/steps/fill_mdfe.py` — `fill_mdfe(profile, codigo_ncm)`

Preenche formulário principal do MDF-e:

| Campo | Chave do perfil | Seção |
|---|---|---|
| Prestador de Serviço | `prestador_tipo` | `[MDFE]` |
| Emitente | `emitente_codigo` | `[MDFE]` |
| UF Carregamento | `uf_carregamento` | `[MDFE]` |
| UF Descarga | `uf_descarga` | `[MDFE]` |
| Município de Carregamento | `municipio_carregamento` | `[MDFE]` |
| Upload do XML CT-e | *(arquivo mais recente de Downloads)* | — |
| Unidade de Medida | `unidade_medida` | `[MDFE]` |
| Tipo de Carga | `carga_tipo` | `[MDFE]` |
| Descrição do Produto | `codigo_produto_descricao` | `[MDFE]` |
| Código NCM | *(selecionado pelo operador no prompt)* | — |
| CEP Origem | `cep_origem` | `[MDFE]` |
| CEP Destino | `cep_destino` | `[MDFE]` |

---

##### `mdfe/steps/fill_modal_rodo.py` — `fill_modal_rodo(profile)`

Preenche seção "Modal Rodoviário":

| Campo | Chave do perfil | Seção |
|---|---|---|
| RNTRC | `rntrc` | `[MODAL_RODOVIARIO]` |
| Nome do Contratante | `contratante_nome` | `[MODAL_RODOVIARIO]` |
| CNPJ do Contratante | `contratante_cnpj` | `[MODAL_RODOVIARIO]` |

---

##### `mdfe/steps/fill_additional_info.py` — `fill_additional_info(profile)`

Preenche "Informações Adicionais / Opcionais" com dados de seguro, frete e parcelas:

| Campo | Chave do perfil | Seção |
|---|---|---|
| CNPJ Contribuinte | `contribuinte_cnpj` | `[INFORMACOES_ADICIONAIS]` |
| CNPJ Contratante (2ª ocorrência) | `contratante_cnpj` | `[MODAL_RODOVIARIO]` |
| Nome Seguradora | `seguradora_nome` | `[INFORMACOES_ADICIONAIS]` |
| CNPJ Seguradora | `seguradora_cnpj` | `[INFORMACOES_ADICIONAIS]` |
| Número Apólice | `numero_apolice` | `[INFORMACOES_ADICIONAIS]` |
| Nome Contratante (3ª ocorrência) | `contratante_nome` | `[MODAL_RODOVIARIO]` |
| CNPJ Contratante (3ª ocorrência) | `contratante_cnpj` | `[MODAL_RODOVIARIO]` |
| Valor Frete | `frete_valor` | `[INFORMACOES_ADICIONAIS]` |
| Forma de Pagamento | `forma_pagamento` | `[INFORMACOES_ADICIONAIS]` |
| Número Banco | `numero_banco` | `[INFORMACOES_ADICIONAIS]` |
| Agência | `agencia` | `[INFORMACOES_ADICIONAIS]` |
| Tipo Frete (dropdown) | `frete_tipo` | `[INFORMACOES_ADICIONAIS]` |
| Valor Frete (2ª ocorrência) | `frete_valor` | `[INFORMACOES_ADICIONAIS]` |
| Número de Parcelas | `numero_parcelas` | `[INFORMACOES_ADICIONAIS]` |
| Valor Frete (3ª ocorrência) | `frete_valor` | `[INFORMACOES_ADICIONAIS]` |

---

##### `mdfe/steps/averbacao.py` — `perform_averbacao(numero_cte, numero_dt, nf_concat)`

Executa averbação e preenche campo de contribuinte:
1. Ctrl+4 → aba averbação.
2. Busca "OK" → "XML" → "ENVIAR" via Ctrl+F + Enter.
3. Upload XML via `upload_latest_xml()`.
4. Captura número de averbação via Ctrl+A/Ctrl+C + regex.
5. Ctrl+3 → aba MDF-e.
6. Busca "DETALHES" → preenche número de averbação.
7. Busca "CONTRIBUINTE" → preenche `"DT: <dt> CTE: <cte> NF: <nf_concat>"`.

---

### `mdfe/runner.py` — Motor de execução principal

**Responsabilidade:** orquestra o fluxo completo. Contém `main()` com toda a lógica
de coordenação: argumentos (`--profile`), inicialização, prompts,
etapas de preenchimento e resumo final.

**Fluxo em `main()`:**
1. `ensure_single_instance()`
2. `choose_profile()` / `--profile`
3. `start_automation_session()`
4. `start_failsafe_f8()`
5. `focus_browser_if_needed()`
6. ESC x2 + Ctrl+3 + F5 + Ctrl+1 (preparação do navegador)
7. `prompt_dt_blocking()` + preenchimento do DT (Ctrl+F "DO DT" e "mero do DT")
8. `focused_alert()` — aviso para baixar XML do CT-e
9. `prompt_batch_info()` — CT-e, NF1, NF2, NCM
10. `verify_cte_on_page()`
11. `navigate_to_mdfe()` + `wait_for_form("Emissor MDF-e")`
12. `fill_mdfe()` → `fill_modal_rodo()` → `fill_additional_info()` → `perform_averbacao()`
13. Resumo final com tempos (automação × real)
14. **(WIP — próxima tarefa)** Integração do CIOT como etapa obrigatória

---

### `mdfe/__init__.py` — Exportação do pacote

```python
from mdfe.runner import main
__all__ = ["main"]
```

---

### `test_suite.py` — Testes unitários

**Responsabilidade:** suíte de testes unitários usando `unittest` com mocks de
`pyautogui`, `pyperclip`, `pynput` e `ctypes.windll`. Testa todos os 11 módulos do
pacote `mdfe/`. Inclui runner colorido com diagnóstico detalhado de falhas e sugestões.

| Aspecto | Descrição |
|---|---|
| Framework | `unittest` padrão |
| Mocks | `pyautogui`, `pyperclip`, `pynput`, `ctypes.windll` |
| Cobertura | logger, timing, failsafe, console, instance, pause, profile, dialogs, keyboard, browser, steps, runner |
| Execução | `python test_suite.py` |

---

### `script_editor.py` — Editor de perfis (GUI)

**Responsabilidade:** interface gráfica (tkinter) para criar e editar perfis `.txt`.

| Funcionalidade | Descrição |
|---|---|
| Carregar perfis | Lista combo Box com perfis existentes |
| Novo Perfil | Cria arquivo em branco |
| Novo do Template | Copia `template_config.txt` como base |
| Assistente | Formulário guiado com validação (dígitos para CEP/CNPJ/NCM, letras para UF) |
| Salvar / Salvar Como | Grava perfil editado |

---

### CIOT — Extensão obrigatória (WIP — próxima tarefa)

O preenchimento do campo CIOT (Conhecimento de Transporte Intermodal Operacional) será
integrado ao pipeline como **etapa obrigatória** após a conclusão do MDF-e. Esta é a
**próxima tarefa de desenvolvimento** do projeto. Os detalhes de implementação serão
definidos durante sua execução.

---

### `script_editor.py` — Editor de perfis

**Responsabilidade:** interface gráfica para criar e editar arquivos `.txt` de perfil
sem que o operador precise usar um editor de texto genérico.

| Funcionalidade | Descrição |
|---|---|
| Carregar/Atualizar | Recarrega a lista de perfis disponíveis |
| Novo Perfil | Cria arquivo em branco para edição livre |
| Novo do Template | Copia `template_config.txt` como base, limpando valores |
| Assistente | Formulário guiado com campos nomeados e validação básica |
| Salvar / Salvar Como | Grava o perfil editado |

---

## 5. Perfis de Configuração

Os perfis são arquivos `.txt` em `scripts/` com estrutura INI. Cada perfil representa
uma **rota de operação** fixa (ex.: ITU → DHL, Sorocaba → DHL).

### Estrutura do arquivo

```ini
[GENERAL]
dt_prompt_text = Digite o número do DT:

[MDFE]
prestador_tipo         = PRESTADOR
emitente_codigo        = 031560
uf_carregamento        = SP
uf_descarga            = SP
municipio_carregamento = ITU
codigo_produto_descricao = PA/PALLET
ncm_primary            = 19041000
ncm_secondary          = 19059090
ncm_tertiary           = 20052000
cep_origem             = 13300340
cep_destino            = 13315000
unidade_medida         = 1
carga_tipo             = 05

[MODAL_RODOVIARIO]
rntrc             = 45501846
contratante_nome  = PEPSICO
contratante_cnpj  = 02957518000224

[INFORMACOES_ADICIONAIS]
contribuinte_cnpj  = 04898488000177
seguradora_nome    = SEGUROS SURA SA
seguradora_cnpj    = 33065699000127
numero_apolice     = 5400035882
frete_valor        = 1314.27
frete_tipo         = FRETE
frete_identificador = FRETE
numero_banco       = 237
agencia            = 2372/8
forma_pagamento    = 1
numero_parcelas    = 1
```

### Perfis ativos

| Arquivo | Rota |
|---|---|
| `1. itu_x_dhl.txt` | ITU → DHL |
| `2. sorocaba_x_dhl.txt` | Sorocaba → DHL |
| `3. itu_x_sorocaba.txt` | ITU → Sorocaba |
| `4. itu_x_mega_abc.txt` | ITU → Mega ABC |
| `5. itu_x_santos.txt` | ITU → Santos |
| `6. itu_x_curitiba.txt` | ITU → Curitiba |
| `7. itu_x_guarulhos.txt` | ITU → Guarulhos |
| `8. itu_x_taubate.txt` | ITU → Taubaté |

---

## 6. Fluxo de Execução Detalhado

```
Operador → run.bat
    │
    ├─[3] Instalar deps: pip install no .venv local
    ├─[2] Editor: abre script_editor.py
    ├─[4] Sair
    └─[1] Automação:
          │
          ├── 1.  Sleep inicial (SLEEP_LONGER) — aguarda estabilização
          ├── 2.  ensure_single_instance() — bloqueia duplicatas via mutex
          ├── 3.  choose_profile() — operador seleciona a rota (ou --profile via arg)
          ├── 4.  start_automation_session() — inicia log e timers
          ├── 5.  start_failsafe_f8() — ativa F8/F9
          ├── 6.  focus_browser_if_needed() — garante navegador em foco
          ├── 7.  ESC x2 — limpa pop-ups residuais
          ├── 8.  Ctrl+3 + F5 — recarrega aba do InvoiSys
          ├── 9.  Ctrl+1 + ESC — volta à aba 1
          ├── 10. prompt_dt_blocking() — operador informa o DT
          ├── 11. Ctrl+F "DO DT" + paste_text(numero_dt) — preenche campo DT
          ├── 12. Ctrl+F "mero do DT" + paste_text(numero_dt) — GAP: 2ª ocorrência
          ├── 13. focused_alert() — aviso para baixar XML do CT-e
          ├── 14. prompt_batch_info() — CT-e, NF1, NF2, NCM (radiobutton)
          ├── 15. verify_cte_on_page() — confirma CT-e via 3 estratégias de match
          ├── 16. ensure_caps_off() — desliga Caps Lock
          ├── 17. navigate_to_mdfe() — Ctrl+3 → busca "EMITIR NOTA" → "MDF-E"
          ├── 18. wait_for_form("Emissor MDF-e") — aguarda formulário abrir
          ├── 19. fill_mdfe(profile, ncm) — preenche formulário MDF-e
          ├── 20. fill_modal_rodo(profile) — preenche Modal Rodoviário
          ├── 21. fill_additional_info(profile) — preenche Informações Adicionais
          ├── 22. perform_averbacao(cte, dt, nf) — executa averbação
          ├── 23. restore_console_popup() + play_low_beep() — alerta ao operador
          ├── 24. Resumo final (DT, CT-e, NCM, NF, tempos de automação × real)
          └── 25. [WIP — próxima tarefa] Integração do CIOT como etapa obrigatória
```

---

## 7. Log de Decisões de Engenharia

> Este registro deve ser atualizado a cada decisão relevante.
> Format: **[DATA] TÍTULO** — contexto, opções avaliadas, decisão tomada, consequências.

---

### [2025-06] Refatoração Parte 5 — Etapas de Automação, Runner e Wrapper Final

**Tipo:** Refatoração

**Módulo(s) criados:** `mdfe/steps/navigate.py`, `mdfe/steps/fill_mdfe.py`, `mdfe/steps/fill_modal_rodo.py`, `mdfe/steps/fill_additional_info.py`, `mdfe/steps/averbacao.py`, `mdfe/steps/__init__.py`, `mdfe/runner.py`
**Módulo(s) modificados:** `modular_mdfe.py`, `mdfe/__init__.py`, `mdfe/profile.py`

**Problema / contexto:**
Conclusão da modularização do monólito `modular_mdfe.py`. Todas as etapas sequenciais de preenchimento (`navigate_to_mdfe`, `fill_mdfe`, `fill_modal_rodo`, `fill_additional_info`, `perform_averbacao`) e o loop principal (`main`) precisavam ser desacoplados para que o script principal se tornasse apenas um wrapper leve.

**Solução adotada:**
- Criação do subpacote `mdfe/steps/` contendo cada etapa da automação em um arquivo dedicado.
- Criação de `mdfe/runner.py` contendo a lógica central de `main()`.
- O script legado `modular_mdfe.py` foi reduzido a 6 linhas, atuando puramente como um redirecionador (wrapper) para `mdfe.runner.main()`, mantendo total compatibilidade com chamadas externas e com o `run.bat`.
- A função interativa `choose_profile` foi migrada de `modular_mdfe.py` para `mdfe/profile.py`.

**Verificação realizada:**
- `py_compile` em todos os módulos novos, reestruturados e no wrapper: ✓ OK
- Compatibilidade de chamada garantida para `run.bat`.

**Arquivos modificados:**

| Arquivo | Tipo de alteração |
|---|---|
| `mdfe/steps/navigate.py` | Criado |
| `mdfe/steps/fill_mdfe.py` | Criado |
| `mdfe/steps/fill_modal_rodo.py` | Criado |
| `mdfe/steps/fill_additional_info.py` | Criado |
| `mdfe/steps/averbacao.py` | Criado |
| `mdfe/steps/__init__.py` | Criado |
| `mdfe/runner.py` | Criado |
| `mdfe/__init__.py` | Modificado (expõe `main`) |
| `mdfe/profile.py` | Modificado (adicionado `choose_profile`) |
| `modular_mdfe.py` | Modificado (reduzido a wrapper leve) |

---

### [2025-06] Refatoração Parte 4 — Utilitários de Navegação e Browser

**Tipo:** Refatoração

**Módulo(s) criados:** `mdfe/browser.py`
**Módulo(s) modificados:** `modular_mdfe.py`

**Problema / contexto:**
Segregação de responsabilidades. Funções auxiliares para lidar com foco do navegador, detecção de janelas ativas e verificação de campos no clipboard (`_get_foreground_title`, `focus_browser_if_needed`, `wait_for_form`, `verify_cte_on_page`, etc.) estavam misturadas no monólito, dificultando sua leitura.

**Solução adotada:**
Extração das 13 funções utilitárias do navegador para o módulo `mdfe/browser.py`. O arquivo `modular_mdfe.py` passa a importá-las (`focus_browser_if_needed`, `wait_for_form`, `verify_cte_on_page`), eliminando cerca de 340 linhas de código do script principal. A fidelidade do comportamento é mantida idêntica à versão legada.

**Verificação realizada:**
- `py_compile` em todo o pacote `mdfe/` e no `modular_mdfe.py`: ✓ OK
- Execução parcial e análise das assinaturas: ✓ OK (sem alteração de dependências ou ordem dos parâmetros)

**Arquivos modificados:**

| Arquivo | Tipo de alteração |
|---|---|
| `mdfe/browser.py` | Criado |
| `modular_mdfe.py` | Modificado (definições removidas e substituídas por imports) |

---

### [2025-06] Refatoração Parte 3 — Perfis, Diálogos e Teclado

**Tipo:** Refatoração

**Módulo(s) criados:** `mdfe/profile.py`, `mdfe/dialogs.py`, `mdfe/keyboard.py`
**Módulo(s) modificados:** `modular_mdfe.py`

**Problema / contexto:**
Primeiro impacto direto no monólito `modular_mdfe.py`. Era necessário extrair as principais rotinas de entrada de dados (diálogos tkinter), gerenciamento de arquivos de rotas (ConfigProfile) e funções de simulação de digitação/teclado (`press_tab`, `smart_write`, `paste_text`).

**Solução adotada:**
Isolamento em módulos específicos:
- `mdfe/profile.py`: Gerencia carregamento/monitoramento de mudanças em arquivos de rota.
- `mdfe/dialogs.py`: Controla caixas de confirmação, avisos e prompts (inclusive tkinter topmost).
- `mdfe/keyboard.py`: Simulação do comportamento de escrita inteligente e manipulação de clipboard.
A integração foi feita no `modular_mdfe.py` substituindo as implementações originais pelos respectivos imports.

**Verificação realizada:**
- `py_compile` nos novos módulos e no principal: ✓ OK
- Execução completa usando perfil de teste: ✓ OK
- Logs gerados no mesmo formato.

**Arquivos modificados:**

| Arquivo | Tipo de alteração |
|---|---|
| `mdfe/profile.py` | Criado |
| `mdfe/dialogs.py` | Criado |
| `mdfe/keyboard.py` | Criado |
| `modular_mdfe.py` | Modificado |

---

### [2025-06] Refatoração Parte 2 — Módulos de sistema e console

**Tipo:** Refatoração

**Módulos criados:** `mdfe/console.py`, `mdfe/instance.py`, `mdfe/pause.py`

**Problema / contexto:**
Continuação da modularização. Funções de sistema (console, mutex, pausa) estavam
misturadas com automação no monólito. Estas funções não dependem de pyautogui e
podiam ser isoladas sem qualquer risco de alteração de comportamento.

**Solução adotada:**

- **`console.py`**: `hide_console_window`, `restore_console_popup`, `play_low_beep` —
  cópia exata das funções originais. Sem dependências externas ao stdlib.

- **`instance.py`**: `ensure_single_instance` + `_SINGLETON_MUTEX_HANDLE` —
  a chamada a `focused_alert` (que virá em `mdfe.dialogs` na Parte 3) é feita via
  lazy import dentro da função para evitar importação circular em nível de módulo.

- **`pause.py`**: `show_pause_dialog`, `check_pause`, `pause_point`,
  `_verify_last_write_before_pause` —
  importa o estado de pausa (`_pause_requested`, `_pause_active`, `_pause_lock`)
  de `mdfe.failsafe` (Parte 1, que é o dono natural desse estado por ser o listener F9).
  Referências a `paste_text` e `_normalize_text` (Parte 3) são lazy imports dentro
  de `_verify_last_write_before_pause`.

**Impacto no operador:** nenhum — `modular_mdfe.py` não foi alterado.

**Verificação realizada:**
- `py_compile` em `mdfe/console.py`, `mdfe/instance.py`, `mdfe/pause.py`: ✓ OK
- `py_compile` em `modular_mdfe.py`: ✓ intacto

**Arquivos modificados:**

| Arquivo | Tipo de alteração |
|---|---|
| `mdfe/console.py` | Criado |
| `mdfe/instance.py` | Criado |
| `mdfe/pause.py` | Criado |



**Tipo:** Refatoração

**Módulos criados:** `mdfe/__init__.py`, `mdfe/logger.py`, `mdfe/timing.py`, `mdfe/failsafe.py`

**Problema / contexto:**
`modular_mdfe.py` com 2228 linhas misturava logging, timers, failsafe, diálogos, perfis,
automação e fluxo principal em um único arquivo. Isso dificultava localizar funções
e manter a documentação por responsabilidade.

**Solução adotada:**
Criação do pacote `mdfe/` com os primeiros três módulos de infraestrutura.
Nesta parte, `modular_mdfe.py` não foi modificado — os módulos existem mas ainda não
são importados pelo principal. A integração ocorre nas Partes 3–5.

Dependência circular resolvida com lazy import dentro das funções:
- `mdfe/timing.py → mdfe/logger.log` (via import dentro de `resume_automation_timer`)
- `mdfe/logger.py → mdfe/timing` (via import dentro de `start_automation_session`)
- `mdfe/failsafe.py → mdfe/logger.log` (via import dentro de `on_press` e `start_failsafe_f8`)

**Opções descartadas:**
- `state.py` centralizado para os globals de pausa: descartado porque distribui
  responsabilidade nos módulos proprietários (timing owns timing state; failsafe owns
  pause trigger state), o que é mais coeso e rastreável.

**Impacto no operador:** nenhum — `modular_mdfe.py` não foi alterado.

**Impacto técnico:**
- Pacote `mdfe/` criado no projeto.
- `_pause_requested`, `_pause_active`, `_pause_lock` e `request_pause()` vivem em
  `mdfe/failsafe.py` (pois F9 é o gatilho). `mdfe/pause.py` importará daqui na Parte 2.
- `_automation_start_time` e `_automation_time_paused` vivem em `mdfe/timing.py`.
  `start_automation_session` (em `logger.py`) mutua esses globals via lazy import.

**Verificação realizada:**
- `py_compile` em `mdfe/__init__.py`, `mdfe/logger.py`, `mdfe/timing.py`, `mdfe/failsafe.py`: ✓ OK
- `py_compile` em `modular_mdfe.py`: ✓ intacto
- `run.bat` não foi alterado.

**Arquivos modificados:**

| Arquivo | Tipo de alteração |
|---|---|
| `mdfe/__init__.py` | Criado |
| `mdfe/logger.py` | Criado |
| `mdfe/timing.py` | Criado |
| `mdfe/failsafe.py` | Criado |

---



### [2025-06] Design: Python por `.venv` local, sem PATH

**Contexto:** Nos computadores corporativos, o Python nunca está no PATH. Não há
garantia de que `python` ou `py` funcionem em qualquer terminal.

**Decisão:** `run.bat` detecta o Python disponível, cria `.venv` no diretório do
projeto e usa o caminho absoluto `%VENV_DIR%\Scripts\python.exe` para todas as
invocações. Isso isola o ambiente e não requer intervenção da TI por rota.

**Consequência:** Instalação de dependências (opção 3 do menu) é obrigatória na
primeira execução em cada máquina. O operador não precisa de privilégios elevados.

---

## 8. Template de Change Record

> **Instruções:** copie esta seção para o Log de Decisões ao registrar qualquer
> melhoria, correção de bug ou novo módulo entregue ao cliente.

---

### [AAAA-MM] Título da mudança

**Tipo:** `Correção de bug` | `Nova funcionalidade` | `Refatoração` | `Ajuste de perfil`

**Módulo(s) afetado(s):** `<nome do arquivo>`

**Problema / contexto:**  
Descreva o comportamento anterior ou a necessidade que motivou a mudança.

**Solução adotada:**  
Descreva o que foi feito e por quê essa abordagem foi escolhida.

**Opções descartadas (se aplicável):**  
Liste alternativas consideradas e o motivo de não terem sido escolhidas.

**Impacto no operador:**  
Descreva o que muda na experiência de uso (ex.: "novo campo no diálogo de dados").

**Impacto técnico:**  
Descreva efeitos colaterais, novos módulos criados, constantes adicionadas, etc.

**Verificação realizada:**  
Liste os testes executados para validar a mudança (ex.: `py_compile`, execução com
perfil de teste, comparação de sequência de teclas com versão anterior).

**Arquivos modificados:**

| Arquivo | Tipo de alteração |
|---|---|
| `exemplo.py` | Modificado |
| `novo_modulo.py` | Criado |
| `constants.py` | Constante adicionada |
