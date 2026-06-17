# AutoMDFText — Documentação Técnica do Projeto

> **Versão do documento:** 1.0  
> **Data:** Junho 2025  
> **Escopo:** Estado atual do projeto + schema para registrar cada melhoria futura  

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
4. Ao terminar, o operador tem a opção de preencher o campo CIOT via módulo complementar.

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

```
┌──────────────────────────────────────────────────────────────────┐
│  Ponto de entrada do operador                                    │
│  run.bat                                                         │
│  • Detecta / cria .venv                                         │
│  • Guard de instância duplicada                                  │
│  • Menu: [1] MDF-e  [2] Editor  [3] Deps  [4] Sair             │
└───────────┬──────────────────────┬───────────────────────────────┘
            │                      │
            ▼                      ▼
   modular_mdfe.py          script_editor.py
   (automação MDF-e)        (editor de perfis GUI)
            │
            ├── constants.py       ← constantes centralizadas
            ├── scripts/*.txt      ← perfis de rota
            ├── logs/              ← logs de sessão
            └── ciot_filler.py    ← módulo complementar (CIOT)
                   (lançado como subprocesso ao final)
```

### 3.2 Fluxo de chamada entre módulos

```
run.bat
  └─► modular_mdfe.py (main)
        ├── lê constants.py
        ├── lê scripts/<perfil>.txt via ConfigProfile
        ├── prompts (DT, CT-e, NF, NCM)
        ├── navigate_to_mdfe()
        ├── fill_mdfe(profile, ncm)
        ├── fill_modal_rodo(profile)
        ├── fill_additional_info(profile)
        ├── perform_averbacao(cte, dt, nf)
        └── [opcional] subprocess → ciot_filler.py
```

---

## 4. Documentação dos Arquivos e Funções

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
configurações numéricas usados tanto na automação principal quanto em módulos
complementares (ex.: `ciot_filler.py`). Centralizar aqui evita que um ajuste de
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

### `modular_mdfe.py` — Motor principal de automação

**Responsabilidade:** orquestra o fluxo completo de preenchimento do MDF-e.
Contém todas as funções de automação, utilitários de sistema e a função `main()`.

> **Nota:** este arquivo está em processo de refatoração modular planejada.
> A documentação abaixo mapeia as funções *como estão hoje* antes da reestruturação.

#### Seção: Logging e sessão

---

##### `log(msg: str) → None`

Registra uma mensagem no arquivo de log da sessão corrente (`logs/automation_<ts>.log`).
Não imprime no console — saída de console usa `ui_print`. Silencia exceções de I/O
para não interromper a automação por falha de disco.

---

##### `ui_print(msg: str, style: str = "info") → None`

Imprime no console com formatação ANSI colorida de acordo com o `style`:

| style | Visual | Uso |
|---|---|---|
| `"header"` | Azul com separadores `═` | Títulos de seção |
| `"step"` | Azul `▸` | Início de etapa |
| `"success"` | Verde `✓` | Conclusão bem-sucedida |
| `"error"` | Vermelho `✗` | Falha |
| `"warning"` | Amarelo `⚠` | Aviso não-crítico |
| `"info"` (default) | Texto simples indentado | Informações gerais |

---

##### `start_automation_session(selected: str, profile_path: Path) → float`

Inicializa uma nova sessão de log com timestamp único, reinicia os contadores de tempo
(`_automation_start_time`, `_automation_time_paused`) e retorna o instante monotônico
do início real (inclui o tempo de seleção de perfil e prompts, para cálculo de tempo total).

---

#### Seção: Failsafe e pausa

---

##### `start_failsafe_f8() → None`

Inicia um thread listener de teclado global usando `pynput`. Monitora duas teclas:
- **F8:** encerra a automação imediatamente via `os._exit(1)` e exibe alerta Win32.
- **F9:** sinaliza uma pausa para o próximo ponto seguro (não interrompe sequências de Tab/Enter).

Em Windows, usa o filtro `win32_event_filter` para rejeitar eventos injetados
(simulados por outro software) e aceitar apenas teclas físicas. Isso evita que a
própria automação acione o failsafe ao enviar teclas.

---

##### `stop_failsafe_f8() → None`

Para o listener de teclado iniciado por `start_failsafe_f8`. Chamado no bloco `finally`
do `main` para garantir que o listener não fique ativo após o encerramento.

---

##### `request_pause() → None`

Sinaliza `_pause_requested = True` de forma thread-safe. Chamada pelo listener de F9.
A pausa efetiva só acontece quando o fluxo principal chega ao próximo `pause_point()`.

---

##### `show_pause_dialog() → str`

Exibe uma janela tkinter topmost com botões "Retomar" e "Cancelar automação". Inclui
um watchdog que reaplica o status `HWND_TOPMOST` a cada 600 ms para garantir que o
diálogo permaneça visível mesmo que o navegador tente tomar o foco. Retorna `"resume"`
ou `"cancel"`.

---

##### `check_pause() → None`

Verifica se uma pausa foi solicitada. Se sim, pausa o timer de automação, exibe o
diálogo e aguarda a decisão do operador. Se o operador cancelar, levanta `SystemExit(1)`.
Após retomar, marca `_pause_requested = False` e recomeça o timer.

---

##### `pause_point() → None`

Ponto seguro de checagem, chamado entre etapas do fluxo (nunca dentro de sequências
de Tab/Enter). Delega para `check_pause()`. O padrão de uso é inserir `pause_point()`
entre funções de preenchimento para que o operador sempre possa pausar sem risco de
o formulário ficar em estado inconsistente.

---

##### `_verify_last_write_before_pause() → None`

Antes de bloquear na pausa, revalida o último valor enviado por `smart_write` usando
Ctrl+A / Ctrl+C para comparar o conteúdo do campo com o valor esperado. Se houver
divergência, reaplicá via `paste_text`. Isso evita que um campo fique vazio se o
operador pausar imediatamente após um `smart_write`.

---

#### Seção: Timers e métricas

---

##### `pause_automation_timer() → None`

Registra o instante de início de uma pausa de usuário (prompt, diálogo). Não pausa
a automação em si — apenas registra o timestamp para descontar do tempo de automação.

---

##### `resume_automation_timer() → None`

Calcula o tempo decorrido desde o último `pause_automation_timer` e acumula em
`_automation_time_paused`. Chamado ao fechar prompts e diálogos.

---

##### `format_duration(seconds: float) → str`

Formata uma duração em segundos para exibição legível:
- Menos de 60 s → `"42s"`
- 60 s ou mais → `"2m 15s"`

---

#### Seção: Utilitários de teclado e clipboard

---

##### `press_tab(count: int = 1, delay: float = TAB_DELAY) → None`

Pressiona Tab `count` vezes com `delay` entre cada pressionamento. Usado em toda
navegação por formulário onde a posição do campo é conhecida pela contagem de Tabs.

---

##### `skip_tabs(count: int, log_msg: str = "") → None`

Wrapper de `press_tab` com logging opcional. Usado para pular campos não editáveis
ou desnecessários na sequência de navegação.

---

##### `paste_text(text, verify, retries, delay, restore_clipboard) → None`

Cola texto via clipboard com verificação opcional:
1. Salva o conteúdo anterior do clipboard.
2. Copia `text` para o clipboard e pressiona Ctrl+V.
3. Se `verify=True`, lê o campo com Ctrl+A / Ctrl+C e compara com o valor esperado.
4. Tenta até `retries + 1` vezes se a verificação falhar.
5. Restaura o clipboard ao valor anterior (parâmetro `restore_clipboard`).

CPF e CNPJ (11 e 14 dígitos numéricos) têm `verify` desativado automaticamente
porque o formulário aplica máscara e o valor lido difere do digitado.

---

##### `smart_write(value, interval, min_paste_len, verify) → None`

Escolhe automaticamente entre colar (clipboard) e digitar (pyautogui.write):
- **Clipboard** se `len(text) >= min_paste_len` (padrão: 4) ou se o texto contém
  espaço, `/`, `-`, `_`, `:`, `.` ou tab — caracteres que digitação caractere a
  caractere pode não enviar corretamente.
- **Digitação** para textos curtos sem caracteres especiais.

Também registra o último valor para `_verify_last_write_before_pause`.

---

##### `ensure_caps_off() → None`

Verifica o estado do Caps Lock via `GetKeyState` (Win32) e o desativa se estiver
ligado, pressionando e soltando a tecla programaticamente. Chamado antes de iniciar
o preenchimento para evitar que textos maiúsculos esperados (como UFs) sejam enviados
em minúsculas ou vice-versa.

---

##### `upload_latest_xml() → None`

Seleciona o arquivo mais recente da pasta Downloads pelo timestamp de criação e
escreve o caminho completo no campo de upload via `smart_write`, depois confirma
com Enter. Usado tanto no preenchimento do MDF-e quanto na averbação.

---

#### Seção: Gerenciamento de janelas e navegador

---

##### `_get_foreground_title() → str`

Retorna o título da janela atualmente em foco usando `GetForegroundWindow` e
`GetWindowTextW` (Win32).

---

##### `_get_foreground_class() → str`

Retorna a classe de janela (`GetClassNameW`) da janela em foco. A classe `Chrome_WidgetWin_1`
identifica janelas do Chrome/Edge de forma mais confiável que o título.

---

##### `_get_window_process_name(hwnd: int) → str`

Dado um handle de janela, abre o processo correspondente e lê o nome do executável
via `GetModuleBaseNameW`. Retorna strings como `"msedge.exe"` ou `"chrome.exe"`.

---

##### `_is_cloaked_window(hwnd: int) → bool`

Verifica se a janela está "cloaked" (oculta pela DWM, típico de apps UWP em background).
Janelas cloaked são excluídas da busca por navegador para evitar falsos positivos.

---

##### `_is_top_level_app_window(hwnd: int) → bool`

Filtra janelas utilitárias: owned windows (pop-ups filhos) e janelas com estilo
`WS_EX_TOOLWINDOW` são excluídas. Mantém apenas janelas de aplicativo de nível raiz.

---

##### `_is_standard_window(hwnd: int) → bool`

Verifica se a janela tem o estilo `WS_OVERLAPPEDWINDOW` (título + bordas), filtrando
janelas sem decoração que não representam uma aplicação interativa.

---

##### `_is_browser_window(title, cls, process_name) → bool`

Combina título, classe e nome de processo para determinar se uma janela é um navegador:
- Processo `msedge.exe` ou `chrome.exe` → sempre positivo.
- Título contendo `"chrome"`, `"edge"`, `"invoisys"`, etc. → positivo.
- Classe `Chrome_WidgetWin_1` → positivo (sem depender só do título).

---

##### `_find_browser_windows() → list[int]`

Enumera todas as janelas visíveis do sistema com `EnumWindows` e filtra usando
`_is_cloaked_window`, `_is_top_level_app_window`, `_is_standard_window` e
`_is_browser_window`. Retorna a lista de handles das janelas de navegador encontradas.

---

##### `focus_browser_if_needed() → None`

Usa `_find_browser_windows` para detectar quantas janelas de navegador existem:
- **Mais de uma janela:** exibe aviso ao operador e usa Win+1 para ir para a primeira.
- **Navegador já em foco:** não faz nada (evita minimizar acidentalmente).
- **Navegador fora de foco:** pressiona Win+1 e verifica o resultado.

**Decisão de design:** Win+1 ativa o primeiro aplicativo fixado na barra de tarefas.
A convenção é que o navegador seja o primeiro item fixado, o que é válido no ambiente
corporativo deste projeto. Essa abordagem evita Alt+Tab não-determinístico.

---

##### `_click_below_edge_searchbar(offset) → None`

Clica na região de conteúdo da página (abaixo da barra de endereço do Edge) para
garantir que Ctrl+A / Ctrl+C copiem o conteúdo da página e não o conteúdo da barra
de endereço. Usa o RECT da janela em foco para calcular a posição.

---

##### `_focus_page_for_copy() → None`

Sequência: ESC (fecha qualquer pop-up residual) + `_click_below_edge_searchbar`.
Garante que o foco esteja no corpo da página antes de copiar.

---

##### `wait_for_form(target_text, tempo_maximo, intervalo, copy_attempts) → str`

Aguarda um formulário abrir verificando o conteúdo da página via clipboard.
A cada `intervalo` segundos, faz `copy_attempts` ciclos de Ctrl+A / Ctrl+C e
verifica se `target_text` está no conteúdo copiado (normalizado, case-insensitive).
Se não detectar em `tempo_maximo`, encerra com `SystemExit(1)`.

---

##### `verify_cte_on_page(numero_cte, tempo_maximo, intervalo) → None`

Verifica se o número do CT-e informado pelo operador aparece na página após o DT
ser inserido. Tenta três estratégias de match em ordem:
1. String direta.
2. Regex de dígitos isolados (sem dígitos adjacentes).
3. Comparação de strings apenas com dígitos.

Se não encontrar em `tempo_maximo`, exibe alerta e encerra.

---

#### Seção: Perfis de configuração

---

##### `parse_profile(path: Path) → dict[str, dict[str, str]]`

Lê um arquivo `.txt` de perfil linha por linha e organiza em dicionário de seções.
Ignora linhas em branco e comentários (`#`). Seções são demarcadas por `[NOME_SECAO]`.
Chaves e valores são separados por `=`.

---

##### `class ConfigProfile`

Representa um perfil carregado em memória com detecção automática de mudança no arquivo.

| Método | Descrição |
|---|---|
| `__init__(path)` | Carrega o perfil e registra o mtime |
| `reload()` | Relê o arquivo e atualiza o mtime |
| `ensure_current()` | Relê automaticamente se o mtime mudou (hot-reload) |
| `get_value(section, key, default)` | Retorna o valor da chave na seção; `default=""` se ausente |

---

##### `list_profiles() → list[str]`

Retorna os nomes de todos os arquivos `.txt` em `scripts/`, ordenados alfabeticamente.

---

##### `choose_profile(interactive_list) → str`

Menu interativo no terminal para seleção de perfil. Aceita número (índice) ou nome
exato do arquivo (case-insensitive). Opção `0` levanta `SystemExit(99)` para retornar
ao menu do `run.bat`. Limita a 100 tentativas para evitar loops infinitos. Oculta a
janela do terminal após seleção (a automação roda em background a partir desse ponto).

---

#### Seção: Diálogos e prompts

---

##### `focused_alert(text, title, button) → str`

Wrapper de `pyautogui.alert` que pausa o timer de automação durante a exibição e
o retoma ao fechar. Isso garante que o tempo de resposta do operador não seja
contabilizado como "tempo de automação".

---

##### `focused_confirm(text, title, buttons) → str`

Wrapper de `pyautogui.confirm`. Retorna a **string do botão clicado** (ex.:
`"Sim, preencher CIOT"`), não um índice numérico. Essa distinção é crítica: código
que compare o retorno com um inteiro sempre falhará silenciosamente.

---

##### `focused_prompt(text, title, default) → str | None`

Wrapper de `pyautogui.prompt` com controle de timer.

---

##### `prompt_dt_blocking(text, title) → str | None`

Diálogo tkinter personalizado para entrada do número do DT. Tem validação embutida
(não permite OK com campo vazio) e foco automático no campo de entrada. O timer de
automação fica pausado enquanto o diálogo está aberto.

---

##### `prompt_batch_info(ncm_options) → dict[str, str] | None`

Diálogo tkinter unificado para coleta de CT-e, NF1, NF2 e NCM em uma única interação
com o operador. O NCM é selecionado via radiobutton (com opção "Outro" para digitação
livre). Retorna `None` se o operador cancelar.

**Decisão de design:** consolidar todos os dados variáveis em um único diálogo reduz
interrupções no fluxo e o tempo total de entrada de dados pelo operador.

---

#### Seção: Sistema e console

---

##### `ensure_single_instance(name, on_duplicate) → None`

Cria um mutex nomeado no namespace global do Windows (`Global\AutoMDFText_Mutex`).
Se o mutex já existir (outra instância rodando), exibe alerta e encerra com
`SystemExit(0)`. O handle do mutex é mantido em `_SINGLETON_MUTEX_HANDLE` para que
o GC não o libere antes do encerramento do processo.

---

##### `hide_console_window() → None`

Oculta a janela do console (`SW_HIDE`) sem encerrar o processo. Usada após a seleção
de perfil para que a automação rode sem uma janela de terminal visível perturbando o
operador.

---

##### `restore_console_popup() → None`

Restaura e traz o console para o topo brevemente (comportamento de popup) ao final
da automação. Usa uma sequência: ShowWindow → SetForegroundWindow → SetWindowPos
com HWND_TOPMOST → reverter para HWND_NOTOPMOST. Isso garante que o terminal apareça
na frente do navegador para que o operador veja o resumo final.

---

##### `play_low_beep() → None`

Emite um beep de baixa frequência (400 Hz, 180 ms) via `winsound.Beep` ao finalizar
a automação, sinalizando ao operador que o processo terminou sem que ele precise olhar
para a tela constantemente.

---

#### Seção: Etapas de preenchimento

---

##### `navigate_to_mdfe() → None`

Navega do estado atual do navegador até o formulário do MDF-e:
1. Ctrl+3 → aba 3 (InvoiSys).
2. Ctrl+F → busca "EMITIR NOTA" → Esc → Enter (abre seção).
3. Ctrl+F → busca "MDF-E" → Esc → Enter (abre formulário).

---

##### `fill_mdfe(profile: ConfigProfile, codigo_ncm: str) → None`

Preenche o formulário principal do MDF-e na ordem exata de campos do InvoiSys:

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

##### `fill_modal_rodo(profile: ConfigProfile) → None`

Preenche a seção "Modal Rodoviário" do formulário:

| Campo | Chave do perfil | Seção |
|---|---|---|
| RNTRC | `rntrc` | `[MODAL_RODOVIARIO]` |
| Nome do Contratante | `contratante_nome` | `[MODAL_RODOVIARIO]` |
| CNPJ do Contratante | `contratante_cnpj` | `[MODAL_RODOVIARIO]` |

Usa Ctrl+F para localizar a seção "MODAL RODO" antes de iniciar o preenchimento.

---

##### `fill_additional_info(profile: ConfigProfile) → None`

Preenche a seção "Informações Adicionais / Opcionais" com dados de seguro, frete e
parcelas. É a etapa mais complexa, com múltiplos sub-formulários abertos via Ctrl+F:

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

##### `perform_averbacao(numero_cte, numero_dt, nf_concat) → None`

Executa a averbação e preenche o campo de contribuinte com os dados da operação:

1. Ctrl+4 → aba de averbação.
2. Busca sequencial por "OK", "XML", "ENVIAR" via Ctrl+F + Enter (cada um abre uma
   etapa do processo de averbação).
3. Upload do XML via `upload_latest_xml()`.
4. Ctrl+A / Ctrl+C para capturar o número de averbação via regex.
5. Ctrl+3 → aba MDF-e.
6. Busca por "DETALHES" → preenche o número de averbação.
7. Busca por "CONTRIBUINTE" → preenche texto combinado:
   `"DT: <numero_dt> CTE: <numero_cte> NF: <nf_concat>"`.

---

### `ciot_filler.py` — Módulo complementar CIOT

**Responsabilidade:** preenchimento do campo CIOT (Conhecimento de Transporte
Intermodal Operacional) que foi adicionado recentemente ao formulário do InvoiSys.
É executado como **subprocesso separado** ao final da automação principal, permitindo
que o operador decida se deseja preencher o CIOT ou não.

#### Por que é um processo separado?

- O campo CIOT é opcional e foi adicionado depois da automação principal.
- Separar como subprocesso permite que o módulo seja executado independentemente
  sem iniciar o fluxo completo do MDF-e.
- O `modular_mdfe.py` encerra com `os._exit(0)` antes de lançar o subprocesso,
  liberando o mutex de instância para que o novo processo possa adquiri-lo se necessário.

#### Funções do `ciot_filler.py`

| Função | Descrição |
|---|---|
| `log(msg)` | Log próprio em `logs/extension_filler_<ts>.log` |
| `ui_print(msg, style)` | Console formatado (idêntico ao do módulo principal) |
| `focused_alert(text, title, button)` | Alert sem controle de timer (módulo simples) |
| `prompt_field_value(field_name, field_label)` | Diálogo genérico para entrada de valor |
| `smart_write(value, interval, min_paste_len)` | Versão simplificada (sem verify) |
| `press_tab(count, delay)` | Navegação por Tab |
| `navigate_to_page_bottom(delay)` | Pressiona End para ir ao fim da página |
| `find_and_click_ciot_toggle()` | Ctrl+F "INFORMAR DADOS DO CIOT" → Esc → Tab → Space |
| `restore_console_popup()` | Restaura o console após conclusão |
| `focus_browser_window(preferred_titles)` | Foca o navegador com múltiplas estratégias |
| `main()` | Fluxo: focar navegador → End → toggle → prompt CIOT → (próx. fases) |

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
    └─[1] Automação:
          │
          ├── 1. Sleep inicial (aguarda foco do operador na tela)
          ├── 2. ensure_single_instance() — bloqueia duplicatas via mutex
          ├── 3. choose_profile() — operador seleciona a rota
          ├── 4. start_automation_session() — inicia log e timers
          ├── 5. start_failsafe_f8() — ativa F8/F9
          ├── 6. focus_browser_if_needed() — garante navegador em foco
          ├── 7. ESC x2 — limpa qualquer pop-up residual
          ├── 8. Ctrl+3 + F5 — recarrega aba do InvoiSys
          ├── 9. Ctrl+1 + ESC — volta à aba 1
          ├── 10. prompt_dt_blocking() — operador informa o DT
          ├── 11. Ctrl+F "DO DT" → preenche o DT no campo
          ├── 12. Ctrl+F "mero do DT" → preenche novamente (GAP do formulário)
          ├── 13. focused_alert() — aviso para baixar o XML do CT-e
          ├── 14. prompt_batch_info() — CT-e, NF1, NF2, NCM
          ├── 15. verify_cte_on_page() — confirma CT-e na página
          ├── 16. navigate_to_mdfe() — navega ao formulário
          ├── 17. wait_for_form("Emissor MDF-e") — aguarda abertura
          ├── 18. fill_mdfe(profile, ncm) — preenche formulário MDF-e
          ├── 19. fill_modal_rodo(profile) — preenche Modal Rodoviário
          ├── 20. fill_additional_info(profile) — preenche Informações Adicionais
          ├── 21. perform_averbacao(cte, dt, nf) — executa averbação
          ├── 22. restore_console_popup() + play_low_beep() — alerta ao operador
          ├── 23. Exibição de resumo (DT, CT-e, NCM, NF, tempos)
          └── 24. focused_confirm("Preencher CIOT?")
                   ├─[Sim] subprocess → ciot_filler.py → os._exit(0)
                   └─[Não] encerramento normal
```

---

## 7. Log de Decisões de Engenharia

> Este registro deve ser atualizado a cada decisão relevante.
> Format: **[DATA] TÍTULO** — contexto, opções avaliadas, decisão tomada, consequências.

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
automação e fluxo principal em um único arquivo. Isso dificultava localizar funções,
reutilizar código no `ciot_filler.py` e manter a documentação por responsabilidade.

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

### [2025-06] Correção: `focused_confirm` retorna string, não inteiro

**Contexto:** O módulo `ciot_filler.py` não era invocado quando o operador clicava
"Sim, preencher CIOT" ao final da automação.

**Causa identificada:** A condição `if ciot_choice == 1` sempre falhava porque
`pyautogui.confirm` retorna a **string do botão clicado**, não seu índice numérico.

**Decisão:** Alterar a comparação para `if ciot_choice == ciot_buttons[0]`,
usando a lista de botões como referência para evitar hardcode de string.

**Consequência:** A condição agora funciona corretamente. O uso de `ciot_buttons[0]`
garante que uma eventual mudança no texto do botão seja refletida automaticamente.

---

### [2025-06] Design: Módulo CIOT como subprocesso separado

**Contexto:** O campo CIOT foi adicionado ao formulário do InvoiSys depois que a
automação principal já estava em produção.

**Opções avaliadas:**
1. Integrar o CIOT diretamente no `main()` como mais uma etapa.
2. Criar um módulo separado executado como subprocesso.

**Decisão:** Subprocesso separado (`ciot_filler.py`).

**Razões:**
- O CIOT é opcional — nem toda emissão exige preenchimento.
- A separação permite executar o módulo CIOT isoladamente para testes e manutenção.
- O `modular_mdfe.py` pode encerrar com `os._exit(0)` liberando o mutex, e o novo
  processo ganha controle limpo do terminal.

**Consequência:** Maior complexidade de comunicação entre processos (via subprocesso
e relançamento do principal com `--skip-ciot`), mas maior flexibilidade operacional.

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
