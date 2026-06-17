"""Motor de execução principal (runner) da automação MDF-e."""

import argparse
import time
import os
import sys
import pyautogui
from pathlib import Path

# Importar constantes
from constants import (
    SLEEP_MEDIUM, SLEEP_LONG, SLEEP_LONGER, SLEEP_ONE, SLEEP_SHORT,
    FORM_WAIT_TIMEOUT, FORM_COPY_ATTEMPTS, SLEEP_ONE_HALF
)

# Importar utilitários do pacote mdfe
from mdfe.logger import log, ui_print, start_automation_session
from mdfe.timing import (
    format_duration, _automation_start_time, _automation_time_paused
)
import mdfe.timing as _timing
from mdfe.failsafe import start_failsafe_f8, stop_failsafe_f8, request_pause
from mdfe.pause import pause_point
import mdfe.failsafe as _failsafe
from mdfe.console import hide_console_window, restore_console_popup, play_low_beep
from mdfe.instance import ensure_single_instance
from mdfe.dialogs import focused_alert, focused_confirm, prompt_dt_blocking, prompt_batch_info
from mdfe.keyboard import paste_text, smart_write, ensure_caps_off
from mdfe.profile import ConfigProfile, list_profiles, choose_profile, CONFIG_DIR
from mdfe.browser import focus_browser_if_needed, wait_for_form, verify_cte_on_page

# Importar etapas da automação
from mdfe.steps import (
    navigate_to_mdfe, fill_mdfe, fill_modal_rodo,
    fill_additional_info, perform_averbacao
)

BASE_DIR = Path(__file__).parent.parent


def main() -> None:
    """Fluxo principal da automacao MDF-e."""
    global _automation_start_time, _automation_time_paused

    prev_failsafe = pyautogui.FAILSAFE
    pyautogui.FAILSAFE = False
    try:
        # Iniciar contadores de tempo após seleção do script
        real_start_time = 0.0
        _timing._automation_time_paused = 0.0
        
        # Sleep inicial para aguardar inicialização
        time.sleep(SLEEP_LONGER)
        
        # Limpar tela e mostrar header
        os.system('cls' if os.name == 'nt' else 'clear')
        ui_print("AUTOMAÇÃO MDF-e", style="header")
        
        # Impedir execução duplicada
        ensure_single_instance()
        
        # Selecionar perfil de configuracao
        parser = argparse.ArgumentParser()
        parser.add_argument("--profile", help="Name of profile file inside scripts/", default=None)
        parser.add_argument("--skip-ciot", action="store_true", help="Do not prompt to run CIOT extension (used when relaunched after CIOT)")
        args = parser.parse_args()

        selected = args.profile
        if not selected:
            log("Exibindo menu de seleção de perfil")
            selected = choose_profile(list_profiles())
        ui_print(f"Perfil: {selected}", style="success")
        profile_path = CONFIG_DIR / selected
        if not profile_path.exists():
            log(f"Perfil {profile_path} não encontrado.")
            try:
                focused_alert(
                    f"O arquivo de perfil não foi encontrado:\n{profile_path}\n\nCorrija o nome do script e tente novamente.",
                    title="Perfil ausente"
                )
            except Exception:
                pass
            raise SystemExit(1)

        profile = ConfigProfile(profile_path)
        real_start_time = start_automation_session(selected, profile_path)
        start_failsafe_f8()
        pause_point()

        ui_print("Iniciando preenchimento...", style="step")
        
        # Preparar navegador e validar pagina inicial
        # Abrir/focar navegador sem minimizar (usa Win+1 só se não estiver em foco)
        log("Focando navegador (evitando minimizar)")
        focus_browser_if_needed()
        pause_point()
        
        # GAP - Pressionar ESC 2x
        log("Enviando ESC x2")
        for _ in range(2):
            pyautogui.press("esc")
            time.sleep(SLEEP_MEDIUM)

        # Recarregar aba 3
        log("Recarregando aba 3 (Ctrl+3, F5)")
        pyautogui.hotkey("ctrl", "3")
        time.sleep(SLEEP_LONG)
        pyautogui.press("f5")
        time.sleep(SLEEP_ONE)
        pause_point()

        # Voltar para aba 1 (uma vez, como no legado)
        log("Voltando para aba 1")
        pyautogui.hotkey("ctrl", "1")
        time.sleep(SLEEP_LONG)
        # Tab para garantir foco correto
        pyautogui.press("esc")
        time.sleep(SLEEP_LONG)
        pause_point()

        pause_point()

        # Coleta de informacoes obrigatorias
        # Prompt para DT ANTES de buscar o campo
        prompt_text = profile.get_value("general", "dt_prompt_text")
        if not prompt_text:
            log("Perfil sem 'general.dt_prompt_text' obrigatório")
            focused_alert(
                "O perfil está faltando a chave obrigatória: general.dt_prompt_text",
                title="Perfil inválido"
            )
            raise SystemExit(1)
        log("Exibindo prompt de DT")
        numero_dt = prompt_dt_blocking(text=prompt_text, title="DT")
        if not numero_dt:
            focused_alert("Nenhum código DT informado. O script foi pausado.")
            return
        log(f"DT informado: {numero_dt}")
        log("Focando a primeira aba do navegador (Ctrl+1) apos o DT")
        pyautogui.hotkey("ctrl", "1")
        time.sleep(SLEEP_LONG)
        pause_point()

        # Posicionar em "serie final" e Tab 2x
        log("Posicionando em 'serie final' e tabulando")
        pyautogui.hotkey("ctrl", "f")
        time.sleep(1)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.15)
        pyautogui.press("backspace")
        time.sleep(0.15)
        smart_write("DO DT", interval=0.12)
        time.sleep(SLEEP_MEDIUM)
        pyautogui.press("esc")
        time.sleep(SLEEP_LONG)
        pyautogui.press("tab")
        time.sleep(SLEEP_LONG)
        
        # Usar o DT armazenado previamente
        log(f"Preenchendo campo DT com valor armazenado: {numero_dt}")
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.15)
        paste_text(numero_dt.upper(), verify=True)
        time.sleep(SLEEP_MEDIUM)
        pyautogui.press("enter")
        time.sleep(SLEEP_LONG)

        #GAP_CTRL+F
        # Posicionar em "serie final" e Tab 2x
        log("Posicionando em 'serie final' e tabulando")
        pyautogui.hotkey("ctrl", "f")
        time.sleep(1)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.15)
        pyautogui.press("backspace")
        time.sleep(0.15)
        smart_write("mero do DT", interval=0.12)
        time.sleep(SLEEP_MEDIUM)
        pyautogui.press("esc")
        time.sleep(SLEEP_LONG)
        pyautogui.press("tab")
        time.sleep(SLEEP_LONG)
        
        # Usar o DT armazenado previamente
        log(f"Preenchendo campo DT com valor armazenado: {numero_dt}")
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.15)
        paste_text(numero_dt.upper(), verify=True)
        time.sleep(SLEEP_MEDIUM)
        pyautogui.press("enter")
        time.sleep(SLEEP_LONG)

        # Aviso inicial para baixar o CT-e, antes do prompt unificado
        log("Exibindo aviso para baixar o CT-e antes de seguir")
        focused_alert(
            text=(
                "Antes de prosseguir, faça o download do XML do CT-e correspondente \n"
                "à DT e mantenha-o salvo. Em seguida clique em OK para continuar."
            ),
            title="Aviso: Baixe o CT-e"
        )
        pause_point()

        ncm_primary = profile.get_value("mdfe", "ncm_primary")
        ncm_secondary = profile.get_value("mdfe", "ncm_secondary")
        ncm_tertiary = profile.get_value("mdfe", "ncm_tertiary")
        if not (ncm_primary and ncm_secondary and ncm_tertiary):
            log("Perfil sem códigos NCM obrigatórios (mdfe.ncm_primary/secondary/tertiary)")
            focused_alert(
                "O perfil está faltando códigos NCM obrigatórios:\n"
                "mdfe.ncm_primary, mdfe.ncm_secondary, mdfe.ncm_tertiary",
                title="Perfil inválido"
            )
            raise SystemExit(1)

        log("Exibindo prompt unificado para CT-e, NFs e NCM")
        batch = prompt_batch_info([ncm_primary, ncm_secondary, ncm_tertiary])
        if not batch:
            focused_alert("Nenhuma informação foi informada. O script foi pausado.")
            raise SystemExit(1)

        log("Focando a primeira aba do navegador (Ctrl+1) apos o prompt de dados")
        pyautogui.hotkey("ctrl", "1")
        time.sleep(SLEEP_LONG)

        numero_cte = batch.get("cte", "")
        nf1 = batch.get("nf1", "")
        nf2 = batch.get("nf2", "")
        codigo_ncm = batch.get("ncm", "")

        log(f"CT-e informado: '{numero_cte}'")
        if not numero_cte:
            log("Nenhum número de CT-e informado; encerrando automação conforme solicitado")
            focused_alert(
                "ERRO: Nenhum número de CT-e foi informado.\n\n"
                "A automação será encerrada.",
                title="CT-e obrigatório"
            )
            raise SystemExit(1)

        # Verificar se o CT-e informado aparece na página após inserir a DT
        verify_cte_on_page(numero_cte)
        pause_point()

        # Concatenar apenas se ambas informadas; caso contrário usar a que foi preenchida
        if nf1 and nf2:
            nf_concat = f"{nf1}/{nf2}"
        elif nf1:
            nf_concat = nf1
        elif nf2:
            nf_concat = nf2
        else:
            nf_concat = ""
        log(f"NF coletadas: '{nf1}' e '{nf2}' => '{nf_concat}'")
        log(f"NCM selecionado e armazenado: {codigo_ncm}")
        ui_print(f"NCM selecionado: {codigo_ncm}", style="success")
        time.sleep(SLEEP_LONG)
        pyautogui.press("esc")
        time.sleep(0.15)
        ensure_caps_off()
        pause_point()

        ui_print("Preenchendo formulário MDF-e...", style="step")
        
        # Preencher formularios principais
        # Navegar para MDF-e e detectar formulário (lógica e tempos do legado)
        navigate_to_mdfe()
        wait_for_form("Emissor MDF-e", tempo_maximo=FORM_WAIT_TIMEOUT, intervalo=3.0, copy_attempts=FORM_COPY_ATTEMPTS)
        pause_point()
        
        # Preencher formulário (passando código NCM já selecionado)
        log("Iniciando preenchimento dos dados MDF-e")
        fill_mdfe(profile, codigo_ncm)
        log("Dados MDF-e preenchidos com sucesso")
        ui_print("Dados MDF-e preenchidos", style="success")
        pause_point()
        
        log("Iniciando preenchimento do modal rodoviário")
        fill_modal_rodo(profile)
        log("Modal rodoviário preenchido com sucesso")
        ui_print("Modal rodoviário preenchido", style="success")
        pause_point()
        
        log("Iniciando preenchimento de informações adicionais")
        fill_additional_info(profile)
        log("Informações adicionais preenchidas com sucesso")
        ui_print("Informações adicionais preenchidas", style="success")
        pause_point()
        
        log("Iniciando processamento de averbação")
        ui_print("Processando averbação...", style="step")
        perform_averbacao(numero_cte, numero_dt, nf_concat)
        log("Averbação processada com sucesso")
        pause_point()
        
        # Encerramento com resumo
        # Ao finalizar com sucesso, calcular tempos
        real_end_time = time.monotonic()
        automation_end_time = time.monotonic()
        
        real_duration = real_end_time - real_start_time
        automation_duration = (automation_end_time - _timing._automation_start_time) - _timing._automation_time_paused
        
        log(f"[DEBUG TIMING] real_end_time={real_end_time}, automation_end_time={automation_end_time}")
        log(f"[DEBUG TIMING] _automation_start_time={_timing._automation_start_time}, _automation_time_paused={_timing._automation_time_paused}")
        log(f"[DEBUG TIMING] real_duration calc: {real_end_time} - {real_start_time} = {real_duration}")
        log(f"[DEBUG TIMING] automation_duration calc: ({automation_end_time} - {_timing._automation_start_time}) - {_timing._automation_time_paused} = {automation_duration}")
        
        # Restaurar o terminal como popup e emitir beep baixo
        restore_console_popup()
        play_low_beep()
        
        # Exibir resumo no terminal estilo GUI
        GREEN = "\033[92m"
        CYAN = "\033[96m"
        YELLOW = "\033[93m"
        BOLD = "\033[1m"
        RESET = "\033[0m"
        
        print(f"\n{CYAN}{'═' * 60}{RESET}")
        print(f"{BOLD}{GREEN}  ✓ AUTOMAÇÃO CONCLUÍDA COM SUCESSO!{RESET}")
        print(f"{CYAN}{'═' * 60}{RESET}\n")
        print(f"{BOLD}Resumo das Informações:{RESET}\n")
        print(f"  {YELLOW}DT:{RESET}      {numero_dt}")
        print(f"  {YELLOW}CT-e:{RESET}    {numero_cte if numero_cte else 'Não capturado'}")
        print(f"  {YELLOW}NCM:{RESET}     {codigo_ncm}")
        print(f"  {YELLOW}NF:{RESET}      {nf_concat if nf_concat else 'Não informado'}")
        print(f"\n{BOLD}Tempo de Execução:{RESET}\n")
        print(f"  {YELLOW}Tempo de Automação:{RESET}  {format_duration(automation_duration)} (apenas automação)")
        print(f"  {YELLOW}Tempo Real:{RESET}          {format_duration(real_duration)} (incluindo prompts)")
        print(f"\n{CYAN}{'─' * 60}{RESET}")
        print(f"{BOLD}Próximos passos:{RESET}")
        print(f"  • Preencha os dados do motorista")
        print(f"\n{CYAN}{'═' * 60}{RESET}\n")
        
        # Pausar 3 segundos para permitir leitura do resumo
        time.sleep(3)
        
        log(f"Automação finalizada com sucesso - Tempo automação: {format_duration(automation_duration)}, Tempo real: {format_duration(real_duration)}")
        log("═" * 60)
        log(f"Resumo final: DT={numero_dt}, CT-e={numero_cte if numero_cte else 'Não capturado'}, NCM={codigo_ncm}, NF={nf_concat if nf_concat else 'Não informado'}")
        log("═" * 60)
        
        # Perguntar ao usuário se deseja preencher o CIOT (pular se --skip-ciot)
        if not getattr(args, "skip_ciot", False):
            log("Perguntando ao usuário sobre preenchimento de CIOT")
            ciot_buttons = ["Sim, preencher CIOT", "Não, encerrar"]
            ciot_choice = focused_confirm(
                text=(
                    "Deseja preencher o campo CIOT (Conhecimento de Transporte Intermodal Operacional) agora?\n\n"
                    "Clique em 'Sim' para abrir a extensão de preenchimento ou 'Não' para encerrar."
                ),
                title="Preencher CIOT?",
                buttons=ciot_buttons
            )

            if ciot_choice == ciot_buttons[0]:  # Sim, preencher CIOT
                log("Usuário escolheu preencher CIOT - chamando script ciot_filler.py")
                ui_print("Abrindo extensão CIOT...", style="step")
                time.sleep(SLEEP_MEDIUM)

                try:
                    import subprocess
                    ciot_script_path = BASE_DIR / "ciot_filler.py"
                    if ciot_script_path.exists():
                        log(f"Executando: {ciot_script_path}")
                        try:
                            log("Tentando focar navegador antes de iniciar CIOT")
                            focus_browser_if_needed()
                            time.sleep(SLEEP_SHORT)
                        except Exception:
                            pass
                        try:
                            hide_console_window()
                        except Exception:
                            pass

                        cmd = [str(sys.executable), str(ciot_script_path)]
                        try:
                            if os.name == "nt":
                                subprocess.Popen(cmd, cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NEW_CONSOLE)
                            else:
                                subprocess.Popen(cmd, cwd=str(BASE_DIR))

                            log("CIOT iniciado em nova janela; encerrando automação principal.")
                            ui_print("CIOT iniciado em nova janela. Encerrando automação principal.", style="success")
                            # Usar os._exit para garantir encerramento imediato do processo
                            try:
                                os._exit(0)
                            except Exception:
                                raise SystemExit(0)
                        except Exception as e:
                            log(f"Erro ao iniciar script CIOT: {e}")
                            focused_alert(
                                f"Erro ao iniciar a extensão CIOT:\n{str(e)}",
                                title="Erro na extensão"
                            )
                    else:
                        log(f"Script CIOT não encontrado: {ciot_script_path}")
                        focused_alert(
                            f"O script de preenchimento de CIOT não foi encontrado:\n{ciot_script_path}",
                            title="Arquivo não encontrado"
                        )
                except Exception as e:
                    log(f"Erro ao executar script CIOT: {e}")
                    focused_alert(
                        f"Erro ao executar a extensão CIOT:\n{str(e)}",
                        title="Erro na extensão"
                    )
            else:
                log("Usuário escolheu não preencher CIOT - encerrando automação principal")
        else:
            log("--skip-ciot presente: pulando prompt de CIOT")
    except SystemExit as e:
        # Capturar saídas como exit code 99 (menu), 1 (erro), etc
        if e.code == 99:
            log("Programa finalizado - usuário retornou ao menu (código 99)")
        else:
            log(f"Programa finalizado com código de saída: {e.code}")
        raise
    except Exception as e:
        import traceback
        error_msg = f"ERRO FATAL: {e}\n{traceback.format_exc()}"
        log(error_msg)
        # Exibir alerta do erro
        try:
            focused_alert(
                f"Erro durante a automação:\n\n{str(e)}\n\nVer log para detalhes completos.",
                title="Erro na automação"
            )
        except Exception:
            pass
        # Restaurar terminal para que o usuário possa ver o erro
        restore_console_popup()
        raise
    finally:
        stop_failsafe_f8()
        pyautogui.FAILSAFE = prev_failsafe
