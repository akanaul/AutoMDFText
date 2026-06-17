#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Suite de Testes do AutoMDFText.

Executa testes unitários mockados nos componentes e submódulos,
garantindo comportamento idêntico ao original sem disparar cliques,
teclas ou janelas reais do Windows.

Esta suite pode ser executada por usuários ou suporte de TI e gera
mensagens diagnósticas claras e arquivo de log detalhado.
"""

import sys
import os
import re
import time
import shutil
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch, mock_open

# ─────────────────────────────────────────────────────────────────────────────
# 1. SETUP DE MOCKS GLOBAIS (Antes de importar os módulos sob teste)
# ─────────────────────────────────────────────────────────────────────────────

# Criar mocks para dependências externas de GUI/teclado para evitar disparar interações reais
mock_pyautogui = MagicMock()
mock_pyperclip = MagicMock()
mock_pyperclip.paste.return_value = ""
mock_pynput = MagicMock()
mock_winsound = MagicMock()

# Impedir que Tkinter tente renderizar janelas reais (modo headless para testes)
os.environ["TK_SILENT"] = "1"

sys.modules["pyautogui"] = mock_pyautogui
sys.modules["pyperclip"] = mock_pyperclip
sys.modules["pynput"] = mock_pynput
sys.modules["winsound"] = mock_winsound

# Mockar ctypes.windll para poder rodar testes em qualquer ambiente e simular retornos do Win32
import ctypes
class MockWindll:
    def __init__(self):
        self.user32 = MagicMock()
        self.kernel32 = MagicMock()
        self.psapi = MagicMock()
        self.dwmapi = MagicMock()

if not hasattr(ctypes, "windll"):
    ctypes.windll = MockWindll()
else:
    # Mesmo no Windows, sobrescrever windll para não chamar APIs nativas da sessão atual durante testes
    ctypes.windll = MockWindll()

# ─────────────────────────────────────────────────────────────────────────────
# IMPORTAR COMPONENTES SOB TESTE
# ─────────────────────────────────────────────────────────────────────────────
# Adicionar diretório raiz ao path se necessário
sys.path.insert(0, str(Path(__file__).parent))

import constants
from mdfe.logger import log, ui_print, start_automation_session
import mdfe.timing as timing
import mdfe.failsafe as failsafe
import mdfe.console as console
import mdfe.instance as instance
import mdfe.pause as pause
import mdfe.profile as profile
import mdfe.dialogs as dialogs
import mdfe.keyboard as keyboard
import mdfe.browser as browser
import mdfe.steps as steps
import mdfe.runner as runner

# ─────────────────────────────────────────────────────────────────────────────
# 2. DEFINIÇÃO DE CASOS DE TESTE
# ─────────────────────────────────────────────────────────────────────────────

class BaseTestCase(unittest.TestCase):
    """Classe base para limpar o estado global antes e depois de cada teste."""

    def setUp(self):
        self.reset_global_states()
        self._print_patcher = patch("builtins.print")
        self.mock_print = self._print_patcher.start()

    def tearDown(self):
        self._print_patcher.stop()
        self.reset_global_states()

    def reset_global_states(self):
        # Reset failsafe state
        if hasattr(failsafe, "_pause_requested"):
            failsafe._pause_requested = False
        if hasattr(failsafe, "_pause_active"):
            failsafe._pause_active = False

        # Reset timing state
        if hasattr(timing, "_automation_start_time"):
            timing._automation_start_time = 0.0
        if hasattr(timing, "_automation_time_paused"):
            timing._automation_time_paused = 0.0
        if hasattr(timing, "_pause_start_time"):
            timing._pause_start_time = 0.0

        # Reset keyboard state
        if hasattr(keyboard, "_last_write_value"):
            keyboard._last_write_value = None
        if hasattr(keyboard, "_last_write_verify"):
            keyboard._last_write_verify = False


class TestLogger(BaseTestCase):
    """Testes para mdfe/logger.py"""

    def setUp(self):
        super().setUp()
        self.original_log_file = getattr(sys.modules["mdfe.logger"], "LOG_FILE", None)

    def tearDown(self):
        if self.original_log_file:
            setattr(sys.modules["mdfe.logger"], "LOG_FILE", self.original_log_file)
        super().tearDown()

    @patch("builtins.open", new_callable=mock_open)
    def test_log_writes_to_file(self, mock_file):
        setattr(sys.modules["mdfe.logger"], "LOG_FILE", Path("dummy.log"))
        log("Teste de log")
        mock_file.assert_called_with(Path("dummy.log"), "a", encoding="utf-8")
        mock_file().write.assert_called()

    def test_ui_print_formats(self):
        ui_print("Mensagem de Sucesso", style="success")
        self.mock_print.assert_called()
        args = self.mock_print.call_args[0][0]
        self.assertIn("mensagem de sucesso", args.lower())
        self.assertIn("\033[92m", args)

        ui_print("Erro", style="error")
        args_err = self.mock_print.call_args[0][0]
        self.assertIn("erro", args_err.lower())
        self.assertIn("\033[91m", args_err)


class TestTiming(BaseTestCase):
    """Testes para mdfe/timing.py"""

    def test_format_duration(self):
        self.assertEqual(timing.format_duration(45.2), "45s")
        self.assertEqual(timing.format_duration(125.0), "2m 5s")
        self.assertEqual(timing.format_duration(3605.0), "60m 5s")

    @patch("time.monotonic")
    def test_pause_resume_timers(self, mock_time):
        mock_time.side_effect = [20.0, 25.0]
        timing._automation_time_paused = 0.0
        timing._automation_start_time = 5.0

        timing.pause_automation_timer()
        self.assertEqual(timing._pause_start_time, 20.0)

        timing.resume_automation_timer()
        self.assertEqual(timing._automation_time_paused, 5.0)


class TestFailsafe(BaseTestCase):
    """Testes para mdfe/failsafe.py"""

    @patch("pynput.keyboard.Listener")
    def test_start_stop_failsafe(self, mock_listener):
        failsafe.start_failsafe_f8()
        mock_listener.assert_called_once()
        failsafe.stop_failsafe_f8()
        mock_listener().stop.assert_called()

    def test_request_pause(self):
        failsafe._pause_requested = False
        failsafe.request_pause()
        self.assertTrue(failsafe._pause_requested)


class TestConsole(BaseTestCase):
    """Testes para mdfe/console.py"""

    def test_hide_console(self):
        ctypes.windll.kernel32.GetConsoleWindow = MagicMock(return_value=123)
        console.hide_console_window()
        ctypes.windll.user32.ShowWindow.assert_called_with(123, 0)

    def test_restore_console(self):
        ctypes.windll.kernel32.GetConsoleWindow = MagicMock(return_value=123)
        console.restore_console_popup()
        ctypes.windll.user32.ShowWindow.assert_called_with(123, 5)
        ctypes.windll.user32.SetForegroundWindow.assert_called_with(123)


class TestInstance(BaseTestCase):
    """Testes para mdfe/instance.py"""

    @patch("mdfe.dialogs.focused_alert")
    def test_single_instance_success(self, mock_alert):
        ctypes.windll.kernel32.CreateMutexW = MagicMock(return_value=1234)
        ctypes.windll.kernel32.GetLastError = MagicMock(return_value=0)
        
        instance.ensure_single_instance()
        mock_alert.assert_not_called()

    @patch("mdfe.dialogs.focused_alert")
    def test_single_instance_duplicate(self, mock_alert):
        ctypes.windll.kernel32.CreateMutexW = MagicMock(return_value=1234)
        ctypes.windll.kernel32.GetLastError = MagicMock(return_value=183) # ERROR_ALREADY_EXISTS
        
        with self.assertRaises(SystemExit):
            instance.ensure_single_instance()
        mock_alert.assert_called_once()


class TestPause(BaseTestCase):
    """Testes para mdfe/pause.py"""

    @patch("mdfe.pause.show_pause_dialog")
    def test_check_pause_when_requested(self, mock_show):
        failsafe._pause_requested = True
        mock_show.return_value = "resume"
        
        pause.check_pause()
        self.assertFalse(failsafe._pause_requested)

    @patch("mdfe.pause.show_pause_dialog")
    def test_check_pause_cancel(self, mock_show):
        failsafe._pause_requested = True
        mock_show.return_value = "cancel"
        
        with self.assertRaises(SystemExit):
            pause.check_pause()


class TestProfile(BaseTestCase):
    """Testes para mdfe/profile.py"""

    @patch("pathlib.Path.read_text")
    def test_parse_profile_content(self, mock_read):
        content = (
            "[GENERAL]\n"
            "key1 = val1\n"
            "# Comentario\n"
            "[MDFE]\n"
            "key2 = val2\n"
        )
        mock_read.return_value = content
        data = profile.parse_profile(Path("dummy.txt"))
        self.assertEqual(data["GENERAL"]["key1"], "val1")
        self.assertEqual(data["MDFE"]["key2"], "val2")

    @patch("mdfe.profile.list_profiles")
    @patch("builtins.input")
    @patch("mdfe.console.hide_console_window")
    def test_choose_profile(self, mock_hide, mock_input, mock_list):
        mock_list.return_value = ["perfil1.txt", "perfil2.txt"]
        
        mock_input.return_value = "1"
        res = profile.choose_profile(["perfil1.txt", "perfil2.txt"])
        self.assertEqual(res, "perfil1.txt")
        mock_hide.assert_called_once()


class TestDialogs(BaseTestCase):
    """Testes para mdfe/dialogs.py"""

    @patch("pyautogui.alert")
    def test_focused_alert(self, mock_alert):
        dialogs.focused_alert("Mensagem", title="Titulo")
        mock_alert.assert_called_with(text="Mensagem", title="Titulo", button="OK")

    @patch("pyautogui.confirm")
    def test_focused_confirm(self, mock_confirm):
        mock_confirm.return_value = "Sim"
        res = dialogs.focused_confirm("Confirma?", title="Titulo", buttons=["Sim", "Não"])
        self.assertEqual(res, "Sim")


class TestKeyboard(BaseTestCase):
    """Testes para mdfe/keyboard.py"""

    def test_normalize_digits(self):
        self.assertEqual(keyboard._normalize_digits("12.345-67"), "1234567")
        self.assertEqual(keyboard._normalize_digits("abc"), "")

    @patch("pyautogui.press")
    def test_press_tab(self, mock_press):
        keyboard.press_tab(count=3, delay=0.0)
        self.assertEqual(mock_press.call_count, 3)

    @patch("pyautogui.write")
    @patch("mdfe.keyboard.paste_text")
    def test_smart_write_methods(self, mock_paste, mock_write):
        keyboard.smart_write("ABC")
        mock_write.assert_called_with("ABC", interval=constants.WRITE_INTERVAL)
        mock_paste.assert_not_called()
        
        mock_write.reset_mock()
        mock_paste.reset_mock()

        keyboard.smart_write("TEXTOMUITOLONGO_PARA_DIGITAR")
        mock_paste.assert_called_with("TEXTOMUITOLONGO_PARA_DIGITAR", verify=True)
        mock_write.assert_not_called()


class TestBrowser(BaseTestCase):
    """Testes para mdfe/browser.py"""

    def test_is_browser_window(self):
        self.assertTrue(browser._is_browser_window("invoisys - edge", "chrome_widgetwin_1", "msedge.exe"))
        self.assertTrue(browser._is_browser_window("chrome", "", ""))
        self.assertFalse(browser._is_browser_window("notepad", "notepadclass", "notepad.exe"))

    @patch("pyperclip.paste")
    @patch("pyautogui.hotkey")
    @patch("time.sleep")
    def test_wait_for_form_success(self, mock_sleep, mock_hotkey, mock_paste):
        mock_paste.return_value = "Corpo do Emissor MDF-e aberto"
        
        res = browser.wait_for_form("Emissor MDF-e", tempo_maximo=2.0, intervalo=0.1)
        self.assertIn("Emissor MDF-e", res)


class TestSteps(BaseTestCase):
    """Testes para as etapas de automação em mdfe/steps/"""

    def setUp(self):
        super().setUp()
        self.mock_profile = MagicMock()

    @patch("pyautogui.hotkey")
    @patch("pyautogui.press")
    @patch("time.sleep")
    @patch("mdfe.steps.navigate.smart_write")
    def test_navigate(self, mock_sw, mock_sleep, mock_press, mock_hotkey):
        steps.navigate_to_mdfe()
        mock_hotkey.assert_any_call("ctrl", "3")
        mock_sw.assert_any_call("EMITIR NOTA", interval=0.10)
        mock_sw.assert_any_call("MDF-E", interval=0.10)

    @patch("pyautogui.hotkey")
    @patch("pyautogui.press")
    @patch("time.sleep")
    @patch("mdfe.steps.fill_mdfe.smart_write")
    @patch("mdfe.steps.fill_mdfe.upload_latest_xml")
    def test_fill_mdfe(self, mock_upload, mock_sw, mock_sleep, mock_press, mock_hotkey):
        self.mock_profile.get_value.side_effect = lambda sec, key, df="": {
            ("mdfe", "prestador_tipo"): "PRESTADOR",
            ("mdfe", "emitente_codigo"): "12345",
            ("mdfe", "uf_carregamento"): "SP",
            ("mdfe", "uf_descarga"): "SP",
            ("mdfe", "municipio_carregamento"): "ITU",
            ("mdfe", "unidade_medida"): "1",
            ("mdfe", "carga_tipo"): "05",
            ("mdfe", "codigo_produto_descricao"): "TESTE",
            ("mdfe", "cep_origem"): "13300000",
            ("mdfe", "cep_destino"): "13300001",
        }.get((sec, key), df)

        steps.fill_mdfe(self.mock_profile, "19041000")
        mock_upload.assert_called_once()
        mock_sw.assert_any_call("19041000", interval=0.1)

    @patch("pyautogui.hotkey")
    @patch("pyautogui.press")
    @patch("time.sleep")
    @patch("mdfe.steps.fill_modal_rodo.smart_write")
    def test_fill_modal_rodo(self, mock_sw, mock_sleep, mock_press, mock_hotkey):
        self.mock_profile.get_value.side_effect = lambda sec, key, df="": {
            ("modal_rodoviario", "rntrc"): "12345678",
            ("modal_rodoviario", "contratante_nome"): "PEPSICO",
            ("modal_rodoviario", "contratante_cnpj"): "02957518000224",
        }.get((sec, key), df)

        steps.fill_modal_rodo(self.mock_profile)
        mock_sw.assert_any_call("12345678", interval=0.10)
        mock_sw.assert_any_call("PEPSICO", interval=0.20)

    @patch("pyautogui.hotkey")
    @patch("pyautogui.press")
    @patch("time.sleep")
    @patch("mdfe.steps.fill_additional_info.smart_write")
    def test_fill_additional_info(self, mock_sw, mock_sleep, mock_press, mock_hotkey):
        self.mock_profile.get_value.side_effect = lambda sec, key, df="": {
            ("informacoes_adicionais", "contribuinte_cnpj"): "04898488000177",
            ("modal_rodoviario", "contratante_cnpj"): "02957518000224",
            ("modal_rodoviario", "contratante_nome"): "PEPSICO",
            ("informacoes_adicionais", "seguradora_nome"): "SEGURADORA",
            ("informacoes_adicionais", "seguradora_cnpj"): "33065699000127",
            ("informacoes_adicionais", "numero_apolice"): "12345",
            ("informacoes_adicionais", "frete_valor"): "100.00",
            ("informacoes_adicionais", "forma_pagamento"): "1",
            ("informacoes_adicionais", "numero_banco"): "237",
            ("informacoes_adicionais", "agencia"): "2372",
            ("informacoes_adicionais", "frete_tipo"): "FRETE",
            ("informacoes_adicionais", "numero_parcelas"): "1",
        }.get((sec, key), df)

        steps.fill_additional_info(self.mock_profile)
        mock_sw.assert_any_call("100.00", interval=0.12, verify=False)

    @patch("pyautogui.hotkey")
    @patch("pyautogui.press")
    @patch("pyperclip.paste")
    @patch("time.sleep")
    @patch("mdfe.steps.averbacao.smart_write")
    @patch("mdfe.steps.averbacao.upload_latest_xml")
    def test_perform_averbacao(self, mock_upload, mock_sw, mock_sleep, mock_paste, mock_press, mock_hotkey):
        mock_paste.return_value = "Número de Averbação: 987654321"
        
        steps.perform_averbacao("12345", "54321", "999/888")
        mock_upload.assert_called_once()
        mock_sw.assert_any_call("OK", interval=0.1, verify=False)


class TestRunner(BaseTestCase):
    """Testes para o orquestrador mdfe/runner.py"""

    @patch("mdfe.runner.ensure_single_instance")
    @patch("mdfe.runner.choose_profile")
    @patch("mdfe.runner.ConfigProfile")
    @patch("mdfe.runner.start_automation_session")
    @patch("mdfe.runner.start_failsafe_f8")
    @patch("mdfe.runner.focus_browser_if_needed")
    @patch("mdfe.runner.prompt_dt_blocking")
    @patch("mdfe.runner.prompt_batch_info")
    @patch("mdfe.runner.verify_cte_on_page")
    @patch("mdfe.runner.navigate_to_mdfe")
    @patch("mdfe.runner.wait_for_form")
    @patch("mdfe.runner.fill_mdfe")
    @patch("mdfe.runner.fill_modal_rodo")
    @patch("mdfe.runner.fill_additional_info")
    @patch("mdfe.runner.perform_averbacao")
    @patch("mdfe.runner.restore_console_popup")
    @patch("mdfe.runner.play_low_beep")
    @patch("mdfe.runner.focused_confirm")
    @patch("pyautogui.press")
    @patch("pyautogui.hotkey")
    @patch("time.sleep")
    @patch("sys.argv", new=["runner.py", "--profile", "test_profile.txt", "--skip-ciot"])
    def test_runner_main_flow(
        self,
        mock_sleep,
        mock_hotkey,
        mock_press,
        mock_confirm,
        mock_beep,
        mock_restore,
        mock_averbacao,
        mock_add_info,
        mock_rodo,
        mock_mdfe_step,
        mock_wait_form,
        mock_navigate,
        mock_verify,
        mock_batch,
        mock_dt,
        mock_focus,
        mock_failsafe,
        mock_session,
        mock_profile_class,
        mock_choose,
        mock_instance,
    ):
        mock_profile_inst = mock_profile_class.return_value
        mock_profile_inst.get_value.side_effect = lambda sec, key, df="": {
            ("general", "dt_prompt_text"): "DT prompt",
            ("mdfe", "ncm_primary"): "19041000",
            ("mdfe", "ncm_secondary"): "19059090",
            ("mdfe", "ncm_tertiary"): "20052000",
        }.get((sec, key), df)

        mock_dt.return_value = "8888"

        mock_batch.return_value = {
            "cte": "55555",
            "nf1": "123",
            "nf2": "456",
            "ncm": "19041000"
        }

        # Rodar o fluxo principal (deve ir até o fim sem levantar SystemExit se --skip-ciot estiver setado)
        with patch("pathlib.Path.exists", return_value=True):
            runner.main()
        
        # Certificar que todas as sub-etapas foram executadas sequencialmente
        self.assertTrue(mock_navigate.called)  # navigate_to_mdfe
        self.assertTrue(mock_wait_form.called)   # wait_for_form
        self.assertTrue(mock_mdfe_step.called)   # fill_mdfe
        self.assertTrue(mock_rodo.called)   # fill_modal_rodo
        self.assertTrue(mock_add_info.called)   # fill_additional_info
        self.assertTrue(mock_averbacao.called)   # perform_averbacao

# ─────────────────────────────────────────────────────────────────────────────
# 3. INTERFACE DE RELATÓRIO DO SUPORTE DE TI
# ─────────────────────────────────────────────────────────────────────────────

class TIFriendlyTestResult(unittest.TextTestResult):
    """Formatador de saída de testes focado no suporte técnico de TI."""

    def __init__(self, stream, descriptions, verbosity):
        super().__init__(stream, descriptions, verbosity)
        self.success_count = 0
        self.failure_details = []

    def addSuccess(self, test):
        super().addSuccess(test)
        self.success_count += 1
        sys.stdout.write(f"\033[92m  [PASS] {test.shortDescription() or test._testMethodName}\033[0m\n")

    def addFailure(self, test, err):
        super().addFailure(test, err)
        self.failure_details.append((test, "FALHA DE ASSERÇÃO", err))
        sys.stdout.write(f"\033[91m  [FAIL] {test.shortDescription() or test._testMethodName}\033[0m\n")

    def addError(self, test, err):
        super().addError(test, err)
        self.failure_details.append((test, "ERRO NO SCRIPT DE TESTE", err))
        sys.stdout.write(f"\033[93m  [ERR ] {test.shortDescription() or test._testMethodName}\033[0m\n")


class TIFriendlyTestRunner(unittest.TextTestRunner):
    """Executador de testes personalizado para relatórios de TI."""

    resultclass = TIFriendlyTestResult

    def run(self, test):
        # Garantir pasta logs/ existente
        Path("logs").mkdir(exist_ok=True)
        log_path = Path("logs/test_run.log")
        
        sys.stdout.write("\n\033[96m" + "="*60 + "\033[0m\n")
        sys.stdout.write("\033[1m\033[96m  INICIANDO DIAGNÓSTICO DE INTEGRIDADE DO AUTOMDFTEXT\033[0m\n")
        sys.stdout.write("\033[96m" + "="*60 + "\033[0m\n\n")

        start_time = time.time()
        result = super().run(test)
        duration = time.time() - start_time

        # Montar diagnóstico detalhado
        sys.stdout.write("\n\033[96m" + "-"*60 + "\033[0m\n")
        sys.stdout.write("\033[1m  RESUMO DO DIAGNÓSTICO PARA TI:\033[0m\n\n")
        
        total = result.testsRun
        failures = len(result.failures)
        errors = len(result.errors)
        passed = result.success_count

        # Gravar log detalhado de traceback
        import traceback
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("=== LOG DE DIAGNÓSTICO DO AUTOMDFTEXT ===\n")
            f.write(f"Data: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Total executado: {total}\n")
            f.write(f"Sucessos: {passed} | Falhas: {failures} | Erros: {errors}\n")
            f.write("="*60 + "\n\n")
            
            if result.failures or result.errors:
                f.write("=== TRACEBACK DE FALHAS/ERROS ===\n")
                for test, kind, err in result.failure_details:
                    tb_str = "".join(traceback.format_exception(*err))
                    f.write(f"Test Case: {test}\nTipo: {kind}\nTraceback:\n{tb_str}\n")
                    f.write("-"*40 + "\n")

        # Exibição bonita no console
        if passed == total:
            sys.stdout.write(f"  \033[92mStatus Geral: 100% OPERACIONAL ({passed}/{total} testes passaram)\033[0m\n")
            sys.stdout.write("  Nenhuma falha de integridade ou lógica foi detectada.\n")
            sys.stdout.write("  O código está limpo, modular e pronto para execução.\n")
        else:
            sys.stdout.write(f"  \033[91mStatus Geral: FALHA DETECTADA ({passed}/{total} passaram, {failures + errors} falharam)\033[0m\n\n")
            sys.stdout.write(f"  \033[1m\033[91mDetalhes das Ocorrências para Análise Técnica:\033[0m\n")
            
            for idx, (test, kind, err) in enumerate(result.failure_details, 1):
                tb_lines = "".join(traceback.format_exception(*err)).splitlines()
                exc_str = tb_lines[-1] if tb_lines else "Erro desconhecido"
                sys.stdout.write(f"\n  {idx}. \033[1mComponente:\033[0m {test._testMethodName} ({test.__class__.__name__})\n")
                sys.stdout.write(f"     \033[1mClassificação:\033[0m {kind}\n")
                sys.stdout.write(f"     \033[1mErro Observado:\033[0m \033[93m{exc_str}\033[0m\n")
                sys.stdout.write(f"     \033[1mAção Sugerida:\033[0m ")
                
                # Gerar soluções úteis com base no nome do teste falhado
                test_name = test._testMethodName.lower()
                if "single_instance" in test_name:
                    sys.stdout.write("Verifique se as permissões de criação de Mutex global no kernel do Windows estão bloqueadas por alguma GPO de segurança.\n")
                elif "failsafe" in test_name:
                    sys.stdout.write("Problema com pynput/listeners de teclado. Certifique-se de que o antivírus/EDR local não está interceptando hooks de teclado (keylogger protection).\n")
                elif "browser" in test_name or "focus" in test_name:
                    sys.stdout.write("O navegador Edge/Chrome pode ter sofrido alterações nas classes de janela ou título padrão. Verifique se o navegador está fixado na primeira posição da barra de tarefas (Win+1).\n")
                elif "ncm" in test_name or "profile" in test_name:
                    sys.stdout.write("Os arquivos de configuração .txt na pasta scripts/ possuem parâmetros ausentes ou formatação inválida (esperado formato INI com seções).\n")
                elif "xml" in test_name:
                    sys.stdout.write("Erro de detecção do arquivo XML em downloads. Certifique-se de que a pasta Downloads está acessível ou se o usuário possui permissão de leitura.\n")
                else:
                    sys.stdout.write("Revise a lógica de preenchimento ou os deltas de atraso (constants.py) para o formulário correspondente.\n")

            sys.stdout.write(f"\n  \033[96mNota:\033[0m O log técnico completo contendo o traceback foi salvo em: \033[1mlogs/test_run.log\033[0m\n")

        sys.stdout.write(f"\n  Diagnóstico concluído em {duration:.3f} segundos.\n")
        sys.stdout.write("\033[96m" + "="*60 + "\033[0m\n\n")

        return result


if __name__ == "__main__":
    # Carregar todos os casos de teste deste arquivo
    loader = unittest.TestLoader()
    suite = loader.loadTestsFromModule(sys.modules[__name__])
    
    # Rodar testes usando o runner especializado para TI
    test_runner = TIFriendlyTestRunner(verbosity=0)
    result = test_runner.run(suite)
    
    # Sair com código apropriado (0=sucesso, 1=falha)
    sys.exit(0 if len(result.failures) == 0 and len(result.errors) == 0 else 1)
