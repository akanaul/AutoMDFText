# Extensão CIOT - Guia de Uso

## O que é CIOT?
CIOT (Conhecimento de Transporte Intermodal Operacional) é um campo adicional que pode ser preenchido após a automação principal do MDF-e.

## Como usar

### Fluxo Automático
1. **Execute o script principal**: `modular_mdfe.py` (via `run.bat` opção 1)
2. **Ao finalizar a automação MDF-e**, você verá a seguinte pergunta:
   ```
   Deseja preencher o campo CIOT (Conhecimento de Transporte Intermodal Operacional) agora?
   ```
3. **Clique em "Sim, preencher CIOT"** para abrir a extensão
4. **Informe o número do CIOT** na caixa de diálogo que aparecerá
5. O script preencherá o campo automaticamente e retornará

### Fluxo Manual (Independente)
Se necessário, você pode executar o preenchimento de CIOT independentemente:

```bash
python ciot_filler.py
```

Ou através do `run.bat` (quando adicionada a opção no menu).

## Arquivos da Extensão

- **`ciot_filler.py`**: Script desacoplado que preenche o campo CIOT
- **Logs**: Registrados em `logs/ciot_filler_YYYYMMDD_HHMMSS.log`

## Característica Técnica

A extensão é **completamente desacoplada** do script principal:
- Executa como um **processo separado** (subprocess)
- Usa a **mesma versão do Python** da máquina (mesmo venv/conda)
- **Pode ser executada independentemente** a qualquer momento
- Mantém seus **próprios logs** separados

## Próximas Melhorias (Opcional)

- Adicionar opção no menu principal do `run.bat`
- Integrar CIOT no perfil de configuração (scripts/*.txt)
- Adicionar validação do formato do CIOT
- Criar suporte para múltiplos CIOTs em uma mesma sessão

## Troubleshooting

### "Script CIOT não encontrado"
- Certifique-se de que `ciot_filler.py` está no mesmo diretório que `modular_mdfe.py`
- Verifique se não houve renomeação ou movimentação do arquivo

### "Erro ao executar a extensão"
- Verifique se as dependências Python estão instaladas (pyautogui, pyperclip, Pillow)
- Consulte o log em `logs/ciot_filler_*.log` para detalhes completos

### Campo CIOT não encontrado
- Verifique se você está na página correta
- O script procura por "CIOT" usando Ctrl+F - certifique-se de que essa página contém esse campo
- Ajuste a lógica em `ciot_filler.py` se necessário
