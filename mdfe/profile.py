"""
Leitura e gerenciamento de perfis de configuração para AutoMDFText.

Responsabilidades:
- Parsear arquivos .txt de perfil em dicionário de seções (parse_profile)
- Representar um perfil carregado com hot-reload automático (ConfigProfile)
- Listar os perfis disponíveis em scripts/ (list_profiles)

Os perfis seguem a estrutura INI com seções como [MDFE], [MODAL_RODOVIARIO] etc.
"""

from pathlib import Path

# Diretório de perfis relativo à raiz do projeto
_BASE_DIR  = Path(__file__).parent.parent
CONFIG_DIR = _BASE_DIR / "scripts"
CONFIG_DIR.mkdir(exist_ok=True)


def parse_profile(path: Path) -> dict[str, dict[str, str]]:
    """Lê um arquivo .txt de perfil e organiza em dicionário de seções."""
    sections: dict[str, dict[str, str]] = {}
    current_section = "GENERAL"
    sections[current_section] = {}

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("[") and "]" in line:
            current_section = line[1:line.index("]")].strip().upper()
            sections.setdefault(current_section, {})
            continue
        if "=" in line:
            key, value = line.split("=", 1)
            sections[current_section][key.strip().lower()] = value.strip()

    return sections


class ConfigProfile:
    """Perfil de configuração com detecção automática de mudança no arquivo (hot-reload)."""

    def __init__(self, path: Path):
        self.path = path
        self._data: dict[str, dict[str, str]] = {}
        self._mtime = 0.0
        self.reload()

    def reload(self) -> None:
        self._data = parse_profile(self.path)
        try:
            self._mtime = self.path.stat().st_mtime
        except OSError:
            self._mtime = 0.0

    def ensure_current(self) -> None:
        try:
            current_mtime = self.path.stat().st_mtime
        except OSError:
            current_mtime = 0.0
        if current_mtime != self._mtime:
            self.reload()

    def get_value(self, section: str, key: str, default: str = "") -> str:
        self.ensure_current()
        return self._data.get(section.upper(), {}).get(key.lower(), default)


def list_profiles() -> list[str]:
    """Retorna os nomes de todos os arquivos .txt em scripts/, ordenados."""
    return sorted(p.name for p in CONFIG_DIR.glob("*.txt"))
