"""Pacote de etapas de execução da automação MDF-e."""

from mdfe.steps.navigate import navigate_to_mdfe
from mdfe.steps.fill_mdfe import fill_mdfe
from mdfe.steps.fill_modal_rodo import fill_modal_rodo
from mdfe.steps.fill_additional_info import fill_additional_info
from mdfe.steps.averbacao import perform_averbacao

__all__ = [
    "navigate_to_mdfe",
    "fill_mdfe",
    "fill_modal_rodo",
    "fill_additional_info",
    "perform_averbacao",
]
