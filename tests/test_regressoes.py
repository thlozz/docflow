"""Regressões de CNPJ e leitura de planilhas."""
from pathlib import Path

import pytest
from openpyxl import Workbook

from docflow.core import extract_fields, match_by_spreadsheet, valid_cnpj


def test_cnpj_valido():
    assert valid_cnpj("11.222.333/0001-81")
    assert valid_cnpj("11222333000181")


def test_cnpj_protocolo_invalido():
    assert not valid_cnpj("20261020123456")
    assert not valid_cnpj("00000000000000")


def test_extracao_nao_confunde_protocolo_com_cnpj(monkeypatch):
    monkeypatch.setattr("docflow.core.read_pdf", lambda _: "Protocolo 20261020123456")
    assert extract_fields(Path("ficticio.pdf"))["cnpj"] is None


def test_extracao_ignora_invalido_e_aceita_valido(monkeypatch):
    monkeypatch.setattr("docflow.core.read_pdf", lambda _: "Protocolo 20261020123456 CNPJ 11.222.333/0001-81")
    assert extract_fields(Path("ficticio.pdf"))["cnpj"] == "11.222.333/0001-81"


def salvar(tmp_path, rows):
    wb = Workbook()
    for row in rows:
        wb.active.append(row)
    path = tmp_path / "entrada.xlsx"
    wb.save(path)
    wb.close()
    return path


def test_planilha_sem_linhas(tmp_path):
    path = salvar(tmp_path, [])
    with pytest.raises(ValueError, match="vazia"):
        match_by_spreadsheet(path, [], [], tmp_path / "saida")


def test_planilha_so_cabecalho(tmp_path):
    path = salvar(tmp_path, [["NF"]])
    assert match_by_spreadsheet(path, [], [], tmp_path / "saida") == []


def test_coluna_inexistente(tmp_path):
    path = salvar(tmp_path, [["Documento"], [1001]])
    with pytest.raises(ValueError, match="Coluna"):
        match_by_spreadsheet(path, [], [], tmp_path / "saida")
