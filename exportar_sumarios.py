import json
import os
import re
import sys
import zipfile
from pathlib import Path
from typing import Dict, List, Tuple
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

# Namespaces utilizados nos ficheiros DOCX
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DEFAULT_CONFIG_PATH = Path("config_exportar_sumarios.json")


def _collect_table_text(doc_root: ET.Element) -> List[List[List[str]]]:
    """Extrai o texto de todas as tabelas do documento.

    Returns:
        Lista de tabelas, onde cada tabela é uma lista de linhas e cada linha é
        uma lista com o texto de cada célula.
    """

    tables: List[List[List[str]]] = []

    def get_text(element: ET.Element) -> str:
        parts: List[str] = []

        for node in element.iter():
            tag = node.tag

            if tag == f"{{{WORD_NAMESPACE}}}t":
                parts.append(node.text or "")
            elif tag in {f"{{{WORD_NAMESPACE}}}br", f"{{{WORD_NAMESPACE}}}cr"}:
                parts.append("\n")

            if node.tail:
                parts.append(node.tail)

        return "".join(parts)

    for tbl in doc_root.iter(f"{{{WORD_NAMESPACE}}}tbl"):
        rows: List[List[str]] = []
        for tr in tbl.iter(f"{{{WORD_NAMESPACE}}}tr"):
            row_cells: List[str] = []
            for tc in tr.iter(f"{{{WORD_NAMESPACE}}}tc"):
                row_cells.append(get_text(tc).strip())
            rows.append(row_cells)
        tables.append(rows)
    return tables


def _extract_turma(header_cells: List[str]) -> str:
    """Obtém o nome da turma a partir da primeira célula da tabela inicial."""
    if not header_cells:
        return ""

    match = re.match(r"([\wÁ-Úá-úçÇ]+)", header_cells[0])
    return match.group(1) if match else header_cells[0]


def _extract_data(cells: List[str]) -> str:
    """Procura a data no texto das células fornecidas."""
    joined = " ".join(cells)
    match = re.search(r"(\d{2}/\d{2}/\d{4})", joined)
    return match.group(1) if match else ""


def _extract_sumario(tables: List[List[List[str]]]) -> str:
    """Localiza a tabela de sumário e devolve o conteúdo agregado."""
    for table in tables:
        if not table:
            continue

        header_text = " ".join(table[0]).lower()
        if "sumário" in header_text or "sumario" in header_text:
            content_rows = table[1:]
            contents: List[str] = []
            for row in content_rows:
                row_text = " ".join(part for part in row if part).strip()
                if row_text:
                    contents.append(row_text)
            return "\n".join(contents)
    return ""


def parse_docx(path: Path) -> Tuple[str, str, str]:
    """Lê um ficheiro DOCX e devolve (turma, data, sumário)."""
    with zipfile.ZipFile(path) as docx_zip:
        xml = docx_zip.read("word/document.xml")

    root = ET.fromstring(xml)
    tables = _collect_table_text(root)

    turma = _extract_turma(tables[0][0] if tables and tables[0] else [])
    data = _extract_data(tables[0][0] if tables and tables[0] else [])

    if not data:
        data = _extract_data([cell for table in tables for row in table for cell in row])

    sumario = _extract_sumario(tables)
    return turma, data, sumario


def _column_letter(index: int) -> str:
    """Converte um índice (0-based) para a letra da coluna do Excel."""
    letters = ""
    index += 1  # Excel usa base 1
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _build_sheet_xml(rows: List[List[str]]) -> bytes:
    """Cria o XML da folha de cálculo com os dados fornecidos."""

    def format_cell_value(value: str) -> str:
        """Escapa texto e converte quebras de linha para o formato esperado pelo Excel."""

        escaped = escape(value)
        return escaped.replace("\n", "&#10;")
    sheet_lines = [
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        "  <sheetData>",
    ]

    for row_idx, row in enumerate(rows, start=1):
        sheet_lines.append(f"    <row r=\"{row_idx}\">")
        for col_idx, value in enumerate(row):
            cell_ref = f"{_column_letter(col_idx)}{row_idx}"
            cell_value = format_cell_value(value)
            sheet_lines.append(
                f"      <c r=\"{cell_ref}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">{cell_value}</t></is></c>"
            )
        sheet_lines.append("    </row>")
    sheet_lines.append("  </sheetData>")
    sheet_lines.append("</worksheet>")

    return "\n".join(sheet_lines).encode("utf-8")


def write_xlsx(rows: List[List[str]], output_path: Path) -> None:
    """Gera um ficheiro XLSX mínimo apenas com os dados fornecidos."""
    sheet_xml = _build_sheet_xml(rows)

    workbook_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <sheets>
    <sheet name=\"Sumarios\" sheetId=\"1\" r:id=\"rId1\"/>
  </sheets>
</workbook>""".encode("utf-8")

    workbook_rels = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>
</Relationships>""".encode("utf-8")

    root_rels = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>
</Relationships>""".encode("utf-8")

    content_types = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
</Types>""".encode("utf-8")

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as xlsx:
        xlsx.writestr("[Content_Types].xml", content_types)
        xlsx.writestr("_rels/.rels", root_rels)
        xlsx.writestr("xl/workbook.xml", workbook_xml)
        xlsx.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        xlsx.writestr("xl/worksheets/sheet1.xml", sheet_xml)



def process_directory(base_dir: Path, output_file: Path) -> None:
    """Percorre subpastas, extrai dados dos DOCX e cria o XLSX final."""
    rows: List[List[str]] = [["Ficheiro", "Turma", "Data", "Sumário"]]

    for root, _, files in os.walk(base_dir):
        for filename in files:
            if not filename.lower().endswith(".docx"):
                continue

            path = Path(root) / filename
            try:
                turma, data, sumario = parse_docx(path)
            except Exception as exc:  # noqa: BLE001
                turma = data = sumario = f"Erro ao ler: {exc}"
            rows.append([filename, turma, data, sumario])

    write_xlsx(rows, output_file)
    print(f"Ficheiro gerado: {output_file}")


def _load_paths_from_config(config_path: Path) -> Tuple[Path, Path]:
    """Lê o ficheiro JSON de configuração e devolve diretórios de entrada e saída."""

    if not config_path.exists():
        raise FileNotFoundError(
            f"Configuração não encontrada em {config_path}. Crie o ficheiro de acordo com o formato inicial."
        )

    with config_path.open("r", encoding="utf-8") as config_file:
        config: Dict[str, str] = json.load(config_file)

    docs_dir = config.get("docs_dir")
    output_xlsx = config.get("output_xlsx")

    if not isinstance(docs_dir, str) or not isinstance(output_xlsx, str):
        raise ValueError(
            "Configuração inválida: as chaves 'docs_dir' e 'output_xlsx' são obrigatórias e devem ser strings."
        )

    return Path(docs_dir), Path(output_xlsx)


def main() -> None:
    config_path = Path(os.environ.get("EXPORTAR_SUMARIOS_CONFIG", DEFAULT_CONFIG_PATH))

    try:
        base_dir, output_file = _load_paths_from_config(config_path)
    except Exception as exc:  # noqa: BLE001
        print(f"Erro ao ler configuração: {exc}")
        sys.exit(1)

    process_directory(base_dir, output_file)


if __name__ == "__main__":
    main()
