from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
import csv
import re
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import load_workbook


@dataclass
class CombinationRow:
    robot_id: str
    name: str
    analysis_type: str
    combination_type: str
    nature: str
    definition: str


@dataclass
class ReactionRow:
    node: int
    case_id: int
    is_combination: bool
    fx: float
    fy: float
    fz: float
    mx: float
    my: float
    mz: float
    raw_key: str


@dataclass
class MemberResult:
    member: str
    raw_member: str
    profile: str
    material: str
    lay: Optional[float]
    laz: Optional[float]
    ratio_elu: Optional[float]
    case_elu: str
    ratio_uy: Optional[float]
    case_uy: str
    ratio_uz: Optional[float]
    case_uz: str
    ratio_vx: Optional[float]
    case_vx: str
    ratio_vy: Optional[float]
    case_vy: str
    auto_named: bool = False


@dataclass
class CaseMapping:
    case_id: int
    load_name: str
    abbreviation: str
    include_reactions: bool = True
    group: str = ""


@dataclass
class ValidationItem:
    level: str  # info | warning | error | success
    message: str


@dataclass
class ProcessedData:
    combinations_elu: List[List[object]]
    combinations_els: List[List[object]]
    table8: List[List[object]]
    table9: List[List[object]]
    table10: List[List[object]]
    validations: List[ValidationItem]
    summary: Dict[str, object]
    raw_combinations: List[CombinationRow]
    raw_reactions: List[ReactionRow]
    raw_members: List[MemberResult]


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------


def _decode_robot_text(data: bytes) -> str:
    """Decode Robot CSV bytes robustly.

    Robot commonly exports UTF-16 LE with BOM. Fallbacks are included for files
    resaved by Excel/Notepad.
    """
    encodings = ["utf-16", "utf-8-sig", "cp1252", "latin1"]
    last_err = None
    for enc in encodings:
        try:
            text = data.decode(enc)
            # A successful utf-16 decode of a non-utf16 file often produces many NULs.
            if "\x00" in text and enc != "utf-16":
                continue
            return text
        except UnicodeError as exc:
            last_err = exc
    raise ValueError(f"Não foi possível decodificar o CSV. Último erro: {last_err}")


def _read_semicolon_csv(data: bytes) -> List[List[str]]:
    text = _decode_robot_text(data)
    return [row for row in csv.reader(text.splitlines(), delimiter=";") if any(c.strip() for c in row)]


def _norm(s: object) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s).strip()).casefold()


def _header_index(headers: Sequence[str], candidates: Sequence[str], fallback: Optional[int] = None) -> int:
    norm_headers = [_norm(h) for h in headers]
    for cand in candidates:
        nc = _norm(cand)
        for i, h in enumerate(norm_headers):
            if h == nc or nc in h:
                return i
    if fallback is not None and fallback < len(headers):
        return fallback
    raise ValueError(f"Coluna não encontrada. Esperado algo como: {', '.join(candidates)}")


def _to_float(value: object) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s or s == "-":
        return None
    s = s.replace(" ", "")
    # Robot exports decimals using comma in CSV, but copied XLSX can use dot-as-text.
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    elif "," in s and "." in s:
        # Assume pt-BR thousands + decimal comma only when comma appears after dot.
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def _clean_case_label(value: object) -> str:
    s = "" if value is None else str(value).strip()
    if not s or s == "-":
        return "-"
    m = re.match(r"^\s*\d+\s+(.+)$", s)
    return m.group(1).strip() if m else s


def _clean_member(value: object) -> Tuple[str, bool]:
    s = "" if value is None else str(value).strip()
    if not s:
        return "", False

    # Normal case: '17  Long. Escada_7' -> 'Long. Escada_7'
    m = re.match(r"^(\d+)\s+(.+)$", s)
    if m and m.group(2).strip():
        return m.group(2).strip(), False

    # Robot can copy an unnamed member as a plain number.
    if re.fullmatch(r"\d+(?:\.0+)?", s):
        number = int(float(s))
        return f"Barra_{number}", True

    return s, False


def _clean_material(value: object) -> str:
    s = "" if value is None else str(value).strip()
    mapping = {
        "STEEL A36": "ASTM A36",
        "STEEL A572-50": "ASTM A572 Gr. 50",
        "STEEL A572-5": "ASTM A572 Gr. 50",
        "STEEL A572-50 ": "ASTM A572 Gr. 50",
    }
    return mapping.get(s, s)


def _clean_profile(value: object) -> str:
    s = "" if value is None else str(value).strip()
    # Preserve Robot nomenclature, but use Brazilian decimal comma for presentation.
    return s.replace(".", ",")


def _format_coef(value: str) -> str:
    try:
        x = float(value)
    except ValueError:
        return value.replace(".", ",")
    return f"{x:.2f}".replace(".", ",")


# -----------------------------------------------------------------------------
# Parsers
# -----------------------------------------------------------------------------


def parse_combinations_csv(data: bytes) -> List[CombinationRow]:
    rows = _read_semicolon_csv(data)
    if len(rows) < 2:
        raise ValueError("CSV de combinações sem linhas de dados.")

    headers = rows[0]
    idx_id = _header_index(headers, ["Combinações", "Combinação"], 0)
    idx_name = _header_index(headers, ["Nome"], 1)
    idx_analysis = _header_index(headers, ["Tipo de análise", "Análise"], 2)
    idx_type = _header_index(headers, ["Tipo de combinação", "Combinação"], 3)
    idx_nature = _header_index(headers, ["Natureza do caso", "Natureza"], 4)
    idx_def = _header_index(headers, ["Definição", "Definition"], 5)

    out: List[CombinationRow] = []
    for r in rows[1:]:
        r = list(r) + [""] * (max(idx_id, idx_name, idx_analysis, idx_type, idx_nature, idx_def) + 1 - len(r))
        definition = r[idx_def].strip()
        name = r[idx_name].strip()
        if not definition and not name:
            continue
        out.append(
            CombinationRow(
                robot_id=r[idx_id].strip(),
                name=name,
                analysis_type=r[idx_analysis].strip(),
                combination_type=r[idx_type].strip(),
                nature=r[idx_nature].strip(),
                definition=definition,
            )
        )
    if not out:
        raise ValueError("Nenhuma combinação foi reconhecida no CSV.")
    return out


def parse_reactions_csv(data: bytes) -> List[ReactionRow]:
    rows = _read_semicolon_csv(data)
    if len(rows) < 2:
        raise ValueError("CSV de reações sem linhas de dados.")

    headers = rows[0]
    idx_key = _header_index(headers, ["Nó/Caso", "No/Caso", "Node/Case"], 0)
    idx_fx = _header_index(headers, ["FX"], 1)
    idx_fy = _header_index(headers, ["FY"], 2)
    idx_fz = _header_index(headers, ["FZ"], 3)
    idx_mx = _header_index(headers, ["MX"], 4)
    idx_my = _header_index(headers, ["MY"], 5)
    idx_mz = _header_index(headers, ["MZ"], 6)

    out: List[ReactionRow] = []
    key_re = re.compile(r"^\s*(\d+)\s*/\s*(\d+)\s*(\(C\))?\s*$", re.I)
    for r in rows[1:]:
        r = list(r) + [""] * (max(idx_key, idx_fx, idx_fy, idx_fz, idx_mx, idx_my, idx_mz) + 1 - len(r))
        key = r[idx_key].strip()
        m = key_re.match(key)
        if not m:
            continue
        vals = [_to_float(r[i]) for i in (idx_fx, idx_fy, idx_fz, idx_mx, idx_my, idx_mz)]
        vals = [0.0 if v is None else v for v in vals]
        out.append(
            ReactionRow(
                node=int(m.group(1)),
                case_id=int(m.group(2)),
                is_combination=bool(m.group(3)),
                fx=vals[0], fy=vals[1], fz=vals[2], mx=vals[3], my=vals[4], mz=vals[5],
                raw_key=key,
            )
        )
    if not out:
        raise ValueError("Nenhuma reação no formato 'Nó/Caso' foi reconhecida.")
    return out


def parse_member_results_xlsx(data: bytes) -> List[MemberResult]:
    try:
        wb = load_workbook(BytesIO(data), data_only=True, read_only=True)
    except Exception as exc:
        raise ValueError(f"Não foi possível abrir o Excel ELU/ELS: {exc}") from exc

    ws = wb[wb.sheetnames[0]]
    rows: List[List[object]] = []
    for row in ws.iter_rows(values_only=True):
        vals = list(row)
        if any(v not in (None, "") for v in vals):
            rows.append(vals)
    if not rows:
        raise ValueError("O Excel ELU/ELS está vazio.")

    # Header is optional. Detect and skip if the user pasted it as well.
    first_norm = [_norm(v) for v in rows[0]]
    has_header = any("membro" in x for x in first_norm) and any("seção" in x or "secao" in x for x in first_norm)
    data_rows = rows[1:] if has_header else rows

    out: List[MemberResult] = []
    for row in data_rows:
        row = list(row) + [None] * (16 - len(row))
        if len(row) < 16:
            continue
        if not any(v not in (None, "") for v in row[:16]):
            continue

        member, auto_named = _clean_member(row[0])
        if not member:
            continue

        out.append(
            MemberResult(
                member=member,
                raw_member="" if row[0] is None else str(row[0]),
                profile=_clean_profile(row[2]),
                material=_clean_material(row[3]),
                lay=_to_float(row[4]),
                laz=_to_float(row[5]),
                ratio_elu=_to_float(row[6]),
                case_elu=_clean_case_label(row[7]),
                ratio_uy=_to_float(row[8]),
                case_uy=_clean_case_label(row[9]),
                ratio_uz=_to_float(row[10]),
                case_uz=_clean_case_label(row[11]),
                ratio_vx=_to_float(row[12]),
                case_vx=_clean_case_label(row[13]),
                ratio_vy=_to_float(row[14]),
                case_vy=_clean_case_label(row[15]),
                auto_named=auto_named,
            )
        )
    if not out:
        raise ValueError(
            "Nenhuma linha ELU/ELS foi reconhecida. O arquivo deve conter as 16 colunas copiadas da janela "
            "'Verificação de membro (ELS; ELU)' do Robot."
        )
    return out


# -----------------------------------------------------------------------------
# Combination translation / case discovery
# -----------------------------------------------------------------------------


def _split_top_level_plus(expr: str) -> List[str]:
    parts: List[str] = []
    buf: List[str] = []
    depth = 0
    for ch in expr:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth = max(0, depth - 1)
        if ch == "+" and depth == 0:
            part = "".join(buf).strip()
            if part:
                parts.append(part)
            buf = []
        else:
            buf.append(ch)
    part = "".join(buf).strip()
    if part:
        parts.append(part)
    return parts


def _parse_combination_terms(definition: str) -> List[Tuple[List[int], str]]:
    """Return [([case ids], factor_as_text), ...].

    Supports Robot's typical syntax, e.g. 1*1.25+(2+3)*1.50+5*0.72.
    Raises ValueError instead of silently mis-translating unsupported expressions.
    """
    expr = re.sub(r"\s+", "", definition)
    if not expr:
        return []

    terms: List[Tuple[List[int], str]] = []
    for raw in _split_top_level_plus(expr):
        m_group = re.fullmatch(r"\((\d+(?:\+\d+)*)\)\*([+-]?\d+(?:\.\d+)?)", raw)
        if m_group:
            ids = [int(x) for x in m_group.group(1).split("+")]
            terms.append((ids, m_group.group(2)))
            continue

        m_single = re.fullmatch(r"(\d+)\*([+-]?\d+(?:\.\d+)?)", raw)
        if m_single:
            terms.append(([int(m_single.group(1))], m_single.group(2)))
            continue

        # Accept an un-factored case as 1.00.
        m_plain = re.fullmatch(r"(\d+)", raw)
        if m_plain:
            terms.append(([int(m_plain.group(1))], "1.00"))
            continue

        raise ValueError(f"Termo de combinação não reconhecido: '{raw}' em '{definition}'")
    return terms


def combination_case_ids(definition: str) -> List[int]:
    ids: List[int] = []
    try:
        for group, _factor in _parse_combination_terms(definition):
            ids.extend(group)
    except ValueError:
        # Conservative fallback: only integers directly followed by '*', plus integers inside (...) before '*'.
        # Used for discovery only; translation will still warn/fail explicitly.
        for m in re.finditer(r"(?<![\d.])(\d+)\s*\*", definition):
            ids.append(int(m.group(1)))
        for m in re.finditer(r"\((\d+(?:\s*\+\s*\d+)*)\)\s*\*", definition):
            ids.extend(int(x.strip()) for x in m.group(1).split("+"))
    return sorted(set(ids))


def detect_basic_case_ids(combinations: Sequence[CombinationRow], reactions: Sequence[ReactionRow]) -> List[int]:
    ids = {r.case_id for r in reactions if not r.is_combination}
    for c in combinations:
        ids.update(combination_case_ids(c.definition))
    return sorted(ids)


def translate_combination(definition: str, mapping: Dict[int, CaseMapping]) -> str:
    out: List[str] = []
    for case_ids, factor in _parse_combination_terms(definition):
        coef = _format_coef(factor)
        for cid in case_ids:
            item = mapping.get(cid)
            abbr = item.abbreviation.strip() if item and item.abbreviation.strip() else f"C{cid}"
            out.append(f"{coef}{abbr}")
    return " + ".join(out)


# -----------------------------------------------------------------------------
# Processor
# -----------------------------------------------------------------------------


def process_data(
    combinations: Sequence[CombinationRow],
    reactions: Sequence[ReactionRow],
    members: Sequence[MemberResult],
    case_mappings: Sequence[CaseMapping],
    structure_name: str,
    lambda_limit: float = 200.0,
    force_conversion_factor: float = 0.01,
    moment_conversion_factor: Optional[float] = None,
) -> ProcessedData:
    if moment_conversion_factor is None:
        moment_conversion_factor = force_conversion_factor

    mapping = {m.case_id: m for m in case_mappings}
    validations: List[ValidationItem] = []

    detected = detect_basic_case_ids(combinations, reactions)
    unmapped = [cid for cid in detected if cid not in mapping or not mapping[cid].abbreviation.strip()]
    if unmapped:
        validations.append(ValidationItem("error", f"Casos sem abreviação definida: {', '.join(map(str, unmapped))}."))

    abbreviations = [m.abbreviation.strip().casefold() for m in case_mappings if m.abbreviation.strip()]
    duplicates = sorted({a for a in abbreviations if abbreviations.count(a) > 1})
    if duplicates:
        validations.append(ValidationItem("warning", "Existem abreviações duplicadas no mapeamento de casos."))

    comb_elu: List[List[object]] = []
    comb_els: List[List[object]] = []
    failed_translations = 0
    for c in combinations:
        try:
            expr = translate_combination(c.definition, mapping)
        except ValueError as exc:
            expr = c.definition
            failed_translations += 1
            validations.append(ValidationItem("warning", f"{c.name}: {exc}. Mantida a definição original do Robot."))

        row = [c.name or c.robot_id, expr, c.definition]
        ctype = _norm(c.combination_type + " " + c.name)
        if "elu" in ctype or c.name.upper().startswith("ELU"):
            comb_elu.append(row)
        elif "els" in ctype or c.name.upper().startswith("ELS"):
            comb_els.append(row)
        else:
            validations.append(ValidationItem("warning", f"Combinação '{c.name}' não classificada como ELU ou ELS e foi omitida das tabelas finais."))

    table8: List[List[object]] = []
    elu_fail = 0
    slender_fail = 0
    auto_named = 0
    for m in members:
        if m.auto_named:
            auto_named += 1
        max_l = max([x for x in (m.lay, m.laz) if x is not None], default=0.0)
        status_elu = "OK" if m.ratio_elu is not None and m.ratio_elu <= 1.0 else "NÃO OK"
        status_slender = "OK" if max_l <= lambda_limit else "NÃO OK"
        if status_elu != "OK":
            elu_fail += 1
        if status_slender != "OK":
            slender_fail += 1
        table8.append([
            m.member,
            m.profile,
            m.material,
            m.lay,
            m.laz,
            m.ratio_elu,
            m.case_elu,
            status_elu,
            status_slender,
        ])

    table9: List[List[object]] = []
    els_fail = 0
    for m in members:
        ratios = [m.ratio_uy, m.ratio_uz, m.ratio_vx, m.ratio_vy]
        numeric = [v for v in ratios if v is not None]
        if not numeric:
            continue
        status = "OK" if max(numeric) <= 1.0 else "NÃO OK"
        if status != "OK":
            els_fail += 1
        table9.append([
            m.member,
            m.profile,
            m.ratio_uy if m.ratio_uy is not None else "-", m.case_uy,
            m.ratio_uz if m.ratio_uz is not None else "-", m.case_uz,
            m.ratio_vx if m.ratio_vx is not None else "-", m.case_vx,
            m.ratio_vy if m.ratio_vy is not None else "-", m.case_vy,
            status,
        ])

    table10: List[List[object]] = []
    included_cases = {m.case_id for m in case_mappings if m.include_reactions}
    for r in reactions:
        if r.is_combination or r.case_id not in included_cases:
            continue
        mp = mapping.get(r.case_id)
        load_name = mp.load_name.strip() if mp and mp.load_name.strip() else f"CASO {r.case_id}"
        group = mp.group.strip() if mp and str(mp.group).strip() else str(r.case_id)
        table10.append([
            r.node,
            group,
            load_name,
            r.fx * force_conversion_factor,
            r.fy * force_conversion_factor,
            r.fz * force_conversion_factor,
            r.mx * moment_conversion_factor,
            r.my * moment_conversion_factor,
            r.mz * moment_conversion_factor,
        ])
    def _group_sort_key(value: object):
        text = str(value).strip()
        try:
            return (0, float(text))
        except ValueError:
            return (1, text.casefold())

    table10.sort(key=lambda x: (x[0], _group_sort_key(x[1]), x[2].casefold()))

    nodes = sorted({r.node for r in reactions if not r.is_combination})
    if not comb_elu:
        validations.append(ValidationItem("error", "Nenhuma combinação ELU foi identificada."))
    else:
        validations.append(ValidationItem("success", f"{len(comb_elu)} combinações ELU identificadas."))
    if not comb_els:
        validations.append(ValidationItem("warning", "Nenhuma combinação ELS foi identificada."))
    else:
        validations.append(ValidationItem("success", f"{len(comb_els)} combinações ELS identificadas."))

    validations.append(ValidationItem("success", f"{len(members)} membros lidos na verificação ELU/ELS."))
    validations.append(ValidationItem("success", f"{len(table9)} membros possuem resultados ELS."))
    validations.append(ValidationItem("success", f"{len(nodes)} nós de apoio identificados no CSV de reações."))

    if auto_named:
        validations.append(ValidationItem("warning", f"{auto_named} membro(s) sem nome no Robot foram renomeados automaticamente como 'Barra_N'."))
    if elu_fail:
        validations.append(ValidationItem("warning", f"{elu_fail} membro(s) com índice ELU > 1,00 ou sem índice numérico."))
    if slender_fail:
        validations.append(ValidationItem("warning", f"{slender_fail} membro(s) excedem λ = {lambda_limit:g}."))
    if els_fail:
        validations.append(ValidationItem("warning", f"{els_fail} membro(s) com razão ELS > 1,00."))
    if failed_translations == 0:
        validations.append(ValidationItem("success", "Todas as definições de combinação foram traduzidas para as abreviações informadas."))

    summary = {
        "structure_name": structure_name,
        "n_combinations_elu": len(comb_elu),
        "n_combinations_els": len(comb_els),
        "n_members": len(members),
        "n_members_els": len(table9),
        "n_support_nodes": len(nodes),
        "support_nodes": nodes,
        "lambda_limit": lambda_limit,
        "force_conversion_factor": force_conversion_factor,
        "moment_conversion_factor": moment_conversion_factor,
    }

    return ProcessedData(
        combinations_elu=comb_elu,
        combinations_els=comb_els,
        table8=table8,
        table9=table9,
        table10=table10,
        validations=validations,
        summary=summary,
        raw_combinations=list(combinations),
        raw_reactions=list(reactions),
        raw_members=list(members),
    )
