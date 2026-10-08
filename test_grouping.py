from core import CaseMapping, parse_combinations_csv, parse_member_results_xlsx, parse_reactions_csv, process_data


def test_case_mapping_keeps_editable_group():
    m = CaseMapping(5, "TEMPERATURA -20°C", "T-°C", True, "4")
    assert m.group == "4"


def test_table10_uses_mapping_group():
    # Lightweight synthetic objects are easier here than file fixtures.
    from core import CombinationRow, MemberResult, ReactionRow

    reactions = [
        ReactionRow(1, 4, False, 1, 2, 3, 4, 5, 6, "1/4"),
        ReactionRow(1, 5, False, 7, 8, 9, 10, 11, 12, "1/5"),
    ]
    mappings = [
        CaseMapping(4, "TEMP +", "T+", True, "4"),
        CaseMapping(5, "TEMP -", "T-", True, "4"),
    ]
    out = process_data([], reactions, [], mappings, "Teste", 200, 0.01, 0.01)
    assert [row[1] for row in out.table10] == ["4", "4"]
