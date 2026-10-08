from core import CombinationRow, CaseMapping, translate_combination


def test_translate_grouped_combination():
    mapping = {
        1: CaseMapping(1, "PESO PRÓPRIO", "PP"),
        2: CaseMapping(2, "CARGA PERMANENTE", "CP"),
        3: CaseMapping(3, "SOBRECARGA", "SC"),
        5: CaseMapping(5, "TEMPERATURA -20°C", "T-°C"),
    }
    assert translate_combination("1*1.25+(2+3)*1.50+5*0.72", mapping) == "1,25PP + 1,50CP + 1,50SC + 0,72T-°C"
