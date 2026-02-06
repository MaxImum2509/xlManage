"""Tests pour le parser d'arguments de macros VBA."""

import pytest

from xlmanage.macro_runner import _parse_macro_args
from xlmanage.exceptions import VBAMacroError


def test_parse_empty_args():
    """Test avec une chaîne vide."""
    assert _parse_macro_args("") == []
    assert _parse_macro_args("   ") == []


def test_parse_single_string():
    """Test avec une seule chaîne."""
    assert _parse_macro_args('"hello"') == ["hello"]
    assert _parse_macro_args("'world'") == ["world"]


def test_parse_string_with_comma():
    """Test chaîne avec virgule interne."""
    assert _parse_macro_args('"hello, world"') == ["hello, world"]
    assert _parse_macro_args("'a,b,c'") == ["a,b,c"]


def test_parse_integers():
    """Test conversion en int."""
    assert _parse_macro_args("42") == [42]
    assert _parse_macro_args("-100") == [-100]
    assert _parse_macro_args("+99") == [99]


def test_parse_floats():
    """Test conversion en float."""
    assert _parse_macro_args("3.14") == [3.14]
    assert _parse_macro_args("-0.5") == [-0.5]
    assert _parse_macro_args("123.456") == [123.456]


def test_parse_booleans():
    """Test conversion en bool."""
    assert _parse_macro_args("true") == [True]
    assert _parse_macro_args("false") == [False]
    assert _parse_macro_args("True") == [True]
    assert _parse_macro_args("FALSE") == [False]
    assert _parse_macro_args("TrUe") == [True]


def test_parse_mixed_types():
    """Test avec plusieurs types mélangés."""
    result = _parse_macro_args('"hello",42,3.14,true,"world",false,-10')
    assert result == ["hello", 42, 3.14, True, "world", False, -10]


def test_parse_complex_string():
    """Test chaîne complexe avec virgules et guillemets."""
    result = _parse_macro_args('"hello, world",42,"foo,bar,baz",3.14')
    assert result == ["hello, world", 42, "foo,bar,baz", 3.14]


def test_parse_with_spaces():
    """Test avec espaces autour des valeurs."""
    result = _parse_macro_args('  "hello"  ,  42  ,  3.14  ')
    assert result == ["hello", 42, 3.14]


def test_parse_unquoted_string():
    """Test chaîne sans guillemets (fallback)."""
    result = _parse_macro_args("hello,world")
    assert result == ["hello", "world"]


def test_parse_number_as_string():
    """Test nombre qui n'est pas valide → str."""
    result = _parse_macro_args("123abc")
    assert result == ["123abc"]


def test_parse_too_many_args():
    """Test erreur si > 30 arguments."""
    # Créer 31 arguments
    args_str = ",".join([str(i) for i in range(31)])

    with pytest.raises(VBAMacroError) as exc_info:
        _parse_macro_args(args_str)

    assert "31" in str(exc_info.value)
    assert "30" in str(exc_info.value)


def test_parse_edge_case_single_quote_in_double():
    """Test guillemet simple dans une chaîne entre guillemets doubles."""
    result = _parse_macro_args('''"it's working"''')
    assert result == ["it's working"]


def test_parse_edge_case_double_quote_in_single():
    """Test guillemet double dans une chaîne entre guillemets simples."""
    result = _parse_macro_args("""'he said "hello"'""")
    assert result == ['he said "hello"']


def test_parse_realistic_scenario():
    """Test scénario réaliste avec plusieurs types."""
    args = '''"Report_2024",100,"Sheet1",true,3.5,"C:\\\\Users\\\\test.xlsx"'''
    result = _parse_macro_args(args)

    assert result == [
        "Report_2024",
        100,
        "Sheet1",
        True,
        3.5,
        "C:\\\\Users\\\\test.xlsx",
    ]
    assert isinstance(result[0], str)
    assert isinstance(result[1], int)
    assert isinstance(result[2], str)
    assert isinstance(result[3], bool)
    assert isinstance(result[4], float)
    assert isinstance(result[5], str)
