import pytest
from . import excelparser as xlp
from openpyxl import Workbook, load_workbook
import os

def test_tokenizer():
    text = """=IF($A$1,"then True",MAX(DEFAULT_VAL,'Sheet 2'!B1))"""
    tokens = list(xlp.generate_tokens(text))
    assert len(tokens)==11
    assert tokens[0].type == "FUNC"
    assert tokens[1].type == "OPERAND"


def test_columnind():
    assert xlp.columnind("A")==1
    assert xlp.columnind("AA")==27
    assert sorted(["A","AA","B","C","AAA"], key=xlp.columnind) == ["A","B","C","AA","AAA"]


def test_ExpressionEvaluator():
    e = xlp.ExpressionEvaluator()
    assert e.parse("=2+3") == 5
    assert e.parse("=2*(3+4)") == 14
    assert e.parse("=2*3") == 6
    assert e.parse("=3*(2*(1+1)+3*(1+0))") == 21
    assert e.parse("=SUM(1,2,3,4)")==10
    assert e.parse("=SUM(1,2,3,4, SUM(2,3))")== 15
    assert e.parse("=1+SUM(1,2,3,4, SUM(2,3))")== 16
    assert e.parse("=1+ SUM(1,2,3,4, SUM(2,3))")== 16
    assert e.parse("=1*SUM(1,2,3,4, SUM(2,3))")== 15
    assert e.parse("=1 *SUM(1,2,3,4, SUM(2,3))")== 15
    assert e.parse("=IF(3=4,1,2)") == 2
    assert e.parse('=IF("YES"="NO",1,2)')
    assert e.parse("=IF(3>4,1,2)") == 2
    assert e.parse("=IF(3<4,1,2)") == 1
    assert e.parse("=2^3+4") == 12
    assert e.parse("=3%") == 0.03
    assert e.parse("=3%*2") == 0.06
    assert e.parse("=-3") == -3
    assert e.parse("=-SUM(1,2)") == -3
    assert e.parse("=2=2") == True
    assert e.parse("=3>2") == True
    assert e.parse("=-3>2") == False
    assert e.parse("=2<2+1") == True
    assert e.parse("=2--3") == 5

    
def test_ExpressionTreeBuilder():
    e = xlp.ExpressionTreeBuilder()
    assert e.parse("=2+3") == ("+", 2, 3)
    assert e.parse("=2*(3+4)") == ("*", 2, ("+", 3, 4))
    assert e.parse("=1+SUM(1)") == ("+", 1, ("SUM", 1))
    assert e.parse("=1+SUM(1.1, 2, 3)") == ("+", 1, ("SUM", 1.1, 2, 3))
    assert e.parse("=SUM(1.1, 2,SUM(1,2))") == ("SUM", 1.1, 2, ("SUM",1,2))
    assert e.parse("=IF(1,2,3)") == ("IF",1,2,3)
    assert e.parse("=1+IF(1,2,3)") == ("+", 1, ("IF", 1, 2, 3))
    assert e.parse("=2<1+3") == ("<" , 2 ,("+", 1, 3))
    assert e.parse("=IF(2<1+3,IF(1>0,10/2,0),0)") == ("IF",("<",2,("+", 1, 3)),("IF",(">", 1, 0), ("/", 10, 2), 0),0)


@pytest.fixture
def workbook():
    w = load_workbook("/home/vikrant/Documents/prayas/RE_Tariff_and_Financial_Analysis_Tool_v2.1-unprotected.xlsx")
    yield w


def test_expand(workbook):
    e = xlp.ExpressionTreeBuilder(workbook)
    assert xlp.expand(()) == ()
    assert xlp.expand(("*", 1, 2)) == ("*", 1, 2)
    assert xlp.expand((1,2,3,4)) == (1,2,3,4)
    assert xlp.expand(("*", 1, ("+", 1, 2))) == ("*", 1, ("+", 1, 2))
    assert xlp.expand(("SUM", 1,2,3,("*",1,2))) == ("SUM", 1,2,3,("*",1,2))
    def two():
        return 2
    assert xlp.expand(("*", two, two)) == ("*" , 2, 2)
    assert xlp.expand(e.parse("=SUM(A1:D1)")) == ("SUM", (("Introduction!A1", "Introduction!B1", "Introduction!C1", "Introduction!D1"),))
    assert xlp.expand(("*", 1, ("+" , 2, lambda: 3))) == ("*", 1, ("+", 2, 3))

def test_is_expanded():
    assert xlp.is_expanded(("*", 1, 1))
    assert xlp.is_expanded(("*", lambda : 2, lambda: 3)) == False
    assert xlp.is_expanded(("*", 1, ("+" , 2, lambda: 3))) == False
    
def test_excel(workbook):
    e = xlp.ExpressionTreeBuilder(workbook)
    assert e.parse("=SUM(A1:D1)") == ("SUM", (("Introduction!A1", "Introduction!B1", "Introduction!C1", "Introduction!D1"),))
    expr = e.parse("='Inputs&Summary!M4'")
    
    

