#!/usr/bin/python
# encoding: utf-8
from __future__ import print_function
from openpyxl import load_workbook

# Nome do arquivo Excel
wb = load_workbook(filename = 'Moodle-Q.xlsx')
ws = wb.active

# Cabeçalho do quiz
quizHeader = """
<?xml version="1.0" encoding="UTF-8"?>
<quiz>"""

# Cabeçalho da questão
questionHeader = """
    <question type="multichoice">
        <name>
            <text>{0}-{1}-Q"""
# Questão
question = """{0}</text>
        </name>
        <questiontext format="html">
            <text><![CDATA[<p>{1}</p>]]></text>
        </questiontext>
        <generalfeedback format="html">
            <text></text>
        </generalfeedback>
        <defaultgrade>1.0000000</defaultgrade>
        <penalty>0.3333333</penalty>
        <hidden>0</hidden>
        <single>true</single>
        <shuffleanswers>true</shuffleanswers>
        <answernumbering>abc</answernumbering>
        <correctfeedback format="html">
            <text></text>
        </correctfeedback>
        <partiallycorrectfeedback format="html">
            <text></text>
        </partiallycorrectfeedback>
        <incorrectfeedback format="html">
            <text></text>
        </incorrectfeedback>
        <answer fraction="100" format="html">
            <text><![CDATA[<p><span>{2}</span></p>]]></text>
            <feedback format="html">
            <text></text>
            </feedback>
        </answer>
        <answer fraction="0" format="html">
            <text><![CDATA[<p><span>{3}</span></p>]]></text>
            <feedback format="html">
            <text></text>
            </feedback>
        </answer>
        <answer fraction="0" format="html">
            <text><![CDATA[<p><span>{4}</span></p>]]></text>
            <feedback format="html">
            <text></text>
            </feedback>
        </answer>
        <answer fraction="0" format="html">
            <text><![CDATA[<p><span>{5}</span></p>]]></text>
            <feedback format="html">
            <text></text>
            </feedback>
        </answer>
    </question>
"""
# Rodapé do questionário
quizFooter ="""</quiz>"""

print(quizHeader, end = '')
print(questionHeader.format(ws["B1"].value, ws["B2"].value), end = '')
print(question.format(ws["A5"].value, ws["B5"].value, ws["C5"].value, ws["D5"].value, ws["E5"].value, ws["F5"].value, ws["G5"].value), end = '')

print(questionHeader.format(ws["B1"].value, ws["B2"].value), end = '')
print(question.format(ws["A6"].value, ws["B6"].value, ws["C6"].value, ws["D6"].value, ws["E6"].value, ws["F6"].value, ws["G6"].value), end = '')

print(questionHeader.format(ws["B1"].value, ws["B2"].value), end = '')
print(question.format(ws["A7"].value, ws["B7"].value, ws["C7"].value, ws["D7"].value, ws["E7"].value, ws["F7"].value, ws["G7"].value), end = '')

print(questionHeader.format(ws["B1"].value, ws["B2"].value), end = '')
print(question.format(ws["A8"].value, ws["B8"].value, ws["C8"].value, ws["D8"].value, ws["E8"].value, ws["F8"].value, ws["G8"].value), end = '')
print(quizFooter)
