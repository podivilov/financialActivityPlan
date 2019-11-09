#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import uuid
import pytz
import xlrd
from xml.dom import minidom
from datetime import datetime

# Устанавливаем стандартную кодировку
reload(sys)
sys.setdefaultencoding('utf8')

# Полезные переменные
local_tz = pytz.timezone('Europe/Moscow')

# Полезные функции
def utc_to_local(utc_dt):
    local_dt = utc_dt.replace(tzinfo=pytz.utc).astimezone(local_tz)
    return local_tz.normalize(local_dt)

def insert_str(string, str_to_insert, index):
    return string[:index] + str_to_insert + string[index:]

# Открываем рабочий файл
workbook = xlrd.open_workbook(sys.argv[1])
worksheet = workbook.sheet_by_index(3)

# Создаём коренной элемент
doc = minidom.Document()
root = doc.createElement('ns2:financialActivityPlan2017')
doc.appendChild(root)

# Header
header = doc.createElement('header')
root.appendChild(header)

# ID
id = doc.createElement('id')
id.appendChild(doc.createTextNode(str(uuid.uuid4())))
header.appendChild(id)

# createDateTime
createDateTime = doc.createElement('createDateTime')
createDateTime.appendChild(doc.createTextNode(str(utc_to_local(datetime.utcnow()).strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + "+03:00")))
header.appendChild(createDateTime)

# ns2:body
ns2_body = doc.createElement('ns2:body')
root.appendChild(ns2_body)

# ns2:position
ns2_position = doc.createElement('ns2:position')
ns2_body.appendChild(ns2_position)

# positionId
positionId = doc.createElement('positionId')
positionId.appendChild(doc.createTextNode(str(uuid.uuid4())))
ns2_position.appendChild(positionId)

# changeDate
changeDate = doc.createElement('changeDate')
changeDate.appendChild(doc.createTextNode(str(utc_to_local(datetime.utcnow()).strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + "+03:00")))
ns2_position.appendChild(changeDate)

# placer
placer = doc.createElement('placer')
ns2_position.appendChild(placer)

# regNum
placer_regNum = doc.createElement('regNum')
placer_regNum.appendChild(doc.createTextNode('462D1140'))
placer.appendChild(placer_regNum)

# fullName
placer_fullName = doc.createElement('fullName')
placer_fullName.appendChild(doc.createTextNode('ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ МОСКОВСКОЙ ОБЛАСТИ «КОЛЛЕДЖ «КОЛОМНА»'))
placer.appendChild(placer_fullName)

# inn
placer_inn = doc.createElement('inn')
placer_inn.appendChild(doc.createTextNode('5022049898'))
placer.appendChild(placer_inn)

# kpp
placer_kpp = doc.createElement('kpp')
placer_kpp.appendChild(doc.createTextNode('502201001'))
placer.appendChild(placer_kpp)

# initiator
initiator = doc.createElement('initiator')
ns2_position.appendChild(initiator)

# regNum
initiator_regNum = doc.createElement('regNum')
initiator_regNum.appendChild(doc.createTextNode('462D1140'))
initiator.appendChild(initiator_regNum)

# fullName
initiator_fullName = doc.createElement('fullName')
initiator_fullName.appendChild(doc.createTextNode('ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ МОСКОВСКОЙ ОБЛАСТИ «КОЛЛЕДЖ «КОЛОМНА»'))
initiator.appendChild(initiator_fullName)

# inn
initiator_inn = doc.createElement('inn')
initiator_inn.appendChild(doc.createTextNode('5022049898'))
initiator.appendChild(initiator_inn)

# kpp
initiator_kpp = doc.createElement('kpp')
initiator_kpp.appendChild(doc.createTextNode('502201001'))
initiator.appendChild(initiator_kpp)

# versionNumber
versionNumber = doc.createElement('versionNumber')
versionNumber.appendChild(doc.createTextNode('0'))
ns2_position.appendChild(versionNumber)

# now
now = datetime.now()

# financialYear
financialYear = doc.createElement('financialYear')
financialYear.appendChild(doc.createTextNode(str(now.year - 1)))
ns2_position.appendChild(financialYear)

# planFirstYear
planFirstYear = doc.createElement('planFirstYear')
planFirstYear.appendChild(doc.createTextNode(str(now.year)))
ns2_position.appendChild(planFirstYear)

# planLastYear
planLastYear = doc.createElement('planLastYear')
planLastYear.appendChild(doc.createTextNode(str(now.year + 1)))
ns2_position.appendChild(planLastYear)

# financialIndex
financialIndex = doc.createElement('financialIndex')
ns2_position.appendChild(financialIndex)

# nonfinancialAssets
nonfinancialAssets = doc.createElement('nonfinancialAssets')
financialIndex.appendChild(nonfinancialAssets)

# realAssets
realAssets = doc.createElement('realAssets')
realAssets.appendChild(doc.createTextNode(str(worksheet.cell(10, 71).value)))
nonfinancialAssets.appendChild(realAssets)

# realAssetsResidual
realAssetsResidual = doc.createElement('realAssetsResidual')
realAssetsResidual.appendChild(doc.createTextNode(str(worksheet.cell(11, 71).value)))
nonfinancialAssets.appendChild(realAssetsResidual)

# highValuePersonalAssets
highValuePersonalAssets = doc.createElement('highValuePersonalAssets')
highValuePersonalAssets.appendChild(doc.createTextNode(str(worksheet.cell(12, 71).value)))
nonfinancialAssets.appendChild(highValuePersonalAssets)

# highValuePersonalAssetsResidual
highValuePersonalAssetsResidual = doc.createElement('highValuePersonalAssetsResidual')
highValuePersonalAssetsResidual.appendChild(doc.createTextNode(str(worksheet.cell(12, 71).value)))
nonfinancialAssets.appendChild(highValuePersonalAssetsResidual)

# total
total = doc.createElement('total')
total.appendChild(doc.createTextNode(str(worksheet.cell(9, 71).value)))
nonfinancialAssets.appendChild(total)

# financialAssets
financialAssets = doc.createElement('financialAssets')
financialIndex.appendChild(financialAssets)

# cash
cash = doc.createElement('cash')
cash.appendChild(doc.createTextNode(str(worksheet.cell(15, 71).value)))
financialAssets.appendChild(cash)

# accountsCash
accountsCash = doc.createElement('accountsCash')
accountsCash.appendChild(doc.createTextNode(str(worksheet.cell(16, 71).value)))
financialAssets.appendChild(accountsCash)

# depositCash
depositCash = doc.createElement('depositCash')
depositCash.appendChild(doc.createTextNode(str(worksheet.cell(18, 71).value)))
financialAssets.appendChild(depositCash)

# others
others = doc.createElement('others')
others.appendChild(doc.createTextNode(str(worksheet.cell(19, 71).value)))
financialAssets.appendChild(others)

# debit
debit = doc.createElement('debit')
financialAssets.appendChild(debit)

# income
income = doc.createElement('income')
income.appendChild(doc.createTextNode(str(worksheet.cell(20, 71).value)))
debit.appendChild(income)

# expense
expense = doc.createElement('expense')
expense.appendChild(doc.createTextNode(str(worksheet.cell(21, 71).value)))
debit.appendChild(expense)

# total
total = doc.createElement('total')
total.appendChild(doc.createTextNode(str(worksheet.cell(14, 71).value)))
debit.appendChild(total)

# financialCircumstances
financialCircumstances = doc.createElement('financialCircumstances')
financialIndex.appendChild(financialCircumstances)

# debentures
debentures = doc.createElement('debentures')
debentures.appendChild(doc.createTextNode(str(worksheet.cell(23, 71).value)))
financialCircumstances.appendChild(debentures)

# kredit
kredit = doc.createElement('kredit')
kredit.appendChild(doc.createTextNode(str(worksheet.cell(24, 71).value)))
financialCircumstances.appendChild(kredit)

# kreditExpired
kreditExpired = doc.createElement('kreditExpired')
kreditExpired.appendChild(doc.createTextNode(str(worksheet.cell(25, 71).value)))
financialCircumstances.appendChild(kreditExpired)

# total
total = doc.createElement('total')
total.appendChild(doc.createTextNode(str(worksheet.cell(22, 71).value)))
financialCircumstances.appendChild(total)

# Работаем с 8 страницей (индекс = 8-1)
worksheet = workbook.sheet_by_index(7)

# expensePaymentIndex (1)
expensePaymentIndex = doc.createElement('expensePaymentIndex')
ns2_position.appendChild(expensePaymentIndex)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode(str(worksheet.cell(10, 1).value)))
expensePaymentIndex.appendChild(name)

# lineCode
lineCode = doc.createElement('lineCode')
lineCode.appendChild(doc.createTextNode(str(worksheet.cell(10, 22).value)))
expensePaymentIndex.appendChild(lineCode)

# totalSum
totalSum = doc.createElement('totalSum')
expensePaymentIndex.appendChild(totalSum)

# nextYear
nextYear = doc.createElement('nextYear')
nextYear.appendChild(doc.createTextNode(str(worksheet.cell(10, 41).value)))
totalSum.appendChild(nextYear)

# firstPlanYear
firstPlanYear = doc.createElement('firstPlanYear')
firstPlanYear.appendChild(doc.createTextNode(str(worksheet.cell(10, 55).value)))
totalSum.appendChild(firstPlanYear)

# secondPlanYear
secondPlanYear = doc.createElement('secondPlanYear')
secondPlanYear.appendChild(doc.createTextNode(str(worksheet.cell(10, 69).value)))
totalSum.appendChild(secondPlanYear)

# fz44Sum
fz44Sum = doc.createElement('fz44Sum')
expensePaymentIndex.appendChild(fz44Sum)

# nextYear
nextYear = doc.createElement('nextYear')
nextYear.appendChild(doc.createTextNode(str(worksheet.cell(10, 83).value)))
fz44Sum.appendChild(nextYear)

# firstPlanYear
firstPlanYear = doc.createElement('firstPlanYear')
firstPlanYear.appendChild(doc.createTextNode(str(worksheet.cell(10, 97).value)))
fz44Sum.appendChild(firstPlanYear)

# secondPlanYear
secondPlanYear = doc.createElement('secondPlanYear')
secondPlanYear.appendChild(doc.createTextNode(str(worksheet.cell(10, 111).value)))
fz44Sum.appendChild(secondPlanYear)

# fz223Sum
fz223Sum = doc.createElement('fz223Sum')
expensePaymentIndex.appendChild(fz223Sum)

# nextYear
nextYear = doc.createElement('nextYear')
nextYear.appendChild(doc.createTextNode('0'))
fz223Sum.appendChild(nextYear)

# firstPlanYear
firstPlanYear = doc.createElement('firstPlanYear')
firstPlanYear.appendChild(doc.createTextNode('0'))
fz223Sum.appendChild(firstPlanYear)

# secondPlanYear
secondPlanYear = doc.createElement('secondPlanYear')
secondPlanYear.appendChild(doc.createTextNode('0'))
fz223Sum.appendChild(secondPlanYear)

# expensePaymentIndex (2)
expensePaymentIndex = doc.createElement('expensePaymentIndex')
ns2_position.appendChild(expensePaymentIndex)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode(str(worksheet.cell(11, 1).value)))
expensePaymentIndex.appendChild(name)

# lineCode
lineCode = doc.createElement('lineCode')
lineCode.appendChild(doc.createTextNode(str(worksheet.cell(11, 22).value)))
expensePaymentIndex.appendChild(lineCode)

# totalSum
totalSum = doc.createElement('totalSum')
expensePaymentIndex.appendChild(totalSum)

# nextYear
nextYear = doc.createElement('nextYear')
nextYear.appendChild(doc.createTextNode(str(worksheet.cell(11, 41).value)))
totalSum.appendChild(nextYear)

# firstPlanYear
firstPlanYear = doc.createElement('firstPlanYear')
firstPlanYear.appendChild(doc.createTextNode(str(worksheet.cell(11, 55).value)))
totalSum.appendChild(firstPlanYear)

# secondPlanYear
secondPlanYear = doc.createElement('secondPlanYear')
secondPlanYear.appendChild(doc.createTextNode(str(worksheet.cell(11, 69).value)))
totalSum.appendChild(secondPlanYear)

# fz44Sum
fz44Sum = doc.createElement('fz44Sum')
expensePaymentIndex.appendChild(fz44Sum)

# nextYear
nextYear = doc.createElement('nextYear')
nextYear.appendChild(doc.createTextNode(str(worksheet.cell(11, 83).value)))
fz44Sum.appendChild(nextYear)

# firstPlanYear
firstPlanYear = doc.createElement('firstPlanYear')
firstPlanYear.appendChild(doc.createTextNode(str(worksheet.cell(11, 97).value)))
fz44Sum.appendChild(firstPlanYear)

# secondPlanYear
secondPlanYear = doc.createElement('secondPlanYear')
secondPlanYear.appendChild(doc.createTextNode(str(worksheet.cell(11, 111).value)))
fz44Sum.appendChild(secondPlanYear)

# fz223Sum
fz223Sum = doc.createElement('fz223Sum')
expensePaymentIndex.appendChild(fz223Sum)

# nextYear
nextYear = doc.createElement('nextYear')
nextYear.appendChild(doc.createTextNode('0'))
fz223Sum.appendChild(nextYear)

# firstPlanYear
firstPlanYear = doc.createElement('firstPlanYear')
firstPlanYear.appendChild(doc.createTextNode('0'))
fz223Sum.appendChild(firstPlanYear)

# secondPlanYear
secondPlanYear = doc.createElement('secondPlanYear')
secondPlanYear.appendChild(doc.createTextNode('0'))
fz223Sum.appendChild(secondPlanYear)

# expensePaymentIndex (3)
expensePaymentIndex = doc.createElement('expensePaymentIndex')
ns2_position.appendChild(expensePaymentIndex)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode(str(worksheet.cell(12, 1).value)))
expensePaymentIndex.appendChild(name)

# lineCode
lineCode = doc.createElement('lineCode')
lineCode.appendChild(doc.createTextNode(str(worksheet.cell(12, 22).value)))
expensePaymentIndex.appendChild(lineCode)

# totalSum
totalSum = doc.createElement('totalSum')
expensePaymentIndex.appendChild(totalSum)

# nextYear
nextYear = doc.createElement('nextYear')
nextYear.appendChild(doc.createTextNode(str(worksheet.cell(12, 41).value)))
totalSum.appendChild(nextYear)

# firstPlanYear
firstPlanYear = doc.createElement('firstPlanYear')
firstPlanYear.appendChild(doc.createTextNode(str(worksheet.cell(12, 55).value)))
totalSum.appendChild(firstPlanYear)

# secondPlanYear
secondPlanYear = doc.createElement('secondPlanYear')
secondPlanYear.appendChild(doc.createTextNode(str(worksheet.cell(12, 69).value)))
totalSum.appendChild(secondPlanYear)

# fz44Sum
fz44Sum = doc.createElement('fz44Sum')
expensePaymentIndex.appendChild(fz44Sum)

# nextYear
nextYear = doc.createElement('nextYear')
nextYear.appendChild(doc.createTextNode(str(worksheet.cell(12, 83).value)))
fz44Sum.appendChild(nextYear)

# firstPlanYear
firstPlanYear = doc.createElement('firstPlanYear')
firstPlanYear.appendChild(doc.createTextNode(str(worksheet.cell(12, 97).value)))
fz44Sum.appendChild(firstPlanYear)

# secondPlanYear
secondPlanYear = doc.createElement('secondPlanYear')
secondPlanYear.appendChild(doc.createTextNode(str(worksheet.cell(12, 111).value)))
fz44Sum.appendChild(secondPlanYear)

# fz223Sum
fz223Sum = doc.createElement('fz223Sum')
expensePaymentIndex.appendChild(fz223Sum)

# nextYear
nextYear = doc.createElement('nextYear')
nextYear.appendChild(doc.createTextNode('0'))
fz223Sum.appendChild(nextYear)

# firstPlanYear
firstPlanYear = doc.createElement('firstPlanYear')
firstPlanYear.appendChild(doc.createTextNode('0'))
fz223Sum.appendChild(firstPlanYear)

# secondPlanYear
secondPlanYear = doc.createElement('secondPlanYear')
secondPlanYear.appendChild(doc.createTextNode('0'))
fz223Sum.appendChild(secondPlanYear)

# Работаем с 9 страницей (индекс = 9-1)
worksheet = workbook.sheet_by_index(8)

# temporaryResources (1)
temporaryResources = doc.createElement('temporaryResources')
expensePaymentIndex.appendChild(temporaryResources)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode(str(worksheet.cell(5, 1).value)))
temporaryResources.appendChild(name)

# lineCode
lineCode = doc.createElement('lineCode')
lineCode.appendChild(doc.createTextNode(str(worksheet.cell(5, 75).value)))
temporaryResources.appendChild(lineCode)

# total
total = doc.createElement('total')
total.appendChild(doc.createTextNode(str(worksheet.cell(5, 90).value)))
temporaryResources.appendChild(total)

# temporaryResources (2)
temporaryResources = doc.createElement('temporaryResources')
expensePaymentIndex.appendChild(temporaryResources)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode(str(worksheet.cell(6, 1).value)))
temporaryResources.appendChild(name)

# lineCode
lineCode = doc.createElement('lineCode')
lineCode.appendChild(doc.createTextNode(str(worksheet.cell(6, 75).value)))
temporaryResources.appendChild(lineCode)

# total
total = doc.createElement('total')
total.appendChild(doc.createTextNode(str(worksheet.cell(6, 90).value)))
temporaryResources.appendChild(total)

# temporaryResources (3)
temporaryResources = doc.createElement('temporaryResources')
expensePaymentIndex.appendChild(temporaryResources)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode(str(worksheet.cell(7, 1).value)))
temporaryResources.appendChild(name)

# lineCode
lineCode = doc.createElement('lineCode')
lineCode.appendChild(doc.createTextNode(str(worksheet.cell(7, 75).value)))
temporaryResources.appendChild(lineCode)

# total
total = doc.createElement('total')
total.appendChild(doc.createTextNode(str(worksheet.cell(7, 90).value)))
temporaryResources.appendChild(total)

# temporaryResources (3)
temporaryResources = doc.createElement('temporaryResources')
expensePaymentIndex.appendChild(temporaryResources)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode(str(worksheet.cell(8, 1).value)))
temporaryResources.appendChild(name)

# lineCode
lineCode = doc.createElement('lineCode')
lineCode.appendChild(doc.createTextNode(str(worksheet.cell(8, 75).value)))
temporaryResources.appendChild(lineCode)

# total
total = doc.createElement('total')
total.appendChild(doc.createTextNode(str(worksheet.cell(8, 90).value)))
temporaryResources.appendChild(total)

# reference (1)
reference = doc.createElement('reference')
expensePaymentIndex.appendChild(reference)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode(str(worksheet.cell(16, 1).value)))
reference.appendChild(name)

# lineCode
lineCode = doc.createElement('lineCode')
lineCode.appendChild(doc.createTextNode(str(worksheet.cell(16, 75).value)))
reference.appendChild(lineCode)

# total
total = doc.createElement('total')
total.appendChild(doc.createTextNode(str(worksheet.cell(16, 90).value)))
reference.appendChild(total)

# reference (2)
reference = doc.createElement('reference')
expensePaymentIndex.appendChild(reference)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode(str(worksheet.cell(17, 1).value)))
reference.appendChild(name)

# lineCode
lineCode = doc.createElement('lineCode')
lineCode.appendChild(doc.createTextNode(str(worksheet.cell(17, 75).value)))
reference.appendChild(lineCode)

# total
total = doc.createElement('total')
total.appendChild(doc.createTextNode(str(worksheet.cell(17, 90).value)))
reference.appendChild(total)

# reference (3)
reference = doc.createElement('reference')
expensePaymentIndex.appendChild(reference)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode(str(worksheet.cell(18, 1).value)))
reference.appendChild(name)

# lineCode
lineCode = doc.createElement('lineCode')
lineCode.appendChild(doc.createTextNode(str(worksheet.cell(18, 75).value)))
reference.appendChild(lineCode)

# total
total = doc.createElement('total')
total.appendChild(doc.createTextNode(str(worksheet.cell(18, 90).value)))
reference.appendChild(total)

xml_str = doc.toprettyxml(indent="    ")
with open(sys.argv[2], "w") as f:
    f.write(xml_str)
