import random #to avoid segmentation error https://bugs.mysql.com/bug.php?id=89889
import mysql.connector as mariadb
import json
from openpyxl import Workbook
import sys
import settings

data = dict()


def load_data(connection):
    cursor = connection.cursor()
    cursor.execute("SELECT * FROM nlp_intents")
    dataIntents = cursor.fetchall()

    for row in dataIntents:
        intent = dict()
        intentId = row[0]
        intent['shortQuestion'] = row[4]
        intent['configuration'] = row[1]
        intent['skill'] = row[5]
        intent['skillState'] = row[6]
        intent['targetSkillPayload'] = row[7]
        intent['confirmationQuestion'] = row[8]
        data[intentId] = intent

        load_synonym(cursor, intentId)
        load_named_entities(cursor, intentId)
        load_answers(cursor, intentId)


def load_synonym(cursor, intentId):
    sql_select_query = """select * from nlp_question_synonyms where intentId = %s"""
    cursor.execute(sql_select_query, (intentId,))
    dataSynonyms = cursor.fetchall()

    if len(dataSynonyms) == 0:
        data[intentId]['synonyms'] = ''

    synonyms = list()
    for syn in dataSynonyms:
        synonyms.append(syn[3])
        data[intentId]['synonyms'] = synonyms


def load_named_entities(cursor, intentId):
    sql_select_query = """select * from nlp_named_entities where intentId = %s"""
    cursor.execute(sql_select_query, (intentId,))
    dataNamedEntities = cursor.fetchall()

    if len(dataNamedEntities) == 0:
        data[intentId]['entities'] = ''

    for namedEnt in dataNamedEntities:
        question = namedEnt[2]
        typeEnt = namedEnt[4]
        name = namedEnt[5]
        specification = json.loads(namedEnt[3])
        ent = [
            {
                "question": question,
                "type": typeEnt,
                "specification": specification,
                "name": name
            }
        ]
        data[intentId]['entities'] = json.dumps(ent, ensure_ascii=False)


def load_answers(cursor, intentId):
    sql_select_query = """select * from nlp_intent_answers where intentId = %s"""
    cursor.execute(sql_select_query, (intentId,))
    dataAnswers = cursor.fetchall()
    if len(dataAnswers) == 0:
        data[intentId]['answers'] = list()
    namedAnswers = list()
    for namedEnt in dataAnswers:
        namedAnswers.append(namedEnt[2])
        data[intentId]['answers'] = namedAnswers


def export_to_excel():
    book = Workbook()
    sheet = book.active

    headers = ['Эталон вопроса', 'Подтверждающий вопрос',
               'Синонимы', 'Переменные', 'Ответы', 'Skill', 'SkillState', 'targetSkillPayload', 'Дополнения']
    sheet.append(headers)
    sheet.column_dimensions['A'].width = 50
    sheet.column_dimensions['B'].width = 50
    sheet.column_dimensions['C'].width = 50
    sheet.column_dimensions['D'].width = 50
    sheet.column_dimensions['E'].width = 50
    sheet.column_dimensions['F'].width = 50
    sheet.column_dimensions['G'].width = 50
    sheet.column_dimensions['H'].width = 50

    a = 2
    for intentId in data:
        array = data[intentId]
        if array['shortQuestion'] is not None:
            sheet['A' + str(a)] = str(array['shortQuestion'])
        if array['confirmationQuestion'] is not None:
            sheet['B' + str(a)] = str(array['confirmationQuestion'])
        if array['synonyms'] is not None:
            sheet['C' + str(a)] = '|'.join(map(str, array['synonyms']))
        if array['entities'] is not None:
            sheet['D' + str(a)] = array['entities']
        if array['answers'] is not None:
            sheet['E' + str(a)] = '|'.join(map(str, array['answers']))
        if array['skill'] is not None:
            sheet['F' + str(a)] = str(array['skill'])
        if array['skillState'] is not None:
            sheet['G' + str(a)] = str(array['skillState'])
        if array['targetSkillPayload'] is not None:
            sheet['H' + str(a)] = ''.join(str(array['targetSkillPayload']))
        a += 1

    book.save(str(sys.argv[sys.argv.index("--HR_EXCEL_FILE_PATH") + 1]))


if __name__ == "__main__":
    db_connection = mariadb.connect(
        host=str(sys.argv[sys.argv.index("--DB_HOST") + 1]),
        port=int(sys.argv[sys.argv.index("--DB_PORT") + 1]),
        user=str(sys.argv[sys.argv.index("--DB_USER") + 1]),
        password=str(sys.argv[sys.argv.index("--DB_PASSWORD") + 1]),
        database=str(sys.argv[sys.argv.index("--DB_NAME") + 1])
    )
    load_data(db_connection)
    export_to_excel()
