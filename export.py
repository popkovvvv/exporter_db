#! /usr/bin/env python
import argparse
import json
import random  # to avoid segmentation error https://bugs.mysql.com/bug.php?id=89889

import mysql.connector as mariadb
from openpyxl import Workbook

data = dict()


def load_data(connection):
    cursor = connection.cursor()
    cursor.execute("SELECT * FROM nlp_intents")
    data_intents = cursor.fetchall()

    for row in data_intents:
        intent = dict()
        intent_id = row[0]
        intent['short_question'] = row[4]
        intent['configuration'] = row[1]
        intent['skill'] = row[5]
        intent['skill_state'] = row[6]
        intent['target_skill_payload'] = row[7]
        intent['confirmation_question'] = row[8]
        intent['knowledge_base'] = row[3]
        data[intent_id] = intent

        load_synonym(cursor, intent_id)
        load_named_entities(cursor, intent_id)
        load_answers(cursor, intent_id)


def load_synonym(cursor, intent_id):
    sql_select_query = """select * from nlp_question_synonyms where intentId = %s"""
    cursor.execute(sql_select_query, (intent_id,))
    data_synonyms = cursor.fetchall()

    if len(data_synonyms) == 0:
        data[intent_id]['synonyms'] = ''

    synonyms = list()
    for syn in data_synonyms:
        synonym = str(syn[3]).strip()
        synonyms.append(synonym)
        data[intent_id]['synonyms'] = synonyms


def load_named_entities(cursor, intent_id):
    sql_select_query = """select * from nlp_named_entities where intentId = %s"""
    cursor.execute(sql_select_query, (intent_id,))
    data_named_entities = cursor.fetchall()

    if len(data_named_entities) == 0:
        data[intent_id]['entities'] = ''

    for namedEnt in data_named_entities:
        question = namedEnt[2]
        type_ent = namedEnt[4]
        name = namedEnt[5]
        specification = json.loads(namedEnt[3])
        ent = [
            {
                "question": question,
                "type": type_ent,
                "specification": specification,
                "name": name
            }
        ]
        data[intent_id]['entities'] = json.dumps(ent, ensure_ascii=False)


def load_answers(cursor, intent_id):
    sql_select_query = """select * from nlp_intent_answers where intentId = %s"""
    cursor.execute(sql_select_query, (intent_id,))
    data_answers = cursor.fetchall()

    if len(data_answers) == 0:
        data[intent_id]['answers'] = list()

    named_answers = list()
    default_answer = ""
    condition = ""
    for namedEnt in data_answers:
        if namedEnt[4] is not None:
            condition = namedEnt[4]
        if bool(namedEnt[3]):
            if condition != "":
                default_answer = str(namedEnt[2]).strip() + "__conditions__:" + condition
            else:
                default_answer = str(namedEnt[2]).strip()
        else:
            if condition != "":
                named_answers.append(str(namedEnt[2]).strip() + "__conditions__:" + condition)
            else:
                named_answers.append(str(namedEnt[2]).strip())

    data[intent_id]['answers'] = {
        "answers": named_answers,
        "default": default_answer
    }


def export_to_excel(path):
    book = Workbook()
    sheet = book.active

    headers = ['Эталон вопроса', 'Подтверждающий вопрос',
               'Синонимы', 'Переменные', 'Ответы', 'Ответ по умолчанию', 'Skill', 'SkillState', 'targetSkillPayload',
               'Knowledgebase']
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
        if array['short_question'] is not None:
            sheet['A' + str(a)] = str(array['short_question'])
        if array['confirmation_question'] is not None:
            sheet['B' + str(a)] = str(array['confirmation_question'])
        if array['synonyms'] is not None:
            sheet['C' + str(a)] = '|'.join(map(str, array['synonyms']))
        if array['entities'] is not None:
            sheet['D' + str(a)] = array['entities']
        if array['answers']['answers'] is not None:
            sheet['E' + str(a)] = '|'.join(map(str, array['answers']['answers']))
        if array['answers']['default'] is not None:
            sheet['F' + str(a)] = array['answers']['default']
        if array['skill'] is not None:
            sheet['G' + str(a)] = str(array['skill'])
        if array['skill_state'] is not None:
            sheet['H' + str(a)] = str(array['skill_state'])
        if array['target_skill_payload'] is not None:
            sheet['I' + str(a)] = ''.join(str(array['target_skill_payload']))
        if array['knowledge_base'] is not None:
            sheet['J' + str(a)] = ''.join(str(array['knowledge_base']))
        a += 1

    book.save(path)


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        'db_user',
        metavar='db-user',
        help='database user',
    )
    parser.add_argument(
        'db_password',
        metavar='db-password',
        help='database password',
    )
    parser.add_argument(
        'db_name',
        metavar='db-name',
        help='database name',
    )
    parser.add_argument(
        '--db-host',
        dest='db_host',
        help='database host (default "localhost")',
        default='localhost',
    )
    parser.add_argument(
        '--db-port',
        dest='db_port',
        help='database port (default 3306)',
        default=3306,
        type=int,
    )
    parser.add_argument(
        '--file',
        dest='export_path',
        help='path to export file (default "export.xlsx")',
        default="export.xlsx",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    db_connection = mariadb.connect(
        host=args.db_host,
        port=args.db_port,
        user=args.db_user,
        password=args.db_password,
        database=args.db_name,
    )
    load_data(db_connection)
    export_to_excel(args.export_path)
