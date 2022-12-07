#!/usr/bin/env python3

import sys
import os
import json
import logging
import logging.handlers
import xmltodict
import xlsxwriter

def main():

    log_dir = "logs"
    log_file = "%s/tmpl2xlsx.log" % log_dir
    tmpl_dir = "templates"
    xlsx_dir = "excel"

    tmpl_files = os.listdir(tmpl_dir)

    excelfile = '%s/items.xlsx' % xlsx_dir

    #for tmpl_file in tmpl_files:
    infile = "%s/%s" % (tmpl_dir, tmpl_files)

    workbook = xlsxwriter.Workbook(excelfile)

    for tmpl_file in tmpl_files:
        infile = "%s/%s" % (tmpl_dir, tmpl_file)
        tmpl2xlsx(infile, xlsx_dir, workbook)
    workbook.close()

def tmpl2xlsx(infile, xlsx_dir, workbook):
    with open(infile, 'r', encoding="utf-8") as f:
        file_content = f.read()
        f.close

# Описываем шрифты ----=====НАЧАЛО=====------

    bold = workbook.add_format(
        {
            'bold': True
            }
        )
    item_header = workbook.add_format(
        {
            'bg_color': '#CCCCCC',
            'bold': True,
            'border': 1
            }
        )
    item_text = workbook.add_format(
        {
            'border': 1
            }
        )
    trigger_severity = {
        '0': workbook.add_format({'bg_color': '#97AAB3'}),
        '1': workbook.add_format({'bg_color': '#7499FF'}),
        '2': workbook.add_format({'bg_color': '#FFC859'}),
        '3': workbook.add_format({'bg_color': '#FFA059'}),
        '4': workbook.add_format({'bg_color': '#E97659'}),
        '5': workbook.add_format({'bg_color': '#E45959'}),
        # '0' : workbook.add_format({'bg_color' : '#97AAB3'}),
        'INFO': workbook.add_format({'bg_color': '#7499FF'}),
        'WARNING': workbook.add_format({'bg_color': '#FFC859'}),
        'AVERAGE': workbook.add_format({'bg_color': '#FFA059'}),
        'HIGH': workbook.add_format({'bg_color': '#E97659'}),
        'DISASTER': workbook.add_format({'bg_color': '#E45959'})
    }
        
# Описываем шрифты ----=====КОНЕЦ=====------

    content = xmltodict.parse(file_content)
    pcontent = json.dumps(content, indent=4, separators=(',', ': '))
    content = json.loads(pcontent)
    templates = content['zabbix_export']['templates']['template']
    if type(templates) is dict:
        templates = [templates]
    
    n = 0
    t = 0

    for template in templates:
        max_length_A = 0
        max_length_B = 0
        max_length_C = 0
        try:
            worksheet = workbook.add_worksheet(template['template'][0:30])
        except Exception as ex:

            continue

        worksheet.write('A1', template['name'], bold)
        worksheet.write('A2', template.get('description', 'No description'))

        worksheet.write('A4', "Items:", bold)
        worksheet.write('A5', "Name", item_header)
        worksheet.write('B5', "Key", item_header)
        worksheet.write('C5', "Description", item_header)
        worksheet.write('E5', "Trigger name", item_header)

        worksheet.merge_range('A6:C6', "System items", item_header)

        index = 7

        for items_c in (content['zabbix_export']['templates']['template']['items']['item']):
            t = t + 1
#            worksheet.write('A%s' % index, items_c['name'], item_text)               # Записываем метрику В excel

            if "triggers" in items_c:                                                # Условие при котором у метрики есть триггер
                worksheet.write('A%s' % index, items_c['name'], item_text)
                if type(items_c['triggers']['trigger']) is dict:
                    worksheet.write('E%s' % index, items_c['triggers']['trigger']['name'], item_text)
                    index = index + 1

                else:
                    worksheet.write('A%s' % index, items_c['name'], item_text)
                    type(items_c['triggers']['trigger']) is list
                    triggers_circle = (items_c['triggers']['trigger'])
                    for triggers_c in triggers_circle:
                        worksheet.write('E%s' % index, triggers_c['name'], item_text)
                        index = index + 1

            else:                                                                    # Условие при котором у метрики отсутствует триггер
                worksheet.write('A%s' % index, items_c['name'], item_text)
                n = n + 1
                index = index + 1


        templates = content['zabbix_export']['templates']['template']

        #### PROTOTIPES ####

        if 'discovery_rules' in template:
            if template['discovery_rules'] is None:
                continue
            drs = template['discovery_rules']['discovery_rule']
            if type(drs) is dict:                                                                   ##### БУДЕТ DICT если есть только одно правило обнаружения ####
                drs = [drs]

                for dr in drs:
                    worksheet.merge_range('A%s:C%s' % (index, index),dr['name'], item_header)       ##### Вписываем название правила обнаружения ####
                    index = index + 1

                    if 'item_prototypes' in dr:                                                     ##### Условие. Если есть 'item_prototypes' то .. ####
                        items = dr['item_prototypes']['item_prototype']
                        if type(items) is dict:                                                     ##### БУДЕТ DICT если есть только один item прототип ####

                            worksheet.write('A%s' % index, items['name'], item_text)                ##### Заполняем поля name, key, description для метрики ####
                            worksheet.write('B%s' % index, items['key'], item_text)
                            worksheet.write('C%s' % index, items.get('description', ''), item_text)
                            max_length_A = max(max_length_A, len(items['name']))
                            max_length_B = max(max_length_B, len(items['key']))
                            max_length_C = max(
                                max_length_C,
                                len(str(items.get('description', '')))
                                )

                            if 'trigger_prototypes' in items:                                           #### Если есть прототипы триггеров, записываем их в файл ####
                                triggers_p_circle = items['trigger_prototypes']['trigger_prototype']
                                for trigger_p_c in triggers_p_circle:
                                    worksheet.write('E%s' % index, trigger_p_c['name'], item_text)
                                    index = index + 1

                        else:                                                                           ##### БУДЕТ LIST если несколько item прототип ####

                            for items_circle in items:
                                worksheet.write('A%s' % index, items_circle['name'], item_text)                ##### Заполняем поля name, key, description для метрики ####
                                worksheet.write('B%s' % index, items_circle['key'], item_text)
                                worksheet.write('C%s' % index, items_circle.get('description', ''), item_text)
                                max_length_A = max(max_length_A, len(items_circle['name']))
                                max_length_B = max(max_length_B, len(items_circle['key']))
                                max_length_C = max(
                                max_length_C,
                                len(str(items_circle.get('description', '')))
                                )
                                if 'trigger_prototypes' in items_circle:                                           #### Если есть прототипы триггеров, записываем их в файл ####
                                    triggers_p_circle_n = items_circle['trigger_prototypes']['trigger_prototype']
                                    if type(triggers_p_circle_n) is dict:

                                        worksheet.write('E%s' % index, triggers_p_circle_n['name'], item_text)
                                        index = index + 1
                                    else:
                                        for trigger_p_c in triggers_p_circle_n:
                                            worksheet.write('E%s' % index, trigger_p_c['name'], item_text)
                                            index = index + 1

                                else:
                                    index = index + 1

            else:                                                                                        # Если в дискавери больше одного правила
                print(type(drs))
                for dr in drs:
                    print(dr)
                    print()
                    worksheet.merge_range('A%s:C%s' % (index, index),dr['name'], item_header)
                    index = index + 1
                    
                    if 'item_prototypes' in dr:                                                     ##### Условие. Если есть 'item_prototypes' то .. ####
                        items = dr['item_prototypes']['item_prototype']
                        if type(items) is dict:                                                     ##### БУДЕТ DICT если есть только один item прототип ####

                            worksheet.write('A%s' % index, items['name'], item_text)                ##### Заполняем поля name, key, description для метрики ####
                            worksheet.write('B%s' % index, items['key'], item_text)
                            worksheet.write('C%s' % index, items.get('description', ''), item_text)
                            max_length_A = max(max_length_A, len(items['name']))
                            max_length_B = max(max_length_B, len(items['key']))
                            max_length_C = max(
                                max_length_C,
                                len(str(items.get('description', '')))
                                )

                            if 'trigger_prototypes' in items:                                           #### Если есть прототипы триггеров, записываем их в файл ####
                                triggers_p_circle = items['trigger_prototypes']['trigger_prototype']
                                for trigger_p_c in triggers_p_circle:
                                    worksheet.write('E%s' % index, trigger_p_c['name'], item_text)
                                    index = index + 1
                            
                            else:
                                index = index + 1

                        else:                                                                           ##### БУДЕТ LIST если несколько item прототип ####

                            for items_circle in items:
                                worksheet.write('A%s' % index, items_circle['name'], item_text)                ##### Заполняем поля name, key, description для метрики ####
                                worksheet.write('B%s' % index, items_circle['key'], item_text)
                                worksheet.write('C%s' % index, items_circle.get('description', ''), item_text)
                                max_length_A = max(max_length_A, len(items_circle['name']))
                                max_length_B = max(max_length_B, len(items_circle['key']))
                                max_length_C = max(
                                max_length_C,
                                len(str(items_circle.get('description', '')))
                                )
                                if 'trigger_prototypes' in items_circle:                                           #### Если есть прототипы триггеров, записываем их в файл ####
                                    triggers_p_circle_n = items_circle['trigger_prototypes']['trigger_prototype']
                                    if type(triggers_p_circle_n) is dict:
                                        worksheet.write('E%s' % index, triggers_p_circle_n['name'], item_text)
                                        index = index + 1
                                    else:
                                        for trigger_p_c in triggers_p_circle_n:
                                            worksheet.write('E%s' % index, trigger_p_c['name'], item_text)
                                            index = index + 1

                                else:
                                    index = index + 1

if __name__ == "__main__":
    main()