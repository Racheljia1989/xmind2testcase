#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import csv
import logging
import os
from xmind2testcase.utils import get_xmind_testcase_list, get_absolute_path

"""
Convert XMind fie to Zentao testcase csv file 

Zentao official document about import CSV testcase file: https://www.zentao.net/book/zentaopmshelp/243.mhtml 
"""


def xmind_to_zentao_csv_file(xmind_file):
    """Convert XMind file to a zentao csv file"""
    xmind_file = get_absolute_path(xmind_file)
    logging.info('Start converting XMind file(%s) to zentao file...', xmind_file)
    testcases = get_xmind_testcase_list(xmind_file)

    fileheader = ["用例ID（选填，不填将新建，否则会更新旧用例）","用例名称（必填，不允许重名）","测试类型（必填，可取值：功能测试、性能效率测试、"
                                                                                         "可维护性测试、可移植性测试、安全性测试、兼容性测试、易用性测试、可靠性测试）",
                  "级别（必填，可取值：P1、P2、P3、P4）","评审状态（必填，可取值：草稿、待评审、通过、不通过）","关联需求（选填）",
                  "功能模块（选填）","前置条件（选填）","步骤描述（选填，多个步骤间用双空行分隔）","预期结果（选填，多个步骤结果间用双空行分隔）","备注（选填）"]
    #fileheader = ["所属模块", "用例标题", "前置条件", "步骤", "预期", "关键词", "优先级", "用例类型", "适用阶段"]
    zentao_testcase_rows = [fileheader]
    for testcase in testcases:
        row = gen_a_testcase_row(testcase)
        zentao_testcase_rows.append(row)

    zentao_file = xmind_file[:-6] + '.csv'
    if os.path.exists(zentao_file):
        os.remove(zentao_file)
        # logging.info('The zentao csv file already exists, return it directly: %s', zentao_file)
        # return zentao_file

    with open(zentao_file, 'w', encoding='utf8') as f:
        writer = csv.writer(f)
        writer.writerows(zentao_testcase_rows)
        logging.info('Convert XMind file(%s) to a zentao csv file(%s) successfully!', xmind_file, zentao_file)

    return zentao_file


def gen_a_testcase_row(testcase_dict):
    case_number = ''
    case_review_status = '草稿'
    case_related_requirements = ''
    case_module = gen_case_module(testcase_dict['suite'])
    case_title = testcase_dict['name']
    case_precontion = testcase_dict['preconditions']
    case_step, case_expected_result = gen_case_step_and_expected_result(testcase_dict['steps'])
    case_keyword = ''
    case_priority = gen_case_priority(testcase_dict['importance'])
    case_type = gen_case_type(testcase_dict['execution_type'])
    case_apply_phase = '迭代测试'
    row = [case_number, case_title, case_type, case_priority, case_review_status, case_related_requirements,
           case_module,  case_precontion, case_step, case_expected_result,  case_apply_phase]
    #row = [case_module, case_title, case_precontion, case_step, case_expected_result, case_keyword, case_priority, case_type, case_apply_phase]
    return row


def gen_case_module(module_name):
    if module_name:
        module_name = module_name.replace('（', '(')
        module_name = module_name.replace('）', ')')
    else:
        module_name = '/'
    return module_name


def gen_case_step_and_expected_result(steps):
    case_step = ''
    case_expected_result = ''

    for step_dict in steps:
        case_step += str(step_dict['step_number']) + '. ' + step_dict['actions'].replace('\n', '').strip() + '\n'
        case_expected_result += str(step_dict['step_number']) + '. ' + \
            step_dict['expectedresults'].replace('\n', '').strip() + '\n' \
            if step_dict.get('expectedresults', '') else ''

    return case_step, case_expected_result


def gen_case_priority(priority):
    mapping = {1: 'P1', 2: 'P2', 3: 'P3', 4:'P4'}
    if priority in mapping.keys():
        return mapping[priority]
    else:
        return 'P2'


def gen_case_type(case_type):
    mapping = {1: '功能测试', 2: '性能效率测试', 3: '可维护性测试', 4:'可移植性测试', 5:'安全性测试',6:'兼容性测试', 7:'易用性测试', 8:'可靠性测试'}
    # mapping = {1: '手动', 2: '自动'}
    if case_type in mapping.keys():
        return mapping[case_type]
    else:
        return '功能测试'




# if __name__ == '__main__':
#     xmind_file = '../docs/zentao_testcase_template.xmind'
#     zentao_csv_file = xmind_to_zentao_csv_file(xmind_file)
#     print('Conver the xmind file to a zentao csv file succssfully: %s', zentao_csv_file)
