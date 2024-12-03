#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import json
import logging
import os
from xmind2testcase.utils import (
    get_xmind_testcase_list,
    get_absolute_path,
    dict_list_to_excel,
)
from xmind2testcase.config import excel_dropdown, excel_header, merge_header

"""
Convert XMind fie to Excel file 

表头字段：
1.用例编号
2.一级功能模块
3.二级功能模块
4.三级功能模块
5.优先级，下拉列表，值：Z、A、B、C
6.用例标题
7.前置条件
8.操作步骤
9.预期结果
10.测试结果，下拉列表，值：PASS、NG、阻塞、未执行
11.JIRA号
12.编写人
13.执行人
14.备注
"""


def xmind_to_excel_file(xmind_file):
    """Convert XMind file to a excel file"""
    xmind_file = get_absolute_path(xmind_file)
    logging.info("Start converting XMind file(%s) to excel file...", xmind_file)
    testcases = get_xmind_testcase_list(xmind_file)

    # 写入时，自动添加表头
    excel_testcase_rows = []
    for testcase in testcases:
        row = gen_a_testcase_row(testcase)
        # header 和 row 大小一致
        row_dict = {}
        for index, header in enumerate(excel_header):
            row_dict[header] = row[index] if len(row) >= index + 1 else ""
        excel_testcase_rows.append(row_dict)

    excel_file = xmind_file[:-6] + ".xlsx"
    if os.path.exists(excel_file):
        os.remove(excel_file)
        # logging.info('The excel_file already exists, return it directly: %s', excel_file)
        # return excel_file

    # 写入到 excel 中
    dict_list_to_excel(
        excel_testcase_rows,
        excel_file,
        dropdown_fields=excel_dropdown.keys(),
        dropdown_options=excel_dropdown,
        merge_fields=merge_header,
    )
    logging.info(
        "Convert XMind file(%s) to a excel file(%s) successfully!",
        xmind_file,
        excel_file,
    )

    return excel_file


def gen_a_testcase_row(testcase_dict):
    case_first_module = gen_case_module(testcase_dict["suite"])
    case_sencond_module = gen_case_module(testcase_dict["second_suite"])
    case_third_module = gen_case_module(testcase_dict["third_suite"])
    case_priority = testcase_dict["importance"]
    case_title = testcase_dict["name"]
    case_precontion = testcase_dict["preconditions"]
    case_step, case_expected_result = gen_case_step_and_expected_result(
        testcase_dict["steps"]
    )
    case_test_result = testcase_dict["result"]
    # JIRA 号为空
    case_jira = ""
    case_writer = testcase_dict["writer"]
    # 执行人为空，暂不支持
    case_executor = testcase_dict["executor"]
    # 备注为空，暂不支持
    case_comment = ""
    # TODO 变成 dict
    row = [
        "=ROW()-1",
        case_first_module,
        case_sencond_module,
        case_third_module,
        case_priority,
        case_title,
        case_precontion,
        case_step,
        case_expected_result,
        case_test_result,
        case_jira,
        case_writer,
        case_executor,
        case_comment,
    ]
    return row


def gen_case_module(module_name):
    if module_name:
        module_name = module_name.replace("（", "(")
        module_name = module_name.replace("）", ")")
    else:
        module_name = "/"
    return module_name


def gen_case_step_and_expected_result(steps):
    case_step = ""
    case_expected_result = ""

    for step_dict in steps:
        case_step += (
            str(step_dict["step_number"])
            + ". "
            + step_dict["actions"].replace("\n", "").strip()
            + "\n"
        )
        case_expected_result += (
            str(step_dict["step_number"])
            + ". "
            + step_dict["expectedresults"].replace("\n", "").strip()
            + "\n"
            if step_dict.get("expectedresults", "")
            else ""
        )
    # 去掉最后的换行符
    return case_step[:-1], case_expected_result[:-1]


if __name__ == "__main__":
    xmind_file = "../docs/zentao_testcase_template.xmind"
    excel_file = xmind_to_excel_file(xmind_file)
    print("Conver the xmind file to a excel file succssfully: %s", excel_file)
