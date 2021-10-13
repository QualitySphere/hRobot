#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import logging
import os

# import hrobot.hcore
import xlrd
import xlwt
import robot
from selenium import webdriver
import requests
import platform
import paramiko
import inspect
import json
import time
import re
from datetime import datetime


def timestamp_before_hour():
    return int(datetime.now().timestamp()) * 1000 - 3600000


class HRobot(object):
    def __init__(self):
        self.env = {
            "WORKDIR": os.path.abspath('.'),
            "HROBOT_PROJECT_FILE": ".hrobot",
            "TESTCASES_DIR": "testcases",
            "TESTCASES_FILE": "suites.xls",
            "KEYWORDS_DIR": "keywords",
            "KEYWORDS_FILE": "keywords.xls",
            "VARIABLES_DIR": "variables",
            "VARIABLES_FILE": "variables.xls",
            "ROBOT_DIR": os.path.basename(os.path.abspath('.')),
            "HROBOT_KEYWORDS_ROBOT_FILE": "hrobot.robot",
            "OUTPUT_DIR": "output",
        }
        self.book_style = xlwt.XFStyle()
        self.book_font = xlwt.Font()
        self.book_font.name = u'黑体'
        self.book_style.font = self.book_font

    def generate_testcase_xl(self, xl_file):
        book = xlwt.Workbook(encoding="utf-8")
        # 开始 定义 Sheet 用例
        sheet_case = book.add_sheet(sheetname=u'用例')
        sheet_case.write(0, 0, label=u'用例标题', style=self.book_style)
        sheet_case.write(0, 1, label=u'关键字类型', style=self.book_style)
        sheet_case.write(0, 2, label=u'关键字', style=self.book_style)
        sheet_case.write(0, 3, label=u'参数', style=self.book_style)
        # 完成 定义 Sheet 用例
        # 开始 定义 Sheet 变量
        sheet_variable = book.add_sheet(sheetname=u'变量')
        # 完成 定义 Sheet 变量
        # 开始 定义 Sheet 前置
        sheet_setup = book.add_sheet(sheetname=u'前置')
        # 完成 定义 Sheet 前置
        # 开始 定义 Sheet 后置
        sheet_teardown = book.add_sheet(sheetname=u'后置')
        # 完成 定义 Sheet 后置
        # 开始 定义 Sheet 内置关键字
        sheet_keyword = book.add_sheet(sheetname=u'内置关键字')
        sheet_keyword.write(0, 0, label=u'关键字类型', style=self.book_style)
        sheet_keyword.write(0, 1, label=u'关键字', style=self.book_style)
        keyword_row = 0
        # <开始提取内置关键字> 提取出 hRobot 中内置的关键字列表
        # robot_keywords = list()
        hkeywords = HKeywords()
        for hkw in hkeywords.__dir__():
            if not hkw.startswith('_'):
                keyword_row += 1
                keyword_name = u'%s' % inspect.getdoc(hkeywords.__getattribute__(hkw)).split(u'内置.')[1]
                sheet_keyword.write(keyword_row, 0, label=u'内置', style=self.book_style)
                sheet_keyword.write(keyword_row, 1, label=keyword_name, style=self.book_style)
                print(u'发现内置关键字 %s' % keyword_name)
                # keyword_args = inspect.getfullargspec(hkeywords.__getattribute__(hkw)).args[1:]
                # robot_keywords.append(keyword_name)
                # if len(keyword_args) == 0:
                #     robot_keywords.append(u'    %s' % hkw.replace('_', ' '))
                # else:
                #     robot_keywords.append(u'    [Arguments]    ${%s}' % '}    ${'.join(keyword_args))
                #     robot_keywords.append(u'    %s    ${%s}' % (
                #         hkw.replace('_', ' '),
                #         '}    ${'.join(keyword_args)
                #     ))
        # <提取完成>
        # 完成 定义 Sheet 内置关键字
        # suite_sheet.write(1, 1, label=xlwt.Formula(u'内置关键字!A2'), style=book_style)
        book.save(xl_file)

    def generate_variable_xl(self, xl_file):
        book = xlwt.Workbook(encoding="utf-8")
        sheet = book.add_sheet(sheetname=u'变量集')
        sheet.write(0, 0, label=u'变量名')
        sheet.write(0, 1, label=u'变量类型')
        sheet.write(0, 2, label=u'变量值')
        book.save(xl_file)

    def generate_keyword_xl(self, xl_file):
        book = xlwt.Workbook(encoding="utf-8")
        sheet = book.add_sheet(sheetname=u'关键字集')
        sheet.write(0, 0, label=u'关键字')
        sheet.write(0, 1, label=u'参数')
        book.save(xl_file)

    def init_project(self, cmd_args: dict):
        """
        初始化项目目录，若存在则终止
        :param: cmd_args
        :return:
        """
        project_path = os.path.join(self.env['WORKDIR'], cmd_args['project'])
        if os.path.exists(project_path):
            print(u"项目目录 %s 已经存在" % project_path)
            exit(1)
        os.mkdir(project_path)
        os.mkdir(os.path.join(project_path, self.env['TESTCASES_DIR']))
        os.mkdir(os.path.join(project_path, self.env['VARIABLES_DIR']))
        os.mkdir(os.path.join(project_path, self.env['KEYWORDS_DIR']))
        self.generate_testcase_xl(os.path.join(project_path, self.env['TESTCASES_DIR'], self.env['TESTCASES_FILE']))
        self.generate_variable_xl(os.path.join(project_path, self.env['VARIABLES_DIR'], self.env['VARIABLES_FILE']))
        self.generate_keyword_xl(os.path.join(project_path, self.env['KEYWORDS_DIR'], self.env['KEYWORDS_FILE']))
        with open(os.path.join(project_path, self.env['HROBOT_PROJECT_FILE']), 'w', encoding='utf-8') as f:
            f.write(cmd_args['project'])
        with open(os.path.join(project_path, '.gitignore'), 'w', encoding='utf-8') as f:
            f.write('\n'.join([
                '.DS_Store',
                '__pycache__/',
                '.pytest_cache__/',
                '.idea/',
                'robotframework/'
            ]))
        return True

    def xl_to_robot_case(self, xl_file):
        """
        excel 文件转换为 RobotFramework 用例文件 .robot
        :param xl_file:
        :return:
        """
        book = xlrd.open_workbook(xl_file)
        robot_file_name_prefix = os.path.basename(xl_file).split('.')[0]
        sheet_case = book.sheet_by_name(u'用例')
        # <开始表头解析> 第0行是表头，处理成字典格式，表头与列号的对应关系，好在后续用例解析的时候灵活使用
        sheet_header = dict()
        col_num = 0
        for col_name in sheet_case.row(rowx=0):
            sheet_header[col_name.value] = col_num
            col_num += 1
        # <完成表头解析>
        nrows = sheet_case.nrows
        robot_file = os.path.join(
            self.env['WORKDIR'],
            self.env['ROBOT_DIR'],
            self.env['TESTCASES_DIR'],
            '%s' % robot_file_name_prefix
        )
        robot_json = {
            'settings': {
                'documentation': robot_file,
                'resource': set(),
                'test_setup': set(),
                'test_teardown': set(),
            },
            'variables': {},
            'testcases': [],
            'keywords': set()
        }
        hrobot_keywords_file = os.path.join('..', self.env['KEYWORDS_DIR'], self.env['HROBOT_KEYWORDS_ROBOT_FILE'])
        robot_json['settings']['resource'].add(hrobot_keywords_file)

        # 第1行开始是测试用例数据
        for n in range(1, nrows):
            row_data = sheet_case.row(rowx=n)
            # Excel 表头是"用例标题"的列号单元格中数据
            case_title = row_data[sheet_header[u'用例标题']].value
            # 如果测试用例数据尚无记录，或者A列不为空且用例标题与记录中最后一个不同，就初始化一个新的用例记录，虽然可以简单粗暴的在 .robot 加空行，但似乎这样处理更美观，待后续再看看有无更好的方案
            if len(robot_json['testcases']) == 0 or \
                    len(case_title) != 0 and case_title != robot_json['testcases'][-1]['title']:
                print(u'发现测试用例 %s' % case_title)
                robot_json['testcases'].append({
                    'title': case_title,
                    'steps': []
                })
            # Excel 中表头是"关键字类型" + 表头是"关键字" 的单元格数据拼装出真正的关键字
            step_kw_type = row_data[sheet_header[u'关键字类型']].value
            step_kw_name = row_data[sheet_header[u'关键字']].value
            step_kw = '.'.join([step_kw_type, step_kw_name])
            # 如果关键字不是内置的，就需要在 .robot 开头导入自定义关键字库文件路径，在这里记录到 robot_json 中
            if step_kw_type != u'内置':
                robot_json['settings']['resource'].add(os.path.join('..', 'keywords', step_kw_type))
            # Excel 中表头从"参数"开始后面都是参数，添加到用例记录的最后一个用例中去
            step_args = list()
            for step_arg in row_data[sheet_header[u'参数']:]:
                step_args.append(str(step_arg.value))
            robot_json['testcases'][-1]['steps'].append('\t'.join([
                step_kw,
                '\t'.join(step_args)
            ]))

        robot_libs = 'Resource         '.join(robot_json['settings']['resource'])
        robot_settings = '\n'.join([
            '*** Settings ***',
            'Documentation    %s' % robot_json['settings']['documentation'],
            'Resource         %s' % robot_libs,
        ])
        robot_variables = '\n'.join([
            '*** Variables ***',
            '\n'
        ])
        robot_testcases = '\n'.join([
            '*** Test Cases ***',
        ])
        for tc in robot_json['testcases']:
            robot_steps = '\n\t'.join(tc['steps'])
            robot_testcases = '\n'.join([
                robot_testcases,
                tc['title'],
                '\t%s' % robot_steps
            ])
        robot_keywords = '\n'.join([
            '*** Keywords ***',
            '\n',
        ])
        robot_content = '\n'.join([
            robot_settings,
            robot_variables,
            robot_testcases,
            robot_keywords
        ])
        with open('%s.robot' % robot_file, 'w', encoding='utf-8') as f:
            f.write(robot_content)

    def cls_to_robot_builtin_keyword(self):
        """
        内置 hrobot.hcore.HKeywords 转换为 RobotFramework 关键字文件 .robot
        :return:
        """
        robot_file = os.path.join(
            self.env['WORKDIR'],
            self.env['ROBOT_DIR'],
            self.env['KEYWORDS_DIR'],
            self.env['HROBOT_KEYWORDS_ROBOT_FILE']
        )
        robot_keywords = list()
        hkeywords = HKeywords()
        for hkw in hkeywords.__dir__():
            if not hkw.startswith('_'):
                keyword_name = u'%s' % inspect.getdoc(hkeywords.__getattribute__(hkw))
                keyword_args = inspect.getfullargspec(hkeywords.__getattribute__(hkw)).args[1:]
                print(u'初始化内置关键字 %s %s' % (keyword_name, keyword_args))
                robot_keywords.append(keyword_name)
                if len(keyword_args) == 0:
                    robot_keywords.append(u'    %s' % hkw.replace('_', ' '))
                else:
                    robot_keywords.append(u'    [Arguments]    ${%s}' % '}    ${'.join(keyword_args))
                    robot_keywords.append(u'    %s    ${%s}' % (
                        hkw.replace('_', ' '),
                        '}    ${'.join(keyword_args)
                    ))
        robot_content = '\n'.join([
            '*** Settings ***',
            'Documentation    hRobot Keywords',
            'Library          hrobot.hcore.HKeywords',
            '\n',
            '*** Keywords ***',
            '\n'.join(robot_keywords),
            '\n',
        ])
        with open(robot_file, 'w', encoding='utf-8') as f:
            f.write(robot_content)

    def xl_to_robot_keyword(self):
        pass

    def xl_to_robot_variable(self):
        pass

    def run_project(self, cmd_args: dict):
        """
        执行项目测试，先转换成 RobotFramework 结构目录文件，然后调用 robot 执行
        :param: cmd_args
        :return:
        """
        if not os.path.exists(self.env["HROBOT_PROJECT_FILE"]):
            print(u'这不是一个 hRobot 项目目录')
            return False
        robot_path = os.path.join(self.env['WORKDIR'], self.env['ROBOT_DIR'])
        os.system('rm -rf %s' % robot_path)
        os.mkdir(robot_path)
        os.mkdir(os.path.join(robot_path, self.env['TESTCASES_DIR']))
        os.mkdir(os.path.join(robot_path, self.env['KEYWORDS_DIR']))
        os.mkdir(os.path.join(robot_path, self.env['VARIABLES_DIR']))
        self.cls_to_robot_builtin_keyword()
        for case_file in os.listdir(os.path.join(self.env['WORKDIR'], self.env['TESTCASES_DIR'])):
            self.xl_to_robot_case(os.path.join(self.env['WORKDIR'], self.env['TESTCASES_DIR'], case_file))
        allure_results_dir = os.path.join(robot_path, self.env['OUTPUT_DIR'], 'allure-results')
        robot.run(
            self.env['ROBOT_DIR'],
            consolewidth=80,
            consolecolors='on',
            outputdir=os.path.join(robot_path, self.env['OUTPUT_DIR']),
            listener='hrobot.Listener.allure_robotframework;%s' % allure_results_dir,
            reporttitle='Hybrid Robot Report',
            variablefile=os.listdir(os.path.join(robot_path, self.env['VARIABLES_DIR']))
        )

    def generate_report(self):
        if not os.path.exists(self.env["HROBOT_PROJECT_FILE"]):
            print(u'这不是一个 hRobot 项目目录')
            return False
        allure_results_path = os.path.join(
            self.env['WORKDIR'],
            self.env['ROBOT_DIR'],
            self.env['OUTPUT_DIR'],
            'allure-results'
        )
        if not os.path.exists(allure_results_path):
            print(u'尚未发现测试用例执行记录，你可以尝试使用 hrobot run 来执行测试用例')
            return False
        os.system('allure generate %s -o report --clean' % allure_results_path)
        os.system('allure open -p 80 report')

    def debug_project(self):
        pass


class HKeywords(object):
    def __init__(self):
        self.__webdriver = None
        self.__session = None
        self.__response = None
        self.__ssh_output = None
        self.__variables = dict()

    def __smart_content(self, content):
        print(u'源内容: %s' % content)
        var_pattern = re.compile("{{[a-zA-Z0-9 _-]*}}")
        for var_key in var_pattern.findall(content):
            var_value = self.__variables[var_key.strip('{{').strip('}}').strip()]
            print('Replace %s to %s' % (var_key, var_value))
            content = content.replace(var_key, str(var_value))
        fun_pattern = re.compile("{%[a-zA-Z0-9 ()_.*/+-]*%}")
        for fun_name in fun_pattern.findall(content):
            fun_result = eval(fun_name.strip('{%').strip('%}').strip())
            print('Replace %s to %s' % (fun_name, fun_result))
            content = content.replace(fun_name, str(fun_result))
        print(u'渲染后: %s' % content)
        return content

    def kw_timestamp(self):
        """内置.时间戳"""
        return True

    def kw_assert(self, key1, assert_key, key2):
        """内置.断言"""
        if assert_key == "==":
            assert key1 == key2
        elif assert_key == ">=":
            assert key1 >= key2

    def kw_def_var(self, key, value):
        """内置.变量.赋值"""
        var_value = self.__smart_content(value)
        print('%s=%s' % (key, var_value))
        self.__variables[key] = var_value
        return True

    def kw_sleep(self, seconds):
        """内置.休眠"""
        print(u'休眠 %s 秒' % seconds)
        if type(seconds) is str:
            time.sleep(int(float(seconds)))
        else:
            time.sleep(int(seconds))
        return True

    def kw_print(self, content):
        """内置.输出"""
        print(content)
        return True

    def kw_webdriver_open(self):
        """内置.浏览器.启动"""
        if self.__webdriver:
            return True
        opts = webdriver.ChromeOptions()
        opts.add_argument('-lang=zh-cn')
        opts.add_argument('--ignore-certificate-errors')
        if platform.system() == 'Linux':
            width = 1920
            height = 1200
            opts.add_argument('--head-less')
            opts.add_argument('--no-sandbox')
            self.__webdriver = webdriver.Chrome(options=opts)
            self.__webdriver.set_window_size(width, height)
        else:
            self.__webdriver = webdriver.Chrome(options=opts)
            self.__webdriver.maximize_window()
        return True

    def kw_webdriver_close(self):
        """内置.浏览器.关闭"""
        if self.__webdriver:
            self.__webdriver.close()
        return True

    def kw_webdriver_access(self, url):
        """内置.浏览器.访问"""
        self.__webdriver.get(url)
        return True

    def kw_webdriver_find(self, xpath):
        """内置.浏览器.查找"""
        pass

    def kw_webdriver_click(self, xpath):
        """内置.浏览器.点击"""
        element = self.__webdriver.find_element_by_xpath(xpath)
        element.click()
        return True

    def kw_webdriver_input(self, xpath, text):
        """内置.浏览器.输入"""
        element = self.__webdriver.find_element_by_xpath(xpath)
        element.send_keys(text)
        return True

    def kw_request_open(self):
        """内置.接口.开启会话"""
        if self.__session:
            return True
        self.__session = requests.session()
        return True

    def kw_request_close(self):
        """内置.接口.关闭会话"""
        if self.__session:
            self.__session.close()
        return True

    def kw_request_get(self, url, headers, params):
        """内置.接口.GET"""
        url = self.__smart_content(url)
        headers = json.loads(self.__smart_content(headers))
        params = json.loads(self.__smart_content(params))
        self.kw_request_open()
        if headers is None:
            headers = {
                "Content-Type": "Application/json"
            }
        self.__response = self.__session.get(url=url, headers=headers, params=params, verify=False)
        print(self.__response.content)
        return True

    def kw_request_post(self, url, headers, body):
        """内置.接口.POST"""
        url = self.__smart_content(url)
        print('URL: %s' % url)
        headers = json.loads(self.__smart_content(headers))
        print('Headers: %s' % headers)
        body = json.loads(self.__smart_content(body))
        print('Body: %s' % body)
        self.kw_request_open()
        self.__response = self.__session.post(url=url, headers=headers, json=body, verify=False)
        print(self.__response.content)
        return True

    def __get_response_smart_value(self, smart_key):
        if smart_key.startswith('status_code'):
            smart_value = self.__response.status_code
        elif smart_key.startswith('body'):
            smart_value = self.__response.json()
        elif smart_key.startswith("headers"):
            smart_value = self.__response.headers
        elif smart_key.startswith("cookies"):
            smart_value = self.__response.cookies
        else:
            print(u'指定的检查路径 %s 不存在' % smart_key)
            raise KeyError
        for sk in smart_key.split('.')[1:]:
            try:
                smart_value = smart_value[sk]
            except TypeError:
                smart_value = smart_value[int(sk)]
            except KeyError:
                print(u'指定的检查路径 %s 下的 %s 不存在' % (smart_key, sk))
                raise KeyError
        return smart_value

    def kw_response_get_value(self, smart_key, var_name):
        """内置.接口.响应.取值"""
        print(u'获取 %s 的值并赋予 %s' % (smart_key, var_name))
        self.__variables[var_name] = self.__get_response_smart_value(smart_key)
        print(u'%s=%s' % (var_name, self.__variables[var_name]))
        return True

    def kw_response_assert(self, smart_key, assert_key, expected_value):
        """内置.接口.响应.断言"""
        print(u'检查 %s 是否符合预期值 %s' % (smart_key, expected_value))
        smart_value = self.__get_response_smart_value(smart_key)
        if type(smart_value) is int:
            smart_expected_value = int(float(expected_value))
        elif type(smart_value) is str:
            smart_expected_value = str(expected_value)
        else:
            smart_expected_value = expected_value
        if assert_key.lower() in ['=', '==', u'等于']:
            assert smart_value == smart_expected_value, u"预期:%s \n实际:%s" % (expected_value, smart_value)
        elif assert_key.lower() in ['in', u'被包含']:
            assert smart_value in smart_expected_value, u"预期:%s \n实际:%s" % (expected_value, smart_value)
        elif assert_key.lower() in ['not in', u'不包含']:
            assert smart_value not in smart_expected_value, u"预期:%s \n实际:%s" % (expected_value, smart_value)
        elif assert_key.lower() in ['contains', 'contain', u'包含']:
            assert smart_expected_value in smart_value, u"预期:%s \n实际:%s" % (expected_value, smart_value)
        elif assert_key.lower() in ['contains', 'contain', u'不包含']:
            assert smart_expected_value not in smart_value, u"预期:%s \n实际:%s" % (expected_value, smart_value)
        else:
            raise KeyError(u'不支持的断言关系 %s' % assert_key)

    def kw_ssh_exec(self, host, user, password, cmd):
        """内置.远程.执行"""
        ssh = paramiko.SSHClient()
        ssh_output = {
            "stdin": '',
            "stdout": '',
            "stderr": '',
        }
        self.__ssh_output = ssh_output['stdout']
        try:
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            print('%s@%s %s' % (user, host, cmd))
            ssh.connect(
                hostname=host,
                port=22,
                username=user,
                password=password
            )
            stdin, stdout, stderr = ssh.exec_command(cmd)
            __ssh_error = stderr.read().decode('utf-8')
            assert not __ssh_error, u'远程执行命令有误:\n%s' % __ssh_error
            self.__ssh_output = stdout.read().decode('utf-8')
        finally:
            ssh.close()
        print(self.__ssh_output)
        return True


if __name__ == '__main__':
    print('This is hRobot Core')
