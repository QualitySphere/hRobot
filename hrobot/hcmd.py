#!/usr/bin/env python3
# -*- coding: utf-8 -*-


import os
import sys
import getopt
# import optparse
from hrobot import hcore

cmd_list = [
    'init',
    'run',
    'debug',
    'report',
    'version',
]


def help_doc_cmd(sys_args):
    help_doc = '\n'.join([
        os.path.basename(sys_args[0]),
        u'  init     初始化项目目录',
        u'  run      执行测试用例',
        u'  debug    调试测试用例，支持选择用例集、测试用例、标签',
        u'  report   生成并展示测试报告',
        u'  version  显示版本信息',
    ])
    print(help_doc)
    exit(1)


def help_doc_init(sys_args):
    help_doc = '\n'.join([
        '%s init' % os.path.basename(sys_args[0]),
        u'  -p    project  定义一个测试项目目录的名称',
    ])
    print(help_doc)
    exit(1)


def help_doc_debug(sys_args):
    help_doc = '\n'.join([
        '%s debug' % os.path.basename(sys_args[0]),
        u'  -s    suite  测试用例集，测试用例 Excel 文件的文件名（不包含 .xlsx 后缀）',
        u'  -c    case   测试用例, 测试用例 Excel 文件中的 "用例" Sheet 中 "用例标题" 列的单元格内容',
        u'  -t    tag    标签, 测试用例 Excel 文件中的 "用例" Sheet 中 "标签" 列的单元格内容',
    ])
    print(help_doc)
    exit(1)


def cmd_init(sys_args):
    _args = {
        "project": "",
    }
    try:
        for opt, arg in getopt.getopt(sys_args[2:], "p:")[0]:
            if opt == '-p':
                _args['project'] = arg
    except getopt.GetoptError:
        help_doc_init(sys_args)
    if not _args['project']:
        help_doc_init(sys_args)
    _robot = hcore.HRobot()
    _robot.init_project(_args)


def cmd_run_full():
    _args = {}
    _robot = hcore.HRobot()
    _robot.run_project(_args)


def cmd_run(sys_args):
    _args = {
        "suite": None,
        "case": None,
        "tag": None,
    }
    try:
        for opt, arg in getopt.getopt(sys_args[2:], "s:c:t:")[0]:
            if opt == '-s':
                _args['suite'] = arg
            elif opt == '-c':
                _args['case'] = arg
            elif opt == '-t':
                _args['tag'] = arg
    except getopt.GetoptError:
        help_doc_debug(sys_args)
    if not _args['suite'] and not _args['case'] and not _args['tag']:
        help_doc_debug(sys_args)
    _robot = hcore.HRobot()
    _robot.run_project(_args)


def cmd_report():
    _robot = hcore.HRobot()
    _robot.generate_report()


def main():
    if len(sys.argv[1:]) == 0:
        help_doc_cmd(sys.argv)
    _cmd = sys.argv[1]
    if _cmd not in cmd_list:
        help_doc_cmd(sys.argv)
    if _cmd == 'init':
        cmd_init(sys.argv)
    elif _cmd == 'run':
        cmd_run_full()
    elif _cmd == 'debug':
        cmd_run(sys.argv)
    elif _cmd == 'report':
        cmd_report()
    elif _cmd == 'version':
        os.system('%s -m pip show hrobot' % sys.executable)


if __name__ == '__main__':
    print(u'这是 Hybrid Robot 命令行工具 hrobot 的代码')
