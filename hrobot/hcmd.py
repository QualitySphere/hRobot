#!/usr/bin/env python3
# -*- coding: utf-8 -*-


import os
import sys
import getopt
# import optparse
from hrobot import hcore

cmd_list = ['init', 'run', 'report']


def help_doc_cmd(sys_args):
    help_doc = '\n'.join([
        os.path.basename(sys_args[0]),
        u'  init     初始化项目目录',
        u'  run      执行测试用例',
        u'  report   生成并展示测试报告',
    ])
    print(help_doc)
    exit(1)


def help_doc_init(sys_args):
    help_doc = '\n'.join([
        '%s init' % os.path.basename(sys_args[0]),
        u'  -p\tproject',
    ])
    print(help_doc)
    exit(1)


def help_doc_run(sys_args):
    help_doc = '\n'.join([
        '%s run' % os.path.basename(sys_args[0]),
        u'  -c\ttestcase',
        u'  -v\tvariable',
        u'  -k\tkeyword',
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


def cmd_run(sys_args):
    _args = {
        "suite": None,
        "case": None,
        "tag": None,
        "debug": False,
    }
    try:
        for opt, arg in getopt.getopt(sys_args[2:], "s:c:t:d")[0]:
            if opt == '-s':
                _args['suite'] = arg
            elif opt == '-c':
                _args['case'] = arg
            elif opt == '-t':
                _args['tag'] = arg
            elif opt == '-d':
                _args['debug'] = True
    except getopt.GetoptError:
        help_doc_run(sys_args)
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
        cmd_run(sys.argv)
    elif _cmd == 'report':
        cmd_report()


if __name__ == '__main__':
    print('This is hRobot Command Line')
