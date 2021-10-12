#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Author: v.stone@163.com
"""
RobotFramework Listener 的使用指南：
http://robotframework.org/robotframework/latest/RobotFrameworkUserGuide.html#listener-interface

Allure RobotFramework 插件的 GitHub 仓库：
https://github.com/allure-framework/allure-python/tree/master/allure-robotframework

Allure RobotFramework 插件的 RF Listener 源码
https://github.com/allure-framework/allure-python/blob/master/allure-robotframework/src/listener/robot_listener.py

这里直接复制上述监听器内容，加以修改使得在 CLI 界面上也能看到日志
"""


import os
import allure_commons

from allure_commons.lifecycle import AllureLifecycle
from allure_commons.logger import AllureFileLogger
from allure_robotframework.allure_listener import AllureListener
from allure_robotframework.types import RobotKeywordType


DEFAULT_OUTPUT_PATH = os.path.join('output', 'allure-results')


# noinspection PyPep8Naming
class allure_robotframework(object):
    ROBOT_LISTENER_API_VERSION = 2

    def __init__(self, logger_path=DEFAULT_OUTPUT_PATH):
        self.messages = Messages()

        self.logger = AllureFileLogger(logger_path)
        self.lifecycle = AllureLifecycle()
        self.listener = AllureListener(self.lifecycle)

        allure_commons.plugin_manager.register(self.logger)
        allure_commons.plugin_manager.register(self.listener)

    def start_suite(self, name, attributes):
        self.messages.start_context()
        self.listener.start_suite_container(name, attributes)

    def end_suite(self, name, attributes):
        self.messages.stop_context()
        self.listener.stop_suite_container(name, attributes)

    def start_test(self, name, attributes):
        self.messages.start_context()
        self.listener.start_test_container(name, attributes)
        self.listener.start_test(name, attributes)
        """以下是修改的地方"""
        print(u'\n开始执行')

    def end_test(self, name, attributes):
        messages = self.messages.stop_context()
        self.listener.stop_test(name, attributes, messages)
        self.listener.stop_test_container(name, attributes)
        """以下是修改的地方"""
        print(u'执行完毕')

    def start_keyword(self, name, attributes):
        self.messages.start_context()
        keyword_type = attributes.get('type')
        # Todo fix value assign
        keyword_name = '{} = {}'.format(attributes.get('assign')[0], name) if attributes.get('assign') else name
        if keyword_type == RobotKeywordType.SETUP:
            self.listener.start_before_fixture(keyword_name)
        elif keyword_type == RobotKeywordType.TEARDOWN:
            self.listener.start_after_fixture(keyword_name)
        else:
            self.listener.start_keyword(name)

    def end_keyword(self, _, attributes):
        messages = self.messages.stop_context()
        keyword_type = attributes.get('type')
        if keyword_type == RobotKeywordType.SETUP:
            self.listener.stop_before_fixture(attributes, messages)
        elif keyword_type == RobotKeywordType.TEARDOWN:
            self.listener.stop_after_fixture(attributes, messages)
        else:
            self.listener.stop_keyword(attributes, messages)

    def log_message(self, message):
        self.messages.push(message)
        """以下是修改的地方"""
        if message['html'] == 'no':
            print('[%s] %s | %s' % (
                message['timestamp'].split('.')[0],
                message['level'],
                message['message'].replace('\n', '\n[%s] %s | ' % (
                    message['timestamp'].split('.')[0],
                    message['level']
                ))
            ))

    def close(self):
        for plugin in [self.logger, self.listener]:
            name = allure_commons.plugin_manager.get_name(plugin)
            allure_commons.plugin_manager.unregister(name=name)
        """以下是修改的地方"""
        if str(os.getenv('RF_REPORT_TYPE')).lower() in ['local-allure', 'jenkins-allure']:
            print('Allure:\t %s' % os.path.join(os.path.abspath('.'), DEFAULT_OUTPUT_PATH))
        if str(os.getenv('RF_REPORT_TYPE')).lower() in ['jenkins-allure']:
            print('Report:\t %s/allure' % os.getenv('BUILD_URL'))


class Messages(object):
    def __init__(self):
        self._stack = []

    def start_context(self):
        self._stack.append([])

    def stop_context(self):
        return self._stack.pop() if self._stack else list()

    def push(self, message):
        self._stack[-1].append(message) if self._stack else self._stack.append([message])