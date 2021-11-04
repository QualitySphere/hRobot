#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# from robot.libraries import BuiltIn as robotBuiltIn
# from robot.libraries import Collections as robotCollections
# from robot.libraries import DateTime as robotDateTime
from robot.api import logger
import re
import os
import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning
# from selenium import webdriver
import json
# import platform
import paramiko
from datetime import datetime


def print_info(text):
    """
    RobotFramework 的日志输出，默认开启 HTML 代码显示，开启控制台显示
    :param text:
    :return:
    """
    logger.info(text, html=True, also_console=True)


def smart_content(content):
    """
    转换内容中 {%%} 的部分，替换成函数执行的结果
    :param content:
    :return:
    """
    new_content = content
    # 开始解析替换 函数执行结果
    fun_pattern = re.compile("{%[a-zA-Z0-9 ()_.*/+-]*%}")
    for fun_name in fun_pattern.findall(content):
        fun_result = eval(fun_name.strip('{%').strip('%}').strip())
        print_info(u'%s = %s' % (fun_name, fun_result))
        new_content = new_content.replace(fun_name, str(fun_result))
    # 完成解析替换 函数执行结果
    if new_content != content:
        print_info(u'源内容: %s' % content)
        print_info(u'渲染后: %s' % new_content)
    return new_content


class BuiltIn(object):
    """关键字"""
    def __init__(self):
        """内置"""
        pass

    def log(self, message, level='INFO', html=False, console=False):
        """打印日志"""
        pass

    def sleep(self, time_):
        """休眠"""
        pass


# class DateTime(object):
#     """关键字"""
#     def __init__(self):
#         """时间"""
#         pass
#
#     def get_current_date(self, time_zone='local', increment=0,
#                          result_format='timestamp', exclude_millis=False):
#         """当前时间戳"""
#         pass


class HttpRequests(object):
    """关键字"""
    def __init__(self):
        """接口"""
        requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
        self.__session = None
        self.__response = None
        self.__variables = dict()

    def request_open(self):
        """开启会话"""
        if self.__session:
            return True
        print_info(u'开启会话')
        self.__session = requests.session()
        return True

    def request_close(self):
        """关闭会话"""
        if self.__session:
            print_info(u'关闭会话')
            self.__session.close()
        return True

    def request_get(self, url, headers=None, params=None, cookies=None):
        """GET"""
        self.request_open()
        self.__response = None
        if headers is None:
            headers = {
                "Content-Type": "Application/json",
                "User-Agent": "Hybrid Robot",
            }
        print_info('\n'.join([
            u'请求',
            u'   Method     : GET',
            u'   URL        : %s' % url,
            u'   Query      : %s' % params,
            u'   Cookies    : %s' % cookies,
            u'   Headers    : %s' % headers
        ]))
        # if params or params != 'None':
        #     params = json.loads(params)
        # else:
        #     params = None
        # if cookies or cookies != 'None':
        #     cookies = json.loads(cookies)
        # else:
        #     cookies = None
        self.__response = self.__session.get(url=url, headers=headers, params=params, cookies=cookies, verify=False)
        print_info('\n'.join([
            u'响应',
            u'   Status Code: %s' % self.__response.status_code,
            u'   Headers    : %s' % self.__response.headers,
            u'   Body       : %s' % self.__response.content.decode()
        ]))
        return True

    def request_post(self, url, headers, body, cookies=None):
        """POST"""
        # headers = json.loads(self.__smart_content(headers))
        # body = json.loads(self.__smart_content(body))
        self.request_open()
        self.__response = None
        print_info('\n'.join([
            u'请求',
            u'   Method     : POST',
            u'   URL        : %s' % url,
            # u'   Cookies    : %s' % cookies,
            u'   Headers    : %s' % headers,
            u'   Body       : %s' % body
        ]))
        headers = json.loads(headers)
        body = json.loads(body)
        if cookies:
            cookies = json.loads(cookies)
        # else:
        #     cookies = None
        # print(type(cookies))
        self.__response = self.__session.post(url=url, headers=headers, json=body, cookies=cookies, verify=False)
        # self.__response = self.__session.post(url=url, headers=headers, json=body, verify=False)
        print_info('\n'.join([
            u'响应',
            u'   Status Code: %s' % self.__response.status_code,
            u'   Headers    : %s' % self.__response.headers,
            u'   Body       : %s' % self.__response.content.decode()
        ]))
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
            print_info(u'指定的检查路径 %s 不存在' % smart_key)
            raise KeyError
        for sk in smart_key.split('.')[1:]:
            try:
                smart_value = smart_value[sk]
            except TypeError:
                smart_value = smart_value[int(sk)]
            except KeyError:
                print_info(u'指定的检查路径 %s 下的 %s 不存在' % (smart_key, sk))
                raise KeyError
        return smart_value

    def response_get_value(self, smart_key, var_name):
        """响应.取值"""
        print_info(u'获取 %s 的值并赋予 %s' % (smart_key, var_name))
        return self.__get_response_smart_value(smart_key)
        # self.__variables[var_name] = self.__get_response_smart_value(smart_key)
        # print_info(u'%s=%s' % (var_name, self.__variables[var_name]))
        # return self.__variables[var_name]

    def response_assert(self, smart_key, assert_key, expected_value):
        """响应.断言"""
        print_info(u'检查 %s 是否符合预期值 %s' % (smart_key, expected_value))
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


class SshRemote(object):
    """关键字"""
    def __init__(self):
        """远程"""
        pass

    def ssh_push(self, host, user, password, local_path, remote_path):
        """上传"""
        transport = paramiko.Transport(sock=(host, 22))
        transport.connect(username=user, password=password)
        sftp = paramiko.SFTPClient.from_transport(transport)
        if os.path.isdir(local_path):
            print_info(u'上传目录 %s 到 %s:%s' % (local_path, host, remote_path))
            for _f in os.listdir(local_path):
                __local_path = os.path.join(local_path, _f)
                __remote_path = os.path.join(remote_path, _f)
                print_info(u'上传文件 %s 到 %s:%s' % (__local_path, host, __remote_path))
                if os.path.isfile(__local_path):
                    sftp.put(localpath=__local_path, remotepath=__remote_path)
        else:
            print_info(u'上传文件 %s 到 %s:%s' % (local_path, host, remote_path))
            sftp.put(localpath=local_path, remotepath=remote_path)
        # try:
        #     if os.path.isdir(local_path):
        #         print_info(u'上传目录 %s 到 %s:%s' % (local_path, host, remote_path))
        #         for _f in os.listdir():
        #             if os.path.isfile(os.path.join(local_path, _f)):
        #                 sftp.put(localpath=os.path.join(local_path, _f), remotepath=os.path.join(remote_path, _f))
        #     else:
        #         print_info(u'上传文件 %s 到 %s:%s' % (local_path, host, remote_path))
        #         sftp.put(localpath=local_path, remotepath=remote_path)
        # finally:
        sftp.close()
        transport.close()
        return True

    def ssh_pull(self, host, user, password, remote_path, local_path):
        """下载"""
        transport = paramiko.Transport(sock=(host, 22))
        transport.connect(username=user, password=password)
        sftp = paramiko.SFTPClient.from_transport(transport)
        try:
            print_info(u'下载文件 %s:%s 到 %s' % (host, remote_path, local_path))
            sftp.get(remotepath=remote_path, localpath=local_path)
        finally:
            sftp.close()
            transport.close()
        return True

    def ssh_exec(self, host, user, password, cmd):
        """执行"""
        ssh = paramiko.SSHClient()
        ssh_output = {
            "stdin": '',
            "stdout": '',
            "stderr": '',
        }
        __ssh_out = ssh_output['stdout']
        try:
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            print_info('%s@%s %s' % (user, host, cmd))
            ssh.connect(
                hostname=host,
                port=22,
                username=user,
                password=password
            )
            stdin, stdout, stderr = ssh.exec_command(cmd)
            __ssh_error = stderr.read().decode('utf-8')
            # assert not __ssh_error, u'远程执行命令有误:\n%s' % __ssh_error
            assert stdout.channel.recv_exit_status() == 0, u'远程执行命令有误:\n%s' % __ssh_error
            __ssh_out = stdout.read().decode('utf-8')
        finally:
            ssh.close()
        print_info(__ssh_out)
        return True


class HRobot(object):
    """关键字"""
    def __init__(self):
        """HRobot"""
        self.__webdriver = None
        self.__session = None
        self.__response = None
        self.__ssh_output = None
        self.__variables = {
            'test_variables': dict(),
            'suite_variables': dict(),
            'global_variables': dict(),
        }

    def hrobot_get_current_timestamp(self):
        """当前时间戳"""
        return int(datetime.now().timestamp()) * 1000

    def hrobot_set_test_variable(self, name, value):
        """设置变量"""
        pass

    def hrobot_set_suite_variable(self, name, value):
        """设置用例集变量"""
        pass

    def hrobot_set_global_variable(self, name, value):
        """设置全局变量"""
        pass

    def __smart_content(self, content):
        new_content = content
        # 开始解析替换 变量
        var_pattern = re.compile("{{[a-zA-Z0-9 _-]*}}")
        for var_key_str in var_pattern.findall(content):
            var_key = var_key_str.strip('{{').strip('}}').strip()
            try:
                # 尝试在当前类中的变量列表中查找
                var_value = self.__variables[var_key]
            except KeyError:
                # 若不存在再到全局变量中查找
                var_value = os.getenv('HROBOT_%s' % var_key)
            print_info(u'%s = %s' % (var_key_str, var_value))
            new_content = new_content.replace(var_key_str, str(var_value))
        # 完成解析替换 变量
        # 开始解析替换 函数执行结果
        fun_pattern = re.compile("{%[a-zA-Z0-9 ()_.*/+-]*%}")
        for fun_name in fun_pattern.findall(content):
            fun_result = eval(fun_name.strip('{%').strip('%}').strip())
            print_info(u'%s = %s' % (fun_name, fun_result))
            new_content = new_content.replace(fun_name, str(fun_result))
        # 完成解析替换 函数执行结果
        if new_content != content:
            print_info(u'源内容: %s' % content)
            print_info(u'渲染后: %s' % new_content)
        return new_content

    # def kw_timestamp(self):
    #     """时间戳"""
    #     return True

    # def kw_assert(self, key1, assert_key, key2):
    #     """断言"""
    #     if assert_key == "==":
    #         assert key1 == key2
    #     elif assert_key == ">=":
    #         assert key1 >= key2
    #
    # def kw_set_env(self, key, value):
    #     """全局变量.赋值"""
    #     var_value = self.__smart_content(value)
    #     print_info('%s=%s' % (key, var_value))
    #     os.environ.setdefault('HROBOT_%s' % key, var_value)
    #     return True
    #
    # def kw_def_var(self, key, value):
    #     """变量.赋值"""
    #     var_value = self.__smart_content(value)
    #     print_info('%s=%s' % (key, var_value))
    #     self.__variables[key] = var_value
    #     return True

    # def kw_sleep(self, seconds):
    #     """休眠"""
    #     print_info(u'小睡 %s 秒' % seconds)
    #     if type(seconds) is str:
    #         time.sleep(int(float(seconds)))
    #     else:
    #         time.sleep(int(seconds))
    #     return True


#
#     def kw_webdriver_open(self):
#         """内置.浏览器.启动"""
#         if self.__webdriver:
#             return True
#         opts = webdriver.ChromeOptions()
#         opts.add_argument('-lang=zh-cn')
#         opts.add_argument('--ignore-certificate-errors')
#         if platform.system() == 'Linux':
#             width = 1920
#             height = 1200
#             opts.add_argument('--head-less')
#             opts.add_argument('--no-sandbox')
#             self.__webdriver = webdriver.Chrome(options=opts)
#             self.__webdriver.set_window_size(width, height)
#         else:
#             self.__webdriver = webdriver.Chrome(options=opts)
#             self.__webdriver.maximize_window()
#         return True
#
#     def kw_webdriver_close(self):
#         """内置.浏览器.关闭"""
#         if self.__webdriver:
#             self.__webdriver.close()
#         return True
#
#     def kw_webdriver_access(self, url):
#         """内置.浏览器.访问"""
#         self.__webdriver.get(url)
#         return True
#
#     def kw_webdriver_find(self, xpath):
#         """内置.浏览器.查找"""
#         pass
#
#     def kw_webdriver_click(self, xpath):
#         """内置.浏览器.点击"""
#         element = self.__webdriver.find_element_by_xpath(xpath)
#         element.click()
#         return True
#
#     def kw_webdriver_input(self, xpath, text):
#         """内置.浏览器.输入"""
#         element = self.__webdriver.find_element_by_xpath(xpath)
#         element.send_keys(text)
#         return True
#
#     def kw_request_open(self):
#         """内置.接口.开启会话"""
#         if self.__session:
#             return True
#         self.__session = requests.session()
#         return True
#
#     def kw_request_close(self):
#         """内置.接口.关闭会话"""
#         if self.__session:
#             self.__session.close()
#         return True
#
#     def kw_request_get(self, url, headers, params):
#         """内置.接口.GET"""
#         url = self.smart_content(url)
#         headers = json.loads(self.smart_content(headers))
#         params = json.loads(self.smart_content(params))
#         self.kw_request_open()
#         if headers is None:
#             headers = {
#                 "Content-Type": "Application/json"
#             }
#         self.__response = self.__session.get(url=url, headers=headers, params=params, verify=False)
#         print_info('\n'.join([
#             u'请求',
#             u'   Method     : %s' % self.__response.request.method,
#             u'   URL        : %s' % url,
#             u'   Query      : %s' % params,
#             u'   Cookies    : %s' % self.__response.request._cookies._cookies,
#             u'   Headers    : %s' % self.__response.request.headers
#         ]))
#         print_info('\n'.join([
#             u'响应',
#             u'   Status Code: %s' % self.__response.status_code,
#             u'   Headers    : %s' % self.__response.headers,
#             u'   Body       : %s' % self.__response.content.decode()
#         ]))
#         return True
#
#     def kw_request_post(self, url, headers, body):
#         """内置.接口.POST"""
#         url = self.smart_content(url)
#         headers = json.loads(self.smart_content(headers))
#         body = json.loads(self.smart_content(body))
#         self.kw_request_open()
#         self.__response = self.__session.post(url=url, headers=headers, json=body, verify=False)
#         print_info('\n'.join([
#             u'请求',
#             u'   Method     : %s' % self.__response.request.method,
#             u'   URL        : %s' % url,
#             u'   Cookies    : %s' % self.__response.request._cookies._cookies,
#             u'   Headers    : %s' % self.__response.request.headers,
#             u'   Body       : %s' % body
#         ]))
#         print_info('\n'.join([
#             u'响应',
#             u'   Status Code: %s' % self.__response.status_code,
#             u'   Headers    : %s' % self.__response.headers,
#             u'   Body       : %s' % self.__response.content.decode()
#         ]))
#         return True

#     def __get_response_smart_value(self, smart_key):
#         if smart_key.startswith('status_code'):
#             smart_value = self.__response.status_code
#         elif smart_key.startswith('body'):
#             smart_value = self.__response.json()
#         elif smart_key.startswith("headers"):
#             smart_value = self.__response.headers
#         elif smart_key.startswith("cookies"):
#             smart_value = self.__response.cookies
#         else:
#             print_info(u'指定的检查路径 %s 不存在' % smart_key)
#             raise KeyError
#         for sk in smart_key.split('.')[1:]:
#             try:
#                 smart_value = smart_value[sk]
#             except TypeError:
#                 smart_value = smart_value[int(sk)]
#             except KeyError:
#                 print_info(u'指定的检查路径 %s 下的 %s 不存在' % (smart_key, sk))
#                 raise KeyError
#         return smart_value
# #
#     def kw_response_get_value(self, smart_key, var_name):
#         """内置.接口.响应.取值"""
#         print_info(u'获取 %s 的值并赋予 %s' % (smart_key, var_name))
#         self.__variables[var_name] = self.get_response_smart_value(smart_key)
#         print_info(u'%s=%s' % (var_name, self.__variables[var_name]))
#         return True
#
#     def kw_response_assert(self, smart_key, assert_key, expected_value):
#         """内置.接口.响应.断言"""
#         print_info(u'检查 %s 是否符合预期值 %s' % (smart_key, expected_value))
#         smart_value = self.get_response_smart_value(smart_key)
#         if type(smart_value) is int:
#             smart_expected_value = int(float(expected_value))
#         elif type(smart_value) is str:
#             smart_expected_value = str(expected_value)
#         else:
#             smart_expected_value = expected_value
#         if assert_key.lower() in ['=', '==', u'等于']:
#             assert smart_value == smart_expected_value, u"预期:%s \n实际:%s" % (expected_value, smart_value)
#         elif assert_key.lower() in ['in', u'被包含']:
#             assert smart_value in smart_expected_value, u"预期:%s \n实际:%s" % (expected_value, smart_value)
#         elif assert_key.lower() in ['not in', u'不包含']:
#             assert smart_value not in smart_expected_value, u"预期:%s \n实际:%s" % (expected_value, smart_value)
#         elif assert_key.lower() in ['contains', 'contain', u'包含']:
#             assert smart_expected_value in smart_value, u"预期:%s \n实际:%s" % (expected_value, smart_value)
#         elif assert_key.lower() in ['contains', 'contain', u'不包含']:
#             assert smart_expected_value not in smart_value, u"预期:%s \n实际:%s" % (expected_value, smart_value)
#         else:
#             raise KeyError(u'不支持的断言关系 %s' % assert_key)
#
#     def kw_ssh_exec(self, host, user, password, cmd):
#         """内置.远程.执行"""
#         ssh = paramiko.SSHClient()
#         ssh_output = {
#             "stdin": '',
#             "stdout": '',
#             "stderr": '',
#         }
#         self.__ssh_output = ssh_output['stdout']
#         try:
#             ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
#             print_info('%s@%s %s' % (user, host, cmd))
#             ssh.connect(
#                 hostname=host,
#                 port=22,
#                 username=user,
#                 password=password
#             )
#             stdin, stdout, stderr = ssh.exec_command(cmd)
#             __ssh_error = stderr.read().decode('utf-8')
#             assert not __ssh_error, u'远程执行命令有误:\n%s' % __ssh_error
#             self.__ssh_output = stdout.read().decode('utf-8')
#         finally:
#             ssh.close()
#         print_info(self.__ssh_output)
#         return True


if __name__ == '__main__':
    print(u'这是 Hybrid Robot 中文关键字转换')

