# Hybrid Robot

基于 RobotFramework 的二次封装

> **注意**
> 
> 该方案是对 RobotFramework 的二次封装，以减轻中国用户使用符合欧美用户习惯的原生 RobotFramework 时的不适，因此，该方案并**不适合**所有场景。针对技术能力暂时薄弱、英文能力差的中小团队，可能具有一定的疗效。

[![org](https://img.shields.io/static/v1?style=for-the-badge&label=org&message=Assurance%20Sphere&color=ff582c)](https://will.bowxeon.com)
![License](https://img.shields.io/github/license/qualitysphere/hrobot?style=for-the-badge)
![author](https://img.shields.io/static/v1?style=for-the-badge&label=author&message=362639663@qq.com&color=blue)
![author](https://img.shields.io/static/v1?style=for-the-badge&label=author&message=v.stone@163.com&color=blue)
[![Python](https://img.shields.io/static/v1?style=for-the-badge&logo=python&label=Python&message=3.8&color=306ba1)](https://docs.python.org/release/3.8.11/)
[![PyPI](https://img.shields.io/pypi/v/hrobot.svg?style=for-the-badge)](https://pypi.org/project/hrobot/)

## 重点改造描述

- 用 Excel 取代 .robot 文件进行测试用例的编写
- 增加命令行脚手架工具，一键初始化测试项目，不需要额外学习如何自己搭建目录结构
- 精简固化用法，高级用法被封装隐藏
- 集成 Allure 测试报告框架
- 检测到系统为非 Windows/Mac 系统时，将增加环境变量 `HROBOT_BROWSER_MODE=headless`，Chrome 将不启动图形界面
- 安装时会集成常用于测试场景的 Python 库，如: requests, selenium, paramiko

## 快速开始

#### 安装

```bash
pip install hrobot
```

#### 执行 `hrobot` 检查

```text
hrobot
  init     初始化项目目录
  run      执行测试用例
  debug    调试测试用例，支持选择用例集、测试用例、标签
  report   生成并展示测试报告
  version  显示版本信息
```

#### 初始化测试用例项目

使用 `hrobot init` 初始化项目目录

```text
hrobot init 
  -p    project  定义一个测试项目目录的名称
```

进入到项目目录中后，文件树接口如下：

```text
<projectName>
├── testcases
│   └── suites.xlsx
├── keywords
│   └── keywords.xlsx
└── variables
    └── variables.xlsx
```

*PS: keywords 暂无实际作用*  

#### 编写测试用例 

通过 Excel 打开 testcases 目录中的 suite.xlsx 文件，有 5 个 Sheet，每个 Sheet 有自己的表头:

- 用例

  <img src="https://qualitysphere.gitee.io/img/hrobot-suite.png" width="800">

- 变量

  <img src="https://qualitysphere.gitee.io/img/hrobot-variable.png" width="480">
  
- 前置

  <img src="https://qualitysphere.gitee.io/img/hrobot-setup.png" width="480">

- 后置

  <img src="https://qualitysphere.gitee.io/img/hrobot-setup.png" width="480">
  
- 可用关键字

  <img src="https://qualitysphere.gitee.io/img/hrobot-setup.png" width="480">

例子：

  <br>|A|B|C|D|E|F|G|H
  ----|----|----|----|----|----|----|----|---
  1 |用例标题|标签|用例描述|关键字库|关键字|参数|
  2 |SSH远程后再调用 HTTP 接口 |<br>|<br>| 远程 | 执行 | root | password | whoami
  3 |<br> |<br>|<br>|接口| GET | https://xxx/api/info | {"Content-Type":"application/json"}
  4 |HTTP 接口请求后断言 |<br>|<br>| 接口 | POST | https://xxx/api/products | {"Content-Type":"application/json | {"project_owner":"jing"}
  5 |<br> |<br>|<br>| 接口 | 响应.断言 | status_code | 等于 | 200
  6 |<br> |<br>|<br>| 接口| 响应.断言 | body.data.0.name | 等于 | hrobot
  7 |HTTP 返回值获取 |<br>|<br>| 接口 | POST | https://xxx/api/login | {"Content-Type":"application/json"} | {"username":"jing"}
  8 |<br> |<br>|<br>| 接口| 响应.取值 | body.token | AUTH
  9 |<br> |<br>|<br>| 接口| GET | https://xxx/api/info | {"Content-Type":"application/json","Authorization":"${AUTH}"} 
  10|<br> |<br>|<br>| 接口| 响应.断言 | status_code | 等于 | 200
  11|<br> |<br>|<br>| 接口 | 响应.断言 | body.username | 等于 | jing

#### 执行用例

```bash
cd <projectName>
hrobot run
```

#### 生成报告

```bash
hrobot report
```

## 定义变量

变量有三个级别：用例、用例集、全局，均可以在测试用例的 Excel 编写中通过关键字定义。

除此之外，也可以在用例集的 Excel 的 "变量" Sheet 中提前定义好用例集级别的变量。

全局变量还可以在 variables 目录下的 Excel 中定义，该处定义的全局变量会在报告的环境信息段落中展现。

<img src="https://qualitysphere.gitee.io/img/hrobot-variable.png" width="480">

## 自定义关键字

- 待设计
