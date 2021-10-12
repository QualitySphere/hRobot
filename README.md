# Hybrid Robot
基于 RobotFramework 的二次封装

> **警告**
> 
> 该方案是对 RobotFramework 的二次封装，以减轻中国用户使用符合欧美用户习惯的原生 RobotFramework 时的不适，因此，该方案并**不适合**所有场景。针对技术能力暂时薄弱、英文能力差的中小团队，可能具有一定的疗效。

## 重点改造描述

- 用 Excel 取代 .robot 文件进行测试用例的编写
- 增加命令行脚手架工具，一键初始化测试项目，不需要额外学习如何自己搭建目录结构
- 精简固化用法，高级用法被封装隐藏
- 集成 Allure 测试报告框架
- 检测到系统为非 Windows/Mac 系统时，将增加环境变量 `HROBOT_BROWSER_MODE=headless`，Chrome 将不启动图形界面
- 安装时会集成常用于测试场景的 Python 库，如: requests, selenium, paramiko

## 使用手册

#### 安装

```bash
pip install hrobot
```

#### 初始化测试用例项目

使用 `hrobot` 工具初始化项目目录

```bash
hrobot init -p <projectName>
```

进入到项目目录中后，文件树接口如下：

```text
<projectName>
├── testcases
│   └── suites.xls
├── keywords
│   └── keywords.xls
└── variables
    └── variables.xls
```

#### 编写测试用例 

通过 Excel 打开 testcases 目录中的 suites.xls 文件，按照定义好的列进行填写

例子：

用例标题|关键字类型|关键字|参数| | | |
----|----|----|----|----|----|----
SSH远程后再调用 HTTP 接口 | 内置 | 远程.执行 | root | password | whoami
<ba> | 内置 | 接口.GET | https://xxx/api/info | {"Content-Type":"application/json"}
HTTP 接口请求后断言 | 内置 | 接口.POST | https://xxx/api/products | {"Content-Type":"application/json | {"project_owner":"jing"}
<ba> | 内置 | 接口.响应.断言 | status_code | 等于 | 200
<ba> | 内置 | 接口.响应.断言 | body.data.0.name | 等于 | hrobot
HTTP 返回值获取 | 内置 | 接口.POST | https://xxx/api/login | {"Content-Type":"application/json"} | {"username":"jing"}
<ba> | 内置 | 接口.响应.取值 | body.token | AUTH
<ba> | 内置 | 接口.GET | https://xxx/api/info | {"Content-Type":"application/json","Authorization":"{{AUTH}}"} 
<ba> | 内置 | 接口.响应.断言 | status_code | 等于 | 200
<ba> | 内置 | 接口.响应.断言 | body.username | 等于 | jing

#### 自定义变量

除了在测试用例执行过程中获取返回值作为变量传递，也可以通过 Excel 打开 variables 目录中的 variables.xls 文件，按照定义好的列进行填写，预先定义一些变量

例子:

变量名|变量类型|变量值
----|----|----
NAME|string|于静
MONTH|int|10

### 自定义关键字

- 待设计
