# docx2msg
使用几行代码将docx转换为Outlook邮件。

[![PyPI](https://img.shields.io/pypi/v/docx2msg)](https://pypi.org/project/docx2msg/)
[![PyPI - License](https://img.shields.io/pypi/l/docx2msg)](https://pypi.org/project/docx2msg/)

## 描述

`docx2msg`是一个Python包，允许您将Word .docx文件转换为Outlook邮件和.msg文件。它提供了一种简单高效的方式来自动化转换过程，使得在工作流中自动化Outlook电子邮件更加容易。

## 特点

- 只使用Word和Outlook应用程序，将docx转换为Outlook邮件，无需任何第三方库。
- 从docx的页眉以YAML格式设置邮件属性。
- 能够使用docx-template渲染docx正文并动态设置邮件属性。

## 要求

- 操作系统：Windows
- 应用程序：Microsoft Word，Microsoft Outlook
- Python：3.8+
- Python包：pywin32，python-docx，docx-template，pyyaml

## 安装

```shell
pip install docx2msg
```

## 用户指南

### 快速入门

> 示例的docx文件即将推出...

1. 编辑docx文件的正文，将邮件正文设置为普通的docx文件，并保存在"path\to\your\docx"路径下。

    > 建议使用Microsoft Word应用程序的Web布局模式编辑docx文件，以避免意外的格式问题。

2. 编辑docx文件的头部，设置邮件属性。头部应以YAML格式编写，如下所示：

    ```yaml
    Subject: 示例邮件
    To: anyone@example.com
    CC: p1@example.com;p2@example.com
    Attachments: path\to\your\附件1.docx;path\to\your\附件2.msg
    Importance: High
    Sensitivity: Confidential
    ReadReceiptRequested: True
    Categories: 红色类别, 蓝色类别
    FlagRequest: 例会
    ReminderTime: 2024-02-29 14:00:00
    ```

3. 使用几行代码将docx转换为Outlook邮件：

    ```python
    import win32com.client
    from docx2msg import Docx2Msg

    outlook = win32com.client.Dispatch("Outlook.Application")
    word = win32com.client.Dispatch("Word.Application")
    docx_path = r"path\to\your\docx"
    with Docx2Msg(docx_path, outlook=outlook, word=word) as docx:
        # 设置 display=True 以在Outlook应用程序中显示邮件
        mail = docx.convert(display=True)
    ```

4. 邮件将在Outlook应用程序中显示，您可以查看输出。

### 使用模板的高级用法

由于`docx2msg`使用`docx-template`来渲染docx正文，您可以使用相同的语法来渲染docx正文并动态设置邮件属性。

您可以访问`Docx2Msg`对象的`template`属性来利用`docx-template`的功能。

> 请参阅[python-docx-template的文档](https://docxtpl.readthedocs.io/en/latest/)了解更多详细信息。

运行以下代码将docx转换为带有模板的Outlook邮件：

```python
# the context to render the docx
context = {
    "姓名": "张三",
    "年龄": 30
}
with Docx2Msg(docx_path, outlook=outlook, word=word) as docx:
    # 使用template属性渲染docx正文
    docx.template.render(context)
    # 将docx转换为Outlook邮件
    mail = docx.convert()
    # 在Outlook应用程序中展示邮件
    mail.Display()
    # 将邮件保存在草稿文件夹中
    mail.Save()
    # 将邮件另保存为.msg文件
    mail.SaveAs(r"path\to\your\output.msg")
```

`convert` 方法的输出将是一个 `MailItem` 对象，您可以参考 [Outlook API](https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem) 了解更多详细信息。

### 邮件头部语法

邮件头部以 YAML 格式编写在 docx 文件的头部。支持以下属性：


| 属性                        | 描述                                   | 类型     | 示例                                          |
|-----------------------------|----------------------------------------|----------|-----------------------------------------------|
| To                          | 收件人                                 | str      | anyone@example.com                            |
| CC                          | 抄送人                                 | str      | p1@example.com;p2@example.com                 |
| BCC                          | 密送人                                 | str      | p1@example.com;p2@example.com                 |
| Subject                     | 主题                                   | str      | Demo email                                    |
| Attachments                 | 附件                                   | str(路径)      | path\to\your\file1.docx;path\to\your\file2.msg |
| Categories                  | 分类                                   | str      | 红色类别                   |
| Importance                  | 重要程度                               | str\|int      | High                                         |
| Sensitivity                 | 机密程度                               | str\|int      | Confidential                                  |
| ReadReceiptRequested        | 请求回执                               | bool     | True                                          |
| OriginatorDeliveryReportRequested| 请求发送人传递报告                 | bool     | True                                          |
| FlagRequest                 | 标记请求                               | str      | 例会                                     |
| VotingOptions               | 投票选项                               | str      | 是;否                                 |
| ReminderTime                | 提醒时间                               | datetime      | 2024-02-29 14:00:00                           |
| DeferredDeliveryTime         | 延迟发送时间                           | datetime      | 2024-02-29 14:00:00                           |
| ExpiryTime                   | 过期时间                               | datetime      | 2024-02-29 14:00:00                           |
| FlagDueBy                    | 标记到期时间                           | datetime      | 2024-02-29 14:00:00                           |
| ReplyRecipients              | 回复收件人                             | str      | p1@example.com;p2@example.com                      |
| SaveSentMessageFolder        | 保存已发送邮件的文件夹                 | str      | 1/自动生成/新                      |

**注意：**

对于一些属性，如`Attachments`、`To`、`CC`、`ReplyRecipients`，它们可能有多个值，您可以使用分号`;`来分隔它们。

对于`SaveSentMessageFolder`属性，示例中的"1/自动生成/新"是指在Python代码中通过`outlook.Session.Folders[1].Folders["自动生成"].Folder["新"]`访问的文件夹，这是`SaveSentMessageFolder`属性的语法糖。

所有属性都是`Outlook.MailItem`对象的有效属性，因此您可以参考https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem#properties了解更多详细信息。

### API文档

> API文档即将推出...

## 参考资料
- pywin32: https://pypi.org/project/pywin32/
- Outlook API: https://learn.microsoft.com/en-us/office/vba/api/overview/outlook
- Word API: https://learn.microsoft.com/en-us/office/vba/api/overview/word
- docxtpl: http://docxtpl.readthedocs.org/
- jinja2: https://jinja.palletsprojects.com/en/3.0.x/
