# docx2msg

Converts a docx to an Outlook Mail-Item with few lines of code.

## Description

`docx2msg` is a python package that allows you to convert Microsoft Word .docx files to Outlook Mail-Item and .msg files. It provides a simple and efficient way to automate the conversion process, making it easier to work with outlook email in your workflows.

## Features

- Convert a docx to an Outlook Mail-Item just using Word and Outlook Application without any third-party library.
- Set mail properties from the header of the docx in YAML format.
- Able to use docx-template to render the docx body and set mail properties dynamically.

## Installation

```shell
pip install docx2msg
```

## Usage

### Simple Usage

1. Edit the body of the docx file to set the mail body as a normal docx file saving at "path\to\your\docx".

    > You are recommand to edit the docx file with Microsoft Word Application in web layout mode to avoid the unexpected format issue.

2. Edit the header of the docx file to set the mail properties. The header should be in YAML format in this way:

    ```yaml
    Subject: Demo email
    To: anyone@example.com
    CC: p1@example.com;p2@example.com
    Attachments: path\to\your\attachment1.docx;path\to\your\attachment2.msg
    Importance: High
    Sensitivity: Confidential
    ReadReceiptRequested: True
    Categories: RED CATEGORY, BLUE CATEGORY
    FlagRequest: Test Flag
    ReminderTime: 2024-02-29 14:00:00
    ```

3. Convert a docx to an Outlook Mail-Item with few lines of code:

    ```python
    import win32com.client
    from docx2msg import Docx2Msg

    outlook = win32com.client.Dispatch("Outlook.Application")
    word = win32com.client.Dispatch("Word.Application")
    docx_path = r"path\to\your\docx"
    with Docx2Msg(docx_path, outlook=outlook, word=word) as docx:
        # set display=True to display the mail in Outlook Application
        mail = docx.convert(display=True)
    ```

4. The mail will be displayed in Outlook Application and you can see the output.

### Advanced Usage with template

Since `docx2msg` uses `docx-template` to render the docx body, you can use the same syntax to render the docx body and set mail properties dynamically.

You can access the `template` attribute of the `Docx2Msg` object to utilize the `docx-template` features.

> Go to [python-docx-template’s documentation](https://docxtpl.readthedocs.io/en/latest/) for more details.

Run the following code to convert a docx to an Outlook Mail-Item with a template:

```python
# the context to render the docx
context = {
    "name": "John Doe",
    "age": 30,
    "address": "123 Main St."
}
with Docx2Msg(docx_path, outlook=outlook, word=word) as docx:
    # use template attribute to render the docx body
    docx.template.render(context)
    # convert the docx to an Outlook Mail-Item
    mail = docx.convert()
    # display the mail in Outlook Application
    mail.Display()
    # save mail in draft folder
    mail.Save()
    # save mail as .msg file
    mail.SaveAs(r"path\to\your\output.msg")
```

The output from `convert` method will be an `MailItem` object, for further development, you can refer to the [Outlook API](https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem) for more details.

## References
- pywin32: https://pypi.org/project/pywin32/
- Outlook API: https://learn.microsoft.com/en-us/office/vba/api/overview/outlook
- Word API: https://learn.microsoft.com/en-us/office/vba/api/overview/word
- docxtpl: http://docxtpl.readthedocs.org/
- jinja2: https://jinja.palletsprojects.com/en/3.0.x/