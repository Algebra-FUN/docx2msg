"""
docx2msg
========
Convert a docx to an Outlook Mail-Item.

Features
--------
- Convert a docx to an Outlook Mail-Item just using Word and Outlook Application without any third-party library.
- Set mail properties from the header of the docx in YAML format.
- Able to use docx-template to render the docx body and set mail properties dynamically.

Usage
-----

Easily convert a docx to an Outlook Mail-Item with few lines of code:

```python
import win32com.client
from docx2msg import Docx2Msg

outlook = win32com.client.Dispatch("Outlook.Application")
word = win32com.client.Dispatch("Word.Application")
docx_path = r"path\\to\your\docx"
with Docx2Msg(docx_path, outlook=outlook, word=word) as docx:
    mail = docx.convert(display=True)
```

"""

__version__ = "0.1.0"

import warnings
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any, Dict, Optional, Union

import yaml
from docx import Document
from docxtpl import DocxTemplate
from win32com.client import CDispatch

from .mail_props import *


class Docx2Msg:
    """Class for converting a docx to an Outlook Mail-Item."""

    def __init__(
        self, docx: Union[str, Path], outlook: CDispatch = None, word: CDispatch = None
    ) -> None:
        """Class for converting a docx to an Outlook Mail-Item.

        Parameters
        ----------
        docx : Union[str, Path]
            The path of the docx file.
        outlook : CDispatch
            The instance of Outlook Application.
        word : CDispatch
            The instance of Word Application.

        References
        ----------
        - pywin32: https://pypi.org/project/pywin32/
        - Outlook API: https://learn.microsoft.com/en-us/office/vba/api/overview/outlook
        - Word API: https://learn.microsoft.com/en-us/office/vba/api/overview/word
        - docxtpl: http://docxtpl.readthedocs.org/
        """
        self.original_docx_path = Path(docx)
        self.docx_path = Path(docx)
        if outlook is None or word is None:
            raise ValueError("The outlook and word parameters are required.")
        self.outlook = outlook
        self.word = word
        self.word.Visible = True
        self.word.DisplayAlerts = 0
        self.__docx_template: DocxTemplate = None
        self.__temp_dir: TemporaryDirectory = None

    def __enter__(self):
        self.temp_dir
        return self

    @property
    def temp_dir(self) -> Path:
        if not self.__temp_dir:
            self.__temp_dir = TemporaryDirectory()
        return Path(self.__temp_dir.name)

    def __exit__(self, exc_type, exc_value, traceback):
        self.__temp_dir.cleanup()
        self.__temp_dir = None

    @property
    def template(self) -> DocxTemplate:
        """Return the instance of DocxTemplate."""
        if not self.__docx_template:
            self.__docx_template = DocxTemplate(self.docx_path)
        return self.__docx_template

    def __template_save(self):
        if self.__docx_template:
            self.docx_path = self.temp_dir / "temp.docx"
            self.template.save(self.docx_path)

    def __get_headers(self) -> Optional[Dict[str, Any]]:
        doc = Document(self.docx_path)
        # get header of the first section of the document
        header = doc.sections[0].header
        header_text = "\n".join(p.text for p in header.paragraphs)
        attrs = yaml.safe_load(header_text)
        return attrs

    @property
    def headers(self) -> Optional[Dict[str, Any]]:
        """Get mail properties from the header of the docx.

        Notes
        -----
        The header of the docx should be in YAML format.
        """
        self.__template_save()
        return self.__get_headers()

    def __get_html(self) -> str:
        docx = self.word.Documents.Open(str(self.docx_path))
        docx.SaveEncoding = 65001
        html_path = self.temp_dir / "temp.html"
        docx.SaveAs2(str(html_path), FileFormat=10, Encoding=65001)
        docx.Close()
        return html_path.read_text(encoding="utf-8")

    @property
    def html(self) -> str:
        """Get the HTMLBody of the docx."""
        self.__template_save()
        return self.__get_html()

    def convert(self, display=False, force_render=False) -> "_MailItem":
        """
        Convert the docx(or docx template) to an Outlook Mail-Item.

        Parameters
        ----------
        display : bool (default False)
            Whether to display the created Mail-Item.
        force_render : bool (default False)
            Whether to force render the HTMLBody of the Mail-Item by saving the Mail-Item in the default draft folder in Outlook.

        Returns
        -------
        mail : _MailItem
            The converted Outlook Mail-Item.

        Raises
        ------
        AttributeError
            If the mail property is not supported by the Mail-Item in Outlook.

        Notes
        -----
        Sometimes, the HTMLBody of generated .msg mail file may not be well rendered, you can set the `force_render` parameter to True to ensure the HTMLBody is well rendered.
        However, this will causes the Mail-Item to be saved in the default draft folder in Outlook, which means you need to cleanup the Mail-Items manually if you don't want to keep it.
        """
        self.__template_save()

        self.mail = self.outlook.CreateItem(0)
        # set mail attributes
        headers = self.__get_headers()
        for k, v in headers.items():
            if k in SET_SUPPORTED_PROPERTIES:
                SET_SUPPORTED_PROPERTIES[k](self.mail, v, outlook=self.outlook)
            else:
                warnings.warn(
                    f"""The mail property "{k}" is not guaranteed to be set correctly, which may cause unexpected behavior or error.
                        Go https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem#properties to check the supported properties. 
                        Please pay attention to the Mail-Item before sending!"""
                )
                if not hasattr(self.mail, k):
                    raise AttributeError(
                        f"The mail property {k} is not supported by the Mail-Item in Outlook. Go https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem#properties to check the supported properties."
                    )
                setattr(self.mail, k, v)
        # set mail body
        self.mail.HTMLBody = self.__get_html()
        if display:
            self.mail.Display()
        if force_render:
            self.mail.Save()
            warnings.warn(
                "The Mail-Item has been saved in the default draft folder in Outlook in order to well render the HTMLBody."
            )
        return self.mail
