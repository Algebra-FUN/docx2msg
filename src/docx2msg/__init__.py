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
from docx2msg import Docx2Msg

docx_path = r"path/to/your/docx"
with Docx2Msg(docx_path) as docx:
    mail = docx.convert(display=True)
```

"""

__version__ = "0.1.2"

import base64
import warnings
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any, Dict, Literal, Optional, Union

import win32com.client
import yaml
from bs4 import BeautifulSoup
from docx import Document
from docxtpl import DocxTemplate

from .mail_props import *


class Docx2Msg:
    """Class for converting a docx to an Outlook Mail-Item."""

    def __init__(self, docx: Union[str, Path]) -> None:
        """Class for converting a docx to an Outlook Mail-Item.

        Parameters
        ----------
        docx : Union[str, Path]
            The path of the docx file.

        References
        ----------
        - pywin32: https://pypi.org/project/pywin32/
        - Outlook API: https://learn.microsoft.com/en-us/office/vba/api/overview/outlook
        - Word API: https://learn.microsoft.com/en-us/office/vba/api/overview/word
        - docxtpl: http://docxtpl.readthedocs.org/
        """
        self.original_docx_path = Path(docx)
        self.docx_path = Path(docx)
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.word = win32com.client.Dispatch("Word.Application")
        self.word.Visible = True
        self.word.DisplayAlerts = 0
        self.__headers: Dict[str, Any] = None
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

    def __extract_headers(self) -> Optional[Dict[str, Any]]:
        doc = Document(self.docx_path)
        # get header of the first section of the document
        header = doc.sections[0].header
        header_text = "\n".join(p.text for p in header.paragraphs)
        attrs = yaml.safe_load(header_text)
        return attrs

    @property
    def headers(self) -> Optional[Dict[str, Any]]:
        """Get mail properties from the header of the docx."""
        if self.__headers:
            return self.__headers
        self.__template_save()
        return self.__extract_headers()

    @headers.setter
    def headers(self, value: Union[str, Dict[str, Any]]) -> None:
        """Set mail properties from a Dict."""
        if isinstance(value, str):
            value = yaml.safe_load(value)
        if isinstance(value, Dict):
            self.__headers = value
        else:
            raise ValueError("The value must be a Dict or a string for YAML.")

    def load_headers(self, path: Union[str, Path]) -> None:
        """Load mail properties from a YAML file instead of the header of the docx.

        Parameters
        ----------
        path : Union[str, Path]
            The path of a YAML file.
        """
        try:
            context = Path(path).read_text()
        except Exception:
            raise ValueError(
                "The path must be a string for a Path or a Path object for a YAML file."
            )
        self.headers = context
        return context

    def __extract_html(self) -> str:
        docx = self.word.Documents.Open(str(self.docx_path))
        docx.SaveEncoding = 65001
        html_path = self.temp_dir / "temp.html"
        docx.SaveAs2(str(html_path), FileFormat=10, Encoding=65001)
        docx.Close()
        html = html_path.read_text(encoding="utf-8")
        return self.__revise_html(html)

    @property
    def html(self) -> str:
        """Get the desired HTMLBody of the docx."""
        self.__template_save()
        return self.__extract_html()

    def __revise_html(self, html: str) -> str:
        """Revise the HTMLBody of the docx."""
        soup = BeautifulSoup(html, "html.parser")
        self.__base64_img(soup)
        return str(soup)

    def __base64_img(self, soup: BeautifulSoup) -> None:
        """Convert the img src to base64."""
        for img in soup.find_all("img"):
            img_path: Path = self.temp_dir / img["src"]
            ext = img_path.suffix.lstrip(".")
            if img_path.exists():
                with open(img_path, "rb") as f:
                    img_data = base64.b64encode(f.read()).decode()
                img["src"] = f"data:image/{ext};base64,{img_data}"

    def convert(
        self,
        reply_on: Optional["_MailItem"] = None,
        reply_mode: Literal["Reply", "ReplyAll"] = "Reply",
        display=False,
        force_render=False,
    ) -> "_MailItem":
        """
        Convert the docx(or docx template) to an Outlook Mail-Item.

        Parameters
        ----------
        reply_on : Optional[_MailItem] (default None)
            The Mail-Item to reply on. If None, create a new Mail-Item.
        reply_mode : Literal['Reply', 'ReplyAll'] (default 'Reply')
            The mode to reply on the Mail-Item. Only available when `reply_on` is not None.
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
        - Render Issue: Sometimes, the HTMLBody of generated .msg mail file may not be well rendered, you can set the `force_render` parameter to True to ensure the HTMLBody is well rendered.
        However, this will causes the Mail-Item to be saved in the default draft folder in Outlook, which means you need to cleanup the Mail-Items manually if you don't want to keep it.
        - Recipients Priority: The recipients in the docx header will have a higher priority than the recipients in the reply mail of `reply_on` Mail-Item. So if you want to reply on a mail to sender, please don't set the `To` in the docx header.
        """
        self.__template_save()

        # create or reply a mail
        if reply_on is None:
            self.mail = self.outlook.CreateItem(0)
        else:
            try:
                self.mail = (
                    reply_on.ReplyAll()
                    if reply_mode == "ReplyAll"
                    else reply_on.Reply()
                )
            except Exception:
                raise ValueError("The reply_on must be a Mail-Item.")

        # set mail attributes
        headers = self.__extract_headers()
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
        if reply_on is None:
            self.mail.HTMLBody = self.__extract_html()
        else:
            # when reply a mail, add new body before the original body
            self.mail.HTMLBody = self.__extract_html() + "\n" + self.mail.HTMLBody

        if display:
            self.mail.Display()
        if force_render:
            self.mail.Save()
            warnings.warn(
                "The Mail-Item has been saved in the default draft folder in Outlook in order to well render the HTMLBody."
            )
        return self.mail
