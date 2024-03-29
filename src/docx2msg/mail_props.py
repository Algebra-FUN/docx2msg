from datetime import datetime
from enum import IntEnum
from typing import Any, Callable, Dict, List, Type, Union

from win32com.client import CDispatch


def OnlyTypeSet(k: str, T: Type, parser: Callable[[Any], Any] = None):
    def __set_attr(mail: object, val: T, **kwargs):
        if isinstance(val, T):
            if parser:
                val = parser(val)
            setattr(mail, k, val)
        else:
            raise ValueError(f"This mail property must be a {T}, not {type(val)}")

    return __set_attr


def StrIntEnumSet(cls: IntEnum, set_attr: Callable[[object, str, Any], None] = setattr):
    def __set_attr(mail: object, val: Union[str, int], **kwargs):
        if isinstance(val, str):
            set_attr(mail, cls.__name__, cls[val].value)
        elif isinstance(val, int):
            set_attr(mail, cls.__name__, cls(val).value)
        else:
            raise ValueError(
                f"This mail property must be a string or an integer, not {type(val)}"
            )

    return __set_attr


class Importance(IntEnum):
    Low = 0
    Normal = 1
    High = 2


class Sensitivity(IntEnum):
    Normal = 0
    Personal = 1
    Private = 2
    Confidential = 3


def StrListSet(attr: str):
    def __set_attr(mail: object, item: Union[str, List[str]], **kwargs):
        if isinstance(item, List):
            item = ";".join(item)
        if isinstance(item, str):
            setattr(mail, attr, item)
        else:
            raise ValueError(
                f"This mail property must be a string or a List of string , not {type(item)}"
            )

    return __set_attr


def AttrsAdd(attr: str):
    def __set_attr(mail: object, items: Union[str, List[str]], **kwargs):
        if isinstance(items, str):
            items = items.split(";")
        if isinstance(items, List):
            for item in items:
                getattr(mail, attr).Add(item)
        else:
            raise ValueError(
                f"This mail property must be a string or a List of string , not {type(items)}"
            )

    return __set_attr


def parse_datetime(val: datetime):
    return f"{val:%Y-%m-%d %H:%M}"


def set_save_sent_folder(mail: object, path: str, outlook: CDispatch, **kwargs):
    folder = outlook.Session
    for folder_name in path.split("/"):
        try:
            folder_name = int(folder_name)
        except ValueError:
            pass
        folder = folder.Folders[folder_name]
    mail.SaveSentMessageFolder = folder


TYPED_PROPERTIES = {
    "Subject": str,
    "Categories": str,
    "OriginatorDeliveryReportRequested": bool,
    "ReadReceiptRequested": bool,
    "FlagRequest": str,
    "VotingOptions": str,
}
MAILADDRESS_PROPERTIES = {"To", "CC", "BCC"}
DATETIME_PROPERTIES = {"DeferredDeliveryTime", "ExpiryTime", "FlagDueBy"}
SPECIAL_PROPERTIES = {
    "Attachments": AttrsAdd("Attachments"),
    "ReplyRecipients": AttrsAdd("ReplyRecipients"),
    "Importance": StrIntEnumSet(Importance),
    "Sensitivity": StrIntEnumSet(Sensitivity),
    "ReminderTime": OnlyTypeSet("FlagDueBy", datetime, parser=parse_datetime),
    "SaveSentMessageFolder": set_save_sent_folder,
}

MAILADDRESS_PROPERTIES = {k: StrListSet(k) for k in MAILADDRESS_PROPERTIES}
TYPED_PROPERTIES = {k: OnlyTypeSet(k, T) for k, T in TYPED_PROPERTIES.items()}
DATETIME_PROPERTIES = {
    k: OnlyTypeSet(k, datetime, parser=parse_datetime) for k in DATETIME_PROPERTIES
}
SET_SUPPORTED_PROPERTIES: Dict[str, Callable[[object, Any], None]] = {
    **MAILADDRESS_PROPERTIES,
    **TYPED_PROPERTIES,
    **DATETIME_PROPERTIES,
    **SPECIAL_PROPERTIES,
}
