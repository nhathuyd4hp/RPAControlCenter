from playwright._impl._errors import Error, TimeoutError
from playwright.sync_api import Page


def apply_patch():
    if getattr(Page, "_patched_goto", False):
        return
    _original_goto = Page.goto

    def safe_goto(self, url, **kwargs):
        try:
            return _original_goto(self, url, **kwargs)
        except TimeoutError:
            return _original_goto(self, url, **kwargs)
        except Error as e:
            if "net::ERR_ABORTED" in e.message:
                return _original_goto(self, url, **kwargs)
            raise e

    Page.goto = safe_goto
    Page._patched_goto = True
