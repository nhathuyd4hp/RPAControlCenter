import logging
import os
import re
from contextlib import suppress
from typing import List
from urllib.parse import urlparse

from playwright._impl._errors import TimeoutError
from playwright.sync_api import Locator, sync_playwright


class SharePoint:
    def __init__(
        self,
        domain: str,
        email: str,
        password: str,
        context=None,
        browser=None,
        playwright=None,
        logger: logging.Logger = logging.getLogger("SharePoint"),
        headless: bool = False,
        timeout: float = 5000,
    ):
        self._external_context = context is not None
        self._external_browser = browser is not None
        self._external_pw = playwright is not None

        if playwright:
            self._pw = playwright
        else:
            self._pw = sync_playwright().start()

        if browser:
            self.browser = browser
        else:
            self.browser = self._pw.chromium.launch(
                headless=headless,
                timeout=timeout,
                args=["--start-maximized"],
            )

        if context:
            self.context = context
        else:
            self.context = self.browser.new_context(
                no_viewport=True,
            )

        self.page = self.context.new_page()
        self.domain = domain
        self.email = email
        self.password = password
        self.timeout = timeout
        self.logger = logger
        if not self.__authentication():
            raise PermissionError("Authentication failed")

    def __authentication(self) -> bool:
        try:
            self.page.bring_to_front()
            self.page.goto(self.domain, wait_until="domcontentloaded")
            if urlparse(self.page.url).netloc != urlparse(self.domain).netloc:
                try:
                    self.page.wait_for_selector("input[type='email']", state="visible", timeout=10000).fill(self.email)
                    self.page.locator(
                        selector="input[type='submit']",
                        has_text="Next",
                    ).click(timeout=self.timeout)
                    self.page.wait_for_selector("input[type='password']", state="visible", timeout=10000).fill(
                        self.password
                    )
                    self.page.locator(
                        selector="input[type='submit']",
                        has_text="Sign in",
                    ).click(timeout=self.timeout)
                    with self.page.expect_navigation(
                        url=re.compile(f"^{self.domain}"),
                        wait_until="load",
                        timeout=30000,
                    ):
                        self.page.locator("input[type='submit']", has_text="Yes").click()
                except TimeoutError:
                    error: Locator = self.page.wait_for_selector(
                        "div#usernameError, div#passwordError",
                    )
                    self.logger.error(error.text_content())
                    return False
            while True:
                try:
                    self.page.wait_for_selector("div[id='HeaderButtonRegion']", state="visible", timeout=10000)
                    with suppress(TimeoutError):
                        self.page.wait_for_selector(
                            selector="div[id='O365_MainLink_MePhoto']", timeout=1000, state="visible"
                        ).dblclick()
                        self.page.wait_for_selector(
                            selector="div[id='O365_MainLink_MePhoto']", timeout=1000, state="visible"
                        ).dblclick()
                        self.page.wait_for_selector(
                            selector="div[id='O365_MainLink_MePhoto']", timeout=1000, state="visible"
                        ).dblclick()
                        self.page.wait_for_selector(
                            selector="div[id='O365_MainLink_MePhoto']", timeout=1000, state="visible"
                        ).dblclick()
                        self.page.wait_for_selector(
                            selector="div[id='O365_MainLink_MePhoto']", timeout=1000, state="visible"
                        ).dblclick()
                    currentAccount_primary = (
                        self.page.locator("div[id='mectrl_currentAccount_primary']").text_content().lower()
                    )
                    currentAccount_secondary = (
                        self.page.locator("div[id='mectrl_currentAccount_secondary']").text_content().lower()
                    )
                    self.logger.info(f"Logged in as: {currentAccount_primary} ({currentAccount_secondary})")
                    break
                except TimeoutError:
                    self.page.reload(wait_until="domcontentloaded")
            return True
        except TimeoutError:
            return self.__authentication()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        if not self._external_context:
            self.context.close()
        if not self._external_browser:
            self.browser.close()
        if not self._external_pw:
            self._pw.stop()

    def download_files(
        self,
        url: str,
        file: re.Pattern | str,
        steps: List[str | re.Pattern] | None = None,
        save_to: str | None = None,
    ) -> List[str]:
        try:
            self.page.bring_to_front()
            downloads = []
            self.page.goto(url=url)
            self.page.wait_for_selector(
                selector="div[class='app-container']",
                timeout=10000,
                state="visible",
            )
            for step in steps or []:
                span: Locator = self.page.locator(
                    selector="span[role='button']",
                    has_text=step,
                )
                try:
                    span.first.wait_for(timeout=self.timeout, state="visible")
                except TimeoutError:
                    return []
                if span.count() != 1:
                    return []
                text = span.text_content()
                span.click()
                self.page.locator(
                    selector="div[type='button'][data-automationid='breadcrumb-crumb']:visible",
                    has_text=text,
                ).wait_for(
                    timeout=self.timeout,
                    state="visible",
                )
            items: Locator = self.page.locator(
                selector="span[role='button'][data-id='heroField']",
                has_text=file,
            )
            for i in range(items.count()):
                item: Locator = items.nth(i)
                item.click(button="right")
                download_btn = self.page.locator(
                    "button[data-automationid='downloadCommand'][role='menuitem']:not([type='button'])"
                )
                download_btn.wait_for(state="visible", timeout=self.timeout)
                with self.page.expect_download() as download_info:
                    download_btn.click()
                download = download_info.value
                if save_to:
                    save_path = os.path.join(save_to, download.suggested_filename)
                else:
                    save_path = os.path.abspath(download.suggested_filename)
                os.makedirs(os.path.dirname(save_path), exist_ok=True)
                download.save_as(save_path)
                downloads.append(save_path)
            return downloads
        except TimeoutError:
            if self.page.locator("div[id='ms-error-content']").count() == 1:
                notification: str = (
                    self.page.locator("div[id='ms-error-content']").text_content().strip().split("\n")[0]
                )
                self.logger.error(notification)
                if (
                    notification
                    == "このドキュメントにアクセスできません。このドキュメントを共有するユーザーに連絡してください。"
                ):
                    return []
            return self.download_files(url, file, steps, save_to)

    def upload_folder(
        self,
        url: str,
        folder_path: str,
    ):
        def empty_folder(path: str) -> bool:
            for _, _, files in os.walk(path):
                if files:
                    return False
            return True

        try:
            self.page.bring_to_front()
            self.page.goto(url=url)
            self.page.wait_for_selector(
                selector="div[class='app-container']",
                timeout=10000,
                state="visible",
            )
            if not empty_folder(folder_path):
                self.page.locator("button[data-automationid='uploadCommand']").click()
                with self.page.expect_file_chooser() as expect_file_chooser:
                    self.page.locator("button[data-automationid='uploadFolder']").click()
                    file_chooser = expect_file_chooser.value
                    file_chooser.set_files(folder_path)
                    self.page.wait_for_selector(
                        "div[class^='toastInnerContainer-'] i[data-icon-name='Cancel']", state="visible", timeout=15000
                    )
                    message: str = self.page.wait_for_selector(
                        "div[class^='toastInnerContainer-'] div[class^='messageRow-']", state="visible"
                    ).text_content()
                    self.logger.info(message)
                    return True
            else:
                self.page.locator("button[data-automationid='newCommand']").click()
                self.page.locator("button[data-automationid='newFolderCommand']").click()
                self.page.locator("input[aria-label='Name']").fill(os.path.basename(folder_path))
                self.page.locator("button[data-automation-id='Create']").click()
                try:
                    self.page.wait_for_selector("div[role='dialog']", state="detached")
                    self.logger.info(f"Upload folder: {os.path.basename(folder_path)}")
                    return True
                except TimeoutError:
                    error_message: str = self.page.locator("span[data-automation-id='error-message']").text_content()
                    self.logger.error(error_message)
                    return False
        except TimeoutError:
            return self.upload_folder(url, folder_path)
