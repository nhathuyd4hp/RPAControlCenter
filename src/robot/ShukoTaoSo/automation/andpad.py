import logging
import time

from playwright._impl._errors import TimeoutError
from playwright.sync_api import sync_playwright


class AndPad:
    def __init__(
        self,
        domain: str,
        username: str,
        password: str,
        context=None,
        browser=None,
        playwright=None,
        logger: logging.Logger = logging.getLogger("AndPad"),
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
        self.username = username
        self.password = password
        self.timeout = timeout
        self.logger = logger
        if not self.__authentication():
            raise PermissionError("Authentication failed")

    def __authentication(self) -> bool:
        try:
            self.page.bring_to_front()
            self.page.goto(self.domain, wait_until="networkidle")
            with self.page.expect_navigation(wait_until="domcontentloaded"):
                self.page.locator("input[value='ログイン画面へ']").click()
            self.page.locator("input[id='email']").fill(self.username)
            self.page.locator("input[id='password']").fill(self.password)
            time.sleep(1)
            with self.page.expect_navigation(wait_until="networkidle"):
                self.page.locator("button[id='btn-login']").click()
            self.page.locator("span", has_text="ログアウト").wait_for(timeout=10000, state="visible")
            self.logger.info(f"Logged in as: {self.username}")
            return True
        except TimeoutError:
            if self.page.locator("p[id='error-message-login-failed-text']").count() == 1:
                error = self.page.locator("p[id='error-message-login-failed-text']").text_content()
                self.logger.error(error)
                return False
            self.logger.error("RETRY __authentication")
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

    def send_message(
        self, object_name: str, message: str, tags: list[str] | None = None, attachments: list[str] | None = None
    ):
        try:
            self.page.bring_to_front()
            with self.page.expect_navigation(wait_until="networkidle"):
                self.page.goto(self.domain)
            textbox = self.page.locator("input[class='search__textbox']")
            textbox.fill(object_name)
            with self.page.expect_navigation(wait_until="networkidle"):
                textbox.press("Enter")
            self.page.wait_for_selector("table[class='table']", state="visible")
            if self.page.locator("table[class='table'] > tbody > tr").count() != 1:
                return False
            with self.page.expect_navigation(wait_until="networkidle"):
                self.page.locator("table[class='table'] > tbody > tr").click()
            with self.page.expect_popup() as popup:
                self.page.locator("a", has_text="チャット").last.click()
                popup = popup.value
                popup.wait_for_load_state(state="domcontentloaded")
                message_box = popup.wait_for_selector("textarea[placeholder='メッセージを入力']", state="visible")
                time.sleep(1)
                message_box.fill(message)
                if tags:
                    for tag in tags:
                        popup.locator("button", has=popup.locator("span", has_text="お知らせ")).click()
                        time.sleep(1.5)
                        popup.locator("input[placeholder='氏名で絞込']").fill(tag)
                        time.sleep(1.5)
                        notify_list = popup.locator("label[data-test='notify-member-cell']")
                        if notify_list.count() != 1:
                            return False
                        if not popup.locator(
                            "label[data-test='notify-member-cell'] > input[type='checkbox']"
                        ).is_checked():
                            popup.locator("label[data-test='notify-member-cell'] > input[type='checkbox']").click()
                        popup.locator(selector="wc-tsukuri-text", has_text="選択").click()
                popup.close()
                return True
        except TimeoutError:
            self.logger.error("RETRY send_message")
            return self.send_message(object_name, message)
