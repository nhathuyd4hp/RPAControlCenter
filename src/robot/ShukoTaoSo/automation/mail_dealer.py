import logging
import os
import re
import shutil
import time
from contextlib import suppress

import pandas as pd
from playwright._impl._errors import TimeoutError
from playwright.sync_api import sync_playwright


class MailDealer:
    def __init__(
        self,
        domain: str,
        username: str,
        password: str,
        context=None,
        browser=None,
        playwright=None,
        logger: logging.Logger = logging.getLogger("MailDealer"),
        headless: bool = False,
        timeout: float = 5000,
    ):
        self._external_context = context is not None
        self._external_browser = browser is not None
        self._external_pw = playwright is not None
        self.timeout = timeout
        if playwright:
            self._pw = playwright
        else:
            self._pw = sync_playwright().start()

        if browser:
            self.browser = browser
        else:
            self.browser = self._pw.chromium.launch(
                headless=headless,
                timeout=self.timeout,
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
            self.page.goto(self.domain, wait_until="domcontentloaded")
            self.page.locator(selector="input[id='fUName']").fill(self.username)
            self.page.locator(selector="input[id='fPassword']").fill(self.password)
            with self.page.expect_navigation(
                wait_until="networkidle",
                timeout=30000,
            ):
                self.page.locator(selector="input[type='submit'][value='ログイン']").click()
            self.page.locator("button[title='設定']").click(click_count=2, delay=0.25)
            with suppress(TimeoutError):
                self.page.wait_for_selector("div[id='md_dialog']", timeout=5000)
                self.page.locator("input[type='button'][id='md_dialog_submit']").click()
            account: str = self.page.wait_for_selector(
                selector="span[class*='olv-u-fc--user-icon_text'][class*='olv-u-bg--user-icon_bg'][class*='thumbnail__text']",
                state="visible",
                timeout=self.timeout,
            ).get_attribute("title")
            self.logger.info(f"Logged in as: {account}")
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

    def mail_lists(self, mailbox: str) -> pd.DataFrame:
        try:
            # Switch Mailbox
            self.page.bring_to_front()
            self.page.frame(name="side").locator(f"span[title='{mailbox}']").click()
            # Get Table
            self.page.frame(name="main").wait_for_selector("table", state="visible")
            columns = self.page.frame(name="main").locator("table thead")
            columns: list[str] = [
                columns.locator("th").nth(i).text_content() for i in range(columns.locator("th").count())
            ]
            Rows = self.page.frame(name="main").locator("table tbody")
            data = [
                [Rows.nth(i).locator("td").nth(j).text_content() for j in range(Rows.nth(i).locator("td").count())]
                for i in range(Rows.count())
            ]
            if self.page.frame(name="main").locator("div[class='olv-p-maillist__no-data']").count() == 1:
                self.logger.warning("条件に一致するデータがありません。")
                return pd.DataFrame(columns=columns)
            return pd.DataFrame(data=data, columns=columns)
        except TimeoutError:
            return self.mail_lists(mailbox)

    def generate(self, mail_id: str) -> tuple[str, str, str]:
        try:
            self.page.bring_to_front()
            self.page.locator("input[name='fDQuery[B]']").clear()
            self.page.locator("input[name='fDQuery[B]']").fill(mail_id)
            self.page.locator("button[title='検索']").click()
            self.page.wait_for_selector("div[class='loader']", state="visible", timeout=10000)
            self.page.wait_for_selector("div[class='loader']", state="hidden", timeout=30000)
            # ---- #
            body = self.page.frame(name="main").locator("div[class='olv-p-mail-page__content']").text_content()
            trigger = self.page.frame(name="main").locator("div[class='dropdown__trigger']").text_content()
            if trigger == "案件化済":
                return False, None, None, trigger, None
            andpad_link = re.search(r"https://andpad.jp/my/orders/(\d*)", body)
            if not andpad_link:
                return False, None, None, "Không tìm thấy link andpad", None
            with self.context.new_page() as andpad:
                andpad.goto(andpad_link.group(0))
                orders_name = andpad.locator("div[class='order-info__name']").text_content()
                address = andpad.locator("tr:has(th:has-text('住所')) td").text_content()
                address = " ".join(address.split(" ")[1:])
            self.page.bring_to_front()
            # ---- #
            self.page.locator("button[title='その他']").click()
            with self.page.expect_popup() as popup:
                self.page.locator("button", has_text="案件管理").click()
                popup = popup.value
                # ---- Check
                popup.locator("select[name='fDQuery[statusid]']").select_option("すべて")
                time.sleep(1 / 3)
                popup.locator("select[id='itemType']").select_option("案件名")
                time.sleep(1 / 3)
                popup.locator("input[id='searchText']").fill(orders_name)
                time.sleep(1 / 3)
                with popup.expect_navigation(wait_until="networkidle"):
                    popup.locator("button[type='submit']").click()
                with suppress(TimeoutError):
                    table = popup.locator("table[class='mokoForTabTable']")
                    time.sleep(1 / 3)
                    if table.locator("table[class='mail-list-table']").count() != 1:
                        raise TimeoutError("")
                    table = table.locator("table[class='mail-list-table']")
                    for i in range(table.locator("tr").count()):
                        if i == 0:
                            continue
                        data = table.locator("tr").nth(i).text_content().replace("\xa0", "").split("\n")
                        cong_trinh = data[3]
                        if cong_trinh == "秀光ビルド":
                            popup.close()
                            return False, None, None, "Đã có số", None
                popup.locator(selector="form a[href^='viewMatter.php?fCID=']").click()
                popup.locator("input[name='fDQuery[custmername]']").fill("秀光ビルド")
                time.sleep(1 / 3)
                popup.locator("input[name='fDQuery[mattername]']").fill(orders_name)
                time.sleep(1 / 3)
                popup.locator("input[value='登録']").click()
                time.sleep(1 / 3)
                fMatterID = popup.locator("th:has-text('案件ID') + td").text_content().strip()
                popup.close()
            self.page.frame(name="main").locator("button[title='一括操作']").click()
            self.page.frame(name="main").locator("div[class='pop-panel__content'] input[id='fMatterID_add']").fill(
                fMatterID
            )
            self.page.frame(name="main").locator(
                "div[class='pop-panel__content'] input[name='fAddMatterRelByMGID'] + div.checkbox__indicator"
            ).check()
            self.page.frame(name="main").locator(
                "div[class='pop-panel__content'] input[id='fMatterID_add'] + button", has_text="関連付ける"
            ).click()
            save_path = f"downloads/{fMatterID} {orders_name}"
            os.makedirs(save_path, exist_ok=True)
            with self.context.new_page() as andpad:
                andpad.goto(andpad_link.group(0))
                andpad.locator("a", has_text=re.compile("^資料$")).click()
                andpad.locator("a", has_text=re.compile("^一覧$")).click()
                for keyword in ["案内図", "平面図", "プレカット図", "依頼書", "工程表", "矩計図", "電気図"]:
                    andpad.locator("input[name='q[filename_cont]']").fill(keyword)
                    with andpad.expect_navigation(wait_until="networkidle"):
                        andpad.locator("input[id='search-btn']").click()
                    for i in range(andpad.locator("a[class='btn-small'][download]").count()):
                        with andpad.expect_download() as expect_download:
                            andpad.locator("a[class='btn-small'][download]").nth(i).click()
                            download = expect_download.value
                            download.save_as(os.path.join(save_path, download.suggested_filename))
            save_path = os.path.abspath(save_path)
            for file in os.listdir(save_path):
                if "案内図" in file:
                    dst = os.path.join(save_path, "★案内図")
                    os.makedirs(dst, exist_ok=True)
                    shutil.move(src=os.path.join(save_path, file), dst=os.path.join(dst, file))
                else:
                    dst = os.path.join(save_path, "資料")
                    os.makedirs(dst, exist_ok=True)
                    shutil.move(src=os.path.join(save_path, file), dst=os.path.join(dst, file))
            self.page.bring_to_front()
            self.page.frame(name="main").locator("div[class='dropdown__trigger']").click()
            self.page.frame(name="main").locator(
                "li[class='list__item is-multiple has-value']", has_text=re.compile("^案件化済$")
            ).click()
            return True, fMatterID, orders_name, save_path, address
        except TimeoutError:
            if self.page.frame(name="main").locator("div[class='alert alert-error']").count() == 1:
                notification = self.page.frame(name="main").locator("div[class='alert alert-error']").text_content()
                self.logger.error(notification)
                time.sleep(2.5)
                return self.generate(mail_id)
            if self.page.frame(name="main").locator("div[class^='mailList-no-tab']").count() == 0:
                return self.generate(mail_id)
            notification = self.page.frame(name="main").locator("div[class^='mailList-no-tab']").text_content().strip()
            if notification == "検索条件に一致するデータがありません。":
                self.logger.error(notification)
                return False, None, None, notification, None
            return self.generate(mail_id)
