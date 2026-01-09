import os
import re
import time
from datetime import date
from typing import List
from urllib.parse import urljoin

import pandas as pd
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

from src.robot.Sakura.automation.bot.base import IBot, IBotMeta
from src.robot.Sakura.automation.common import retry_if_exception


class WebAccessMeta(IBotMeta):
    def __new__(cls, name, bases, class_dict):
        for attr_name, attr_value in class_dict.items():
            if callable(attr_value) and attr_name not in [
                "__init__",
                "_authentication",
            ]:
                class_dict[attr_name] = cls.login(attr_value)
        return super().__new__(cls, name, bases, class_dict)

    @staticmethod
    def login(func):
        def wrapper(self, *args, **kwargs):
            if not self.authenticated:
                raise ConnectionRefusedError("Yêu cầu xác thực")
            if self.browser.current_url.endswith(".com/") or self.browser.current_url == "data:,":
                self.authenticated = self._authentication(self._username, self._password)
            return func(self, *args, **kwargs)

        return wrapper


class WebAccess(IBot, metaclass=WebAccessMeta):
    def __init__(
        self,
        url: str,
        username: str,
        password: str,
        **kwargs,
    ):
        super().__init__(**kwargs)
        self.url = url
        self._username = username
        self._password = password
        self.authenticated = self._authentication(self._username, self._password)
        if not self.authenticated:
            raise ConnectionRefusedError("The username or password is incorrect.")

    @retry_if_exception(
        delay=2,
        failure_return=False,
    )
    def _authentication(self, username: str, password: str) -> bool:
        self.navigate(self.url)
        self.wait.until(
            method=EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='text']")),
            message="Không tìm thấy usernameField",
        ).send_keys(username)
        self.wait.until(
            method=EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='password']")),
            message="Không tìm thấy passwordField",
        ).send_keys(password)
        self.wait.until(
            method=EC.element_to_be_clickable((By.CSS_SELECTOR, "button[class='btn login']")),
            message="Không tìm thấy loginButton",
        ).click()
        try:
            error_box = self.wait.until(
                method=EC.presence_of_element_located((By.CSS_SELECTOR, "div[id='f-error-box']"))
            )
            data = error_box.find_element(By.CSS_SELECTOR, "div[class='data']")
            self.logger.info(f"Xác thực thất bại!: {data.text}")
            return False
        except TimeoutException:
            self.logger.info("Authenticated")
            return True

    @retry_if_exception(
        exceptions=(
            StaleElementReferenceException,
            ElementClickInterceptedException,
        ),
    )
    def modify(self, case_number: int, **kwargs) -> bool:
        self.logger.info(f"Update {case_number}: {kwargs}")
        # --- Redirect --- #
        f_menus = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[id='f-menus']")))
        if order_list := f_menus.find_elements(By.CSS_SELECTOR, "a[title='受注一覧']"):
            href = order_list[0].get_attribute("href")
            self.navigate(href)
        else:
            raise Exception("Can't find order view")
        # Clear
        clear_btn = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[type='reset']")))
        self.browser.execute_script(
            "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
            clear_btn,
        )
        time.sleep(self.retry_interval)
        self.wait.until(EC.element_to_be_clickable(clear_btn)).click()
        # Input Case Number
        case_number_td = self.wait.until(
            EC.presence_of_element_located((By.XPATH, "//th[span[text()='案件番号']]/following-sibling::td"))
        )
        self.browser.execute_script(
            "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
            case_number_td,
        )
        time.sleep(self.retry_interval)
        input_field = case_number_td.find_element(By.TAG_NAME, "input")
        self.wait.until(EC.element_to_be_clickable(input_field)).send_keys(case_number)
        # Search
        self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))).click()
        time.sleep(self.retry_interval)
        while self.browser.execute_script("return document.readyState") != "complete":
            time.sleep(self.retry_interval)
        table = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table[id='orderlist']")))
        # Data in tbody
        tbody = table.find_element(By.TAG_NAME, "tbody")
        trs = tbody.find_elements(By.TAG_NAME, "tr")
        if len(trs) == 0:
            raise LookupError(f"Not Found Case Number: {case_number}")
        if trs[0].text == "検索結果はありません":
            raise LookupError(f"Not Found Case Number: {case_number}")
        hrefs = []
        for tr in trs:
            refer_btn = tr.find_element(By.CSS_SELECTOR, 'input[value="参照"]')
            self.browser.execute_script(
                "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                refer_btn,
            )
            time.sleep(self.retry_interval)
            self.wait.until(EC.element_to_be_clickable(refer_btn))
            onclick = refer_btn.get_attribute("onclick")
            match = re.search(r"location\.href='([^']+)'", onclick)
            if not match:
                raise Exception(f"Update {case_number} failed")
            hrefs.append(urljoin(self.url, match.group(1)))
        time.sleep(self.retry_interval * 2)
        for href in hrefs:
            self.navigate(href)
            # --- Serializer Kwargs --- #
            for key in kwargs.keys():
                if not isinstance(kwargs.get(key), list):
                    kwargs[key] = [kwargs.get(key)]
            # --- Update Web Access --- #
            for key in kwargs.keys():
                try:
                    ths = self.wait.until(EC.presence_of_all_elements_located((By.XPATH, f"//th[text()='{key}']")))[
                        : len(kwargs.get(key, None))
                    ]
                    for insert_index, th in enumerate(ths):
                        self.browser.execute_script(
                            "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                            th,
                        )
                        time.sleep(self.retry_interval)
                        table = th.find_element(By.XPATH, "./ancestor::table[1]")
                        tbody = table.find_element(By.TAG_NAME, "tbody")
                        start_index = None  # Cột Header
                        end_index = None  # Cột Header Tiếp theo
                        for index, tr in enumerate(tbody.find_elements(By.TAG_NAME, "tr")):
                            if tr.find_elements(By.XPATH, f".//th[text()='{key}']"):
                                start_index = index
                                break
                        if start_index is None:
                            raise LookupError("Không tìm được vị trí dòng dữ liệu!")
                        for index, tr in enumerate(tbody.find_elements(By.TAG_NAME, "tr")[start_index + 1 :]):
                            if tr.find_elements(By.CSS_SELECTOR, "th[class='c']"):
                                end_index = index + start_index + 1
                                break
                        table = tbody.find_elements(By.TAG_NAME, "tr")[start_index:end_index]
                        header = table[0]
                        pos_index = None
                        for index, column in enumerate(header.find_elements(By.TAG_NAME, "th")):
                            if column.text == key:
                                pos_index = index
                                break
                        if pos_index is None:
                            raise LookupError("No Data Column")
                        for tr in tbody.find_elements(By.TAG_NAME, "tr")[start_index + 1 : end_index]:
                            td = tr.find_elements(By.TAG_NAME, "td")[pos_index]
                            input_field = td.find_element(By.XPATH, ".//input | .//select")
                            self.wait.until(EC.element_to_be_clickable(input_field)).click()
                            self.browser.execute_script(
                                "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                                input_field,
                            )
                            time.sleep(self.retry_interval)
                            if input_field.tag_name == "input":
                                if kwargs[key][insert_index].get("mode") == "replace":
                                    input_field.clear()
                                    insert_data = kwargs[key][insert_index].get("data")
                                    input_field.send_keys(insert_data)
                                    input_field.send_keys(Keys.ENTER)
                                if kwargs[key][insert_index].get("mode") == "append":
                                    input_field.click()
                                    insert_data = f" {kwargs[key][insert_index].get('data')}"
                                    if not input_field.get_attribute("value").endswith(insert_data):
                                        input_field.send_keys(insert_data)
                                        input_field.send_keys(Keys.ENTER)
                                if kwargs[key][insert_index].get("mode") == "insert":
                                    input_field.click()
                                    input_field.send_keys(Keys.HOME)
                                    insert_data = f"{kwargs[key][insert_index].get('data')} "
                                    if not input_field.get_attribute("value").startswith(insert_data):
                                        input_field.send_keys(insert_data)
                                        input_field.send_keys(Keys.ENTER)
                            elif input_field.tag_name == "select":
                                select = Select(input_field)
                                select.select_by_visible_text(kwargs[key][insert_index].get("data"))
                except TimeoutException:
                    pass
            time.sleep(self.retry_interval)
            update_btn = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[id='order_update']")))
            self.browser.execute_script("window.scrollTo(0, 0);")
            time.sleep(self.retry_interval * 2)
            self.wait.until(EC.element_to_be_clickable(update_btn)).click()
            time.sleep(self.retry_interval)
            while self.browser.execute_script("return document.readyState") != "complete":
                time.sleep(self.retry_interval)
            notice = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[id='f-notice-box']")))
            self.logger.info(f"Update Case: {case_number}: {repr(notice.text)}")
            if notice.text != "案件情報を更新しました。\n閉じる":
                return False
        return True

    @retry_if_exception(
        exceptions=(
            StaleElementReferenceException,
            ElementClickInterceptedException,
            NotImplementedError,
        ),
        delay=2,
        failure_return=pd.DataFrame(),
    )
    def get_order_list(
        self,
        building_name: str | None = None,
        drawing: list[str] | None = None,
        delivery_date: list[str] = [date.today().strftime("%Y/%m/%d"), ""],  # noqa
        fields: list[str] | None = None,
    ) -> pd.DataFrame:
        f_menus = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[id='f-menus']")))
        if order_list := f_menus.find_elements(By.CSS_SELECTOR, "a[title='受注一覧']"):
            href = order_list[0].get_attribute("href")
            self.navigate(href)
        else:
            self.logger.error("Can't find Order List")
            return pd.DataFrame()
        # Clear
        self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[type='reset']")))
        self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='reset']"))).click()
        # Filter
        if building_name:
            time.sleep(1)
            self.wait.until(EC.element_to_be_clickable((By.ID, "select2-search_builder_cd-container"))).click()
            time.sleep(1)
            self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[class="select2-search__field"]'))
            ).send_keys(building_name)
            options: List[WebElement] = self.wait.until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul#select2-search_builder_cd-results > li"))
            )
            if len(options) > 1:
                for option in options:
                    self.logger.info(f"Found: {option.text}")
                raise LookupError(f"Found {len(options)} options")
            option = options[0]
            if option.text == "No results found":
                raise LookupError("No results found")
            self.logger.info(f"ビルダー名: {option.text}")
            self.wait.until(EC.element_to_be_clickable(option)).click()
            time.sleep(1)
        if delivery_date:  # Delivery date
            self.logger.info(f"確定納品日: {delivery_date[0]} ~ {delivery_date[1]}")
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="search_fix_deliver_date_from"]'))
            ).clear()
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="search_fix_deliver_date_from"]'))
            ).send_keys(delivery_date[0])
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="search_fix_deliver_date_to"]'))
            ).clear()
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="search_fix_deliver_date_to"]'))
            ).send_keys(delivery_date[1])
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="search_fix_deliver_date_to"]'))
            ).send_keys(Keys.ESCAPE)
        if drawing:  # Drawing
            self.logger.info(f"Filter Drawing: {drawing}")
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[id='search_drawing_type_ms']")))
            self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[id='search_drawing_type_ms']"))
            ).click()
            for e in drawing:
                xpath = f"//span[text()='{e}']"
                self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
            self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[id='search_drawing_type_ms']"))
            ).send_keys(Keys.ESCAPE)
        # Search
        submit_btn = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[type='submit']")))
        self.browser.execute_script(
            "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
            submit_btn,
        )
        time.sleep(self.retry_interval)
        self.wait.until(EC.element_to_be_clickable(submit_btn)).click()
        time.sleep(self.retry_interval)
        while self.browser.execute_script("return document.readyState") != "complete":
            time.sleep(self.retry_interval)
        # Download File
        download_btn = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a[class='button fa fa-download']"))
        )
        self.browser.execute_script(
            "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
            download_btn,
        )
        time.sleep(self.retry_interval)
        self.wait.until(EC.element_to_be_clickable(download_btn)).click()
        time.sleep(self.timeout)
        file_name, tag = self.wait_for_download_to_finish()
        if tag:
            self.logger.error(tag)
            time.sleep(self.timeout)
            raise NotImplementedError(f"Lấy danh sách đơn hàng thất bại: {tag}")
        file_path = os.path.join(self.download_directory, file_name)
        try:
            df = pd.read_csv(file_path, encoding="CP932", low_memory=False)
        except UnicodeDecodeError:
            df = pd.read_csv(file_path, low_memory=False)
        df = df[fields] if fields else df
        os.remove(file_path)
        return df


__all__ = ["WebAccess"]
