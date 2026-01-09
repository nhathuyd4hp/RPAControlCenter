import os
import re
import time
from pathlib import Path
from typing import List, Literal

from pywinauto import Application
from pywinauto.application import WindowSpecification
from pywinauto.timings import wait_until_passes
from selenium.common.exceptions import (ElementClickInterceptedException,
                                        StaleElementReferenceException,
                                        TimeoutException)
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

from src.robot.Sakura.automation.bot.base import IBot
from src.robot.Sakura.automation.common import retry_if_exception


class MailDealer(IBot):
    
    def __init__(self, url: str, username: str, password: str, **kwargs):
        super().__init__(**kwargs)
        self.url = url
        self.authenticated = self._authentication(username, password)
        if not self.authenticated:
            raise ConnectionRefusedError("The username or password is incorrect.")
        
    @retry_if_exception(
        failure_return=False
    )
    def _authentication(self,username:str,password:str) -> bool:
        self.navigate(self.url)
        self.wait.until(
            method = EC.element_to_be_clickable((By.ID, "fUName")),
            message = "Not Found Username Field",
        ).send_keys(username)
        self.wait.until(
            method = EC.element_to_be_clickable((By.ID, "fPassword")),
            message="Not Found Password Field",
        ).send_keys(password)
        self.wait.until(
            method=EC.element_to_be_clickable((By.CSS_SELECTOR, "input[value='ログイン']")),
            message="Not Found Login Button",
        ).click()
        time.sleep(self.retry_interval)
        while self.browser.execute_script("return document.readyState") != "complete":
            time.sleep(self.retry_interval)
        try:
            error = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div[class='d_error_area ']"),
                ),
            )
            self.logger.error(error.text)
            return False
        except TimeoutException:
            if self.browser.current_url.find("app") != -1:
                self.logger.info("Authenticated")
                return True
            return True

    @retry_if_exception(
        exceptions=(
            ElementClickInterceptedException,
            StaleElementReferenceException,
            TimeoutException,
        ),
        failure_return=False
    )
    def send_mail(self,fr:str,to:str,subject:str,content:str,attachments:List[str] = [],mode:Literal["送信確認","承認依頼確認","一時保存"] = "一時保存") -> bool:
        for window in self.browser.window_handles:
            if window != self.root_window:
                self.browser.switch_to.window(window)
                self.browser.close()
        self.browser.switch_to.window(self.root_window)
        self.logger.info(f"Send mail from {fr} to {to} (subject: {subject})")
        self.wait.until(
            method = EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe[id='ifmSide']")),
            message = "Not Found Side Frame",
        )
        self.wait.until(
            method = EC.element_to_be_clickable((By.CSS_SELECTOR, "span[title='メール作成']")),
            message="Not Found Create Mail Button",
        ).click()
        while len(self.browser.window_handles) == 1:
            continue
        new_window = self.browser.window_handles[-1]
        self.browser.switch_to.window(new_window)
        self.browser.maximize_window()
        while self.browser.execute_script("return document.readyState") != "complete":
            time.sleep(self.retry_interval)
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[text()='次へ']"))
        ).click()
        # From
        fr_input = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='fFrom']"))
        )
        fr_input.clear()
        fr_input.send_keys(fr)
        # To
        to_input = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='fTo[]']"))
        )
        to_input.clear()
        to_input.send_keys(to)
        # Subject
        subject_input = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR,"input[id='fSubject']"))
        )
        subject_input.clear()
        subject_input.send_keys(subject)
        # Body
        self.wait.until(
            EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe"))
        )
        body = self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        self.browser.execute_script(
            "arguments[0].innerText = arguments[1];", body, content
        )
        self.browser.switch_to.default_content()
        for attachment in attachments:
            # Check
            if not os.path.exists(attachment):
                raise FileNotFoundError(f"Attachment Not Found: {attachment}")
            # Attachment
            self.wait.until(
                method = EC.element_to_be_clickable((By.CSS_SELECTOR, "button[title='添付ファイル']")),
                message="Not Found Attachment Button",
            ).click()
            file_uploader: WindowSpecification = wait_until_passes(
                timeout=self.timeout,
                retry_interval=self.retry_interval,
                func=lambda: Application()
                .connect(title="Open")
                .window(title="Open"),
            )
            wait_until_passes(
                timeout=self.timeout,
                retry_interval=self.retry_interval,
                func=lambda dialog=file_uploader: dialog 
                .child_window(class_name="ComboBoxEx32")
                .child_window(class_name="Edit")
                .wrapper_object(),
            ).type_keys(re.sub(r"([{}^%~()])", r"{\1}", str(Path(attachment))),with_spaces=True)
            wait_until_passes(
                timeout=self.timeout,
                retry_interval=1,
                func=lambda dialog=file_uploader: dialog.child_window(
                    title="&Open", class_name="Button"
                ).wrapper_object(),
            ).click_input()
            uploader_itembox = self.wait.until(
                EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    "div[class='md-uploader-itembox']"
                ))
            )
            items = uploader_itembox.find_elements(By.XPATH,'./div')
            items = [item.text.strip() for item in items]
            if os.path.basename(attachment) not in items:
                return self.send_mail(
                    fr=fr,
                    to=to,
                    subject=subject,
                    content=content,
                    attachments=attachments,
                    mode=mode,
                )
        for button in self.wait.until(
            method = EC.presence_of_all_elements_located((By.CSS_SELECTOR,"button[class='accent-btn__btn']")),
            message="Not Found Send Button"
        ):
            if button.text.find(mode) != -1:
                self.browser.execute_script(
                    "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                    button,
                )
                time.sleep(1)
                self.wait.until(EC.element_to_be_clickable(button)).click()
                time.sleep(1)
                snackbar = self.wait.until(
                    method = EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='snackbar__msg']")),
                    message= "Snackbar Not Found ", 
                )
                self.logger.info(snackbar.text)
                break
        self.browser.close()
        self.browser.switch_to.window(self.root_window)
        self.logger.info("Mail sent successfully")
        return True
