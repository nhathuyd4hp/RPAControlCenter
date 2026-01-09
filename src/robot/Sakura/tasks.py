import os
import re
from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal

import pandas as pd
import xlwings as xw
from celery import shared_task
from selenium import webdriver
from xlwings.main import Sheet
from xlwings.utils import col_name

from src.core.config import settings
from src.robot.Sakura.automation.bot import SharePoint, WebAccess
from src.service import ResultService as minio


def from_date_default() -> datetime:
    now = datetime.now()
    if now.month == 1:
        return now.replace(year=now.year - 1, month=12, day=21)
    return now.replace(month=now.month - 1, day=21)


def to_date_default() -> datetime:
    now = datetime.now()
    return now.replace(day=20)


# -- Chrome Options
options = webdriver.ChromeOptions()
options.add_argument("--disable-notifications")
options.add_argument("--disable-logging")
options.add_argument("--log-level=3")
options.add_argument("--silent")
options.add_experimental_option("excludeSwitches", ["enable-logging"])
options.add_experimental_option(
    "prefs",
    {
        "download.default_directory": os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads"),
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True,
    },
)


def main():
    # From To
    to_date = datetime.now().replace(day=20)
    if to_date.month == 1:
        from_date = to_date.replace(year=to_date.year - 1, month=12, day=21)
    else:
        from_date = to_date.replace(month=to_date.month - 1, day=21)
    from_date = from_date.strftime("%Y/%m/%d")
    to_date = to_date.strftime("%Y/%m/%d")
    # Download
    with WebAccess(
        url="https://webaccess.nsk-cad.com/",
        username="hanh0704",
        password="159753",
        log_name="Web Access",
        options=options,
    ) as web_access:
        data = web_access.get_order_list(
            building_name="009300",
            delivery_date=[from_date, to_date],
            fields=[
                "案件番号",
                "得意先名",
                "物件名",
                "確未",
                "確定納期",
                "曜日",
                "追加不足",
                "配送先住所",
                "階",
                "資料リンク",
            ],
        )
    if data.empty:
        return
    # ---- Download files ----
    prices = []
    data = data[data["追加不足"] != "不足"]
    with SharePoint(
        url="https://nskkogyo.sharepoint.com/",
        username="vietnamrpa@nskkogyo.onmicrosoft.com",
        password="Robot159753",
        log_name="SharePoint",
        options=options,
    ) as share_point:
        for url in data["資料リンク"]:
            downloads = share_point.download(
                site_url=url,
                file_pattern="(見積書|見積もり)/.*.(xlsm|xlsx|xls)$",
            )
            if not downloads:
                raise RuntimeError("Không có file")
            if all(status for _, _, status in downloads):
                prices.append(downloads[0][2])
                continue
            for _, file, status in downloads:
                price = None
                found = False
                if status:
                    continue
                for sheet in pd.ExcelFile(file, engine="openpyxl", engine_kwargs={"read_only": True}).sheet_names:
                    sheet: pd.DataFrame = pd.read_excel(file, sheet_name=sheet)
                    for _, row in sheet.iterrows():
                        row: str = " ".join(str(cell) for cell in row)
                        if match := re.search(r"税抜金額[^\d]*([\d,]+(?:\.\d+)?)", row):
                            price = match.group(1)
                            price = Decimal(price).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
                            if price != 0:
                                found = True
                            break
                    if found:
                        break
                prices.append(price)
    # Process
    data["金額（税抜）"] = prices
    data["金額（税抜）"] = pd.to_numeric(data["金額（税抜）"], errors="coerce").fillna(0)
    data.drop(columns=["資料リンク"], inplace=True)
    # Append Row
    empty_row = pd.Series({col: pd.NA for col in data.columns})
    append_row = {col: pd.NA for col in data.columns}
    append_row[list(data.columns)[-3]] = "合計"
    append_row[list(data.columns)[-1]] = data["金額（税抜）"].sum()
    data = pd.concat([data, pd.DataFrame([empty_row.to_dict(), append_row])], ignore_index=True)
    # Save
    excel_file = f"{datetime.today().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    data.to_excel(
        excel_file,
        index=False,
    )
    try:
        app = xw.App(visible=False)
        wb = app.books.open(excel_file)
        sheet: Sheet = wb.sheets[0]
        # AutoFitColumn
        sheet.autofit()
        # Header
        sheet.api.PageSetup.LeftHeader = f"さくら建設　鋼製野縁納材報告（{from_date}-{to_date}）　"
        sheet.api.PageSetup.RightHeader = datetime.now().strftime("%Y/%m/%d")
        # Format
        ## - Tô màu Header
        sheet.range(f"A1:{col_name(data.shape[1])}1").color = (166, 166, 166)
        ## - Tô màu ô "合計"
        sheet.range(f"H{data.shape[0] + 1}").color = (166, 166, 166)
        ## - Định dang cột J 金額（税抜）(12345 -> 12,345)
        sheet.range(f"J2:J{data.shape[0] + 1}").number_format = "#,##0"
        # Landspace
        sheet.api.PageSetup.Orientation = 2
        # All Border
        rng = sheet.range(f"A1:J{data.shape[0] + 1}")
        for i in [7, 8, 9, 10, 11, 12]:
            border = rng.api.Borders(i)
            border.LineStyle = 1
            border.Weight = 2
            border.ColorIndex = 0
        wb.save()
        pdfFile = f"さくら建設　鋼製野縁納材報告（{from_date} - {to_date}).pdf".replace("/", "-")
        # ---- Export in one page #
        sheet.api.PageSetup.Zoom = False
        sheet.api.PageSetup.FitToPagesWide = 1
        sheet.api.PageSetup.FitToPagesTall = 1
        sheet.to_pdf(pdfFile)
    finally:
        wb.close()
        app.quit()
        os.remove(excel_file)
        # with MailDealer(
        #     url = "https://mds3310.maildealer.jp/",
        #     username="vietnamrpa",
        #     password="nsk159753",
        #     log_name="MailDealer",
        #     options=options,
        # ) as mail_dealer:
        #     mail_dealer.send_mail(
        #         fr="kantou@nsk-cad.com",
        #         to="ikeda.k@jkenzai.com",
        #         subject=f"さくら建設　鋼製野縁納材報告書（{from_date}～{to_date}）",
        #         content=f"""
        #         ジャパン建材　池田様

        #         いつもお世話になっております。

        #         さくら建設　鋼製野縁納材報告書（{from_date}～{to_date}）
        #         を送付致しましたので、ご査収の程よろしくお願い致します。

        #         ----・・・・・----------・・・・・----------・・・・・-----

        #         　エヌ・エス・ケー工業㈱　横浜営業所
        #         中山　知凡
        #
        #         　〒222-0033
        #         　横浜市港北区新横浜２-４-６　マスニ第一ビル８F-B
        #         　TEL:(045)595-9165 / FAX:(045)577-0012
        #
        #         -----・・・・・----------・・・・・----------・・・・・-----
        #         """,
        #         attachments=[
        #             os.path.abspath(pdfFile),
        #         ],
        #     )
    return os.path.abspath(pdfFile)


@shared_task(
    bind=True,
    name="Sakura",
)
def Sakura(self):
    pdfFile = main()
    result = minio.fput_object(
        bucket_name=settings.MINIO_BUCKET,
        object_name=f"Sakura/{self.request.id}/{os.path.basename(pdfFile)}",
        file_path=pdfFile,
        content_type="application/pdf",
    )
    return result.object_name
