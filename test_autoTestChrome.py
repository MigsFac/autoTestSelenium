from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pytest_check as check
import unicodedata
import openpyxl
import xlwings as xw
import time
import csv
import pytest
import platform
import datetime
import re

os_info = platform.system()
if os_info == "Darwin":
    os_info = "macOS"
    os_version, _, _ = platform.mac_ver()
elif os_info == "Windows":
    os_version = platform.release()
else:
    os_version = platform.version()

# errors = []
file_path = "test_case.xlsm"
app = xw.App(visible=False)  # バックグラウンドで開く
app.display_alerts = False
workbook = app.books.open(file_path)
sheet = workbook.sheets["テストケース"]

search_number = "テストNo."
search_ri = "入力値(薬剤名)"
search_ri = unicodedata.normalize("NFKC", search_ri)
search_weight = "入力値(体重)"
search_weight = unicodedata.normalize("NFKC", search_weight)
search_result = "結果"
search_comment = "コメント"
search_selenium = "Selenium使用"
search_nuclide = "核種"
search_medicine = "薬剤正式名"
search_excel_calc = "計算最小値比較"
search_load_medicine = "負荷薬剤名"
search_load_dose = "負荷薬剤量"
search_ri_id = "薬剤id"
initial_row = 153  # 起動時間のテストケース行
calc_row = 185  # 計算時間のテストケース行
cnt = 0

title_row = 3
for row in range(1, sheet.range("A1").end("down").row + 1):
    if str(sheet.cells(row, 1).value).replace("\n", "").strip() == search_number:
        title_row = row
        break

for col in range(1, sheet.used_range.columns.count + 1):
    if cnt == 12:
        break
    cell = sheet.cells(title_row, col)
    cell = unicodedata.normalize(
        "NFKC", str(cell.value or "").replace("\n", "").strip()
    )

    if cell == search_number:
        number_column = col
        cnt += 1
        continue
    if cell == search_ri:
        input_ri_column = col
        cnt += 1
        continue
    if cell == search_weight:
        input_weight_column = col
        cnt += 1
        continue
    if cell == search_result:
        result_column = col
        cnt += 1
        continue
    if cell == search_comment:
        comment_column = col
        cnt += 1
        continue
    if cell == search_selenium:
        selenium_column = col
        cnt += 1
        continue
    if cell == search_nuclide:
        nuclide_column = col
        cnt += 1
        continue
    if cell == search_medicine:
        medicine_column = col
        cnt += 1
        continue
    if cell == search_excel_calc:
        excel_calc_dose_column = col
        cnt += 1
        continue
    if cell == search_load_medicine:
        load_medicine_column = col
        cnt += 1
        continue
    if cell == search_load_dose:
        load_dose_column = col
        cnt += 1
        continue
    if cell == search_ri_id:
        ri_id_column = col
        cnt += 1
        continue

data = []

for i in range(
    title_row + 1, sheet.used_range.rows.count + 1
):  # 小数で挙動見る時はここの値を10とか指定する　sheet.used_range.columns.count + 1
    if sheet.cells(i, selenium_column).value == "◯":
        data.append(
            (
                sheet.cells(i, number_column).value,
                sheet.cells(i, input_weight_column).value,
                sheet.cells(i, nuclide_column).value,
                sheet.cells(i, medicine_column).value,
                sheet.cells(i, excel_calc_dose_column).value,
                sheet.cells(i, load_medicine_column).value,
                sheet.cells(i, load_dose_column).value,
                sheet.cells(i, ri_id_column).value,
            )
        )  # テストNo.正式薬剤名,入力体重,核種,計算最小値比較,負荷薬剤名,負荷薬剤量


@pytest.fixture(scope="session", autouse=True)
def driver():
    # Chromeのドライバーを指定
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-extensions")
    driver = webdriver.Chrome(options=options)

    chrome_version = driver.capabilities["browserVersion"]

    sheet.range("H1").value = str(os_info) + " " + str(os_version)
    sheet.range("H2").value = "Google Chrome " + chrome_version

    yield driver

    now = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    workbook.save(now + "test_case.xlsm")
    workbook.close()
    app.quit()
    driver.quit()


@pytest.mark.parametrize(
    "test_number, input_weight,nuclide,medicine,excel_calc_dose,load_medicine,load_dose,ri_id",
    data,
)
def test_seleTest(
    driver,
    test_number,
    input_weight,
    nuclide,
    medicine,
    ri_id,
    excel_calc_dose,
    load_medicine,
    load_dose,
):
    driver.get("http://13.236.193.159:8080/")
    errors = []
    this_row = int(title_row) + int(test_number)
    start_time = time.time()
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "button.btn"))
    )
    end_time = time.time()
    initial_time = round(end_time - start_time, 2)
    if (
        sheet.cells(initial_row, result_column).value is None
        or sheet.cells(initial_row, result_column).value == ""
    ):
        if initial_time < 3:
            sheet.cells(initial_row, result_column).value = "OK"
        else:
            sheet.cells(initial_row, result_column).value = "NG"
        sheet.cells(initial_row, comment_column).value = (
            "起動時間：" + str(initial_time) + "秒"
        )

    input_ri = nuclide + " " + medicine
    # 体重入力
    weight_input = driver.find_element(By.ID, "kg")
    weight_input.clear()
    weight_input.send_keys(input_weight)
    time.sleep(1)

    # 薬のラジオボタンを選択
    ri_selector = f"input[value={ri_id}]"
    medicine_radio = driver.find_element(By.CSS_SELECTOR, ri_selector)
    medicine_radio.click()

    # 実行ボタンを押す
    calculate_btn = driver.find_element(By.CSS_SELECTOR, "button.btn-primary")
    calculate_btn.click()
    calc_start = time.time()

    # 計算結果の確認

    result_message = WebDriverWait(driver, 5).until(
        EC.visibility_of_element_located((By.ID, "resultMessage"))
    )
    calc_end = time.time()
    calc_time = round(calc_end - calc_start, 2)
    if (
        sheet.cells(calc_row, result_column).value is None
        or sheet.cells(calc_row, result_column).value == ""
    ):
        sheet.cells(calc_row, comment_column).value = (
            "計算時間：" + str(calc_time) + "秒"
        )
        if calc_time < 1:
            sheet.cells(calc_row, result_column).value = "OK"
        else:
            sheet.cells(calc_row, result_column).value = "NG"

    result_text = result_message.get_attribute("innerHTML")

    result_weight = float(
        re.search(r"体重：\s?(\d+\.\d+|\d+)\s?kg", result_text).group(1)
    )

    result_ri = re.search(r"薬剤名：([^<]+)", result_text).group(1)

    result_dose = float(
        re.search(r"投与量：\s?(\d+\.\d+|\d+)\s?MBq", result_text).group(1)
    )
    result_load_name = None
    result_load_amount = None
    if ri_id == "ecd" or ri_id == "imp":
        result_load_amount = float(
            re.search(r"負荷薬剤ダイアモックス（ACZ）：(\d+)\s?mg", result_text).group(
                1
            )
        )
        result_load_name = re.search(r"負荷薬剤([^：]+)", result_text).group(1)
        result_load_name = "負荷薬剤" + str(result_load_name)
    elif ri_id == "mag3":
        result_load_amount = float(
            re.search(
                r"負荷薬剤フロセミド（ラシックス）：(\d+\.\d)\s?ml", result_text
            ).group(1)
        )
        result_load_name = re.search(r"負荷薬剤([^：]+)", result_text).group(1)
        result_load_name = "負荷薬剤" + str(result_load_name)

    close_btn = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn-secondary"))
    )
    close_btn.click()

    if (
        sheet.cells(this_row, result_column).value is None
        or sheet.cells(this_row, result_column).value == ""
    ):
        if (
            result_weight == input_weight
            and result_ri == input_ri
            and result_dose == excel_calc_dose
        ):
            if result_load_name:

                if (
                    result_load_name == load_medicine
                    and result_load_amount == load_dose
                ):
                    sheet.cells(this_row, result_column).value = "OK"
                else:
                    sheet.cells(this_row, result_column).value = "NG"
            else:
                sheet.cells(this_row, result_column).value = "OK"
        else:
            sheet.cells(this_row, result_column).value = "NG"

    try:
        assert float(result_weight) == float(input_weight)

    except AssertionError as e:
        errors.append(
            f"体重が一致しません。計算結果：{result_weight} 期待値：{input_weight}"
        )

    try:
        assert str(result_ri) == input_ri

    except AssertionError as e:
        errors.append(
            f"薬剤名が一致しません。実行結果:"
            + str(result_ri)
            + "期待値:"
            + str(input_ri)
        )

    try:
        assert float(result_dose) == float(excel_calc_dose)

    except AssertionError as e:
        errors.append(
            f"投与量が一致しません。計算結果:{result_dose}MBq 期待値：{excel_calc_dose}MBq"
        )

    if ri_id == "ecd" or ri_id == "imp" or ri_id == "mag3":
        if ri_id == "mag3":
            unit = " ml"
        else:
            unit = " mg"
        try:
            assert str(result_load_name) == str(load_medicine)

        except AssertionError as e:
            errors.append(
                f"負荷薬剤名が一致しません。実行結果:{result_load_name} 期待値：{load_medicine}"
            )

        try:
            assert float(result_load_amount) == float(load_dose)

        except AssertionError as e:
            errors.append(
                f"負荷薬剤量が一致しません。計算結果:{result_load_amount}{unit} 期待値：{load_dose}{unit}"
            )

        result_load_name = ""
        result_load_amount = ""

    if errors:
        sheet.cells(this_row, comment_column).value = "\n".join(errors)
        pytest.fail("テストNo." + str(test_number) + "\n".join(errors))
