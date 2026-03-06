
# 開發者手冊

## 1. 項目概述

本項目是一個基於 Flask 和 openpyxl的程式。系統包含兩個主要的申請表單（Form1 和 Form2），並將收集到的數據整理到一個5個工作表的 Excel 文件 `HY-applicants.xlsx` 中。

為了方便非技術人員本地測試，整個應用程序已被打包成單一的可執行文件 `v1.exe`。

### 1.1. 主要技術應用

- **後端**: Python 3, Flask
- **數據存儲**: `openpyxl` 庫，用於讀寫 `.xlsx` 文件
- **打包工具**: PyInstaller
- **前端**: HTM（表單架構）, CSS（界面設計）

### 1.2. 文件結構

```
re
├── HY-applicantse-exe/  (打包後的可執行文件目錄)
│   └── v1.exe
├── static/
│   └── css/
│       ├── final.css(鏈接form1.html和form2.html的css 文件)
│       ├── form1.css(之前鏈接form1.html的css文件，只做參考)
│       ├── form2.css(之前鏈接form1.htm2的css文件，只做參考)
│       └── main.css(鏈接index.html和success.html的css 文件)
├── templates/
│   ├── index.html       (主頁)
│   ├── form1.html       (前線工作人員申請表單)
│   ├── form2.html       (普通工作人員申請表單)
│   └── success.html     (提交成功頁面-現在只鏈接form1.html)
├── v1.py                (應用主程序)
├── requirements.txt     (Python 依賴列表)
└── HY-applicants.xlsx   (數據存儲文件-Excel文件)
```

---

## 2. 後端詳解 (`v1.py`)

`v1.py` 是應用的核心，負責處理前端收到的資料和自動存放資料到對應的列表。

### 2.1. 主要依賴庫

- `flask`: Web 框架。
- `openpyxl`: 用於操作 Excel 文件。
- `datetime`: 用於生成帶時間的唯一員工 ID。
- `os`, `sys`: 用於處理文件路徑
- `random`: 負責生成隨機Employee ID
- `filelock`: 用於防止多個請求同時寫入 Excel 文件，解決並發衝突。

### 2.2. 關鍵函數

#### `resource_path(relative_path)` 處

- **目的**: 解決 PyInstaller 打包後，程序在臨時目錄 (`_MEIPASS`) 中運行時無法找到 `static` 和 `templates` 等資源文件的問題。
- **機制**: 判斷程序是否在打包環境下運行。如果是，則返回指向臨時目錄中資源的絕對路徑；否則，返回正常的相對路徑。
- **應用**: 在創建 Flask 實例時，`template_folder` 和 `static_folder` 都通過此函數進行路徑轉換。

```python
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# ...

app = Flask(__name__,
            static_folder=resource_path('static'),
            template_folder=resource_path('templates'))
```

#### `get_next_employee_id(formtype)`

- **目的**: 為每個提交的申請生成一個唯一的員工 ID。
- **機制**: ID 由 `YYYYMMDDHHMMSS`+3個隨機數字 組成。例如：2025年10月24日 10時45分47秒提交會生成`20251024104547990`。其中`990`是隨機數字
- **限制**: 雖然隨機數降低了碰撞概率，但在高並發場景下理論上仍有極小可能重複。
```python
def get_next_employee_id():
    now = datetime.now() 
    timestamp_part = now.strftime("%Y%m%d%H%M%S")# format: YYYYMMDDHHMMSS
    random_part = ''.join(random.choices(string.digits, k=3))# add three ramdom number at the end
    employee_id = f"{timestamp_part}{random_part}"
    return employee_id
```
#### `fill_excel_with_form_data(form_data)`

這是整個應用的核心，負責數據處理表單收到的資料並自動寫入Excel。

- **目的**: 將前端表單提交的數據解析並寫入 `HY-applicants.xlsx` 的三個不同工作表中。
- **工作流程**:
    1.  **文件鎖**: 使用 `FileLock` 鎖定 `HY-applicants.xlsx.lock`，確保同一時間只有一個線程能寫入 Excel 文件，防止數據損壞。
    2.  **加載/創建 Excel**:
        - 如果 `HY-applicants.xlsx` 已存在，則加載它。
        - 如果不存在，則創建一個新的Excel
    3.  **獲取表單類型**: 通過 `form_data.get('formtype')` 判斷數據來源於 `form1` 還是 `form2`。
    4.  **數據寫入**:
        - **"basic info of applicant Sheet**: 寫入所有申請者的基本信息（員工ID、姓名、電話、申請職位等）。
        - **Education experience Sheet**: 寫入學歷並且自動隔一行，但是employee id和formtype會繼續寫入。允許前端自動添加。
        - **Professional Qualification**:寫入專業證書並寫入學歷並且自動隔一行，但是employee id和formtype會繼續寫入。允許前端自動添加。
        - **Working experience**: 寫入工作經驗，原理同上。
        - **Others information**: 寫入除基本資料的其他資料包括但不限於語言能力，軟件應用等。詳情請看**5.資料庫**
        - **Referees Information**: 寫入諮詢人資料，每個申請者只會最寫入兩個諮詢人，也就是兩行。
    5.  **保存文件**: 保存對 Excel 文件的修改。
    6.  **釋放鎖**: `with` 語句結束後，文件鎖自動釋放。
```python
def fill_excel_with_form_data(form_data):
    """fill data into Excel"""
    try:
        logger.info(f"Starting to fill Excel data... applicants: {form_data.get('chinese_name', 'unknown')}")
#...
next_row = ws_main.max_row + 1
        typeform = form_data.get('typeform', '')
        submission_time = datetime.now()
        
        
        # 基本資料
        main_data = [
            employee_id,  # Employee ID
            form_data.get('formtype', ''),  # 表單類型
            form_data.get('position', ''),  # 申請職位
        #...
        ]
 for col_idx, value in enumerate(main_data, start=1):
            safe_set_cell_value(ws_main, next_row, col_idx, value)
        
```
### 2.3. 路由 (Routes)

- `GET /`: 主頁。
- `GET /form1`, `GET /form2`: 顯示兩個不同的申請表單。
- `POST /submit`:
    - 接收來自 `form1` 和 `form2` 的 `POST` 請求。
- `GET /success`: 顯示提交成功頁面。
- `GET /test_excel`: **開發者專用**的調試端點。用於在不經過前端表單的情況下，直接用預設的數據測試，目的只是為了測試 `fill_excel_with_form_data()` function

---

## 3. 前端詳解

### 3.1. 模板引擎

- 使用 Flask 自帶的 Jinja2 模板引擎。
- **`url_for()`**: 所有指向 CSS 文件和內部頁面的鏈接都使用 `url_for()` 生成，以確保無論應用部署在何處，URL 都能保持正確。
    - **CSS**: `<link rel="stylesheet" href="{{ url_for('static', filename='css/form1.css') }}">`
    - **頁面跳轉**: `<a href="{{ url_for('form1') }}">`

### 3.2. 表單數據結構

- **`formtype`**: 每個表單 (`form1.html`, `form2.html`) 都包含一個隱藏字段 `<input type="hidden" name="formtype" value="form1">`。這是後端區分數據來源的關鍵。
### 4. 動態字段
- 教育、工作經歷等部分使用 JavaScript 實現動態添加/刪除行。
- 為了讓後端能正確接收，所有動態生成的 `input` 字段都使用相同的 `name` 屬性。Flask 會自動將它們收集為一個列表。


---

## 5. 數據庫 (`HY-applicants.xlsx`)


### 5.1. 工作表結構

**i. basic info of applicant Sheet**: 所有申請者的基本資料
- Employee ID
- Type of form(frontline/general)
- Position
- Salary
- Available Date
- Title
- Chinese Name
- English Name
- ID/Passport Number
- Marital Status
- Birth Date
- Birth Place
- Arrival Date
- Nationality
- Race
- Residential Phone
- Mobile Phone
- Email Address
- Address
- Correspondence Address
- Passport Issued Place
- Passport Issued Date
- Need Visa?
- criminal offence
- work in HY
- relatives?
- Submission Time

**ii. Education experience Sheet**: 學歷
- Employee ID
- type of form(general/frontline)
- Start Date
- Finish Date
- School/Institution
- Qualification/Certificate obtained

**iii. Professional Qualification**:專業證書
- Employee ID
- type of form(general/frontline)
- Date Obtained
- Issuing Authority
- Qualification
- License

**iv. Working experience**: 工作經驗
- Employee ID
- formtype
- Start Date
- Finish Date
- Name of Employer
- Previous Position
- Previous Salary
- Reason of Leaving
- Nature of Business

**v.  Others information**: 除基本資料外其他資料包括但不限於語言能力，軟件應用等
- Employee ID
- formtype
- Written in Chinese
- Spoken in Chinese
- Written in English
- Spoken in English
- Others language
- Written in others
- Spoken in others
- Software Applications
- Programming Languages
- Others Skills
- Company worked before
- Department
- Position
- Duration
- Name of relative working in company
- Relationship with relative
- Relative's company
- Relative's department
- Relative's position
- criminal offence date
- Criminal offence place
- learn position
- voluntary job-date
- voluntary job union
- voluntary job position

**vi. Referees Information**: 諮詢人資料
- Employee ID
- formtype
- Referee Name
- Phone Number
- Company
- Relationship

### 6. 注意事項

- **並發問題**: 儘管使用了 `filelock`，但它只對 `v1.py` 程序內的並發請求有效。如果用戶在程序運行的同時手動打開 `HY-applicants.xlsx` 文件，**將會導致程序寫入失敗並可能崩潰**

- **前線工作人員申請表修改**:`form1.html`的資料尚未完成對接，因此需要進行更新
- **圖片處理**:關於form的圖片上載，尚未進行處理。
