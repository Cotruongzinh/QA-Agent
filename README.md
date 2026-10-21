# QA-Agent

QA-Agent là một project thử nghiệm xây dựng **AI Agent hỗ trợ kiểm thử Web/PHP** bằng cách kết hợp **Agent Skills**, **Tencent BrowserSkill** và quy trình tạo **QA Report Excel** theo template.

Mục tiêu của project không phải thay thế QA Tester, mà là chuẩn hóa các bước lặp lại trong quá trình kiểm thử: đọc hoặc tạo testcase, thao tác trình duyệt, thu thập evidence, đánh giá kết quả và xuất báo cáo.

---

## 1. Kiến trúc tổng quát

```text
User / QA Tester
      |
      | prompt
      v
+-------------------+
|     qa-tester     |
|  QA orchestration |
+---------+---------+
          |
          | browser interaction
          v
+-------------------+
|   browser-skill   |
| Tencent BrowserSkill
+---------+---------+
          |
          | bsk CLI / daemon / extension
          v
      Chrome / Edge
          |
          | UI + Network + Console + Screenshot
          v
+-------------------+
|     qa-tester     |
| PASS / FAIL /     |
| UNVERIFIED        |
+---------+---------+
          |
          | qa-result.json
          v
+-------------------+
|   excel-report    |
| validate + report |
+---------+---------+
          |
          v
      QA Report.xlsx
```

Project tách riêng ba trách nhiệm chính:

| Thành phần | Vai trò |
|---|---|
| `qa-tester` | Điều phối nghiệp vụ QA, tạo/đọc testcase, thu thập evidence và quyết định verdict. |
| `browser-skill` | Cung cấp hướng dẫn sử dụng Tencent BrowserSkill để thao tác trình duyệt thật. |
| `excel-report` | Nhận dữ liệu QA đã chốt và ghi vào bản copy của Excel template. |

---

## 2. QA Tester Skill

Skill chính nằm tại:

```text
.agents/skills/qa-tester/
```

`qa-tester` có hai chế độ hoạt động.

### EXECUTE

Dùng khi đã có testcase.

Agent phải:

1. Giữ nguyên `Test Steps`, `Test Data` và `Expected Result`.
2. Thực thi testcase đúng thứ tự.
3. Thu thập evidence có thể quan sát được.
4. Ghi `Actual Result`.
5. Kết luận `PASS`, `FAIL` hoặc `UNVERIFIED`.

### DISCOVER

Dùng khi chưa có testcase.

Agent sẽ:

1. Quan sát trang hoặc workflow được giao.
2. Xác định các chức năng và trạng thái có thể kiểm thử.
3. Lập test plan theo rủi ro.
4. Tạo testcase phù hợp như happy path, negative test, boundary/edge case, validation, empty state và error state khi có thể xác minh.
5. Ưu tiên testcase read-only và reversible trước các thao tác làm thay đổi dữ liệu.
6. Thực thi và đánh giá bằng evidence.

### Quy tắc verdict

- **PASS**: mọi phần quan trọng của Expected Result đã được chứng minh bằng quan sát.
- **FAIL**: có ít nhất một quan sát mâu thuẫn với Expected Result.
- **UNVERIFIED**: không đủ điều kiện hoặc evidence để xác minh đầy đủ.

Lỗi của BrowserSkill hoặc element reference **không được tự động xem là lỗi website**. Nếu tool action thất bại, Agent phải re-observe, lấy fresh reference và chỉ retry action lỗi tối đa một lần trước khi chuyển testcase sang `UNVERIFIED` nếu vẫn không thể thực thi.

---

## 3. BrowserSkill

Project sử dụng **Tencent BrowserSkill** cho toàn bộ browser interaction.

Luồng kỹ thuật:

```text
Codex / Agent
    -> bsk CLI
    -> BrowserSkill daemon
    -> BrowserSkill extension
    -> Chrome / Edge
```

BrowserSkill được dùng để:

- observe giao diện;
- click / fill / press;
- borrow tab trình duyệt đang mở;
- chụp screenshot;
- kiểm tra Network / response / console khi cần;
- thu thập debug evidence.

`qa-tester` không sử dụng Playwright hoặc Selenium trong workflow hiện tại.

---

## 4. Excel Report Skill

Skill báo cáo nằm tại:

```text
.agents/skills/excel-report/
```

`excel-report` không tự quyết định testcase đúng hay sai. Nó chỉ xử lý dữ liệu đã được `qa-tester` xác nhận.

Luồng:

```text
qa-tester
  -> qa-result.json
  -> validate contract
  -> copy Excel template
  -> write Testcase / NoteBug / Summary
  -> insert evidence
  -> reopen and verify output
```

Template nguồn:

```text
templates/QA_Report_Template.xlsx
```

Template được xem là **read-only**. Script chỉ được ghi vào bản copy trong thư mục output.

---

## 5. Cấu trúc project

```text
QA-Agent/
├── .agents/
│   └── skills/
│       ├── browser-skill/
│       │   ├── SKILL.md
│       │   └── references/
│       │
│       ├── qa-tester/
│       │   ├── SKILL.md
│       │   ├── references/
│       │   │   ├── bsk-windows-cli.md
│       │   │   ├── messenger-runner.md
│       │   │   ├── methodology.md
│       │   │   └── report-handoff.md
│       │   └── scripts/
│       │       └── messenger_runner.py
│       │
│       └── excel-report/
│           ├── SKILL.md
│           ├── references/
│           │   ├── qa-result-schema.md
│           │   └── template-rules.md
│           └── scripts/
│               ├── inspect_template.py
│               ├── write_results.py
│               └── write_smoke_test.py
│
├── templates/
│   └── QA_Report_Template.xlsx
│
├── build_fresh_report.py
├── .gitignore
└── README.md
```

---

## 6. Yêu cầu môi trường

- Windows
- Python 3.x
- Google Chrome hoặc Microsoft Edge
- Tencent BrowserSkill CLI (`bsk`)
- BrowserSkill extension
- Codex / Agent environment có hỗ trợ Agent Skills

Python packages dùng cho phần Excel:

```powershell
python -m pip install openpyxl pillow
```

Khuyến nghị tạo virtual environment:

```powershell
cd C:\QA-Agent
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install openpyxl pillow
```

---

## 7. Cài BrowserSkill

Trên Windows PowerShell:

```powershell
irm https://raw.githubusercontent.com/Tencent/BrowserSkill/main/install.ps1 | iex
```

Kiểm tra CLI:

```powershell
bsk --version
bsk status --json
```

Cài skill BrowserSkill cho Agent nếu môi trường chưa có:

```powershell
bsk install-skill
```

Xem các target cài đặt được hỗ trợ:

```powershell
bsk install-skill --list
```

Ngoài CLI, cần cài và bật BrowserSkill extension trên Chrome/Edge.

---

## 8. Ví dụ sử dụng

### 8.1 DISCOVER — chưa có testcase

```text
$qa-tester

Test module <tên module> trên tab Chrome đang mở.
Không có testcase sẵn: DISCOVER, tự tạo testcase phù hợp rồi thực thi.

Dùng browser-skill và không thao tác ngoài phạm vi module.
Nếu cần thì kiểm tra Network/Fetch-XHR để xác minh.

Sau khi xong, tạo report bằng excel-report từ template.
Không sửa template gốc.
```

### 8.2 EXECUTE — đã có testcase

```text
$qa-tester

Chạy các testcase TC_001 đến TC_005 trên tab Chrome đang mở.
Giữ nguyên Test Steps, Test Data và Expected Result.

Dùng browser-skill cho toàn bộ thao tác trình duyệt.
Sau khi hoàn tất, lưu qa-result.json và tạo report bằng excel-report.
Không sửa template gốc.
```

---

## 9. QA result handoff

Sau một QA run, `qa-tester` có thể lưu kết quả vào:

```text
runs/<run_id>/qa-result.json
```

File JSON là data contract giữa `qa-tester` và `excel-report`.

Các nhóm dữ liệu chính gồm:

```text
run
  - run_id
  - target
  - mode
  - environment

testcases[]
  - tc_id
  - module
  - test_title
  - test_steps
  - test_data
  - expected_result
  - actual_result
  - status
  - evidence

bugs[]
summary
```

Các verdict hợp lệ:

```text
PASS
FAIL
UNVERIFIED
```

Evidence phải tham chiếu đến file local và không nhúng base64 trực tiếp vào JSON.

---

## 10. Chạy Excel tools trực tiếp

### Inspect template ở chế độ read-only

```powershell
python .\.agents\skills\excel-report\scripts\inspect_template.py `
  .\templates\QA_Report_Template.xlsx
```

Script chỉ đọc và in cấu trúc workbook, không save template.

### Ghi report từ `qa-result.json`

```powershell
python .\.agents\skills\excel-report\scripts\write_results.py `
  --input .\runs\<run_id>\qa-result.json `
  --template .\templates\QA_Report_Template.xlsx `
  --output .\output\QA_Report.xlsx
```

### Tạo fresh report từ template design

```powershell
python .\build_fresh_report.py `
  --input .\runs\<run_id>\qa-result.json `
  --template .\templates\QA_Report_Template.xlsx `
  --output .\output\QA_Report_Fresh.xlsx
```

`build_fresh_report.py` giữ cấu trúc/format của template nhưng loại dữ liệu lịch sử trên **bản copy output** trước khi ghi kết quả run mới.

---

## 11. Messenger runner

Project có thêm một runner chuyên dụng cho testcase Messenger:

```text
.agents/skills/qa-tester/scripts/messenger_runner.py
```

Runner này chuẩn hóa chuỗi:

```text
observe -> focus -> fill -> verify draft -> press Enter -> observe -> screenshot
```

Xem trợ giúp:

```powershell
python .\.agents\skills\qa-tester\scripts\messenger_runner.py --help
```

Chạy plan:

```powershell
python .\.agents\skills\qa-tester\scripts\messenger_runner.py run-case `
  --session <SESSION_ID> `
  --tab-id <TAB_ID> `
  --plan <PLAN_JSON> `
  --evidence-dir <EVIDENCE_DIR>
```

Runner Messenger là công cụ bổ sung; nó không bắt buộc cho các module Web/PHP thông thường.

---

## 12. Safety và nguyên tắc kiểm thử

Project đặt một số guardrail quan trọng:

- Không tự bịa credentials, quyền truy cập hoặc business rules.
- Không tự sửa testcase chính thức trong chế độ EXECUTE.
- Không xem lỗi automation là lỗi của website.
- Không kết luận PASS chỉ vì "không thấy lỗi".
- Không âm thầm retry testcase thất bại để che lỗi intermittent.
- Ưu tiên read-only và reversible testcase.
- Không thực hiện thao tác destructive hoặc có tác động thực tế ngoài phạm vi được phép.
- Không sửa trực tiếp Excel template gốc.
- Không commit run artifacts, debug logs hoặc dữ liệu nhạy cảm lên repository public.

---

## 13. Các thư mục không nên commit

`.gitignore` nên loại các artifact runtime như:

```gitignore
.venv/
.bsk-runtime/
runs/
output/
__pycache__/
*.pyc
*.log
*.lock
.env
TC-*.png
```

Nếu debug artifact từng chứa credential, token hoặc thông tin nhạy cảm, cần thu hồi credential đó thay vì chỉ xóa file khỏi Git history.

---

## 14. Phạm vi hiện tại

Project hiện tập trung vào:

- Web functional testing;
- exploratory QA bằng DISCOVER;
- testcase execution bằng EXECUTE;
- browser automation thông qua BrowserSkill;
- UI / Network / Console evidence;
- QA result handoff bằng JSON;
- tạo báo cáo Excel theo template.

Project chưa nhằm thay thế hoàn toàn QA Tester và vẫn cần human review đối với nghiệp vụ, Expected Result mơ hồ, severity và quyết định cuối cùng trước khi gửi báo cáo chính thức.

---

## 15. Repository

```text
https://github.com/Cotruongzinh/QA-Agent
```

---

## Author

**Cổ Trường Vinh**  
Software Engineering Internship Project
