# IB Data Tool — Headless Aki Skill

Run the IB Data Batch Sending Tool via conversation (no GUI).

## When to Use

- User says "跑IB Data"、"发IB数据"、"IB batch" or similar

## Prerequisites

- Project at `C:\Users\{USERNAME}\Desktop\IB_Data_Tool\` on branch `headless`
- Python 3.13: `C:\Program Files\Python313\python.exe`
- Outlook running with target mailbox accessible
- Firefox logged into `trans-logistics-cn.amazon.com` (OC Cookie)
- `Mars_LCL_Package\BookingFilePack\BCFile\PendingList\` on Desktop

## Workflow

### Step 0: Check & Clean Output

First check if Output has leftover files:

```bash
powershell.exe -Command "cd $env:USERPROFILE\Desktop\IB_Data_Tool; & 'C:\Program Files\Python313\python.exe' src/main.py --check-output"
```

Output:
- `[OUTPUT_EMPTY]` → silently proceed to Step 1
- `[OUTPUT_FILES]...[/OUTPUT_FILES]` → ask user: "Output 目录有以下 {n} 个旧文件：{file list}，是否清理？"

If user confirms, run cleanup:
```bash
powershell.exe -Command "cd $env:USERPROFILE\Desktop\IB_Data_Tool; & 'C:\Program Files\Python313\python.exe' src/main.py --cleanup"
```

If user declines, proceed without cleaning (files will remain).

### Step 1: Select Outlook Store

```bash
powershell.exe -Command "cd $env:USERPROFILE\Desktop\IB_Data_Tool; & 'C:\Program Files\Python313\python.exe' src/main.py --list-stores"
```

Output format:
```
[STORES]
  1. user@amazon.com
  2. shared-mailbox@amazon.com
[/STORES]
```

Present the list and ask user:
"请选择搜索邮箱（输入编号，多选用逗号分隔，如 `1,2`）："

User replies with number(s). Map numbers back to store display names, join with comma.

### Step 2: Set Store

Supports multi-select (comma-separated):

```bash
powershell.exe -Command "cd $env:USERPROFILE\Desktop\IB_Data_Tool; & 'C:\Program Files\Python313\python.exe' src/main.py --set-store '<name1>,<name2>'"
```

### Step 3: Process Batches (loop)

Run single batch:
```bash
powershell.exe -Command "cd $env:USERPROFILE\Desktop\IB_Data_Tool; & 'C:\Program Files\Python313\python.exe' src/main.py --headless --single-batch"
```

Parse the `[BATCH_SUMMARY]` output block:
```
[BATCH_SUMMARY]
  processed: 10
  success: 8
  odm: 1
  skipped: 1
  cda: 0
  remaining: 15
  success_list:
    - AL0-xxx
  ...
[/BATCH_SUMMARY]
```

Present results to user in a readable table, then ask:
- **remaining > 0**: "还有 {remaining} 票待处理，继续下一批？"
- **remaining == 0**: "全部处理完毕！"

If user says continue → run again. If stop/finish → go to Step 4.

If output contains "无待处理 AL0，流程结束" → all done, go to Step 4.

### Step 4: Finish

Tell user:
- Event CSV location: `Output/event_list_{date}.csv`
- Remind to upload at OC Pending Tasks page
- Link: https://trans-logistics-cn.amazon.com/aglt/appViews/app#/pending-tasks

## Notes

- `--check-output` lists Output files (no delete); `--cleanup` executes deletion
- `--single-batch` processes exactly one batch then exits
- Event CSV uses append mode — accumulates across batches
- `[PROGRESS] idx/total AL0` lines appear during processing
- If "OC Cookie 已过期" appears → ask user to login Firefox first
- Timeout: 60s per batch should be sufficient for batch_size=10
- OC Cookie expires every Monday — login Firefox before first run of the week
