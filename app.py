#!/usr/bin/env python3
"""
System Task Dashboard - Windows v17

修正:
1. 表格分頁顯示（每頁50筆）
2. 超期圖表邏輯修正（只顯示超期 vs 未超期）
3. 下拉選單移到標題同一行
4. 移除表格欄位凍結
"""

import os
import re
import io
import sys
import json
import tempfile
from datetime import datetime, timedelta
from collections import defaultdict
from typing import List, Dict, Optional, Set
from dataclasses import dataclass, field

IS_WINDOWS = sys.platform == 'win32'

if IS_WINDOWS:
    import win32com.client
    import pythoncom

from flask import Flask, render_template_string, request, jsonify, send_file, Response

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except:
    HAS_OPENPYXL = False

try:
    import extract_msg
    HAS_EXTRACT_MSG = True
except:
    extract_msg = None
    HAS_EXTRACT_MSG = False

app = Flask(__name__)
app.secret_key = 'realtek-2025-v17'

FOLDERS = []
FOLDER_TREE = []
OUTLOOK_OK = False
LAST_RESULT = None
LAST_DATA = None

PRIORITY_WEIGHTS = {'high': 3, 'medium': 2, 'normal': 1}

# 儲存 mail 內容的全域字典
MAIL_CONTENTS = {}

@dataclass
class Task:
    title: str
    owners: List[str]
    priority: str = "normal"
    due_date: Optional[str] = None
    status: Optional[str] = None
    mail_date: str = ""
    mail_subject: str = ""
    module: str = ""  # 大模組 [公版], [DIAS] 等
    mail_id: str = ""  # mail 的唯一識別碼

@dataclass
class TaskTracker:
    title: str
    owners: List[str]
    priority: str = "normal"
    due_date: Optional[str] = None
    status: Optional[str] = None
    first_seen: str = ""
    last_seen: str = ""
    appearances: List[str] = field(default_factory=list)
    in_last_mail: bool = False
    
    def days_spent(self) -> int:
        if not self.first_seen or not self.last_seen:
            return 0
        try:
            first = datetime.strptime(self.first_seen, "%Y-%m-%d")
            last = datetime.strptime(self.last_seen, "%Y-%m-%d")
            return (last - first).days + 1
        except:
            return 0
    
    def is_overdue(self) -> bool:
        if not self.due_date:
            return False
        try:
            today = datetime.now()
            parts = self.due_date.split('/')
            if len(parts) == 2:
                month, day = int(parts[0]), int(parts[1])
                year = today.year
                due = datetime(year, month, day)
                if (today - due).days > 180:
                    due = datetime(year + 1, month, day)
                return today > due
        except:
            pass
        return False
    
    def get_task_status(self) -> str:
        if not self.in_last_mail:
            return "completed"
        elif self.status and 'pending' in self.status.lower():
            return "pending"
        else:
            return "in_progress"

def load_folders():
    global FOLDERS, FOLDER_TREE, OUTLOOK_OK
    
    print("[1] 初始化 COM...")
    pythoncom.CoInitialize()
    
    print("[2] 連接 Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    
    print("[3] 取得 MAPI Namespace...")
    namespace = outlook.GetNamespace("MAPI")
    
    print("[4] 列出資料夾...")
    
    folders = []
    tree = []
    
    for account in namespace.Folders:
        account_name = account.Name
        is_archive = '封存' in account_name or 'Archive' in account_name.lower()
        account_node = {"name": account_name, "children": [], "entry_id": "", "store_id": "", "is_archive": is_archive}
        
        try:
            store_id = account.StoreID
        except:
            store_id = ""
        
        try:
            for subfolder in account.Folders:
                sf_name = subfolder.Name
                try:
                    entry_id = subfolder.EntryID
                    sf_store_id = subfolder.StoreID
                except:
                    entry_id = ""
                    sf_store_id = store_id
                
                folders.append({"name": sf_name, "path": f"{account_name}/{sf_name}", "entry_id": entry_id, "store_id": sf_store_id, "is_archive": is_archive})
                subfolder_node = {"name": sf_name, "entry_id": entry_id, "store_id": sf_store_id, "children": [], "is_archive": is_archive}
                
                try:
                    for sub2 in subfolder.Folders:
                        s2_name = sub2.Name
                        try:
                            s2_entry = sub2.EntryID
                            s2_store = sub2.StoreID
                        except:
                            s2_entry = ""
                            s2_store = sf_store_id
                        
                        folders.append({"name": s2_name, "path": f"{account_name}/{sf_name}/{s2_name}", "entry_id": s2_entry, "store_id": s2_store, "is_archive": is_archive})
                        subfolder_node["children"].append({"name": s2_name, "entry_id": s2_entry, "store_id": s2_store, "children": [], "is_archive": is_archive})
                except:
                    pass
                
                account_node["children"].append(subfolder_node)
        except Exception as e:
            pass
        
        tree.append(account_node)
    
    FOLDERS = folders
    FOLDER_TREE = tree
    OUTLOOK_OK = True
    print(f"    ✅ 共載入 {len(folders)} 個資料夾")

def get_messages(entry_id, store_id, start_date, end_date, exclude_after_5pm: bool = True):
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    folder = namespace.GetFolderFromID(entry_id, store_id)
    items = folder.Items
    items.Sort("[ReceivedTime]", True)
    
    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    end_dt = datetime.strptime(end_date, "%Y-%m-%d") + timedelta(days=1)
    
    try:
        items = items.Restrict(f"[ReceivedTime] >= '{start_dt.strftime('%m/%d/%Y')}' AND [ReceivedTime] < '{end_dt.strftime('%m/%d/%Y')}'")
    except:
        pass
    
    messages = []
    for item in items:
        try:
            rt = item.ReceivedTime
            if hasattr(rt, 'date') and not (start_dt.date() <= rt.date() < end_dt.date()):
                continue
            
            # 檢查是否是下午 5:00 後的 mail
            if exclude_after_5pm and hasattr(rt, 'hour'):
                if rt.hour >= 17:  # 17:00 = 下午 5:00
                    continue
            
            # 取得 HTML 內容
            html_body = ""
            try:
                html_body = item.HTMLBody or ""
            except:
                pass
            
            messages.append({
                "subject": item.Subject or "", 
                "body": item.Body or "",
                "html_body": html_body,
                "date": rt.strftime("%Y-%m-%d") if hasattr(rt, 'strftime') else "",
                "time": rt.strftime("%H:%M") if hasattr(rt, 'strftime') else ""
            })
        except:
            continue
    
    return messages

class TaskParser:
    def __init__(self, exclude_middle_priority: bool = True):
        self.tasks: List[Task] = []
        self.current_module: str = ""  # 當前的大模組
        self.exclude_middle_priority = exclude_middle_priority
        self.stop_parsing = False  # 遇到 Middle priority 後停止解析
    
    def _is_valid_module(self, text: str) -> bool:
        """檢查是否是有效的大模組標題"""
        # 排除 [status: xxx], [Due: xxx], [Pending], [Resolved] 等
        invalid_patterns = [
            r'^\[status\s*:', r'^\[due\s*:', r'^\[duedate\s*:',
            r'^\[pending\]$', r'^\[resolved\]$', r'^\[done\]$',
            r'^\[completed\]$', r'^\[in\s*progress\]$'
        ]
        text_lower = text.lower()
        for pattern in invalid_patterns:
            if re.match(pattern, text_lower):
                return False
        return True
    
    def _is_middle_priority_marker(self, line: str) -> bool:
        """檢查是否是 Middle priority 標記"""
        line_lower = line.lower().strip()
        return 'middle priority' in line_lower or 'low priority' in line_lower
    
    def parse(self, subject: str, body: str, mail_date: str = "", mail_time: str = "", html_body: str = ""):
        # 生成 mail_id
        import hashlib
        mail_id = hashlib.md5(f"{mail_date}_{mail_time}_{subject}".encode()).hexdigest()[:12]
        
        # 儲存原始 mail 內容（包含 HTML）
        MAIL_CONTENTS[mail_id] = {
            "subject": subject,
            "body": body,
            "html_body": html_body,
            "date": mail_date,
            "time": mail_time
        }
        
        original_body = body
        if '<html' in body.lower() or '<' in body:
            body = re.sub(r'<style[^>]*>.*?</style>', '', body, flags=re.DOTALL | re.IGNORECASE)
            body = re.sub(r'<[^>]+>', '\n', body)
            body = re.sub(r'&nbsp;', ' ', body)
            body = re.sub(r'&[a-z]+;', ' ', body)
        
        self.current_module = ""  # 重置
        self.stop_parsing = False  # 重置
        
        for line in body.split('\n'):
            line = line.strip()
            
            # 檢查是否遇到 Middle priority 標記
            if self.exclude_middle_priority and self._is_middle_priority_marker(line):
                self.stop_parsing = True
                break  # 停止解析這封 mail 的後續內容
            
            # 檢查是否是大模組標題（如 [公版]、[DIAS][AN11 Mac8q 2816A 2GB AOSP] 等）
            # 必須是獨立一行，且不包含數字開頭的任務格式
            module_match = re.match(r'^(\[[^\]]+\](?:\[[^\]]+\])*)\s*$', line)
            if module_match:
                potential_module = module_match.group(1)
                # 檢查第一個 [...] 是否是有效的模組標題
                first_bracket = re.match(r'^(\[[^\]]+\])', potential_module)
                if first_bracket and self._is_valid_module(first_bracket.group(1)):
                    self.current_module = potential_module
                continue
            
            # 解析任務
            match = re.match(r'^(\d+)[.\)、]\s*(.+)$', line)
            if match:
                content = match.group(2).strip()
                task = self._parse_task(content, mail_date, subject)
                if task:
                    task.module = self.current_module
                    task.mail_id = mail_id
                    self.tasks.append(task)
    
    def _parse_task(self, content: str, mail_date: str, mail_subject: str) -> Optional[Task]:
        # 必須有 Due date 才算任務（支援多種格式）
        # [Due date: 1126], [Due: 1126], [Duedate: 1126], [Due date: 11/26]
        due_match = re.search(r'\[\s*[Dd]ue\s*(?:date)?\s*[:\s]\s*(\d{2,4}[/]?\d{0,2})\s*\]', content, re.IGNORECASE)
        if not due_match:
            return None
        
        due_date = due_match.group(1)
        if '/' not in due_date and len(due_date) >= 3:
            if len(due_date) == 3:
                due_date = due_date[0] + '/' + due_date[1:]
            elif len(due_date) == 4:
                due_date = due_date[:2] + '/' + due_date[2:]
        
        content_without_due = content[:due_match.start()] + content[due_match.end():]
        content_without_due = content_without_due.strip()
        
        priority = "normal"
        star_match = re.match(r'^(\*{1,3})\s*', content_without_due)
        if star_match:
            stars = len(star_match.group(1))
            priority = "high" if stars >= 3 else ("medium" if stars == 2 else "normal")
            content_without_due = content_without_due[star_match.end():].strip()
        
        status = None
        status_match = re.search(r'\[Status[:\s]*([^\]]+)\]', content_without_due, re.IGNORECASE)
        if status_match:
            status = status_match.group(1).strip()
            content_without_due = content_without_due[:status_match.start()] + content_without_due[status_match.end():]
            content_without_due = content_without_due.strip()
        
        parts = re.split(r'\s*–\s*', content_without_due, maxsplit=1)
        if len(parts) < 2:
            parts = re.split(r'\s+-\s+', content_without_due, maxsplit=1)
        
        if len(parts) < 2:
            return None
        
        task_name = parts[0].strip()
        members_str = parts[1].strip()
        members_str = re.sub(r'\[.*?\]', '', members_str).strip()
        
        owners = self._parse_members(members_str)
        if not owners:
            return None
        
        task_name = re.sub(r'\s+', ' ', task_name).strip()
        task_name = re.sub(r'\[.*?\]', '', task_name).strip()
        
        if not task_name or len(task_name) < 2:
            return None
        
        return Task(title=task_name, owners=owners, priority=priority, due_date=due_date, status=status, mail_date=mail_date, mail_subject=mail_subject)
    
    def _parse_members(self, text: str) -> List[str]:
        if not text:
            return []
        parts = re.split(r'[/,、]', text)
        members = []
        for p in parts:
            p = p.strip()
            if not p:
                continue
            if re.match(r'^[\u4e00-\u9fff]{1,10}$', p):
                members.append(p)
            elif re.match(r'^[A-Za-z][A-Za-z0-9_]{0,19}$', p):
                members.append(p)
        return members

class Stats:
    def __init__(self):
        self.raw_tasks: List[Dict] = []  # 原始任務記錄
        self.unique_members: Set[str] = set()
        self.last_mail_date: str = ""
    
    def _task_key(self, title: str, due: str, owners: List[str]) -> str:
        """任務唯一識別：標題 + Due date + 負責人"""
        return f"{title.strip().lower()}|{due}|{','.join(sorted(owners))}"
    
    def add(self, task: Task):
        """先收集所有原始任務"""
        self.raw_tasks.append({
            "title": task.title,
            "owners": task.owners,
            "owners_str": "/".join(task.owners),
            "priority": task.priority,
            "due": task.due_date,
            "status": task.status or "-",
            "mail_date": task.mail_date,
            "mail_subject": task.mail_subject,
            "mail_id": task.mail_id,
            "module": task.module or "",
            "_key": self._task_key(task.title, task.due_date, task.owners)
        })
        
        for owner in task.owners:
            self.unique_members.add(owner)
        
        if task.mail_date > self.last_mail_date:
            self.last_mail_date = task.mail_date
    
    def _process_tasks(self) -> List[Dict]:
        """
        處理任務生命週期：
        1. 按日期排序所有 mail
        2. 同一天的 mail 合併（去重）
        3. 追蹤每個任務的出現與消失
        4. 計算超期天數：從第一次出現到被移除的那天 vs Due date
        """
        if not self.raw_tasks:
            return []
        
        # 按 mail_date 分組
        from collections import defaultdict
        tasks_by_date = defaultdict(list)
        for t in self.raw_tasks:
            tasks_by_date[t["mail_date"]].append(t)
        
        # 排序日期
        sorted_dates = sorted(tasks_by_date.keys())
        
        # 追蹤任務生命週期
        # key -> {"first_seen", "task_data", "active"}
        task_tracker = {}
        
        # 最終任務列表（每個任務實例）
        final_tasks = []
        
        # 上一個日期的任務 keys
        prev_date_keys = set()
        
        for date_idx, mail_date in enumerate(sorted_dates):
            # 同一天的任務去重（只保留一個）
            day_tasks = tasks_by_date[mail_date]
            day_task_map = {}
            for t in day_tasks:
                key = t["_key"]
                if key not in day_task_map:
                    day_task_map[key] = t
                else:
                    # 同一天重複的任務，保留 priority 較高的
                    existing = day_task_map[key]
                    priority_order = {"high": 3, "medium": 2, "normal": 1}
                    if priority_order.get(t["priority"], 0) > priority_order.get(existing["priority"], 0):
                        day_task_map[key] = t
            
            current_date_keys = set(day_task_map.keys())
            
            # 檢查哪些任務在這一天消失了（完成了）
            for key in prev_date_keys:
                if key not in current_date_keys and key in task_tracker and task_tracker[key]["active"]:
                    # 任務完成！計算超期
                    tracker = task_tracker[key]
                    task_data = tracker["task_data"].copy()
                    
                    # 計算超期天數：完成日期（上一個日期）vs Due date
                    # 實際完成日是上一個還有這個任務的日期
                    prev_date = sorted_dates[date_idx - 1] if date_idx > 0 else mail_date
                    task_data["first_seen"] = tracker["first_seen"]
                    task_data["last_seen"] = prev_date
                    task_data["completed_date"] = prev_date
                    task_data["task_status"] = "completed"
                    task_data["overdue_days"] = self._calc_overdue_days_v2(
                        task_data["due"], tracker["first_seen"], prev_date
                    )
                    task_data["days_spent"] = self._calc_days_between(tracker["first_seen"], prev_date)
                    
                    final_tasks.append(task_data)
                    task_tracker[key]["active"] = False
            
            # 處理這一天的任務
            for key, task_data in day_task_map.items():
                if key not in task_tracker or not task_tracker[key]["active"]:
                    # 新任務或重新出現的任務
                    task_tracker[key] = {
                        "first_seen": mail_date,
                        "task_data": task_data,
                        "active": True
                    }
                else:
                    # 任務繼續存在，更新最新資料
                    task_tracker[key]["task_data"] = task_data
            
            prev_date_keys = current_date_keys
        
        # 處理最後一天仍然存在的任務（進行中或 Pending）
        last_date = sorted_dates[-1] if sorted_dates else ""
        last_date_keys = prev_date_keys
        
        for key in last_date_keys:
            if key in task_tracker and task_tracker[key]["active"]:
                tracker = task_tracker[key]
                task_data = tracker["task_data"].copy()
                task_data["first_seen"] = tracker["first_seen"]
                task_data["last_seen"] = last_date
                
                # 判斷是 Pending 還是進行中
                if task_data["status"] and 'pending' in task_data["status"].lower():
                    task_data["task_status"] = "pending"
                else:
                    task_data["task_status"] = "in_progress"
                
                # 進行中/Pending 的超期計算：用今天 vs Due date
                task_data["overdue_days"] = self._calc_overdue_from_today(task_data["due"])
                task_data["days_spent"] = self._calc_days_between(tracker["first_seen"], last_date)
                
                final_tasks.append(task_data)
                task_tracker[key]["active"] = False
        
        return final_tasks
    
    def _calc_overdue_days_v2(self, due_date: str, first_seen: str, completed_date: str) -> int:
        """
        計算已完成任務的超期天數：
        超期天數 = 完成日期 - Due date（正數表示超期）
        """
        if not due_date or not completed_date:
            return 0
        try:
            parts = due_date.split('/')
            if len(parts) == 2:
                month, day = int(parts[0]), int(parts[1])
                completed_dt = datetime.strptime(completed_date, "%Y-%m-%d")
                first_dt = datetime.strptime(first_seen, "%Y-%m-%d")
                year = first_dt.year
                due_dt = datetime(year, month, day)
                
                # 如果 due date 比 first_seen 早超過 6 個月，可能是明年的
                if (first_dt - due_dt).days > 180:
                    due_dt = datetime(year + 1, month, day)
                
                diff = (completed_dt - due_dt).days
                return max(0, diff)
        except:
            pass
        return 0
    
    def _calc_overdue_from_today(self, due_date: str) -> int:
        """計算進行中任務的超期天數（相對於今天）"""
        if not due_date:
            return 0
        try:
            parts = due_date.split('/')
            if len(parts) == 2:
                month, day = int(parts[0]), int(parts[1])
                today = datetime.now()
                year = today.year
                due_dt = datetime(year, month, day)
                
                if (today - due_dt).days > 180:
                    due_dt = datetime(year + 1, month, day)
                elif (due_dt - today).days > 180:
                    due_dt = datetime(year - 1, month, day)
                
                diff = (today - due_dt).days
                return max(0, diff)
        except:
            pass
        return 0
    
    def _calc_days_between(self, start_date: str, end_date: str) -> int:
        """計算兩個日期之間的天數"""
        if not start_date or not end_date:
            return 0
        try:
            start = datetime.strptime(start_date, "%Y-%m-%d")
            end = datetime.strptime(end_date, "%Y-%m-%d")
            return (end - start).days + 1
        except:
            return 0
    
    def _is_overdue(self, due_date: str) -> bool:
        if not due_date:
            return False
        try:
            today = datetime.now()
            parts = due_date.split('/')
            if len(parts) == 2:
                month, day = int(parts[0]), int(parts[1])
                year = today.year
                due = datetime(year, month, day)
                if (today - due).days > 180:
                    due = datetime(year + 1, month, day)
                return today > due
        except:
            pass
        return False
    
    def summary(self):
        # 處理任務生命週期
        all_tasks = self._process_tasks()
        
        completed_count = 0
        pending_count = 0
        in_progress_count = 0
        overdue_count = 0
        not_overdue_count = 0
        
        member_stats = defaultdict(lambda: {
            "total": 0, "completed": 0, "pending": 0, "in_progress": 0,
            "high": 0, "medium": 0, "normal": 0, 
            "score": 0, "tasks": []
        })
        
        # 按 last_seen 降序排序
        sorted_tasks = sorted(all_tasks, key=lambda x: x.get("last_seen", ""), reverse=True)
        
        for task in sorted_tasks:
            task_status = task["task_status"]
            overdue_days = task.get("overdue_days", 0)
            is_overdue = overdue_days > 0
            task["is_overdue"] = is_overdue
            
            if task_status == "completed":
                completed_count += 1
                # 已完成任務也計入超期統計
                if is_overdue:
                    overdue_count += 1
                else:
                    not_overdue_count += 1
            elif task_status == "pending":
                pending_count += 1
                if is_overdue:
                    overdue_count += 1
                else:
                    not_overdue_count += 1
            else:
                in_progress_count += 1
                if is_overdue:
                    overdue_count += 1
                else:
                    not_overdue_count += 1
            
            for owner in task["owners"]:
                d = member_stats[owner]
                d["total"] += 1
                d[task_status] += 1
                d[task["priority"]] += 1
                
                if task_status == "completed":
                    d["score"] += PRIORITY_WEIGHTS.get(task["priority"], 1)
                
                d["tasks"].append(task)
        
        total_tasks = len(all_tasks)
        
        members = []
        for n, s in sorted(member_stats.items(), key=lambda x: -x[1]["total"]):
            members.append({
                "name": n, "total": s["total"], 
                "completed": s["completed"], "pending": s["pending"], "in_progress": s["in_progress"],
                "high": s["high"], "medium": s["medium"], "normal": s["normal"],
                "score": s["score"], "tasks": s["tasks"]
            })
        
        contribution = []
        overdue_by_member = {}  # 每個成員的超期統計
        
        for n, s in member_stats.items():
            # 計算非 Pending 的任務（已完成 + 進行中）
            non_pending_tasks = [t for t in s["tasks"] if t["task_status"] != "pending"]
            task_count = len(non_pending_tasks)
            
            # 權重分數：非 Pending 任務的優先級權重
            high_count = sum(1 for t in non_pending_tasks if t["priority"] == "high")
            med_count = sum(1 for t in non_pending_tasks if t["priority"] == "medium")
            nor_count = sum(1 for t in non_pending_tasks if t["priority"] == "normal")
            weighted_score = high_count * 3 + med_count * 2 + nor_count * 1
            
            # 計算超期統計
            overdue_tasks = [t for t in non_pending_tasks if t.get("overdue_days", 0) > 0]
            overdue_task_count = len(overdue_tasks)
            total_overdue_days = sum(t.get("overdue_days", 0) for t in overdue_tasks)
            avg_overdue_days = total_overdue_days / overdue_task_count if overdue_task_count > 0 else 0
            
            # 分別計算已完成和未完成的超期天數
            completed_overdue_tasks = [t for t in overdue_tasks if t.get("task_status") == "completed"]
            active_overdue_tasks = [t for t in overdue_tasks if t.get("task_status") != "completed"]
            completed_overdue_days = sum(t.get("overdue_days", 0) for t in completed_overdue_tasks)
            active_overdue_days = sum(t.get("overdue_days", 0) for t in active_overdue_tasks)
            
            # 超期減分公式：
            # - 每個超期任務扣 0.5 分
            # - 平均超期天數 > 7 天，額外扣 (平均天數 / 7) 分
            # - 超期率 > 30%，額外扣 2 分
            overdue_penalty = 0
            if task_count > 0:
                overdue_rate = overdue_task_count / task_count
                overdue_penalty += overdue_task_count * 0.5  # 每個超期任務扣 0.5 分
                if avg_overdue_days > 7:
                    overdue_penalty += avg_overdue_days / 7  # 平均超期天數懲罰
                if overdue_rate > 0.3:
                    overdue_penalty += 2  # 超期率高額外扣分
            
            final_score = max(0, weighted_score - overdue_penalty)
            
            overdue_by_member[n] = {
                "overdue_count": overdue_task_count,
                "total_overdue_days": total_overdue_days,
                "avg_overdue_days": round(avg_overdue_days, 1),
                "overdue_rate": round(overdue_task_count / task_count * 100, 1) if task_count > 0 else 0
            }
            
            contribution.append({
                "name": n,
                "task_count": task_count,
                "high": high_count, "medium": med_count, "normal": nor_count,
                "base_score": weighted_score,
                "overdue_count": overdue_task_count,
                "overdue_days": total_overdue_days,
                "completed_overdue_days": completed_overdue_days,  # 已完成超期天數
                "active_overdue_days": active_overdue_days,        # 未完成超期天數
                "overdue_penalty": round(overdue_penalty, 1),
                "score": round(final_score, 1)
            })
        
        contribution.sort(key=lambda x: -x["score"])
        for i, c in enumerate(contribution):
            c["rank"] = i + 1
        
        priority_counts = {"high": 0, "medium": 0, "normal": 0}
        for task in all_tasks:
            priority_counts[task["priority"]] += 1
        
        # 計算模組統計
        module_stats = defaultdict(int)
        for task in all_tasks:
            module = task.get("module", "") or "未分類"
            module_stats[module] += 1
        
        return {
            "total_tasks": total_tasks, 
            "total_members": len(self.unique_members),
            "completed_count": completed_count,
            "pending_count": pending_count,
            "in_progress_count": in_progress_count,
            "overdue_count": overdue_count,
            "not_overdue_count": not_overdue_count,
            "last_mail_date": self.last_mail_date,
            "priority_counts": priority_counts,
            "module_stats": dict(module_stats),
            "overdue_by_member": overdue_by_member,
            "members": members, 
            "all_tasks": sorted_tasks,
            "member_list": sorted(list(self.unique_members)), 
            "contribution": contribution
        }
    
    def excel(self):
        wb = Workbook()
        hfill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
        hfont = Font(bold=True, color="FFFFFF")
        redfont = Font(color="FF0000", bold=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        summary = self.summary()
        
        ws = wb.active
        ws.title = "成員統計"
        headers = ["成員", "總數", "已完成", "Pending", "進行中", "High", "Medium", "Normal", "貢獻分數"]
        for i, h in enumerate(headers, 1):
            c = ws.cell(1, i, h); c.fill, c.font, c.border = hfill, hfont, border
        for r, m in enumerate(summary["members"], 2):
            for i, v in enumerate([m["name"], m["total"], m["completed"], m["pending"], m["in_progress"], m["high"], m["medium"], m["normal"], m["score"]], 1):
                ws.cell(r, i, v).border = border
        
        ws2 = wb.create_sheet("任務明細")
        headers2 = ["模組", "任務", "負責人", "優先級", "Due Date", "超期天數", "狀態", "任務狀態", "首次出現", "最後出現", "花費天數"]
        for i, h in enumerate(headers2, 1):
            c = ws2.cell(1, i, h); c.fill, c.font, c.border = hfill, hfont, border
        status_map = {"completed": "已完成", "pending": "Pending", "in_progress": "進行中"}
        for r, t in enumerate(summary["all_tasks"], 2):
            overdue_days = t.get("overdue_days", 0)
            values = [
                t.get("module", "") or "", 
                t["title"], 
                t["owners_str"], 
                t["priority"], 
                t["due"] or "", 
                overdue_days if overdue_days > 0 else "",
                t["status"], 
                status_map.get(t["task_status"], ""), 
                t.get("first_seen", "") or "", 
                t.get("last_seen", "") or "", 
                t.get("days_spent", 0)
            ]
            for i, v in enumerate(values, 1):
                cell = ws2.cell(r, i, v)
                cell.border = border
                if i == 5 and overdue_days > 0:  # Due Date 欄位
                    cell.font = redfont
                if i == 6 and overdue_days > 0:  # 超期天數欄位
                    cell.font = redfont
        
        ws3 = wb.create_sheet("貢獻度排名")
        headers3 = ["排名", "成員", "任務數", "基礎分", "超期任務數", "總超期天數", "扣分", "總分"]
        for i, h in enumerate(headers3, 1):
            c = ws3.cell(1, i, h); c.fill, c.font, c.border = hfill, hfont, border
        for r, c in enumerate(summary["contribution"], 2):
            for i, v in enumerate([c["rank"], c["name"], c["task_count"], c["base_score"], c["overdue_count"], c["overdue_days"], c["overdue_penalty"], c["score"]], 1):
                cell = ws3.cell(r, i, v)
                cell.border = border
                if i in [5, 6, 7] and v > 0:  # 超期相關欄位標紅
                    cell.font = redfont
        
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

HTML = '''
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>System Task Dashboard</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root { --primary: #2E75B6; --primary-dark: #1a4f7a; }
        body { background: #f5f7fa; font-size: 14px; }
        .navbar { background: #2E75B6; }
        .card { border: none; border-radius: 10px; box-shadow: 0 2px 12px rgba(0,0,0,0.08); margin-bottom: 12px; }
        .card-header { background: #2E75B6; color: white; border-radius: 10px 10px 0 0 !important; padding: 8px 12px; display: flex; justify-content: space-between; align-items: center; }
        .card-header-title { font-weight: 500; }
        .stat-card { text-align: center; padding: 10px; cursor: pointer; transition: all 0.2s; height: 85px; display: flex; flex-direction: column; justify-content: center; }
        .stat-card:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(0,0,0,0.15); }
        .stat-number { font-size: 1.5rem; font-weight: bold; color: var(--primary); }
        .stat-number.danger { color: #dc3545; }
        .stat-number.warning { color: #FFA500; }
        .stat-number.success { color: #28a745; }
        .stat-number.info { color: #17a2b8; }
        .stat-label { color: #666; font-size: 0.7rem; }
        .badge-high { background: #FF6B6B !important; }
        .badge-medium { background: #FFE066 !important; color: #333 !important; }
        .badge-normal { background: #74C0FC !important; }
        .badge-completed { background: #28a745 !important; }
        .badge-pending { background: #FFA500 !important; }
        .badge-in_progress { background: #17a2b8 !important; }
        .config-ok { background: #d4edda; color: #155724; padding: 6px 12px; border-radius: 6px; margin-bottom: 8px; font-size: 0.8rem; }
        .drop-zone { border: 2px dashed #dee2e6; border-radius: 6px; padding: 15px; text-align: center; cursor: pointer; }
        .drop-zone.dragover { border-color: var(--primary); background: #f0f7ff; }
        .loading { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(255,255,255,0.9); display: flex; justify-content: center; align-items: center; z-index: 9999; }
        .tree-box { max-height: 150px; overflow-y: auto; border: 1px solid #dee2e6; border-radius: 6px; padding: 5px; background: #fff; font-size: 0.75rem; }
        .tree ul { list-style: none; padding-left: 12px; margin: 0; }
        .tree > ul { padding-left: 0; }
        .tree-toggle { cursor: pointer; color: #666; }
        .tree-toggle::before { content: "▶ "; font-size: 7px; }
        .tree-toggle.open::before { content: "▼ "; }
        .tree-item { cursor: pointer; padding: 1px 4px; border-radius: 3px; }
        .tree-item:hover { background: #e8f4fc; }
        .tree-item.selected { background: var(--primary); color: white; }
        .tree-item::before { content: "📁 "; }
        
        .data-table { width: 100%; font-size: 0.8rem; border-collapse: collapse; table-layout: auto; }
        .data-table thead th { background: #4a4a4a !important; color: white !important; font-weight: 600; cursor: pointer; padding: 8px 5px; white-space: nowrap; border: 1px solid #666; resize: horizontal; overflow: auto; min-width: 60px; }
        .data-table thead th:hover { background: #333 !important; }
        .data-table tbody td { padding: 6px 5px; vertical-align: middle; border: 1px solid #ddd; }
        .data-table tbody tr { cursor: pointer; }
        .data-table tbody tr:nth-child(even) { background: #f9f9f9; }
        .data-table tbody tr:hover { background: #e8f4fc !important; }
        .data-table tbody tr.row-pending { background: #fff8e1; }
        .data-table tbody tr.row-in_progress { background: #e3f2fd; }
        .data-table tbody tr.row-overdue { background: #ffebee; }
        .table-toolbar { display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #f8f9fa; border-bottom: 1px solid #dee2e6; }
        .table-toolbar input { max-width: 180px; font-size: 0.75rem; }
        .table-container { overflow-x: auto; }
        .text-overdue { color: #dc3545 !important; font-weight: bold; }
        
        .pagination-controls { display: flex; justify-content: space-between; align-items: center; padding: 8px 10px; background: #f8f9fa; border-top: 1px solid #dee2e6; font-size: 0.75rem; }
        .pagination-controls button { padding: 3px 10px; font-size: 0.75rem; }
        .pagination-controls select { font-size: 0.75rem; padding: 2px 5px; width: 70px; }
        
        .footer { text-align: center; padding: 12px; color: #999; font-size: 0.7rem; border-top: 1px solid #eee; margin-top: 10px; }
        .rank-badge { display: inline-block; width: 22px; height: 22px; line-height: 22px; border-radius: 50%; text-align: center; font-weight: bold; color: white; font-size: 0.7rem; }
        .rank-1 { background: #FFD700; }
        .rank-2 { background: #C0C0C0; }
        .rank-3 { background: #CD7F32; }
        .rank-other { background: #6c757d; }
        .member-badge { display: inline-block; padding: 4px 8px; margin: 2px; background: var(--primary); color: white; border-radius: 10px; cursor: pointer; font-size: 0.75rem; }
        .progress { height: 18px; }
        .chart-container { height: 220px; }
        .chart-select { font-size: 0.75rem; padding: 3px 8px; width: 75px; }
    </style>
</head>
<body>
    <nav class="navbar navbar-dark mb-2 py-1">
        <div class="container-fluid">
            <span class="navbar-brand py-0 fs-6"><i class="bi bi-bar-chart-fill me-2"></i>System Task Dashboard</span>
            <div>
                <button class="btn btn-outline-light btn-sm" onclick="exportHTML()" title="匯出 HTML"><i class="bi bi-filetype-html"></i></button>
                <span class="text-white-50 small ms-2">v17</span>
            </div>
        </div>
    </nav>

    <div class="loading" id="loading" style="display:none;">
        <div class="text-center">
            <div class="spinner-border text-primary"></div>
            <p class="mt-2">處理中...</p>
        </div>
    </div>

    <div class="container-fluid px-2">
        <div class="config-ok"><i class="bi bi-check-circle me-1"></i>已連接 Outlook ({{ fc }} 個資料夾)</div>

        <div class="card">
            <div class="card-header"><span class="card-header-title"><i class="bi bi-folder me-1"></i>資料來源</span></div>
            <div class="card-body py-2">
                <div class="row g-2">
                    <div class="col-md-4">
                        <div class="tree-box"><div class="tree" id="folderTree"></div></div>
                        <small class="text-muted" style="font-size:0.7rem">已選: <span id="selectedName" class="text-primary">-</span></small>
                    </div>
                    <div class="col-md-8">
                        <div class="row g-2 mb-2">
                            <div class="col-3"><input type="date" class="form-control form-control-sm" id="startDate"></div>
                            <div class="col-3"><input type="date" class="form-control form-control-sm" id="endDate"></div>
                            <div class="col-3"><button class="btn btn-primary btn-sm w-100" onclick="analyze()"><i class="bi bi-search me-1"></i>分析</button></div>
                            <div class="col-3"><button class="btn btn-outline-secondary btn-sm w-100" onclick="toggleFilterSettings()"><i class="bi bi-gear me-1"></i>篩選設定</button></div>
                        </div>
                        <div id="filterSettings" style="display:none;" class="mb-2 p-2 bg-light rounded">
                            <div class="form-check form-check-inline">
                                <input class="form-check-input" type="checkbox" id="excludeMiddlePriority" checked>
                                <label class="form-check-label small" for="excludeMiddlePriority">排除 Middle priority 以下的任務</label>
                            </div>
                            <div class="form-check form-check-inline">
                                <input class="form-check-input" type="checkbox" id="excludeAfter5pm" checked>
                                <label class="form-check-label small" for="excludeAfter5pm">排除下午 5:00 後的 Mail</label>
                            </div>
                        </div>
                        <div class="drop-zone py-2" id="dropZone">
                            <i class="bi bi-cloud-upload text-muted"></i>
                            <span class="small text-muted">拖放 .msg</span>
                            <input type="file" id="fileInput" multiple accept=".msg" style="display:none;">
                        </div>
                        <div id="fileList"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="results" style="display:none;">
            <!-- 統計卡片 -->
            <div class="row g-2 mb-2">
                <div class="col"><div class="card stat-card" onclick="showAllTasks()"><div class="stat-number" id="totalTasks">0</div><div class="stat-label">總任務</div></div></div>
                <div class="col"><div class="card stat-card" onclick="showByStatus('pending')"><div class="stat-number warning" id="pendingCount">0</div><div class="stat-label">Pending</div></div></div>
                <div class="col"><div class="card stat-card" onclick="showByStatus('in_progress')"><div class="stat-number info" id="inProgressCount">0</div><div class="stat-label">進行中</div></div></div>
                <div class="col"><div class="card stat-card" onclick="showOverdue()"><div class="stat-number danger" id="overdueCount">0</div><div class="stat-label">超期</div></div></div>
                <div class="col"><div class="card stat-card" onclick="showMembers()"><div class="stat-number" id="totalMembers">0</div><div class="stat-label">成員</div></div></div>
                <div class="col"><div class="card stat-card" onclick="exportExcel()"><i class="bi bi-file-excel fs-5 text-success"></i><div class="stat-label">Excel</div></div></div>
            </div>

            <!-- 進度條 -->
            <div class="card mb-2">
                <div class="card-body py-2">
                    <div class="d-flex justify-content-between small mb-1">
                        <strong>任務狀態（未完成任務）</strong>
                        <span>最後郵件: <span id="lastMailDate">-</span></span>
                    </div>
                    <div class="progress">
                        <div class="progress-bar bg-info" id="inProgressBar" title="進行中"></div>
                        <div class="progress-bar bg-warning" id="pendingBar" title="Pending"></div>
                    </div>
                    <div class="d-flex justify-content-between mt-1" style="font-size:0.7rem">
                        <span class="text-info">🔄 進行中 <span id="inProgressPct">0</span>%</span>
                        <span class="text-warning">⏳ Pending <span id="pendingPct">0</span>%</span>
                    </div>
                </div>
            </div>

            <!-- 圖表 -->
            <div class="row g-2 mb-2">
                <div class="col-md-4">
                    <div class="card">
                        <div class="card-header">
                            <span class="card-header-title"><i class="bi bi-pie-chart me-1"></i>狀態分佈</span>
                            <select class="form-select chart-select" id="chart1Type" onchange="updateChart1()">
                                <option value="doughnut">環形</option>
                                <option value="pie">圓餅</option>
                                <option value="bar">長條</option>
                                <option value="polarArea">極區</option>
                            </select>
                        </div>
                        <div class="card-body py-2">
                            <div class="chart-container"><canvas id="chart1"></canvas></div>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card">
                        <div class="card-header">
                            <span class="card-header-title"><i class="bi bi-bar-chart me-1"></i>優先級</span>
                            <select class="form-select chart-select" id="chart2Type" onchange="updateChart2()">
                                <option value="doughnut">環形</option>
                                <option value="pie">圓餅</option>
                                <option value="bar">長條</option>
                                <option value="polarArea">極區</option>
                            </select>
                        </div>
                        <div class="card-body py-2">
                            <div class="chart-container"><canvas id="chart2"></canvas></div>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card">
                        <div class="card-header">
                            <span class="card-header-title"><i class="bi bi-exclamation-triangle me-1"></i>超期狀況</span>
                            <select class="form-select chart-select" id="chart3Type" onchange="updateChart3()">
                                <option value="doughnut">環形</option>
                                <option value="pie">圓餅</option>
                                <option value="bar">長條</option>
                                <option value="polarArea">極區</option>
                            </select>
                        </div>
                        <div class="card-body py-2">
                            <div class="chart-container"><canvas id="chart3"></canvas></div>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card">
                        <div class="card-header">
                            <span class="card-header-title"><i class="bi bi-person-exclamation me-1"></i>成員超期天數</span>
                            <select class="form-select chart-select" id="chart4Type" onchange="updateChart4()">
                                <option value="stacked" selected>堆疊長條</option>
                                <option value="bar">長條</option>
                                <option value="line">折線</option>
                                <option value="doughnut">環形</option>
                            </select>
                        </div>
                        <div class="card-body py-2">
                            <div class="chart-container"><canvas id="chart4"></canvas></div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- 任務列表 -->
            <div class="card mb-2">
                <div class="card-header"><span class="card-header-title"><i class="bi bi-list-task me-1"></i>任務列表</span></div>
                <div class="table-toolbar">
                    <input type="text" class="form-control form-control-sm" placeholder="🔍 搜尋..." id="taskSearch" onkeyup="filterAndRenderTaskTable()">
                    <button class="btn btn-outline-secondary btn-sm" onclick="exportTableCSV('task')"><i class="bi bi-download"></i></button>
                </div>
                <div class="table-container">
                    <table class="table table-sm data-table mb-0">
                        <thead>
                            <tr>
                                <th onclick="sortData('task', 'last_seen')">Mail日期 ↕</th>
                                <th onclick="sortData('task', 'module')">模組 ↕</th>
                                <th onclick="sortData('task', 'title')">任務 ↕</th>
                                <th onclick="sortData('task', 'owners_str')">負責人 ↕</th>
                                <th onclick="sortData('task', 'priority')">優先級 ↕</th>
                                <th onclick="sortData('task', 'due')">Due ↕</th>
                                <th onclick="sortData('task', 'overdue_days')">超期 ↕</th>
                                <th onclick="sortData('task', 'task_status')">狀態 ↕</th>
                            </tr>
                        </thead>
                        <tbody id="taskTableBody"></tbody>
                    </table>
                </div>
                <div class="pagination-controls">
                    <div>
                        <span>每頁</span>
                        <select class="form-select form-select-sm d-inline-block" style="width:70px" id="taskPageSize" onchange="changePageSize('task')">
                            <option value="50" selected>50</option>
                            <option value="100">100</option>
                            <option value="200">200</option>
                            <option value="500">500</option>
                            <option value="1000">1000</option>
                        </select>
                        <span>筆</span>
                    </div>
                    <span id="taskPageInfo">-</span>
                    <div>
                        <button class="btn btn-outline-secondary btn-sm" onclick="changePage('task', -1)">◀ 上一頁</button>
                        <button class="btn btn-outline-secondary btn-sm" onclick="changePage('task', 1)">下一頁 ▶</button>
                    </div>
                </div>
            </div>

            <!-- 成員統計 & 貢獻度 -->
            <div class="row g-2">
                <div class="col-md-7">
                    <div class="card">
                        <div class="card-header"><span class="card-header-title"><i class="bi bi-people me-1"></i>成員統計</span></div>
                        <div class="table-toolbar">
                            <input type="text" class="form-control form-control-sm" placeholder="🔍 搜尋..." id="memberSearch" onkeyup="filterAndRenderMemberTable()">
                            <button class="btn btn-outline-secondary btn-sm" onclick="exportTableCSV('member')"><i class="bi bi-download"></i></button>
                        </div>
                        <div class="table-container">
                            <table class="table table-sm data-table mb-0">
                                <thead>
                                    <tr>
                                        <th onclick="sortData('member', 'name')">成員 ↕</th>
                                        <th onclick="sortData('member', 'total')">總數 ↕</th>
                                        <th onclick="sortData('member', 'completed')">完成 ↕</th>
                                        <th onclick="sortData('member', 'in_progress')">進行 ↕</th>
                                        <th onclick="sortData('member', 'pending')">Pend ↕</th>
                                        <th onclick="sortData('member', 'high')">H ↕</th>
                                        <th onclick="sortData('member', 'medium')">M ↕</th>
                                        <th onclick="sortData('member', 'normal')">N ↕</th>
                                    </tr>
                                </thead>
                                <tbody id="memberTableBody"></tbody>
                            </table>
                        </div>
                        <div class="pagination-controls">
                            <div>
                                <span>每頁</span>
                                <select class="form-select form-select-sm d-inline-block" style="width:70px" id="memberPageSize" onchange="changePageSize('member')">
                                    <option value="50" selected>50</option>
                                    <option value="100">100</option>
                                    <option value="200">200</option>
                                    <option value="500">500</option>
                                    <option value="1000">1000</option>
                                </select>
                                <span>筆</span>
                            </div>
                            <span id="memberPageInfo">-</span>
                            <div>
                                <button class="btn btn-outline-secondary btn-sm" onclick="changePage('member', -1)">◀ 上一頁</button>
                                <button class="btn btn-outline-secondary btn-sm" onclick="changePage('member', 1)">下一頁 ▶</button>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="col-md-5">
                    <div class="card">
                        <div class="card-header"><span class="card-header-title"><i class="bi bi-trophy me-1"></i>貢獻度 <small class="text-warning">(含超期減分)</small></span></div>
                        <div class="table-toolbar">
                            <input type="text" class="form-control form-control-sm" placeholder="🔍 搜尋..." id="contribSearch" onkeyup="filterAndRenderContribTable()">
                            <button class="btn btn-outline-secondary btn-sm" onclick="exportTableCSV('contrib')"><i class="bi bi-download"></i></button>
                        </div>
                        <div class="table-container">
                            <table class="table table-sm data-table mb-0">
                                <thead>
                                    <tr>
                                        <th onclick="sortData('contrib', 'rank')"># ↕</th>
                                        <th onclick="sortData('contrib', 'name')">成員 ↕</th>
                                        <th onclick="sortData('contrib', 'task_count')">任務 ↕</th>
                                        <th onclick="sortData('contrib', 'base_score')">基礎分 ↕</th>
                                        <th onclick="sortData('contrib', 'overdue_count')">超期數 ↕</th>
                                        <th onclick="sortData('contrib', 'overdue_penalty')">扣分 ↕</th>
                                        <th onclick="sortData('contrib', 'score')">總分 ↕</th>
                                    </tr>
                                </thead>
                                <tbody id="contribTableBody"></tbody>
                            </table>
                        </div>
                        <div class="pagination-controls">
                            <div>
                                <span>每頁</span>
                                <select class="form-select form-select-sm d-inline-block" style="width:70px" id="contribPageSize" onchange="changePageSize('contrib')">
                                    <option value="50" selected>50</option>
                                    <option value="100">100</option>
                                    <option value="200">200</option>
                                    <option value="500">500</option>
                                    <option value="1000">1000</option>
                                </select>
                                <span>筆</span>
                            </div>
                            <span id="contribPageInfo">-</span>
                            <div>
                                <button class="btn btn-outline-secondary btn-sm" onclick="changePage('contrib', -1)">◀ 上一頁</button>
                                <button class="btn btn-outline-secondary btn-sm" onclick="changePage('contrib', 1)">下一頁 ▶</button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div class="footer">© 2025 Vince Lin. All rights reserved.</div>
    </div>

    <!-- Modal -->
    <div class="modal fade" id="detailModal" tabindex="-1">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header py-2">
                    <h6 class="modal-title" id="modalTitle">明細</h6>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body" style="max-height:70vh;overflow-y:auto;"><div id="modalContent"></div></div>
                <div class="modal-footer py-1">
                    <button class="btn btn-outline-secondary btn-sm" onclick="exportModalCSV()"><i class="bi bi-download me-1"></i>CSV</button>
                    <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">關閉</button>
                </div>
            </div>
        </div>
    </div>
    
    <!-- Mail Preview Modal -->
    <div class="modal fade" id="mailModal" tabindex="-1">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header py-2 bg-primary text-white">
                    <h6 class="modal-title"><i class="bi bi-envelope me-1"></i>Mail 預覽</h6>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body p-0">
                    <div class="p-2 bg-light border-bottom d-flex justify-content-between align-items-center">
                        <div>
                            <div><strong>主旨：</strong><span id="mailSubject">-</span></div>
                            <div><strong>日期：</strong><span id="mailDate">-</span> <span id="mailTime" class="text-muted"></span></div>
                        </div>
                        <div class="btn-group btn-group-sm">
                            <button class="btn btn-outline-secondary active" onclick="setMailView('html')" id="btnHtml">HTML</button>
                            <button class="btn btn-outline-secondary" onclick="setMailView('text')" id="btnText">純文字</button>
                        </div>
                    </div>
                    <div id="mailBodyHtml" style="height:60vh;overflow:hidden;">
                        <iframe id="mailIframe" style="width:100%;height:100%;border:none;"></iframe>
                    </div>
                    <div id="mailBodyText" style="max-height:60vh;overflow-y:auto;padding:15px;font-family:monospace;font-size:13px;white-space:pre-wrap;background:#fafafa;display:none;"></div>
                </div>
                <div class="modal-footer py-1">
                    <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">關閉</button>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        const treeData = {{ tree | tojson }};
        let selectedEntry = null, selectedStore = null, resultData = null;
        let chart1 = null, chart2 = null, chart3 = null, chart4 = null, currentModal = null;
        
        // 分頁設定（每個表格獨立）
        let tableState = {
            task: { data: [], filtered: [], page: 0, pageSize: 50, sortKey: '', sortDir: 1 },
            member: { data: [], filtered: [], page: 0, pageSize: 50, sortKey: '', sortDir: 1 },
            contrib: { data: [], filtered: [], page: 0, pageSize: 50, sortKey: '', sortDir: 1 }
        };

        // Tree
        function buildTree(data) {
            const ul = document.createElement('ul');
            data.forEach(node => {
                const li = document.createElement('li');
                if (node.children && node.children.length > 0) {
                    const toggle = document.createElement('span');
                    toggle.className = 'tree-toggle' + (node.is_archive ? ' archive' : '');
                    toggle.textContent = node.name;
                    toggle.onclick = function(e) { e.stopPropagation(); this.classList.toggle('open'); this.nextElementSibling.style.display = this.nextElementSibling.style.display === 'none' ? 'block' : 'none'; };
                    li.appendChild(toggle);
                    const childUl = buildTree(node.children);
                    childUl.style.display = 'none';
                    li.appendChild(childUl);
                } else if (node.entry_id) {
                    const item = document.createElement('span');
                    item.className = 'tree-item' + (node.is_archive ? ' archive' : '');
                    item.textContent = node.name;
                    item.onclick = function(e) { e.stopPropagation(); document.querySelectorAll('.tree-item.selected').forEach(el => el.classList.remove('selected')); this.classList.add('selected'); selectedEntry = node.entry_id; selectedStore = node.store_id; document.getElementById('selectedName').textContent = node.name; };
                    if (node.name === 'Dias-System team 協助事項' && !node.is_archive) setTimeout(() => { item.click(); let p = item.parentElement; while (p) { const t = p.querySelector(':scope > .tree-toggle'); if (t && !t.classList.contains('open')) t.click(); p = p.parentElement?.closest('li'); } }, 100);
                    li.appendChild(item);
                }
                ul.appendChild(li);
            });
            return ul;
        }
        document.getElementById('folderTree').appendChild(buildTree(treeData));
        document.querySelectorAll('.tree > ul > li > .tree-toggle').forEach(t => { if (!t.classList.contains('archive')) t.click(); });

        const today = new Date(), monthAgo = new Date(today.getTime() - 30*24*60*60*1000);
        document.getElementById('endDate').value = today.toISOString().split('T')[0];
        document.getElementById('startDate').value = monthAgo.toISOString().split('T')[0];

        // Analyze
        function toggleFilterSettings() {
            const el = document.getElementById('filterSettings');
            el.style.display = el.style.display === 'none' ? 'block' : 'none';
        }
        
        async function analyze() {
            if (!selectedEntry) { alert('請選擇資料夾'); return; }
            document.getElementById('loading').style.display = 'flex';
            try {
                const excludeMiddlePriority = document.getElementById('excludeMiddlePriority').checked;
                const excludeAfter5pm = document.getElementById('excludeAfter5pm').checked;
                const r = await fetch('/api/outlook', { 
                    method: 'POST', 
                    headers: {'Content-Type': 'application/json'}, 
                    body: JSON.stringify({
                        entry_id: selectedEntry, 
                        store_id: selectedStore, 
                        start: document.getElementById('startDate').value, 
                        end: document.getElementById('endDate').value,
                        exclude_middle_priority: excludeMiddlePriority,
                        exclude_after_5pm: excludeAfter5pm
                    }) 
                });
                const data = await r.json();
                document.getElementById('loading').style.display = 'none';
                if (r.ok) { resultData = data; renderResults(data); } else alert(data.error || '分析失敗');
            } catch (e) { document.getElementById('loading').style.display = 'none'; alert('錯誤: ' + e); }
        }

        // Drop zone
        const dropZone = document.getElementById('dropZone'), fileInput = document.getElementById('fileInput');
        dropZone.onclick = () => fileInput.click();
        ['dragenter','dragover','dragleave','drop'].forEach(e => dropZone.addEventListener(e, ev => { ev.preventDefault(); ev.stopPropagation(); }));
        dropZone.addEventListener('dragover', () => dropZone.classList.add('dragover'));
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
        dropZone.addEventListener('drop', e => { dropZone.classList.remove('dragover'); handleFiles(e.dataTransfer.files); });
        fileInput.addEventListener('change', e => handleFiles(e.target.files));
        function handleFiles(files) { const msgs = Array.from(files).filter(f => f.name.endsWith('.msg')); if (!msgs.length) return; window.uploadFiles = msgs; document.getElementById('fileList').innerHTML = `<div class="alert alert-info py-1 mt-1 small">${msgs.length} 檔 <button class="btn btn-primary btn-sm ms-2" onclick="uploadAnalyze()">分析</button></div>`; }
        async function uploadAnalyze() { 
            if (!window.uploadFiles) return; 
            const fd = new FormData(); 
            window.uploadFiles.forEach(f => fd.append('f', f)); 
            fd.append('exclude_middle_priority', document.getElementById('excludeMiddlePriority').checked);
            fd.append('exclude_after_5pm', document.getElementById('excludeAfter5pm').checked);
            document.getElementById('loading').style.display = 'flex'; 
            try { 
                const r = await fetch('/api/upload', { method: 'POST', body: fd }); 
                const data = await r.json(); 
                document.getElementById('loading').style.display = 'none'; 
                if (r.ok) { resultData = data; renderResults(data); } else alert(data.error); 
            } catch (e) { document.getElementById('loading').style.display = 'none'; alert(e); } 
        }

        // Render
        function renderResults(data) {
            const total = data.total_tasks || 1;
            const activeTotal = data.pending_count + data.in_progress_count || 1;
            
            document.getElementById('totalTasks').textContent = data.total_tasks;
            document.getElementById('pendingCount').textContent = data.pending_count;
            document.getElementById('inProgressCount').textContent = data.in_progress_count;
            document.getElementById('overdueCount').textContent = data.overdue_count;
            document.getElementById('totalMembers').textContent = data.total_members;
            document.getElementById('lastMailDate').textContent = data.last_mail_date || '-';

            const pendingPct = (data.pending_count / activeTotal * 100).toFixed(1);
            const inProgressPct = (data.in_progress_count / activeTotal * 100).toFixed(1);
            document.getElementById('pendingBar').style.width = pendingPct + '%';
            document.getElementById('inProgressBar').style.width = inProgressPct + '%';
            document.getElementById('pendingPct').textContent = pendingPct;
            document.getElementById('inProgressPct').textContent = inProgressPct;
            document.getElementById('results').style.display = '';

            // 初始化表格資料
            tableState.task.data = data.all_tasks;
            tableState.task.filtered = [...data.all_tasks];
            tableState.task.page = 0;
            
            tableState.member.data = data.members;
            tableState.member.filtered = [...data.members];
            tableState.member.page = 0;
            
            // 貢獻度加上 rank
            tableState.contrib.data = data.contribution.map((c, i) => ({...c, rank: i + 1}));
            tableState.contrib.filtered = [...tableState.contrib.data];
            tableState.contrib.page = 0;

            renderTaskTable();
            renderMemberTable();
            renderContribTable();

            updateChart1();
            updateChart2();
            updateChart3();
            updateChart4();
        }

        // 分頁渲染
        const statusLabels = { completed: '已完成', pending: 'Pending', in_progress: '進行中' };

        function renderTaskTable() {
            const state = tableState.task;
            const start = state.page * state.pageSize;
            const pageData = state.filtered.slice(start, start + state.pageSize);
            
            document.getElementById('taskTableBody').innerHTML = pageData.map(t => `
                <tr class="row-${t.task_status} ${t.overdue_days > 0 ? 'row-overdue' : ''}">
                    <td>${t.last_seen || '-'}</td>
                    <td><span class="badge bg-secondary" style="font-size:0.65rem">${t.module || '-'}</span></td>
                    <td>
                        <span style="cursor:pointer" onclick="showTaskDetail('${esc(t.title)}')">${t.title}</span>
                        ${t.mail_id ? `<i class="bi bi-envelope ms-1 text-primary" style="cursor:pointer;font-size:0.8rem" onclick="showMailPreview('${t.mail_id}', event)" title="預覽 Mail"></i>` : ''}
                    </td>
                    <td>${t.owners_str}</td>
                    <td><span class="badge badge-${t.priority}">${t.priority}</span></td>
                    <td class="${t.overdue_days > 0 ? 'text-overdue' : ''}">${t.due || '-'}</td>
                    <td class="${t.overdue_days > 0 ? 'text-overdue' : ''}">${t.overdue_days > 0 ? '+' + t.overdue_days + '天' : '-'}</td>
                    <td><span class="badge badge-${t.task_status}">${statusLabels[t.task_status]}</span></td>
                </tr>
            `).join('');
            
            const totalPages = Math.ceil(state.filtered.length / state.pageSize);
            document.getElementById('taskPageInfo').textContent = `第 ${state.page + 1}/${totalPages || 1} 頁 (共 ${state.filtered.length} 筆)`;
        }

        function renderMemberTable() {
            const state = tableState.member;
            const start = state.page * state.pageSize;
            const pageData = state.filtered.slice(start, start + state.pageSize);
            
            document.getElementById('memberTableBody').innerHTML = pageData.map(m => `
                <tr>
                    <td><strong style="cursor:pointer" onclick="showMemberTasks('${esc(m.name)}')">${m.name}</strong></td>
                    <td style="cursor:pointer" onclick="showMemberTasks('${esc(m.name)}')">${m.total}</td>
                    <td style="cursor:pointer" onclick="showMemberTasksByStatus('${esc(m.name)}', 'completed')"><span class="badge badge-completed">${m.completed}</span></td>
                    <td style="cursor:pointer" onclick="showMemberTasksByStatus('${esc(m.name)}', 'in_progress')"><span class="badge badge-in_progress">${m.in_progress}</span></td>
                    <td style="cursor:pointer" onclick="showMemberTasksByStatus('${esc(m.name)}', 'pending')"><span class="badge badge-pending">${m.pending}</span></td>
                    <td style="cursor:pointer" onclick="showMemberTasksByPriority('${esc(m.name)}', 'high')"><span class="badge badge-high">${m.high}</span></td>
                    <td style="cursor:pointer" onclick="showMemberTasksByPriority('${esc(m.name)}', 'medium')"><span class="badge badge-medium">${m.medium}</span></td>
                    <td style="cursor:pointer" onclick="showMemberTasksByPriority('${esc(m.name)}', 'normal')"><span class="badge badge-normal">${m.normal}</span></td>
                </tr>
            `).join('');
            
            const totalPages = Math.ceil(state.filtered.length / state.pageSize);
            document.getElementById('memberPageInfo').textContent = `第 ${state.page + 1}/${totalPages || 1} 頁 (共 ${state.filtered.length} 筆)`;
        }

        function renderContribTable() {
            const state = tableState.contrib;
            const start = state.page * state.pageSize;
            const pageData = state.filtered.slice(start, start + state.pageSize);
            
            document.getElementById('contribTableBody').innerHTML = pageData.map(c => `
                <tr onclick="showMemberNonPendingTasks('${esc(c.name)}')">
                    <td><span class="rank-badge ${c.rank <= 3 ? 'rank-' + c.rank : 'rank-other'}">${c.rank}</span></td>
                    <td>${c.name}</td>
                    <td>${c.task_count}</td>
                    <td>${c.base_score}</td>
                    <td class="${c.overdue_count > 0 ? 'text-overdue' : ''}">${c.overdue_count}</td>
                    <td class="${c.overdue_penalty > 0 ? 'text-overdue' : ''}">-${c.overdue_penalty}</td>
                    <td><strong>${c.score}</strong></td>
                </tr>
            `).join('');
            
            const totalPages = Math.ceil(state.filtered.length / state.pageSize);
            document.getElementById('contribPageInfo').textContent = `第 ${state.page + 1}/${totalPages || 1} 頁 (共 ${state.filtered.length} 筆)`;
        }

        // 改變每頁筆數
        function changePageSize(table) {
            const select = document.getElementById(table + 'PageSize');
            const newSize = parseInt(select.value);
            tableState[table].pageSize = newSize;
            tableState[table].page = 0;  // 重置到第一頁
            if (table === 'task') renderTaskTable();
            else if (table === 'member') renderMemberTable();
            else renderContribTable();
        }

        // 分頁切換
        function changePage(table, delta) {
            const state = tableState[table];
            const totalPages = Math.ceil(state.filtered.length / state.pageSize);
            state.page = Math.max(0, Math.min(totalPages - 1, state.page + delta));
            if (table === 'task') renderTaskTable();
            else if (table === 'member') renderMemberTable();
            else renderContribTable();
        }

        // 排序
        function sortData(table, key) {
            const state = tableState[table];
            if (state.sortKey === key) state.sortDir *= -1;
            else { state.sortKey = key; state.sortDir = 1; }
            
            state.filtered.sort((a, b) => {
                let av = a[key], bv = b[key];
                if (typeof av === 'number') return (av - bv) * state.sortDir;
                return String(av || '').localeCompare(String(bv || ''), 'zh-TW') * state.sortDir;
            });
            
            state.page = 0;
            if (table === 'task') renderTaskTable();
            else if (table === 'member') renderMemberTable();
            else renderContribTable();
        }

        // 搜尋
        function filterAndRenderTaskTable() {
            const q = document.getElementById('taskSearch').value.toLowerCase();
            tableState.task.filtered = tableState.task.data.filter(t => 
                t.title.toLowerCase().includes(q) || t.owners_str.toLowerCase().includes(q)
            );
            tableState.task.page = 0;
            renderTaskTable();
        }

        function filterAndRenderMemberTable() {
            const q = document.getElementById('memberSearch').value.toLowerCase();
            tableState.member.filtered = tableState.member.data.filter(m => m.name.toLowerCase().includes(q));
            tableState.member.page = 0;
            renderMemberTable();
        }

        function filterAndRenderContribTable() {
            const q = document.getElementById('contribSearch').value.toLowerCase();
            tableState.contrib.filtered = tableState.contrib.data.filter(c => c.name.toLowerCase().includes(q));
            tableState.contrib.page = 0;
            renderContribTable();
        }

        // Charts
        function updateChart1() {
            const type = document.getElementById('chart1Type').value;
            if (chart1) chart1.destroy();
            const ctx = document.getElementById('chart1').getContext('2d');
            
            chart1 = new Chart(ctx, {
                type: type === 'bar' ? 'bar' : type,
                data: { labels: ['已完成', '進行中', 'Pending'], datasets: [{ data: [resultData.completed_count, resultData.in_progress_count, resultData.pending_count], backgroundColor: ['#28a745', '#17a2b8', '#FFA500'] }] },
                options: { maintainAspectRatio: false, plugins: { legend: { display: type !== 'bar', position: 'right' } }, onClick: (e, el) => { if (el.length) showByStatus(['completed', 'in_progress', 'pending'][el[0].index]); } }
            });
        }

        function updateChart2() {
            const type = document.getElementById('chart2Type').value;
            if (chart2) chart2.destroy();
            const ctx = document.getElementById('chart2').getContext('2d');
            
            chart2 = new Chart(ctx, {
                type: type === 'bar' ? 'bar' : type,
                data: { labels: ['High', 'Medium', 'Normal'], datasets: [{ data: [resultData.priority_counts.high, resultData.priority_counts.medium, resultData.priority_counts.normal], backgroundColor: ['#FF6B6B', '#FFE066', '#74C0FC'] }] },
                options: { maintainAspectRatio: false, plugins: { legend: { display: type !== 'bar', position: 'right' } }, onClick: (e, el) => { if (el.length) showByPriority(['high', 'medium', 'normal'][el[0].index]); } }
            });
        }

        function updateChart3() {
            const type = document.getElementById('chart3Type').value;
            if (chart3) chart3.destroy();
            const ctx = document.getElementById('chart3').getContext('2d');
            
            // 超期圖：只看未完成的任務（進行中+Pending）
            const overdue = resultData.overdue_count;
            const notOverdue = resultData.not_overdue_count;
            
            chart3 = new Chart(ctx, {
                type: type === 'bar' ? 'bar' : type,
                data: { labels: ['超期', '未超期'], datasets: [{ data: [overdue, notOverdue], backgroundColor: ['#dc3545', '#28a745'] }] },
                options: { maintainAspectRatio: false, plugins: { legend: { display: type !== 'bar', position: 'right' } }, onClick: (e, el) => { if (el.length && el[0].index === 0) showOverdue(); else if (el.length && el[0].index === 1) showNotOverdue(); } }
            });
        }

        function updateChart4() {
            const type = document.getElementById('chart4Type').value;
            if (chart4) chart4.destroy();
            const ctx = document.getElementById('chart4').getContext('2d');
            
            // 成員超期天數統計（取超期天數前 10 名）
            const overdueData = resultData.contribution
                .filter(c => c.overdue_days > 0)
                .sort((a, b) => b.overdue_days - a.overdue_days)
                .slice(0, 10);
            
            if (overdueData.length === 0) {
                chart4 = new Chart(ctx, {
                    type: 'bar',
                    data: { labels: ['無超期'], datasets: [{ data: [0], backgroundColor: '#28a745' }] },
                    options: { maintainAspectRatio: false, plugins: { legend: { display: false } } }
                });
                return;
            }
            
            const labels = overdueData.map(c => c.name);
            
            if (type === 'stacked') {
                // 堆疊長條圖：已完成 vs 未完成
                const completedData = overdueData.map(c => c.completed_overdue_days || 0);
                const activeData = overdueData.map(c => c.active_overdue_days || 0);
                
                chart4 = new Chart(ctx, {
                    type: 'bar',
                    data: {
                        labels: labels,
                        datasets: [
                            { label: '已完成超期', data: completedData, backgroundColor: '#6c757d', stack: 'stack1' },
                            { label: '未完成超期', data: activeData, backgroundColor: '#dc3545', stack: 'stack1' }
                        ]
                    },
                    options: {
                        maintainAspectRatio: false,
                        indexAxis: 'y',
                        plugins: { legend: { display: true, position: 'top' } },
                        scales: { x: { stacked: true, beginAtZero: true }, y: { stacked: true } },
                        onClick: (e, el) => { if (el.length) showMemberOverdueTasks(labels[el[0].index]); }
                    }
                });
            } else if (type === 'line') {
                // 折線圖
                const completedData = overdueData.map(c => c.completed_overdue_days || 0);
                const activeData = overdueData.map(c => c.active_overdue_days || 0);
                
                chart4 = new Chart(ctx, {
                    type: 'line',
                    data: {
                        labels: labels,
                        datasets: [
                            { label: '已完成超期', data: completedData, borderColor: '#6c757d', backgroundColor: 'rgba(108,117,125,0.2)', fill: true, tension: 0.3 },
                            { label: '未完成超期', data: activeData, borderColor: '#dc3545', backgroundColor: 'rgba(220,53,69,0.2)', fill: true, tension: 0.3 }
                        ]
                    },
                    options: {
                        maintainAspectRatio: false,
                        plugins: { legend: { display: true, position: 'top' } },
                        scales: { y: { beginAtZero: true } },
                        onClick: (e, el) => { if (el.length) showMemberOverdueTasks(labels[el[0].index]); }
                    }
                });
            } else if (type === 'doughnut') {
                const data = overdueData.map(c => c.overdue_days);
                const colors = ['#FF6B6B', '#FFA500', '#FFE066', '#74C0FC', '#69DB7C', '#B197FC', '#F783AC', '#20C997', '#ADB5BD', '#868E96'];
                
                chart4 = new Chart(ctx, {
                    type: 'doughnut',
                    data: { labels: labels, datasets: [{ data: data, backgroundColor: colors.slice(0, labels.length) }] },
                    options: {
                        maintainAspectRatio: false,
                        plugins: { legend: { display: true, position: 'right' } },
                        onClick: (e, el) => { if (el.length) showMemberOverdueTasks(labels[el[0].index]); }
                    }
                });
            } else {
                // 一般長條圖
                const data = overdueData.map(c => c.overdue_days);
                const maxDays = Math.max(...data);
                const highThreshold = maxDays * 0.7;
                const midThreshold = maxDays * 0.4;
                
                chart4 = new Chart(ctx, {
                    type: 'bar',
                    data: { 
                        labels: labels, 
                        datasets: [{ 
                            label: '超期天數',
                            data: data, 
                            backgroundColor: data.map(d => d >= highThreshold ? '#dc3545' : d >= midThreshold ? '#FFA500' : '#FFE066')
                        }] 
                    },
                    options: { 
                        maintainAspectRatio: false, 
                        indexAxis: 'y',
                        plugins: { legend: { display: false } },
                        scales: { x: { beginAtZero: true } },
                        onClick: (e, el) => { if (el.length) showMemberOverdueTasks(labels[el[0].index]); }
                    }
                });
            }
        }

        // CSV Export
        function exportTableCSV(table) {
            const state = tableState[table];
            let csv = [];
            let headers, getData;
            
            if (table === 'task') {
                headers = ['Mail日期', '模組', '任務', '負責人', '優先級', 'Due', '超期天數', '狀態'];
                getData = t => [t.last_seen || '', t.module || '', t.title, t.owners_str, t.priority, t.due || '', t.overdue_days || 0, statusLabels[t.task_status]];
            } else if (table === 'member') {
                headers = ['成員', '總數', '完成', '進行', 'Pending', 'High', 'Medium', 'Normal'];
                getData = m => [m.name, m.total, m.completed, m.in_progress, m.pending, m.high, m.medium, m.normal];
            } else {
                headers = ['排名', '成員', '任務數', '基礎分', '超期數', '扣分', '總分'];
                getData = c => [c.rank, c.name, c.task_count, c.base_score, c.overdue_count, c.overdue_penalty, c.score];
            }
            
            csv.push(headers.join(','));
            state.filtered.forEach(item => csv.push(getData(item).map(v => '"' + String(v).replace(/"/g, '""') + '"').join(',')));
            downloadCSV(csv.join('\\n'), table + '.csv');
        }

        function exportModalCSV() {
            const table = document.querySelector('#modalContent table');
            if (!table) return;
            let csv = [];
            csv.push(Array.from(table.querySelectorAll('thead th')).map(th => th.textContent.replace(/[↕]/g, '').trim()).join(','));
            table.querySelectorAll('tbody tr').forEach(row => csv.push(Array.from(row.cells).map(td => '"' + td.textContent.trim().replace(/"/g, '""') + '"').join(',')));
            downloadCSV(csv.join('\\n'), 'export.csv');
        }

        function downloadCSV(content, filename) {
            const blob = new Blob(['\\uFEFF' + content], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = filename;
            link.click();
        }

        // Modal
        function esc(s) { return s.replace(/'/g, "\\'").replace(/"/g, '\\"'); }
        function showModal(title, content) { if (currentModal) currentModal.hide(); document.getElementById('modalTitle').textContent = title; document.getElementById('modalContent').innerHTML = content; currentModal = new bootstrap.Modal(document.getElementById('detailModal')); currentModal.show(); }
        
        function modalTable(tasks, id = 'modalTableBody') {
            return `
                <div class="mb-2"><input type="text" class="form-control form-control-sm" style="max-width:250px" placeholder="🔍 搜尋..." onkeyup="filterModalTable('${id}', this.value)"></div>
                <div style="max-height: 50vh; overflow-y: auto;">
                    <table class="table table-sm data-table">
                        <thead><tr><th>Mail日期</th><th>模組</th><th>任務</th><th>負責人</th><th>優先級</th><th>Due</th><th>超期</th><th>狀態</th></tr></thead>
                        <tbody id="${id}">${tasks.map(t => `
                            <tr class="row-${t.task_status} ${t.overdue_days > 0 ? 'row-overdue' : ''}">
                                <td>${t.last_seen || t.mail_date || '-'}</td>
                                <td><span class="badge bg-secondary" style="font-size:0.6rem">${t.module || '-'}</span></td>
                                <td>
                                    ${t.title}
                                    ${t.mail_id ? `<i class="bi bi-envelope ms-1 text-primary" style="cursor:pointer;font-size:0.8rem" onclick="showMailPreview('${t.mail_id}', event)" title="預覽 Mail"></i>` : ''}
                                </td>
                                <td>${t.owners_str || (t.owners ? t.owners.join('/') : '-')}</td>
                                <td><span class="badge badge-${t.priority}">${t.priority}</span></td>
                                <td class="${t.overdue_days > 0 ? 'text-overdue' : ''}">${t.due || '-'}</td>
                                <td class="${t.overdue_days > 0 ? 'text-overdue' : ''}">${t.overdue_days > 0 ? '+' + t.overdue_days + '天' : '-'}</td>
                                <td><span class="badge badge-${t.task_status}">${statusLabels[t.task_status] || t.task_status}</span></td>
                            </tr>
                        `).join('')}</tbody>
                    </table>
                </div>
            `;
        }

        function filterModalTable(id, q) {
            q = q.toLowerCase();
            for (let row of document.getElementById(id).rows) row.style.display = row.textContent.toLowerCase().includes(q) ? '' : 'none';
        }

        function showAllTasks() { if (!resultData) return; showModal(`全部任務 (${resultData.total_tasks})`, modalTable(resultData.all_tasks)); }
        function showByStatus(status) { if (!resultData) return; const tasks = resultData.all_tasks.filter(t => t.task_status === status); showModal(`${statusLabels[status]} (${tasks.length})`, modalTable(tasks, status + 'Table')); }
        function showOverdue() { if (!resultData) return; const tasks = resultData.all_tasks.filter(t => t.is_overdue && t.task_status !== 'completed'); showModal(`超期任務 (${tasks.length})`, modalTable(tasks, 'overdueTable')); }
        function showNotOverdue() { if (!resultData) return; const tasks = resultData.all_tasks.filter(t => !t.is_overdue && t.task_status !== 'completed'); showModal(`未超期任務 (${tasks.length})`, modalTable(tasks, 'notOverdueTable')); }
        function showByPriority(p) { if (!resultData) return; const tasks = resultData.all_tasks.filter(t => t.priority === p); showModal(`${p.toUpperCase()} 優先級 (${tasks.length})`, modalTable(tasks, 'priorityTable')); }
        function showMembers() { if (!resultData) return; showModal(`成員列表 (${resultData.total_members} 人)`, `<div class="d-flex flex-wrap">${resultData.member_list.map(m => `<span class="member-badge" onclick="currentModal.hide();setTimeout(()=>showMemberTasks('${esc(m)}'),300)">${m}</span>`).join('')}</div>`); }
        function showMemberTasks(name) { if (!resultData) return; const m = resultData.members.find(x => x.name === name); if (!m) return; const tasks = m.tasks.map(t => ({...t, owners_str: t.owners?.join('/')})); showModal(`${name} 的任務 (${tasks.length})`, modalTable(tasks, 'memberTaskTable')); }
        
        // 依狀態顯示成員任務
        function showMemberTasksByStatus(name, status) {
            if (!resultData) return;
            const m = resultData.members.find(x => x.name === name);
            if (!m) return;
            const tasks = m.tasks.filter(t => t.task_status === status).map(t => ({...t, owners_str: t.owners?.join('/')}));
            const statusLabel = statusLabels[status] || status;
            showModal(`${name} 的${statusLabel}任務 (${tasks.length})`, modalTable(tasks, 'memberStatusTable'));
        }
        
        // 依優先級顯示成員任務
        function showMemberTasksByPriority(name, priority) {
            if (!resultData) return;
            const m = resultData.members.find(x => x.name === name);
            if (!m) return;
            const tasks = m.tasks.filter(t => t.priority === priority).map(t => ({...t, owners_str: t.owners?.join('/')}));
            showModal(`${name} 的 ${priority.toUpperCase()} 優先級任務 (${tasks.length})`, modalTable(tasks, 'memberPriorityTable'));
        }
        
        // 貢獻度點擊：顯示非 Pending 的任務
        function showMemberNonPendingTasks(name) {
            if (!resultData) return;
            const m = resultData.members.find(x => x.name === name);
            if (!m) return;
            
            // 過濾非 Pending 的任務（已完成 + 進行中）
            const nonPendingTasks = m.tasks.filter(t => t.task_status !== 'pending').map(t => ({...t, owners_str: t.owners?.join('/')}));
            
            // 計算權重分數和超期扣分
            const highCount = nonPendingTasks.filter(t => t.priority === 'high').length;
            const medCount = nonPendingTasks.filter(t => t.priority === 'medium').length;
            const norCount = nonPendingTasks.filter(t => t.priority === 'normal').length;
            const baseScore = highCount * 3 + medCount * 2 + norCount * 1;
            
            const overdueTasks = nonPendingTasks.filter(t => t.overdue_days > 0);
            const overdueCount = overdueTasks.length;
            const totalOverdueDays = overdueTasks.reduce((s, t) => s + (t.overdue_days || 0), 0);
            const avgOverdueDays = overdueCount > 0 ? totalOverdueDays / overdueCount : 0;
            const overdueRate = nonPendingTasks.length > 0 ? overdueCount / nonPendingTasks.length : 0;
            
            // 詳細計算扣分
            let penaltyDetails = [];
            let penalty = 0;
            
            // 1. 每個超期任務扣 0.5 分
            const penaltyPerTask = overdueCount * 0.5;
            penalty += penaltyPerTask;
            if (penaltyPerTask > 0) {
                penaltyDetails.push(`超期任務數×0.5 = ${overdueCount}×0.5 = ${penaltyPerTask.toFixed(1)}`);
            }
            
            // 2. 平均超期天數 > 7 天，額外扣 (平均天數 / 7) 分
            if (avgOverdueDays > 7) {
                const penaltyAvg = avgOverdueDays / 7;
                penalty += penaltyAvg;
                penaltyDetails.push(`平均超期>7天懲罰 = ${avgOverdueDays.toFixed(1)}/7 = ${penaltyAvg.toFixed(1)}`);
            }
            
            // 3. 超期率 > 30%，額外扣 2 分
            if (overdueRate > 0.3) {
                penalty += 2;
                penaltyDetails.push(`超期率>${(overdueRate*100).toFixed(0)}%>30%，扣2分`);
            }
            
            const header = `
                <div class="alert alert-info py-2 mb-2">
                    <strong>📊 貢獻度計算公式：</strong><br>
                    <hr class="my-1">
                    <strong>基礎分：</strong> High(${highCount})×3 + Medium(${medCount})×2 + Normal(${norCount})×1 = <strong>${baseScore}</strong><br>
                    <hr class="my-1">
                    <strong>超期統計：</strong><br>
                    • 超期任務數: ${overdueCount} 筆 / 總任務 ${nonPendingTasks.length} 筆 (超期率: ${(overdueRate*100).toFixed(1)}%)<br>
                    • 總超期天數: ${totalOverdueDays} 天<br>
                    • 平均超期: ${avgOverdueDays.toFixed(1)} 天/筆<br>
                    <hr class="my-1">
                    <strong>扣分計算：</strong><br>
                    ${penaltyDetails.length > 0 ? penaltyDetails.map(d => `• ${d}`).join('<br>') : '• 無扣分'}<br>
                    <strong>總扣分: <span class="text-danger">-${penalty.toFixed(1)}</span></strong><br>
                    <hr class="my-1">
                    <strong>最終得分: ${baseScore} - ${penalty.toFixed(1)} = <span class="text-success">${Math.max(0, baseScore - penalty).toFixed(1)}</span></strong>
                </div>
            `;
            
            showModal(`${name} 的任務（不含Pending） (${nonPendingTasks.length})`, header + modalTable(nonPendingTasks, 'memberNonPendingTable'));
        }
        
        // 顯示成員的超期任務
        function showMemberOverdueTasks(name) {
            if (!resultData) return;
            const m = resultData.members.find(x => x.name === name);
            if (!m) return;
            
            const overdueTasks = m.tasks.filter(t => t.overdue_days > 0 && t.task_status !== 'pending')
                .map(t => ({...t, owners_str: t.owners?.join('/')}))
                .sort((a, b) => b.overdue_days - a.overdue_days);
            
            const totalDays = overdueTasks.reduce((s, t) => s + (t.overdue_days || 0), 0);
            
            const header = `
                <div class="alert alert-danger py-2 mb-2">
                    <strong>${name} 的超期統計：</strong><br>
                    超期任務數: ${overdueTasks.length} 筆<br>
                    總超期天數: ${totalDays} 天<br>
                    平均超期: ${(overdueTasks.length > 0 ? totalDays / overdueTasks.length : 0).toFixed(1)} 天
                </div>
            `;
            
            showModal(`${name} 的超期任務 (${overdueTasks.length})`, header + modalTable(overdueTasks, 'memberOverdueTable'));
        }
        
        function showTaskDetail(title) {
            if (!resultData) return;
            const t = resultData.all_tasks.find(x => x.title === title);
            if (!t) return;
            const firstSeen = t.first_seen || t.mail_date || '-';
            const lastSeen = t.last_seen || t.mail_date || '-';
            const overdueDays = t.overdue_days || 0;
            showModal('任務詳情', `
                <div class="row">
                    <div class="col-md-6">
                        <p><strong>任務:</strong> ${t.title}</p>
                        <p><strong>模組:</strong> ${t.module || '-'}</p>
                        <p><strong>負責人:</strong> ${t.owners_str}</p>
                        <p><strong>優先級:</strong> <span class="badge badge-${t.priority}">${t.priority}</span></p>
                    </div>
                    <div class="col-md-6">
                        <p><strong>Due:</strong> <span class="${overdueDays > 0 ? 'text-overdue' : ''}">${t.due || '-'}</span></p>
                        <p><strong>狀態:</strong> <span class="badge badge-${t.task_status}">${statusLabels[t.task_status]}</span></p>
                        <p><strong>超期:</strong> <span class="${overdueDays > 0 ? 'text-overdue' : ''}">${overdueDays > 0 ? '+' + overdueDays + ' 天' : '無'}</span></p>
                        <p><strong>花費:</strong> ${t.days_spent || 0} 天 (${firstSeen} ~ ${lastSeen})</p>
                    </div>
                </div>
            `);
        }
        
        // Mail 預覽
        let currentMailData = null;
        
        async function showMailPreview(mailId, event) {
            if (event) event.stopPropagation();
            if (!mailId) { alert('此任務沒有關聯的 Mail'); return; }
            
            try {
                const r = await fetch(`/api/mail/${mailId}`);
                if (!r.ok) { alert('無法取得 Mail 內容'); return; }
                const mail = await r.json();
                currentMailData = mail;
                
                document.getElementById('mailSubject').textContent = mail.subject || '-';
                document.getElementById('mailDate').textContent = mail.date || '-';
                document.getElementById('mailTime').textContent = mail.time ? `(${mail.time})` : '';
                
                // 檢查是否有 HTML 內容
                const hasHtml = mail.html_body && mail.html_body.trim().length > 0;
                
                if (hasHtml) {
                    // 使用 HTML 內容
                    setMailView('html');
                    const iframe = document.getElementById('mailIframe');
                    iframe.srcdoc = mail.html_body;
                } else {
                    // 只有純文字，轉換為簡單 HTML
                    setMailView('html');
                    const iframe = document.getElementById('mailIframe');
                    const textAsHtml = `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>body{font-family:Segoe UI,Arial,sans-serif;font-size:14px;padding:20px;line-height:1.6;}</style></head><body><pre style="white-space:pre-wrap;font-family:inherit;">${escapeHtml(mail.body || '(無內容)')}</pre></body></html>`;
                    iframe.srcdoc = textAsHtml;
                }
                
                // 儲存純文字內容
                document.getElementById('mailBodyText').textContent = mail.body || '(無內容)';
                
                new bootstrap.Modal(document.getElementById('mailModal')).show();
            } catch (e) {
                alert('錯誤: ' + e);
            }
        }
        
        function escapeHtml(text) {
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }
        
        function setMailView(mode) {
            const htmlView = document.getElementById('mailBodyHtml');
            const textView = document.getElementById('mailBodyText');
            const btnHtml = document.getElementById('btnHtml');
            const btnText = document.getElementById('btnText');
            
            if (mode === 'html') {
                htmlView.style.display = 'block';
                textView.style.display = 'none';
                btnHtml.classList.add('active');
                btnText.classList.remove('active');
            } else {
                htmlView.style.display = 'none';
                textView.style.display = 'block';
                btnHtml.classList.remove('active');
                btnText.classList.add('active');
            }
        }

        function exportExcel() { window.location.href = '/api/excel'; }
        function exportHTML() { 
            if (!resultData) { alert('請先分析'); return; } 
            // 直接開啟，後端使用 LAST_DATA
            window.open('/api/export-html', '_blank');
        }

        document.getElementById('detailModal').addEventListener('hidden.bs.modal', () => { currentModal = null; });
    </script>
</body>
</html>
'''

# 簡化的 HTML Export
HTML_EXPORT = '''
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Task Report - {{ date }}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <style>
        body { background: #f5f7fa; padding: 15px; font-size: 14px; }
        .card { border: none; border-radius: 10px; box-shadow: 0 2px 12px rgba(0,0,0,0.08); margin-bottom: 12px; }
        .card-header { background: #2E75B6; color: white; padding: 8px 12px; display: flex; justify-content: space-between; align-items: center; }
        .stat-card { text-align: center; padding: 12px; cursor: pointer; transition: all 0.2s; }
        .stat-card:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(0,0,0,0.15); }
        .stat-number { font-size: 1.5rem; font-weight: bold; color: #2E75B6; }
        .stat-number.danger { color: #dc3545; }
        .stat-number.warning { color: #FFA500; }
        .stat-number.success { color: #28a745; }
        .stat-number.info { color: #17a2b8; }
        .badge-high { background: #FF6B6B !important; }
        .badge-medium { background: #FFE066 !important; color: #333 !important; }
        .badge-normal { background: #74C0FC !important; }
        .badge-completed { background: #28a745 !important; }
        .badge-pending { background: #FFA500 !important; }
        .badge-in_progress { background: #17a2b8 !important; }
        .data-table { border-collapse: collapse; width: 100%; }
        .data-table thead th { background: #4a4a4a !important; color: white !important; padding: 8px; border: 1px solid #666; cursor: pointer; }
        .data-table tbody td { padding: 6px; border: 1px solid #ddd; }
        .data-table tbody tr:nth-child(even) { background: #f9f9f9; }
        .data-table tbody tr:hover { background: #e9ecef; }
        .text-overdue { color: #dc3545 !important; font-weight: bold; }
        .row-pending { background: #fff8e6 !important; }
        .row-in_progress { background: #e6f7ff !important; }
        .row-overdue { border-left: 3px solid #dc3545 !important; }
        .chart-container { height: 200px; }
        .chart-select { width: auto; padding: 2px 8px; font-size: 0.75rem; }
        .rank-badge { display: inline-block; width: 24px; height: 24px; line-height: 24px; border-radius: 50%; text-align: center; font-weight: bold; color: white; font-size: 0.75rem; }
        .rank-1 { background: linear-gradient(135deg, #FFD700, #FFA500); }
        .rank-2 { background: linear-gradient(135deg, #C0C0C0, #A0A0A0); }
        .rank-3 { background: linear-gradient(135deg, #CD7F32, #8B4513); }
        .rank-other { background: #6c757d; }
        .table-toolbar { display: flex; gap: 8px; padding: 8px; background: #f8f9fa; align-items: center; flex-wrap: wrap; }
        .footer { text-align: center; padding: 15px; color: #999; font-size: 0.75rem; border-top: 1px solid #eee; margin-top: 20px; }
        @media print { .no-print { display: none !important; } }
    </style>
</head>
<body>
    <div class="container-fluid">
        <div class="text-center mb-3">
            <h3 style="color:#2E75B6"><i class="bi bi-clipboard-data me-2"></i>Task Dashboard Report</h3>
            <p class="text-muted mb-1">{{ date }} | 最後郵件: {{ data.last_mail_date }}</p>
            <button class="btn btn-primary btn-sm no-print" onclick="window.print()"><i class="bi bi-printer me-1"></i>列印</button>
        </div>

        <div class="row g-2 mb-2">
            <div class="col"><div class="card stat-card" onclick="filterByStatus('all')"><div class="stat-number">{{ data.total_tasks }}</div><div class="small">總任務</div></div></div>
            <div class="col"><div class="card stat-card" onclick="filterByStatus('pending')"><div class="stat-number warning">{{ data.pending_count }}</div><div class="small">Pending</div></div></div>
            <div class="col"><div class="card stat-card" onclick="filterByStatus('in_progress')"><div class="stat-number info">{{ data.in_progress_count }}</div><div class="small">進行中</div></div></div>
            <div class="col"><div class="card stat-card" onclick="filterByStatus('completed')"><div class="stat-number success">{{ data.completed_count }}</div><div class="small">已完成</div></div></div>
            <div class="col"><div class="card stat-card" onclick="filterByOverdue()"><div class="stat-number danger">{{ data.overdue_count }}</div><div class="small">超期</div></div></div>
        </div>

        <div class="row g-2 mb-2">
            <div class="col-md-3">
                <div class="card">
                    <div class="card-header">
                        <span><i class="bi bi-pie-chart me-1"></i>狀態分佈</span>
                        <select class="form-select chart-select no-print" id="c1Type" onchange="updateC1()">
                            <option value="doughnut">環形</option><option value="pie">圓餅</option><option value="bar">長條</option><option value="polarArea">極區</option>
                        </select>
                    </div>
                    <div class="card-body py-2"><div class="chart-container"><canvas id="c1"></canvas></div></div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card">
                    <div class="card-header">
                        <span><i class="bi bi-bar-chart me-1"></i>優先級</span>
                        <select class="form-select chart-select no-print" id="c2Type" onchange="updateC2()">
                            <option value="doughnut">環形</option><option value="pie">圓餅</option><option value="bar">長條</option><option value="polarArea">極區</option>
                        </select>
                    </div>
                    <div class="card-body py-2"><div class="chart-container"><canvas id="c2"></canvas></div></div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card">
                    <div class="card-header">
                        <span><i class="bi bi-exclamation-triangle me-1"></i>超期狀況</span>
                        <select class="form-select chart-select no-print" id="c3Type" onchange="updateC3()">
                            <option value="doughnut">環形</option><option value="pie">圓餅</option><option value="bar">長條</option><option value="polarArea">極區</option>
                        </select>
                    </div>
                    <div class="card-body py-2"><div class="chart-container"><canvas id="c3"></canvas></div></div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card">
                    <div class="card-header">
                        <span><i class="bi bi-person-exclamation me-1"></i>成員超期</span>
                        <select class="form-select chart-select no-print" id="c4Type" onchange="updateC4()">
                            <option value="stacked">堆疊</option><option value="bar">長條</option><option value="line">折線</option><option value="doughnut">環形</option>
                        </select>
                    </div>
                    <div class="card-body py-2"><div class="chart-container"><canvas id="c4"></canvas></div></div>
                </div>
            </div>
        </div>

        <div class="card mb-2">
            <div class="card-header"><span><i class="bi bi-list-task me-1"></i>任務列表</span></div>
            <div class="table-toolbar no-print">
                <input type="text" class="form-control form-control-sm" style="width:200px" placeholder="🔍 搜尋..." id="taskSearch" onkeyup="filterTasks()">
                <select class="form-select form-select-sm" style="width:100px" id="statusFilter" onchange="filterTasks()">
                    <option value="">全部狀態</option><option value="in_progress">進行中</option><option value="pending">Pending</option><option value="completed">已完成</option>
                </select>
                <select class="form-select form-select-sm" style="width:100px" id="priorityFilter" onchange="filterTasks()">
                    <option value="">全部優先級</option><option value="high">High</option><option value="medium">Medium</option><option value="normal">Normal</option>
                </select>
                <select class="form-select form-select-sm" style="width:80px" id="pageSize" onchange="renderTaskTable()">
                    <option value="50">50</option><option value="100">100</option><option value="200">200</option><option value="500">500</option>
                </select>
                <button class="btn btn-outline-secondary btn-sm" onclick="exportCSV()"><i class="bi bi-download"></i> CSV</button>
                <span class="ms-auto text-muted small" id="taskInfo"></span>
            </div>
            <div style="max-height:400px;overflow:auto">
                <table class="table table-sm data-table mb-0" id="taskTable">
                    <thead>
                        <tr>
                            <th onclick="sortTasks('last_seen')">Mail日期 ↕</th>
                            <th onclick="sortTasks('module')">模組 ↕</th>
                            <th onclick="sortTasks('title')">任務 ↕</th>
                            <th onclick="sortTasks('owners_str')">負責人 ↕</th>
                            <th onclick="sortTasks('priority')">優先級 ↕</th>
                            <th onclick="sortTasks('due')">Due ↕</th>
                            <th onclick="sortTasks('overdue_days')">超期 ↕</th>
                            <th onclick="sortTasks('task_status')">狀態 ↕</th>
                        </tr>
                    </thead>
                    <tbody id="taskBody"></tbody>
                </table>
            </div>
            <div class="d-flex justify-content-between align-items-center p-2 no-print">
                <button class="btn btn-sm btn-outline-secondary" onclick="prevPage()">上一頁</button>
                <span id="pageInfo" class="small"></span>
                <button class="btn btn-sm btn-outline-secondary" onclick="nextPage()">下一頁</button>
            </div>
        </div>

        <div class="row g-2">
            <div class="col-md-7">
                <div class="card">
                    <div class="card-header"><span><i class="bi bi-people me-1"></i>成員統計</span></div>
                    <div style="max-height:300px;overflow:auto">
                        <table class="table table-sm data-table mb-0">
                            <thead><tr><th>成員</th><th>總數</th><th>完成</th><th>進行</th><th>Pend</th><th>H</th><th>M</th><th>N</th></tr></thead>
                            <tbody>{% for m in data.members %}<tr style="cursor:pointer" onclick="filterByMember('{{ m.name }}')"><td><strong>{{ m.name }}</strong></td><td>{{ m.total }}</td><td>{{ m.completed }}</td><td>{{ m.in_progress }}</td><td>{{ m.pending }}</td><td>{{ m.high }}</td><td>{{ m.medium }}</td><td>{{ m.normal }}</td></tr>{% endfor %}</tbody>
                        </table>
                    </div>
                </div>
            </div>
            <div class="col-md-5">
                <div class="card">
                    <div class="card-header"><span><i class="bi bi-trophy me-1"></i>貢獻度排名</span></div>
                    <div style="max-height:300px;overflow:auto">
                        <table class="table table-sm data-table mb-0">
                            <thead><tr><th>#</th><th>成員</th><th>任務</th><th>基礎</th><th>超期</th><th>扣分</th><th>總分</th></tr></thead>
                            <tbody>{% for c in data.contribution %}<tr style="cursor:pointer" onclick="filterByMember('{{ c.name }}')"><td><span class="rank-badge {{ 'rank-' ~ loop.index if loop.index <= 3 else 'rank-other' }}">{{ loop.index }}</span></td><td>{{ c.name }}</td><td>{{ c.task_count }}</td><td>{{ c.base_score }}</td><td class="{{ 'text-overdue' if c.overdue_count > 0 else '' }}">{{ c.overdue_count }}</td><td class="{{ 'text-overdue' if c.overdue_penalty > 0 else '' }}">-{{ c.overdue_penalty }}</td><td><strong>{{ c.score }}</strong></td></tr>{% endfor %}</tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>

        <div class="footer">© 2025 Task Dashboard | Generated: {{ date }}</div>
    </div>

    <script>
        const statusLabels = { completed: '已完成', pending: 'Pending', in_progress: '進行中' };
        const allTasks = {{ data.all_tasks | tojson }};
        const contribution = {{ data.contribution | tojson }};
        let filteredTasks = [...allTasks];
        let currentPage = 0;
        let sortKey = 'last_seen', sortDir = -1;
        let chart1, chart2, chart3, chart4;

        function filterTasks() {
            const search = document.getElementById('taskSearch').value.toLowerCase();
            const status = document.getElementById('statusFilter').value;
            const priority = document.getElementById('priorityFilter').value;
            filteredTasks = allTasks.filter(t => {
                if (search && !JSON.stringify(t).toLowerCase().includes(search)) return false;
                if (status && t.task_status !== status) return false;
                if (priority && t.priority !== priority) return false;
                return true;
            });
            currentPage = 0;
            renderTaskTable();
        }

        function filterByStatus(status) {
            document.getElementById('statusFilter').value = status === 'all' ? '' : status;
            filterTasks();
        }

        function filterByOverdue() {
            filteredTasks = allTasks.filter(t => t.overdue_days > 0);
            currentPage = 0;
            renderTaskTable();
        }

        function filterByMember(name) {
            document.getElementById('taskSearch').value = name;
            filterTasks();
        }

        function sortTasks(key) {
            if (sortKey === key) sortDir *= -1; else { sortKey = key; sortDir = 1; }
            filteredTasks.sort((a, b) => {
                let va = a[key] || '', vb = b[key] || '';
                if (typeof va === 'number') return (va - vb) * sortDir;
                return String(va).localeCompare(String(vb)) * sortDir;
            });
            renderTaskTable();
        }

        function renderTaskTable() {
            const pageSize = parseInt(document.getElementById('pageSize').value);
            const start = currentPage * pageSize;
            const pageData = filteredTasks.slice(start, start + pageSize);
            
            document.getElementById('taskBody').innerHTML = pageData.map(t => `
                <tr class="row-${t.task_status} ${t.overdue_days > 0 ? 'row-overdue' : ''}">
                    <td>${t.last_seen || '-'}</td>
                    <td><span class="badge bg-secondary" style="font-size:0.6rem">${t.module || '-'}</span></td>
                    <td>${t.title}</td>
                    <td>${t.owners_str}</td>
                    <td><span class="badge badge-${t.priority}">${t.priority}</span></td>
                    <td class="${t.overdue_days > 0 ? 'text-overdue' : ''}">${t.due || '-'}</td>
                    <td class="${t.overdue_days > 0 ? 'text-overdue' : ''}">${t.overdue_days > 0 ? '+' + t.overdue_days + '天' : '-'}</td>
                    <td><span class="badge badge-${t.task_status}">${statusLabels[t.task_status]}</span></td>
                </tr>
            `).join('');
            
            const totalPages = Math.ceil(filteredTasks.length / pageSize) || 1;
            document.getElementById('pageInfo').textContent = `第 ${currentPage + 1}/${totalPages} 頁`;
            document.getElementById('taskInfo').textContent = `共 ${filteredTasks.length} 筆`;
        }

        function prevPage() { if (currentPage > 0) { currentPage--; renderTaskTable(); } }
        function nextPage() { 
            const pageSize = parseInt(document.getElementById('pageSize').value);
            if ((currentPage + 1) * pageSize < filteredTasks.length) { currentPage++; renderTaskTable(); }
        }

        function exportCSV() {
            const headers = ['Mail日期', '模組', '任務', '負責人', '優先級', 'Due', '超期天數', '狀態'];
            const rows = filteredTasks.map(t => [t.last_seen||'', t.module||'', t.title, t.owners_str, t.priority, t.due||'', t.overdue_days||0, statusLabels[t.task_status]]);
            let csv = [headers.join(','), ...rows.map(r => r.map(v => '"'+String(v).replace(/"/g,'""')+'"').join(','))].join('\\n');
            const blob = new Blob([csv], {type:'text/csv'});
            const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'tasks.csv'; a.click();
        }

        function updateC1() {
            const type = document.getElementById('c1Type').value;
            if (chart1) chart1.destroy();
            chart1 = new Chart(document.getElementById('c1'), {
                type: type, data: { labels: ['進行中', 'Pending', '已完成'], datasets: [{ data: [{{ data.in_progress_count }}, {{ data.pending_count }}, {{ data.completed_count }}], backgroundColor: ['#17a2b8', '#FFA500', '#28a745'] }] },
                options: { maintainAspectRatio: false, plugins: { legend: { display: type !== 'bar', position: 'right' } } }
            });
        }

        function updateC2() {
            const type = document.getElementById('c2Type').value;
            if (chart2) chart2.destroy();
            chart2 = new Chart(document.getElementById('c2'), {
                type: type, data: { labels: ['High', 'Medium', 'Normal'], datasets: [{ data: [{{ data.priority_counts.high }}, {{ data.priority_counts.medium }}, {{ data.priority_counts.normal }}], backgroundColor: ['#FF6B6B', '#FFE066', '#74C0FC'] }] },
                options: { maintainAspectRatio: false, plugins: { legend: { display: type !== 'bar', position: 'right' } } }
            });
        }

        function updateC3() {
            const type = document.getElementById('c3Type').value;
            if (chart3) chart3.destroy();
            chart3 = new Chart(document.getElementById('c3'), {
                type: type, data: { labels: ['超期', '未超期'], datasets: [{ data: [{{ data.overdue_count }}, {{ data.not_overdue_count }}], backgroundColor: ['#dc3545', '#28a745'] }] },
                options: { maintainAspectRatio: false, plugins: { legend: { display: type !== 'bar', position: 'right' } } }
            });
        }

        function updateC4() {
            const type = document.getElementById('c4Type').value;
            if (chart4) chart4.destroy();
            const overdueData = contribution.filter(c => c.overdue_days > 0).sort((a, b) => b.overdue_days - a.overdue_days).slice(0, 10);
            const labels = overdueData.map(c => c.name);
            
            if (type === 'stacked') {
                chart4 = new Chart(document.getElementById('c4'), {
                    type: 'bar', data: { labels, datasets: [
                        { label: '已完成超期', data: overdueData.map(c => c.completed_overdue_days || 0), backgroundColor: '#6c757d', stack: 's' },
                        { label: '未完成超期', data: overdueData.map(c => c.active_overdue_days || 0), backgroundColor: '#dc3545', stack: 's' }
                    ]}, options: { maintainAspectRatio: false, indexAxis: 'y', plugins: { legend: { display: true, position: 'top' } }, scales: { x: { stacked: true }, y: { stacked: true } } }
                });
            } else if (type === 'line') {
                chart4 = new Chart(document.getElementById('c4'), {
                    type: 'line', data: { labels, datasets: [
                        { label: '已完成超期', data: overdueData.map(c => c.completed_overdue_days || 0), borderColor: '#6c757d', backgroundColor: 'rgba(108,117,125,0.2)', fill: true },
                        { label: '未完成超期', data: overdueData.map(c => c.active_overdue_days || 0), borderColor: '#dc3545', backgroundColor: 'rgba(220,53,69,0.2)', fill: true }
                    ]}, options: { maintainAspectRatio: false, plugins: { legend: { display: true } } }
                });
            } else if (type === 'doughnut') {
                chart4 = new Chart(document.getElementById('c4'), {
                    type: 'doughnut', data: { labels, datasets: [{ data: overdueData.map(c => c.overdue_days), backgroundColor: ['#FF6B6B','#FFA500','#FFE066','#74C0FC','#69DB7C','#B197FC','#F783AC','#20C997','#ADB5BD','#868E96'] }] },
                    options: { maintainAspectRatio: false, plugins: { legend: { position: 'right' } } }
                });
            } else {
                const data = overdueData.map(c => c.overdue_days);
                const max = Math.max(...data);
                chart4 = new Chart(document.getElementById('c4'), {
                    type: 'bar', data: { labels, datasets: [{ data, backgroundColor: data.map(d => d >= max*0.7 ? '#dc3545' : d >= max*0.4 ? '#FFA500' : '#FFE066') }] },
                    options: { maintainAspectRatio: false, indexAxis: 'y', plugins: { legend: { display: false } } }
                });
            }
        }

        // 初始化
        updateC1(); updateC2(); updateC3(); updateC4();
        renderTaskTable();
    </script>
</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML, tree=FOLDER_TREE, fc=len(FOLDERS))

@app.route('/api/outlook', methods=['POST'])
def api_outlook():
    global LAST_RESULT, LAST_DATA, MAIL_CONTENTS
    MAIL_CONTENTS.clear()  # 清空之前的 mail 內容
    try:
        j = request.json
        exclude_middle_priority = j.get('exclude_middle_priority', True)
        exclude_after_5pm = j.get('exclude_after_5pm', True)
        
        msgs = get_messages(j['entry_id'], j['store_id'], j['start'], j['end'], exclude_after_5pm)
        parser = TaskParser(exclude_middle_priority=exclude_middle_priority)
        for m in msgs:
            parser.parse(m['subject'], m['body'], m['date'], m.get('time', ''), m.get('html_body', ''))
        stats = Stats()
        for t in parser.tasks:
            stats.add(t)
        LAST_RESULT = stats
        LAST_DATA = stats.summary()
        return jsonify(LAST_DATA)
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/upload', methods=['POST'])
def api_upload():
    global LAST_RESULT, LAST_DATA, MAIL_CONTENTS
    MAIL_CONTENTS.clear()  # 清空之前的 mail 內容
    if not HAS_EXTRACT_MSG:
        return jsonify({'error': 'extract-msg not installed'}), 500
    
    exclude_middle_priority = request.form.get('exclude_middle_priority', 'true').lower() == 'true'
    exclude_after_5pm = request.form.get('exclude_after_5pm', 'true').lower() == 'true'
    
    parser = TaskParser(exclude_middle_priority=exclude_middle_priority)
    for f in request.files.getlist('f'):
        if not f.filename.endswith('.msg'): continue
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.msg') as tmp:
                f.save(tmp.name)
                msg = extract_msg.Message(tmp.name)
                
                # 檢查時間
                mail_time = msg.date
                if exclude_after_5pm and mail_time and hasattr(mail_time, 'hour'):
                    if mail_time.hour >= 17:
                        os.unlink(tmp.name)
                        continue
                
                mail_date_str = mail_time.strftime("%Y-%m-%d") if mail_time else ""
                mail_time_str = mail_time.strftime("%H:%M") if mail_time else ""
                
                # 取得 HTML 內容
                html_body = ""
                try:
                    html_body = msg.htmlBody or ""
                except:
                    pass
                
                parser.parse(msg.subject or "", msg.body or "", mail_date_str, mail_time_str, html_body)
                os.unlink(tmp.name)
        except: pass
    stats = Stats()
    for t in parser.tasks:
        stats.add(t)
    LAST_RESULT = stats
    LAST_DATA = stats.summary()
    return jsonify(LAST_DATA)

@app.route('/api/excel')
def api_excel():
    if not LAST_RESULT:
        return jsonify({'error': 'No data'}), 400
    return send_file(LAST_RESULT.excel(), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f'task_report_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx')

@app.route('/api/mail/<mail_id>')
def api_mail(mail_id):
    """取得 mail 內容"""
    if mail_id in MAIL_CONTENTS:
        return jsonify(MAIL_CONTENTS[mail_id])
    return jsonify({'error': 'Mail not found'}), 404

@app.route('/api/export-html')
def api_export_html():
    global LAST_DATA
    
    if not LAST_DATA:
        return "請先執行分析", 400
    
    # 限制任務列表筆數（避免 HTML 太大）
    limited_data = dict(LAST_DATA)
    if len(limited_data.get('all_tasks', [])) > 200:
        limited_data['all_tasks'] = limited_data['all_tasks'][:200]
        limited_data['_truncated'] = True
    
    html = render_template_string(HTML_EXPORT, data=limited_data, date=datetime.now().strftime("%Y-%m-%d %H:%M"))
    return Response(html, mimetype='text/html', headers={'Content-Disposition': f'attachment; filename=task_report_{datetime.now().strftime("%Y%m%d_%H%M")}.html'})

if __name__ == '__main__':
    print("=" * 50)
    print("Task Dashboard v19")
    print("=" * 50)
    load_folders()
    print("開啟: http://127.0.0.1:5000")
    print("=" * 50)
    from werkzeug.serving import run_simple
    run_simple('127.0.0.1', 5000, app, use_reloader=False, threaded=False)
