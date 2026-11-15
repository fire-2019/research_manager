#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
科研项目与每日工作管理系统 (Windows 桌面本地版)
================================================
纯 Python 标准库实现（零第三方依赖），使用 SQLite 本地持久化存储，
启动后自动以本地 Web 应用 / Edge 应用窗口形式运行。

运行方式 (Windows):
    python research_manager.py
"""

import base64
import datetime
import json
import mimetypes
import os
import platform
import re
import shutil
import socket
import sqlite3
import subprocess
import threading
import urllib.parse
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "research_data"
EVIDENCE_DIR = DATA_DIR / "evidence"
DB_PATH = DATA_DIR / "research_pm.db"


def ensure_dirs():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    EVIDENCE_DIR.mkdir(parents=True, exist_ok=True)


def get_conn():
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def ensure_column(conn, table, column, col_def):
    cols = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()}
    if column not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {col_def}")


def init_db():
    ensure_dirs()
    with get_conn() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                category TEXT DEFAULT '科研项目',
                status TEXT DEFAULT 'active',
                objectives TEXT DEFAULT '',
                significance TEXT DEFAULT '',
                progress_mode TEXT DEFAULT 'auto',
                manual_progress INTEGER DEFAULT 0,
                start_date TEXT DEFAULT '',
                target_date TEXT DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS project_steps (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                step_order INTEGER DEFAULT 1,
                title TEXT NOT NULL,
                description TEXT DEFAULT '',
                status TEXT DEFAULT 'todo',
                progress INTEGER DEFAULT 0,
                due_date TEXT DEFAULT '',
                subtasks_json TEXT DEFAULT '[]',
                created_at TEXT NOT NULL,
                FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS step_evidences (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                step_id INTEGER NOT NULL,
                project_id INTEGER NOT NULL,
                title TEXT NOT NULL,
                evidence_type TEXT DEFAULT 'data_note',
                file_rel_path TEXT DEFAULT '',
                original_filename TEXT DEFAULT '',
                local_path TEXT DEFAULT '',
                url TEXT DEFAULT '',
                data_summary TEXT DEFAULT '',
                created_at TEXT NOT NULL,
                FOREIGN KEY (step_id) REFERENCES project_steps(id) ON DELETE CASCADE,
                FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS daily_tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                task_date TEXT NOT NULL,
                project_id INTEGER,
                title TEXT NOT NULL,
                priority TEXT DEFAULT 'medium',
                estimated_hours REAL DEFAULT 1.0,
                notes TEXT DEFAULT '',
                subtasks_json TEXT DEFAULT '[]',
                is_completed INTEGER DEFAULT 0,
                completed_at TEXT DEFAULT '',
                created_at TEXT NOT NULL,
                FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE SET NULL
            );

            CREATE TABLE IF NOT EXISTS feedback_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                module TEXT DEFAULT '科研项目详情',
                feedback_type TEXT DEFAULT 'ux',
                priority TEXT DEFAULT 'medium',
                status TEXT DEFAULT 'open',
                description TEXT DEFAULT '',
                resolution_note TEXT DEFAULT '',
                target_version TEXT DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                resolved_at TEXT DEFAULT ''
            );
            """
        )
        # 兼容旧版数据库平滑迁移
        ensure_column(conn, "project_steps", "subtasks_json", "TEXT DEFAULT '[]'")
        ensure_column(conn, "daily_tasks", "subtasks_json", "TEXT DEFAULT '[]'")
        conn.commit()


def now_iso():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today_str():
    return datetime.date.today().strftime("%Y-%m-%d")


def parse_subtasks(raw):
    if isinstance(raw, list):
        items = raw
    else:
        try:
            items = json.loads(raw or "[]")
        except Exception:
            items = []
    cleaned = []
    for it in items:
        if isinstance(it, dict):
            text = str(it.get("text", "")).strip()
            if text:
                cleaned.append({"text": text, "done": bool(it.get("done", False))})
    return cleaned


def compute_project_progress(project_row, steps):
    if project_row["progress_mode"] == "manual":
        return max(0, min(100, int(project_row["manual_progress"] or 0)))
    if not steps:
        return 100 if project_row["status"] == "completed" else 0
    total = 0
    for s in steps:
        if s["status"] == "done":
            total += 100
        elif s["status"] == "todo":
            total += 0
        else:
            p = int(s["progress"] or 0)
            total += p if p > 0 else 50
    return round(total / len(steps))


def fetch_all_projects_summary():
    with get_conn() as conn:
        proj_rows = conn.execute(
            "SELECT * FROM projects ORDER BY CASE status "
            "WHEN 'active' THEN 1 WHEN 'planning' THEN 2 WHEN 'paused' THEN 3 ELSE 4 END, updated_at DESC"
        ).fetchall()
        step_rows = conn.execute(
            "SELECT * FROM project_steps ORDER BY step_order ASC, id ASC"
        ).fetchall()
        ev_counts = conn.execute(
            "SELECT project_id, COUNT(*) AS cnt FROM step_evidences GROUP BY project_id"
        ).fetchall()

    steps_by_proj = {}
    for s in step_rows:
        steps_by_proj.setdefault(s["project_id"], []).append(dict(s))

    ev_map = {r["project_id"]: r["cnt"] for r in ev_counts}
    result = []
    for p in proj_rows:
        d = dict(p)
        p_steps = steps_by_proj.get(p["id"], [])
        d["computed_progress"] = compute_project_progress(p, p_steps)
        d["step_count"] = len(p_steps)
        d["done_step_count"] = sum(1 for s in p_steps if s["status"] == "done")
        d["evidence_count"] = ev_map.get(p["id"], 0)
        result.append(d)
    return result


def fetch_project_detail(project_id):
    with get_conn() as conn:
        proj = conn.execute("SELECT * FROM projects WHERE id = ?", (project_id,)).fetchone()
        if not proj:
            return None
        steps = []
        for r in conn.execute(
            "SELECT * FROM project_steps WHERE project_id = ? ORDER BY step_order ASC, id ASC",
            (project_id,),
        ).fetchall():
            sd = dict(r)
            sd["subtasks"] = parse_subtasks(sd.get("subtasks_json"))
            steps.append(sd)

        evidences = [
            dict(r)
            for r in conn.execute(
                "SELECT * FROM step_evidences WHERE project_id = ? ORDER BY created_at DESC, id DESC",
                (project_id,),
            ).fetchall()
        ]

    ev_by_step = {}
    for ev in evidences:
        ev_by_step.setdefault(ev["step_id"], []).append(ev)

    for s in steps:
        s["evidences"] = ev_by_step.get(s["id"], [])

    d = dict(proj)
    d["computed_progress"] = compute_project_progress(proj, steps)
    d["steps"] = steps
    return d


def sanitize_filename(name):
    name = os.path.basename(name or "file")
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    return name[:120] or "evidence_file"


def seed_sample_data():
    ts = now_iso()
    today = today_str()
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO projects (title, category, status, objectives, significance, progress_mode, manual_progress, start_date, target_date, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "UAV LiDAR 林下点云多视角赋色优化研究",
                "学术论文 / ISPRS",
                "active",
                "1. 量化多视角影像观测冗余度与点云跨视角颜色不一致性的关系；\n2. 提出基于几何可见性与辐射一致性的自适应颜色聚合方法；\n3. 在7个样地航飞块（2.61亿点、466张影像）上完成留一视角验证。",
                "解决复杂森林冠层遮挡导致的无人机点云赋色伪影问题，为单木分割、树种识别与森林生物量反演提供高保真三维真彩色点云数据基础。",
                "auto",
                0,
                "2026-09-01",
                "2026-12-30",
                ts,
                ts,
            ),
        )
        p1_id = cur.lastrowid

        steps_p1 = [
            (
                1,
                "数据采集与航飞预处理",
                "完成外业无人机 LiDAR 与 RGB 同步采集，整理航带与 POS 数据。",
                "done",
                100,
                "2026-09-20",
                [{"text": "完成点云航带拼接与去噪", "done": True}, {"text": "完成POS与影像时间戳对齐", "done": True}],
            ),
            (
                2,
                "多视角可见性与颜色聚合算法实现",
                "实现基于射线追踪/Z-buffer 的遮挡剔除与多视角颜色加权聚合。",
                "done",
                100,
                "2026-10-18",
                [{"text": "实现Z-buffer遮挡检测", "done": True}, {"text": "实现多视角加权聚合", "done": True}],
            ),
            (
                3,
                "跨视角颜色一致性对比实验",
                "在7个样地航飞块上运行消融实验与基线对比。",
                "in_progress",
                67,
                "2026-11-25",
                [
                    {"text": "完成样地1-5基线对比", "done": True},
                    {"text": "完成样地6-7消融实验", "done": True},
                    {"text": "汇总PSNR与ΔE2000统计表", "done": False},
                ],
            ),
            (
                4,
                "论文撰写与图表规范化排版",
                "按 ISPRS 期刊要求绘制 180mm 宽图件并完成 LaTeX 初稿。",
                "todo",
                0,
                "2026-12-20",
                [],
            ),
        ]
        step_ids = []
        for order, title, desc, st, prog, due, subs in steps_p1:
            cur.execute(
                """
                INSERT INTO project_steps (project_id, step_order, title, description, status, progress, due_date, subtasks_json, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (p1_id, order, title, desc, st, prog, due, json.dumps(subs, ensure_ascii=False), ts),
            )
            step_ids.append(cur.lastrowid)

        cur.execute(
            """
            INSERT INTO step_evidences (step_id, project_id, title, evidence_type, local_path, data_summary, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                step_ids[0],
                p1_id,
                "外业样地与航飞数据清单",
                "data_note",
                "",
                "共整理 7 个核心样地-航飞单元：累计 261.1M LiDAR 点，466 张高分辨率 RGB 影像；补充采集 298.6M 点 / 507 帧混交林数据。",
                ts,
            ),
        )
        cur.execute(
            """
            INSERT INTO step_evidences (step_id, project_id, title, evidence_type, local_path, data_summary, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                step_ids[2],
                p1_id,
                "留一视角评估指标初版汇总",
                "data_note",
                "",
                "PSNR 较最近视角基线提升 +2.4 dB，跨视角色差 ΔE2000 下降 31.5%（待补充第 6、7 航飞块最终统计）。",
                ts,
            ),
        )

        cur.execute(
            """
            INSERT INTO projects (title, category, status, objectives, significance, progress_mode, manual_progress, start_date, target_date, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "困难立地胁迫诊断与引水造林协同决策项目",
                "国际合作基金项目",
                "planning",
                "构建两年期四季度里程碑实施方案，明确多智能体协同决策架构、数据共享合规机制与知识产权分配。",
                "支撑干旱/困难立地条件下的精准引水造林工程，形成可复制的国际联合研究与技术示范方案。",
                "auto",
                0,
                "2026-11-01",
                "2028-10-31",
                ts,
                ts,
            ),
        )
        p2_id = cur.lastrowid
        cur.execute(
            """
            INSERT INTO project_steps (project_id, step_order, title, description, status, progress, due_date, subtasks_json, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                p2_id,
                1,
                "申报书核心章节定稿与甘特图绘制",
                "完成组织管理、保障措施、知识产权与风险分析章节。",
                "in_progress",
                80,
                "2026-11-20",
                json.dumps([], ensure_ascii=False),
                ts,
            ),
        )

        cur.execute(
            """
            INSERT INTO daily_tasks (task_date, project_id, title, priority, estimated_hours, notes, subtasks_json, is_completed, completed_at, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                today,
                p1_id,
                "整理第 6、7 航飞块跨视角色差统计表",
                "high",
                2.0,
                "输出 CSV 并核对异常值",
                json.dumps(
                    [
                        {"text": "导出第6航飞块CSV", "done": True},
                        {"text": "导出第7航飞块CSV", "done": False},
                        {"text": "核对异常值并生成对比图", "done": False},
                    ],
                    ensure_ascii=False,
                ),
                0,
                "",
                ts,
            ),
        )
        cur.execute(
            """
            INSERT INTO daily_tasks (task_date, project_id, title, priority, estimated_hours, notes, subtasks_json, is_completed, completed_at, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (today, p2_id, "核对两年期季度甘特图配色与节点文字", "medium", 1.0, "确保与申报书文本完全一致", "[]", 1, ts, ts),
        )
        cur.execute(
            """
            INSERT INTO daily_tasks (task_date, project_id, title, priority, estimated_hours, notes, subtasks_json, is_completed, completed_at, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (today, None, "跟踪 3D Gaussian Splatting 林业重建本周新论文", "low", 0.5, "重点关注遮挡处理与点云融合方向", "[]", 0, "", ts),
        )
        conn.commit()


HTML_PAGE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>科研项目与每日工作管理系统</title>
<style>
  :root {
    --bg: #f1f5f9;
    --surface: #ffffff;
    --sidebar: #0f172a;
    --sidebar-hover: #1e293b;
    --text: #0f172a;
    --muted: #64748b;
    --border: #e2e8f0;
    --primary: #2563eb;
    --primary-hover: #1d4ed8;
    --primary-soft: #eff6ff;
    --success: #10b981;
    --success-soft: #ecfdf5;
    --warning: #f59e0b;
    --warning-soft: #fffbeb;
    --danger: #ef4444;
    --danger-soft: #fef2f2;
    --purple: #7c3aed;
    --purple-soft: #f5f3ff;
    --radius: 10px;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    font-family: "Microsoft YaHei UI", "PingFang SC", "Segoe UI", system-ui, sans-serif;
    background: var(--bg);
    color: var(--text);
    display: flex;
    height: 100vh;
    overflow: hidden;
  }
  /* Sidebar */
  .sidebar {
    width: 240px;
    background: var(--sidebar);
    color: #e2e8f0;
    display: flex;
    flex-direction: column;
    flex-shrink: 0;
  }
  .brand {
    padding: 20px 18px;
    border-bottom: 1px solid rgba(255,255,255,0.08);
  }
  .brand h1 {
    margin: 0;
    font-size: 16px;
    font-weight: 700;
    color: #fff;
    letter-spacing: 0.3px;
  }
  .brand p {
    margin: 4px 0 0;
    font-size: 12px;
    color: #94a3b8;
  }
  .nav {
    padding: 12px 10px;
    flex: 1;
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  .nav-btn {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    width: 100%;
    padding: 10px 14px;
    border: none;
    border-radius: 8px;
    background: transparent;
    color: #cbd5e1;
    font-size: 14px;
    cursor: pointer;
    text-align: left;
    transition: all 0.15s;
  }
  .nav-btn:hover { background: var(--sidebar-hover); color: #fff; }
  .nav-btn.active { background: var(--primary); color: #fff; font-weight: 600; }
  .nav-badge {
    background: rgba(245, 158, 11, 0.25);
    color: #fcd34d;
    font-size: 11px;
    padding: 1px 7px;
    border-radius: 999px;
    font-weight: 700;
  }
  .nav-btn.active .nav-badge {
    background: rgba(255,255,255,0.22);
    color: #fff;
  }
  .sidebar-footer {
    padding: 14px;
    border-top: 1px solid rgba(255,255,255,0.08);
    display: flex;
    flex-direction: column;
    gap: 8px;
  }
  .sidebar-footer button {
    width: 100%;
    padding: 8px 10px;
    border-radius: 6px;
    border: 1px solid rgba(255,255,255,0.16);
    background: rgba(255,255,255,0.05);
    color: #e2e8f0;
    font-size: 12px;
    cursor: pointer;
  }
  .sidebar-footer button:hover { background: rgba(255,255,255,0.12); }

  /* Main Area */
  .main {
    flex: 1;
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }
  .topbar {
    height: 58px;
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    padding: 0 24px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-shrink: 0;
  }
  .topbar-title { font-size: 17px; font-weight: 700; }
  .topbar-actions { display: flex; gap: 10px; align-items: center; }

  .content {
    flex: 1;
    overflow-y: auto;
    padding: 22px 24px;
  }

  /* Common UI */
  .btn {
    padding: 8px 14px;
    border-radius: 7px;
    border: 1px solid var(--border);
    background: var(--surface);
    color: var(--text);
    font-size: 13px;
    font-weight: 500;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 6px;
    transition: all 0.15s;
  }
  .btn:hover { background: #f8fafc; }
  .btn-primary { background: var(--primary); color: #fff; border-color: var(--primary); }
  .btn-primary:hover { background: var(--primary-hover); }
  .btn-danger { color: var(--danger); border-color: #fecaca; background: #fff; }
  .btn-danger:hover { background: var(--danger-soft); }
  .btn-sm { padding: 4px 10px; font-size: 12px; }

  .card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 18px;
    box-shadow: 0 1px 2px rgba(15,23,42,0.03);
  }
  .grid-4 {
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 14px;
    margin-bottom: 18px;
  }
  .grid-2 {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 16px;
  }
  .stat-label { font-size: 12px; color: var(--muted); margin-bottom: 6px; }
  .stat-val { font-size: 26px; font-weight: 700; }
  .stat-sub { font-size: 12px; color: var(--muted); margin-top: 4px; }

  .badge {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 600;
  }
  .badge-active { background: var(--primary-soft); color: var(--primary); }
  .badge-done { background: var(--success-soft); color: #059669; }
  .badge-plan { background: #f1f5f9; color: #475569; }
  .badge-paused { background: var(--warning-soft); color: #b45309; }
  .badge-high { background: var(--danger-soft); color: #dc2626; }
  .badge-medium { background: var(--warning-soft); color: #d97706; }
  .badge-low { background: #f1f5f9; color: #475569; }
  .badge-purple { background: var(--purple-soft); color: var(--purple); }

  .progress-wrap {
    width: 100%;
    height: 8px;
    background: #e2e8f0;
    border-radius: 999px;
    overflow: hidden;
  }
  .progress-bar {
    height: 100%;
    background: var(--primary);
    border-radius: 999px;
    transition: width 0.25s ease;
  }
  .progress-bar.done { background: var(--success); }

  /* Click-to-edit styling */
  .editable {
    cursor: text;
    border-radius: 6px;
    padding: 2px 5px;
    margin: -2px -5px;
    border: 1px dashed transparent;
    transition: background 0.15s, border-color 0.15s;
  }
  .editable:hover {
    background: #f8fafc;
    border-color: #cbd5e1;
  }
  .inline-edit-box {
    width: 100%;
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
  .inline-edit-actions {
    display: flex;
    gap: 6px;
    justify-content: flex-end;
  }

  /* Subtasks checklist */
  .subtasks-box {
    margin-top: 8px;
    padding: 8px 10px;
    background: #f8fafc;
    border: 1px solid var(--border);
    border-radius: 7px;
  }
  .subtask-row {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 3px 0;
    font-size: 12.5px;
  }
  .subtask-row.done span.subtask-text {
    text-decoration: line-through;
    color: var(--muted);
  }
  .subtask-add-row {
    display: flex;
    gap: 6px;
    margin-top: 6px;
  }
  .subtask-add-row input {
    padding: 4px 8px;
    font-size: 12px;
  }

  /* Projects Split Layout */
  .proj-layout {
    display: grid;
    grid-template-columns: 310px 1fr;
    gap: 18px;
    align-items: start;
  }
  .proj-list-item {
    padding: 13px 14px;
    border: 1px solid var(--border);
    border-radius: 9px;
    background: var(--surface);
    cursor: pointer;
    margin-bottom: 10px;
    transition: all 0.15s;
  }
  .proj-list-item:hover { border-color: #93c5fd; }
  .proj-list-item.selected {
    border-color: var(--primary);
    background: #f8fbff;
    box-shadow: 0 0 0 2px rgba(37,99,235,0.1);
  }
  .section-title {
    font-size: 14px;
    font-weight: 700;
    margin: 0 0 10px;
    display: flex;
    align-items: center;
    justify-content: space-between;
  }
  .pre-text {
    white-space: pre-wrap;
    font-size: 13.5px;
    line-height: 1.65;
    color: #334155;
    margin: 0;
  }

  /* Step Card */
  .step-card {
    border: 1px solid var(--border);
    border-radius: 9px;
    padding: 14px 16px;
    margin-bottom: 12px;
    background: #fcfdff;
  }
  .step-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
    flex-wrap: wrap;
  }
  .evidence-list {
    margin-top: 12px;
    padding-top: 10px;
    border-top: 1px dashed var(--border);
    display: flex;
    flex-direction: column;
    gap: 8px;
  }
  .evidence-item {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 7px;
    padding: 10px 12px;
    font-size: 13px;
  }
  .evidence-img-preview {
    margin-top: 8px;
    max-width: 260px;
    max-height: 160px;
    border-radius: 6px;
    border: 1px solid var(--border);
    cursor: pointer;
  }

  /* Daily Tasks */
  .task-item {
    display: flex;
    align-items: flex-start;
    gap: 10px;
    padding: 11px 12px;
    border: 1px solid var(--border);
    border-radius: 8px;
    margin-bottom: 8px;
    background: var(--surface);
  }
  .task-item.completed {
    background: #f8fafc;
    opacity: 0.85;
  }
  .task-item.completed .task-title {
    text-decoration: line-through;
    color: var(--muted);
  }
  .task-check {
    width: 18px;
    height: 18px;
    margin-top: 2px;
    cursor: pointer;
    accent-color: var(--primary);
  }

  /* Feedback Item */
  .fb-item {
    border: 1px solid var(--border);
    border-radius: 9px;
    padding: 14px 16px;
    margin-bottom: 10px;
    background: var(--surface);
  }
  .fb-item.resolved {
    background: #f8fafc;
    opacity: 0.82;
  }

  /* Forms & Modals */
  .modal-mask {
    position: fixed;
    inset: 0;
    background: rgba(15,23,42,0.45);
    display: none;
    align-items: center;
    justify-content: center;
    z-index: 100;
  }
  .modal-mask.open { display: flex; }
  .modal {
    background: var(--surface);
    width: 560px;
    max-width: 92vw;
    max-height: 88vh;
    overflow-y: auto;
    border-radius: 12px;
    padding: 20px 22px;
    box-shadow: 0 20px 40px rgba(0,0,0,0.18);
  }
  .form-group { margin-bottom: 12px; }
  .form-group label {
    display: block;
    font-size: 12.5px;
    font-weight: 600;
    margin-bottom: 5px;
    color: #334155;
  }
  input[type="text"], input[type="date"], input[type="number"], select, textarea {
    width: 100%;
    padding: 8px 10px;
    border: 1px solid #cbd5e1;
    border-radius: 7px;
    font-size: 13px;
    font-family: inherit;
    background: #fff;
  }
  textarea { resize: vertical; min-height: 76px; }
  .empty-state {
    text-align: center;
    padding: 40px 20px;
    color: var(--muted);
    font-size: 14px;
  }
  .toast {
    position: fixed;
    bottom: 22px;
    right: 24px;
    background: #0f172a;
    color: #fff;
    padding: 10px 16px;
    border-radius: 8px;
    font-size: 13px;
    opacity: 0;
    pointer-events: none;
    transition: opacity 0.2s;
    z-index: 200;
  }
  .toast.show { opacity: 1; }
  @media (max-width: 960px) {
    .grid-4, .grid-2, .proj-layout { grid-template-columns: 1fr; }
    .sidebar { width: 200px; }
  }
</style>
</head>
<body>

<aside class="sidebar">
  <div class="brand">
    <h1>科研项目工作台</h1>
    <p>Windows 本地科研管理与证据库</p>
  </div>
  <nav class="nav">
    <button class="nav-btn active" data-tab="dashboard" onclick="switchTab('dashboard')"><span>📊 全部项目总览</span></button>
    <button class="nav-btn" data-tab="projects" onclick="switchTab('projects')"><span>🔬 科研项目详情</span></button>
    <button class="nav-btn" data-tab="daily" onclick="switchTab('daily')"><span>✅ 每日工作记录</span></button>
    <button class="nav-btn" data-tab="feedback" onclick="switchTab('feedback')">
      <span>🛠️ 意见改进</span>
      <span class="nav-badge" id="fb-open-badge" style="display:none;">0</span>
    </button>
  </nav>
  <div class="sidebar-footer">
    <button onclick="seedDemo()">✨ 导入示例科研项目</button>
    <button onclick="exportBackup()">💾 导出全库 JSON 备份</button>
    <button onclick="openDataFolder()">📂 打开本地数据目录</button>
  </div>
</aside>

<main class="main">
  <header class="topbar">
    <div class="topbar-title" id="topbar-title">全部项目总览</div>
    <div class="topbar-actions">
      <button class="btn" onclick="openFeedbackModal()">💡 快速提意见</button>
      <button class="btn btn-primary" onclick="openProjectModal()">＋ 新建科研项目</button>
    </div>
  </header>
  <div class="content" id="app-content"></div>
</main>

<!-- Modal Container -->
<div class="modal-mask" id="modal-mask" onclick="if(event.target===this) closeModal()">
  <div class="modal" id="modal-body"></div>
</div>
<div class="toast" id="toast"></div>

<script>
let state = {
  tab: 'dashboard',
  projects: [],
  selectedProjectId: null,
  selectedProjectDetail: null,
  selectedDate: new Date().toISOString().slice(0, 10),
  dailyTasks: [],
  feedbacks: [],
  filterStatus: 'all',
  fbFilterStatus: 'all',
  fbFilterType: 'all',
  fbSearch: ''
};

const STATUS_MAP = {
  active: { label: '进行中', cls: 'badge-active' },
  planning: { label: '规划中', cls: 'badge-plan' },
  completed: { label: '已结题', cls: 'badge-done' },
  paused: { label: '已暂停', cls: 'badge-paused' }
};

const STEP_STATUS_MAP = {
  todo: { label: '未开始', cls: 'badge-plan' },
  in_progress: { label: '进行中', cls: 'badge-active' },
  done: { label: '已完成', cls: 'badge-done' }
};

const PRIORITY_MAP = {
  high: { label: '高优先级', cls: 'badge-high' },
  medium: { label: '中优先级', cls: 'badge-medium' },
  low: { label: '低优先级', cls: 'badge-low' }
};

const FB_TYPE_MAP = {
  bug: { label: '问题缺陷', cls: 'badge-high' },
  feature: { label: '新功能需求', cls: 'badge-active' },
  ux: { label: '体验优化', cls: 'badge-purple' },
  perf: { label: '性能与稳定', cls: 'badge-medium' }
};

const FB_STATUS_MAP = {
  open: { label: '待处理', cls: 'badge-paused' },
  planned: { label: '已排期', cls: 'badge-plan' },
  in_progress: { label: '处理中', cls: 'badge-active' },
  resolved: { label: '已解决', cls: 'badge-done' },
  wontfix: { label: '暂不处理', cls: 'badge-low' }
};

const FB_MODULES = ['全部项目总览', '科研项目详情', '数据与证据', '每日工作记录', '界面与交互', '导出与备份', '其他'];

function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function showToast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), 2400);
}

async function api(url, method = 'GET', body = null) {
  const opts = { method, headers: {} };
  if (body !== null) {
    opts.headers['Content-Type'] = 'application/json';
    opts.body = JSON.stringify(body);
  }
  const res = await fetch(url, opts);
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || '请求失败');
  return data;
}

/* Generic Click-to-Edit Helper */
function startInlineEdit(el, currentVal, multiline, onSave) {
  if (el.dataset.editing === '1') return;
  el.dataset.editing = '1';
  const origHtml = el.innerHTML;
  el.innerHTML = '';

  const box = document.createElement('div');
  box.className = 'inline-edit-box';
  box.onclick = (e) => e.stopPropagation();

  const input = document.createElement(multiline ? 'textarea' : 'input');
  if (!multiline) input.type = 'text';
  input.value = currentVal || '';
  if (multiline) input.style.minHeight = '84px';
  box.appendChild(input);

  const actions = document.createElement('div');
  actions.className = 'inline-edit-actions';
  const btnCancel = document.createElement('button');
  btnCancel.type = 'button';
  btnCancel.className = 'btn btn-sm';
  btnCancel.textContent = '取消';
  const btnSave = document.createElement('button');
  btnSave.type = 'button';
  btnSave.className = 'btn btn-sm btn-primary';
  btnSave.textContent = '保存';
  actions.appendChild(btnCancel);
  actions.appendChild(btnSave);
  box.appendChild(actions);

  el.appendChild(box);
  input.focus();
  if (!multiline) input.select();

  const finish = async (save) => {
    if (el.dataset.editing !== '1') return;
    el.dataset.editing = '0';
    if (save) {
      const val = input.value;
      if (val !== (currentVal || '')) {
        await onSave(val);
        return;
      }
    }
    el.innerHTML = origHtml;
  };

  btnCancel.onclick = () => finish(false);
  btnSave.onclick = () => finish(true);
  input.onkeydown = (e) => {
    if (!multiline && e.key === 'Enter') { e.preventDefault(); finish(true); }
    if (multiline && e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); finish(true); }
    if (e.key === 'Escape') { e.preventDefault(); finish(false); }
  };
}

async function loadAll() {
  const [projects, dailyTasks, feedbacks] = await Promise.all([
    api('/api/projects'),
    api('/api/daily-tasks?date=' + encodeURIComponent(state.selectedDate)),
    api('/api/feedbacks')
  ]);
  state.projects = projects;
  state.dailyTasks = dailyTasks;
  state.feedbacks = feedbacks;

  const openCount = feedbacks.filter(f => f.status === 'open' || f.status === 'in_progress' || f.status === 'planned').length;
  const badge = document.getElementById('fb-open-badge');
  if (openCount > 0) {
    badge.textContent = openCount;
    badge.style.display = 'inline-block';
  } else {
    badge.style.display = 'none';
  }

  if (state.selectedProjectId) {
    try {
      state.selectedProjectDetail = await api('/api/projects/' + state.selectedProjectId);
    } catch {
      state.selectedProjectId = null;
      state.selectedProjectDetail = null;
    }
  }
  if (!state.selectedProjectId && state.projects.length > 0) {
    state.selectedProjectId = state.projects[0].id;
    state.selectedProjectDetail = await api('/api/projects/' + state.selectedProjectId);
  }
  render();
}

function switchTab(tab) {
  state.tab = tab;
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  const titles = {
    dashboard: '全部项目总览',
    projects: '科研项目详情与数据证据',
    daily: '每日工作记录与任务管理',
    feedback: '系统意见与改进记录'
  };
  document.getElementById('topbar-title').textContent = titles[tab];
  render();
}

function render() {
  const el = document.getElementById('app-content');
  if (state.tab === 'dashboard') renderDashboard(el);
  else if (state.tab === 'projects') renderProjects(el);
  else if (state.tab === 'daily') renderDaily(el);
  else renderFeedback(el);
}

function renderDashboard(el) {
  const total = state.projects.length;
  const active = state.projects.filter(p => p.status === 'active').length;
  const completed = state.projects.filter(p => p.status === 'completed').length;
  const avgProg = total ? Math.round(state.projects.reduce((a, p) => a + p.computed_progress, 0) / total) : 0;
  const todayPending = state.dailyTasks.filter(t => !t.is_completed).length;
  const todayDone = state.dailyTasks.filter(t => t.is_completed).length;

  let filtered = state.projects;
  if (state.filterStatus !== 'all') filtered = filtered.filter(p => p.status === state.filterStatus);

  const projectRows = filtered.length ? filtered.map(p => {
    const st = STATUS_MAP[p.status] || STATUS_MAP.active;
    return `
      <div style="padding:14px 0; border-bottom:1px solid var(--border); cursor:pointer;" onclick="selectProjectAndGo(${p.id})">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
          <div>
            <span style="font-weight:700; font-size:15px;">${esc(p.title)}</span>
            <span class="badge ${st.cls}" style="margin-left:8px;">${st.label}</span>
            <span style="font-size:12px; color:var(--muted); margin-left:8px;">${esc(p.category)}</span>
          </div>
          <div style="font-size:13px; font-weight:700; color:${p.computed_progress===100?'var(--success)':'var(--primary)'};">
            ${p.computed_progress}%
          </div>
        </div>
        <div class="progress-wrap">
          <div class="progress-bar ${p.computed_progress===100?'done':''}" style="width:${p.computed_progress}%"></div>
        </div>
        <div style="display:flex; gap:16px; margin-top:7px; font-size:12px; color:var(--muted);">
          <span>步骤完成: ${p.done_step_count}/${p.step_count}</span>
          <span>数据与证据: ${p.evidence_count} 条</span>
          <span>计划周期: ${esc(p.start_date || '未设')} ~ ${esc(p.target_date || '未设')}</span>
        </div>
      </div>`;
  }).join('') : `<div class="empty-state">暂无科研项目，点击右上角「新建科研项目」或左侧「导入示例科研项目」开始使用。</div>`;

  const todayTasksHtml = state.dailyTasks.length ? state.dailyTasks.slice(0, 6).map(t => {
    const subs = t.subtasks || [];
    const subDone = subs.filter(x => x.done).length;
    return `
      <div class="task-item ${t.is_completed?'completed':''}">
        <input type="checkbox" class="task-check" ${t.is_completed?'checked':''} onchange="toggleTask(${t.id}, this.checked)" />
        <div style="flex:1;">
          <div class="task-title" style="font-size:13.5px; font-weight:600;">${esc(t.title)}</div>
          <div style="font-size:12px; color:var(--muted); margin-top:2px;">
            ${t.project_title ? '📁 ' + esc(t.project_title) : '📝 日常科研'} · 预计 ${t.estimated_hours}h
            ${subs.length ? ` · 分点进度 ${subDone}/${subs.length}` : ''}
          </div>
        </div>
        <span class="badge ${(PRIORITY_MAP[t.priority]||PRIORITY_MAP.medium).cls}">${(PRIORITY_MAP[t.priority]||PRIORITY_MAP.medium).label}</span>
      </div>`;
  }).join('') : `<div class="empty-state" style="padding:20px;">今日暂未记录任务，<a href="javascript:void(0)" onclick="switchTab('daily')">去添加每日工作</a></div>`;

  el.innerHTML = `
    <div class="grid-4">
      <div class="card">
        <div class="stat-label">科研项目总数</div>
        <div class="stat-val">${total}</div>
        <div class="stat-sub">进行中 ${active} · 已结题 ${completed}</div>
      </div>
      <div class="card">
        <div class="stat-label">全部项目平均进度</div>
        <div class="stat-val" style="color:var(--primary);">${avgProg}%</div>
        <div class="progress-wrap" style="margin-top:8px;"><div class="progress-bar" style="width:${avgProg}%"></div></div>
      </div>
      <div class="card">
        <div class="stat-label">今日待完成工作</div>
        <div class="stat-val" style="color:var(--warning);">${todayPending}</div>
        <div class="stat-sub">日期：${esc(state.selectedDate)}</div>
      </div>
      <div class="card">
        <div class="stat-label">今日已完成工作</div>
        <div class="stat-val" style="color:var(--success);">${todayDone}</div>
        <div class="stat-sub">完成率：${(todayPending+todayDone)?Math.round(todayDone*100/(todayPending+todayDone)):0}%</div>
      </div>
    </div>

    <div class="grid-2" style="grid-template-columns: 1.45fr 1fr;">
      <div class="card">
        <div class="section-title">
          <span>全部项目当前进度</span>
          <select style="width:auto; padding:4px 8px; font-size:12px;" onchange="state.filterStatus=this.value; render();">
            <option value="all" ${state.filterStatus==='all'?'selected':''}>全部状态</option>
            <option value="active" ${state.filterStatus==='active'?'selected':''}>进行中</option>
            <option value="planning" ${state.filterStatus==='planning'?'selected':''}>规划中</option>
            <option value="completed" ${state.filterStatus==='completed'?'selected':''}>已结题</option>
            <option value="paused" ${state.filterStatus==='paused'?'selected':''}>已暂停</option>
          </select>
        </div>
        ${projectRows}
      </div>
      <div class="card">
        <div class="section-title">
          <span>今日工作速览 (${esc(state.selectedDate)})</span>
          <button class="btn btn-sm" onclick="switchTab('daily')">进入每日工作 →</button>
        </div>
        ${todayTasksHtml}
      </div>
    </div>`;
}

async function selectProjectAndGo(id) {
  state.selectedProjectId = id;
  state.selectedProjectDetail = await api('/api/projects/' + id);
  switchTab('projects');
}

async function selectProject(id) {
  state.selectedProjectId = id;
  state.selectedProjectDetail = await api('/api/projects/' + id);
  render();
}

/* Inline Project Field Updates */
async function patchProjectField(pid, patch) {
  await api('/api/projects/' + pid, 'PUT', patch);
  showToast('已保存修改');
  await loadAll();
}

async function patchStepField(sid, patch) {
  await api('/api/steps/' + sid, 'PUT', patch);
  showToast('步骤已更新');
  await loadAll();
}

async function patchEvidenceField(eid, patch) {
  await api('/api/evidences/' + eid, 'PUT', patch);
  showToast('证据已更新');
  await loadAll();
}

async function toggleStepSubtask(stepId, idx, checked) {
  const step = state.selectedProjectDetail.steps.find(x => x.id === stepId);
  if (!step) return;
  const subs = (step.subtasks || []).map((it, i) => i === idx ? { ...it, done: checked } : it);
  await api('/api/steps/' + stepId, 'PUT', { subtasks: subs });
  showToast('分点进度已同步');
  await loadAll();
}

async function addStepSubtask(e, stepId) {
  e.preventDefault();
  const inp = document.getElementById('step-sub-inp-' + stepId);
  const text = (inp.value || '').trim();
  if (!text) return;
  const step = state.selectedProjectDetail.steps.find(x => x.id === stepId);
  const subs = [...(step.subtasks || []), { text, done: false }];
  await api('/api/steps/' + stepId, 'PUT', { subtasks: subs });
  showToast('已添加步骤分点');
  await loadAll();
}

async function editStepSubtaskText(el, stepId, idx) {
  const step = state.selectedProjectDetail.steps.find(x => x.id === stepId);
  if (!step) return;
  const subs = step.subtasks || [];
  startInlineEdit(el, subs[idx].text, false, async (val) => {
    const next = subs.map((it, i) => i === idx ? { ...it, text: val.trim() || it.text } : it);
    await api('/api/steps/' + stepId, 'PUT', { subtasks: next });
    showToast('分点内容已更新');
    await loadAll();
  });
}

async function removeStepSubtask(stepId, idx) {
  const step = state.selectedProjectDetail.steps.find(x => x.id === stepId);
  if (!step) return;
  const subs = (step.subtasks || []).filter((_, i) => i !== idx);
  await api('/api/steps/' + stepId, 'PUT', { subtasks: subs });
  await loadAll();
}

function renderProjects(el) {
  if (!state.projects.length) {
    el.innerHTML = `<div class="card empty-state">还没有科研项目。点击右上角「新建科研项目」创建您的第一个项目。</div>`;
    return;
  }

  const listHtml = state.projects.map(p => {
    const st = STATUS_MAP[p.status] || STATUS_MAP.active;
    return `
      <div class="proj-list-item ${p.id===state.selectedProjectId?'selected':''}" onclick="selectProject(${p.id})">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px;">
          <div style="font-weight:700; font-size:14px; line-height:1.35;">${esc(p.title)}</div>
          <span class="badge ${st.cls}">${st.label}</span>
        </div>
        <div style="margin-top:8px; display:flex; align-items:center; gap:8px;">
          <div class="progress-wrap" style="flex:1;"><div class="progress-bar ${p.computed_progress===100?'done':''}" style="width:${p.computed_progress}%"></div></div>
          <span style="font-size:12px; font-weight:700;">${p.computed_progress}%</span>
        </div>
        <div style="font-size:11.5px; color:var(--muted); margin-top:6px;">
          ${esc(p.category)} · ${p.done_step_count}/${p.step_count} 步 · ${p.evidence_count} 条证据
        </div>
      </div>`;
  }).join('');

  const p = state.selectedProjectDetail;
  if (!p) {
    el.innerHTML = `<div class="proj-layout"><div>${listHtml}</div><div class="card empty-state">请选择左侧项目查看详情</div></div>`;
    return;
  }

  const st = STATUS_MAP[p.status] || STATUS_MAP.active;
  const stepsHtml = p.steps.length ? p.steps.map(s => {
    const sst = STEP_STATUS_MAP[s.status] || STEP_STATUS_MAP.todo;
    const subs = s.subtasks || [];
    const subDone = subs.filter(x => x.done).length;

    const subsHtml = `
      <div class="subtasks-box">
        <div style="font-size:12px; font-weight:600; color:var(--muted); margin-bottom:4px;">
          分点确认 (${subDone}/${subs.length})${subs.length ? ' · 全部勾选后该步骤自动完成' : ''}
        </div>
        ${subs.map((it, idx) => `
          <div class="subtask-row ${it.done?'done':''}">
            <input type="checkbox" ${it.done?'checked':''} onchange="toggleStepSubtask(${s.id}, ${idx}, this.checked)" />
            <span class="subtask-text editable" style="flex:1;" title="点击直接编辑" onclick="editStepSubtaskText(this, ${s.id}, ${idx})">${esc(it.text)}</span>
            <button class="btn btn-sm" style="padding:1px 6px; font-size:11px;" onclick="removeStepSubtask(${s.id}, ${idx})">×</button>
          </div>
        `).join('')}
        <form class="subtask-add-row" onsubmit="addStepSubtask(event, ${s.id})">
          <input type="text" id="step-sub-inp-${s.id}" placeholder="输入分点内容后回车添加..." />
          <button type="submit" class="btn btn-sm">＋ 分点</button>
        </form>
      </div>`;

    const evs = (s.evidences || []).map(ev => {
      const isImg = ev.file_rel_path && /\.(png|jpe?g|gif|webp|bmp)$/i.test(ev.file_rel_path);
      return `
        <div class="evidence-item">
          <div style="display:flex; justify-content:space-between; align-items:center; gap:8px;">
            <div>
              <strong class="editable" title="点击直接编辑标题" onclick="startInlineEdit(this, ${JSON.stringify(ev.title)}, false, v => patchEvidenceField(${ev.id}, {title: v}))">📎 ${esc(ev.title)}</strong>
              <span style="font-size:11.5px; color:var(--muted); margin-left:6px;">${esc(ev.created_at)}</span>
            </div>
            <div style="display:flex; gap:6px;">
              ${ev.file_rel_path ? `<a class="btn btn-sm" href="/evidence/${encodeURIComponent(ev.file_rel_path)}" target="_blank">预览/下载</a>` : ''}
              ${ev.local_path ? `<button class="btn btn-sm" onclick="openLocalPath(${JSON.stringify(ev.local_path)})">打开本地路径</button>` : ''}
              ${ev.url ? `<a class="btn btn-sm" href="${esc(ev.url)}" target="_blank">打开链接</a>` : ''}
              <button class="btn btn-sm btn-danger" onclick="deleteEvidence(${ev.id})">删除</button>
            </div>
          </div>
          <p class="pre-text editable" style="margin-top:6px; font-size:12.5px;" title="点击直接编辑数据摘要" onclick="startInlineEdit(this, ${JSON.stringify(ev.data_summary||'')}, true, v => patchEvidenceField(${ev.id}, {data_summary: v}))">${esc(ev.data_summary || '点击补充数据指标或结论摘要...')}</p>
          ${ev.local_path ? `<div style="margin-top:4px; font-size:12px; color:var(--muted);">本地路径：<code>${esc(ev.local_path)}</code></div>` : ''}
          ${isImg ? `<img class="evidence-img-preview" src="/evidence/${encodeURIComponent(ev.file_rel_path)}" onclick="window.open(this.src, '_blank')" />` : ''}
        </div>`;
    }).join('');

    return `
      <div class="step-card">
        <div class="step-header">
          <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap; flex:1;">
            <span style="font-weight:700; font-size:14.5px;">Step ${s.step_order}.</span>
            <span class="editable" style="font-weight:700; font-size:14.5px;" title="点击直接编辑步骤名称" onclick="startInlineEdit(this, ${JSON.stringify(s.title)}, false, v => patchStepField(${s.id}, {title: v}))">${esc(s.title)}</span>
            <span class="badge ${sst.cls}">${sst.label}</span>
            ${s.due_date ? `<span style="font-size:12px; color:var(--muted);">截止 ${esc(s.due_date)}</span>` : ''}
          </div>
          <div style="display:flex; gap:6px; align-items:center;">
            <select style="width:auto; padding:4px 8px; font-size:12px;" onchange="quickStepStatus(${s.id}, this.value)">
              <option value="todo" ${s.status==='todo'?'selected':''}>未开始 (0%)</option>
              <option value="in_progress" ${s.status==='in_progress'?'selected':''}>进行中 (${s.progress||50}%)</option>
              <option value="done" ${s.status==='done'?'selected':''}>已完成 (100%)</option>
            </select>
            <button class="btn btn-sm btn-primary" onclick="openEvidenceModal(${p.id}, ${s.id})">＋ 数据/证据</button>
            <button class="btn btn-sm" onclick="openStepModal(${p.id}, ${s.id})">属性</button>
            <button class="btn btn-sm btn-danger" onclick="deleteStep(${s.id})">删除</button>
          </div>
        </div>
        <p class="pre-text editable" style="margin-top:8px; font-size:13px;" title="点击直接编辑步骤说明" onclick="startInlineEdit(this, ${JSON.stringify(s.description||'')}, true, v => patchStepField(${s.id}, {description: v}))">${esc(s.description || '点击添加步骤说明...')}</p>
        ${subsHtml}
        <div class="evidence-list">
          ${evs || `<div style="font-size:12px; color:var(--muted);">暂无数据与证据，点击右上角「＋ 数据/证据」挂载实验数据、图表、本地文件或结论记录。</div>`}
        </div>
      </div>`;
  }).join('') : `<div class="empty-state" style="padding:24px;">暂无研究步骤，点击「＋ 添加研究步骤」拆解项目里程碑并记录每一步的数据与证据。</div>`;

  el.innerHTML = `
    <div class="proj-layout">
      <div>${listHtml}</div>
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:12px;">
            <div style="flex:1;">
              <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                <h2 class="editable" style="margin:0; font-size:19px;" title="点击直接编辑项目名称" onclick="startInlineEdit(this, ${JSON.stringify(p.title)}, false, v => patchProjectField(${p.id}, {title: v}))">${esc(p.title)}</h2>
                <span class="badge ${st.cls}">${st.label}</span>
              </div>
              <div style="font-size:12.5px; color:var(--muted); margin-top:6px;">
                类型：<span class="editable" title="点击直接编辑类型" onclick="startInlineEdit(this, ${JSON.stringify(p.category)}, false, v => patchProjectField(${p.id}, {category: v}))">${esc(p.category)}</span>
                · 周期：${esc(p.start_date||'未设')} 至 ${esc(p.target_date||'未设')}
                · 进度模式：${p.progress_mode==='auto'?'按步骤自动计算':'手动指定'}
              </div>
            </div>
            <div style="display:flex; gap:8px;">
              <button class="btn btn-sm" onclick="openProjectModal(${p.id})">✏️ 更多设置</button>
              <button class="btn btn-sm btn-danger" onclick="deleteProject(${p.id})">删除</button>
            </div>
          </div>
          <div style="margin-top:14px; display:flex; align-items:center; gap:12px;">
            <div class="progress-wrap" style="height:10px; flex:1;">
              <div class="progress-bar ${p.computed_progress===100?'done':''}" style="width:${p.computed_progress}%"></div>
            </div>
            <strong style="font-size:16px; color:${p.computed_progress===100?'var(--success)':'var(--primary)'};">${p.computed_progress}%</strong>
          </div>
        </div>

        <div class="grid-2">
          <div class="card">
            <div class="section-title"><span>🎯 研究目标</span><span style="font-size:11.5px; color:var(--muted); font-weight:400;">点击文本直接编辑</span></div>
            <p class="pre-text editable" onclick="startInlineEdit(this, ${JSON.stringify(p.objectives||'')}, true, v => patchProjectField(${p.id}, {objectives: v}))">${esc(p.objectives || '点击填写研究目标...')}</p>
          </div>
          <div class="card">
            <div class="section-title"><span>💡 研究意义</span><span style="font-size:11.5px; color:var(--muted); font-weight:400;">点击文本直接编辑</span></div>
            <p class="pre-text editable" onclick="startInlineEdit(this, ${JSON.stringify(p.significance||'')}, true, v => patchProjectField(${p.id}, {significance: v}))">${esc(p.significance || '点击填写研究意义...')}</p>
          </div>
        </div>

        <div class="card">
          <div class="section-title">
            <span>🧪 研究步骤、数据与证据 (${p.steps.length})</span>
            <button class="btn btn-primary btn-sm" onclick="openStepModal(${p.id})">＋ 添加研究步骤</button>
          </div>
          ${stepsHtml}
        </div>
      </div>
    </div>`;
}

/* Daily Tasks + Subtasks + Click-to-edit */
async function patchTaskField(tid, patch) {
  await api('/api/daily-tasks/' + tid, 'PUT', patch);
  showToast('任务已更新');
  await loadAll();
}

async function toggleTaskSubtask(tid, idx, checked) {
  const task = state.dailyTasks.find(x => x.id === tid);
  if (!task) return;
  const subs = (task.subtasks || []).map((it, i) => i === idx ? { ...it, done: checked } : it);
  await api('/api/daily-tasks/' + tid, 'PUT', { subtasks: subs });
  showToast(subs.length && subs.every(x => x.done) ? '所有分点已完成，任务已自动标记完成！' : '分点状态已更新');
  await loadAll();
}

async function addTaskSubtask(e, tid) {
  e.preventDefault();
  const inp = document.getElementById('task-sub-inp-' + tid);
  const text = (inp.value || '').trim();
  if (!text) return;
  const task = state.dailyTasks.find(x => x.id === tid);
  const subs = [...(task.subtasks || []), { text, done: false }];
  await api('/api/daily-tasks/' + tid, 'PUT', { subtasks: subs });
  showToast('已添加分点');
  await loadAll();
}

async function editTaskSubtaskText(el, tid, idx) {
  const task = state.dailyTasks.find(x => x.id === tid);
  if (!task) return;
  const subs = task.subtasks || [];
  startInlineEdit(el, subs[idx].text, false, async (val) => {
    const next = subs.map((it, i) => i === idx ? { ...it, text: val.trim() || it.text } : it);
    await api('/api/daily-tasks/' + tid, 'PUT', { subtasks: next });
    showToast('分点内容已更新');
    await loadAll();
  });
}

async function removeTaskSubtask(tid, idx) {
  const task = state.dailyTasks.find(x => x.id === tid);
  if (!task) return;
  const subs = (task.subtasks || []).filter((_, i) => i !== idx);
  await api('/api/daily-tasks/' + tid, 'PUT', { subtasks: subs });
  await loadAll();
}

function renderDaily(el) {
  const pending = state.dailyTasks.filter(t => !t.is_completed);
  const completed = state.dailyTasks.filter(t => t.is_completed);
  const projOpts = `<option value="">日常科研事务（不关联特定项目）</option>` +
    state.projects.map(p => `<option value="${p.id}">${esc(p.title)}</option>`).join('');

  const renderTaskCard = (t) => {
    const pr = PRIORITY_MAP[t.priority] || PRIORITY_MAP.medium;
    const subs = t.subtasks || [];
    const subDone = subs.filter(x => x.done).length;
    return `
      <div class="task-item ${t.is_completed?'completed':''}">
        <input type="checkbox" class="task-check" ${t.is_completed?'checked':''} onchange="toggleTask(${t.id}, this.checked)" title="${subs.length?'点击将全部设为空/完成':'切换任务完成状态'}" />
        <div style="flex:1; min-width:0;">
          <div class="task-title editable" style="font-size:14px; font-weight:600;" title="点击直接编辑任务名称" onclick="startInlineEdit(this, ${JSON.stringify(t.title)}, false, v => patchTaskField(${t.id}, {title: v}))">${esc(t.title)}</div>
          <div style="font-size:12px; color:var(--muted); margin-top:3px;">
            ${t.project_title ? '📁 ' + esc(t.project_title) : '📝 日常科研'} · 预计
            <span class="editable" title="点击修改工时" onclick="startInlineEdit(this, ${JSON.stringify(String(t.estimated_hours))}, false, v => patchTaskField(${t.id}, {estimated_hours: Number(v)||1}))">${t.estimated_hours}</span>h
            ${subs.length ? ` · 分点完成 <strong>${subDone}/${subs.length}</strong>` : ''}
            ${t.completed_at ? ' · 完成于 ' + esc(t.completed_at.slice(11,16)) : ''}
          </div>
          <div class="editable" style="font-size:12.5px; color:#475569; margin-top:5px; white-space:pre-wrap;" title="点击直接编辑备注" onclick="startInlineEdit(this, ${JSON.stringify(t.notes||'')}, true, v => patchTaskField(${t.id}, {notes: v}))">${esc(t.notes || '点击添加备注 / 产出说明...')}</div>
          <div class="subtasks-box">
            <div style="font-size:11.5px; font-weight:600; color:var(--muted); margin-bottom:3px;">
              分点确认 ${subs.length ? `(${subDone}/${subs.length}，全部勾选后任务自动完成)` : '(可选，添加分点后需全部完成才算完成)'}
            </div>
            ${subs.map((it, idx) => `
              <div class="subtask-row ${it.done?'done':''}">
                <input type="checkbox" ${it.done?'checked':''} onchange="toggleTaskSubtask(${t.id}, ${idx}, this.checked)" />
                <span class="subtask-text editable" style="flex:1;" title="点击直接编辑分点" onclick="editTaskSubtaskText(this, ${t.id}, ${idx})">${esc(it.text)}</span>
                <button class="btn btn-sm" style="padding:1px 6px; font-size:11px;" onclick="removeTaskSubtask(${t.id}, ${idx})">×</button>
              </div>
            `).join('')}
            <form class="subtask-add-row" onsubmit="addTaskSubtask(event, ${t.id})">
              <input type="text" id="task-sub-inp-${t.id}" placeholder="添加分点确认项并回车..." />
              <button type="submit" class="btn btn-sm">＋ 分点</button>
            </form>
          </div>
        </div>
        <div style="display:flex; flex-direction:column; align-items:flex-end; gap:6px; flex-shrink:0;">
          <span class="badge ${pr.cls}">${pr.label}</span>
          <button class="btn btn-sm btn-danger" onclick="deleteTask(${t.id})">删除</button>
        </div>
      </div>`;
  };

  el.innerHTML = `
    <div class="card" style="margin-bottom:16px;">
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px;">
        <div style="display:flex; align-items:center; gap:10px;">
          <strong>工作日期：</strong>
          <input type="date" style="width:auto;" value="${esc(state.selectedDate)}" onchange="changeDailyDate(this.value)" />
          <button class="btn btn-sm" onclick="changeDailyDate('${new Date().toISOString().slice(0,10)}')">回到今天</button>
        </div>
        <div style="display:flex; gap:8px;">
          <button class="btn btn-sm" onclick="carryOverTasks()">↪ 顺延昨日未完成任务到当前日期</button>
          <button class="btn btn-sm" onclick="exportDailyMarkdown()">📋 导出当日工作日报 (MD)</button>
        </div>
      </div>

      <form onsubmit="addDailyTask(event)" style="margin-top:14px; padding-top:14px; border-top:1px solid var(--border); display:grid; grid-template-columns: 2fr 1.2fr 0.7fr 0.6fr auto; gap:10px; align-items:end;">
        <div>
          <label style="font-size:12px; font-weight:600; display:block; margin-bottom:4px;">今日需完成工作内容 *</label>
          <input type="text" id="new-task-title" required placeholder="例如：完成样地3点云赋色对比脚本测试..." />
        </div>
        <div>
          <label style="font-size:12px; font-weight:600; display:block; margin-bottom:4px;">关联科研项目</label>
          <select id="new-task-proj">${projOpts}</select>
        </div>
        <div>
          <label style="font-size:12px; font-weight:600; display:block; margin-bottom:4px;">优先级</label>
          <select id="new-task-prio">
            <option value="high">高</option>
            <option value="medium" selected>中</option>
            <option value="low">低</option>
          </select>
        </div>
        <div>
          <label style="font-size:12px; font-weight:600; display:block; margin-bottom:4px;">预计(h)</label>
          <input type="number" id="new-task-hours" step="0.5" min="0.5" value="1.5" />
        </div>
        <button type="submit" class="btn btn-primary">＋ 添加任务</button>
        <div style="grid-column: 1 / -1; display:grid; grid-template-columns: 1fr 1fr; gap:10px;">
          <input type="text" id="new-task-notes" placeholder="补充备注 / 预期产出（选填）" />
          <input type="text" id="new-task-subs" placeholder="分点确认项（选填，多个分点用英文或中文分号 ; / ； 隔开）" />
        </div>
      </form>
    </div>

    <div class="grid-2">
      <div class="card">
        <div class="section-title">
          <span>⏳ 未完成任务 (${pending.length})</span>
          <span style="font-size:11.5px; color:var(--muted); font-weight:400;">点击文字即可直接编辑</span>
        </div>
        ${pending.length ? pending.map(renderTaskCard).join('') : `<div class="empty-state" style="padding:24px;">当前日期所有任务均已完成 🎉</div>`}
      </div>
      <div class="card">
        <div class="section-title">
          <span>✅ 已完成任务 (${completed.length})</span>
        </div>
        ${completed.length ? completed.map(renderTaskCard).join('') : `<div class="empty-state" style="padding:24px;">暂无已完成任务</div>`}
      </div>
    </div>`;
}

/* Feedback & Improvement Tracking */
function renderFeedback(el) {
  const total = state.feedbacks.length;
  const openCnt = state.feedbacks.filter(f => f.status === 'open').length;
  const activeCnt = state.feedbacks.filter(f => f.status === 'planned' || f.status === 'in_progress').length;
  const resolvedCnt = state.feedbacks.filter(f => f.status === 'resolved').length;

  let list = state.feedbacks;
  if (state.fbFilterStatus !== 'all') list = list.filter(f => f.status === state.fbFilterStatus);
  if (state.fbFilterType !== 'all') list = list.filter(f => f.feedback_type === state.fbFilterType);
  if (state.fbSearch.trim()) {
    const q = state.fbSearch.trim().toLowerCase();
    list = list.filter(f =>
      (f.title || '').toLowerCase().includes(q) ||
      (f.description || '').toLowerCase().includes(q) ||
      (f.module || '').toLowerCase().includes(q)
    );
  }

  const modOpts = FB_MODULES.map(m => `<option value="${esc(m)}">${esc(m)}</option>`).join('');

  const itemsHtml = list.length ? list.map(f => {
    const tp = FB_TYPE_MAP[f.feedback_type] || FB_TYPE_MAP.ux;
    const pr = PRIORITY_MAP[f.priority] || PRIORITY_MAP.medium;
    const st = FB_STATUS_MAP[f.status] || FB_STATUS_MAP.open;
    return `
      <div class="fb-item ${f.status==='resolved'?'resolved':''}">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:12px;">
          <div style="flex:1;">
            <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
              <strong style="font-size:15px;">#${f.id} ${esc(f.title)}</strong>
              <span class="badge ${tp.cls}">${tp.label}</span>
              <span class="badge ${pr.cls}">${pr.label}</span>
              <span class="badge ${st.cls}">${st.label}</span>
              <span class="badge badge-plan">模块：${esc(f.module)}</span>
              ${f.target_version ? `<span class="badge badge-active">目标版本 ${esc(f.target_version)}</span>` : ''}
            </div>
            ${f.description ? `<p class="pre-text" style="margin-top:8px; font-size:13px;">${esc(f.description)}</p>` : ''}
            ${f.resolution_note ? `<div style="margin-top:8px; padding:8px 10px; background:var(--success-soft); border-radius:6px; font-size:12.5px; color:#065f46;"><strong>处理备注：</strong>${esc(f.resolution_note)}</div>` : ''}
            <div style="font-size:11.5px; color:var(--muted); margin-top:8px;">
              记录时间：${esc(f.created_at)} · 更新于：${esc(f.updated_at)}
              ${f.resolved_at ? ` · 解决于：${esc(f.resolved_at)}` : ''}
            </div>
          </div>
          <div style="display:flex; gap:6px; align-items:center; flex-shrink:0;">
            <select style="width:auto; padding:4px 8px; font-size:12px;" onchange="quickFeedbackStatus(${f.id}, this.value)">
              <option value="open" ${f.status==='open'?'selected':''}>待处理</option>
              <option value="planned" ${f.status==='planned'?'selected':''}>已排期</option>
              <option value="in_progress" ${f.status==='in_progress'?'selected':''}>处理中</option>
              <option value="resolved" ${f.status==='resolved'?'selected':''}>已解决</option>
              <option value="wontfix" ${f.status==='wontfix'?'selected':''}>暂不处理</option>
            </select>
            <button class="btn btn-sm" onclick="openFeedbackModal(${f.id})">编辑/备注</button>
            <button class="btn btn-sm btn-danger" onclick="deleteFeedback(${f.id})">删除</button>
          </div>
        </div>
      </div>`;
  }).join('') : `<div class="empty-state" style="padding:28px;">当前筛选条件下暂无意见改进记录。</div>`;

  el.innerHTML = `
    <div class="grid-4">
      <div class="card">
        <div class="stat-label">累计记录总数</div>
        <div class="stat-val">${total}</div>
      </div>
      <div class="card">
        <div class="stat-label">待处理 (Open)</div>
        <div class="stat-val" style="color:var(--warning);">${openCnt}</div>
      </div>
      <div class="card">
        <div class="stat-label">已排期 / 处理中</div>
        <div class="stat-val" style="color:var(--primary);">${activeCnt}</div>
      </div>
      <div class="card">
        <div class="stat-label">已解决 (Resolved)</div>
        <div class="stat-val" style="color:var(--success);">${resolvedCnt}</div>
      </div>
    </div>

    <div class="card" style="margin-bottom:16px;">
      <div class="section-title">
        <span>📝 快速记录使用问题或改进建议</span>
        <button class="btn btn-sm" onclick="exportFeedbackMarkdown()">📋 导出待处理改进清单 (MD)</button>
      </div>
      <form onsubmit="quickAddFeedback(event)" style="display:grid; grid-template-columns: 2fr 1fr 0.9fr 0.8fr auto; gap:10px; align-items:end;">
        <div>
          <label style="font-size:12px; font-weight:600; display:block; margin-bottom:4px;">问题 / 改进标题 *</label>
          <input type="text" id="fb-q-title" required placeholder="例如：项目详情页希望支持批量上传实验图片..." />
        </div>
        <div>
          <label style="font-size:12px; font-weight:600; display:block; margin-bottom:4px;">所属模块</label>
          <select id="fb-q-mod">${modOpts}</select>
        </div>
        <div>
          <label style="font-size:12px; font-weight:600; display:block; margin-bottom:4px;">类型</label>
          <select id="fb-q-type">
            <option value="ux">体验优化</option>
            <option value="feature">新功能需求</option>
            <option value="bug">问题缺陷</option>
            <option value="perf">性能与稳定</option>
          </select>
        </div>
        <div>
          <label style="font-size:12px; font-weight:600; display:block; margin-bottom:4px;">优先级</label>
          <select id="fb-q-prio">
            <option value="high">高</option>
            <option value="medium" selected>中</option>
            <option value="low">低</option>
          </select>
        </div>
        <button type="submit" class="btn btn-primary">＋ 记录意见</button>
        <div style="grid-column: 1 / -1;">
          <input type="text" id="fb-q-desc" placeholder="详细描述：触发场景、痛点说明或期望实现方式（选填）" />
        </div>
      </form>
    </div>

    <div class="card">
      <div class="section-title" style="flex-wrap:wrap; gap:10px;">
        <span>📋 改进清单 (${list.length})</span>
        <div style="display:flex; gap:8px; flex-wrap:wrap;">
          <input type="text" style="width:200px; padding:5px 9px; font-size:12px;" placeholder="搜索标题/描述..." value="${esc(state.fbSearch)}" oninput="state.fbSearch=this.value; render();" />
          <select style="width:auto; padding:5px 8px; font-size:12px;" onchange="state.fbFilterStatus=this.value; render();">
            <option value="all" ${state.fbFilterStatus==='all'?'selected':''}>全部状态</option>
            <option value="open" ${state.fbFilterStatus==='open'?'selected':''}>待处理</option>
            <option value="planned" ${state.fbFilterStatus==='planned'?'selected':''}>已排期</option>
            <option value="in_progress" ${state.fbFilterStatus==='in_progress'?'selected':''}>处理中</option>
            <option value="resolved" ${state.fbFilterStatus==='resolved'?'selected':''}>已解决</option>
            <option value="wontfix" ${state.fbFilterStatus==='wontfix'?'selected':''}>暂不处理</option>
          </select>
          <select style="width:auto; padding:5px 8px; font-size:12px;" onchange="state.fbFilterType=this.value; render();">
            <option value="all" ${state.fbFilterType==='all'?'selected':''}>全部类型</option>
            <option value="bug" ${state.fbFilterType==='bug'?'selected':''}>问题缺陷</option>
            <option value="feature" ${state.fbFilterType==='feature'?'selected':''}>新功能需求</option>
            <option value="ux" ${state.fbFilterType==='ux'?'selected':''}>体验优化</option>
            <option value="perf" ${state.fbFilterType==='perf'?'selected':''}>性能与稳定</option>
          </select>
        </div>
      </div>
      ${itemsHtml}
    </div>`;
}

async function quickAddFeedback(e) {
  e.preventDefault();
  await api('/api/feedbacks', 'POST', {
    title: document.getElementById('fb-q-title').value.trim(),
    module: document.getElementById('fb-q-mod').value,
    feedback_type: document.getElementById('fb-q-type').value,
    priority: document.getElementById('fb-q-prio').value,
    description: document.getElementById('fb-q-desc').value.trim(),
    status: 'open'
  });
  showToast('改进意见已记录');
  await loadAll();
}

function openFeedbackModal(editId = null) {
  const f = editId ? state.feedbacks.find(x => x.id === editId) : null;
  const modOpts = FB_MODULES.map(m => `<option value="${esc(m)}" ${f?.module===m?'selected':''}>${esc(m)}</option>`).join('');
  openModal(`
    <h3 style="margin:0 0 14px;">${f ? '编辑改进记录 #' + f.id : '记录系统意见 / 改进需求'}</h3>
    <form onsubmit="saveFeedbackModal(event, ${editId || 'null'})">
      <div class="form-group"><label>标题 *</label><input type="text" id="fb-m-title" required value="${esc(f?.title||'')}" placeholder="简述遇到的问题或希望新增的功能" /></div>
      <div class="grid-2">
        <div class="form-group"><label>所属模块</label><select id="fb-m-mod">${modOpts}</select></div>
        <div class="form-group"><label>类型</label>
          <select id="fb-m-type">
            <option value="ux" ${f?.feedback_type==='ux'?'selected':''}>体验优化</option>
            <option value="feature" ${f?.feedback_type==='feature'?'selected':''}>新功能需求</option>
            <option value="bug" ${f?.feedback_type==='bug'?'selected':''}>问题缺陷</option>
            <option value="perf" ${f?.feedback_type==='perf'?'selected':''}>性能与稳定</option>
          </select>
        </div>
      </div>
      <div class="grid-2">
        <div class="form-group"><label>优先级</label>
          <select id="fb-m-prio">
            <option value="high" ${f?.priority==='high'?'selected':''}>高</option>
            <option value="medium" ${(!f||f.priority==='medium')?'selected':''}>中</option>
            <option value="low" ${f?.priority==='low'?'selected':''}>低</option>
          </select>
        </div>
        <div class="form-group"><label>处理状态</label>
          <select id="fb-m-status">
            <option value="open" ${(!f||f.status==='open')?'selected':''}>待处理</option>
            <option value="planned" ${f?.status==='planned'?'selected':''}>已排期</option>
            <option value="in_progress" ${f?.status==='in_progress'?'selected':''}>处理中</option>
            <option value="resolved" ${f?.status==='resolved'?'selected':''}>已解决</option>
            <option value="wontfix" ${f?.status==='wontfix'?'selected':''}>暂不处理</option>
          </select>
        </div>
      </div>
      <div class="form-group"><label>计划处理版本（选填）</label><input type="text" id="fb-m-ver" value="${esc(f?.target_version||'')}" placeholder="例如：v1.1" /></div>
      <div class="form-group"><label>详细描述 / 复现场景 / 期望效果</label><textarea id="fb-m-desc">${esc(f?.description||'')}</textarea></div>
      <div class="form-group"><label>后续处理方案 / 解决备注</label><textarea id="fb-m-res" placeholder="统一处理时记录修改位置或解决说明...">${esc(f?.resolution_note||'')}</textarea></div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:14px;">
        <button type="button" class="btn" onclick="closeModal()">取消</button>
        <button type="submit" class="btn btn-primary">保存</button>
      </div>
    </form>`);
}

async function saveFeedbackModal(e, editId) {
  e.preventDefault();
  const payload = {
    title: document.getElementById('fb-m-title').value.trim(),
    module: document.getElementById('fb-m-mod').value,
    feedback_type: document.getElementById('fb-m-type').value,
    priority: document.getElementById('fb-m-prio').value,
    status: document.getElementById('fb-m-status').value,
    target_version: document.getElementById('fb-m-ver').value.trim(),
    description: document.getElementById('fb-m-desc').value,
    resolution_note: document.getElementById('fb-m-res').value
  };
  if (editId) {
    await api('/api/feedbacks/' + editId, 'PUT', payload);
    showToast('改进记录已更新');
  } else {
    await api('/api/feedbacks', 'POST', payload);
    showToast('改进意见已记录');
  }
  closeModal();
  await loadAll();
}

async function quickFeedbackStatus(id, status) {
  await api('/api/feedbacks/' + id, 'PUT', { status });
  showToast('状态已更新');
  await loadAll();
}

async function deleteFeedback(id) {
  if (!confirm('确定删除该条改进记录吗？')) return;
  await api('/api/feedbacks/' + id, 'DELETE');
  showToast('记录已删除');
  await loadAll();
}

function exportFeedbackMarkdown() {
  const pending = state.feedbacks.filter(f => f.status !== 'resolved' && f.status !== 'wontfix');
  const lines = [
    `# 科研项目管理系统 — 待统一处理改进清单 (${new Date().toISOString().slice(0,10)})`,
    '',
    `共 ${pending.length} 项待处理/排期中需求与问题：`,
    ''
  ];
  pending.forEach(f => {
    const tp = (FB_TYPE_MAP[f.feedback_type] || FB_TYPE_MAP.ux).label;
    const pr = (PRIORITY_MAP[f.priority] || PRIORITY_MAP.medium).label;
    const st = (FB_STATUS_MAP[f.status] || FB_STATUS_MAP.open).label;
    lines.push(`## #${f.id} ${f.title}`);
    lines.push(`- **模块**：${f.module} | **类型**：${tp} | **优先级**：${pr} | **状态**：${st}${f.target_version ? ` | **目标版本**：${f.target_version}` : ''}`);
    if (f.description) lines.push(`- **详细说明**：${f.description}`);
    lines.push('');
  });
  if (!pending.length) lines.push('当前没有待处理的改进项。');
  const blob = new Blob([lines.join('\n')], { type: 'text/markdown;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `待处理改进清单_${new Date().toISOString().slice(0,10)}.md`;
  a.click();
}

/* Modal Helpers */
function openModal(html) {
  document.getElementById('modal-body').innerHTML = html;
  document.getElementById('modal-mask').classList.add('open');
}
function closeModal() {
  document.getElementById('modal-mask').classList.remove('open');
}

function openProjectModal(editId = null) {
  const p = editId ? state.selectedProjectDetail : null;
  openModal(`
    <h3 style="margin:0 0 14px;">${p ? '编辑科研项目' : '新建科研项目'}</h3>
    <form onsubmit="saveProject(event, ${editId || 'null'})">
      <div class="form-group"><label>项目名称 *</label><input type="text" id="p-title" required value="${esc(p?.title||'')}" /></div>
      <div class="grid-2">
        <div class="form-group"><label>项目类型/来源</label><input type="text" id="p-cat" value="${esc(p?.category||'科研项目')}" /></div>
        <div class="form-group"><label>当前状态</label>
          <select id="p-status">
            <option value="active" ${p?.status==='active'?'selected':''}>进行中</option>
            <option value="planning" ${p?.status==='planning'?'selected':''}>规划中</option>
            <option value="completed" ${p?.status==='completed'?'selected':''}>已结题</option>
            <option value="paused" ${p?.status==='paused'?'selected':''}>已暂停</option>
          </select>
        </div>
      </div>
      <div class="grid-2">
        <div class="form-group"><label>开始日期</label><input type="date" id="p-start" value="${esc(p?.start_date||'')}" /></div>
        <div class="form-group"><label>目标完成日期</label><input type="date" id="p-target" value="${esc(p?.target_date||'')}" /></div>
      </div>
      <div class="grid-2">
        <div class="form-group"><label>进度计算模式</label>
          <select id="p-pmode">
            <option value="auto" ${p?.progress_mode!=='manual'?'selected':''}>按研究步骤自动计算</option>
            <option value="manual" ${p?.progress_mode==='manual'?'selected':''}>手动指定百分比</option>
          </select>
        </div>
        <div class="form-group"><label>手动进度 (0-100)</label><input type="number" id="p-mprog" min="0" max="100" value="${p?.manual_progress??0}" /></div>
      </div>
      <div class="form-group"><label>研究目标</label><textarea id="p-obj" placeholder="核心科学问题、量化指标、阶段交付物...">${esc(p?.objectives||'')}</textarea></div>
      <div class="form-group"><label>研究意义</label><textarea id="p-sig" placeholder="理论价值、应用前景、创新点...">${esc(p?.significance||'')}</textarea></div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:14px;">
        <button type="button" class="btn" onclick="closeModal()">取消</button>
        <button type="submit" class="btn btn-primary">保存</button>
      </div>
    </form>`);
}

async function saveProject(e, editId) {
  e.preventDefault();
  const payload = {
    title: document.getElementById('p-title').value.trim(),
    category: document.getElementById('p-cat').value.trim(),
    status: document.getElementById('p-status').value,
    start_date: document.getElementById('p-start').value,
    target_date: document.getElementById('p-target').value,
    progress_mode: document.getElementById('p-pmode').value,
    manual_progress: Number(document.getElementById('p-mprog').value || 0),
    objectives: document.getElementById('p-obj').value,
    significance: document.getElementById('p-sig').value
  };
  if (editId) {
    await api('/api/projects/' + editId, 'PUT', payload);
    showToast('项目已更新');
  } else {
    const res = await api('/api/projects', 'POST', payload);
    state.selectedProjectId = res.id;
    showToast('项目已创建');
  }
  closeModal();
  await loadAll();
}

async function deleteProject(id) {
  if (!confirm('确定删除该项目及其全部步骤与证据记录吗？')) return;
  await api('/api/projects/' + id, 'DELETE');
  if (state.selectedProjectId === id) state.selectedProjectId = null;
  showToast('项目已删除');
  await loadAll();
}

function openStepModal(projectId, stepId = null) {
  const s = stepId ? state.selectedProjectDetail.steps.find(x => x.id === stepId) : null;
  const nextOrder = (state.selectedProjectDetail?.steps?.length || 0) + 1;
  openModal(`
    <h3 style="margin:0 0 14px;">${s ? '编辑研究步骤' : '添加研究步骤'}</h3>
    <form onsubmit="saveStep(event, ${projectId}, ${stepId || 'null'})">
      <div class="grid-2">
        <div class="form-group"><label>步骤序号</label><input type="number" id="s-order" min="1" value="${s?.step_order ?? nextOrder}" /></div>
        <div class="form-group"><label>步骤状态</label>
          <select id="s-status">
            <option value="todo" ${s?.status==='todo'?'selected':''}>未开始</option>
            <option value="in_progress" ${s?.status==='in_progress'?'selected':''}>进行中</option>
            <option value="done" ${s?.status==='done'?'selected':''}>已完成</option>
          </select>
        </div>
      </div>
      <div class="form-group"><label>步骤名称 *</label><input type="text" id="s-title" required value="${esc(s?.title||'')}" /></div>
      <div class="grid-2">
        <div class="form-group"><label>步骤完成度 (进行中时生效 0-100)</label><input type="number" id="s-prog" min="0" max="100" value="${s?.progress ?? 50}" /></div>
        <div class="form-group"><label>计划截止日期</label><input type="date" id="s-due" value="${esc(s?.due_date||'')}" /></div>
      </div>
      <div class="form-group"><label>步骤说明 / 方法</label><textarea id="s-desc">${esc(s?.description||'')}</textarea></div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:14px;">
        <button type="button" class="btn" onclick="closeModal()">取消</button>
        <button type="submit" class="btn btn-primary">保存步骤</button>
      </div>
    </form>`);
}

async function saveStep(e, projectId, stepId) {
  e.preventDefault();
  const payload = {
    project_id: projectId,
    step_order: Number(document.getElementById('s-order').value || 1),
    status: document.getElementById('s-status').value,
    title: document.getElementById('s-title').value.trim(),
    progress: Number(document.getElementById('s-prog').value || 0),
    due_date: document.getElementById('s-due').value,
    description: document.getElementById('s-desc').value
  };
  if (stepId) await api('/api/steps/' + stepId, 'PUT', payload);
  else await api('/api/steps', 'POST', payload);
  closeModal();
  showToast('研究步骤已保存');
  await loadAll();
}

async function quickStepStatus(stepId, status) {
  await api('/api/steps/' + stepId, 'PUT', { status, progress: status === 'done' ? 100 : (status === 'todo' ? 0 : 50) });
  showToast('步骤状态已更新');
  await loadAll();
}

async function deleteStep(stepId) {
  if (!confirm('确定删除该研究步骤及关联的数据证据吗？')) return;
  await api('/api/steps/' + stepId, 'DELETE');
  showToast('步骤已删除');
  await loadAll();
}

function openEvidenceModal(projectId, stepId) {
  openModal(`
    <h3 style="margin:0 0 14px;">添加数据与证据</h3>
    <form onsubmit="saveEvidence(event, ${projectId}, ${stepId})">
      <div class="form-group"><label>证据/数据标题 *</label><input type="text" id="ev-title" required placeholder="例如：实验结果对比图 / 样地3点云参数表" /></div>
      <div class="form-group"><label>上传文件（图片/PDF/CSV/文档，自动归档到本地证据库）</label><input type="file" id="ev-file" /></div>
      <div class="form-group"><label>或关联 Windows 本地文件/文件夹路径（可一键打开）</label><input type="text" id="ev-local" placeholder="例如：D:\\Research\\Exp01\\metrics.csv" /></div>
      <div class="form-group"><label>参考链接 URL（论文/代码仓库/在线文档）</label><input type="text" id="ev-url" placeholder="https://..." /></div>
      <div class="form-group"><label>关键数据指标 / 实验结论摘要</label><textarea id="ev-summary" placeholder="记录关键数值、参数设置、实验现象或支撑结论..."></textarea></div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:14px;">
        <button type="button" class="btn" onclick="closeModal()">取消</button>
        <button type="submit" class="btn btn-primary">保存证据</button>
      </div>
    </form>`);
}

async function saveEvidence(e, projectId, stepId) {
  e.preventDefault();
  const fileInput = document.getElementById('ev-file');
  let filename = '';
  let content_base64 = '';
  if (fileInput.files && fileInput.files[0]) {
    const f = fileInput.files[0];
    filename = f.name;
    content_base64 = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result).split(',')[1] || '');
      reader.onerror = reject;
      reader.readAsDataURL(f);
    });
  }
  await api('/api/evidences', 'POST', {
    project_id: projectId,
    step_id: stepId,
    title: document.getElementById('ev-title').value.trim(),
    local_path: document.getElementById('ev-local').value.trim(),
    url: document.getElementById('ev-url').value.trim(),
    data_summary: document.getElementById('ev-summary').value,
    filename,
    content_base64
  });
  closeModal();
  showToast('数据与证据已保存');
  await loadAll();
}

async function deleteEvidence(id) {
  if (!confirm('确定删除该条数据与证据吗？')) return;
  await api('/api/evidences/' + id, 'DELETE');
  showToast('证据已删除');
  await loadAll();
}

async function openLocalPath(path) {
  try {
    await api('/api/open-local', 'POST', { path });
    showToast('已在 Windows 中打开');
  } catch (err) {
    alert('打开失败：' + err.message);
  }
}

async function openDataFolder() {
  try {
    await api('/api/open-local', 'POST', { path: '__DATA_DIR__' });
  } catch (err) {
    alert(err.message);
  }
}

/* Daily Tasks */
async function changeDailyDate(d) {
  state.selectedDate = d;
  state.dailyTasks = await api('/api/daily-tasks?date=' + encodeURIComponent(d));
  render();
}

async function addDailyTask(e) {
  e.preventDefault();
  const projVal = document.getElementById('new-task-proj').value;
  const subsRaw = document.getElementById('new-task-subs').value.trim();
  const subtasks = subsRaw
    ? subsRaw.split(/[;；]+/).map(s => s.trim()).filter(Boolean).map(text => ({ text, done: false }))
    : [];
  await api('/api/daily-tasks', 'POST', {
    task_date: state.selectedDate,
    project_id: projVal ? Number(projVal) : null,
    title: document.getElementById('new-task-title').value.trim(),
    priority: document.getElementById('new-task-prio').value,
    estimated_hours: Number(document.getElementById('new-task-hours').value || 1),
    notes: document.getElementById('new-task-notes').value.trim(),
    subtasks
  });
  showToast('每日任务已添加');
  await loadAll();
}

async function toggleTask(id, completed) {
  const task = state.dailyTasks.find(x => x.id === id);
  if (task && task.subtasks && task.subtasks.length > 0) {
    const subs = task.subtasks.map(it => ({ ...it, done: completed }));
    await api('/api/daily-tasks/' + id, 'PUT', { subtasks: subs });
  } else {
    await api('/api/daily-tasks/' + id, 'PUT', { is_completed: completed ? 1 : 0 });
  }
  await loadAll();
}

async function deleteTask(id) {
  await api('/api/daily-tasks/' + id, 'DELETE');
  showToast('任务已删除');
  await loadAll();
}

async function carryOverTasks() {
  const res = await api('/api/daily-tasks/carry-over', 'POST', { target_date: state.selectedDate });
  showToast(`已顺延 ${res.moved} 条未完成任务`);
  await loadAll();
}

function exportDailyMarkdown() {
  const pending = state.dailyTasks.filter(t => !t.is_completed);
  const done = state.dailyTasks.filter(t => t.is_completed);
  const lines = [`# 科研工作日报 (${state.selectedDate})`, '', '## ✅ 已完成任务'];
  done.forEach(t => {
    lines.push(`- [x] ${t.title}${t.project_title ? ` (${t.project_title})` : ''}${t.notes ? ` — ${t.notes}` : ''}`);
    (t.subtasks || []).forEach(st => lines.push(`  - [${st.done?'x':' '}] ${st.text}`));
  });
  if (!done.length) lines.push('- （无）');
  lines.push('', '## ⏳ 未完成任务');
  pending.forEach(t => {
    lines.push(`- [ ] ${t.title}${t.project_title ? ` (${t.project_title})` : ''}${t.notes ? ` — ${t.notes}` : ''}`);
    (t.subtasks || []).forEach(st => lines.push(`  - [${st.done?'x':' '}] ${st.text}`));
  });
  if (!pending.length) lines.push('- （无）');
  const blob = new Blob([lines.join('\n')], { type: 'text/markdown;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `科研工作日报_${state.selectedDate}.md`;
  a.click();
}

async function seedDemo() {
  if (!confirm('将导入示例科研项目、今日任务与意见改进记录用于体验，是否继续？')) return;
  await api('/api/seed-demo', 'POST', {});
  showToast('示例数据已导入');
  await loadAll();
}

function exportBackup() {
  window.open('/api/export-backup', '_blank');
}

loadAll();
</script>
</body>
</html>
"""


class RequestHandler(BaseHTTPRequestHandler):
    server_version = "ResearchPM/1.2"

    def log_message(self, format, *args):
        pass

    def _send_json(self, data, status=HTTPStatus.OK):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self):
        length = int(self.headers.get("Content-Length", "0") or "0")
        raw = self.rfile.read(length) if length > 0 else b"{}"
        return json.loads(raw.decode("utf-8"))

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == "/" or path == "/index.html":
            body = HTML_PAGE.encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        if path == "/api/projects":
            self._send_json(fetch_all_projects_summary())
            return

        m = re.fullmatch(r"/api/projects/(\d+)", path)
        if m:
            detail = fetch_project_detail(int(m.group(1)))
            if not detail:
                self._send_json({"error": "项目不存在"}, HTTPStatus.NOT_FOUND)
            else:
                self._send_json(detail)
            return

        if path == "/api/daily-tasks":
            qs = urllib.parse.parse_qs(parsed.query)
            date_val = qs.get("date", [today_str()])[0]
            with get_conn() as conn:
                rows = conn.execute(
                    """
                    SELECT dt.*, p.title AS project_title
                    FROM daily_tasks dt
                    LEFT JOIN projects p ON dt.project_id = p.id
                    WHERE dt.task_date = ?
                    ORDER BY dt.is_completed ASC,
                             CASE dt.priority WHEN 'high' THEN 1 WHEN 'medium' THEN 2 ELSE 3 END,
                             dt.id DESC
                    """,
                    (date_val,),
                ).fetchall()
            out = []
            for r in rows:
                d = dict(r)
                d["subtasks"] = parse_subtasks(d.get("subtasks_json"))
                out.append(d)
            self._send_json(out)
            return

        if path == "/api/feedbacks":
            with get_conn() as conn:
                rows = conn.execute(
                    """
                    SELECT * FROM feedback_items
                    ORDER BY CASE status
                               WHEN 'open' THEN 1
                               WHEN 'in_progress' THEN 2
                               WHEN 'planned' THEN 3
                               WHEN 'resolved' THEN 4
                               ELSE 5
                             END,
                             CASE priority WHEN 'high' THEN 1 WHEN 'medium' THEN 2 ELSE 3 END,
                             id DESC
                    """
                ).fetchall()
            self._send_json([dict(r) for r in rows])
            return

        if path == "/api/export-backup":
            with get_conn() as conn:
                backup = {
                    "exported_at": now_iso(),
                    "projects": [dict(r) for r in conn.execute("SELECT * FROM projects").fetchall()],
                    "project_steps": [dict(r) for r in conn.execute("SELECT * FROM project_steps").fetchall()],
                    "step_evidences": [dict(r) for r in conn.execute("SELECT * FROM step_evidences").fetchall()],
                    "daily_tasks": [dict(r) for r in conn.execute("SELECT * FROM daily_tasks").fetchall()],
                    "feedback_items": [dict(r) for r in conn.execute("SELECT * FROM feedback_items").fetchall()],
                }
            body = json.dumps(backup, ensure_ascii=False, indent=2).encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Disposition", f'attachment; filename="research_backup_{today_str()}.json"')
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        if path.startswith("/evidence/"):
            rel = urllib.parse.unquote(path[len("/evidence/"):])
            target = (EVIDENCE_DIR / rel).resolve()
            if not str(target).startswith(str(EVIDENCE_DIR.resolve())) or not target.is_file():
                self.send_error(HTTPStatus.NOT_FOUND, "File not found")
                return
            mime, _ = mimetypes.guess_type(str(target))
            data = target.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", mime or "application/octet-stream")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return

        self._send_json({"error": "Not found"}, HTTPStatus.NOT_FOUND)

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        try:
            data = self._read_json()
        except Exception as exc:
            self._send_json({"error": f"无效的 JSON 请求: {exc}"}, HTTPStatus.BAD_REQUEST)
            return

        ts = now_iso()
        if path == "/api/projects":
            title = (data.get("title") or "").strip()
            if not title:
                self._send_json({"error": "项目名称不能为空"}, HTTPStatus.BAD_REQUEST)
                return
            with get_conn() as conn:
                cur = conn.execute(
                    """
                    INSERT INTO projects (title, category, status, objectives, significance, progress_mode, manual_progress, start_date, target_date, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        title,
                        data.get("category", "科研项目"),
                        data.get("status", "active"),
                        data.get("objectives", ""),
                        data.get("significance", ""),
                        data.get("progress_mode", "auto"),
                        int(data.get("manual_progress", 0) or 0),
                        data.get("start_date", ""),
                        data.get("target_date", ""),
                        ts,
                        ts,
                    ),
                )
                conn.commit()
                self._send_json({"id": cur.lastrowid}, HTTPStatus.CREATED)
            return

        if path == "/api/steps":
            title = (data.get("title") or "").strip()
            project_id = int(data.get("project_id"))
            if not title:
                self._send_json({"error": "步骤名称不能为空"}, HTTPStatus.BAD_REQUEST)
                return
            subs = parse_subtasks(data.get("subtasks"))
            status = data.get("status", "todo")
            prog = int(data.get("progress", 0) or 0)
            if subs:
                done_cnt = sum(1 for x in subs if x["done"])
                prog = round(done_cnt * 100 / len(subs))
                status = "done" if done_cnt == len(subs) else ("todo" if done_cnt == 0 else "in_progress")
            with get_conn() as conn:
                cur = conn.execute(
                    """
                    INSERT INTO project_steps (project_id, step_order, title, description, status, progress, due_date, subtasks_json, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        project_id,
                        int(data.get("step_order", 1)),
                        title,
                        data.get("description", ""),
                        status,
                        prog,
                        data.get("due_date", ""),
                        json.dumps(subs, ensure_ascii=False),
                        ts,
                    ),
                )
                conn.execute("UPDATE projects SET updated_at = ? WHERE id = ?", (ts, project_id))
                conn.commit()
                self._send_json({"id": cur.lastrowid}, HTTPStatus.CREATED)
            return

        if path == "/api/evidences":
            title = (data.get("title") or "").strip()
            project_id = int(data.get("project_id"))
            step_id = int(data.get("step_id"))
            if not title:
                self._send_json({"error": "证据标题不能为空"}, HTTPStatus.BAD_REQUEST)
                return

            file_rel_path = ""
            orig_name = data.get("filename") or ""
            b64 = data.get("content_base64") or ""
            if orig_name and b64:
                proj_dir = EVIDENCE_DIR / f"proj_{project_id}"
                proj_dir.mkdir(parents=True, exist_ok=True)
                safe_name = f"{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{sanitize_filename(orig_name)}"
                dest = proj_dir / safe_name
                dest.write_bytes(base64.b64decode(b64))
                file_rel_path = f"proj_{project_id}/{safe_name}"

            ev_type = "file" if file_rel_path else ("path" if data.get("local_path") else "data_note")
            with get_conn() as conn:
                cur = conn.execute(
                    """
                    INSERT INTO step_evidences (step_id, project_id, title, evidence_type, file_rel_path, original_filename, local_path, url, data_summary, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        step_id,
                        project_id,
                        title,
                        ev_type,
                        file_rel_path,
                        orig_name,
                        data.get("local_path", ""),
                        data.get("url", ""),
                        data.get("data_summary", ""),
                        ts,
                    ),
                )
                conn.execute("UPDATE projects SET updated_at = ? WHERE id = ?", (ts, project_id))
                conn.commit()
                self._send_json({"id": cur.lastrowid}, HTTPStatus.CREATED)
            return

        if path == "/api/daily-tasks":
            title = (data.get("title") or "").strip()
            if not title:
                self._send_json({"error": "任务内容不能为空"}, HTTPStatus.BAD_REQUEST)
                return
            subs = parse_subtasks(data.get("subtasks"))
            is_comp = 1 if (subs and all(x["done"] for x in subs)) else 0
            comp_at = ts if is_comp else ""
            with get_conn() as conn:
                cur = conn.execute(
                    """
                    INSERT INTO daily_tasks (task_date, project_id, title, priority, estimated_hours, notes, subtasks_json, is_completed, completed_at, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        data.get("task_date") or today_str(),
                        data.get("project_id"),
                        title,
                        data.get("priority", "medium"),
                        float(data.get("estimated_hours", 1.0) or 1.0),
                        data.get("notes", ""),
                        json.dumps(subs, ensure_ascii=False),
                        is_comp,
                        comp_at,
                        ts,
                    ),
                )
                conn.commit()
                self._send_json({"id": cur.lastrowid}, HTTPStatus.CREATED)
            return

        if path == "/api/daily-tasks/carry-over":
            target_date = data.get("target_date") or today_str()
            with get_conn() as conn:
                cur = conn.execute(
                    "UPDATE daily_tasks SET task_date = ? WHERE is_completed = 0 AND task_date < ?",
                    (target_date, target_date),
                )
                conn.commit()
                self._send_json({"moved": cur.rowcount})
            return

        if path == "/api/feedbacks":
            title = (data.get("title") or "").strip()
            if not title:
                self._send_json({"error": "意见标题不能为空"}, HTTPStatus.BAD_REQUEST)
                return
            status = data.get("status", "open")
            resolved_at = ts if status == "resolved" else ""
            with get_conn() as conn:
                cur = conn.execute(
                    """
                    INSERT INTO feedback_items (title, module, feedback_type, priority, status, description, resolution_note, target_version, created_at, updated_at, resolved_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        title,
                        data.get("module", "科研项目详情"),
                        data.get("feedback_type", "ux"),
                        data.get("priority", "medium"),
                        status,
                        data.get("description", ""),
                        data.get("resolution_note", ""),
                        data.get("target_version", ""),
                        ts,
                        ts,
                        resolved_at,
                    ),
                )
                conn.commit()
                self._send_json({"id": cur.lastrowid}, HTTPStatus.CREATED)
            return

        if path == "/api/open-local":
            raw_path = (data.get("path") or "").strip()
            target = DATA_DIR if raw_path == "__DATA_DIR__" else Path(raw_path)
            if not target.exists():
                self._send_json({"error": f"路径不存在: {target}"}, HTTPStatus.NOT_FOUND)
                return
            if platform.system() == "Windows":
                os.startfile(str(target))  # type: ignore[attr-defined]
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", str(target)])
            else:
                subprocess.Popen(["xdg-open", str(target)])
            self._send_json({"ok": True})
            return

        if path == "/api/seed-demo":
            seed_sample_data()
            self._send_json({"ok": True})
            return

        self._send_json({"error": "Not found"}, HTTPStatus.NOT_FOUND)

    def do_PUT(self):
        path = urllib.parse.urlparse(self.path).path
        data = self._read_json()
        ts = now_iso()

        m = re.fullmatch(r"/api/projects/(\d+)", path)
        if m:
            pid = int(m.group(1))
            with get_conn() as conn:
                row = conn.execute("SELECT * FROM projects WHERE id=?", (pid,)).fetchone()
                if not row:
                    self._send_json({"error": "项目不存在"}, HTTPStatus.NOT_FOUND)
                    return
                conn.execute(
                    """
                    UPDATE projects
                    SET title=?, category=?, status=?, objectives=?, significance=?,
                        progress_mode=?, manual_progress=?, start_date=?, target_date=?, updated_at=?
                    WHERE id=?
                    """,
                    (
                        data.get("title", row["title"]),
                        data.get("category", row["category"]),
                        data.get("status", row["status"]),
                        data.get("objectives", row["objectives"]),
                        data.get("significance", row["significance"]),
                        data.get("progress_mode", row["progress_mode"]),
                        int(data.get("manual_progress", row["manual_progress"]) or 0),
                        data.get("start_date", row["start_date"]),
                        data.get("target_date", row["target_date"]),
                        ts,
                        pid,
                    ),
                )
                conn.commit()
            self._send_json({"ok": True})
            return

        m = re.fullmatch(r"/api/steps/(\d+)", path)
        if m:
            sid = int(m.group(1))
            with get_conn() as conn:
                row = conn.execute("SELECT * FROM project_steps WHERE id=?", (sid,)).fetchone()
                if not row:
                    self._send_json({"error": "步骤不存在"}, HTTPStatus.NOT_FOUND)
                    return
                if "subtasks" in data:
                    subs = parse_subtasks(data.get("subtasks"))
                    subs_json = json.dumps(subs, ensure_ascii=False)
                else:
                    subs = parse_subtasks(row["subtasks_json"])
                    subs_json = row["subtasks_json"] or "[]"

                new_status = data.get("status", row["status"])
                new_prog = int(data.get("progress", row["progress"]) or 0)
                if "subtasks" in data and subs:
                    done_cnt = sum(1 for x in subs if x["done"])
                    new_prog = round(done_cnt * 100 / len(subs))
                    new_status = "done" if done_cnt == len(subs) else ("todo" if done_cnt == 0 else "in_progress")

                conn.execute(
                    """
                    UPDATE project_steps
                    SET step_order=?, title=?, description=?, status=?, progress=?, due_date=?, subtasks_json=?
                    WHERE id=?
                    """,
                    (
                        int(data.get("step_order", row["step_order"])),
                        data.get("title", row["title"]),
                        data.get("description", row["description"]),
                        new_status,
                        new_prog,
                        data.get("due_date", row["due_date"]),
                        subs_json,
                        sid,
                    ),
                )
                conn.execute("UPDATE projects SET updated_at=? WHERE id=?", (ts, row["project_id"]))
                conn.commit()
            self._send_json({"ok": True})
            return

        m = re.fullmatch(r"/api/evidences/(\d+)", path)
        if m:
            eid = int(m.group(1))
            with get_conn() as conn:
                row = conn.execute("SELECT * FROM step_evidences WHERE id=?", (eid,)).fetchone()
                if not row:
                    self._send_json({"error": "证据不存在"}, HTTPStatus.NOT_FOUND)
                    return
                conn.execute(
                    "UPDATE step_evidences SET title=?, data_summary=?, local_path=?, url=? WHERE id=?",
                    (
                        data.get("title", row["title"]),
                        data.get("data_summary", row["data_summary"]),
                        data.get("local_path", row["local_path"]),
                        data.get("url", row["url"]),
                        eid,
                    ),
                )
                conn.commit()
            self._send_json({"ok": True})
            return

        m = re.fullmatch(r"/api/daily-tasks/(\d+)", path)
        if m:
            tid = int(m.group(1))
            with get_conn() as conn:
                row = conn.execute("SELECT * FROM daily_tasks WHERE id=?", (tid,)).fetchone()
                if not row:
                    self._send_json({"error": "任务不存在"}, HTTPStatus.NOT_FOUND)
                    return
                if "subtasks" in data:
                    subs = parse_subtasks(data.get("subtasks"))
                    subs_json = json.dumps(subs, ensure_ascii=False)
                    if subs:
                        is_comp = 1 if all(x["done"] for x in subs) else 0
                    else:
                        is_comp = 1 if data.get("is_completed", row["is_completed"]) else 0
                else:
                    subs_json = row["subtasks_json"] or "[]"
                    is_comp = 1 if data.get("is_completed", row["is_completed"]) else 0

                comp_at = ts if is_comp and not row["is_completed"] else ("" if not is_comp else row["completed_at"])
                conn.execute(
                    """
                    UPDATE daily_tasks
                    SET title=?, priority=?, estimated_hours=?, notes=?, subtasks_json=?, is_completed=?, completed_at=?
                    WHERE id=?
                    """,
                    (
                        data.get("title", row["title"]),
                        data.get("priority", row["priority"]),
                        float(data.get("estimated_hours", row["estimated_hours"]) or 1.0),
                        data.get("notes", row["notes"]),
                        subs_json,
                        is_comp,
                        comp_at,
                        tid,
                    ),
                )
                conn.commit()
            self._send_json({"ok": True})
            return

        m = re.fullmatch(r"/api/feedbacks/(\d+)", path)
        if m:
            fid = int(m.group(1))
            with get_conn() as conn:
                row = conn.execute("SELECT * FROM feedback_items WHERE id=?", (fid,)).fetchone()
                if not row:
                    self._send_json({"error": "记录不存在"}, HTTPStatus.NOT_FOUND)
                    return
                new_status = data.get("status", row["status"])
                resolved_at = ts if new_status == "resolved" and row["status"] != "resolved" else ("" if new_status != "resolved" else row["resolved_at"])
                conn.execute(
                    """
                    UPDATE feedback_items
                    SET title=?, module=?, feedback_type=?, priority=?, status=?,
                        description=?, resolution_note=?, target_version=?, updated_at=?, resolved_at=?
                    WHERE id=?
                    """,
                    (
                        data.get("title", row["title"]),
                        data.get("module", row["module"]),
                        data.get("feedback_type", row["feedback_type"]),
                        data.get("priority", row["priority"]),
                        new_status,
                        data.get("description", row["description"]),
                        data.get("resolution_note", row["resolution_note"]),
                        data.get("target_version", row["target_version"]),
                        ts,
                        resolved_at,
                        fid,
                    ),
                )
                conn.commit()
            self._send_json({"ok": True})
            return

        self._send_json({"error": "Not found"}, HTTPStatus.NOT_FOUND)

    def do_DELETE(self):
        path = urllib.parse.urlparse(self.path).path

        m = re.fullmatch(r"/api/projects/(\d+)", path)
        if m:
            pid = int(m.group(1))
            with get_conn() as conn:
                conn.execute("DELETE FROM projects WHERE id=?", (pid,))
                conn.commit()
            proj_dir = EVIDENCE_DIR / f"proj_{pid}"
            if proj_dir.exists():
                shutil.rmtree(proj_dir, ignore_errors=True)
            self._send_json({"ok": True})
            return

        m = re.fullmatch(r"/api/steps/(\d+)", path)
        if m:
            with get_conn() as conn:
                conn.execute("DELETE FROM project_steps WHERE id=?", (int(m.group(1)),))
                conn.commit()
            self._send_json({"ok": True})
            return

        m = re.fullmatch(r"/api/evidences/(\d+)", path)
        if m:
            eid = int(m.group(1))
            with get_conn() as conn:
                row = conn.execute("SELECT file_rel_path FROM step_evidences WHERE id=?", (eid,)).fetchone()
                if row and row["file_rel_path"]:
                    f = EVIDENCE_DIR / row["file_rel_path"]
                    if f.is_file():
                        f.unlink(missing_ok=True)
                conn.execute("DELETE FROM step_evidences WHERE id=?", (eid,))
                conn.commit()
            self._send_json({"ok": True})
            return

        m = re.fullmatch(r"/api/daily-tasks/(\d+)", path)
        if m:
            with get_conn() as conn:
                conn.execute("DELETE FROM daily_tasks WHERE id=?", (int(m.group(1)),))
                conn.commit()
            self._send_json({"ok": True})
            return

        m = re.fullmatch(r"/api/feedbacks/(\d+)", path)
        if m:
            with get_conn() as conn:
                conn.execute("DELETE FROM feedback_items WHERE id=?", (int(m.group(1)),))
                conn.commit()
            self._send_json({"ok": True})
            return

        self._send_json({"error": "Not found"}, HTTPStatus.NOT_FOUND)


def find_free_port(preferred=8765):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        if s.connect_ex(("127.0.0.1", preferred)) != 0:
            return preferred
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def launch_windows_ui(url):
    """优先使用 Windows 自带的 Edge App 模式打开（无地址栏，类原生桌面窗口）"""
    if platform.system() == "Windows":
        edge_candidates = [
            Path(os.environ.get("ProgramFiles(x86)", "")) / "Microsoft/Edge/Application/msedge.exe",
            Path(os.environ.get("ProgramFiles", "")) / "Microsoft/Edge/Application/msedge.exe",
            Path(os.environ.get("LocalAppData", "")) / "Microsoft/Edge/Application/msedge.exe",
        ]
        for exe in edge_candidates:
            if exe.is_file():
                subprocess.Popen([str(exe), f"--app={url}"])
                return
    webbrowser.open(url)


def main():
    init_db()
    port = find_free_port(8765)
    url = f"http://127.0.0.1:{port}"
    server = ThreadingHTTPServer(("127.0.0.1", port), RequestHandler)
    print(f"科研项目管理系统已启动: {url}")
    print(f"本地数据库位置: {DB_PATH}")
    print("关闭此控制台窗口或按 Ctrl+C 即可退出。")

    threading.Timer(0.6, launch_windows_ui, args=(url,)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n正在安全关闭服务...")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()