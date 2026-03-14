"""
MS Project MCP Server
Controls local Microsoft Project via COM automation.
Install: pip install mcp
Run:     python server.py
Register in claude_desktop_config.json (see bottom of file).
"""

import json
import datetime
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("MS Project")

# ---------------------------------------------------------------------------
# COM helpers
# ---------------------------------------------------------------------------

def get_app(require_project=True):
    """Get running MS Project instance. Raises if not running."""
    import win32com.client
    try:
        app = win32com.client.GetActiveObject("MSProject.Application")
    except Exception:
        raise RuntimeError(
            "MS Project is not running. Open MS Project and load a file first."
        )
    if require_project and app.Projects.Count == 0:
        raise RuntimeError(
            "No project file is open in MS Project. Please open a file first."
        )
    return app


def get_proj(app):
    return app.ActiveProject


def _get_mpd(proj):
    """Get MinutesPerDay safely (falls back to 480 for freshly-imported XML)."""
    try:
        return proj.MinutesPerDay
    except Exception:
        return 480


def _parse_date(s):
    """Parse YYYY-MM-DD string to datetime for COM. Returns None if empty."""
    if not s:
        return None
    return datetime.datetime.strptime(s, "%Y-%m-%d")


def task_to_dict(t, proj):
    """Convert a COM Task object to a plain dict."""
    mpd = _get_mpd(proj)

    def fmt(dt):
        try:
            if dt is None:
                return None
            return str(dt)[:19]
        except Exception:
            return None

    # Reverse map for constraint type integers
    CONSTRAINT_NAMES = {
        0: "ASAP", 1: "ALAP", 2: "MSO", 3: "MFO",
        4: "SNET", 5: "SNLT", 6: "FNET", 7: "FNLT",
    }
    # Reverse map for task type integers
    TASK_TYPE_NAMES = {0: "FixedUnits", 1: "FixedDuration", 2: "FixedWork"}

    # Safe reads for fields that may not be available on all task types
    def safe(prop, default=None):
        try:
            return prop
        except Exception:
            return default

    return {
        "unique_id":              t.UniqueID,
        "id":                     t.ID,
        "name":                   t.Name,
        "outline_level":          t.OutlineLevel,
        "wbs":                    t.WBS,
        "summary":                bool(t.Summary),
        "milestone":              bool(t.Milestone),
        "start":                  fmt(t.Start),
        "finish":                 fmt(t.Finish),
        "duration_days":          round(t.Duration / mpd, 2) if t.Duration else 0,
        "percent_complete":       t.PercentComplete,
        "actual_start":           fmt(safe(t.ActualStart)),
        "actual_finish":          fmt(safe(t.ActualFinish)),
        "remaining_duration_days": round(safe(t.RemainingDuration, 0) / mpd, 2),
        "total_slack_days":       round(safe(t.TotalSlack, 0) / mpd, 2),
        "free_slack_days":        round(safe(t.FreeSlack, 0) / mpd, 2),
        "deadline":               fmt(safe(t.Deadline)),
        "priority":               safe(t.Priority, 500),
        "constraint_type":        CONSTRAINT_NAMES.get(safe(t.ConstraintType, 0), "ASAP"),
        "constraint_date":        fmt(safe(t.ConstraintDate)),
        "manual":                 bool(safe(t.Manual, False)),
        "type":                   TASK_TYPE_NAMES.get(safe(t.Type, 0), "FixedUnits"),
        "predecessors":           t.Predecessors,
        "resource_names":         t.ResourceNames,
        "notes":                  t.Notes,
        "critical":               bool(t.Critical),
        "active":                 bool(t.Active),
        "rag":                    t.Text1 or "",
        "text1":                  t.Text1 or "",
        "text2":                  t.Text2 or "",
        "text3":                  t.Text3 or "",
        "flag1":                  bool(t.Flag1),
        "flag2":                  bool(t.Flag2),
        "hyperlink":              safe(t.HyperlinkAddress, "") or "",
        "hyperlink_text":         safe(t.Hyperlink, "") or "",
    }


def _count_resources(proj):
    """Count resources safely — returns 0 if resource pool is empty/inaccessible."""
    try:
        return sum(1 for r in proj.Resources if r is not None)
    except Exception:
        return 0


def _fmt_date(dt):
    """Format a COM date to 'YYYY-MM-DD' string. Returns None on failure."""
    try:
        return str(dt)[:10] if dt else None
    except Exception:
        return None


def _to_naive(dt):
    """Strip timezone from COM datetime for safe comparison with datetime.now()."""
    if dt is None:
        return None
    try:
        if hasattr(dt, 'tzinfo') and dt.tzinfo is not None:
            return dt.replace(tzinfo=None)
    except Exception:
        pass
    return dt


def _find_task(proj, unique_id):
    """Find a task by UniqueID. Returns the COM Task object or None."""
    for t in proj.Tasks:
        if t is not None and t.UniqueID == unique_id:
            return t
    return None


def _custom_field_id(field_name):
    """
    Map field name like 'Text5', 'Number1', 'Flag3', 'Date1', 'Duration2'
    to the COM pjCustomTask* field ID constant.
    Returns (field_id, field_type) or raises ValueError.
    """
    name = field_name.strip()
    lower = name.lower()

    # COM field ID bases (pjCustomTask* constants)
    bases = {
        "text":     (188743731, 30),   # Text1-30
        "number":   (188743767, 20),   # Number1-20
        "date":     (188743945, 10),   # Date1-10
        "flag":     (188743752, 20),   # Flag1-20
        "duration": (188743783, 10),   # Duration1-10
    }

    for prefix, (base_id, max_n) in bases.items():
        if lower.startswith(prefix):
            num_str = lower[len(prefix):]
            if num_str.isdigit():
                num = int(num_str)
                if 1 <= num <= max_n:
                    return base_id + (num - 1), prefix
    raise ValueError(f"Unknown custom field: '{field_name}'. Use Text1-30, Number1-20, Date1-10, Flag1-20, Duration1-10.")


# ---------------------------------------------------------------------------
# TOOLS — Project management
# ---------------------------------------------------------------------------

@mcp.tool()
def open_project(file_path: str) -> str:
    """
    Open a Microsoft Project file (.mpp or .xml).
    MS Project must already be running (it is launched automatically if not).
    """
    import win32com.client
    try:
        app = win32com.client.GetActiveObject("MSProject.Application")
    except Exception:
        app = win32com.client.Dispatch("MSProject.Application")
        app.Visible = True
        app.DisplayAlerts = False

    app.FileOpen(file_path)
    proj = app.ActiveProject
    return json.dumps({
        "status":     "opened",
        "name":       proj.Name,
        "full_path":  proj.FullName,
        "task_count": proj.Tasks.Count,
        "start":      str(proj.ProjectStart)[:10],
        "finish":     str(proj.ProjectFinish)[:10],
    }, indent=2)


@mcp.tool()
def new_project(title: str = "New Project", start: str = "") -> str:
    """
    Create a new blank project without needing a file on disk.

    Args:
        title: Project title (default "New Project").
        start: Project start date as YYYY-MM-DD (optional).
    """
    import win32com.client
    try:
        app = win32com.client.GetActiveObject("MSProject.Application")
    except Exception:
        app = win32com.client.Dispatch("MSProject.Application")
        app.Visible = True
        app.DisplayAlerts = False

    app.FileNew()
    proj = app.ActiveProject
    proj.Title = title
    if start:
        proj.ProjectStart = _parse_date(start)

    return json.dumps({
        "status": "created",
        "title":  proj.Title,
        "name":   proj.Name,
        "start":  str(proj.ProjectStart)[:10],
    }, indent=2)


@mcp.tool()
def get_project_info() -> str:
    """Get summary information about the currently active project."""
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    def fmt(dt):
        try:
            return str(dt)[:10] if dt else None
        except Exception:
            return None

    task_count    = sum(1 for t in proj.Tasks if t is not None)
    summary_count = sum(1 for t in proj.Tasks if t is not None and t.Summary)
    mile_count    = sum(1 for t in proj.Tasks if t is not None and t.Milestone)
    critical      = sum(1 for t in proj.Tasks if t is not None and t.Critical and not t.Summary)

    # Safe reads for optional metadata
    def safe_read(attr):
        try:
            return getattr(proj, attr, "") or ""
        except Exception:
            return ""

    return json.dumps({
        "name":            proj.Name,
        "full_path":       proj.FullName,
        "title":           safe_read("Title"),
        "manager":         safe_read("Manager"),
        "company":         safe_read("Company"),
        "author":          safe_read("Author"),
        "subject":         safe_read("Subject"),
        "start":           fmt(proj.ProjectStart),
        "finish":          fmt(proj.ProjectFinish),
        "status_date":     fmt(getattr(proj, "StatusDate", None)),
        "calendar":        str(proj.Calendar) if proj.Calendar else "",
        "tasks_total":     task_count,
        "summary_tasks":   summary_count,
        "milestones":      mile_count,
        "critical_tasks":  critical,
        "resources":       _count_resources(proj),
        "minutes_per_day": mpd,
    }, indent=2)


@mcp.tool()
def set_project_properties(properties_json: str) -> str:
    """
    Set project metadata properties.

    Args:
        properties_json: JSON string with fields to set. All optional:
            title, manager, company, author, subject, status_date (YYYY-MM-DD),
            start (YYYY-MM-DD).
            Example: '{"title": "EXPO 2030", "manager": "John", "company": "ERC"}'
    """
    props = json.loads(properties_json)
    app   = get_app()
    proj  = get_proj(app)

    changed = []
    if "title" in props:
        proj.Title = props["title"];       changed.append("title")
    if "manager" in props:
        proj.Manager = props["manager"];   changed.append("manager")
    if "company" in props:
        proj.Company = props["company"];   changed.append("company")
    if "author" in props:
        proj.Author = props["author"];     changed.append("author")
    if "subject" in props:
        proj.Subject = props["subject"];   changed.append("subject")
    if "start" in props and props["start"]:
        proj.ProjectStart = _parse_date(props["start"]); changed.append("start")
    if "status_date" in props and props["status_date"]:
        proj.StatusDate = _parse_date(props["status_date"]); changed.append("status_date")

    app.FileSave()
    return json.dumps({"status": "updated", "changed": changed}, indent=2)


@mcp.tool()
def save_project() -> str:
    """Save the active project (in place)."""
    app = get_app()
    app.FileSave()
    return "Project saved."


@mcp.tool()
def save_project_as(file_path: str, format: str = "mpp") -> str:
    """
    Save the active project to a new path.
    format: 'mpp' (default), 'xml', 'csv'
    """
    fmt_map = {"mpp": 0, "xml": 22, "csv": 23}
    fmt_id  = fmt_map.get(format.lower(), 0)
    app     = get_app()
    app.FileSaveAs(Name=file_path, Format=fmt_id, Backup=False, ReadOnly=False)
    return f"Project saved as: {file_path}"


@mcp.tool()
def close_project(save: bool = False) -> str:
    """Close the active project. Set save=True to save before closing."""
    app = get_app()
    app.FileClose(Save=1 if save else 0)
    return "Project closed."


# ---------------------------------------------------------------------------
# TOOLS — Reading tasks
# ---------------------------------------------------------------------------

@mcp.tool()
def get_tasks(
    include_summary: bool = False,
    outline_level: int = 0,
    keyword: str = ""
) -> str:
    """
    Get all tasks from the active project.

    Args:
        include_summary: Include summary/parent tasks (default False).
        outline_level:   Filter to a specific outline level (0 = all).
        keyword:         Filter tasks whose name contains this string (case-insensitive).
    """
    app  = get_app()
    proj = get_proj(app)

    results = []
    for t in proj.Tasks:
        if t is None:
            continue
        if not include_summary and t.Summary:
            continue
        if outline_level > 0 and t.OutlineLevel != outline_level:
            continue
        if keyword and keyword.lower() not in t.Name.lower():
            continue
        results.append(task_to_dict(t, proj))

    return json.dumps({"count": len(results), "tasks": results}, indent=2)


@mcp.tool()
def get_task(unique_id: int) -> str:
    """Get full details for a single task by its UniqueID."""
    app  = get_app()
    proj = get_proj(app)

    for t in proj.Tasks:
        if t is not None and t.UniqueID == unique_id:
            return json.dumps(task_to_dict(t, proj), indent=2)

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


@mcp.tool()
def get_critical_path() -> str:
    """Return all tasks on the critical path (non-summary)."""
    app  = get_app()
    proj = get_proj(app)

    results = []
    for t in proj.Tasks:
        if t is not None and t.Critical and not t.Summary:
            results.append(task_to_dict(t, proj))

    return json.dumps({"count": len(results), "tasks": results}, indent=2)


@mcp.tool()
def get_tasks_by_rag(rag: str = "Red") -> str:
    """
    Return tasks filtered by RAG status stored in the Text1 custom field.
    rag: 'Red', 'Amber', or 'Green'
    """
    app  = get_app()
    proj = get_proj(app)

    results = []
    for t in proj.Tasks:
        if t is not None and not t.Summary:
            if (t.Text1 or "").strip().lower() == rag.strip().lower():
                results.append(task_to_dict(t, proj))

    return json.dumps({"rag": rag, "count": len(results), "tasks": results}, indent=2)


@mcp.tool()
def get_overdue_tasks() -> str:
    """Return incomplete tasks whose Finish date is in the past."""
    import datetime
    today = datetime.datetime.now()
    app   = get_app()
    proj  = get_proj(app)

    results = []
    for t in proj.Tasks:
        if t is None or t.Summary or t.Milestone:
            continue
        if t.PercentComplete >= 100:
            continue
        try:
            finish = _to_naive(t.Finish)
            if finish and finish < today:
                results.append(task_to_dict(t, proj))
        except Exception:
            continue

    return json.dumps({"count": len(results), "tasks": results}, indent=2)


@mcp.tool()
def get_tasks_by_resource(resource_name: str) -> str:
    """Return all tasks assigned to a named resource (case-insensitive substring match)."""
    app  = get_app()
    proj = get_proj(app)

    results = []
    name_lower = resource_name.lower()
    for t in proj.Tasks:
        if t is not None and not t.Summary:
            if name_lower in (t.ResourceNames or "").lower():
                results.append(task_to_dict(t, proj))

    return json.dumps({
        "resource": resource_name,
        "count":    len(results),
        "tasks":    results,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Modifying tasks
# ---------------------------------------------------------------------------

@mcp.tool()
def update_task(
    unique_id:        int,
    name:             str = "",
    percent_complete: int = -1,
    notes:            str = "",
    start:            str = "",
    finish:           str = "",
    duration_days:    int = -1,
    manual:           bool = None,
    rag:              str = "",
    text2:            str = "",
    text3:            str = "",
    flag1:            bool = None,
    flag2:            bool = None,
    priority:         int = -1,
    task_type:        str = "",
) -> str:
    """
    Update one or more properties of a task identified by UniqueID.
    Only the fields you provide are changed.

    Args:
        unique_id:        Task UniqueID (required).
        name:             New task name.
        percent_complete: 0-100.
        notes:            Free-text notes.
        start:            Start date as YYYY-MM-DD.
        finish:           Finish date as YYYY-MM-DD.
        duration_days:    Duration in working days (0+ to set).
        manual:           True for manually scheduled, False for auto-scheduled.
        rag:              RAG status: 'Red', 'Amber', or 'Green' (stored in Text1).
        text2:            Custom Text2 field.
        text3:            Custom Text3 field.
        flag1:            Custom Flag1 boolean.
        flag2:            Custom Flag2 boolean.
        priority:         Leveling priority 0-1000 (default 500). 0+ to set.
        task_type:        'FixedUnits', 'FixedDuration', or 'FixedWork'.
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    for t in proj.Tasks:
        if t is None or t.UniqueID != unique_id:
            continue

        changed = []
        if name:
            t.Name = name;              changed.append("name")
        if percent_complete >= 0:
            t.PercentComplete = percent_complete; changed.append("percent_complete")
        if notes:
            t.Notes = notes;            changed.append("notes")
        if start:
            t.Start = _parse_date(start); changed.append("start")
        if finish:
            t.Finish = _parse_date(finish); changed.append("finish")
        if duration_days >= 0:
            t.Duration = duration_days * mpd; changed.append("duration_days")
        if manual is not None:
            t.Manual = manual;          changed.append("manual")
        if rag:
            t.Text1 = rag;              changed.append("rag/text1")
        if text2:
            t.Text2 = text2;            changed.append("text2")
        if text3:
            t.Text3 = text3;            changed.append("text3")
        if flag1 is not None:
            t.Flag1 = flag1;            changed.append("flag1")
        if flag2 is not None:
            t.Flag2 = flag2;            changed.append("flag2")
        if priority >= 0:
            t.Priority = priority;      changed.append("priority")
        if task_type:
            TYPE_MAP = {"fixedunits": 0, "fixedduration": 1, "fixedwork": 2}
            tt = TYPE_MAP.get(task_type.lower())
            if tt is not None:
                t.Type = tt;            changed.append("type")

        app.FileSave()
        return json.dumps({
            "status":   "updated",
            "unique_id": unique_id,
            "name":     t.Name,
            "changed":  changed,
        }, indent=2)

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


@mcp.tool()
def bulk_update_rag(updates: str) -> str:
    """
    Update RAG status for multiple tasks at once.

    Args:
        updates: JSON string — list of {unique_id, rag} objects.
                 Example: '[{"unique_id": 42, "rag": "Red"}, {"unique_id": 55, "rag": "Green"}]'
    """
    items = json.loads(updates)
    app   = get_app()
    proj  = get_proj(app)

    uid_map = {t.UniqueID: t for t in proj.Tasks if t is not None}

    results = []
    for item in items:
        uid = item["unique_id"]
        rag = item["rag"]
        if uid in uid_map:
            uid_map[uid].Text1 = rag
            results.append({"unique_id": uid, "status": "updated", "rag": rag})
        else:
            results.append({"unique_id": uid, "status": "not_found"})

    app.FileSave()
    return json.dumps({"updated": len([r for r in results if r["status"] == "updated"]),
                       "results": results}, indent=2)


@mcp.tool()
def bulk_update_tasks(updates_json: str) -> str:
    """
    Update multiple tasks in one call. Suspends auto-calc for performance.

    Args:
        updates_json: JSON string — list of objects with fields:
            unique_id (required), name, start, finish, duration_days,
            percent_complete, rag, text2, text3, notes, manual (bool).
            Example: '[{"unique_id": 42, "rag": "Red", "percent_complete": 50}]'
    """
    items = json.loads(updates_json)
    app   = get_app()
    proj  = get_proj(app)
    mpd   = _get_mpd(proj)

    app.Calculation = 0
    try:
        uid_map = {t.UniqueID: t for t in proj.Tasks if t is not None}
        updated = 0
        not_found = []

        for item in items:
            uid = item["unique_id"]
            t = uid_map.get(uid)
            if t is None:
                not_found.append(uid)
                continue

            if "name" in item and item["name"]:
                t.Name = item["name"]
            if "start" in item and item["start"]:
                t.Start = _parse_date(item["start"])
            if "finish" in item and item["finish"]:
                t.Finish = _parse_date(item["finish"])
            if "duration_days" in item and item["duration_days"] is not None and item["duration_days"] >= 0:
                t.Duration = item["duration_days"] * mpd
            if "percent_complete" in item and item["percent_complete"] is not None and item["percent_complete"] >= 0:
                t.PercentComplete = item["percent_complete"]
            if "rag" in item and item["rag"]:
                t.Text1 = item["rag"]
            if "text2" in item and item["text2"]:
                t.Text2 = item["text2"]
            if "text3" in item and item["text3"]:
                t.Text3 = item["text3"]
            if "notes" in item and item["notes"]:
                t.Notes = item["notes"]
            if "manual" in item and item["manual"] is not None:
                t.Manual = item["manual"]

            updated += 1
    finally:
        app.CalculateProject()
        app.Calculation = -1

    app.FileSave()
    return json.dumps({
        "updated":   updated,
        "not_found": not_found,
    }, indent=2)


@mcp.tool()
def add_task(
    name:          str,
    outline_level: int  = 1,
    start:         str  = "",
    finish:        str  = "",
    duration_days: int  = 1,
    milestone:     bool = False,
    notes:         str  = "",
    resource:      str  = "",
    rag:           str  = "",
    after_unique_id: int = 0,
) -> str:
    """
    Add a new task to the active project.

    Args:
        name:            Task name (required).
        outline_level:   WBS level (1 = top-level, 2 = sub-task, etc.).
        start:           Start date YYYY-MM-DD (optional).
        finish:          Finish date YYYY-MM-DD (optional).
        duration_days:   Duration in days (default 1).
        milestone:       True to create as a milestone.
        notes:           Free-text notes.
        resource:        Resource name to assign.
        rag:             RAG status stored in Text1.
        after_unique_id: Insert after this task's UniqueID (0 = append at end).
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    app.Calculation = 0

    try:
        task = proj.Tasks.Add(name)
        task.OutlineLevel  = outline_level
        task.Milestone     = milestone
        if milestone:
            task.Duration = 0
        else:
            task.Duration = max(duration_days, 1) * mpd

        if start:
            task.Start = _parse_date(start)
        if finish:
            task.Finish = _parse_date(finish)
        if notes:
            task.Notes = notes
        if resource:
            task.ResourceNames = resource
        if rag:
            task.Text1 = rag

    finally:
        app.CalculateProject()
        app.Calculation = -1

    app.FileSave()
    return json.dumps({
        "status":    "created",
        "unique_id": task.UniqueID,
        "id":        task.ID,
        "name":      task.Name,
    }, indent=2)


@mcp.tool()
def bulk_add_tasks(tasks_json: str) -> str:
    """
    Add multiple tasks in one call. Critical for roadmap generation.
    Suspends auto-calc for performance. Tasks are added sequentially;
    outline_level controls WBS hierarchy.

    Args:
        tasks_json: JSON string — list of task objects with fields:
            name (required), outline_level (default 1), start, finish,
            duration_days (default 1), milestone (bool), resource,
            rag, text2, text3, notes, manual (bool).
            Example: '[{"name": "Phase 1", "outline_level": 1},
                       {"name": "Task A", "outline_level": 2, "start": "2026-04-01"}]'
    """
    tasks = json.loads(tasks_json)
    app   = get_app()
    proj  = get_proj(app)
    mpd   = _get_mpd(proj)

    app.Calculation = 0
    created = []

    try:
        for item in tasks:
            t = proj.Tasks.Add(item["name"])
            t.OutlineLevel = item.get("outline_level", 1)
            t.Milestone = item.get("milestone", False)

            dur = item.get("duration_days", 1)
            if t.Milestone:
                t.Duration = 0
            else:
                t.Duration = max(dur, 1) * mpd

            if item.get("start"):
                t.Start = _parse_date(item["start"])
            if item.get("finish"):
                t.Finish = _parse_date(item["finish"])
            if item.get("resource"):
                t.ResourceNames = item["resource"]
            if item.get("rag"):
                t.Text1 = item["rag"]
            if item.get("text2"):
                t.Text2 = item["text2"]
            if item.get("text3"):
                t.Text3 = item["text3"]
            if item.get("notes"):
                t.Notes = item["notes"]
            if item.get("manual") is not None:
                t.Manual = item["manual"]

            created.append({
                "unique_id": t.UniqueID,
                "id":        t.ID,
                "name":      t.Name,
            })
    finally:
        app.CalculateProject()
        app.Calculation = -1

    app.FileSave()
    return json.dumps({
        "created": len(created),
        "tasks":   created,
    }, indent=2)


@mcp.tool()
def delete_task(unique_id: int) -> str:
    """Delete a task by its UniqueID. This cannot be undone after save."""
    app  = get_app()
    proj = get_proj(app)

    for t in proj.Tasks:
        if t is not None and t.UniqueID == unique_id:
            task_name = t.Name
            task_id   = t.ID
            app.SelectRow(task_id, False)
            app.EditDelete()
            app.FileSave()
            return json.dumps({
                "status":    "deleted",
                "unique_id": unique_id,
                "name":      task_name,
            }, indent=2)

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


# ---------------------------------------------------------------------------
# TOOLS — Task scheduling control
# ---------------------------------------------------------------------------

@mcp.tool()
def set_task_mode(unique_id: int, manual: bool = True) -> str:
    """
    Set a task to manually or automatically scheduled.

    Args:
        unique_id: Task UniqueID (required).
        manual:    True for manually scheduled (default), False for auto-scheduled.
    """
    app  = get_app()
    proj = get_proj(app)

    for t in proj.Tasks:
        if t is not None and t.UniqueID == unique_id:
            t.Manual = manual
            app.FileSave()
            return json.dumps({
                "status":    "updated",
                "unique_id": unique_id,
                "name":      t.Name,
                "manual":    manual,
            }, indent=2)

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


@mcp.tool()
def bulk_set_task_mode(updates_json: str) -> str:
    """
    Set manual/auto schedule mode for multiple tasks, or by scope.

    Args:
        updates_json: JSON string — EITHER:
            A list of {unique_id, manual} objects:
              '[{"unique_id": 42, "manual": true}]'
            OR a scope object:
              '{"mode": "manual", "scope": "all"}'
              '{"mode": "auto", "scope": "summary"}'
              '{"mode": "manual", "scope": "non_summary"}'
    """
    data = json.loads(updates_json)
    app  = get_app()
    proj = get_proj(app)

    updated = 0

    if isinstance(data, dict) and "scope" in data:
        manual = data.get("mode", "manual").lower() == "manual"
        scope  = data.get("scope", "all").lower()

        for t in proj.Tasks:
            if t is None:
                continue
            if scope == "all":
                t.Manual = manual; updated += 1
            elif scope == "summary" and t.Summary:
                t.Manual = manual; updated += 1
            elif scope == "non_summary" and not t.Summary:
                t.Manual = manual; updated += 1

    else:
        items = data if isinstance(data, list) else [data]
        uid_map = {t.UniqueID: t for t in proj.Tasks if t is not None}
        for item in items:
            uid = item["unique_id"]
            t = uid_map.get(uid)
            if t is not None:
                t.Manual = item.get("manual", True)
                updated += 1

    app.FileSave()
    return json.dumps({"updated": updated}, indent=2)


@mcp.tool()
def set_constraint(unique_id: int, constraint_type: str = "SNET", constraint_date: str = "") -> str:
    """
    Set a scheduling constraint on a task.

    Args:
        unique_id:       Task UniqueID (required).
        constraint_type: One of: ASAP, ALAP, MSO, MFO, SNET, SNLT, FNET, FNLT (default SNET).
        constraint_date: Date as YYYY-MM-DD (required for all types except ASAP/ALAP).
    """
    CONSTRAINT_MAP = {
        "ASAP": 0, "ALAP": 1, "MSO": 2, "MFO": 3,
        "SNET": 4, "SNLT": 5, "FNET": 6, "FNLT": 7,
    }

    ct = constraint_type.upper()
    if ct not in CONSTRAINT_MAP:
        return json.dumps({"error": f"Unknown constraint type '{constraint_type}'. Use: {list(CONSTRAINT_MAP.keys())}"})

    app  = get_app()
    proj = get_proj(app)

    for t in proj.Tasks:
        if t is not None and t.UniqueID == unique_id:
            t.ConstraintType = CONSTRAINT_MAP[ct]
            if constraint_date and ct not in ("ASAP", "ALAP"):
                t.ConstraintDate = _parse_date(constraint_date)
            app.FileSave()
            return json.dumps({
                "status":          "updated",
                "unique_id":       unique_id,
                "name":            t.Name,
                "constraint_type": ct,
                "constraint_date": constraint_date or "N/A",
            }, indent=2)

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


@mcp.tool()
def clear_estimated_flags() -> str:
    """
    Remove the estimated '?' flag from all task dates in the active project.
    This cleans up the question marks that appear on dates in MS Project.
    """
    app  = get_app()
    proj = get_proj(app)

    count = 0
    for t in proj.Tasks:
        if t is not None:
            try:
                t.Estimated = False
                count += 1
            except Exception:
                pass

    app.FileSave()
    return json.dumps({"status": "cleared", "tasks_updated": count}, indent=2)


@mcp.tool()
def rename_custom_fields(fields_json: str) -> str:
    """
    Rename custom text fields (Text1-Text30) to meaningful labels.
    Uses the MS Project CustomFieldRename method.

    Args:
        fields_json: JSON string — object mapping field names to display names.
            Example: '{"text1": "RAG Status", "text2": "Technology Required"}'
            Supported fields: text1-text30.
    """
    fields = json.loads(fields_json)
    app    = get_app()
    proj   = get_proj(app)

    # pjCustomTaskText1 = 188743731, each subsequent +1
    BASE_TEXT_FIELD_ID = 188743731

    renamed = []
    for field_key, display_name in fields.items():
        key_lower = field_key.lower().strip()
        if not key_lower.startswith("text"):
            continue
        try:
            num = int(key_lower[4:])
            if num < 1 or num > 30:
                continue
            field_id = BASE_TEXT_FIELD_ID + (num - 1)
            app.CustomFieldRename(field_id, display_name)
            renamed.append({"field": field_key, "display_name": display_name})
        except Exception as e:
            renamed.append({"field": field_key, "error": str(e)})

    return json.dumps({"renamed": len([r for r in renamed if "error" not in r]),
                       "results": renamed}, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Dependencies
# ---------------------------------------------------------------------------

@mcp.tool()
def add_predecessor(
    successor_unique_id:   int,
    predecessor_unique_id: int,
    link_type:             str = "FS",
    lag_days:              int = 0,
) -> str:
    """
    Add a predecessor link between two tasks.

    Args:
        successor_unique_id:   The task that depends on the predecessor.
        predecessor_unique_id: The task that must finish/start first.
        link_type:             'FS' (default), 'SS', 'FF', or 'SF'.
        lag_days:              Lag in days (positive = lag, negative = lead).
    """
    app  = get_app()
    proj = get_proj(app)

    uid_to_id = {t.UniqueID: t.ID for t in proj.Tasks if t is not None}

    if successor_unique_id not in uid_to_id:
        return json.dumps({"error": f"Successor UniqueID {successor_unique_id} not found."})
    if predecessor_unique_id not in uid_to_id:
        return json.dumps({"error": f"Predecessor UniqueID {predecessor_unique_id} not found."})

    pred_id = uid_to_id[predecessor_unique_id]
    succ_task = None
    for t in proj.Tasks:
        if t is not None and t.UniqueID == successor_unique_id:
            succ_task = t
            break

    lag_str = ""
    if lag_days > 0:
        lag_str = f"+{lag_days}d"
    elif lag_days < 0:
        lag_str = f"{lag_days}d"

    existing = succ_task.Predecessors.strip()
    new_pred  = f"{pred_id}{link_type}{lag_str}"

    if existing:
        succ_task.Predecessors = existing + "," + new_pred
    else:
        succ_task.Predecessors = new_pred

    app.FileSave()
    return json.dumps({
        "status":       "linked",
        "successor":    successor_unique_id,
        "predecessor":  predecessor_unique_id,
        "link":         new_pred,
        "predecessors": succ_task.Predecessors,
    }, indent=2)


@mcp.tool()
def bulk_add_predecessors(links_json: str) -> str:
    """
    Add multiple predecessor links in one call.

    Args:
        links_json: JSON string — list of link objects:
            [{successor_unique_id, predecessor_unique_id, link_type (default "FS"), lag_days (default 0)}]
            Example: '[{"successor_unique_id": 10, "predecessor_unique_id": 5, "link_type": "FS"}]'
    """
    links = json.loads(links_json)
    app   = get_app()
    proj  = get_proj(app)

    uid_to_id = {t.UniqueID: t.ID for t in proj.Tasks if t is not None}
    uid_to_task = {t.UniqueID: t for t in proj.Tasks if t is not None}

    linked = 0
    errors = []

    for link in links:
        succ_uid = link["successor_unique_id"]
        pred_uid = link["predecessor_unique_id"]
        lt       = link.get("link_type", "FS")
        lag      = link.get("lag_days", 0)

        if succ_uid not in uid_to_id:
            errors.append({"successor_unique_id": succ_uid, "error": "not found"})
            continue
        if pred_uid not in uid_to_id:
            errors.append({"predecessor_unique_id": pred_uid, "error": "not found"})
            continue

        pred_id   = uid_to_id[pred_uid]
        succ_task = uid_to_task[succ_uid]

        lag_str = ""
        if lag > 0:
            lag_str = f"+{lag}d"
        elif lag < 0:
            lag_str = f"{lag}d"

        new_pred = f"{pred_id}{lt}{lag_str}"
        existing = succ_task.Predecessors.strip()

        if existing:
            succ_task.Predecessors = existing + "," + new_pred
        else:
            succ_task.Predecessors = new_pred

        linked += 1

    app.FileSave()
    return json.dumps({
        "linked": linked,
        "errors": errors,
    }, indent=2)


@mcp.tool()
def remove_predecessor(
    successor_unique_id:   int,
    predecessor_unique_id: int,
) -> str:
    """Remove a specific predecessor link from a task."""
    app  = get_app()
    proj = get_proj(app)

    uid_to_id = {t.UniqueID: t.ID for t in proj.Tasks if t is not None}

    if successor_unique_id not in uid_to_id:
        return json.dumps({"error": f"Successor UniqueID {successor_unique_id} not found."})

    pred_id   = uid_to_id.get(predecessor_unique_id)
    succ_task = None
    for t in proj.Tasks:
        if t is not None and t.UniqueID == successor_unique_id:
            succ_task = t
            break

    existing = succ_task.Predecessors.strip()
    if not existing:
        return json.dumps({"status": "no_change", "message": "Task has no predecessors."})

    parts     = [p.strip() for p in existing.split(",")]
    filtered  = [p for p in parts if not p.startswith(str(pred_id))]
    succ_task.Predecessors = ",".join(filtered)

    app.FileSave()
    return json.dumps({
        "status":              "unlinked",
        "successor":           successor_unique_id,
        "removed_predecessor": predecessor_unique_id,
        "predecessors_now":    succ_task.Predecessors,
    }, indent=2)


@mcp.tool()
def get_task_dependencies(unique_id: int) -> str:
    """Get all predecessor and successor dependencies for a task."""
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    target = None
    for t in proj.Tasks:
        if t is not None and t.UniqueID == unique_id:
            target = t
            break

    if target is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    preds = []
    try:
        for dep in target.TaskDependencies:
            if dep.To.UniqueID == unique_id:
                preds.append({
                    "unique_id": dep.From.UniqueID,
                    "name":      dep.From.Name,
                    "type":      dep.Type,
                    "lag_days":  round(dep.Lag / mpd, 2),
                })
    except Exception:
        pass

    succs = []
    try:
        for dep in target.TaskDependencies:
            if dep.From.UniqueID == unique_id:
                succs.append({
                    "unique_id": dep.To.UniqueID,
                    "name":      dep.To.Name,
                    "type":      dep.Type,
                    "lag_days":  round(dep.Lag / mpd, 2),
                })
    except Exception:
        pass

    return json.dumps({
        "task":        {"unique_id": unique_id, "name": target.Name},
        "predecessors": preds,
        "successors":   succs,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Resources
# ---------------------------------------------------------------------------

@mcp.tool()
def get_resources() -> str:
    """Get all resources in the active project."""
    app  = get_app()
    proj = get_proj(app)

    results = []
    for r in proj.Resources:
        if r is not None:
            results.append({
                "unique_id":  r.UniqueID,
                "id":         r.ID,
                "name":       r.Name,
                "initials":   r.Initials,
                "type":       r.Type,
                "max_units":  r.MaxUnits,
                "cost":       r.StandardRate,
                "task_count": r.Assignments.Count,
            })

    return json.dumps({"count": len(results), "resources": results}, indent=2)


@mcp.tool()
def add_resource(
    name:          str,
    type:          int   = 0,
    max_units:     float = 1.0,
    standard_rate: str   = "",
    cost_per_use:  float = 0.0,
) -> str:
    """
    Add a resource to the project resource pool.

    Args:
        name:          Resource name (required).
        type:          0=Work (default), 1=Material, 2=Cost.
        max_units:     Maximum allocation units (default 1.0 = 100%).
        standard_rate: Standard rate as string (e.g. "50/h", "100/d").
        cost_per_use:  Fixed cost per use (default 0).
    """
    app  = get_app()
    proj = get_proj(app)

    r = proj.Resources.Add(name)
    r.Type = type
    if type == 0:  # MaxUnits only valid for Work resources
        r.MaxUnits = max_units
    if standard_rate:
        r.StandardRate = standard_rate
    if cost_per_use > 0:
        r.CostPerUse = cost_per_use

    app.FileSave()
    return json.dumps({
        "status":    "created",
        "unique_id": r.UniqueID,
        "id":        r.ID,
        "name":      r.Name,
        "type":      r.Type,
        "max_units": r.MaxUnits,
    }, indent=2)


@mcp.tool()
def assign_resource(task_unique_id: int, resource_name: str, units: float = 1.0) -> str:
    """
    Assign a resource to a task. If the resource doesn't exist, it is created.
    If the task already has resources, the new one is appended.

    Args:
        task_unique_id: Task UniqueID (required).
        resource_name:  Resource name to assign (required).
        units:          Allocation units, e.g. 1.0 = 100% (default 1.0).
    """
    app  = get_app()
    proj = get_proj(app)

    # Find task
    task = None
    for t in proj.Tasks:
        if t is not None and t.UniqueID == task_unique_id:
            task = t
            break
    if task is None:
        return json.dumps({"error": f"Task UniqueID {task_unique_id} not found."})

    # Check if resource exists, create if not
    res_exists = False
    for r in proj.Resources:
        if r is not None and r.Name.lower() == resource_name.lower():
            res_exists = True
            break
    if not res_exists:
        proj.Resources.Add(resource_name)

    # Append to ResourceNames (handles existing assignments)
    existing = (task.ResourceNames or "").strip()
    if existing:
        # Check if already assigned
        existing_names = [n.strip().lower() for n in existing.split(",")]
        if resource_name.lower() not in existing_names:
            task.ResourceNames = existing + "," + resource_name
    else:
        task.ResourceNames = resource_name

    app.FileSave()
    return json.dumps({
        "status":         "assigned",
        "task_unique_id": task_unique_id,
        "task_name":      task.Name,
        "resource_name":  resource_name,
        "resource_names": task.ResourceNames,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Import / Export
# ---------------------------------------------------------------------------

@mcp.tool()
def import_xml(file_path: str) -> str:
    """
    Open an MS Project XML file (e.g. the consolidated EXPO 2030 roadmap).
    MS Project must be running.
    """
    return open_project(file_path)


@mcp.tool()
def export_xml(output_path: str) -> str:
    """Export the active project to MS Project XML format."""
    return save_project_as(output_path, format="xml")


# ---------------------------------------------------------------------------
# TOOLS — Utilities
# ---------------------------------------------------------------------------

@mcp.tool()
def search_tasks(query: str, include_summary: bool = False) -> str:
    """
    Search for tasks by name (case-insensitive substring match).
    Returns matching tasks with their UniqueIDs for use in other tools.
    """
    return get_tasks(include_summary=include_summary, keyword=query)


@mcp.tool()
def get_progress_summary() -> str:
    """
    Return a high-level progress summary:
    - Count by % complete bucket (0%, 1-99%, 100%)
    - Count by RAG status (Text1)
    - Count of overdue, critical tasks
    """
    import datetime
    today = datetime.datetime.now()
    app   = get_app()
    proj  = get_proj(app)

    not_started = in_progress = complete = 0
    rag_counts  = {"Red": 0, "Amber": 0, "Green": 0, "Other": 0}
    overdue     = critical = 0

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue

        pct = t.PercentComplete
        if pct == 0:
            not_started += 1
        elif pct < 100:
            in_progress += 1
        else:
            complete += 1

        rag = (t.Text1 or "").strip()
        if rag in rag_counts:
            rag_counts[rag] += 1
        elif rag:
            rag_counts["Other"] += 1

        if t.Critical:
            critical += 1

        try:
            fin = _to_naive(t.Finish)
            if fin and fin < today and pct < 100:
                overdue += 1
        except Exception:
            pass

    total = not_started + in_progress + complete

    return json.dumps({
        "project":     proj.Name,
        "total_tasks": total,
        "by_progress": {
            "not_started": not_started,
            "in_progress": in_progress,
            "complete":    complete,
            "pct_complete": round(complete / total * 100, 1) if total else 0,
        },
        "by_rag": rag_counts,
        "overdue":   overdue,
        "critical":  critical,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — WBS & Hierarchy
# ---------------------------------------------------------------------------

@mcp.tool()
def indent_task(unique_id: int, direction: str = "indent") -> str:
    """
    Promote or demote a task in the WBS hierarchy.

    Args:
        unique_id: Task UniqueID (required).
        direction: 'indent' to demote (increase level) or 'outdent' to promote (decrease level).
    """
    app  = get_app()
    proj = get_proj(app)

    for t in proj.Tasks:
        if t is not None and t.UniqueID == unique_id:
            old_level = t.OutlineLevel
            app.SelectRow(t.ID, False)
            if direction.lower() == "outdent":
                app.OutlineOutdent()
            else:
                app.OutlineIndent()
            app.FileSave()
            return json.dumps({
                "status":    "updated",
                "unique_id": unique_id,
                "name":      t.Name,
                "old_level": old_level,
                "new_level": t.OutlineLevel,
            }, indent=2)

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


@mcp.tool()
def get_wbs_structure(max_level: int = 0) -> str:
    """
    Export the full WBS hierarchy as a nested JSON tree.
    Useful for dashboards, reporting, and verifying project structure.

    Args:
        max_level: Maximum outline level to include (0 = all levels).
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    def fmt(dt):
        try:
            return str(dt)[:10] if dt else None
        except Exception:
            return None

    # Build flat list first
    flat = []
    for t in proj.Tasks:
        if t is None:
            continue
        if max_level > 0 and t.OutlineLevel > max_level:
            continue
        flat.append({
            "unique_id":     t.UniqueID,
            "id":            t.ID,
            "name":          t.Name,
            "level":         t.OutlineLevel,
            "summary":       bool(t.Summary),
            "milestone":     bool(t.Milestone),
            "start":         fmt(t.Start),
            "finish":        fmt(t.Finish),
            "duration_days": round(t.Duration / mpd, 2) if t.Duration else 0,
            "children":      [],
        })

    # Build tree using stack
    root = {"name": proj.Name, "level": 0, "children": []}
    stack = [root]

    for node in flat:
        level = node["level"]
        # Pop stack back to parent level
        while len(stack) > level:
            stack.pop()
        # Append to current parent
        stack[-1]["children"].append(node)
        # Push this node as potential parent
        stack.append(node)

    return json.dumps(root, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Calendars
# ---------------------------------------------------------------------------

@mcp.tool()
def get_calendars() -> str:
    """List all base calendars in the project and which one is active."""
    app  = get_app()
    proj = get_proj(app)

    calendars = []
    try:
        for cal in proj.BaseCalendars:
            if cal is not None:
                calendars.append(str(cal.Name))
    except Exception:
        pass

    active = ""
    try:
        active = str(proj.Calendar)
    except Exception:
        pass

    return json.dumps({
        "active_calendar": active,
        "calendars":       calendars,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Schedule Analysis
# ---------------------------------------------------------------------------

@mcp.tool()
def get_schedule_analysis() -> str:
    """
    Return float/slack metrics and schedule health for all non-summary tasks.
    TotalSlack and FreeSlack are converted from minutes to working days.
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    def fmt(dt):
        try:
            return str(dt)[:10] if dt else None
        except Exception:
            return None

    tasks = []
    zero_float = 0
    negative_float = 0
    total_slack_sum = 0
    count = 0

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue
        count += 1

        try:
            ts = round(t.TotalSlack / mpd, 2) if t.TotalSlack is not None else 0
        except Exception:
            ts = 0
        try:
            fs = round(t.FreeSlack / mpd, 2) if t.FreeSlack is not None else 0
        except Exception:
            fs = 0

        if ts == 0:
            zero_float += 1
        if ts < 0:
            negative_float += 1
        total_slack_sum += ts

        tasks.append({
            "unique_id":       t.UniqueID,
            "name":            t.Name,
            "total_slack_days": ts,
            "free_slack_days":  fs,
            "critical":        bool(t.Critical),
            "start":           fmt(t.Start),
            "finish":          fmt(t.Finish),
        })

    return json.dumps({
        "summary": {
            "total_tasks":     count,
            "zero_float":      zero_float,
            "negative_float":  negative_float,
            "avg_total_slack":  round(total_slack_sum / count, 2) if count else 0,
        },
        "tasks": tasks,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Baselines
# ---------------------------------------------------------------------------

@mcp.tool()
def save_baseline(baseline_number: int = 0, all_tasks: bool = True) -> str:
    """
    Save a baseline snapshot for earned value and variance tracking.

    Args:
        baseline_number: 0 to 10 (Baseline, Baseline1 through Baseline10). Default 0.
        all_tasks:       True to baseline all tasks (default), False for selected only.
    """
    if baseline_number < 0 or baseline_number > 10:
        return json.dumps({"error": "baseline_number must be 0-10."})

    app  = get_app()
    proj = get_proj(app)

    task_count = sum(1 for t in proj.Tasks if t is not None)

    app.BaselineSave(All=all_tasks, Copy=baseline_number, Into=baseline_number)
    app.FileSave()

    return json.dumps({
        "status":          "saved",
        "baseline_number": baseline_number,
        "all_tasks":       all_tasks,
        "tasks_baselined": task_count if all_tasks else "selected",
    }, indent=2)


@mcp.tool()
def clear_baseline(baseline_number: int = 0, all_tasks: bool = True) -> str:
    """
    Clear a previously saved baseline.

    Args:
        baseline_number: 0 to 10. Default 0.
        all_tasks:       True to clear for all tasks (default).
    """
    if baseline_number < 0 or baseline_number > 10:
        return json.dumps({"error": "baseline_number must be 0-10."})

    app = get_app()
    app.BaselineClear(All=all_tasks, From=baseline_number)
    app.FileSave()

    return json.dumps({
        "status":          "cleared",
        "baseline_number": baseline_number,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Earned Value
# ---------------------------------------------------------------------------

@mcp.tool()
def get_earned_value() -> str:
    """
    Return earned value metrics for all non-summary tasks.
    Requires a saved baseline and progress (% complete) to return meaningful data.
    Fields: BCWS (PV), BCWP (EV), ACWP (AC), SV, CV, plus SPI and CPI.
    """
    app  = get_app()
    proj = get_proj(app)

    tasks = []
    totals = {"bcws": 0, "bcwp": 0, "acwp": 0, "sv": 0, "cv": 0}

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue

        try:
            bcws = float(t.BCWS) if t.BCWS else 0
        except Exception:
            bcws = 0
        try:
            bcwp = float(t.BCWP) if t.BCWP else 0
        except Exception:
            bcwp = 0
        try:
            acwp = float(t.ACWP) if t.ACWP else 0
        except Exception:
            acwp = 0
        try:
            sv = float(t.SV) if t.SV else 0
        except Exception:
            sv = 0
        try:
            cv = float(t.CV) if t.CV else 0
        except Exception:
            cv = 0

        totals["bcws"] += bcws
        totals["bcwp"] += bcwp
        totals["acwp"] += acwp
        totals["sv"]   += sv
        totals["cv"]   += cv

        tasks.append({
            "unique_id": t.UniqueID,
            "name":      t.Name,
            "bcws":      bcws,
            "bcwp":      bcwp,
            "acwp":      acwp,
            "sv":        sv,
            "cv":        cv,
        })

    # Derived indices
    spi = round(totals["bcwp"] / totals["bcws"], 3) if totals["bcws"] else 0
    cpi = round(totals["bcwp"] / totals["acwp"], 3) if totals["acwp"] else 0

    if totals["bcws"] == 0 and totals["bcwp"] == 0:
        warning = "No earned value data. Ensure a baseline is saved and progress is entered."
    else:
        warning = None

    result = {
        "project_totals": {
            "bcws": totals["bcws"],
            "bcwp": totals["bcwp"],
            "acwp": totals["acwp"],
            "sv":   totals["sv"],
            "cv":   totals["cv"],
            "spi":  spi,
            "cpi":  cpi,
        },
        "tasks": tasks,
    }
    if warning:
        result["warning"] = warning

    return json.dumps(result, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 3: Multi-Project Navigation
# ---------------------------------------------------------------------------

@mcp.tool()
def list_projects() -> str:
    """
    List all open projects in MS Project with metadata.
    Returns project count, which is active, and details for each.
    """
    app = get_app(require_project=False)

    projects = []
    active_name = ""
    try:
        active_name = app.ActiveProject.Name
    except Exception:
        pass

    for i in range(1, app.Projects.Count + 1):
        p = app.Projects(i)
        projects.append({
            "index":      i,
            "name":       p.Name,
            "full_path":  p.FullName,
            "task_count": p.Tasks.Count,
            "start":      _fmt_date(p.ProjectStart),
            "finish":     _fmt_date(p.ProjectFinish),
            "is_active":  p.Name == active_name,
        })

    return json.dumps({
        "count":    len(projects),
        "active":   active_name,
        "projects": projects,
    }, indent=2)


@mcp.tool()
def switch_project(name_or_index: str) -> str:
    """
    Switch the active project by name (substring match) or 1-based index.

    Args:
        name_or_index: Project name (case-insensitive substring) or numeric index.
    """
    app = get_app(require_project=False)

    if app.Projects.Count == 0:
        return json.dumps({"error": "No projects are open."})

    # Try as index first
    try:
        idx = int(name_or_index)
        if 1 <= idx <= app.Projects.Count:
            p = app.Projects(idx)
            p.Activate()
            return json.dumps({
                "status":     "switched",
                "name":       p.Name,
                "full_path":  p.FullName,
                "task_count": p.Tasks.Count,
                "start":      _fmt_date(p.ProjectStart),
                "finish":     _fmt_date(p.ProjectFinish),
            }, indent=2)
        else:
            return json.dumps({"error": f"Index {idx} out of range. Projects: 1-{app.Projects.Count}."})
    except ValueError:
        pass

    # Substring match on name
    query = name_or_index.lower()
    matches = []
    for i in range(1, app.Projects.Count + 1):
        p = app.Projects(i)
        if query in p.Name.lower():
            matches.append((i, p))

    if len(matches) == 0:
        names = [app.Projects(i).Name for i in range(1, app.Projects.Count + 1)]
        return json.dumps({"error": f"No project matching '{name_or_index}'. Open: {names}"})
    if len(matches) > 1:
        names = [m[1].Name for m in matches]
        return json.dumps({"error": f"Multiple matches for '{name_or_index}': {names}. Be more specific."})

    _, p = matches[0]
    try:
        active = app.ActiveProject.Name
        if active == p.Name:
            return json.dumps({
                "status":     "already_active",
                "name":       p.Name,
                "full_path":  p.FullName,
                "task_count": p.Tasks.Count,
                "start":      _fmt_date(p.ProjectStart),
                "finish":     _fmt_date(p.ProjectFinish),
            }, indent=2)
    except Exception:
        pass

    p.Activate()
    return json.dumps({
        "status":     "switched",
        "name":       p.Name,
        "full_path":  p.FullName,
        "task_count": p.Tasks.Count,
        "start":      _fmt_date(p.ProjectStart),
        "finish":     _fmt_date(p.ProjectFinish),
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 3: Advanced Filtering & Queries
# ---------------------------------------------------------------------------

@mcp.tool()
def filter_tasks(filters_json: str) -> str:
    """
    Powerful AND-logic filtering across all task fields with sort and pagination.

    Args:
        filters_json: JSON string with filter keys (all optional):
            rag, resource, start_after, start_before, finish_after, finish_before,
            min_pct, max_pct, outline_level, critical (bool), milestone (bool),
            active (bool), summary (bool), name_contains,
            text1, text2, text3, flag1 (bool), flag2 (bool),
            sort_by (field name), sort_desc (bool), limit (int), offset (int).
            Example: '{"rag": "Red", "critical": true, "limit": 20}'
    """
    f = json.loads(filters_json)
    app  = get_app()
    proj = get_proj(app)

    predicates = []

    if "rag" in f:
        v = f["rag"].lower()
        predicates.append(lambda t, _v=v: (t.Text1 or "").strip().lower() == _v)
    if "resource" in f:
        v = f["resource"].lower()
        predicates.append(lambda t, _v=v: _v in (t.ResourceNames or "").lower())
    if "start_after" in f:
        d = _parse_date(f["start_after"])
        predicates.append(lambda t, _d=d: t.Start is not None and t.Start >= _d)
    if "start_before" in f:
        d = _parse_date(f["start_before"])
        predicates.append(lambda t, _d=d: t.Start is not None and t.Start <= _d)
    if "finish_after" in f:
        d = _parse_date(f["finish_after"])
        predicates.append(lambda t, _d=d: t.Finish is not None and t.Finish >= _d)
    if "finish_before" in f:
        d = _parse_date(f["finish_before"])
        predicates.append(lambda t, _d=d: t.Finish is not None and t.Finish <= _d)
    if "min_pct" in f:
        v = f["min_pct"]
        predicates.append(lambda t, _v=v: t.PercentComplete >= _v)
    if "max_pct" in f:
        v = f["max_pct"]
        predicates.append(lambda t, _v=v: t.PercentComplete <= _v)
    if "outline_level" in f:
        v = f["outline_level"]
        predicates.append(lambda t, _v=v: t.OutlineLevel == _v)
    if "critical" in f:
        v = f["critical"]
        predicates.append(lambda t, _v=v: bool(t.Critical) == _v)
    if "milestone" in f:
        v = f["milestone"]
        predicates.append(lambda t, _v=v: bool(t.Milestone) == _v)
    if "active" in f:
        v = f["active"]
        predicates.append(lambda t, _v=v: bool(t.Active) == _v)
    if "summary" in f:
        v = f["summary"]
        predicates.append(lambda t, _v=v: bool(t.Summary) == _v)
    if "name_contains" in f:
        v = f["name_contains"].lower()
        predicates.append(lambda t, _v=v: _v in t.Name.lower())
    if "text1" in f:
        v = f["text1"].lower()
        predicates.append(lambda t, _v=v: (t.Text1 or "").strip().lower() == _v)
    if "text2" in f:
        v = f["text2"].lower()
        predicates.append(lambda t, _v=v: (t.Text2 or "").strip().lower() == _v)
    if "text3" in f:
        v = f["text3"].lower()
        predicates.append(lambda t, _v=v: (t.Text3 or "").strip().lower() == _v)
    if "flag1" in f:
        v = f["flag1"]
        predicates.append(lambda t, _v=v: bool(t.Flag1) == _v)
    if "flag2" in f:
        v = f["flag2"]
        predicates.append(lambda t, _v=v: bool(t.Flag2) == _v)

    # Collect matching tasks
    matched = []
    for t in proj.Tasks:
        if t is None:
            continue
        if all(p(t) for p in predicates):
            matched.append(task_to_dict(t, proj))

    # Sort
    sort_by = f.get("sort_by", "")
    if sort_by and matched:
        desc = f.get("sort_desc", False)
        try:
            matched.sort(key=lambda x: x.get(sort_by) or "", reverse=desc)
        except Exception:
            pass

    total = len(matched)
    offset = f.get("offset", 0)
    limit  = f.get("limit", 0)
    if offset > 0:
        matched = matched[offset:]
    if limit > 0:
        matched = matched[:limit]

    return json.dumps({
        "total_matching": total,
        "returned":       len(matched),
        "offset":         offset,
        "limit":          limit or total,
        "tasks":          matched,
    }, indent=2)


@mcp.tool()
def group_tasks_by(field: str, include_tasks: bool = False) -> str:
    """
    Group non-summary tasks by a field and return counts per group.

    Args:
        field:         Field to group by: 'rag', 'resource', 'outline_level', 'critical',
                       'milestone', 'percent_complete', 'text1', 'text2', 'text3',
                       'flag1', 'flag2'.
        include_tasks: If true, include task list per group (default false).
    """
    app  = get_app()
    proj = get_proj(app)

    groups = {}
    total  = 0

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue
        total += 1

        td = task_to_dict(t, proj) if include_tasks else None

        if field == "resource":
            # Split comma-separated resource names
            names = [n.strip() for n in (t.ResourceNames or "").split(",") if n.strip()]
            if not names:
                names = ["(unassigned)"]
            keys = names
        elif field == "percent_complete":
            pct = t.PercentComplete
            if pct == 0:
                keys = ["0%"]
            elif pct <= 25:
                keys = ["1-25%"]
            elif pct <= 50:
                keys = ["26-50%"]
            elif pct <= 75:
                keys = ["51-75%"]
            elif pct < 100:
                keys = ["76-99%"]
            else:
                keys = ["100%"]
        elif field in ("rag", "text1"):
            keys = [(t.Text1 or "").strip() or "(blank)"]
        elif field == "text2":
            keys = [(t.Text2 or "").strip() or "(blank)"]
        elif field == "text3":
            keys = [(t.Text3 or "").strip() or "(blank)"]
        elif field == "outline_level":
            keys = [str(t.OutlineLevel)]
        elif field == "critical":
            keys = [str(bool(t.Critical))]
        elif field == "milestone":
            keys = [str(bool(t.Milestone))]
        elif field == "flag1":
            keys = [str(bool(t.Flag1))]
        elif field == "flag2":
            keys = [str(bool(t.Flag2))]
        else:
            keys = [str(getattr(t, field, "(unknown)"))]

        for k in keys:
            if k not in groups:
                groups[k] = {"value": k, "count": 0}
                if include_tasks:
                    groups[k]["tasks"] = []
            groups[k]["count"] += 1
            if include_tasks and td:
                groups[k]["tasks"].append(td)

    result = sorted(groups.values(), key=lambda g: g["count"], reverse=True)

    return json.dumps({
        "field":       field,
        "groups":      result,
        "total_tasks": total,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 3: Calendar Management (Write)
# ---------------------------------------------------------------------------

@mcp.tool()
def set_calendar_exception(
    calendar_name: str,
    name: str,
    start: str,
    finish: str,
    working: bool = False,
) -> str:
    """
    Add a holiday or exception to a base calendar.

    Args:
        calendar_name: Name of the base calendar (e.g. 'Standard').
        name:          Exception name (e.g. 'National Day').
        start:         Start date as YYYY-MM-DD.
        finish:        End date as YYYY-MM-DD (same as start for single day).
        working:       True for a working exception, False for non-working/holiday (default).
    """
    app  = get_app()
    proj = get_proj(app)

    # Validate calendar exists
    valid_cals = []
    try:
        for cal in proj.BaseCalendars:
            if cal is not None:
                valid_cals.append(str(cal.Name))
    except Exception:
        pass

    if calendar_name not in valid_cals:
        return json.dumps({"error": f"Calendar '{calendar_name}' not found. Available: {valid_cals}"})

    try:
        # Use the Calendar.Exceptions collection for date-range exceptions
        cal = None
        for c in proj.BaseCalendars:
            if c is not None and str(c.Name) == calendar_name:
                cal = c
                break

        # pjCalendarExceptionDaily = 1
        start_dt  = _parse_date(start)
        finish_dt = _parse_date(finish)
        exc = cal.Exceptions.Add(1, start_dt, finish_dt, name)
    except Exception as e:
        return json.dumps({"error": f"Failed to set exception: {e}"})

    app.FileSave()
    return json.dumps({
        "status":   "created",
        "calendar": calendar_name,
        "exception": name,
        "start":    start,
        "finish":   finish,
        "working":  working,
    }, indent=2)


@mcp.tool()
def set_project_calendar(calendar_name: str) -> str:
    """
    Switch the active project's base calendar.

    Args:
        calendar_name: Name of the base calendar to use (e.g. '24 Hours', 'Standard').
    """
    app  = get_app()
    proj = get_proj(app)

    # Validate
    valid_cals = []
    try:
        for cal in proj.BaseCalendars:
            if cal is not None:
                valid_cals.append(str(cal.Name))
    except Exception:
        pass

    if calendar_name not in valid_cals:
        return json.dumps({"error": f"Calendar '{calendar_name}' not found. Available: {valid_cals}"})

    previous = ""
    try:
        previous = str(proj.Calendar)
    except Exception:
        pass

    # proj.Calendar is read-only in some COM bindings;
    # try multiple approaches to set it
    set_ok = False
    errors = []

    # Approach 1: direct property set
    try:
        proj.Calendar = calendar_name
        set_ok = True
    except Exception as e:
        errors.append(f"direct: {e}")

    # Approach 2: use the Calendar object from BaseCalendars
    if not set_ok:
        try:
            for cal in proj.BaseCalendars:
                if cal is not None and str(cal.Name) == calendar_name:
                    proj.Calendar = cal
                    set_ok = True
                    break
        except Exception as e:
            errors.append(f"object: {e}")

    # Approach 3: use _oleobj_ InvokeTypes to force property set
    if not set_ok:
        try:
            import pythoncom
            # Calendar property dispid — try to find via QueryInterface
            proj._oleobj_.InvokeTypes(
                0x30, 0, pythoncom.DISPATCH_PROPERTYPUT,
                (24, 0),  # VT_VOID return
                ((8, 1),),  # VT_BSTR input
                calendar_name
            )
            set_ok = True
        except Exception as e:
            errors.append(f"oleobj: {e}")

    if not set_ok:
        return json.dumps({"error": f"Could not set calendar. Tried: {errors}"})

    app.FileSave()

    return json.dumps({
        "status":   "updated",
        "calendar": calendar_name,
        "previous": previous,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 3: Extended Custom Fields
# ---------------------------------------------------------------------------

@mcp.tool()
def update_custom_fields(unique_id: int, fields_json: str) -> str:
    """
    Write any custom field on a task: Text1-30, Number1-20, Date1-10, Flag1-20, Duration1-10.

    Args:
        unique_id:   Task UniqueID (required).
        fields_json: JSON object mapping field names to values.
                     Example: '{"Text5": "Phase A", "Number1": 42, "Flag3": true, "Date1": "2026-06-01"}'
    """
    fields = json.loads(fields_json)
    app    = get_app()
    proj   = get_proj(app)
    mpd    = _get_mpd(proj)

    t = _find_task(proj, unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    changed = []
    errors  = []

    for field_name, value in fields.items():
        try:
            # Validate field name and get the type
            _field_id, field_type = _custom_field_id(field_name)

            # Normalize the property name for setattr (e.g. "text5" -> "Text5")
            prop_name = field_type.capitalize() + field_name[len(field_type):]

            if field_type == "date":
                value = _parse_date(value)
            elif field_type == "flag":
                value = bool(value)
            elif field_type == "duration":
                value = float(value) * mpd
            elif field_type == "number":
                value = float(value)
            else:
                value = str(value)

            setattr(t, prop_name, value)
            changed.append({"field": field_name, "value": str(value)})
        except Exception as e:
            errors.append({"field": field_name, "error": str(e)})

    app.FileSave()
    return json.dumps({
        "status":    "updated",
        "unique_id": unique_id,
        "name":      t.Name,
        "changed":   changed,
        "errors":    errors,
    }, indent=2)


@mcp.tool()
def get_custom_field_values(field_name: str) -> str:
    """
    Get all unique values for a custom field across all tasks.
    Useful for validation, audit, and understanding what values exist.

    Args:
        field_name: Custom field name, e.g. 'Text1', 'Number1', 'Flag3'.
    """
    try:
        _field_id, field_type = _custom_field_id(field_name)
    except ValueError as e:
        return json.dumps({"error": str(e)})

    # Normalize property name (e.g. "text1" -> "Text1")
    prop_name = field_type.capitalize() + field_name.strip()[len(field_type):]

    app  = get_app()
    proj = get_proj(app)

    value_counts = {}
    total = 0

    for t in proj.Tasks:
        if t is None:
            continue
        total += 1
        try:
            val = getattr(t, prop_name, None)
            if val is None:
                val = "(blank)"
            val = str(val).strip()
            if not val:
                val = "(blank)"
        except Exception:
            val = "(error)"
        value_counts[val] = value_counts.get(val, 0) + 1

    unique_values = sorted(value_counts.keys())

    return json.dumps({
        "field":         field_name,
        "unique_values": unique_values,
        "value_counts":  value_counts,
        "total_tasks":   total,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 3: Schedule & Dependency Intelligence
# ---------------------------------------------------------------------------

@mcp.tool()
def validate_schedule() -> str:
    """
    Comprehensive schedule health check — the PMO's best friend.
    Checks for orphan tasks, missing resources, past-due zero-progress,
    empty summaries, missing dates, and negative slack.
    Returns a health score (0-100) and categorized issues.
    """
    app   = get_app()
    proj  = get_proj(app)
    mpd   = _get_mpd(proj)
    today = datetime.datetime.now()

    issues = {
        "orphan_tasks":       {"count": 0, "tasks": []},
        "no_resources":       {"count": 0, "tasks": []},
        "past_due_no_progress": {"count": 0, "tasks": []},
        "empty_summaries":    {"count": 0, "tasks": []},
        "missing_dates":      {"count": 0, "tasks": []},
        "negative_slack":     {"count": 0, "tasks": []},
    }

    total_tasks = 0
    tasks_list  = []  # (task, is_summary, outline_level)

    for t in proj.Tasks:
        if t is None:
            continue
        total_tasks += 1
        tasks_list.append(t)

    for i, t in enumerate(tasks_list):
        tid = {"unique_id": t.UniqueID, "name": t.Name}

        # Empty summaries: summary with no children
        if t.Summary:
            has_child = False
            if i + 1 < len(tasks_list):
                next_t = tasks_list[i + 1]
                if next_t.OutlineLevel > t.OutlineLevel:
                    has_child = True
            if not has_child:
                issues["empty_summaries"]["count"] += 1
                issues["empty_summaries"]["tasks"].append(tid)
            continue  # Skip non-leaf checks for summaries

        # Orphan tasks: no predecessors AND no successors
        preds = (t.Predecessors or "").strip()
        # Check if this task is a predecessor for any other task
        has_successor = False
        task_id_str = str(t.ID)
        for other in tasks_list:
            if other is None or other.UniqueID == t.UniqueID:
                continue
            other_preds = (other.Predecessors or "").strip()
            if other_preds:
                # Check if our task ID appears in other's predecessors
                for part in other_preds.split(","):
                    part = part.strip()
                    # Extract the numeric ID from predecessor string like "5FS" or "5"
                    num = ""
                    for ch in part:
                        if ch.isdigit():
                            num += ch
                        else:
                            break
                    if num == task_id_str:
                        has_successor = True
                        break
            if has_successor:
                break

        if not preds and not has_successor:
            issues["orphan_tasks"]["count"] += 1
            issues["orphan_tasks"]["tasks"].append(tid)

        # No resources (non-milestone)
        if not t.Milestone and not (t.ResourceNames or "").strip():
            issues["no_resources"]["count"] += 1
            issues["no_resources"]["tasks"].append(tid)

        # Past due, zero progress
        try:
            fin = _to_naive(t.Finish)
            if fin and fin < today and t.PercentComplete == 0:
                issues["past_due_no_progress"]["count"] += 1
                issues["past_due_no_progress"]["tasks"].append(tid)
        except Exception:
            pass

        # Missing dates
        try:
            if not t.Start or not t.Finish:
                issues["missing_dates"]["count"] += 1
                issues["missing_dates"]["tasks"].append(tid)
        except Exception:
            pass

        # Negative slack
        try:
            if t.TotalSlack is not None and t.TotalSlack < 0:
                issues["negative_slack"]["count"] += 1
                issues["negative_slack"]["tasks"].append(tid)
        except Exception:
            pass

    total_issues = sum(cat["count"] for cat in issues.values())
    health_score = max(0, round(100 - (total_issues / total_tasks * 100))) if total_tasks else 0

    return json.dumps({
        "project":      proj.Name,
        "health_score": health_score,
        "issues":       issues,
        "summary": {
            "total_tasks":  total_tasks,
            "total_issues": total_issues,
        },
    }, indent=2)


@mcp.tool()
def get_milestone_report(days_ahead: int = 30, upcoming_count: int = 10) -> str:
    """
    Milestone-focused status report for executive dashboards.
    Categorizes milestones as complete, overdue, at_risk, or on_track.
    Includes baseline variance if a baseline is saved.

    Args:
        days_ahead:     Number of days ahead to consider 'at risk' (default 30).
        upcoming_count: Max upcoming milestones to return (default 10).
    """
    app   = get_app()
    proj  = get_proj(app)
    today = datetime.datetime.now()
    horizon = today + datetime.timedelta(days=days_ahead)

    by_status = {"complete": 0, "overdue": 0, "at_risk": 0, "on_track": 0}
    upcoming  = []
    overdue   = []
    total     = 0

    for t in proj.Tasks:
        if t is None or not t.Milestone:
            continue
        total += 1

        finish = None
        try:
            finish = _to_naive(t.Finish)
        except Exception:
            pass

        pct = t.PercentComplete

        # Baseline variance
        variance_days = None
        try:
            bf = _to_naive(t.BaselineFinish)
            if bf and finish:
                delta = finish - bf
                variance_days = delta.days if hasattr(delta, "days") else None
        except Exception:
            pass

        entry = {
            "unique_id":      t.UniqueID,
            "name":           t.Name,
            "finish":         _fmt_date(finish),
            "percent":        pct,
            "variance_days":  variance_days,
        }

        if pct >= 100:
            by_status["complete"] += 1
        elif finish and finish < today:
            by_status["overdue"] += 1
            overdue.append(entry)
        elif finish and finish <= horizon and pct < 100:
            by_status["at_risk"] += 1
            upcoming.append(entry)
        else:
            by_status["on_track"] += 1
            if finish:
                upcoming.append(entry)

    # Sort upcoming by finish asc, overdue by finish asc
    upcoming.sort(key=lambda x: x["finish"] or "")
    overdue.sort(key=lambda x: x["finish"] or "")

    return json.dumps({
        "total_milestones": total,
        "by_status":        by_status,
        "upcoming":         upcoming[:upcoming_count],
        "overdue":          overdue,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 3: Resource Intelligence
# ---------------------------------------------------------------------------

@mcp.tool()
def get_resource_workload(resource_name: str, start_date: str = "", end_date: str = "") -> str:
    """
    Resource allocation view with conflict detection.
    Shows all assignments for a resource and identifies overlapping assignments.

    Args:
        resource_name: Resource name (case-insensitive exact match).
        start_date:    Filter assignments starting after this date (YYYY-MM-DD, optional).
        end_date:      Filter assignments ending before this date (YYYY-MM-DD, optional).
    """
    app  = get_app()
    proj = get_proj(app)

    # Find resource
    resource = None
    for r in proj.Resources:
        if r is not None and r.Name.lower() == resource_name.lower():
            resource = r
            break

    if resource is None:
        # List available resources
        avail = []
        for r in proj.Resources:
            if r is not None:
                avail.append(r.Name)
        return json.dumps({"error": f"Resource '{resource_name}' not found. Available: {avail}"})

    filter_start = _parse_date(start_date) if start_date else None
    filter_end   = _parse_date(end_date) if end_date else None

    assignments = []
    for a in resource.Assignments:
        try:
            a_start  = a.Start
            a_finish = a.Finish
            task     = a.Task

            if filter_start and a_finish and a_finish < filter_start:
                continue
            if filter_end and a_start and a_start > filter_end:
                continue

            work_hours = 0
            try:
                work_hours = round(a.Work / 60, 2)  # minutes to hours
            except Exception:
                pass

            assignments.append({
                "task_unique_id": task.UniqueID if task else None,
                "task_name":      task.Name if task else "(unknown)",
                "start":          _fmt_date(a_start),
                "finish":         _fmt_date(a_finish),
                "units":          a.Units,
                "work_hours":     work_hours,
            })
        except Exception:
            continue

    # Conflict detection: find overlapping date ranges
    conflicts = []
    for i in range(len(assignments)):
        for j in range(i + 1, len(assignments)):
            a = assignments[i]
            b = assignments[j]
            if a["start"] and a["finish"] and b["start"] and b["finish"]:
                if a["start"] <= b["finish"] and b["start"] <= a["finish"]:
                    overlap_start = max(a["start"], b["start"])
                    overlap_finish = min(a["finish"], b["finish"])
                    combined = (a.get("units") or 0) + (b.get("units") or 0)
                    conflicts.append({
                        "task_a":          a["task_name"],
                        "task_b":          b["task_name"],
                        "overlap_start":   overlap_start,
                        "overlap_finish":  overlap_finish,
                        "combined_units":  combined,
                    })

    overallocated = False
    try:
        overallocated = bool(resource.Overallocated)
    except Exception:
        pass

    max_units = 1.0
    try:
        max_units = resource.MaxUnits
    except Exception:
        pass

    return json.dumps({
        "resource":      resource.Name,
        "overallocated": overallocated,
        "max_units":     max_units,
        "assignments":   assignments,
        "conflicts":     conflicts,
    }, indent=2)


@mcp.tool()
def level_resources() -> str:
    """
    Run MS Project's built-in resource leveling algorithm.
    WARNING: This may shift task dates. Save a baseline first if tracking variance.
    """
    app  = get_app()
    proj = get_proj(app)

    app.LevelNow()
    app.FileSave()

    return json.dumps({
        "status":  "leveled",
        "project": proj.Name,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 3: Deadline & Active Task Management
# ---------------------------------------------------------------------------

@mcp.tool()
def set_deadline(unique_id: int, deadline_date: str) -> str:
    """
    Set a soft deadline on a task. Shows a visual indicator if finish > deadline.
    Unlike hard constraints, deadlines don't affect scheduling.

    Args:
        unique_id:     Task UniqueID (required).
        deadline_date: Deadline as YYYY-MM-DD, or 'clear' to remove.
    """
    app  = get_app()
    proj = get_proj(app)

    t = _find_task(proj, unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    if deadline_date.lower() == "clear":
        t.Deadline = "NA"
        app.FileSave()
        return json.dumps({
            "status":    "cleared",
            "unique_id": unique_id,
            "name":      t.Name,
            "deadline":  None,
        }, indent=2)

    dl = _parse_date(deadline_date)
    t.Deadline = dl
    app.FileSave()

    deadline_missed = False
    try:
        if t.Finish and t.Finish > dl:
            deadline_missed = True
    except Exception:
        pass

    return json.dumps({
        "status":          "set",
        "unique_id":       unique_id,
        "name":            t.Name,
        "deadline":        deadline_date,
        "finish":          _fmt_date(t.Finish),
        "deadline_missed": deadline_missed,
    }, indent=2)


@mcp.tool()
def set_task_active(unique_id: int, active: bool = True) -> str:
    """
    Activate or deactivate a task. Deactivated tasks are excluded from scheduling
    but remain visible (soft-delete).

    Args:
        unique_id: Task UniqueID (required).
        active:    True to activate (default), False to deactivate.
    """
    app  = get_app()
    proj = get_proj(app)

    t = _find_task(proj, unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    t.Active = active
    app.FileSave()

    return json.dumps({
        "status":    "updated",
        "unique_id": unique_id,
        "name":      t.Name,
        "active":    active,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 3: Safe Bulk Operations
# ---------------------------------------------------------------------------

@mcp.tool()
def dry_run_bulk_update(updates_json: str) -> str:
    """
    Preview bulk changes without modifying anything — the enterprise safety net.
    Shows what would change for each task without actually applying updates.

    Args:
        updates_json: JSON string — list of objects with fields:
            unique_id (required), name, start, finish, duration_days,
            percent_complete, rag, text2, text3, notes.
            Same format as bulk_update_tasks.
    """
    items = json.loads(updates_json)
    app   = get_app()
    proj  = get_proj(app)
    mpd   = _get_mpd(proj)

    changes   = []
    not_found = []
    no_change = []

    for item in items:
        uid = item.get("unique_id")
        if uid is None:
            continue

        t = _find_task(proj, uid)
        if t is None:
            not_found.append(uid)
            continue

        field_changes = []

        if "name" in item and item["name"] and item["name"] != t.Name:
            field_changes.append({"field": "name", "old": t.Name, "new": item["name"]})
        if "start" in item and item["start"]:
            old_start = _fmt_date(t.Start)
            if old_start != item["start"]:
                field_changes.append({"field": "start", "old": old_start, "new": item["start"]})
        if "finish" in item and item["finish"]:
            old_finish = _fmt_date(t.Finish)
            if old_finish != item["finish"]:
                field_changes.append({"field": "finish", "old": old_finish, "new": item["finish"]})
        if "duration_days" in item and item["duration_days"] is not None:
            old_dur = round(t.Duration / mpd, 2) if t.Duration else 0
            if old_dur != item["duration_days"]:
                field_changes.append({"field": "duration_days", "old": old_dur, "new": item["duration_days"]})
        if "percent_complete" in item and item["percent_complete"] is not None:
            if t.PercentComplete != item["percent_complete"]:
                field_changes.append({"field": "percent_complete", "old": t.PercentComplete, "new": item["percent_complete"]})
        if "rag" in item and item["rag"]:
            old_rag = (t.Text1 or "").strip()
            if old_rag != item["rag"]:
                field_changes.append({"field": "rag", "old": old_rag, "new": item["rag"]})
        if "text2" in item and item["text2"]:
            old = (t.Text2 or "").strip()
            if old != item["text2"]:
                field_changes.append({"field": "text2", "old": old, "new": item["text2"]})
        if "text3" in item and item["text3"]:
            old = (t.Text3 or "").strip()
            if old != item["text3"]:
                field_changes.append({"field": "text3", "old": old, "new": item["text3"]})
        if "notes" in item and item["notes"]:
            old = (t.Notes or "").strip()
            if old != item["notes"]:
                field_changes.append({"field": "notes", "old": old[:50], "new": item["notes"][:50]})

        if field_changes:
            changes.append({
                "unique_id": uid,
                "name":      t.Name,
                "fields":    field_changes,
            })
        else:
            no_change.append(uid)

    total_changes = sum(len(c["fields"]) for c in changes)

    return json.dumps({
        "preview":              True,
        "changes":              changes,
        "not_found":            not_found,
        "no_change":            no_change,
        "total_changes":        total_changes,
        "total_tasks_affected": len(changes),
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 4, Tier 1: High Impact
# ---------------------------------------------------------------------------

@mcp.tool()
def compare_baselines(baseline_a: int = 0, baseline_b: int = -1) -> str:
    """
    Variance report between two baselines, or baseline vs current schedule.

    Args:
        baseline_a: First baseline number (0-10). Default 0.
        baseline_b: Second baseline number (0-10), or -1 for current schedule (default).
    """
    if baseline_a < 0 or baseline_a > 10:
        return json.dumps({"error": "baseline_a must be 0-10."})
    if baseline_b < -1 or baseline_b > 10:
        return json.dumps({"error": "baseline_b must be -1 to 10 (-1 = current schedule)."})

    app  = get_app()
    proj = get_proj(app)

    def _bl_attr(n, suffix):
        if n == 0:
            return f"Baseline{suffix}"
        return f"Baseline{n}{suffix}"

    tasks = []
    tasks_with_variance = 0
    total_finish_variance = 0
    max_slippage = 0
    total_tasks = 0

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue
        total_tasks += 1

        try:
            a_start  = _to_naive(getattr(t, _bl_attr(baseline_a, "Start"), None))
            a_finish = _to_naive(getattr(t, _bl_attr(baseline_a, "Finish"), None))
        except Exception:
            a_start = a_finish = None

        if baseline_b == -1:
            b_start  = _to_naive(t.Start)
            b_finish = _to_naive(t.Finish)
        else:
            try:
                b_start  = _to_naive(getattr(t, _bl_attr(baseline_b, "Start"), None))
                b_finish = _to_naive(getattr(t, _bl_attr(baseline_b, "Finish"), None))
            except Exception:
                b_start = b_finish = None

        start_delta = finish_delta = None
        if a_start and b_start:
            try:
                start_delta = (b_start - a_start).days
            except Exception:
                pass
        if a_finish and b_finish:
            try:
                finish_delta = (b_finish - a_finish).days
            except Exception:
                pass

        if finish_delta is not None and finish_delta != 0:
            tasks_with_variance += 1
            total_finish_variance += finish_delta
            if finish_delta > max_slippage:
                max_slippage = finish_delta

        tasks.append({
            "unique_id":    t.UniqueID,
            "name":         t.Name,
            "a_start":      _fmt_date(a_start),
            "a_finish":     _fmt_date(a_finish),
            "b_start":      _fmt_date(b_start),
            "b_finish":     _fmt_date(b_finish),
            "start_delta":  start_delta,
            "finish_delta": finish_delta,
        })

    # Sort by finish variance desc (worst slippages first)
    tasks.sort(key=lambda x: x["finish_delta"] if x["finish_delta"] is not None else 0, reverse=True)

    return json.dumps({
        "baseline_a": baseline_a,
        "baseline_b": baseline_b if baseline_b >= 0 else "current",
        "summary": {
            "total_tasks":          total_tasks,
            "tasks_with_variance":  tasks_with_variance,
            "avg_finish_variance":  round(total_finish_variance / tasks_with_variance, 1) if tasks_with_variance else 0,
            "max_slippage":         max_slippage,
        },
        "tasks": tasks,
    }, indent=2)


@mcp.tool()
def get_dependency_chain(unique_id: int, direction: str = "successors", max_depth: int = 50) -> str:
    """
    Recursive walk of the dependency chain — 'what's downstream if this slips?'

    Args:
        unique_id: Starting task UniqueID (required).
        direction: 'successors' (default) or 'predecessors'.
        max_depth: Maximum depth to walk (default 50, safety cap).
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    root = _find_task(proj, unique_id)
    if root is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    visited = set()
    chain = []
    queue = [(root, 0)]  # (task, depth)
    visited.add(root.UniqueID)

    while queue:
        current, depth = queue.pop(0)
        if depth > 0:
            chain.append({
                "unique_id": current.UniqueID,
                "name":      current.Name,
                "depth":     depth,
                "start":     _fmt_date(current.Start),
                "finish":    _fmt_date(current.Finish),
                "critical":  bool(current.Critical),
            })
        if depth >= max_depth:
            continue

        try:
            for dep in current.TaskDependencies:
                if direction.lower() == "successors":
                    if dep.From.UniqueID == current.UniqueID:
                        next_task = dep.To
                    else:
                        continue
                else:
                    if dep.To.UniqueID == current.UniqueID:
                        next_task = dep.From
                    else:
                        continue

                if next_task.UniqueID not in visited:
                    visited.add(next_task.UniqueID)
                    # Add link info to the chain entry
                    link_type = dep.Type
                    lag_days = round(dep.Lag / mpd, 2) if dep.Lag else 0
                    queue.append((next_task, depth + 1))
                    # Update last chain entry with link info when it's added
        except Exception:
            pass

    return json.dumps({
        "root":          {"unique_id": unique_id, "name": root.Name},
        "direction":     direction,
        "depth_reached": max(e["depth"] for e in chain) if chain else 0,
        "chain":         chain,
    }, indent=2)


@mcp.tool()
def bulk_assign_resources(assignments_json: str) -> str:
    """
    Assign resources to multiple tasks in one call.

    Args:
        assignments_json: JSON string — list of {task_unique_id, resource_name, units (optional)}.
            Example: '[{"task_unique_id": 42, "resource_name": "Alice"},
                       {"task_unique_id": 55, "resource_name": "Bob", "units": 0.5}]'
    """
    items = json.loads(assignments_json)
    app   = get_app()
    proj  = get_proj(app)

    app.Calculation = 0
    try:
        uid_map = {t.UniqueID: t for t in proj.Tasks if t is not None}
        # Get existing resource names
        existing_resources = set()
        for r in proj.Resources:
            if r is not None:
                existing_resources.add(r.Name.lower())

        assigned = 0
        errors = []
        created_resources = []

        for item in items:
            task_uid = item["task_unique_id"]
            res_name = item["resource_name"]

            t = uid_map.get(task_uid)
            if t is None:
                errors.append({"task_unique_id": task_uid, "error": "task not found"})
                continue

            # Create resource if needed
            if res_name.lower() not in existing_resources:
                proj.Resources.Add(res_name)
                existing_resources.add(res_name.lower())
                created_resources.append(res_name)

            # Append to ResourceNames
            existing = (t.ResourceNames or "").strip()
            if existing:
                existing_names = [n.strip().lower() for n in existing.split(",")]
                if res_name.lower() not in existing_names:
                    t.ResourceNames = existing + "," + res_name
            else:
                t.ResourceNames = res_name

            assigned += 1
    finally:
        app.CalculateProject()
        app.Calculation = -1

    return json.dumps({
        "assigned":          assigned,
        "errors":            errors,
        "created_resources": created_resources,
    }, indent=2)


@mcp.tool()
def remove_resource_assignment(task_unique_id: int, resource_name: str) -> str:
    """
    Remove a specific resource from a task.

    Args:
        task_unique_id: Task UniqueID (required).
        resource_name:  Resource name to remove (required).
    """
    app  = get_app()
    proj = get_proj(app)

    t = _find_task(proj, task_unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {task_unique_id} not found."})

    existing = (t.ResourceNames or "").strip()
    if not existing:
        return json.dumps({"error": f"Task '{t.Name}' has no resources assigned."})

    names = [n.strip() for n in existing.split(",")]
    filtered = [n for n in names if n.lower() != resource_name.lower()]

    if len(filtered) == len(names):
        return json.dumps({"error": f"Resource '{resource_name}' not assigned to task '{t.Name}'. Current: {existing}"})

    t.ResourceNames = ",".join(filtered) if filtered else ""

    return json.dumps({
        "status":           "removed",
        "task_name":        t.Name,
        "removed":          resource_name,
        "resource_names_now": t.ResourceNames,
    }, indent=2)


@mcp.tool()
def update_resource(resource_name: str, new_name: str = "", max_units: float = -1, standard_rate: str = "", cost_per_use: float = -1) -> str:
    """
    Modify an existing resource's properties.

    Args:
        resource_name: Current resource name (required, case-insensitive).
        new_name:      New name for the resource (optional).
        max_units:     Maximum allocation units, e.g. 2.0 = 200% (optional, -1 = no change).
        standard_rate: Standard rate as string, e.g. '50/h' (optional).
        cost_per_use:  Fixed cost per use (optional, -1 = no change).
    """
    app  = get_app()
    proj = get_proj(app)

    resource = None
    for r in proj.Resources:
        if r is not None and r.Name.lower() == resource_name.lower():
            resource = r
            break

    if resource is None:
        avail = [r.Name for r in proj.Resources if r is not None]
        return json.dumps({"error": f"Resource '{resource_name}' not found. Available: {avail}"})

    changed = []

    if new_name:
        resource.Name = new_name
        changed.append("name")
    if max_units >= 0:
        resource.MaxUnits = max_units
        changed.append("max_units")
    if standard_rate:
        resource.StandardRate = standard_rate
        changed.append("standard_rate")
    if cost_per_use >= 0:
        resource.CostPerUse = cost_per_use
        changed.append("cost_per_use")

    return json.dumps({
        "status":  "updated",
        "name":    resource.Name,
        "changed": changed,
    }, indent=2)


@mcp.tool()
def move_task(unique_id: int, after_unique_id: int) -> str:
    """
    Reposition a task to appear after another task.

    Args:
        unique_id:       Task UniqueID to move (required).
        after_unique_id: Place moved task after this task's UniqueID (required).
    """
    app  = get_app()
    proj = get_proj(app)

    t = _find_task(proj, unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    after_t = _find_task(proj, after_unique_id)
    if after_t is None:
        return json.dumps({"error": f"Target task UniqueID {after_unique_id} not found."})

    old_id = t.ID
    old_level = t.OutlineLevel

    # Select and cut the source task row
    app.SelectRow(t.ID, False)
    app.EditCut()

    # Re-find the after task (IDs may have shifted after cut)
    after_t = _find_task(proj, after_unique_id)
    if after_t is None:
        # Paste back at original position as fallback
        app.EditPaste()
        return json.dumps({"error": "Target task not found after cut. Task restored."})

    # Select row after the target and paste
    target_row = after_t.ID + 1
    if target_row > proj.Tasks.Count:
        target_row = proj.Tasks.Count
    app.SelectRow(target_row, False)
    app.EditPaste()

    # Re-find moved task and verify
    moved = _find_task(proj, unique_id)
    new_id = moved.ID if moved else None

    return json.dumps({
        "status":    "moved",
        "unique_id": unique_id,
        "name":      moved.Name if moved else "(unknown)",
        "old_id":    old_id,
        "new_id":    new_id,
    }, indent=2)


@mcp.tool()
def get_progress_by_wbs(max_level: int = 2) -> str:
    """
    Rolled-up % complete per WBS branch — the PMO dashboard.
    Returns summary tasks at or below max_level with their percent complete.

    Args:
        max_level: Maximum outline level to report (default 2).
    """
    app  = get_app()
    proj = get_proj(app)

    branches = []
    tasks_list = [t for t in proj.Tasks if t is not None]

    for i, t in enumerate(tasks_list):
        if not t.Summary:
            continue
        if t.OutlineLevel > max_level:
            continue

        # Count children
        child_count = 0
        milestones_complete = 0
        milestones_total = 0
        for j in range(i + 1, len(tasks_list)):
            child = tasks_list[j]
            if child.OutlineLevel <= t.OutlineLevel:
                break
            if not child.Summary:
                child_count += 1
                if child.Milestone:
                    milestones_total += 1
                    if child.PercentComplete >= 100:
                        milestones_complete += 1

        branches.append({
            "unique_id":          t.UniqueID,
            "name":               t.Name,
            "level":              t.OutlineLevel,
            "percent_complete":   t.PercentComplete,
            "start":              _fmt_date(t.Start),
            "finish":             _fmt_date(t.Finish),
            "child_count":        child_count,
            "milestones_complete": milestones_complete,
            "milestones_total":   milestones_total,
        })

    return json.dumps({
        "max_level": max_level,
        "branches":  branches,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 4, Tier 2: Nice to Have
# ---------------------------------------------------------------------------

@mcp.tool()
def copy_task_structure(source_unique_id: int, copies: int = 1) -> str:
    """
    Duplicate a task subtree (reusable programme templates).

    Args:
        source_unique_id: UniqueID of the task to copy (and its children if summary).
        copies:           Number of copies to create (default 1).
    """
    app  = get_app()
    proj = get_proj(app)

    source = _find_task(proj, source_unique_id)
    if source is None:
        return json.dumps({"error": f"Task UniqueID {source_unique_id} not found."})

    # Determine range: source + all children (higher outline level)
    source_id = source.ID
    source_level = source.OutlineLevel
    end_id = source_id

    for t in proj.Tasks:
        if t is None:
            continue
        if t.ID > source_id:
            if t.OutlineLevel > source_level:
                end_id = t.ID
            else:
                break

    all_copied = []
    for _copy_num in range(copies):
        # Select the range
        app.SelectRow(source_id, False)
        app.SelectRow(end_id, True)  # extend selection
        app.EditCopy()

        # Paste at end
        last_id = proj.Tasks.Count
        app.SelectRow(last_id, False)
        app.EditPaste()

        # Collect newly created tasks
        new_count = proj.Tasks.Count
        for t in proj.Tasks:
            if t is not None and t.ID > last_id:
                all_copied.append({
                    "unique_id": t.UniqueID,
                    "name":      t.Name,
                })

    return json.dumps({
        "status":       "copied",
        "source_name":  source.Name,
        "copied_tasks": all_copied,
    }, indent=2)


@mcp.tool()
def cross_project_link(source_project: str, source_unique_id: int, target_project: str, target_unique_id: int, link_type: str = "FS") -> str:
    """
    Create a dependency link across open projects.

    Args:
        source_project:    Name of the predecessor's project.
        source_unique_id:  UniqueID of the predecessor task.
        target_project:    Name of the successor's project.
        target_unique_id:  UniqueID of the successor task.
        link_type:         'FS' (default), 'SS', 'FF', or 'SF'.
    """
    app = get_app(require_project=False)

    # Find source project and task
    src_proj = None
    for i in range(1, app.Projects.Count + 1):
        p = app.Projects(i)
        if p.Name.lower() == source_project.lower() or source_project.lower() in p.Name.lower():
            src_proj = p
            break
    if src_proj is None:
        return json.dumps({"error": f"Source project '{source_project}' not found."})

    src_task = _find_task(src_proj, source_unique_id)
    if src_task is None:
        return json.dumps({"error": f"Source task UniqueID {source_unique_id} not found in '{src_proj.Name}'."})

    # Find target project and task
    tgt_proj = None
    for i in range(1, app.Projects.Count + 1):
        p = app.Projects(i)
        if p.Name.lower() == target_project.lower() or target_project.lower() in p.Name.lower():
            tgt_proj = p
            break
    if tgt_proj is None:
        return json.dumps({"error": f"Target project '{target_project}' not found."})

    tgt_task = _find_task(tgt_proj, target_unique_id)
    if tgt_task is None:
        return json.dumps({"error": f"Target task UniqueID {target_unique_id} not found in '{tgt_proj.Name}'."})

    # Set cross-project predecessor using "ProjectName\TaskID" format
    pred_str = f"{src_proj.Name}\\{src_task.ID}{link_type}"
    existing = (tgt_task.Predecessors or "").strip()
    if existing:
        tgt_task.Predecessors = existing + "," + pred_str
    else:
        tgt_task.Predecessors = pred_str

    return json.dumps({
        "status":  "linked",
        "source":  {"project": src_proj.Name, "task": src_task.Name, "unique_id": source_unique_id},
        "target":  {"project": tgt_proj.Name, "task": tgt_task.Name, "unique_id": target_unique_id},
        "link_type": link_type,
    }, indent=2)


@mcp.tool()
def export_csv(output_path: str, columns_json: str = "", filters_json: str = "") -> str:
    """
    Export filtered task data to CSV for PowerBI / Excel dashboards.

    Args:
        output_path:  Full path for the output CSV file (required).
        columns_json: JSON list of column names to include (optional, default: standard set).
                      Available: any key from task_to_dict output.
        filters_json: JSON object with filter criteria (same format as filter_tasks).
    """
    import csv

    app  = get_app()
    proj = get_proj(app)

    # Default columns
    default_cols = ["unique_id", "name", "outline_level", "start", "finish",
                    "duration_days", "percent_complete", "resource_names", "rag", "critical"]

    columns = default_cols
    if columns_json:
        columns = json.loads(columns_json)

    # Apply filters if provided
    tasks = []
    if filters_json:
        f = json.loads(filters_json)
        # Reuse filter_tasks logic inline
        result = json.loads(filter_tasks(json.dumps(f)))
        tasks = result.get("tasks", [])
    else:
        for t in proj.Tasks:
            if t is not None:
                tasks.append(task_to_dict(t, proj))

    # Write CSV
    with open(output_path, "w", newline="", encoding="utf-8") as fp:
        writer = csv.writer(fp)
        writer.writerow(columns)
        for task in tasks:
            row = [task.get(col, "") for col in columns]
            writer.writerow(row)

    return json.dumps({
        "status":  "exported",
        "path":    output_path,
        "rows":    len(tasks),
        "columns": columns,
    }, indent=2)


@mcp.tool()
def bulk_set_deadlines(deadlines_json: str) -> str:
    """
    Set deadlines on multiple tasks at once.

    Args:
        deadlines_json: JSON string — list of {unique_id, deadline_date}.
            deadline_date can be 'clear' to remove the deadline.
            Example: '[{"unique_id": 42, "deadline_date": "2026-06-01"},
                       {"unique_id": 55, "deadline_date": "clear"}]'
    """
    items = json.loads(deadlines_json)
    app   = get_app()
    proj  = get_proj(app)

    set_count = 0
    cleared = 0
    not_found = []
    errors = []

    for item in items:
        uid = item["unique_id"]
        dd  = item["deadline_date"]

        t = _find_task(proj, uid)
        if t is None:
            not_found.append(uid)
            continue

        try:
            if dd.lower() == "clear":
                t.Deadline = "NA"
                cleared += 1
            else:
                t.Deadline = _parse_date(dd)
                set_count += 1
        except Exception as e:
            errors.append({"unique_id": uid, "error": str(e)})

    return json.dumps({
        "set":       set_count,
        "cleared":   cleared,
        "not_found": not_found,
        "errors":    errors,
    }, indent=2)


@mcp.tool()
def find_available_slack(min_days: int = 5) -> str:
    """
    Find tasks with positive float — where can we absorb delay?

    Args:
        min_days: Minimum total slack in working days (default 5).
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    tasks = []
    for t in proj.Tasks:
        if t is None or t.Summary:
            continue

        try:
            ts = t.TotalSlack
            if ts is None:
                continue
            ts_days = round(ts / mpd, 2)
            if ts_days < min_days:
                continue

            fs = 0
            try:
                fs = round(t.FreeSlack / mpd, 2) if t.FreeSlack else 0
            except Exception:
                pass

            tasks.append({
                "unique_id":        t.UniqueID,
                "name":             t.Name,
                "total_slack_days": ts_days,
                "free_slack_days":  fs,
                "start":            _fmt_date(t.Start),
                "finish":           _fmt_date(t.Finish),
                "resource_names":   t.ResourceNames or "",
            })
        except Exception:
            continue

    tasks.sort(key=lambda x: x["total_slack_days"], reverse=True)

    return json.dumps({
        "min_days": min_days,
        "count":    len(tasks),
        "tasks":    tasks,
    }, indent=2)


@mcp.tool()
def set_task_calendar(unique_id: int, calendar_name: str) -> str:
    """
    Set a task-level calendar override (e.g. 24/7 for commissioning phase).

    Args:
        unique_id:     Task UniqueID (required).
        calendar_name: Calendar name to apply, or empty string '' to clear.
    """
    app  = get_app()
    proj = get_proj(app)

    t = _find_task(proj, unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    # Validate calendar exists (if not clearing)
    if calendar_name:
        valid_cals = []
        try:
            for cal in proj.BaseCalendars:
                if cal is not None:
                    valid_cals.append(str(cal.Name))
        except Exception:
            pass
        if calendar_name not in valid_cals:
            return json.dumps({"error": f"Calendar '{calendar_name}' not found. Available: {valid_cals}"})

    previous = ""
    try:
        previous = str(t.Calendar) if t.Calendar else ""
    except Exception:
        pass

    try:
        if calendar_name:
            t.Calendar = calendar_name
        else:
            t.Calendar = ""
    except Exception as e:
        return json.dumps({"error": f"Failed to set task calendar: {e}"})

    return json.dumps({
        "status":    "updated",
        "unique_id": unique_id,
        "name":      t.Name,
        "calendar":  calendar_name or "(cleared)",
        "previous":  previous,
    }, indent=2)


@mcp.tool()
def get_cost_summary() -> str:
    """
    Budget rollup with cost fields.
    Returns project cost totals, cost breakdown by resource, and tasks with cost data.
    Requires cost data to be entered on tasks or resources.
    """
    app  = get_app()
    proj = get_proj(app)

    totals = {"cost": 0, "actual_cost": 0, "remaining_cost": 0, "baseline_cost": 0}
    by_resource = {}
    tasks_with_cost = []

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue

        cost = actual = remaining = baseline = 0
        try:
            cost = float(t.Cost) if t.Cost else 0
        except Exception:
            pass
        try:
            actual = float(t.ActualCost) if t.ActualCost else 0
        except Exception:
            pass
        try:
            remaining = float(t.RemainingCost) if t.RemainingCost else 0
        except Exception:
            pass
        try:
            baseline = float(t.BaselineCost) if t.BaselineCost else 0
        except Exception:
            pass

        totals["cost"] += cost
        totals["actual_cost"] += actual
        totals["remaining_cost"] += remaining
        totals["baseline_cost"] += baseline

        if cost > 0:
            tasks_with_cost.append({
                "unique_id":      t.UniqueID,
                "name":           t.Name,
                "cost":           cost,
                "actual_cost":    actual,
                "remaining_cost": remaining,
                "baseline_cost":  baseline,
            })

    # Cost by resource
    try:
        for r in proj.Resources:
            if r is not None:
                r_cost = 0
                r_actual = 0
                try:
                    r_cost = float(r.Cost) if r.Cost else 0
                except Exception:
                    pass
                try:
                    r_actual = float(r.ActualCost) if r.ActualCost else 0
                except Exception:
                    pass
                if r_cost > 0 or r_actual > 0:
                    by_resource[r.Name] = {"cost": r_cost, "actual_cost": r_actual}
    except Exception:
        pass

    totals["variance"] = totals["baseline_cost"] - totals["cost"]

    return json.dumps({
        "project": proj.Name,
        "totals":  totals,
        "by_resource":     list({"name": k, **v} for k, v in by_resource.items()),
        "tasks_with_cost": tasks_with_cost,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 4, Tier 3: Power User
# ---------------------------------------------------------------------------

@mcp.tool()
def undo_last(count: int = 1) -> str:
    """
    Safety net — undo last N operations in MS Project.

    Args:
        count: Number of undo steps (default 1, max 10).
    """
    if count < 1:
        count = 1
    if count > 10:
        count = 10

    app = get_app()
    for _ in range(count):
        try:
            app.EditUndo()
        except Exception:
            break

    return json.dumps({"status": "undone", "undo_count": count}, indent=2)


@mcp.tool()
def create_calendar(name: str, copy_from: str = "Standard") -> str:
    """
    Create a new base calendar, optionally copying from an existing one.

    Args:
        name:      Name for the new calendar (required).
        copy_from: Existing calendar to copy from (default 'Standard').
    """
    app  = get_app()
    proj = get_proj(app)

    # Validate copy_from exists
    valid_cals = []
    try:
        for cal in proj.BaseCalendars:
            if cal is not None:
                valid_cals.append(str(cal.Name))
    except Exception:
        pass

    if copy_from not in valid_cals:
        return json.dumps({"error": f"Calendar '{copy_from}' not found. Available: {valid_cals}"})

    if name in valid_cals:
        return json.dumps({"error": f"Calendar '{name}' already exists."})

    try:
        app.BaseCalendarCreate(Name=name, FromName=copy_from)
    except Exception:
        try:
            app.BaseCalendarCreate(name, copy_from)
        except Exception as e:
            return json.dumps({"error": f"Failed to create calendar: {e}"})

    # Re-read calendar list
    calendars = []
    try:
        for cal in proj.BaseCalendars:
            if cal is not None:
                calendars.append(str(cal.Name))
    except Exception:
        pass

    return json.dumps({
        "status":     "created",
        "name":       name,
        "copied_from": copy_from,
        "calendars":  calendars,
    }, indent=2)


@mcp.tool()
def insert_subproject(file_path: str, after_unique_id: int = 0) -> str:
    """
    Insert an external .mpp file as a subproject.

    Args:
        file_path:       Full path to the .mpp file to insert (required).
        after_unique_id: Insert after this task's UniqueID (0 = insert at end).
    """
    import os
    if not os.path.exists(file_path):
        return json.dumps({"error": f"File not found: {file_path}"})

    app  = get_app()
    proj = get_proj(app)

    count_before = proj.Tasks.Count

    # Insert by adding a task and setting its SubProject property.
    # (SubprojectInsert is not reliably exposed via COM in all versions.)
    try:
        import os as _os
        basename = _os.path.splitext(_os.path.basename(file_path))[0]
        if after_unique_id > 0:
            t = _find_task(proj, after_unique_id)
            if t is None:
                return json.dumps({"error": f"Task UniqueID {after_unique_id} not found."})
            app.SelectRow(t.ID + 1, False)
            new_t = proj.Tasks.Add(basename)
        else:
            new_t = proj.Tasks.Add(basename)
        new_t.SubProject = file_path
    except Exception as e:
        return json.dumps({"error": f"Failed to insert subproject: {e}"})

    count_after = proj.Tasks.Count

    return json.dumps({
        "status":           "inserted",
        "file_path":        file_path,
        "inserted_after":   after_unique_id or "end",
        "task_count_before": count_before,
        "task_count_after":  count_after,
    }, indent=2)


@mcp.tool()
def apply_filter(filter_name: str) -> str:
    """
    Apply MS Project's built-in or custom named filter to the Gantt view.
    Use 'All Tasks' to clear any active filter.

    Args:
        filter_name: Name of the filter (e.g. 'Critical', 'Incomplete Tasks', 'All Tasks').
    """
    app = get_app()

    try:
        app.FilterApply(filter_name)
    except Exception as e:
        return json.dumps({"error": f"Failed to apply filter '{filter_name}': {e}"})

    return json.dumps({
        "status": "applied",
        "filter": filter_name,
    }, indent=2)


@mcp.tool()
def snapshot_to_json(output_path: str, include_resources: bool = True) -> str:
    """
    Full project state dump for version control / diff.
    Exports all tasks and optionally resources to a JSON file.

    Args:
        output_path:       Full path for the output JSON file (required).
        include_resources: Include resource data (default True).
    """
    app  = get_app()
    proj = get_proj(app)

    tasks = []
    for t in proj.Tasks:
        if t is not None:
            tasks.append(task_to_dict(t, proj))

    resources = []
    if include_resources:
        try:
            for r in proj.Resources:
                if r is not None:
                    resources.append({
                        "unique_id":  r.UniqueID,
                        "id":         r.ID,
                        "name":       r.Name,
                        "initials":   r.Initials,
                        "type":       r.Type,
                        "max_units":  r.MaxUnits,
                        "cost":       r.StandardRate,
                        "task_count": r.Assignments.Count,
                    })
        except Exception:
            pass

    project_meta = {
        "name":   proj.Name,
        "start":  _fmt_date(proj.ProjectStart),
        "finish": _fmt_date(proj.ProjectFinish),
    }
    try:
        project_meta["title"]   = proj.Title or ""
        project_meta["manager"] = proj.Manager or ""
    except Exception:
        pass

    snapshot = {
        "project":   project_meta,
        "tasks":     tasks,
        "resources": resources,
    }

    with open(output_path, "w", encoding="utf-8") as fp:
        json.dump(snapshot, fp, indent=2, default=str)

    return json.dumps({
        "status":    "exported",
        "path":      output_path,
        "tasks":     len(tasks),
        "resources": len(resources),
        "project":   project_meta["name"],
    }, indent=2)


# ---------------------------------------------------------------------------
# Phase 5 — get_constraints, delete_resource, get_actual_work
# ---------------------------------------------------------------------------

@mcp.tool()
def get_constraints() -> str:
    """Return all tasks with non-default (non-ASAP) scheduling constraints."""
    CONSTRAINT_NAMES = {
        0: "ASAP", 1: "ALAP", 2: "MSO", 3: "MFO",
        4: "SNET", 5: "SNLT", 6: "FNET", 7: "FNLT",
    }
    app  = get_app()
    proj = get_proj(app)

    results = []
    for t in proj.Tasks:
        if t is None or t.Summary:
            continue
        try:
            ct = t.ConstraintType
            if ct != 0:  # 0 = ASAP (default)
                results.append({
                    "unique_id":       t.UniqueID,
                    "name":            t.Name,
                    "constraint_type": CONSTRAINT_NAMES.get(ct, f"Unknown({ct})"),
                    "constraint_date": _fmt_date(t.ConstraintDate),
                })
        except Exception:
            continue

    return json.dumps({"count": len(results), "tasks": results}, indent=2)


@mcp.tool()
def delete_resource(resource_name: str) -> str:
    """
    Delete a resource from the project pool (case-insensitive match).
    All task assignments referencing this resource are cleared first.
    """
    app  = get_app()
    proj = get_proj(app)

    target = None
    for r in proj.Resources:
        if r is not None and r.Name.lower() == resource_name.lower():
            target = r
            break

    if target is None:
        return json.dumps({"error": f"Resource '{resource_name}' not found."})

    # Clear assignments first
    assignments_cleared = 0
    # Iterate in reverse to avoid index shifting
    for i in range(target.Assignments.Count, 0, -1):
        try:
            target.Assignments(i).Delete()
            assignments_cleared += 1
        except Exception:
            pass

    name = target.Name
    try:
        target.Delete()
    except Exception as e:
        return json.dumps({"error": f"Failed to delete resource: {e}"})

    app.FileSave()
    return json.dumps({
        "status":              "deleted",
        "name":                name,
        "assignments_cleared": assignments_cleared,
    }, indent=2)


@mcp.tool()
def get_actual_work() -> str:
    """
    Return actual vs remaining work per task and project totals.
    Work values are converted from COM minutes to hours.
    """
    app  = get_app()
    proj = get_proj(app)

    tasks = []
    total_work = 0.0
    total_actual = 0.0
    total_remaining = 0.0

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue
        try:
            w = (t.Work or 0) / 60.0
            a = (t.ActualWork or 0) / 60.0
            r = (t.RemainingWork or 0) / 60.0
            total_work += w
            total_actual += a
            total_remaining += r
            tasks.append({
                "unique_id":      t.UniqueID,
                "name":           t.Name,
                "work_hours":     round(w, 2),
                "actual_hours":   round(a, 2),
                "remaining_hours": round(r, 2),
                "pct_work_complete": t.PercentWorkComplete,
            })
        except Exception:
            continue

    pct = round(total_actual / total_work * 100, 1) if total_work else 0.0

    return json.dumps({
        "totals": {
            "work_hours":        round(total_work, 2),
            "actual_hours":      round(total_actual, 2),
            "remaining_hours":   round(total_remaining, 2),
            "pct_work_complete": pct,
        },
        "count": len(tasks),
        "tasks": tasks,
    }, indent=2)


# ---------------------------------------------------------------------------
# Phase 6 — Gap Analysis & Missing Features
# ---------------------------------------------------------------------------


@mcp.tool()
def calculate_project() -> str:
    """
    Recalculate the active project schedule.
    Use after bulk manual changes to ensure dates, slack, and critical path are up to date.
    """
    app = get_app()
    app.CalculateProject()
    return json.dumps({"status": "calculated", "project": app.ActiveProject.Name})


@mcp.tool()
def list_calendar_exceptions(calendar_name: str = "") -> str:
    """
    List all exceptions (holidays, non-working days) defined on a calendar.
    If calendar_name is empty, uses the project calendar.

    Args:
        calendar_name: Name of the base calendar. Empty = project calendar.
    """
    app  = get_app()
    proj = get_proj(app)

    if not calendar_name:
        try:
            calendar_name = str(proj.Calendar)
        except Exception:
            calendar_name = "Standard"

    cal = None
    for c in proj.BaseCalendars:
        if c is not None and str(c.Name).lower() == str(calendar_name).lower():
            cal = c
            break
    if cal is None:
        return json.dumps({"error": f"Calendar '{calendar_name}' not found."})

    exceptions = []
    try:
        for exc in cal.Exceptions:
            exceptions.append({
                "name":   str(exc.Name),
                "start":  _fmt_date(exc.Start),
                "finish": _fmt_date(exc.Finish),
                "type":   exc.Type,
            })
    except Exception:
        pass  # Some calendars have no Exceptions collection

    return json.dumps({
        "calendar":   str(cal.Name),
        "count":      len(exceptions),
        "exceptions": exceptions,
    }, indent=2)


@mcp.tool()
def health_check() -> str:
    """
    Lightweight connectivity test. Returns MS Project version, whether a project
    is open, and basic project info if available.
    """
    import win32com.client
    try:
        app = win32com.client.GetActiveObject("MSProject.Application")
    except Exception:
        return json.dumps({"status": "disconnected", "error": "MS Project is not running."})

    result = {
        "status":  "connected",
        "version": str(app.Version),
    }

    if app.Projects.Count > 0:
        proj = app.ActiveProject
        result["project_open"] = True
        result["project_name"] = proj.Name
        result["task_count"]   = proj.Tasks.Count
    else:
        result["project_open"] = False

    return json.dumps(result, indent=2)


@mcp.tool()
def update_project(complete_through: str, set_0_or_100: bool = False) -> str:
    """
    Mark all tasks complete through a given date (the weekly PMO ritual).
    Tasks that should have finished by the date get their % complete updated.

    Args:
        complete_through: Date as YYYY-MM-DD — tasks scheduled through this date are updated.
        set_0_or_100:     If True, tasks are set to 0% or 100% only (no partial). Default False.
    """
    app  = get_app()
    proj = get_proj(app)
    dt   = _parse_date(complete_through)
    if dt is None:
        return json.dumps({"error": "complete_through date is required (YYYY-MM-DD)."})

    # COM VBA signature: UpdateProject(All, UpdateDate, Action)
    # All = True (entire project), UpdateDate = date, Action:
    #   pjUpdateProjectStatusPctComplete = 0
    #   pjUpdateProject0or100 = 1
    action = 1 if set_0_or_100 else 0
    try:
        app.UpdateProject(True, dt, action)
    except Exception:
        try:
            # Alternative: just date
            app.UpdateProject(True, dt)
        except Exception:
            app.UpdateProject(dt)
    app.FileSave()

    return json.dumps({
        "status": "updated",
        "complete_through": complete_through,
        "set_0_or_100": set_0_or_100,
        "project": proj.Name,
    }, indent=2)


@mcp.tool()
def reschedule_incomplete_work(reschedule_from: str = "") -> str:
    """
    Move remaining work on incomplete tasks to start after the given date
    (or the project status date if not specified).

    Args:
        reschedule_from: Date as YYYY-MM-DD. Empty = use project status date.
    """
    app  = get_app()
    proj = get_proj(app)

    if reschedule_from:
        dt = _parse_date(reschedule_from)
    else:
        dt = proj.StatusDate
        if dt is None or str(dt) == "NA":
            dt = datetime.datetime.now()

    # Set status date, then use UpdateProject to reschedule
    # The reschedule action = updating incomplete tasks from the status date
    try:
        proj.StatusDate = dt
        # UpdateProject with All=True, date, action=0 to push remaining work
        app.UpdateProject(True, dt, 0)
    except Exception:
        try:
            app.UpdateProject(True, dt)
        except Exception:
            proj.StatusDate = dt  # At minimum set the status date
    app.FileSave()

    return json.dumps({
        "status": "rescheduled",
        "reschedule_from": str(dt)[:10],
        "project": proj.Name,
    }, indent=2)


@mcp.tool()
def delete_calendar(calendar_name: str) -> str:
    """
    Delete a base calendar by name. Cannot delete the project calendar.

    Args:
        calendar_name: Name of the calendar to delete.
    """
    app  = get_app()
    proj = get_proj(app)

    try:
        proj_cal = str(proj.Calendar).lower()
    except Exception:
        proj_cal = "standard"
    if proj_cal == calendar_name.lower():
        return json.dumps({"error": "Cannot delete the active project calendar."})

    for c in proj.BaseCalendars:
        if c is not None and str(c.Name).lower() == calendar_name.lower():
            c.Delete()
            app.FileSave()
            return json.dumps({"status": "deleted", "calendar": calendar_name})

    return json.dumps({"error": f"Calendar '{calendar_name}' not found."})


@mcp.tool()
def delete_calendar_exception(calendar_name: str, exception_name: str) -> str:
    """
    Remove a specific exception (holiday/non-working day) from a calendar.

    Args:
        calendar_name:  Name of the base calendar.
        exception_name: Name of the exception to remove.
    """
    app  = get_app()
    proj = get_proj(app)

    cal = None
    for c in proj.BaseCalendars:
        if c is not None and str(c.Name).lower() == calendar_name.lower():
            cal = c
            break
    if cal is None:
        return json.dumps({"error": f"Calendar '{calendar_name}' not found."})

    try:
        for exc in cal.Exceptions:
            if str(exc.Name).lower() == exception_name.lower():
                exc.Delete()
                app.FileSave()
                return json.dumps({
                    "status":    "deleted",
                    "calendar":  calendar_name,
                    "exception": exception_name,
                })
    except Exception as e:
        return json.dumps({"error": f"Failed to access exceptions: {e}"})

    return json.dumps({"error": f"Exception '{exception_name}' not found in calendar '{calendar_name}'."})


@mcp.tool()
def set_resource_calendar(resource_name: str, calendar_name: str) -> str:
    """
    Assign a specific base calendar to a resource (e.g., part-time, different timezone).

    Args:
        resource_name: Name of the resource.
        calendar_name: Name of the base calendar to assign.
    """
    app  = get_app()
    proj = get_proj(app)

    # Verify calendar exists
    cal_found = False
    for c in proj.BaseCalendars:
        if c is not None and str(c.Name).lower() == calendar_name.lower():
            cal_found = True
            break
    if not cal_found:
        return json.dumps({"error": f"Calendar '{calendar_name}' not found."})

    for r in proj.Resources:
        if r is not None and r.Name and r.Name.lower() == resource_name.lower():
            r.BaseCalendar = calendar_name
            app.FileSave()
            return json.dumps({
                "status":   "updated",
                "resource": r.Name,
                "calendar": calendar_name,
            }, indent=2)

    return json.dumps({"error": f"Resource '{resource_name}' not found."})


@mcp.tool()
def get_timephased_data(
    unique_id:  int,
    start_date: str,
    end_date:   str,
    timescale:  str = "weekly",
    data_type:  str = "work",
) -> str:
    """
    Get period-by-period timephased data for a task. Essential for S-curves,
    resource loading charts, and cash flow forecasts.

    Args:
        unique_id:  Task UniqueID.
        start_date: Period start as YYYY-MM-DD.
        end_date:   Period end as YYYY-MM-DD.
        timescale:  'daily', 'weekly', or 'monthly' (default 'weekly').
        data_type:  'work', 'cost', 'actual_work', 'actual_cost',
                    'remaining_work', 'baseline_work', 'baseline_cost' (default 'work').
    """
    app  = get_app()
    proj = get_proj(app)

    TIMESCALE_MAP = {"daily": 3, "weekly": 4, "monthly": 5}
    ts = TIMESCALE_MAP.get(timescale.lower())
    if ts is None:
        return json.dumps({"error": f"Unknown timescale '{timescale}'. Use: daily, weekly, monthly."})

    # pjTaskTimescaledWork=1, Cost=2, ActualWork=3, ActualCost=4,
    # RemainingWork=9, BaselineWork=22, BaselineCost=23
    TYPE_MAP = {
        "work": 1, "cost": 2, "actual_work": 3, "actual_cost": 4,
        "remaining_work": 9, "baseline_work": 22, "baseline_cost": 23,
    }
    dt = TYPE_MAP.get(data_type.lower())
    if dt is None:
        return json.dumps({"error": f"Unknown data_type '{data_type}'. Use: {list(TYPE_MAP.keys())}."})

    t = _find_task(proj, unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    sd = _parse_date(start_date)
    ed = _parse_date(end_date)
    if sd is None or ed is None:
        return json.dumps({"error": "Both start_date and end_date are required (YYYY-MM-DD)."})

    periods = []
    try:
        tsd = t.TimeScaleData(sd, ed, dt, ts)
        for item in tsd:
            val = item.Value
            periods.append({
                "start": _fmt_date(item.StartDate),
                "end":   _fmt_date(item.EndDate),
                "value": float(val) if val else 0.0,
            })
    except Exception as e:
        return json.dumps({"error": f"TimeScaleData failed: {e}"})

    return json.dumps({
        "unique_id": unique_id,
        "name":      t.Name,
        "data_type": data_type,
        "timescale": timescale,
        "periods":   periods,
    }, indent=2)


@mcp.tool()
def set_working_hours(calendar_name: str, day: int, shifts_json: str) -> str:
    """
    Modify working hours for a specific day of the week in a calendar.

    Args:
        calendar_name: Name of the base calendar.
        day:           Day number (1=Sunday, 2=Monday, ..., 7=Saturday).
        shifts_json:   JSON array of shifts, e.g. [["08:00","12:00"],["13:00","17:00"]].
                       Empty array [] marks the day as non-working.
    """
    app  = get_app()
    proj = get_proj(app)

    if day < 1 or day > 7:
        return json.dumps({"error": "day must be 1 (Sunday) through 7 (Saturday)."})

    cal = None
    for c in proj.BaseCalendars:
        if c is not None and str(c.Name).lower() == calendar_name.lower():
            cal = c
            break
    if cal is None:
        return json.dumps({"error": f"Calendar '{calendar_name}' not found."})

    shifts = json.loads(shifts_json)

    wd = cal.WeekDays(day)

    if not shifts:
        # Mark as non-working
        wd.Working = False
        app.FileSave()
        return json.dumps({
            "status":   "updated",
            "calendar": cal.Name,
            "day":      day,
            "working":  False,
        }, indent=2)

    wd.Working = True

    # Set shifts using ShiftN Start/Finish properties (1-indexed)
    # Clear all 5 shifts first, then set the provided ones
    shift_attrs = [
        ("Shift1Start", "Shift1Finish"),
        ("Shift2Start", "Shift2Finish"),
        ("Shift3Start", "Shift3Finish"),
        ("Shift4Start", "Shift4Finish"),
        ("Shift5Start", "Shift5Finish"),
    ]

    for i, (s_attr, f_attr) in enumerate(shift_attrs):
        try:
            if i < len(shifts):
                setattr(wd, s_attr, shifts[i][0])
                setattr(wd, f_attr, shifts[i][1])
            else:
                # Clear unused shifts
                setattr(wd, s_attr, "")
                setattr(wd, f_attr, "")
        except Exception:
            pass

    app.FileSave()
    return json.dumps({
        "status":   "updated",
        "calendar": cal.Name,
        "day":      day,
        "working":  True,
        "shifts":   shifts[:5],
    }, indent=2)


@mcp.tool()
def get_resource_availability(
    resource_name: str,
    start_date:    str,
    end_date:      str,
    timescale:     str = "weekly",
) -> str:
    """
    Show resource allocation vs capacity per period. Shows max units, allocated
    work, and free capacity windows.

    Args:
        resource_name: Name of the resource.
        start_date:    Period start as YYYY-MM-DD.
        end_date:      Period end as YYYY-MM-DD.
        timescale:     'daily', 'weekly', or 'monthly' (default 'weekly').
    """
    app  = get_app()
    proj = get_proj(app)

    TIMESCALE_MAP = {"daily": 3, "weekly": 4, "monthly": 5}
    ts = TIMESCALE_MAP.get(timescale.lower())
    if ts is None:
        return json.dumps({"error": f"Unknown timescale '{timescale}'. Use: daily, weekly, monthly."})

    res = None
    for r in proj.Resources:
        if r is not None and r.Name and r.Name.lower() == resource_name.lower():
            res = r
            break
    if res is None:
        return json.dumps({"error": f"Resource '{resource_name}' not found."})

    sd = _parse_date(start_date)
    ed = _parse_date(end_date)

    max_units = res.MaxUnits  # e.g. 1.0 = 100%

    periods = []
    try:
        # Resource TimeScaleData types differ from Task types:
        # Type 13 = pjResourceTimescaledWork (minutes), Type 4 = Availability (units)
        tsd = res.TimeScaleData(sd, ed, 13, ts)
        for item in tsd:
            try:
                val = item.Value
                allocated_hrs = float(val) / 60.0 if val else 0.0
            except Exception:
                allocated_hrs = 0.0
            periods.append({
                "start":          _fmt_date(item.StartDate),
                "end":            _fmt_date(item.EndDate),
                "allocated_hours": round(allocated_hrs, 2),
            })
    except Exception as e:
        return json.dumps({"error": f"TimeScaleData failed: {e}"})

    return json.dumps({
        "resource":  res.Name,
        "max_units": max_units,
        "timescale": timescale,
        "periods":   periods,
    }, indent=2)


@mcp.tool()
def get_variance_report(baseline: int = 0) -> str:
    """
    Schedule and cost variance per task compared to a baseline.

    Args:
        baseline: Baseline number (0-10). Default 0.
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    # Baseline property name mapping
    BL_START  = f"Baseline{'' if baseline == 0 else baseline}Start"
    BL_FINISH = f"Baseline{'' if baseline == 0 else baseline}Finish"
    BL_COST   = f"Baseline{'' if baseline == 0 else baseline}Cost"

    tasks = []
    for t in proj.Tasks:
        if t is None or t.Summary:
            continue
        try:
            bl_start  = getattr(t, BL_START, None)
            bl_finish = getattr(t, BL_FINISH, None)
            bl_cost   = getattr(t, BL_COST, 0) or 0

            sv_days = round(t.StartVariance / mpd, 2) if t.StartVariance else 0
            fv_days = round(t.FinishVariance / mpd, 2) if t.FinishVariance else 0
            cv      = round((t.Cost or 0) - bl_cost, 2)

            tasks.append({
                "unique_id":        t.UniqueID,
                "name":             t.Name,
                "start":            _fmt_date(t.Start),
                "finish":           _fmt_date(t.Finish),
                "baseline_start":   _fmt_date(bl_start),
                "baseline_finish":  _fmt_date(bl_finish),
                "start_variance_days":  sv_days,
                "finish_variance_days": fv_days,
                "cost":             round(t.Cost or 0, 2),
                "baseline_cost":    round(bl_cost, 2),
                "cost_variance":    cv,
            })
        except Exception:
            continue

    # Filter to only tasks with actual variance
    with_variance = [t for t in tasks if t["start_variance_days"] != 0 or t["finish_variance_days"] != 0 or t["cost_variance"] != 0]

    return json.dumps({
        "baseline":       baseline,
        "total_tasks":    len(tasks),
        "with_variance":  len(with_variance),
        "tasks":          with_variance if with_variance else tasks[:50],
    }, indent=2)


@mcp.tool()
def snapshot_diff(path_a: str, path_b: str) -> str:
    """
    Compare two JSON snapshot files (from snapshot_to_json) and return
    additions, deletions, and changes.

    Args:
        path_a: Path to the earlier snapshot JSON.
        path_b: Path to the later snapshot JSON.
    """
    import os

    for p in (path_a, path_b):
        if not os.path.exists(p):
            return json.dumps({"error": f"File not found: {p}"})

    with open(path_a, "r", encoding="utf-8") as f:
        snap_a = json.load(f)
    with open(path_b, "r", encoding="utf-8") as f:
        snap_b = json.load(f)

    tasks_a = {t["unique_id"]: t for t in snap_a.get("tasks", [])}
    tasks_b = {t["unique_id"]: t for t in snap_b.get("tasks", [])}

    ids_a = set(tasks_a.keys())
    ids_b = set(tasks_b.keys())

    added   = [tasks_b[uid] for uid in sorted(ids_b - ids_a)]
    deleted = [tasks_a[uid] for uid in sorted(ids_a - ids_b)]

    changed = []
    for uid in sorted(ids_a & ids_b):
        a, b = tasks_a[uid], tasks_b[uid]
        diffs = {}
        for key in set(list(a.keys()) + list(b.keys())):
            va, vb = a.get(key), b.get(key)
            if va != vb:
                diffs[key] = {"from": va, "to": vb}
        if diffs:
            changed.append({"unique_id": uid, "name": b.get("name", a.get("name")), "changes": diffs})

    return json.dumps({
        "added_count":   len(added),
        "deleted_count": len(deleted),
        "changed_count": len(changed),
        "added":         added,
        "deleted":       deleted,
        "changed":       changed,
    }, indent=2)


@mcp.tool()
def set_task_hyperlink(unique_id: int, url: str, text: str = "", sub_address: str = "") -> str:
    """
    Set a hyperlink on a task.

    Args:
        unique_id:   Task UniqueID.
        url:         The hyperlink URL or file path.
        text:        Display text for the hyperlink (optional).
        sub_address: Sub-address / bookmark within the target (optional).
    """
    app  = get_app()
    proj = get_proj(app)
    t = _find_task(proj, unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    t.HyperlinkAddress    = url
    t.HyperlinkScreenTip  = text or url
    if text:
        t.Hyperlink = text
    if sub_address:
        t.HyperlinkSubAddress = sub_address

    app.FileSave()
    return json.dumps({
        "status":    "updated",
        "unique_id": unique_id,
        "name":      t.Name,
        "hyperlink": url,
        "text":      text or url,
    }, indent=2)


@mcp.tool()
def add_recurring_task(
    name:            str,
    recurrence_type: str = "weekly",
    start_date:      str = "",
    end_date:        str = "",
    duration_days:   float = 1,
    day_of_week:     int = 2,
) -> str:
    """
    Add a recurring task to the project.

    Args:
        name:            Task name.
        recurrence_type: 'daily', 'weekly', or 'monthly' (default 'weekly').
        start_date:      Recurrence range start (YYYY-MM-DD).
        end_date:        Recurrence range end (YYYY-MM-DD).
        duration_days:   Duration of each occurrence in days (default 1).
        day_of_week:     For weekly: 1=Sun, 2=Mon, ..., 7=Sat (default 2=Monday).
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    sd = _parse_date(start_date)
    ed = _parse_date(end_date)
    if sd is None or ed is None:
        return json.dumps({"error": "Both start_date and end_date are required (YYYY-MM-DD)."})

    dur = int(duration_days * mpd)

    # pjRecurType: 0=daily, 1=weekly, 2=monthly, 3=yearly
    RECUR_MAP = {"daily": 0, "weekly": 1, "monthly": 2, "yearly": 3}
    rt = RECUR_MAP.get(recurrence_type.lower())
    if rt is None:
        return json.dumps({"error": f"Unknown recurrence_type '{recurrence_type}'. Use: daily, weekly, monthly, yearly."})

    try:
        # MS Project's RecurringTaskInsert is a dialog-only COM method;
        # it cannot be driven programmatically. Instead, we create individual
        # task occurrences under a summary task to simulate recurrence.

        from dateutil.rrule import rrule, DAILY, WEEKLY, MONTHLY, YEARLY
        FREQ_MAP = {"daily": DAILY, "weekly": WEEKLY, "monthly": MONTHLY, "yearly": YEARLY}
        freq = FREQ_MAP.get(recurrence_type.lower(), WEEKLY)

        # Map day_of_week (1=Sun..7=Sat) to dateutil byweekday (0=Mon..6=Sun)
        WEEKDAY_MAP = {1: 6, 2: 0, 3: 1, 4: 2, 5: 3, 6: 4, 7: 5}
        byday = WEEKDAY_MAP.get(day_of_week, 0)

        if freq == WEEKLY:
            dates = list(rrule(freq, dtstart=sd, until=ed, byweekday=byday))
        else:
            dates = list(rrule(freq, dtstart=sd, until=ed))

        if not dates:
            return json.dumps({"error": "No occurrences generated for the given range."})

        # Create summary task
        summary = proj.Tasks.Add(name)
        summary_uid = summary.UniqueID

        # Create each occurrence as a subtask
        occurrence_uids = []
        for i, dt_occ in enumerate(dates, 1):
            occ = proj.Tasks.Add(f"{name} #{i}")
            occ.OutlineIndent()
            occ.Start = dt_occ
            occ.Duration = dur
            occurrence_uids.append(occ.UniqueID)

        app.CalculateProject()
        app.FileSave()

        return json.dumps({
            "status":          "created",
            "name":            name,
            "unique_id":       summary_uid,
            "recurrence_type": recurrence_type,
            "occurrences":     len(dates),
            "start":           start_date,
            "end":             end_date,
        }, indent=2)

    except ImportError:
        return json.dumps({"error": "python-dateutil is required for recurring tasks. Install with: pip install python-dateutil"})
    except Exception as e:
        return json.dumps({"error": f"Failed to create recurring task: {e}"})


@mcp.tool()
def get_resource_rate_tables(resource_name: str) -> str:
    """
    Get cost rate tables (A through E) for a resource.

    Args:
        resource_name: Name of the resource.
    """
    app  = get_app()
    proj = get_proj(app)

    res = None
    for r in proj.Resources:
        if r is not None and r.Name and r.Name.lower() == resource_name.lower():
            res = r
            break
    if res is None:
        return json.dumps({"error": f"Resource '{resource_name}' not found."})

    tables = {}
    TABLE_NAMES = ["A", "B", "C", "D", "E"]

    for idx, tname in enumerate(TABLE_NAMES):
        try:
            table = res.CostRateTables(idx + 1)
            rates = []
            for pay_rate in table.PayRates:
                rates.append({
                    "effective_date":  _fmt_date(pay_rate.EffectiveDate),
                    "standard_rate":   str(pay_rate.StandardRate),
                    "overtime_rate":   str(pay_rate.OvertimeRate),
                    "cost_per_use":    float(pay_rate.CostPerUse) if pay_rate.CostPerUse else 0.0,
                })
            tables[tname] = rates
        except Exception:
            tables[tname] = []

    return json.dumps({
        "resource": res.Name,
        "tables":   tables,
    }, indent=2)


@mcp.tool()
def set_resource_rate_table(
    resource_name: str,
    table:         str = "A",
    standard_rate: str = "",
    overtime_rate: str = "",
    cost_per_use:  float = -1,
    effective_date: str = "",
) -> str:
    """
    Set or add a cost rate entry in a resource's rate table.

    Args:
        resource_name:  Name of the resource.
        table:          Rate table letter: A, B, C, D, or E (default A).
        standard_rate:  Standard rate as string, e.g. '50/h' or '400/d'.
        overtime_rate:  Overtime rate as string, e.g. '75/h'.
        cost_per_use:   Per-use cost (default -1 = don't change).
        effective_date: When this rate takes effect (YYYY-MM-DD). Empty = first entry.
    """
    app  = get_app()
    proj = get_proj(app)

    TABLE_MAP = {"A": 1, "B": 2, "C": 3, "D": 4, "E": 5}
    tbl_idx = TABLE_MAP.get(table.upper())
    if tbl_idx is None:
        return json.dumps({"error": f"Invalid table '{table}'. Use A-E."})

    res = None
    for r in proj.Resources:
        if r is not None and r.Name and r.Name.lower() == resource_name.lower():
            res = r
            break
    if res is None:
        return json.dumps({"error": f"Resource '{resource_name}' not found."})

    try:
        rate_table = res.CostRateTables(tbl_idx)
        pay_rates  = rate_table.PayRates

        if effective_date:
            # Add a new rate entry with effective date
            ed = _parse_date(effective_date)
            new_rate = pay_rates.Add(ed)
            if standard_rate:
                new_rate.StandardRate = standard_rate
            if overtime_rate:
                new_rate.OvertimeRate = overtime_rate
            if cost_per_use >= 0:
                new_rate.CostPerUse = cost_per_use
        else:
            # Update the first (default) rate entry
            first = pay_rates(1)
            if standard_rate:
                first.StandardRate = standard_rate
            if overtime_rate:
                first.OvertimeRate = overtime_rate
            if cost_per_use >= 0:
                first.CostPerUse = cost_per_use

        app.FileSave()
        return json.dumps({
            "status":   "updated",
            "resource": res.Name,
            "table":    table.upper(),
            "standard_rate": standard_rate or "(unchanged)",
            "overtime_rate": overtime_rate or "(unchanged)",
            "cost_per_use":  cost_per_use if cost_per_use >= 0 else "(unchanged)",
        }, indent=2)

    except Exception as e:
        return json.dumps({"error": f"Failed to update rate table: {e}"})


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("Starting MS Project MCP Server...")
    print("MS Project must be running with a file open before using tools.")
    mcp.run()


# ---------------------------------------------------------------------------
# REGISTRATION — add this to claude_desktop_config.json:
#
# {
#   "mcpServers": {
#     "msproject": {
#       "command": "python",
#       "args": ["/path/to/msproject/server.py"]
#     }
#   }
# }
# ---------------------------------------------------------------------------
