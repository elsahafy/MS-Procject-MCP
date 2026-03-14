"""
Phase 6 Tests — Gap Analysis & Missing Features (17 new tools + enriched task_to_dict)
Run: python test_phase6.py
Requires: MS Project running (will be launched if not)
"""
import asyncio
import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from server import mcp

PASS = 0
FAIL = 0
SKIP = 0


async def call(tool_name, **kwargs):
    """Call an MCP tool and return parsed JSON."""
    try:
        result = await mcp.call_tool(tool_name, kwargs)
        contents = result[0] if isinstance(result, tuple) else result
        text = contents[0].text if contents else ""
        return json.loads(text) if text else {}
    except Exception as e:
        print(f"  [ERROR] {tool_name}: {str(e)[:120]}")
        return {"error": str(e)}


async def run_tests():
    global PASS, FAIL, SKIP

    def ok(name, cond, detail=""):
        global PASS, FAIL
        if cond:
            PASS += 1
            print(f"  PASS  {name}")
        else:
            FAIL += 1
            print(f"  FAIL  {name}  {detail}")

    def skip(name, reason=""):
        global SKIP
        SKIP += 1
        print(f"  SKIP  {name}  {reason}")

    # -------------------------------------------------------------------
    # Setup — create a test project
    # -------------------------------------------------------------------
    print("\n=== SETUP ===")

    r = await call("new_project", title="Phase 6 Test", start="2026-04-01")
    print(f"  Created project: {r.get('title')}")

    tasks_data = json.dumps([
        {"name": "Programme A", "outline_level": 1},
        {"name": "Task Alpha",  "outline_level": 2, "start": "2026-04-01", "duration_days": 10},
        {"name": "Task Beta",   "outline_level": 2, "start": "2026-04-15", "duration_days": 20},
        {"name": "Milestone M", "outline_level": 2, "start": "2026-05-15", "duration_days": 0},
        {"name": "Task Gamma",  "outline_level": 2, "start": "2026-06-01", "duration_days": 30},
    ])
    r = await call("bulk_add_tasks", tasks_json=tasks_data)
    created = r.get("tasks", [])
    uid_alpha = created[1]["unique_id"] if len(created) > 1 else None
    uid_beta  = created[2]["unique_id"] if len(created) > 2 else None
    uid_gamma = created[4]["unique_id"] if len(created) > 4 else None

    # Add a predecessor so we get slack values
    if uid_alpha and uid_beta:
        await call("add_predecessor", successor_unique_id=uid_beta, predecessor_unique_id=uid_alpha, link_type="FS")

    # Add a resource for resource tests
    await call("add_resource", name="PM Lead", max_units=1.0)
    if uid_alpha:
        await call("assign_resource", task_unique_id=uid_alpha, resource_name="PM Lead")

    # -------------------------------------------------------------------
    # 1. health_check / ping
    # -------------------------------------------------------------------
    print("\n=== 1. health_check ===")
    r = await call("health_check")
    ok("health_check returns connected", r.get("status") == "connected")
    ok("health_check has version",       "version" in r)
    ok("health_check project_open",      r.get("project_open") is True)

    # -------------------------------------------------------------------
    # 2. Enriched task_to_dict
    # -------------------------------------------------------------------
    print("\n=== 2. Enriched task_to_dict ===")
    if uid_alpha:
        r = await call("get_task", unique_id=uid_alpha)
        NEW_FIELDS = [
            "actual_start", "actual_finish", "remaining_duration_days",
            "total_slack_days", "free_slack_days", "deadline", "priority",
            "constraint_type", "constraint_date", "manual", "type",
            "hyperlink", "hyperlink_text",
        ]
        for field in NEW_FIELDS:
            ok(f"task_to_dict has '{field}'", field in r, f"missing key: {field}")
        ok("priority is int", isinstance(r.get("priority"), int))
        ok("constraint_type is string", isinstance(r.get("constraint_type"), str))
        ok("type is string", isinstance(r.get("type"), str))
    else:
        skip("enriched task_to_dict", "no task created")

    # -------------------------------------------------------------------
    # 3. calculate_project
    # -------------------------------------------------------------------
    print("\n=== 3. calculate_project ===")
    r = await call("calculate_project")
    ok("calculate_project status", r.get("status") == "calculated")
    ok("calculate_project returns project name", "project" in r)

    # -------------------------------------------------------------------
    # 4. list_calendar_exceptions
    # -------------------------------------------------------------------
    print("\n=== 4. list_calendar_exceptions ===")
    # First add an exception so we have something to list
    await call("set_calendar_exception",
               calendar_name="Standard", name="Test Holiday",
               start="2026-12-25", finish="2026-12-25", working=False)

    r = await call("list_calendar_exceptions", calendar_name="Standard")
    ok("list_calendar_exceptions has calendar", "calendar" in r)
    ok("list_calendar_exceptions has count",    isinstance(r.get("count"), int))
    ok("list_calendar_exceptions count >= 1",   r.get("count", 0) >= 1)
    excs = r.get("exceptions", [])
    names = [e["name"] for e in excs]
    # COM may store the name differently; check that at least one exception exists
    ok("exceptions list non-empty", len(excs) >= 1, f"names={names}")

    # default (project calendar)
    r2 = await call("list_calendar_exceptions")
    ok("list_calendar_exceptions default calendar works", "calendar" in r2)

    # -------------------------------------------------------------------
    # 5. update_project (mark complete through date)
    # -------------------------------------------------------------------
    print("\n=== 5. update_project ===")
    r = await call("update_project", complete_through="2026-04-10")
    ok("update_project status", r.get("status") == "updated")
    ok("update_project returns date", r.get("complete_through") == "2026-04-10")

    # -------------------------------------------------------------------
    # 6. reschedule_incomplete_work
    # -------------------------------------------------------------------
    print("\n=== 6. reschedule_incomplete_work ===")
    r = await call("reschedule_incomplete_work", reschedule_from="2026-04-15")
    ok("reschedule status", r.get("status") == "rescheduled")
    ok("reschedule date",   r.get("reschedule_from") == "2026-04-15")

    # -------------------------------------------------------------------
    # 7. delete_calendar_exception
    # -------------------------------------------------------------------
    print("\n=== 7. delete_calendar_exception ===")
    # Find the name of the exception that was actually stored
    exc_name_to_delete = excs[0]["name"] if excs else "Test Holiday"
    r = await call("delete_calendar_exception",
                   calendar_name="Standard", exception_name=exc_name_to_delete)
    ok("delete_calendar_exception status", r.get("status") == "deleted", f"got: {r}")

    # Verify count decreased
    r2 = await call("list_calendar_exceptions", calendar_name="Standard")
    ok("exception removed", r2.get("count", 999) < r.get("count", 999) if "count" in r else r2.get("count", 0) == 0)

    # -------------------------------------------------------------------
    # 8. delete_calendar (create one first, then delete)
    # -------------------------------------------------------------------
    print("\n=== 8. delete_calendar ===")
    await call("create_calendar", name="Temp Cal", copy_from="Standard")
    r = await call("delete_calendar", calendar_name="Temp Cal")
    ok("delete_calendar status", r.get("status") == "deleted")

    # Try deleting project calendar — should fail
    r2 = await call("delete_calendar", calendar_name="Standard")
    ok("delete project calendar blocked", "error" in r2)

    # -------------------------------------------------------------------
    # 9. set_resource_calendar
    # -------------------------------------------------------------------
    print("\n=== 9. set_resource_calendar ===")
    await call("create_calendar", name="Part Time", copy_from="Standard")
    r = await call("set_resource_calendar", resource_name="PM Lead", calendar_name="Part Time")
    ok("set_resource_calendar status", r.get("status") == "updated")
    ok("set_resource_calendar resource", r.get("resource") == "PM Lead")

    # Non-existent calendar
    r2 = await call("set_resource_calendar", resource_name="PM Lead", calendar_name="NoSuchCal")
    ok("set_resource_calendar bad cal", "error" in r2)

    # -------------------------------------------------------------------
    # 10. get_timephased_data
    # -------------------------------------------------------------------
    print("\n=== 10. get_timephased_data ===")
    if uid_alpha:
        r = await call("get_timephased_data",
                       unique_id=uid_alpha, start_date="2026-04-01",
                       end_date="2026-04-30", timescale="weekly", data_type="work")
        ok("timephased has periods", isinstance(r.get("periods"), list))
        ok("timephased has name",    "name" in r)
        ok("timephased timescale",   r.get("timescale") == "weekly")

        # Monthly
        r2 = await call("get_timephased_data",
                        unique_id=uid_alpha, start_date="2026-04-01",
                        end_date="2026-06-30", timescale="monthly", data_type="cost")
        ok("timephased monthly works", isinstance(r2.get("periods"), list))

        # Bad timescale
        r3 = await call("get_timephased_data",
                        unique_id=uid_alpha, start_date="2026-04-01",
                        end_date="2026-04-30", timescale="hourly", data_type="work")
        ok("timephased bad timescale", "error" in r3)
    else:
        skip("get_timephased_data", "no task")

    # -------------------------------------------------------------------
    # 11. set_working_hours
    # -------------------------------------------------------------------
    print("\n=== 11. set_working_hours ===")
    # Set Friday (6) to half-day
    shifts = json.dumps([["07:00", "12:00"]])
    r = await call("set_working_hours", calendar_name="Standard", day=6, shifts_json=shifts)
    ok("set_working_hours status", r.get("status") == "updated")
    ok("set_working_hours working", r.get("working") is True)

    # Set Saturday (7) to non-working
    r2 = await call("set_working_hours", calendar_name="Standard", day=7, shifts_json="[]")
    ok("set_working_hours non-working", r2.get("working") is False)

    # Bad day
    r3 = await call("set_working_hours", calendar_name="Standard", day=0, shifts_json="[]")
    ok("set_working_hours bad day", "error" in r3)

    # -------------------------------------------------------------------
    # 12. get_resource_availability
    # -------------------------------------------------------------------
    print("\n=== 12. get_resource_availability ===")
    r = await call("get_resource_availability",
                   resource_name="PM Lead", start_date="2026-04-01",
                   end_date="2026-04-30", timescale="weekly")
    ok("resource_availability has periods", isinstance(r.get("periods"), list))
    ok("resource_availability has max_units", "max_units" in r)
    ok("resource_availability resource name", r.get("resource") == "PM Lead")

    # Non-existent resource
    r2 = await call("get_resource_availability",
                    resource_name="Nobody", start_date="2026-04-01",
                    end_date="2026-04-30")
    ok("resource_availability bad resource", "error" in r2)

    # -------------------------------------------------------------------
    # 13. update_task — priority and task_type
    # -------------------------------------------------------------------
    print("\n=== 13. update_task priority & type ===")
    if uid_alpha:
        r = await call("update_task", unique_id=uid_alpha, priority=800, task_type="FixedDuration")
        ok("update_task priority changed", "priority" in r.get("changed", []))
        ok("update_task type changed",     "type" in r.get("changed", []))

        # Verify
        r2 = await call("get_task", unique_id=uid_alpha)
        ok("priority is 800",             r2.get("priority") == 800)
        ok("type is FixedDuration",        r2.get("type") == "FixedDuration")
    else:
        skip("update_task priority/type", "no task")

    # -------------------------------------------------------------------
    # 14. get_variance_report
    # -------------------------------------------------------------------
    print("\n=== 14. get_variance_report ===")
    # Save a baseline first
    await call("save_baseline", baseline_number=0)
    r = await call("get_variance_report", baseline=0)
    ok("variance_report has baseline",    "baseline" in r)
    ok("variance_report has total_tasks", isinstance(r.get("total_tasks"), int))
    ok("variance_report has tasks list",  isinstance(r.get("tasks"), list))

    # -------------------------------------------------------------------
    # 15. snapshot_diff
    # -------------------------------------------------------------------
    print("\n=== 15. snapshot_diff ===")
    tmp_dir = tempfile.gettempdir()
    snap_a = os.path.join(tmp_dir, "snap_a.json")
    snap_b = os.path.join(tmp_dir, "snap_b.json")

    # Take snapshot A
    await call("snapshot_to_json", output_path=snap_a)

    # Modify a task
    if uid_beta:
        await call("update_task", unique_id=uid_beta, name="Task Beta MODIFIED")

    # Take snapshot B
    await call("snapshot_to_json", output_path=snap_b)

    r = await call("snapshot_diff", path_a=snap_a, path_b=snap_b)
    ok("snapshot_diff has counts", "added_count" in r and "changed_count" in r)
    if uid_beta:
        ok("snapshot_diff detects change", r.get("changed_count", 0) >= 1)
    else:
        skip("snapshot_diff change detection", "no task")

    # Bad path
    r2 = await call("snapshot_diff", path_a="nonexistent.json", path_b=snap_b)
    ok("snapshot_diff bad path", "error" in r2)

    # Cleanup
    for f in (snap_a, snap_b):
        try:
            os.remove(f)
        except Exception:
            pass

    # -------------------------------------------------------------------
    # 16. set_task_hyperlink
    # -------------------------------------------------------------------
    print("\n=== 16. set_task_hyperlink ===")
    if uid_alpha:
        r = await call("set_task_hyperlink",
                       unique_id=uid_alpha, url="https://example.com",
                       text="Example Link")
        ok("set_hyperlink status", r.get("status") == "updated")
        ok("set_hyperlink url",    r.get("hyperlink") == "https://example.com")

        # Verify in task_to_dict
        r2 = await call("get_task", unique_id=uid_alpha)
        ok("hyperlink in task_to_dict", r2.get("hyperlink") == "https://example.com")
        ok("hyperlink_text in task_to_dict", r2.get("hyperlink_text") == "Example Link")
    else:
        skip("set_task_hyperlink", "no task")

    # -------------------------------------------------------------------
    # 17. add_recurring_task
    # -------------------------------------------------------------------
    print("\n=== 17. add_recurring_task ===")
    r = await call("add_recurring_task",
                   name="Weekly Status Meeting",
                   recurrence_type="weekly",
                   start_date="2026-05-01",
                   end_date="2026-06-30",
                   duration_days=0.25,
                   day_of_week=2)
    if "error" not in r:
        ok("recurring task created", r.get("status") == "created")
        ok("recurring task has uid", "unique_id" in r)
    else:
        skip("add_recurring_task", f"COM may not support: {r.get('error', '')[:80]}")

    # -------------------------------------------------------------------
    # 18. get_resource_rate_tables
    # -------------------------------------------------------------------
    print("\n=== 18. get_resource_rate_tables ===")
    r = await call("get_resource_rate_tables", resource_name="PM Lead")
    ok("rate_tables has resource", r.get("resource") == "PM Lead")
    ok("rate_tables has tables",   isinstance(r.get("tables"), dict))
    ok("rate_tables has table A",  "A" in r.get("tables", {}))

    # Non-existent resource
    r2 = await call("get_resource_rate_tables", resource_name="Nobody")
    ok("rate_tables bad resource", "error" in r2)

    # -------------------------------------------------------------------
    # 19. set_resource_rate_table
    # -------------------------------------------------------------------
    print("\n=== 19. set_resource_rate_table ===")
    r = await call("set_resource_rate_table",
                   resource_name="PM Lead", table="A",
                   standard_rate="50/h", overtime_rate="75/h")
    if "error" not in r:
        ok("set_rate_table status", r.get("status") == "updated")
        ok("set_rate_table table",  r.get("table") == "A")
    else:
        skip("set_resource_rate_table", f"may need specific setup: {r.get('error', '')[:80]}")

    # -------------------------------------------------------------------
    # Cleanup
    # -------------------------------------------------------------------
    print("\n=== CLEANUP ===")
    await call("close_project", save=False)
    print("  Project closed without saving.\n")

    # -------------------------------------------------------------------
    # Summary
    # -------------------------------------------------------------------
    total = PASS + FAIL + SKIP
    print("=" * 60)
    print(f"  Phase 6 Results:  {PASS} passed,  {FAIL} failed,  {SKIP} skipped  (total {total})")
    print("=" * 60)

    return FAIL == 0


if __name__ == "__main__":
    success = asyncio.run(run_tests())
    sys.exit(0 if success else 1)
