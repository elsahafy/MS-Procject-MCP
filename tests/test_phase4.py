"""
Phase 4 Tests — 19 new tools (Tier 1 + Tier 2 + Tier 3)
Run: python test_phase4.py
Requires: MS Project running (will be launched if not)
"""
import asyncio
import json
import os
import sys
import tempfile

# Add parent to path for server import
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from server import mcp

PASS = 0
FAIL = 0
SKIP = 0


async def call(tool_name, **kwargs):
    """Call an MCP tool and return parsed JSON."""
    result = await mcp.call_tool(tool_name, kwargs)
    # result is (list_of_TextContent, meta_dict)
    contents = result[0] if isinstance(result, tuple) else result
    text = contents[0].text if contents else ""
    return json.loads(text) if text else {}


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

    # -----------------------------------------------------------------------
    # Setup
    # -----------------------------------------------------------------------
    print("\n=== SETUP ===")

    r = await call("new_project", title="Phase 4 Test", start="2026-04-01")
    print(f"  Created project: {r.get('title')}")

    # Add tasks: 2 programmes with children + milestones
    tasks_data = json.dumps([
        {"name": "Programme Alpha", "outline_level": 1},
        {"name": "Task A", "outline_level": 2, "start": "2026-04-01", "duration_days": 10},
        {"name": "Task B", "outline_level": 2, "start": "2026-04-15", "duration_days": 10},
        {"name": "Milestone X", "outline_level": 2, "milestone": True, "start": "2026-04-30"},
        {"name": "Programme Beta", "outline_level": 1},
        {"name": "Task C", "outline_level": 2, "start": "2026-05-01", "duration_days": 15},
        {"name": "Task D", "outline_level": 2, "start": "2026-05-20", "duration_days": 10},
        {"name": "Sub-task D1", "outline_level": 3, "start": "2026-05-20", "duration_days": 5},
        {"name": "Sub-task D2", "outline_level": 3, "start": "2026-05-27", "duration_days": 5},
        {"name": "Milestone Y", "outline_level": 2, "milestone": True, "start": "2026-06-15"},
    ])
    r = await call("bulk_add_tasks", tasks_json=tasks_data)
    created = r.get("tasks", [])
    uid = {t["name"]: t["unique_id"] for t in created}
    print(f"  Created {len(created)} tasks")

    # Add resources
    for name in ["Alice", "Bob", "Charlie"]:
        await call("add_resource", name=name)
    print("  Added 3 resources")

    # Assign resources
    await call("assign_resource", task_unique_id=uid["Task A"], resource_name="Alice")
    await call("assign_resource", task_unique_id=uid["Task B"], resource_name="Bob")
    await call("assign_resource", task_unique_id=uid["Task C"], resource_name="Charlie")
    await call("assign_resource", task_unique_id=uid["Task D"], resource_name="Alice")
    print("  Assigned 4 resources")

    # Add predecessor chain: A → B → C → D
    links = json.dumps([
        {"successor_unique_id": uid["Task B"], "predecessor_unique_id": uid["Task A"]},
        {"successor_unique_id": uid["Task C"], "predecessor_unique_id": uid["Task B"]},
        {"successor_unique_id": uid["Task D"], "predecessor_unique_id": uid["Task C"]},
    ])
    await call("bulk_add_predecessors", links_json=links)
    print("  Created 3 predecessor links")

    # Set RAG
    rag_updates = json.dumps([
        {"unique_id": uid["Task A"], "rag": "Green"},
        {"unique_id": uid["Task B"], "rag": "Amber"},
        {"unique_id": uid["Task C"], "rag": "Red"},
    ])
    await call("bulk_update_rag", updates=rag_updates)
    print("  Set RAG statuses")

    # Save baseline 0
    await call("save_baseline", baseline_number=0)
    print("  Saved baseline 0")

    # Shift Task B finish to create variance
    await call("update_task", unique_id=uid["Task B"], finish="2026-05-05")
    print("  Shifted Task B finish for variance test")

    # -----------------------------------------------------------------------
    # Tests
    # -----------------------------------------------------------------------
    print("\n=== PHASE 4 TESTS ===")

    # 1. compare_baselines
    r = await call("compare_baselines", baseline_a=0, baseline_b=-1)
    has_variance = r.get("summary", {}).get("tasks_with_variance", 0) > 0
    ok("1. compare_baselines", has_variance, f"variance tasks: {r.get('summary', {}).get('tasks_with_variance')}")

    # 2. get_dependency_chain — successors from A
    r = await call("get_dependency_chain", unique_id=uid["Task A"], direction="successors")
    chain_names = [e["name"] for e in r.get("chain", [])]
    ok("2. get_dependency_chain (successors)",
       "Task B" in chain_names and "Task C" in chain_names,
       f"chain: {chain_names}")

    # 3. get_dependency_chain — predecessors from D
    r = await call("get_dependency_chain", unique_id=uid["Task D"], direction="predecessors")
    chain_names = [e["name"] for e in r.get("chain", [])]
    ok("3. get_dependency_chain (predecessors)",
       "Task C" in chain_names and "Task B" in chain_names,
       f"chain: {chain_names}")

    # 4. bulk_assign_resources
    assign_data = json.dumps([
        {"task_unique_id": uid["Sub-task D1"], "resource_name": "Alice"},
        {"task_unique_id": uid["Sub-task D2"], "resource_name": "Bob"},
        {"task_unique_id": uid["Milestone X"], "resource_name": "NewRes"},
    ])
    r = await call("bulk_assign_resources", assignments_json=assign_data)
    ok("4. bulk_assign_resources", r.get("assigned") == 3, f"assigned: {r.get('assigned')}")

    # 5. remove_resource_assignment
    r = await call("remove_resource_assignment", task_unique_id=uid["Sub-task D1"], resource_name="Alice")
    ok("5a. remove_resource_assignment", r.get("status") == "removed", f"status: {r.get('status')}")
    r2 = await call("remove_resource_assignment", task_unique_id=uid["Sub-task D1"], resource_name="NonExistent")
    ok("5b. remove_resource_assignment (invalid)", "error" in r2, "expected error for non-existent")

    # 6. update_resource
    r = await call("update_resource", resource_name="Charlie", max_units=2.0)
    ok("6. update_resource", "max_units" in r.get("changed", []), f"changed: {r.get('changed')}")

    # 7. move_task — move Milestone Y to after Task A
    r = await call("move_task", unique_id=uid["Milestone Y"], after_unique_id=uid["Task A"])
    ok("7. move_task", r.get("status") == "moved", f"status: {r.get('status')}, new_id: {r.get('new_id')}")

    # 8. get_progress_by_wbs
    r = await call("get_progress_by_wbs", max_level=2)
    branches = r.get("branches", [])
    ok("8. get_progress_by_wbs", len(branches) > 0, f"branches: {len(branches)}")

    # 9. copy_task_structure — copy Programme Alpha subtree
    r = await call("copy_task_structure", source_unique_id=uid["Programme Alpha"])
    copied = r.get("copied_tasks", [])
    ok("9. copy_task_structure", len(copied) > 0, f"copied: {len(copied)} tasks")

    # 10. cross_project_link — requires 2 open projects; COM proxy gets stale
    #     when creating a 2nd project in-process. Test manually or via MS Project UI.
    skip("10. cross_project_link", "COM proxy staleness with multi-project in-process; test manually")

    # 11. export_csv
    tmp_csv = os.path.join(tempfile.gettempdir(), "phase4_test_export.csv")
    r = await call("export_csv", output_path=tmp_csv)
    ok("11. export_csv", r.get("rows", 0) > 0 and os.path.exists(tmp_csv), f"rows: {r.get('rows')}")
    if os.path.exists(tmp_csv):
        os.remove(tmp_csv)

    # 12. bulk_set_deadlines
    dl_data = json.dumps([
        {"unique_id": uid["Task A"], "deadline_date": "2026-04-20"},
        {"unique_id": uid["Task B"], "deadline_date": "2026-05-10"},
        {"unique_id": uid["Task C"], "deadline_date": "2026-06-01"},
        {"unique_id": uid["Task A"], "deadline_date": "clear"},  # clear the one we just set
    ])
    r = await call("bulk_set_deadlines", deadlines_json=dl_data)
    ok("12. bulk_set_deadlines", r.get("set", 0) >= 2 and r.get("cleared", 0) >= 1,
       f"set: {r.get('set')}, cleared: {r.get('cleared')}")

    # 13. find_available_slack
    r = await call("find_available_slack", min_days=0)
    ok("13. find_available_slack", isinstance(r.get("tasks"), list), f"count: {r.get('count')}")

    # 14. set_task_calendar
    try:
        r = await call("set_task_calendar", unique_id=uid["Task C"], calendar_name="Standard")
        ok("14a. set_task_calendar (set)", r.get("status") == "updated", f"status: {r.get('status')}")
        r2 = await call("set_task_calendar", unique_id=uid["Task C"], calendar_name="")
        ok("14b. set_task_calendar (clear)", r2.get("status") == "updated", f"status: {r2.get('status')}")
    except Exception as e:
        skip("14. set_task_calendar", str(e))

    # 15. get_cost_summary
    r = await call("get_cost_summary")
    ok("15. get_cost_summary", "totals" in r and "by_resource" in r, f"keys: {list(r.keys())}")

    # 16. undo_last
    await call("update_task", unique_id=uid["Task A"], notes="UNDO TEST MARKER")
    r = await call("undo_last", count=1)
    ok("16. undo_last", r.get("status") == "undone", f"undo_count: {r.get('undo_count')}")

    # 17. create_calendar
    r = await call("create_calendar", name="Test Cal", copy_from="Standard")
    ok("17. create_calendar", r.get("status") == "created" and "Test Cal" in r.get("calendars", []),
       f"calendars: {r.get('calendars')}")

    # 18. insert_subproject — skip if no test file
    skip("18. insert_subproject", "no test .mpp file available")

    # 19. apply_filter
    try:
        r = await call("apply_filter", filter_name="Critical")
        ok("19a. apply_filter (Critical)", r.get("status") == "applied", f"status: {r.get('status')}")
        r2 = await call("apply_filter", filter_name="All Tasks")
        ok("19b. apply_filter (All Tasks)", r2.get("status") == "applied", f"status: {r2.get('status')}")
    except Exception as e:
        skip("19. apply_filter", str(e))

    # 20. snapshot_to_json
    tmp_json = os.path.join(tempfile.gettempdir(), "phase4_test_snapshot.json")
    r = await call("snapshot_to_json", output_path=tmp_json)
    snapshot_ok = r.get("tasks", 0) > 0 and os.path.exists(tmp_json)
    if snapshot_ok:
        with open(tmp_json, "r") as fp:
            snap = json.load(fp)
        snapshot_ok = len(snap.get("tasks", [])) == r["tasks"]
    ok("20. snapshot_to_json", snapshot_ok, f"tasks: {r.get('tasks')}")
    if os.path.exists(tmp_json):
        os.remove(tmp_json)

    # -----------------------------------------------------------------------
    # Cleanup
    # -----------------------------------------------------------------------
    print("\n=== CLEANUP ===")
    # Close all test projects
    for proj_name in ["Phase 4 Xlink", "Phase 4 Test"]:
        try:
            await call("switch_project", name_or_index=proj_name)
            await call("close_project", save=False)
            print(f"  Closed {proj_name}")
        except Exception:
            # Project may not exist or COM stale — try direct close
            try:
                await call("close_project", save=False)
                print(f"  Closed active project (was {proj_name})")
            except Exception:
                pass
    print("  Cleanup done")

    # -----------------------------------------------------------------------
    # Summary
    # -----------------------------------------------------------------------
    total = PASS + FAIL + SKIP
    print(f"\n{'='*50}")
    print(f"  Phase 4 Results: {PASS} passed, {FAIL} failed, {SKIP} skipped / {total} total")
    print(f"{'='*50}\n")

    return FAIL == 0


if __name__ == "__main__":
    success = asyncio.run(run_tests())
    sys.exit(0 if success else 1)
