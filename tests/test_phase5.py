"""
Phase 5 Tests — 2 bug fixes + 3 new tools
Run: python test_phase5.py
Requires: MS Project running (will be launched if not)
"""
import asyncio
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from server import mcp

PASS = 0
FAIL = 0
SKIP = 0


async def call(tool_name, **kwargs):
    """Call an MCP tool and return parsed JSON."""
    result = await mcp.call_tool(tool_name, kwargs)
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

    # -------------------------------------------------------------------
    # Setup — create a project with tasks + constraints + resources
    # -------------------------------------------------------------------
    print("\n=== SETUP ===")

    r = await call("new_project", title="Phase 5 Test", start="2026-04-01")
    print(f"  Created project: {r.get('title')}")

    # Add tasks: some overdue (finish in the past), some not
    tasks_data = json.dumps([
        {"name": "Programme A", "outline_level": 1},
        {"name": "Past Task",   "outline_level": 2, "start": "2025-01-01", "duration_days": 10},
        {"name": "Future Task", "outline_level": 2, "start": "2027-06-01", "duration_days": 20},
        {"name": "Work Task",   "outline_level": 2, "start": "2026-05-01", "duration_days": 15},
    ])
    r = await call("bulk_add_tasks", tasks_json=tasks_data)
    created = r.get("tasks", [])
    uid = {t["name"]: t["unique_id"] for t in created}
    print(f"  Created {len(created)} tasks")

    # Add a resource and assign it
    await call("add_resource", name="Alice", type=0)
    if "Work Task" in uid:
        await call("assign_resource", task_unique_id=uid["Work Task"], resource_name="Alice")
    await call("add_resource", name="ToDelete", type=0)

    # Set a constraint on future task
    if "Future Task" in uid:
        await call("set_constraint",
                   unique_id=uid["Future Task"],
                   constraint_type="SNET",
                   constraint_date="2027-06-01")

    # -------------------------------------------------------------------
    # Test 1: get_overdue_tasks uses _to_naive (bug fix validation)
    # -------------------------------------------------------------------
    print("\n=== TEST 1: get_overdue_tasks _to_naive fix ===")

    r = await call("get_overdue_tasks")
    # Should not crash; past_task should be in overdue list
    overdue_ids = [t["unique_id"] for t in r.get("tasks", [])]
    ok("get_overdue_tasks returns without error", "error" not in r)
    if "Past Task" in uid:
        ok("past task is overdue", uid["Past Task"] in overdue_ids,
           f"expected uid {uid['Past Task']} in {overdue_ids}")
    else:
        skip("past task is overdue", "Past Task not created")

    # -------------------------------------------------------------------
    # Test 2: get_progress_summary uses _to_naive (bug fix validation)
    # -------------------------------------------------------------------
    print("\n=== TEST 2: get_progress_summary _to_naive fix ===")

    r = await call("get_progress_summary")
    ok("get_progress_summary returns without error", "error" not in r)
    ok("overdue count is populated", r.get("overdue", -1) >= 0,
       f"got overdue={r.get('overdue')}")

    # -------------------------------------------------------------------
    # Test 3: get_constraints
    # -------------------------------------------------------------------
    print("\n=== TEST 3: get_constraints ===")

    r = await call("get_constraints")
    ok("get_constraints returns count", "count" in r, str(r))
    constrained = r.get("tasks", [])
    snet_found = any(c["constraint_type"] == "SNET" for c in constrained)
    ok("SNET constraint found", snet_found, f"constraints: {constrained}")

    # -------------------------------------------------------------------
    # Test 4: delete_resource
    # -------------------------------------------------------------------
    print("\n=== TEST 4: delete_resource ===")

    r = await call("delete_resource", resource_name="ToDelete")
    ok("delete_resource status=deleted", r.get("status") == "deleted", str(r))

    # Verify it's gone
    r2 = await call("delete_resource", resource_name="ToDelete")
    ok("re-delete returns error (not found)", "error" in r2, str(r2))

    # -------------------------------------------------------------------
    # Test 5: get_actual_work
    # -------------------------------------------------------------------
    print("\n=== TEST 5: get_actual_work ===")

    r = await call("get_actual_work")
    ok("get_actual_work returns totals", "totals" in r, str(r))
    ok("totals has work_hours", "work_hours" in r.get("totals", {}), str(r))
    ok("tasks list returned", r.get("count", 0) > 0, f"count={r.get('count')}")

    # -------------------------------------------------------------------
    # Cleanup
    # -------------------------------------------------------------------
    print("\n=== CLEANUP ===")
    try:
        await call("close_project")
    except Exception:
        pass
    print("  Project closed.")

    # -------------------------------------------------------------------
    # Summary
    # -------------------------------------------------------------------
    print(f"\n{'='*50}")
    print(f"  Phase 5 Results:  PASS={PASS}  FAIL={FAIL}  SKIP={SKIP}")
    print(f"{'='*50}")
    return FAIL == 0


if __name__ == "__main__":
    success = asyncio.run(run_tests())
    sys.exit(0 if success else 1)
