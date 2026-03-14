"""Integration test for Phase 2 MCP server tools (10 new features)."""
import asyncio
import json
import importlib.util

spec = importlib.util.spec_from_file_location(
    "server", r"C:\Users\Ibrahim Elsahafy\mcp-servers\msproject\server.py"
)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)


async def call(name, args=None):
    """Call an MCP tool and return parsed JSON or raw text."""
    r = await mod.mcp.call_tool(name, args or {})
    if isinstance(r, tuple):
        r = r[0]
    if isinstance(r, list):
        item = r[0]
        text = item.text if hasattr(item, "text") else str(item)
    elif hasattr(r, "text"):
        text = r.text
    else:
        text = str(r)
    try:
        return json.loads(text)
    except (json.JSONDecodeError, TypeError):
        return text


async def test():
    results = {}
    task_uids = {}

    # Setup: create test project
    print("=== Setup: new_project ===")
    r = await call("new_project", {"title": "Phase 2 Test", "start": "2026-04-01"})
    print(f"  Project: {r['name']}")

    # Create tasks for testing
    print("=== Setup: bulk_add_tasks ===")
    tasks = [
        {"name": "Programme A", "outline_level": 1},
        {"name": "Design Phase", "outline_level": 2, "start": "2026-04-01", "finish": "2026-05-01", "duration_days": 22},
        {"name": "Build Phase", "outline_level": 2, "start": "2026-05-02", "finish": "2026-06-30", "duration_days": 43},
        {"name": "Design Review", "outline_level": 3, "start": "2026-04-15", "finish": "2026-04-20", "duration_days": 5},
        {"name": "Go-Live", "outline_level": 2, "milestone": True, "start": "2026-06-30"},
    ]
    r = await call("bulk_add_tasks", {"tasks_json": json.dumps(tasks)})
    print(f"  Created: {r['created']} tasks")
    task_uids = {t["name"]: t["unique_id"] for t in r["tasks"]}

    # Add predecessor links
    links = [
        {"successor_unique_id": task_uids["Build Phase"], "predecessor_unique_id": task_uids["Design Phase"], "link_type": "FS"},
        {"successor_unique_id": task_uids["Go-Live"], "predecessor_unique_id": task_uids["Build Phase"], "link_type": "FS"},
    ]
    await call("bulk_add_predecessors", {"links_json": json.dumps(links)})
    print("  Links: 2 added")

    # ---------------------------------------------------------------
    # Test 1: set_project_properties
    # ---------------------------------------------------------------
    print("\n=== Test 1: set_project_properties ===")
    try:
        r = await call("set_project_properties", {"properties_json": json.dumps({
            "title": "EXPO 2030 Test Schedule",
            "manager": "Ibrahim Elsahafy",
            "company": "ERC",
            "author": "PMO",
            "subject": "Integration Test",
        })})
        print(f"  Changed: {r['changed']}")

        # Verify via get_project_info
        info = await call("get_project_info")
        assert info["title"] == "EXPO 2030 Test Schedule", f"title mismatch: {info['title']}"
        assert info["manager"] == "Ibrahim Elsahafy", f"manager mismatch: {info['manager']}"
        assert info["company"] == "ERC", f"company mismatch: {info['company']}"
        print(f"  Verified: title={info['title']}, manager={info['manager']}, company={info['company']}")
        results["set_project_properties"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["set_project_properties"] = "FAIL"

    # ---------------------------------------------------------------
    # Test 2: add_resource
    # ---------------------------------------------------------------
    print("\n=== Test 2: add_resource ===")
    try:
        r1 = await call("add_resource", {"name": "Senior PM", "type": 0, "max_units": 1.0})
        r2 = await call("add_resource", {"name": "Concrete", "type": 1, "max_units": 100.0})
        print(f"  Created: {r1['name']} (UID={r1['unique_id']}), {r2['name']} (UID={r2['unique_id']})")
        assert r1["type"] == 0, "Work resource type wrong"
        assert r2["type"] == 1, "Material resource type wrong"
        results["add_resource"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["add_resource"] = "FAIL"

    # ---------------------------------------------------------------
    # Test 3: assign_resource
    # ---------------------------------------------------------------
    print("\n=== Test 3: assign_resource ===")
    try:
        r = await call("assign_resource", {
            "task_unique_id": task_uids["Design Phase"],
            "resource_name": "Senior PM",
        })
        print(f"  Assigned: {r['resource_name']} -> {r['task_name']}")
        assert "Senior PM" in r["resource_names"], "Resource not in names"

        # Assign second resource to same task
        r2 = await call("assign_resource", {
            "task_unique_id": task_uids["Design Phase"],
            "resource_name": "Concrete",
        })
        print(f"  Assigned: {r2['resource_name']} -> {r2['task_name']} (now: {r2['resource_names']})")

        # Assign to different task (resource auto-created)
        r3 = await call("assign_resource", {
            "task_unique_id": task_uids["Build Phase"],
            "resource_name": "Contractor A",
        })
        print(f"  Assigned: {r3['resource_name']} -> {r3['task_name']} (auto-created resource)")
        results["assign_resource"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["assign_resource"] = "FAIL"

    # ---------------------------------------------------------------
    # Test 4: indent_task
    # ---------------------------------------------------------------
    print("\n=== Test 4: indent_task ===")
    try:
        # Get current level of Design Review (should be 3)
        r = await call("indent_task", {
            "unique_id": task_uids["Design Review"],
            "direction": "outdent",
        })
        print(f"  Outdented: {r['name']} level {r['old_level']} -> {r['new_level']}")
        assert r["new_level"] < r["old_level"], "Outdent didn't decrease level"

        # Indent it back
        r2 = await call("indent_task", {
            "unique_id": task_uids["Design Review"],
            "direction": "indent",
        })
        print(f"  Indented back: {r2['name']} level {r2['old_level']} -> {r2['new_level']}")
        results["indent_task"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["indent_task"] = "FAIL"

    # ---------------------------------------------------------------
    # Test 5: get_wbs_structure
    # ---------------------------------------------------------------
    print("\n=== Test 5: get_wbs_structure ===")
    try:
        r = await call("get_wbs_structure")
        print(f"  Root: {r['name']}, children: {len(r['children'])}")
        # Walk tree
        def count_nodes(node):
            c = len(node.get("children", []))
            for child in node.get("children", []):
                c += count_nodes(child)
            return c
        total = count_nodes(r)
        print(f"  Total nodes in tree: {total}")
        assert total >= 5, f"Expected >= 5 nodes, got {total}"

        # Test with max_level filter
        r2 = await call("get_wbs_structure", {"max_level": 2})
        total2 = count_nodes(r2)
        print(f"  Nodes at max_level=2: {total2}")
        assert total2 <= total, "max_level filter should reduce nodes"
        results["get_wbs_structure"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["get_wbs_structure"] = "FAIL"

    # ---------------------------------------------------------------
    # Test 6: get_calendars
    # ---------------------------------------------------------------
    print("\n=== Test 6: get_calendars ===")
    try:
        r = await call("get_calendars")
        print(f"  Active: {r['active_calendar']}")
        print(f"  Available: {r['calendars']}")
        assert len(r["calendars"]) > 0, "No calendars found"
        results["get_calendars"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["get_calendars"] = "FAIL"

    # ---------------------------------------------------------------
    # Test 7: get_schedule_analysis
    # ---------------------------------------------------------------
    print("\n=== Test 7: get_schedule_analysis ===")
    try:
        r = await call("get_schedule_analysis")
        s = r["summary"]
        print(f"  Tasks: {s['total_tasks']}, Zero-float: {s['zero_float']}, Avg slack: {s['avg_total_slack']}d")
        assert s["total_tasks"] > 0, "No tasks analyzed"
        # Check per-task data
        for t in r["tasks"][:2]:
            print(f"    {t['name']}: total_slack={t['total_slack_days']}d, free_slack={t['free_slack_days']}d")
        results["get_schedule_analysis"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["get_schedule_analysis"] = "FAIL"

    # ---------------------------------------------------------------
    # Test 8: save_baseline
    # ---------------------------------------------------------------
    print("\n=== Test 8: save_baseline ===")
    try:
        r = await call("save_baseline", {"baseline_number": 0, "all_tasks": True})
        print(f"  Baseline {r['baseline_number']} saved for {r['tasks_baselined']} tasks")
        results["save_baseline"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["save_baseline"] = "FAIL"

    # ---------------------------------------------------------------
    # Test 9: get_earned_value
    # ---------------------------------------------------------------
    print("\n=== Test 9: get_earned_value ===")
    try:
        r = await call("get_earned_value")
        pt = r["project_totals"]
        print(f"  BCWS={pt['bcws']}, BCWP={pt['bcwp']}, ACWP={pt['acwp']}")
        print(f"  SPI={pt['spi']}, CPI={pt['cpi']}")
        if "warning" in r:
            print(f"  Warning: {r['warning']}")
        results["get_earned_value"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["get_earned_value"] = "FAIL"

    # ---------------------------------------------------------------
    # Test 10: clear_baseline
    # ---------------------------------------------------------------
    print("\n=== Test 10: clear_baseline ===")
    try:
        r = await call("clear_baseline", {"baseline_number": 0})
        print(f"  Baseline {r['baseline_number']} {r['status']}")
        results["clear_baseline"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["clear_baseline"] = "FAIL"

    # ---------------------------------------------------------------
    # Summary
    # ---------------------------------------------------------------
    print("\n" + "=" * 50)
    print("PHASE 2 TEST SUMMARY")
    print("=" * 50)
    passed = failed = 0
    for name, status in results.items():
        print(f"  [{status:4s}] {name}")
        if status == "PASS":
            passed += 1
        else:
            failed += 1
    print(f"\n  {passed} passed, {failed} failed / {len(results)} total")

    # Cleanup
    print("\nCleaning up...")
    try:
        await call("close_project", {"save": False})
        print("Test project closed.")
    except Exception:
        pass


if __name__ == "__main__":
    asyncio.run(test())
