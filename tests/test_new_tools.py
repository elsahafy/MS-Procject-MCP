"""Integration test for new MCP server tools."""
import asyncio
import json
import importlib.util
import os
import traceback

_server_path = os.path.join(os.path.dirname(__file__), "..", "server.py")
spec = importlib.util.spec_from_file_location("server", _server_path)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)


async def call(name, args=None):
    """Call an MCP tool and return parsed JSON or raw text."""
    r = await mod.mcp.call_tool(name, args or {})
    # call_tool may return (list, dict) tuple or just list
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

    # Test 1: new_project
    print("=== Test 1: new_project ===")
    try:
        r = await call("new_project", {"title": "MCP Test Project", "start": "2026-04-01"})
        print(json.dumps(r, indent=2))
        results["new_project"] = "PASS"
    except Exception as e:
        print(f"FAIL: {e}")
        results["new_project"] = "FAIL"

    # Test 2: bulk_add_tasks (5 tasks with hierarchy)
    print("\n=== Test 2: bulk_add_tasks ===")
    try:
        tasks = [
            {"name": "Phase 1 - Planning", "outline_level": 1},
            {"name": "Requirements Gathering", "outline_level": 2, "start": "2026-04-01", "finish": "2026-04-15", "duration_days": 10},
            {"name": "Stakeholder Review", "outline_level": 2, "start": "2026-04-16", "finish": "2026-04-20", "duration_days": 5},
            {"name": "Approval Gate", "outline_level": 2, "milestone": True, "start": "2026-04-20"},
            {"name": "Phase 2 - Execution", "outline_level": 1},
        ]
        r = await call("bulk_add_tasks", {"tasks_json": json.dumps(tasks)})
        print(f"Created: {r['created']} tasks")
        for t in r["tasks"]:
            print(f"  UID={t['unique_id']} ID={t['id']} {t['name']}")
        task_uids = {t["name"]: t["unique_id"] for t in r["tasks"]}
        results["bulk_add_tasks"] = "PASS"
    except Exception as e:
        print(f"FAIL: {e}")
        traceback.print_exc()
        results["bulk_add_tasks"] = "FAIL"

    # Test 3: get_project_info (MinutesPerDay bug fix)
    print("\n=== Test 3: get_project_info ===")
    try:
        r = await call("get_project_info")
        print(f"Project: {r['name']}, Tasks: {r['tasks_total']}, MPD: {r['minutes_per_day']}")
        results["get_project_info"] = "PASS"
    except Exception as e:
        print(f"FAIL: {e}")
        results["get_project_info"] = "FAIL"

    # Test 4: get_tasks (verifies task_to_dict with _get_mpd)
    print("\n=== Test 4: get_tasks ===")
    try:
        r = await call("get_tasks", {"include_summary": True})
        print(f"Tasks returned: {r['count']}")
        results["get_tasks"] = "PASS"
    except Exception as e:
        print(f"FAIL: {e}")
        results["get_tasks"] = "FAIL"

    # Test 5: bulk_add_predecessors
    print("\n=== Test 5: bulk_add_predecessors ===")
    if task_uids:
        try:
            links = [
                {"successor_unique_id": task_uids["Stakeholder Review"], "predecessor_unique_id": task_uids["Requirements Gathering"], "link_type": "FS"},
                {"successor_unique_id": task_uids["Approval Gate"], "predecessor_unique_id": task_uids["Stakeholder Review"], "link_type": "FS"},
            ]
            r = await call("bulk_add_predecessors", {"links_json": json.dumps(links)})
            print(f"Linked: {r['linked']}, Errors: {len(r['errors'])}")
            results["bulk_add_predecessors"] = "PASS"
        except Exception as e:
            print(f"FAIL: {e}")
            results["bulk_add_predecessors"] = "FAIL"
    else:
        print("SKIP (no tasks)")
        results["bulk_add_predecessors"] = "SKIP"

    # Test 6: bulk_update_tasks
    print("\n=== Test 6: bulk_update_tasks ===")
    if task_uids:
        try:
            updates = [
                {"unique_id": task_uids["Requirements Gathering"], "rag": "Green", "text2": "Low Risk"},
                {"unique_id": task_uids["Stakeholder Review"], "rag": "Amber", "percent_complete": 25},
            ]
            r = await call("bulk_update_tasks", {"updates_json": json.dumps(updates)})
            print(f"Updated: {r['updated']}, Not found: {r['not_found']}")
            results["bulk_update_tasks"] = "PASS"
        except Exception as e:
            print(f"FAIL: {e}")
            results["bulk_update_tasks"] = "FAIL"
    else:
        print("SKIP (no tasks)")
        results["bulk_update_tasks"] = "SKIP"

    # Test 7: bulk_set_task_mode (scope-based)
    print("\n=== Test 7: bulk_set_task_mode ===")
    try:
        r = await call("bulk_set_task_mode", {"updates_json": json.dumps({"mode": "manual", "scope": "all"})})
        print(f"Set to manual: {r['updated']} tasks")
        results["bulk_set_task_mode"] = "PASS"
    except Exception as e:
        print(f"FAIL: {e}")
        results["bulk_set_task_mode"] = "FAIL"

    # Test 8: clear_estimated_flags
    print("\n=== Test 8: clear_estimated_flags ===")
    try:
        r = await call("clear_estimated_flags")
        print(f"Cleared estimated flags on {r['tasks_updated']} tasks")
        results["clear_estimated_flags"] = "PASS"
    except Exception as e:
        print(f"FAIL: {e}")
        results["clear_estimated_flags"] = "FAIL"

    # Test 9: set_constraint
    print("\n=== Test 9: set_constraint ===")
    if task_uids:
        try:
            r = await call("set_constraint", {
                "unique_id": task_uids["Requirements Gathering"],
                "constraint_type": "SNET",
                "constraint_date": "2026-04-01",
            })
            print(f"Constraint: {r['constraint_type']} on {r['name']}")
            results["set_constraint"] = "PASS"
        except Exception as e:
            print(f"FAIL: {e}")
            results["set_constraint"] = "FAIL"
    else:
        print("SKIP (no tasks)")
        results["set_constraint"] = "SKIP"

    # Test 10: rename_custom_fields
    print("\n=== Test 10: rename_custom_fields ===")
    try:
        fields = {"text1": "RAG Status", "text2": "Risk Profile", "text3": "Technology Required"}
        r = await call("rename_custom_fields", {"fields_json": json.dumps(fields)})
        print(f"Renamed: {r['renamed']} fields")
        for item in r["results"]:
            print(f"  {item}")
        results["rename_custom_fields"] = "PASS"
    except Exception as e:
        print(f"FAIL: {e}")
        results["rename_custom_fields"] = "FAIL"

    # Test 11: update_task with new duration_days and manual params
    print("\n=== Test 11: update_task (duration_days + manual) ===")
    if task_uids:
        try:
            r = await call("update_task", {
                "unique_id": task_uids["Requirements Gathering"],
                "duration_days": 12,
                "manual": False,
            })
            print(f"Updated: {r['name']}, changed: {r['changed']}")
            results["update_task_enhanced"] = "PASS"
        except Exception as e:
            print(f"FAIL: {e}")
            results["update_task_enhanced"] = "FAIL"
    else:
        print("SKIP (no tasks)")
        results["update_task_enhanced"] = "SKIP"

    # Summary
    print("\n" + "=" * 50)
    print("TEST SUMMARY")
    print("=" * 50)
    passed = failed = skipped = 0
    for test_name, status in results.items():
        print(f"  [{status:4s}] {test_name}")
        if status == "PASS":
            passed += 1
        elif status == "SKIP":
            skipped += 1
        else:
            failed += 1
    print(f"\n  {passed} passed, {failed} failed, {skipped} skipped / {len(results)} total")

    # Cleanup: close without saving
    print("\nCleaning up test project...")
    try:
        await call("close_project", {"save": False})
        print("Test project closed.")
    except Exception:
        pass


if __name__ == "__main__":
    asyncio.run(test())
