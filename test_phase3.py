"""Integration test for Phase 3 MCP server tools (15 new features)."""
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

    # =======================================================================
    # SETUP: Create test project with hierarchy, resources, links, RAG
    # =======================================================================
    print("=== Setup: new_project ===")
    r = await call("new_project", {"title": "Phase 3 Test", "start": "2026-04-01"})
    main_project_name = r["name"]  # Actual name (e.g. "Project1")
    print(f"  Project: {main_project_name}")

    print("=== Setup: bulk_add_tasks ===")
    tasks = [
        {"name": "Programme Alpha", "outline_level": 1},
        {"name": "Design Work", "outline_level": 2, "start": "2026-04-01", "finish": "2026-05-01", "duration_days": 22},
        {"name": "Build Work", "outline_level": 2, "start": "2026-05-02", "finish": "2026-06-30", "duration_days": 43},
        {"name": "Testing", "outline_level": 3, "start": "2026-06-01", "finish": "2026-06-20", "duration_days": 15},
        {"name": "Go-Live Milestone", "outline_level": 2, "milestone": True, "start": "2026-06-30"},
        {"name": "Programme Beta", "outline_level": 1},
        {"name": "Planning", "outline_level": 2, "start": "2026-04-01", "finish": "2026-04-15", "duration_days": 11},
        {"name": "Execution", "outline_level": 2, "start": "2026-04-16", "finish": "2026-07-31", "duration_days": 77},
    ]
    r = await call("bulk_add_tasks", {"tasks_json": json.dumps(tasks)})
    print(f"  Created: {r['created']} tasks")
    task_uids = {t["name"]: t["unique_id"] for t in r["tasks"]}

    # Add resources
    print("=== Setup: add_resource ===")
    await call("add_resource", {"name": "Alice", "type": 0, "max_units": 1.0})
    await call("add_resource", {"name": "Bob", "type": 0, "max_units": 1.0})
    print("  Resources: Alice, Bob")

    # Assign resources
    print("=== Setup: assign_resource ===")
    await call("assign_resource", {"task_unique_id": task_uids["Design Work"], "resource_name": "Alice"})
    await call("assign_resource", {"task_unique_id": task_uids["Build Work"], "resource_name": "Alice"})
    await call("assign_resource", {"task_unique_id": task_uids["Testing"], "resource_name": "Bob"})
    print("  Assignments: Alice->Design+Build, Bob->Testing")

    # Add predecessor links
    print("=== Setup: bulk_add_predecessors ===")
    links = [
        {"successor_unique_id": task_uids["Build Work"], "predecessor_unique_id": task_uids["Design Work"], "link_type": "FS"},
        {"successor_unique_id": task_uids["Go-Live Milestone"], "predecessor_unique_id": task_uids["Build Work"], "link_type": "FS"},
        {"successor_unique_id": task_uids["Execution"], "predecessor_unique_id": task_uids["Planning"], "link_type": "FS"},
    ]
    await call("bulk_add_predecessors", {"links_json": json.dumps(links)})
    print("  Links: 3 added")

    # Set RAG on some tasks
    print("=== Setup: bulk_update_rag ===")
    rag_updates = [
        {"unique_id": task_uids["Design Work"], "rag": "Green"},
        {"unique_id": task_uids["Build Work"], "rag": "Amber"},
        {"unique_id": task_uids["Testing"], "rag": "Red"},
        {"unique_id": task_uids["Planning"], "rag": "Green"},
    ]
    await call("bulk_update_rag", {"updates": json.dumps(rag_updates)})
    print("  RAG set on 4 tasks")

    # ===================================================================
    # Test 1: list_projects
    # ===================================================================
    print("\n=== Test 1: list_projects ===")
    try:
        r = await call("list_projects")
        assert r["count"] >= 1, f"Expected >= 1 project, got {r['count']}"
        assert r["active"], "No active project"
        proj = r["projects"][0]
        assert "name" in proj and "task_count" in proj and "start" in proj
        has_active = any(p["is_active"] for p in r["projects"])
        assert has_active, "No project marked as active"
        print(f"  Count: {r['count']}, Active: {r['active']}")
        results["list_projects"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["list_projects"] = "FAIL"

    # ===================================================================
    # Test 2: switch_project (open 2nd project, switch, switch back)
    # ===================================================================
    print("\n=== Test 2: switch_project ===")
    try:
        # Create a second project
        r2 = await call("new_project", {"title": "Switch Test", "start": "2026-01-01"})
        proj2_name = r2["name"]
        print(f"  Created 2nd project: {proj2_name}")

        # Switch back to first by name substring (use actual name)
        r = await call("switch_project", {"name_or_index": main_project_name})
        assert r["status"] in ("switched", "already_active"), f"Unexpected status: {r['status']}"
        print(f"  Switched to: {r['name']}")

        # Switch by index to project 2
        # Find which index the 2nd project is at
        lp = await call("list_projects")
        proj2_idx = None
        for p in lp["projects"]:
            if p["name"] == proj2_name:
                proj2_idx = str(p["index"])
                break
        assert proj2_idx, f"Could not find {proj2_name} in project list"
        r = await call("switch_project", {"name_or_index": proj2_idx})
        assert r["status"] == "switched", f"Expected switched, got {r['status']}"
        print(f"  Switched by index to: {r['name']}")

        # Test invalid name
        r = await call("switch_project", {"name_or_index": "NONEXISTENT_XYZ"})
        assert "error" in r, "Expected error for invalid name"
        print(f"  Invalid name handled: {r['error'][:50]}")

        # Switch back to main project for remaining tests
        await call("switch_project", {"name_or_index": main_project_name})
        print(f"  Switched back to: {main_project_name}")

        results["switch_project"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["switch_project"] = "FAIL"
    finally:
        # Safety: always ensure we're back on the main project
        try:
            await call("switch_project", {"name_or_index": main_project_name})
        except Exception:
            pass

    # ===================================================================
    # Test 3: filter_tasks
    # ===================================================================
    print("\n=== Test 3: filter_tasks ===")
    try:
        # Filter by outline_level
        r = await call("filter_tasks", {"filters_json": json.dumps({"outline_level": 2, "summary": False})})
        assert r["total_matching"] > 0, "No tasks at outline level 2"
        print(f"  Level 2 tasks: {r['total_matching']}")

        # Filter milestones
        r = await call("filter_tasks", {"filters_json": json.dumps({"milestone": True})})
        assert r["total_matching"] >= 1, "No milestones found"
        print(f"  Milestones: {r['total_matching']}")

        # Test pagination
        r = await call("filter_tasks", {"filters_json": json.dumps({"limit": 2, "offset": 0})})
        assert r["returned"] <= 2, f"Limit not respected: got {r['returned']}"
        print(f"  Pagination: returned={r['returned']}, total={r['total_matching']}")

        # Filter by RAG
        r = await call("filter_tasks", {"filters_json": json.dumps({"rag": "Red"})})
        assert r["total_matching"] >= 1, "No Red RAG tasks"
        print(f"  Red RAG tasks: {r['total_matching']}")

        results["filter_tasks"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["filter_tasks"] = "FAIL"

    # ===================================================================
    # Test 4: group_tasks_by
    # ===================================================================
    print("\n=== Test 4: group_tasks_by ===")
    try:
        # Group by outline_level
        r = await call("group_tasks_by", {"field": "outline_level"})
        total_from_groups = sum(g["count"] for g in r["groups"])
        assert total_from_groups == r["total_tasks"], f"Group counts don't sum: {total_from_groups} vs {r['total_tasks']}"
        print(f"  By outline_level: {len(r['groups'])} groups, {r['total_tasks']} tasks")

        # Group by milestone
        r = await call("group_tasks_by", {"field": "milestone"})
        print(f"  By milestone: {[(g['value'], g['count']) for g in r['groups']]}")

        # Group by critical
        r = await call("group_tasks_by", {"field": "critical"})
        print(f"  By critical: {[(g['value'], g['count']) for g in r['groups']]}")

        results["group_tasks_by"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["group_tasks_by"] = "FAIL"

    # ===================================================================
    # Test 5: set_calendar_exception
    # ===================================================================
    print("\n=== Test 5: set_calendar_exception ===")
    try:
        r = await call("set_calendar_exception", {
            "calendar_name": "Standard",
            "name": "Test Holiday",
            "start": "2026-12-25",
            "finish": "2026-12-25",
            "working": False,
        })
        assert r.get("status") == "created", f"Expected created, got: {r}"
        print(f"  Exception added: {r['exception']} on {r['start']}")

        # Test invalid calendar
        r = await call("set_calendar_exception", {
            "calendar_name": "FAKE_CALENDAR",
            "name": "Test",
            "start": "2026-01-01",
            "finish": "2026-01-01",
        })
        assert "error" in r, "Expected error for invalid calendar"
        print(f"  Invalid calendar handled correctly")

        results["set_calendar_exception"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["set_calendar_exception"] = "FAIL"

    # ===================================================================
    # Test 6: set_project_calendar
    # ===================================================================
    print("\n=== Test 6: set_project_calendar ===")
    try:
        # First get available calendars
        cals = await call("get_calendars")
        active_cal = cals["active_calendar"]
        other_cal = None
        for c in cals["calendars"]:
            if c != active_cal:
                other_cal = c
                break

        if not other_cal:
            # Only one calendar — test set to same (should still work)
            r = await call("set_project_calendar", {"calendar_name": active_cal})
            assert r["status"] == "updated", f"Expected updated, got {r}"
            print(f"  Only one calendar '{active_cal}' — set to same, status OK")

            # Test invalid calendar
            r = await call("set_project_calendar", {"calendar_name": "FAKE_CAL_XYZ"})
            assert "error" in r, "Expected error for invalid calendar"
            print(f"  Invalid calendar handled correctly")
        else:
            # Switch to other calendar
            r = await call("set_project_calendar", {"calendar_name": other_cal})
            assert r["status"] == "updated", f"Expected updated, got {r}"
            assert r["calendar"] == other_cal
            print(f"  Switched to '{other_cal}' from '{r['previous']}'")

            # Switch back
            r = await call("set_project_calendar", {"calendar_name": active_cal})
            assert r["status"] == "updated"
            print(f"  Switched back to '{active_cal}'")

        results["set_project_calendar"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["set_project_calendar"] = "FAIL"

    # ===================================================================
    # Test 7: update_custom_fields
    # ===================================================================
    print("\n=== Test 7: update_custom_fields ===")
    try:
        uid = task_uids["Design Work"]
        r = await call("update_custom_fields", {
            "unique_id": uid,
            "fields_json": json.dumps({
                "Text5": "Phase A",
                "Number1": 42,
                "Flag3": True,
            }),
        })
        assert r["status"] == "updated", f"Expected updated, got {r}"
        if r["errors"]:
            print(f"  Errors: {r['errors']}")
        assert len(r["changed"]) == 3, f"Expected 3 changes, got {len(r['changed'])}. Errors: {r['errors']}"
        assert len(r["errors"]) == 0, f"Unexpected errors: {r['errors']}"
        print(f"  Set 3 custom fields on '{r['name']}': {[c['field'] for c in r['changed']]}")

        # Verify Text5 via get_custom_field_values
        r2 = await call("get_custom_field_values", {"field_name": "Text5"})
        assert "Phase A" in r2["unique_values"], f"'Phase A' not in Text5 values: {r2['unique_values']}"
        print(f"  Verified Text5 contains 'Phase A'")

        results["update_custom_fields"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["update_custom_fields"] = "FAIL"

    # ===================================================================
    # Test 8: get_custom_field_values
    # ===================================================================
    print("\n=== Test 8: get_custom_field_values ===")
    try:
        # Query Text1 (RAG field) — should have Green, Amber, Red from setup
        r = await call("get_custom_field_values", {"field_name": "Text1"})
        assert r["total_tasks"] > 0, "No tasks found"
        print(f"  Text1 values: {r['unique_values']}")
        print(f"  Counts: {r['value_counts']}")
        # We set Green, Amber, Red — at least those should exist
        for rag in ["Green", "Amber", "Red"]:
            assert rag in r["value_counts"], f"Missing '{rag}' in Text1 values"
        print(f"  Verified: Green, Amber, Red all present")

        results["get_custom_field_values"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["get_custom_field_values"] = "FAIL"

    # ===================================================================
    # Test 9: validate_schedule
    # ===================================================================
    print("\n=== Test 9: validate_schedule ===")
    try:
        r = await call("validate_schedule")
        assert "health_score" in r, "Missing health_score"
        assert "issues" in r, "Missing issues"
        assert r["summary"]["total_tasks"] > 0, "No tasks found"
        print(f"  Health score: {r['health_score']}")
        print(f"  Total tasks: {r['summary']['total_tasks']}, Total issues: {r['summary']['total_issues']}")
        for cat, data in r["issues"].items():
            if data["count"] > 0:
                print(f"    {cat}: {data['count']}")

        # We have tasks without resources (like Planning, Execution) — verify detection
        assert r["issues"]["no_resources"]["count"] > 0, "Should detect tasks without resources"
        print(f"  Verified: no_resources detected ({r['issues']['no_resources']['count']})")

        results["validate_schedule"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["validate_schedule"] = "FAIL"

    # ===================================================================
    # Test 10: get_milestone_report
    # ===================================================================
    print("\n=== Test 10: get_milestone_report ===")
    try:
        r = await call("get_milestone_report", {"days_ahead": 365, "upcoming_count": 5})
        assert r["total_milestones"] >= 1, f"Expected >= 1 milestone, got {r['total_milestones']}"
        print(f"  Total milestones: {r['total_milestones']}")
        print(f"  By status: {r['by_status']}")
        if r["upcoming"]:
            print(f"  Upcoming: {r['upcoming'][0]['name']} ({r['upcoming'][0]['finish']})")
        results["get_milestone_report"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["get_milestone_report"] = "FAIL"

    # ===================================================================
    # Test 11: get_resource_workload
    # ===================================================================
    print("\n=== Test 11: get_resource_workload ===")
    try:
        r = await call("get_resource_workload", {"resource_name": "Alice"})
        assert len(r["assignments"]) >= 2, f"Alice should have >= 2 assignments, got {len(r['assignments'])}"
        print(f"  Alice: {len(r['assignments'])} assignments, overallocated={r['overallocated']}")
        for a in r["assignments"]:
            print(f"    {a['task_name']}: {a['start']} - {a['finish']}")
        if r["conflicts"]:
            print(f"  Conflicts: {len(r['conflicts'])}")

        # Test invalid resource
        r2 = await call("get_resource_workload", {"resource_name": "NONEXISTENT"})
        assert "error" in r2, "Expected error for invalid resource"
        print(f"  Invalid resource handled correctly")

        results["get_resource_workload"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["get_resource_workload"] = "FAIL"

    # ===================================================================
    # Test 12: level_resources
    # ===================================================================
    print("\n=== Test 12: level_resources ===")
    try:
        r = await call("level_resources")
        assert r["status"] == "leveled", f"Expected leveled, got {r}"
        print(f"  Leveled project: {r['project']}")
        results["level_resources"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["level_resources"] = "FAIL"

    # ===================================================================
    # Test 13: set_deadline
    # ===================================================================
    print("\n=== Test 13: set_deadline ===")
    try:
        uid = task_uids["Build Work"]
        # Set deadline
        r = await call("set_deadline", {"unique_id": uid, "deadline_date": "2026-07-15"})
        assert r["status"] == "set", f"Expected set, got {r}"
        assert r["deadline"] == "2026-07-15"
        print(f"  Set deadline on '{r['name']}': {r['deadline']}, missed={r['deadline_missed']}")

        # Set a tight deadline that should be missed
        r = await call("set_deadline", {"unique_id": uid, "deadline_date": "2026-04-01"})
        print(f"  Tight deadline: missed={r['deadline_missed']}")

        # Clear deadline
        r = await call("set_deadline", {"unique_id": uid, "deadline_date": "clear"})
        assert r["status"] == "cleared", f"Expected cleared, got {r}"
        print(f"  Deadline cleared")

        results["set_deadline"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["set_deadline"] = "FAIL"

    # ===================================================================
    # Test 14: set_task_active
    # ===================================================================
    print("\n=== Test 14: set_task_active ===")
    try:
        uid = task_uids["Execution"]
        # Deactivate
        r = await call("set_task_active", {"unique_id": uid, "active": False})
        assert r["status"] == "updated", f"Expected updated, got {r}"
        assert r["active"] is False, "Task should be inactive"
        print(f"  Deactivated '{r['name']}'")

        # Verify via get_task
        t = await call("get_task", {"unique_id": uid})
        assert t["active"] is False, f"get_task shows active={t['active']}, expected False"
        print(f"  Verified: active={t['active']}")

        # Reactivate
        r = await call("set_task_active", {"unique_id": uid, "active": True})
        assert r["active"] is True, "Task should be active"
        print(f"  Reactivated '{r['name']}'")

        results["set_task_active"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["set_task_active"] = "FAIL"

    # ===================================================================
    # Test 15: dry_run_bulk_update
    # ===================================================================
    print("\n=== Test 15: dry_run_bulk_update ===")
    try:
        updates = [
            {"unique_id": task_uids["Design Work"], "rag": "Red", "percent_complete": 50},
            {"unique_id": task_uids["Build Work"], "name": "Construction Phase"},
            {"unique_id": 999999, "rag": "Green"},  # non-existent
        ]
        r = await call("dry_run_bulk_update", {"updates_json": json.dumps(updates)})
        assert r["preview"] is True, "Should be a preview"
        assert r["total_tasks_affected"] >= 2, f"Expected >= 2 affected, got {r['total_tasks_affected']}"
        assert 999999 in r["not_found"], "Should report 999999 as not found"
        print(f"  Preview: {r['total_changes']} changes across {r['total_tasks_affected']} tasks")
        print(f"  Not found: {r['not_found']}")

        # Verify NO actual mutation happened
        t = await call("get_task", {"unique_id": task_uids["Design Work"]})
        assert t["rag"] != "Red" or t["rag"] == "Green", "Design Work RAG should NOT have changed from dry run"
        assert t["percent_complete"] == 0, f"Design Work % should still be 0, got {t['percent_complete']}"
        print(f"  Verified: no mutation occurred (RAG={t['rag']}, %={t['percent_complete']})")

        t2 = await call("get_task", {"unique_id": task_uids["Build Work"]})
        assert t2["name"] == "Build Work", f"Build Work name should NOT have changed, got '{t2['name']}'"
        print(f"  Verified: Build Work name unchanged ('{t2['name']}')")

        results["dry_run_bulk_update"] = "PASS"
    except Exception as e:
        print(f"  FAIL: {e}")
        results["dry_run_bulk_update"] = "FAIL"

    # ===================================================================
    # Summary
    # ===================================================================
    print("\n" + "=" * 50)
    print("PHASE 3 TEST SUMMARY")
    print("=" * 50)
    passed = failed = 0
    for name, status in results.items():
        print(f"  [{status:4s}] {name}")
        if status == "PASS":
            passed += 1
        else:
            failed += 1
    print(f"\n  {passed} passed, {failed} failed / {len(results)} total")

    # Cleanup: close all test projects
    print("\nCleaning up...")
    # Close projects in reverse order to avoid index shifts
    for _ in range(5):  # max 5 attempts
        try:
            r = await call("list_projects")
            if r["count"] == 0:
                break
            await call("close_project", {"save": False})
            print(f"  Closed a project ({r['count']-1} remaining)")
        except Exception:
            break
    print("Cleanup complete.")


if __name__ == "__main__":
    asyncio.run(test())
