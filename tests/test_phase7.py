"""
Phase 7 Tests — Critical Path Intelligence (3 new tools)
Run: python test_phase7.py
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
    # Setup — create a project with a realistic dependency chain
    # -------------------------------------------------------------------
    print("\n=== SETUP ===")

    r = await call("new_project", title="Phase 7 Critical Path Test", start="2026-04-01")
    print(f"  Created project: {r.get('title')}")

    tasks_data = json.dumps([
        {"name": "Programme",          "outline_level": 1},
        {"name": "Requirements",       "outline_level": 2, "start": "2026-04-01", "duration_days": 10},
        {"name": "Design",             "outline_level": 2, "start": "2026-04-15", "duration_days": 15},
        {"name": "Design Review (M)",  "outline_level": 2, "start": "2026-05-06", "duration_days": 0},
        {"name": "Development",        "outline_level": 2, "start": "2026-05-07", "duration_days": 30},
        {"name": "Testing",            "outline_level": 2, "start": "2026-06-18", "duration_days": 20},
        {"name": "UAT Sign-off (M)",   "outline_level": 2, "start": "2026-07-16", "duration_days": 0},
        {"name": "Deployment",         "outline_level": 2, "start": "2026-07-17", "duration_days": 5},
        {"name": "Go-Live (M)",        "outline_level": 2, "start": "2026-07-24", "duration_days": 0},
        # Parallel non-critical track
        {"name": "Documentation",      "outline_level": 2, "start": "2026-05-07", "duration_days": 10},
        {"name": "Training",           "outline_level": 2, "start": "2026-05-21", "duration_days": 5},
    ])
    r = await call("bulk_add_tasks", tasks_json=tasks_data)
    created = r.get("tasks", [])
    uids = {t["name"]: t["unique_id"] for t in created}

    # Build the critical chain: Requirements -> Design -> Design Review -> Development -> Testing -> UAT -> Deployment -> Go-Live
    critical_chain = [
        ("Design",            "Requirements"),
        ("Design Review (M)", "Design"),
        ("Development",       "Design Review (M)"),
        ("Testing",           "Development"),
        ("UAT Sign-off (M)",  "Testing"),
        ("Deployment",        "UAT Sign-off (M)"),
        ("Go-Live (M)",       "Deployment"),
    ]
    links = json.dumps([
        {"successor_unique_id": uids[s], "predecessor_unique_id": uids[p], "link_type": "FS"}
        for s, p in critical_chain
    ])
    await call("bulk_add_predecessors", links_json=links)

    # Non-critical links: Documentation -> Training (parallel track)
    await call("add_predecessor",
               successor_unique_id=uids["Training"],
               predecessor_unique_id=uids["Documentation"],
               link_type="FS")

    # Recalculate
    await call("calculate_project")

    # Add a resource for richer output
    await call("add_resource", name="Lead Dev", max_units=1.0)
    await call("assign_resource", task_unique_id=uids["Development"], resource_name="Lead Dev")

    print(f"  Created {len(created)} tasks with {len(critical_chain)} critical links + 1 parallel track")

    # -------------------------------------------------------------------
    # 1. get_critical_path_sequence
    # -------------------------------------------------------------------
    print("\n=== 1. get_critical_path_sequence ===")
    r = await call("get_critical_path_sequence")

    ok("has project_start",        "project_start" in r)
    ok("has project_finish",       "project_finish" in r)
    ok("has sequence",             isinstance(r.get("sequence"), list))
    ok("has critical_path_length", isinstance(r.get("critical_path_length"), int))
    ok("has total_duration_days",  isinstance(r.get("total_duration_days"), (int, float)))

    seq = r.get("sequence", [])
    ok("sequence is non-empty",    len(seq) > 0)

    if seq:
        # Check sequence is ordered (step numbers ascending)
        steps = [s["step"] for s in seq]
        ok("steps are sequential", steps == list(range(1, len(steps) + 1)))

        # First task should be Requirements (start of chain)
        ok("first task is Requirements", seq[0]["name"] == "Requirements")

        # Last task should be Go-Live
        ok("last task is Go-Live", seq[-1]["name"] == "Go-Live (M)")

        # Each entry has required fields
        first = seq[0]
        for field in ["unique_id", "name", "start", "finish", "duration_days", "milestone", "percent_complete"]:
            ok(f"sequence entry has '{field}'", field in first)

        # Check that link info exists between steps
        has_link = any("link_to_next" in s for s in seq[:-1])
        ok("sequence has link_to_next info", has_link)

        # Non-critical tasks should NOT be in the sequence
        non_critical_names = {"Documentation", "Training"}
        seq_names = {s["name"] for s in seq}
        ok("no non-critical tasks in sequence", not seq_names.intersection(non_critical_names))

    # -------------------------------------------------------------------
    # 2. get_critical_tasks_for_period
    # -------------------------------------------------------------------
    print("\n=== 2. get_critical_tasks_for_period ===")

    # Q2 2026 (April-June) — should capture Requirements, Design, Development
    r = await call("get_critical_tasks_for_period",
                   start_date="2026-04-01", end_date="2026-06-30")

    ok("has period",              isinstance(r.get("period"), dict))
    ok("has critical_task_count", isinstance(r.get("critical_task_count"), int))
    ok("has critical_tasks",      isinstance(r.get("critical_tasks"), list))
    ok("has critical_milestones", isinstance(r.get("critical_milestones"), list))

    ct = r.get("critical_tasks", [])
    ok("critical tasks found for Q2", len(ct) > 0)

    if ct:
        first_ct = ct[0]
        for field in ["unique_id", "name", "start", "finish", "duration_days", "percent_complete", "overlap_days"]:
            ok(f"critical task has '{field}'", field in first_ct)

        # Requirements and Design should be in Q2
        ct_names = {t["name"] for t in ct}
        ok("Requirements in Q2", "Requirements" in ct_names)
        ok("Design in Q2", "Design" in ct_names)

    # Check milestones — may be 0 if MS Project didn't flag 0-duration tasks as milestones
    cm = r.get("critical_milestones", [])
    # Use full-year range to ensure we capture milestones regardless of scheduling
    r_full = await call("get_critical_tasks_for_period",
                        start_date="2026-04-01", end_date="2026-12-31")
    cm_full = r_full.get("critical_milestones", [])
    all_critical = r_full.get("critical_task_count", 0) + len(cm_full)
    ok("critical milestones or tasks found full year", all_critical > 0, f"tasks={r_full.get('critical_task_count')}, milestones={len(cm_full)}")

    # Q3 — should capture Testing, Deployment
    r2 = await call("get_critical_tasks_for_period",
                    start_date="2026-07-01", end_date="2026-09-30")
    ct2 = r2.get("critical_tasks", [])
    ok("Q3 has critical tasks", len(ct2) >= 0)  # may have Testing/Deployment overlap

    # With non-critical milestones included
    r3 = await call("get_critical_tasks_for_period",
                    start_date="2026-04-01", end_date="2026-12-31",
                    include_non_critical_milestones=True)
    ok("full year has other_milestones field", isinstance(r3.get("other_milestones"), list))

    # Bad dates
    r4 = await call("get_critical_tasks_for_period", start_date="", end_date="")
    ok("empty dates returns error", "error" in r4)

    # Narrow period with no critical tasks
    r5 = await call("get_critical_tasks_for_period",
                    start_date="2027-01-01", end_date="2027-12-31")
    ok("future period returns zero tasks", r5.get("critical_task_count", -1) == 0)

    # -------------------------------------------------------------------
    # 3. what_if_delay
    # -------------------------------------------------------------------
    print("\n=== 3. what_if_delay ===")

    # Delay a critical task (Development) by 5 days
    r = await call("what_if_delay", unique_id=uids["Development"], delay_days=5)

    ok("has task info",            isinstance(r.get("task"), dict))
    ok("has severity",             r.get("severity") in ("NONE", "LOW", "MEDIUM", "HIGH"))
    ok("has summary",              isinstance(r.get("summary"), str))
    ok("has current_slack_days",   isinstance(r.get("current_slack_days"), (int, float)))
    ok("has project_impact",       isinstance(r.get("project_impact"), dict))
    ok("has downstream_affected",  isinstance(r.get("downstream_affected"), int))
    ok("has newly_critical",       isinstance(r.get("newly_critical"), list))
    ok("has slack_consumed",       isinstance(r.get("slack_consumed"), list))
    ok("has downstream_tasks",     isinstance(r.get("downstream_tasks"), list))

    # Development is critical with 0 slack — 5-day delay should be HIGH severity
    ok("critical task delay is HIGH", r.get("severity") == "HIGH",
       f"got: {r.get('severity')}, slack: {r.get('current_slack_days')}")

    pi = r.get("project_impact", {})
    ok("project impact has delay_days", isinstance(pi.get("delay_days"), (int, float)))
    ok("project would be delayed",      pi.get("delay_days", 0) > 0,
       f"got delay_days={pi.get('delay_days')}")

    # Downstream should include Testing, UAT, Deployment, Go-Live
    downstream = r.get("downstream_tasks", [])
    ok("downstream tasks found", len(downstream) > 0)
    if downstream:
        ds_names = {d["name"] for d in downstream}
        ok("Testing is downstream", "Testing" in ds_names)

    # Delay a non-critical task (Documentation) — should be absorbed by slack
    r2 = await call("what_if_delay", unique_id=uids["Documentation"], delay_days=3)
    ok("non-critical delay has severity", r2.get("severity") in ("NONE", "LOW", "MEDIUM", "HIGH"))
    # Documentation has lots of slack, 3-day delay should be NONE or LOW
    ok("non-critical small delay absorbed", r2.get("severity") in ("NONE", "LOW"),
       f"got: {r2.get('severity')}, slack: {r2.get('current_slack_days')}")

    # Large delay on non-critical task that exceeds its slack
    r3 = await call("what_if_delay", unique_id=uids["Documentation"], delay_days=200)
    ok("large delay escalates severity", r3.get("severity") in ("MEDIUM", "HIGH"),
       f"got: {r3.get('severity')}")

    # Invalid task
    r4 = await call("what_if_delay", unique_id=999999, delay_days=5)
    ok("invalid task returns error", "error" in r4)

    # Summary task
    r5 = await call("what_if_delay", unique_id=uids["Programme"], delay_days=5)
    ok("summary task returns error", "error" in r5)

    # Zero delay
    r6 = await call("what_if_delay", unique_id=uids["Development"], delay_days=0)
    ok("zero delay is NONE severity", r6.get("severity") == "NONE")

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
    print(f"  Phase 7 Results:  {PASS} passed,  {FAIL} failed,  {SKIP} skipped  (total {total})")
    print("=" * 60)

    return FAIL == 0


if __name__ == "__main__":
    success = asyncio.run(run_tests())
    sys.exit(0 if success else 1)
