# MS Project MCP Server

Control Microsoft Project via COM automation through the Model Context Protocol (MCP).

## Prerequisites

- **Windows** with Microsoft Project installed (tested on MS Project 16.0)
- **Python 3.10+**
- **mcp** package: `pip install mcp`
- **pywin32** for COM: `pip install pywin32`
- **python-dateutil** (optional, for `add_recurring_task`): `pip install python-dateutil`

## Quick Start

1. **Start MS Project** and open a project file (or the server will create one).

2. **Register** in your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "msproject": {
      "command": "python",
      "args": ["/path/to/msproject/server.py"]
    }
  }
}
```

3. **Run standalone** (for testing):

```bash
python server.py
```

## Tool Inventory (99 tools)

### Project Management (7)

| Tool | Description |
|------|-------------|
| `open_project` | Open an existing .mpp file |
| `new_project` | Create a new blank project |
| `get_project_info` | Read project metadata and summary stats |
| `set_project_properties` | Update title, manager, start date, etc. |
| `save_project` | Save current project |
| `save_project_as` | Save as .mpp or .xml |
| `close_project` | Close the active project |

### Task Queries (9)

| Tool | Description |
|------|-------------|
| `get_tasks` | List tasks with optional filters |
| `get_task` | Get a single task by UniqueID |
| `get_critical_path` | Return critical-path tasks |
| `get_tasks_by_rag` | Filter tasks by RAG status (Text1) |
| `get_overdue_tasks` | Incomplete tasks past their finish date |
| `get_tasks_by_resource` | Tasks assigned to a named resource |
| `search_tasks` | Full-text search across task names |
| `get_progress_summary` | Dashboard: progress, RAG counts, overdue |
| `get_wbs_structure` | Hierarchical WBS tree |

### Task Mutations (12)

| Tool | Description |
|------|-------------|
| `update_task` | Update any task field (incl. priority, task type) |
| `bulk_update_rag` | Mass RAG status update |
| `bulk_update_tasks` | Mass field updates across tasks |
| `add_task` | Add a single task |
| `bulk_add_tasks` | Add many tasks at once (JSON array) |
| `add_recurring_task` | Create recurring task occurrences (daily/weekly/monthly) |
| `delete_task` | Remove a task |
| `set_task_mode` | Toggle auto/manual scheduling |
| `bulk_set_task_mode` | Mass scheduling mode update |
| `set_constraint` | Set scheduling constraint (ASAP/ALAP/SNET/etc.) |
| `clear_estimated_flags` | Clear estimated flag on all tasks |
| `indent_task` | Indent or outdent a task |

### Dependencies (4)

| Tool | Description |
|------|-------------|
| `add_predecessor` | Create FS/SS/FF/SF link with optional lag |
| `bulk_add_predecessors` | Mass predecessor creation |
| `remove_predecessor` | Delete a dependency link |
| `get_task_dependencies` | Show predecessors and successors |

### Resources (7)

| Tool | Description |
|------|-------------|
| `get_resources` | List all resources in the pool |
| `add_resource` | Add a work/material/cost resource |
| `assign_resource` | Assign a resource to a task |
| `update_resource` | Modify resource properties |
| `delete_resource` | Remove a resource (clears assignments) |
| `set_resource_calendar` | Assign a specific calendar to a resource |
| `get_resource_availability` | Allocation vs capacity per period |

### Resource Assignments (3)

| Tool | Description |
|------|-------------|
| `bulk_assign_resources` | Mass resource assignment |
| `remove_resource_assignment` | Unassign a resource from a task |
| `get_resource_workload` | Workload and conflict detection for a resource |

### Resource Rate Tables (2)

| Tool | Description |
|------|-------------|
| `get_resource_rate_tables` | Read cost rate tables A-E for a resource |
| `set_resource_rate_table` | Set or add rate entries in a cost rate table |

### Custom Fields (3)

| Tool | Description |
|------|-------------|
| `rename_custom_fields` | Rename Text1-Text30, Number1-Number20, etc. |
| `update_custom_fields` | Set custom field values on a task |
| `get_custom_field_values` | Read all values of a custom field |

### Import / Export (6)

| Tool | Description |
|------|-------------|
| `import_xml` | Import from MS Project XML |
| `export_xml` | Export to MS Project XML |
| `export_csv` | Export tasks to CSV with column selection |
| `snapshot_to_json` | Full project snapshot as JSON |
| `snapshot_diff` | Compare two JSON snapshots — additions, deletions, changes |
| `insert_subproject` | Insert a subproject file |

### Calendars (7)

| Tool | Description |
|------|-------------|
| `get_calendars` | List all project calendars |
| `create_calendar` | Create a new calendar (optionally copy from existing) |
| `delete_calendar` | Remove a base calendar |
| `set_calendar_exception` | Add exception dates (holidays) to a calendar |
| `delete_calendar_exception` | Remove a specific exception from a calendar |
| `list_calendar_exceptions` | List all exceptions on a calendar |
| `set_project_calendar` | Switch the project base calendar |

### Calendar Working Hours (1)

| Tool | Description |
|------|-------------|
| `set_working_hours` | Modify working hours for a specific day of the week |

### Scheduling & Analysis (12)

| Tool | Description |
|------|-------------|
| `get_schedule_analysis` | Critical path length, float analysis |
| `validate_schedule` | Find scheduling issues (missing links, etc.) |
| `calculate_project` | Recalculate the schedule after manual changes |
| `get_milestone_report` | Upcoming and overdue milestones |
| `level_resources` | Run MS Project resource leveling |
| `find_available_slack` | Tasks with free slack above threshold |
| `get_constraints` | Read non-default constraints on all tasks |
| `set_task_calendar` | Assign a calendar to a specific task |
| `set_task_hyperlink` | Set a hyperlink on a task |
| `get_critical_path_sequence` | Ordered critical path chain from start to finish with driving links |
| `get_critical_tasks_for_period` | Critical tasks and milestones within a date range (e.g. Q2) |
| `what_if_delay` | Simulate delaying a task — shows project impact, newly critical tasks, slack loss |

### Status Date & Progress Updates (2)

| Tool | Description |
|------|-------------|
| `update_project` | Mark all tasks complete through a date (weekly PMO ritual) |
| `reschedule_incomplete_work` | Move remaining work to start after a given date |

### Timephased Data (1)

| Tool | Description |
|------|-------------|
| `get_timephased_data` | Period-by-period work/cost data (for S-curves, cash flow) |

### Baselines & Earned Value (4)

| Tool | Description |
|------|-------------|
| `save_baseline` | Save baseline (0-10) |
| `clear_baseline` | Clear a saved baseline |
| `compare_baselines` | Compare two baselines or baseline vs current |
| `get_earned_value` | BCWS, BCWP, ACWP, SPI, CPI per task |

### Variance & Reporting (1)

| Tool | Description |
|------|-------------|
| `get_variance_report` | Schedule + cost variance per task vs a baseline |

### Cost & Work (2)

| Tool | Description |
|------|-------------|
| `get_cost_summary` | Budget vs actual cost breakdown |
| `get_actual_work` | Actual vs remaining work hours per task |

### Progress Tracking (2)

| Tool | Description |
|------|-------------|
| `get_progress_by_wbs` | Completion % rolled up by WBS level |
| `get_dependency_chain` | Walk predecessor/successor chains |

### Advanced Operations (8)

| Tool | Description |
|------|-------------|
| `set_deadline` | Set deadline indicator on a task |
| `bulk_set_deadlines` | Mass deadline assignment |
| `set_task_active` | Activate/inactivate a task |
| `dry_run_bulk_update` | Preview bulk changes without applying |
| `move_task` | Reorder a task after another |
| `copy_task_structure` | Duplicate a task and its subtree |
| `cross_project_link` | Create inter-project dependency |
| `undo_last` | Undo recent operations (up to 10) |

### Multi-Project (3)

| Tool | Description |
|------|-------------|
| `list_projects` | List all open projects |
| `switch_project` | Switch active project by name or index |
| `apply_filter` | Apply a built-in or custom filter |

### Filtering & Grouping (2)

| Tool | Description |
|------|-------------|
| `filter_tasks` | Advanced multi-field filtering |
| `group_tasks_by` | Group tasks by any field with aggregation |

### Connectivity (1)

| Tool | Description |
|------|-------------|
| `health_check` | Lightweight COM connectivity test — version, project status |

## Enriched Task Data

Every task query (`get_task`, `get_tasks`, etc.) returns a rich dict with 35+ fields including:

| Field | Description |
|-------|-------------|
| `actual_start` / `actual_finish` | When work actually began/ended |
| `remaining_duration_days` | How much work is left |
| `total_slack_days` / `free_slack_days` | Schedule flexibility |
| `deadline` | Soft deadline indicator |
| `priority` | Leveling priority (0–1000) |
| `constraint_type` / `constraint_date` | Scheduling constraint |
| `manual` | Auto vs manual scheduling |
| `type` | FixedUnits / FixedDuration / FixedWork |
| `hyperlink` / `hyperlink_text` | Task hyperlink |

## Known Limitations

- **COM proxy staleness**: When multiple projects are open, switching projects invalidates existing COM references. Always call `switch_project` before operating on a different file.
- **Undo stack**: `undo_last` supports up to 10 consecutive undos. MS Project's COM undo is less reliable than the UI's.
- **File locking**: Only one process can hold the COM connection. Don't open MS Project's GUI dialogs while the server is active.
- **Timezone-aware dates**: COM may return timezone-aware datetimes. The server normalizes these via `_to_naive()` for safe comparisons.
- **Recurring tasks**: MS Project's `RecurringTaskInsert` is a dialog-only COM method. The `add_recurring_task` tool simulates recurrence by creating individual occurrences under a summary task. Requires `python-dateutil`.
- **Timephased data**: `get_timephased_data` can be slow on large date ranges. Keep queries to reasonable periods (weeks/months, not years).

## Tests

```bash
# Run all phases
python tests/test_new_tools.py  # 11 tests (core CRUD)
python tests/test_phase2.py     # 10 tests (resources, baselines, WBS)
python tests/test_phase3.py     # 15 tests (custom fields, calendars, scheduling)
python tests/test_phase4.py     # 23 tests (advanced ops, multi-project, filtering)
python tests/test_phase5.py     # 11 tests (bug fixes, cost/work tracking)
python tests/test_phase6.py     # 75 tests (timephased data, calendar mgmt, variance)
python tests/test_phase7.py     # 57 tests (critical path intelligence, what-if)
```

**202 tests total** across 7 test suites. All tests require MS Project to be running (they create and close temporary projects).

## Architecture

Single-file server (`server.py`, ~5,200 lines) using the FastMCP framework. All COM calls go through `get_app()` / `get_proj()` helpers. Dates are normalized with `_to_naive()` and formatted with `_fmt_date()`.

### Development Phases

| Phase | Tools | Focus |
|-------|-------|-------|
| 1–2 | 25 | Core CRUD, dependencies, resources |
| 3 | 44 | Custom fields, calendars, scheduling analysis |
| 4 | 65 | Advanced operations, multi-project, filtering |
| 5 | 79 | Bug fixes, cost/work tracking |
| 6 | 96 | Timephased data, calendar management, resource availability, variance reporting |
| 7 | 99 | Critical path intelligence: ordered sequence, period filtering, what-if analysis |

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, conventions, and how to submit changes.

## License

This project is licensed under the MIT License — see [LICENSE](LICENSE) for details.
