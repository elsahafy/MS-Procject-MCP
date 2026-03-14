# MS Project MCP Server

Control Microsoft Project via COM automation through the Model Context Protocol (MCP).

## Prerequisites

- **Windows** with Microsoft Project installed (tested on MS Project 16.0)
- **Python 3.10+**
- **mcp** package: `pip install mcp`
- **pywin32** for COM: `pip install pywin32`

## Quick Start

1. **Start MS Project** and open a project file (or the server will create one).

2. **Register** in your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "msproject": {
      "command": "python",
      "args": ["C:/Users/Ibrahim Elsahafy/mcp-servers/msproject/server.py"]
    }
  }
}
```

3. **Run standalone** (for testing):

```bash
python server.py
```

## Tool Inventory (79 tools)

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

### Task Mutations (11)

| Tool | Description |
|------|-------------|
| `update_task` | Update any task field |
| `bulk_update_rag` | Mass RAG status update |
| `bulk_update_tasks` | Mass field updates across tasks |
| `add_task` | Add a single task |
| `bulk_add_tasks` | Add many tasks at once (JSON array) |
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

### Resources (5)

| Tool | Description |
|------|-------------|
| `get_resources` | List all resources in the pool |
| `add_resource` | Add a work/material/cost resource |
| `assign_resource` | Assign a resource to a task |
| `update_resource` | Modify resource properties |
| `delete_resource` | Remove a resource (clears assignments) |

### Resource Assignments (3)

| Tool | Description |
|------|-------------|
| `bulk_assign_resources` | Mass resource assignment |
| `remove_resource_assignment` | Unassign a resource from a task |
| `get_resource_workload` | Workload and conflict detection for a resource |

### Custom Fields (3)

| Tool | Description |
|------|-------------|
| `rename_custom_fields` | Rename Text1-Text30, Number1-Number20, etc. |
| `update_custom_fields` | Set custom field values on a task |
| `get_custom_field_values` | Read all values of a custom field |

### Import / Export (5)

| Tool | Description |
|------|-------------|
| `import_xml` | Import from MS Project XML |
| `export_xml` | Export to MS Project XML |
| `export_csv` | Export tasks to CSV with column selection |
| `snapshot_to_json` | Full project snapshot as JSON |
| `insert_subproject` | Insert a subproject file |

### Calendars (4)

| Tool | Description |
|------|-------------|
| `get_calendars` | List all project calendars |
| `create_calendar` | Create a new calendar (optionally copy from existing) |
| `set_calendar_exception` | Add exception dates to a calendar |
| `set_project_calendar` | Switch the project base calendar |

### Scheduling & Analysis (7)

| Tool | Description |
|------|-------------|
| `get_schedule_analysis` | Critical path length, float analysis |
| `validate_schedule` | Find scheduling issues (missing links, etc.) |
| `get_milestone_report` | Upcoming and overdue milestones |
| `level_resources` | Run MS Project resource leveling |
| `find_available_slack` | Tasks with free slack above threshold |
| `get_constraints` | Read non-default constraints on all tasks |
| `set_task_calendar` | Assign a calendar to a specific task |

### Baselines & Earned Value (4)

| Tool | Description |
|------|-------------|
| `save_baseline` | Save baseline (0-10) |
| `clear_baseline` | Clear a saved baseline |
| `compare_baselines` | Compare two baselines or baseline vs current |
| `get_earned_value` | BCWS, BCWP, ACWP, SPI, CPI per task |

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

## Known Limitations

- **COM proxy staleness**: When multiple projects are open, switching projects invalidates existing COM references. Always call `switch_project` before operating on a different file.
- **Undo stack**: `undo_last` supports up to 10 consecutive undos. MS Project's COM undo is less reliable than the UI's.
- **File locking**: Only one process can hold the COM connection. Don't open MS Project's GUI dialogs while the server is active.
- **Timezone-aware dates**: COM may return timezone-aware datetimes. The server normalizes these via `_to_naive()` for safe comparisons.

## Tests

```bash
# Run all phases
python test_phase2.py   # 10 tests
python test_phase3.py   # 15 tests
python test_phase4.py   # 21 tests (+2 skipped)
python test_phase5.py   #  5 tests (bug fixes + new tools)
```

All tests require MS Project to be running (they create and close temporary projects).

## Architecture

Single-file server (`server.py`, ~3,900 lines) using the FastMCP framework. All COM calls go through `get_app()` / `get_proj()` helpers. Dates are normalized with `_to_naive()` and formatted with `_fmt_date()`.
