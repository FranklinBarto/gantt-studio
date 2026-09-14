# Gantt Studio — MSPDI Edition

This version uses the **Microsoft Office Project XML Data Interchange (MSPDI)** schema as its project interchange format.

Microsoft documents MSPDI as the XML interchange format for Microsoft Project. The schema is centered on a `<Project>` root with project properties, calendars, tasks, resources and assignments. This app exports those structures as `.xml` and can import MSPDI XML back into the editor.

## MSPDI implementation

The XML export includes:

- `<Project>` with `SaveVersion`, `GUID`, `UID`, `Name`, `Title`, `Subject`, `StartDate`, `FinishDate`, calendar settings and project defaults.
- `<Calendars>` with a Standard base calendar and Monday–Friday 08:00–12:00 / 13:00–17:00 working periods.
- `<Tasks>` with `UID`, `ID`, `Name`, `WBS`, `OutlineNumber`, `Start`, `Finish`, `Duration`, `DurationFormat`, `Work`, `CalendarUID`, `Notes`, `PercentComplete`, `Milestone`, `Summary` and `ExtendedAttribute`.
- `<Resources>` with an unassigned resource.
- `<Assignments>` linking each task to the unassigned resource.
- Gantt Studio theme metadata is stored in an MSPDI task `ExtendedAttribute` so the visual preset can survive an XML round-trip.

The task duration is represented as ISO 8601 XML duration data, using Microsoft Project's `DurationFormat=7` for working days.

## Import / export

Use **Open XML** to load an MSPDI project. Use **Export → Microsoft Project XML (MSPDI)** to create an XML project intended for Microsoft Project interoperability.

PNG, PDF, PowerPoint and Excel exports remain available as presentation formats.

## Run

```bash
npm install
npm run dev
```

Then:

```bash
npm run build
```

## Important interoperability note

MSPDI is a project interchange format, not a visual Gantt-chart styling standard. Therefore, scheduling information is represented in MSPDI fields, while the Gantt Studio visual theme remains application presentation metadata. The core project schedule is still expressed using MSPDI's project/task/calendar/resource/assignment model.


## Dependencies / task relationships

Each task can have a predecessor and an MSPDI task link type:

- **FS — Finish-to-Start**
- **FF — Finish-to-Finish**
- **SF — Start-to-Finish**
- **SS — Start-to-Start**

The application writes these as `<PredecessorLink>` elements under the dependent `<Task>`, using the MSPDI `Type` values documented by Microsoft: `0=FF`, `1=FS`, `2=SF`, `3=SS`. Optional lag is exported using `LinkLag` and `LagFormat`.

The editor includes a Dependency column where the predecessor, relationship type and lag can be changed without editing XML manually. Imported MSPDI predecessor links are read back into the editor.

## Text layout

The Initiative and Objective columns are now flexible and wrap long text instead of clipping it. The grid uses minimum column widths, expandable text fields and a dedicated Dependency column. Excel export uses wrapped cells and increased row height; PowerPoint uses shrink-to-fit text for dense slides.

### Initiative settings and relationships

The main Gantt grid intentionally keeps relationship controls out of the timeline. Use the **initiative settings** button on any initiative row to open a dedicated modal. The modal has:

- Initiative name and objective editing
- A Relationships tab
- Multiple predecessor relationships per initiative
- Finish-to-Finish (FF), Finish-to-Start (FS), Start-to-Finish (SF), and Start-to-Start (SS)
- Positive or negative lag in days
- Add/remove relationship controls

MSPDI export writes each relationship as a separate `PredecessorLink`, matching the Microsoft Project XML structure. Microsoft documents `PredecessorLink` as supporting multiple occurrences and the four link types. citeturn0search0turn0search4

### Adaptive timeline scale
The visible timeline automatically selects a useful unit from the selected timeframe:
- 0–3 days: hourly
- 4–21 days: daily
- 22–120 days: weekly
- More than 120 days: monthly

Task start and finish can be edited with date and time in Initiative Settings. The same date/time values are used for MSPDI XML export.
# gantt-studio
