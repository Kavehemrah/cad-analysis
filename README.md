# CAD Drainage Analysis

Python + pywin32 tool for detecting `Drainage Pipe E` against railway sleepers in AutoCAD, calculating required transverse movement, optionally applying the move, and generating an Excel report.

## Current geometric model

Pipe diameter is detected from the two parallel `E Pipe` side lines in the `Drainage Pipe E` block. Sleeper dimensions are detected from the geometry of `AG_t` and `AG_tttt` block definitions. The pipe center is transformed into the sleeper coordinate system.

Required final center distance from the sleeper longitudinal axis is:

`half sleeper width + pipe radius + required clearance`

Movement is perpendicular to the pipe axis and is selected from the pipe-side relationship to the sleeper.

## Detection policy

The program scans ModelSpace once, indexes sleepers spatially, and evaluates only nearby candidates. A sleeper is accepted only when its geometry and axis satisfy the configured tolerances. Ambiguous or unmatched cases are reported and are never moved automatically.

## Usage

Open the target DWG in AutoCAD and run:

`python run_analysis.py`

This is analysis-only and creates an `.xlsx` report.

To apply calculated moves:

`python run_analysis.py --apply`

To apply and save the DWG:

`python run_analysis.py --apply --save`

For a controlled test:

`python run_analysis.py --limit 10`

Reports are written to the current directory unless `--report` is supplied.

## Dependencies

- Python 3.10+
- pywin32
- AutoCAD with the drawing open
- Microsoft Excel for `.xlsx` report generation

## Safety

The default mode does not move anything. `--apply` is required for real movement. `--save` is required to save the DWG. Applied movements are placed inside an AutoCAD undo mark, and movement evidence lines are written to layer `CAD_ANALYSIS_MOVE`.
