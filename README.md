# ExportAssemblyParts

A SOLIDWORKS VBA macro that batch-exports every part and subassembly in an assembly to STEP, PNG, PDF and BOM CSV in one run.

Open an assembly, run the macro, choose what you want, and it writes:

| Output | For | File name |
| --- | --- | --- |
| STEP | each part, each assembly (always as a single file) | `<Model>-<Config>.step` |
| PNG | each model: isometric view, zoomed to fit | `<Model>-<Config>.png` |
| PDF | each model that has a drawing beside it, all sheets | `<Model>.pdf` |
| CSV | BOM tables from those drawings, chosen from a list at the end | `<Drawing>-<Sheet>.csv` |
| Log | one row per output file, for downstream scripts | `<Assembly>_export_log.csv` |

`-<Config>` is only added when the model has more than one configuration.

## Behaviour

- **All levels of the assembly.** This includes parts inside subassemblies, the subassemblies themselves and the top-level assembly.
- **Only the configurations the assembly uses.** Each model/configuration pair is exported once, however many instances there are. If you run the macro on a single part, it exports every configuration of that part.
- **Suppressed components are skipped,** and so is anything inside a suppressed subassembly. Lightweight components are resolved first.
- **Drawings** are found by name: `Bracket.sldprt` → `Bracket.slddrw` in the same folder. Each drawing is opened once, read-only, for both the PDF and the BOM tables.
- **Nothing is saved.** The macro switches each model's active configuration when it needs to and switches it back afterwards. Models and drawings are only ever written with *Save As Copy*.
- **One STEP per assembly.** For the run, the system option *Export assembly components as separate STEP files* is turned off, then restored at the end.
- **Errors don't stop the run.** If a model fails, the macro logs it and moves on to the next one.

### Where files go

Output goes next to the model file it came from, sorted into subfolders:

```
Parts\
  Bracket.sldprt
  Bracket.slddrw
  Export\
    STEP\Bracket.step
    PNG\Bracket.png
    PDF\Bracket.pdf
    CSV\Bracket-Sheet1.csv
```

If you turn subfolders off, files go straight into the model's folder. Virtual parts (stored inside the assembly) have no folder of their own, so their output goes next to the assembly. The log is always written next to the assembly you ran the macro on.

## Installation

The macro needs SOLIDWORKS 2026 (it was written against the 2026 API type libraries).

1. In SOLIDWORKS, choose **Tools → Macro → New…** and save the macro as `ExportAssemblyParts.swp`. The VBA editor opens.
2. Choose **File → Import File…** and import `ExportAssemblyParts.bas`. Then delete the default module that SOLIDWORKS created, so that only one `main` exists.
3. Create the settings form:
   - Choose **Insert → UserForm**, then in the Properties window set its **(Name)** to `frmSettings`.
   - Add two **CommandButtons** and name them `cmdOK` and `cmdCancel`. Their size, position and captions don't matter, because the macro sets them at run time.
   - Right-click the form, choose **View Code**, and replace the contents with [`forms/frmSettings.vba`](forms/frmSettings.vba).
4. Repeat step 3 for a second form named `frmPickBoms`, using [`forms/frmPickBoms.vba`](forms/frmPickBoms.vba).
5. Choose **Debug → Compile VBAProject** to check for errors, then save.

All the other controls on both forms are created by the macro when it runs. You don't need to add anything else in the form designer.

> If **File → Import** gives you one long line or garbled code, the `.bas` file has LF line endings. Clone the repository with Git (the `.gitattributes` file converts it to CRLF), or paste the code in directly.

## Usage

Open an assembly or part that has been saved, then choose **Tools → Macro → Run** and pick `ExportAssemblyParts.swp`. You can also assign the macro to a toolbar button under **Tools → Customize → Commands → Macro**.

A settings dialog opens at the start of each run:

| Option | Default |
| --- | --- |
| STEP file for each part | on |
| STEP file for each assembly (always one file) | on |
|   Top-level assembly only, not every subassembly | off |
| Isometric PNG for each model | on |
| PDF of the drawing, where one exists | on |
| BOM tables from drawings, as CSV | on |
| Include subassemblies | on |
| Include the top-level assembly itself | on |
| Add the configuration to file names | on |
| Sort output into `Export\STEP`, `\PNG`, `\PDF` | on |
| Export folder name | `Export` |
| Stop on the offending line instead of logging it | off |

If BOM tables were found, a list appears when the export finishes. Tick the tables you want written as CSV. None are ticked to start with.

A summary at the end shows how many files were exported, how many problems there were, how long the run took, and where the log is.

### Changing the defaults

The defaults are the `DEF_*` constants under **Settings** at the top of `ExportAssemblyParts.bas`. If you set `SHOW_SETTINGS_DIALOG` to `False`, the dialog is skipped and every run uses those constants.

## Export log

`<Assembly>_export_log.csv` has one row per output file, plus rows for skipped or failed items:

| Column | Meaning |
| --- | --- |
| `timestamp` | when the row was written |
| `source_type` | `part` or `assembly` |
| `source_name` | source file name: the model for STEP/PNG rows, the drawing for PDF/CSV rows |
| `configuration` | the configuration exported |
| `source_path` | folder holding the source file |
| `output_type` | `step`, `png`, `pdf` or `csv` |
| `output_path` | the output file, relative to `source_path` |
| `status` | `ok`, `failed`, `skipped` or `error` |
| `message` | detail, mostly for failures |

## Troubleshooting

- Check the log first. Each failure records what the macro was doing at the time.
- While the VBA editor is open, every log row is also printed to the Immediate window (**View → Immediate Window**).
- To debug a failure, turn on **Stop on the offending line** in the settings dialog. The macro then breaks into the debugger where the error happens instead of logging it and moving on.
- If a PDF or BOM is missing, look for a `skipped` row. It tells you whether there was no drawing named like the model, or the drawing had no BOM table.
