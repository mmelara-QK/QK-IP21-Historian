
# Repository review: QK-IP21-Historian

## What the current VBA repository does well

The repository is cleanly separated into:
- `modAtHistory.bas` for direct worksheet-writing pulls and log sheet output,
- `modAtUDF.bas` for spillable UDFs and cache,
- `modAtTime.bas` for local/UTC conversion,
- `frmAtHistory.frm` plus `modAtUI.bas` for a usable desktop UI,
- `modRibbon.bas` for ribbon commands,
- `JsonConverter.bas` vendor dependency.

The strongest improvements already present in the repo are:
- batched history requests instead of one POST per tag,
- support for `f="D"` response mode,
- support for non-numeric values such as `OFF`,
- `"Invalid Tag"` marker behavior,
- reduced refresh blast radius by using sheet recalc instead of full rebuild.

## Desktop-only assumptions in the VBA code

The original plugin depends on:
- VBA / XLAM execution,
- `WinHttp.WinHttpRequest.5.1`,
- current Windows credentials on the business network,
- workbook/UI interactions that only exist in desktop Excel.

Those assumptions do not carry over directly to Excel on the web.

## What had to change for an online version

The online version uses:
- Office Add-in manifest,
- JavaScript custom functions,
- browser `fetch`,
- task pane UI instead of VBA UserForm,
- session cache in shared runtime memory,
- browser/network/CORS-compatible request flow.
