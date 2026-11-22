# Repository instructions

## Project type and entry points

- This repository contains Excel VBA source exported as `.bas` files; there is no application build system, package manifest, test framework, or automated lint configuration.
- `ActualizarBD_Optimizada.bas` is a standard-module implementation. Import it into the VBA project of the Excel workbook that owns the target `ListObject`.
- The public entry point is `ActualizarBD_Optimizada_Batch`. Run it from the Excel VBA Editor or by assigning it to a workbook macro button.
- The repository README only defines this as a location for VBA code worked on with GitHub Copilot.

## Build, validation, and testing

- There are no CLI build, test, or lint commands. Validate changes in Microsoft Excel by importing or refreshing the module in the VBE and selecting a cell in the target table before running `ActualizarBD_Optimizada_Batch`.
- Use Debug > Compile VBAProject in the Excel VBA Editor as the compile check for VBA syntax and unresolved project references.
- There is no test suite or single-test runner. Manual smoke tests should cover `Actualizar`, `Insertar`, and `Eliminar` rows plus a parameter-error case. Because the macro writes to a database and commits batches, test against a disposable copy of the workbook/database.
- `Option Explicit` is enabled. Keep declared and initialized variables consistent with VBA and ADODB types.
- The module calls `AbreConexion cn`, but no implementation of `AbreConexion` exists in this repository. An Excel VBA project using this module must provide that helper or it must be added to the module/project. Do not assume a connection string or provider is already defined here.

## Architecture

- `ActualizarBD_Optimizada_Batch` is the orchestration routine:
  1. Resolve the active table as `ActiveCell.ListObject` and validate required columns.
  2. Read the table body into a `Variant` array once, and inspect Excel's `AccionEnBD` values.
  3. Query `Bd_Ppto` metadata through ADODB to map Excel fields to database ADO types.
  4. Process `Actualizar`, `Insertar`, and `Eliminar` actions using parameterized ADODB commands.
  5. Commit in batches of `TAMANO_LOTE = 20`, update worksheet confirmation cells after each commit, then run follow-up SQL and the `Stp_2026_Actualiza_AnalizaForm_Azul` stored procedure.
- Database operations use ADODB with a transaction, parameter metadata, and explicit conversions. The module is not an Excel formula or generic table helper; it is tied to the `Bd_Ppto` Access table and the workbook's named columns.
- `CamposActualizarBD` is the canonical list of fields synchronized from Excel to `Bd_Ppto`. The list is assembled as a pipe-delimited string and split in VBA, preserving exact column names including accents and spaces.
- `RegistrarLog` appends errors to `ActualizarBD_Log.txt` beside `ThisWorkbook`.
- Batch state is kept in arrays and worksheet state is changed only after a database transaction is committed.

## VBA conventions

- Required Excel table columns are `AccionEnBD`, `Id`, `ID_ActualizarBD`, and `Fecha_Actualizacion`. Additions must update both `CamposActualizarBD` and its corresponding Excel column validation when a database field is added.
- Processable actions are compared case-insensitively and only include `Actualizar`, `Insertar`, and `Eliminar`. Confirmation values are Spanish and include `Actualizar-` and `Este renglón fue eliminado de la BD`.
- Column lookup is name-based through `ListObject.ListColumns`, not fixed worksheet indexes.
- Treat worksheet values as `Variant` data. Preserve the helpers that distinguish `Error`, `Null`, `Empty`, strings, and ADO numeric/string types before assigning ADODB parameters.
- Preserve dynamic parameter sizing in `AjustarTamParametro` and database-type conversion in `ConvertirValorParametro` when changing fields or parameter types.
- Keep SQL identifiers bracketed and escape closing brackets in field names. Use parameter placeholders rather than interpolating Excel values.
- During processing, preserve and restore `Application.Calculation`, `ScreenUpdating`, `EnableEvents`, and `StatusBar`; restore ADO objects and roll back an active transaction through the existing error path.
- The current batch behavior commits every 20 rows and then confirms them on the sheet. An error rollbacks only the active batch; already committed batches remain confirmed and are not rolled back.
