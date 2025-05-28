# 📊 Tally File Automation in Excel

This Excel macro-enabled solution automates the filtering, updating, and summarizing of Tally data across multiple sheets. It combines the power of **VBA macros**, **Excel dynamic arrays**, and **XLOOKUP-based formulas** to streamline data processing for Box and KG tallies.

---

## 🚀 Features

- 🔁 Automated update flow triggered via `Ctrl + Q`
- ✅ Input data validation and alert prompts
- 🔍 Dynamic lookups with `XLOOKUP` for fast reference matching
- 📊 Real-time metrics with summary calculations
- 📎 Easy integration with structured Tally exports

---

## 📁 Workbook Structure

| Sheet Name       | Purpose                                      |
|------------------|----------------------------------------------|
| `TALLY BOX DUMP` | Raw tally data (box units)                   |
| `TALLY KG DUMP`  | Raw tally data (weight in KG)                |
| `OVERALL ACH`    | Output summary with achievement metrics      |
| `CODE`           | Validation, status messages, SKU mappings    |
| `ACH CALC`       | Helper for formulas like SUMIF, FILTER, etc. |

---

## ⚙️ Main Macro: `Summarize`

### Triggered by: `Ctrl + Q`

### Workflow:

1. Validates if all conditions are met in the `CODE` sheet
2. If valid:
   - Clears & recalculates columns in `TALLY BOX DUMP` and `TALLY KG DUMP`
   - Uses `XLOOKUP` to fetch matching SKU data from `CODE`
   - AutoFills, copies, and pastes values to overwrite old data
3. If invalid:
   - Displays user prompts with specific instructions
   - Ends by requesting SKU updates if missing

---

## Macros Overview

```vba
Sub Summarize()
    Dim kgtally As Worksheet
    Dim val1 As String
    Set kgtally = ThisWorkbook.Sheets("TALLY KG DUMP")
    Dim boxtally As Worksheet
    Dim B As String
    Set boxtally = ThisWorkbook.Sheets("TALLY BOX DUMP")
    Dim code As Worksheet
    Dim answer As Integer
    Dim val2 As String
    Set code = ThisWorkbook.Sheets("CODE")
    Dim overall As Worksheet
    Dim A As String
    Set overall = ThisWorkbook.Sheets("OVERALL ACH")
    Dim C As Integer
    Dim D As Integer
        If code.Range("CB2") = "NO DATA" Then
            If code.Range("CF2") = 0 Then
            MsgBox code.Range("CF1")
            
            ElseIf code.Range("CG2") = 0 Then
            MsgBox code.Range("CG1")
            
            ElseIf code.Range("CH2") < 2 Then
            MsgBox code.Range("CH1")
            
            ElseIf code.Range("CI2") = 1 Then
            MsgBox code.Range("CI1")

            Else
            Application.ScreenUpdating = True
            E = code.Range("CH5").Value
            D = code.Range("CG5").Value
            kgtally.Select
            Range("A1").Select
            Cells(2, E).Clear
            Cells(2, E + 1).Clear

                code.Select
                val1 = code.Range("CF11").Value
                
                A = val1
                val2 = code.Range("CG11").Value
                B = val2
                Application.ScreenUpdating = False
                boxtally.Select
                boxtally.Rows("1:1").Select
                Selection.UnMerge
                boxtally.Range(A).Select
                ActiveCell.FormulaR1C1 = "=XLOOKUP(R1C,CODE!C76,CODE!C77,"""")"
                boxtally.Range("E10000").Select
                ActiveCell.FormulaR1C1 = "=XLOOKUP(R1C[-1],CODE!C78,CODE!C79,"""")"
                boxtally.Range(A & ":E10000").Select
                Selection.AutoFill Destination:=Range(A & ":" & B), Type:= _
                    xlFillDefault
                
                boxtally.Range(A & ":" & B).Calculate
                Application.Wait (Now() + TimeValue("00:00:01"))
                boxtally.Range(A & ":" & B).Select
                
                Application.CutCopyMode = False
                Selection.Copy
                boxtally.Range("D1").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                boxtally.Range(A & ":" & B).Select
                Selection.Clear
                boxtally.Range("A1").Select
                
                kgtally.Select
                kgtally.Rows("1:1").Select
                Selection.UnMerge
                kgtally.Range(A).Select
                ActiveCell.FormulaR1C1 = "=XLOOKUP(R1C,CODE!C76,CODE!C77,"""")"
                kgtally.Range("E10000").Select
                ActiveCell.FormulaR1C1 = "=XLOOKUP(R1C[-1],CODE!C78,CODE!C79,"""")"
                kgtally.Range(A & ":E10000").Calculate
                Application.Wait (Now() + TimeValue("00:00:01"))
                kgtally.Range(A & ":E10000").Select
                Selection.AutoFill Destination:=Range(A & ":" & B), Type:= _
                    xlFillDefault
                    
                kgtally.Range(A & ":" & B).Select
                Application.CutCopyMode = False
                Selection.Copy
                kgtally.Range("D1").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                kgtally.Range(A & ":" & B).Select
                Selection.Clear
                kgtally.Range("A1").Select
                code.Select
                code.Range("A1").Select
                Application.ScreenUpdating = True
                
                overall.Select
                overall.Range("B1").Select
                MsgBox "All are Updated"
                End If
        Else
            code.Select
            If code.Range("CB3") = "" Then
            code.Range("CB2").Copy
            Else
            code.Range("CB2").Select
            Application.CutCopyMode = False
            code.Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            End If
        code.Range("A1").Select
        Application.Goto Reference:="R10000C1"
        Selection.End(xlUp).Select
        ActiveCell.Offset(1, 0).Range("A1").PasteSpecial xlPasteValues
        ActiveCell.Offset(0, 1).Range("A1").Select
        MsgBox "Please Update SKU Code"
        End If
End Sub
```

---

## 📐 Key Excel Formulas Used

| Formula | Description |
|--------|-------------|
| `=OFFSET('TALLY BOX DUMP'!$A$1,1,0,COUNTA('TALLY KG DUMP'!A:A)-1,3)` | Dynamically expands range based on number of KG entries |
| `=IF($D4="","",(SUMIF('TALLY KG DUMP'!$2:$2,"Qty",'TALLY KG DUMP'!3:3)/1000)-(SUMIF('TALLY KG DUMP'!$1:$1,"GL",'TALLY KG DUMP'!3:3)/1000))` | Calculates net weight |
| `=IF($D4="","",SUMIF('TALLY BOX DUMP'!$1:$1,'ACH CALC'!M$2,'TALLY BOX DUMP'!3:3)+AD4)` | Adds prior values with current SUMIF |
| `=FILTER($A1:$B600,$C1:$C600="Qty")` | Filters rows where column C has "Qty" |
| `=SORT(UNIQUE(FILTER($CC$2:$CC$2000,$CD$2:$CD$2000="NOT ADDED","NO DATA")))` | Extracts unsynced SKUs |
| `=IFS(CC2="","",COUNTIF(A:A,CC2)>0,"ADDED",COUNTIF(B:B,CC2)>0,"ADDED",COUNTIF(A:A,CC2)=0,"NOT ADDED")` | Checks status of SKUs |
| `=UNIQUE(TRANSPOSE('TALLY BOX DUMP'!D1:IN1))` | Lists unique headers horizontally |
| `=((COUNTA('TALLY KG DUMP'!$A:$A)>0)*1)+((COUNTA('TALLY BOX DUMP'!$A:$A)>0)*1)` | Simple check for data presence |
| `=COLUMN(XLOOKUP($CH$4,'TALLY BOX DUMP'!1:1,'TALLY BOX DUMP'!1:1))` | Gets column index from header match |

---

## 📸 Screenshots

Include:
- Code screen
- Before/after of `TALLY BOX DUMP`
- Messages triggered by macro

---

## ✅ How to Use

1. Download the `.xlsm` file.
2. Open in **Microsoft Excel** with **Macros Enabled**.
3. Press `Ctrl + Q` to run the macro.
4. Follow on-screen prompts and update SKU codes if required.

---

## 🛠 Built With

- **VBA (Excel Macros)**
- **Excel Formulas** (`XLOOKUP`, `FILTER`, `OFFSET`, `SUMIF`, `IFS`, `UNIQUE`)
- **Excel Dynamic Arrays (Office 365 or newer)**

---

## 📂 Files

- `Tally Shorter for Box & KG 2.0.xlsm` – main macro-enabled workbook

---

## 👨‍💻 Author

- Sampath_SK

