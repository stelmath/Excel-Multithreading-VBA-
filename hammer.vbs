Option Explicit
On Error Resume Next

Dim ExcelObject, GetMessagePos, x, y, Count, Position 
set ExcelObject = getObject( ,"Excel.Application")

Do While ExcelObject.Worksheets("Sheet1").Range("A1").Value <> "Stopped"

GetMessagePos = ExcelObject.ExecuteExcel4Macro("CALL(""user32"",""GetMessagePos"",""J"")")

x = CLng("&H" & Right(Hex(GetMessagePos), 4))
y = CLng("&H" & Left(Hex(GetMessagePos), (Len(Hex(GetMessagePos)) - 4)))

ExcelObject.Worksheets("Sheet1").Shapes("hammer").Left = 0.75 * (x - 79) ' 0.75 
ExcelObject.Worksheets("Sheet1").Shapes("hammer").Top = 0.75 * (y - 277)
Loop
