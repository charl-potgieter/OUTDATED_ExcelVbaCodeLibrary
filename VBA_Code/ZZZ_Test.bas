Attribute VB_Name = "ZZZ_Test"
Option Explicit


Sub TestDelete()

    Dim s As zLIB_ListStorage
    
    Set s = New zLIB_ListStorage
    
    s.CreateStorage ActiveWorkbook, "Test", Array("a", "b")
    s.Delete

End Sub
