Attribute VB_Name = "ReportBookCstr"
'@Folder "GoodsCollectorProject.src.Core.ReportBook"
Option Explicit

Public Function NewReportBook(ByVal Path As String) As ReportBook
    Set NewReportBook = New ReportBook
    NewReportBook.Path = Path
End Function
