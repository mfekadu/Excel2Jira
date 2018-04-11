

' initial inspiration : https://gist.github.com/DBremen/0ba67c6ec894ee581d98
Sub ConvertToJiraTable()
    Dim strKey As Variant


    ' declare an array of strings for BOLD keywords

    Dim boldArray(20) As String
    boldArray(0) = "Click"
    boldArray(1) = "Navigate"
    boldArray(2) = "Open"
    boldArray(3) = "Save"
    boldArray(4) = "Load"
    boldArray(5) = "Delete"
    boldArray(6) = "Hello"
    boldArray(7) = "Select"
    boldArray(8) = "Fill"
    boldArray(9) = "Enter"
    boldArray(10) = "Drag"
    boldArray(11) = "Type"

    Dim undlnArray(20) As String
    undlnArray(0) = "Asset"


    Dim sh As Worksheet
    Dim rw As Range
    Dim RowCount As Integer

    RowCount = 0

    Set sh = ActiveSheet

    ' Add Bold Markup
    For Each strKey In boldArray
        sh.Cells.Replace what:=strKey, Replacement:="*" & strKey & "*", _
            LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
            SearchFormat:=False, ReplaceFormat:=False
    Next strKey
    
    ' Add Underline Markup
    For Each strKey In undlnArray
        sh.Cells.Replace what:=strKey, Replacement:="+" & strKey & "+", _
            LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
            SearchFormat:=False, ReplaceFormat:=False
    Next strKey

End Sub
