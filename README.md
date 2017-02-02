# ProspectBank
prospect bank for OM group project

Option Explicit
'clint maynard
'seperate membername into first and last name
'first name lower case
'last name all CAPS
'all other data transfered over as is

Sub namesplit()
   
    'move customer name data into column B
    
    Sheets("sunshine").Range("MemberName").Copy Destination:=Sheets("LoadData").Range("B1")
    
    Dim name As String
    Dim comma As Integer
    'dim i as long because of size
    Dim i As Long
    For i = 2 To Rows.Count
        name = Cells(i, 2).Value
        comma = InStr(name, ",")
        'overlap data transfered into column B originally to be first name
        Cells(i, 2).Value = Mid(name, comma + 2)
        Cells(i, 3).Value = Left(name, comma - 1)
        'skip blank name cells
        On Error Resume Next
        
    Next i
    
    'all other data coppied over and pasted in colum approiate column
    Sheets("sunshine").Range("AccountNumbers").Copy Destination:=Sheets("LoadData").Range("A1")
    Sheets("sunshine").Range("AccountBalance").Copy Destination:=Sheets("LoadData").Range("D1")
    Sheets("sunshine").Range("memberstate").Copy Destination:=Sheets("LoadData").Range("e1")
    Sheets("sunshine").Range("AccountType").Copy Destination:=Sheets("LoadData").Range("f1")
    Sheets("sunshine").Range("AccountPin").Copy Destination:=Sheets("LoadData").Range("g1")
    Sheets("sunshine").Range("MemberRace").Copy Destination:=Sheets("LoadData").Range("h1")
    Sheets("sunshine").Range("MaritalStatus").Copy Destination:=Sheets("LoadData").Range("i1")
    Sheets("sunshine").Range("HometownStatus").Copy Destination:=Sheets("LoadData").Range("j1")
    

End Sub



'lower case first name of all customer
'skip blanks
Sub lowercasefirst()
    Dim rng As Range
    Dim memfirstname As Range
    Set rng = ActiveSheet.Range("b2:b50045")
    For Each memfirstname In rng
        memfirstname.Value = LCase(memfirstname.Value)
    
        On Error Resume Next
    Next memfirstname


End Sub



'uppercase last name of all customers
'skip blank cells and continue to next

Sub uppercaselast()

    Dim rng As Range
    Dim memfirstname As Range
    'until to dynamic code bottom of list
    Set rng = ActiveSheet.Range("c2:c50045")
    For Each memfirstname In rng
        memfirstname.Value = UCase(memfirstname.Value)
    
        On Error Resume Next
    Next memfirstname
End Sub


Sub Add_Zeros_2()
    Dim a
    Dim i As Long
    
    With Range("A2", Range("A" & Rows.Count).End(xlUp))
       .NumberFormat = "@"
       a = .Value
       For i = 1 To UBound(a, 1)
        a(i, 1) = Right("00000000" & a(i, 1), 8)
       Next i
       .Value = a
    End With


    
End Sub

Sub Add_Zeros_3()
Dim r As Range
Set r = ActiveSheet("A2:A50045")
    r = function.Text(r, "00000000")
    
    
    
End Sub
    
