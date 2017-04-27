Attribute VB_Name = "Module1"
Sub ToMinutes()

Dim text As String
Dim totalHours() As String
Dim hours As String
Dim mins As String
Dim a As Long
Dim b As Long

Dim x As Long
Dim y As Long
Dim arrayLength As Integer

Dim RangeInput As Variant
Dim i As Integer
Dim tempString As String

'Get the total number of projects in the .csv file
RangeInput = InputBox("How many projects are in the file?")

'Convert the given string into an int so as to increment off the first row
'Then convert back to a string to concatenate onto the cellrange.
i = CInt(RangeInput)
i = i + 1
tempString = CStr(i)

'Declare the range of cells to work with
Dim projectRange As Range
Dim cellRange As String
cellRange = "F2:F" & tempString
Set projectRange = Range(cellRange)


'''' work out the total minutes ''''
'Make our new column say .. Total Minutes
Range("H1").Value = "Total Minutes"


' for each cell , split the text into an array
' e.g. the text 406 hrs 32 min becomes an array of {"406"}, {"hrs"}, {"32"}, {"min"},
    For x = 1 To projectRange.Rows.Count
        For y = 1 To projectRange.Columns.Count
        
            text = projectRange.Cells(x, y).Value
            totalHours = Split(text, " ")
            hours = totalHours(0)
            
            ' if the first cell in the array is "&mdash" simply output null in the cell
            If totalHours(0) = "&mdash;" Then
                projectRange.Cells(x, y + 2).Value = "Null"
                
            ' if the second cell in the array is "min" , then there is no hours in the cell
            ' and no need to multiply by 60 to get the minutes
            ElseIf totalHours(1) = "min" Then
                a = CLng(hours)
                projectRange.Cells(x, y + 2).Value = a
                
            ' if the condtions above do not apply, then convert the hours and minute strings into longs
            ' multiply the hours by 60 to get the minutes, then add the leftover minutes to get the total
            ' output the total minutes into the cell
            Else
                a = CLng(hours)
                a = a * 60
                
                arrayLength = Application.CountA(totalHours)
                    If arrayLength > 2 Then
                        mins = totalHours(2)
                        b = CLng(mins)
                        a = (a + b)
                    End If
                
                projectRange.Cells(x, y + 2).Value = a
            End If
        Next y
    Next x


'''' Work out the total minutes for each client ''''
' Reset our cell range
cellRange = "B2:B" & tempString
Set projectRange = Range(cellRange)

' declare our variables
Dim clientName As String
Dim tempValue As Long
Dim clientDict As Scripting.Dictionary

' Create a dictionary to store our unique values
Set clientDict = New Scripting.Dictionary

' loop through our cells of clients
    For x = 1 To projectRange.Rows.Count
        For y = 1 To projectRange.Columns.Count
                ' clientName = the text in the current cell
                clientName = projectRange.Cells(x, y).Value
                
                ' if the client already exists in the dictionary
                ' add the new minutes to the current total minutes
                If clientDict.Exists(clientName) Then
                    tempValue = clientDict(clientName)
                    ' if the value isnt null
                    If projectRange.Cells(x, y + 6).Value <> "Null" Then
                        tempValue = tempValue + projectRange.Cells(x, y + 6).Value
                        clientDict(clientName) = tempValue
                    End If
                ' else create a new key and value for that client
                Else
                    ' if the value isnt null
                    If projectRange.Cells(x, y + 6).Value <> "Null" Then
                        clientDict(clientName) = projectRange.Cells(x, y + 6).Value
                    End If
            End If
        Next y
    Next x
    
    'Make a new column say .. Clients and Total Minutes
        Range("I1").Value = "Clients"
        Range("J1").Value = "Total Minutes"
        Range("K1").Value = "Total Hours"
    ' loop through the dictionary and print out the keys and their values to the worksheets cells
    Dim c As Long
    For c = 0 To clientDict.Count - 1
        Debug.Print clientDict.Keys(c), clientDict.Items(c)
            Cells(2 + c, 9).Value = clientDict.Keys(c)
            Cells(2 + c, 10).Value = clientDict.Items(c)
    Next c
    
    
    ' convert the total minutes for each client into hours and minutes
    Dim hoursAndMinutes() As String
    Dim tempText As String
    Dim tempDouble As Double
    Dim tempCalculation As Double
    
    For c = 0 To clientDict.Count - 1
        tempCalculation = Cells(2 + c, 10).Value / 60
        tempText = CStr(tempCalculation)
        hoursAndMinutes = Split(tempText, ".")
        arrayLength = Application.CountA(hoursAndMinutes)
        If arrayLength > 1 Then
            tempDouble = CDbl("0." & hoursAndMinutes(1))
            tempDouble = tempDouble * 60
            Cells(2 + c, 11).Value = hoursAndMinutes(0) & " Hours " & CStr(tempDouble) & " Mins"
        Else
            Cells(2 + c, 11).Value = hoursAndMinutes(0) & " Hours"
        End If
    Next c
    
End Sub
