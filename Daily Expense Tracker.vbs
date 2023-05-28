Do
    Dim result
    result = ""
    
    Dim totalSpending
    totalSpending = 0
    
    Dim userInput1
    Do
        userInput1 = InputBox("SPENT ON?", "EXPENSE NAME :", "Chai-S")
        
        ' Check if the first input is empty
        If Trim(userInput1) = "" Then
            MsgBox "PROVIDE A VALID EXPENSE NAME.", vbExclamation
        End If
    Loop Until Trim(userInput1) <> ""
    
    Dim userInput2
    Do
        userInput2 = InputBox("HOW MUCH?", "AMOUNT SPENT :")
        
        ' Validate the user input
        If Not IsNumeric(userInput2) Then
            MsgBox "INVALID INPUT! ENTER A NUMERIC VALUE.", vbExclamation
        End If
    Loop Until IsNumeric(userInput2)
    
    ' Convert the input to a numeric value
    Dim numericValue
    numericValue = CDbl(userInput2)
    
    Dim currentDate
    currentDate = Date
    
    Dim filePath
    filePath = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\")) & "Daily Expense Records.txt" ' File path in the same location as the script
    
    Dim fileSystem
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the file exists
    If Not fileSystem.FileExists(filePath) Then
        ' Create a new file named "Records.txt"
        Set file = fileSystem.CreateTextFile(filePath)
        file.WriteLine "-: THIS HAS ALL THE EXPENSES DATE WISE :-" ' Add the first line
        file.Close
    End If
    
    Dim fileContent
    Set file = fileSystem.OpenTextFile(filePath, 1, True)
    fileContent = file.ReadAll
    file.Close
    
    ' Check if today's date is already present in the file
    If InStr(fileContent, currentDate) = 0 Then
        ' Append today's date to the file
        Set file = fileSystem.OpenTextFile(filePath, 8, True)
        file.WriteLine vbNewLine & currentDate
        file.Close
    
        MsgBox "TODAY'S DATE ADDED!", vbInformation
    End If  ' End of the inner If block
    
    Set file = fileSystem.OpenTextFile(filePath, 8, True)
    file.WriteLine(userInput1 & "   -->   Rs. " & userInput2) ' Add "Rs." to the amount
    file.Close
    
    Set file = fileSystem.OpenTextFile(filePath, 1, True)
    Dim resultLine
    Dim isAfterToday
    isAfterToday = False
    
    Do Until file.AtEndOfStream
        resultLine = file.ReadLine
        
        ' Check if the line contains today's date or a date after today
        If IsDate(resultLine) Then
            If CDate(resultLine) >= currentDate Then
                isAfterToday = True
            Else
                isAfterToday = False
            End If
        End If
        
        ' Process the lines after today's date
        If isAfterToday Then
            ' Do something with the line
            result = result & resultLine & vbCrLf
            
            ' Extract the numeric value from the line and add it to the total spending
            Dim pos
            pos = InStr(resultLine, "-->")
            If pos > 0 Then
                Dim amount
                amount = Trim(Mid(resultLine, pos + 3))
                amount = Replace(amount, "Rs. ", "") ' Remove "Rs." from the amount
                If IsNumeric(amount) Then
                    totalSpending = totalSpending + CDbl(amount)
                End If
            End If
        End If
    Loop
    
    ' Close the file
    file.Close
    
    Dim answer
    answer = MsgBox("DO YOU HAVE ANOTHER TRANSACTION TO LOG?", vbQuestion + vbYesNo)
    
    If answer = vbNo Then
        Exit Do ' Exit the loop
    End If
Loop

MsgBox "TODAY'S SPENDING LIST :" & vbCrLf & vbCrLf & "-----------------------------------" & vbCrLf & "        " & result & "-----------------------------------", vbInformation

MsgBox "TODAY'S TOTAL SPENDING : " & vbCrLf & vbCrLf & "          Rs. " & totalSpending, vbInformation
