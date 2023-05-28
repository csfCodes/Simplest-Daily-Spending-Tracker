' Get the folder path for storing the files
Dim folderPath
folderPath = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\")) & "All Records\"

' Create the "All Records" folder if it doesn't exist
Dim fileSystem
Set fileSystem = CreateObject("Scripting.FileSystemObject")
If Not fileSystem.FolderExists(folderPath) Then
    fileSystem.CreateFolder(folderPath)
End If

Do
    ' Initialize variables
    Dim result
    result = ""
    
    Dim totalSpending
    totalSpending = 0
    
    ' Get the expense name from the user
    Dim userInput1
    Do
        userInput1 = InputBox("SPENT ON?", "EXPENSE NAME :", "Chai-S")
		userInput1 = UCase(userInput1)
        
        ' Check if the first input is empty
        If Trim(userInput1) = "" Then
            MsgBox "PROVIDE A VALID EXPENSE NAME.", vbExclamation
        End If
    Loop Until Trim(userInput1) <> ""
    
    ' Get the amount spent from the user
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
    
    ' Get the current date
    Dim currentDate
    currentDate = Date
    
    ' Set the file path for storing expense records
    Dim expenseFilePath
    expenseFilePath = folderPath & "Daily Expense Records.txt"
    
    ' Check if the file exists
    If Not fileSystem.FileExists(expenseFilePath) Then
        ' Create a new file named "Daily Expense Records.txt"
        Set file = fileSystem.CreateTextFile(expenseFilePath)
        file.WriteLine "-: THIS HAS ALL THE EXPENSES DATE WISE :-" ' Add the first line
        file.Close
    End If
    
    ' Read the file content
    Dim fileContent
    Set file = fileSystem.OpenTextFile(expenseFilePath, 1, True)
    fileContent = file.ReadAll
    file.Close
    
    ' Check if today's date is already present in the file
    If InStr(fileContent, currentDate) = 0 Then
        ' Append today's date to the file
        Set file = fileSystem.OpenTextFile(expenseFilePath, 8, True)
        file.WriteLine vbNewLine & currentDate
        file.Close
    End If  ' End of the inner If block
    
    ' Write the expense details to the file
    Set file = fileSystem.OpenTextFile(expenseFilePath, 8, True)
    file.WriteLine(userInput1 & "   -->   Rs. " & userInput2) ' Add "Rs." to the amount
    file.Close
    
    ' Read the file content again to extract spending details
    Set file = fileSystem.OpenTextFile(expenseFilePath, 1, True)
    Dim resultLine
    Dim isAfterToday
    isAfterToday = False
    
    ' Process the lines after today's date
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
            ' Append the line to the result (spending list)
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
    
    ' Set the file path for storing total spending
    Dim spendingFilePath
    spendingFilePath = folderPath & "Daily Total Spendings.txt"
    
    ' Create a new file named "Daily Total Spendings.txt" if it doesn't exist
    If Not fileSystem.FileExists(spendingFilePath) Then
        ' Create the file and add the initial line
        Set file = fileSystem.CreateTextFile(spendingFilePath)
        file.WriteLine "-: TOTAL SPENT RECORDS OF EACH DAY :-"
        file.Close
    End If
    
    ' Read the total spending file content
    Dim totalSpendingContent
    Set file = fileSystem.OpenTextFile(spendingFilePath, 1, True)
    totalSpendingContent = file.ReadAll
    file.Close
    
    ' Check if today's date is already present in the total spending file
    If InStr(totalSpendingContent, currentDate) = 0 Then
        ' Append today's date and total spending to the file
        Set file = fileSystem.OpenTextFile(spendingFilePath, 8, True)
        file.WriteLine(currentDate & "   -->   Rs. " & totalSpending)
        file.Close
    Else
        ' Replace the existing total spending for today's date with the latest total
        Dim regexPattern
        regexPattern = currentDate & ".*"
        
        Dim regex
        Set regex = New RegExp
        regex.Pattern = regexPattern
        regex.Global = False
        regex.IgnoreCase = True
        
        totalSpendingContent = regex.Replace(totalSpendingContent, currentDate & "   -->   Rs. " & totalSpending)
        
        ' Write the updated total spending content to the file
        Set file = fileSystem.OpenTextFile(spendingFilePath, 2, True)
        file.Write totalSpendingContent
        file.Close
    End If
    
    ' Ask the user if they have another transaction to log
    Dim answer
    answer = MsgBox("DO YOU HAVE ANOTHER TRANSACTION TO LOG?", vbQuestion + vbYesNo)
    
    If answer = vbNo Then
        Exit Do ' Exit the loop
    End If
Loop

' Display today's spending list
MsgBox "TODAY'S SPENDING LIST :" & vbCrLf & "-----------------------------------" & vbCrLf & "        " & result & "-----------------------------------" & vbCrLf & "TODAY'S TOTAL : Rs. " & totalSpending & vbCrLf & "-----------------------------------", vbInformation