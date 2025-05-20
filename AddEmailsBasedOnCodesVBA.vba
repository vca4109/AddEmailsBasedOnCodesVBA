Sub AddEmailsBasedOnCodes()
    Dim wsITR As Worksheet
    Dim wsEmail As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim emailDict As Object
    Dim code As Variant
    Dim emailList As String
    Dim uniqueEmailDict As Object
    Dim emailArr As Variant
    Dim email As Variant
    Dim uniqueEmailList As String

    ' Set worksheets
    Set wsITR = ThisWorkbook.Sheets("ITR")
    Set wsEmail = ThisWorkbook.Sheets("EMAIL")
    
    ' Clear previous content in A1 and A2 on the EMAIL sheet
    wsEmail.Range("A1:A2").ClearContents
    
    ' Dictionary to store codes and their corresponding emails
	Set emailDict = CreateObject("Scripting.Dictionary")
	emailDict.Add "AB", "test1.email@example.com"
	emailDict.Add "QR", "user2.fake@example.com; someone3.mail@demo.com; contact4.user@sample.net; test5.mail@fakemail.org; name6.demo@demo.com; user7.test@tryme.com"
	emailDict.Add "DD", "engineer8@sample.org"
	emailDict.Add "QA", "qa9.team@checkmail.com"
	emailDict.Add "ME", "mep10.support@example.net"
	emailDict.Add "DO", "doc11.control@archive.org"
	emailDict.Add "EN", "eng12.services@projmail.com"
	emailDict.Add "HR", "hr13.recruitment@workplace.net"
	emailDict.Add "PM", "pm14.management@projecthub.org"
	emailDict.Add "FI", "finance15.billing@corpemail.net"
        
    
    ' Initialize email list
    emailList = ""
    
    ' Get the last row in the ITR sheet for column B
    lastRow = wsITR.Cells(wsITR.Rows.Count, "B").End(xlUp).Row
    
    ' Loop through each cell in column B
    For Each cell In wsITR.Range("B2:B" & lastRow)
        ' Loop through each code in the dictionary
        For Each code In emailDict.Keys
            If InStr(cell.Value, code) > 0 Then
                ' If the code is found, append the emails if not already added
                If InStr(emailList, emailDict(code)) = 0 Then
                    If emailList = "" Then
                        emailList = emailDict(code)
                    Else
                        emailList = emailList & "; " & emailDict(code)
                    End If
                End If
            End If
        Next code
    Next cell
    
    ' Output the entire email list in A1 on the EMAIL sheet
    wsEmail.Range("A1").Value = emailList
    
    ' Remove duplicates from the email list in A1
    Set uniqueEmailDict = CreateObject("Scripting.Dictionary")
    emailArr = Split(emailList, "; ")
    
    ' Add each email to the dictionary to filter out duplicates
    For Each email In emailArr
        If Not uniqueEmailDict.exists(Trim(email)) Then
            uniqueEmailDict.Add Trim(email), Nothing
        End If
    Next email
    
    ' Combine unique emails into a single string
    uniqueEmailList = Join(uniqueEmailDict.Keys, "; ")
    
    ' Output the unique email list in A2 on the EMAIL sheet
    wsEmail.Range("A2").Value = uniqueEmailList
End Sub

