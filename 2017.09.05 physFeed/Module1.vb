'20100105 - ITWNewFeed.exe
Imports System
Imports System.IO
Imports System.Collections
Imports System.data.sqlclient

'20130103 - changed ITW connection string. Old: 10.48.10.246,1433.  New: 10.48.242.249,1433.
'20140130 - changed Mcare connections for db conversion.

'20150708 - mods to write errors to text file in preps for move to ITWListen.
Module Module1
    '7/14/2010 Pharmacy feed program from ITW feed shell
    Dim McareConnectionString As String = ""
    Dim ITWConnectionString As String = ""
    Dim ShelbyConnectionString As String = ""

    Dim dictNVP As New Hashtable
    Dim sql As String
    Dim sql2 As String
    Dim myfile As StreamReader
    Dim gblLogString As String = ""
    Dim functionError As Boolean = False
    Dim dbError As Boolean = False
    Dim globalError As Boolean = False
    Dim orphanfound As Boolean = False
    Dim phoneArray() As String
    Dim phoneItemArray() As String

    Dim addressArray() As String
    Dim addressItemArray() As String
    Dim docIDArray() As String
    Dim docIDItemArray() As String

    Dim specialtyArray() As String
    Dim specialtyItemArray() As String
    Dim staffIDCodeArray() As String
    Dim staffIDCodeArrayItem() As String

    'Public objIniFile As New iniFile("C:\KY2 Test Environment\HL7Mapper.ini") 'Local
    'Public objIniFile As New iniFile("C:\ULHTest\HL7Mapper.ini") 'Test
    Public objIniFile As New iniFile("C:\newfeeds\HL7Mapper.ini") '20140805 Prod

    'Public conIniFile As New iniFile("C:\KY2 Test Environment\KY2ConnDev.ini") 'Local
    'Public conIniFile As New iniFile("C:\ULHTest\KY2ConnTest.ini") 'Test
    Public conIniFile As New iniFile("C:\newfeeds\KY2ConnProd.ini") '20140805 Prod

    Dim strOutputDirectory As String = ""
    Dim strMapperFile As String = ""
    Dim dir As String
    Public thefile As FileInfo


    'global database fields
    Dim gblPhysnum As Integer = 0
    Dim gblFirstName As String = ""
    Dim gblMiddle As String = ""
    Dim gblLastName As String = ""
    Dim gblLicNo As String = ""
    Dim gblPager As String = ""

    Dim gblPhone As String = ""
    Dim gblPhone_2 As String = ""
    Dim gblPhone_3 As String = ""

    Dim gblMobile As String = ""

    Dim gblFax As String = ""
    Dim gblFax_2 As String = ""
    Dim gblFax_3 As String = ""

    Dim GblHome As String = ""

    Dim gblAddr1 As String = ""
    Dim gblAddr2 As String = ""
    Dim gblAddr1_2 As String = ""
    Dim gblAddr2_2 As String = ""
    Dim gblAddr1_3 As String = ""
    Dim gblAddr2_3 As String = ""

    Dim gblSpeciality As String = ""
    Dim gblDEANum As String = ""
    Dim gblBlueCrossNum As String = ""
    Dim gblMedicareNum As String = ""
    Dim gblNPINum As String = ""
    Dim gblInactive As Boolean = False
    Dim gblNotes As String = ""
    '20110524 - added global upin number and status flag
    Dim gblUPIN As String = ""
    Dim gblStatus As String = ""
    Dim strInputDirectory As String = ""
    Dim gblID As Integer = 0
    '20111005
    Dim gblCurrentStatus As String = ""
    Dim gblStatusCategory As String = ""
	Dim strLogDirectory As String = "" '20150708



    Sub Main()
        
        'declarations for split function
        Dim delimStr As String = "="
        Dim delimiter As Char() = delimStr.ToCharArray()

        'declarations for stream reader
        Dim strLine As String
        'Dim sql As String = ""
        'Dim datareader As SqlDataReader
        Try
            strOutputDirectory = objIniFile.GetString("Physicians", "Physiciansoutputdirectory", "(none)") 'c:\feeds\ltw\Physicians\
            strMapperFile = objIniFile.GetString("Physicians", "Physiciansmapper", "(none)") 'c:\newfeeds\map\Physicians.txt
			strLogDirectory = objIniFile.GetString("Settings", "logs", "(none)") '20150708

            'setup directory
            Dim dirs As String() = Directory.GetFiles(strOutputDirectory, "LTW.*")
            
            'declarations and external assignments for local database operations
            'McareConnectionString = "server=(local);database=pHYSICIANS;trusted_connection=true"
            'ITWConnectionString = "server=(local);database=pHYSICIANS;trusted_connection=true"
            'ShelbyConnectionString = "server=(local);database=pHYSICIANS;trusted_connection=true"

            'connectionString = "server=10.48.10.246,1433;database=ITWtest;uid=sysmax;pwd=Condor!"
            'connectionString = "server=10.48.10.246,1433;database=ITW;uid=sysmax;pwd=Condor!"

            '20110511 - test connections at JHSMH
            'McareConnectionString = "server=10.48.242.249,1433;database=sqlmcare;uid=sysmax;pwd=Condor!"
            'ITWConnectionString = "server=10.48.242.249,1433;database=itw;uid=sysmax;pwd=Condor!"
            'ShelbyConnectionString = "server=10.48.242.249,1433;database=mcareShelby;uid=sysmax;pwd=Condor!"

            McareConnectionString = conIniFile.GetString("Strings", "PHYSMCARE", "(none)")
            ITWConnectionString = conIniFile.GetString("Strings", "PHYSITW", "(none)")
            ShelbyConnectionString = conIniFile.GetString("Strings", "PHYSHELBY", "(none)")

            'Dim myConnection As New SqlConnection(connectionString)
            'Dim objCommand As New SqlCommand
            'Dim updatecommand As New SqlCommand
            'updatecommand.Connection = myConnection
            'objCommand.Connection = myConnection

            For Each dir In dirs
                theFile = New FileInfo(dir)
                If theFile.Extension <> ".$#$" Then
                    '1.set up the streamreader to get a file
                    myfile = File.OpenText(dir)
                    'and read the first line
                    'strLine = myfile.ReadLine()

                    '20100119 - Catch a problem if the LTW file is messes up
                    Try
                        'Do While Not strLine Is Nothing
                        Do While Not myfile.EndOfStream
                            Dim myArray As String() = Nothing
                            strLine = myfile.ReadLine()
                            If strLine <> "" Then
                                myArray = strLine.Split(delimiter, 2)
                                'add array key and item to hashtable
                                Try
                                    dictNVP.Add(myArray(0), myArray(1))
                                Catch
                                End Try
                            End If
                        Loop
                    Catch ex As Exception
                        'make copy in the problems directory delete any previous ones with same name
                        Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "problems\" & theFile.Name)
                        fi2.Delete()
                        theFile.CopyTo(strOutputDirectory & "problems\" & theFile.Name)

                        gblLogString = gblLogString & "Dictionary Error" & " - " & theFile.Name & vbCrLf
                        gblLogString = gblLogString & ex.Message & vbCrLf
                        writeToLog(gblLogString, 1)
                        'get rid of the file so it doesn't mess up the next run.
                        myfile.Close()
                        If theFile.Exists Then
                            theFile.Delete()
                            Exit Sub
                        End If
                    End Try
                    '20100119 - Catch a problem if the LTW file is messes up

                    myfile.Close()

                    
                    'gblID = 0
                    '20101004 capture the epnum
                    'sql = "select epnum as epnum  from [001episode] where panum = '" & gblORIGINAL_PA_NUMBER & "'"
                    'objCommand.CommandText = sql
                    'myConnection.Open()
                    ''datareader = objCommand.ExecuteReader()
                    'While datareader.Read()
                    'gblEPNum = datareader.GetInt32(0)
                    'End While
                    'myConnection.Close()
                    'datareader.Close()
                    '===================================================================================================
                    'call subdirectories here
                    globalError = False
                    Call initializeVariables()
                    Call processStaffID(dictNVP)

                    If gblPhysnum <> 0 Then
                        Call processName(dictNVP)
                        '20111005:
                        Call processCurrentStatus(dictNVP)

                        Call processAddress(dictNVP)
                        Call processPhone(dictNVP)
                        Call processDocID(dictNVP)
                        Call processSpecialty(dictNVP)
                        Call updateMcare(dictNVP)
                        Call updateITW(dictNVP)
                        Call updateShelby(dictNVP)
                    End If

                    '===================================================================================================
                    dictNVP.Clear()
                    If functionError Then

                        gblLogString = "Function Error - " & thefile.Name & vbCrLf & gblLogString
                        writeToLog(gblLogString, 2)
                        gblLogString = ""

                        Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "backup\" & thefile.Name)
                        fi2.Delete()
                        thefile.CopyTo(strOutputDirectory & "backup\" & thefile.Name)
                        thefile.Delete()

                    ElseIf dbError Then

                        gblLogString = "dbError Error - " & thefile.Name & vbCrLf & gblLogString
                        writeToLog(gblLogString, 2)
                        gblLogString = ""

                        Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "reprocess\" & thefile.Name)
                        fi2.Delete()
                        thefile.CopyTo(strOutputDirectory & "reprocess\" & thefile.Name)
                        thefile.Delete()

                    ElseIf orphanfound Then


                        writeToLog(gblLogString, 3)
                        gblLogString = ""

                        Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "orphans\" & thefile.Name)
                        fi2.Delete()
                        thefile.CopyTo(strOutputDirectory & "orphans\" & thefile.Name)
                        thefile.Delete()
                    Else

                        thefile.Delete()

                    End If



                End If 'If theFile.Extension <> ".$#$"
            Next
            
        Catch ex As Exception
            globalError = True

            gblLogString = gblLogString & "Main Routine Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            writeToLog(gblLogString, 1)
            gblLogString = ""
            '20091117 - get rid of the problem file if it exists
            If thefile.Exists Then
                thefile.Delete()
            End If

            Exit Sub
        End Try
    End Sub



    Public Sub insertString(ByVal theString As String)
        Try
            If theString <> "" Then
                sql = sql & "'" & Replace(theString, "'", "''") & "', "
            Else
                sql = sql & "NULL, "
            End If
        Finally
        End Try
    End Sub

    Public Sub insertNumber(ByVal theString As String)
        Try
            If IsNumeric(theString) Then
                sql = sql & theString & ", "
            Else
                sql = sql & "NULL, "
            End If
        Finally
        End Try
    End Sub

    Public Function ConvertDate(ByVal datedata As String) As String
        'convert the hl7 date to a database date
        'hl7 in format: yyyymmdd or yyyymmddhhmm
        'returns now if the string is not in one
        'of the two formats.
        '
        Try
            Dim strYear As String = ""
            Dim strMonth As String = ""
            Dim strDay As String = ""
            Dim strHour As String = ""
            Dim strMinute As String = ""
            Dim strSecond As String = ""

            If Len(Trim(datedata)) = 8 Then
                strYear = Mid$(datedata, 1, 4)
                strMonth = Mid$(datedata, 5, 2)
                strDay = Mid$(datedata, 7, 2)
                ConvertDate = strMonth & "/" & strDay & "/" & strYear

            ElseIf Len(Trim(datedata)) >= 12 Then
                strYear = Mid$(datedata, 1, 4)
                strMonth = Mid$(datedata, 5, 2)
                strDay = Mid$(datedata, 7, 2)
                strHour = Mid$(datedata, 9, 2)
                strMinute = Mid$(datedata, 11, 2)
                strSecond = Mid$(datedata, 13, 2)

                'If strHour = "24" Then
                'ConvertDate = strMonth & "/" & strDay & "/" & strYear
                'Else
                ConvertDate = strMonth & "/" & strDay & "/" & strYear & " " & strHour & ":" & strMinute & ":" & strSecond
                'End If


            Else
                ConvertDate = DateTime.Now

            End If
        Finally
        End Try

    End Function

    Public Function ConvertSupDate(ByVal datedata As String) As String
        'convert the hl7 date to a database date
        'hl7 in format: yyyymmdd or yyyymmddhhmm
        'returns now if the string is not in one
        'of the two formats.
        '
        Try
            Dim strYear As String = ""
            Dim strMonth As String = ""
            Dim strDay As String = ""
            Dim strHour As String = ""
            Dim strMinute As String = ""

            If Len(Trim(datedata)) = 8 Then
                'strYear = Left$(datedata, 4)
                'strMonth = Mid$(datedata, 5, 2)
                'strDay = Mid$(datedata, 7, 2)
                strYear = Mid$(datedata, 5, 4)
                strMonth = Left$(datedata, 2)
                strDay = Mid$(datedata, 3, 2)
                ConvertSupDate = strMonth & "/" & strDay & "/" & strYear



            Else
                ConvertSupDate = DateTime.Now

            End If

        Finally
        End Try

    End Function

    Public Function ConvertSOS(ByVal data As String) As String
        'converts a sos number without delimiters to:
        ' sss-ss-ssss
        Try
            If Len(data) = 9 Then
                ConvertSOS = Mid$(data, 1, 3) & "-" & Mid$(data, 4, 2) & "-" & Mid$(data, 6, 4)
            Else
                ConvertSOS = ""
            End If
        Finally
        End Try
    End Function
    
	Public Sub writeTolog(ByVal strMsg As String, ByVal eventType As Integer)
        '20150708 - use a text file to log errors instead of the event log
        Dim file As System.IO.StreamWriter
        Dim tempLogFileName As String = strLogDirectory & "Physicians_log.txt"
        file = My.Computer.FileSystem.OpenTextFileWriter(tempLogFileName, True)
        file.WriteLine(DateTime.Now & " : " & strMsg)
        file.Close()
    End Sub
    
    

    
    
    Public Sub writeToLog2(ByVal logText As String, ByVal eventType As Integer)
        Dim myLog As New EventLog()
        Try
            ' check for the existence of the log that the user wants to create.
            ' Create the source, if it does not already exist.
            If Not EventLog.SourceExists("physicians") Then
                EventLog.CreateEventSource("physicians", "physicians")
            End If

            ' Create an EventLog instance and assign its source.

            myLog.Source = "physicians"

            ' Write an informational entry to the event log.
            If eventType = 1 Then
                myLog.WriteEntry(logText, EventLogEntryType.Error, 1)
            ElseIf eventType = 2 Then
                myLog.WriteEntry(logText, EventLogEntryType.Warning, 2)
            ElseIf eventType = 3 Then
                myLog.WriteEntry(logText, EventLogEntryType.Information, 3)
            End If


        Finally
            myLog.Close()
        End Try
    End Sub
    Public Sub initializeVariables()
        gblPhysnum = 0
        gblFirstName = ""
        gblMiddle = ""
        gblLastName = ""
        gblLicNo = ""
        gblPager = ""
        gblPhone = ""
        gblMobile = ""
        gblFax = ""
        GblHome = ""
        gblAddr1 = ""
        gblAddr2 = ""
        gblSpeciality = ""
        gblDEANum = ""
        gblBlueCrossNum = ""
        gblMedicareNum = ""
        gblInactive = False
        gblNotes = "Additional Information:" & vbCrLf

        
    End Sub
    Public Sub processAddress(ByVal dictNVP As Hashtable)
        Dim strAddress As String = ""
        Dim strAddressItem As String = ""
        Dim i As Integer = 0
        Dim tempstr As String = ""
        Dim intMRNUM As Long = 0
        Dim addit As Boolean
        Dim updateit As Boolean
        addit = False
        updateit = False

        Dim j As Integer = 0
        Try
            strAddress = dictNVP.Item("Office/Home Address")
            If strAddress <> "" Then ' 20111101
                addressArray = Split(strAddress, "~")
                For j = 0 To UBound(addressArray)
                    'txtOutput.Text = txtOutput.Text & "+++++++++++++++++++++++++" & vbCrLf
                    'txtOutput.Text = txtOutput.Text & addressArray(j) & vbCrLf
                    strAddressItem = addressArray(j)
                    addressItemArray = Split(strAddressItem, "^")
                    For i = 0 To UBound(addressItemArray)
                        'txtOutput.Text = txtOutput.Text & i & "  -  " & addressItemArray(i) & vbCrLf
                    Next

                    Select Case Mid(addressItemArray(6), 1, 1)
                        Case "O"
                            If j = 0 Then
                                gblAddr1 = addressItemArray(0) & " " & addressItemArray(1)
                                gblAddr2 = addressItemArray(2) & ", " & addressItemArray(3) & "  " & addressItemArray(4)
                            ElseIf j = 1 Then
                                gblAddr1_2 = addressItemArray(0) & " " & addressItemArray(1)
                                gblAddr2_2 = addressItemArray(2) & ", " & addressItemArray(3) & "  " & addressItemArray(4)
                            ElseIf j = 2 Then
                                gblAddr1_3 = addressItemArray(0) & " " & addressItemArray(1)
                                gblAddr2_3 = addressItemArray(2) & ", " & addressItemArray(3) & "  " & addressItemArray(4)
                            End If

                    End Select
                Next
            End If 'if strAddress <> ""
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Process Address Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
        Finally
        End Try
    End Sub
    Public Sub processPhone(ByVal dictNVP As Hashtable)
        Dim strPhone As String = ""
        Dim strPhoneItem As String = ""
        Dim i As Integer = 0
        Dim tempstr As String = ""
        Dim intMRNUM As Long = 0
        Dim addit As Boolean
        Dim updateit As Boolean
        addit = False
        updateit = False
        Dim intFaxCounter As Integer = 0
        Dim j As Integer = 0
        Try
            strPhone = dictNVP.Item("Phone")
            If strPhone <> "" Then ' 20111101
                phoneArray = Split(strPhone, "~")
                For j = 0 To UBound(phoneArray)

                    strPhoneItem = phoneArray(j)
                    phoneItemArray = Split(strPhoneItem, "^")
                    For i = 0 To UBound(phoneItemArray)
                        'txtOutput.Text = txtOutput.Text & i & "  -  " & phoneItemArray(i) & vbCrLf
                    Next

                    Select Case phoneItemArray(8)
                        Case "CO"
                            If j = 0 Then
                                'txtOutput.Text = txtOutput.Text & "Office Phone: " & phoneItemArray(6) & vbCrLf
                                gblPhone = phoneItemArray(6)
                            ElseIf j = 1 Then
                                gblPhone_2 = phoneItemArray(6)
                            ElseIf j + 2 Then
                                gblPhone_3 = phoneItemArray(6)
                            End If
                        Case "CF"
                            'txtOutput.Text = txtOutput.Text & "Office Fax: " & phoneItemArray(6) & vbCrLf
                            If intFaxCounter = 0 Then
                                gblFax = phoneItemArray(6)
                                intFaxCounter += 1
                            ElseIf intFaxCounter = 1 Then
                                gblFax_2 = phoneItemArray(6)
                                intFaxCounter += 1
                            ElseIf intFaxCounter = 2 Then
                                gblFax_3 = phoneItemArray(6)
                                intFaxCounter += 1
                            End If
                        Case "CH"
                            'txtOutput.Text = txtOutput.Text & "Home Phone: " & phoneItemArray(6) & vbCrLf
                            GblHome = phoneItemArray(6)
                        Case "CB"
                            'txtOutput.Text = txtOutput.Text & "Beeper: " & phoneItemArray(6) & vbCrLf

                        Case "CP"
                            'txtOutput.Text = txtOutput.Text & "Pager: " & phoneItemArray(6) & vbCrLf
                            gblPager = phoneItemArray(6)
                        Case "CC"
                            'txtOutput.Text = txtOutput.Text & "Mobile Phone: " & phoneItemArray(6) & vbCrLf
                            gblMobile = phoneItemArray(6)
                    End Select
                Next
            End If 'if strPhone <> ""
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Process Phone Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
        Finally
        End Try
    End Sub
    Public Sub processName(ByVal dictNVP As Hashtable)
        Try

            gblFirstName = dictNVP("Staff Name Given Name")
            gblMiddle = dictNVP("Staff Name Middle Name")
            gblLastName = dictNVP("Staff Name Family Name")
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Process Name Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
        Finally
        End Try
    End Sub

    Public Sub processCurrentStatus(ByVal dictNVP As Hashtable)
        '20111005 - added to handle additional status fields
        Try

            gblCurrentStatus = dictNVP("CurrentStatus")
            gblStatusCategory = dictNVP("StatusCategory")

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Process Current Status Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
        Finally
        End Try
    End Sub

    Public Sub processDocID(ByVal dictNVP As Hashtable)
        Dim strDocID As String = ""
        Dim strDocIDItem As String = ""
        Dim i As Integer = 0
        Dim tempstr As String = ""
        Dim intMRNUM As Long = 0
        Dim addit As Boolean
        Dim updateit As Boolean
        addit = False
        updateit = False

        Dim j As Integer = 0
        Try
            strDocID = dictNVP.Item("Practitioner ID Numbers")
            docIDArray = Split(strDocID, "~")
            If UBound(docIDArray) > 0 Then
                For j = 0 To UBound(docIDArray)
                    'txtOutput.Text = txtOutput.Text & "<><><><><><><><><><><><><>" & vbCrLf
                    'txtOutput.Text = txtOutput.Text & docIDArray(j) & vbCrLf
                    strDocIDItem = docIDArray(j)
                    docIDItemArray = Split(strDocIDItem, "^")
                    For i = 0 To UBound(docIDItemArray)
                        'txtOutput.Text = txtOutput.Text & i & "  -  " & docIDItemArray(i) & vbCrLf
                    Next

                    Select Case Trim(Mid(docIDItemArray(1), 1, 3))
                        Case "SL"
                            'txtOutput.Text = txtOutput.Text & "State Licence: " & docIDItemArray(0) & vbCrLf
                            'txtOutput.Text = txtOutput.Text & "State: " & docIDItemArray(2) & vbCrLf
                            'txtOutput.Text = txtOutput.Text & "Expiration Date: " & docIDItemArray(3) & vbCrLf
                            gblLicNo = docIDItemArray(0) & " (" & docIDItemArray(2) & ")"
                        Case "DEA"
                            'txtOutput.Text = txtOutput.Text & "DEA Licence: " & docIDItemArray(0) & vbCrLf
                            'txtOutput.Text = txtOutput.Text & "State: " & docIDItemArray(2) & vbCrLf
                            'txtOutput.Text = txtOutput.Text & "Expiration Date: " & docIDItemArray(3) & vbCrLf
                            gblDEANum = docIDItemArray(0)
                        Case "UPI"
                            'txtOutput.Text = txtOutput.Text & "Unique Physician ID No: " & docIDItemArray(0) & vbCrLf
                            '20110524 added global upin number
                            gblUPIN = docIDItemArray(0)
                        Case "BCN"
                            'txtOutput.Text = txtOutput.Text & "Blue Cross No: " & docIDItemArray(0) & vbCrLf
                            gblBlueCrossNum = docIDItemArray(0)
                        Case "MCD"
                            'txtOutput.Text = txtOutput.Text & "Medicaid No: " & docIDItemArray(0) & vbCrLf
                        Case "GL"
                            'txtOutput.Text = txtOutput.Text & "General Ledger No: " & docIDItemArray(0) & vbCrLf
                        Case "CY"
                            'txtOutput.Text = txtOutput.Text & "Country No: " & docIDItemArray(0) & vbCrLf
                        Case "TAX"
                            'txtOutput.Text = txtOutput.Text & "Tax ID No: " & docIDItemArray(0) & vbCrLf
                        Case "MCR"
                            'txtOutput.Text = txtOutput.Text & "Medicare No: " & docIDItemArray(0) & vbCrLf
                            gblMedicareNum = docIDItemArray(0)
                        Case "L&I"
                            'txtOutput.Text = txtOutput.Text & "Labor and Industries No: " & docIDItemArray(0) & vbCrLf
                        Case "QA"
                            'txtOutput.Text = txtOutput.Text & "QA No: " & docIDItemArray(0) & vbCrLf
                        Case "TRL"
                            'txtOutput.Text = txtOutput.Text & "Training Licence No: " & docIDItemArray(0) & vbCrLf
                        Case "NPI"
                            'txtOutput.Text = txtOutput.Text & "Training Licence No: " & docIDItemArray(0) & vbCrLf
                            gblNPINum = docIDItemArray(0)
                    End Select
                Next
            End If 'if ubound > 0
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Process Doc ID Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
        Finally
        End Try
    End Sub
    Public Sub processSpecialty(ByVal dictNVP As Hashtable)
        Dim strSpecialty As String = ""
        Dim strSpecialtyItem As String = ""
        Dim i As Integer = 0
        Dim tempstr As String = ""
        Dim intMRNUM As Long = 0
        Dim addit As Boolean
        Dim updateit As Boolean
        addit = False
        updateit = False

        Dim j As Integer = 0
        Try
            strSpecialty = dictNVP.Item("Specialty")
            specialtyArray = Split(strSpecialty, "~")
            For j = 0 To UBound(specialtyArray)
                'txtOutput.Text = txtOutput.Text & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                'txtOutput.Text = txtOutput.Text & specialtyArray(j) & vbCrLf
                strSpecialtyItem = specialtyArray(j)
                specialtyItemArray = Split(strSpecialtyItem, "^")
                For i = 0 To UBound(specialtyItemArray)
                    'txtOutput.Text = txtOutput.Text & i & "  -  " & specialtyItemArray(i) & vbCrLf
                Next
                'txtOutput.Text = txtOutput.Text & "Specialty: " & specialtyItemArray(0) & vbCrLf
                'txtOutput.Text = txtOutput.Text & "Governing Board: " & specialtyItemArray(1) & vbCrLf
                'txtOutput.Text = txtOutput.Text & "Eligible/Certified: " & specialtyItemArray(2) & vbCrLf
                'txtOutput.Text = txtOutput.Text & "Date of Certification: " & specialtyItemArray(3) & vbCrLf

                gblSpeciality = gblSpeciality & "  " & specialtyItemArray(0)

            Next
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Process Specialty Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
        Finally
        End Try
    End Sub
    Public Sub processStaffID(ByVal dictNVP As Hashtable)
        Dim strStaffIDCode As String = ""
        Dim strStaffIDCodeItem As String = ""
        Dim i As Integer = 0
        Dim tempstr As String = ""
        Dim intMRNUM As Long = 0
        Dim addit As Boolean
        Dim updateit As Boolean
        addit = False
        updateit = False

        'Dim staffIDCodeArray() As String
        'Dim staffIDCodeArrayItem() As String

        Dim j As Integer = 0
        Try
            gblStatus = dictNVP.Item("Status Flag ID")
            strStaffIDCode = dictNVP.Item("Staff ID Code")
            staffIDCodeArray = Split(strStaffIDCode, "~")
            For j = 0 To UBound(staffIDCodeArray)
                'txtOutput.Text = txtOutput.Text & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                'txtOutput.Text = txtOutput.Text & staffIDCodeArray(j) & vbCrLf
                strStaffIDCodeItem = staffIDCodeArray(j)
                staffIDCodeArrayItem = Split(strStaffIDCodeItem, "^")
                For i = 0 To UBound(staffIDCodeArrayItem)

                    If j = 1 And i = 0 Then
                        'txtOutput.Text = txtOutput.Text & "PhysNum" & "  -  " & staffIDCodeArrayItem(i) & vbCrLf
                        gblPhysnum = staffIDCodeArrayItem(i)
                    Else

                        'txtOutput.Text = txtOutput.Text & i & "  -  " & staffIDCodeArrayItem(i) & vbCrLf
                    End If
                Next
            Next
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Process Staff ID Error Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
        Finally
        End Try
    End Sub
    Public Sub updateMcare(ByVal dictNVP As Hashtable)
        Dim i As Integer = 0
        Dim Locsql As String = ""
        Dim tempstr As String = ""
        Dim updateIT As Boolean = False
        Dim recordExists As Boolean = False
        Dim strTempName As String = ""

        Try
            Dim myConnection As New SqlConnection(McareConnectionString)
            Dim objCommand As New SqlCommand
            objCommand.Connection = myConnection
            Dim dataReader As SqlDataReader

            Locsql = "select * from [115phys] "
            Locsql = Locsql & "where physnum = " & gblPhysnum
            objCommand.CommandText = Locsql
            myConnection.Open()
            dataReader = objCommand.ExecuteReader()

            If dataReader.HasRows Then
                recordExists = True
            Else
                recordExists = False

            End If


            dataReader.Close()
            myConnection.Close()

            If gblAddr1_2 <> "" Then
                gblNotes = gblNotes & gblAddr1_2 & vbCrLf
                gblNotes = gblNotes & gblAddr2_2 & vbCrLf
                gblNotes = gblNotes & "Phone: " & gblPhone_2 & vbCrLf
                gblNotes = gblNotes & "Fax: " & gblFax_2 & vbCrLf

            End If

            If gblAddr1_3 <> "" Then
                gblNotes = gblNotes & gblAddr1_3 & vbCrLf
                gblNotes = gblNotes & gblAddr2_3 & vbCrLf
                gblNotes = gblNotes & "Phone: " & gblPhone_3 & vbCrLf
                gblNotes = gblNotes & "Fax: " & gblFax_3 & vbCrLf

            End If
            If Mid$(Replace(gblMiddle, "'", "''"), 1, 1) <> "" Then
                strTempName = Replace(gblLastName, "'", "''") & ", " & Replace(gblFirstName, "'", "''") & " " & Mid$(Replace(gblMiddle, "'", "''"), 1, 1) & "."
            Else
                strTempName = Replace(gblLastName, "'", "''") & ", " & Replace(gblFirstName, "'", "''") & " " & Mid$(Replace(gblMiddle, "'", "''"), 1, 1)
            End If
            If recordExists Then
                
                'update record
                sql = "UPDATE [115phys] "
                sql = sql & "SET modified = '" & DateTime.Now & "'"

                sql = sql & ", firstname = '" & Replace(gblFirstName, "'", "''") & "'"
                sql = sql & ", MI = '" & Replace(gblMiddle, "'", "''") & "'"
                sql = sql & ", lname = '" & Replace(gblLastName, "'", "''") & "'"
                sql = sql & ", lastname = '" & strTempName & "'"

                If gblAddr1 <> "" Then ' 20111101
                    sql = sql & ", addr1 = '" & Replace(Mid$(gblAddr1, 1, 99), "'", "''") & "'"
                    sql = sql & ", addr2 = '" & Replace(gblAddr2, "'", "''") & "'"
                End If

                sql = sql & ", licno = '" & Replace(gblLicNo, "'", "''") & "'"

                If Replace(gblPhone, "'", "''") <> "" Then ' 20111101
                    sql = sql & ", pager = '" & Replace(gblPager, "'", "''") & "'"
                    sql = sql & ", phone = '" & Replace(gblPhone, "'", "''") & "'"
                    sql = sql & ", fax = '" & Replace(gblFax, "'", "''") & "'"
                    sql = sql & ", mobile = '" & Replace(gblMobile, "'", "''") & "'"
                    sql = sql & ", home = '" & Replace(GblHome, "'", "''") & "'"
                End If

                sql = sql & ", specialty = '" & Replace(gblSpeciality, "'", "''") & "'"

                sql = sql & ", DEANum = '" & Replace(gblDEANum, "'", "''") & "'"
                '20110524 added upin processing and statusFlag
                sql = sql & ", UPIN = '" & Replace(gblUPIN, "'", "''") & "'"
                sql = sql & ", NPI = '" & Replace(gblNPINum, "'", "''") & "'"
                sql = sql & ", statusFlag = '" & Replace(gblStatus, "'", "''") & "'"

                '20111005 - added code to process CurrentStatus and StatusCategory
                sql = sql & ", CurrentStatus = '" & Replace(gblCurrentStatus, "'", "''") & "'"
                sql = sql & ", StatusCategory = '" & Replace(gblStatusCategory, "'", "''") & "'"

                sql = sql & ", blueCrossNum = '" & Replace(gblBlueCrossNum, "'", "''") & "'"
                sql = sql & ", medicareNum = '" & Replace(gblMedicareNum, "'", "''") & "'"
                sql = sql & ", notes = '" & Replace(gblNotes, "'", "''") & "'"
                sql = sql & " Where physnum = " & gblPhysnum


                objCommand.CommandText = sql
                myConnection.Open()
                objCommand.ExecuteNonQuery()
                myConnection.Close()
            Else
                'Insert record
                '20110524 added upin processing
                '20111005 - added CurrentStatus and StatusCategory
                sql = "Insert [115phys] "
                sql = sql & "(physnum, firstname, MI, lname, lastname, addr1, addr2, licno, pager, phone, "
                sql = sql & "fax, mobile, home, specialty, DEANum, upin, npi, statusFlag, blueCrossNum, notes, "
                sql = sql & "CurrentStatus, StatusCategory, "
                sql = sql & "medicareNum) "
                sql = sql & "VALUES ("
                insertNumber(gblPhysnum)
                insertString(gblFirstName)
                insertString(gblMiddle)
                insertString(gblLastName)
                insertString(strTempName)
                insertString(Mid$(gblAddr1, 1, 99))
                insertString(gblAddr2)
                insertString(gblLicNo)
                insertString(gblPager)
                insertString(gblPhone)
                insertString(gblFax)
                insertString(gblMobile)
                insertString(GblHome)
                insertString(gblSpeciality)
                insertString(gblDEANum)
                insertString(gblUPIN)
                insertString(gblNPINum)
                insertString(gblStatus)
                insertString(gblBlueCrossNum)
                insertString(gblNotes)
                '20111005
                insertString(gblCurrentStatus)
                insertString(gblStatusCategory)

                insertLastString(gblMedicareNum)

                objCommand.CommandText = sql
                myConnection.Open()

                objCommand.ExecuteNonQuery()
                myConnection.Close()
            End If






        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Mcare Update Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub

        End Try
    End Sub
    Public Sub insertLastString(ByVal theString As String)
        Try
            If theString <> "" Then
                sql = sql & "'" & Replace(theString, "'", "''") & "') "
            Else
                sql = sql & "NULL) "
            End If
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Insert Last String Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub
    Public Sub updateITW(ByVal dictNVP As Hashtable)
        Dim i As Integer = 0
        Dim Locsql As String = ""
        Dim tempstr As String = ""
        Dim updateIT As Boolean = False
        Dim recordExists As Boolean = False
        Dim strTempName As String = ""

        Try
            Dim myConnection As New SqlConnection(ITWConnectionString)
            Dim objCommand As New SqlCommand
            objCommand.Connection = myConnection
            Dim dataReader As SqlDataReader

            Locsql = "select * from [114phys] "
            Locsql = Locsql & "where physnum = " & gblPhysnum
            objCommand.CommandText = Locsql
            myConnection.Open()
            dataReader = objCommand.ExecuteReader()

            If dataReader.HasRows Then
                recordExists = True
            Else
                recordExists = False

            End If


            dataReader.Close()
            myConnection.Close()

            If gblAddr1_2 <> "" Then
                gblNotes = gblNotes & gblAddr1_2 & vbCrLf
                gblNotes = gblNotes & gblAddr2_2 & vbCrLf
                gblNotes = gblNotes & "Phone: " & gblPhone_2 & vbCrLf
                gblNotes = gblNotes & "Fax: " & gblFax_2 & vbCrLf

            End If

            If gblAddr1_3 <> "" Then
                gblNotes = gblNotes & gblAddr1_3 & vbCrLf
                gblNotes = gblNotes & gblAddr2_3 & vbCrLf
                gblNotes = gblNotes & "Phone: " & gblPhone_3 & vbCrLf
                gblNotes = gblNotes & "Fax: " & gblFax_3 & vbCrLf

            End If

            If Mid$(Replace(gblMiddle, "'", "''"), 1, 1) <> "" Then
                strTempName = Replace(gblLastName, "'", "''") & ", " & Replace(gblFirstName, "'", "''") & " " & Mid$(Replace(gblMiddle, "'", "''"), 1, 1) & "."
            Else
                strTempName = Replace(gblLastName, "'", "''") & ", " & Replace(gblFirstName, "'", "''") & " " & Mid$(Replace(gblMiddle, "'", "''"), 1, 1)
            End If

            If recordExists Then
                'update record
                sql = "UPDATE [114phys] "
                sql = sql & "SET modified = '" & DateTime.Now & "'"

                sql = sql & ", firstname = '" & Replace(gblFirstName, "'", "''") & "'"
                sql = sql & ", middle = '" & Mid$(Replace(gblMiddle, "'", "''"), 1, 1) & "'"
                sql = sql & ", lastname = '" & strTempName & "'"
                sql = sql & ", lname = '" & Replace(gblLastName, "'", "''") & "'"

                If gblAddr1 <> "" Then ' 20111101
                    sql = sql & ", addr1 = '" & Replace(Mid$(gblAddr1, 1, 99), "'", "''") & "'"
                    sql = sql & ", addr2 = '" & Replace(gblAddr2, "'", "''") & "'"
                End If

                sql = sql & ", licno = '" & Replace(gblLicNo, "'", "''") & "'"

                If Replace(gblPhone, "'", "''") <> "" Then ' 20111101
                    sql = sql & ", pager = '" & Replace(gblPager, "'", "''") & "'"
                    sql = sql & ", phone = '" & Replace(gblPhone, "'", "''") & "'"
                    sql = sql & ", fax = '" & Replace(gblFax, "'", "''") & "'"
                    sql = sql & ", mobile = '" & Replace(gblMobile, "'", "''") & "'"
                    sql = sql & ", home = '" & Replace(GblHome, "'", "''") & "'"
                End If

                sql = sql & ", specialty = '" & Replace(gblSpeciality, "'", "''") & "'"

                sql = sql & ", DEANum = '" & Replace(gblDEANum, "'", "''") & "'"
                '20110524 added upin processing and statusFlag
                sql = sql & ", UPIN = '" & Replace(gblUPIN, "'", "''") & "'"
                sql = sql & ", NPI = '" & Replace(gblNPINum, "'", "''") & "'"
                sql = sql & ", statusFlag = '" & Replace(gblStatus, "'", "''") & "'"

                '20111005 - added code to process CurrentStatus and StatusCategory
                sql = sql & ", CurrentStatus = '" & Replace(gblCurrentStatus, "'", "''") & "'"
                sql = sql & ", StatusCategory = '" & Replace(gblStatusCategory, "'", "''") & "'"


                sql = sql & ", blueCrossNum = '" & Replace(gblBlueCrossNum, "'", "''") & "'"
                sql = sql & ", medicareNum = '" & Replace(gblMedicareNum, "'", "''") & "'"
                sql = sql & ", notes = '" & Replace(gblNotes, "'", "''") & "'"
                sql = sql & " Where physnum = " & gblPhysnum


                objCommand.CommandText = sql
                myConnection.Open()
                objCommand.ExecuteNonQuery()
                myConnection.Close()
            Else
                'Insert record
                '20110524 added upin processing and statusFlag
                '20111005 - added CurrentStatus and StatusCategory
                sql = "Insert [114phys] "
                sql = sql & "(physnum, firstname, middle, lname, lastname, addr1, addr2, licno, pager, phone, "
                sql = sql & "fax, mobile, home, specialty, DEANum, upin, npi, statusFlag, blueCrossNum, notes, "
                sql = sql & "CurrentStatus, StatusCategory, "
                sql = sql & "medicareNum) "
                sql = sql & "VALUES ("
                insertNumber(gblPhysnum)
                insertString(gblFirstName)
                insertString(Mid$(gblMiddle, 1, 1))
                insertString(gblLastName)
                insertString(strTempName)

                insertString(Mid$(gblAddr1, 1, 99))
                insertString(gblAddr2)
                insertString(gblLicNo)
                insertString(gblPager)
                insertString(gblPhone)
                insertString(gblFax)
                insertString(gblMobile)
                insertString(GblHome)
                insertString(gblSpeciality)
                insertString(gblDEANum)

                insertString(gblUPIN)
                insertString(gblNPINum)
                insertString(gblStatus)
                insertString(gblBlueCrossNum)
                insertString(gblNotes)

                '20111005
                insertString(gblCurrentStatus)
                insertString(gblStatusCategory)

                insertLastString(gblMedicareNum)

                objCommand.CommandText = sql
                myConnection.Open()

                objCommand.ExecuteNonQuery()
                myConnection.Close()
            End If






        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "ITW Update Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub

        End Try
    End Sub
    Public Sub updateShelby(ByVal dictNVP As Hashtable)
        Dim i As Integer = 0
        Dim Locsql As String = ""
        Dim tempstr As String = ""
        Dim updateIT As Boolean = False
        Dim recordExists As Boolean = False
        Dim strTempName As String = ""

        Try
            Dim myConnection As New SqlConnection(ShelbyConnectionString)
            Dim objCommand As New SqlCommand
            objCommand.Connection = myConnection
            Dim dataReader As SqlDataReader

            Locsql = "select * from [115phys] "
            Locsql = Locsql & "where physnum = " & gblPhysnum
            objCommand.CommandText = Locsql
            myConnection.Open()
            dataReader = objCommand.ExecuteReader()

            If dataReader.HasRows Then
                recordExists = True
            Else
                recordExists = False

            End If


            dataReader.Close()
            myConnection.Close()

            If gblAddr1_2 <> "" Then
                gblNotes = gblNotes & gblAddr1_2 & vbCrLf
                gblNotes = gblNotes & gblAddr2_2 & vbCrLf
                gblNotes = gblNotes & "Phone: " & gblPhone_2 & vbCrLf
                gblNotes = gblNotes & "Fax: " & gblFax_2 & vbCrLf

            End If

            If gblAddr1_3 <> "" Then
                gblNotes = gblNotes & gblAddr1_3 & vbCrLf
                gblNotes = gblNotes & gblAddr2_3 & vbCrLf
                gblNotes = gblNotes & "Phone: " & gblPhone_3 & vbCrLf
                gblNotes = gblNotes & "Fax: " & gblFax_3 & vbCrLf

            End If
            If Mid$(Replace(gblMiddle, "'", "''"), 1, 1) <> "" Then
                strTempName = Replace(gblLastName, "'", "''") & ", " & Replace(gblFirstName, "'", "''") & " " & Mid$(Replace(gblMiddle, "'", "''"), 1, 1) & "."
            Else
                strTempName = Replace(gblLastName, "'", "''") & ", " & Replace(gblFirstName, "'", "''") & " " & Mid$(Replace(gblMiddle, "'", "''"), 1, 1)
            End If
            If recordExists Then
                
                'update record
                sql = "UPDATE [115phys] "
                sql = sql & "SET modified = '" & DateTime.Now & "'"

                sql = sql & ", firstname = '" & Replace(gblFirstName, "'", "''") & "'"
                sql = sql & ", MI = '" & Replace(gblMiddle, "'", "''") & "'"
                sql = sql & ", lname = '" & Replace(gblLastName, "'", "''") & "'"
                sql = sql & ", lastname = '" & strTempName & "'"

                If gblAddr1 <> "" Then ' 20111101
                    sql = sql & ", addr1 = '" & Replace(Mid$(gblAddr1, 1, 99), "'", "''") & "'"
                    sql = sql & ", addr2 = '" & Replace(gblAddr2, "'", "''") & "'"
                End If

                sql = sql & ", licno = '" & Replace(gblLicNo, "'", "''") & "'"

                If Replace(gblPhone, "'", "''") <> "" Then ' 20111101
                    sql = sql & ", pager = '" & Replace(gblPager, "'", "''") & "'"
                    sql = sql & ", phone = '" & Replace(gblPhone, "'", "''") & "'"

                    sql = sql & ", fax = '" & Replace(gblFax, "'", "''") & "'"
                    sql = sql & ", mobile = '" & Replace(gblMobile, "'", "''") & "'"
                    sql = sql & ", home = '" & Replace(GblHome, "'", "''") & "'"
                End If

                sql = sql & ", specialty = '" & Replace(gblSpeciality, "'", "''") & "'"

                sql = sql & ", DEANum = '" & Replace(gblDEANum, "'", "''") & "'"
                '20110524 added upin processing and statusFlag
                sql = sql & ", UPIN = '" & Replace(gblUPIN, "'", "''") & "'"
                sql = sql & ", NPI = '" & Replace(gblNPINum, "'", "''") & "'"
                sql = sql & ", statusFlag = '" & Replace(gblStatus, "'", "''") & "'"

                '20111005 - added code to process CurrentStatus and StatusCategory
                sql = sql & ", CurrentStatus = '" & Replace(gblCurrentStatus, "'", "''") & "'"
                sql = sql & ", StatusCategory = '" & Replace(gblStatusCategory, "'", "''") & "'"

                sql = sql & ", blueCrossNum = '" & Replace(gblBlueCrossNum, "'", "''") & "'"
                sql = sql & ", medicareNum = '" & Replace(gblMedicareNum, "'", "''") & "'"
                sql = sql & ", notes = '" & Replace(gblNotes, "'", "''") & "'"
                sql = sql & " Where physnum = " & gblPhysnum


                objCommand.CommandText = sql
                myConnection.Open()
                objCommand.ExecuteNonQuery()
                myConnection.Close()
            Else
                'Insert record
                '20110524 added upin processing and statusFlag
                '20111005 - added CurrentStatus and StatusCategory
                sql = "Insert [115phys] "
                sql = sql & "(physnum, firstname, MI, lname, lastname, addr1, addr2, licno, pager, phone, "
                sql = sql & "fax, mobile, home, specialty, DEANum, upin, npi, statusFlag, blueCrossNum, notes, "
                sql = sql & "CurrentStatus, StatusCategory, "
                sql = sql & "medicareNum) "
                sql = sql & "VALUES ("
                insertNumber(gblPhysnum)
                insertString(gblFirstName)
                insertString(gblMiddle)
                insertString(gblLastName)
                insertString(strTempName)
                insertString(Mid$(gblAddr1, 1, 99))
                insertString(gblAddr2)
                insertString(gblLicNo)
                insertString(gblPager)
                insertString(gblPhone)
                insertString(gblFax)
                insertString(gblMobile)
                insertString(GblHome)
                insertString(gblSpeciality)
                insertString(gblDEANum)
                insertString(gblUPIN)
                insertString(gblNPINum)
                insertString(gblStatus)
                insertString(gblBlueCrossNum)
                insertString(gblNotes)
                '20111005
                insertString(gblCurrentStatus)
                insertString(gblStatusCategory)

                insertLastString(gblMedicareNum)

                objCommand.CommandText = sql
                myConnection.Open()

                objCommand.ExecuteNonQuery()
                myConnection.Close()
            End If






        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Shelby Update Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub

        End Try
    End Sub
End Module
