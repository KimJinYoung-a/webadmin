<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbBulkInsopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/orderInput/incUploadFunction.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

'���丮�� ������ ����� �Լ�
Function CreateDirectoryIfNotExists(sPath)
	Dim oFSO : Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	If Not oFSO.FolderExists(sPath) Then
		oFSO.CreateFolder(sPath)
	End If

	Set oFSO = Nothing
End Function

'��Ű�� ������ ���� �����ϴ� �Լ�
Function UnzipFile(sUploadPath, isubfoldername, ifilename)
	Dim oZip : Set oZip = Server.CreateObject("Chilkat.Zip2")
	oZip.UnlockComponent("10X10CZIP_4HmoweDQnXfy")
	oZip.OpenZip(sUploadPath&"\"&ifilename)

	CreateDirectoryIfNotExists(sUploadPath & "\" & isubfoldername)

	Dim i
	Dim n : n = oZip.NumEntries
	For i = 0 To n - 1
		Dim oEntry : Set oEntry = oZip.GetEntryByIndex(i)

		oEntry.Extract(sUploadPath & "\" & isubfoldername)

		Set oEntry = Nothing
	Next

	oZip.CloseZip

	Set oZip = Nothing

	'' ���ϻ���.
	Dim oFSO : Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	If oFSO.FileExists(sUploadPath&"\"&ifilename) Then
		oFSO.deletefile(sUploadPath&"\"&ifilename)
	End If
	Set oFSO = Nothing

End Function

Class TXLRowObj
	Public FItemArray

	Public Function setArrayLength(ln)
		Redim FItemArray(ln)
	End Function
End Class

Function IsSKipRow(ixlRow, skipCol0Str)
	If Not IsArray(ixlRow) Then
		IsSKipRow = true
		Exit Function
	End if

	If  LCASE(ixlRow(0)) = LCASE(skipCol0Str) Then
		IsSKipRow = true
		Exit Function
	End If

    If  LCASE(ixlRow(0)) = "��ǰ��" Then
		IsSKipRow = true
		Exit Function
	End If

	IsSKipRow = false
End Function 

Function fnGetXLFileArray(byref xlRowALL, sFilePath, aSheetName, iArrayLen, ixlPosArr, iskipString)
	Dim conDB, Rs, strQry, iResult, i, J, iObj
	Dim irowObj, strTable
	'' on Error ���� ���� �ȵ�.. ���� ���ѷ��� ���µ�.
	Set conDB = Server.CreateObject("ADODB.Connection")
		'conDB.Provider = "Microsoft.Jet.oledb.4.0"
		conDB.Provider = "Microsoft.ace.oledb.12.0"
		'conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;IMEX=1"  ''';IMEX=1 �߰� 2013/12/19
		conDB.Properties("ExtEnded Properties").Value = "Excel 12.0;IMEX=1"  ''';IMEX=1 �߰� 2013/12/19
	'On Error Resume Next
		conDB.Open sFilePath
		If (ERR) Then
			fnGetXLFileArray=false
			'/������ �˼� ���� ������ ������. "����ġ ���� ����. �ܺ� ��ü�� Ʈ�� ������ ����(C0000005)�� �߻��߽��ϴ�. ��ũ��Ʈ�� ��� ������ �� �����ϴ�"
			set conDB = nothing
			Exit Function
		End If
	'On Error Goto 0
		'' get First Sheet Name=============''��Ʈ�� �������ΰ�� ������ �� ����.
		Set Rs = conDB.OpenSchema(adSchemaTables)
			If Not Rs.Eof Then
				aSheetName = Rs.Fields("table_name").Value
				''rw "aSheetName="&aSheetName
			End If
		Set Rs = Nothing
		''==================================
		Set Rs = Server.CreateObject("ADODB.Recordset")
			''strQry = "Select * From [sheet1$]"
			strQry = "Select * From ["&aSheetName&"]"
			ReDim xlRowALL(0)
			fnGetXLFileArray = true

		'On Error Resume Next
			Rs.Open strQry, conDB
			IF (ERR) then
				fnGetXLFileArray=false
				Rs.Close
				Set Rs = Nothing
				Set conDB = Nothing
				Exit Function
			End if

			If Not Rs.Eof Then
				Do Until Rs.Eof
					If (ERR) Then
						fnGetXLFileArray=false
						Rs.Close
						Set Rs = Nothing
						Set conDB = Nothing
						Exit Function
					End if

					Set irowObj = new TXLRowObj
						irowObj.setArrayLength(iArrayLen)

						For i=0 to iArrayLen
							If (ixlPosArr(i) < 0) Then
								irowObj.FItemArray(i) = ""
							Else
								irowObj.FItemArray(i) = Replace(null2blank(Rs(ixlPosArr(i))),"*","")
							End If
							''rw irowObj.FItemArray(i)
						Next

						IF (Not IsSKipRow(irowObj.FItemArray,iskipString)) then
							ReDim Preserve xlRowALL(UBound(xlRowALL)+1)
							set xlRowALL(UBound(xlRowALL)) =  irowObj
							''xlRowALL(UBound(xlRowALL)).arrayObj = xlRow
						END IF
					set irowObj = Nothing
					Rs.MoveNext
				Loop
			Else
				fnGetXLFileArray=false
			End If
			''''On Error Goto 0
			If (ERR) Then
				fnGetXLFileArray=false
			End If
			Rs.Close
		'On Error Goto 0
		Set Rs = Nothing
	Set conDB = Nothing
	If Ubound(xlRowALL) <  1 Then fnGetXLFileArray = false
End Function 


Function ConvXlsxToCsv(iFilePath,imonthFolder,isubfolder,iFileName,byref newCvsFileName,byRef RetErr)
	Dim xlPosArr : xlPosArr = Array(0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16)
	Dim ArrayLen : ArrayLen = UBound(xlPosArr)
	Dim skipString : skipString = "Sheet1"
	Dim afile : afile = iFilePath&iFileName
	Dim aSheetName : aSheetName = ""
	Dim newCvsFilePath
	ConvXlsxToCsv = False

	Dim xlRowALL
	Dim ret : ret = fnGetXLFileArray(xlRowALL, afile, aSheetName, ArrayLen, xlPosArr, skipString)
	If (Not ret) or (Not IsArray(xlRowALL)) then
		RetErr = "������ �ùٸ��� �ʰų� ������ �����ϴ�."
		Exit function 
	End If

	Dim i,j
	Dim pcnt : pcnt = UBound(xlRowALL)

	' If (application("Svr_Info")	= "Dev") Then
	' 	newCvsFilePath = iFilePath&"\"&replace(iFileName,".xlsx",".csv")
	' else
	' 	Dim objfile
	' 	Set objfile = Server.CreateObject("Scripting.FileSystemObject")
	' 	If NOT  objfile.FolderExists("\\192.168.0.103\naverlowprc\"&imonthFolder) Then
	' 		objfile.CreateFolder "\\192.168.0.103\naverlowprc\"&imonthFolder
	' 	End If

	' 	If NOT  objfile.FolderExists("\\192.168.0.103\naverlowprc\"&imonthFolder&"\"&isubfolder) Then
	' 		objfile.CreateFolder "\\192.168.0.103\naverlowprc\"&imonthFolder&"\"&isubfolder
	' 	End If
	' 	Set objfile = Nothing
	' 	newCvsFilePath = "\\192.168.0.103\naverlowprc\"&imonthFolder&"\"&isubfolder&"\"&replace(iFileName,".xlsx",".csv")
	' end if
	newCvsFilePath = iFilePath&"\"&replace(iFileName,".xlsx",".csv")

	newCvsFileName = replace(iFileName,".xlsx",".csv")

	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	Dim tFile : Set tFile = fso.CreateTextFile(newCvsFilePath,true,true)  ''�����ڵ��
	Dim oneCsvLine, iLine
	For i = 0 to pcnt
		If IsObject(xlRowALL(i)) Then
			Set iLine = xlRowALL(i)
			oneCsvLine = ""
			for j=LBOUND(iLine.FItemArray) to UBOUND(iLine.FItemArray)
				''oneCsvLine = oneCsvLine &replace(replace(iLine.FItemArray(j),"""",""),",","")&CHR(9)
				oneCsvLine = oneCsvLine &replace(iLine.FItemArray(j),CHR(9),"")&CHR(9)
			next
			if Right(oneCsvLine,1)=CHR(9) then oneCsvLine=LEFT(oneCsvLine,LEN(oneCsvLine)-1)

			if NOT isNULL(oneCsvLine) then
				tFile.WriteLine(oneCsvLine)
			end if

			Set iLine = Nothing
		end if
	Next
	ConvXlsxToCsv = True

	''�� XL���� ����.
	If fso.FileExists(afile) Then
		fso.deletefile(afile)
	End If

	tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
end Function

Function CurrTBLToMovePreTBL()
	Dim cmd : set cmd = server.CreateObject("ADODB.Command")
	Dim intResult
			
	cmd.ActiveConnection = dbBULKINSget
	cmd.CommandText = "[db_analyze_etc].dbo.[sp_Ten_Naver_Lowproce_Exl_MoveToPre]"
	cmd.CommandType = adCmdStoredProc
	cmd.Parameters.Append cmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
	cmd.Execute
	intResult = cmd.Parameters("returnValue").Value

	CurrTBLToMovePreTBL = intResult
End Function

Function CurrTBLAll_Truncate()
	Dim cmd : set cmd = server.CreateObject("ADODB.Command")
	Dim intResult
			
	cmd.ActiveConnection = dbBULKINSget
	cmd.CommandText = "[db_analyze_etc].dbo.[sp_Ten_Naver_Lowproce_Exl_NoMapAll_Truncate]"
	cmd.CommandType = adCmdStoredProc
	cmd.Parameters.Append cmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
	cmd.Execute
	intResult = cmd.Parameters("returnValue").Value

	CurrTBLAll_Truncate = intResult
end function

Function CSVFILEAddToDB(iLocalPath,iFolderPath,imonthFolder,isubfolder,iFileName,byRef RetErr,byVal isNoMapAllItemInsert)
	Dim xlPosArr : xlPosArr = Array(0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16)
	Dim ArrayLen : ArrayLen = UBound(xlPosArr)
	Dim skipString : skipString = "Sheet1"
	Dim aSheetName : aSheetName = ""
	Dim iCvsNetFilePath 
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

	CSVFILEAddToDB = False

	' If (application("Svr_Info")= "Dev") Then
	' 	'admin\etc\naverEp\upFiles
	' 	iCvsNetFilePath = "\\192.168.50.2\testweb\TESTscm\admin\etc\naverEp\upFiles\"
	' Else
	' 	iCvsNetFilePath = "\\192.168.0.94\cube1010\admin2009scm\admin\etc\naverEp\upFiles\"
	' 	''iCvsNetFilePath = "\\192.168.0.103\naverlowprc\"
	' end if

	If (application("Svr_Info")= "Dev") Then
		iCvsNetFilePath = "\\192.168.50.2\naverEp\" 
	else
		iCvsNetFilePath = "\\192.168.0.94\naverEp\" 
	end if

	iCvsNetFilePath = iCvsNetFilePath & imonthFolder &"\"& isubfolder &"\"& iFileName

	'' ���ѻ� ��Ʈ�� ����̹��� ���� ����.
	If (application("Svr_Info")= "Dev") Then
		If NOT(fso.FileExists(iLocalPath)) Then
			Set fso = Nothing
			RetErr = "������ �������� �ʽ��ϴ�.:"&iLocalPath
			Exit function
		end if
	end if 

	Dim i,j
	Dim cmd : set cmd = server.CreateObject("ADODB.Command")
	Dim intResult
			
	cmd.ActiveConnection = dbBULKINSget
	if (isNoMapAllItemInsert) then
		cmd.CommandText = "[db_analyze_etc].dbo.[sp_Ten_Naver_Lowproce_Exl_BulkInsert_NoMapAll]"
		cmd.CommandType = adCmdStoredProc
		cmd.Parameters.Append cmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
		cmd.Parameters.Append cmd.CreateParameter("@csvFilePath", adVarWchar, adParamInput, 200, iCvsNetFilePath) 
		cmd.Execute
		intResult = cmd.Parameters("returnValue").Value
	else
		cmd.CommandText = "[db_analyze_etc].dbo.sp_Ten_Naver_Lowproce_Exl_BulkInsert"
		cmd.CommandType = adCmdStoredProc
		cmd.Parameters.Append cmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
		cmd.Parameters.Append cmd.CreateParameter("@csvFilePath", adVarWchar, adParamInput, 200, iCvsNetFilePath) 
		cmd.Execute
		intResult = cmd.Parameters("returnValue").Value
	end if
	
			

	if (intResult>0) then
		CSVFILEAddToDB = True
	end if

	
	If fso.FileExists(iLocalPath) Then
		fso.deletefile(iLocalPath)
	End If
	Set fso = Nothing
end Function

Dim otime : otime = Timer()
Dim uploadform, objfile, sDefaultPath, sFolderPath ,monthFolder
Dim sFile, sFilePath, iMaxLen, orgFileName, maybeSheetName, i, iuploadform, isNoMapAllItemInsert
''Dim sCode, egCode, eCode
monthFolder = Replace(Left(CStr(now()),7),"-","")

' If (application("Svr_Info")	= "Dev") Then
' 	if (G_IsLocalDev) then
' 		Set uploadform = Server.CreateObject("TABSUpload4.Upload")
' 		''Set uploadform = Server.CreateObject("TABSUpload4.UploadSingle")  '' TabsUpload5���� �����ϴµ�.
' 	else
'     	Set uploadform = Server.CreateObject("TABS.Upload")			'' - TEST : TABS.Upload
' 	end if
' Else
'     Set uploadform = Server.CreateObject("TABSUpload4.Upload")	''REAL : TABSUpload4.Upload
' End If

Set uploadform = Server.CreateObject("TABSUpload4.Upload")

Set objfile = Server.CreateObject("Scripting.FileSystemObject")
	''sDefaultPath = Server.MapPath("/admin/etc/naverEp/upFiles/")
	sDefaultPath = "C:\etcUpFiles\naverEp" '' �ٸ����� ��������.
	uploadform.Start sDefaultPath '�ӽ����� ���ε���

	iMaxLen = 50*1024*1024 ''uploadform.Form("iML")	'�̹�������ũ��

	isNoMapAllItemInsert = (uploadform("isall")="A")
	'' UploadSingle �� �� �ٸ���.
	If (fnChkFile(uploadform("sFile"), iMaxLen,"zip")) Then	'����üũ
		sFolderPath = sDefaultPath&"\"
		If NOT  objfile.FolderExists(sFolderPath) Then
			objfile.CreateFolder sFolderPath
		End If

		sFolderPath = sDefaultPath&"\"&monthFolder&"\"
		If NOT  objfile.FolderExists(sFolderPath) Then
			objfile.CreateFolder sFolderPath
		End If

		sFile = fnMakeFileName(uploadform("sFile"))
		sFilePath = sFolderPath&sFile
		sFilePath = uploadform("sFile").SaveAs(sFilePath, False)

		'orgFileName = uploadform("sFile").FileName
		'maybeSheetName = Replace(orgFileName,"."&uploadform("sFile").FileType,"")

	end if

	
Set objfile = Nothing
Set uploadform = Nothing

dim subfolder : subfolder=replace(sFile,".zip","")
''���� ����
call UnzipFile(sFolderPath, subfolder ,sFile)

dim folderObj, xlfiles, xlfile
Set objfile = Server.CreateObject("Scripting.FileSystemObject")
Set folderObj = objfile.GetFolder(sFolderPath&"\"&subfolder)
Set xlfiles = folderObj.Files 

dim retval
dim retCsvFileNm, retErr, ifilename

''CSV ���� ��ȯ
For Each xlfile in xlfiles
	
	ifilename = xlfile.name
	retval = retval & ifilename & "|"

	''CSV�� ��ȯ�� ����.
	retCsvFileNm = ""

	if (ConvXlsxToCsv(sFolderPath&"\"&subfolder&"\",monthFolder,subfolder,ifilename,retCsvFileNm,retErr)) then
		rw ifilename&":::"&retCsvFileNm
	else
		rw ifilename
	end if
	
	response.flush
Next

SET xlfiles = Nothing
SET objfile = Nothing

Set objfile = Server.CreateObject("Scripting.FileSystemObject")
Set folderObj = objfile.GetFolder(sFolderPath&"\"&subfolder)
Set xlfiles = folderObj.Files 

rw "�ѰǼ�:"&xlfiles.count


''���� ���([dbo].[tbl_nvshop_mapItem])�� �������([dbo].[tbl_nvshop_mapItem_Pre]) ���� �̰��Ѵ�.
Dim retErrNo
if (isNoMapAllItemInsert) then
	retErrNo = CurrTBLAll_Truncate()
	if (retErrNo="1") then
		rw "�������� ����"&retErrNo
	else

	end if
	response.flush
else
	retErrNo = CurrTBLToMovePreTBL()
	rw "�������� �̰�"&retErrNo
	response.flush
end if



'' csv ���� ���ε�
For Each xlfile in xlfiles
	
	ifilename = xlfile.name
	
	if (CSVFILEAddToDB(sFolderPath&"\"&subfolder&"\"&ifilename,sFolderPath,monthFolder,subfolder,ifilename,retErr,isNoMapAllItemInsert)) then
	 	rw ifilename
	else
		rw "ERR:"&retErr
	end if
	
	response.flush
Next

''TENDB �̰�
dim iisql, cmd2
if (isNoMapAllItemInsert) then
	set cmd2 = server.CreateObject("ADODB.Command")
	
	cmd2.ActiveConnection = dbBULKINSget
	cmd2.CommandText = "[db_analyze_etc].dbo.[sp_Ten_Naver_Lowproce_Exl_NoMapAll_ToMap]"
	
	cmd2.CommandType = adCmdStoredProc
	cmd2.Execute
	set cmd2 = Nothing

	rw "�����ڷ��̰�"
	response.flush
end if

iisql = "exec [db_etcmall].[dbo].[usp_TEN_NvMapItem_Job_BatchUpdate] "
dbget.Execute iisql

rw "TENDB �̰�"


''response.write "<script>parent.execFileArr('"&subfolder&"','"&monthFolder&"','"&retval&"');</script>"

dim ttltime : ttltime = FormatNumber(Timer()-otime,0)
rw "FIN - " &ttltime&" ��"
rw "<script>alert('�Ϸ� - "&ttltime&"�� ')</script>"
%>
<!-- #include virtual="/lib/db/dbBulkInsclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->