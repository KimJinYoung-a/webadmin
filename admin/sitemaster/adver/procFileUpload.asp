<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/orderInput/incUploadFunction.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim uploadform, objfile, sDefaultPath, sFolderPath ,monthFolder
Dim sFile, sFilePath, iMaxLen, orgFileName, maybeSheetName
monthFolder = Replace(Left(CStr(now()),7),"-","")

'�Ҽ������� ���ڿ��� üũ
'-----------------------------------------
Function FIsNum(ByVal iValue)
	Dim iLength , i , retValue
	For i = 1 To Len(iValue)
		If Not (( asc(Mid(iValue,i,1)) > 47 And asc(Mid(iValue,i,1)) < 58 ) or asc(Mid(iValue,i,1))  = 46 ) Then
			FIsNum  = False
			Exit For
		Else
			FIsNum = True
		End If
	Next
End Function
'-----------------------------------------
If (application("Svr_Info")	= "Dev") Then
    Set uploadform = Server.CreateObject("TABS.Upload")			'' - TEST : TABS.Upload
Else
    Set uploadform = Server.CreateObject("TABSUpload4.Upload")	''REAL : TABSUpload4.Upload
End If

Set objfile = Server.CreateObject("Scripting.FileSystemObject")
	sDefaultPath = Server.MapPath("/admin/sitemaster/adver/upFiles/")
	uploadform.Start sDefaultPath '���ε���
	iMaxLen = uploadform.Form("iML")	'�̹�������ũ��

	If (fnChkFile(uploadform("sFile"), iMaxLen,"xls")) Then	'����üũ
	    sFolderPath = sDefaultPath&"/"
	    If NOT  objfile.FolderExists(sFolderPath) Then
	    	objfile.CreateFolder sFolderPath
	    End If

	    sFolderPath = sDefaultPath&"/"&monthFolder&"/"
	    If NOT  objfile.FolderExists(sFolderPath) Then
	    	objfile.CreateFolder sFolderPath
	    End If

		sFile = fnMakeFileName(uploadform("sFile"))
		sFilePath = sFolderPath&sFile
		sFilePath = uploadform("sFile").SaveAs(sFilePath, False)

		orgFileName = uploadform("sFile").FileName
		maybeSheetName = Replace(orgFileName,"."&uploadform("sFile").FileType,"")
	End If

Set objfile = Nothing
Set uploadform = Nothing

Dim xlPosArr, ArrayLen, skipString, afile, aSheetName ,i
''''	0		1		2
''''������, 	������, 	��ǰ�ڵ�
xlPosArr = Array(0, 1, 2)
ArrayLen = UBound(xlPosArr)
skipString = "Sheet1"
afile = sFilePath
aSheetName = ""

Dim xlRowALL
Dim ret : ret = fnGetXLFileArray(xlRowALL, afile, aSheetName, ArrayLen)
If (Not ret) or (Not IsArray(xlRowALL)) then
	response.write "<script>alert('������ �ùٸ��� �ʰų� ������ �����ϴ�. "&Replace(Err.Description,"'","")&"');</script>"

	If (Err.Description="�ܺ� ���̺� ������ �߸��Ǿ����ϴ�.") Then
		response.write "<script>alert('�������� Save As Excel 97 -2003 ���չ��� ���·� ������ ����ϼ���.');</script>"
	End If
	response.write "<script>history.back();</script>"
	response.end
End If

''������ ó��.
Dim iLine, iResult, strSql, addSqlDB, addSql, RetErr
Dim errCNT, objCmd, strStatus
Dim pcnt : pcnt = UBound(xlRowALL)
Dim itemArr, itemidarr, arrList, intLoop
Dim dSDate, dEDate, iSaleRate, iSaleMargin, iSaleMarginValue, saleStatus
errCNT = 0
addSqlDB = ""
addSql = ""
For i = 0 to pcnt
	If IsObject(xlRowALL(i)) Then
		Set iLine = xlRowALL(i)
			If (Not FIsNum(iLine.FItemArray(2))) Then
				errCNT = errCNT + 1
				RetErr = RetErr & "���� ������ ��ǰ�ڵ� �� ���ڿ��� ���ڰ� �ֽ��ϴ�." & " \n"
			End If
			itemArr				= itemArr & iLine.FItemArray(2) & ","			'��ǰ�ڵ�
		Set iLine = xlRowALL(i)
	End If
Next

If errCNT > 0 Then
	response.write "<script>alert('"&errCNT&"�� ����.\n\n"&RetErr&"');opener.location.reload();self.close();</script>"
	response.end
Else
	errCNT = 0
	RetErr = ""
	If Right(itemArr,1) = "," Then
		itemidarr = Left(itemArr, Len(itemArr) - 1)
	End If

	Dim vItemid, vAlarmyn
	For i = 0 to pcnt
		If IsObject(xlRowALL(i)) Then
			Set iLine = xlRowALL(i)
				vItemid = ""
				vAlarmyn = ""
			
				strSql = ""
				strSql = strSql & " SELECT Top 1 itemid, alarmyn, startdate, enddate FROM db_sitemaster.[dbo].[tbl_adver_item] WHERE itemid= '"&iLine.FItemArray(2)&"' "
				rsget.Open strSql,dbget,1
				If not rsget.EOF Then
					vItemid		= rsget("itemid")
					vAlarmyn	= rsget("alarmyn")
				End If
				rsget.Close
				If vItemid = "" Then
					strSql = ""
					strSql = strSql & " INSERT INTO db_sitemaster.[dbo].[tbl_adver_item] "
					strSql = strSql & " (itemid, startdate, enddate, alarmyn, regdate) "
					strSql = strSql & " VALUES "
					strSql = strSql & " ('"&iLine.FItemArray(2)&"'"
					strSql = strSql & "	, '"&iLine.FItemArray(0)&"'"
					strSql = strSql & "	, '"&iLine.FItemArray(1)&"'"
					strSql = strSql & "	, 'N'"
					strSql = strSql & "	, getdate()) "
					dbget.execute strSql
				ElseIf vItemid <> "" AND vAlarmyn = "Y" Then
					strSql = ""
					strSql = strSql & " UPDATE db_sitemaster.[dbo].[tbl_adver_item] SET "
					strSql = strSql & " alarmyn='N'"
					strSql = strSql & " , alarmdate = null "
					strSql = strSql & " , lastupdate = getdate() "
					strSql = strSql & " , startdate = '"&iLine.FItemArray(0)&"' "
					strSql = strSql & " , enddate = '"&iLine.FItemArray(1)&"' "
					strSql = strSql & " WHERE itemid = '"&iLine.FItemArray(2)&"' and alarmyn = 'Y' "
					dbget.execute strSql
				ElseIf vItemid <> "" AND vAlarmyn = "N" Then
					strSql = ""
					strSql = strSql & " UPDATE db_sitemaster.[dbo].[tbl_adver_item] SET "
					strSql = strSql & " lastupdate = getdate() "
					strSql = strSql & " , startdate = '"&iLine.FItemArray(0)&"' "
					strSql = strSql & " , enddate = '"&iLine.FItemArray(1)&"' "
					strSql = strSql & " WHERE itemid = '"&iLine.FItemArray(2)&"' "
					dbget.execute strSql
				End If
			Set iLine = xlRowALL(i)
		End If
	Next
	response.write "<script>opener.location.reload();self.close();</script>"
End If
'''====================================================================================
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
	IsSKipRow = false
End Function

Function fnGetXLFileArray(byref xlRowALL, sFilePath, aSheetName, iArrayLen)
	Dim conDB, Rs, strQry, iResult, i, J, iObj
	Dim irowObj, strTable
	'' on Error ���� ���� �ȵ�.. ���� ���ѷ��� ���µ�.
	Set conDB = Server.CreateObject("ADODB.Connection")
		conDB.Provider = "Microsoft.Jet.oledb.4.0"
		'conDB.Provider = "Microsoft.ACE.OLEDB.12.0"
		conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;IMEX=1"  ''';IMEX=1 �߰� 2013/12/19
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

						For i=0 to ArrayLen
							If (xlPosArr(i) < 0) Then
								irowObj.FItemArray(i) = ""
							Else
								irowObj.FItemArray(i) = Replace(null2blank(Rs(xlPosArr(i))),"*","")
							End If
							''rw irowObj.FItemArray(i)
						Next

						IF (Not IsSKipRow(irowObj.FItemArray,skipString)) then
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
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->