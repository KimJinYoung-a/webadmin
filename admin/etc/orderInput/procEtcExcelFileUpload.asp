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
Dim mallid
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
	sDefaultPath = Server.MapPath("/admin/etc/orderInput/upEtcExcelFiles/")
	uploadform.Start sDefaultPath '���ε���

	iMaxLen = uploadform.Form("iML")	'�̹�������ũ��
	mallid = uploadform.Form("mallid")
	If (fnChkFile(uploadform("sFile"), iMaxLen,"xls")) Then	'����üũ
	    sFolderPath = sDefaultPath&"/"&mallid&"/"
	    If NOT  objfile.FolderExists(sFolderPath) Then
	    	objfile.CreateFolder sFolderPath
	    End If

	     sFolderPath = sDefaultPath&"/"&mallid&"/"&monthFolder&"/"
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
''''	0			1				2				3			4			5
''''TEN ��ǰ��ȣ	TEN �ɼǹ�ȣ		���� ��ǰ��ȣ	���� ��ǰ��	���޿ɼǸ�	�ǸŰ�

xlPosArr = Array(0,1,2,3,4,5)
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
Dim iLine, iResult, RetErr, sqlStr
Dim errCNT, objCmd, errItem
Dim pcnt : pcnt = UBound(xlRowALL)
Dim iExists, itemid, itemoption, outmallitemid, outmallitemname, outmallitemOptionname, outmallPrice
errCNT = 0

For i = 0 to pcnt
	If IsObject(xlRowALL(i)) Then
		Set iLine = xlRowALL(i)
			itemid					= Trim(iLine.FItemArray(0))
			itemoption				= Trim(iLine.FItemArray(1))
			outmallitemid			= Trim(iLine.FItemArray(2))
			outmallitemname			= Trim(iLine.FItemArray(3))
			outmallitemOptionname	= Trim(iLine.FItemArray(4))
			outmallPrice			= Trim(iLine.FItemArray(5))

			iExists = false
			sqlStr = "SELECT COUNT(*) as CNT FROM db_item.dbo.tbl_item WHERE itemid="&itemid
			rsget.Open sqlStr,dbget,1
			If rsget("CNT") < 1 Then
				errCNT = errCNT + 1
				RetErr = RetErr & "��ǰ�ڵ尡 �������� �ʽ��ϴ�." & " \n"
			End If
			rsget.close

			If (itemoption<>"") Then
				If (itemoption<>"0000") Then
					sqlStr = "SELECT COUNT(*) as CNT FROM db_item.dbo.tbl_item_option WHERE itemid="&itemid&" and itemoption='"&itemoption&"'"
					rsget.Open sqlStr,dbget,1
					If Not rsget.Eof Then
						iExists = rsget("CNT")>0
					End If
					rsget.close

					If (Not iExists) Then
						errCNT = errCNT + 1
						RetErr = RetErr & "�ɼ��ڵ尡 �������� �ʽ��ϴ�. �ɼ��� ���� ��� �Ǵ� �ɼǺ� ��Ī�� �ʿ��� ��츸 �Է�" & " \n"
					End If
				Else
					sqlStr = "SELECT COUNT(*) as CNT FROM db_item.dbo.tbl_item_option where itemid="&itemid
					rsget.Open sqlStr,dbget,1
					If Not rsget.Eof Then
						iExists = rsget("CNT")>0
					End If
					rsget.close

					If (iExists) Then
						errCNT = errCNT + 1
						RetErr = RetErr & "�ɼ��� �����ϴ� ��ǰ�Դϴ�. 0000 �Է� �Ұ�" & " \n"
					End If
				End If
			End If

		    sqlStr = "SELECT COUNT(*) as CNT "
		    sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xSite_EtcItemLink"
		    sqlStr = sqlStr & " WHERE mallid='"&mallid&"'"
		    sqlStr = sqlStr & " and itemid='"&itemid&"'"
		    sqlStr = sqlStr & " and itemoption='"&itemoption&"'"
		    rsget.Open sqlStr,dbget,1
		    If Not rsget.Eof Then
		        iExists = rsget("CNT")>0
		    End If
		    rsget.close

		    If (iExists) Then
				errCNT = errCNT + 1
				RetErr = RetErr & "�̹� ��ϵ� ��ǰ�ڵ� [�ɼǸ�] �Դϴ�." & " \n"
		    End If
		Set iLine = xlRowALL(i)
	End If
Next
If errCNT > 0 Then
	response.write "<script>alert('"&errCNT&"�� ����.\n\n"&RetErr&"');opener.location.reload();self.close();</script>"
	response.end
Else
    sqlStr = "Insert Into db_temp.dbo.tbl_xSite_EtcItemLink"
    sqlStr = sqlStr & " (itemid,itemoption,mallID,outmallitemid,outmallitemname,outmallitemOptionname,outmallPrice,outmallSellYn)"
    sqlStr = sqlStr & " values("
    sqlStr = sqlStr & " "&itemid&VbCRLF
    sqlStr = sqlStr & " ,'"&itemoption&"'"&VbCRLF
    sqlStr = sqlStr & " ,'"&mallid&"'"&VbCRLF
    sqlStr = sqlStr & " ,'"&outmallitemid&"'"&VbCRLF
    sqlStr = sqlStr & " ,'"&html2DB(outmallitemname)&"'"&VbCRLF
    sqlStr = sqlStr & " ,'"&html2DB(outmallitemOptionname)&"'"&VbCRLF
    sqlStr = sqlStr & " ,"&(outmallPrice)&""&VbCRLF
    sqlStr = sqlStr & " ,'Y'"&VbCRLF
    sqlStr = sqlStr & " )"
    dbget.execute sqlStr
End If

response.write "<script>opener.location.reload();self.close();</script>"
'End If
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