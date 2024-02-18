<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/orderInput/incUploadFunction.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Function Math_RoundOff( su1, decimalPlaces)
    Dim sTemp, i, antilog, fraction '����, �Ҽ�
    antilog = 1

    If decimalPlaces > 0 Then ' 10�� 0�� �̻� ó��
        antilog = 10 ^ decimalPlaces
        sTemp = Fix( su1 / antilog + 0.5 ) * antilog
    Else ' 1 �ڸ��� �� ���� ó��
        sTemp = Round( su1 + 0.000001 , -(decimalPlaces))
    End if
    Math_RoundOff = sTemp
End Function

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
Set uploadform = Server.CreateObject("TABSUpload4.Upload")
Set objfile = Server.CreateObject("Scripting.FileSystemObject")
	sDefaultPath = Server.MapPath("/admin/etc/mustprice/upFiles/")
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
''''	0		1		2		3			4			5				6			7
''''������	��ǰ�ڵ�	Ư��	Ư���ø���	�Ⱓ������	�Ⱓ������	���󰡽�����	����������
xlPosArr = Array(0,1,2,3,4,5,6,7)
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
Dim errCNT, objCmd, strStatus, errItem, cnt
Dim pcnt : pcnt = UBound(xlRowALL)
Dim arrList, intLoop, calcuBuyPrice
Dim dSDate, dEDate, iSaleRate, iSaleMargin, iSaleMarginValue, saleStatus
Dim isMustPriceData : isMustPriceData = False
errCNT = 0
addSqlDB = ""
addSql = ""

For i = 0 to pcnt
	If IsObject(xlRowALL(i)) Then
		Set iLine = xlRowALL(i)
			If (Len(iLine.FItemArray(0)) >= 3) Then
				isMustPriceData = True
			End If

			If isMustPriceData = True Then
				iLine.FItemArray(1) = replace(iLine.FItemArray(1) ,",","")
				iLine.FItemArray(2) = replace(iLine.FItemArray(2) ,",","")

				If (Not FIsNum(iLine.FItemArray(1))) OR (Not FIsNum(iLine.FItemArray(2))) Then
					errCNT = errCNT + 1
					RetErr = RetErr & "��ǰ�ڵ�, Ư�� �� ���� ���� ���ڰ� �ֽ��ϴ�." & " \n"
				Else

					If (Not isDate(iLine.FItemArray(4))) OR (Not isDate(iLine.FItemArray(5))) Then
						errCNT = errCNT + 1
						RetErr = RetErr & "����/�������� ��¥ ������ �ƴմϴ�." & " \n"
					End If

                    if iLine.FItemArray(4)<>"" and not(isnull(iLine.FItemArray(4))) and iLine.FItemArray(5)<>"" and not(isnull(iLine.FItemArray(5))) then
						If CDate(iLine.FItemArray(4)) > CDate(iLine.FItemArray(5)) then
							errCNT = errCNT + 1
							RetErr = RetErr & "�Ⱓ �������� �Ⱓ �����Ϻ��� ���� �� �����ϴ�." & " \n"
						End If
					End If

					If iLine.FItemArray(0) = "nvstorefarm" Then
						If (Not isDate(iLine.FItemArray(6))) OR (Not isDate(iLine.FItemArray(7))) Then
							errCNT = errCNT + 1
							RetErr = RetErr & "���� ����/�������� ��¥ ������ �ƴմϴ�." & " \n"
						End If

						if iLine.FItemArray(6)<>"" and not(isnull(iLine.FItemArray(6))) and iLine.FItemArray(7)<>"" and not(isnull(iLine.FItemArray(7))) then
							If CDate(iLine.FItemArray(6)) > CDate(iLine.FItemArray(7)) then
								errCNT = errCNT + 1
								RetErr = RetErr & "���� �Ⱓ �������� �Ⱓ �����Ϻ��� ���� �� �����ϴ�." & " \n"
							End If
						End If
					End If

					If iLine.FItemArray(3) = "" Then
						strSql = ""
						strSql = strSql & " SELECT COUNT(*) as cnt "
						strSql = strSql & " FROM db_item.dbo.tbl_item "
						strSql = strSql & " WHERE itemid in ("& iLine.FItemArray(1) &") "
						strSql = strSql & " and mwdiv <> 'M' "
						rsget.CursorLocation = adUseClient
						rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
							cnt = rsget("cnt")
						rsget.Close

						If cnt > 0 Then
							errCNT = errCNT + 1
							RetErr = RetErr & "���� �ƴ� ��ǰ�� Ư�������� �Էµ��� �ʾҽ��ϴ�." & " \n"
						End If
					End If
				End If
			Else
				errCNT = errCNT + 1
				RetErr = RetErr & "�� ������ �� ���� �����Ͱ� �ֽ��ϴ�." & " \n"
			End If
		Set iLine = nothing
	End If
Next

If errCNT > 0 Then
	response.write "<script>alert('"&errCNT&"�� ����.\n\n"&RetErr&"');opener.location.reload();self.close();</script>"
	response.end
Else
	For i = 0 to pcnt
		If IsObject(xlRowALL(i)) Then
			Set iLine = xlRowALL(i)
				calcuBuyPrice = 0
				If iLine.FItemArray(3) <> "" Then
					calcuBuyPrice = Math_RoundOff(iLine.FItemArray(2) - (iLine.FItemArray(2) * (iLine.FItemArray(3) / 100)), 0)
				End If

				strSql = ""
				strSql = strSql & " IF EXISTS(SELECT TOP 1 itemid from db_etcmall.dbo.tbl_outmall_mustPriceItem WHERE itemid = '"& iLine.FItemArray(1) &"' and mallgubun = '"& iLine.FItemArray(0) &"' ) "
				strSql = strSql & " BEGIN "
				strSql = strSql & "     UPDATE db_etcmall.dbo.tbl_outmall_mustPriceItem SET "
				strSql = strSql & "     mustPrice = '"& iLine.FItemArray(2) &"' " & VBCRLF
				strSql = strSql & "     ,mustBuyPrice = '"& calcuBuyPrice &"' " & VBCRLF
				strSql = strSql & "     ,mustMargin = '"& iLine.FItemArray(3) &"' " & VBCRLF
				strSql = strSql & "     ,startDate = '"& iLine.FItemArray(4) &"' " & VBCRLF
				strSql = strSql & "     ,endDate = '"& iLine.FItemArray(5) &"' " & VBCRLF
				If iLine.FItemArray(0) = "nvstorefarm" Then
					strSql = strSql & " ,orgpricestartDate = '"& iLine.FItemArray(6) &"' " & VBCRLF
					strSql = strSql & " ,orgpriceendDate = '"& iLine.FItemArray(7) &"' " & VBCRLF
				End If
				strSql = strSql & "     ,lastUpdate = getdate() " & VBCRLF
				strSql = strSql & "     ,lastUpdateUserId = '"& session("ssBctID") &"' " & VBCRLF
				strSql = strSql & "     WHERE itemid = '"& iLine.FItemArray(1) &"' and mallgubun = '"& iLine.FItemArray(0) &"' " & VBCRLF
				strSql = strSql & " END ELSE "
				strSql = strSql & " BEGIN "
				If iLine.FItemArray(0) = "nvstorefarm" Then
					strSql = strSql & "     INSERT INTO db_etcmall.dbo.tbl_outmall_mustPriceItem (mallgubun, itemid, mustPrice, mustBuyPrice, mustMargin, startDate, endDate, orgpricestartDate, orgpriceendDate, regDate, regUserId) VALUES " & VBCRLF
					strSql = strSql & "     ('"& iLine.FItemArray(0) &"', '"& iLine.FItemArray(1) &"', '"& iLine.FItemArray(2) &"', '"& calcuBuyPrice &"', '"& iLine.FItemArray(3) &"', '"& iLine.FItemArray(4) &"', '"& iLine.FItemArray(5) &"', '"& iLine.FItemArray(6) &"', '"& iLine.FItemArray(7) &"', getdate(), '"&  session("ssBctID") &"' ) "
				Else
					strSql = strSql & "     INSERT INTO db_etcmall.dbo.tbl_outmall_mustPriceItem (mallgubun, itemid, mustPrice, mustBuyPrice, mustMargin, startDate, endDate, regDate, regUserId) VALUES " & VBCRLF
					strSql = strSql & "     ('"& iLine.FItemArray(0) &"', '"& iLine.FItemArray(1) &"', '"& iLine.FItemArray(2) &"', '"& calcuBuyPrice &"', '"& iLine.FItemArray(3) &"', '"& iLine.FItemArray(4) &"', '"& iLine.FItemArray(5) &"', getdate(), '"&  session("ssBctID") &"' ) "
				End If
				strSql = strSql & " END "
				dbget.Execute(strSql)
			Set iLine = nothing
		End If
	Next
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