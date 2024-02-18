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
Dim sCode, egCode, eCode
monthFolder = Replace(Left(CStr(now()),7),"-","")

'�Ҽ������� ���ڿ��� üũ
'-----------------------------------------
Function FIsNum(ByVal iValue)
    Dim iLength , i , retValue
    For i = 1 To Len(iValue)
     If Not (( asc(Mid(iValue,i,1)) > 47 And asc(Mid(iValue,i,1)) < 58 ) or asc(Mid(iValue,i,1))  = 46 ) Then
       FIsNum  = False
            Exit For
     else
        FIsNum = True
    end if
    Next


End Function
'-----------------------------------------
If (application("Svr_Info")	= "Dev") Then
    Set uploadform = Server.CreateObject("TABSUpload4.Upload")			'' - TEST : TABS.Upload
Else
    Set uploadform = Server.CreateObject("TABSUpload4.Upload")	''REAL : TABSUpload4.Upload
End If

Set objfile = Server.CreateObject("Scripting.FileSystemObject")
	sDefaultPath = Server.MapPath("/admin/etc/kakaogift/upFiles/")
	uploadform.Start sDefaultPath '���ε���

	iMaxLen = uploadform.Form("iML")	'�̹�������ũ��
	sCode	= uploadform.Form("sC")		'���� �ڵ�
	eCode	= uploadform.Form("eC")		'�̺�Ʈ �ڵ�
	egCode	= uploadform.Form("egC")	'�̺�Ʈ �׷� �ڵ�

	If (fnChkFile(uploadform("sFile"), iMaxLen,"xlsx")) Then	'����üũ
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
''''	0			1			2			3				4
''''ä�λ�ǰ��ȣ, 	�Ǹ��ڸ�, 	��ǰ��,		�Ǹ��ڻ�ǰ��ȣ	�ǸŻ���
''''	5			6			7			8				9				10
''''�ɼ�,			�ǸŰ�,	���� ���ΰ�		ä�� ������		�Ǹ� ������		�Ǹ� ������


xlPosArr = Array(5, 3, 3, 11, 6, 12, 40, 18, 19, 36, 42, 43)
ArrayLen = UBound(xlPosArr)
skipString = "product"
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
Dim errCNT, objCmd, strStatus, errItem
Dim pcnt : pcnt = UBound(xlRowALL)
Dim itemArr, itemidarr, arrList, intLoop
Dim dSDate, dEDate, iSaleRate, iSaleMargin, iSaleMarginValue, saleStatus
errCNT = 0
addSqlDB = ""
addSql = ""
'1.tmp���̺� ���� �ִ� �״�� ���
strSql = ""
strSql = strSql & " TRUNCATE TABLE db_temp.dbo.Tbl_kakaotemp "
dbget.execute strSql
response.flush
response.clear
'2. ���ε�� db_temp.dbo.Tbl_kakaotemp �̰��� �ϸ� truncate�� insert �Ѵ�
For i = 0 to pcnt
	If IsObject(xlRowALL(i)) Then
		Set iLine = xlRowALL(i)
			If iLine.FItemArray(5) = "�Ǹű���" Then
				iLine.FItemArray(5) = "�Ǹ�����" 
			End If
			strSql = ""
			strSql = strSql & " INSERT INTO db_temp.dbo.Tbl_kakaotemp (KakaoNo, chanelname, EventName, ProductName, SellNumber, SellState, options, NormalPrice, SellPrice, SusuPrice, SellStart, SellEnd) "
			strSql = strSql & " VALUES ('"&html2db(iLine.FItemArray(0))&"', '"&html2db(iLine.FItemArray(1))&"', '"&html2db(iLine.FItemArray(1))&"', '"&html2db(iLine.FItemArray(3))&"', '"&html2db(iLine.FItemArray(4))&"', "
			strSql = strSql & " '"&html2db(iLine.FItemArray(5))&"', '"&html2db(iLine.FItemArray(6))&"', '"&html2db(iLine.FItemArray(7))&"', '"&html2db(iLine.FItemArray(8))&"', '"&html2db(iLine.FItemArray(9))&"', '"&html2db(iLine.FItemArray(10))&"', '"&html2db(iLine.FItemArray(11))&"') "
			dbget.execute strSql
		Set iLine = xlRowALL(i)
	End If
Next
response.flush
response.clear
'3. insert�� ������ db_etcmall.dbo.usp_API_kakaoregitem_TempSync �����Ѵ�
strSql = ""
strSql = strSql & "EXEC db_etcmall.dbo.usp_API_kakaoregitem_TempSync "
dbget.execute strSql
rw "OK"

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
		conDB.Provider = "Microsoft.ACE.OLEDB.12.0"
		conDB.Properties("ExtEnded Properties").Value = "Excel 12.0;IMEX=1"
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