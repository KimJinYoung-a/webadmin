<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/orderInput/incUploadFunction.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
''������
1=a 

Dim monthFolder : monthFolder = request("monthFolder")
Dim subfldr : subfldr = request("subfldr")
Dim filename : filename = request("filename")


Dim sDefaultPath : sDefaultPath = Server.MapPath("/admin/etc/naverEp/upFiles/"&monthFolder&"/"&subfldr&"/")
Dim maybeSheetName : maybeSheetName = replace(filename,".xlsx","")
Dim sFilePath : sFilePath = sDefaultPath&"\"&filename

''rw subfldr&"\"&monthFolder&"\"&filename


Dim xlPosArr, ArrayLen, skipString, afile, aSheetName ,i
''''    =======================================================
''''	0			1			2					
''''    ��ǰ��(TEN), 	���̹����� ī�װ�, 	���θ� ī�װ�
''''	3				4				5			6				
''''    ���̹����� ��ǰID,	���θ� ��ǰID,	��ǰ����,	�ؿܿ���
''''	7			8				9				10
''''	�ǸŹ��,	�귣��(���̹�����),	������(���̹�����),	�ǸŰ�
''''	11			12				13				14
''''	�������,	���ݺ񱳸�,	���ݺ� ������	,���� ��ǰ ����
''''	15				16	
''''	���ݺ� ��û����,	���ݺ� ID
''''    =======================================================

xlPosArr = Array(0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16)
ArrayLen = UBound(xlPosArr)
skipString = "Sheet1"
afile = sFilePath
aSheetName = ""

'rw sFilePath&":"&aSheetName&":"
'response.end


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
Dim intResult, iretVal, iAssignedTTL

errCNT = 0
addSqlDB = ""
addSql = ""
For i = 0 to pcnt
	If IsObject(xlRowALL(i)) Then
		Set iLine = xlRowALL(i)
			' If Right(iLine.FItemArray(2), 1) = "%" Then
			' 	iLine.FItemArray(2) = Left(iLine.FItemArray(2), Len(iLine.FItemArray(2)) - 1)
			' End If

			' If Right(iLine.FItemArray(3), 1) = "%" Then
			' 	iLine.FItemArray(3) = Left(iLine.FItemArray(3), Len(iLine.FItemArray(3)) - 1)
			' End If

			' If iLine.FItemArray(1) = "" Then
			' 	iLine.FItemArray(1) = 0
			' ElseIf iLine.FItemArray(2) = "" Then
			' 	iLine.FItemArray(2) = 0
			' End If

''��ǰ��(0),���̹����� ī�װ�(1)���θ� ī�װ�(2)���̹����� ��ǰID(3)���θ� ��ǰID(4)
''��ǰ����(5)-�Ϲ�	�ؿܿ���(6)	�ǸŹ��(7)	�귣��(���̹�����)(8)	������(���̹�����)(9)
''�ǸŰ�(10) �������(11) ���ݺ񱳸�(12) ���ݺ� ������(13)	���� ��ǰ ����(14)���ݺ񱳸�Ī�Ϸ�,���ݺ� ��û����(15)-ó���Ϸ�,���ݺ� ID(16)

			Dim cmd : set cmd = server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = dbget
			cmd.CommandText = "db_etcmall.dbo.usp_ten_AddNvMapItem_BULKROW"
			cmd.CommandType = adCmdStoredProc

			cmd.Parameters.Append cmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
			cmd.Parameters.Append cmd.CreateParameter("@tenitemid", adVarchar, adParamInput, 10, iLine.FItemArray(4))  ''TEN ��ǰ�ڵ�
			cmd.Parameters.Append cmd.CreateParameter("@matchNvMid", adVarchar, adParamInput, 16, iLine.FItemArray(16))	''���ݺ� ID
			cmd.Parameters.Append cmd.CreateParameter("@nvitemid", adVarchar, adParamInput, 16, iLine.FItemArray(3))	'NV ��ǰ�ڵ�
			cmd.Parameters.Append cmd.CreateParameter("@minprice", adCurrency, adParamInput, , iLine.FItemArray(13))	''���ݺ� ������
			cmd.Parameters.Append cmd.CreateParameter("@nvbrandName", adVarWchar, adParamInput, 80, iLine.FItemArray(8))	''�귣��(���̹�����)
			cmd.Parameters.Append cmd.CreateParameter("@nvmakername", adVarWchar, adParamInput, 80, iLine.FItemArray(9))	''������(���̹�����)
			cmd.Parameters.Append cmd.CreateParameter("@nvcatename", adVarWchar, adParamInput, 300, iLine.FItemArray(1))	''���̹����� ī�װ�(1)
			cmd.Parameters.Append cmd.CreateParameter("@nvitemname", adVarWchar, adParamInput, 300, iLine.FItemArray(12))	''���ݺ񱳸�
			''cmd.Parameters.Append cmd.CreateParameter("@retVal", adVarWchar, adParamOutput, 384, "")
			cmd.Execute
'	rw  iLine.FItemArray(4) &":"&iLine.FItemArray(16) 		
			intResult = cmd.Parameters("returnValue").Value
			''iretVal = cmd.Parameters("@retVal").Value
			
			set cmd = Nothing
			
			if (intResult>0) then
				iAssignedTTL = iAssignedTTL + 1
			end if

			' If (Not FIsNum(iLine.FItemArray(0))) OR (Not FIsNum(iLine.FItemArray(1))) OR (Not FIsNum(iLine.FItemArray(2))) OR (Not FIsNum(iLine.FItemArray(3))) Then
			' 	errCNT = errCNT + 1
			' 	RetErr = RetErr & "���� ������ �� ���ڿ��� ���ڰ� �ֽ��ϴ�." & " \n"

			' else
			' 	if 	iLine.FItemArray(2) >99 then
			' 		errCNT = errCNT + 1
			' 		RetErr = RetErr & "��ǰ�ڵ� : " & iLine.FItemArray(0) & " �������� 100%�� �ѱ� �� �����ϴ�." & " \n"
			' 	end if
			'     If iLine.FItemArray(3) < 0 Then
			' 			errCNT = errCNT + 1
			' 			RetErr = RetErr & "��ǰ�ڵ� : " & iLine.FItemArray(0) & "������ ��ǰ�Դϴ�." & " \n"

			'     End If
			'      dim sMargin
			'     strSql = "select sale_margin from db_Event.dbo.tbl_sale where sale_code = "&sCode
			'      rsget.Open strSql,dbget,1
			' 		 if not rsget.eof Then
			' 		 	sMargin = rsget(0)
			' 		end if
			' 		rsget.close

			'     If iLine.FItemArray(3) < 1 and sMargin <> 4 Then
			' 			errCNT = errCNT + 1
			' 			RetErr = RetErr & "��ǰ�ڵ� : " & iLine.FItemArray(0) & " ���θ������� 1% �̸��� ��ǰ�Դϴ�. " & " \n"

			'     End If
			' End If

			''itemArr				= itemArr & iLine.FItemArray(0) & ","			'��ǰ�ڵ�
		Set iLine = xlRowALL(i)
	End If
Next
response.write maybeSheetName&"::"
response.write "����Ǽ�:"&iAssignedTTL&"/"&pcnt&"��"

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

    If  LCASE(ixlRow(0)) = "��ǰ��" Then
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