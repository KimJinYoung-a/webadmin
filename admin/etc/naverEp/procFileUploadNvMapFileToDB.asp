<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/orderInput/incUploadFunction.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
''사용안함
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
''''    상품명(TEN), 	네이버쇼핑 카테고리, 	쇼핑몰 카테고리
''''	3				4				5			6				
''''    네이버쇼핑 상품ID,	쇼핑몰 상품ID,	상품상태,	해외여부
''''	7			8				9				10
''''	판매방식,	브랜드(네이버쇼핑),	제조사(네이버쇼핑),	판매가
''''	11			12				13				14
''''	등록일자,	가격비교명,	가격비교 최저가	,서비스 상품 상태
''''	15				16	
''''	가격비교 요청상태,	가격비교 ID
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
	response.write "<script>alert('파일이 올바르지 않거나 내용이 없습니다. "&Replace(Err.Description,"'","")&"');</script>"

	If (Err.Description="외부 테이블 형식이 잘못되었습니다.") Then
		response.write "<script>alert('엑셀에서 Save As Excel 97 -2003 통합문서 형태로 저장후 사용하세요.');</script>"
	End If
	response.write "<script>history.back();</script>"
	response.end
End If

''데이터 처리.
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

''상품명(0),네이버쇼핑 카테고리(1)쇼핑몰 카테고리(2)네이버쇼핑 상품ID(3)쇼핑몰 상품ID(4)
''상품상태(5)-일반	해외여부(6)	판매방식(7)	브랜드(네이버쇼핑)(8)	제조사(네이버쇼핑)(9)
''판매가(10) 등록일자(11) 가격비교명(12) 가격비교 최저가(13)	서비스 상품 상태(14)가격비교매칭완료,가격비교 요청상태(15)-처리완료,가격비교 ID(16)

			Dim cmd : set cmd = server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = dbget
			cmd.CommandText = "db_etcmall.dbo.usp_ten_AddNvMapItem_BULKROW"
			cmd.CommandType = adCmdStoredProc

			cmd.Parameters.Append cmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
			cmd.Parameters.Append cmd.CreateParameter("@tenitemid", adVarchar, adParamInput, 10, iLine.FItemArray(4))  ''TEN 상품코드
			cmd.Parameters.Append cmd.CreateParameter("@matchNvMid", adVarchar, adParamInput, 16, iLine.FItemArray(16))	''가격비교 ID
			cmd.Parameters.Append cmd.CreateParameter("@nvitemid", adVarchar, adParamInput, 16, iLine.FItemArray(3))	'NV 상품코드
			cmd.Parameters.Append cmd.CreateParameter("@minprice", adCurrency, adParamInput, , iLine.FItemArray(13))	''가격비교 최저가
			cmd.Parameters.Append cmd.CreateParameter("@nvbrandName", adVarWchar, adParamInput, 80, iLine.FItemArray(8))	''브랜드(네이버쇼핑)
			cmd.Parameters.Append cmd.CreateParameter("@nvmakername", adVarWchar, adParamInput, 80, iLine.FItemArray(9))	''제조사(네이버쇼핑)
			cmd.Parameters.Append cmd.CreateParameter("@nvcatename", adVarWchar, adParamInput, 300, iLine.FItemArray(1))	''네이버쇼핑 카테고리(1)
			cmd.Parameters.Append cmd.CreateParameter("@nvitemname", adVarWchar, adParamInput, 300, iLine.FItemArray(12))	''가격비교명
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
			' 	RetErr = RetErr & "엑셀 데이터 중 숫자외의 문자가 있습니다." & " \n"

			' else
			' 	if 	iLine.FItemArray(2) >99 then
			' 		errCNT = errCNT + 1
			' 		RetErr = RetErr & "상품코드 : " & iLine.FItemArray(0) & " 할인율은 100%를 넘길 수 없습니다." & " \n"
			' 	end if
			'     If iLine.FItemArray(3) < 0 Then
			' 			errCNT = errCNT + 1
			' 			RetErr = RetErr & "상품코드 : " & iLine.FItemArray(0) & "역마진 상품입니다." & " \n"

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
			' 			RetErr = RetErr & "상품코드 : " & iLine.FItemArray(0) & " 할인마진율이 1% 미만인 상품입니다. " & " \n"

			'     End If
			' End If

			''itemArr				= itemArr & iLine.FItemArray(0) & ","			'상품코드
		Set iLine = xlRowALL(i)
	End If
Next
response.write maybeSheetName&"::"
response.write "적용건수:"&iAssignedTTL&"/"&pcnt&"건"

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

    If  LCASE(ixlRow(0)) = "상품명" Then
		IsSKipRow = true
		Exit Function
	End If

	IsSKipRow = false
End Function 

Function fnGetXLFileArray(byref xlRowALL, sFilePath, aSheetName, iArrayLen)
	Dim conDB, Rs, strQry, iResult, i, J, iObj
	Dim irowObj, strTable
	'' on Error 구문 쓰면 안됨.. 서버 무한루프 도는듯.
	Set conDB = Server.CreateObject("ADODB.Connection")
		'conDB.Provider = "Microsoft.Jet.oledb.4.0"
		conDB.Provider = "Microsoft.ace.oledb.12.0"
		'conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;IMEX=1"  ''';IMEX=1 추가 2013/12/19
		conDB.Properties("ExtEnded Properties").Value = "Excel 12.0;IMEX=1"  ''';IMEX=1 추가 2013/12/19
	'On Error Resume Next
		conDB.Open sFilePath
		If (ERR) Then
			fnGetXLFileArray=false
			'/이유를 알수 없는 서버단 에러남. "예기치 않은 오류. 외부 개체에 트랩 가능한 오류(C0000005)가 발생했습니다. 스크립트를 계속 실행할 수 없습니다"
			set conDB = nothing
			Exit Function
		End If
	'On Error Goto 0
		'' get First Sheet Name=============''시트가 여러개인경우 오류날 수 있음.
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