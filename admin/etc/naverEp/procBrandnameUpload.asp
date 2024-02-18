<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/orderInput/incUploadFunction.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Function Math_RoundOff( su1, decimalPlaces)
    Dim sTemp, i, antilog, fraction '진수, 소수
    antilog = 1

    If decimalPlaces > 0 Then ' 10의 0승 이상 처리
        antilog = 10 ^ decimalPlaces
        sTemp = Fix( su1 / antilog + 0.5 ) * antilog
    Else ' 1 자리와 그 이하 처리
        sTemp = Round( su1 + 0.000001 , -(decimalPlaces))
    End if
    Math_RoundOff = sTemp
End Function

Dim uploadform, objfile, sDefaultPath, sFolderPath ,monthFolder
Dim sFile, sFilePath, iMaxLen, orgFileName, maybeSheetName
monthFolder = Replace(Left(CStr(now()),7),"-","")

'소숫점포함 숫자여부 체크
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
	sDefaultPath = Server.MapPath("/admin/etc/naverEp/upFiles/")
	uploadform.Start sDefaultPath '업로드경로
	iMaxLen = uploadform.Form("iML")	'이미지파일크기

	If (fnChkFile(uploadform("sFile"), iMaxLen,"xls")) Then	'파일체크
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
''''	0		1
''''상품코드	브랜드명
xlPosArr = Array(0,1)
ArrayLen = UBound(xlPosArr)
skipString = "Sheet1"
afile = sFilePath
aSheetName = ""

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
			If (Len(iLine.FItemArray(0)) >= 1) Then
				isMustPriceData = True
			End If

			If isMustPriceData = True Then
				iLine.FItemArray(0) = replace(iLine.FItemArray(0) ,",","")

				If (Not FIsNum(iLine.FItemArray(0))) Then
					errCNT = errCNT + 1
					RetErr = RetErr & "상품코드 중 숫자 외의 문자가 있습니다." & " \n"
				End If
			End If
		Set iLine = nothing
	End If
Next

If errCNT > 0 Then
	response.write "<script>alert('"&errCNT&"건 오류.\n\n"&RetErr&"');opener.location.reload();self.close();</script>"
	response.end
Else
	For i = 0 to pcnt
		If IsObject(xlRowALL(i)) Then
			Set iLine = xlRowALL(i)
				strSql = ""
				strSql = strSql & " IF EXISTS(SELECT TOP 1 itemid from db_outmall.dbo.tbl_EpShop_itemid_Socname WHERE itemid = '"& iLine.FItemArray(0) &"' and mallgubun = 'naverep' ) "
				strSql = strSql & " BEGIN "
				strSql = strSql & "     UPDATE db_outmall.dbo.tbl_EpShop_itemid_Socname SET "
				strSql = strSql & "     socname = '"& iLine.FItemArray(1) &"' " & VBCRLF
				strSql = strSql & "     ,socname_kor = '"& iLine.FItemArray(1) &"' " & VBCRLF
				strSql = strSql & "     ,lastUpdate = getdate() " & VBCRLF
				strSql = strSql & "     ,updateid = '"& session("ssBctID") &"' " & VBCRLF
				strSql = strSql & "     WHERE itemid = '"& iLine.FItemArray(0) &"' and mallgubun = 'naverep' " & VBCRLF
				strSql = strSql & " END ELSE "
				strSql = strSql & " BEGIN "
				strSql = strSql & "     INSERT INTO db_outmall.dbo.tbl_EpShop_itemid_Socname (itemid, mallgubun, socname, socname_kor, isusing, regdate, regid) VALUES " & VBCRLF
				strSql = strSql & "     ('"& iLine.FItemArray(0) &"', 'naverep', '"& iLine.FItemArray(1) &"', '"& iLine.FItemArray(1) &"', 'Y' ,getdate(), '"&session("ssBctID")&"') "
				strSql = strSql & " END "
				dbCTget.Execute(strSql)

				strSql = ""
				strSql = strSql & " UPDATE I SET lastupdate=getdate() from [db_AppWish].dbo.tbl_item I where I.itemid="&iLine.FItemArray(0)&""
				dbCTget.Execute(strSql)
			Set iLine = nothing
			If (i mod 1000) = 0 Then
				rw "----ing"
				response.flush
			End If
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
	'' on Error 구문 쓰면 안됨.. 서버 무한루프 도는듯.
	Set conDB = Server.CreateObject("ADODB.Connection")
		conDB.Provider = "Microsoft.Jet.oledb.4.0"
		'conDB.Provider = "Microsoft.ACE.OLEDB.12.0"
		conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;IMEX=1"  ''';IMEX=1 추가 2013/12/19
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
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->