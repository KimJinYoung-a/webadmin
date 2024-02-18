<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/orderInput/incUploadFunction.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim uploadform, objfile, sDefaultPath, sFolderPath ,monthFolder
Dim sFile, sFilePath, iMaxLen, orgFileName, maybeSheetName
Dim sCode, egCode, eCode,adminid, iSalePercent, chgSalePercent
monthFolder = Replace(Left(CStr(now()),7),"-","")
adminid  = session("ssBctId")
'소숫점포함 숫자여부 체크
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
    Set uploadform = Server.CreateObject("TABS.Upload")			'' - TEST : TABS.Upload
Else
    Set uploadform = Server.CreateObject("TABSUpload4.Upload")	''REAL : TABSUpload4.Upload
End If

Set objfile = Server.CreateObject("Scripting.FileSystemObject")
	sDefaultPath = Server.MapPath("/admin/shopmaster/sale/upFiles/")
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
''''	0			1				2
''''상품코드, 	할인판매가, 	할인매입가
xlPosArr = Array(0,1,2,3)
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
Dim iLine, iResult, strSql, addSqlDB, addSql, RetErr, RetErr2, isOK
Dim errCNT, objCmd, strStatus, errItem
Dim pcnt : pcnt = UBound(xlRowALL)
Dim itemArr, itemidarr, arrList, intLoop
Dim dSDate, dEDate, iSaleRate, iSaleMargin, iSaleMarginValue, saleStatus
errCNT = 0
addSqlDB = ""
addSql = ""
isOK = "Y"
For i = 0 to pcnt
	If IsObject(xlRowALL(i)) Then
		Set iLine = xlRowALL(i)
		if getNumeric(iLine.FItemArray(0))<>"" then
			If iLine.FItemArray(1) = "" Then
				iLine.FItemArray(1) = 0
			ElseIf iLine.FItemArray(2) = "" Then
				iLine.FItemArray(2) = 0
			End If

			iLine.FItemArray(0) = round(iLine.FItemArray(0))
			iLine.FItemArray(1) = round(iLine.FItemArray(1))
			iLine.FItemArray(2) = round(iLine.FItemArray(2))


			If (Not FIsNum(iLine.FItemArray(0))) OR (Not FIsNum(iLine.FItemArray(1))) OR (Not FIsNum(iLine.FItemArray(2)))   Then
				errCNT = errCNT + 1
				RetErr2 = RetErr2 & "[" & i+1 & "행] 엑셀 데이터 중 숫자외의 문자가 있습니다.<br />"
			End If

			itemArr				= itemArr & iLine.FItemArray(0) & ","			'상품코드
		end if
		Set iLine = xlRowALL(i)
	End If
Next

If errCNT > 0 Then
	response.write "ERROR : 등록되지 않았습니다. <br />"
	response.write RetErr2
	dbget.close() : response.end
End If

If Right(itemArr,1) = "," Then
	itemidarr = Left(itemArr, Len(itemArr) - 1)
End If

'1.tmp테이블에 내용 있는 그대로 등록
strSql = "create table #tmpATSale(itemid int, sellprice money, buyprice money, errCodes varchar(150))"
dbget.execute strSql

strSql = ""
For i = 0 to pcnt
	If IsObject(xlRowALL(i)) Then
		Set iLine = xlRowALL(i)
		if getNumeric(iLine.FItemArray(0))<>"" then
			If Right(int(iLine.FItemArray(1)),1) <> "0" Then			'할인 판매가의 원단위 절삭위해 분기
				iLine.FItemArray(1) = Left(iLine.FItemArray(1),Len(iLine.FItemArray(1))-1)&"0"
			End If
			strSql = strSql & " IF NOT EXISTS (Select itemid from #tmpATSale where itemid='"&iLine.FItemArray(0)&"') BEGIN INSERT INTO #tmpATSale (itemid, sellprice, buyprice, errCodes) values ('"&iLine.FItemArray(0)&"', '"& int(iLine.FItemArray(1)) &"', '"&iLine.FItemArray(2)&"', '') END "
		end if
		Set iLine = xlRowALL(i)
	End If
Next
dbget.execute strSql

Dim errMarginArr
strSql = " select itemid, ((sellprice - buyprice) / sellprice) * 100 as margin "
strSql = strSql & " FROM #tmpATSale"
rsget.CursorLocation = adUseClient
rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
If Not rsget.Eof Then
	errMarginArr = rsget.getRows()
End If
rsget.Close

If IsArray(errMarginArr) Then
	For intLoop = 0 To UBound(errMarginArr,2)
		If errMarginArr(1, intLoop) < 0 Then
			strSql = ""
			strSql = strSql & " UPDATE #tmpATSale "
			strSql = strSql & " SET errCodes = errCodes + '0|' "
			strSql = strSql & " WHERE itemid = '"& errMarginArr(0,intLoop) &"' "
			dbget.execute strSql
			isOK = "N"
		ElseIf errMarginArr(1, intLoop) < 1 Then
			strSql = ""
			strSql = strSql & " UPDATE #tmpATSale "
			strSql = strSql & " SET errCodes = errCodes + '1|' "
			strSql = strSql & " WHERE itemid = '"& errMarginArr(0,intLoop) &"' "
			dbget.execute strSql
			isOK = "N"
		End If
	Next
End If

'' 2018/08/14 여기서 막을걸 좀 막자.. (이버전도 안되면 중간 페이지를 하나 만드는게 맞지 않을까..?)
''CASE 1 상품코드를 잘못등록하는 케이스가 많음. => 동일브랜드끼리만 수정 가능하게 하자.
dim MultipleMakerArr
strSql = " select i.makerid,count(*) CNT from #tmpATSale T"
strSql = strSql & " Join db_item.dbo.tbl_item i"
strSql = strSql & " on T.itemid=i.itemid"
strSql = strSql & " group by i.makerid"
rsget.CursorLocation = adUseClient
rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
if Not rsget.Eof then
	MultipleMakerArr = rsget.getRows()
end if
rsget.Close

if IsArray(MultipleMakerArr) then
	if UBound(MultipleMakerArr,2)>0 then
		strSql = ""
		strSql = strSql & " UPDATE #tmpATSale "
		strSql = strSql & " SET errCodes = errCodes + '2|' "
		dbget.execute strSql
		isOK = "N"
	end if
else
	response.write "<script>alert('Oops! 등록될 상품이 없습니다. 상품코드가 올바르지 않거나 내역이 없습니다.')</script>"
	strSql = "drop table #tmpATSale "
	dbget.execute strSql
	dbget.close() : response.end
end if

''CASE 2 소비자가보다 할인이 높은경우
dim SalePriceHighArr
strSql = " select T.itemid, T.sellprice, i.orgprice from #tmpATSale T"
strSql = strSql & " Join db_item.dbo.tbl_item i"
strSql = strSql & " on T.itemid=i.itemid"
strSql = strSql & " where T.sellprice>i.orgprice"
rsget.CursorLocation = adUseClient
rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
if Not rsget.Eof then
	SalePriceHighArr = rsget.getRows()
end if
rsget.Close

if IsArray(SalePriceHighArr) then
	FOR intLoop = 0 TO UBound(SalePriceHighArr,2)
		strSql = ""
		strSql = strSql & " UPDATE #tmpATSale "
		strSql = strSql & " SET errCodes = errCodes + '3|' "
		strSql = strSql & " WHERE itemid = '"& SalePriceHighArr(0,intLoop) &"' "
		dbget.execute strSql
		isOK = "N"
	Next
end if

''CASE 3 기존판매가보다 가격이 N% 이상 할인될경우 - 노티만
dim SalePriceLowArr
strSql = " select T.itemid, T.sellprice, i.sellcash from #tmpATSale T"
strSql = strSql & " Join db_item.dbo.tbl_item i"
strSql = strSql & " on T.itemid=i.itemid"
strSql = strSql & " where i.sellcash<>0"
strSql = strSql & " and (100-T.sellprice/sellcash*100>88)"
rsget.CursorLocation = adUseClient
rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
if Not rsget.Eof then
	SalePriceLowArr = rsget.getRows()
end if
rsget.Close

if IsArray(SalePriceLowArr) then
	FOR intLoop = 0 TO UBound(SalePriceLowArr,2)
		strSql = ""
		strSql = strSql & " UPDATE #tmpATSale "
		strSql = strSql & " SET errCodes = errCodes + '4|' "
		strSql = strSql & " WHERE itemid = '"& SalePriceLowArr(0,intLoop) &"' "
		dbget.execute strSql
		isOK = "N"
	Next
end if
'' CASE END

''CASE 4 원매입가보다 할인매입가가 높을 경우(마진하락)
dim buyPriceOverArr
strSql = " select T.itemid, T.buyprice, i.orgsuplycash from #tmpATSale T"
strSql = strSql & " Join db_item.dbo.tbl_item i"
strSql = strSql & " on T.itemid=i.itemid"
strSql = strSql & " where i.orgsuplycash<T.buyprice "
rsget.CursorLocation = adUseClient
rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
if Not rsget.Eof then
	buyPriceOverArr = rsget.getRows()
end if
rsget.Close

if IsArray(buyPriceOverArr) then
	FOR intLoop = 0 TO UBound(buyPriceOverArr,2)
		strSql = ""
		strSql = strSql & " UPDATE #tmpATSale "
		strSql = strSql & " SET errCodes = errCodes + '5|' "
		strSql = strSql & " WHERE itemid = '"& buyPriceOverArr(0,intLoop) &"' "
		dbget.execute strSql
		isOK = "N"
	Next
end if
'' CASE END

strSql = ""
strSql = strSql & " SELECT i.itemid, i.makerid, i.itemname, i.smallimage ,i.sailyn,i.sellcash, i.buycash,i.orgprice, i.orgsuplycash, i.sailprice, i.sailsuplycash, i.mwdiv,i.limityn,i.limitno, i.limitsold,i.isusing   "
strSql = strSql & " ,i.itemCouponyn, i.itemCoupontype, i.itemCouponvalue"
strSql = strSql & " ,  Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail with(noLock) Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
strSql = strSql & " , st.sale_code, st.itemid as atItemid , st.orgsailprice as atorgsp, st.orgsailsuplycash as atorgsc, st.orgsailyn as atorgsyn"
strSql = strSql & ", i.lastupdate, T.sellprice, T.buyprice, T.errCodes "
strSql = strSql & " FROM #tmpATSale as T "
strSql = strSql & " JOIN db_item.dbo.tbl_item as i with(noLock) on T.itemid = i.itemid "
strSql = strSql & "   left outer join "
strSql = strSql & " 	  (	select si.itemid "
strSql = strSql & " 	  	, min(s.sale_code) sale_code "
strSql = strSql & " 	  	, min(si.orgsailprice) as orgsailprice "
strSql = strSql & " 	  	, min(si.orgsailsuplycash) as orgsailsuplycash "
strSql = strSql & " 	  	, min(si.orgsailyn) as orgsailyn "
strSql = strSql & " 			from db_event.dbo.tbl_sale as s with(noLock)  "
strSql = strSql & " 			inner join 	db_event.dbo.tbl_saleitem as si with(noLock) on s.sale_code = si.sale_code  "
strSql = strSql & " 			where s.sale_status = 6  "
strSql = strSql & " 				and si.saleItem_status = 6  "
strSql = strSql & " 				and s.sale_using =1  "
strSql = strSql & " 				and s.sale_startdate<=convert(varchar(10),getdate(),121) and s.sale_enddate >=convert(varchar(10),getdate(),121) "
strSql = strSql & "			group by si.itemid "
strSql = strSql & " 	  ) as st on i.itemid = st.itemid "
strSql = strSql & "  where i.isusing ='Y' and i.itemid <> 0    "
rsget.Open strSql,dbget
IF not rsget.EOF THEN
	arrList = rsget.getRows()
END IF
rsget.close

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
		'conDB.Provider = "MicrosofACE.OLEDB.12.0"
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

						For i=0 to ArrayLen-1
							If (xlPosArr(i) < 0) Then
								irowObj.FItemArray(i) = ""
							Else
								irowObj.FItemArray(i) = Replace(null2blank(Rs(xlPosArr(i))),"*","")
							End If
							'' rw irowObj.FItemArray(i)

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
<script type="text/javascript">
function confirmReg(){
	if(confirm("<%= RetErr %>저장하시겠습니까?")) {
		document.frmSvArr.target = "xLink";
		document.frmSvArr.action = "excelConfirm_Process.asp"
		document.frmSvArr.submit();
	}
}
</script>
※ 비고란에 코드가 적혀 있다면 저장 버튼이 활성화 되지 않습니다. 이하는 코드 정리 입니다.<br />
※ 비고 0 : 역마진 상품 <br />
※ 비고 1 : 할인 마진율 1% 미만 <br />
※ 비고 2 : 업로드 엑셀에 브랜드가 2개 이상 / 단일 브랜드만 업로드 가능<br />
※ 비고 3 : 할인가가 소비자가 보다 높음<br />
※ 비고 4 : 판매가보다 할인율이 88% 이상에 해당<br />
※ 비고 5 : 할인매입가가 계약매입가보다 높음
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method = "POST" onSubmit="return false;">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">상품ID</td>
	<td width="60">이미지</td>
	<td width="100">브랜드</td>
	<td>상품명</td>
	<td width="100">계약구분</td>
	<td width="100">할인상태</td>
	<td width="200">현재판매가</td>
	<td width="100">현재매입가</td>
	<td width="100">현재마진율</td>
	<td width="100">현재할인율</td>
	<td width="200">변경판매가</td>
	<td width="100">변경매입가</td>
	<td width="100">변경할인율</td>
	<td width="200">비고</td>
</tr>
<%
If isArray(arrList) Then
	Dim tmpArrItemid
	For i =0 To UBound(arrList,2)
		if (arrList(7, i)=0) then
			iSalePercent = 0
			chgSalePercent=0
		else
			iSalePercent = (1-(clng(arrList(9, i))/clng(arrList(7, i))))*100
			chgSalePercent = (1-(clng(arrList(26, i))/clng(arrList(7, i))))*100
		end if

		If (arrList(9, i)=0) Then
			iSalePercent = 0
		End If

		tmpArrItemid = tmpArrItemid & arrList(0, i) & ","

		If Right(arrList(28, i),1) = "|" Then
			arrList(28, i) = Left(arrList(28, i), Len(arrList(28, i)) - 1)
		End If
%>
<tr align="center" bgcolor="#FFFFFF">
	<input type="hidden" name="itemid" value="<%= arrList(0, i) %>">
	<input type="hidden" name="salePrice_<%= arrList(0, i) %>" value="<%= arrList(26, i) %>">
	<input type="hidden" name="saleBuyPrice_<%= arrList(0, i) %>" value="<%= arrList(27, i) %>">
	<td><%= arrList(0, i) %></td>
	<td><img src="http://webimage.10x10.co.kr/image/small/<%= GetImageSubFolderByItemid(arrList(0, i)) %>/<%= arrList(3, i) %>" width="50" ></td>

	<td><%= arrList(1, i) %></td>
	<td><%= arrList(2, i) %></td>
	<td><%=fnColor(arrList(11, i),"mw") %></td>
	<td><%=fnColor(arrList(4, i),"yn")%></td>
	<td><%=FormatNumber(arrList(7, i),0)%>
		<% 		'할인가(할인율=(소비자가-할인가)/소비자가*100)
		if arrList(4, i) ="Y" then %>
			<% if (arrList(7, i)<>0) then %>
		<br><font color=#F08050>(<%=CLng((arrList(7, i)-arrList(9, i))/arrList(7, i)*100) %>%할)<%=FormatNumber(arrList(9, i),0)%></font>
			<% end if %>
		<% end if %>
		<%'쿠폰가
		if arrList(16, i)="Y" then

			Select Case arrList(17, i)
				Case "1" '% 쿠폰
		%>
			<br><font color=#5080F0>(쿠)<%=FormatNumber(arrList(5, i)-(CLng(arrList(18, i)*arrList(5, i)/100)),0)%></font>
		<%
				Case "2" '원 쿠폰
		%>
			<br><font color=#5080F0>(쿠)<%=FormatNumber(arrList(5, i),0)%></font>
		<%
			end Select
		end if
		%>
	</td>
	<td><%=FormatNumber(arrList(8, i),0)%>
		<% '할인가
			if arrList(4, i) ="Y" then
		%>
				<br><font color=#F08050><%=FormatNumber(arrList(10, i),0) %></font>
		<%
			end if
			'쿠폰가
			if  arrList(16, i)="Y" then
				if arrList(17, i)="1" or arrList(17, i)="2" then
					if  arrList(19, i)=0 or isNull(arrList(19, i)) then
						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(6, i),0) & "</font>"
					else
						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(19, i),0) & "</font>"
					end if
				end if
			end if
	%>
	</td>
	<td>
		<%=fnPercent(arrList(8, i),arrList(7, i),1)%>
		<%
			'할인가
			if arrList(4, i) ="Y"  then
				Response.Write "<br><font color=#F08050>" & fnPercent(arrList(10, i),arrList(9, i),1) & "</font>"
			end if
			'쿠폰가
			if arrList(16, i)="Y" then
				Select Case  arrList(17, i)
					Case "1"
						if arrList(19, i)=0 or isNull(arrList(19, i)) then
							Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(6, i),arrList(5, i)-(CLng(arrList(18, i)*arrList(5, i)/100)),1) & "</font>"
						else
							Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(19, i),arrList(5, i)-(CLng(arrList(18, i)*arrList(5, i)/100)),1) & "</font>"
						end if
					Case "2"
						if arrList(19, i)=0 or isNull(arrList(19, i)) then
							Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(6, i),arrList(5, i),1) & "</font>"
						else
							Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(19, i),arrList(5, i),1) & "</font>"
						end if
				end Select
		end if
	%>
	</td>
	<td id="lyrSpct<%=arrList(0, i)%>" style="<%=chkIIF(iSalePercent>=50,"color:#EE0000;font-weight:bold;","")%>"><%=formatnumber(iSalePercent,0)%> %</td>
	<td><%=FormatNumber(arrList(26, i),0)%></td>
	<td><%=FormatNumber(arrList(27, i),0)%></td>
	<td style="<%=chkIIF(chgSalePercent>=50,"color:#EE0000;font-weight:bold;","")%>"><%=formatnumber(chgSalePercent,0)%> %</td>
	<td>
		<%
			response.write arrList(28, i)
		%>
	</td>
</tr>
<%
	Next
%>
<input type="hidden" name="tmpArrItemid" value="<%= tmpArrItemid %>" >
</form>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan="15">
	    <input type="button" class="button" value="등록" onClick="confirmReg();" <%= Chkiif(isOK = "N", "disabled", "") %> >
	</td>
</tr>
<%
End If
%>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="0"></iframe>
<!-- #include virtual="/lib/db/dbclose.asp" -->