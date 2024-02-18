<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<%
dim taxIdx : taxIdx = RequestCheckVar(request("taxIdx"),10)
dim result : result = RequestCheckVar(request("result"),32)
dim no_tax : no_tax = RequestCheckVar(request("no_tax"),32)
dim result_msg  : result_msg = RequestCheckVar(request("result_msg"),100)
dim write_date :  write_date = RequestCheckVar(request("write_date"),10)
dim no_iss     :  no_iss = RequestCheckVar(request("no_iss"),24)
dim strSql, sqlStr

'response.write "result: " & result & "<Br>"
'response.write "taxIdx: " & taxIdx & "<Br>"
'response.write "no_tax: " & no_tax & "<Br>"
'response.write "result_msg: " & result_msg & "<Br>"
'response.write "write_date: " & write_date & "<Br>"
'response.write "no_iss: " & no_iss & "<Br>"
dim oTax


set oTax = new CTax
oTax.FRecttaxIdx = taxIdx

oTax.GetTaxRead



if result="00000" then
    '데이터 처리
    strSql = " Update A  " & VbCrlf
    strSql = strSql & " SET confirmYn='Y'" & VbCrlf
    strSql = strSql & " from db_order.[dbo].tbl_busiinfo A " & VbCrlf
    strSql = strSql & "   Join db_order.[dbo].tbl_taxSheet S " & VbCrlf
    strSql = strSql & "   on A.busiIdx=S.busiIdx" & VbCrlf
    strSql = strSql & " Where S.taxIdx=" & taxIdx & VbCrlf
	''response.write strSql & "<br><br>"
    dbget.Execute(strSql)

    strSql =    " Update db_order.[dbo].tbl_taxSheet Set " & VbCrlf
    strSql = strSql & "	neoTaxNo = '" & no_tax & "' " & VbCrlf
    strSql = strSql & "	,curUserId = '" & Session("ssBctId") & "' " & VbCrlf
    strSql = strSql & "	,printDate = getdate() " & VbCrlf
    strSql = strSql & "	,isueYn = 'Y' " & VbCrlf
    strSql = strSql & "	,isueDate = '" & write_date & "' " & VbCrlf
    strSql = strSql & "	,no_iss= '" & no_iss & "' " & VbCrlf
    strSql = strSql & " Where taxIdx=" & taxIdx & VbCrlf
	''response.write strSql & "<br><br>"
    dbget.Execute(strSql)

	if (Not IsNull(oTax.FOneItem.Forderserial)) and (oTax.FOneItem.Forderserial <> "") then '고객 세금계산서 신청
		if (Left(oTax.FOneItem.Forderserial, 2) <> "SO") then

			if oTax.FOneItem.Fbilldiv = "01" then
				sqlStr = " update [db_order].[dbo].tbl_order_master" & VbCrlf
				sqlStr = sqlStr & " set " & VbCrlf
				sqlStr = sqlStr & " authcode = (case when accountdiv in ('7', '20') then '" & taxIdx & "' else authcode end) " + VbCrlf
				sqlStr = sqlStr & " , cashreceiptreq = (case when accountdiv in ('7', '20') then 'T' else 'U' end) " + VbCrlf
				sqlStr = sqlStr & " where orderserial='" & oTax.FOneItem.Forderserial & "'"
				dbget.Execute sqlStr
			end if

		end if
	end if

    ''response.write "<script>parent.closeMe();</script>"


    if (oTax.FOneItem.Fbilldiv = "02") or (oTax.FOneItem.Fbilldiv = "51") or (oTax.FOneItem.Fbilldiv = "99") then	' 오프 가맹점, 기타매출

		'수정된 데이타를 다시 읽어온다.
		oTax.GetTaxRead

		''response.write "------------" & oTax.FOneItem.FisueDate

		sqlStr = " update db_shop.dbo.tbl_fran_meachuljungsan_master" + vbCrlf
		sqlStr = sqlStr + " set taxlinkidx=" + CStr(oTax.FOneItem.FtaxIdx) + vbCrlf
		sqlStr = sqlStr + " ,neotaxno='" + CStr(oTax.FOneItem.FneoTaxNo) + "'" + vbCrlf
		sqlStr = sqlStr + " ,issuestatecd='9'"  + vbCrlf		' 계산서발행완료
		sqlStr = sqlStr + " ,taxdate = '" + CStr(oTax.FOneItem.FisueDate) + "'"  + vbCrlf
		sqlStr = sqlStr + " ,taxregdate = getdate() " + vbCrlf
		sqlStr = sqlStr + " ,eseroTaxKey='" & no_iss & "' " & VbCrlf  '''국세청번호. /서동석 추가.
		sqlStr = sqlStr + " where idx=" + CStr(oTax.FOneItem.Forderidx) '고객계산서의 경우 주문번호 / 가맹점의 경우 인덱스코드
		''response.write sqlStr & "<br><br>"
		dbget.Execute(sqlStr)

		if CStr(oTax.FOneItem.Forderidx) = "0" then
			sqlStr = " update db_shop.dbo.tbl_fran_meachuljungsan_master" + vbCrlf
			sqlStr = sqlStr + " set taxlinkidx=" + CStr(oTax.FOneItem.FtaxIdx) + vbCrlf
			sqlStr = sqlStr + " ,neotaxno='" + CStr(oTax.FOneItem.FneoTaxNo) + "'" + vbCrlf
			sqlStr = sqlStr + " ,issuestatecd='9'"  + vbCrlf		' 계산서발행완료
			sqlStr = sqlStr + " ,taxdate = '" + CStr(oTax.FOneItem.FisueDate) + "'"  + vbCrlf
			sqlStr = sqlStr + " ,taxregdate = getdate() " + vbCrlf
			sqlStr = sqlStr + " ,eseroTaxKey='" & no_iss & "' " & VbCrlf  '''국세청번호. /서동석 추가.
			sqlStr = sqlStr + " where idx in (select matchlinkkey from db_order.dbo.tbl_taxSheet_Match where matchtype = 'E' and taxidx = " & oTax.FOneItem.Ftaxidx & ") "

			dbget.Execute(sqlStr)
		end if

		sqlStr = " UPDATE A " + vbCrlf
		sqlStr = sqlStr + " SET a.segumDate = '" + CStr(oTax.FOneItem.FisueDate) + "'"  + vbCrlf
		sqlStr = sqlStr + " FROM db_storage.dbo.tbl_ordersheet_master a " & vbCrLf
		sqlStr = sqlStr + " INNER JOIN [db_shop].[dbo].tbl_fran_meachuljungsan_submaster b " & vbCrLf
		sqlStr = sqlStr + " ON a.baljucode = b.code02 " & vbCrLf
		sqlStr = sqlStr + " WHERE b.masterIdx = " + CStr(oTax.FOneItem.Forderidx) '고객계산서의 경우 주문번호 / 가맹점의 경우 인덱스코드

		dbget.Execute(sqlStr)

	end if



end IF
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
