<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 이벤트 상품추가
' History : 2010.03.11 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode , itemid ,evt_code ,addSql ,sqlStr , itemoption , itemgubun
dim itemidarr, itemoptionarr ,itemgubunarr , i ,result
	mode = requestCheckVar(Request("mode"),32)
	itemid = requestCheckVar(Request("itemid"),10)
	itemoption = requestCheckVar(Request("itemoption"),4)
	itemgubun = requestCheckVar(Request("itemgubun"),2)
	itemidarr = Request("itemidarr")
	itemoptionarr = Request("itemoptionarr")
	itemgubunarr = Request("itemgubunarr")
	evt_code = requestCheckVar(request("evt_code"),10)
	'response.write mode &"<br>"
	'response.end

dim referer
referer = request.ServerVariables("HTTP_REFERER")

'// 상품추가
if mode = "itemadd" then

	itemidarr = split(itemidarr,",")
	itemoptionarr = split(itemoptionarr,",")
	itemgubunarr = split(itemgubunarr,",")

	dbget.begintrans

	for i = 0 to ubound(itemidarr)-1

		result = ""
		if IsNumeric(trim(itemidarr(i))) = TRUE and ((len(trim(itemgubunarr(i))&trim(Format00(6,itemidarr(i)))&trim(itemoptionarr(i))) = 12) or (len(trim(itemgubunarr(i))&trim(Format00(8,itemidarr(i)))&trim(itemoptionarr(i))) = 14)) then

			sqlStr = "SELECT top 1 shopitemid"
			sqlStr = sqlStr & " FROM [db_shop].[dbo].[tbl_shop_item]"
			sqlStr = sqlStr & " WHERE itemgubun = '" & requestCheckVar(trim(itemgubunarr(i)),2) & "'"
			sqlStr = sqlStr & " AND shopitemid = '" & requestCheckVar(trim(itemidarr(i)),10) & "'"
			sqlStr = sqlStr & " AND itemoption = '" & requestCheckVar(trim(itemoptionarr(i)),4) & "'"

			'response.write sqlStr & "<Br>"
			rsget.Open sqlStr,dbget,1

			If Not rsget.Eof Then
				result = "LOGICSBARCODE"
			End If

			rsget.close()
		end if

		if result = "" then
			sqlStr = "SELECT top 1 shopitemid"
			sqlStr = sqlStr & " FROM [db_shop].[dbo].[tbl_shop_item]"
			sqlStr = sqlStr & " WHERE extbarcode = '"& requestCheckVar(trim(itemgubunarr(i)),2) & requestCheckVar(trim(itemidarr(i)),10) & requestCheckVar(trim(itemoptionarr(i)),4) &"'"

			'response.write sqlStr & "<Br>"
			rsget.Open sqlStr,dbget,1

			If Not rsget.Eof Then
				result = "EXTBARCODE"
			End If

			rsget.close()
		End IF

		sqlStr = " insert into [db_shop].[dbo].tbl_eventitem_off" + VbCrlf
		sqlStr = sqlStr + " (evt_code, itemid, itemgubun, itemoption)" + VbCrlf
		sqlStr = sqlStr + " 	select " + CStr(evt_code) + ", i.shopitemid, i.itemgubun ,i.itemoption" + VbCrlf
		sqlStr = sqlStr + " 	from [db_shop].dbo.tbl_shop_item i" + VbCrlf
		sqlStr = sqlStr + " 	where 1=1 " + VbCrlf

		if result = "LOGICSBARCODE" then
			sqlStr = sqlStr + " 	and i.itemgubun = '"&itemgubunarr(i)&"'" + VbCrlf
			sqlStr = sqlStr + " 	and i.shopitemid = "&itemidarr(i)&"" + VbCrlf
			sqlStr = sqlStr + " 	and i.itemoption = '"&itemoptionarr(i)&"'" + VbCrlf
		else
			sqlStr = sqlStr + " 	and i.extbarcode = '"&trim(itemgubunarr(i))&trim(itemidarr(i))&trim(itemoptionarr(i))&"'" + VbCrlf
		end if

		sqlStr = sqlStr + " 	and i.shopitemid not in (" + VbCrlf
		sqlStr = sqlStr + "			select ei.itemid" + VbCrlf
		sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_eventitem_off ei" + VbCrlf
		sqlStr = sqlStr + " 		join [db_shop].dbo.tbl_shop_item ii" + VbCrlf
		sqlStr = sqlStr + " 			on ei.itemid = ii.shopitemid" + VbCrlf
		sqlStr = sqlStr + " 			and ei.itemoption = ii.itemoption" + VbCrlf
		sqlStr = sqlStr + " 			and ei.itemgubun = ii.itemgubun" + VbCrlf
		sqlStr = sqlStr + " 		where 1=1 " + VbCrlf

		if result = "LOGICSBARCODE" then
			sqlStr = sqlStr + " 		and ei.itemid = "&itemidarr(i)&"" + VbCrlf
			sqlStr = sqlStr + " 		and ei.itemoption = '"&itemoptionarr(i)&"'" + VbCrlf
			sqlStr = sqlStr + " 		and ei.itemgubun = '"&itemgubunarr(i)&"'" + VbCrlf
		else
			sqlStr = sqlStr + " 		and ii.extbarcode = '"&trim(itemgubunarr(i))&trim(itemidarr(i))&trim(itemoptionarr(i))&"'" + VbCrlf
		end if

		sqlStr = sqlStr + " 		and ei.evt_code = "&evt_code&"" + VbCrlf
		sqlStr = sqlStr + " 	)"

		'response.write sqlStr &"<br>"
		dbget.execute sqlStr
    next

	IF Err.Number = 0 THEN
		dbget.CommitTrans
		response.write "<script type='text/javascript'>alert('OK'); parent.location.reload(); parent.opener.location.reload();</script>"
		dbget.close()	:	response.End
	Else
   		dbget.RollBackTrans
   		response.write "<script type='text/javascript'>alert('데이터 처리에 문제가 발생하였습니다.'); history.back(-1);</script>"
   		dbget.close()	:	response.End
   	end if

'// 선택상품 삭제
elseif mode = "itemdel" then

	itemidarr = split(itemidarr,",")
	itemoptionarr = split(itemoptionarr,",")
	itemgubunarr = split(itemgubunarr,",")

	dbget.begintrans

	for i = 0 to ubound(itemidarr)-1

		sqlStr = "Delete From [db_shop].[dbo].tbl_eventitem_off" + VbCrlf
		sqlStr = sqlStr + " WHERE evt_code = "&evt_code&"" + VbCrlf
		sqlStr = sqlStr + " and itemid = "& requestCheckVar(itemidarr(i),10) &"" + VbCrlf
		sqlStr = sqlStr + " and itemoption = '"& requestCheckVar(itemoptionarr(i),4) &"'" + VbCrlf
		sqlStr = sqlStr + " and itemgubun = '"& requestCheckVar(itemgubunarr(i),2) &"'" + VbCrlf

		'response.write sqlStr &"<br>"
		dbget.execute sqlStr
    next

	IF Err.Number = 0 THEN
		dbget.CommitTrans
		response.write "<script type='text/javascript'>alert('OK'); location.replace('" + referer + "');</script>"
		dbget.close()	:	response.End
	Else
   		dbget.RollBackTrans
   		response.write "<script type='text/javascript'>alert('데이터 처리에 문제가 발생하였습니다.'); history.back(-1);</script>"
   		dbget.close()	:	response.End
   	end if

end if
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
