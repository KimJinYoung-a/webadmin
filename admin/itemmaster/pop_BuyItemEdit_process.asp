<%@  codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<% Session.codepage="65001" %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description :  상품고시
' History : 2013.12.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr, mode
dim itemgubun, itemid, itemoption, makerid, upchemanagecode, buyitemname, buyitemoptionname, currencyUnit, buyitemprice
dim itemgubunArr, itemidArr, itemoptionArr, makeridArr, upchemanagecodeArr, buyitemnameArr, buyitemoptionnameArr, currencyUnitArr, buyitempriceArr

dim i, j, k

mode 	= request("mode")

itemgubunArr 	= Split(request("itemgubun"), ",")
itemidArr 		= Split(request("itemid"), ",")
itemoptionArr 	= Split(request("itemoption"), ",")
makeridArr 		= Split(request("makerid"), ",")
upchemanagecodeArr 		= Split(request("upchemanagecode"), ",")
buyitemnameArr 			= Split(request("buyitemname"), ",")
buyitemoptionnameArr 	= Split(request("buyitemoptionname"), ",")
currencyUnitArr 		= Split(request("currencyUnit"), ",")
buyitempriceArr 		= Split(request("buyitemprice"), ",")

select case mode
	case "ins"
		for i = 0 to ubound(itemgubunArr)
			itemgubun = trim(itemgubunArr(i))
			itemid = trim(itemidArr(i))
			itemoption = trim(itemoptionArr(i))
			makerid = trim(html2db(makeridArr(i)))
			upchemanagecode = trim(html2db(upchemanagecodeArr(i)))
			buyitemname = trim(html2db(buyitemnameArr(i)))
			buyitemoptionname = trim(html2db(buyitemoptionnameArr(i)))
			currencyUnit = trim(html2db(currencyUnitArr(i)))
			buyitemprice = trim(html2db(buyitempriceArr(i)))

			if (itemgubun <> "") then
				sqlStr = " if not exists( "
				sqlStr = sqlStr & "	select top 1 itemid from [db_item].[dbo].[tbl_item_option_stock] "
				sqlStr = sqlStr & "		where "
				sqlStr = sqlStr & "			1 = 1 "
				sqlStr = sqlStr & "			and itemgubun = '" & itemgubun & "' "
				sqlStr = sqlStr & "			and itemid = '" & itemid & "' "
				sqlStr = sqlStr & "			and itemoption = '" & itemoption & "' "
				sqlStr = sqlStr & "	) "
				sqlStr = sqlStr & "	begin "
				sqlStr = sqlStr & "		insert into [db_item].[dbo].[tbl_item_option_stock](itemgubun, itemid, itemoption, upchemanagecode) "
				sqlStr = sqlStr & "		values('" & itemgubun & "', '" & itemid & "', '" & itemoption & "', N'" & upchemanagecode & "') "
				sqlStr = sqlStr & "	end "
				sqlStr = sqlStr & "	else "
				sqlStr = sqlStr & "	begin "
				sqlStr = sqlStr & "		update [db_item].[dbo].[tbl_item_option_stock] "
				sqlStr = sqlStr & "		set upchemanagecode = '" & upchemanagecode & "' "
				sqlStr = sqlStr & "		where "
				sqlStr = sqlStr & "			1 = 1 "
				sqlStr = sqlStr & "			and itemgubun = '" & itemgubun & "' "
				sqlStr = sqlStr & "			and itemid = '" & itemid & "' "
				sqlStr = sqlStr & "			and itemoption = '" & itemoption & "' "
				sqlStr = sqlStr & "	end "
				''response.write sqlStr & "<BR>"
				dbget.execute sqlStr

				sqlStr =  "	if not exists( "
				sqlStr = sqlStr & "		select top 1 buyitemid from [db_shop].[dbo].[tbl_buy_item] "
				sqlStr = sqlStr & "		where "
				sqlStr = sqlStr & "			1 = 1 "
				sqlStr = sqlStr & "			and itemgubun = '" & itemgubun & "' "
				sqlStr = sqlStr & "			and buyitemid = '" & itemid & "' "
				sqlStr = sqlStr & "			and itemoption = '" & itemoption & "' "
				sqlStr = sqlStr & "	) "
				sqlStr = sqlStr & "	begin "
				sqlStr = sqlStr & "		insert into [db_shop].[dbo].[tbl_buy_item](itemgubun, buyitemid, itemoption, makerid, buyitemname, buyitemoptionname, buyitemprice, currencyUnit, isusing, regdate, updt) "
				sqlStr = sqlStr & "		values('" & itemgubun & "', '" & itemid & "', '" & itemoption & "', N'" & makerid & "', N'" & buyitemname & "', N'" & buyitemoptionname & "', '" & buyitemprice & "', '" & currencyUnit & "', 'Y', getdate(), getdate()) "
				sqlStr = sqlStr & "	end "
				sqlStr = sqlStr & "	else "
				sqlStr = sqlStr & "	begin "
				sqlStr = sqlStr & "		update [db_shop].[dbo].[tbl_buy_item] "
				sqlStr = sqlStr & "		set "
				sqlStr = sqlStr & "			makerid = N'" & makerid & "', "
				sqlStr = sqlStr & "			buyitemname = N'" & buyitemname & "', "
				sqlStr = sqlStr & "			buyitemoptionname = N'" & buyitemoptionname & "', "
				sqlStr = sqlStr & "			buyitemprice = '" & buyitemprice & "', "
				sqlStr = sqlStr & "			currencyUnit = '" & currencyUnit & "', "
				sqlStr = sqlStr & "			updt = getdate() "
				sqlStr = sqlStr & "		where "
				sqlStr = sqlStr & "			1 = 1 "
				sqlStr = sqlStr & "			and itemgubun = '" & itemgubun & "' "
				sqlStr = sqlStr & "			and buyitemid = '" & itemid & "' "
				sqlStr = sqlStr & "			and itemoption = '" & itemoption & "' "
				sqlStr = sqlStr & "	end "
				''response.write sqlStr & "<BR>"
				dbget.execute sqlStr
			end if
		next

		response.write "<script language='javascript'>"
		response.write "	alert('저장되었습니다');"
		response.write "	document.location.href='"& refer &"'"
		response.write "</script>"
	case else
		response.write "잘못된 접근입니다."
end select

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<% Session.codepage="949" %>
