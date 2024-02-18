<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 정기구독 상품
' History : 2016.06.16 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim strSql, i, lastuserid, menupos, mode, identikey, startreserveidx, endreserveidx, vreservecount
dim itemid, itemoption, sendkey, reserveDlvDate, reserveidx, reserveitemgubun, reserveItemID, reserveItemOption, reserveItemName
	lastuserid=session("ssBctId")
	menupos = getNumeric(requestcheckvar(request("menupos"),10))
	mode = requestcheckvar(request("mode"),32)
	itemid = getNumeric(requestcheckvar(request("itemid"),10))
	itemoption = requestcheckvar(request("itemoption"),32)
	startreserveidx = requestcheckvar(request("startreserveidx"),10)
	endreserveidx = requestcheckvar(request("endreserveidx"),10)
	reserveidx = getNumeric(requestcheckvar(request("reserveidx"),10))

vreservecount = 0

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

if InStr(referer,"10x10.co.kr")<1 and session("ssBctId")<>"tozzinet" then
	response.write "not valid Referer"
    response.end
end if

if mode="standinglistedit" then
	for i=1 to request.form("identikey").count
		identikey = request.form("identikey")(i)
		itemid = request.form("itemid_"&identikey)
		itemoption = request.form("itemoption_"&identikey)
		reserveDlvDate = request.form("reserveDlvDate_"&identikey)
		reserveidx = request.form("reserveidx_"&identikey)
		reserveitemgubun = request.form("reserveitemgubun_"&identikey)
		reserveItemOption = request.form("reserveItemOption_"&identikey)
		reserveItemID = request.form("reserveItemID_"&identikey)
		reserveItemName = request.form("reserveItemName_"&identikey)

		if reserveItemName="" or isnull(reserveItemName) then
			reserveItemName="NULL"
		else
			reserveItemName="'" & html2db(trim(reserveItemName)) & "'"
		end if
		if reserveDlvDate="" or isnull(reserveDlvDate) then
			reserveDlvDate="NULL"
		else
			reserveDlvDate="'" & trim(reserveDlvDate) & "'"
		end if
		if reserveItemGubun="" or isnull(reserveItemGubun) then
			reserveItemGubun="NULL"
		else
			reserveItemGubun="'" & trim(reserveItemGubun) & "'"
		end if
		if reserveItemOption="" or isnull(reserveItemOption) then
			reserveItemOption="NULL"
		else
			reserveItemOption="'" & trim(reserveItemOption) & "'"
		end if
		if reserveidx="" or isnull(reserveidx) then
			reserveidx="NULL"
		end if
		if reserveItemID="" or isnull(reserveItemID) then
			reserveItemID="NULL"
		end if

		if getNumeric(itemid)="" then
			response.write "상품코드가 없습니다."
			dbget.close()	:	response.end
		end if
		if itemoption="" then
			response.write "옵션코드가 없습니다."
			dbget.close()	:	response.end
		end if
		if getNumeric(reserveidx)="" then
			response.write "회차가 없습니다."
			dbget.close()	:	response.end
		end if

		strSql = "Update db_item.[dbo].[tbl_item_standing_order]" & vbcrlf
		strSql = strSql & " Set reserveDlvDate=" & reserveDlvDate & "" & vbcrlf
		strSql = strSql & " ,reserveItemGubun=" & reserveItemGubun & "" & vbcrlf
		strSql = strSql & " ,reserveItemID=" & trim(reserveItemID) & "" & vbcrlf
		strSql = strSql & " ,reserveItemOption=" & reserveItemOption & "" & vbcrlf
		strSql = strSql & " ,reserveItemName="& reserveItemName &"" & vbcrlf
		strSql = strSql & " ,lastupdate=getdate()" & vbcrlf
		strSql = strSql & " ,lastadminid='"& lastuserid &"' Where" & vbcrlf
		strSql = strSql & " orgitemid="& trim(itemid) &" and orgitemoption='"& trim(itemoption) &"' and reserveidx="& trim(reserveidx) &"" & vbcrlf

		'response.write strSql & "<br>"
		dbget.Execute strSql
	next

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	location.replace('"& getSCMSSLURL &"/admin/itemmaster/standing/pop_standingIteminfo.asp?itemid="& itemid &"&itemoption="& itemoption &"&menupos="& menupos &"');"
	response.write "</script>"
	dbget.close()	:	response.end

elseif mode="standingdel" then
	if getNumeric(itemid)="" then
		response.write "상품코드가 없습니다."
		dbget.close()	:	response.end
	end if
	if itemoption="" then
		response.write "옵션코드가 없습니다."
		dbget.close()	:	response.end
	end if
	if getNumeric(reserveidx)="" then
		response.write "회차가 없습니다."
		dbget.close()	:	response.end
	end if

	strSql = "delete from db_item.[dbo].[tbl_item_standing_order] where" & vbcrlf
	strSql = strSql & " orgitemid="& trim(itemid) &" and orgitemoption='"& trim(itemoption) &"' and reserveidx="& trim(reserveidx) &"" & vbcrlf

	'response.write strSql & "<br>"
	dbget.Execute strSql

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	location.replace('"& getSCMSSLURL &"/admin/itemmaster/standing/pop_standingIteminfo.asp?itemid="& itemid &"&itemoption="& itemoption &"&menupos="& menupos &"');"
	response.write "</script>"
	dbget.close()	:	response.end

elseif mode="standingitemedit" then
	if getNumeric(itemid)="" then
		response.write "상품코드가 없습니다."
		dbget.close()	:	response.end
	end if
	if itemoption="" then
		response.write "옵션코드가 없습니다."
		dbget.close()	:	response.end
	end if
	if getNumeric(startreserveidx)="" then
		response.write "정기구독 진행 시작 회차 VOL(번호)가 없습니다."
		dbget.close()	:	response.end
	end if
	if getNumeric(endreserveidx)="" then
		response.write "정기구독 진행 종료 회차 VOL(번호)가 없습니다."
		dbget.close()	:	response.end
	end if

	vreservecount = (endreserveidx-startreserveidx)+1
	if vreservecount < 2 then
		response.write "정기구독 진행 회차 VOL(번호)가 잘못 지정되었습니다. 총 진행 횟수(종료회차-시작회차)를 2회 이상으로 지정하세요."
		dbget.close()	:	response.end
	end if

	strSql = "if exists(" & vbcrlf
	strSql = strSql & " 	select top 1 orgitemid" & vbcrlf
	strSql = strSql & " 	from db_item.[dbo].[tbl_item_standing_item]" & vbcrlf
	strSql = strSql & " 	where orgitemid="& trim(itemid) &" and orgitemoption='"& trim(itemoption) &"'" & vbcrlf
	strSql = strSql & " )" & vbcrlf
	strSql = strSql & " 	Update db_item.[dbo].[tbl_item_standing_item]" & vbcrlf
	strSql = strSql & " 	Set startreserveidx=" & trim(startreserveidx) & "" & vbcrlf
	strSql = strSql & " 	,endreserveidx=" & endreserveidx & " Where" & vbcrlf
	strSql = strSql & " 	orgitemid="& trim(itemid) &" and orgitemoption='"& trim(itemoption) &"'" & vbcrlf
	strSql = strSql & " else" & vbcrlf
	strSql = strSql & " 	insert into db_item.[dbo].[tbl_item_standing_item] (" & vbcrlf
	strSql = strSql & " 	orgitemid, orgitemoption, startreserveidx, endreserveidx) values (" & vbcrlf
	strSql = strSql & " 	"& trim(itemid) &", '"& trim(itemoption) &"', "& trim(startreserveidx) &", "& trim(endreserveidx) &"" & vbcrlf
	strSql = strSql & " 	)" & vbcrlf

	'response.write strSql & "<br>"
	dbget.Execute strSql

	' 정기구독 진행 회차를 루프 돌면서 꽂아 넣음.
	for i = startreserveidx to endreserveidx
		sendkey = (i - startreserveidx)+1	' 차수

		strSql = "if not exists(" & vbcrlf
		strSql = strSql & " 	select top 1 orgitemid" & vbcrlf
		strSql = strSql & " 	from db_item.[dbo].[tbl_item_standing_order]" & vbcrlf
		strSql = strSql & " 	where orgitemid="& trim(itemid) &" and orgitemoption='"& trim(itemoption) &"' and reserveidx="& trim(i) &"" & vbcrlf
		strSql = strSql & " )" & vbcrlf
		strSql = strSql & " 	insert into db_item.[dbo].[tbl_item_standing_order] (" & vbcrlf
		strSql = strSql & " 	orgitemid, orgitemoption, reserveidx, reserveItemName, reserveDlvDate, reserveItemGubun, reserveItemID, reserveItemOption" & vbcrlf
		strSql = strSql & " 	, regadminid, lastadminid)" & vbcrlf
		strSql = strSql & " 		select top 1" & vbcrlf
		strSql = strSql & " 		"& trim(itemid) &", '"& trim(itemoption) &"', "& trim(i) &", NULL, NULL, NULL, NULL, NULL" & vbcrlf
		strSql = strSql & " 		, '"& lastuserid &"', '"& lastuserid &"'" & vbcrlf

		'response.write strSql & "<br>"
		dbget.Execute strSql
	next

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	location.replace('"& getSCMSSLURL &"/admin/itemmaster/standing/pop_standingIteminfo.asp?itemid="& itemid &"&itemoption="& itemoption &"&menupos="& menupos &"');"
	response.write "</script>"
	dbget.close()	:	response.end

else
	response.write "<script type='text/javascript'>"
	response.write "	alert('구분자가 없습니다.');"
	response.write "</script>"
	dbget.close()	:	response.end
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->