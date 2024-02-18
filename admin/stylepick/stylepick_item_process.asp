<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 스타일픽 관리
' Hieditor : 2011.04.05 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->

<%
dim itemidarr , mode , sqlstr , i , catetype , tmpitem , tmpitemid , cd1 ,cd2, cd3
dim totalcount , itemidxarr
	itemidarr = request("itemidarr")
	itemidxarr = request("itemidxarr")
	mode = request("mode")
	menupos = request("menupos")
	catetype = request("catetype")
	cd1 = request("cd1")
	cd2 = request("cd2")
	cd3 = request("cd3")

dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	
'/상품 신규등록
if mode = "itemadd" then

	if itemidarr = "" or cd1 = "" or cd2 = "" then
		response.write "<script language='javascript'>"
		response.write	"	alert('코드에 문제가 있습니다.관리자 문의 하세요');"		
		response.write "</script>"	
		dbget.close()	:	response.End	
	end if
	
	'/다른카테고리에 있는 중복 상품 체크	
	sqlstr = "select"
	sqlstr = sqlstr & " si.itemid"
	sqlstr = sqlstr & ",(select top 1 catename"
	sqlstr = sqlstr & "		from db_giftplus.dbo.tbl_stylepick_cate_cd1"
	sqlstr = sqlstr & "		where isusing='Y' and si.cd1 = cd1) as cd1name"
	sqlstr = sqlstr & ",(select top 1 catename"
	sqlstr = sqlstr & "		from db_giftplus.dbo.tbl_stylepick_cate_cd2"
	sqlstr = sqlstr & "		where isusing='Y' and si.cd2 = cd2) as cd2name"
	sqlstr = sqlstr & " FROM [db_giftplus].dbo.tbl_stylepick_item si"
	sqlstr = sqlstr & " where si.isusing='Y'"
	sqlstr = sqlstr & " and si.itemid in ("&itemidarr&")"
	sqlstr = sqlstr & " and si.cd1 = "&cd1&" and si.cd2 <> "&cd2&""
				
	'response.write sqlstr &"<Br>"
	rsget.open sqlstr ,dbget,1
	
	totalcount = rsget.recordcount
	
	if not rsget.EOF then
		do until rsget.EOF
		
		i = i + 1
		
		if tmpitem = "" then tmpitem = "\n\n타카테고리 중복 등록상품 입니다. 참고하세요\n※10건 까지 노출됩니다\n\n"
		
		'/10건 까지만 조회
		if i+1 <= 10 then
			tmpitem = tmpitem & "["& rsget("cd1name") &" / " & rsget("cd2name") & "] 상품코드:" & rsget("itemid") & "\n"
		end if
		
		tmpitemid = tmpitemid & rsget("itemid")
		
		if totalcount <> i then tmpitemid = tmpitemid &","
					
		rsget.movenext
		loop
	end if
	
	rsget.Close

	sqlstr = "insert into [db_giftplus].dbo.tbl_stylepick_item (itemid,isusing, cd1,cd2,cd3)" + vbcrlf
	sqlstr = sqlstr & "	select" + vbcrlf	
	sqlstr = sqlstr & "	i.itemid , 'Y' ,'"&cd1&"','"&cd2&"',''" + vbcrlf	
	sqlstr = sqlstr & "	from db_item.dbo.tbl_item i" + vbcrlf
	sqlstr = sqlstr & "	left join [db_giftplus].dbo.tbl_stylepick_item si" + vbcrlf
	sqlstr = sqlstr & "	on i.itemid = si.itemid" + vbcrlf
	sqlstr = sqlstr & "	and si.isusing='Y'" + vbcrlf
	sqlstr = sqlstr & "	and si.cd1 = '"&cd1&"' and si.cd2 = '"&cd2&"' and si.cd3=''" + vbcrlf
	sqlstr = sqlstr & "	where i.isusing = 'Y'" + vbcrlf
	sqlstr = sqlstr & "	and i.itemid in ("&itemidarr&")" + vbcrlf
	sqlstr = sqlstr & "	and si.itemid is null" + vbcrlf		'/같은 카테고리내 중복 상품 제낌
	
	if tmpitemid <> "" then
		'sqlstr = sqlstr & "	and si.itemid not in ("&tmpitemid&")" + vbcrlf		'/다른 카테고리내 중복 상품 제낌
	end if

	'response.write sqlstr &"<Br>"
	dbget.execute sqlstr

	response.write	"<script language='javascript'>"
	response.write	"	alert('저장되었습니다"&tmpitem&"');"
	response.write "	location.replace('about:blank');"
	response.write "	parent.opener.location.reload();"
	response.write "	self.focus();"
	response.write	"</script>"
	dbget.close()	:	response.End

'/상품 삭제
elseif mode = "itemdel" then

	if itemidxarr = "" then
		response.write "<script language='javascript'>"
		response.write	"	alert('코드에 문제가 있습니다.관리자 문의 하세요');"		
		response.write "</script>"	
		dbget.close()	:	response.End	
	end if
	
	sqlstr = "delete [db_giftplus].dbo.tbl_stylepick_item " + vbcrlf
	sqlstr = sqlstr & " where itemidx in ("&itemidxarr&")"
	
	'response.write sqlstr &"<Br>"
    dbget.Execute sqlStr		
	
	response.write	"<script language='javascript'>"
	response.write	"	alert('삭제되었습니다');"
	response.write "	location.href='/admin/stylepick/stylepick_item.asp?menupos="&menupos&"';"
	response.write	"</script>"
	dbget.close()	:	response.End	
end if
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->