<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ��Ÿ���� ����
' Hieditor : 2011.04.05 �ѿ�� ����
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
	
'/��ǰ �űԵ��
if mode = "itemadd" then

	if itemidarr = "" or cd1 = "" or cd2 = "" then
		response.write "<script language='javascript'>"
		response.write	"	alert('�ڵ忡 ������ �ֽ��ϴ�.������ ���� �ϼ���');"		
		response.write "</script>"	
		dbget.close()	:	response.End	
	end if
	
	'/�ٸ�ī�װ��� �ִ� �ߺ� ��ǰ üũ	
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
		
		if tmpitem = "" then tmpitem = "\n\nŸī�װ� �ߺ� ��ϻ�ǰ �Դϴ�. �����ϼ���\n��10�� ���� ����˴ϴ�\n\n"
		
		'/10�� ������ ��ȸ
		if i+1 <= 10 then
			tmpitem = tmpitem & "["& rsget("cd1name") &" / " & rsget("cd2name") & "] ��ǰ�ڵ�:" & rsget("itemid") & "\n"
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
	sqlstr = sqlstr & "	and si.itemid is null" + vbcrlf		'/���� ī�װ��� �ߺ� ��ǰ ����
	
	if tmpitemid <> "" then
		'sqlstr = sqlstr & "	and si.itemid not in ("&tmpitemid&")" + vbcrlf		'/�ٸ� ī�װ��� �ߺ� ��ǰ ����
	end if

	'response.write sqlstr &"<Br>"
	dbget.execute sqlstr

	response.write	"<script language='javascript'>"
	response.write	"	alert('����Ǿ����ϴ�"&tmpitem&"');"
	response.write "	location.replace('about:blank');"
	response.write "	parent.opener.location.reload();"
	response.write "	self.focus();"
	response.write	"</script>"
	dbget.close()	:	response.End

'/��ǰ ����
elseif mode = "itemdel" then

	if itemidxarr = "" then
		response.write "<script language='javascript'>"
		response.write	"	alert('�ڵ忡 ������ �ֽ��ϴ�.������ ���� �ϼ���');"		
		response.write "</script>"	
		dbget.close()	:	response.End	
	end if
	
	sqlstr = "delete [db_giftplus].dbo.tbl_stylepick_item " + vbcrlf
	sqlstr = sqlstr & " where itemidx in ("&itemidxarr&")"
	
	'response.write sqlstr &"<Br>"
    dbget.Execute sqlStr		
	
	response.write	"<script language='javascript'>"
	response.write	"	alert('�����Ǿ����ϴ�');"
	response.write "	location.href='/admin/stylepick/stylepick_item.asp?menupos="&menupos&"';"
	response.write	"</script>"
	dbget.close()	:	response.End	
end if
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->