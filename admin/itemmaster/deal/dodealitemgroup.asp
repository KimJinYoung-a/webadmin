<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : /admin/itemmaster/deal/dodealitemgroup.asp
' Description :  �� ��ǰ  �׷� - ���, ����
' History : 2022.10.17 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<%
'--------------------------------------------------------
' �������� & �Ķ���� �� �ޱ�
'--------------------------------------------------------
Dim k, sqlStr, i, tmpSort, cnt
Dim idx : idx = requestCheckVar(Request.Form("idx"),9)
Dim title : title = requestCheckVar(Request.Form("title"),128)
Dim sort : sort = requestCheckVar(Request.Form("sort"),2)
Dim mode : mode = requestCheckVar(Request.Form("mode"),10)
Dim group_code : group_code = requestCheckVar(Request.Form("groupCode"),10)
Dim selGroup : selGroup = requestCheckVar(Request.Form("selGroup"),10)
Dim itemidarr : itemidarr = Request.Form("itemidarr")
Dim sortarr : sortarr = Request.Form("sortarr")

if Request.Form("title") <> "" then
	if checkNotValidHTML(Request.Form("title")) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if

if mode="add" then
	sqlStr = "INSERT INTO [db_event].[dbo].[tbl_deal_event_group] (deal_code, title, sort)" & vbcrlf
	sqlStr = sqlStr + " VALUES (" & idx & ", '" & title &"'," & sort & ")"
	dbget.execute sqlStr
	response.write "<script>"
	response.write "	location.replace('/admin/itemmaster/deal/pop_dealitem_group.asp?idx=" + idx + "');"
	response.write "	parent.opener.fnLoadItems();"
	response.write "</script>"
	response.End
elseif mode="update" then
	sqlStr = "update [db_event].[dbo].[tbl_deal_event_group] set title='" & title & "', sort='" & sort & "' WHERE group_code=" & group_code & vbcrlf
	dbget.execute sqlStr
	response.write "<script>"
	response.write "	location.replace('/admin/itemmaster/deal/pop_dealitem_group.asp?idx=" + idx + "');"
	response.write "</script>"
	response.End
elseif mode="move" then
	sqlStr = "update [db_event].[dbo].[tbl_deal_event_item] set group_code='" & selGroup & "' WHERE dealcode=" & idx & " and itemid in (" & itemidarr & ")" & vbcrlf
	dbget.execute sqlStr
	response.write "<script>"
	response.write "	location.replace('/admin/itemmaster/deal/dealitem_regist.asp?idx=" + idx + "');"
	response.write "</script>"
	response.End
elseif mode="sort" then
	If sortarr="" THEN
		Response.Write "<script language='javascript'>history.back(-1);</script>"
		dbget.close() : response.End
	end if

	'���û�ǰ �ľ�
	itemidarr = split(itemidarr,",")
	cnt = ubound(itemidarr)
	if cnt > 0 then
		sortarr =  split(sortarr,",")
	end if
	'// ���ļ��� ����  
	for i=0 to cnt
		tmpSort = "NULL"
		if cnt > 0 then
			if sortarr(i)<> "" then
				tmpSort = sortarr(i)
			end if
		else
			tmpSort = sortarr
		end if 
		sqlStr = "UPDATE [db_event].[dbo].[tbl_deal_event_group]"
		sqlStr = sqlStr & " SET sort = " & tmpSort
		sqlStr = sqlStr & "	WHERE deal_code=" & idx & " and group_code=" & itemidarr(i)
		dbget.execute sqlStr
	next
	response.write "<script>"
	response.write "	location.replace('/admin/itemmaster/deal/pop_dealitem_group.asp?idx=" + idx + "');"
	response.write "</script>"
	response.End
elseif mode="edit" then
	If sortarr="" THEN
		Response.Write "<script language='javascript'>history.back(-1);</script>"
		dbget.close() : response.End
	end if

	'���û�ǰ �ľ�
	itemidarr = split(itemidarr,",")
	cnt = ubound(itemidarr)
	if cnt > 0 then
		sortarr =  split(sortarr,",")
	end if
	'// ���ļ��� ����  
	for i=0 to cnt
		tmpSort = "NULL"
		if cnt > 0 then
			if sortarr(i)<> "" then
				tmpSort = sortarr(i)
			end if
		else
			tmpSort = sortarr
		end if 
		sqlStr = "UPDATE [db_event].[dbo].[tbl_deal_event_item]"
		sqlStr = sqlStr & " SET viewidx = " & tmpSort
		sqlStr = sqlStr & "	WHERE dealcode=" & idx & " and itemid=" & itemidarr(i)
		dbget.execute sqlStr
	next
	response.write "<script>"
	response.write "	location.replace('/admin/itemmaster/deal/dealitem_regist.asp?idx=" + idx + "');"
	response.write "	parent.opener.fnLoadItems();"
	response.write "</script>"
	response.End
elseif mode="del" then
	sqlStr = "update [db_event].[dbo].[tbl_deal_event_group] set isusing='N' WHERE group_code=" & group_code & vbcrlf
	dbget.execute sqlStr
	response.write "<script>"
	response.write "	location.replace('/admin/itemmaster/deal/pop_dealitem_group.asp?idx=" + idx + "');"
	response.write "	parent.opener.fnLoadItems();"
	response.write "</script>"
	response.End
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->