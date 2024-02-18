<%@ language=vbscript %>
<% option explicit %>
<%
'############### 2008년 11월 4일 한용민 생성
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->

<% 
dim mode ,  event_type , evt_code, idx_order, isusing , idx , event_link , itemid, CateCode
	mode = request("mode")
	event_type = request("event_type")
	evt_code = request("evt_code")
	idx_order = request("idx_order")
	isusing = request("isusing")
	idx = request("idx")
	event_link = request("event_link")
	itemid = request("itemid")
	CateCode = request("cate")
		
	
if event_type = "" or event_link = "" then
	response.write "<script>"
	response.write "alert('[오류]파라메타가없습니다.빠진거 없나 다시 확인후 시도하세요');"
	response.write "history.go(-1);"
	response.write "</script>"		
end if	

dim sql

'//신규추가
if mode = "new" then

sql = ""
sql = "insert into db_diary2010.dbo.tbl_event (evt_code, event_type, isusing, idx_order, event_link, itemid, cate) values (" & vbcrlf
sql = sql & " '"&evt_code&"'" & vbcrlf
sql = sql & " ,'"&event_type&"'" & vbcrlf
sql = sql & " ,'"&isusing&"'" & vbcrlf
sql = sql & " ,'"&idx_order&"'" & vbcrlf
sql = sql & " ,'"&event_link&"'" & vbcrlf
sql = sql & " ,'"&itemid&"'" & vbcrlf
sql = sql & " ,'"&CateCode&"'" & vbcrlf
sql = sql & " )" & vbcrlf

response.write sql&"<br>"
dbget.execute sql
%>
<script language="javascript">
	opener.location.reload();
	window.close();
</script>

<%
elseif mode = "edit" then

'//수정모드
if idx = "" then
	response.write "<script>"
	response.write "alert('idx 파라메타 값이없습니다');"
	response.write "history.go(-1);"
	response.write "</script>"		
end if

sql = ""
sql = "update db_diary2010.dbo.tbl_event set" & vbcrlf
sql = sql & " evt_code = '"& evt_code &"'"  & vbcrlf
sql = sql & " ,event_type = '"& event_type &"'"  & vbcrlf
sql = sql & " ,isusing = '"& isusing &"'"  & vbcrlf
sql = sql & " ,idx_order = '"& idx_order &"'" & vbcrlf
sql = sql & " ,event_link = '"& event_link &"'"  & vbcrlf
sql = sql & " ,itemid = '"& itemid &"'" & vbcrlf
sql = sql & " ,cate = '"& CateCode &"'" & vbcrlf
sql = sql & " where idx = "& idx &""  & vbcrlf

response.write sql&"<br>"
dbget.execute sql
%>
<script language="javascript">
	opener.location.reload();
	window.close();
</script>

<%
end if
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->