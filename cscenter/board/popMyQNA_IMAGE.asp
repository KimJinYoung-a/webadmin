<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/popheaderUTF8.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/myqnacls.asp" -->
<%

'==============================================================================
'나의 1:1질문답변
dim boardqna, imgidx
set boardqna = New CMyQNA

imgidx = request("imgidx")

boardqna.read(request("idx"))

%>
<% if (imgidx = 0) then %>
<img src="<%= uploadUrl %><%= boardqna.results(0).FattachFile %>">
<% elseif (imgidx = 1) then %>
<img src="<%= uploadUrl %><%= boardqna.results(0).FattachFile2 %>">
<% end if %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
