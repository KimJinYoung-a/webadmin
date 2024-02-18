<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/event/etcsongjangcls.asp"-->

<%
'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=SVC" & request("sjYYYYMMDD") & ".xls"
Response.CacheControl = "public"


dim buf, i
buf = "발주번호"&VbTab&"송장번호"&VbCrlf

dim onesongjang
set onesongjang = new CEventsBeasong
onesongjang.FPageSize = 2000
onesongjang.FCurrPage = 1

onesongjang.getSVCSongjangList(request("sjYYYYMMDD"))

for i=0 to onesongjang.FREsultCount-1
    buf = buf & onesongjang.FItemList(i).FetcBaljuNo&VbTab&onesongjang.FItemList(i).Fsongjangno&VbCrlf
Next
set onesongjang = Nothing

response.write buf
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->