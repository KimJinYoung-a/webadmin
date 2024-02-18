<%@ language=vbscript %>
<% option explicit %>
<%
Server.ScriptTimeOut = 60*10		' 10분
%>
<%
'#######################################################
' Description : cs센터 cs처리리스트
' History	:  2007.06.01 이상구 생성
'              2017.07.05 한용민 수정
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/checknoticls.asp"-->
<%
dim menupos, select_type, arrlist, i
	menupos = requestCheckvar(request("menupos"),10)
    select_type = requestCheckvar(request("select_type"),32)

dim onoti
set onoti = New CNoti
    onoti.frectselect_type=select_type
    onoti.getchulgobogo
    arrlist = onoti.farrlist

Response.Buffer=true
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
if select_type="samedaymichulgo" then
    Response.AddHeader "Content-Disposition", "attachment; filename=TEN_출고보고(당일미출고주문)_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
elseif select_type="delaychulgo" then
    Response.AddHeader "Content-Disposition", "attachment; filename=TEN_출고보고(지연출고주문)_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
elseif select_type="delaychulgodate" then
    Response.AddHeader "Content-Disposition", "attachment; filename=TEN_출고보고(지연출고_결제일빠른날짜)_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
elseif select_type="delaychulgocnt" then
    Response.AddHeader "Content-Disposition", "attachment; filename=TEN_출고보고(지연출고_결제일빠른주문)_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
else
    Response.AddHeader "Content-Disposition", "attachment; filename=TEN_출고보고_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
end if
Response.CacheControl = "public"
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 12">
<style type="text/css">
 td {font-size:8.0pt;}
 .txt {mso-number-format:"\@";}
 .num {mso-number-format:"0";}
 .prc {mso-number-format:"\#\,\#\#0";}
</style>
</head>
<body>
<!--[if !excel]>　　<![endif]-->
<div align=center x:publishsource="Excel">

<table width="100%" border="1" align="center" class="a csH15" cellpadding="2" cellspacing="1" bgcolor="#BABABA" style="table-layout:fixed">
<tr bgcolor="<%= adminColor("tabletop") %>">
    <% if select_type="delaychulgodate" then %>
        <td align="center">가장빠른결제일</td>
    <% else %>
        <td align="center">주문번호</td>
    <% end if %>
</tr>
<% if isarray(arrlist) then %>
<% for i = 0 to ubound(arrlist,2) %>
    <tr bgcolor="#FFFFFF" align="center" >
        <td><%= arrlist(0,i) %></td>
    </tr>
<%
    if i mod 1000 = 0 then
        Response.Flush		' 버퍼리플래쉬
    end if
next
end if
%>

</table>
</div>
</body>
</html>
<%
set onoti = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
