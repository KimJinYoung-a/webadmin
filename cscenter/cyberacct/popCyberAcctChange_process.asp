<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/DcCyberAcctUtil.asp"-->
<%


dim orderserial, subtotalprice, CLOSEDATE, buf1
orderserial = request.form("orderserial")
subtotalprice = request.form("subtotalprice")
CLOSEDATE = request.form("CLOSEDATE")
buf1 = request.form("buf1")

'response.write orderserial & "<br>"
'response.write subtotalprice & "<br>"
'response.write Replace(CLOSEDATE,"-","") + Replace(buf1,":","") & "<br>"

dim ret : ret = ChangeCyberAcct(orderserial, subtotalprice, Replace(CLOSEDATE,"-","") + Replace(buf1,":",""))

dim ref : ref = request.SERVERVariables("HTTP_REFERER")
if (ret) then
    response.write "<script>alert('수정 되었습니다.');</script>"
    response.write "<script>location.replace('" + ref + "');</script>"
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->