<%@ language=vbscript %>
<% option explicit %>

<%
'###########################################################
' Description : 오프라인 배송
' Hieditor : 2011.02.28 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
If session("ssBctId") = "" then
    response.write "<script language='javascript'>alert('세션이 종료되었습니다.');</script>"
    dbget.close()	:	response.End
end if

dim detailidxArr , ojumun ,ix,sqlStr ,listitem
dim iSall, SheetType
	detailidxArr = request.Form("detailidxArr")
	listitem =  request("orderserial")
	iSall   =  request("isall")


''기본택배사
dim defaultSongjangdiv
sqlStr = "select defaultSongjangdiv from [db_partner].[dbo].tbl_partner"
sqlStr = sqlStr + " where id='" & session("ssBctID") & "'"

'response.write sqlStr &"<Br>"
rsget.Open sqlStr,dbget,1
if Not Rsget.Eof then
    defaultSongjangdiv = rsget("defaultSongjangdiv")

    if IsNULL(defaultSongjangdiv) then defaultSongjangdiv=""
end if
rsget.close

set ojumun = new cupchebeasong_list
	ojumun.frectdetailidxarr = detailidxArr
	ojumun.FRectDesignerID = session("ssBctID")
	ojumun.FRectIsAll       = iSall
	ojumun.fReDesignerSelectBaljuList()

function ReplaceSCVStr(oStr)
    ReplaceSCVStr = ""
    if IsNULL(oStr) then Exit function
    ReplaceSCVStr = Replace(oStr, chr(34),"'")
    ReplaceSCVStr = Replace(oStr, ",","")
end function

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENDLV_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<title>송장파일</title>
<style>

	br
	    {mso-data-placement:same-cell;}
	tr
	    {mso-height-source:auto;
	    mso-ruby-visibility:none;}
	td
	    {white-space:normal;}

</style>
</head>

<body leftmargin="10">
<table border=1>
<tr>
    <td align="center" x:str >일련번호</td>
	<td align="center" height="25" x:str >주문번호</td>
	<td align="center" x:str >수령인</td>
	<td align="center" x:str >상품명</td>
	<td align="center" x:str >옵션명</td>
	<td align="center" x:str >택배사코드</td>
	<td align="center" x:str >송장번호</td>

</tr>
<% for ix=0 to ojumun.FResultCount - 1 %>
<tr>
    <td align="center" x:str><%= ojumun.FItemList(ix).FDetailidx %></td>
	<td align="center" x:str><%= ojumun.FItemList(ix).forderno %></td>
	<td align="center" x:str><%= ReplaceSCVStr(ojumun.FItemList(ix).FReqName) %></td>
	<td align="center" x:str><%= ReplaceSCVStr(ojumun.FItemList(ix).FItemName) %></td>
	<td align="center" x:str><%= ReplaceSCVStr(ojumun.FItemList(ix).FItemoptionName) %></td>
	<td align="center" x:str>
	<% if IsNULL(ojumun.FItemList(ix).FSongjangdiv) then %>
	<%= defaultSongjangdiv %>
	<% end if %>
	</td>
	<td align="center" x:str><%= ojumun.FItemList(ix).Fsongjangno %></td>
</tr>
<% next %>
</table>
</body>
</html>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->