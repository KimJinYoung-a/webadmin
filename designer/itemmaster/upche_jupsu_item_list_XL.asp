<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<%

dim itemid
dim page
dim searchtype

itemid  = RequestCheckVar(request("itemid"),10)
page = RequestCheckVar(request("page"),10)
searchtype = RequestCheckVar(request("searchtype"),32)

if (page="") then page=1


'상품코드 유효성 검사(2008.08.01;허진원)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

'==============================================================================
dim oitem

set oitem = new CItem

oitem.FRectMakerId = session("ssBctID")
oitem.FRectItemid = itemid
oitem.FRectSearchType = searchtype
oitem.GetJupsuProductListQuick

dim i

dim jupsuSUM, ipkumSUM, notifySUM, confirmSUM


'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_ORDSUM" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"

function replaceXlText(org)
    dim reText
    reText = replace(org,"<","&lt;")
    replaceXlText = replace(reText,">","&gt;")
end function
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<title>미배송상품합계</title>
<style>
    <!--
	br
	    {mso-data-placement:same-cell;}
	tr
	    {mso-height-source:auto;
	    mso-ruby-visibility:none;}
	td
	    {white-space:normal;}
	-->
</style>
</head>

<body leftmargin="10">

	<table width=1200 cellspacing=0 cellpadding=1 border=1>
	    <tr align="center" height="25" >
			<td width="90" x:str >상품코드</td>
			<td width="300" x:str>상품명</td>
			<td width="180" x:str>옵션명</td>
			<td width="70" x:str>주문접수</td>
			<td width="70" x:str>결재완료</td>
			<td width="70" x:str>업체통보</td>
			<td width="70" x:str>업체확인</td>
	    </tr>
<% if oitem.FresultCount<1 then %>
	    <tr bgcolor="#FFFFFF">
	    	<td colspan="7" align="center" x:str>[검색결과가 없습니다.]</td>
	    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
	<%
	jupsuSUM = 0
	ipkumSUM = 0
	notifySUM = 0
	confirmSUM = 0
	%>
    <% for i=0 to oitem.FresultCount-1 %>
		<tr class="a" height="25" >
			<td align="center" x:str><%= oitem.FItemList(i).Fitemid %></td>
			<td align="left" x:str>
				<% =oitem.FItemList(i).Fitemname %>
			</td>
			<td align="left" x:str>
				<%= oitem.FItemList(i).Fitemoptionname %>
			</td>
		    <td align="center" x:num="<%= oitem.FItemList(i).FjupsuCNT %>" >
				<%= oitem.FItemList(i).FjupsuCNT %>
		    </td>
		    <td align="center" x:num="<%= oitem.FItemList(i).FipkumCNT %>" >
				<%= oitem.FItemList(i).FipkumCNT %>
		    </td>
		    <td align="center" x:num="<%= oitem.FItemList(i).FnotifyCNT %>" >
				<%= oitem.FItemList(i).FnotifyCNT %>
		    </td>
		    <td align="center" x:num="<%= oitem.FItemList(i).FconfirmCNT %>" >
				<%= oitem.FItemList(i).FconfirmCNT %>
		    </td>
		</tr>
			<%
			jupsuSUM = jupsuSUM + oitem.FItemList(i).FjupsuCNT
			ipkumSUM = ipkumSUM + oitem.FItemList(i).FipkumCNT
			notifySUM = notifySUM + oitem.FItemList(i).FnotifyCNT
			confirmSUM = confirmSUM + oitem.FItemList(i).FconfirmCNT
			%>
		<% next %>
		<tr class="a" height="40" bgcolor="#FFFFFF">
			<td align="center" colspan="3" x:str></td>
		    <td align="center" x:num="<%= jupsuSUM %>">
				<%= jupsuSUM %>
		    </td>
		    <td align="center" x:num="<%= ipkumSUM %>">
				<%= ipkumSUM %>
		    </td>
		    <td align="center" x:num="<%= notifySUM %>">
				<%= notifySUM %>
		    </td>
		    <td align="center" x:num="<%= confirmSUM %>">
				<%= confirmSUM %>
		    </td>
		</tr>
	</table>
<% end if %>

</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->