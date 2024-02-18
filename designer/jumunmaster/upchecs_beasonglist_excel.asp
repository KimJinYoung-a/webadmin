<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/designer_cs_baljucls.asp"-->

<%
''통합 .2008-05-20
If (session("ssBctId") = "") or (session("ssBctDiv") <> "9999" and session("ssBctDiv") > "9") then
    response.write "<script language='javascript'>alert('세션이 종료되었습니다.');</script>"
    dbget.close()	:	response.End
end if

dim ojumun
dim ix,sql
Dim listitemlist,listitem,listitemcount
dim iSall, SheetType

listitem =  Replace(request("orderserial"), " ", "")
iSall   =  requestCheckVar(request("isall"), 32)
SheetType  =  requestCheckVar(request("SheetType"), 32)

set ojumun = new CCSJumunMaster

ojumun.FRectOrderSerial = listitem
ojumun.FRectIsAll       = iSall
ojumun.FRectDesignerID = session("ssBctID")
ojumun.reDesignerCS_SelectBaljuList

dim i, j

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_CS" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
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
<title>송장파일</title>
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
<table width=1200 cellspacing=0 cellpadding=1 border=0>
<tr>
	<td align="center" x:str >접수구분</td>
    <% if (SheetType="V4") then %>
    <td align="center" height="25" x:str >일련번호</td>
    <% end if %>
	<td align="center" x:str >CS주문번호</td>
	<td align="center" x:str >접수일</td>
	<td align="center" x:str >구매자명</td>
	<td align="center" x:str >구매자전화</td>
	<td align="center" x:str >구매자핸드폰</td>
	<td align="center" x:str >수령인</td>
	<td align="center" x:str >수령인전화</td>
	<td align="center" x:str >수령인핸드폰</td>
	<td align="center" x:str >우편번호</td>
	<% if (SheetType="V2") then %>
	<td align="center" x:str >배송지주소</td>
	<% else %>
	<td align="center" x:str >배송지주소1</td>
	<td align="center" x:str >배송지주소2</td>
	<% end if %>
	<td align="center" x:str >상품코드</td>
	<td align="center" x:str >상품명</td>
	<td align="center" x:str >옵션</td>
	<td align="center" x:str >판매가</td>
	<td align="center" x:str >수량</td>
	<% if (SheetType="V3") or (SheetType="V4") then %>
	<td align="center" x:str >업체상품코드</td>
	<% end if %>
</tr>
<% for ix=0 to ojumun.FResultCount - 1 %>
<tr>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).Fdivcdname %></td>
    <% if (SheetType="V4") then %>
    <td align="center" x:str><%= ojumun.FMasterItemList(ix).Fcsdetailidx %></td>
    <% end if %>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FOrgOrderSerial %>-<%= ojumun.FMasterItemList(ix).Fcsmasteridx %></td>
	<td align="center" x:str><%= Left(CStr(ojumun.FMasterItemList(ix).FRegDate),10) %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FBuyName %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FBuyPhone %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FBuyHp %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqName %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqPhone %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqHp %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqZipCode %></td>
	<% if (SheetType="V2") then %>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqZipAddr %><%=chr(32)%><%= ojumun.FMasterItemList(ix).FReqAddress %></td>
	<% else %>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqZipAddr %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqAddress %></td>
	<% end if %>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).Fitemid %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FItemName %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FItemoptionName %></td>
	<td align="center" x:num="<%= ojumun.FMasterItemList(ix).FItemCost %>" ><%= ojumun.FMasterItemList(ix).FItemCost %></td>
	<td align="center" x:num="<%= ojumun.FMasterItemList(ix).FItemNo %>" ><%= ojumun.FMasterItemList(ix).FItemNo %></td>
	<% if (SheetType="V3") or (SheetType="V4") then %>
	<td align="center" x:str ><%= ojumun.FMasterItemList(ix).FupcheManageCode %></td>
	<% end if %>
</tr>
<% next %>
</table>
</body>
</html>
<%
set ojumun = Nothing
set oGift = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->