<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  온라인 출고지시
' History : 2020.07.08 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_baljuipgocls.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_baljucls.asp" -->
<!-- #include virtual="/lib/BarcodeFunction.asp" -->
<%
dim baljukey ,section ,obaljupage, sitebaljukey, i, menupos
	baljukey = requestcheckvar(getNumeric(request("baljukey")),10)
	sitebaljukey = requestcheckvar(getNumeric(request("sitebaljukey")),10)
    menupos = requestcheckvar(getNumeric(request("menupos")),10)

set obaljupage = new CBaljuIpgo
	obaljupage.FRectBaljuKey = sitebaljukey
	obaljupage.GetBaljuIpgoitem

%>
<script type='text/javascript'>

function jsLogisticsBaljuitem_excel(sitebaljukey) {
	frmview.action="/admin/ordermaster/pop_logistics_baljuitem_excel.asp?sitebaljukey=" + sitebaljukey + "&menupos=<%= menupos %>"
	frmview.target="view";
	frmview.sitebaljukey.value=sitebaljukey;
	frmview.submit();
}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left"></td>
    <td align="right">
        <input type="button" onclick="jsLogisticsBaljuitem_excel('<%= sitebaljukey %>');" value="엑셀다운로드" class="button">
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="3" bgcolor="FFFFFF">
	<td colspan="4">
		검색결과 : <b><%= obaljupage.FTotalCount%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>물류코드</td>
	<td>수량</td>
	<td>브랜드ID</td>
    <td>재고구분</td>
</tr>
<% if obaljupage.FResultCount>0 then %>
    <% for i = 0 to obaljupage.FResultCount-1 %>
    <tr class="a" height="25" bgcolor="#FFFFFF" align="center">
        <td><%= BF_MakeTenBarcode(obaljupage.FItemList(i).fitemgubun,obaljupage.FItemList(i).fitemid,obaljupage.FItemList(i).fitemoption) %></td>
        <td><%= obaljupage.FItemList(i).fitemno %></td>
        <td><%= obaljupage.FItemList(i).fmakerid %></td>
        <td><%= obaljupage.FItemList(i).FwarehouseCd %></td>
    </tr>
    <% next %>
<% else %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="4" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>

</table>

<form name="frmview" method="get" action="" style="margin:0px;">
<input type="hidden" name="sitebaljukey" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
</form>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width="100%" height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>

<%
set obaljupage = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
