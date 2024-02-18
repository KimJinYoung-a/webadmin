<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/shoplinker/shoplinkercls.asp"-->
<%
Dim oshoplinker, page, i
page = request("page")

If page = "" Then page = 1
	
SET oshoplinker = new CShoplinker
	oshoplinker.FCurrPage = page
	oshoplinker.FPageSize = 20
	oshoplinker.getShoplinkerOutmallList
%>
<script language="javascript">
function OutmallSettingREG(){
	var popwin3=window.open('/admin/etc/shoplinker/popOutmallsettingREG.asp','REGoutmall','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin3.focus();
}
function OutmallSettingEDT(makerid){
	var popwin3=window.open('/admin/etc/shoplinker/popOutmallsettingREG.asp?mode=U&makerid='+makerid+'','EDToutmall','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin3.focus();
}
function goPage(pg){
	var frm = document.frm;
    frm.page.value = pg;
    frm.submit();
}
</script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<input type="button" class="button" value="등록" onclick="OutmallSettingREG()">
<p>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr align="center" bgcolor="#F3F3FF" height="20">
	<td width="18%">브랜드ID</td>
	<td width="18%">제휴몰 어드민 ID</td>
	<td width="28%">브랜드명</td>
	<td width="36%">조건</td>
</tr>
<%
If oshoplinker.FResultCount > 0 Then
	For i=0 to oshoplinker.FResultCount - 1
%>
<tr align="center" bgcolor="#FFFFFF" height="20" onclick="OutmallSettingEDT('<%= oshoplinker.FItemList(i).FMakerid %>');" style="cursor:pointer;">
	<td width="18%"><%= oshoplinker.FItemList(i).FMakerid %></td>
	<td width="18%"><%= oshoplinker.FItemList(i).FMall_user_id %></td>
	<td width="28%"><%= oshoplinker.FItemList(i).FMall_name %></td>
	<td width="36%">
		<%= oshoplinker.FItemList(i).FDefaultFreeBeasongLimit %>원 미만 구매시 배송료<%=oshoplinker.FItemList(i).FDefaultDeliverPay%>원
	</td>
</tr>
<%  Next %>
<tr height="20">
    <td colspan="17" align="center" bgcolor="#FFFFFF">
        <% If oshoplinker.HasPreScroll then %>
		<a href="javascript:goPage('<%= oshoplinker.StartScrollPage-1 %>');">[pre]</a>
    	<% Else %>
    		[pre]
    	<% End If %>

    	<% For i = 0 + oshoplinker.StartScrollPage to oshoplinker.FScrollCount + oshoplinker.StartScrollPage - 1 %>
    		<% If i>oshoplinker.FTotalpage Then Exit For %>
    		<% If CStr(page) = CStr(i) Then %>
    		<font color="red">[<%= i %>]</font>
    		<% Else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% End If %>
    	<% Next %>

    	<% If oshoplinker.HasNextScroll Then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% Else %>
    		[next]
    	<% End If %>
    </td>
</tr>
<% Else %>
<tr bgcolor="#FFFFFF" height="50" align="center">
    <td colspan="17">등록된 브랜드가 없습니다.</td>
</tr>
<% End If %>
</table>
<% SET oshoplinker = nothing %>
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
