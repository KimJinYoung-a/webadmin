<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 상품등록
' History : 서동석 생성
'			2017.04.12 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_confirmitemcls.asp"-->
<!-- #include virtual="/admin/offshop/shopitemmodi_top.asp"-->

<%
dim designer
dim acttype, ckonlyusing, ckonlyoff, imageview, pricediff
designer 	= requestCheckVar(request("designer"),32)
acttype 	= requestCheckVar(request("acttype"),10)
ckonlyusing	= requestCheckVar(request("ckonlyusing"),2)
ckonlyoff	= requestCheckVar(request("ckonlyoff"),2)
imageview	= requestCheckVar(request("imageview"),2)
pricediff	= requestCheckVar(request("pricediff"),10)

dim ooffitem
set ooffitem = new COffShopConfirm
ooffitem.FPageSize = 200
ooffitem.FRectDesigner = designer
ooffitem.GetOnOffDiffItemBrandList

dim i
%>
<script language='javascript'>
function SaveItems(frm){
	if (confirm('선택 브랜드를 변경하시겠습니까?')){
		frm.submit();
	}
}
</script>
<table width="98%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<tr>
		<td class="a" >
		업체:<% drawSelectBoxDesignerOffShopContract "designer",designer  %>
<!--
		<br>
		옵션:

		<input type="checkbox" name="ckonlyusing" value="on" <% if ckonlyusing="on" then response.write "checked" %> >사용중인상품만
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >이미지보기
-->
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="98%" border="0" cellspacing="1" cellpadding="2 bgcolor="#3d3d3d" class="a">
<% if ooffitem.FResultCount>0 then %>
<tr>
	<td colspan="6" align=right><input type=button value="저장" onclick="SaveItems(frmarr);"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=20><input type="checkbox" name="ckall" onClick="AnSelectAll(frmarr,this.checked)"></td>
	<td>상품코드</td>
	<td>ON브랜드</td>
	<td>OFF브랜드</td>
	<td>상품명</td>
</tr>
<% else %>
<tr>
	<td colspan="6" align=center> [ 검색결과가 없습니다. ] </td>
</tr>
<% end if %>

<form name="frmarr" method=post action="shopitem_process.asp">
<input type="hidden" name="mode" value="makeridmodiarr">
<% for i=0 to ooffitem.FResultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" value="<%= ooffitem.FItemList(i).FShopItemId %>"></td>
	<td><%= ooffitem.FItemList(i).FShopItemId %></td>
	<td><%= ooffitem.FItemList(i).Fonlinemakerid %></td>
	<td><%= ooffitem.FItemList(i).FMakerid %></td>
	<td><%= ooffitem.FItemList(i).FShopItemName %></td>
</tr>
<% next %>
</table>
</form>
<%
set ooffitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->