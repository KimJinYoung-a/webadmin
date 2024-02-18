<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  매장 선택 팝업
' History : 2011.11.24 
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->
<%
dim ochargeuser ,i, shopdiv, isusing, research
shopdiv = requestCheckvar(request("shopdiv"),32)
isusing = requestCheckvar(request("isusing"),10)
research = requestCheckvar(request("research"),10)

if (research="") then isusing="Y"

set ochargeuser = new COffShopChargeUser
    ochargeuser.FRectShopDiv2 = shopdiv
    ochargeuser.FRectIsUsing = isusing
    ochargeuser.FRectNotProtoTypeShop ="on"
	ochargeuser.GetOffShopList
%>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	매장 구분 : <% Call DrawShopDivCombo("shopdiv",shopdiv) %>
	사용 구분 : <% Call drawSelectBoxUsingYN("isusing",isusing) %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">

	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		
		</td>
		<td align="right">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if ochargeuser.FresultCount > 0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ochargeuser.fresultcount %></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">매장ID</td>
	<td>매장명</td>
	<td width="80">구분</td>
	<td width="80">국가</td>
	<td width="50">사용여부</td>
	<td width="50">선택</td>
</tr>
<%
for i=0 to ochargeuser.FresultCount - 1
%>
<% if ochargeuser.FItemList(i).FIsUsing="N" then %>
<tr align="center" bgcolor="<%= adminColor("dgray") %>">
<% else %>
<tr align="center" bgcolor="#FFFFFF">
<% end if %>
	<td><%= ochargeuser.FItemList(i).Fuserid %></td>
	<td><%= ochargeuser.FItemList(i).Fshopname %></td>
	<td><%= ochargeuser.FItemList(i).GetShopdivName %></td>
	<td><%= ochargeuser.FItemList(i).FcountryNamekr %></td>
	<td><%= ochargeuser.FItemList(i).FIsUsing %></td>
	<td><input type="button" class="button" value="선택" onclick="opener.addSelectedShop('<%= ochargeuser.FItemList(i).Fuserid %>','<%= ochargeuser.FItemList(i).Fshopname %>')"></td>
</tr>
<%
next
else
%>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=10>검색 결과가 없습니다</td>
</tr>
<%
end if
%>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=10><input type="button" value="닫 기" onClick="window.close()"></td>
</tr>
</table>

<%
set ochargeuser = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->