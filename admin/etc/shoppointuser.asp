<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/offshoppointcls.asp"-->
<%
dim opointuser, targetpointuser,mode
opointuser = request("opointuser")
targetpointuser = request("targetpointuser")
mode = request("mode")

dim sqlStr
if mode="updatepoint" then
	sqlStr = " update [db_shop].[dbo].tbl_shop_pointlog" + VbCrlf
	sqlStr = sqlStr + " set pointuserno='" + targetpointuser + "'" + VbCrlf
	sqlStr = sqlStr + " where pointuserno='" + opointuser + "'" + VbCrlf

	rsget.Open sqlStr,dbget,1

	sqlStr = " update [db_shop].[dbo].tbl_shop_pointuser" + VbCrlf
	sqlStr = sqlStr + " set shoppoint=IsNULL(T.sumpoint,0)" + VbCrlf
	sqlStr = sqlStr + " from (" + VbCrlf
	sqlStr = sqlStr + " select sum(point) as sumpoint from [db_shop].[dbo].tbl_shop_pointlog" + VbCrlf
	sqlStr = sqlStr + " where deleteyn='N'" + VbCrlf
	sqlStr = sqlStr + " and pointuserno='" + targetpointuser + "'" + VbCrlf
	sqlStr = sqlStr + " ) as T" + VbCrlf
	sqlStr = sqlStr + " where  [db_shop].[dbo].tbl_shop_pointuser.pointuserno='" + targetpointuser + "'"

	rsget.Open sqlStr,dbget,1
end if

dim apointuser
set apointuser = new COffShopPoint
apointuser.FRectPointuserNo = opointuser
apointuser.GetOffShopPointUser

dim opoint
set opoint= new COffShopPoint
opoint.FRectPointuserNo = opointuser
opoint.GetOffShopPointlog

dim i
%>
<script language='javascript'>
function searchPoint(){
	frm.submit();
}

function changePoint(){
	frm.targetpointuser.value = frmtarget.targetpointuser.value;
	frm.mode.value="updatepoint";

	if (frm.opointuser.value.length!=13){
		alert('not valid pointuser');
		return;
	}

	if (frm.targetpointuser.value.length!=13){
		alert('not valid targetpointuser');
		return;
	}

	frm.submit();
}
</script>

<table border=0 cellspacing=0 cellpadding=0>
<form name=frm method=get action="">
<input type=hidden name="menupos" value="<%= menupos %>">
<input type=hidden name="mode" value="">
<input type=hidden name="targetpointuser" value="">
<tr>
	<td><input type=text name="opointuser" value="<%= opointuser %>"></td>
	<td><input type=button value="검색" onclick="searchPoint()"></td>
</tr>
</form>
</table>
<br>
<table width=700 border=0 cellpadding=1 cellspacing="1"  class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF">
	<td>idx</td>
	<td>pointuserno</td>
	<td>주민번호</td>
	<td>등록샾</td>
	<td>성명</td>
	<td>현재포인트</td>
	<td>등록일</td>
</tr>
<% for i=0 to apointuser.FResultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td><%= apointuser.FItemList(i).Fidx %></td>
	<td><%= apointuser.FItemList(i).Fpointuserno %></td>
	<td><%= apointuser.FItemList(i).Fjuminno %></td>
	<td><%= apointuser.FItemList(i).Fregshopid %></td>
	<td><%= apointuser.FItemList(i).Fpointusername %></td>
	<td><%= apointuser.FItemList(i).Fshoppoint %></td>
	<td><%= apointuser.FItemList(i).Fregdate %></td>
</tr>
<% next %>
<% if apointuser.FResultCount<1 then %>
<form name=frmtarget>
<tr bgcolor="#FFFFFF">
	<td colspan=7 align=right><input type=text name="targetpointuser" value="<%= targetpointuser %>">
	<input type=button value="변경" onclick="changePoint()">
	</td>
</tr>
</form>
<% end if %>
</table>
<br>
<table width=700 border=0 cellpadding=1 cellspacing="1"  class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF">
	<td>idx</td>
	<td>pointuserno</td>
	<td>point</td>
	<td>shopid</td>
	<td>적요</td>
	<td>등록일</td>
	<td>삭제</td>
</tr>
<% for i=0 to opoint.FResultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td><%= opoint.FItemList(i).Fidx %></td>
	<td><%= opoint.FItemList(i).Fpointuserno %></td>
	<td><%= opoint.FItemList(i).Fpoint %></td>
	<td><%= opoint.FItemList(i).Fshopid %></td>
	<td><%= opoint.FItemList(i).Fjukyo %></td>
	<td><%= opoint.FItemList(i).Fregdate %></td>
	<td><%= opoint.FItemList(i).Fdeleteyn %></td>
</tr>
<% next %>
</table>
<%
set apointuser = Nothing
set opoint= Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->