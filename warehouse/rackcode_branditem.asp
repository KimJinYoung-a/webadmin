<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/rackipgocls.asp"-->

<%
dim makerid, mwdiv, sellyn, isusing, diffrackcode, research
dim upbaeZeroStock
makerid = request("makerid")
mwdiv = request("mwdiv")
sellyn = request("sellyn")
isusing = request("isusing")
diffrackcode = request("diffrackcode")
research = request("research")
upbaeZeroStock = request("upbaeZeroStock")

dim i
if mwdiv="" then mwdiv="MW"
if (research="") and (diffrackcode="") then diffrackcode="on"
if (research="") and (isusing="") then isusing="Y"

if (upbaeZeroStock = "on") then
	mwdiv = "U"
end if

dim orackcode_branditem
set orackcode_branditem = new CRackIpgo
orackcode_branditem.FRectMakerid = makerid
orackcode_branditem.FRectMwdiv = mwdiv
orackcode_branditem.FRectSellYN = sellyn
orackcode_branditem.FRectIsUsingYN = isusing
orackcode_branditem.FRectdiffrackcode = diffrackcode
''orackcode_branditem.FRectUpbaeZeroStock = upbaeZeroStock
orackcode_branditem.GetRackBrandItemList

%>

<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
}

function PopBrandInfo(v){
    PopBrandInfoEdit(v);

	//var popwin = window.open("/admin/lib/popbrandinfoonly.asp?designer=" + v,"popbrandinfoonly","width=640 height=580 scrollbars=yes resizable=yes");
	//popwin.focus();
}

</script>


<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>
	     	&nbsp;
			���:<% drawSelectBoxUsingYN "isusing", isusing %>
			&nbsp;
	     	�ŷ�����:<% drawSelectBoxMWU "mwdiv", mwdiv %>
	     	&nbsp;
        	<input type=checkbox name="diffrackcode" value="on" <% if diffrackcode="on" then response.write "checked" %>>�귣�巢�ڵ�� ������ ��ǰ��
			&nbsp;
        	<input type=checkbox name="upbaeZeroStock" value="on" <% if upbaeZeroStock="on" then response.write "checked" %>>���ڵ��ϵ� ������ ��ǰ��(����)
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="120">�귣��ID</td>
    	<td width="40">��ǰID</td>
    	<td width="50">�귣��<br>���ڵ�</td>
    	<td width="50">��ǰ<br>���ڵ�</td>
    	<td width="50">�̹���</td>
    	<td>��ǰ��</td>

    	<td width="30">�ŷ�<br>����</td>

		<td width="30">�Ǹ�<br>����</td>
		<td width="30">���<br>����</td>
		<td width="30">����<br>����</td>

		<td>���</td>
    </tr>
<% for i=0 to orackcode_branditem.FResultCount - 1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><a href="javascript:PopBrandInfo('<%= orackcode_branditem.FItemList(i).Fmakerid %>')"><%= orackcode_branditem.FItemList(i).Fmakerid %></a></td>
		<td><a href="javascript:PopItemSellEdit('<%= orackcode_branditem.FItemList(i).Fitemid %>');"><%= orackcode_branditem.FItemList(i).Fitemid %></a></td>
		<td><%= orackcode_branditem.FItemList(i).Frackcode %></td>
		<td>
			<% if (orackcode_branditem.FItemList(i).Fitemrackcode <> orackcode_branditem.FItemList(i).Frackcode) then %>
			<b><font color="red"><%= orackcode_branditem.FItemList(i).Fitemrackcode %></font></b>
			<% else %>
			<%= orackcode_branditem.FItemList(i).Fitemrackcode %>
			<% end if %>
		</td>
		<td><img src="<%= orackcode_branditem.FItemList(i).Fimgsmall %>" width=50 height=50></td>
		<td align="left"><%= orackcode_branditem.FItemList(i).Fitemname %></td>

		<td><font color="<%= mwdivColor(orackcode_branditem.FItemList(i).Fmwdiv) %>"><%= mwdivName(orackcode_branditem.FItemList(i).Fmwdiv) %></font></td>


		<td><font color="<%= yncolor(orackcode_branditem.FItemList(i).Fsellyn) %>"><%= orackcode_branditem.FItemList(i).Fsellyn %></font></td>
		<td><font color="<%= yncolor(orackcode_branditem.FItemList(i).FIsusing) %>"><%= orackcode_branditem.FItemList(i).FIsusing %></font></td>
		<td><font color="<%= yncolor(orackcode_branditem.FItemList(i).Flimityn) %>"><%= orackcode_branditem.FItemList(i).Flimityn %></font></td>
		<td></td>
	</tr>
<% next %>
</table>




<%
set orackcode_branditem = Nothing
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
