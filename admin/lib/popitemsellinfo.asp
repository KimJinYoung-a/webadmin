<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<!-- include virtual="/lib/classes/realjaegocls.asp"-->

<%
dim itemid
itemid = request("itemid")


response.redirect "/common/pop_simpleitemedit.asp?itemid=" + CStr(itemid)
''������

dim ojaego
set ojaego = new CRealJaeGo
ojaego.FRectItemID = itemid

if itemid<>"" then
	ojaego.GetItemInfoWithDailyRealJaeGo
end if

dim i
%>
<script language='javascript'>

function EnabledCheck(comp){
	var frm = document.frm2;

	if (comp.value=="Y"){
		frm.limitno.disabled = false;
		frm.limitsold.disabled = false;
	}else{
		frm.limitno.disabled = true;
		frm.limitsold.disabled = true;
	}
}

function SaveItem(frm){
	if ((frm.itemrackcode.value.length>0)&&(frm.itemrackcode.value.length!=6)){
		alert('��ǰ ���ڵ�� 6�ڸ��� �����Ǿ��ֽ��ϴ�.');
		frm.itemrackcode.focus();
		return;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if(ret){
		frm.submit();
	}
}

function popoptionEdit(iid){
	var popwin = window.open('/admin/shopmaster/popitemoptionedit.asp?menupos=239&itemid=' + iid,'popitemoptionedit','width=440 height=500 scrollbars=yes resizable=yes');
	popwin.focus();
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	��ǰ�ڵ� : <input type="text" name="itemid" value="<%= itemid %>" Maxlength="12" size="12">
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->


<% if ojaego.FResultCount>0 then %>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name=frm2 method=post action="doitemsellinfo.asp">
<input type=hidden name=itemid value="<%= itemid %>">
	<tr bgcolor="#FFFFFF">
		<td width=90 bgcolor="#DDDDFF">��ǰ�ڵ�</td>
		<td><%= itemid %></td>
		<td width=100 rowspan=5><img src="<%= ojaego.FITemList(0).FImageList %>" width=100></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width=80 bgcolor="#DDDDFF">�귣��</td>
		<td><%= ojaego.FITemList(0).Fmakerid %> (<%= ojaego.FITemList(0).FBrandName %>)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width=80 bgcolor="#DDDDFF">�ǸŰ�/���԰�</td>
		<td>
		    <%= FormatNumber(ojaego.FITemList(0).FSellcash,0) %> / <%= FormatNumber(ojaego.FITemList(0).FBuycash,0) %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width=80 bgcolor="#DDDDFF">���Ա���</td>
		<td>
		<font color="<%= ojaego.FITemList(0).getMwDivColor %>"><%= ojaego.FITemList(0).getMwDivName %></font>
		&nbsp;
		<% if ojaego.FITemList(0).FSellcash<>0 then %>
		<%= CLng((1- ojaego.FITemList(0).FBuycash/ojaego.FITemList(0).FSellcash)*100) %> %
		<% end if %>
    	</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width=80 bgcolor="#DDDDFF">���ڵ�</td>
		<td>
		<input type="text" name="itemrackcode" value="<%= ojaego.FITemList(0).FitemRackCode %>" size="6" maxlength="6" > (6�ڸ� Fix)
    	</td>
	</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
	<tr>
		<td bgcolor="#FFFFFF" colspan="2">
		<table width="100%" cellspacing=1 cellpadding=1 class=a bgcolor=#BABABA>
			<tr bgcolor="#FFFFFF">
				<td colspan="4" align="right">
				����������Ʈ (<%= ojaego.FITemList(0).FLastupdate %>)
				</td>
			</tr>
			<tr bgcolor="#FFDDDD">
				<td>��ǰ��</td>
				<td>�ɼǸ�</td>
				<td>OLD SYS</td>
				<td>NEW SYS</td>
			</tr>
		<% for i=0 to ojaego.FResultCount - 1 %>
			<% if ojaego.FITemList(i).FOptionUsing="N" then %>
			<tr bgcolor="#DDDDDD">
			<% else %>
			<tr bgcolor="#FFFFFF">
			<% end if %>
				<td><%= ojaego.FITemList(i).FItemName %></td>
				<td><%= ojaego.FITemList(i).FItemOptionName %></td>
				<td><%= ojaego.FITemList(i).Foldstockcurrno %></td>
				<td><%= ojaego.FITemList(i).GetCheckStockNo %></td>
			</tr>
		<% next %>
		</table>
		</td>
	</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
	<tr>
		<td width=90 bgcolor="#DDDDFF">�ɼ�</td>
		<td bgcolor="#FFFFFF">
		(<%= ojaego.FITemList(0).FOptionCnt %>)
		<input type=button value="�ɼǼ���" onclick="popoptionEdit('<%= itemid %>');">
		</td>
	</tr>

	<tr>
		<td width=80 bgcolor="#DDDDFF">��۱���</td>
		<td bgcolor="#FFFFFF">
		<% if ojaego.FITemList(0).IsUpcheBeasong then %>
		<b>��ü</b>���
		<% else %>
		�ٹ����ٹ��
		<% end if %>
		</td>
	</tr>

	<tr>
		<td width=80 bgcolor="#DDDDFF">SoldOut</td>
		<td bgcolor="#FFFFFF">
		<% if ojaego.FITemList(0).IsSoldOut then %>
		<font color=red><b>Sold Out</b></font>
		<% end if %>
		</td>
	</tr>
	<tr>
		<td width=80 bgcolor="#DDDDFF">�����Ǹſ���</td>
		<td bgcolor="#FFFFFF">
		<% if ojaego.FITemList(0).FLimitYn="Y" then %>
		<input type="radio" name="limityn" value="Y" checked onclick="EnabledCheck(this)">�����Ǹ�
		<input type="radio" name="limityn" value="N" onclick="EnabledCheck(this)">�������Ǹ�
		<% else %>
		<input type="radio" name="limityn" value="Y" onclick="EnabledCheck(this)">�����Ǹ�
		<input type="radio" name="limityn" value="N" checked onclick="EnabledCheck(this)">�������Ǹ�
		<% end if %>
		</td>
	</tr>
	<tr>
		<td width=80 bgcolor="#DDDDFF">����������</td>
		<td bgcolor="#FFFFFF"><input type="text" name="limitno" value="<%= ojaego.FITemList(0).FLimitNo %>" size="5" maxlength=5 <% if ojaego.FITemList(0).FLimitYn="N" then response.write "disabled" %> >��</td>
	</tr>
	<tr>
		<td width=80 bgcolor="#DDDDFF">�Ǹŵȼ���</td>
		<td bgcolor="#FFFFFF"><input type="text" name="limitsold" value="<%= ojaego.FITemList(0).FLimitSold %>" size="5" maxlength=5 <% if ojaego.FITemList(0).FLimitYn="N" then response.write "disabled" %> >��</td>
	</tr>
	<tr>
		<td width=80 bgcolor="#DDDDFF">��������</td>
		<td bgcolor="#FFFFFF"><input type="text" name="remainno" value="<%= ojaego.FITemList(0).GetLimitEa %>" size="5" maxlength=5 disabled >��</td>
	</tr>
	<tr>
		<td width=80 bgcolor="#DDDDFF">���ÿ���</td>
		<td bgcolor="#FFFFFF">
		<% if ojaego.FITemList(0).FDispYn="Y" then %>
		<input type="radio" name="dispyn" value="Y" checked >������
		<input type="radio" name="dispyn" value="N" >���þ���
		<% else %>
		<input type="radio" name="dispyn" value="Y" >������
		<input type="radio" name="dispyn" value="N" checked ><font color="red">���þ���</font>
		<% end if %>
		</td>
	</tr>
	<tr>
		<td width=80 bgcolor="#DDDDFF">�Ǹſ���</td>
		<td bgcolor="#FFFFFF">
		<% if ojaego.FITemList(0).FSellYn="Y" then %>
		<input type="radio" name="sellyn" value="Y" checked >�Ǹ���
		<input type="radio" name="sellyn" value="N" >�Ǹž���
		<% else %>
		<input type="radio" name="sellyn" value="Y" >�Ǹ���
		<input type="radio" name="sellyn" value="N" checked ><font color="red">�Ǹž���</font>
		<% end if %>
		</td>
	</tr>
	<tr>
		<td width=80 bgcolor="#DDDDFF">��뿩��</td>
		<td bgcolor="#FFFFFF">
		<% if ojaego.FITemList(0).FIsUsing="Y" then %>
		<input type="radio" name="isusing" value="Y" checked >�����
		<input type="radio" name="isusing" value="N" >������
		<% else %>
		<input type="radio" name="isusing" value="Y" >�����
		<input type="radio" name="isusing" value="N" checked ><font color="red">������</font>
		<% end if %>
		</td>
	</tr>
	<input type=hidden name="pojangok" value="<%= ojaego.FITemList(0).FPojangOK %>">
<!--
	<tr>
		<td width=80 bgcolor="#DDDDFF">���忩��</td>
		<td bgcolor="#FFFFFF">
		<% if ojaego.FITemList(0).FPojangOK="Y" then %>
		<input type="radio" name="pojangok" value="Y" checked >���尡��
		<input type="radio" name="pojangok" value="N" >����Ұ�
		<% else %>
		<input type="radio" name="pojangok" value="Y" >���尡��
		<input type="radio" name="pojangok" value="N" checked ><font color="red">����Ұ�</font>
		<% end if %>
		</td>
	</tr>
-->
</form>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center" bgcolor="F4F4F4"><input type="button" value="����" onclick="SaveItem(frm2)"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->


<% end if %>
<%
set ojaego = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->