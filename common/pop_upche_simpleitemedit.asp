<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ����
' Hieditor : 2009.04.07 ������ ����
'			 2011.04.28 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<%
dim itemid ,i
	itemid = requestCheckvar(request("itemid"),10)  ''requestCheckvar 2016/02/11

if itemid = "" then
	response.write "<script>"
	response.write "	alert('��ǰ�ڵ尡 �����ϴ�');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
end if


'####### ��ǰ��ù��� ���� ������üũ
If IsNumeric(itemid) = false Then
	response.write "<script>"
	response.write "	alert('�߸��� ��ǰ�ڵ��Դϴ�');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
End IF
Dim vQuery, vIsOK
''vQuery = "EXEC [db_item].[dbo].[sp_Ten_ItemNotificationRaw_Check] '" & itemid & "'"
''rsget.open vQuery,dbget,1
''2015/06/18 ������
vQuery = "[db_item].[dbo].[sp_Ten_ItemNotificationRaw_Check]('" & itemid & "')"
rsget.CursorLocation = adUseClient
rsget.Open vQuery, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
If Not rsget.Eof Then
	vIsOK = rsget(0)
Else
	vIsOK = "x"
End IF
rsget.close()
'rw vIsOK
'####### ��ǰ��ù��� ���� ������üũ


dim oitem
set oitem = new CItemInfo

oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItemInfo
end if

''2016/02/11 �߰�.
if (oitem.FResultCount<1) then
    response.write "<script>"
	response.write "	alert('�߸��� ��ǰ�ڵ��̰ų� �ش��ǰ�� �����ϴ�.');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if
%>

<script language='javascript'>

function EnabledCheck(comp){
	var frm = document.frm2;

	for (i = 0; i < frm.elements.length; i++) {
		  var e = frm.elements[i];
		  if ((e.type == 'text') && ((e.name.substring(0,"optlimitno".length) == "optlimitno")||(e.name.substring(0,"optlimitsold".length) == "optlimitsold"))) {
				e.disabled = (comp.value=="N");
		  }
  	}

}

function SaveItem(frm){
	frm.itemoptionarr.value = ""
	frm.optlimitnoarr.value = ""
	frm.optlimitsoldarr.value = ""
	frm.optisusingarr.value = ""

    var option_isusing_count = 0;
	for (i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];
		if ((e.type == 'text')||(e.type == 'radio')) {
		  	if ((e.name.substring(0,"optlimitno".length)) == "optlimitno"){

		  	    if (!IsDigit(e.value)){
		  	        alert('���������� ���ڸ� �����մϴ�.');
		  	        e.focus();
		  	        return;
		  	    }

				frm.itemoptionarr.value = frm.itemoptionarr.value + e.id + "," ;
				frm.optlimitnoarr.value = frm.optlimitnoarr.value + e.value + "," ;

				if (e.id == "0000") {
				    option_isusing_count = 1;
                }
		  	}

		  	if ((e.name.substring(0,"optlimitsold".length)) == "optlimitsold") {
		  	    if (!IsDigit(e.value)){
		  	        alert('���������� ���ڸ� �����մϴ�.');
		  	        e.focus();
		  	        return;
		  	    }

				frm.optlimitsoldarr.value = frm.optlimitsoldarr.value + e.value + "," ;
			}

			if ((e.name.substring(0,"optisusing".length)) == "optisusing") {
				if (e.checked) {
					if (e.value == "Y") {
					    option_isusing_count = option_isusing_count + 1;
                    }
					frm.optisusingarr.value = frm.optisusingarr.value + e.value + "," ;
				}
			}
		}
  	}
    if (option_isusing_count < 1) {
        alert("��� �ɼ��� ���������� �Ҽ� �����ϴ�. ��ǰ������ ���������� �����ϰų�, ���þ��� �����ϼ���.");
        //alert(frm.itemoptionarr.value);
        return;
    }

<% if (oitem.FOneItem.Fmwdiv <> "U") then %>
    if (frm.reqstring.value == "") {
        alert("���������� �Է����ּ���.");
        return;
    }
<% end if %>

<%
	If vIsOK = "x" Then
		If oitem.FOneItem.FSellYn <> "Y" Then
%>
			if(frm.sellyn[0].checked)
			{
				alert("��ǰ��ó����� ��� �ԷµǾ� ���� ���� �����Դϴ�.\n��� �Է��ϼž� �Ǹ������� ���� �����մϴ�.\n��� �Է��Ͻ� �� �� â�� ���� ���ðų� ���ΰ�ħ �Ͻø� ������ �����մϴ�.");
				return;
			}
<%
		End If
	End If
%>

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if(ret){
		frm.submit();
	}
}

function PopOptionEdit(itemid){
	var popwin = window.open('/common/pop_adminitemoptionedit.asp?itemid=' + itemid,'PopOptionEdit','width=700 height=500 scrollbars=yes resizable=yes');
	popwin.focus();
}

function CloseWindow() {
    window.close();
}

function editItemInfo(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/designer/itemmaster/upche_item_infomodify.asp?' + param ,'editItemInfo','width=1100,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<tr>
	<td align="left">
		<strong>��ǰ���� ����</strong><br>
		<br>- �ٹ�(�ٹ����ٹ��)��ǰ�� �ٹ����� Ȯ���� <font color=red>������ ���������� �ݿ�</font>�˴ϴ�.
		<br>- ����(��ü���) ��ǰ�� ��� <font color=red>��ùݿ�</font>�˴ϴ�.
		<br>- �����̳�, ��ǰ�� �� ��Ÿ ���� �Ͻ� ������ <font color=red>��翥��</font>���� ������ �ּ���.
	</td>
	<td align="right">
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<% if oitem.FResultCount>0 then %>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name="frm2" method="post" action="do_upche_simpleiteminfoedit.asp">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="optisusingarr" value="">
<input type="hidden" name="optlimitnoarr" value="">
<input type="hidden" name="optlimitsoldarr" value="">
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">��ǰ�ڵ�</td>
	<td width=76% bgcolor="#FFFFFF"><%= itemid %></td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">��ǰ��</td>
	<td width=76% bgcolor="#FFFFFF"><%= oitem.FOneItem.Fitemname %></td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">�귣��</td>
	<td bgcolor="#FFFFFF"><%= oitem.FOneItem.Fmakerid %> (<%= oitem.FOneItem.FBrandName %>)</td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">�ǸŰ�/���԰�</td>
	<td bgcolor="#FFFFFF">
	<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
	</td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">���Ա���</td>
	<td bgcolor="#FFFFFF">
	<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
	&nbsp;
	<% if oitem.FOneItem.FSellcash<>0 then %>
	<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
	<% end if %>
	</td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">���ɼ�</td>
	<td bgcolor="#FFFFFF">
	(<%= oitem.FOneItem.FOptionCnt %> ��)
	&nbsp;
	<% if oitem.FOneItem.IsUpcheBeasong then %>
	<input type=button value="�ɼǼ���" onclick="PopOptionEdit('<%= itemid %>');" class="button">
	<% else %>
	<font color=red>* �ɼ� �߰�/������ ���MD</font>���� �����ϼ���.
	<% end if %>
	</td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">��۱���</td>
	<td bgcolor="#FFFFFF">
	<% if oitem.FOneItem.IsUpcheBeasong then %>
	<b>��ü</b>���
	<% else %>
	�ٹ����ٹ��
	<% end if %>
	</td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">��ǰ ǰ������</td>
	<td bgcolor="#FFFFFF">
	<% if (oitem.FOneItem.IsSoldOut) or (oitem.FOneItem.FSellYn="S") then %>
	<font color=red><b>ǰ��</b></font>
	<% end if %>
	</td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">��ǰ �Ǹſ���</td>
	<td bgcolor="#FFFFFF">
	<% if oitem.FOneItem.FSellYn="Y" then %>
	<input type="radio" name="sellyn" value="Y" checked >�Ǹ���
	<input type="radio" name="sellyn" value="S" >�Ͻ�ǰ��
	<input type="radio" name="sellyn" value="N" >�Ǹž���
	<% elseif oitem.FOneItem.FSellYn="S" then %>
	<input type="radio" name="sellyn" value="Y" >�Ǹ���
	<input type="radio" name="sellyn" value="S" checked ><font color="blue">�Ͻ�ǰ��</font>
	<input type="radio" name="sellyn" value="N" >�Ǹž���
	<% else %>
	<input type="radio" name="sellyn" value="Y" >�Ǹ���
	<input type="radio" name="sellyn" value="S" >�Ͻ�ǰ��
	<input type="radio" name="sellyn" value="N" checked ><font color="red">�Ǹž���</font>
	<% end if %>
	<% If vIsOK = "x" Then %>
    	&nbsp;&nbsp;<input type="button" class="button" value="��ǰ��ó����Է�" style="width:110px;" onClick="editItemInfo('<%=itemid%>');">
	<% End If %>
	</td>
</tr>
<input type="hidden" name="isusing" value="<%= oitem.FOneItem.FIsUsing %>">
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">�����Ǹſ���</td>
	<td bgcolor="#FFFFFF">
	<% if oitem.FOneItem.FLimitYn="Y" then %>
	<input type="radio" name="limityn" value="Y" checked onclick="EnabledCheck(this)"><font color="blue">�����Ǹ�</font>
	<input type="radio" name="limityn" value="N" onclick="EnabledCheck(this)">�������Ǹ�
	<% else %>
	<input type="radio" name="limityn" value="Y" onclick="EnabledCheck(this)">�����Ǹ�
	<input type="radio" name="limityn" value="N" checked onclick="EnabledCheck(this)">�������Ǹ�
	<% end if %>
	</td>
</tr>
<tr>
	<td colspan="2" height="25" bgcolor="#FFFFFF">
		<table width="100%" cellspacing=1 cellpadding=1 class=a bgcolor=#BABABA>
		<tr bgcolor="#FFDDDD">
			<td height="25">�ɼǸ�</td>
			<td width="100">�ɼǻ�뿩��</td>
			<td>�������� - �Ǹż��� = �������</td>
			<td width="40">���</td>
		</tr>
		<% if oitemoption.FResultCount>0 then %>
			<% for i=0 to oitemoption.FResultCount - 1 %>
				<% if oitemoption.FITemList(i).FOptIsUsing="N" then %>
				<tr bgcolor="#EEEEEE">
				<% else %>
				<tr bgcolor="#FFFFFF">
				<% end if %>
					<td height="25"><%= oitemoption.FITemList(i).FOptionName %>(<%= oitemoption.FITemList(i).FItemOption %>)</td>
					<td>
						<% if oitemoption.FITemList(i).Foptisusing="Y" then %>
						<input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" checked >����� <input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" >������
						<% else %>
						<input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" >����� <input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" checked ><font color="red">������</font>
						<% end if %>
					</td>
					<td>
					<input type="text" id="<%= oitemoption.FITemList(i).FItemOption %>" name="optlimitno<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).FOptLimitNo %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
					-
					<input type="text" id="<%= oitemoption.FITemList(i).FItemOption %>" name="optlimitsold<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).FOptLimitSold %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
					=
					<input type="text" name="optremainno<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).GetOptLimitEa %>" size="4" maxlength=5 disabled >
				</td>
				<td>
				<% if (oitemoption.FITemList(i).FOptIsUsing="N") or (oitemoption.FITemList(i).Foptsellyn="N") or (oitemoption.FITemList(i).Foptlimityn="Y" and oitemoption.FITemList(i).GetOptLimitEa<1) then %>
				<font color=red>ǰ��</font>
				<% end if %>
				</td>
				</tr>
			<% next %>
		<% else %>
			<tr bgcolor="#FFFFFF">
				<td height="25" colspan="2">�ɼǾ��� (0000)</td>
				<td>
				<input type="text" id="0000" name="optlimitno" value="<%= oitem.FOneItem.FLimitNo %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
				-
				<input type="text" id="0000" name="optlimitsold" value="<%= oitem.FOneItem.FLimitSold %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
				=
				<input type="text" name="optremainno" value="<%= oitem.FOneItem.GetLimitEa %>" size="4" maxlength=5 disabled >
			</td>
			<td>
			    <% if oitem.FOneItem.isSoldOut() then %>
			    <font color=red>ǰ��</font>
			    <% end if %>
			</td>
			</tr>
		<% end if %>
		</table>
	</td>
</tr>
<input type="hidden" name="pojangok" value="<%= oitem.FOneItem.FPojangOK %>">
<tr>
	<td width=80 bgcolor="#DDDDFF">�̹���</td>
	<td bgcolor="#FFFFFF">
	<img src="<%= oitem.FOneItem.FListImage %>" width=100>
	</td>
</tr>
<% if (oitem.FOneItem.Fmwdiv <> "U") then %>
<tr>
	<td width=80 bgcolor="#DDDDFF">��������</td>
	<td bgcolor="#FFFFFF">
	  <input type="text" name="reqstring" value="" size="30"><br>(ex: ����, ����Ͻú���(�԰����� 2003-05-15), ���԰�..)
	</td>
</tr>
<% end if %>
<tr>
	<td bgcolor="#FFFFFF" colspan=2 align="center">
		<% if (oitem.FOneItem.Fmwdiv = "U") then %>
      		<input type="button" value="�����ϱ�" onclick="SaveItem(frm2)" class="button">
		<% else %>
     		<input type="button" value="������û" onclick="SaveItem(frm2)" class="button">
		<% end if %>
		<input type="button" value=" �� �� " onclick="CloseWindow()" class="button">
	</td>
</tr>
</form>
</table>
<% end if %>

<%
set oitemoption = Nothing
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->