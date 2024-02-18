<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ɼǰ���
' Hieditor : ������ ����
'			 2022.07.06 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/optionmanagecls.asp"-->
<%
dim cdl, cdm
cdl = request("cdl")
cdm = request("cdm")

dim pmode
dim mode, t_cdl, t_cdm
dim codename, optiondispyn, disporder
pmode = request("pmode")
mode = request("mode")
t_cdl = request("t_cdl")
t_cdm = request("t_cdm")

codename = db2html(request("codename"))
optiondispyn = request("optiondispyn")
disporder = request("disporder")

if not IsNUmeric(disporder) then disporder="0"

dim sqlstr
dim alreadyCodeExists
if mode="modismall" then
	if codename <> "" and not(isnull(codename)) then
		codename = ReplaceBracket(codename)
	end If

	sqlstr = "update [db_item].[dbo].tbl_option_div02"
	sqlstr = sqlstr + " set codeview='" + codename + "'"
	sqlstr = sqlstr + " , optiondispyn='" + optiondispyn + "'"
	sqlstr = sqlstr + " , disporder=" + disporder + ""
	sqlstr = sqlstr + " where optioncode01='" + t_cdl + "'"
	sqlstr = sqlstr + " and optioncode02='" + t_cdm + "'"

	dbget.execute sqlstr

	sqlstr = "update [db_item].[dbo].tbl_item_option"
	sqlstr = sqlstr + " set optionname=v.codeview"
	sqlstr = sqlstr + " from [db_item].[dbo].vw_all_option v"
	sqlstr = sqlstr + " where [db_item].[dbo].tbl_item_option.itemoption='" + t_cdl + t_cdm + "'"
	sqlstr = sqlstr + " and [db_item].[dbo].tbl_item_option.itemoption=v.optioncode"

	dbget.execute sqlstr

	response.write "<script type='text/javascript'>alert('���� �Ǿ����ϴ�.');</script>"
elseif mode="modilarge" then
	if codename <> "" and not(isnull(codename)) then
		codename = ReplaceBracket(codename)
	end If

	sqlstr = "update [db_item].[dbo].tbl_option_div01"
	sqlstr = sqlstr + " set codename='" + codename + "'"
	sqlstr = sqlstr + " , optiondispyn='" + optiondispyn + "'"
	sqlstr = sqlstr + " , disporder=" + disporder + ""
	sqlstr = sqlstr + " where optioncode01='" + t_cdl + "'"

	dbget.execute sqlstr
	response.write "<script type='text/javascript'>alert('���� �Ǿ����ϴ�.');</script>"
elseif mode="addlarge" then

	''check code already Exists
	sqlstr = "select count(optioncode01) as cnt"
	sqlstr = sqlstr + "  from [db_item].[dbo].tbl_option_div01 with (nolock)"
	sqlstr = sqlstr + "  where optioncode01='" + t_cdl + "'"

	rsget.CursorLocation = adUseClient
	rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
	alreadyCodeExists = rsget("cnt")>0
	rsget.Close

	if alreadyCodeExists then
		response.write "<script type='text/javascript'>alert('�ڵ尡 �̹������մϴ�.');history.back();</script>"
		dbget.close()	:	response.End
	end if

	if codename <> "" and not(isnull(codename)) then
		codename = ReplaceBracket(codename)
	end If

	sqlstr = "insert into [db_item].[dbo].tbl_option_div01"
	sqlstr = sqlstr + " (optioncode01,codename,optiondispyn,disporder)"
	sqlstr = sqlstr + " values ("
	sqlstr = sqlstr + " '" + t_cdl + "'"
	sqlstr = sqlstr + " ,'" + codename + "'"
	sqlstr = sqlstr + " ,'" + optiondispyn + "'"
	sqlstr = sqlstr + " ," + disporder + ""
	sqlstr = sqlstr + " )"

	'response.write sqlstr
	dbget.execute sqlstr
	response.write "<script type='text/javascript'>alert('���� �Ǿ����ϴ�.');</script>"
elseif mode="addmid" then
	''check code already Exists
	sqlstr = "select count(optioncode01) as cnt"
	sqlstr = sqlstr + "  from [db_item].[dbo].tbl_option_div02 with (nolock)"
	sqlstr = sqlstr + "  where optioncode01='" + t_cdl + "'"
	sqlstr = sqlstr + "  and optioncode02='" + t_cdm + "'"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
	alreadyCodeExists = rsget("cnt")>0
	rsget.Close

	if alreadyCodeExists then
		response.write "<script type='text/javascript'>alert('�ڵ尡 �̹������մϴ�.');history.back();</script>"
		dbget.close()	:	response.End
	end if

	if codename <> "" and not(isnull(codename)) then
		codename = ReplaceBracket(codename)
	end If

	sqlstr = "insert into [db_item].[dbo].tbl_option_div02"
	sqlstr = sqlstr + " (optioncode01,optioncode02,codevalue,codeview,optiondispyn,disporder)"
	sqlstr = sqlstr + " values ("
	sqlstr = sqlstr + " '" + t_cdl + "'"
	sqlstr = sqlstr + " ,'" + t_cdm + "'"
	sqlstr = sqlstr + " ,'" + codename + "'"
	sqlstr = sqlstr + " ,'" + codename + "'"
	sqlstr = sqlstr + " ,'" + optiondispyn + "'"
	sqlstr = sqlstr + " ," + disporder + ""
	sqlstr = sqlstr + " )"

	dbget.execute sqlstr
	response.write "<script type='text/javascript'>alert('���� �Ǿ����ϴ�.');</script>"
end if

dim ooption
set ooption = new COptionManager

if (cdl<>"") and (cdm<>"") then
	ooption.GetOption02 cdl,cdm
elseif (cdl<>"") then
	ooption.GetOption01 cdl
end if
%>
<script type='text/javascript'>
function saveOpt(frm){
	if (frm.codename.value.length<1){
		alert('�ڵ���� �����ּ���.');
		return;
	}

	//if (frm.disporder.value.length<1){
	//	alert('���Ǽ����� �����ּ���.');
	//	return;
	//}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

function addOptLarge(frm){
	if (frm.t_cdl.value.length!=2){
		alert('�ڵ�2�ڸ��� �����ּ���.');
		return;
	}

	if (frm.codename.value.length<1){
		alert('�ڵ���� �����ּ���.');
		return;
	}

	//if (frm.disporder.value.length<1){
	//	alert('���Ǽ����� �����ּ���.');
	//	return;
	//}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

function AddOptMid(frm){
	if (frm.t_cdm.value.length!=2){
		alert('�ڵ�2�ڸ��� �����ּ���.');
		return;
	}

	if (frm.codename.value.length<1){
		alert('�ڵ���� �����ּ���.');
		return;
	}

	//if (frm.disporder.value.length<1){
	//	alert('���Ǽ����� �����ּ���.');
	//	return;
	//}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

</script>
<% if (pmode<>"add") and (cdl<>"") and (cdm<>"") then %>
<form name=frm method=post action="" style="margin:0px;">
<input type="hidden" name="mode" value="modismall">
<input type="hidden" name="t_cdl" value="<%= cdl %>">
<input type="hidden" name="t_cdm" value="<%= cdm %>">
<table border=0 width=500 cellpadding="5" cellspacing="1" bgcolor="#CCCCCC" class="a">
<tr align=center>
	<td width=50>�ڵ�</td>
	<td width=50>�ڵ�02</td>
	<td width=100>�ڵ��</td>
	<td width=60>���</td>
	<td width=60>����</td>
</tr>
<tr align=center bgcolor="#FFFFFF">
	<td align=center><%= cdl %></td>
	<td align=center><%= cdm %></td>
	<td><input type=text name="codename" value="<%= ReplaceBracket(ooption.FItemList(0).Fcodeview) %>"  size="20" maxlength=20></td>
	<td>
		<select name=optiondispyn>
		<option value="Y" <% if ooption.FItemList(0).Foptiondispyn02="Y" then response.write "selected" %> >Y
		<option value="N" <% if ooption.FItemList(0).Foptiondispyn02="N" then response.write "selected" %> >N
		</select>
	</td>
	<td><input type=text name="disporder" value="<%= ooption.FItemList(0).Fdisporder02 %>" size="3" maxlength="3"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="5" align=center><input type="button" value="����" onclick="saveOpt(frm);"></td>
</tr>
</table>
</form>
<% elseif (pmode<>"add") and (cdl<>"") then %>
<form name=frm method=post action="" style="margin:0px;">
<input type="hidden" name="mode" value="modilarge">
<input type="hidden" name="t_cdl" value="<%= cdl %>">
<table border=0 width=500 cellpadding="5" cellspacing="1" bgcolor="#CCCCCC" class="a">
<tr align=center>
	<td width=50>�ڵ�01</td>
	<td width=90>�ڵ��</td>
	<td width=60>���</td>
	<td width=60>����</td>
</tr>
<tr align=center bgcolor="#FFFFFF">
	<td align=center><%= cdl %></td>
	<td><input type=text name="codename" value="<%= ReplaceBracket(ooption.FItemList(0).Fcodename) %>" size="16" maxlength=16></td>
	<td>
		<select name=optiondispyn>
		<option value="Y" <% if ooption.FItemList(0).Foptiondispyn01="Y" then response.write "selected" %> >Y
		<option value="N" <% if ooption.FItemList(0).Foptiondispyn01="N" then response.write "selected" %> >N
		</select>
	</td>
	<td><input type=text name="disporder" value="<%= ooption.FItemList(0).Fdisporder01 %>" size="3" maxlength="3"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="4" align=center><input type="button" value="����" onClick="saveOpt(frm);"></td>
</tr>
</table>
</form>
<% elseif (cdl<>"") then %>
<form name=frm method=post action="" style="margin:0px;">
<input type="hidden" name="mode" value="addmid">
<input type="hidden" name="t_cdl" value="<%= cdl %>">
<table border=0 width=500 cellpadding="5" cellspacing="1" bgcolor="#CCCCCC" class="a">
<tr align=center>
    <td width=50>�ڵ�01</td>
	<td width=50>�ڵ�02</td>
	<td width=90>�ڵ��</td>
	<td width=60>���</td>
	<td width=60>����</td>
</tr>
<tr align=center bgcolor="#FFFFFF">
	<td align=center><%= cdl %></td>
	<td align=center><input type="text" name="t_cdm" value="" size="2" maxlength="2"></td>
	<td><input type=text name="codename" value="" size="8" maxlength=16></td>
	<td>
		<select name=optiondispyn>
		<option value="Y" selected >Y
		<option value="N" >N
		</select>
	</td>
	<td><input type=text name="disporder" value="" size="3" maxlength="3"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="5" align=center><input type="button" value="����" onClick="AddOptMid(frm);"></td>
</tr>
</table>
</form>
<% else %>
<form name=frm method=post action="" style="margin:0px;">
<input type="hidden" name="mode" value="addlarge">
<table border=0 width=500 cellpadding="5" cellspacing="1" bgcolor="#CCCCCC" class="a">
<tr align=center>
	<td width=50>�ڵ�01</td>
	<td width=90>�ڵ��</td>
	<td width=60>���</td>
	<td width=60>����</td>
</tr>
<tr align=center bgcolor="#FFFFFF">
	<td align=center><input type="text" name="t_cdl" value="" size="2" maxlength="2"></td>
	<td><input type=text name="codename" value="" size="8" maxlength=16></td>
	<td>
		<select name=optiondispyn>
		<option value="Y" selected >Y
		<option value="N" >N
		</select>
	</td>
	<td><input type=text name="disporder" value="" size="3" maxlength="3"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="4" align=center><input type="button" value="����" onClick="addOptLarge(frm);"></td>
</tr>
</table>
</form>
<% end if %>

<%
set ooption = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->