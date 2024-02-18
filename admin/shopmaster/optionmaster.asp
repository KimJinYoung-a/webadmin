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
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/optionmanagecls.asp"-->
<%
dim cdl, cdm
cdl = request("cdl")
cdm = request("cdm")

dim onlyusing, ordertype, research
onlyusing = request("onlyusing")
ordertype = request("ordertype")
research = request("research")

if (onlyusing="") and (research="") then onlyusing="on"
if ordertype="" then ordertype="c"

dim ooption
set ooption = new COptionManager
ooption.FRectOnlyUsing = onlyusing
ooption.FRectOrderType= ordertype
ooption.GetOption01List


dim subooption
set subooption = new COptionManager
subooption.FRectOnlyUsing = onlyusing
subooption.FRectOrderType= ordertype
if cdl<>"" then
	subooption.GetOption02List cdl
end if
dim i
%>
<script type='text/javascript'>

function AddCode(cdl,cdm){
	var popwin = window.open('editoptioncode.asp?pmode=add&cdl=' + cdl + '&cdm=' + cdm,'editoptioncode','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function EditCode(cdl,cdm){
	var popwin = window.open('editoptioncode.asp?cdl=' + cdl + '&cdm=' + cdm,'editoptioncode','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function DelCode(cdl,cdm){
	alert('���� ���� �Ұ��� �մϴ�.');
	return;

	if (confirm('�ɼ� �ڵ带 ���� �Ͻðڽ��ϱ�?')){
		frmdel.cdl.value = cdl;
		frmdel.cdm.value = cdm;
		frmdel.submit();
	}
}
</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<input type="radio" name="ordertype" value="c" <% if ordertype="c" then response.write "checked" %> >�ڵ�
			<input type="radio" name="ordertype" value="d" <% if ordertype="d" then response.write "checked" %> >����					
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type=checkbox name="onlyusing" <% if onlyusing="on" then response.write "checked" %> >����ϴ¿ɼǸ�����			
		</td>
	</tr>
</table>
</form>
<!-- �˻� �� -->
<br>
<form name=frmdel method=post action="" style="margin:0px;">
<input type="hidden" name=mode value="delcode">
<input type="hidden" name=cdl value="">
<input type="hidden" name=cdm value="">
</form>
<table width=700 class="a">
<tr>
	<td><input type=button value="��з��߰�" onclick="AddCode('','');"></td>
	<td></td>
	<% if cdl<>"" then %>
	<td><input type=button value="�ߺз��߰�" onclick="AddCode('<%= cdl %>','');"></td>
	<% else %>
	<td></td>
	<% end if %>
</tr>
<tr>
	<td valign=top>
		<table width=390 class="a" border=0 cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
		<tr>
			<td>�ڵ��</td>
			<td>���</td>
			<td>����</td>
			<td>����</td>
			<td>����</td>
		</tr>
		<% for i=0 to ooption.FResultCount-1 %>
		<tr bgcolor="#FFFFFF">
		<% if cdl=ooption.FItemList(i).Foptioncode01 then %>
			<td>
				<b><a href="?cdl=<%= ooption.FItemList(i).Foptioncode01 %>&onlyusing=<%= onlyusing %>&ordertype=<%= ordertype %>&research=<%= research %>">
				[<%= ooption.FItemList(i).Foptioncode01 %>]<%= ReplaceBracket(ooption.FItemList(i).Fcodename) %></a></b>
			</td>
		<% else %>
			<td>
				<a href="?cdl=<%= ooption.FItemList(i).Foptioncode01 %>&onlyusing=<%= onlyusing %>&ordertype=<%= ordertype %>&research=<%= research %>">
				[<%= ooption.FItemList(i).Foptioncode01 %>]<%= ReplaceBracket(ooption.FItemList(i).Fcodename) %></a>
			</td>
		<% end if %>
			<td width=40><%= ooption.FItemList(i).Foptiondispyn01 %></td>
			<td width=40><%= ooption.FItemList(i).Fdisporder01 %></td>
			<% if ooption.FItemList(i).Foptioncode01="00" then %>
			<td width=30 align=center>&nbsp;</td>
			<td width=30 align=center>&nbsp;</td>
			<% else %>
			<td width=30 align=center><a href="javascript:EditCode('<%= ooption.FItemList(i).Foptioncode01 %>','');">����</a></td>
			<td width=30 align=center><a href="javascript:DelCode('<%= ooption.FItemList(i).Foptioncode01 %>','');">x</a></td>
			<% end if %>
		</tr>
		<% next %>
		</table>
	</td>
	<td width=20></td>
	<td valign=top>
		<table width=390 class="a" border=0 cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
		<tr>
			<td>�ڵ��</td>
			<td>���</td>
			<td>����</td>
			<td>����</td>
			<td>����</td>
		</tr>
		<% for i=0 to subooption.FResultCount-1 %>
		<tr bgcolor="#FFFFFF">
		<% if cdm=subooption.FItemList(i).Foptioncode02 then %>
			<td>
				<b><a href="?cdl=<%= subooption.FItemList(i).Foptioncode01 %>&cdm=<%= subooption.FItemList(i).Foptioncode02 %>&onlyusing=<%= onlyusing %>&ordertype=<%= ordertype %>&research=<%= research %>">
				[<%= subooption.FItemList(i).Foptioncode02 %>]<%= ReplaceBracket(subooption.FItemList(i).Fcodeview) %></a>
			</td>
		<% else %>
			<td>
				<a href="?cdl=<%= subooption.FItemList(i).Foptioncode01 %>&cdm=<%= subooption.FItemList(i).Foptioncode02 %>&onlyusing=<%= onlyusing %>&ordertype=<%= ordertype %>&research=<%= research %>">
				[<%= subooption.FItemList(i).Foptioncode02 %>]<%= ReplaceBracket(subooption.FItemList(i).Fcodeview) %></a>
			</td>
		<% end if %>
			<td width=40><%= subooption.FItemList(i).Foptiondispyn02 %></td>
			<td width=40><%= subooption.FItemList(i).Fdisporder02 %></td>
			<td width=30 align=center><a href="javascript:EditCode('<%= subooption.FItemList(i).Foptioncode01 %>','<%= subooption.FItemList(i).Foptioncode02 %>');">����</a></td>
			<td width=30 align=center><a href="javascript:DelCode('<%= subooption.FItemList(i).Foptioncode01 %>','<%= subooption.FItemList(i).Foptioncode02 %>');">x</a></td>
		</tr>
		<% next %>
		</table>
	</td>
</tr>
</table>
<%
set ooption = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->