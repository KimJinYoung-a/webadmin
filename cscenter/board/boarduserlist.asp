<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs����
' History : 2009.04.17 �̻� ����
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_boardusercls.asp" -->
<%
dim userid, i

dim occscenterboarduser
set occscenterboarduser = new CCSCenterBoardUser
	occscenterboarduser.FPageSize = 50
	occscenterboarduser.FCurrPage = 1
	occscenterboarduser.GetCSCenterBoardUserList

'// ������ �̻�:2 �� �ý�����:7
if Not ((session("ssAdminLsn") <= 2) or (session("ssAdminPsn") = 7) or C_CSPowerUser or C_ADMIN_AUTH or C_CSpermanentUser) then
	response.write "<br><br>������ �����ϴ�."
	response.end
end if

dim IsSystemPsn	: IsSystemPsn = False
if (session("ssAdminPsn") = 7) then
	IsSystemPsn = True
end if

%>
<script type="text/javascript">

function ModifyIppbxInfo(frm){
	if ((frm.userid.value == "") && (frm.useyn.value == "Y")) {
		alert("���̵� �����ϼ���.\n\n�Ǵ� ���������� �����ϼ���.");
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function doVVipReOrganize(frm){
	if (confirm("VVIP��� ����� �����й� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "reorganizevvipone2one";
		frm.submit();
	}
}

function doVipReOrganize(frm){
	if (confirm("VIP��� ����� �����й� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "reorganizevipone2one";
		frm.submit();
	}
}

function doVipReOrganizeNoCharge(frm){
	if (confirm("�̺й� VIP��� ����� �����й� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "reorganizevipone2onenocharge";
		frm.submit();
	}
}

function doReOrganize(frm){
	if (confirm("�Ϲݻ�� ����� �����й� �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function doReOrganizeNoCharge(frm) {
	if (confirm("�̺й� �Ϲݻ�� ����� �����й� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "reorganizeone2onenocharge";
		frm.submit();
	}
}

function doReOrganizeMichulgoNoCharge(frm){
	if (confirm("�̺й� �귣�� ����� �����й� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "reorganizemichulgonocharge";
		frm.submit();
	}
}

function doReOrganizeNotReturnNoCharge(frm){
	if (confirm("�̺й� �귣�� ����� �����й� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "reorganizenotreturnnocharge";
		frm.submit();
	}
}

function doReOrganizeMichulgoAll(frm){
	if (confirm("��ü ����� �����й� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "reorganizemichulgoall";
		frm.submit();
	}
}

function doReOrganizeNotReturnAll(frm){
	if (confirm("��ü ����� �����й� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "reorganizenotreturnall";
		frm.submit();
	}
}

function doReOrganizeMichulgoAvgAll(frm){
	if (confirm("����� �귣�� �����ֱ� �����й� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "reorganizemichulgoavg";
		frm.submit();
	}
}

function doReOrganizeStockout(frm){
	if (confirm("ǰ�����Ұ� ��ü ��й� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "reorganizestockoutall";
		frm.submit();
	}
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<!--���̵� : <input type="text" class="text" name="userid" value="<%= userid %>">-->
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
      	<input type="button" class="button_s" value="�˻�" onclick="document.frm.submit()">
	</td>
</tr>
</form>
</table>

<br>

* �ް����´� �ް��޷�(�ް� <font color=red>��û���� ����</font>)���� �����ɴϴ�.<br>
* �ް����°� �ް��� �̰ų�, ������ �����ϸ� ����� �ڵ��й迡�� ���ܵ˴ϴ�.<br>
* 1:1 ��� ����� �й�� <font color=red>�� �� �ۼ���</font> �ڵ��й�˴ϴ�.<br><br>

* <font color=red>����� �����й�</font>�� �Ͽ� ����� �й� �� �� �ֽ��ϴ�.

<br>

<b>* 1:1 ��� �Խ���</b>
<input type=button class=button value="VVIP ��� ��ü ��й�" onClick="doVVipReOrganize(frmAction)">
&nbsp;
<input type=button class=button value="VIP ��� ��ü ��й�" onClick="doVipReOrganize(frmAction)">
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
<input type=button class=button value="VIP ��� ������ ��й�" onClick="doVipReOrganizeNoCharge(frmAction)">
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
<input type=button class=button value="�Ϲݻ�� ������ ��й�" onClick="doReOrganizeNoCharge(frmAction)">
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
<input type=button class=button value="�Ϲݻ�� ��ü ��й�" onClick="doReOrganize(frmAction)">
<% if IsSystemPsn then %>
	<input type=button class=button value="��ü ��й�[�ý�����]" onClick="doReOrganize(frmAction)">
<% end if %>
<br><br>

<b>* D+3 �̹߼۰�(����)</b> <input type=button class=button value="�������� �й�" onClick="doReOrganizeMichulgoNoCharge(frmAction)">
<input type=button class=button value="��ü ��й�" onClick="doReOrganizeMichulgoAll(frmAction)">
<input type=button class=button value="������ �귣�� �����ֱ�" onClick="doReOrganizeMichulgoAvgAll(frmAction)">

<br><br>

<b>* ǰ����ҿ�û��(����������, �� �ȳ�����)</b>
<input type=button class=button value="��ü ��й�<% if (C_CSPowerUser or C_ADMIN_AUTH) then %>[�����ڱ���]<% end if %>" onClick="doReOrganizeStockout(frmAction)" <% if Not(C_CSPowerUser or C_ADMIN_AUTH) then %>disabled<% end if %> >

<br><br>

<b>* D+3, D+7 ��ǰ ��ó����(����)</b> <input type=button class=button value="�������� �й�" onClick="doReOrganizeNotReturnNoCharge(frmAction)">
<input type=button class=button value="��ü ��й�" onClick="doReOrganizeNotReturnAll(frmAction)">

<br>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmAction" method="post" action="boarduser_process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="reorganizeone2one">
</form>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width="60">����</td>
    <td width="150">���̵�</td>
    <td width="80">�ް�����</td>
    <td width="90"><font color="red">VVIP</font> 1:1���</td>
    <td width="90">VIP 1:1���</td>
	<td width="90">1:1���</td>
    <td width="90">�����</td>
    <td width="90">ǰ�����</td>
    <td width="90">��ǰ</td>
    <td width="90">���</td>
    <td width="150">������</td>
    <td>���</td>
</tr>
<% if occscenterboarduser.FTotalCount > 0 then %>
	<% for i = 0 to (occscenterboarduser.FResultCount - 1) %>

	<% if (occscenterboarduser.FItemList(i).Fuseyn = "N") then %>
		<tr align="center" bgcolor="#DDDDDD" height="25">
	<% else %>
		<tr align="center" bgcolor="#FFFFFF" height="25">
	<% end if %>

		<form name="frm<%= i %>" method="post" action="/cscenter/board/boarduser_process.asp">
		<input type="hidden" class="text" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="mode" value="modify">
		<input type="hidden" class="text" name="indexno" value="<%= occscenterboarduser.FItemList(i).Findexno %>">
	    <td><%= occscenterboarduser.FItemList(i).Findexno %></td>
	    <td><input type="text" class="text" name="userid" value="<%= occscenterboarduser.FItemList(i).Fuserid %>" size="16"></td>
	    <td>
	    	<% if (occscenterboarduser.FItemList(i).Fuserid <> "") then %>
	        	<% if (occscenterboarduser.FItemList(i).Fvacationyn = "Y") then %>
	        		<font color=red>�ް���</font>
	        	<% else %>
	        		�ٹ���
	        	<% end if %>
	        <% end if %>
	    </td>
	    <td bgcolor="#ABF200">
			<select name="vvipone2oneyn" class="select">
				<option value="Y" <% if (occscenterboarduser.FItemList(i).fvvipone2oneyn = "Y") then %>selected<% end if %>>�й���
				<option value="T" <% if (occscenterboarduser.FItemList(i).fvvipone2oneyn = "T") then %>selected<% end if %>>�й�����
				<option value="N" <% if (occscenterboarduser.FItemList(i).fvvipone2oneyn = "N") then %>selected<% end if %>>�й����
			</select>
	    </td>
	    <td>
			<select name="vipone2oneyn" class="select">
				<option value="Y" <% if (occscenterboarduser.FItemList(i).Fvipone2oneyn = "Y") then %>selected<% end if %>>�й���
				<option value="T" <% if (occscenterboarduser.FItemList(i).Fvipone2oneyn = "T") then %>selected<% end if %>>�й�����
				<option value="N" <% if (occscenterboarduser.FItemList(i).Fvipone2oneyn = "N") then %>selected<% end if %>>�й����
			</select>
	    </td>
	    <td>
			<select name="one2oneyn" class="select">
				<option value="Y" <% if (occscenterboarduser.FItemList(i).Fone2oneyn = "Y") then %>selected<% end if %>>�й���
				<option value="T" <% if (occscenterboarduser.FItemList(i).Fone2oneyn = "T") then %>selected<% end if %>>�й�����
				<option value="N" <% if (occscenterboarduser.FItemList(i).Fone2oneyn = "N") then %>selected<% end if %>>�й����
			</select>
	    </td>
	    <td>
			<select name="michulgoyn" class="select">
				<option value="Y" <% if (occscenterboarduser.FItemList(i).Fmichulgoyn = "Y") then %>selected<% end if %>>�й���
				<option value="T" <% if (occscenterboarduser.FItemList(i).Fmichulgoyn = "T") then %>selected<% end if %>>�й�����
				<option value="N" <% if (occscenterboarduser.FItemList(i).Fmichulgoyn = "N") then %>selected<% end if %>>�й����
			</select>
	    </td>
	    <td>
			<select name="stockoutyn" class="select">
				<option value="Y" <% if (occscenterboarduser.FItemList(i).Fstockoutyn = "Y") then %>selected<% end if %>>�й���
				<option value="T" <% if (occscenterboarduser.FItemList(i).Fstockoutyn = "T") then %>selected<% end if %>>�й�����
				<option value="N" <% if (occscenterboarduser.FItemList(i).Fstockoutyn = "N") then %>selected<% end if %>>�й����
			</select>
	    </td>
	    <td>
			<select name="returnyn" class="select">
				<option value="Y" <% if (occscenterboarduser.FItemList(i).Freturnyn = "Y") then %>selected<% end if %>>�й���
				<option value="T" <% if (occscenterboarduser.FItemList(i).Freturnyn = "T") then %>selected<% end if %>>�й�����
				<option value="N" <% if (occscenterboarduser.FItemList(i).Freturnyn = "N") then %>selected<% end if %>>�й����
			</select>
	    </td>
	    <td>
			<select name="useyn" class="select">
				<option value="Y" <% if (occscenterboarduser.FItemList(i).Fuseyn = "Y") then %>selected<% end if %>>�����
				<option value="N" <% if (occscenterboarduser.FItemList(i).Fuseyn = "N") then %>selected<% end if %>>������
			</select>
	    </td>
	    <td><%= occscenterboarduser.FItemList(i).Flastupdate %></td>
	    <td>
	    	<input type="button" class="button" value="����" onClick="ModifyIppbxInfo(frm<%= i %>)">
	    </td>
	    </form>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF" align="center">
	    <td height="25" colspan="13">�˻������ �����ϴ�.</td>
	</tr>
<% end if %>

</table>

<%
set occscenterboarduser = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
