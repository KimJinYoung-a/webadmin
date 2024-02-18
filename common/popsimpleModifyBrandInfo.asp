<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : CS����
' History : ������ ����
'           2021.06.18 �ѿ�� ����(����� �޴���,�̸��� �������� �������ʿ��� �߰�)
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<%
dim ogroup,opartner,i, makerid, groupid, mode
	makerid = requestCheckVar(request("makerid"),32)
	mode = requestCheckVar(request("mode"),32)

set opartner = new CPartnerUser
	opartner.FRectDesignerID = makerid
	opartner.GetOnePartnerNUser

groupid = opartner.FOneItem.FGroupid

set ogroup = new CPartnerGroup
	if opartner.FResultCount>0 then
		ogroup.FRectGroupid = groupid
		ogroup.GetOneGroupInfo
	end if

dim OReturnAddr
set OReturnAddr = new CCSReturnAddress
	OReturnAddr.FRectMakerid = makerid
	OReturnAddr.GetBrandReturnAddress

dim OCSBrandMemo
set OCSBrandMemo = new CCSBrandMemo
	OCSBrandMemo.FRectMakerid = makerid
	OCSBrandMemo.GetBrandMemo

dim insertOrUpdate
if (OCSBrandMemo.Fbrandid = "") then
	insertOrUpdate = "ins"
else
	insertOrUpdate = "mod"
end if

%>
<script type="text/javascript" src="/cscenter/js/cscenter.js"></script>
<script type='text/javascript'>

function CopyZip(flag,post1,post2,add,dong){
	var frm = eval(flag);

	frm.returnZipcode.value= post1 + "-" + post2;
	frm.returnZipaddr.value= add;
	frm.returnEtcaddr.value= dong;
}

function popZip(flag){
	var popwin = window.open("/lib/searchzip3.asp?target=" + flag,"searchzip3","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function jsSubmitForm(frm) {
	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.submit();
	}
}

</script>

<form name="frmAct" method="post" action="/common/popsimpleModifyBrandInfo_process.asp" style="margin:0px;">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="submode" value="<%= insertOrUpdate %>">
<input type="hidden" name="groupid" value="<%= groupid %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="4">
		<b>�귣�� ����</b>
	</td>
</tr>
<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">[ �귣�� �⺻���� ]</td>
</tr>
<tr height="25">
	<td width="18%" bgcolor="<%= adminColor("tabletop") %>" >�귣��ID</td>
	<td width="30%" bgcolor="#FFFFFF"><b><%= opartner.FOneItem.FID %></b></td>
	<td width="18%" bgcolor="<%= adminColor("tabletop") %>">��Ʈ��Ʈ��</td>
	<td bgcolor="#FFFFFF"><b><%= opartner.FOneItem.Fsocname_kor %></b></td>
</tr>
<tr height="5">
	<td colspan="4" bgcolor="#FFFFFF"></td>
</tr>

<% if (mode = "modifyReturnCharge") then %>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�����</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" size="15" name="returnName" value="<%= OReturnAddr.FreturnName %>">
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��ȭ</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" size="15" name="returnPhone" value="<%= OReturnAddr.FreturnPhone %>">
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ���</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" size="15" name="returnhp" value="<%= OReturnAddr.Freturnhp %>">
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�̸���</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" size="20" name="returnEmail" value="<%= OReturnAddr.FreturnEmail %>">
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ �ּ�</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" class="text" name="returnZipcode" value="<%= OReturnAddr.FreturnZipcode %>" size="7" maxlength="7">
			<input type="button" class="button_s" value="�˻�" onClick="FnFindZipNew('frmAct','J')">
			<input type="button" class="button_s" value="�˻�(��)" onClick="TnFindZipNew('frmAct','J')"><br>
			<% '<input type="button" class="button" value="�˻�(��)" onclick="javascript:popZip('frmAct');"><br> %>
			<input type="text" class="text" name="returnZipaddr" value="<%= OReturnAddr.FreturnZipaddr %>" size="25" maxlength="64">
			<input type="text" class="text" name="returnEtcaddr" value="<%= OReturnAddr.FreturnEtcaddr %>" size="40" maxlength="128">
		</td>
	</tr>
<% elseif (mode = "modifyCSCharge") then %>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">CS�����</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" size="15" name="csName" value="<%= OCSBrandMemo.FcsName %>">
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">CS��ȭ</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" size="15" name="csPhone" value="<%= OCSBrandMemo.FcsPhone %>">
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">CS�ڵ���</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" size="15" name="cshp" value="<%= OCSBrandMemo.Fcshp %>">
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">CS�̸���</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" size="20" name="csEmail" value="<%= OCSBrandMemo.FcsEmail %>">
		</td>
	</tr>
<% end if %>

<tr align="center">
	<td colspan="4" bgcolor="#FFFFFF" height="30">
		<input type="button" class="button" value="�����ϱ�" onClick="jsSubmitForm(frmAct)">
		&nbsp;
		<input type="button" class="button" value="�ݱ�" onclick="self.close();">
	</td>
</tr>
</table>
</form>

<%
set opartner = Nothing
set ogroup = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
