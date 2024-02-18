<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/giftcard/giftcard_cls.asp"-->
<%
Dim i, idx, mode, oGiftCard
idx = request("idx")
If idx ="" Then
	mode = "I"
Else
	mode = "U"
End If

Set oGiftCard = new cGiftCard
	oGiftCard.FRectIdx = idx
	oGiftCard.getGiftCardOneItem
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function targetOpt(v){
	if (v == "0000"){
		$("#sugiTarget").show();
		$("#sugiPrice").val("");
	}else{
		$("#sugiTarget").hide();
		$("#sugiPrice").val("");
	}
}
function regGift(){
	<% 'if not(C_ADMIN_AUTH or C_PSMngPart) then %>
	<% if not(C_ADMIN_AUTH) then %>
	if ($("#eappidx").val()== ""){
		alert('ǰ�Ǽ�IDX�� �Է��ϼ���');
		$("#eappidx").focus();
		return;
	}
	<% end if %>

	if ($("#reqTitle").val()== ""){
		alert('������ �Է��ϼ���');
		$("#reqTitle").focus();
		return;
	}

	if ($("#reqContent").val()== ""){
		alert('������ �Է��ϼ���');
		$("#reqContent").focus();
		return;
	}


	if ($("#userid").val()== ""){
		alert('�ٹ�����ID�� �Է��ϼ���');
		$("#userid").focus();
		return;
	}

	if ($("#makeCnt").val()== ""){
		alert('�߱��� ī�� ������ �Է��ϼ���');
		$("#makeCnt").focus();
		return;
	}

	if ($("#opt").val()== ""){
		alert('�ɼ��� �����ϼ���');
		$("#opt").focus();
		return;
	}

	if ($("#MMSTitle").val()== ""){
		alert('MMS ������ �Է��ϼ���');
		$("#MMSTitle").focus();
		return;
	}

	if ($("#MMSContent").val()== ""){
		alert('MMS ������ �Է��ϼ���');
		$("#MMSContent").focus();
		return;
	}

	if (confirm("���� �Ͻðڽ��ϱ�?")){
		var frm = document.frm;
		frm.action = "/admin/giftcard/giftcardProc.asp";
		frm.submit();
	}
}
function pop_checkId(){
	if ($("#userid").val()== ""){
		alert('�ٹ�����ID�� �Է��ؾ� Ȯ�� �����մϴ�.');
		$("#userid").focus();
		return;
	} else {
		var str = $("#userid").val();
		str = str.replace(/(?:\r\n|\r|\n)/g, ',');

		var popwin = window.open("/admin/giftcard/pop_checkId.asp?userid="+str,"popcheckId","width=1200,height=600,scrollbars=yes,resizable=yes");
		popwin.focus();
	}
}
function pop_eappView(){
	if ($("#eappidx").val()== ""){
		alert('ǰ�Ǽ�IDX�� �Է��ؾ� Ȯ�� �����մϴ�.');
		$("#eappidx").focus();
		return;
	} else {
		var iridx = $("#eappidx").val();
		var popwin = window.open("/admin/approval/eapp/vieweapp.asp?iridx="+iridx,"popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
		popwin.focus();
	}
}
</script>
<form name="frm" method="post">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="idx" value="<%= idx %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">ǰ�Ǽ�IDX</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" id="eappidx" name="eappidx" class="text" size="10" value="<%= oGiftCard.FOneItem.FEappIdx %>">
		&nbsp;<input type="button" class="button" value="ǰ�Ǽ�����" onclick="pop_eappView();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">����</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" id="reqTitle" name="reqTitle" class="text" size="100" value="<%= oGiftCard.FOneItem.FReqTitle %>">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">����</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<textarea class="textarea" id="reqContent" name="reqContent" cols="150" rows="20"><%= oGiftCard.FOneItem.FReqContent %></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">�ٹ�����ID</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<textarea name="userid" id="userid" class="textarea" cols="32" rows="5"><%= Chkiif(mode="U", getUserids(idx), "") %></textarea>
		<strong>(�� �ݵ�� ���ͷ� �����ؼ� �־��ּ���)</strong>
		<br />
		<input type="button" class="button" value="IDüũ" onclick="pop_checkId();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">�߱��� ī�� ����</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" id="makeCnt" name="makeCnt" class="text" size="10" value="<%= oGiftCard.FOneItem.FMakeCnt %>">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">�ɼ�</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<select id="opt" name="opt" class="select" onchange="targetOpt(this.value);">
			<option value="">-Choice-</option>
			<option value="0001" <%= Chkiif(oGiftCard.FOneItem.FOpt="0001", "selected", "") %> >1������</option>
			<option value="0002" <%= Chkiif(oGiftCard.FOneItem.FOpt="0002", "selected", "") %>>2������</option>
			<option value="0003" <%= Chkiif(oGiftCard.FOneItem.FOpt="0003", "selected", "") %>>3������</option>
			<option value="0004" <%= Chkiif(oGiftCard.FOneItem.FOpt="0004", "selected", "") %>>5������</option>
			<option value="0005" <%= Chkiif(oGiftCard.FOneItem.FOpt="0005", "selected", "") %>>8������</option>
			<option value="0006" <%= Chkiif(oGiftCard.FOneItem.FOpt="0006", "selected", "") %>>10������</option>
			<option value="0007" <%= Chkiif(oGiftCard.FOneItem.FOpt="0007", "selected", "") %>>15������</option>
			<option value="0008" <%= Chkiif(oGiftCard.FOneItem.FOpt="0008", "selected", "") %>>20������</option>
			<option value="0009" <%= Chkiif(oGiftCard.FOneItem.FOpt="0009", "selected", "") %>>30������</option>
			<option value="0000" <%= Chkiif(oGiftCard.FOneItem.FOpt="0000", "selected", "") %>>�����Է�</option>
		</select>
		<span id="sugiTarget" <%= Chkiif(oGiftCard.FOneItem.FSugiPrice = "" , "style='display:none;'", "") %>  >
			<input type="text" id="sugiPrice" name="sugiPrice" value="<%= oGiftCard.FOneItem.FSugiPrice %>">��
		</span>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">MMS ����</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" class="text" size="100" id="MMSTitle" name="MMSTitle" value="<%= oGiftCard.FOneItem.FMMSTitle %>">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">MMS ����</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<textarea class="textarea" id="MMSContent" name="MMSContent" cols="150" rows="20"><%= oGiftCard.FOneItem.FMMSContent %></textarea>
	</td>
</tr>
<tr align="center" height="25" bgcolor="FFFFFF">
	<td colspan="3">
		<% If oGiftCard.FOneItem.FIsSend <> "Y" Then %>
			<input type="button" class="button" value="����" onClick="regGift();">
		<% End If %>
		<input type="button" class="button" value="����Ʈ��" onClick="location.href='/admin/giftcard/list.asp?menupos=<%= menupos %>';">
	</td>
</tr>
</table>
</form>
<% Set oGiftCard = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->