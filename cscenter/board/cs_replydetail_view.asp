<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_replycls.asp"-->
<%

dim idx, currmode
dim masterIdx, gubunCode, masterUseYN

idx = request("idx")
masterIdx = request("masterIdx")
gubunCode = request("gubunCode")
masterUseYN = request("masterUseYN")


dim oCReply
Set oCReply = new CReply

if (idx <> "") then
	currmode = "modiDetail"
	oCReply.FRectDetailIDX = idx
	oCReply.GetReplyDetailOne()
else
	currmode = "insDetail"
	oCReply.GetReplyDetailEmptyOne()

	oCReply.FOneItem.Fmasteridx = masterIdx
	oCReply.FOneItem.FgubunCode = gubunCode
	oCReply.FOneItem.FmasterUseYN = masterUseYN
end if

%>
<script language='javascript'>

function fnSaveReplyMaster() {
	var frm = document.frm;

	if (frm.masterIdx.value == "") {
		alert("�⺻ ī�װ��� �����ϼ���.");
		frm.masterIdx.focus();
		return;
	}

	if (frm.subtitle.value == "") {
		alert("�� ī�װ����� �Է��ϼ���.");
		frm.subtitle.focus();
		return;
	}

	/*
	if (frm.contents.value == "") {
		alert("�ȳ������� �Է��ϼ���.");
		frm.contents.focus();
		return;
	}
	*/

	if (frm.dispOrderNo.value == "") {
		alert("ǥ�ü����� �Է��ϼ���.");
		frm.dispOrderNo.focus();
		return;
	}

	if (frm.dispOrderNo.value*0 != 0) {
		alert("ǥ�ü����� ���ڸ� �����մϴ�.");
		frm.dispOrderNo.focus();
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function fnGotoList() {
	document.location.href = "cs_replydetail_list.asp?menupos=<%= menupos %>";
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
	        	<img src="/images/icon_arrow_down.gif" align="absbottom">
	        	<font color="red"><strong>�� ī�װ� <% if (currmode = "insMaster") then %>�ۼ�<% else %>����<% end if %></strong></font>
	        </td>
	        <td align="right">

	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="cs_reply_process.asp">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="mode" value="<%= currmode %>">
	<input type="hidden" name="detailidx" value="<%= oCReply.FOneItem.Fidx %>">
	<input type="hidden" name="gubunCode" value="<%= oCReply.FOneItem.FgubunCode %>">
	<tr bgcolor="#FFFFFF">
		<td width=100 bgcolor="<%= adminColor("tabletop") %>" height="30">�⺻ ī�װ���</td>
		<td>
			<% Call drawSelectBoxReplyMaster("masterIdx", oCReply.FOneItem.Fmasteridx, oCReply.FOneItem.FgubunCode, oCReply.FOneItem.FmasterUseYN) %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width=100 bgcolor="<%= adminColor("tabletop") %>" height="30">�� ī�װ���</td>
		<td>
			<input type="text" class="text" name="subtitle" value="<%= oCReply.FOneItem.Fsubtitle %>" size="40">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" height="30">�ȳ�����</td>
		<td>
			<textarea class="textarea" name="contents" cols="100" rows="10"><%= oCReply.FOneItem.Fcontents %></textarea>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" height="30">ǥ�ü���</td>
		<td>
			<input type="text" class="text" name="dispOrderNo" value="<%= oCReply.FOneItem.FdispOrderNo %>" size="4">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" height="30">��뱸��</td>
		<td>
			<select class="select" name="useYN">
				<option value="Y" <% if (oCReply.FOneItem.FuseYN = "Y") then %>selected<% end if %> >�����</option>
				<option value="N" <% if (oCReply.FOneItem.FuseYN = "N") then %>selected<% end if %> >������</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>" height="30">�����</td>
		<td>
			<%= oCReply.FOneItem.Freguserid %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td>
			<%= oCReply.FOneItem.Flastupdate %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" height="35" align="center">
			<input type="button" class="button" value="�����ϱ�" onclick="fnSaveReplyMaster()">
			&nbsp;
			<input type="button" class="button" value="�������" onclick="fnGotoList()">
		</td>
	</tr>
	</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
