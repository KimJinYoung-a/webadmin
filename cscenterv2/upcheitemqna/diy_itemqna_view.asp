<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/upcheitemqna/diy_item_qnacls.asp"-->
<%
dim id
id= RequestCheckvar(request("id"),10)

dim itemqna
set itemqna = new CItemQna
itemqna.FRectID = id
itemqna.FRectMakerid = session("ssBctID")
itemqna.getOneItemQna

%>
<script language='javascript'>
function ActReply(frm){


	if(frm.replycontents.value.length < 1){
		alert("�亯 ������ �����ּ��� �մϴ�.");
		frm.replycontents.focus();
		return;
	}

	if(confirm("��ǰ�� ���� �亯 �Ͻðڽ��ϱ�?")){
		frm.submit();
	}
}
</script>

<!-- ǥ ��ܹ� ����-->
<table width="650" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
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
	        	<b>�ۼ���</b> : <%= itemqna.FOneItem.Fusername %>(<%= itemqna.FOneItem.Fuserid %>)
	        	&nbsp;&nbsp;&nbsp;&nbsp;
	        	<b>�����</b> : <%= itemqna.FOneItem.GetUserLevelStr %>
	        </td>
	        <td align="right">
	        	<b>�ۼ���</b> : <%= itemqna.FOneItem.Fregdate %>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->


<table width="650" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=post action="diy_itemqna_process.asp">
	<input type="hidden" name="id" value="<%= itemqna.FOneItem.Fid %>">
	<% if itemqna.FOneItem.IsReplyOk then %>
	<input type="hidden" name="mode" value="reply">
	<% else %>
	<input type="hidden" name="mode" value="firstreply">
	<% end if %>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center"><b>������<br>��ǰ����</b>
		<td>
			<!-- ��ǰ���� -->
			<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
				<tr valign="top">
					<td rowspan="3" width="105"><img src="<%= itemqna.FOneItem.Flistimage %>">
					<td>
						��ǰ�ڵ� : <%= itemqna.FOneItem.FItemID %><br>
						��ǰ�� : <%= itemqna.FOneItem.FItemName %><br>
						�귣�� : <%= itemqna.FOneItem.FMakerid %><br>
						���� : <%= FormatNumber(itemqna.FOneItem.FSellcash,0) %><br>
					</td>
				</tr>
			</table>
			<!-- ��ǰ���� -->
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>���ǳ���</b></td>
		<td>
			<%= nl2br(itemqna.FOneItem.FContents) %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>�̸��ϼ��ſ�û</b></td>
		<td>
			<input type="radio" name="emailok" value="Y" checked>�߼� <input type="radio" name="emailok" value="N">�̹߼�
		</td>
	</tr>
	<tr height="25" bgcolor="<%= adminColor("topbar") %>">
		<td colspan="2">&nbsp;<img src="/images/icon_arrow_down.gif" align="absbottom">&nbsp;<b><font color="red">�亯�ۼ�</font></b></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align=center><b>�亯����</b></td>
		<td>
			<textarea name="replycontents" cols="80" rows="8" class="input_01"><%= (itemqna.FOneItem.FReplyContents) %></textarea><br>
			<font color=red>�� �亯 �ۼ��� ���̸� ��� �����̵� ����ϼ���.(�������� ������ ����� �ֽ��ϴ�.)</font>
		</td>
	</tr>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="650" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<a href="javascript:ActReply(frm);"><img src="/images/icon_reply.gif" border="0" align="absbottom"></a>
        	<a href="/designer/board/newitemqna_list.asp"><img src="/images/icon_list.gif" border="0" align="absbottom"></a>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
</form>


<%
set itemqna = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->