<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ֹ� Ŭ����
' Hieditor : 2009.04.17 �̻� ����
'			 2016.07.19 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/item_qnacls.asp" -->
<%
dim id
	id= requestCheckvar(request("id"),10)

dim itemqna
set itemqna = new CItemQna
	itemqna.FRectID = id
	itemqna.FRectMakerid = session("ssBctID")
	itemqna.getOneItemQna

if itemqna.FResultCount<=0 then
	Call Alert_return("�����Ǿ��ų� �߸��� ���ǹ�ȣ�Դϴ�.")
	Response.End
end if

if IsNull(itemqna.FOneItem.FContents) then
	itemqna.FOneItem.FContents = ""
end if

if (itemqna.FOneItem.FContents = "") then
	itemqna.FOneItem.FContents = "(�������)"
end if

%>
<script type="text/javascript">

function ActReply(frm){
	var userid, username;
	userid = "<%= Replace(itemqna.FOneItem.Fuserid, Chr(34), "") %>";
	username = "<%= Replace(itemqna.FOneItem.Fusername, Chr(34), "") %>";

	if(frm.replycontents.value.length < 1){
		alert("�亯 ������ �����ּ��� �մϴ�.");
		frm.replycontents.focus();
		return;
	}

	if (userid.length>1){
		if (frm.replycontents.value.indexOf(userid) >= 0) {
			alert("�ԷºҰ�!!\n\n�� ���̵� �亯���뿡 �Է��� �� �����ϴ�.");
			return;
		}
	}
	if (username.length>1){
		if (frm.replycontents.value.indexOf(username) >= 0) {
			alert("�ԷºҰ�!!\n\n�� �̸��� �亯���뿡 �Է��� �� �����ϴ�.");
			return;
		}
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
    	<b>�����</b> : <%= getUserLevelStrByDate(itemqna.FOneItem.fUserLevel, left(itemqna.FOneItem.Fregdate,10)) %>
    </td>
    <td align="right">
    	<b>�ۼ���</b> : <%= itemqna.FOneItem.Fregdate %>
    </td>
    <td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="650" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=post action="newitemqna_process.asp">
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
				<td rowspan="3" width="105"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= itemqna.FOneItem.FItemID %>" target="_blank"><img src="<%= itemqna.FOneItem.Flistimage %>" border="0"></a>
				<td>
					��ǰ�ڵ� : <a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= itemqna.FOneItem.FItemID %>" target="_blank"><%= itemqna.FOneItem.FItemID %></a><br>
					��ǰ�� : <a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= itemqna.FOneItem.FItemID %>" target="_blank"><%= itemqna.FOneItem.FItemName %></a><br>
					�귣�� : <%= itemqna.FOneItem.FMakerid %>(<%= itemqna.FOneItem.FbrandName %>)<br>
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
		<%= nl2br(Replace(itemqna.FOneItem.FContents, "<", "&lt;")) %>
	</td>
</tr>
<% if (FALSE) then %>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>�̸��ϼ��ſ�û</b></td>
	<td>
		<input type="radio" name="emailok" value="Y" checked>�߼� <input type="radio" name="emailok" value="N">�̹߼�
	</td>
</tr>
<% end if %>
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
	<td colspan="2">&nbsp;<img src="/images/icon_arrow_down.gif" align="absbottom">&nbsp;<b><font color="red">�亯�ۼ�</font></b></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align=center><b>�亯����</b></td>
	<td>
		<textarea name="replycontents" cols="80" rows="8" class="input_01"><%= (itemqna.FOneItem.FReplyContents) %></textarea>
		<br><br>
		* �亯 �ۼ��� <font color=red>���̸�, �����̵� �ԷºҰ�</font>(�������� ������ ����� �ֽ��ϴ�.)
		<br>&nbsp;
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
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
