<%
'###########################################################
' Description : �系��������
' Hieditor : �̻� ����
'			 2022.07.12 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<%

Dim lBoardScmNotice
Set lBoardScmNotice = new board
	lBoardScmNotice.fnGetScmNoticeList

%>
<script type='text/javascript'>

function jsPopModiScmNotice()
{
	var win = window.open("/admin/member/popScmNoticeModi.asp","jsPopModiScmNotice","width=1400,height=768,scrollbars=yes");
	win.focus();
}
</script>
<table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr height="25">
		    <td style="border-bottom:1px solid #BABABA">
		        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�系��������</b>
		    </td>
		    <td align="right" style="border-bottom:1px solid #BABABA">
				<input type="button" class="button" value="�����ϱ�" onClick="jsPopModiScmNotice()" <%= CHKIIF(C_OP Or C_PSMngPart Or C_SYSTEM_Part or C_ADMIN_AUTH, "", "disabled") %>>
		    </td>
		</tr>
		<tr height="25">
		    <td colspan="2">
				<table width="100%" border="0" align="center" cellpadding="1" cellspacing="2" class="a">
				<tr align="left">
					<td bgcolor="#DCDCDC">����</td>
					<td bgcolor="#DCDCDC">����</td>
					<td bgcolor="#DCDCDC">����</td>
				</tr>
			<% for i = 0 to lBoardScmNotice.FResultCount - 1 %>
				<tr align="left">
					<td bgcolor="#EFEFEF"><%= ReplaceBracket(lBoardScmNotice.FbrdList(i).FscheduleDate) %></td>
					<td bgcolor="#EFEFEF"><%= ReplaceBracket(lBoardScmNotice.FbrdList(i).Ftitle) %></td>
					<td bgcolor="#EFEFEF"><%= nl2br(ReplaceBracket(lBoardScmNotice.FbrdList(i).Fcontents)) %></td>
				</tr>
			<% next %>
				</table>
			</td>
		</tr>
        </table>
    </td>
</tr>
</table>
