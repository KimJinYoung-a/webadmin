<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �˾� ����
' History : 2011.01.28 ������ ����
'			2018.08.10 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/partpersonCls.asp"-->
<%

dim board, isCsCenter

board = requestCheckvar(request("board"),1)

%>

<script language="javascript">

<% if (board = "U") then %>

function workerselect(userid, username)
{
	opener.focus();
	opener.document.frm.workername.value = username;
	opener.document.frm.workerid.value = userid;
	window.close();
}

<% end if %>

</script>
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr height="10" valign="bottom" bgcolor="F4F4F4">
    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_02.gif"></td>
    <td background="/images/tbl_blue_round_02.gif"></td>
    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="20" valign="bottom" bgcolor="F4F4F4">
    <td background="/images/tbl_blue_round_04.gif"></td>
    <td valign="top" bgcolor="F4F4F4" align="center"><b>�ٹ����� ��Ʈ�� ����� ����ó</b></td>
    <td valign="top" align="right" bgcolor="F4F4F4"></td>
    <td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr height="15" valign="bottom" bgcolor="F4F4F4">
    <td background="/images/tbl_blue_round_04.gif"></td>
    <td valign="top" bgcolor="F4F4F4"></td>
    <td valign="top" align="right" bgcolor="F4F4F4"></td>
    <td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" align="center" cellpadding="2" cellspacing="1" bgcolor=#bababa class="a">
<%
Dim clist,clist2, arlist, arlist2, arlist3, i, j, gubun, sabun, idx

'Partlist Ŭ���� ����
Set clist = new Partlist
	'Partlist Ŭ������ ������ Y�� ��쿣 ��� ����Ʈ ���̰� �̻��� ����Ʈ���� ����
	clist.FGubun = "Y"
	arlist = clist.fnGetlist

For i = 0 to Ubound(arlist,2)
	clist.idx = arlist(0,i)
%>
	<tr bgcolor="#FFDDDD" height="25">
		<td colspan=4><b><%= i+1 %>.<%= arlist(1,i) %></b></td>
	</tr>
<%
	arlist2 = clist.fnGetmolist2
	If IsArray(arlist2) = "True" Then
		For j = 0 to Ubound(arlist2,2)
            isCsCenter = (arlist2(1,j) = 220)
%>
	<tr bgcolor="#FFFFFF" height="22" <% if (board = "U") then %>style="cursor:pointer" onClick="workerselect('<%= arlist2(9,j)%>', '<%= arlist2(3,j)%>')"<% end if %> >
    	<td><%= arlist2(2,j)%></td>
    	<td>
			<% if isCsCenter then %>
			������
			<% else %>
			<%= arlist2(3,j)%>
			<% end if %>
		</td>
    	<td>
    		<%
    		'/cs�� �б�. ������ȭ�� �ȹް�. ���� ��ȣ�� �޴´ٰ���.
    		if isCsCenter then
    		%>
    			070-4868-1799 (���ֹ����ù���)
    		<% else %>
    			<%= arlist2(4,j)%>&nbsp;#<%= arlist2(5,j)%>
    		<% end if %>
    	</td>
    	<td>
    		<%
    		if isCsCenter then
    		%>
    			<a href="mailto:customer@10x10.co.kr">customer@10x10.co.kr</a>
    		<% else %>
    			<a href="mailto:<%= arlist2(6,j)%>"><%= arlist2(6,j)%></a>
    		<% end if %>
		</td>
    </tr>
		<% Next %>
	<% elseif (arlist(1,i) = "���ູ����") then %>
	<tr bgcolor="#FFFFFF" height="22">
    	<td>������</td>
    	<td>������</td>
    	<td>070-4868-1799</td>
    	<td></td>
    </tr>
	<%End If%>
<% Next %>
<% Set clist = nothing %>
</table>
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr valign="top" bgcolor="F4F4F4" height="30">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left" bgcolor="F4F4F4">
    	<b>* ����ó</b> <br>&nbsp;&nbsp; ���� : 02-554-2033 &nbsp;&nbsp; �������� : 1644-1851</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" bgcolor="F4F4F4" height="30">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left" bgcolor="F4F4F4">
    	<b>* �ѽ���ȣ</b> <br>&nbsp;&nbsp; ���з� : 02-2179-9244 (MD��Ʈ), 02-2179-9245(������), 02-2179-9058(��������)
    	<br>&nbsp;&nbsp; �������� : 02-3493-1032</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" bgcolor="F4F4F4" height="30">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left" bgcolor="F4F4F4" >
    	<b>* �ּ�</b> <br>
    	&nbsp;&nbsp; ���з� : (03082) ����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ�����<br>

		&nbsp;&nbsp; �������� : ��⵵ ��õ�� ������ ����������2�� 83 �ٹ����� ��������
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" bgcolor="F4F4F4" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
