<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/upche_qnacls.asp" -->
<%

dim i, j, page, gubun

page = request("page")
gubun = request("gubun")
if page="" then page=1


'==============================================================================
'���� 1:1�����亯
dim boardqna
set boardqna = New CUpcheQnA

boardqna.FPageSize = 20
boardqna.FCurrPage = page
boardqna.FScrollCount = 10
boardqna.FRectGubun = gubun

boardqna.list
%>
<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 9pt; }
    INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #CACACA; color: #000000; }
-->
</STYLE>
<script language='javascript'>
function  TnSearch(frm){
	if (frm.rectuserid.length<1){
		alert('�˻�� �Է��ϼ���.');
		return;
	}
	frm.method="get";
	frm.submit();
}
function NextPage(ipage){
	document.frmSrc.page.value= ipage;
	document.frmSrc.submit();
}
</script>
<table width="720" border="0">
<form name="frmSrc" method="get" action="">
<input type="hidden" name="page" value="<% = page %>">
<tr>
  <td>
	&nbsp;&nbsp;������������ :
		  <select name="gubun">
			<option value="">����</option>
			<option value="01" <% if gubun="01" then response.write "selected" %> >��۹���</option>
			<option value="02" <% if gubun="02" then response.write "selected" %> >��ǰ����</option>
			<option value="03" <% if gubun="03" then response.write "selected" %> >��ȯ����</option>
			<option value="04" <% if gubun="04" then response.write "selected" %> >���깮��</option>
			<option value="05" <% if gubun="05" then response.write "selected" %> >�԰���</option>
			<option value="06" <% if gubun="06" then response.write "selected" %> >�����</option>
			<option value="07" <% if gubun="07" then response.write "selected" %> >��ǰ��Ϲ���</option>
			<option value="08" <% if gubun="08" then response.write "selected" %> >�̺�Ʈ����</option>
			<option value="20" <% if gubun="20" then response.write "selected" %> >��Ÿ����</option>
		  </select>&nbsp;<input type="submit" value="�˻�">
  </td>
  <td align="right"><a href="upche_qna_board_list.asp">��ó������Ʈ</a></td>
</tr>
</form>
</table>

<table width="720" border="1" bordercolordark="White" bordercolorlight="black" cellpadding="0" cellspacing="0">
  <tr bgcolor="#DDDDFF" height="25">
    <td width="150" align="center">��ü��</td>
    <td align="center">����</td>
    <td width="100" align="center">����</td>
    <td width="70" align="center">��ü����</td>
	<td width="70" align="center">ó������</td>
    <td width="100" align="center">�ۼ���</td>
  </tr>
<% for i = 0 to (boardqna.FResultCount - 1) %>
  <tr height="20">
    <td align="center"><%= boardqna.FItemList(i).Fusername %>(<%= boardqna.FItemList(i).Fuserid %>)</td>
    <td>&nbsp;<a href="upche_qna_board_reply.asp?idx=<%= boardqna.FItemList(i).Fidx %>"><%= boardqna.FItemList(i).Ftitle %></a></td>
    <td align="center"><%= boardqna.FItemList(i).GubunName %></td>
    <td align="center"><%= boardqna.FItemList(i).UpcheGubun %></td>
    <% if (boardqna.FItemList(i).Freplyn="") then %>
    <td>&nbsp;</td>
    <% else %>
    <td align="center">�Ϸ�</td>
    <% end if %>
    <td align="center"><%= FormatDate(boardqna.FItemList(i).Fregdate, "0000.00.00") %></td>
  </tr>
<% next %>
<tr>
	<td colspan="6" align="center" height="30">
		<% if boardqna.HasPreScroll then %>
			<a href="javascript:NextPage('<%= boardqna.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + boardqna.StartScrollPage to boardqna.FScrollCount + boardqna.StartScrollPage - 1 %>
			<% if i>boardqna.FtotalPage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if boardqna.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>
<br><br>
<!-- #include virtual="/lib/db/dbclose.asp" -->