<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/classes/board/upche_qnacls.asp" -->
<%

dim i, j, page

page = requestCheckVar(request("page"),10)
if page="" then page=1


'==============================================================================
'���� 1:1�����亯
dim boardqna
set boardqna = New CUpcheQnA

boardqna.FPageSize = 20
boardqna.FCurrPage = page
boardqna.FScrollCount = 10
boardqna.FRectUserid = session("ssBctId")
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
function NextPage(ipage){
	document.frmSrc.page.value= ipage;
	document.frmSrc.submit();
}
</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="�۾���" onClick="javascript:location.href='upche_qna_board_write.asp?mode=write&menupos=<%=menupos %>';">
		</td>
		<td align="right">
			
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="200">��ü(�귣��ID)</td>
	    <td>����</td>
	    <td width="100">����</td>
	    <td width="70">�����</td>
	    <td width="100">�亯����</td>
	    <td width="120">�ۼ���</td>
	</tr>
	<% for i = 0 to (boardqna.FResultCount - 1) %>
	<tr align="center" bgcolor="#FFFFFF" >
	    <td><%= boardqna.FItemList(i).Fusername %>(<%= boardqna.FItemList(i).Fuserid %>)</td>
	    <td align="left"><a href="upche_qna_board_reply.asp?idx=<%= boardqna.FItemList(i).Fidx %>&menupos=<%=menupos %>"><%= ReplaceBracket(boardqna.FItemList(i).Ftitle) %></a></td>
	    <td><%= boardqna.FItemList(i).GubunName %></td>
	    <td><%= boardqna.FItemList(i).Fworker %></td>
	    <td>
	    	<% if (boardqna.FItemList(i).Freplyn<>"") then %>
	    		�亯�Ϸ�
	    	<% end if %>
	    </td>
	    <td><%= FormatDate(boardqna.FItemList(i).Fregdate, "0000.00.00") %></td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF" >
		<td colspan="5" align="center" height="30">
				<% if boardqna.HasPreScroll then %>
				<a href="?page=<%= CStr(boardqna.StartScrollPage - 1) %>">��</a>
				<% else %>
				<% end if %>
				<% for i = boardqna.StartScrollPage to (boardqna.StartScrollPage + boardqna.FScrollCount - 1) %>
				<% if (i > boardqna.FTotalPage) then Exit For %>
				<% if CStr(i) = CStr(boardqna.FCurrPage) then %>
				[<font color="red"><%= i %></font>]
				<% else %>
				<a href="?page=<%= i %>" class="verdana-small">[<%= i %>]</a>
				<% end if %>
	
				<% next %>
				<% if boardqna.HasNextScroll then %>
				<a href="?page=<%= CStr(boardqna.StartScrollPage + boardqna.FScrollCount) %>">��</a>
				<% else %>
				<% end if %>
		</td>
		<td align="center"></td>
	</tr>
</table>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->