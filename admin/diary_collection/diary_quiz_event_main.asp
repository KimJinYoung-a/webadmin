<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db2open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
		dim sql,strHtml
		sql =	" select top 20 eventidx,eventname from [db_cts].[dbo].[tbl_2007_diary_event_master]" &_
					" order by eventidx desc"
		db2_rsget.open sql, db2_dbget, 1

		if not db2_rsget.eof then




				do until db2_rsget.eof
					strHtml = strHtml + "<tr bgcolor='#FFFFFF'>"
					strHtml = strHtml + "	<td align='center'>" & db2_rsget("eventidx") & "</td> "
					strHtml = strHtml + "	<td align='center'>" &_
															" <span onclick=" & chr(34) & "TnEditEvent('" & db2_rsget("eventidx") & "')" & chr(34) & " style='cursor:pointer'>" & db2_rsget("eventName") & "</span>" &_
															"	<span onclick=" & chr(34) & "TnDelEvent('" & db2_rsget("eventidx") & "')" & chr(34) & " style='cursor:pointer'>[����]</span></td> "
					strHtml = strHtml + "	<td align='center'><span onclick=" & chr(34) & "TnGoentrant('" & db2_rsget("eventidx") & "')" & chr(34) & " style='cursor:pointer'>����</span></td>"
					'strHtml = strHtml + "	<td align='center'><span onclick=" & chr(34) & "TnGoWinnerList('" & db2_rsget("eventidx") & "')" & chr(34) & " style='cursor:pointer'>����</span></td> "
					strHtml = strHtml + "</tr>"

				db2_rsget.movenext
				loop

		end if

		db2_rsget.close
%>

<script language="javascript" type="text/javascript">
//������ ����
function TnGoentrant(idx){
	document.location.href="diary_quiz_event_winnerList.asp?eventidx=" + idx;

}
//�ű� ���
function popNewReg(){
	window.open('diary_quiz_event_Pop_new.asp?mode=add','newpop','width=500,height=500,scrollbars=yes,resizable=yes');
}
//�̺�Ʈ ����
function TnEditEvent(idx){
	window.open('diary_quiz_event_Pop_new.asp?mode=edit&eventidx=' + idx,'editpop','width=500,height=500,scrollbars=yes,resizable=yes');
}
//�̺�Ʈ ����
function TnDelEvent(idx){
	var ret = confirm('���� �Ͻðڽ��ϱ�.?');
	if (ret) {
		document.delFrm.mode.value='del';
		document.delFrm.eventidx.value= idx ;
		document.delFrm.submit();
	}
}

</script>
<table width="600" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC" class="a">
	<tr>
		<td colspan="3" align="right" bgcolor="#EDEDED"><span style="cursor:pointer" onclick="popNewReg()">�űԵ��</span>&nbsp;&nbsp;</td>
	</tr>
	<tr bgcolor="#EDEDED">
		<td width="70" align="center">��ȣ</td>
		<td align="center">����</td>
		<td width="100" align="center">������ ����</td>
		<!--<td width="100" align="center">��÷�� ����</td>  -->
	</tr>

	<%= strHtml %>

</table>

<form name="delFrm" method="post" action="http://imgstatic.10x10.co.kr/linkweb/doDiary_quiz_event.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="del">
<input type="hidden" name="eventidx" value="">
</form>
<table width="600" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC" class="a">
	<tr>
		<td><a href="/admin/sitemaster/diary_collection_2007/diary_eval_event_main.asp">���̾ ��ǰ�ı��̺�Ʈ ��÷�� ���</a></td>
	</tr>
	<tr>
		<td><a href="/admin/sitemaster/diary_collection_2007/diary_eval_event_winList.asp">���̾ ��ǰ�ı��̺�Ʈ ��÷�� ����</a></td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db2close.asp" -->