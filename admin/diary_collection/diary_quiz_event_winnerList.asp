<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db2open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim eventidx
eventidx= request("eventidx")

if eventidx="" then
response.write "<h1>�߸��� �����Դϴ�</h1>"
dbget.close()	:	response.End
end if

		dim SQL,strHtml,FResultCount

		SQL = " select win_idx,userid" &_
					" from [db_cts].[dbo].[tbl_2007_diary_event_winner] " &_
					" where  eventidx=" & eventidx &_
					" order by userid "


		db2_rsget.open SQL,db2_dbget,1

				if not db2_rsget.eof then
							FResultCount = db2_rsget.recordCount
				end if

				if not db2_rsget.eof then

						do until db2_rsget.eof
							strHtml = strHtml + "<tr bgcolor='#FFFFFF'>"
							strHtml = strHtml + "	<td align='center'>" & db2_rsget("userid") & "</td> "
							strHtml = strHtml + "	<td align='center'><span style='cursor:pointer; color:blue;' onclick=" & chr(34) & "FnDelWinner('" & db2_rsget("win_idx") & "');" & chr(34) &">[del]</span></td> "
							strHtml = strHtml + "</tr>"

						db2_rsget.movenext
						loop

				end if

		db2_rsget.close

%>

<table width="900" border="0" cellpadding="0" cellspacing="0" class="a">
<form name="winnerFrm" method="post" action="/admin/sitemaster/diary_collection_2007/doDiary_quiz_event_winner.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="eventidx" value="<%= eventidx %>">
<input type="hidden" name="win_idx" value="">
	<tr>
		<td bgcolor="#EDEDED">
			<font color="red">notice:</font><br>
			��ܿ� "�����ڸ� ���� �ڽ�"�� üũ�Ͻ��� "�����Է�" ���� ������ �ۼ��Ͻ��� �˻��ϼ���.<br>
			���̵� Ŭ���Ͻø� ���� �ڽ��� ���̵� �ӽ� ����˴ϴ�.<br>
			�������� �ѱ�ŵ� ����Ȱ��� �����˴ϴ�.<br>
			���̵� ������ �����ϴܿ� ���� ��ư�� �����ø� ��÷�ڷ� ���� �˴ϴ�.
			</td>
	</tr>
	<tr>
		<td width="700"><iframe src="/admin/sitemaster/diary_collection_2007/diary_quiz_event_entryList.asp?eventidx=<%= eventidx %>" width="700" height="700" frameborder="0"></iframe></td>
		<td width="200" align="left" valign="top">
			<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td>
						<table width="150" border="0" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" class="a">
							<%= strHtml %>
							<tr>
								<td colspan="2">��÷�ڼ� : �� <font color="red"><%= FresultCount %></font>��</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<textarea name="winnerList" value="" cols="20" rows="20"></textarea>
						<br>
						<input type="button" value="����"  onclick="FnsaveWinner();" />
					</td>
				</tr>
			</table>
		</td>

	</tr>
</form>
</table>

<script language="javascript" type="text/javascript">
function FnsaveWinner(){
	var frm = document.winnerFrm;

	if (frm.winnerList.value.length<1){
		alert('��÷�ڸ� ������ �ּ���	');
		return false;
	}

	var conf = confirm('�����Ͻðڽ��ϱ�?\n' + frm.winnerList.value);

	if (conf) {
		frm.mode.value="write"
		frm.submit();
	}
}

function FnDelWinner(idx){
	var frm = document.winnerFrm;
	var conf = confirm("���� �Ͻðڽ��ϱ�?");

	if (conf){
		frm.mode.value="del"
		frm.win_idx.value=idx
		frm.submit();
	}


}
</script>





<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db2close.asp" -->