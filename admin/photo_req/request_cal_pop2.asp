<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���̳��� ����
' History : 2009.04.07 ������ ����
'			2010.12.27 �ѿ�� ����
'           2012.01.10 ������ ����; ���̳��� ����/�߰�
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/photo_req/shedulecls.asp"-->
<%
Dim rno, sno, query1, sche
rno = request("rno")
sno = request("sno")
%>
<script>
function goList(){
	opener.location.href= 'request_modi.asp?req_no=<%=rno%>&udate=A';
	window.close();
}
function goModify(){
	document.getElementById('G').style.display= "none";
	document.getElementById('S').style.display= "block";
}
function go_submit(){
	if(document.frm.req_status.value == "0"){
		alert('������¸� �����ϼ���');
		document.frm.req_status.focus();
		return;
	}
	document.frm.submit();
}
</script>
<%
	query1 = " select start_date, end_date, status from db_partner.dbo.tbl_photo_schedule "
	query1 = query1 + " where schedule_no= '"&sno&"'"
	rsget.Open query1,dbget,1
	IF not rsget.EOF THEN
		sche = rsget.getRows()
	End IF
	rsget.Close
%>
<table id="G" width="400" border="1" cellpadding="0" cellspacing="0" class="a" bordercolordark="White" bordercolorlight="black" align="center"  style="display:block;">
<tr align="center">
	<td height="50" bgcolor="#DDDDFF">�Կ���û ����Ʈ�� ����</td>
	<td height="50" bgcolor="#DDDDFF">�ش� ������ ���� ����</td>
</tr>
<tr align="center">
	<td height="50"><label><input type="radio" name="G" onclick="goList();">GoGo</label></td>
	<td height="50"><label><input type="radio" name="G" onclick="goModify();">GoGo</label></td>
</tr>
</table>

<table id="S" width="400" border="1" cellpadding="0" cellspacing="0" class="a" bordercolordark="White" bordercolorlight="black" align="center" style="display:none;">
<form name="frm" action="request_cal_proc2.asp">
<input type="hidden" name="rno" value="<%=rno%>">
<input type="hidden" name="sno" value="<%=sno%>">
<tr>
	<td height="30" width="15%" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<select name="req_status" class="select">
			<option value="0">--������¼���--</option>
			<option value="4" <%= chkIIF(sche(2,0)="4","selected","") %>>�߰����� ��û</option>
			<option value="1" <%= chkIIF(sche(2,0)="1","selected","") %>>�Կ������� ����</option>
			<option value="2" <%= chkIIF(sche(2,0)="2","selected","") %>>�Կ���</option>
			<option value="3" <%= chkIIF(sche(2,0)="3","selected","") %>>�Կ��Ϸ�</option>
		</select>
	</td>
</tr>
<tr>
	<td height="30" width="15%" bgcolor="#DDDDFF">�ð�</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<%=sche(0,0)%> ~ <%=sche(1,0)%>
	</td>
</tr>
<tr align="center">
	<td height="30" colspan="4">
		<input type="button" value="Ȯ��" onclick="go_submit();" class="button">
		<input type="button" value="�ݱ�" onclick="window.close();" class="button">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->