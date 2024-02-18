<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->
<link rel="stylesheet" href="/css/scm.css" type="text/css">

<%
	Dim vAction, vEvtID, vTitle, oPart, lp
	vAction = RequestCheckvar(Request("action"),16)
	vEvtID = RequestCheckvar(Request("evtId"),10)
	vTitle = RequestCheckvar(Request("title"),64)
	
	If vAction = "insert" OR vAction = "delete" Then
		Call Proc()
	Else
		set oPart = new CPart
		oPart.FRectevtId = vEvtID
		oPart.GetWinnerList
	End If
%>

<script language="JavaScript">
function checkform(f) {
	if (f.winner.value == "")
	{
		alert("��÷�ڸ� �Է��ϼ���!")
		f.winner.focus();
		return false;
	}
	
	var tmp = f.winner.value.replace(/[,]/gi,'\n');

	if(confirm("�Է��Ͻ� ��÷�ڰ� �½��ϱ�?\n\n"+tmp+"") == true) {
		f.action.value = "insert";
		return true;
     } else {
     	return false;
     }
}

function delWinner(tmp) {
	if(confirm("����Ͻ� ��÷�ڰ� �½��ϱ�?\n\n"+tmp+"") == true) {
		frm1.action.value = "delete";
		frm1.winner.value = tmp;
		frm1.submit();
		return true;
     } else {
     	return false;
     }
}
</script>

[<%=vEvtID%>] <%=db2html(vTitle)%> ��÷��
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm1" method="post" onSubmit="return checkform(this);">
<input type="hidden" name="evtId" value="<%=vEvtID%>">
<input type="hidden" name="title" value="<%=vTitle%>">
<input type="hidden" name="action" value="">
<tr>
	<td style="padding:5 0 0 0">
		<input type="text" name="winner" size="30"> <input type="submit" value="����" class="button">
		�� <b>�۹�ȣ-���̵�</b>(��: 123-abcde,135-wxyz11)<br>
		&nbsp;&nbsp;&nbsp;&nbsp;�ݵ�� <b>�۹�ȣ����</b>�� <b>������(-)</b>�� ��������.<br>
		&nbsp;&nbsp;&nbsp;&nbsp;<b>�θ� �̻�</b>�϶��� <b>��ǥ(,)</b> �� �� ��������.
	</td>
</tr>
</form>
</table>
<br>
<table border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<%
	If oPart.FTotalCount = 0 Then
		Response.Write "<tr bgcolor='#FFFFFF'><td colspan=2 align=center>��÷�ڰ� �����ϴ�.</td></tr>"
	Else
		For lp=0 to oPart.FTotalCount - 1
			Response.Write "<tr bgcolor='#FFFFFF'><td>" & oPart.FPartList(lp).FprtId & " ����</td><td>" & oPart.FPartList(lp).FprtUserId & "</td>"
			Response.Write "<td><input type='button' value='���' onClick=delWinner('"&oPart.FPartList(lp).FprtId&"-"&oPart.FPartList(lp).FprtUserId&"')></td></tr>"
		next
	End If
%>
</table>

<%
Function Proc()
	Dim vAction, vEvtID, vTitle, vWinner, vPrdId, vUserID, i
	vAction = RequestCheckvar(Request("action"),16)
	vEvtID = RequestCheckvar(Request("evtId"),10)
	vTitle = RequestCheckvar(Request("title"),64)
	vWinner = Split(Request("winner"),",")
	
	If vAction = "insert" Then
		For i = LBound(vWinner) To UBound(vWinner)
			vPrdId 	= Trim(Split(Trim(vWinner(i)),"-")(0))
			vUserID	= Trim(Split(Trim(vWinner(i)),"-")(1))
			
			dbACADEMYget.Execute " UPDATE [db_academy].[dbo].tbl_eventSub SET isWinner = 'o', winnerDate = getdate() WHERE evtId = '" & vEvtID & "' AND prtId = '" & vPrdId & "' AND userid = '" & vUserID & "' "
			
			vPrdId = ""
			vUserID = ""
		Next
	ElseIf vAction = "delete" Then
		dbACADEMYget.Execute " UPDATE [db_academy].[dbo].tbl_eventSub SET isWinner = null, winnerDate = null WHERE evtId = '" & vEvtID & "' AND prtId = '" & Split(vWinner(0),"-")(0) & "' AND userid = '" & Split(vWinner(0),"-")(1) & "' "
	End If
	
	Response.Redirect "event_winner.asp?evtId=" & vEvtID & "&title=" & vTitle & ""
	Response.End
	
End Function
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->