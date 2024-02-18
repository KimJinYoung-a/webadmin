<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim yyyymmdd
dim i

yyyymmdd	= req("yyyymmdd", "")

Dim strSql
strSql = " db_datamart.dbo.sp_Ten_Call_Absence_List ('" & yyyymmdd & "')"

db3_rsget.CursorLocation = adUseClient
db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc	

Dim rs 
If Not db3_rsget.EOF Then
	rs = db3_rsget.getRows()
End If 
db3_rsget.close


%>

<script language='javascript'>

window.onload = function () 
{
	document.getElementById("divShow").innerHTML = document.getElementById("divHide").innerHTML;
}

</script>


<div id="divShow"></div>
<p>
 
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">

        <td></td>
        <td>���Ź�ȣ</td>
        <td>�߽Ź�ȣ</td>
        <td>��ȭ�½ð�</td>
        <td>���</td>

	</tr>
<%

'' ���ϸ� ����
Function getWeekDay(ByVal val)
	Select Case Weekday(val)
	Case "1" getWeekDay = "<span style='color:red;'>�Ͽ���</span>"
	Case "2" getWeekDay = "������"
	Case "3" getWeekDay = "ȭ����"
	Case "4" getWeekDay = "������"
	Case "5" getWeekDay = "�����"
	Case "6" getWeekDay = "�ݿ���"
	Case "7" getWeekDay = "<span style='color:red;'>�����</span>"
	End Select 
End Function


Dim rowCnt
Dim sRs(20)

Dim dststr, dst, bigo, dcontext, lastapp, lastdata

If IsArray(rs) Then 
%>

	<%For i=0 To UBound(rs,2)%>
    <tr align="center" bgcolor="#FFFFFF">
	<%
		dst = rs(1,i)
		dcontext	= rs(3,i)
		lastapp		= rs(4,i)
		lastdata	= rs(5,i)
        
        dststr = ""
		If dst <> "" Then
			If InStr(dst,"07075490429") > 0 Then 
				dststr = "�ݼ���_��Ʈ"
				sRs(1) = sRs(1) + 1
			ElseIf InStr(dst,"07075490556") > 0 Then 
				dststr = "�繫��_��Ʈ"
				sRs(2) = sRs(2) + 1
			ElseIf InStr(dst,"07075490449") > 0 Then 
				dststr = "��ǥ��ȣ"
				sRs(3) = sRs(3) + 1
			ElseIf InStr(dst,"07075490448") > 0 Then 
				dststr = "��ǥ��ȣ"
				sRs(4) = sRs(4) + 1
			End If 
		End If 
		dststr = dststr & "(" & dst & ")"
		sRs(5) = sRs(5) + 1


		bigo = "��Ʈû���� ����"
		If dcontext = "tr_context" Then 
			bigo = "��������ȭ"
		ElseIf lastapp = "Busy" Or lastapp = "BackGround" Then 
			Select Case Replace(lastdata,"tenbyten/","")
				Case "tenbyten_call_main"
				bigo = "��ǥ��ȭ�ȳ���Ʈ�߲���"
				Case "tenbyten_call_recall"
				bigo = "��������ȭ�߸�Ʈ�߲���"
				Case "tenbyten_main"
				bigo = "��ǥ��ȭ�ȳ���Ʈ�߲���"
				Case "tenbyten_call_lunch"
				bigo = "���ɽð��ȳ���Ʈ�߲���"
				Case "tenbyten_call_workafter"
				bigo = "�����ľȳ���Ʈ�߲���"
				Case "tenbyten_call_workbefore"
				bigo = "�������ȳ���Ʈ�߲���"
				Case "tenbyten_forword"
				bigo = "�������ȳ���Ʈ�߲���"
				Case Else 
				bigo = ""
			End Select 
		End If 

	%>
		<td><%=i+1%></td>
		<td><%=dststr%></td>
		<td><%=rs(2,i)%></td>
		<td><%=rs(0,i)%></td>
		<td><%=bigo%></td>
	</tr>
	<%Next%>
<%
End If 
%>
</table>


<!-- ���Ӹ� -->
<div id="divHide" style="display:none;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>�Ϻ� ��������ȭ ����</td>
		<td>���Ź�ȣ</td>
		<td>�ݼ�</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td rowspan="5"><%=yyyymmdd%><br><%=getWeekDay(yyyymmdd)%></td>
		<td>�ݼ���_��Ʈ(07075490429)</td>
		<td align="right"><%=FormatNumber(sRs(1),0)%></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>�繫��_��Ʈ(07075490556)</td>
		<td align="right"><%=FormatNumber(sRs(2),0)%></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>��ǥ��ȣ(07075490449)</td>
		<td align="right"><%=FormatNumber(sRs(3),0)%></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>��ǥ��ȣ(07075490448)</td>
		<td align="right"><%=FormatNumber(sRs(4),0)%></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>�հ�</td>
		<td align="right"><%=FormatNumber(sRs(5),0)%></td>
	</tr>
</table>
</div>
<!-- ���Ӹ� -->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
