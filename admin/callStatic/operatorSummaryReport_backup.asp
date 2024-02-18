<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim tenUserID, vSessionID
dim i

Dim sDate	: sDate	= req("sDate", Date() - 30)
Dim eDate	: eDate	= req("eDate", Date())

tenUserID = req("tenUserID", "")
vSessionID = session("ssBctId")

Dim strSql
strSql = " db_datamart.dbo.sp_Ten_Call_Report ('" & sDate & "', '" & eDate & "', '" & tenUserID & "')"

db3_rsget.CursorLocation = adUseClient
db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc	

Dim rs 
If Not db3_rsget.EOF Then
	rs = db3_rsget.getRows()
End If 
db3_rsget.close


%>

<script language='javascript'>


function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);
	
	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

function showDays(yyyymmdd, days)
{
	var f = document.frmWrite;
	
	f.yyyymmdd.value = yyyymmdd;
	f.workDays.value = days;

	document.getElementById("today").innerHTML = yyyymmdd;
	document.getElementById("days").style.display = "block";
}

function saveDays()
{
	var f = document.frmWrite;
	
	if (f.tenUserID.value)
		f.submit();

}

function searchTenUser(tenUserID)
{
	var arr = tenUserID.split(" - ");
	var f = document.frm;

	f.tenUserID.value = arr[1];
	f.submit();

}
</script>

<div id="days" style="display:none; position:absolute; z-index: 1; left: 500px; top: 180px; background-color:#FFFFFF; border:solid 1px #000000; padding:20px; width:200px; " align="center">
	<form name="frmWrite" action="operatorWorkDayProc.asp">
	<input type="hidden" name="sDate" value="<%=sDate%>">
	<input type="hidden" name="eDate" value="<%=eDate%>">
	<input type="hidden" name="tenUserID" value="<%=tenUserID%>">
	<input type="hidden" name="yyyymmdd">
		<span id="today"></span>
		&nbsp;
		�ٹ�
		<br><br>
		<select name="workDays">
			<option value="0.0">0</option>
			<option value="0.5">0.5</option>
			<option value="1.0">1</option>
		</select>
		��
		<br><br>
		<input type="button" class="button" value="����" onclick="saveDays();">
		&nbsp;
		<input type="button" class="button" value="���" onclick="document.getElementById('days').style.display = 'none';">
	</form>
</div>


<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
            ��¥ : <input type="text" size="10" name="sDate" value="<%=sDate%>" onClick="jsPopCal('frm','sDate');" style="cursor:hand;">
            ~<input type="text" size="10" name="eDate" value="<%=eDate%>" onClick="jsPopCal('frm','eDate');" style="cursor:hand;">
			&nbsp;

			����ID
			<input type="text" class="text" name="tenUserID" value="<%=tenUserID%>">
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->
 
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td rowspan="2" width="120">
		<%If tenUserID <> "" Then %>
			��¥
		<%Else %>
			����ID
		<%End If%>
		</td>
        <td colspan="3">�����հ�</td>
        <td colspan="3">�߽��հ�</td>
        <td colspan="3">�հ�</td>
		<td rowspan="2">�ٹ��ϼ�</td>
		<td rowspan="2">�Խ���ó����</td>
		<td rowspan="2">�������ȭ��</td>
		<td rowspan="2">����հԽ���</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">

        <td>�Ǽ�</td>
        <td>�ð�</td>
        <td>���</td>

        <td>�Ǽ�</td>
        <td>�ð�</td>
        <td>���</td>

        <td>�Ǽ�</td>
        <td>�ð�</td>
        <td>���</td>

</tr>
<%
Dim servicePoint, servicePointText, pnt1, pnt2, pnt3, pnt4
Dim rowCnt
Dim sRs(20)

' �ʵ����͸� �ú��� �������� ��ȯ
Function sec2time(ByVal sec)
	sec2time = Int(sec / 3600) & ":" & Right("0"&(Int(sec/60) Mod 60),2) & ":" & Right("0"&(sec Mod 60),2)
End Function 

If IsArray(rs) Then 
	rowCnt = UBound(rs,2) + 1
%>

	<%For i=0 To UBound(rs,2)%>
    <tr align="right" bgcolor="#FFFFFF">
	<%
		' Row �ջ�
		sRs(1) = sRs(1) + CDbl(rs(1,i))
		sRs(2) = sRs(2) + CDbl(rs(2,i))
		sRs(3) = sRs(3) + CDbl(rs(3,i))
		sRs(4) = sRs(4) + CDbl(rs(4,i))
		sRs(5) = sRs(5) + CDbl(rs(5,i))
		sRs(6) = sRs(6) + CDbl(rs(6,i))
		sRs(7) = sRs(7) + CDbl(rs(7,i))
		sRs(8) = sRs(8) + CDbl(rs(8,i))
		sRs(9) = sRs(9) + CDbl(rs(9,i))
		sRs(10) = sRs(10) + CDbl(rs(10,i))
		sRs(11) = sRs(11) + CDbl(rs(11,i))
		sRs(12) = sRs(12) + CDbl(rs(12,i))
		sRs(13) = sRs(13) + CDbl(rs(13,i))
		sRs(14) = sRs(14) + CDbl(rs(14,i))

		rs(7,i) = rs(1,i) + rs(4,i)
		rs(8,i) = rs(2,i) + rs(5,i)

		' ��� �׸� ����
		If CDbl(rs(1,i)) > 0 Then
			rs(3,i) = CInt( CDbl(rs(2,i)) / CDbl(rs(1,i)) )
		End If 
		If CDbl(rs(4,i)) > 0 Then
			rs(6,i) = CInt( CDbl(rs(5,i)) / CDbl(rs(4,i)) )
		End If 
		If CDbl(rs(7,i)) > 0 Then
			rs(9,i) = CInt( CDbl(rs(8,i)) / CDbl(rs(7,i)) )
		End If 

		' �������ȭ��, ����հԽ���ó����
		If CDbl(rs(11,i)) > 0 Then
			rs(12,i) = CInt( CDbl(rs(8,i)) / CDbl(rs(11,i)) )	' �հ�ð� / �ٹ��ϼ�
			rs(14,i) = CDbl( CDbl(rs(13,i)) / CDbl(rs(11,i)) )	' �Խ���ó���� / �ٹ��ϼ�
		End If 
	%>
		<td align="left"><%=rs(0,i)%></td>
		<td><%=FormatNumber(rs(1,i),0)%></td>
		<td><%=sec2time(rs(2,i))%></td>
		<td><%=sec2time(rs(3,i))%></td>
		<td><%=FormatNumber(rs(4,i),0)%></td>
		<td><%=sec2time(rs(5,i))%></td>
		<td><%=sec2time(rs(6,i))%></td>
		<td><%=FormatNumber(rs(7,i),0)%></td>
		<td><%=sec2time(rs(8,i))%></td>
		<td><%=sec2time(rs(9,i))%></td>

		<td>
		<%If tenUserID <> "" Then %>
			<a href="javascript:showDays('<%=rs(0,i)%>','<%=FormatNumber(rs(11,i),1)%>');"><%=FormatNumber(rs(11,i),1)%></a>
		<%Else %>
			<a href="javascript:searchTenUser('<%=rs(0,i)%>');">
				<%=FormatNumber(rs(11,i),1)%>
			</a>
		<%End If%>
		</td>

		<td><%=rs(13,i)%></td>
		<td><%=sec2time(rs(12,i))%></td>
		<td><%=FormatNumber(rs(14,i),1)%></td>
	</tr>
	<%Next%>
    <tr align="right" bgcolor="#FFFFFF">
	<%
		' �հ� ��� �׸� ����
		If CDbl(sRs(1)) > 0 Then
			sRs(3) = CInt( CDbl(sRs(2)) / CDbl(sRs(1)) )
		End If 
		If CDbl(sRs(4)) > 0 Then
			sRs(6) = CInt( CDbl(sRs(5)) / CDbl(sRs(4)) )
		End If 

		sRs(7) = sRs(1) + sRs(4)
		sRs(8) = sRs(2) + sRs(5)

		If CDbl(sRs(7)) > 0 Then
			sRs(9) = CInt( CDbl(sRs(8)) / CDbl(sRs(7)) )
		End If 

		' �������ȭ��, ����հԽ���ó����
		If CDbl(sRs(11)) > 0 Then
			sRs(12) = CInt( CDbl(sRs(8)) / CDbl(sRs(11)) )	' �հ�ð� / �ٹ��ϼ�
			sRs(14) = CDbl( CDbl(sRs(13)) / CDbl(sRs(11)) )	' �Խ���ó���� / �ٹ��ϼ�
		End If 
	%>
    	<td align="center"><b>�հ� or ���</b></td>
		<td><b><%=FormatNumber(sRs(1),0)%></b></td>
		<td><b><%=sec2time(sRs(2))%></b></td>
		<td><b><%=sec2time(sRs(3))%></b></td>
		<td><b><%=FormatNumber(sRs(4),0)%></b></td>
		<td><b><%=sec2time(sRs(5))%></b></td>
		<td><b><%=sec2time(sRs(6))%></b></td>
		<td><b><%=FormatNumber(sRs(7),0)%></b></td>
		<td><b><%=sec2time(sRs(8))%></b></td>
		<td><b><%=sec2time(sRs(9))%></b></td>

		<td><b><%=FormatNumber(sRs(11),1)%></b></td>
		<td><b><%=sRs(13)%></b></td>
		<td><b><%=sec2time(sRs(12))%></b></td>
		<td><b><%=FormatNumber(sRs(14),1)%></b></td>
    </tr>
<%
End If 
%>
</table>



<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
