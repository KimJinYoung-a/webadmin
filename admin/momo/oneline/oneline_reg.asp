<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ���ٳ���
' Hieditor : 2010.11.23 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ooneline,i , onelineid,startdate,enddate,winnerdate,comment,regdate,isusing , winner ,winnercomment
	onelineid = requestcheckvar(request("onelineid"),8)

'//��
set ooneline = new coneline_list
	ooneline.frectonelineid = onelineid
	
	'//�����ϰ�쿡�� ����
	if onelineid <> "" then
		ooneline.foneline_oneitem()
	end if
		
	if ooneline.ftotalcount > 0 then
		onelineid = ooneline.FOneItem.fonelineid
		startdate = ooneline.FOneItem.fstartdate
		enddate = ooneline.FOneItem.fenddate
		winnerdate = ooneline.FOneItem.fwinnerdate
		comment = ooneline.FOneItem.fcomment
		regdate = ooneline.FOneItem.fregdate
		isusing = ooneline.FOneItem.fisusing
		winner = ooneline.FOneItem.fwinner
		winnercomment = ooneline.FOneItem.fwinnercomment			
	end if
%>

<script language="javascript">

	//����
	function reg(){
		if (frm.startdate.value==''){
		alert('�������� �Է����ּ���');
		frm.startdate.focus();
		return;
		}		
		if (frm.enddate.value==''){
		alert('�������� �Է����ּ���');
		frm.enddate.focus();
		return;
		}
		if (frm.winnerdate.value==''){
		alert('��÷���� �Է����ּ���');
		frm.winnerdate.focus();
		return;
		}							
		if (frm.isusing.value==''){
		alert('��뿩�θ� �������ּ���');
		return;
		}
		
		frm.action='oneline_process.asp';
		frm.mode.value='edit';
		frm.submit();
	}
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>ID</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= onelineid %><input type="hidden" name="onelineid" value="<%= onelineid %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>�Ⱓ</b><br></td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="startdate" size=10 value="<%= startdate %>">			
		<a href="javascript:calendarOpen3(frm.startdate,'������',frm.startdate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a> -
		<input type="text" name="enddate" size=10  value="<%= left(enddate,10) %>">
		<a href="javascript:calendarOpen3(frm.enddate,'��������',frm.enddate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>	
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>��÷��</b><br></td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="winnerdate" size=10 value="<%= winnerdate %>">			
		<a href="javascript:calendarOpen3(frm.winnerdate,'��÷��',frm.winnerdate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="comment" style="width:450px; height:100px;"><%=comment%></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��÷��ID</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="winner" value="<%=winner%>"> �س����� �ʿ��� ��츸 �Է��ϼ���	
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��÷����</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="winnercomment" style="width:450px; height:100px;"><%=winnercomment%></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��뿩��</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="isusing" value="<%=isusing%>">
			<option value="" <% if isusing = "" then response.write " selected" %>>��뿩��</option>
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>			
	</td>
</tr>
<tr align="center" bgcolor="FFFFFF">
	<td colspan=2><input type="button" onclick="reg();" value="����" class="button"></td>
</tr>
</form>
</table>
<%
	set ooneline = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->