<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������
' Hieditor : 2010.11.15 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim oforecast,i
dim cardidx ,startdate ,enddate ,isusing ,regdate
	cardidx = requestcheckvar(request("cardidx"),8)

'//��
set oforecast = new cforecast_list
	oforecast.frectcardidx = cardidx
	
	'//�����ϰ�쿡�� ����
	if cardidx <> "" then
	oforecast.fcard_oneitem()
	end if
	
	if oforecast.ftotalcount > 0 then
		cardidx = oforecast.FOneItem.fcardidx
		startdate = oforecast.FOneItem.fstartdate
		enddate = oforecast.FOneItem.fenddate
		isusing = oforecast.FOneItem.fisusing
		regdate = oforecast.FOneItem.fregdate
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
		if (frm.isusing.value==''){
		alert('��뿩�θ� �������ּ���');
		return;
		}
		
		frm.action='/admin/momo/forecast/card_process.asp';
		frm.mode.value='add';
		frm.submit();
	}

</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode" >
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ȣ</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= cardidx %><input type="hidden" name="cardidx" value="<%= cardidx %>">		
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
	set oforecast = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->