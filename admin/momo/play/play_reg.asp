<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������
' Hieditor : 2010.12.22 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim oplay,i
dim playSn ,startdate ,enddate, playLinkType, linkURL, evt_code ,isusing ,regdate
	playSn = requestcheckvar(request("playSn"),8)

'//��
set oplay = new cplayList
	oplay.frectplaySn = playSn
	oplay.FPageSize = 1
	oplay.FCurrPage = 1
	
	'//�����ϰ�쿡�� ����
	if playSn <> "" then
		oplay.fplay_list()
	end if
	
	if oplay.ftotalcount > 0 then
		playSn = oplay.FItemList(0).fplaySn
		startdate = oplay.FItemList(0).fstartdate
		enddate = oplay.FItemList(0).fenddate
		playLinkType = oplay.FItemList(0).fplayLinkType
		linkURL = oplay.FItemList(0).flinkURL
		evt_code = oplay.FItemList(0).fevtCode
		isusing = oplay.FItemList(0).fisusing
		regdate = oplay.FItemList(0).fregdate
	end if
	
	If isusing = "" Then
		isusing = "Y"
	End IF
%>

<script language="javascript">

	//����
	function reg(){

		if(frm.playLinkType[0].checked&&frm.evt_code.value=='') {
		alert('�̺�Ʈ��ȣ�� �Է����ּ���');
		frm.evt_code.focus();
		return;
		}
		if(frm.playLinkType[1].checked&&frm.linkURL.value=='') {
		alert('��Ʈ�� ������ URL�� �Է����ּ���');
		frm.linkURL.focus();
		return;
		}
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

		frm.action='/admin/momo/play/play_process.asp';
		frm.mode.value='add';
		frm.submit();
	}

	function chgLinkType(tp) {
		if(tp=="E") {
			document.getElementById("lyEvent").style.display="";
			document.getElementById("lyItem").style.display="none";
		} else {
			document.getElementById("lyItem").style.display="";
			document.getElementById("lyEvent").style.display="none";
		}
	}
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode" >
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">��ȣ</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= playSn %><input type="hidden" name="playSn" value="<%= playSn %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����Ÿ��</td>
	<td bgcolor="#FFFFFF" align="left">
		<label onClick="chgLinkType('E');"><input type="radio" name="playLinkType" value="E" <%=chkIIF(playLinkType="E" or playLinkType="","checked","")%>>�̺�Ʈ</label>&nbsp;
		<label onClick="chgLinkType('I');"><input type="radio" name="playLinkType" value="I" <%=chkIIF(playLinkType="I","checked","")%>>��������</label>
	</td>
</tr>
<tr align="center" id="lyEvent" bgcolor="<%= adminColor("tabletop") %>"> <%=chkIIF(playLinkType="E" or playLinkType="","","style='display:none;'")%>
	<td>�̺�Ʈ��ȣ</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="evt_code" size=10 value="<%= evt_code %>">			
	</td>
</tr>
<tr align="center" id="lyItem" bgcolor="<%= adminColor("tabletop") %>" <%=chkIIF(playLinkType="I","","style='display:none;'")%>>
	<td>��Ʈ��ũ</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="linkURL" size=50 value="<%= linkURL %>">			
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
			<option value="Y" <%=chkIIF(isusing="Y","selected","")%>>���</option>
			<option value="N" <%=chkIIF(isusing="N","selected","")%>>������</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="FFFFFF">
	<td colspan=2><input type="button" onclick="reg();" value="����" class="button"></td>
</tr>
</form>
</table>

<%
	set oplay = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->