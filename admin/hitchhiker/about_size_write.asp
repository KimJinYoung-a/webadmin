<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : HITCHHIKER ADMIN
'	History		: 2014.07.09 ���¿� ����
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/about_hitchhiker_contentsCls.asp"-->
<%
Dim mode, i
dim deviceidx, device_name, contents_size, sortnum, isusing
	isusing = request("isusing")
	deviceidx = request("deviceidx")
	mode = requestCheckvar(Request("mode"),10)
	sortnum = requestCheckvar(Request("sortnum"),5)
	device_name = requestCheckvar(Request("device_name"),32)
	contents_size = requestCheckvar(Request("contents_size"),32)

Dim ohitpc
set ohitpc = new CAbouthitchhiker
	ohitpc.Frectgubun="1"
	ohitpc.fnGetDeviceList
	
dim ohitm
set ohitm = new CAbouthitchhiker
	ohitm.Frectgubun="2"
	ohitm.fnGetDeviceList
%>

<script type='text/javascript'>
//�ű��Է� ���	
function frmreg(gubun){
	if (gubun==""){
		alert("��񱸺��� �������� �ʾҽ��ϴ�.������ ���� ���ּ���.");
		return;
	}
	
	//��񱸺п� ���� ó����
	//PC
	if (gubun=="1"){
		frm.contents_size.value=frm.newpcsizetextbox.value;
		if (frm.newpcsizetextbox.value==""){
			alert("����� �Է����ּ���");
			frm.newpcsizetextbox.focus();
			return;
		}
		
		var tempchkvalue = "";
		for (var i=0;i<frm.newpcsizeisusing.length;i++) {
			if (frm.newpcsizeisusing[i].checked==true) {
				tempchkvalue=frm.newpcsizeisusing[i].value;
			}
		}
		if(tempchkvalue==""){
			alert("��뿩�θ� ���� ���ּ���");
			return;
		}
		frm.isusing.value=tempchkvalue;
		frm.sortnum.value=frm.newpcsizesortnum.value;
		frm.contents_size.value=frm.newpcsizetextbox.value;
		frm.device_name.value=frm.newmobiledevicetextbox.value;
	//�����
	}else{
		if (frm.newmobiledevicetextbox.value==""){
			alert("��ǥ������ �Է����ּ���");
			frm.newmobiledevicetextbox.focus();
			return;
		}
		
		if (frm.newmobilesizetextbox.value==""){
			alert("����� �Է����ּ���");
			frm.newmobilesizetextbox.focus();
			return;
		}
		
		var temmobilechkvalue = "";
		for (var i=0;i<frm.newmobilesizeisusing.length;i++) {
			if (frm.newmobilesizeisusing[i].checked==true) {
				temmobilechkvalue=frm.newmobilesizeisusing[i].value;
			}
		}
		if(temmobilechkvalue==""){
			alert("��뿩�θ� ���� ���ּ���");
			return;
		}
		
		frm.isusing.value=temmobilechkvalue;
		frm.sortnum.value=frm.newmobilesizesortnum.value;
		frm.contents_size.value=frm.newmobilesizetextbox.value;
		frm.device_name.value=frm.newmobiledevicetextbox.value;
	}
	
	frm.deviceidx.value="";
	frm.gubun.value=gubun;
	frm.mode.value="sizeedit"
	frm.submit();
}

//������ ���� ���	
function frmedit(gubun,ix){
	if (gubun==""){
		alert("��񱸺��� �������� �ʾҽ��ϴ�.������ ���� ���ּ���.");
		return;
	}
	if (ix==""){
		alert("���������� �������� �ʾҽ��ϴ�.������ ���� ���ּ���.");
		return;
	}

	//��񱸺п� ���� ó����
	//PC
	if (gubun=="1"){
		var tmpdeviceidx = eval("frm.pcsizedeviceidx_"+ix);  //gubun=1(PC)idx
		frm.deviceidx.value=tmpdeviceidx.value;

		var tmppcsizetextbox = eval("frm.pcsizetextbox_"+ix); //gubun=1(PC) �������� ������
		frm.contents_size.value=tmppcsizetextbox.value;
		if (tmppcsizetextbox.value==""){
			alert("����� �Է����ּ���");
			eval("frm.pcsizetextbox_"+ix).focus();
			return;
		}

		var tmpsortnum = eval("frm.pcsizesortnum_"+ix); //gubun=1(PC) �켱����
		frm.sortnum.value=tmpsortnum.value;
		if (tmpsortnum.value==""){
			alert("�켱������ �Է��� �ּ���");
			eval("frm.pcsizesortnum_"+ix).focus();
			return;
		}
		
		var tmppcsizeisusing = eval("frm.pcsizeisusing_"+ix); //gubun=1(PC)�������� ��뿩��
		var tempchkvalue = "";
		for (var i=0;i<tmppcsizeisusing.length;i++) {
			if (tmppcsizeisusing[i].checked==true) {
				tempchkvalue=tmppcsizeisusing[i].value;
			}
		}
		if(tempchkvalue==""){
			alert("��뿩�θ� ���� ���ּ���");
			return;
		}

		frm.isusing.value=tempchkvalue;
	//�����
	}else{
		var tmpdeviceidx = eval("frm.mobiledeviceidx_"+ix); //gubun=2(�����)idx
		frm.deviceidx.value=tmpdeviceidx.value;
		
		var tmpdevicename = eval("frm.mobiledevicetextbox_"+ix); //����ϱ���
		frm.device_name.value=tmpdevicename.value;
		if (tmpdevicename.value==""){
			alert("����� ������ �Է��� �ּ���");
			eval("frm.mobiledevicetextbox_"+ix).focus();
			return;
		}

		var tmpmsizetextbox = eval("frm.mobilesizetextbox_"+ix); //gubun=2(�����)�������� ������
		frm.contents_size.value=tmpmsizetextbox.value;
		if (tmpdevicename.value==""){
			alert("����� ������ �Է��� �ּ���");
			eval("frm.mobiledevicetextbox_"+ix).focus();
			return;
		}

		var tmpsortnum = eval("frm.mobilesizesortnum_"+ix); //gubun=2(�����)�켱����
		frm.sortnum.value=tmpsortnum.value;
		if (tmpsortnum.value==""){
			alert("�켱������ �Է��� �ּ���");
			eval("frm.mobilesizesortnum_"+ix).focus();
			return;
		}


		var tmpmsizeisusing = eval("frm.mobilesizeisusing_"+ix); //gubun=2(�����)�������� ��뿩��
		var temmobilechkvalue = "";
		for (var i=0;i<tmpmsizeisusing.length;i++) {
			if (tmpmsizeisusing[i].checked==true) {
				temmobilechkvalue=tmpmsizeisusing[i].value;
			}
		}
		if(tempchkvalue==""){
			alert("��뿩�θ� ���� ���ּ���");
			return;
		}

		frm.isusing.value=temmobilechkvalue;
	}
	
	frm.gubun.value=gubun;
	frm.mode.value="sizeedit"
	frm.submit();
}
	
	function onlyNumDecimalInput(){  //�ѱ� �Է� �ȵǰ�
	var code = window.event.keyCode; 
	
	if ((code >= 48 && code <= 57) || (code >= 96 && code <= 105) || code == 110 || code == 190 || code == 8 || code == 9 || code == 13 || code == 46){ 
		window.event.returnValue = true; 
		return; 
	} 
	window.event.returnValue = false; 
	}

</script>

<form name="frm" method="post" action="about_size_proc.asp" >
<input type="hidden" name="mode" >
<input type="hidden" name="gubun" >
<input type="hidden" name="isusing" >
<input type="hidden" name="sortnum" >
<input type="hidden" name="deviceidx" >
<input type="hidden" name="device_name" >
<input type="hidden" name="contents_size" >
<input type="hidden" name="menupos" value="<%=menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td align="left">
			<b>�ؿ������� �⺻ ������</b>
		</td>
	</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%=adminColor("tabletop")%>">
		<td>�� PC �������� ������</td>
	</tr>
	
<!--PC�������� �ű� ������ ���-->
	<tr bgcolor="FFFFFF">
		<td>
			������	 <input type="text" name="newpcsizetextbox" value="" />
			�켱���� <input type="text" name="newpcsizesortnum" value="99" size="3" maxlength="3" onkeydown='onlyNumDecimalInput();' style="ime-mode:disabled" />
			��뿩�� <input type="radio" name="newpcsizeisusing" value="Y" /> Y <input type="radio" name="newpcsizeisusing" value="N" /> N
					 <input type="button" value="�űԵ��" class="button" onclick="frmreg('1')" />
		</td>
	</tr>
<!--PC�������� �ű� ������ ��� ��-->
	
	<tr bgcolor="FFFFFF">
		<td height=10></td>
	</tr>
	
<!--����PC�������� ������ ����Ʈ-->
	<% if ohitpc.FResultCount > 0 then %>
		<% for i = 0 to ohitpc.FResultCount - 1 %>
			<tr bgcolor="FFFFFF">
				<td<% if ohitpc.FItemList(i).FIsusing = "N" then %> bgcolor="CCCCCC" <% else %> bgcolor="FFFFFF" <% end if %>>
							<input type="hidden" name="pcsizedeviceidx_<%= i %>" value="<%= ohitpc.FItemList(i).FDeviceidx %>" />
					������	<input type="text" name="pcsizetextbox_<%= i %>" value="<%= trim(ohitpc.FItemList(i).FContentsSize) %>" />
					�켱����	<input type="text" name="pcsizesortnum_<%= i %>" value="<%= trim(ohitpc.FItemList(i).FSortnum) %>" size="3" maxlength="3" onkeydown='onlyNumDecimalInput();' style="ime-mode:disabled" />
					��뿩��	<input type="radio" name="pcsizeisusing_<%= i %>" value="Y" <% if ohitpc.FItemList(i).FIsusing = "Y" then response.write " checked" %> /> Y
							<input type="radio" name="pcsizeisusing_<%= i %>" value="N" <% if ohitpc.FItemList(i).FIsusing = "N" then response.write " checked" %> /> N
							<input type="button" value="����" class="button" onclick="frmedit('1','<%=i%>')"/>
				</td>
			</tr>
		<% next %>
	<% end if %>
<!--����PC�������� ������ ����Ʈ ��-->
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%=adminColor("tabletop")%>">
		<td>�� MOBILE �������� ���� �� ������</td>
	</tr>	
	
<!-- ����� �ű� ���� �� �������� ������ ���-->
	<tr bgcolor="FFFFFF">
		<td>
			��ǥ���� <input type="text" name="newmobiledevicetextbox" value="" />
			������	 <input type="text" name="newmobilesizetextbox" value="" />
			�켱���� <input type="text" name="newmobilesizesortnum" value="99" size="3" maxlength="3" onkeydown='onlyNumDecimalInput();' style="ime-mode:disabled" />
			��뿩�� <input type="radio" name="newmobilesizeisusing" value="Y" /> Y <input type="radio" name="newmobilesizeisusing" value="N" /> N
					 <input type="button" value="�űԵ��"  class="button" onclick="frmreg('2')"/>
		</td>
	</tr>
<!-- ����� �ű� ���� �� �������� ������ ��� ��-->
	
	<tr bgcolor="FFFFFF">
		<td height=10></td>
	</tr>
	
<!--��������� ���� �� �������� ������ ����Ʈ-->
	<% if ohitm.FResultCount > 0 then %>
		<% for i = 0 to ohitm.FResultCount - 1 %>
			<tr bgcolor="FFFFFF">
				<td <% if ohitm.FItemList(i).FIsusing = "N" then %> bgcolor="CCCCCC" <% else %> bgcolor="FFFFFF" <% end if %>>		
							<input type="hidden" name="mobiledeviceidx_<%= i %>" value="<%= ohitm.FItemList(i).FDeviceidx %>" />
					��ǥ����	<input type="text" name="mobiledevicetextbox_<%= i %>" value="<%= trim(ohitm.FItemList(i).FDevicename) %>" />
					������	<input type="text" name="mobilesizetextbox_<%= i %>" value="<%= trim(ohitm.FItemList(i).FContentsSize) %>" />
					�켱����	<input type="text" name="mobilesizesortnum_<%= i %>" value="<%= trim(ohitm.FItemList(i).FSortnum) %>" size="3" maxlength="3" onkeydown='onlyNumDecimalInput();' style="ime-mode:disabled"/>
					��뿩��	<input type="radio" name="mobilesizeisusing_<%= i %>" value="Y" <% if ohitm.FItemList(i).FIsusing = "Y" then response.write " checked" %> /> Y
							<input type="radio" name="mobilesizeisusing_<%= i %>" value="N" <% if ohitm.FItemList(i).FIsusing = "N" then response.write " checked" %> /> N
							<input type="button" value="����"  class="button" onclick="frmedit('2','<%=i%>')"/>
				</td>
			</tr>
		<% next %>
	<% end if %>
<!--��������� ���� �� �������� ������ ����Ʈ ��-->
</table>
</form>

<%
set ohitpc = nothing
set ohitm = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->