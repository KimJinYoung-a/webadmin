<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ν�Ÿ�׷� �̺�Ʈ�� ���� ���������
' Hieditor : 2016.06.23 ���¿� ����
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/admin/etc/only_sys/instagrameventCls.asp"-->

<%
Dim i, mode, contentsidx , ecode
dim instaidx,  evt_code, imgurl, instauserid, linkurl, isusing
	contentsidx = request("contentsidx")

dim oinsta
	set oinsta = new CInstagramevent
		oinsta.Frectcontentsidx=contentsidx
		if contentsidx <> "" then
			oinsta.fnGetinstagramevent_oneitem
			if oinsta.FResultCount > 0 then
				instaidx = oinsta.Foneitem.Fcontentsidx
				evt_code = oinsta.Foneitem.Fevt_code
				imgurl = oinsta.Foneitem.Fimgurl
				instauserid = oinsta.Foneitem.Fuserid
				linkurl = oinsta.Foneitem.Flinkurl
				isusing = oinsta.Foneitem.FIsusing
			end if
		end if

		
'���� idx���� �������(�űԵ��) NEW, �ƴҰ��(����) EDIT	
if instaidx = "" then 
	mode="NEW"
	ecode = contentsidx
else
	mode="EDIT"
	ecode =	evt_code
end if
%>

<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	
function frmedit(){
	if(frm.evt_code.value==""){
		alert('�̺�Ʈ �ڵ带 �Է��� �ּ���');
		frm.evt_code.focus();
		return;
	}

	if(frm.userid.value==""){
		alert('�Խ��� ID�� �Է��� �ּ���');
		frm.userid.focus();
		return;
	}

	if(frm.imgurl.value==""){
		alert('�̹��� �ּҸ� �Է��� �ּ���');
		frm.imgurl.focus();
		return;
	}
	
	if(frm.linkurl.value==""){
		alert('�Խù� ��ũ�� �Է��� �ּ���');
		frm.linkurl.focus();
		return;
	}
	
	frm.submit();
}

</script>

<form name="frm" method="post" action="instagramevent_proc.asp">
<input type="hidden" name="mode" value="<%=mode %>">
<input type="hidden" name="menupos" value="<%=menupos %>">
<input type="hidden" name="contentsidx" value="<%=contentsidx %>">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td align="left">
			<b>���ν�Ÿ�׷� �̺�Ʈ ������ ���� ���</b>
		</td>
	</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% IF contentsidx <> "" THEN%>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">��ȣ</td>
		<td colspan="2"><%=instaidx%></td>
	</tr>
	<% End if %>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�̺�Ʈ�ڵ�</td>
		<td colspan="2">
			<input type="text" name="evt_code" size="10" value="<%= ecode %>"/>
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�Խ���ID</td>
		<td colspan="2">
			<input type="text" name="userid" size="25" value="<% if mode="NEW" then response.write "10x10" else response.write instauserid end if %>"/>
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�̹���URL</td>
		<td colspan="2">
				<input type="text" name="imgurl" size="100" value="<%= imgurl %>" />
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�Խù���ũ</td>
		<td colspan="2">
				<input type="text" name="linkurl" size="100" value="<%= linkurl %>" />
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center"> ��뿩�� </td>
		<td colspan="2">
			<input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; 
			<input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������
		</td>
	</tr>

	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="3">
			<% if mode = "EDIT" or mode = "NEW" then %>
				<input type="button" name="editsave" value="����" onclick="frmedit()" />	
			<% end if %>
			
			<input type="button" name="editclose" value="���" onclick="self.close()" />
		</td>
	</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->