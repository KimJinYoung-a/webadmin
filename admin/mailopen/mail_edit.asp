<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������ ���
' History : 2008.01.21 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->

<%
Dim omd ,idx,mode
	idx = request("idx")
	mode = request("mode")
	
	If idx = "" Then idx=0

set omd = New CMailzineOne
	omd.GetMailingOne idx
%>

<script language="JavaScript">

function TnMailDataReg(){
	if(frm.title.value == ""){
		alert("�߼��̸��� �����ּ���");
		frm.title.focus();
	}
	else if(frm.gubun.value == ""){
		alert("�߼۱����� �����ּ���");
		frm.gubun.focus();
	}
	else if(frm.startdate.value == ""){
		alert("�߼۽��۽ð��� �����ּ���");
		frm.startdate.focus();
	}
	else if(frm.enddate.value == ""){
		alert("�߼�����ð��� �����ּ���");
		frm.enddate.focus();
	}
	else if(frm.reenddate.value == ""){
		alert("��߼�����ð��� �����ּ���");
		frm.reenddate.focus();
	}
	else if(frm.totalcnt.value == ""){
		alert("�Ѵ���ڼ��� �����ּ���");
		frm.totalcnt.focus();
	}
	else if(frm.realcnt.value == ""){
		alert("�ǹ߼������ �����ּ���");
		frm.realcnt.focus();
	}
	else if(frm.realpct.value == ""){
		alert("�ǹ߼ۺ����� �����ּ���");
		frm.realpct.focus();
	}
	else if(frm.filteringcnt.value == ""){
		alert("���͸������ �����ּ���");
		frm.filteringcnt.focus();
	}
	else if(frm.filteringpct.value == ""){
		alert("���͸������� �����ּ���");
		frm.filteringpct.focus();
	}
	else if(frm.successcnt.value == ""){
		alert("�����߼������ �����ּ���");
		frm.successcnt.focus();
	}
	else if(frm.successpct.value == ""){
		alert("�������� �����ּ���");
		frm.successpct.focus();
	}
	else if(frm.failcnt.value == ""){
		alert("���й߼������ �����ּ���");
		frm.failcnt.focus();
	}
	else if(frm.failpct.value == ""){
		alert("�������� �����ּ���");
		frm.failpct.focus();
	}
	else if(frm.opencnt.value == ""){
		alert("��������� �����ּ���");
		frm.opencnt.focus();
	}
	else if(frm.openpct.value == ""){
		alert("�������� �����ּ���");
		frm.openpct.focus();
	}
	else if(frm.noopencnt.value == ""){
		alert("�̿�������� �����ּ���");
		frm.noopencnt.focus();
	}
	else if(frm.noopenpct.value == ""){
		alert("�̿������� �����ּ���");
		frm.noopenpct.focus();
	}
	else if(frm.isusing.value == ""){
		alert("��뿩�θ� ������ �ּ���");
		frm.isusing.focus();		
	}
	else if(frm.mailergubun.value == ""){
		alert("�߼۸��Ϸ��� ������ �ּ���");
		frm.mailergubun.focus();		
	}	
	else{
		frm.submit();
	}
}

</script>

<table width="100%" border="0" class="a" cellpadding="1" cellspacing="1" bgcolor="#BABABA" align="center">
<form action="/admin/mailopen/mail_process.asp" name="frm" method="post">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="idx" value="<% = idx %>">
<tr bgcolor=#FFFFFF>
	<td align="center">�߼۱���</td>
	<td align="left">
		<input name="gubun" size="25" type="text" value="<% = omd.fgubun %>">
	</td>
</tr>
<tr bgcolor=#FFFFFF>
	<td align="center">�߼��̸�</td>
	<td align="left"><input name="title" size="25" type="text" value="<% = omd.Ftitle %>"></td>
</tr>
<tr bgcolor=#FFFFFF>
	<td align="center">�߼۽��۽ð�</td>
	<td align="left"><input name="startdate" size="25" type="text" value="<% = omd.Fstartdate %>"></td>
</tr>	
<tr bgcolor=#FFFFFF>
	<td align="center">�߼�����ð�</td>
	<td align="left"><input name="enddate" size="25" type="text" value="<% = omd.Fenddate %>"></td>
</tr>			
<tr bgcolor=#FFFFFF>
	<td align="center">��߼�����ð�</td>
	<td align="left"><input name="reenddate" size="25" type="text" value="<% = omd.Freenddate %>"></td>
</tr>	
<tr bgcolor=#FFFFFF>
	<td align="center">�Ѵ���ڼ�</td>
	<td align="left"><input name="totalcnt" size="15" type="text" value="<% = omd.Ftotalcnt %>"></td>
</tr>			
<tr bgcolor=#FFFFFF>
	<td align="center">�ǹ߼����(����)</td>
	<td align="left"><input name="realcnt" size="15" type="text" value="<% = omd.Frealcnt %>">
	<input name="realpct" size="10" type="text" value="<% = omd.Frealpct %>">
	</td>
</tr>	
<tr bgcolor=#FFFFFF>
	<td align="center">���͸����</td>
	<td align="left"><input name="filteringcnt" size="15" type="text" value="<% = omd.Ffilteringcnt %>">
	<input name="filteringpct" size="10" type="text" value="<% = omd.Ffilteringpct %>">
	</td>
</tr>	
<tr bgcolor=#FFFFFF>
	<td align="center">�����߼����(����)</td>
	<td align="left"><input name="successcnt" size="15" type="text" value="<% = omd.Fsuccesscnt %>">
	<input name="successpct" size="10" type="text" value="<% = omd.Fsuccesspct %>">
	</td>
</tr>		
<tr bgcolor=#FFFFFF>
	<td align="center">���й߼����</td>
	<td align="left"><input name="failcnt" size="15" type="text" value="<% = omd.Ffailcnt %>">
	<input name="failpct" size="10" type="text" value="<% = omd.Ffailpct %>">
	</td>
</tr>		
<tr bgcolor=#FFFFFF>
	<td align="center">�������</td>
	<td align="left"><input name="opencnt" size="15" type="text" value="<% = omd.Fopencnt %>">
	<input name="openpct" size="10" type="text" value="<% = omd.Fopenpct %>">
	</td>
</tr>			
<tr bgcolor=#FFFFFF>
	<td align="center">�̿������</td>
	<td align="left"><input name="noopencnt" size="15" type="text" value="<% = omd.Fnoopencnt %>">
	<input name="noopenpct" size="10" type="text" value="<% = omd.Fnoopenpct %>">
	</td>
</tr>
<tr bgcolor=#FFFFFF>
	<td align="center">��뿩��</td>
	<td align="left">
		<select name="isusing">
			<option value="Y" <% if omd.Fisusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if omd.Fisusing = "N" then response.write " selected" %>>N</option>
		</select>
	</td>
</tr>
<tr bgcolor=#FFFFFF>
	<td align="center">�߼۸��Ϸ�</td>
	<td align="left">
		<% drawmailergubun "mailergubun" , omd.fmailergubun , "" %>
	</td>
</tr>
<tr bgcolor=#FFFFFF>
	<td align="center" colspan=2>
		<input type="button" value="����" onclick="javascript:TnMailDataReg();" class="button">
	</td>	
</tr>
</form>
</table>

<% set omd = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->