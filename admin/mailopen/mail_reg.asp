<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������ ���
' History : 2008.01.21 �ѿ�� ����
'			2012.05.09 ������ ���� �߰�
'			2012.12.04 ������ ���������׸� �߰�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
' ���系������ �ʿ��� �κи� ������ ����. ###############################################
Dim ret_txt,ret_txt_re , i
	ret_txt = request("ret_txt")
	ret_txt  = replace(ret_txt,vbcrlf," bbb ")
	ret_txt_re = split(ret_txt," bbb ")

Dim ret_txt_seach,ret_txt_seach_re
	ret_txt_seach = ""
	If ret_txt <> "" Then
		ret_txt_seach = ret_txt_seach + ret_txt_re(1)&" bbb "
		ret_txt_seach = ret_txt_seach + ret_txt_re(8)&" bbb "
		ret_txt_seach = ret_txt_seach + ret_txt_re(9)&" bbb "
		ret_txt_seach = ret_txt_seach + ret_txt_re(10)&" bbb "
		ret_txt_seach = ret_txt_seach + ret_txt_re(11)&" bbb "
		ret_txt_seach = ret_txt_seach + ret_txt_re(13)&" bbb "
		ret_txt_seach_re = split(ret_txt_seach," ")

		If ret_txt_seach_re(0) <> "mailzine" and ret_txt_seach_re(0) <> "mailzine_��ȸ��" and ret_txt_seach_re(0) <> "mailzine_fingers" and ret_txt_seach_re(0) <> "mailzine_�ΰŽ���ȸ��" and ret_txt_seach_re(0) <> "OFFLINE" Then
			response.write "<script>" &_
						   "	alert('ķ���������� Ȯ���ϼ���');" &_
						   "	location.replace('/admin/mailopen/mail_reg.asp?mode=add');" &_
						   "</script>" &_
			response.end
		End If
	End If
' ���系������ �ʿ��� �κи� ������ ����. ###############################################
%>
<script language="JavaScript">
function TnMailDataReg(){
	if(frm.title.value == ""){
		alert("�߼��̸��� �����ּ���");
		frm.title.focus();
	}else if(frm.gubun.value == ""){
		alert("�߼۱����� �����ּ���");
		frm.gubun.focus();
	}else if(frm.startdate.value == ""){
		alert("�߼۽��۽ð��� �����ּ���");
		frm.startdate.focus();
	}else if(frm.enddate.value == ""){
		alert("�߼�����ð��� �����ּ���");
		frm.enddate.focus();
	}else if(frm.reenddate.value == ""){
		alert("��߼�����ð��� �����ּ���");
		frm.reenddate.focus();
	}else if(frm.totalcnt.value == ""){
		alert("�Ѵ���ڼ��� �����ּ���");
		frm.totalcnt.focus();
	}else if(frm.realcnt.value == ""){
		alert("�ǹ߼������ �����ּ���");
		frm.realcnt.focus();
	}else if(frm.realpct.value == ""){
		alert("�ǹ߼ۺ����� �����ּ���");
		frm.realpct.focus();
	}else if(frm.filteringcnt.value == ""){
		alert("���͸������ �����ּ���");
		frm.filteringcnt.focus();
	}else if(frm.filteringpct.value == ""){
		alert("���͸������� �����ּ���");
		frm.filteringpct.focus();
	}else if(frm.successcnt.value == ""){
		alert("�����߼������ �����ּ���");
		frm.successcnt.focus();
	}else if(frm.successpct.value == ""){
		alert("�������� �����ּ���");
		frm.successpct.focus();
	}else if(frm.failcnt.value == ""){
		alert("���й߼������ �����ּ���");
		frm.failcnt.focus();
	}else if(frm.failpct.value == ""){
		alert("�������� �����ּ���");
		frm.failpct.focus();
	}else if(frm.opencnt.value == ""){
		alert("��������� �����ּ���");
		frm.opencnt.focus();
	}else if(frm.openpct.value == ""){
		alert("�������� �����ּ���");
		frm.openpct.focus();
	}else if(frm.noopencnt.value == ""){
		alert("�̿�������� �����ּ���");
		frm.noopencnt.focus();
	}else if(frm.noopenpct.value == ""){
		alert("�̿������� �����ּ���");
		frm.noopenpct.focus();
	}else{
		frm.submit();
	}
}
</script>

<!-- ǥ �߰��� ����-->
<table width="100%" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
		<font color="red"><strong>�� THUNDERMAIL �߼۵��</strong></font>
		<br>�� (OFFLINE���� �߼۵�Ͻÿ��� ķ���������� OFFLINE���� ��ĥ ��~!!)<br>
		��� : textarea�� �ι�°���� �ٹ����ٰ����� ���� OFFLINE���� ����~! ���� �ؾ���(���⸦ split�ϱ� ����)<br>
		ex)OFFLINE �ٹ����ٰ�����
    </td>
    <td align="right">
    </td>        
</tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form action="" name="frm1" method="post">
<tr bgcolor=#FFFFFF>
	<td>
		<textarea name="ret_txt" cols="120" rows="10"></textarea>
	</td>
	<td>
		<input type="button" value="����" onclick="javascript:frm1.submit();" class="button">
	</td>
</tr>
</form>
</table>

<% If ret_txt <> "" Then %>
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<form action="/admin/mailopen/mail_process.asp" name="frm" method="post">
	<input type="hidden" name="mode" value="add">
	<tr bgcolor=#FFFFFF>
		<td align="center">�߼۱���</td>
		<td align="left">
			<select name="gubun">
				<option>�����ϼ���</option>
				<option value="mailzine">mailzine</option>
				<option value="fingers">fingers</option>
				<option value="mailzine_not">mailzine_not</option>
				<option value="fingers_not">fingers_not</option>
				<option value="OFFLINE">OFFLINE</option>
			<!-- <option value="academy">academy</option> -->
			</select>
		</td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">�߼��̸�</td>
		<td align="left"><input name="title" size="25" type="text" value="<%= ret_txt_seach_re(0) %>_<%=ret_txt_seach_re(11)%>"></td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">�߼۽��۽ð�</td>
		<td align="left"><input name="startdate" size="25" type="text" value="<%= ret_txt_seach_re(11) %>"></td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">�߼�����ð�</td>
		<td align="left"><input name="enddate" size="25" type="text" value="<%= ret_txt_seach_re(21) %>"></td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">��߼�����ð�</td>
		<td align="left"><input name="reenddate" size="25" type="text" value="<%= ret_txt_seach_re(34) %>"></td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">�Ѵ���ڼ�</td>
		<td align="left"><input name="totalcnt" size="15" type="text" value="<%= ret_txt_seach_re(4) %>"></td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">�ǹ߼����(����)</td>
		<td align="left"><input name="realcnt" size="15" type="text" value="<%= ret_txt_seach_re(15) %>">
		<input name="realpct" size="10" type="text" value="<%= round((ret_txt_seach_re(15)/ret_txt_seach_re(4))*100,1) %>">
		</td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">���͸����</td>
		<td align="left"><input name="filteringcnt" size="15" type="text" value="<%= ret_txt_seach_re(4)- ret_txt_seach_re(15) %>">
		<input name="filteringpct" size="10" type="text" value="<%= round(((ret_txt_seach_re(4)- ret_txt_seach_re(15))/ret_txt_seach_re(4))*100,1) %>">
		</td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">�����߼����(����)</td>
		<td align="left"><input name="successcnt" size="15" type="text" value="<%= ret_txt_seach_re(28) %>">
		<input name="successpct" size="10" type="text" value="<%= round((ret_txt_seach_re(28)/ret_txt_seach_re(15))*100,1) %>">
		</td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">���й߼����</td>
		<td align="left"><input name="failcnt" size="15" type="text" value="<%= ret_txt_seach_re(41) %>">
		<input name="failpct" size="10" type="text" value="<%= round((ret_txt_seach_re(41)/ret_txt_seach_re(15))*100,1) %>">
		</td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">�������</td>
		<td align="left"><input name="opencnt" size="15" type="text" value="<%= ret_txt_seach_re(56) %>">
		<input name="openpct" size="10" type="text" value="<%= round(( ret_txt_seach_re(56)/ret_txt_seach_re(28))*100,1) %>">
		</td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">�̿������</td>
		<td align="left"><input name="noopencnt" size="15" type="text" value="<%= ret_txt_seach_re(28)-ret_txt_seach_re(56) %>">
		<input name="noopenpct" size="10" type="text" value="<%= round(((ret_txt_seach_re(28)-ret_txt_seach_re(56))/ret_txt_seach_re(28))*100,1) %>">
		</td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">�߼۸��Ϸ�</td>
		<td align="left">
			<input type="hidden" name="mailergubun" value="THUNDERMAIL">
			THUNDERMAIL
		</td>
	</tr>
	<!--<tr bgcolor=#FFFFFF>
		<td align="center">���� ����</td>
		<td align="left">
			<% for i = 0 to 73 %>
				<% response.write ret_txt_seach_re(i)& "_"&i&"<br>" %>
			<% next %>
		</td>
	</tr>-->
	<tr bgcolor=#FFFFFF>
		<td align="center" colspan=2><input type="button" value="����" onclick="TnMailDataReg();" class="button"></td>
	</tr>
	</form>
	</table>
<% End If %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->