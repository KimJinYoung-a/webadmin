<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/etc/outmallConfirmCls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
Dim mallid, sellyn, makerid, isRegCnt, mallidArr, i
mallid	= request("mallid")
sellyn	= request("sellyn")
makerid	= request("makerid")

'####### �̹� ���� ������ ��ϵǾ��� �� �˻� #######
mallidArr = Split(mallid, ",")
for i = 0 to UBound(mallidArr)
	if (Trim(mallidArr(i)) <> "") then
		isRegCnt = fnIsRegedHopeCnt(Trim(mallidArr(i)), makerid)
		If isRegCnt > 0 Then
			response.write "<script language='javascript'>alert('�̹� �����û �ϼ̽��ϴ�[" & Trim(mallidArr(i)) & "]');window.close();</script>"
			response.end
		End If
	end if
next
'###### �̹� ���� ������ ��ϵǾ��� �� �˻� �� #####

%>
<script language='javascript'>
function frmsubmit(){
	var frm = document.frm;
	if(frm.sugiadminText.value == ''){
		alert('������ �Է��ϼ���');
		frm.sugiadminText.focus();
		return;
	}
	frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="POST" action="/admin/etc/outmall/confirm_process.asp" style="margin:0px;">
<input type="hidden" name="cmdparam" value="sugiOK">
<input type="hidden" name="sugimallid" value="<%=mallid%>">
<input type="hidden" name="sugisellyn" value="<%=sellyn%>">
<input type="hidden" name="sugimakerid" value="<%=makerid%>">
<input type="hidden" name="sugiadminid" value="<%=session("ssBctID")%>">
<tr bgcolor="#FFFFFF" height="30">
	<td>
		[ �귣��ID : <%= makerid %> ]<br><br>
		��ü���� ������ �ڸ�Ʈ �Է� �� ���� ��ư<br>
		<input type="text" class="text" name="sugiadminText" size="100" onkeydown="if(event.keyCode==13){alert('���� ��ư�� Ŭ�����ּ���.');event.returnValue=false;}">
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="30" align="center">
	<td>
		<input type="button" class="button" value="����" onclick="frmsubmit();">&nbsp;&nbsp;
		<input type="button" class="button" value="���" onclick="self.close();">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
