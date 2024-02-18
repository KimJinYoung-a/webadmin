<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/seminar/seminarCls.asp"-->
<%
'####################################################
' Description : ���̳��� ��� �� ����Ʈ ������
' History : 2012.10.24 ������ ����
'####################################################

Dim Semi
Dim sMode, i
Dim idx, mRoomName, mMaxSu, mOrderNo, mIsusing
idx = request("idx")

SET Semi = new CSeminarManage

IF idx <> "" THEN
	sMode = "U"
	Semi.Fidx = idx
	Semi.Modify

	mRoomName = Semi.Froomname
	mMaxSu = Semi.FMaxSu
	mOrderNo = Semi.ForderNo
	mIsusing = Semi.FIsusing
Else
	sMode = "I"
END IF
	Semi.List
%>
<script language="javascript">
function jsModiCode(no){
	self.location.href = "popSeminarRoom.asp?idx="+no;
}

function jsRegCode(){
	var frm = document.frmReg;
	if(!frm.roomname.value) {
		alert("���̳��� �̸��� �Է��� �ּ���");
		frm.roomname.focus();
		return false;
	}

	if(!frm.MaxSu.value) {
		alert("���̳��� �ִ� �����ο��� �Է��� �ּ���");
		frm.MaxSu.focus();
		return false;
	}

	if(!frm.orderNo.value) {
		alert("���̳��� ���Ĺ�ȣ�� �Է��� �ּ���");
		frm.orderNo.focus();
		return false;
	}

	if(!frm.isusing(0).checked && !frm.isusing(1).checked) {
		alert("��뿩�ο� üũ�ϼ���");
		frm.isusing(0).focus();
		return false;
	}
	return true;
}
</script>
<table width="385" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td colspan="2"><!--//�ڵ� ��� �� ����-->
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a" >
		<form name="frmReg" method="post" action="procRoom.asp" onSubmit="return jsRegCode();">
		<input type="hidden" name="sM" value="<%=sMode%>">
<% IF idx <> "" THEN %>
		<input type="hidden" name="idx" value="<%=Semi.Fidx%>">
<% End If %>
		<tr>
			<td>+ ���̳��� ��� �� ����</td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
				<tr>
					<td bgcolor="#EFEFEF"   align="center">���̳��� ��</td>
					<td bgcolor="#FFFFFF">
						<input type="text" size="30" maxlength="30" name="roomname" value="<%=mRoomName%>">
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">�ִ� �����ο�</td>
					<td bgcolor="#FFFFFF">
						<input type="text" size="10" maxlength="5" name="MaxSu" value="<%=mMaxSu%>">
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">���Ĺ�ȣ</td>
					<td bgcolor="#FFFFFF">
						<input type="text" size="10" maxlength="5" name="orderNo" value="<%=mOrderNo%>"><br>
						40���ϸ� ������ġ, 40�ʰ��� ����
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"  align="center">��뿩��</td>
					<td bgcolor="#FFFFFF">
						<input type="radio" value="Y" name="isusing" <%=chkIIF(mIsusing = "Y","checked","")%>>���
						<input type="radio" value="N" name="isusing" <%=chkIIF(mIsusing = "N","checked","")%> >������
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="right"><input type="image" src="/images/icon_save.gif"></td>
		</tr>
		<tr>
			<td colspan="2"><hr width="100%"></td>
		</tr>
		</form>
		</table>
	</td>
</tr>
</table>
<table width="385" border="0" cellpadding="3" cellspacing="0" class="a" >
<form name="frmSearch" method="post" action="PopUserList.asp">
<tr>
	<td colspan="2">+ ���̳��� ����Ʈ</td>
</tr>
<tr>
	<td colspan="2">
		<div id="divList" style="height:305px;overflow-y:scroll;">
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
		<tr bgcolor="#EFEFEF">
			<td align="center">���̳��� ��</td>
			<td align="center">�ִ� �����ο�</td>
			<td align="center">��뿩��</td>
			<td align="center">���Ĺ�ȣ</td>
			<td align="center">ó��</td>
		</tr>
<%
	If Semi.Fresultcount > 0 THEN
		For i = 0 To Semi.Fresultcount -1
%>
		<tr bgcolor="#FFFFFF">
			<td align="center"><%=Semi.FItemList(i).Froomname%></td>
			<td align="center"><%=Semi.FItemList(i).FMaxSu%></td>
			<td align="center"><%=Semi.FItemList(i).Fisusing%></td>
			<td align="center"><%=Semi.FItemList(i).ForderNo%></td>
			<td align="center"><input type="button" value="����" onClick="javascript:jsModiCode('<%=Semi.FItemList(i).Fidx%>');" class="input_b"></td>
		</tr>
<%
		Next
	Else
%>
		<tr bgcolor="#FFFFFF"><td colspan="5" align="center">��ϵ� ������ �����ϴ�.</td></tr>
<%End if%>
		</table>
		</div>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->