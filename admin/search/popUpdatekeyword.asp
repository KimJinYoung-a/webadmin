<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/search/itemCls.asp" -->
<%
Dim arritemid, arritemCnt
arritemid	= request("arritemid")
If Right(arritemid,1) = "," Then
	arritemid	= Left(arritemid, Len(arritemid) - 1)
End If
arritemCnt	= Ubound(Split(arritemid, ",")) + 1
%>
<style>
input:-ms-input-placeholder { color: #ADADAD; }
input::-webkit-input-placeholder { color: #ADADAD; }
input::-moz-placeholder { color: #ADADAD; }
input::-moz-placeholder { color: #ADADAD; }
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}
function closePop(){
	if (confirm('���������� �������� �ʰ� ����Ͻðڽ��ϱ�?')){
		self.close();
	}
}
function AllconfirmProcess(){
	if ($("#nextkeyword").val() == ""){
		alert("���� Ű���带 �Է��ϼ���");
		return false;
	}
	if ($("#etc").val() == ""){
		alert("��� �Է��ϼ���");
		return false;
	}

	if( $("#mode").val() == "U") {
		if ($("#prekeyword").val() == ""){
			alert("������ Ű���带 �Է��ϼ���");
			return false;
		}

		if( $("#prekeyword").val().indexOf(",")  > 0) {
			alert("������ Ű���忡 ,�� �Է��� �� �����ϴ�.");
			$("#prekeyword").val("");
			return false;
		}
	}

	if( $("#nextkeyword").val().indexOf(",")  > 0) {
		alert("���� Ű���忡 ,�� �Է��� �� �����ϴ�.");
		$("#nextkeyword").val("");
		return false;
	}

	if (confirm('<%=arritemCnt%>���� ��ǰ Ű���� ������ �ϰ������Ͻðڽ��ϱ�?')){
		document.frm.action = "/admin/search/keywordProc.asp"
		document.frm.submit();
	}
}
function chgSelectSH(v){
	if(v == 'U'){
		$("#prekeyword").show();
	}else{
		$("#prekeyword").hide();
		$("#prekeyword").val("");
	}
}
</script>
<table width="100%">
<form name="frm" method="POST">
<input type="hidden" name="cmdparam" value="allchk">
<input type="hidden" name="cksel" value="<%= arritemid %>">
<tr>
	<td align="LEFT"><strong>Ű���� ���� ����</strong></td>
</tr>
<tr>
	<td colspan="2">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="LEFT" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="20%">���� ����</td>
			<td bgcolor="#FFFFFF" align="LEFT">
				<select name="mode" class="select" id="mode" onchange="chgSelectSH(this.value);">
					<option value="I">���</option>
					<option value="U">����</option>
					<option value="D" selected>����</option>
				</select>
			</td>
		</tr>
		<tr align="LEFT" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="20%">���� Ű����</td>
			<td bgcolor="#FFFFFF" align="LEFT">
				<input type="text" size="25" class="text" id="prekeyword" name="prekeyword" placeholder="������ Ű���� �Է�" style="display:none;">
				<input type="text" size="25" class="text" id="nextkeyword" name="nextkeyword" placeholder="������ ���� Ű���� �Է�">
			</td>
		</tr>
		<tr align="LEFT" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="20%">���</td>
			<td bgcolor="#FFFFFF" align="LEFT">
				<input type="text" size="70" class="text" name="etc" id="etc" placeholder="�̷� ������ �����ϰ� �� �� �ֵ��� ��� �Է�">
			</td>
		</tr>
	</td>
</tr>	
</form>
</table>
<br/>
<table width="100%">
<tr>
	<td align="LEFT">* Ű����� ��,���� ������ �Ѱ��� Ű���带 �Է����ּ���.</td>
</tr>
<tr>
	<td align="center">
		<input type="button" class="button" value="�ϰ� ����" onclick="AllconfirmProcess();">&nbsp;
		<input type="button" class="button" value="���" onclick="closePop();">
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->