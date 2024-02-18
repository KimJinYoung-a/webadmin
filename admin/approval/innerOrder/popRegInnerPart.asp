<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<%
'###########################################################
' Description : ������û�� ���
' History : 2011.03.14 ������  ����
' 0 ��û/1 ������/ 5 �ݷ�/7 ����/ 9 �Ϸ�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/innerPartcls.asp"-->
<%

dim idx, mode

idx = requestCheckvar(Request("idx"),32)

if (idx = "") then
	idx = 0
	mode = "regnewpart"
else
	mode = "modifypart"
end if

'==============================================================================
dim oinnerpart
set oinnerpart = New CInnerPart

oinnerpart.FCurrPage = 1
oinnerpart.FPageSize = 1

oinnerpart.FRectIdx = idx

oinnerpart.GetInnerPartOne

if (mode = "modifypart") and (oinnerpart.FOneItem.Fidx = "") then
	response.write "�߸��� �����Դϴ�."
	response.end
end if

%>
<script language="javascript">

function jsReg(frm) {
	if (frm.divcd.value == "") {
		alert("���κμ� ������ �����ϼ���.");
		return;
	}

	if (frm.BIZSECTION_CD.value == "") {
		alert("ERP�μ��ڵ带 �����ϼ���.");
		return;
	}

	if (frm.scmid.value == "") {
		alert("���κμ��ڵ带 �����ϼ���.");
		return;
	}



	if (confirm("���κμ��� ��� �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsDel(frm) {
	if (confirm("������ �����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "delpart";
		frm.submit();
	}
}

function jsRegInsertUpcheShopWitak(frm) {
	if (confirm("�ϰ����� �Ͻðڽ��ϱ�?\n\n������ �ð��� �ҿ�˴ϴ�.(5~10��)") == true) {
		frm.mode.value = "reginsertupcheshopwitak";
		frm.submit();
	}
}

function jsRegInsertPartToOnline(frm) {
	if (confirm("�ϰ����� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "reginsertparttoonline";
		frm.submit();
	}
}

function jsRegInsertPartToOffline(frm) {
	if (confirm("�ϰ����� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "reginsertparttooffline";
		frm.submit();
	}
}

</script>
<table width="100%" cellpadding="5" cellspacing="1" class="a"  style="padding-bottom:50px;" >
<tr>
	<td>
		<table width="100%" align="left" cellpadding="1" cellspacing="1" class="a"   border="0" >
		<form name="frm" method="post" action="innerpart_process.asp">
		<input type="hidden" name="mode" value="<%= mode %>">
		<tr>
			<td>
				<table width="100%" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td height=30><b>���κμ� ���</b></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="30" width="100">
						IDX
					</td>
					<input type="hidden" name="idx" value="<%= oinnerpart.FOneItem.Fidx %>">
					<td bgcolor="#FFFFFF">
						<%= oinnerpart.FOneItem.Fidx %>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="30" width="100">
						����
					</td>
					<td bgcolor="#FFFFFF">
						<select name="divcd">
						<option value="">--����--</option>
						<option value="S" <% if (oinnerpart.FOneItem.Fdivcd = "S") then %>selected<% end if %>>����</option>
						<option value="M" <% if (oinnerpart.FOneItem.Fdivcd = "M") then %>selected<% end if %>>���Ժμ�</option>
						</select>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="30" width="100">
						ERP�μ���
					</td>
					<td bgcolor="#FFFFFF">
						<%= oinnerpart.FOneItem.FBIZSECTION_NM %>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="30" width="100">
						ERP�μ��ڵ�
					</td>
					<td bgcolor="#FFFFFF">
						<input type="text" class="text" name="BIZSECTION_CD" value="<%= oinnerpart.FOneItem.FBIZSECTION_CD %>">
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="30" width="100">
						���κμ��ڵ�
					</td>
					<td bgcolor="#FFFFFF">
						<input type="text" class="text" name="scmid" value="<%= oinnerpart.FOneItem.Fscmid %>">
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFFFFF" colspan="2" align=center height="40">

						<% if (mode = "regnewpart") then %>
						<input type="button" class="button" value="���" onClick="jsReg(frm)">
						<% else %>
						<!--
						<input type="button" class="button" value="����">
						&nbsp;
						-->
						<input type="button" class="button" value="����" onClick="jsDel(frm)">
						<% end if %>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
