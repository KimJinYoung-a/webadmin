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
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<!-- #include virtual="/lib/classes/approval/innerOrdercls.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<%

dim idx
idx	= requestCheckvar(Request("idx"),10)

if (idx = "") then
	idx = 0
end if


'==============================================================================
dim oinnerorder
set oinnerorder = New CInnerOrder

oinnerorder.FRectIdx = idx
oinnerorder.GetInnerOrderOne


'==============================================================================
dim mode
if oinnerorder.FOneItem.Fidx = "" then
	mode = "ins"

	oinnerorder.FOneItem.Freguserid = session("ssBctId")
	oinnerorder.FOneItem.Fselluserid = session("ssBctId")
	oinnerorder.FOneItem.Fregdate = date()

else
	mode = "mod"
end if


'==============================================================================
Dim clsBS, arrBizList, intLoop
Set clsBS = new CBizSection
	clsBS.FUSE_YN = "Y"
	clsBS.FOnlySub = "Y"
	arrBizList = clsBS.fnGetBizSectionList
Set clsBS = nothing

%>
<script type="text/javascript" src="/admin/approval/eapp/eapp.js"></script>
<script type="text/javascript">
function jsSetARAP(a, b, acc_cd, acc_nm) {
	frm.acc_cd.value = acc_cd;
	frm.acc_nm.value = acc_nm;
}

function jsReqInnerOrderMannually(frm) {
	if (frm.mode.value != "ins") {
		return;
	}

	if (frm.SELLBIZSECTION_CD.value == "") {
		alert("����μ��� �����ϼ���.");
		return;
	}

	if (frm.BUYBIZSECTION_CD.value == "") {
		alert("���Ժμ��� �����ϼ���.");
		return;
	}

	if (frm.divcd.value == "") {
		alert("������ �����ϼ���.");
		return;
	}

	if (frm.appDate.value == "") {
		alert("�ŷ����ڸ� �Է��ϼ���.");
		return;
	}

	if (frm.appDate.value == "") {
		alert("�ŷ����ڸ� �Է��ϼ���.");
		return;
	}

	if (frm.acc_cd.value == "") {
		alert("���������� �Է��ϼ���.");
		return;
	}

	if (frm.supplySum.value == "") {
		alert("�ŷ����� �Է��ϼ���.");
		return;
	}

	if (frm.supplySum.value*0 != 0) {
		alert("�ŷ����� ���ڸ� �����մϴ�.");
		return;
	}

	if (frm.taxSum.value == "") {
		alert("�ΰ����� �Է��ϼ���.");
		return;
	}

	if (frm.taxSum.value*0 != 0) {
		alert("�ΰ����� ���ڸ� �����մϴ�.");
		return;
	}

	if (frm.totalSum.value == "") {
		alert("�հ踦 �Է��ϼ���.");
		return;
	}

	if (frm.totalSum.value*0 != 0) {
		alert("�հ�� ���ڸ� �����մϴ�.");
		return;
	}

	if ((frm.supplySum.value*1 + frm.taxSum.value*1) != frm.totalSum.value*1) {
		alert("�ŷ��� + �ΰ��� �ݾ��� �հ�� ��ġ���� �ʽ��ϴ�.");
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}


function jsConfirmInnerOrder(frm) {
	if (confirm("Ȯ���Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "confirminnerorder";
		frm.submit();
	}
}

</script>

<table width="100%" cellpadding="5" cellspacing="1" class="a"  style="padding-bottom:50px;" >
<tr>
	<td>
		<form name="frm" method="post" action="innerorder_process.asp">
		<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a"   border="0" >
		<input type="hidden" name="idx" value="<%= idx %>">
		<input type="hidden" name="mode" value="<%= mode %>">
		<tr>
			<td colspan="2">
				<table width="100%" cellpadding="5" cellspacing="1" class="a">
				<tr>
					<td class="verdana-large"><b>���ΰŷ� ���/����</b></td>
					<td align="right"></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td width="50%">
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td height="25" bgcolor="<%= adminColor("tabletop") %>" colspan="2" align="center"><b>�� �� �� ��</b></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center" width="100">��/�μ�</td>
					<td bgcolor="#FFFFFF">
						<% if (oinnerorder.FOneItem.FSELLBIZSECTION_NM = "") then %>
		                    <select class="select" name="SELLBIZSECTION_CD">
		                    <option value="">--����--</option>
		                    <% For intLoop = 0 To UBound(arrBizList,2)	%>
		                		<option value="<%=arrBizList(0,intLoop)%>"><%=arrBizList(1,intLoop)%></option>
		                	<% Next %>
		                    </select>
						<% else %>
							<%= oinnerorder.FOneItem.FSELLBIZSECTION_NM %>
						<% end if %>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">�ۼ���</td>
					<td bgcolor="#FFFFFF">
						<%= oinnerorder.FOneItem.Fselluserid %>
					</td>
				</tr>
				</table>
			</td>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td height="25" bgcolor="<%= adminColor("tabletop") %>" colspan="2" align="center"><b>�� �� �� ��</b></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center" width="100">��/�μ�</td>
					<td bgcolor="#FFFFFF">
						<% if (oinnerorder.FOneItem.FBUYBIZSECTION_NM = "") then %>
		                    <select class="select" name="BUYBIZSECTION_CD">
		                    <option value="">--����--</option>
		                    <% For intLoop = 0 To UBound(arrBizList,2)	%>
		                		<option value="<%=arrBizList(0,intLoop)%>"><%=arrBizList(1,intLoop)%></option>
		                	<% Next %>
		                    </select>
						<% else %>
							<%= oinnerorder.FOneItem.FBUYBIZSECTION_NM %>
						<% end if %>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">Ȯ����</td>
					<td bgcolor="#FFFFFF">
						<%= oinnerorder.FOneItem.Fbuyuserid %>
					</td>
				</tr>
				</table>
			</td>
		</tr>

		<tr>
			<td colspan="2">
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>

				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td width="60" rowspan="6">����</td>
					<td width="20%">����</td>
					<td width="20%">�ŷ�����</td>
					<td width="20%">��������</td>
					<td>���</td>
				</tr>

				<tr align="center"  bgcolor="#FFFFFF">
					<td>
						<% if (oinnerorder.FOneItem.Fdivcd = "") then %>
		                    <select class="select" name="divcd">
			                    <option value="">--����--</option>
		                		<option value="401">��������</option>
		                    </select>
						<% else %>
							<%= oinnerorder.FOneItem.GetDivcdName %>
						<% end if %>
					</td>
					<td>
						<input type="text" name="appDate" value="<%= oinnerorder.FOneItem.FappDate %>" size="10" style="border:0" readonly>
						<% if (oinnerorder.FOneItem.FappDate = "") then %>
							<img src="/images/calicon.gif" align="absmiddle" border="0" onClick="jsPopCal('appDate');"  style="cursor:hand;">
						<% end if %>
					</td>
					<td>
						<input type="hidden" name="acc_cd" value="<%= oinnerorder.FOneItem.Facc_cd %>" size="10" style="border:0" readonly>
						<input type="text" name="acc_nm" value="<%= oinnerorder.FOneItem.Facc_nm %>" size="10" style="border:0" readonly>
					</td>
					<td>
						<% if (oinnerorder.FOneItem.Facc_cd = "") then %>
							<input type="button" class="button" value="�������� ���" onClick="jsGetARAP();">
						<% end if %>
					</td>
				</tr>

				<tr align="center"  bgcolor="<%= adminColor("tabletop") %>">
					<td>�ŷ���</td>
					<td>�ΰ���</td>
					<td>�հ�</td>
					<td>���</td>
				</tr>

				<tr align="center"  bgcolor="#FFFFFF">
					<td>
					<input type="text" class="text" name="supplySum" value="<%= oinnerorder.FOneItem.FsupplySum %>" size="12" <% if oinnerorder.FOneItem.FsupplySum <> "" then %>readonly<% end if %>>
					</td>
					<td>
						<input type="text" class="text" name="taxSum" value="<%= oinnerorder.FOneItem.FtaxSum %>" size="12" <% if oinnerorder.FOneItem.FtaxSum <> "" then %>readonly<% end if %>>
					</td>
					<td>
						<input type="text" class="text" name="totalSum" value="<%= oinnerorder.FOneItem.FtotalSum %>" size="12" <% if oinnerorder.FOneItem.FtotalSum <> "" then %>readonly<% end if %>>
					</td>
					<td>

					</td>
				</tr>

				</table>
			</td>
		</tr>

		<tr>
			<td colspan="2" height="40" align="center">
				<% if mode = "ins" then %>
					<input type="button" value="�ۼ��Ϸ�" class="button"  onClick="jsReqInnerOrderMannually(frm);">
				<% end if %>
				<% if (mode <> "ins") and (Not IsNull(oinnerorder.FOneItem.Fselluserid)) and IsNull(oinnerorder.FOneItem.Fbuyuserid) then %>
					<input type="button" value="Ȯ���ϱ�" class="button"  onClick="jsConfirmInnerOrder(frm);">
				<% end if %>
			</td>
		</tr>

		</table>
		</form>
	</td>
</tr>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
