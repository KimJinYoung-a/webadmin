<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<%

Dim page, masteridx
dim i
dim part_sn
dim userid

page = Request("page")
masteridx = Request("masteridx")

if page="" then page=1

userid = session("ssBctId")

'// �α�������(���)�� ���� �⺻ �μ� ����(������ �̻�:2 �� �ý�����:7 ����)
'if Not (session("ssAdminLsn")<=2 or session("ssAdminPsn")=7) then
'	part_sn = session("ssAdminPsn")
'end if

dim oVacation
Set oVacation = new CTenByTenVacation

oVacation.FRectMasterIdx = masteridx
oVacation.FRectpart_sn = part_sn

oVacation.GetMasterOne

oVacation.FPageSize = 40

oVacation.GetDetailList

%>
<!-- �˻� ���� -->
<script language="javascript">
<!--

function AddItem()
{
	window.open("pop_vacation_detail_modify.asp?masteridx=<%= masteridx %>","popAddIem","width=500,height=600,scrollbars=yes");
}

function ViewList()
{
	location.href = "/admin/member/tenbyten/tenbyten_vacation_list.asp";
}

function ViewCalendar()
{
	window.open("/admin/member/tenbyten/pop_vacation_calendar.asp","popAddIem","width=800,height=650,scrollbars=yes");
}

function SubmitAllow(masteridx, detailidx)
{
	var frm = document.frmmodify;

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "allowdetail";
		frm.masteridx.value = masteridx;
		frm.detailidx.value = detailidx;
		frm.comment.value = document.frmDetail.comment.value;

		frm.submit();
	}
}

function SubmitDeleteMaster(masteridx)
{
	var frm = document.frmmodify;

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "deletemaster";
		frm.masteridx.value = masteridx;

		frm.submit();
	}
}

function SubmitDeny(masteridx, detailidx)
{
	var frm = document.frmmodify;

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "denydetail";
		frm.masteridx.value = masteridx;
		frm.detailidx.value = detailidx;
		frm.comment.value = document.frmDetail.comment.value;

		frm.submit();
	}
}

function SubmitDelete(masteridx, detailidx)
{
	var frm = document.frmmodify;

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "deletedetail";
		frm.masteridx.value = masteridx;
		frm.detailidx.value = detailidx;

		frm.submit();
	}
}

function SubmitModify(masteridx)
{
	var frmmaster = document.frm;
	var frm = document.frmmodify;

	if (jsCheckIsDigit(frmmaster.totalvacationday) != true) {
		alert("���ڸ� �Է°����մϴ�.");
		frmmaster.totalvacationday.focus();
		return;
	}

	if (jsCheckIsDigit(frmmaster.promotionDay) != true) {
		alert("���ڸ� �Է°����մϴ�.");
		frmmaster.promotionDay.focus();
		return;
	}

	if (jsCheckIsDigit(frmmaster.jungsanDay) != true) {
		alert("���ڸ� �Է°����մϴ�.");
		frmmaster.jungsanDay.focus();
		return;
	}

	if (jsCheckIsDigit(frmmaster.retireJungsanDay) != true) {
		alert("���ڸ� �Է°����մϴ�.");
		frmmaster.retireJungsanDay.focus();
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "modifymaster";
		frm.divcd.value = frmmaster.divcd.value;

		<% if (oVacation.FItemOne.Fposit_sn = 13) then %>
			// �ñް������ �ð��� ���ڷ� �������ش�.
			// 1���� 8�ð�, �ѽð��� 0.125(= 1/8)
			frmmaster.totalvacationday.value = frmmaster.totalvacationday.value * 1.0 * 0.125;
			frmmaster.promotionDay.value = frmmaster.promotionDay.value * 1.0 * 0.125;
			frmmaster.jungsanDay.value = frmmaster.jungsanDay.value * 1.0 * 0.125;
			frmmaster.retireJungsanDay.value = frmmaster.retireJungsanDay.value * 1.0 * 0.125;
		<% end if %>

		frm.totalvacationday.value = frmmaster.totalvacationday.value;
		frm.promotionDay.value = frmmaster.promotionDay.value;
		frm.jungsanDay.value = frmmaster.jungsanDay.value;
		frm.retireJungsanDay.value = frmmaster.retireJungsanDay.value;
		frm.startday.value = frmmaster.startday.value;
		frm.endday.value = frmmaster.endday.value;
		frm.comment.value = frmmaster.comment.value;

		frm.masteridx.value = masteridx;

		frm.submit();
	}
}

function jsCheckIsDigit(obj) {
	if ((obj.value == "") || (obj.value*0 != 0)) {
		return false;
	}

	return true;
}

// ������ �̵�
function goPage(pg)
{
	document.frm.page.value=pg;
	document.frm.submit();
}

function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

//-->
</script>
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">�̸�(���̵�)</td>
		<td align="left">
			<%= oVacation.FItemOne.Fusername %>(<% if (oVacation.FItemOne.Fuserid = "") then response.write oVacation.FItemOne.Fempno else response.write oVacation.FItemOne.Fuserid end if %>)
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">�μ�</td>
		<td align="left">
			<%= oVacation.FItemOne.Fpart_name %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">����</td>
		<td align="left">
			<select class="select" name=divcd>
				<option value="1" <% if (oVacation.FItemOne.Fdivcd = "1") then %>selected<% end if %> >����</option>
				<!--
				<option value="2" <% if (oVacation.FItemOne.Fdivcd = "2") then %>selected<% end if %> >����</option>
				-->
				<option value="3" <% if (oVacation.FItemOne.Fdivcd = "3") then %>selected<% end if %> >����</option>
				<option value="4" <% if (oVacation.FItemOne.Fdivcd = "4") then %>selected<% end if %> >����</option>
				<option value="6" <% if (oVacation.FItemOne.Fdivcd = "6") then %>selected<% end if %> >������</option>
				<option value="5" <% if (oVacation.FItemOne.Fdivcd = "5") then %>selected<% end if %> >���</option>
				<option value="7" <% if (oVacation.FItemOne.Fdivcd = "7") then %>selected<% end if %> >���ϴ�ü</option>
				<option value="8" <% if (oVacation.FItemOne.Fdivcd = "8") then %>selected<% end if %> >��Ÿ�ް�</option>
				<option value="9" <% if (oVacation.FItemOne.Fdivcd = "9") then %>selected<% end if %> >�����ް�</option>
				<option value="9" <% if (oVacation.FItemOne.Fdivcd = "A") then %>selected<% end if %> >�����ް�</option>
			</select>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">���ϼ�</td>
		<td align="left">
			<% if (oVacation.FItemOne.Fposit_sn <> 13) then %>
				<input type="text" class="text" name="totalvacationday" size="2" value="<%= oVacation.FItemOne.Ftotalvacationday %>">
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			<% else %>
				<input type="text" class="text" name="totalvacationday" size="2" value="<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FItemOne.Ftotalvacationday) %>">
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
				(�ñް����)
			<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">��밡�ɱⰣ</td>
		<td align="left">
    		<input type="text" name="startday" class="text_ro" size="11" maxlength="10" value="<%= Left(oVacation.FItemOne.Fstartday,10) %>" onClick="jsPopCal('frm','startday');" style="cursor:hand;">
    		-
    		<input type="text" name="endday" class="text_ro" size="11" maxlength="10" value="<%= Left(oVacation.FItemOne.Fendday,10) %>" onClick="jsPopCal('frm','endday');" style="cursor:hand;">
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">����ϼ�</td>
		<td align="left">
			<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FItemOne.Fusedvacationday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">���δ��</td>
		<td align="left">
			<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FItemOne.Frequestedday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">�����ϼ�</td>
		<td align="left">
			<% if (oVacation.FItemOne.Fposit_sn <> 13) then %>
				<input type="text" class="text<% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>_ro<% end if %>" name="promotionDay" size="2" value="<%= oVacation.FItemOne.FpromotionDay %>" <% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>readonly<% end if %> >
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			<% else %>
				<input type="text" class="text<% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>_ro<% end if %>" name="promotionDay" size="2" value="<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FItemOne.FpromotionDay) %>" <% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>readonly<% end if %> >
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">�����ϼ�</td>
		<td align="left">
			<% if (oVacation.FItemOne.Fposit_sn <> 13) then %>
				<input type="text" class="text<% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>_ro<% end if %>" name="jungsanDay" size="2" value="<%= oVacation.FItemOne.FjungsanDay %>" <% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>readonly<% end if %> >
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			<% else %>
				<input type="text" class="text<% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>_ro<% end if %>" name="jungsanDay" size="2" value="<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FItemOne.FjungsanDay) %>" <% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>readonly<% end if %> >
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			<% end if %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">û���ϼ�</td>
		<td align="left">
			<% if (oVacation.FItemOne.Fposit_sn <> 13) then %>
				<input type="text" class="text<% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>_ro<% end if %>" name="retireJungsanDay" size="2" value="<%= oVacation.FItemOne.FretireJungsanDay %>" <% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>readonly<% end if %> >
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			<% else %>
				<input type="text" class="text<% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>_ro<% end if %>" name="retireJungsanDay" size="2" value="<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FItemOne.FretireJungsanDay) %>" <% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>readonly<% end if %> >
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">�ܿ��ϼ�</td>
		<td align="left">
			<b> 
			<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, (oVacation.FItemOne.GetRemainVacationDay)) %> <%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			</b>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">������ �ڸ�Ʈ</td>
		<td align="left">
			<input type="text" name="comment" value="<%=replace(oVacation.FItemOne.Fcomment ,"""","&quot;")%>" class="text" style="width:96%;" />
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">��밡��</td>
		<td align="left">
			<b><%= oVacation.FItemOne.IsAvailableVacation %></b>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">��������</td>
		<td align="left">
			<%= oVacation.FItemOne.Fdeleteyn %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="40" width="100" bgcolor="<%= adminColor("gray") %>">�����ڱ��</td>
		<td align="center" colspan="3">
			<%
			'// �α�������(���)�� ���� �⺻ �μ� ����(��Ʈ���� �̻�:3 �� �ý�����:7 �濵������:8 ����)
			if (session("ssAdminLsn")<=3 or session("ssAdminPsn")=7 or session("ssAdminPsn")=8) or C_PSMngPart or C_ADMIN_AUTH then
			%>
				<input type="button" class="button" value="�����ϱ�" onClick="javascript:SubmitModify(<%= masteridx %>);">
			<% end if %>
		</td>
	</tr>
</table>
</form>
<!-- �˻� �� -->

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
<%
'// �α�������(���)�� ���� �⺻ �μ� ����(��Ʈ���� �̻�:3 �� �ý�����:7 �濵������:8 ����)
if (session("ssAdminLsn")<=3 or session("ssAdminPsn")=7 or session("ssAdminPsn")=8 or C_PSMngPart or C_ADMIN_AUTH) then
%>
			<input type="button" class="button" value="�������ް��������" onClick="javascript:AddItem('');">
<% end if %>
			<input type="button" class="button" value="�ް��޷º���" onClick="javascript:ViewCalendar('');">
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<!-- ��� �� ���� -->
<form name="frmDetail" method="get" action="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%=oVacation.FtotalCount%></b>
			&nbsp;
			������ : <b><%= page %> / <%=oVacation.FtotalPage%></b>
		</td>
	</tr>
	<tr height=30 align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">idx</td>
		<td width="50">����</td>
		<td width="150">�Ⱓ</td>
		<td width="60">��û�ϼ�</td>
		<td width="60">��������</td>
		<td width="100">�����</td>
		<td width="100">ó����</td>
		<td>���</td>
    </tr>
	<% if oVacation.FResultCount=0 then %>
	<tr height=30>
		<td colspan="15" align="center" bgcolor="#FFFFFF">���(�˻�)�� ������ �����ϴ�.</td>
	</tr>
	<% else %>
		<% for i=0 to oVacation.FResultCount - 1 %>
	<tr height=30 align="center" bgcolor="<% if (oVacation.FitemList(i).Fdeleteyn="N") then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">
		<td><a href="javascript:ModiItem('<%=oVacation.FitemList(i).Fidx%>')"><%=oVacation.FitemList(i).Fidx%></a></td>
		<td><%= oVacation.FitemList(i).GetStateDivCDStr %></td>
		<td><%= Left(oVacation.FitemList(i).Fstartday,10) %> - <%= Left(oVacation.FitemList(i).Fendday,10) %></td>
		<td>
			<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FitemList(i).Ftotalday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
		</td>
		<td><%= oVacation.FitemList(i).Fdeleteyn %></td>
		<td><%= oVacation.FitemList(i).Fregistername %></td>
		<td><%= oVacation.FitemList(i).Fapprovername %></td>
		<td>
<%
'// �α�������(���)�� ���� �⺻ �μ� ����(��Ʈ���� �̻�:3 �� �ý�����:7 �濵������:8 ����)
if (session("ssAdminLsn")<=3 or session("ssAdminPsn")=7 or session("ssAdminPsn")=8 or C_PSMngPart or C_ADMIN_AUTH) then
%>
			<% if (oVacation.FitemList(i).Fdeleteyn="N") and (oVacation.FitemList(i).Fstatedivcd="R") then %>
			�ڸ�Ʈ: <input type="text" name="comment" class="text" value="" style="width:70%" /><br />
			<input type=button value=" �� �� " class="button" onclick="SubmitAllow(<%= masteridx %>, <%=oVacation.FitemList(i).Fidx%>)"> <input type=button value=" �� �� " class="button" onclick="SubmitDeny(<%= masteridx %>, <%=oVacation.FitemList(i).Fidx%>)">
			<% elseif oVacation.FitemList(i).Fcomment<>"" then %>
			�ڸ�Ʈ: <%=oVacation.FitemList(i).Fcomment%><br />
			<% end if %>
			<% if (oVacation.FitemList(i).Fdeleteyn="N") then %>
			<input type=button value=" �� �� " class="button" onclick="SubmitDelete(<%= masteridx %>, <%=oVacation.FitemList(i).Fidx%>)">
			<% end if %>
<% end if %>
		</td>
	</tr>
		<% next %>

	<% end if %>
<!-- ���� ��� �� -->

<!-- ������ ���� -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<%
				if oVacation.HasPreScroll then
					Response.Write "<a href='javascript:goPage(" & oVacation.StartScrollPage-1 & ")'>[pre]</a>"
				else
					Response.Write "[pre]"
				end if

				for i=0 + oVacation.StartScrollPage to oVacation.FScrollCount + oVacation.StartScrollPage - 1

					if i>oVacation.FTotalpage then Exit for

					if CStr(page)=CStr(i) then
						Response.Write " <font color='red'>[" & i & "]</font> "
					else
						Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
					end if

				next

				if oVacation.HasNextScroll then
					Response.Write "<a href='javascript:goPage(" & i & ")'>[next]</a>"
				else
					Response.Write "[next]"
				end if
			%>
		</td>
	</tr>
</table>
</form>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value=" ����Ʈ " onClick="ViewList();">
<%
'// �α�������(���)�� ���� �⺻ �μ� ����(��Ʈ���� �̻�:3 �� �ý�����:7 �濵������:8 ����)
if (session("ssAdminLsn")<=3 or session("ssAdminPsn")=7 or session("ssAdminPsn")=8 or C_PSMngPart or C_ADMIN_AUTH) then
%>
			<input type="button" class="button" value=" �ް����� " onClick="SubmitDeleteMaster(<%= masteridx %>);">
<% end if %>
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- �׼� �� -->

<form name="frmmodify" method="post" action="domodifyvacation.asp" onsubmit="return false;">
	<input type="hidden" name="mode" value="" />
	<input type="hidden" name="masteridx" value="" />
	<input type="hidden" name="detailidx" value="" />
	<input type="hidden" name="divcd" value="" />
	<input type="hidden" name="totalvacationday" value="" />
	<input type="hidden" name="promotionDay" value="" />
	<input type="hidden" name="jungsanDay" value="" />
	<input type="hidden" name="retireJungsanDay" value="" />
	<input type="hidden" name="startday" value="" />
	<input type="hidden" name="endday" value="" />
	<input type="hidden" name="comment" value="" />
</form>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
