<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ް���û
' History : ������ ����
'###########################################################
%>
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
dim useridDb, usernameDb, part_nameDb, posit_nameDb, GetDivCDStrDb, totalvacationdayDb, startdayDb, enddayDb, usedvacationday
dim requesteddayDb, IsAvailableVacationDb, deleteynDb

	page = Request("page")
	masteridx = Request("masteridx")
	part_sn = Request("part_sn")

	if page="" then page=1

	userid = session("ssBctId")

	'// ��å �����̻�. �Ǵ� �ý������� �����ϰ� ���ΰ͸�
	if Not((session("ssAdminPOsn") = "1") or (session("ssAdminPOsn") = "2") or (session("ssAdminPOsn") = "3") or (session("ssAdminPOsn") = "4") or (session("ssAdminPOsn") = "5") or session("ssAdminPsn")=7 or C_ADMIN_AUTH) then
		userid = session("ssBctId")
		part_sn = session("ssAdminPsn")
	end if

	'// �α�������(���)�� ���� �⺻ �μ� ����(������ �̻�:2 �� �ý�����:7 ����)
	'if Not (session("ssAdminLsn")<=2 or session("ssAdminPsn")=7) then
	'	part_sn = session("ssAdminPsn")
	'end if

	dim oVacation
	Set oVacation = new CTenByTenVacation

	oVacation.FRectMasterIdx = masteridx
	'oVacation.FRectpart_sn = part_sn

	'// ��å ��Ʈ���̻�. �Ǵ� �ý������� �����ϰ� ���ΰ͸�
	if Not((session("ssAdminPOsn") = "1") or (session("ssAdminPOsn") = "2") or (session("ssAdminPOsn") = "3") or (session("ssAdminPOsn") = "4") or (session("ssAdminPOsn") = "5") or session("ssAdminPsn")=7 or C_ADMIN_AUTH) then
		oVacation.FRectsearchKey = " t.userid "
		oVacation.FRectsearchString = session("ssBctId")
	end if

	if masteridx<>"" and not(isnull(masteridx)) then
		oVacation.GetMasterOne
	end if

	if masteridx<>"" and not(isnull(masteridx)) then
		oVacation.GetDetailList
	end if

if oVacation.FResultCount>0 then
	useridDb=oVacation.FItemOne.Fuserid
	usernameDb=oVacation.FItemOne.Fusername
	part_nameDb=oVacation.FItemOne.Fpart_name
	posit_nameDb=oVacation.FItemOne.Fposit_name
	GetDivCDStrDb=oVacation.FItemOne.GetDivCDStr
	totalvacationdayDb=oVacation.FItemOne.Ftotalvacationday
	startdayDb=oVacation.FItemOne.Fstartday
	enddayDb=oVacation.FItemOne.Fendday
	usedvacationday=oVacation.FItemOne.Fusedvacationday
	requesteddayDb=oVacation.FItemOne.Frequestedday
	IsAvailableVacationDb=oVacation.FItemOne.IsAvailableVacation
	deleteynDb=oVacation.FItemOne.Fdeleteyn
end if

%>
<!-- �˻� ���� -->
<script language="javascript">
<!--
	function AddItem()
	{
<% if (useridDb = userid) then %>
		window.open("pop_vacation_detail_modify.asp?masteridx=<%= masteridx %>","popAddIem","width=500,height=600,scrollbars=yes");
<% else %>
		alert("�ް��� �ڱ�͸� ��û�� �� �ֽ��ϴ�.");
<% end if %>
	}

	function ViewList(part_sn)
	{
		location.href = "/admin/member/tenbyten/pop_tenbyten_vacation_list_admin.asp?part_sn=" + part_sn;
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




	// ������ �̵�
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}

	//�� ���뺸��
	function jsDetailView(idx){
		var winDetail = window.open("pop_vacation_detail_view.asp?detailidx="+idx,"popDetail","width=500,height=300,scrollbars=yes");
		winDetail.focus();
	}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="part_sn" value="<%= part_sn %>">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">�̸�(���̵�)</td>
		<td align="left">
			<%= usernameDb %>(<%= useridDb %>)
		</td>
		<!--
		<td width="100" bgcolor="<%= adminColor("gray") %>">�μ� / ����</td>
		<td align="left">
			<%= part_nameDb %> / <%= posit_nameDb %>
		</td>
		-->
		<td width="100" bgcolor="<%= adminColor("gray") %>">�μ�</td>
		<td align="left">
			<%= part_nameDb %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">����</td>
		<td align="left">
			<%= GetDivCDStrDb %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">���ϼ�</td>
		<td align="left">
			<%= totalvacationdayDb %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">��밡�ɱⰣ</td>
		<td align="left">
			<%= Left(startdayDb,10) %> - <%= Left(enddayDb,10) %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">����ϼ�</td>
		<td align="left">
			<%= usedvacationday %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">���δ��</td>
		<td align="left">
			<%= requesteddayDb %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">�ܿ��ϼ�</td>
		<td align="left">
			<b><%= (totalvacationdayDb - (usedvacationday + requesteddayDb)) %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">��밡��</td>
		<td align="left">
			<b><%= IsAvailableVacationDb %></b>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">��������</td>
		<td align="left">
			<%= deleteynDb %>
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<!-- ��� �� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%=oVacation.FtotalCount%></b>
			&nbsp;
			������ : <b><%= page %> / <%=oVacation.FtotalPage%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">idx</td>
		<td width="50">����</td>
		<td width="150">�Ⱓ</td>
		<td width="60">��û�ϼ�</td>
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
		<td><%=oVacation.FitemList(i).Fidx%></td>
		<td><%= oVacation.FitemList(i).GetStateDivCDStr %></td>
		<td><%= Left(oVacation.FitemList(i).Fstartday,10) %> - <%= Left(oVacation.FitemList(i).Fendday,10) %></td>
		<td><%= oVacation.FitemList(i).Ftotalday %><% If oVacation.FitemList(i).Ftotalday = "0.5" Then Response.Write CHKIIF(oVacation.FItemList(i).Fhalfgubun="am","[����]","[����]") End If %></td>
		<td><%= oVacation.FitemList(i).Fregistername %></td>
		<td><%= oVacation.FitemList(i).Fapprovername %></td>
		<td>
<%
'// '// ��å �����̻�. �Ǵ� �ý�����
if ((session("ssAdminPOsn") = "1") or (session("ssAdminPOsn") = "2") or (session("ssAdminPOsn") = "3") or (session("ssAdminPOsn") = "4") or (session("ssAdminPOsn") = "5") or session("ssAdminPsn")=7 or C_ADMIN_AUTH) then
%>
			<% if (oVacation.FitemList(i).Fdeleteyn="N") and (oVacation.FitemList(i).Fstatedivcd="R") then %>
			<input type=button value=" �� �� " class="button" onclick="SubmitAllow(<%= masteridx %>, <%=oVacation.FitemList(i).Fidx%>)"> <input type=button value=" �� �� " class="button" onclick="SubmitDeny(<%= masteridx %>, <%=oVacation.FitemList(i).Fidx%>)">
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

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="����Ʈ" onClick="ViewList('<%= part_sn %>');">
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- �׼� �� -->


<form name=frmmodify method=post action="domodifyvacation.asp" onsubmit="return false;">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="masteridx" value="">
	<input type="hidden" name="detailidx" value="">
</form>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
