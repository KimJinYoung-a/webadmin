<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/tenmember/lib/header.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<%

dim empno
dim login_username

Dim page, SearchKey, SearchString, part_sn, research
Dim deleteyn
dim lp

empno = session("ssBctSn")
login_username	= session("ssBctCname")

page = Request("page")
deleteyn = Request("deleteyn")
'SearchKey = Request("SearchKey")
'SearchString = Request("SearchString")
'part_sn = Request("part_sn")
research = Request("research")
deleteyn="N"
'if deleteyn="" and research="" then deleteyn="N"
if page="" then page=1

'// �μ� ����
part_sn = session("ssAdminPsn")

dim oVacation
Set oVacation = new CTenByTenVacation

oVacation.FPagesize = 20
oVacation.FCurrPage = page
oVacation.FRectsearchKey = " t.empno "
oVacation.FRectsearchString = empno
oVacation.FRectIsDelete = "N"
oVacation.FRectpart_sn = part_sn

oVacation.GetMasterList

%>
<!-- �˻� ���� -->
<script language="javascript">

function ViewDetail(masteridx)
{
	location.href = "pop_tenbyten_vacation_detail_list.asp?masteridx=" + masteridx;
}

// ������ �̵�
function goPage(pg)
{
	document.frm.page.value=pg;
	document.frm.submit();
}

</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�̸� : <%= login_username %>(<%= empno %>)
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
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
	<tr height=30 align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">����</td>
		<td>��밡�ɱⰣ</td>
		<td width="60">���ϼ�</td>
		<td width="60">����ϼ�</td>
		<td width="60">���δ��</td>
		<td width="60">�ܿ��ϼ�</td>
		<td width="50">��밡��</td>
		<td width="50">�����ϼ�</td>
		<td width="100">�ް���û</td>
    </tr>
	<% if oVacation.FResultCount=0 then %>
	<tr height=30>
		<td colspan="15" align="center" bgcolor="#FFFFFF">���(�˻�)�� ������ �����ϴ�.</td>
	</tr>
	<% else %>

	<% for lp=0 to oVacation.FResultCount - 1 %>
	<tr height=30 align="center" bgcolor="<% if (oVacation.FitemList(lp).Fdeleteyn="N") and (oVacation.FitemList(lp).IsAvailableVacation="Y") then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">
		<td><%= oVacation.FitemList(lp).GetDivCDStr %></td>
		<td><a href="javascript:ViewDetail(<%=oVacation.FitemList(lp).Fidx%>)"><%= Left(oVacation.FitemList(lp).Fstartday,10) %>-<%= Left(oVacation.FitemList(lp).Fendday,10) %></a></td>
		<td>
			<%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).Ftotalvacationday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %>
		</td>
		<td>
			<%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).Fusedvacationday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %>
		</td>
		<td>
			<%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).Frequestedday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %>
		</td>
		<td>
			<% if (oVacation.FitemList(lp).IsAvailableVacation = "Y") then %>
			<b><%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, (oVacation.FitemList(lp).Ftotalvacationday - (oVacation.FitemList(lp).Fusedvacationday + oVacation.FitemList(lp).Frequestedday))) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %></b>
			<% else %>
			<b>0</b>
			<% end if %>


		</td>
		<td><%= oVacation.FitemList(lp).IsAvailableVacation %></td>
		<td>
			<% if (oVacation.FitemList(lp).IsAvailableVacation = "Y") then %>
			0
			<% else %>
			<%= (oVacation.FitemList(lp).Ftotalvacationday - oVacation.FitemList(lp).Fusedvacationday) %>
			<% end if %>
		</td>
		<td>
			<% if (oVacation.FitemList(lp).IsAvailableVacation = "Y") then %>
			<b><input type="button" class="button" value="�ް���û" onclick="ViewDetail(<%=oVacation.FitemList(lp).Fidx%>)"></b>
			<% else %>
			<b>&nbsp;</b>
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

				for lp=0 + oVacation.StartScrollPage to oVacation.FScrollCount + oVacation.StartScrollPage - 1

					if lp>oVacation.FTotalpage then Exit for

					if CStr(page)=CStr(lp) then
						Response.Write " <font color='red'>[" & lp & "]</font> "
					else
						Response.Write " <a href='javascript:goPage(" & lp & ")'>[" & lp & "]</a> "
					end if

				next

				if oVacation.HasNextScroll then
					Response.Write "<a href='javascript:goPage(" & lp & ")'>[next]</a>"
				else
					Response.Write "[next]"
				end if
			%>
		</td>
	</tr>
</table>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
