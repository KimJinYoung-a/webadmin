<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
'
'if (session("ssAdminPsn") = "10") and (session("ssBctId") <> "bseo") and (session("ssBctId") <> "boyishP") then
'	'// CS����� ��û����, 2015-04-08
'	response.write  "������ �����ϴ�. - �ý����� ���� " ''eastone
'	dbget.close() : response.end
'end if

Dim page, SearchKey, SearchString, part_sn, posit_sn, research
Dim deleteyn
dim lp
dim divcd,statediv

dim userid
dim showonlyavail, iPageSize
dim department_id, inc_subdepartment

page 			= Request("page")
deleteyn 		= Request("deleteyn")
SearchKey 		= Request("SearchKey")
SearchString 	= Request("SearchString")
part_sn 		= Request("part_sn")
posit_sn 		= Request("posit_sn")
research 		= Request("research")
divcd 			= Request("divcd")
statediv		= Request("statediv")
showonlyavail	= Request("showonlyavail")
department_id 	= requestCheckvar(Request("department_id"),10)
inc_subdepartment 	= requestCheckvar(Request("inc_subdepartment"),1)
iPageSize 	= requestCheckvar(Request("pagesize"),10)

if (iPageSize = "") then
	iPageSize = 20
end if
if deleteyn="" and research="" then deleteyn="N"
if page="" then page=1

if (SearchKey = "t.userid") then
	userid = SearchString
end if

'// �α�������(���)�� ���� �⺻ �μ� ����(������ �̻�:2 �� �ý�����:7 ����)
'if Not(session("ssAdminLsn")<=2 or session("ssAdminPsn")=7) then
'	part_sn = session("ssAdminPsn")
'end if

dim oVacation
Set oVacation = new CTenByTenVacation

oVacation.FPagesize = iPageSize
oVacation.FCurrPage = page
oVacation.FRectsearchKey = searchKey
oVacation.FRectsearchString = searchString
oVacation.FRectIsDelete = deleteyn
oVacation.FRectpart_sn = part_sn
oVacation.FRectposit_sn = posit_sn
oVacation.FRectDivCd = divcd
oVacation.FRectStateDiv = statediv
oVacation.FRectShowOnlyAvail = showonlyavail
oVacation.Fdepartment_id 		= department_id
oVacation.Finc_subdepartment 	= inc_subdepartment

oVacation.GetMasterList

%>
<!-- �˻� ���� -->
<script language="javascript">

function ViewDetail(masteridx)
{
	var pop = window.open("/admin/member/tenbyten/tenbyten_vacation_detail_list.asp?masteridx=" + masteridx,"ViewDetail","width=900,height=600,scrollbars=yes");
	pop.focus();
}

function AddItem(userid)
{
	window.open("pop_vacation_modify.asp?userid=" + userid,"popAddIem","width=500,height=600,scrollbars=yes");
}

function AddYearVacationItem(userid)
{
	window.open("pop_vacation_modify.asp?userid=" + userid + "&isyearvacation=Y","popAddIem","width=500,height=600,scrollbars=yes");
}

function AddAllYearVacation(insDivcode)
{
	var strMsg = "";
	if (insDivcode == "R"){
		strMsg = "���⵵ ������ �����˴ϴ�.";
	}

	if (confirm(strMsg + "��ü ���� �ް��� ����Ͻðڽ��ϱ�?") == true) {
		document.frmupdate.mode.value="addallyearvacationNew";
		document.frmupdate.insDivcode.value = insDivcode;

		document.frmupdate.submit();
	}
}

function AddAllLongYearVacation()
{
	if (confirm("��ü ���ټ� �ް��� ����Ͻðڽ��ϱ�?") == true) {
		document.frmupdate.mode.value="addalllongmonthvacation";
		document.frmupdate.submit();
	}
}


function AddReCalVacation()
{
	if (confirm("�ñް���� ������ ������  ������ ����Ͻðڽ��ϱ�?") == true) {
		document.frmupdate.mode.value="addrecalvacation";
		document.frmupdate.submit();
	}
}
// ������ �̵�
function goPage(pg)
{
	document.frm.page.value=pg;
	document.frm.submit();
}

function jsPartList(empno){
	var winPL = window.open("tenbyten_vacation_part_list.asp?empno=" + empno,"popPL","width=800,height=600,scrollbars=yes");
	winPL.focus();
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
			�μ�NEW:
			<%= drawSelectBoxDepartment("department_id", department_id) %>
			<input type="checkbox" name="inc_subdepartment" value="N" <% if (inc_subdepartment = "N") then %>checked<% end if %> > ���� �μ����� ����
			&nbsp;
			��������:
			<select name="deleteyn" class="select">
				<option value="">��ü</option>
				<option value="N">���</option>
				<option value="Y">����</option>
			</select>
			&nbsp;
			��������:
			<select name="statediv" class="select">
				<option value="">��ü</option>
				<option value="Y">����</option>
				<option value="N">���</option>
			</select>
			&nbsp;
			<% if C_ADMIN_AUTH or C_PSMngPart then %>
			����:
			<%=printPositOptionIN90("posit_sn", posit_sn)%>&nbsp;
			&nbsp;
			<% end if %>
			�ް����� :
			<select name=divcd class="select">
				<option value="">��ü</option>
				<option value="1" <% if (divcd = "1") then %>selected<% end if %>>����</option>
				<!--
				<option value="2">����</option>
				-->
				<option value="3" <% if (divcd = "3") then %>selected<% end if %>>����</option>
				<option value="4" <% if (divcd = "4") then %>selected<% end if %>>����</option>
				<option value="6" <% if (divcd = "6") then %>selected<% end if %>>������</option>
				<option value="5" <% if (divcd = "5") then %>selected<% end if %>>���</option>
				<option value="7" <% if (divcd = "7") then %>selected<% end if %>>���ϴ�ü</option>
				<option value="8" <% if (divcd = "8") then %>selected<% end if %>>��Ÿ�ް�</option>
				<option value="9" <% if (divcd = "9") then %>selected<% end if %>>�����ް�</option>
				<option value="A" <% if (divcd = "A") then %>selected<% end if %>>�����ް�</option>
			</select>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			�˻�:
			<select name="SearchKey" class="select">
				<option value="">::����::</option>
				<option value="t.userid">���̵�</option>
				<option value="t.username">����ڸ�</option>
				<option value="t.empno">���</option>
			</select>
			<script language="javascript">
				document.frm.deleteyn.value="<%= deleteyn %>";
				document.frm.SearchKey.value="<%= SearchKey %>";
				document.frm.statediv.value="<%= statediv %>";
			</script>
			<input type="text" class="text" name="SearchString" size="20" value="<%=SearchString%>">	&nbsp;
			<input type="checkbox" name="showonlyavail" value="Y" <% if (showonlyavail = "Y") then %>checked<% end if %> >
			��밡�� �ް���
			&nbsp;
			ǥ�ð���:
			<select class="select" name="pagesize">
				<option value="20" <% if (iPageSize = "20") then %>selected<% end if %> >20 ��</option>
				<option value="50" <% if (iPageSize = "50") then %>selected<% end if %> >50 ��</option>
				<option value="100" <% if (iPageSize = "100") then %>selected<% end if %> >100 ��</option>
				<option value="500" <% if (iPageSize = "500") then %>selected<% end if %> >500 ��</option>
			</select>
		</td>


	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
<%
'// �α�������(���)�� ���� �⺻ �μ� ����(������ �̻�:1 �� �ý�����:7 �λ��ѹ���:20 ����)
if (session("ssAdminLsn")<=1 or session("ssAdminPsn")=7 or C_PSMngPart or C_ADMIN_AUTH) then
%>
			[
			������ :
			<input type="button" class="button" value="�ް� ���" onClick="javascript:AddItem('<%= userid %>');">
			<input type="button" class="button" value="���� ���" onClick="javascript:AddYearVacationItem('<%= userid %>');">
			&nbsp;
			<input type="button" class="button" value="��ü���� ���(������,��1ȸ)" onClick="javascript:AddAllYearVacation('R');">
			<input type="button" class="button" value="��ü���� ���(�����,��1ȸ)" onClick="javascript:AddAllYearVacation('P');">
			<input type="button" class="button" value="��ü���ټ� ���(������,��1ȸ)" onClick="javascript:AddAllLongYearVacation();">
			&nbsp;
			<input type="button" class="button" value="�ñ�/���ް���� ������ ����������" onClick="javascript:AddReCalVacation();">
			]
<% end if %>
<%
'// �ý�����:7, 30
if (session("ssAdminPsn") = 7) or (session("ssAdminPsn") = 30) then
%>
		<!--	[
			�ý����� ���� :
			<input type="button" class="button" value="��ü���� ���" onClick="javascript:AddAllYearVacation('');">
			<input type="button" class="button" value="���ټ� ���" onClick="javascript:AddAllLongYearVacation('');">
			]-->
<% end if %>

		</td>
		<td align="right">
			<!-- <img src="/images/icon_excel.gif" border="0"> -->
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<!-- ��� �� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="21">
			�˻���� : <b><%=oVacation.FtotalCount%></b>
			&nbsp;
			������ : <b><%= page %> / <%=oVacation.FtotalPage%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td>idx</td>
		<td>����</td>
    	<td>���</td>
		<td width="40">�̸�</td>
		<td width="70">�Ի���<br>(������)</td>
		<td width="70">����<br>�Ի���</td>
		<td width="70">���<br>(����)��</td>
		<td>�μ�</td>
		<% if C_ADMIN_AUTH or C_PSMngPart then %><td>����</td><% end if %>
		<td>��å</td>
		<td width="80">��밡�ɱⰣ</td>
		<td width="40">��<br>�ϼ�</td>
		<td width="40">���<br>�ϼ�</td>
		<td width="40">����<br>���</td>

		<td width="50">����<br>�ϼ�</td>
		<td width="50">����<br>�ϼ�</td>
		<td width="50">���<br>����</td>

		<td width="50">�ܿ�<br>�ϼ�</td>
		<td width="30">���<br>����</td>
		<td width="30">����<br>����</td>
		<td>�����</td>
    </tr>
	<% if oVacation.FResultCount=0 then %>
	<tr height="25">
		<td colspan="21" align="center" bgcolor="#FFFFFF">���(�˻�)�� ������ �����ϴ�.</td>
	</tr>
	<% else %>

	<% for lp=0 to oVacation.FResultCount - 1 %>
	<tr align="center" bgcolor="<% if (oVacation.FitemList(lp).Fdeleteyn="N") and (oVacation.FitemList(lp).IsAvailableVacation="Y") then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>" height="30">
		<td><a href="javascript:ViewDetail(<%=oVacation.FitemList(lp).Fidx%>)"><%= oVacation.FitemList(lp).Fidx %></a></td>
		<td nowrap><%= oVacation.FitemList(lp).GetDivCDStr %></td>
		<td nowrap><a href="javascript:ViewDetail(<%=oVacation.FitemList(lp).Fidx%>)"><%=oVacation.FitemList(lp).Fempno%></a></td>
		<td nowrap><%= oVacation.FitemList(lp).Fusername %></td>

		<td nowrap><%= Left(oVacation.FitemList(lp).Fjoinday, 10) %></td>
		<td nowrap>
			<% if Not IsNull(oVacation.FitemList(lp).Frealjoinday) then %>
				<% if (oVacation.FitemList(lp).Fjoinday <> oVacation.FitemList(lp).Frealjoinday) then %>
					<font color="red"><%= oVacation.FitemList(lp).Frealjoinday %></font>
				<% else %>
					<%= oVacation.FitemList(lp).Frealjoinday %>
				<% end if %>
			<% end if %>
		</td>
		<td nowrap>
			<% if Not IsNull(oVacation.FitemList(lp).Fretireday) then %>
				<%= oVacation.FitemList(lp).Fretireday %>
			<% end if %>
		</td>

		<td><%= oVacation.FitemList(lp).FdepartmentNameFull %></td>
		<% if C_ADMIN_AUTH or C_PSMngPart then %><td><%= oVacation.FitemList(lp).Fposit_name %></td><% end if %>
		<td><%= oVacation.FitemList(lp).Fjob_name %></td>
		<td><%= Left(oVacation.FitemList(lp).Fstartday,10) %> ~ <%= Left(oVacation.FitemList(lp).Fendday,10) %></td>
		<td>
			<a href="javascript:jsPartList('<%=oVacation.FitemList(lp).Fempno%>');"><%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).Ftotalvacationday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %></a>
		</td>
		<td>
			<%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).Fusedvacationday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %>
		</td>
		<td>
			<b>
			<%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).Frequestedday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %>
			</b>
		</td>

		<td>
			<% if (oVacation.FitemList(lp).Fdivcd = "1") or (oVacation.FitemList(lp).Fdivcd = "7") then %>
				<%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).FpromotionDay) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %>
			<% end if %>
		</td>
		<td>
			<% if (oVacation.FitemList(lp).Fdivcd = "1") or (oVacation.FitemList(lp).Fdivcd = "7") then %>
				<%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).FjungsanDay) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %>
			<% end if %>
		</td>
		<td>
			<% if (oVacation.FitemList(lp).Fdivcd = "1") or (oVacation.FitemList(lp).Fdivcd = "7") then %>
				<%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).FretireJungsanDay) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %>
			<% end if %>
		</td>

		<td>
			<%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).GetRemainVacationDay) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %>
		</td>
		<td><%= oVacation.FitemList(lp).IsAvailableVacation %></td>
		<td><%= oVacation.FitemList(lp).Fdeleteyn %></td>
		<td><%= oVacation.FitemList(lp).Fregisterid %></td>
	</tr>
	<% next %>

	<% end if %>
<!-- ���� ��� �� -->

<!-- ������ ���� -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="21" align="center">
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
<form name="frmupdate" method="post" action="domodifyvacation.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="insDivcode" value="">
</form>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
