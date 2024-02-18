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
empno = session("ssBctSn")

Dim page, masteridx
dim i
dim part_sn
dim vTmpIdx, vTmpUserid

page = Request("page")
masteridx = Request("masteridx")

if page="" then page=1


'// ============================================================================
dim oVacation
Set oVacation = new CTenByTenVacation

oVacation.FRectMasterIdx = masteridx
oVacation.FRectpart_sn = part_sn
oVacation.FRectsearchKey = " t.empno "
oVacation.FRectsearchString = empno
oVacation.FRectIsDelete = "N"

oVacation.GetMasterOne

oVacation.FPageSize = 15
oVacation.FCurrPage = page
oVacation.GetDetailList

%>
<script language="javascript">

function AddItem() {
	window.open("pop_vacation_detail_modify.asp?masteridx=<%= masteridx %>","popAddIem","width=500,height=600,scrollbars=yes");
}

function ViewList(masteridx)
{
	location.href = "pop_tenbyten_vacation_list.asp?masteridx=" + masteridx;
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

//���ڰ��� ǰ�Ǽ� ���
function jsRegEapp(scmidx, adate, aday){
	<% if (Not C_IS_SCM_LOGIN) then %>
		alert("���̵�� �α����� ���Ŀ��� ǰ�Ǽ��� �ۼ��� �� �ֽ��ϴ�.");
		return;
	<% end if %>
	
	document.frmEapp.iSL.value = scmidx; 		

	document.frmEapp.dDate.value = adate; 
	document.frmEapp.dDay.value = aday; 
	document.frmEapp.target = "popE";
	var winEapp = window.open("","popE","width=1000,height=600,scrollbars=yes");
	document.frmEapp.submit();
	winEapp.focus();
}

//���ڰ��� ǰ�Ǽ� ���뺸��
function jsViewEapp(reportidx,reportstate){
	var winEapp = window.open("/admin/approval/eapp/popIndex.asp?iRM=M01"+reportstate+"&iridx="+reportidx,"popE","");
	winEapp.focus();
}

//�� ���뺸��
function jsDetailView(idx){
	var winDetail = window.open("/admin/member/tenbyten/pop_vacation_detail_view.asp?detailidx="+idx,"popDetail","width=500,height=300,scrollbars=yes");
	winDetail.focus();
}

</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="masteridx" value="<%=masteridx%>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">�̸�(���)</td>
		<td align="left">
			<%= oVacation.FItemOne.Fusername %>(<%= oVacation.FItemOne.Fempno %>)
		</td>
		<!--
		<td width="100" bgcolor="<%= adminColor("gray") %>">�μ� / ����</td>
		<td align="left">
			<%= oVacation.FItemOne.Fpart_name %> / <%= oVacation.FItemOne.Fposit_name %>
		</td>
		-->
		<td width="100" bgcolor="<%= adminColor("gray") %>">�μ�</td>
		<td align="left">
			<%= oVacation.FItemOne.Fpart_name %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">����</td>
		<td align="left">
			<%= oVacation.FItemOne.GetDivCDStr %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">���ϼ�</td>
		<td align="left">
			<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FItemOne.Ftotalvacationday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">��밡�ɱⰣ</td>
		<td align="left">
			<%= Left(oVacation.FItemOne.Fstartday,10) %> - <%= Left(oVacation.FItemOne.Fendday,10) %>
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
		<td width="100" bgcolor="<%= adminColor("gray") %>">�ܿ��ϼ�</td>
		<td align="left">
			<b><%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, (oVacation.FItemOne.Ftotalvacationday - (oVacation.FItemOne.Fusedvacationday + oVacation.FItemOne.Frequestedday))) %> <%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %></b>
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
	</form>
</table>
<!-- �˻� �� -->

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="�ް���û" onClick="javascript:AddItem('');">
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- �׼� �� -->

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
		<td width="80">�����</td>
		<td width="80">ó����</td>
		<td>���</td>
    </tr>
	<% if oVacation.FResultCount=0 then %>
	<tr height=30>
		<td colspan="15" align="center" bgcolor="#FFFFFF">���(�˻�)�� ������ �����ϴ�.</td>
	</tr>
	<% else %>
		<% for i=0 to oVacation.FResultCount - 1 %>
	<tr height=30 align="center" bgcolor="<% if (oVacation.FitemList(i).Fdeleteyn="N") then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">
		<td><a href="#" onclick="jsDetailView('<%=oVacation.FitemList(i).Fidx%>');return false;"><%=oVacation.FitemList(i).Fidx%></a></td>
		<td><%= oVacation.FitemList(i).GetStateDivCDStr %></td>
		<td><%= Left(oVacation.FitemList(i).Fstartday,10) %> - <%= Left(oVacation.FitemList(i).Fendday,10) %></td>
		<td>
			<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FitemList(i).Ftotalday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			<% If oVacation.FitemList(i).Ftotalday = "0.5" Then Response.Write CHKIIF(oVacation.FItemList(i).Fhalfgubun="am","[����]","[����]") End If %>
		</td>
		<td><%= oVacation.FitemList(i).Fregistername %><% If oVacation.FitemList(i).Ftotalday = "0.5" Then Response.Write CHKIIF(empno=oVacation.FitemList(i).Fregisterempno,"<br>[<a href='#' onclick='jsDetailView("&oVacation.FitemList(i).Fidx&");return false;'>��������</a>]","") End If %></td>
		<td><%= oVacation.FitemList(i).Fapprovername %></td>
		<td>
			<% if (oVacation.FitemList(i).Fdeleteyn="N") and (oVacation.FitemList(i).Fstatedivcd="R") then %>
			<input type=button value=" �� �� " class="button" onclick="SubmitDelete(<%= masteridx %>, <%=oVacation.FitemList(i).Fidx%>)">
			<% end if %>

			<% if C_IS_SCM_LOGIN then %>
				<% if isNull(oVacation.FitemList(i).Freportidx) then %>
				<input type="button" class="button"  value="ǰ�Ǽ� �ۼ�" onClick="jsRegEapp('<%=oVacation.FitemList(i).Fidx%>','<%= Left(oVacation.FitemList(i).Fstartday,10) %> - <%= Left(oVacation.FitemList(i).Fendday,10) %>','<%= oVacation.FitemList(i).Ftotalday %>');" >
				<% else %>
				<input type="button" class="button"  value="ǰ�Ǽ� ����" onClick="jsViewEapp('<%=oVacation.FitemList(i).Freportidx%>','<%= oVacation.FitemList(i).Freportstate %>');" <% if (Not C_IS_SCM_LOGIN) then %>disabled<% end if %> >
				<% end if%>
			<% end if%>
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
			<input type="button" class="button" value="����Ʈ" onClick="ViewList(<%= masteridx %>);">
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!--���ڰ���--> 
	<form name="frmEapp" method="post" action="/tenmember/member/tenbyten_vacation_regeapp.asp">  
	<input type="hidden" name="iSL" value="">
	<input type="hidden" name="ieidx" value="<%if oVacation.FitemOne.Fdivcd = "5" then%>61<%else%>22<%end if%>"> <!-- ������ȣ ����!! ������ȣ���� 2013.11.25 -->
	<input type="hidden" name="uid" value="<%=oVacation.FItemOne.Fuserid%>">
	<input type="hidden" name="divcd" value="<%= oVacation.FItemOne.Fdivcd%>">
	<input type="hidden" name="uday" value="<%=oVacation.FItemOne.Fusedvacationday %>">
	<input type="hidden" name="rday" value="<%=oVacation.FItemOne.Frequestedday %>">
	<input type="hidden" name="tday" value="<%=oVacation.FItemOne.Ftotalvacationday%>"> 
	<input type="hidden" name="dDate" value="">
	<input type="hidden" name="dDay" value="">
	</form> 
	<!--/���ڰ���-->
 
<form name=frmmodify method=post action="modifyvacation_process.asp" onsubmit="return false;">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="empno" value="<%= empno %>">
	<input type="hidden" name="masteridx" value="">
	<input type="hidden" name="detailidx" value="">
</form>
<%Set oVacation = nothing%>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
