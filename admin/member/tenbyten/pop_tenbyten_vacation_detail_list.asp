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
	dim userid, vTmpIdx, vTmpUserid

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
	oVacation.FRectsearchKey = " t.userid "
	oVacation.FRectsearchString = session("ssBctId")
	oVacation.FRectIsDelete = "N"

	oVacation.GetMasterOne

	oVacation.GetDetailList



%>
<!-- �˻� ���� -->
<script language="javascript">
<!--
	function AddItem()
	{
<% if (oVacation.FItemOne.Fuserid = userid) then %>
	 window.open("pop_vacation_detail_modify.asp?masteridx=<%= masteridx %>","popAddIem","width=500,height=600,scrollbars=yes"); 
<% else %>
		alert("�ް��� ���θ� ��û�� �� �ֽ��ϴ�.");
<% end if %>
	}

	function ViewList(masteridx)
	{
		location.href = "/admin/member/tenbyten/pop_tenbyten_vacation_list.asp?masteridx=" + masteridx;
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
	
	//���ڰ��� ǰ�Ǽ� ���
	function jsRegEapp(scmidx, adate, aday){ 
		var winEapp = window.open("/admin/approval/eapp/regeapp.asp","popE","width=1000,height=600,scrollbars=yes");
		document.frmEapp.iSL.value = scmidx;
		document.all.divSL.innerHTML = scmidx; 
		document.all.divDate.innerHTML = adate + " ("+aday+"��)"; 
		document.frmEapp.tC.value = document.all.divEapp.innerHTML.replace(/\r|\n/g,"");
		document.frmEapp.target = "popE";
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
		var winDetail = window.open("pop_vacation_detail_view.asp?detailidx="+idx,"popDetail","width=500,height=300,scrollbars=yes");
		winDetail.focus();
	}

//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">�̸�(���̵�)</td>
		<td align="left">	
			<%= oVacation.FItemOne.Fusername %>(<%= oVacation.FItemOne.Fuserid %>)
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
			<%= oVacation.FItemOne.Ftotalvacationday %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">��밡�ɱⰣ</td>
		<td align="left">
			<%= Left(oVacation.FItemOne.Fstartday,10) %> - <%= Left(oVacation.FItemOne.Fendday,10) %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">����ϼ�</td>
		<td align="left">
			<%= oVacation.FItemOne.Fusedvacationday %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">���δ��</td>
		<td align="left">
			<%= oVacation.FItemOne.Frequestedday %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">�ܿ��ϼ�</td>
		<td align="left">
			<b><%= (oVacation.FItemOne.Ftotalvacationday - (oVacation.FItemOne.Fusedvacationday + oVacation.FItemOne.Frequestedday)) %></b>
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
		<td><a href="javascript:jsDetailView('<%=oVacation.FitemList(i).Fidx%>')"><%=oVacation.FitemList(i).Fidx%></a></td>
		<td><%= oVacation.FitemList(i).GetStateDivCDStr %></td>
		<td><%= Left(oVacation.FitemList(i).Fstartday,10) %> - <%= Left(oVacation.FitemList(i).Fendday,10) %></td>
		<td><%= oVacation.FitemList(i).Ftotalday %><% If oVacation.FitemList(i).Ftotalday = "0.5" Then Response.Write CHKIIF(oVacation.FItemList(i).Fhalfgubun="am","[����]","[����]") End If %></td>
		<td><%= oVacation.FitemList(i).Fregisterid %><% If oVacation.FitemList(i).Ftotalday = "0.5" Then Response.Write CHKIIF(userid=oVacation.FitemList(i).Fregisterid,"<br>[<a href='javascript:jsDetailView("&oVacation.FitemList(i).Fidx&");'>��������</a>]","") End If %></td>
		<td><%= oVacation.FitemList(i).Fapproverid %></td>
		<td>
			<% if (oVacation.FitemList(i).Fdeleteyn="N") and (oVacation.FitemList(i).Fstatedivcd="R") then %>
			<input type=button value=" �� �� " class="button" onclick="SubmitDelete(<%= masteridx %>, <%=oVacation.FitemList(i).Fidx%>)">
			<% end if %>
		 
			<% if isNull(oVacation.FitemList(i).Freportidx) then %>
			<input type="button" class="button"  value="ǰ�Ǽ� �ۼ�" onClick="jsRegEapp('<%=oVacation.FitemList(i).Fidx%>','<%= Left(oVacation.FitemList(i).Fstartday,10) %> - <%= Left(oVacation.FitemList(i).Fendday,10) %>','<%= oVacation.FitemList(i).Ftotalday %>');">
			<% else %>
			<input type="button" class="button"  value="ǰ�Ǽ� ����" onClick="jsViewEapp('<%=oVacation.FitemList(i).Freportidx%>','<%= oVacation.FitemList(i).Freportstate %>');">
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
	<form name="frmEapp" method="post" action="/admin/approval/eapp/regeapp.asp">
	<input type="hidden" name="tC" value=""> 
	<input type="hidden" name="ieidx" value="1"> <!-- ������ȣ ����!! -->
	<input type="hidden" name="iSL" value="">
	</form> 
	<div id="divEapp" style="display:none;"> 
	<table width="500" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
	<tr height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">idx</td>
		<td bgcolor="#FFFFFF" width="300">
			<div id="divSL"></div>
		</td>
	</tr>
	<tr height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">SCM ���̵�</td>
		<td bgcolor="#FFFFFF">
			<%= oVacation.FItemOne.Fuserid %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF">
			<%= oVacation.FItemOne.GetDivCDStr %>
		</td>
	</tr>
	<tr height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">����ϼ�/���δ��/���ϼ� </td>
    	<td bgcolor="#FFFFFF">
    		<%=oVacation.FItemOne.Fusedvacationday %> / <%=oVacation.FItemOne.Frequestedday%> / <%=  oVacation.FItemOne.Ftotalvacationday%>
    	</td>
    </tr>
	<tr height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">��û�Ⱓ</td>
    	<td bgcolor="#FFFFFF">
    		<div id="divDate"></div>
    	</td>
    </tr> 
	</table> 
	</div>
	<!--/���ڰ���-->

<form name=frmmodify method=post action="domodifyvacation.asp" onsubmit="return false;">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="masteridx" value="">
	<input type="hidden" name="detailidx" value="">
	<input type="hidden" name="userid" value="<%=userid%>">
</form>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->