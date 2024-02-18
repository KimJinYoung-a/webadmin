<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
	Dim page, SearchKey, SearchString, isUsing, part_sn, research, orderby
	Dim job_sn, posit_sn, continuous_service_year, employeeonly

	page = Request("page")
	isUsing = Request("isUsing")
	SearchKey = Request("SearchKey")
	SearchString = Request("SearchString")
	part_sn = Request("part_sn")
	job_sn = Request("job_sn")
	posit_sn = Request("posit_sn")
	employeeonly = Request("employeeonly")
	research = Request("research")
	
	orderby = Request("orderby")
	
	if isUsing="" and research="" then isUsing="Y"
	if employeeonly="" and research="" then employeeonly="Y"
	if page="" then page=1



	'// ���� ����
	dim oMember, lp
	Set oMember = new CTenByTenMember

	oMember.FPagesize = 20
	oMember.FCurrPage = page
	oMember.FRectsearchKey = searchKey
	oMember.FRectsearchString = searchString
	oMember.FRectisUsing = isUsing
	oMember.FRectpart_sn = part_sn
	oMember.FRectjob_sn = job_sn
	oMember.FRectposit_sn = posit_sn
	oMember.FRectemployeeonly = employeeonly
	'oMember.FRectpart_sn = part_sn
	oMember.FRectOrderBy = orderby

	oMember.GetList
%>
<!-- �˻� ���� -->
<script language="javascript">

<!--
	// ������ �̵�
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}
//-->
</script>

<title>��󿬶���</title>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="#999999">
		<tr>
			<td width="400" style="padding:5; border-top:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999"  background="/images/menubar_1px.gif">
				<font color="#333333"><b>���� ��󿬶���</b></font>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<br><p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�μ�:
			<%=printPartOption("part_sn", part_sn)%>&nbsp;
			����:
			<%=printPositOption("posit_sn", posit_sn)%>&nbsp;
			��å:
			<%=printJobOption("job_sn", job_sn)%>&nbsp;
			
			<br>
			
			��뿩��:
			<select name="isUsing" class="select">
				<option value="">��ü</option>
				<option value="Y">���</option>
				<option value="N">����</option>
			</select>
			&nbsp;
			�˻�:
			<select name="SearchKey" class="select">
				<option value="">::����::</option>
				<option value="id">���̵�</option>
				<option value="company_name">����ڸ�</option>
			</select>
			<input type="text" class="text" name="SearchString" size="12" value="<%=SearchString%>">
			&nbsp;
			����:
			<select name="orderby" class="select">
				<option value="">����</option>
				<option value="t.joinday,p.posit_sn,t.username">�Ի���</option>
				<option value="t.username,p.posit_sn">�̸�</option>
				<option value="p.posit_sn,t.joinday,t.username">����</option>
				<option value="p.job_sn,p.posit_sn,t.username">��å</option>
				<option value="t.extension,t.username">����</option>
			</select>
			&nbsp;
			<input type="checkbox" name="employeeonly" value="Y"> ����̻�
			&nbsp;
			<script language="javascript">
				document.frm.isUsing.value="<%= isUsing %>";
				document.frm.SearchKey.value="<%= SearchKey %>";
				document.frm.orderby.value="<%= orderby %>";
				if ("Y" == "<%= employeeonly %>") {
					document.frm.employeeonly.checked = true;
				}
			</script>
			
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<br><p>

<!-- ��� �� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%=oMember.FtotalCount%></b>
			&nbsp;
			������ : <b><%= page %> / <%=oMember.FtotalPage%></b>
			&nbsp;&nbsp;&nbsp;
			�� �̸��� ���콺�� �������� ������ ��Ÿ���ϴ�.
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="70">��å</td>
		<td width="80">�̸�</td>
		<td width="190">�μ�</td>
		<td width="90">�ڵ�����ȣ</td>
		<td width="100">����ȭ��ȣ</td>
		<td width="85">ȸ����ȭ</td>
		<td width="35">����</td>
		<td width="90">�����ȣ(070)</td>
		<td>�̸���</td>
		<td>MSN�޽���</td>
    </tr>
	<% if oMember.FResultCount=0 then %>
	<tr>
		<td colspan="15" align="center" bgcolor="#FFFFFF">���(�˻�)�� ����ڰ� �����ϴ�.</td>
	</tr>
	<% else %>

	<% for lp=0 to oMember.FResultCount - 1 %>
	<tr height=30 align="center" bgcolor="<% if oMember.FitemList(lp).FisUsing="Y" then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
		<td><%=oMember.FitemList(lp).Fjob_name%></td>
		<td>
			<table cellpadding="0" cellspacing="0" border="0" class="a">
			<tr>
				<td id="photo<%=lp%>" alt="<img src='<%=oMember.FitemList(lp).FUserImage%>' width='110'>"><%=oMember.FitemList(lp).Fusername%>(<%=oMember.FitemList(lp).Fposit_name%>)</td>
			</tr>
			</table>
			<div id="ddd0" style="background-color:white; border-width:1px; border-style:solid; width:110; position:absolute; left:10; top:10; z-index:1; display:none"></div>
		</td>
		<td><%=oMember.FitemList(lp).Fpart_name%></td>
		<td><%=oMember.FitemList(lp).Fusercell%></td>
		<td><%=oMember.FitemList(lp).Fuserphone%></td>
		<td><%=oMember.FitemList(lp).Finterphoneno%></td>
		<td><%=oMember.FitemList(lp).Fextension%></td>
		<td><%=oMember.FitemList(lp).Fdirect070%></td>
		<td><%=oMember.FitemList(lp).Fusermail%></td>
		<td><%=oMember.FitemList(lp).Fmsnmail%></td>
	</tr>
	<% next %>

	<% end if %>
<!-- ���� ��� �� -->

<!-- ������ ���� -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<%
				if oMember.HasPreScroll then
					Response.Write "<a href='javascript:goPage(" & oMember.StartScrollPage-1 & ")'>[pre]</a>"
				else
					Response.Write "[pre]"
				end if

				for lp=0 + oMember.StartScrollPage to oMember.FScrollCount + oMember.StartScrollPage - 1

					if lp>oMember.FTotalpage then Exit for

					if CStr(page)=CStr(lp) then
						Response.Write " <font color='red'>[" & lp & "]</font> "
					else
						Response.Write " <a href='javascript:goPage(" & lp & ")'>[" & lp & "]</a> "
					end if

				next

				if oMember.HasNextScroll then
					Response.Write "<a href='javascript:goPage(" & lp & ")'>[next]</a>"
				else
					Response.Write "[next]"
				end if
			%>
		</td>
	</tr>
</table>


<script language="javascript">
document.onmousemove=function(){ 
	oElement = document.elementFromPoint(event.x, event.y);
	var ddd0 = document.getElementById("ddd0");
	if(oElement.id.indexOf('photo')!=-1)
	{
		ddd0.style.display='';
		ddd0.style.pixelLeft=event.x+10 + document.body.scrollLeft;
		ddd0.style.pixelTop=event.y-80 + document.body.scrollTop;
		ddd0.innerHTML=oElement.alt;
	} else { 
		ddd0.style.display='none';
	}
}
</script>

<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->