<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ��������Ʈ
' History : 2017.04.10 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
	Dim page, SearchKey, SearchString, isUsing, part_sn, research, orderby
	Dim job_sn, posit_sn, continuous_service_year, employeeonly
	Dim iTotCnt,iPageSize, iTotalPage
	dim department_id, inc_subdepartment,nodepartonly

	iPageSize = 20
	page = requestCheckVar(Request("page"),10)
	isUsing = requestCheckVar(Request("isUsing"),10)
	SearchKey = requestCheckVar(Request("SearchKey"),32)
	SearchString = requestCheckVar(Request("SearchString"),32)
	part_sn = requestCheckVar(Request("part_sn"),10)
	job_sn = requestCheckVar(Request("job_sn"),10)
	posit_sn = requestCheckVar(Request("posit_sn"),10)
	research = requestCheckVar(Request("research"),2)

	department_id = requestCheckvar(Request("department_id"),10)
	inc_subdepartment = requestCheckvar(Request("inc_subdepartment"),1)
	nodepartonly = requestCheckvar(Request("nodepartonly"),1)

	orderby = requestCheckvar(Request("orderby"),10)

	if isUsing="" and research="" then isUsing="Y"
	if page="" then page=1
	'if posit_sn ="" then posit_sn = 99
	'// ���� ����
	dim oMember, arrList,intLoop
	Set oMember = new CTenByTenMember

	oMember.FPagesize 	= iPageSize
	oMember.FCurrPage 	= page
	oMember.FSearchType 	= searchKey
	oMember.FSearchText 	= searchString
	oMember.Fstatediv 		= isUsing
	oMember.Fpart_sn 		= part_sn
	oMember.Fjob_sn 		= job_sn
	oMember.Fposit_sn 	= posit_sn
	oMember.Forderby 		= orderby

	oMember.Fdepartment_id 		= department_id
	oMember.Finc_subdepartment 	= inc_subdepartment
	oMember.FRectNoDepartOnly 	= nodepartonly

	arrList = oMember.fnGetMemberList
	iTotCnt = oMember.FTotCnt
	set oMember = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
%>
<!-- �˻� ���� -->
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
<!--
	// ������ �̵�
	function jsGoPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}
	function mywork_update(emp){
		var popwin = window.open('pop_mywork_update.asp?empno=' + emp,'pop','width=500,height=200,scrollbars=yes,resizable=yes');
		popwin.focus();
	}

	$(function(){
		$(".colName").mouseenter(function(){
			$(this).find(".colPhoto").show();
		}).mousemove(function(e){
			$(this).find(".colPhoto").css("top",e.pageY-20).css("left",e.pageX+20)
		}).mouseleave(function(){
			$(this).find(".colPhoto").hide();
		});
	});
//-->
</script>
<style type="text/css">
body{ margin:0; }
</style>
<title>��󿬶���</title>

<table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#FFFFFF">
<tr>
	<td width="30%">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
		<tr>
			<td height="26" align="center" width="50%" bgcolor="#FFFFFF" style="cursor:pointer;" onClick="location.href='/common/pop_organization_chart.asp';"><font size="2">������</font></td>
			<td align="center" width="50%" bgcolor="#EDEDED" style="cursor:pointer;" onClick="location.href='/common/pop_tenbyten_member_list.asp';"><strong><font size="2">��󿬶���</font></strong></td>
		</tr>
		</table>
	</td>
	<td width="70%" style="border-bottom: 1px solid #CCCCCC;"></td>
</tr>
<tr>
	<td colspan="2" height="10"></td>
</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�μ�NEW:
			<%= drawChSelectBoxDepartment("department_id", department_id,"") %>
			<input type="checkbox" name="inc_subdepartment" value="N" <% if (inc_subdepartment = "N") then %>checked<% end if %> > ���� �μ����� ����&nbsp;
			<% if C_ADMIN_AUTH or C_PSMngPart then %>
			����:
			<%=printPositOptionIN90("posit_sn", posit_sn)%>&nbsp;
			<% end if %>
			��å:
			<%=printJobOption("job_sn", job_sn)%>&nbsp;

			<br>

			��뿩��:
			<select name="isUsing" class="select">
				<option value="">��ü</option>
				<option value="Y">����</option>
				<option value="N">���</option>
			</select>
			&nbsp;
			�˻�:
			<select name="SearchKey" class="select">
				<option value="">::����::</option>
				<option value="1">���̵�</option>
				<option value="2">����ڸ�</option>
				<option value="3">���</option>
				<option value="4">�ڵ���</option>
			</select>
			<input type="text" class="text" name="SearchString" size="17" value="<%=SearchString%>">
			&nbsp;
			����:
			<select name="orderby" class="select">
				<option value="">�̸�</option>
				<option value="6">�Ի���(�ֱټ�)</option>
				<option value="5">�����(�ֱټ�)</option>
				<!--<option value="2">����</option>-->
				<option value="3">��å</option>
				<option value="4">����</option>
			</select>
			&nbsp;
			<input type="checkbox" name="nodepartonly" value="Y" <% if (nodepartonly = "Y") then %>checked<% end if %> > �μ�NEW �������� (������ ��� �μ� ������ ����)

			<script type="text/javascript">
				document.frm.isUsing.value="<%= isUsing %>";
				document.frm.SearchKey.value="<%= SearchKey %>";
				document.frm.orderby.value="<%= orderby %>";
			</script>

		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
</table>
<!-- �˻� �� -->

<br><p>

<!-- ��� �� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%=iTotCnt%></b>
			&nbsp;
			������ : <b><%= page %> / <%=iTotalPage%></b>
			&nbsp;&nbsp;&nbsp;
			�� �̸��� ���콺�� �������� ������ ��Ÿ���ϴ�.
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80">��å</td>
		<td width="130">�̸�</td>
		<% if C_ADMIN_AUTH or C_PSMngPart then %><td width="80">����</td><% end if %>
		<td>�μ�</td>
		<td>������</td>
		<td width="90">�ڵ�����ȣ</td>
		<td width="85">ȸ����ȭ</td>
		<td width="35">����</td>
		<td width="110">�����ȣ(070)</td>
		<td>�̸���</td>
    </tr>
	<% if not isArray(arrList)  then %>
	<tr>
		<td colspan="15" align="center" bgcolor="#FFFFFF">���(�˻�)�� ����ڰ� �����ϴ�.</td>
	</tr>
	<% else %>

	<% for intLoop = 0 To UBound(arrList,2) %>
	<tr height=30 align="center" bgcolor="<% if  (arrList(15,intLoop)="Y") then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'"  >
		<td><%=arrList(14,intLoop)%></td>
		<td class="colName" alt="<%=intLoop%>">
			<b><%=arrList(1,intLoop)%></b>
			<% if arrList(16,intLoop)<>"" then %>
			<div class="colPhoto" id="lyEmpPhoto<%=intLoop%>" style="background-color:white; position:absolute; left:10; top:10; z-index:1; display:none">
				<img src="<%=replace(arrList(16,intLoop),"http://webimage.10x10.co.kr","/webimage")%>" alt="<%=arrList(1,intLoop)%>" height="110" />
			</div>
			<% end if %>
		</td>
		<% if C_ADMIN_AUTH or C_PSMngPart then %><td><%=arrList(13,intLoop)%></td><% end if %>
		<td align="left"><%=arrList(27,intLoop)%></td>
		<td width="130"><%=arrList(20,intLoop)%><!--% If (session("ssAdminPsn") = "7" or (session("ssAdminPOsn") > 0 and session("ssAdminPOsn") =< "3")) or ((session("ssAdminPOsn") > "0" and session("ssAdminPOsn") < "6") and (session("ssAdminPsn") = arrList(5,intLoop)))  Then %--><!--input type ="button" name="mywork" value="����" onclick="mywork_update('<%=arrList(0,intLoop)%>')" class="button"--><!--%End If%--></td>
		<td><%=arrList(17,intLoop)%></td>
		<td><%=arrList(9,intLoop)%></td>
		<td><b><%=arrList(10,intLoop)%></b></td>
		<td><b><%=arrList(11,intLoop)%></b></td>
		<td><%=arrList(8,intLoop)%></td>
	</tr>
	<% next %>

	<% end if %>
<!-- ���� ��� �� -->

<!-- ������ ���� -->
<%
Dim iStartPage,iEndPage,iX,iPerCnt
iPerCnt = 10

iStartPage = (Int((page-1)/iPerCnt)*iPerCnt) + 1

If (page mod iPerCnt) = 0 Then
	iEndPage = page
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
			    <tr valign="bottom" height="25">
			        <td valign="bottom" align="center">
			         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
					<% else %>[pre]<% end if %>
			        <%
						for ix = iStartPage  to iEndPage
							if (ix > iTotalPage) then Exit for
							if Cint(ix) = Cint(page) then
					%>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
					<%		else %>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
					<%
							end if
						next
					%>
			    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
					<% else %>[next]<% end if %>
			        </td>
			    </tr>
			</table>
		</td>
	</tr>
	</form>
</table>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
