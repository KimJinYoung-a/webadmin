<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �μ�����
' History : 2011.1.19 ������ ����
'			2011.12.16 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
	Dim page, SearchKey, SearchString, isUsing, part_sn, research, orderby
	Dim job_sn, posit_sn, continuous_service_year, employeeonly
	Dim iTotCnt,iPageSize, iTotalPage

	iPageSize = 20
	page = requestCheckvar(Request("page"),10)
	isUsing = requestCheckvar(Request("isUsing"),1)
	SearchKey = requestCheckvar(Request("SearchKey"),1)
	SearchString = requestCheckvar(Request("SearchString"),32)
	part_sn = requestCheckvar(Request("part_sn"),10)
	job_sn = requestCheckvar(Request("job_sn"),10)
	posit_sn = requestCheckvar(Request("posit_sn"),10)

	research = requestCheckvar(Request("research"),2)

	orderby = requestCheckvar(Request("orderby"),1)

	if isUsing="" and research="" then isUsing="Y"
	if page="" then page=1

	'// �α�������(���)�� ���� �⺻ �μ� ����(������ �̻�:2 �� �ý�����:7 ����)
	'SCM �޴����� �������� �����Ѵ�.
	'if Not(session("ssAdminLsn")<=2 or session("ssAdminPsn")=7) then
	'	part_sn = session("ssAdminPsn")
	'end if


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
	oMember.Fposit_sn 		= posit_sn
	oMember.Forderby 		= orderby

	arrList = oMember.fnGetMemberList
	iTotCnt = oMember.FTotCnt
	set oMember = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
%>
<!-- �˻� ���� -->
<script language="javascript">

function ViewVacationByID(userid)
{
	window.open("/admin/member/tenbyten/tenbyten_vacation_list.asp?menupos=1178&research=on&page=&part_sn=&deleteyn=N&SearchKey=t.userid&SearchString=" + userid ,"ViewVacationByID","width=1000,height=450,scrollbars=yes")
}


	// �ű� ����� ���
	function AddItem()
	{
		var w = window.open("pop_member_reg.asp","popAddIem","width=700,height=800,scrollbars=yes");
		w.focus();
	}

	// ����� ����/����
	function ModiItem(empno)
	{
		var w = window.open("pop_member_modify.asp?sEPN="+empno,"ModiItem","width=700,height=800,scrollbars=yes");
		w.focus();
	}

	// ������ �̵�
	function jsGoPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}

	//�޿����� ���
	function ViewPay(empno){
		var w = window.open("pop_payform.asp?sEN="+empno,"ModiItem","width=700,height=800,scrollbars=yes");
		w.focus();
	}

</script>
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
			<%=printPositOptionIN90("posit_sn", posit_sn)%>&nbsp;
			��å:
			<%=printJobOption("job_sn", job_sn)%>&nbsp;

			<br>

			��������:
			<select name="isUsing" class="select">
				<option value="">��ü</option>
				<option value="Y">����</option>
				<option value="N">���</option>
			</select>
			&nbsp;
			�˻�:
			<select name="SearchKey" class="select">
				<option value="">::����::</option>
				<option value="1" >���̵�</option>
				<option value="2">����ڸ�</option>
				<option value="3">���</option>
			</select>
			<input type="text" class="text" name="SearchString" size="17" value="<%=SearchString%>">
			&nbsp;
			����:
			<select name="orderby" class="select">
				<option value="">�̸�</option>
				<option value="1">�Ի���</option>
				<option value="2">����</option>
				<option value="3">��å</option>
				<option value="4">����</option>
			</select>

			&nbsp;
			<script language="javascript">
				document.frm.isUsing.value="<%= isUsing %>";
				document.frm.SearchKey.value="<%= SearchKey %>";
				document.frm.orderby.value="<%= orderby %>";
			</script>

		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
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
			<input type="button" class="button" value="�űԵ��" onClick="javascript:AddItem('');">
		</td>
		<td align="right">
			<!--<img src="/images/icon_excel.gif" border="0">-->
		</td>
	</tr>
</table>
<!-- �׼� �� -->
<p>

<!-- ��� �� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			�˻���� : <b><%=iTotCnt%></b>
			&nbsp;
			������ : <b><%= page %> / <%=iTotalPage%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="70">��å</td>
		<td width="100">���</td>
		<td>�̸�</td>
		<td width="100">���̵�</td>
		<td width="70">�Ի���/<br>�����</td>
		<td width="40">����</td>
		<td width="190">�μ�<Br>��ǥ����(���������)</td>
		<td width="100">����</td>
		<td>�̸���</td>
		<td>ȸ����ȭ</td>
		<td>����</td>
		<td>�����ȣ(070)</td>
		<td>�ް�</td>

    </tr>
	<% if isArray(arrList) then %>
	<% for intLoop=0 to ubound(arrList,2) %>
	<tr height=30 align="center" bgcolor="<% if  (arrList(15,intLoop)="Y") then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">

		<td><%=arrList(14,intLoop)%></td>
		<td><a href="javascript:ModiItem('<%=arrList(0,intLoop)%>')"><%=arrList(0,intLoop)%></td>
		<td><%=arrList(1,intLoop)%></td>
		<td><a href="javascript:ModiItem('<%=arrList(0,intLoop)%>')"><%=arrList(2,intLoop)%></a></td>
		<td><%= Left(arrList(3,intLoop), 10) %><%IF not isNull(arrList(4,intLoop)) and arrList(15,intLoop) ="N" THEN %><br><font color="blue"><%= Left(arrList(4,intLoop), 10) %></font><%END IF%></td>
		<td><%IF Not isNull(arrList(3,intLoop)) and Left(arrList(0,intLoop), 1) = "1" THEN %><%= GetYearDiff(arrList(3,intLoop)) %><%END IF%></td>
		<td>
			<%=arrList(12,intLoop)%>
			<% if arrList(5,intLoop) = "16" or arrList(5,intLoop) = "18" or arrList(5,intLoop) = "19" then %>
				<Br><a href="javascript:shopreg('<%= arrList(0,intLoop) %>');" onfocus="this.blur()">
				<font color="grey">	
				<% if arrList(22,intLoop) <> "" then %>
					<%=arrList(22,intLoop)%>/<%=arrList(21,intLoop)%> (<%=arrList(23,intLoop)%>��)
				<% else %>
					��������
				<% end if %>
			<% end if %>
		</td>
		<td><%=arrList(13,intLoop)%></td>
		<td><%=arrList(8,intLoop)%></td>
		<td><%=arrList(9,intLoop)%></td>
		<td><%=arrList(10,intLoop)%></td>
		<td><%=arrList(11,intLoop)%></td>
		<td><input type="button" class="button" value="�ް�" onClick="javascript:ViewVacationByID('<%=arrList(2,intLoop)%>');"></td>

	</tr>
	<% next %>
	<% else %>
	<tr>
		<td colspan="16" align="center" bgcolor="#FFFFFF">���(�˻�)�� ����ڰ� �����ϴ�.</td>
	</tr>
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
		<td colspan="16" align="center">
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
</table>
<!-- ������ �� -->
 <!-- #include virtual="/lib/db/dbclose.asp" -->