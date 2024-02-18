<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [cs]�����ڵ����
' Hieditor : �̻� ����
'			 2023.08.28 �ѿ�� ����(�����⿩�� �߰�, �ҽ�ǥ���ڵ�� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/CsCommCdcls.asp"-->
<%
dim comm_cd, page, groupCd, searchKey, searchString, comm_isDel, sortType, oComm, i, lp, bgcolor, strUsing
dim dispyn
	comm_cd     = requestCheckVar(request("comm_cd"),32)
	page        = requestCheckVar(getNumeric(request("page")),9)
	groupCd     = requestCheckVar(request("groupCd"),32)
	searchKey   = requestCheckVar(request("searchKey"),32)
	searchString = requestCheckVar(request("searchString"),32)
	comm_isDel  = requestCheckVar(request("comm_isDel"),32)
	sortType	= requestCheckVar(request("sortType"),2)
	dispyn	= requestCheckVar(request("dispyn"),2)

if page="" then page=1
if searchKey="" then searchKey="comm_name"
if sortType="" then sortType="sa"

set oComm = new CCommCd
	oComm.FCurrPage = page
	oComm.FPageSize = 50
	oComm.FRectgroupCd = groupCd
	oComm.FRectsearchKey = searchKey
	oComm.FRectsearchString = searchString
	oComm.FSortType = sortType
	oComm.FRectisDel = comm_isDel
	oComm.FRectdispyn = dispyn
	oComm.GetCommList
%>
<script type='text/javascript'>

function popCsAsGubunHelpEdit(icomm_cd){
	var popwin = window.open('popCsAsGubunHelpEdit.asp?comm_cd=' + icomm_cd,'popCsAsGubunHelpEdit','width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function chk_form(frm){
	if(!frm.searchKey.value)
	{
		alert("�˻� ������ �������ֽʽÿ�.");
		frm.searchKey.focus();
		return false;
	}
	else if(!frm.searchString.value)
	{
		alert("�˻�� �Է����ֽʽÿ�.");
		frm.searchString.focus();
		return false;
	}

	frm.page.value= 1;
	frm.submit();
}

function goPage(pg){
	var frm = document.frm_search;

	frm.page.value= pg;
	frm.submit();
}

function chgSort(t,s) {
	var frm = document.frm_search;
	frm.sortType.value= t+s;
	frm.submit();
}

function popCommCdReg(){
	var popwin = window.open('/cscenter/comm/commCd_write.asp?menupos=<%=menupos%>','popCommCdReg','width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popCommCdEdit(comm_cd){
	var popwin = window.open('/cscenter/comm/CommCd_modi.asp?comm_cd='+comm_cd+'&menupos=<%=menupos%>','popCommCdEdit','width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<!-- �˻� ���� -->
<form name="frm_search" method="GET" action="CommCd_list.asp" onSubmit="return chk_form(this)" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>" />
<input type="hidden" name="menupos" value="<%=menupos%>" />
<input type="hidden" name="sortType" value="<%=sortType%>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			* ����:
			<select class="select" name="comm_isDel" onChange="goPage(1)">
				<option value="">��ü</option>
				<option value="N">���</option>
				<option value="Y">����</option>
			</select>
			&nbsp;
			* �׷�:
			<select class="select" name="groupCd" onChange="goPage(1)">
				<option value="">��ü</option>
				<%
				' ���� ����(2010������)�� ������� �ڵ尡 ������ �ڵ带 ������ ������ ������ cs�� ��� ����̰�, ���尪�� �־ ������ �Ұ�����.
				' comm_group "z999" �ڵ忡 "����" ������ �ڵ带 �߰��� ����� ��� �˻��̶� �ǰ� ���밪�� �켱 ����� ����.
				' comm_cd ����.. comm_group �ڵ尪�� ��ȯ������ ��� ���� ������ ������ �Ǿ� ����.
				' comm_cd ���� db_cs.dbo.tbl_new_as_list ���̺� gubun01 �ʵ忡 �ԷµǴ� ������.
				%>
				<option value="C004" <% if groupCd="C004" then response.write "selected" %>>����</option>
				<%= oComm.optGroupCd(groupCd)%>
			</select>
			&nbsp;
			* ���⿩�� : <% drawSelectBoxisusingYN "dispyn", dispyn,"" %>
			&nbsp;
			* �˻�:
			<select class="select" name="searchKey">
				<option value="comm_cd">�����ڵ�</option>
				<option value="comm_name">�ڵ��</option>
			</select>
			<script language="javascript">
				document.frm_search.comm_isDel.value="<%=comm_isDel%>";
				document.frm_search.searchKey.value="<%=searchKey%>";
			</script>
			&nbsp;
			<input type="text" class="text" name="searchString" size="20" value="<%= searchString %>">
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="goPage(1)" />
		</td>
	</tr>
</table>
</form>

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<font color="red">�����ڵ尡 �߰��ǰ�, �����ڵ��� �����ڵ�� �߰����� �ʽ��ϴ�.</font>
	</td>
	<td align="right">
		<input type="button" value="�űԵ��" onclick="popCommCdReg();" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oComm.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oComm.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" width="60">���� <span onclick="chgSort('s','<%=chkIIF(left(sortType,1)="s",chkIIF(right(sortType,1)="d","a","d"),"a")%>')" style="cursor:pointer;"><%=chkIIF(left(sortType,1)="s",chkIIF(right(sortType,1)="d","��","��"),"��")%></span></td>
	<td align="center" width="140">�׷� <span onclick="chgSort('g','<%=chkIIF(left(sortType,1)="g",chkIIF(right(sortType,1)="d","a","d"),"a")%>')" style="cursor:pointer;"><%=chkIIF(left(sortType,1)="g",chkIIF(right(sortType,1)="d","��","��"),"��")%></span></td>
	<td align="center" width="80">�����ڵ� <span onclick="chgSort('c','<%=chkIIF(left(sortType,1)="c",chkIIF(right(sortType,1)="d","a","d"),"a")%>')" style="cursor:pointer;"><%=chkIIF(left(sortType,1)="c",chkIIF(right(sortType,1)="d","��","��"),"��")%></span></td>
	<td align="center">�ڵ��</td>
	<td align="center" width="90">����Ʈ���⿩��</td>
	<td align="center" width="50">Color</td>
	<td align="center" width="50">����</td>
	<td align="center" width="100">���</td>
</tr>
<%
for lp=0 to oComm.FResultCount - 1
	if oComm.FItemList(lp).Fcomm_isDel="<font color=darkblue>���</font>" then
		bgcolor = "#FFFFFF"
	else
		bgcolor = "#E0E0E0"
	end if
%>
<tr align="center" bgcolor="<%=bgcolor%>">
	<td><%= oComm.FItemList(lp).Fsortno %></td>
	<td><%= oComm.FItemList(lp).Fgroup_name %></td>
	<td><%= oComm.FItemList(lp).Fcomm_cd %></td>
	<td align="left"><%= db2html(oComm.FItemList(lp).Fcomm_name) %></td>
	<td ><%= oComm.FItemList(lp).fdispyn %></td>
	<td ><%= oComm.FItemList(lp).Fcomm_color %></td>
	<td><%= oComm.FItemList(lp).Fcomm_isDel %></td>
	<td>
		<input type="button" value="����" onclick="popCommCdEdit('<%= oComm.FItemList(lp).Fcomm_cd %>');" class="button">

		<% if Left(oComm.FItemList(lp).Fcomm_cd,1)="A" then %>
			<br><input type="button" value="����(�ȳ�����)" onclick="popCsAsGubunHelpEdit('<%= oComm.FItemList(lp).Fcomm_cd %>');" class="button">
		<% end if %>
	</td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<!-- ������ ���� -->
		<%
			if oComm.HasPreScroll then
				Response.Write "<a href='javascript:goPage(" & oComm.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
			else
				Response.Write "[pre] &nbsp;"
			end if

			for i=0 + oComm.StartScrollPage to oComm.FScrollCount + oComm.StartScrollPage - 1

				if i>oComm.FTotalpage then Exit for

				if CStr(page)=CStr(i) then
					Response.Write " <font color='red'>[" & i & "]</font> "
				else
					Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
				end if

			next

			if oComm.HasNextScroll then
				Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
			else
				Response.Write "&nbsp; [next]"
			end if
		%>
		<!-- ������ �� -->
	</td>
</tr>
</table>

<%
set oComm = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->