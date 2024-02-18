<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp" -->
<%
'###############################################
' PageName : Category_left_bestBrand.asp
' Discription : ī�װ� ���� ����Ʈ �귣�� ���
' History : 2008.04.02 ������ : ����
'			2008.05.06 ������ ���� [ fnSearch�Լ�,�Ǽ����������� �߰�(isUsing='Y'))
'###############################################

'// ���� ����
dim cdl, page, isusing, i, lp
cdl = request("cdl")
page = request("page")
isusing = request("isusing")

if page="" then page=1
if isusing="" then isusing="Y"

dim omd
set omd = New CMDSRecommend
omd.FCurrPage = page
omd.FPageSize=10
omd.FRectCDL = cdl
omd.FRectIsusing = isusing
omd.GetBestBrandList

%>
<script language='javascript'>
<!--
function ckAll(icomp){
	var bool = icomp.checked;
	var frm = document.frmarr;

	if(frm.selIdx.length) {
		for (var i=0;i<frm.selIdx.length;i++){
			frm.selIdx[i].checked = bool;
		}
	} else {
		frm.selIdx.checked = bool;
	}
}

function CheckSelected(){
	var pass = false;
	var frm = document.frmarr;

	if(frm.selIdx.length) {
		for (var i=0;i<frm.selIdx.length;i++){
			pass = ((pass)||(frm.selIdx[i].checked));
		}
	} else {
		pass = ((pass)||(frm.selIdx.checked));
	}

	if (!pass) {
		return false;
	}
	return true;
}

function delitems(upfrm){
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}

	if (confirm('���� �������� �����Ͻðڽ��ϱ�?')) {
		upfrm.mode.value="del";
		upfrm.action="doCategoryLeftBestBrand.asp";
		upfrm.submit();
	}
}

function RefreshLeftBestBrandRec(){
	if (document.refreshFrm.cdl.value == ""){
		alert("ī�װ��� �������ּ���");
		document.refreshFrm.cdl.focus();
	}
	else{
		 var popwin = window.open('','refreshPop','');
		 popwin.focus();
		 refreshFrm.target = "refreshPop";
		 refreshFrm.action = "<%=wwwUrl%>/chtml/make_category_left_bestBrand_JS.asp";
		 refreshFrm.submit();
	}
}

function changeSort(upfrm){
	var arrSort="";
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}

	if(confirm('���� �����ۿ� �Է��Ͻ� ������ȣ��� �����մϴ�.')) {

		if(upfrm.selIdx.length) {
			for (var i=0;i<upfrm.selIdx.length;i++){
				if(upfrm.selIdx[i].checked) arrSort = arrSort + upfrm.SortNo[i].value + ",";
			}
		} else {
			if(upfrm.selIdx.checked) arrSort=upfrm.SortNo.value;
		}
		upfrm.arrSort.value = arrSort;

		upfrm.mode.value="changeSort";
		upfrm.action="doCategoryLeftBestBrand.asp";
		upfrm.submit();
	} else {
		return;
	}
}

// �귣�� �߰� �������� �̵�
function addBestBrand()
{
	document.frmarr.cdl.value = document.refreshFrm.cdl.value;
	document.frmarr.mode.value = "add";
	document.frmarr.action="category_left_bestBrand_write.asp";
	document.frmarr.submit();
}

// ������ �̵�
function goPage(pg)
{
	document.refreshFrm.page.value=pg;
	document.refreshFrm.action="category_left_bestBrand.asp";
	document.refreshFrm.submit();
}
function fnSearch()
{
	document.refreshFrm.action='category_left_BestBrand.asp';
	document.refreshFrm.target='';
	document.refreshFrm.submit();
}

// ī�װ� ����� ���
function changecontent(){ }
//-->
</script>
<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" onSubmit="frm_search()" action="category_left_BestBrand.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		ī�װ� <% DrawSelectBoxCategoryLarge "cdl", cdl %> /
		������� <select name="isusing" class="select"><option value="Y">Yes</option><option value="N">No</option></select>
		<script language="javascript">
			document.refreshFrm.isusing.value="<%=isusing%>";
		</script>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" onclick="fnSearch();" value="�˻�">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<form name="frmarr" method="get" action="doCategoryLeftBestBrand.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="cdl" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="idx" value="">
<input type="hidden" name="arrSort" value="">
<tr>
	<td><input type="button" value="���þ����ۻ���" onclick="delitems(frmarr);" class="button"></td>
	<td align="right">
		<img src="/images/icon_reload.gif" onClick="RefreshLeftBestBrandRec()" style="cursor:pointer" align="absmiddle" alt="html�����">
		����Ʈ�� ���� /
		<input type="button" class="button" value="��������" onclick="changeSort(frmarr);">
		/
		<input type="button" value="������ �߰�" onclick="addBestBrand()" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="7">
		�˻���� : <b><%=omd.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=omd.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>ī�װ���</td>
	<td>��ü��</td>
	<td>�̹���</td>
	<td>�������</td>
	<td>����</td>
	<td>�����</td>
</tr>
<%	if omd.FResultCount < 1 then %>
<tr>
	<td colspan="7" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� �������� �����ϴ�.</td>
</tr>
<%
	else
		for i=0 to omd.FResultCount-1
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><input type="checkbox" name="selIdx" value="<%= omd.FItemList(i).Fidx %>"></td>
	<td><%= omd.FItemList(i).Fcode_nm %></td>
	<td><%= omd.FItemList(i).Fmakerid %></td>
	<td>
		<a href="/admin/categorymaster/category_left_bestbrand_write.asp?mode=edit&idx=<%= omd.FItemList(i).Fidx %>&page=<%=page%>">
		<img src="<%= staticImgUrl & "/left/bestbrand/" & omd.FItemList(i).Fimage %>" border="0" height="60"></a>
	</td>
	<td><%= omd.FItemList(i).Fisusing %></td>
	<td><input type="text" class="text" name="SortNo" value="<%=omd.FItemList(i).FsortNo%>" size="2" style="text-align:center"></td>
	<td><%= FormatDateTime(omd.FItemList(i).Fregdate,2) %></td>
</tr>
<%
		next
	end if
%>
<tr bgcolor="#FFFFFF">
	<td colspan="7" align="center">
	<!-- ������ ���� -->
	<%
		if omd.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & omd.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + omd.StartScrollPage to omd.FScrollCount + omd.StartScrollPage - 1

			if lp>omd.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if omd.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- ������ �� -->
	</td>
</tr>
</table>
<%
set omd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->