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
' PageName : Category_left_BrandFocus.asp
' Discription : ī�װ� ���� �귣�� ��Ŀ�� ���
' History : 2008.04.04 ������ : ����
'###############################################

'// ���� ����
dim cdl, page, i, lp
cdl = request("cdl")
page = request("page")

if page="" then page=1

dim omd
set omd = New CMDSRecommend
omd.FCurrPage = page
omd.FPageSize=10
omd.FRectCDL = cdl
omd.GetBrandFocusList

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
		upfrm.action="doCategoryLeftbrandFocus.asp";
		upfrm.submit();
	}
}

function RefreshLeftbrandFocusRec(){
	if (document.refreshFrm.cdl.value == ""){
		alert("ī�װ��� �������ּ���");
		document.refreshFrm.cdl.focus();
	}
	else{
		 var popwin = window.open('','refreshPop','');
		 popwin.focus();
		 refreshFrm.target = "refreshPop";
		 refreshFrm.action = "<%=wwwUrl%>/chtml/make_category_left_brandFocus_JS.asp";
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
		upfrm.action="doCategoryLeftbrandFocus.asp";
		upfrm.submit();
	} else {
		return;
	}
}

// �귣�� �߰� ó��
function addbrandFocus(upfrm)
{
	if(!upfrm.cdl.value) {
		alert("�߰��� ��� ī�װ��� �������ּ���.");
		return;
	}

	if(!upfrm.makerid.value) {
		alert("�귣��ID�� �Է����ּ���.");
		return;
	}

	if (confirm('�����Ͻ� �귣�带 �߰��Ͻðڽ��ϱ�?')) {
		upfrm.mode.value="add";
		upfrm.action="doCategoryLeftbrandFocus.asp";
		upfrm.submit();
	}
}

// �귣�� �˻� �˾�
function popBrandSearch(fm,tg){
	var popup_item = window.open("/admin/member/popBrandSearch.asp?frmName=" + fm + "&compName=" + tg, "popup_brand", "width=800,height=500,scrollbars=yes,status=no");
	popup_item.focus();
}

// ������ �̵�
function goPage(pg)
{
	document.refreshFrm.page.value=pg;
	document.refreshFrm.action="category_left_brandFocus.asp";
	document.refreshFrm.submit();
}

// ī�װ� ����� ���
function changecontent(){
	document.frmarr.cdl.value=refreshFrm.cdl.value;
}
//-->
</script>
<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" onSubmit="frm_search()" action="category_left_brandFocus.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		ī�װ� <% DrawSelectBoxCategoryLarge "cdl", cdl %>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<form name="frmarr" method="get" action="doCategoryLeftbrandFocus.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="cdl" value="<%=cdl%>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="idx" value="">
<input type="hidden" name="arrSort" value="">
<tr>
	<td><input type="button" value="���þ����ۻ���" onclick="delitems(frmarr);" class="button"></td>
	<td align="right">
		<img src="/images/icon_reload.gif" onClick="RefreshLeftbrandFocusRec()" style="cursor:pointer" align="absmiddle" alt="html�����">
		����Ʈ�� ���� /
		<input type="button" class="button" value="��������" onclick="changeSort(frmarr);">
		/
		�귣��ID
		<input type="text" class="text" name="makerid" value="" onClick="popBrandSearch('frmarr','makerid')" style="cursor:pointer">
		<input type="button" value="������ �߰�" onclick="addbrandFocus(frmarr)" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="6">
		�˻���� : <b><%=omd.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=omd.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>��ȣ</td>
	<td>ī�װ���</td>
	<td>��ü��</td>
	<td>�̹���</td>
	<td>����</td>
</tr>
<%	if omd.FResultCount < 1 then %>
<tr>
	<td colspan="6" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� �������� �����ϴ�.</td>
</tr>
<%
	else
		for i=0 to omd.FResultCount-1
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><input type="checkbox" name="selIdx" value="<%= omd.FItemList(i).Fidx %>"></td>
	<td><%= omd.FItemList(i).Fidx %></td>
	<td><%= omd.FItemList(i).Fcode_nm %></td>
	<td><%= omd.FItemList(i).Fmakerid %></td>
	<td><img src="<%= omd.FItemList(i).FImageSmall %>"><img src="<%= omd.FItemList(i).Ftitleimgurl %>" ></td>
	<td><input type="text" class="text" name="SortNo" value="<%=omd.FItemList(i).FsortNo%>" size="2" style="text-align:center"></td>
</tr>
<%
		next
	end if
%>
<tr bgcolor="#FFFFFF">
	<td colspan="6" align="center">
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