<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp"-->
<%
'###############################################
' PageName : Category_left_topKeyword.asp
' Discription : ī�װ� ���� žŰ���� ���
' History : 2008.03.29 ������ : ����
'         : 2008.10.27 ��ī�װ� ó�� �߰�(������)
'         : 2009.04.15 �̹��� �߰�(������)
'###############################################

'// ���� ���� //
dim page,cdl,cdm, SearchString, strUse, lp

page = request("page")
SearchString = request("SearchString")
strUse = request("strUse")
if page = "" then page=1
if strUse = "" then strUse="Y"
cdl = request("cdl")
cdm = request("cdm")

dim ocate
set ocate = New CCategoryKeyWord
ocate.FCurrPage = page
ocate.FPageSize=20
ocate.FRectCDL = cdl
ocate.FRectCDM = cdm
ocate.FRectUsing = strUse
ocate.FRectSearch = SearchString

ocate.GetCaFavKeyWord

dim i
%>
<script language='javascript'>
<!--
function popItemWindow(iid,frm){
	window.open("/admin/pop/viewitemlist.asp?designerid=" + iid + "&target=" + frm, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
}

function ckAll(icomp){
	var bool = icomp.checked;
	var frm = document.frmBuyPrc;

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
	var frm = document.frmBuyPrc;

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

function changeUsing(upfrm){
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}
	
	if (upfrm.allusing.value=='Y'){
		var ret = confirm('���� �������� ��������� �����մϴ�');
	} else {
		var ret = confirm('���� �������� ���������� �����մϴ�');
	}

	if (ret) {
		upfrm.mode.value="changeUsing";
		upfrm.action="doCateTopKeyword.asp";
		upfrm.submit();
	} else {
		return;
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
		upfrm.action="doCateTopKeyword.asp";
		upfrm.submit();
	} else {
		return;
	}
}


function RefreshCaFavKeyWordRec(){
	if (document.refreshFrm.cdl.value==""){
		alert("������ ���Ͻô� ī�װ��� �������ּ���!!");
	}
	else{
	var popwin = window.open('','refreshFrm','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm";
		 refreshFrm.action = "<%=wwwUrl%>/chtml/make_category_TopKeyword_JS.asp";
		 refreshFrm.submit();
	}
}

function RefreshChannelKeyWordRec() {
	if (document.refreshFrm.cdl.value==""){
		alert("������ ���Ͻô� ��ī�װ��� �������ּ���!!");
	}
	else if (document.refreshFrm.cdm.value==""){
		alert("������ ���Ͻô� ��ī�װ��� �������ּ���!!");
	}
	else{
	var popwin = window.open('','refreshFrm','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm";
		 refreshFrm.action = "<%=wwwUrl%>/chtml/make_category_TopKeyword_JS.asp";
		 refreshFrm.submit();
	}
}

function frm_search()
{
	refreshFrm.target = "";
	refreshFrm.action = "category_left_topKeyword.asp";
}

	// ������ �̵�
	function goPage(pg)
	{
		document.refreshFrm.page.value=pg;
		document.refreshFrm.action="category_left_topKeyword.asp";
		document.refreshFrm.submit();
	}

// ī�װ� ����� ���
function changecontent() {
}
//-->
</script>
<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" onSubmit="frm_search()" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
		<td>
			<table width="100%" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td align="left">
					ī�װ� <% DrawSelectBoxCategoryLarge "cdl", cdl %>
					<% if cdl="110" then DrawSelectBoxCategoryMid "cdm", cdl, cdm %>
				</td>
				<td align="right">
					��뿩��
					<select class="select" name="strUse">
						<option value="all">��ü</option>
						<option value="Y">���</option>
						<option value="N">����</option>
					</select>
					/ Ű���� �˻�
					<input type="text" class="text" name="SearchString" size="12" value="<%=SearchString%>">
					<script language="javascript">
						document.refreshFrm.strUse.value="<%=strUse%>";
					</script>
				</td>
			</tr>
			</table>
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
<form name="frmBuyPrc" method="post" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="arrSort" value="">
<tr>
	<td>
		<%
			if cdl<>"" then
				if cdl<>"110" then
		%>
		<a href="javascript:RefreshCaFavKeyWordRec();"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html�����"></a>(�űԵ��, ���� �� ��!! ī�װ� ���� �� ���� ��ư�� �����ּ���)
		<%
				elseif cdm<>"" then
		%>
		<a href="javascript:RefreshChannelKeyWordRec();"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html�����"></a>(�űԵ��, ���� �� ��!! ī�װ� ���� �� ���� ��ư�� �����ּ���)
		<%
				end if
			end if
		%>
	</td>
	<td align="right">
		<select class="select" name="allusing">
			<option value="Y">���� -> Y</option>
			<option value="N">���� -> N</option>
		</select>
		<input type="button" class="button" value="����" onclick="changeUsing(frmBuyPrc);">
		/
		<input type="button" class="button" value="��������" onclick="changeSort(frmBuyPrc);">
		/
		<input type="button" value="������ �߰�" onclick="self.location='category_left_topKeyword_write.asp?menupos=<%= menupos %>'" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		�˻���� : <b><%=ocate.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=ocate.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>ī�װ�</td>
	<td>�̹���</td>
	<td>Ű����</td>
	<td>��ũ����</td>
	<td>�������</td>
	<td>����</td>
	<td>�����</td>
</tr>
<%	if ocate.FResultCount < 1 then %>
<tr>
	<td colspan="8" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� �������� �����ϴ�.</td>
</tr>
<%
	else
		for i=0 to ocate.FResultCount-1
%>
<tr align="center" bgcolor="<% if ocate.FItemList(i).Fisusing = "Y" then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">
	<td><input type="checkbox" name="selIdx" value="<%= ocate.FItemList(i).Fidx %>"></td>
	<td><%
		Response.Write ocate.FItemList(i).FCDL_Nm
		if Not(ocate.FItemList(i).FCDM_Nm="" or isNull(ocate.FItemList(i).FCDM_Nm)) then
			Response.Write "<br>/" & ocate.FItemList(i).FCDM_Nm
		end if
	%></td>
	<td>
	<% if Not(ocate.FItemList(i).FImageSmall="" or isNull(ocate.FItemList(i).FImageSmall)) then %>
		<img src="<%=ocate.FItemList(i).FImageSmall%>" border="0" width="50">
	<% else %>
		<img src="http://fiximage.10x10.co.kr/web2008/category/blank.gif" border="0" width="50">
	<% end if %>
	</td>
	<td><a href="category_left_topKeyword_write.asp?idx=<%= ocate.FItemList(i).Fidx %>&page=<%=page%>"><%= ocate.FItemList(i).Fkeyword %></a></td>
	<td align="left" style="word-break : break-all;">&nbsp;<a href="category_left_topKeyword_write.asp?idx=<%= ocate.FItemList(i).Fidx %>&page=<%=page%>"><%= ocate.FItemList(i).Flinkinfo %></a></td>
	<td><%=ocate.FItemList(i).Fisusing%></td>
	<td><input type="text" class="text" name="SortNo" value="<%=ocate.FItemList(i).FsortNo%>" size="2" style="text-align:center"></td>
	<td><%= FormatDate(ocate.FItemList(i).FRegdate,"0000.00.00") %></td>
</tr>
<%
		next
	end if
%>
<!-- ���� ��� �� -->
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
	<!-- ������ ���� -->
	<%
		if ocate.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & ocate.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + ocate.StartScrollPage to ocate.FScrollCount + ocate.StartScrollPage - 1

			if lp>ocate.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if ocate.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- ������ �� -->
	</td>
</tr>
</form>
</table>
<%
set ocate = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

