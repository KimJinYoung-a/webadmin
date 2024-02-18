<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/categoryCls.asp"-->
<%
'###############################################
' PageName : Category_left_Bigchance.asp
' Discription : ī�װ� ���� ������ ���
' History : 2008.03.31 ������ : ����
'           2008.07.25 ������ ���� : ��ǰ ���ļ��� �߰�
'###############################################

dim cdl, cdm, page, lp
cdl = request("cdl")
cdm = request("cdm")
page = request("page")

if page="" then page=1

dim omd
set omd = New CMDSRecommend
omd.FCurrPage = page
omd.FPageSize=20
omd.FRectCDL = cdl
omd.FRectCDM = cdm
omd.GetCategoryBigChanceList

dim i
%>
<script language='javascript'>
<!--
function popItemWindow(tgf){
	<% if cdl<>"110" then %>
		if (document.refreshFrm.cdl.value == ""){
			alert("ī�װ��� ������ �ּ���!");
			document.refreshFrm.cdl.focus();
		} else if (document.refreshFrm.cdl.value=="110") {
			alert("����ä���� �˻��� �����Ͽ� ��ī�װ��� �����ϼž��մϴ�.");
		} else {
			var popup_item = window.open("/common/pop_CateItemList.asp?cdl=" + document.refreshFrm.cdl.value + "&target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
			popup_item.focus();
		}
	<% else %>
		if (document.refreshFrm.cdl.value == ""){
			alert("ī�װ��� ������ �ּ���!");
			document.refreshFrm.cdl.focus();
		} else if (document.refreshFrm.cdm.value == ""){
			alert("��ī�װ��� ������ �ּ���!");
			document.refreshFrm.cdm.focus();
		} else {
			var popup_item = window.open("/common/pop_CateItemList.asp?cdl=" + document.refreshFrm.cdl.value + "&cdm=" + document.refreshFrm.cdm.value + "&target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
			popup_item.focus();
		}
	<% end if %>
}

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
			if(frm.selIdx[i].checked) frm.arrSort.value = frm.arrSort.value + frm.sortNo[i].value + ",";
		}
	} else {
		pass = ((pass)||(frm.selIdx.checked));
		frm.arrSort.value = frm.sortNo.value;
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
		upfrm.action="doCategoryLeftBigchance.asp";
		upfrm.submit();
	}
}

// ���þ������� ���Ĺ�ȣ ����(2008.07.25; ������ �߰�)
function submitSortNo(upfrm) {
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}

	if (confirm('���� �������� ���Ĺ�ȣ�� �����Ͻðڽ��ϱ�?')) {
		upfrm.mode.value="sort";
		upfrm.action="doCategoryLeftBigchance.asp";
		upfrm.submit();
	}	
}


function AddIttems(){
	<% if cdl<>"110" then %>
		if (document.refreshFrm.cdl.value == ""){
			alert("ī�װ��� �������ּ���");
			document.refreshFrm.cdl.focus();
		} else if (document.refreshFrm.cdl.value == "110"){
			alert("����ä���� �˻��� �����Ͽ� ��ī�װ��� �����ϼž��մϴ�.");
		} else if (confirm(frmarr.itemidarr.value + '�������� �߰��Ͻðڽ��ϱ�?')){
			frmarr.itemid.value = frmarr.itemidarr.value;
			frmarr.cdl.value = refreshFrm.cdl.value;
			frmarr.mode.value="add";
			frmarr.submit();
		}
	<% else %>
		if (document.refreshFrm.cdl.value == ""){
			alert("ī�װ��� �������ּ���");
			document.refreshFrm.cdl.focus();
		} else if(document.refreshFrm.cdm.value == ""){
			alert("��ī�װ��� �������ּ���");
			document.refreshFrm.cdm.focus();
		} else if (confirm(frmarr.itemidarr.value + '�������� �߰��Ͻðڽ��ϱ�?')){
			frmarr.itemid.value = frmarr.itemidarr.value;
			frmarr.cdl.value = refreshFrm.cdl.value;
			frmarr.cdm.value = refreshFrm.cdm.value;
			frmarr.mode.value="add";
			frmarr.submit();
		}
	<% end if %>
}

function RefreshMainRotateEventRec(){
	<% if cdl<>"110" then %>
		if (document.refreshFrm.cdl.value == ""){
			alert("ī�װ��� �������ּ���");
			document.refreshFrm.cdl.focus();
		} else if (document.refreshFrm.cdl.value == "110"){
			alert("����ä���� �˻��� �����Ͽ� ��ī�װ��� �����ϼž��մϴ�.");
		} else{
			 var popwin = window.open('','refreshPop','');
			 popwin.focus();
			 refreshFrm.target = "refreshPop";
			 refreshFrm.action = "<%=wwwUrl%>/chtml/make_category_left_bigchance_JS.asp";
			 refreshFrm.submit();
		}
	<% else %>
		if (document.refreshFrm.cdl.value == ""){
			alert("ī�װ��� �������ּ���");
			document.refreshFrm.cdl.focus();
		} else if (document.refreshFrm.cdm.value == ""){
			alert("��ī�װ��� �������ּ���");
			document.refreshFrm.cdm.focus();
		} else {
			 var popwin = window.open('','refreshPop','');
			 popwin.focus();
			 refreshFrm.target = "refreshPop";
			 refreshFrm.action = "<%=wwwUrl%>/chtml/make_channel_left_bigchance_JS.asp";
			 refreshFrm.submit();
		}
	<% end if %>
	location.reload();
}

// ������ �̵�
function goPage(pg)
{
	document.refreshFrm.page.value=pg;
	document.refreshFrm.action="category_left_Bigchance.asp";
	document.refreshFrm.submit();
}

// ī�װ� ����� ���
function changecontent(){}
//-->
</script>
<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" action="category_left_Bigchance.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		ī�װ� <% DrawSelectBoxCategoryLarge "cdl", cdl %>
		<% if cdl="110" then DrawSelectBoxCategoryMid "cdm", cdl, cdm %>
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
<form name="frmarr" method="post" action="doCategoryLeftBigchance.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="cdl" value="">
<input type="hidden" name="cdm" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="arrSort" value="">
<tr>
	<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td>
				<input type="text" name="itemidarr" value="" size="80" class="text">
				<input type="button" value="������ ���� �߰�" onclick="AddIttems()" class="button">
			</td>
			<td align="right">
				<img src="/images/icon_reload.gif" onClick="RefreshMainRotateEventRec()" style="cursor:pointer" align="absmiddle" alt="html�����">
				����Ʈ�� ����
			</td>
		</tr>
		<tr>
			<td><input type="button" value="���þ����� ����" onClick="delitems(frmarr)" class="button"></td>
			<td align="right">
				<input type="button" value="���Ĺ�ȣ ����" onClick="submitSortNo(frmarr)" class="button">
				<input type="button" value="������ �߰�" onclick="popItemWindow('frmarr.itemidarr')" class="button">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<!-- �׼� �� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		�˻���� : <b><%=omd.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=omd.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>ī�װ���</td>
	<td>Image</td>
	<td>ItemID</td>
	<td>��ǰ��</td>
	<td>���η�</td>
	<td>����</td>
	<td>���Ĺ�ȣ</td>
</tr>
<%	if omd.FResultCount < 1 then %>
<tr>
	<td colspan="8" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� �������� �����ϴ�.</td>
</tr>
<%
	else
		for i=0 to omd.FResultCount-1
%>
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="selIdx" value="<%= omd.FItemList(i).Fidx %>"></td>
	<td align="center"><%
		Response.Write omd.FItemList(i).Fcode_nm
		if Not(omd.FItemList(i).FCDM_Nm="" or isNull(omd.FItemList(i).FCDM_Nm)) then
			Response.Write "<br>/" & omd.FItemList(i).FCDM_Nm
		end if
	%>
	</td>
	<td align="center"><img src="<%= omd.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= omd.FItemList(i).FItemID %></td>
	<td align="center"><%= omd.FItemList(i).FItemname %></td>
	<td align="center"><% if omd.FItemList(i).FsailYn="Y" then Response.Write formatPercent(1-omd.FItemList(i).FsailPrice/omd.FItemList(i).ForgPrice,1) %></td>
	<td align="center"><% if omd.FItemList(i).FsellYn<>"Y" then Response.Write "ǰ��" %></td>
	<td align="center"><input type="text" name="sortNo" value="<%=omd.FItemList(i).FsortNo %>" size="3" style="text-align:right"></td>
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
</form>
</table>
<%
set omd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
