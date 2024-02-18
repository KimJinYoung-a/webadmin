<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  웨딩 엠디 픽
' History : 2018-04-18 정태훈 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/sitemaster/wedding/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/wedding_ContentsManageCls.asp" -->
<%

dim research,isusing, fixtype, linktype, poscode, validdate, prevDate, gubun, i, DateDiv
dim page,strParm
	isusing = request("isusing")
	research= request("research")
	poscode = request("poscode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")
	prevDate = request("prevDate")
	gubun = request("gubun")
	DateDiv = request("DateDiv")

	If gubun = "" Then
		gubun = "index"
	End If

	If DateDiv="" Then DateDiv="Y"

	if page="" then page=1

dim oMDPick
	set oMDPick = new CWeddingContents
	oMDPick.FPageSize = 20
	oMDPick.FCurrPage = page
	oMDPick.FRectIsusing = isusing
	oMDPick.FRectSelDate = prevDate
	oMDPick.FRectDateDiv = DateDiv
	oMDPick.GetMDPickList
%>
<script type="text/javascript">
<!--
function AddNewMainContents(idx){
    var popwin = window.open('/admin/sitemaster/wedding/popWeddingMDPickedit.asp?idx=' + idx+'&<%=strParm%>','mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}
function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}
function ckAll(icomp){
	var bool = icomp.checked;
	var frm = document.frmPrc;

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
	var frm = document.frmPrc;

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
		alert('선택아이템이 없습니다.');
		return;
	}

	if (confirm('선택 아이템을 사용안함으로 변경합니다')) {
		upfrm.mode.value="changeUsing";
		upfrm.action="doWeddingMDPickUpdate.asp";
		upfrm.submit();
	} else {
		return;
	}
}
function changeSort(upfrm){
	var arrSort="";
	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}

	if(confirm('선택 아이템에 입력하신 순서번호대로 저장합니다.')) {

		if(upfrm.selIdx.length) {
			for (var i=0;i<upfrm.selIdx.length;i++){
				if(upfrm.selIdx[i].checked) arrSort = arrSort + upfrm.DispOrder[i].value + ",";
			}
		} else {
			if(upfrm.selIdx.checked) arrSort=upfrm.DispOrder.value;
		}
		upfrm.arrSort.value = arrSort;

		upfrm.mode.value="changeSort";
		upfrm.action="doWeddingMDPickUpdate.asp";
		upfrm.submit();
	} else {
		return;
	}
}
function AddMultiMainContents(idx){
    var popwin = window.open('/admin/sitemaster/wedding/popWeddingMDPickMulti.asp','mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}
//-->
</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />


<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="left"><input type="button" class="button" value="선택삭제" onClick="changeUsing(frmPrc);">&nbsp;<input type="button" class="button" value="선택순서변경" onClick="changeSort(frmPrc);"></td>
    <td align="right">
    	<input type="button" class="button" value="일괄 등록" onClick="AddMultiMainContents(frmPrc);">&nbsp;<a href="javascript:AddNewMainContents('0');"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmPrc" method="post" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="arrSort" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10">
		검색결과 : <b><%=oMDPick.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oMDPick.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
    <td>상품이미지</td>
	<td>상품명</td>
	<td>순번</td>
    <td>등록자</td>
	<td>등록일</td>
	<td>사용유무</td>
</tr>
<%
	for i=0 to oMDPick.FResultCount - 1
%>
<% If oMDPick.FItemList(i).FIsusing="N" Then %>
<tr bgcolor="#e5e5e5">
<% Else %>
<tr bgcolor="#FFFFFF">
<% End If %>
	<td align="center"><input type="checkbox" name="selIdx" value="<%= oMDPick.FItemList(i).Fidx %>"></td>
    <td align="center" onClick="AddNewMainContents('<%= oMDPick.FItemList(i).FIdx %>');"  style="cursor:pointer;"><img src="<%= oMDPick.FItemList(i).Fsmallimage %>" border="0"></td>
	<td align="center" onClick="AddNewMainContents('<%= oMDPick.FItemList(i).FIdx %>');"  style="cursor:pointer;"><%= oMDPick.FItemList(i).Fitemname %></td>
	<td align="center"><input type="text" name="DispOrder" value="<%= oMDPick.FItemList(i).FDispOrder %>" size="3"></td>
    <td align="center"><%= oMDPick.FItemList(i).FLastUser %></td>
	<td align="center"><%= oMDPick.FItemList(i).FRegdate %></td>
	<td align="center"><%= oMDPick.FItemList(i).FIsusing %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="12" align="center" height="30">
    <% if oMDPick.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oMDPick.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oMDPick.StarScrollPage to oMDPick.FScrollCount + oMDPick.StarScrollPage - 1 %>
		<% if i>oMDPick.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oMDPick.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</form>
</table>

<%
set oMDPick = Nothing
%>

<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->