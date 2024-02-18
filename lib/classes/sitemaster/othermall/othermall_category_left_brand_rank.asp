<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2007.11.09 한용민 개발
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/othermall_idx_mdchoice_brandcls.asp"-->
<%
dim cdl, page
cdl = request("cdl")
page = request("page")

if page="" then page=1

dim omd
set omd = New MDChoice
omd.FCurrPage = page
omd.FPageSize=100
omd.FRectCDL = cdl
omd.GetCategoryLeftBrandRank

dim i
%>

<script language='javascript'>

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function delitems(upfrm){
	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}

	var ret = confirm('선택 아이템을 삭제하시겠습니까?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + frm.itemid.value + "," ;
				}
			}
		}

		upfrm.mode.value="del";
		upfrm.submit();

	}
}

function RefreshBestBrand(upfrm){

	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}

	var ret = confirm('10개까지 저장됩니다. 선택 아이템을 적용 하시겠습니까?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + frm.itemid.value + "," ;
				}
			}
		}

		//upfrm.mode.value="del";
		upfrm.action = "<%=othermall%>/chtml/othermall_make_best_friend.asp"
		upfrm.submit();

	}
}

function changecontent(){
    // nothing
}
</script>

<!--표 헤드시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>외부몰 베스트 브랜드</strong></font>
			</td>		
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<% DrawSelectBoxCategoryLarge "cdl", cdl %>&nbsp;
			<a href="javascript:document.frm.submit();">
			<img src="/admin/images/search2.gif" width="74" height="22" border="0" align="absmiddle"></a>
		    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		    리얼적용 시킬  아이템선택후 <a href="javascript:RefreshBestBrand(refreshFrm);">
		    <img src="/images/refreshcpage.gif" width="19" align="absmiddle" border="0"></a> 버튼을 눌러주세요
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr height="10" valign="top">
		<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_06.gif"></td>
		<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
	</tr>
	</tr>
	</form>
</table>
<!--표 헤드끝-->

<table border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC" width="100%">
	<tr bgcolor="#DDDDFF" height="25">
		<td width="50" align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td width="150" align="center">idx</td>
		<td width="150" align="center">카테고리명</td>
		<td width="200" align="center">업체명</td>
		<td width="150" align="center">이미지</td>
	</tr>
	<% for i=0 to omd.FResultCount-1 %>
		<form name="frmBuyPrc_<%=i%>" method="post" action="" >
		<input type="hidden" name="itemid" value="<%= omd.FItemList(i).Fidx %>">
		<tr bgcolor="#FFFFFF">
			<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
			<td align="center"><%= omd.FItemList(i).Fidx %></td>
			<td align="center"><%= omd.FItemList(i).GetCD1Name %></td>
			<td align="center"><%= omd.FItemList(i).Fmakerid %></td>
			<td align="center"><img src="<%= omd.FItemList(i).FImgSmall %>"><img src="<%= omd.FItemList(i).Ftitleimgurl %>" ></td>
		</tr>
		</form>
	<% next %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
			<% if omd.HasPreScroll then %>
				<a href="?page=<%= omd.StarScrollPage-1 %>&menupos=<%= menupos %>&cdl=<%=cdl%>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
		
			<% for i=0 + omd.StarScrollPage to omd.FScrollCount + omd.StarScrollPage - 1 %>
				<% if i>omd.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>&cdl=<%=cdl%>">[<%= i %>]</a>
				<% end if %>
			<% next %>
		
			<% if omd.HasNextScroll then %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>&cdl=<%=cdl%>">[next]</a>
			<% else %>
				[next]
			<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<form name="delform" method="post" action="doleftbrandrank.asp">
<input type="hidden" name="mode">
<input type="hidden" name="itemid">
</form>
<form name="refreshFrm" method=post>
<input type="hidden" name="cdl">
<input type="hidden" name="itemid">
</form>
<%
set omd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->