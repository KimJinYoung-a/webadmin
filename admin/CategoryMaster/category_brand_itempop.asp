<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/category_contents_managecls.asp"-->
<%
dim page, cdl, cdm, isusing, vIdx, vDisp1
cdl = request("cdl")
cdm = request("cdm")
page = request("page")
isusing = request("isusing")
vIdx = request("idx")

if page="" then page=1


dim omd
set omd = New CCateContents
omd.FCurrPage = page
omd.FPageSize=20
omd.FRectIdx = vIdx
omd.FRectIsUsing = isusing
omd.GetBrandItemList

dim i
%>
<script language='javascript'>
<!--
function popItemWindow(tgf){
	var popup_item = window.open("/common/pop_CateItemList.asp?target=" + tgf, "popup_item", "width=1000,height=800,scrollbars=yes,resizable=yes");
	popup_item.focus();
}

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];

		if (frm.name.indexOf('frmBuyPrc')!= -1) {

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

function changeUsing(upfrm){
	if (!CheckSelected()){
		alert('상품을 선택해 주세요.');
		return;
	}

	if (upfrm.allusing.value=='Y'){
		var ret = confirm('선택하신 상품을 사용함 으로 변경합니다.');
	} else {
		var ret = confirm('선택하신 상품을 사용안함 으로 변경합니다.');
	}


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
		upfrm.mode.value="isUsingValue";
		upfrm.submit();

	}
}

// 순서적용
function changeSort(upfrm) {
	if (!CheckSelected()){
		alert('상품을 선택해 주세요.');
		return;
	}
	var ret = confirm('선택하신 상품의 순서를 지정하신 번호로 변경하시겠습니까?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + frm.itemid.value + "," ;
					upfrm.sortNo.value = upfrm.sortNo.value + frm.sortNo.value + "," ;
				}
			}
		}
		upfrm.mode.value="ChangeSort";
		upfrm.submit();

	}
}

function AddIttems(){
	var ret = confirm(arrFrm.itemid.value + '아이템을 추가하시겠습니까?');
	if (ret){
		arrFrm.itemidarr.value = arrFrm.itemid.value;
		arrFrm.submit();
	}
}

function AddIttems2(){
	if (document.arrFrm.itemidarr.value == ""){
		alert("아이템번호를  적어주세요!");
		document.arrFrm.itemidarr.focus();
	}
	else if (confirm(arrFrm.itemidarr.value + '아이템을 추가하시겠습니까?')){
		arrFrm.itemid.value = arrFrm.itemidarr.value;
		arrFrm.submit();
	}
}

//-->
</script>
<!-- 상단 검색폼 시작 -->
<form name="Listfrm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="idx" value="<%=vIdx%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		사용유무 :
		<select name="isusing" onchange="this.form.submit();" class="select">
			<option value=""  <% if isusing="" then response.write "selected" %>>전체</option>
			<option value="Y" <% if isusing="Y" then response.write "selected" %>>사용</option>
			<option value="N" <% if isusing="N" then response.write "selected" %>>사용안함</option>
		</select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td>
		<form name="arrFrm" method="post" action="doCateBrand.asp" style="margin:0px;">
		<input type="hidden" name="cdl">
		<input type="hidden" name="cdm">
		<input type="hidden" name="mode" value="add">
		<input type="hidden" name="itemid">
		<input type="hidden" name="sortNo">
		<input type="hidden" name="idx" value="<%=vIdx%>">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td colspan="2" align="right">
				<input type="text" name="itemidarr" value="" size="80" class="input" onKeyPress="if (event.keyCode == 13) arrFrm.itemid.value = arrFrm.itemidarr.value;">
				<input type="button" value="아이템 직접 추가" onclick="AddIttems2()" class="button">
			</td>
		</tr>
		<tr>
			<td>
				<input type="button" value="선택아이템 삭제" onClick="delitems(arrFrm)" class="button"> /
				<select name="allusing"  class="select">
					<option value="Y">선택 사용 -> Y </option>
					<option value="N">선택 사용 -> N </option>
				</select><input type="button" value="적용" class="button" onclick="changeUsing(arrFrm);"> /
				<input type="button" value="순서적용" class="button" onclick="changeSort(arrFrm);">
			</td>
			<td align="right"><input type="button" value="아이템 추가" onclick="popItemWindow('arrFrm.itemid')" class="button"></td>
		</tr>
		</table>
		</form>
	</td>
</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FAFAFA" height="22">
	<td colspan="8">&nbsp;검색된 상품수 : <%=omd.FTotalCount%> 건</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">카테고리명</td>
	<td align="center">ItemID</td>
	<td align="center">Image</td>
	<td align="center">제품명</td>
	<td align="center">순서</td>
	<td align="center">사용유무</td>
	<td align="center">품절유무</td>
</tr>
<% for i=0 to omd.FResultCount-1 %>
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="itemid" value="<%= omd.FItemList(i).FItemID %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center"><%= omd.FItemList(i).Fcode_nm %>&nbsp;<%= omd.FItemList(i).Fmidcode_nm %></td>
	<td align="center"><img src="<%= omd.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= omd.FItemList(i).FItemID %></td>
	<td align="center"><%= omd.FItemList(i).FItemname %></td>
	<td align="center"><input type="text" name="sortNo" value="<%= omd.FItemList(i).FsortNo %>" size="3" style="text-align:right;" /></td>
	<td align="center"><%= omd.FItemList(i).Fisusing %></td>
	<td align="center">
		<% if omd.FItemList(i).IsSoldOut then %>
		<font color="red">품절</font>
		<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
	<% if omd.HasPreScroll then %>
		<a href="?page=<%= omd.StarScrollPage-1 %>&cdl=<%=cdl%>&idx=<%=vIdx%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + omd.StarScrollPage to omd.FScrollCount + omd.StarScrollPage - 1 %>
		<% if i>omd.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&cdl=<%=cdl%>&idx=<%=vIdx%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if omd.HasNextScroll then %>
		<a href="?page=<%= i %>&cdl=<%=cdl%>&idx=<%=vIdx%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set omd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
