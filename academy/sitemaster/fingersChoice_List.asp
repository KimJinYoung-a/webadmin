<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/sitemaster/fingersChoiceCls.asp"-->
<%
dim page, MenuId, isusing
MenuId = RequestCheckvar(request("MenuId"),6)
page = RequestCheckvar(request("page"),10)
isusing = RequestCheckvar(request("isusing"),1)

if page="" then page=1


dim oFingers
set oFingers = New CFingersChoice
oFingers.FCurrPage = page
oFingers.FPageSize=21
oFingers.FRectMenuId = MenuId
oFingers.FRectIsUsing = isusing
oFingers.GetFingersChoiceList

dim i
%>
<script language='javascript'>
<!--
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
		alert('선택강좌가 없습니다.');
		return;
	}

	var ret = confirm('선택 강좌를 삭제하시겠습니까?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.lec_idx.value = upfrm.lec_idx.value + frm.lec_idx.value + "," ;
				}
			}
		}
		upfrm.mode.value="del";
		upfrm.submit();

	}
}

function changeUsing(upfrm){
	if (!CheckSelected()){
		alert('강좌를 선택해 주세요.');
		return;
	}

	if (upfrm.allusing.value=='Y'){
		var ret = confirm('선택하신 강좌를 사용함 으로 변경합니다.');
	} else {
		var ret = confirm('선택하신 강좌를 사용안함 으로 변경합니다.');
	}


	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.lec_idx.value = upfrm.lec_idx.value + frm.lec_idx.value + "," ;
				}
			}
		}
		upfrm.MenuId.value = Listfrm.MenuId.value;
		upfrm.mode.value="isUsingValue";
		upfrm.submit();

	}
}

// 순서적용
function changeSort(upfrm) {
	if (!CheckSelected()){
		alert('강좌를 선택해 주세요.');
		return;
	}
	var ret = confirm('선택하신 강좌의 순서를 지정하신 번호로 변경하시겠습니까?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.lec_idx.value = upfrm.lec_idx.value + frm.lec_idx.value + "," ;
					upfrm.sortNo.value = upfrm.sortNo.value + frm.sortNo.value + "," ;
				}
			}
		}
		upfrm.MenuId.value = Listfrm.MenuId.value;
		upfrm.mode.value="ChangeSort";
		upfrm.submit();

	}
}

function AddIttems(){
	var ret = confirm(arrFrm.lec_idx.value + '강좌를 추가하시겠습니까?');
	if (ret){
		arrFrm.lec_idx.value = arrFrm.lec_idx.value;
		arrFrm.MenuId.value = Listfrm.MenuId.value;
		arrFrm.mode.value="add";
		arrFrm.submit();
	}
}

function AddIttems2(){
	if (document.Listfrm.MenuId.value == ""){
		alert("입력할 주메뉴를 선택해 주세요!");
		document.Listfrm.MenuId.focus();
	}
	else if (document.arrFrm.lecIdxarr.value == ""){
		alert("강좌번호를  적어주세요!");
		document.arrFrm.lecIdxarr.focus();
	}
	else if (confirm(arrFrm.lecIdxarr.value + '강좌를 추가하시겠습니까?')){
		arrFrm.lec_idx.value = arrFrm.lecIdxarr.value;
		arrFrm.MenuId.value = Listfrm.MenuId.value;
		arrFrm.mode.value="add";
		arrFrm.submit();
	}
}

function RefreshCaFingersChoiceRec(){
	if (document.Listfrm.MenuId.value == ""){
		alert("주메뉴를 선택해 주세요!");
		document.Listfrm.MenuId.focus();
	}
	 else{
			  var popwin = window.open('','refreshFrm','');
			  popwin.focus();
			  refreshFrm.target = "refreshFrm";
			  refreshFrm.MenuId.value = document.Listfrm.MenuId.value;
			  refreshFrm.action = "<%=wwwFingers%>/chtml/make_FingersChoice_JS.asp";
			  refreshFrm.submit();
	 }
}

// 주메뉴 변경시 명령
function changecontent(){}

//-->
</script>
<form name="refreshFrm" method="post">
<input type="hidden" name="MenuId">
</form>
<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="Listfrm" method="get" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		주메뉴 :
		<Select name="MenuId" Class="select">
			<option value="">선택</option>
			<option value="10" <% if MenuId="10" then Response.Write "selected"%>>강좌전체</option>
			<option value="20" <% if MenuId="20" then Response.Write "selected"%>>원데이 클래스</option>
			<option value="30" <% if MenuId="30" then Response.Write "selected"%>>위클리 클래스</option>
			<option value="40" <% if MenuId="40" then Response.Write "selected"%>>스튜디오워크샾</option>			
		</select>
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
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td colspan="2">
				<img src="/images/icon_reload.gif" onClick="RefreshCaFingersChoiceRec()" style="cursor:pointer" align="absmiddle" alt="html만들기">
				프론트에 적용
			</td>
		</tr>
		</form>
		<form name="arrFrm" method="post" action="doFingersChoice.asp">
		<input type="hidden" name="MenuId">
		<input type="hidden" name="mode">
		<input type="hidden" name="lec_idx">
		<input type="hidden" name="sortNo">
		<tr>
			<td>
				<input type="button" value="선택강좌 삭제" onClick="delitems(arrFrm)" class="button"> /
				<select name="allusing"  class="select">
					<option value="Y">선택 사용 -> Y </option>
					<option value="N">선택 사용 -> N </option>
				</select><input type="button" value="적용" class="button" onclick="changeUsing(arrFrm);"> /
				<input type="button" value="순서적용" class="button" onclick="changeSort(arrFrm);">
			</td>
			<td align="right">
				<input type="text" name="lecIdxarr" value="" size="50" class="input">
				<input type="button" value="강좌 추가" onclick="AddIttems2()" class="button">
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FAFAFA" height="22">
	<td colspan="8">&nbsp;검색된 강좌수 : <%=oFingers.FTotalCount%> 건</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">주메뉴명</td>
	<td align="center">Image</td>
	<td align="center">강좌번호</td>
	<td align="center">강좌명</td>
	<td align="center">순서</td>
	<td align="center">사용유무</td>
	<td align="center">마감유무</td>
</tr>
<% for i=0 to oFingers.FResultCount-1 %>
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="lec_idx" value="<%= oFingers.FItemList(i).Flec_idx %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center"><%= getLecMenuName(oFingers.FItemList(i).FMenuId) %></td>
	<td align="center"><img src="<%= oFingers.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= oFingers.FItemList(i).Flec_idx %></td>
	<td align="center"><%= oFingers.FItemList(i).Flec_title %></td>
	<td align="center"><input type="text" name="sortNo" value="<%= oFingers.FItemList(i).FsortNo %>" size="3" style="text-align:right;" /></td>
	<td align="center"><%= oFingers.FItemList(i).Fisusing %></td>
	<td align="center">
		<% if oFingers.FItemList(i).IsSoldOut then %>
		<font color="red">마감</font>
		<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
	<% if oFingers.HasPreScroll then %>
		<a href="?page=<%= oFingers.StarScrollPage-1 %>&MenuId=<%=MenuId%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oFingers.StarScrollPage to oFingers.FScrollCount + oFingers.StarScrollPage - 1 %>
		<% if i>oFingers.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&MenuId=<%=MenuId%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oFingers.HasNextScroll then %>
		<a href="?page=<%= i %>&MenuId=<%=MenuId%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set oFingers = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyClose.asp" -->
