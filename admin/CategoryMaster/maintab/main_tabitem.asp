<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.04.18 한용민 카테고리md픽 이동/ 추가/수정
'	Description : 메인페이지 탭관리
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_tabitem_cls.asp"-->
<%
dim page, cdl, isusing
	cdl = request("cdl")
	page = request("page")
	isusing = request("isusing")
	
	if page="" then page=1


dim oip
	set oip = New Cmain_tabitem_list
	oip.FCurrPage = page
	oip.FPageSize=20
	oip.FRectCDL = cdl
	oip.FRectIsUsing = isusing
	oip.Getmain_tabitem

dim i
%>
<script language='javascript'>

	function popItemWindow(tgf){
		if (document.Listfrm.cdl.value == ""){
			alert("카테고리를 선택해 주세요!");
			document.Listfrm.cdl.focus();
		}
		else{
			var popup_item = window.open("/common/pop_CateItemList.asp?cdl=" + document.refreshFrm.cdl.value + "&target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
			popup_item.focus();
		}
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
			upfrm.cdl.value = Listfrm.cdl.value;
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
			upfrm.cdl.value = Listfrm.cdl.value;
			upfrm.mode.value="ChangeSort";
			upfrm.submit();
	
		}
	}
	
	function AddIttems(){
		var ret = confirm(arrFrm.itemid.value + '아이템을 추가하시겠습니까?');
		if (ret){
			arrFrm.itemid.value = arrFrm.itemid.value;
			arrFrm.cdl.value = Listfrm.cdl.value;
			arrFrm.mode.value="add";
			arrFrm.submit();
		}
	}
	
	function AddIttems2(){
		if (document.Listfrm.cdl.value == ""){
			alert("카테고리를 선택해 주세요!");
			document.Listfrm.cdl.focus();
		}
		else if (document.arrFrm.itemidarr.value == ""){
			alert("아이템번호를  적어주세요!");
			document.arrFrm.itemidarr.focus();
		}
		else if (confirm(arrFrm.itemidarr.value + '아이템을 추가하시겠습니까?')){
			arrFrm.itemid.value = arrFrm.itemidarr.value;
			arrFrm.cdl.value = Listfrm.cdl.value;
			arrFrm.mode.value="add";
			arrFrm.submit();
		}
	}

	function RefreshCaMDChoiceRec(){
		if (document.Listfrm.cdl.value == ""){
			alert("카테고리를 선택해 주세요!");
			document.Listfrm.cdl.focus();
		}
		 else{
				  var popwin = window.open('','refreshFrm','');
				  popwin.focus();
				  refreshFrm.target = "refreshFrm";
				  refreshFrm.cdl.value = document.Listfrm.cdl.value;
				  refreshFrm.action = "<%=wwwUrl%>/chtml/make_main_tabitem.asp";
				  refreshFrm.submit();
		 }
	}

	// 카테고리 변경시 명령
	function changecontent(){}

</script>
<form name="refreshFrm" method="post">
<input type="hidden" name="cdl">
</form>
<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="Listfrm" method="get" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		탭선택 :
		<select name='cdl' class="select">
			<option value="">선택하세요</option>
			<option value=1 <% if cdl = "1" then response.write " selected" %>>디자인/오피스</option>
			<option value=2 <% if cdl = "2" then response.write " selected" %>>키덜트/취미</option>
			<option value=3 <% if cdl = "3" then response.write " selected" %>>리빙</option>
			<option value=4 <% if cdl = "4" then response.write " selected" %>>패션</option>
			<option value=5 <% if cdl = "5" then response.write " selected" %>>베이비/키즈</option>
			<option value=6 <% if cdl = "6" then response.write " selected" %>>감성채널</option>
		</select>
		&nbsp;/&nbsp;
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
				<img src="/images/icon_reload.gif" onClick="RefreshCaMDChoiceRec()" style="cursor:pointer" align="absmiddle" alt="html만들기">
				프론트에 적용
			</td>
		</tr>
		</form>
		<form name="arrFrm" method="post" action="domaintabitem.asp">
		<input type="hidden" name="cdl">
		<input type="hidden" name="mode">
		<input type="hidden" name="itemid">
		<input type="hidden" name="sortNo">
		<tr>
			<td colspan="2" align="right">
				<input type="text" name="itemidarr" value="" size="80" class="input">
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
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FAFAFA" height="22">
	<td colspan="8">&nbsp;검색된 상품수 : <%=oip.FTotalCount%> 건</td>
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
<% for i=0 to oip.FResultCount-1 %>
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="itemid" value="<%= oip.FItemList(i).FItemID %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center">
		<% if  cStr(oip.FItemList(i).Fcdl) = "1" then
		response.write "디자인/오피스"
		elseif cStr(oip.FItemList(i).Fcdl) = "2" then
		response.write "키덜트/취미"
		elseif 	cStr(oip.FItemList(i).Fcdl) = "3" then
		response.write "리빙"
		elseif 	cStr(oip.FItemList(i).Fcdl) = "4" then
		response.write "패션"
		elseif 	cStr(oip.FItemList(i).Fcdl) = "5" then
		response.write "베이비/키즈"
		elseif 	cStr(oip.FItemList(i).Fcdl) = "6" then
		response.write "감성채널"
		end if
		%>
	</td>
	<td align="center"><img src="<%= oip.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= oip.FItemList(i).FItemID %></td>
	<td align="center"><%= oip.FItemList(i).FItemname %></td>
	<td align="center"><input type="text" name="sortNo" value="<%= oip.FItemList(i).FsortNo %>" size="3" style="text-align:right;" /></td>
	<td align="center"><%= oip.FItemList(i).Fisusing %></td>
	<td align="center">
		<% if oip.FItemList(i).IsSoldOut then %>
		<font color="red">품절</font>
		<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
	<% if oip.HasPreScroll then %>
		<a href="?page=<%= oip.StarScrollPage-1 %>&cdl=<%=cdl%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oip.StarScrollPage to oip.FScrollCount + oip.StarScrollPage - 1 %>
		<% if i>oip.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&cdl=<%=cdl%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oip.HasNextScroll then %>
		<a href="?page=<%= i %>&cdl=<%=cdl%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set oip = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
