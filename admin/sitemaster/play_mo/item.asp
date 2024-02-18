<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/play/play_moCls.asp" -->

<%
dim page, isusing, oitem, i, itemid, itemname, playidx, playcate
	page		= request("page")
	isusing		= request("isusing")
	itemid = request("itemid")
	itemname = request("itemname")
	playidx = request("playidx")
	playcate = request("playcate")
	
	if page="" then page=1
	if isusing = "" then isusing = "Y"
	
set oitem = New CPlayMoContents
	oitem.FCurrPage = page
	oitem.FPageSize=20
	oitem.FRectPlayIdx = playidx
	oitem.FRectitemid = itemid
	oitem.FRectitemname = itemname
	oitem.FRectIsUsing = isusing
	oitem.fnPlayItemList
%>

<script>
//사용여부 수정
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
					upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
				}
			}
		}

		upfrm.mode.value="chisusing";
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
					upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
					upfrm.orderno.value = upfrm.orderno.value + frm.orderno.value + "," ;
				}
			}
		}
		upfrm.mode.value="ChangeSort";
		upfrm.submit();

	}
}

//상품삭제
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
					upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
				}
			}
		}
		upfrm.mode.value="delitem";
		upfrm.submit();

	}
}

//상품 검색 팝업 에서 상품 선택후 다음 작업
function AddIttems(){
	var ret = confirm(arrFrm.itemid.value + '아이템을 추가하시겠습니까?');
	if (ret){
		arrFrm.itemid.value = arrFrm.itemid.value;
		arrFrm.mode.value="itemadd";
		arrFrm.submit();
	}
}

//상품 검색 팝업
function popItemWindow(tgf){
	
	if (document.Listfrm.playidx.value == ""){
		alert("PLAY idx값이 없습니다.\n창을 닫고 다시 등록해 주세요!");
		return;
	}

	var popup_item = window.open("/common/pop_CateItemList.asp?target=" + tgf, "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
	popup_item.focus();
}

//상품 수동 직접 추가
function AddIttems2(){

	if (document.arrFrm.itemidarr.value == ""){
		alert("아이템번호를  적어주세요!");
		return;
		document.arrFrm.itemidarr.focus();
	}
	if (confirm(arrFrm.itemidarr.value + '아이템을 추가하시겠습니까?')){
		arrFrm.itemid.value = arrFrm.itemidarr.value;
		arrFrm.mode.value="itemadd";
		arrFrm.submit();
	}
}

//검색
function jsSerach(ipage){
	var frm;
	frm = document.Listfrm;
	
	if(frm.itemid.value!=''){
		if (!IsDouble(frm.itemid.value)){
			alert('상품 코드는 숫자만 가능합니다.');
			frm.itemid.focus();
			return;
		}
	}

	frm.page.value= ipage;
	frm.submit();
}

//선택시 tr 색 변함
function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="Listfrm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="playidx" value="<%=playidx%>">
<input type="hidden" name="playcate" value="<%=playcate%>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		사용 : <% drawSelectBoxUsingYN "isusing", isusing %>
		상품코드 : <input type="text" name="itemid" value="<%= itemid %>" size=10 maxlength=10>
		상품명 : <input type="text" name="itemname" value="<%= itemname %>" size=30 maxlength=30>
	</td>	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSerach('');">
	</td>
</tr>
</form>
</table>
<br>

<table width="100%" align="center" cellpadding="2" cellspacing="0" class="a">
<form name="arrFrm" method="post" action="item_proc.asp">
<input type="hidden" name="mode">
<input type="hidden" name="itemid">
<input type="hidden" name="idx">
<input type="hidden" name="orderno">
<input type="hidden" name="playidx" value="<%=playidx%>">
<input type="hidden" name="playcate" value="<%=playcate%>">
<tr>
	<td align="left">
		<input type="button" value="선택아이템 삭제" onClick="delitems(arrFrm)" class="button"> /
		<select name="allusing"  class="select">
			<option value="Y">선택 사용 -> Y </option>
			<option value="N">선택 사용 -> N </option>
		</select><input type="button" value="적용" class="button" onclick="changeUsing(arrFrm);"> /
		<input type="button" value="순서적용" class="button" onclick="changeSort(arrFrm);">
	</td>	
	<td align="right">
		<input type="text" name="itemidarr" value="" size="80" class="input" onKeyPress="if (event.keyCode == 13){ AddIttems2(); return false;}">
		<input type="button" value="상품 직접추가" onclick="AddIttems2()" class="button">
		<input type="button" value="상품 검색추가" onclick="popItemWindow('arrFrm.itemid')" class="button">
	</td>
</tr>
</form>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr bgcolor="#FFFFFF">
	<td colspan="20">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				검색결과 : <b><%= oitem.FTotalCount%></b>
				&nbsp;
				페이지 : <b><%= page %> /<%=  oitem.FTotalpage %></b>				
			</td>
			<td align="right">
				※카테고리 ALL 일경우에만, 리스트에 있는 상품이 정렬 순서대로 노출 됩니다.		
			</td>			
		</tr>
		</table>
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>이미지</td>
	<td>ItemID</td>
	<td>제품명</td>
	<td>정렬<br>순서</td>
	<td>사용<br>여부</td>
	<td>품절<br>여부</td>
</tr>
<% if oitem.FResultCount > 0 then %>

<% for i=0 to oitem.FResultCount-1 %>
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="idx" value="<%= oitem.FItemList(i).FIDX %>">
<% if oitem.FItemList(i).fisusing = "Y" then %>
	<tr bgcolor="#FFFFFF">
<% else %>
	<tr bgcolor="#f1f1f1">
<% end if %>
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center"><img src="<%= oitem.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= oitem.FItemList(i).FItemID %></td>
	<td align="center"><%= oitem.FItemList(i).FItemname %></td>
	<td align="center">
		<input type="text" name="orderno" value="<%= oitem.FItemList(i).forderno %>" size="3" style="text-align:right;" onKeyup="CheckThis(frmBuyPrc<%= i %>)">
	</td>
	<td align="center"><%= oitem.FItemList(i).Fisusing %></td>
	<td align="center">
		<% if oitem.FItemList(i).IsSoldOut then %>
			<font color="red">품절</font>
		<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if oitem.HasPreScroll then %>
			<a href="javascript:jsSerach('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:jsSerach('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oitem.HasNextScroll then %>
			<a href="javascript:jsSerach('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
</table>

<%
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->