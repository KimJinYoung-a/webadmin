<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/Diary2009/classes/DiaryCls.asp"-->
<%
dim page, isusing, cate
page = request("page")
cate = request("cated")
isusing = request("isusing")

if page="" then page=1

dim omd
set omd = New DiaryCls
omd.FCurrPage = page
omd.FPageSize=20
omd.FRectCDL = cate
omd.FRectIsUsing = isusing
omd.GetWithBuyList

dim i
%>
<script language='javascript'>
<!--
function popItemWindow(tgf){
	if (document.Listfrm.cated.value == ""){
		alert("카테고리를 선택해 주세요!");
		document.Listfrm.cated.focus();
	}
	else{
		var popup_item = window.open("/common/pop_CateItemList.asp?cdl=010&cdm=010&target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
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
		upfrm.mode.value="isUsingValue";
		upfrm.submit();
	}
}

//카테고리 변경 적용
function changeCate(upfrm) {
	if(document.Listfrm.cated.value == ""){
		alert('카테고리를 선택해 주세요.');
		document.Listfrm.cated.focus();
		return;
	}
	if (!CheckSelected()){
		alert('상품을 선택해 주세요.');
		return;
	}
	var ret = confirm('선택하신 다이어리 카테고리로 변경하시겠습니까?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
					upfrm.cate.value = document.Listfrm.cated.value ;
				}
			}
		}
		upfrm.mode.value="modify";
		upfrm.submit();
	}
}

function AddIttems(){
	var ret = confirm(arrFrm.itemid.value + '아이템을 추가하시겠습니까?');
	if (ret){
		arrFrm.itemid.value = arrFrm.itemid.value;
		arrFrm.cate.value = document.Listfrm.cated.value;
		arrFrm.mode.value="add";
		arrFrm.submit();
	}
}

function AddIttems2(){
	if (document.Listfrm.cated.value == ""){
		alert("카테고리를 선택해 주세요!");
		document.Listfrm.cated.focus();
	}
	else if (document.arrFrm.itemidarr.value == ""){
		alert("아이템번호를  적어주세요!");
		document.arrFrm.itemidarr.focus();
	}
	else if (confirm(arrFrm.itemidarr.value + '아이템을 추가하시겠습니까?')){
		arrFrm.itemid.value = arrFrm.itemidarr.value;
		arrFrm.cate.value = document.Listfrm.cated.value;
		arrFrm.mode.value="add";
		arrFrm.submit();
	}
}

function modify_catename(idx,mode){
	window.open('/admin/Diary2009/pop_Diartycate.asp?idx='+idx+'&mode='+mode);
}
//-->
</script>
<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="Listfrm" method="get" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		다이어리 카테고리 :
		<select name="cated">
			<option value="">-선택-
			<option value="10" <% if cate="10" then response.write "selected" %>>심플
			<option value="20" <% if cate="20" then response.write "selected" %>>일러스트
			<option value="30" <% if cate="30" then response.write "selected" %>>패턴
			<option value="40" <% if cate="40" then response.write "selected" %>>포토
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
</form>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
		<form name="arrFrm" method="post" action="doWithBuy.asp">
		<input type="hidden" name="cate">
		<input type="hidden" name="mode">
		<input type="hidden" name="itemid">
		<input type="hidden" name="sortNo">
		<input type="hidden" name="idx">
		<tr>
			<td colspan="2" align="right">
				<input type="text" name="itemidarr" value="" size="80" class="input" onkeypress="if(event.keyCode==13){return false;}">
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
				<input type="button" value="순서적용" class="button" onclick="changeSort(arrFrm);"> /
				<input type="button" value="카테고리적용" class="button" onclick="changeCate(arrFrm);"> /
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
	<td colspan="8">&nbsp;검색된 상품수 : <%=omd.FTotalCount%> 건</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">다이어리 카테고리명</td>
	<td align="center">Image</td>
	<td align="center">ItemID</td>
	<td align="center">제품명</td>
	<td align="center">순서</td>
	<td align="center">사용유무</td>
	<td align="center">품절유무</td>
</tr>
<%
Dim catename
%>
<% for i=0 to omd.FResultCount-1 %>
<%
	If omd.FItemList(i).FCdl = 10 Then catename = "심플"
	If omd.FItemList(i).FCdl = 20 Then catename = "일러스트"
	If omd.FItemList(i).FCdl = 30 Then catename = "패턴"
	If omd.FItemList(i).FCdl = 40 Then catename = "포토"
%>
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="itemid" value="<%= omd.FItemList(i).FItemID %>">
<input type="hidden" name="idx" value="<%= omd.FItemList(i).FIdx %>">
<input type="hidden" name="cate" value="<%= omd.FItemList(i).FCdl %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center"><a href="javascript:modify_catename('<%=omd.FItemList(i).Fidx%>','modify');"><%= catename %></a></td>
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
		<a href="?page=<%= omd.StarScrollPage-1 %>&cdl=<%=cdl%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + omd.StartScrollPage to omd.FScrollCount + omd.StartScrollPage - 1 %>
		<% if i>omd.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&cate=<%=cate%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if omd.HasNextScroll then %>
		<a href="?page=<%= i %>&cate=<%=cate%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[next]</a>
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
