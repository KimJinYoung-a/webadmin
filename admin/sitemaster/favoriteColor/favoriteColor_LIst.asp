<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2010.04.07 허진원 생성
'	Description : Favorite Colore 관리
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/favoriteColorCls.asp"-->
<%
dim page, Category, colorCD, isusing
Dim oitem, lp , schcolorCD
	category	= request("category")
	colorCD		= request("colorCD")
	page		= request("page")
	isusing		= request("isusing")
	schcolorCD	= request("schcolorCD")
	
	if page="" then page=1

dim oip
	set oip = New CfavoriteColor
	oip.FCurrPage = page
	oip.FPageSize=20
	oip.FRectCategory = category
	oip.FRectColorCD = schcolorCD
	oip.FRectIsUsing = isusing
	oip.GetfavoriteColor

dim i
	
	set oitem = new CItemColor
	oitem.FPageSize = 50
	oitem.FRectUsing = "Y"
	oitem.GetColorList

%>
<script language='javascript'>

	function popItemWindow(tgf){
		if (document.Listfrm.category.value == ""){
			alert("탭을 선택해 주세요!");
			document.Listfrm.category.focus();
		}
		else if (document.Listfrm.schcolorCD.value == ""){
			alert("색상을 선택해 주세요!");
		}
		else{
			var popup_item = window.open("/common/pop_CateItemList.asp?category=" + document.refreshFrm.category.value + "&target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
			popup_item.focus();
		}
	}

	function popColorWindow(){
		var popup_item = window.open("/admin/sitemaster/favoriteColor/popManageColorCode.asp", "popup_item", "width=380,height=600,scrollbars=yes,status=no");
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
						upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
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
						upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
					}
				}
			}
			upfrm.category.value = Listfrm.category.value;
			upfrm.colorCD.value = Listfrm.schcolorCD.value;
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
						upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
						upfrm.sortNo.value = upfrm.sortNo.value + frm.sortNo.value + "," ;
					}
				}
			}
			upfrm.category.value = Listfrm.category.value;
			upfrm.colorCD.value = Listfrm.schcolorCD.value;
			upfrm.mode.value="ChangeSort";
			upfrm.submit();
	
		}
	}
	
	function AddIttems(){
		var ret = confirm(arrFrm.itemid.value + '아이템을 추가하시겠습니까?');
		if (ret){
			arrFrm.itemid.value = arrFrm.itemid.value;
			arrFrm.category.value = Listfrm.category.value;
			arrFrm.colorCD.value = Listfrm.schcolorCD.value;
			arrFrm.mode.value="add";
			arrFrm.submit();
		}
	}
	
	function AddIttems2(){
		if (document.Listfrm.category.value == ""){
			alert("탭을 선택해 주세요!");
			document.Listfrm.category.focus();
		}
		else if (document.Listfrm.schcolorCD.value == ""){
			alert("색상을 선택해 주세요!");
		}
		else if (document.arrFrm.itemidarr.value == ""){
			alert("아이템번호를  적어주세요!");
			document.arrFrm.itemidarr.focus();
		}
		else if (confirm(arrFrm.itemidarr.value + '아이템을 추가하시겠습니까?')){
			arrFrm.itemid.value = arrFrm.itemidarr.value;
			arrFrm.category.value = Listfrm.category.value;
			arrFrm.colorCD.value = Listfrm.schcolorCD.value;
			arrFrm.mode.value="add";
			arrFrm.submit();
		}
	}

	// 탭 변경시 명령
	function changecontent(){}

	function chgColorChip(ccd,cnt,idx) {
		document.Listfrm.colorCD.value=ccd;
		document.Listfrm.schcolorCD.value=idx;
		for(var i=0;i<=(cnt-1);i++) {
			if(i==ccd) {
				document.getElementById("tbColor"+i).style.backgroundColor="#000000";
			} else {
				document.getElementById("tbColor"+i).style.backgroundColor="#EDEDED";
			}
		}
	}

</script>
<form name="refreshFrm" method="post">
<input type="hidden" name="category">
<input type="hidden" name="colorCD">
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
		<% DrawSelectBoxCateTab "category", category %>&nbsp;/&nbsp;
		사용유무 :
		<select name="isusing" onchange="this.form.submit();" class="select">
			<option value=""  <% if isusing="" then response.write "selected" %>>전체</option>
			<option value="Y" <% if isusing="Y" then response.write "selected" %>>사용</option>
			<option value="N" <% if isusing="N" then response.write "selected" %>>사용안함</option>
		</select><br>
		<%' DrawSelectBoxColoreBar "colorCD",colorCD %>
		<table border="0" cellspacing="3" cellpadding="0">
		<tr>
		<%
			'For i=1 to 20
			if oitem.FResultCount>0 then
				for lp=0 to oitem.FResultCount-1
		%>
			<td onClick="chgColorChip(<%=lp%>,'<%=oitem.FResultCount%>','<%=oitem.FItemList(lp).FcolorCode%>')" style="cursor:pointer">
				<table id="tbColor<%=lp%>" border="0" cellpadding="0" cellspacing="1" bgcolor="<% if cstr(colorCD)=cstr(lp) then %>#000000<% else %>#EDEDED<% end if %>">
				<tr>
					<td bgcolor="#FFFFFF"><img src="<%=oitem.FItemList(lp).FcolorIcon%>" width="15" height="15" hspace="1" vspace="1" border="0"></td>
				</tr>
				</table>
			</td>
		<%
				Next
			End If 
		%>
		<input type="hidden" name="colorCD" value="<%=colorCD%>">
		<input type="hidden" name="schcolorCD" value="<%=schcolorCD%>">
		</tr>
		</table>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
		<form name="arrFrm" method="post" action="doFavoriteColor.asp">
		<input type="hidden" name="category">
		<input type="hidden" name="colorCD">
		<input type="hidden" name="mode">
		<input type="hidden" name="itemid">
		<input type="hidden" name="idx">
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
			<% If C_ADMIN_AUTH then%>
			<td align="right"><input type="button" value="Color 추가" onclick="popColorWindow();" class="button"></td>
			<% End If %>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FAFAFA" height="22">
	<td colspan="9">&nbsp;검색된 상품수 : <%=oip.FTotalCount%> 건 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;※ 순서 0번 : [메인 basic 이미지 노출] 각 탭별 메인 노출 상품 1개 필수입니다.</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">탭</td>
	<td align="center">컬러</td>
	<td align="center">이미지</td>
	<td align="center">ItemID</td>
	<td align="center">제품명</td>
	<td align="center">순서</td>
	<td align="center">사용유무</td>
	<td align="center">품절유무</td>
</tr>
<% for i=0 to oip.FResultCount-1 %>
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="idx" value="<%= oip.FItemList(i).FIDX %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center">
		<% if  oip.FItemList(i).Fcategory = 1 then
		response.write "Stationary&Persnal"
		elseif oip.FItemList(i).Fcategory = 2 then
		response.write "Home&Living"
		elseif 	oip.FItemList(i).Fcategory = 3 then
		response.write "Fashion&Beauty"
		elseif 	oip.FItemList(i).Fcategory = 4 then
		response.write "Kidult&Hobby"
		elseif 	oip.FItemList(i).Fcategory = 5 then
		response.write "Kids&Baby"
		else
		response.write "미분류"
		end if
		%>
	</td>
	<td align="center"><img src="http://fiximage.10x10.co.kr/web2011/common/color/<%= oip.FItemList(i).FcolorCD%>" width="20" height="20"></td>
	<td align="center"><img src="<%= oip.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= oip.FItemList(i).FItemID %></td>
	<td align="center"><%= oip.FItemList(i).FItemname %></td>
	<td align="center"><input type="text" name="sortNo" value="<%= oip.FItemList(i).FsortNo %>" size="3" style="text-align:right;" /><%If oip.FItemList(i).FsortNo = "0" then%>메인<%End if%></td>
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
	<td colspan="9" align="center">
	<% if oip.HasPreScroll then %>
		<a href="?page=<%= oip.StarScrollPage-1 %>&category=<%=category%>&isusing=<%=isusing%>&menupos=<%= menupos %>&colorCD=<%=colorCD%>&schcolorCD=<%=schcolorCD%>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oip.StarScrollPage to oip.FScrollCount + oip.StarScrollPage - 1 %>
		<% if i>oip.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&category=<%=category%>&isusing=<%=isusing%>&menupos=<%= menupos %>&colorCD=<%=colorCD%>&schcolorCD=<%=schcolorCD%>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oip.HasNextScroll then %>
		<a href="?page=<%= i %>&category=<%=category%>&isusing=<%=isusing%>&menupos=<%= menupos %>&colorCD=<%=colorCD%>&schcolorCD=<%=schcolorCD%>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set oip = Nothing
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
