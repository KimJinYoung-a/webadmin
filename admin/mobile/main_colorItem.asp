<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	Description : 모바일 사이트 컬러별 상품 목록관리
'	History	:  2010.02.258 허진원
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mobile/main_colorItemCls.asp"-->
<%
dim page, ccd
	ccd = request("ccd")
	page = request("page")
	
	if page="" then page=1


dim oip
	set oip = New Cmain_tabitem_list
	oip.FCurrPage = page
	oip.FPageSize=20
	oip.FRectCCD = ccd
	oip.Getmain_tabitem

dim i
%>
<script language='javascript'>

	function popItemWindow(tgf){
		if (document.Listfrm.ccd.value == ""){
			alert("색상을 선택해 주세요!");
			document.Listfrm.ccd.focus();
		}
		else{
			var popup_item = window.open("/common/pop_CateItemList.asp?target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
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
						upfrm.ccd.value = upfrm.ccd.value + frm.ccd.value + "," ;
						upfrm.itemid.value = upfrm.itemid.value + frm.itemid.value + "," ;
					}
				}
			}
			upfrm.mode.value="del";
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
						upfrm.ccd.value = upfrm.ccd.value + frm.ccd.value + "," ;
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
			arrFrm.itemid.value = arrFrm.itemid.value;
			arrFrm.ccd.value = Listfrm.ccd.value;
			arrFrm.mode.value="add";
			arrFrm.submit();
		}
	}
	
	function AddIttems2(){
		if (document.Listfrm.ccd.value == ""){
			alert("색상을 선택해 주세요!");
			document.Listfrm.ccd.focus();
		}
		else if (document.arrFrm.itemidarr.value == ""){
			alert("아이템번호를  적어주세요!");
			document.arrFrm.itemidarr.focus();
		}
		else if (confirm(arrFrm.itemidarr.value + '아이템을 추가하시겠습니까?')){
			arrFrm.itemid.value = arrFrm.itemidarr.value;
			arrFrm.ccd.value = Listfrm.ccd.value;
			arrFrm.mode.value="add";
			arrFrm.submit();
		}
	}

	function RefreshMainCorItemRec(){
		if (document.Listfrm.ccd.value == ""){
			alert("색상을 선택해 주세요!");
			document.Listfrm.ccd.focus();
		}
		 else{
				  var popwin = window.open('','refreshFrm','');
				  popwin.focus();
				  refreshFrm.target = "refreshFrm";
				  refreshFrm.ccd.value = document.Listfrm.ccd.value;
				  refreshFrm.action = "<%=wwwUrl%>/chtml/mobile/make_main_ColorItemXML.asp";
				  refreshFrm.submit();
		 }
	}

	// 색상 변경시 명령
	function changecontent(){}

</script>
<form name="refreshFrm" method="post">
<input type="hidden" name="ccd">
</form>
<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="Listfrm" method="get" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		컬러탭 선택 :
		<% DrawSelectBoxmaintab "ccd", ccd %>
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
				<img src="/images/icon_reload.gif" onClick="RefreshMainCorItemRec()" style="cursor:pointer" align="absmiddle" alt="XML만들기">
				프론트에 적용
			</td>
		</tr>
		</form>
		<form name="arrFrm" method="post" action="do_mainColorItem.asp">
		<input type="hidden" name="ccd">
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
	<td align="center">색상명</td>
	<td align="center">Image</td>
	<td align="center">ItemID</td>
	<td align="center">제품명</td>
	<td align="center">순서</td>
	<td align="center">품절유무</td>
</tr>
<% for i=0 to oip.FResultCount-1 %>
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="ccd" value="<%= oip.FItemList(i).Fccd %>">
<input type="hidden" name="itemid" value="<%= oip.FItemList(i).FItemID %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center">
		<% if  oip.FItemList(i).Fccd = 1 then
		response.write "빨강"
		elseif oip.FItemList(i).Fccd = 2 then
		response.write "주황"
		elseif 	oip.FItemList(i).Fccd = 3 then
		response.write "노랑"
		elseif 	oip.FItemList(i).Fccd = 5 then
		response.write "초록"
		else
		response.write "파랑"
		end if
		%>
	</td>
	<td align="center"><img src="<%= oip.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= oip.FItemList(i).FItemID %></td>
	<td align="center"><%= oip.FItemList(i).FItemname %></td>
	<td align="center"><input type="text" name="sortNo" value="<%= oip.FItemList(i).FsortNo %>" size="3" style="text-align:right;" /></td>
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
		<a href="?page=<%= oip.StarScrollPage-1 %>&ccd=<%=ccd%>&menupos=<%= menupos %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oip.StarScrollPage to oip.FScrollCount + oip.StarScrollPage - 1 %>
		<% if i>oip.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&ccd=<%=ccd%>&menupos=<%= menupos %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oip.HasNextScroll then %>
		<a href="?page=<%= i %>&ccd=<%=ccd%>&menupos=<%= menupos %>">[next]</a>
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
