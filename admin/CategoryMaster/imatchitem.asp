<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp"-->
<%
'###############################################
' PageName : imatchitem.asp
' Discription : 해당 카테고리내 상품 목록
' History : 2008.03.20 허진원 : 이전 Admin에서 이전/수정
'###############################################

dim dispsailyn, itemCateDiv
dim cdl,cdm,cds
cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")

dim cd1,cd2,cd3
cd1 = request("cd1")
cd2 = request("cd2")
cd3 = request("cd3")

dispsailyn = request("dispsailyn")
itemCateDiv = request("itemCateDiv")
if itemCateDiv="" then itemCateDiv="D"

dim mode,ckitem,page
page = request("page")
if page = "" then page = 1
mode = request("mode")
ckitem = request("ckitem")

dim arrItemid, makerid
arrItemid = request("arrItemid")
makerid = request("makerid")

dim sqlStr
if mode="delArr" then
	sqlStr = "delete from [db_temp].[dbo].tbl_temp_itemcategory"
	sqlStr = sqlStr + " where itemid in (" + ckitem + ")"
	rsget.Open sqlStr, dbget, 1

end if


dim oCateItemItem
set oCateItemItem = new CCatemanager
oCateItemItem.FPageSize = 100
oCateItemItem.FCurrPage = page
oCateItemItem.FRectDispSailYN = dispsailyn
oCateItemItem.FRectArrItemid = arrItemid
oCateItemItem.FRectMakerid = makerid
if (cdl<>"") and (cdm<>"") and (cds<>"") then
	if itemCateDiv="D" then
		'// 기본카테고리 사용
		oCateItemItem.GetNewCateItemList cdl,cdm,cds
	else
		'// 추가카테고리 사용
		oCateItemItem.GetAddCateItemList cdl,cdm,cds
	end if
end if

dim nkeyword,keywords
set nkeyword = new CCatemanager
if (cdl<>"") and (cdm<>"") and (cds<>"") then
nkeyword.GetCategoryKeyword cdl,cdm,cds

keywords = nkeyword.FItemList(0).Fkeyword
end if

dim i
%>
<script language="JavaScript">
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
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

// 기본 카테고리 변경
function TnChangeCategory(upfrm){

	if (upfrm.cd1.value == ""){
		alert('대카테고리를 선택해주세요');
		return;
	}

	if (upfrm.cd2.value == ""){
		alert('중카테고리를 선택해주세요');
		return;
	}

	if (upfrm.cd3.value == ""){
		alert('소카테고리를 선택해주세요');
		return;
	}

	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}

	var ret = confirm('선택 아이템의 기본 카테고리를 변경하시겠습니까?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + "," ;
				}
			}
		}
		upfrm.codeDiv.value="D";
		upfrm.submit();
	}
}

// 추가 카테고리 지정
function TnInputAddCategory(upfrm){

	if (upfrm.cd1.value == ""){
		alert('대카테고리를 선택해주세요');
		return;
	}

	if (upfrm.cd2.value == ""){
		alert('중카테고리를 선택해주세요');
		return;
	}

	if (upfrm.cd3.value == ""){
		alert('소카테고리를 선택해주세요');
		return;
	}

	if(upfrm.cd2.value=='<%=cdl&","&cdm%>') {
		alert('같은 중분류 안에서는 추가로 등록할 수 없습니다.');
		return;
	}

	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}

	var ret = confirm('*** 추가카테고리를 선택하셨습니다! ***\n\n선택 아이템의 [추가]카테고리를 지정하시겠습니까?\n\n※기본 카테고리와는 별개이며 추가 후 [현재 카테고리 구분]에서 확인할 수 있습니다.');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + "," ;
				}
			}
		}
		upfrm.codeDiv.value="A";
		upfrm.submit();
	}
}

// 추가 카테고리 삭제
function TnDelAddCategory(upfrm){

	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}

	var ret = confirm('선택 아이템을 [추가]카테고리에서 삭제하시겠습니까?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + "," ;
				}
			}
		}
		upfrm.codeDiv.value="DelA";
		upfrm.submit();
	}
}


function TnCategoryKeyword(upfrm){

	if (upfrm.keyword.value == ""){
		alert('키워드를 입력해주세요');
		return;
	}

	var ret = confirm('키워드를 변경하시겠습니까?');

	if (ret){
		upfrm.submit();
	}
}

function TnDispSailYN(chk){
	if(chk.checked) {
		SrchFrm.dispsailyn.value="on";
	} else {
		SrchFrm.dispsailyn.value="";
	}
	SrchFrm.submit();
}

function TnMovePage(pg) {
	SrchFrm.page.value=pg;
	SrchFrm.submit();
}
function TnCategoryItem() {
	if(!SrchFrm.arrItemid.value) {
		alert('검색할 상품코드를 입력해주세요.');
		return;
	} else {
		SrchFrm.submit();
	}
}

function TnCategoryMaker() {
	if(!SrchFrm.makerid.value) {
		alert('검색할 브랜드ID를 입력해주세요.');
		return;
	} else {
		SrchFrm.submit();
	}
}

function chgItemCateDiv() {
	SrchFrm.submit();
}

//-->
</script>
<body style="margin:0 0 0 0">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
	<td align="center">
		<table width="100%" cellspacing=1 cellpadding=3 class=a border=0 bgcolor="#808080">
		<form method="post" name="keyfrm" action="/admin/CategoryMaster/doCdKeyword.asp">
		<input type="hidden" name="cdl" value="<% = cdl %>">
		<input type="hidden" name="cdm" value="<% = cdm %>">
		<input type="hidden" name="cds" value="<% = cds %>">
		<tr bgcolor="#FFFFFF">
			<td>
				<input type="text" name="keyword" size="40" value="<% = keywords %>" class="text"><br>
				<input type="button" value="검색키워드변경" onclick="TnCategoryKeyword(keyfrm);" class="button">(ex: 노트,다이어리,공책...)
			</td>
		</tr>
		</form>
		<form method="get" name="SrchFrm" action="imatchitem.asp">
		<input type="hidden" name="cdl" value="<% = cdl %>">
		<input type="hidden" name="cdm" value="<% = cdm %>">
		<input type="hidden" name="cds" value="<% = cds %>">
		<input type="hidden" name="page" value="<% = page %>">
		<input type="hidden" name="dispsailyn" value="<% = dispsailyn %>">
		<tr bgcolor="#FFFFFF">
			<td id="tdICtDiv" style="padding-top:10px;" bgcolor="<%=chkIIF(itemCateDiv="D","#FFFFFF","#FFBBAA")%>">
				현재 카테고리 구분
				<label><input type="radio" name="itemCateDiv" value="D" <%=chkIIF(itemCateDiv="D","checked","")%> onClick="chgItemCateDiv();">기본</label>
				<label><input type="radio" name="itemCateDiv" value="A" <%=chkIIF(itemCateDiv="A","checked","")%> onClick="chgItemCateDiv();">추가</label>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>
				<input type="text" name="arrItemid" size="40" value="<% = arrItemid %>" class="text"><br>
				<input type="button" value="상품코드검색" onclick="TnCategoryItem();" class="button">(ex: 123123,26457...)
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>
				<input type="text" name="makerid" size="40" value="<% = makerid %>" class="text"><br>
				<input type="button" value="브랜드ID검색" onclick="TnCategoryMaker();" class="button">(ex: mmmg,ithinkso...)
			</td>
		</tr>
		</form>
		</table>
		<% if itemCateDiv="D" then %>
		<table width="100%" cellspacing=1 cellpadding=3 class=a border=0 bgcolor="#808080">
		<form method="post" name="SubmitFrm" action="/admin/CategoryMaster/doCdChange.asp">
		<input type="hidden" name="itemidarr">
		<input type="hidden" name="codeDiv" value="D">
		<tr bgcolor="#FFFFFF">
			<td width="140" align="center">
				  <select name="cd1" onchange="javascript:searchCD2(this.options[this.selectedIndex].value);" class="select">
				  <option value="">대카테고리선택</option>
				  </select>
				  <select name="cd2" onchange="javascript:searchCD3(this.options[this.selectedIndex].value);" class="select">
				  <option value="">중카테고리선택</option>
				  </select>
				  <select name="cd3" class="select">
				  <option value="">소카테고리선택</option>
				  </select>
			</td>
			<td align="center">
				<input type="button" value="기본 카테고리로 변경" onclick="TnChangeCategory(SubmitFrm);" style="width:125px;height:30px;" class="button"><br>
				<input type="button" value="추가 카테고리로 변경" onclick="TnInputAddCategory(SubmitFrm);" style="width:125px;height:20px;background-color:#FFBBAA;margin-top:5px;" class="button">
			</td>
		</tr>
		</form>
		</table>
		<% elseif itemCateDiv="A" then %>
		<table width="100%" cellspacing=1 cellpadding=3 class=a border=0 bgcolor="#808080">
		<form method="post" name="SubmitFrm" action="/admin/CategoryMaster/doCdChange.asp">
		<input type="hidden" name="itemidarr">
		<input type="hidden" name="codeDiv" value="DelA">
		<input type="hidden" name="cd1" value="<%=cdl%>">
		<input type="hidden" name="cd2" value="<%=cdl&","&cdm%>">
		<input type="hidden" name="cd3" value="<%=cdl&","&cdm&","&cds%>">
		<tr bgcolor="#FFFFFF">
			<td align="right">
				<input type="button" value="선택상품 삭제" onclick="TnDelAddCategory(SubmitFrm);" style="width:125px;background-color:#FFBBAA;margin-top:5px;" class="button">
			</td>
		</tr>
		</form>
		</table>	
		<% end if %>
		<table width="100%" cellspacing=1 cellpadding=0 class=a border=0 bgcolor="#808080">
		<tr bgcolor="#FFFFFF">
			<td colspan=3 align="center">
				 <% if oCateItemItem.HasPreScroll then %>
					 <a href="javascript:TnMovePage(<%= oCateItemItem.StartScrollPage-1 %>)">[pre]</a>
				 <% else %>
					 [pre]
				 <% end if %>
		
				 <% for i=0 + oCateItemItem.StartScrollPage to oCateItemItem.FScrollCount + oCateItemItem.StartScrollPage - 1 %>
					 <% if i>oCateItemItem.FTotalpage then Exit for %>
					 <% if CStr(page)=CStr(i) then %>
					 <font color="red">[<%= i %>]</font>
					 <% else %>
					 <a href="javascript:TnMovePage(<%= i %>)">[<%= i %>]</a>
					 <% end if %>
				 <% next %>
		
				 <% if oCateItemItem.HasNextScroll then %>
					 <a href="javascript:TnMovePage(<%= i %>)">[next]</a>
				 <% else %>
					 [next]
				 <% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td align=left><input type="checkbox" name="ckall" onClick="ckAll(this);"></td>
			<td colspan=2 align=left>&nbsp;<input type="checkbox" name="dispsailyn" onClick="TnDispSailYN(this);" <% if dispsailyn="on" then response.write "checked" %>>판매,전시만 보여주기</td>
		</tr>
		<% for i=0 to oCateItemItem.FResultCount-1 %>
		<form name="frmBuyPrc_<%=i%>" method="post">
		<input type="hidden" name="itemid" value="<%= oCateItemItem.FITemList(i).FItemID %>">
		<tr bgcolor="#FFFFFF">
			<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
			<td width=50><img src="<%= oCateItemItem.FITemList(i).FImgSmall %>" width="50" height="50" border="0"></td>
			<td><font color="#888888"><%= "[" & oCateItemItem.FITemList(i).FItemID & "] " & oCateItemItem.FITemList(i).FItemName %></font><br>(<%= oCateItemItem.FITemList(i).FMakerid %>)<br>
			<% if oCateItemItem.FITemList(i).Fsellyn="N" then %>
			<font color="red">판매x</font>
			<% end if %>
			<% if oCateItemItem.FITemList(i).Fisusing="N" then %>
			사용x
			<% end if %>
			</td>
		</tr>
		</form>
		<% next %>
		<tr bgcolor="#FFFFFF">
			<td colspan=3 align="center">
				 <% if oCateItemItem.HasPreScroll then %>
					 <a href="javascript:TnMovePage(<%= oCateItemItem.StartScrollPage-1 %>)">[pre]</a>
				 <% else %>
					 [pre]
				 <% end if %>
		
				 <% for i=0 + oCateItemItem.StartScrollPage to oCateItemItem.FScrollCount + oCateItemItem.StartScrollPage - 1 %>
					 <% if i>oCateItemItem.FTotalpage then Exit for %>
					 <% if CStr(page)=CStr(i) then %>
					 <font color="red">[<%= i %>]</font>
					 <% else %>
					 <a href="javascript:TnMovePage(<%= i %>)">[<%= i %>]</a>
					 <% end if %>
				 <% next %>
		
				 <% if oCateItemItem.HasNextScroll then %>
					 <a href="javascript:TnMovePage(<%= i %>)">[next]</a>
				 <% else %>
					 [next]
				 <% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<iframe name="FrameSearchCategory" src="/admin/CategoryMaster/frame_category_select.asp?form_name=SubmitFrm&element_name=cd1" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script language="JavaScript">
<!--

//대카테고리선택시 중카테고리 셋팅
function searchCD2(paramCodeLarge) {
		
	resetLeftCountrySelect() ;		
	resetLeftCitySelect() ;
	
	if(paramCodeLarge != '') {
		FrameSearchCategory.location.href="/admin/CategoryMaster/frame_category_select.asp?search_code=" + paramCodeLarge + "&form_name=SubmitFrm&element_name=cd2";
	}
}

//중카테고리 선택시 소카테고리 셋팅	
function searchCD3(paramCodeMid) {	
	resetLeftCitySelect() ;
	
	if(paramCodeMid != '') {
		FrameSearchCategory.location.href="/admin/CategoryMaster/frame_category_select.asp?search_code=" + paramCodeMid + "&form_name=SubmitFrm&element_name=cd3";
	}	 
}

//대카테고리 초기화
function resetLeftCountrySelect() {
	document.SubmitFrm.cd2.length = 1;
	document.SubmitFrm.cd2.selectedIndex = 0 ;
}

		
//중카테고리 초기화
function resetLeftCitySelect() {
	document.SubmitFrm.cd3.length = 1;
	document.SubmitFrm.cd3.selectedIndex = 0 ;
}

//-->
</script>
<%
set oCateItemItem = Nothing
set nkeyword = Nothing
%>
</body>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->