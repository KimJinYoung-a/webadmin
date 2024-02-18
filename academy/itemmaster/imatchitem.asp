<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYCategoryCls.asp"-->
<%
'###############################################
' PageName : imatchitem.asp
' Discription : 해당 카테고리내 상품 목록
' History : 2008.03.20 허진원 : 이전 Admin에서 이전/수정
'###############################################

dim dispsailyn
dim cdl,cdm,cds
cdl = RequestCheckvar(request("cdl"),10)
cdm = RequestCheckvar(request("cdm"),10)
cds = RequestCheckvar(request("cds"),10)

dim cd1,cd2,cd3
cd1 = RequestCheckvar(request("cd1"),10)
cd2 = RequestCheckvar(request("cd2"),10)
cd3 = RequestCheckvar(request("cd3"),10)

dispsailyn = RequestCheckvar(request("dispsailyn"),1)

dim mode,ckitem,page
page = RequestCheckvar(request("page"),10)
if page = "" then page = 1
mode = RequestCheckvar(request("mode"),16)
ckitem = request("ckitem")
if ckitem <> "" then
	if checkNotValidHTML(ckitem) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end if
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
if (cdl<>"") and (cdm<>"") and (cds<>"") then
oCateItemItem.GetNewCateItemList cdl,cdm,cds
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

	var ret = confirm('선택 아이템의 카테고리를 변경하시겠습니까?\n\n※기본 카테고리만 변경되며 추가 카테고리의 변경은 상품관리>상품수정 페이지에서 할 수 있습니다.');

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
		upfrm.submit();
	}
}

function TnDispSailYN(){
	document.frm.submit();
}

//-->
</script>
<body style="margin:0 0 0 0">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
	<td align="center">
<table width=300 cellspacing=1 cellpadding=0 class=a border=0 bgcolor="#808080">
<form method="post" name="SubmitFrm" action="/academy/itemmaster/doCdChange.asp">
<input type="hidden" name="itemidarr">
<input type="hidden" >
<tr bgcolor="#FFFFFF">
	<td width=20>
		  <select name="cd1" onchange="javascript:searchCD2(this.options[this.selectedIndex].value);">
		  <option value="">대카테고리선택</option>
		  </select>
		  <select name="cd2" onchange="javascript:searchCD3(this.options[this.selectedIndex].value);">
		  <option value="">중카테고리선택</option>
		  </select>
		  <select name="cd3">
		  <option value="">소카테고리선택</option>
		  </select>
	</td>
	<td align="center"><input type="button" value="카테고리변경" onclick="TnChangeCategory(SubmitFrm);"></td>
</tr>
</form>
</table>

<table width=300 cellspacing=1 cellpadding=0 class=a border=0 bgcolor="#808080">
<tr bgcolor="#FFFFFF">
	<td colspan=3 align="center">
		 <% if oCateItemItem.HasPreScroll then %>
			 <a href="?page=<%= oCateItemItem.StartScrollPage-1 %>&cdl=<%=cdl%>&cdm=<%=cdm%>&cds=<%=cds%>&dispsailyn=<%=dispsailyn%>">[pre]</a>
		 <% else %>
			 [pre]
		 <% end if %>

		 <% for i=0 + oCateItemItem.StartScrollPage to oCateItemItem.FScrollCount + oCateItemItem.StartScrollPage - 1 %>
			 <% if i>oCateItemItem.FTotalpage then Exit for %>
			 <% if CStr(page)=CStr(i) then %>
			 <font color="red">[<%= i %>]</font>
			 <% else %>
			 <a href="?page=<%= i %>&cdl=<%=cdl%>&cdm=<%=cdm%>&cds=<%=cds%>&dispsailyn=<%=dispsailyn%>">[<%= i %>]</a>
			 <% end if %>
		 <% next %>

		 <% if oCateItemItem.HasNextScroll then %>
			 <a href="?page=<%= i %>&cdl=<%=cdl%>&cdm=<%=cdm%>&cds=<%=cds%>&dispsailyn=<%=dispsailyn%>">[next]</a>
		 <% else %>
			 [next]
		 <% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align=left><input type="checkbox" name="ckall" onClick="ckAll(this);"></td>
		<form method="get" name="frm">
		<input type="hidden" name="cdl" value="<% = cdl %>">
		<input type="hidden" name="cdm" value="<% = cdm %>">
		<input type="hidden" name="cds" value="<% = cds %>">
		<td colspan=2 align=left>&nbsp;<input type="checkbox" name="dispsailyn" onClick="TnDispSailYN();" <% if dispsailyn="on" then response.write "checked" %>>판매,전시만 보여주기</td>
		</form>
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
			 <a href="?page=<%= oCateItemItem.StartScrollPage-1 %>&cdl=<%=cdl%>&cdm=<%=cdm%>&cds=<%=cds%>&dispsailyn=<%=dispsailyn%>">[pre]</a>
		 <% else %>
			 [pre]
		 <% end if %>

		 <% for i=0 + oCateItemItem.StartScrollPage to oCateItemItem.FScrollCount + oCateItemItem.StartScrollPage - 1 %>
			 <% if i>oCateItemItem.FTotalpage then Exit for %>
			 <% if CStr(page)=CStr(i) then %>
			 <font color="red">[<%= i %>]</font>
			 <% else %>
			 <a href="?page=<%= i %>&cdl=<%=cdl%>&cdm=<%=cdm%>&cds=<%=cds%>&dispsailyn=<%=dispsailyn%>">[<%= i %>]</a>
			 <% end if %>
		 <% next %>

		 <% if oCateItemItem.HasNextScroll then %>
			 <a href="?page=<%= i %>&cdl=<%=cdl%>&cdm=<%=cdm%>&cds=<%=cds%>&dispsailyn=<%=dispsailyn%>">[next]</a>
		 <% else %>
			 [next]
		 <% end if %>
	</td>
</tr>
</table>
	</td>
</tr>
</table>
<iframe name="FrameSearchCategory" src="/academy/itemmaster/frame_category_select.asp?form_name=SubmitFrm&element_name=cd1" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script language="JavaScript">
<!--

//대카테고리선택시 중카테고리 셋팅
function searchCD2(paramCodeLarge) {
		
	resetLeftCountrySelect() ;		
	resetLeftCitySelect() ;
	
	if(paramCodeLarge != '') {
		FrameSearchCategory.location.href="/academy/itemmaster/frame_category_select.asp?search_code=" + paramCodeLarge + "&form_name=SubmitFrm&element_name=cd2";
	}
}

//중카테고리 선택시 소카테고리 셋팅	
function searchCD3(paramCodeMid) {	
	resetLeftCitySelect() ;
	
	if(paramCodeMid != '') {
		FrameSearchCategory.location.href="/academy/itemmaster/frame_category_select.asp?search_code=" + paramCodeMid + "&form_name=SubmitFrm&element_name=cd3";
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
%>
</body>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
