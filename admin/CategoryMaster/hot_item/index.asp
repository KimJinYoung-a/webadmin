<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/category_hot_managecls.asp" -->

<%
dim research,isusing, fixtype, linktype, poscode, validdate
dim page, cdl, cdm, imgSize
dim cds

isusing = request("isusing")
research= request("research")
poscode = request("poscode")
fixtype = request("fixtype")
page    = request("page")
validdate= request("validdate")
cdl		= request("cdl")
cdm		= request("cdm")
cds		= request("cds")

if ((research="") and (isusing="")) then 
    isusing = "Y"
    validdate = "on"
end if

if page="" then page=1

dim oposcode
set oposcode = new CCateContentsCode
oposcode.FRectPosCode = poscode

if (poscode<>"") then
    oposcode.GetOneContentsCode
end if

dim oCateContents
set oCateContents = new CCateContents
oCateContents.FPageSize = 10
oCateContents.FCurrPage = page
oCateContents.FRectIsusing = isusing
oCateContents.FRectfixtype = fixtype
oCateContents.FRectPosCode = poscode
oCateContents.FRectvaliddate = validdate
oCateContents.FRectCdl = cdl
oCateContents.FRectCdm = cdm
oCateContents.FRectCds = cds
oCateContents.GetHotCateItemList

dim i
%>
<script language='javascript'>
function NextPage(page){
    frm.page.value = page;
    frm.submit();
}

function popPosCodeManage(){
    var popwin = window.open('/admin/categorymaster/popCatePosCodeEdit.asp','catePosCodeEdit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function AddNewCateContents(idx){
    var popwin = window.open('/admin/categorymaster/hot_item/popCateContentsEdit.asp?idx=' + idx,'cateHotPosCodeEdit','width=900,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}



function AssignReal(vTerm){
	 var popwin = window.open('','refreshFrm_Cate','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Cate";
	 refreshFrm.action = "http://<%=CHKIIF(application("Svr_Info")="Dev","2011www","www1")%>.10x10.co.kr/chtml/make_cate_hot_JS.asp?vTerm=" + vTerm;
	 refreshFrm.submit();
}


function chkConfirm() {
    if (document.frm.poscode.value == ""){
		alert("적용위치를 선택해주세요");
		document.frm.poscode.focus();
		return false;
	}
<% if cdl<>"110" then %>
	else if (document.frm.cdl.value == ""){
		alert("카테고리를 선택해주세요");
		document.frm.cdl.focus();
		return false;
	}
	else if (document.frm.cdl.value == "110"){
		alert("감성채널은 검색을 실행하여 중카테고리를 선택할 수 있도록해야합니다.");
		return false;
	}
	else{
		return true;
	}
<% else %>
	else if (document.frm.cdl.value != "110"){
		alert("카테고리를 선택해주세요");
		document.frm.cdl.focus();
		return false;
	}
	else{
		if(document.frm.cdm.value=="") {
			if(confirm("중카테고리를 선택하지 않았습니다.\n\n중카테고리 없이 처리하시겠습니까?")) {
				return true;
			} else {
				return false;
			}
		} else {
			return true;
		}
	}
<% end if %>
}

// 카테고리 변경시 명령
function changecontent(){
	frm.submit();
}
</script>

<table width="100%" border="0" cellpadding="7" cellspacing="1" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr bgcolor="#FFFFFF">
		<td class="a" width="15%"><input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >종료이전</td>
		<td class="a">
		    사용구분
			<select name="isusing" class="select">
			<option value="">전체
			<option value="Y" <% if isusing="Y" then response.write "selected" %> >사용함
			<option value="N" <% if isusing="N" then response.write "selected" %> >사용안함
			</select>
			&nbsp;&nbsp;
			<br><br>
			카테고리
			<% call DrawSelectBoxCategoryLarge("cdl", cdl) %>
			&nbsp;&nbsp;
			<% if cdl <> "" then DrawSelectBoxCategoryMid "cdm", cdl, cdm %>
			&nbsp;&nbsp;
			<% if cdm <> "" then DrawSelectBoxCategorySmall "cds", cdl, cdm , cds %>
			
		</td>
		<td class="a" align="right" width="10%">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>
<table width="100%" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td>
		<b>※ 실서버 프론트에 바로 노출 됩니다. 작업시 상품코드/카테고리 를 한번 더 확인 해주세요</b>
    </td>
    <td align="right"><a href="javascript:AddNewCateContents('0');"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
</table>
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#DDDDFF" align="center">
    <td width="10%">idx</td>
    <td>카테고리</td>
    <td width="15%">이미지</td>
    <td width="15%">시작일</td>
    <td width="15%">종료일</td>
    <td width="10%">사용<br>여부</td>
</tr>
<%
	for i=0 to oCateContents.FResultCount - 1
%>
<% if (oCateContents.FItemList(i).IsEndDateExpired) or (oCateContents.FItemList(i).FIsusing="N") then %>
<tr bgcolor="#DDDDDD">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td align="center"><%= oCateContents.FItemList(i).Fidx %></td>
    <td align="left" style="padding-left:20px;"><%
		Response.Write "대 : "&oCateContents.FItemList(i).Fcodename & "<br>"
		if Not(oCateContents.FItemList(i).Fcdmname="" or isNull(oCateContents.FItemList(i).Fcdmname)) then
			Response.Write  "중 : "& oCateContents.FItemList(i).Fcdmname & "<br>"
			Response.Write  "소 : "& oCateContents.FItemList(i).Fcdsname
		end if
    %></td>
    <td><a href="javascript:AddNewCateContents('<%= oCateContents.FItemList(i).Fidx %>');"><img src="<%= oCateContents.FItemList(i).Fimg1 %>" border="0"><img src="<%= oCateContents.FItemList(i).Fimg2 %>" border="0"><img src="<%= oCateContents.FItemList(i).Fimg3 %>" border="0"></a></td>
    <td align="center"><%= oCateContents.FItemList(i).FStartdate %></td>
    <td align="center">
    <% if (oCateContents.FItemList(i).IsEndDateExpired) then %>
    <font color="#777777"><%= Left(oCateContents.FItemList(i).FEnddate,10) %></font>
    <% else %>
    <%= Left(oCateContents.FItemList(i).FEnddate,10) %>
    <% end if %>
    </td>
    <td align="center"><%= oCateContents.FItemList(i).FIsusing %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="12" align="center">
    <% if oCateContents.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCateContents.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oCateContents.StarScrollPage to oCateContents.FScrollCount + oCateContents.StarScrollPage - 1 %>
		<% if i>oCateContents.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oCateContents.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
<%
set oposcode = Nothing
set oCateContents = Nothing
%>
<form name="refreshFrm" method="post">
<input type="hidden" name="cdl" value="<%=cdl%>">
<input type="hidden" name="cdm" value="<%=cdm%>">
<input type="hidden" name="cds" value="<%=cds%>">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
