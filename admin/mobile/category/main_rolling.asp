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
<!-- #include virtual="/admin/mobile/category/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/mo_catetoryMainManageCls.asp" -->
<%
'###############################################
' PageName : main_rolling.asp
' Discription : 모바일 카테고리 메인 롤링 관리
' History : 2020.11.30 정태훈 생성
'###############################################

dim research, view_yn, fixtype, linktype, Catecode, validdate, prevDate , sedatechk , prevTime
dim page, imgURL
	view_yn = request("view_yn")
	research= request("research")
	Catecode = request("Catecode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")
	prevDate = request("prevDate")
	prevTime = request("prevTime")

	sedatechk = request("sedatechk")

	if ((research="") and (view_yn="")) then 
	    view_yn = "1"
	    validdate = "on"
	end if
	
	if page="" then page=1
	if prevTime = "" then prevTime = "00"

dim oMainContents
	set oMainContents = new CMainContents
	oMainContents.FPageSize = 10
	oMainContents.FCurrPage = page
	oMainContents.FRectIsusing = view_yn
	oMainContents.FRectCatecode = Catecode
	oMainContents.FRectSelDate = prevDate
	oMainContents.FRectsedatechk= sedatechk '//시작일 기준 체크
	oMainContents.GetMainContentsList

dim i
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
function NextPage(page){
    frm.page.value = page;
    frm.submit();
}

function AddNewMainContents(idx){
	var dateOptionParam
	var frm = document.frm
	dateOptionParam = frm.prevDate.value

    var popwin = window.open('popmaincontentsedit.asp?idx=' + idx + '&dateoption=' + dateOptionParam,'mainposcodeedit','width=1200,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function tnCheckAll(bool, comp){
    var frm = comp.form;
    if (!comp.length){
        comp.checked = bool;
    }else{
        for (var i=0;i<comp.length;i++){
            comp[i].checked = bool;
        }
    }
}

function fnOrderidxEdit(){
var itemcount = 0;
var frm;
var ck=0;
frm = document.frmArrupdate;

	if(typeof(frm.cksel) !="undefined"){
		if(!frm.cksel.length){
			if(!frm.cksel.checked){
				alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
				return;
			}
				frm.orderidxarr.value = frm.cksel.value;
		}else{
			for(i=0;i<frm.cksel.length;i++){
				if(frm.cksel[i].checked) {
					ck=ck+1;	   	    			
					if (frm.orderidxarr.value==""){
						frm.idxarr.value =  frm.cksel[i].value;
						frm.orderidxarr.value =  frm.orderidx[i].value;
					}else{
						frm.idxarr.value = frm.idxarr.value + "," +frm.cksel[i].value;
						frm.orderidxarr.value = frm.orderidxarr.value + "," +frm.orderidx[i].value;
					} 
				}	
			}
			
			if (frm.orderidxarr.value == ""){
				alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
				return;
			}
		}
	}else{
		alert("추가할 상품이 없습니다.");
		return;
	}
	frm.target = "FrameCKP";
	frm.action = "doRollingOrderidx.asp";
	frm.submit();
}
</script>

<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="tabs" value="<%= request("tabs") %>">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
	    <input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >종료이전
	    &nbsp;카테고리&nbsp;
        <% DrawSelectBoxDispCateLarge "catecode", catecode, "" %>
        &nbsp;사용구분&nbsp;
		<select name="view_yn" class="select">
		<option value="">전체
		<option value="1" <% if view_yn="1" then response.write "selected" %> >사용함
		<option value="0" <% if view_yn="0" then response.write "selected" %> >사용안함
		</select>
        &nbsp;&nbsp;
		시작일기준 <input type="checkbox" name="sedatechk" <% if sedatechk="on" then response.write "checked" %> />
        지정일자 <input id="prevDate" name="prevDate" value="<%=prevDate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer;vertical-align:bottom"/>
		<script type="text/javascript">
			var CAL_Start = new Calendar({
				inputField : "prevDate", trigger    : "prevDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="left">
    	<input type="button" value="지정 순서 변경" onClick="fnOrderidxEdit();"/>
    </td>
    <td align="right">
    	<a href="javascript:AddNewMainContents('0');"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!-- 액션 끝 -->
<form name="frmArrupdate" method="post">
<input type="hidden" name="idxarr">
<input type="hidden" name="orderidxarr">
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		검색결과 : <b><%=oMainContents.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oMainContents.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td><input type="checkbox" name="validdate" onClick="tnCheckAll(this.checked,frmArrupdate.cksel);" ></td>
	<td>idx</td>
    <td>카테고리</td>
    <td>이미지</td>
    <td>시작일</td>
    <td>종료일</td>
    <td>사용여부</td>
    <td>우선순위</td>
    <td>등록자</td>
</tr>
<%
	for i=0 to oMainContents.FResultCount - 1
%>
<% if (oMainContents.FItemList(i).IsEndDateExpired) or (oMainContents.FItemList(i).Fview_yn="0") then %>
<tr bgcolor="#DDDDDD">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
	<td align="center"><input type="checkbox" name="cksel" id="cksel" value="<%= oMainContents.FItemList(i).Fidx %>"></td>
    <td align="center"><a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><%= oMainContents.FItemList(i).Fidx %></a></td>
    <td align="center"><a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><%= oMainContents.FItemList(i).Fcatename %></a></td>
    <td align="center">
    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).Fmain_image %>" border="0" width="300"></a>
    </td>
    <td align="center"><%= oMainContents.FItemList(i).FStart_date %></td>
    <td align="center">
    <% if (oMainContents.FItemList(i).IsEndDateExpired) then %>
    <font color="#777777"><%= Left(oMainContents.FItemList(i).FEnd_date,10) %></font>
    <% else %>
    <%= Left(oMainContents.FItemList(i).FEnd_date,10) %>
    <% end if %>
    </td>
    <td align="center"><%= oMainContents.FItemList(i).Fview_yn %></td>
    <td align="center">
		<input type="text" name="orderidx" class="formTxt" size=5 value="<%= oMainContents.FItemList(i).fview_order %>">
    </td>
    <td align="center"><%= oMainContents.FItemList(i).Freguserid %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="9" align="center">
    <% if oMainContents.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oMainContents.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oMainContents.StarScrollPage to oMainContents.FScrollCount + oMainContents.StarScrollPage - 1 %>
		<% if i>oMainContents.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oMainContents.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
</form>

<%
set oMainContents = Nothing
%>
<iframe name="FrameCKP" src="" frameborder="0" width="600" height="400"></iframe>
<form name="refreshFrm" method="post">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->