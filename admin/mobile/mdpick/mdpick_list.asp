<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' Discription : 모바일 mdpick
' History : 2013.12.17 한용민
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/mobile/main/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/mdpick_cls.asp" -->
<%
Dim isusing, page, i, okeyword, reload, itemid, itemname, makerid, sellyn, usingyn, danjongyn
dim mwdiv, limityn, vatyn, sailyn, overSeaYn, itemdiv, cdl, cdm, cds, dispCate, acURL
	page = request("page")
	reload = request("reload")
	isusing = RequestCheckVar(request("isusing"),1)
	itemid      = requestCheckvar(request("itemid"),255)
	itemname    = request("itemname")
	makerid     = requestCheckvar(request("makerid"),32)
	sellyn      = requestCheckvar(request("sellyn"),10)
	usingyn     = requestCheckvar(request("usingyn"),10)
	danjongyn   = requestCheckvar(request("danjongyn"),10)
	mwdiv       = requestCheckvar(request("mwdiv"),10)
	limityn     = requestCheckvar(request("limityn"),10)
	vatyn       = requestCheckvar(request("vatyn"),10)
	sailyn      = requestCheckvar(request("sailyn"),10)
	overSeaYn   = requestCheckvar(request("overSeaYn"),10)
	itemdiv     = requestCheckvar(request("itemdiv"),10)
	cdl = requestCheckvar(request("cdl"),10)
	cdm = requestCheckvar(request("cdm"),10)
	cds = requestCheckvar(request("cds"),10)
	dispCate = requestCheckvar(request("disp"),16)

acURL =Server.HTMLEncode("/admin/mobile/mdpick/mdpick_process.asp?menupos="&menupos)
	
if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	arrTemp = Split(itemid,chr(13))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

if page="" then page=1
if reload="" and isusing="" then isusing="Y"

set okeyword = new cmdpick
	okeyword.FPageSize		= 100
	okeyword.FCurrPage		= page
	okeyword.Frectisusing			= isusing
	okeyword.FRectMakerid      = makerid
	okeyword.FRectItemid       = itemid
	okeyword.FRectItemName     = itemname
	okeyword.FRectSellYN       = sellyn
	okeyword.FRectitemIsUsing      = usingyn
	okeyword.FRectDanjongyn    = danjongyn
	okeyword.FRectLimityn      = limityn
	okeyword.FRectMWDiv        = mwdiv
	okeyword.FRectVatYn        = vatyn
	okeyword.FRectSailYn       = sailyn
	okeyword.FRectIsOversea	= overSeaYn
	okeyword.FRectCate_Large   = cdl
	okeyword.FRectCate_Mid     = cdm
	okeyword.FRectCate_Small   = cds
	okeyword.FRectDispCate		= dispCate
	okeyword.FRectItemDiv      = itemdiv
	okeyword.getmdpick_list()

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>

function totalCheck(){
	var f = document.frmlist;
	var objStr = "idx";
	var chk_flag = true;
	for(var i=0; i<f.elements.length; i++) {
		if(f.elements[i].name == objStr) {
			if(!f.elements[i].checked) {
				chk_flag = f.elements[i].checked;
				break;
			}
		}
	}

	for(var i=0; i < f.elements.length; i++) {
		if(f.elements[i].name == objStr) {
			if(chk_flag) {
				f.elements[i].checked = false;
			} else {
				f.elements[i].checked = true;
			}
		}
	}
}

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

// 상품추가(검색) 팝업
function addnewItem(){
	var popwin;
	popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?acURL=<%=acURL%>&menupos=<%=menupos%>", "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function mdpickedit(idx){
	var mdpickedit = window.open('/admin/mobile/mdpick/mdpick_edit.asp?idx='+idx+'&menupos=<%=menupos%>','mdpickedit','width=1024,height=768,scrollbars=yes,resizable=yes');
	mdpickedit.focus();
}

function AssignXmlReal(){
	if (confirm('모바일사이트 메인 페이지에 적용 하시겠습니까?')){
		 var popwin = window.open('','refreshFrm','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm";
		 refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_mdpick_xml.asp" ;
		 refreshFrm.submit();
	}
}

//주석처리
function AssignXmlAppl(term){
    if (!confirm('새로 반영하시겠습니까?')) return;
     
	 var popwin = window.open('','refreshFrm_Main','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main";
	 refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_mdpick_xml.asp?term=" + term;
	 refreshFrm.submit();
}

</script>

<img src="/images/icon_arrow_link.gif"> <b>MDPICK</b>
<p>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="reload" value="ON">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* MDPICK 사용여부 : <% DrawSelectBoxUsingYN "isusing",isusing %>
		<p>
		* 브랜드 : <% drawSelectBoxDesignerWithName "makerid", makerid %>
		&nbsp;&nbsp;
		* 상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;&nbsp;
		<span style="white-space:nowrap;">* 상품코드 :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(쉼표로 복수입력가능)</span>
		<p>
		<span style="white-space:nowrap;">* 판매 : <% drawSelectBoxSellYN "sellyn", sellyn %></span>
     	&nbsp;&nbsp;
     	<span style="white-space:nowrap;">* 상품사용 : <% drawSelectBoxUsingYN "usingyn", usingyn %></span>
     	&nbsp;&nbsp;
     	<span style="white-space:nowrap;">* 단종 : <% drawSelectBoxDanjongYN "danjongyn", danjongyn %></span>
     	&nbsp;&nbsp;
     	<span style="white-space:nowrap;">* 한정 : <% drawSelectBoxLimitYN "limityn", limityn %></span>
     	&nbsp;&nbsp;
     	<span style="white-space:nowrap;">* 거래구분 : <% drawSelectBoxMWU "mwdiv", mwdiv %></span>
     	&nbsp;&nbsp;
     	<span style="white-space:nowrap;">* 과세 : <% drawSelectBoxVatYN "vatyn", vatyn %></span>
     	&nbsp;&nbsp;
     	<span style="white-space:nowrap;">* 할인 : <% drawSelectBoxSailYN "sailyn", sailyn %></span>
     	&nbsp;&nbsp;
     	<span style="white-space:nowrap;">* 해외배송 : <% drawSelectBoxIsOverSeaYN "overSeaYn", overSeaYn %></span>
        &nbsp;&nbsp;
     	<span style="white-space:nowrap;">* 상품구분 : <% drawSelectBoxItemDiv "itemdiv", itemdiv %></span>
		<p>
		* 관리<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		&nbsp;&nbsp;전시카테고리 : <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:frmsubmit('');">
	</td>
</tr>
</form>
<form name="refreshFrm" method="post">
</form>
</table>
<!-- 검색 끝 -->

<br>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td>
		<a href="javascript:AssignXmlReal();"><img src="/images/refreshcpage.gif" border="0"> Real 적용</a>
		<!--오늘을 포함하여 <input type="text" name="vTerm" value="1" size="1" class="text" style="text-align:right;">일간
		<a href="javascript:AssignXmlAppl(document.all.vTerm.value);"><img src="/images/refreshcpage.gif" border="0"> XML Real 적용(예약)</a>-->
	</td>
    <td align="right">
    	<input type="button" value="상품추가(검색)" onclick="addnewItem();" class="button">
    </td>
</tr>
</table>

<!--  리스트 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		총 등록수 : <b><%=okeyword.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=okeyword.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<!--<td><input type="checkbox" name="ckall" onclick="totalCheck()"></td>-->
	<td>이미지</td>
	<td>상품코드</td>	
    <td>상품명</td>
    <!--<td>시작일</td>
    <td>종료일</td>-->
    <td>정렬순위</td>
    <td>사용여부</td>
    <td>비고</td>
</tr>
<form name="frmlist" method="post">
<%
if okeyword.FResultCount>0 then
	
for i=0 to okeyword.FResultCount - 1 
%>
<tr height="30" align="center" bgcolor="<%=chkIIF(okeyword.FItemList(i).fisusing="Y","#FFFFFF","#F0F0F0")%>">
	<!--<td><input type="checkbox" name="idx" value="<%=okeyword.FItemList(i).Fidx%>" onClick="AnCheckClick(this);"></td>-->
    <td>
    	<img src="<%= okeyword.FItemList(i).fbasicimage %>" width=50 height=50 />
	</td>
	<td>
		<%= okeyword.FItemList(i).fitemid %>
	</td>	
	<td>
		<%= okeyword.FItemList(i).Fitemname %>
	</td>
    <!--<td>
    	<% if okeyword.FItemList(i).FStartdate<>"" then %>
    		<%= okeyword.FItemList(i).FStartdate %>
    	<% end if %>
    </td>
    <td>
    	<% if okeyword.FItemList(i).FStartdate<>"" then %>
		    <% if (okeyword.FItemList(i).IsEndDateExpired) then %>
		    	<font color="#777777"><%= Left(okeyword.FItemList(i).FEnddate,10) %></font>
		    <% else %>
		    	<%= Left(okeyword.FItemList(i).FEnddate,10) %>
		    <% end if %>
    	<% end if %>		    
    </td>-->	
	<td>
		<%= okeyword.FItemList(i).forderno %>
	</td>
	<td><%= okeyword.FItemList(i).fisusing %></td>
	<td>
		<input type="button" onclick="mdpickedit('<%=okeyword.FItemList(i).Fidx%>')" value="수정" class="button">
	</td>
</tr>
<% Next %>

<tr bgcolor="#FFFFFF">
	<td align="center" colspan="20">
		<% if okeyword.HasPreScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= okeyword.StartScrollPage-1 %>')">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + okeyword.StartScrollPage to okeyword.StartScrollPage + okeyword.FScrollCount - 1 %>
			<% if (i > okeyword.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(okeyword.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:frmsubmit('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if okeyword.HasNextScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= i %>')">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</form>
</table>

<%
set okeyword = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->