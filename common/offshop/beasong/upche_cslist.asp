<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 오프라인 cs내역
' Hieditor : 2011.03.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->

<%
dim currstate,research,page,orderno ,searchfield, searchstring , orgmasteridx
	research    = requestCheckVar(request("research"),10)
	currstate   = requestCheckVar(request("currstate"),10)
	page        = requestCheckVar(request("page"),10)
	orgmasteridx        = requestCheckVar(request("orgmasteridx"),10)
	searchfield = requestCheckVar(request("searchfield"),32)
	searchstring = requestCheckVar(request("searchstring"),32)

	if page="" then page=1
	
	if research="" then
		currstate="notfinish"
	end if
	
	if searchstring="" then searchfield="" end if
	if searchfield="" then searchstring="" end if

dim ioneas,i
set ioneas = new corder
	ioneas.FPageSize = 20
	ioneas.FCurrPage = page
	
	if searchfield="01" then
		ioneas.FRectorderno = searchstring
	elseif searchfield="02" then
		ioneas.FRectUserName = searchstring
	elseif searchfield="03" then
		ioneas.frectmasteridx = searchstring
	end if
	      
	ioneas.FRectCurrstate  = currstate
	ioneas.FRectSearchType = "upcheview"
	ioneas.FRectMakerID = session("ssBctID")
	ioneas.fGetCSASMasterList
%>

<script language='javascript'>

function ShowOrderInfo(frm,orgmasteridx){
    var props = "width=600, height=600, location=no, status=yes, resizable=no, scrollbars=yes";
	window.open("about:blank", "orderview", props);
	
    frm.target = "orderview";
    frm.masteridx.value = orgmasteridx;
    frm.action="/common/offshop/beasong/upche_viewordermaster.asp";
	frm.submit();
}


function NextPage(page){
    frm.page.value = page;
    frm.submit();
}

function regdetail(csmasteridx,menupos){
	location.href='/common/offshop/beasong/upche_csdetail.asp?csmasteridx=' + csmasteridx + '&menupos=' + menupos
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="T">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		상태:
		<select class="select" name="currstate">
	     	<option value='' selected>전체</option>
	     	<option value='notfinish' <% if currstate="notfinish" then response.write "selected" %>>미처리</option>
	     	<option value='B007' <% if currstate="B007" then response.write "selected" %>>처리완료</option>
     	</select>
		&nbsp;
		기타검색:
		<select class="select" name="searchfield">
			<option value="">검색조건</option>
			<option value="01" <% if searchfield="01" then response.write "selected" %>>주문번호</option>
			<option value="02" <% if searchfield="02" then response.write "selected" %>>고객명</option>
			<option value="03" <% if searchfield="03" then response.write "selected" %>>일렬번호</option>
		</select>
		<input type="text" class="text" name="searchstring" value="<%= searchstring %>" size="16" maxlength="16">
		
	</td>
	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit()">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><% = ioneas.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= ioneas.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>주문번호</td>
	<td>판매매장</td>	
	<td>고객명</td>
	<td>제목</td>	
	<td>등록일</td>
	<td>처리완료일</td>
	<td>상태</td>
	<td>비고</td>
</tr>
<% if ioneas.FresultCount > 0 then %>
<% for i=0 to ioneas.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><a href="javascript:ShowOrderInfo(frmshow,'<%= ioneas.FItemList(i).forgmasteridx %>');"><%= ioneas.FItemList(i).Forderno %></a></td>
	<td><%= ioneas.FItemList(i).fshopname %></td>
	<td><%= ioneas.FItemList(i).FCustomerName %></td>	
	<td align="left"><%= ioneas.FItemList(i).FTitle %></td>	
	<td><%= Left(CStr(ioneas.FItemList(i).Fregdate),10) %></td>
	<td>
		<% if ioneas.FItemList(i).Ffinishdate<>"" then %>
			<%= Left(CStr(ioneas.FItemList(i).Ffinishdate),10) %>
		<% end if %>	
	</td>	
	<td><%= CsState2Name_off(ioneas.FItemList(i).FCurrstate) %></td>
	<td><input type="button" class="button" value="처리상세" onclick="regdetail('<%= ioneas.FItemList(i).Fmasteridx %>','<%= menupos %>');"></td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
<td colspan="15" align="center">
    <% if ioneas.HasPreScroll then %>
		<a href="javascript:NextPage('<%= CStr(ioneas.StartScrollPage - 1) %>')">[prev]</a>
	<% else %>
		[prev]
	<% end if %>
	<% for i = ioneas.StartScrollPage to (ioneas.StartScrollPage + ioneas.FScrollCount - 1) %>
	  <% if (i > ioneas.FTotalPage) then Exit For %>
	  <% if CStr(i) = CStr(ioneas.FCurrPage) then %>
		 [<%= i %>]
	  <% else %>
		 <a href="javascript:NextPage('<%= i %>')" class="id_link">[<%= i %>]</a>
	  <% end if %>
	<% next %>
	<% if ioneas.HasNextScroll then %>
		<a href="javascript:NextPage('<%= CStr(ioneas.StartScrollPage + ioneas.FScrollCount) %>')">[next]</a>
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
<form name="frmshow" method="post">
	<input type="hidden" name="masteridx" value="">
</form>	
</table>

<%
set ioneas = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->