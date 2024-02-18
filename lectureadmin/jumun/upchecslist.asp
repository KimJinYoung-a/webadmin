<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<%

dim currstate,research,page,orderserial
dim searchfield, searchstring

research    = requestCheckVar(request("research"),10)
currstate   = requestCheckVar(request("currstate"),10)
page        = requestCheckVar(request("page"),10)

searchfield = requestCheckVar(request("searchfield"),32)
searchstring = requestCheckVar(request("searchstring"),32)

if page="" then page=1

if research="" then
	currstate="notfinish"
end if

if searchstring="" then searchfield="" end if
if searchfield="" then searchstring="" end if

dim ioneas,i
set ioneas = new CCSASList
ioneas.FPageSize = 20
ioneas.FCurrPage = page

if searchfield="01" then
	ioneas.FRectOrderserial = searchstring
elseif searchfield="02" then
	ioneas.FRectUserName = searchstring
elseif searchfield="03" then
	ioneas.FRectUserID = searchstring
end if


ioneas.FRectCurrstate  = currstate
ioneas.FRectSearchType = "upcheview"
ioneas.FRectMakerID = session("ssBctID")
ioneas.GetCSASMasterList

%>

<script language='javascript'>

function ShowOrderInfo(frm,orderserial){
    var props = "width=600, height=600, location=no, status=yes, resizable=no, scrollbars=yes";
	window.open("about:blank", "orderdetail", props);
    frm.target = "orderdetail";
    frm.orderserial.value = orderserial;
    frm.action="/designer/common/viewordermaster.asp";
	frm.submit();
}


function NextPage(page){
    frm.page.value = page;
    frm.submit();
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
				<option value="03" <% if searchfield="03" then response.write "selected" %>>고객ID</option>
			</select>
			<input type="text" class="text" name="searchstring" value="<%= searchstring %>" size="13" maxlength="11">

		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit()">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><% = ioneas.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= ioneas.FTotalPage %></b>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80">주문번호</td>
		<td width="50">고객명</td>
		<td width="100">고객ID</td>
		<td>제목</td>
		<td>접수사유</td>
		<td width="70">등록일</td>
		<td width="70">처리완료일</td>
		<td width="50">상태</td>
	</tr>
	<% for i=0 to ioneas.FresultCount-1 %>

	<tr align="center" bgcolor="#FFFFFF">
		<td><a href="javascript:ShowOrderInfo(frmshow,'<%= ioneas.FItemList(i).FOrderSerial %>');"><%= ioneas.FItemList(i).FOrderSerial %></a></td>
		<td><%= ioneas.FItemList(i).FCustomerName %></td>
		<td><%= ioneas.FItemList(i).FUserID %></td>
		<td align="left" ><a href="upchecsdetail.asp?idx=<%= ioneas.FItemList(i).Fid %>&menupos=<%= menupos %>"><%= ioneas.FItemList(i).FTitle %></a></td>
		<td align="left"><%= (ioneas.FItemList(i).Fgubun01Name) %>&gt;&gt;<%= (ioneas.FItemList(i).Fgubun02Name) %></td>
		<td><%= Left(CStr(ioneas.FItemList(i).Fregdate),10) %></td>
		<td>
			<% if ioneas.FItemList(i).Ffinishdate<>"" then %>
			<%= Left(CStr(ioneas.FItemList(i).Ffinishdate),10) %>
			<% else %>
			<input type="button" class="button" value="완료처리" onclick="location.href='upchecsdetail.asp?idx=<%= ioneas.FItemList(i).Fid %>&menupos=<%= menupos %>';">
			<% end if %>
		</td>
		<td><%= CsState2Name(ioneas.FItemList(i).FCurrstate) %></td>
	</tr>
	<% next %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	        <% if ioneas.HasPreScroll then %>
				<a href="javascript:NextPage('<%= CStr(ioneas.StarScrollPage - 1) %>')">[prev]</a>
			<% else %>
				[prev]
			<% end if %>
			<% for i = ioneas.StarScrollPage to (ioneas.StarScrollPage + ioneas.FScrollCount - 1) %>
			  <% if (i > ioneas.FTotalPage) then Exit For %>
			  <% if CStr(i) = CStr(ioneas.FCurrPage) then %>
				 [<%= i %>]
			  <% else %>
				 <a href="javascript:NextPage('<%= i %>')" class="id_link">[<%= i %>]</a>
			  <% end if %>
			<% next %>
			<% if ioneas.HasNextScroll then %>
				<a href="javascript:NextPage('<%= CStr(ioneas.StarScrollPage + ioneas.FScrollCount) %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
	    </td>
	</tr>
</table>

<form name="frmshow" method="post">
<input type="hidden" name="orderserial" value="">

</form>


<%
set ioneas = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->