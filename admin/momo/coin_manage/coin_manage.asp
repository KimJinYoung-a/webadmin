<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모
' Hieditor : 2009.11.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_coincls.asp"-->

<%
dim research,userid, fixtype, linktype, poscode, validdate
dim page

	userid = request("userid")
	research= request("research")
	poscode = request("poscode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")

if page = "" then page = 1

	dim cMomoMngCoinList, PageSize , ttpgsz , CurrPage, i
	CurrPage = requestCheckVar(request("cpg"),9)

	IF CurrPage = "" then CurrPage=1
	if page = "" then page = 1
	

	'### 내가 사용 코인 내역
	set cMomoMngCoinList = new ClsMomoCoin
	cMomoMngCoinList.FPageSize = 30
	cMomoMngCoinList.FCurrPage = page
	cMomoMngCoinList.FCoinMngList
%>

<!-- 검색 시작
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
		    아이디:<input type="text" name="userid" value="<%=userid%>" size="10">
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

		</td>
	</tr>
	</form>
</table>
검색 끝 -->

<br>

<!-- 리스트 시작 -->
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if cMomoMngCoinList.FResultCount > 0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
			<tr>
				<td>검색결과 : <b><%= cMomoMngCoinList.FTotalCount %></b></td>
				<td align="right">
					<input type="button" value="상품관리" onClick="javascript:window.open('pop_prod_list.asp','prod','width=800,height=500,scrollbars=yes');">&nbsp;&nbsp;&nbsp;
					<input type="button" value="신규등록" onClick="javascript:window.open('coin_manage_write.asp','mng','width=400,height=150');">
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td align="center" width="50">idx</td>
	    <td align="center" width="100">Coin</td>
	    <td align="center" width="100">사용여부</td>
	    <td align="center" width="150">등록일</td>
	    <td align="center" width="100"></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
	<% for i=0 to cMomoMngCoinList.FResultCount - 1 %>
	<tr bgcolor="#FFFFFF">	
	    <td align="center"><%= cMomoMngCoinList.FItemList(i).fidx %></td>
	    <td align="center"><%= cMomoMngCoinList.FItemList(i).fcoin %></td>
	    <td align="center">
	    <%
	    	If cMomoMngCoinList.FItemList(i).fuseyn = "y" Then
	    		Response.Write "<b><font color='blue'>" & cMomoMngCoinList.FItemList(i).fuseyn & "</font></b>"
	    	Else
	    		Response.Write cMomoMngCoinList.FItemList(i).fuseyn
	    	End If
	    %>
	    </td>
	    <td align="center"><%= cMomoMngCoinList.FItemList(i).fregdate %></td>
		<td align="center">
			<input type="button" value="수정" onClick="javascript:window.open('coin_manage_write.asp?idx=<%= cMomoMngCoinList.FItemList(i).fidx %>','mng','width=400,height=200');">
			<input type="button" value="Item" onClick="javascript:window.open('coin_manage_item.asp?mng_idx=<%= cMomoMngCoinList.FItemList(i).fidx %>','mng','width=700,height=550,scrollbars=yes');">
		</td>
	</tr>
	<% next %>
    </tr>   
    
<% else %>

	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if cMomoMngCoinList.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= cMomoMngCoinList.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + cMomoMngCoinList.StartScrollPage to cMomoMngCoinList.StartScrollPage + cMomoMngCoinList.FScrollCount - 1 %>
				<% if (i > cMomoMngCoinList.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(cMomoMngCoinList.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if cMomoMngCoinList.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
	set cMomoMngCoinList = nothing	
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
