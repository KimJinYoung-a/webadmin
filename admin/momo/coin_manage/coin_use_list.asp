<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_coincls.asp"-->

<%
dim research,userid, fixtype, linktype, poscode, validdate
dim page, vGubun

	vGubun = request("gubun")
	userid = request("userid")
	research= request("research")
	poscode = request("poscode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")

If vGubun = "" Then vGubun = "i" End If
if page = "" then page = 1 End If

	dim cMomoMngCoinList, PageSize , ttpgsz , CurrPage, i
	CurrPage = requestCheckVar(request("cpg"),9)

	IF CurrPage = "" then CurrPage=1


	'### 내가 사용 코인 내역
	set cMomoMngCoinList = new ClsMomoCoin
	cMomoMngCoinList.FPageSize = 30
	cMomoMngCoinList.FCurrPage = page
	cMomoMngCoinList.FGubun = vGubun
	cMomoMngCoinList.FCoinUseList
%>

<table cellpadding="0" cellspacing="0" border="0" class="a">
<tr height="50">
<% If vGubun = "i" Then %>
	<td><b><u><a href="?menupos=<%=Request("menupos")%>&gubun=i">[상품교환현황]</a></u></b></td>
	<td style="padding-left:30"><a href="?menupos=<%=Request("menupos")%>&gubun=c">[쿠폰교환현황]</a></td>
	<td style="padding-left:30"><a href="coin_log_list.asp?menupos=<%=Request("menupos")%>&gubun=l">[코인적립현황]</a></td>
<% ElseIf vGubun = "c" Then %>
	<td><a href="?menupos=<%=Request("menupos")%>&gubun=i">[상품교환현황]</a></td>
	<td style="padding-left:30"><b><u><a href="?menupos=<%=Request("menupos")%>&gubun=c">[쿠폰교환현황]</a></u></b></td>
	<td style="padding-left:30"><a href="coin_log_list.asp?menupos=<%=Request("menupos")%>&gubun=l">[코인적립현황]</a></td>
<% Else %>
	<td><a href="?menupos=<%=Request("menupos")%>&gubun=i">[상품교환현황]</a></td>
	<td style="padding-left:30"><a href="?menupos=<%=Request("menupos")%>&gubun=c">[쿠폰교환현황]</a></td>
	<td style="padding-left:30"><b><u><a href="coin_log_list.asp?menupos=<%=Request("menupos")%>&gubun=l">[코인적립현황]</a></u></b></td>
<% End If %>
</tr>
</table>

<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if cMomoMngCoinList.FResultCount > 0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
			<tr>
				<td>검색결과 : <b><%= cMomoMngCoinList.FTotalCount %></b></td>
				<td align="right">
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td align="center" width="50"><% If vGubun = "i" Then %>주문번호<% Else %>idx<% End If %></td>
	    <td align="center" width="100">userid</td>
	    <td align="center" width="100">Coin</td>
	    <td align="center" width="300">내역</td>
	    <td align="center" width="150">등록일</td>
	    <td align="center" width="150">출고일</td>
	    <td align="center" width="200"></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
	<% for i=0 to cMomoMngCoinList.FResultCount - 1 %>
	<tr bgcolor="#FFFFFF">	
	    <td align="center"><%= cMomoMngCoinList.FItemList(i).fidx %></td>
	    <td align="center"><%= cMomoMngCoinList.FItemList(i).fuserid %></td>
	    <td align="center"><%= cMomoMngCoinList.FItemList(i).fcoin %></td>
	    <td align="center"><%= cMomoMngCoinList.FItemList(i).fitemname %>(<%= cMomoMngCoinList.FItemList(i).foptionname %>)</td>
	    <td align="center"><%= cMomoMngCoinList.FItemList(i).fregdate %></td>
	    <td align="center"><%= cMomoMngCoinList.FItemList(i).foutputdate %><br>송장:<%= cMomoMngCoinList.FItemList(i).fsongjangno %></td>
		<td align="center">
		<% If vGubun = "i" Then %><% Else %><%= Replace(cMomoMngCoinList.FItemList(i).fetc,"~","<br>~") %><% End If %>
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
				<span class="list_link"><a href="?page=<%= cMomoMngCoinList.StartScrollPage-1 %>&gubun=<%=vGubun%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + cMomoMngCoinList.StartScrollPage to cMomoMngCoinList.StartScrollPage + cMomoMngCoinList.FScrollCount - 1 %>
				<% if (i > cMomoMngCoinList.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(cMomoMngCoinList.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&gubun=<%=vGubun%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if cMomoMngCoinList.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&gubun=<%=vGubun%>">[next]</a></span>
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
