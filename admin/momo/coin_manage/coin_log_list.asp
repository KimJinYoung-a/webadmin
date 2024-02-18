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
dim page, vGubun, vPlusMinus, vSDate, vEDate, vSort, vDelJikwon

	vGubun = request("gubun")
	userid = request("userid")
	research= request("research")
	poscode = request("poscode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")
	vPlusMinus = request("plusminus")
	vSDate = request("sdate")
	vEDate = request("edate")
	vSort = request("sort")
	vDelJikwon = request("deljikwon")
	If vSort = "" Then
		vSort = "now"
	End IF
	If vDelJikwon = "" Then
		vDelJikwon = "o"
	End IF

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
	cMomoMngCoinList.FUserID = userid
	cMomoMngCoinList.FPlusMinus = vPlusMinus
	cMomoMngCoinList.FSDate = vSDate
	cMomoMngCoinList.FEDate = vEDate
	cMomoMngCoinList.FSort = vSort
	cMomoMngCoinList.FDeljikwon = vDelJikwon
	cMomoMngCoinList.FCoinLogList
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
function goDetail(userid)
{
	var detail;
	detail = window.open('coin_log_list_pop.asp?userid='+userid+'','logpop','width=500,height=450,scrollbars=yes')
	detail.focus();
}
function delchk()
{
	if(frm.del_tmp.checked)
	{
		frm.deljikwon.value = "o";
	}
	else
	{
		frm.deljikwon.value = "x";
	}
}
function popcorner()
{
	var corner;
	corner = window.open('coin_log_list_pop.asp?gb=corner','cornerpop','width=500,height=450,scrollbars=yes')
	corner.focus();
}
</script>

<table cellpadding="0" cellspacing="0" border="0" class="a">
<tr height="50">
<% If vGubun = "i" Then %>
	<td><b><u><a href="coin_use_list.asp?menupos=<%=Request("menupos")%>&gubun=i">[상품교환현황]</a></u></b></td>
	<td style="padding-left:30"><a href="coin_use_list.asp?menupos=<%=Request("menupos")%>&gubun=c">[쿠폰교환현황]</a></td>
	<td style="padding-left:30"><a href="coin_log_list.asp?menupos=<%=Request("menupos")%>&gubun=l">[코인적립현황]</a></td>
<% ElseIf vGubun = "c" Then %>
	<td><a href="coin_use_list.asp?menupos=<%=Request("menupos")%>&gubun=i">[상품교환현황]</a></td>
	<td style="padding-left:30"><b><u><a href="coin_use_list.asp?menupos=<%=Request("menupos")%>&gubun=c">[쿠폰교환현황]</a></u></b></td>
	<td style="padding-left:30"><a href="coin_log_list.asp?menupos=<%=Request("menupos")%>&gubun=l">[코인적립현황]</a></td>
<% Else %>
	<td><a href="coin_use_list.asp?menupos=<%=Request("menupos")%>&gubun=i">[상품교환현황]</a></td>
	<td style="padding-left:30"><a href="coin_use_list.asp?menupos=<%=Request("menupos")%>&gubun=c">[쿠폰교환현황]</a></td>
	<td style="padding-left:30"><b><u><a href="coin_log_list.asp?menupos=<%=Request("menupos")%>&gubun=l">[코인적립현황]</a></u></b></td>
<% End If %>
</tr>
</table>

<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="gubun" value="<%=vGubun%>">
<input type="hidden" name="deljikwon" value="<%=vDelJikwon%>">
<tr height="30" bgcolor="FFFFFF">
	<td>
		<table cellpadding="0" cellspacing="0" class="a">
		<tr bgcolor="FFFFFF">
			<td>
				아이디 : <input type="text" name="userid" value="<%=userid%>" size="10">&nbsp;&nbsp;&nbsp;
				<input type="radio" name="plusminus" value="p" <% If vPlusMinus = "p" Then Response.Write " checked" End If %>>적립날짜
				<input type="radio" name="plusminus" value="m" <% If vPlusMinus = "m" Then Response.Write " checked" End If %>>사용날짜
				<input id="sdate" name="sdate" value="<%=vSDate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
				<input id="edate" name="edate" value="<%=vEDate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="edate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "sdate", trigger    : "sdate_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "edate", trigger    : "edate_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
				<input type="submit" value="검색">
			</td>
		</tr>
		<tr>
			<td><input type="checkbox" name="del_tmp" onClick="delchk()" value="o" <% If vDelJikwon = "o" Then Response.Write " checked" End If %>>직원제외</td>
		</tr>
		</table>
	</td>
	<td>
		<select name="sort" onChange="frm.submit();">
			<option value="now" <% If vSort = "now" Then Response.Write "selected" End If %>>현재코인</option>
			<option value="use" <% If vSort = "use" Then Response.Write "selected" End If %>>사용코인</option>
			<option value="save" <% If vSort = "save" Then Response.Write "selected" End If %>>누적코인</option>
		</select>
	</td>
	<td>
		<input type="button" value="코너별 코인적립현황" onClick="popcorner()">
	</td>
</tr>
</form>
</table>
<br>
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if cMomoMngCoinList.FResultCount > 0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
			<tr>
				<td>검색결과 : <b><%= cMomoMngCoinList.FTotalCount %></b></td>
				<td align="right">&nbsp;</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td align="center" width="100">userid</td>
	    <td align="center" width="120">현재코인</td>
	    <td align="center" width="120">사용코인</td>
	    <td align="center" width="120">누적코인</td>
	    <td align="center"></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
	<% for i=0 to cMomoMngCoinList.FResultCount - 1 %>
	<tr bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">	
	    <td align="center"><%= cMomoMngCoinList.FItemList(i).fuserid %></td>
	    <td align="center"><%= FormatNumber(cMomoMngCoinList.FItemList(i).fnowcoin,0) %></td>
	    <td align="center"><%= FormatNumber(cMomoMngCoinList.FItemList(i).fcurrentcoin,0) %></td>
	    <td align="center"><%= FormatNumber(cMomoMngCoinList.FItemList(i).fsavecoin,0) %></td>
	    <td align="center"><input type="button" value="상세보기" onClick="goDetail('<%= cMomoMngCoinList.FItemList(i).fuserid %>')"></td>
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
				<span class="list_link"><a href="?menupos=<%=menupos%>&gubun=<%=vGubun%>&plusminus=<%=vPlusMinus%>&sdate=<%=vSDate%>&edate=<%=vEDate%>&sort=<%=vSort%>&deljikwon=<%=vDelJikwon%>&page=<%= cMomoMngCoinList.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + cMomoMngCoinList.StartScrollPage to cMomoMngCoinList.StartScrollPage + cMomoMngCoinList.FScrollCount - 1 %>
				<% if (i > cMomoMngCoinList.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(cMomoMngCoinList.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?menupos=<%=menupos%>&gubun=<%=vGubun%>&plusminus=<%=vPlusMinus%>&sdate=<%=vSDate%>&edate=<%=vEDate%>&sort=<%=vSort%>&deljikwon=<%=vDelJikwon%>&page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if cMomoMngCoinList.HasNextScroll then %>
				<span class="list_link"><a href="?menupos=<%=menupos%>&gubun=<%=vGubun%>&plusminus=<%=vPlusMinus%>&sdate=<%=vSDate%>&edate=<%=vEDate%>&sort=<%=vSort%>&deljikwon=<%=vDelJikwon%>&page=<%= i %>">[next]</a></span>
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
