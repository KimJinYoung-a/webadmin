<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_coincls.asp"-->

<%
dim userid, vGubun

	userid = request("userid")
	vGubun = request("gb")

	dim cMomoMngCoinList, i, vTotalCount, vTotalSum
	vTotalCount = 0
	vTotalSum	= 0

	'### 내가 사용 코인 내역
	set cMomoMngCoinList = new ClsMomoCoin
	
	If vGubun = "corner" Then
		cMomoMngCoinList.FGubun = vGubun
	Else
		cMomoMngCoinList.FUserID = userid
	End If
	cMomoMngCoinList.FUserCoinLogList
%>
<table cellpadding="0" cellspacing="0" border="0" class="a" width="100%">
<tr height="30" bgcolor="FFFFFF">
	<td width="50%"><% If vGubun <> "corner" Then %>회원아이디 : <%=userid%><% End If %></td>
	<td width="50%" align="right"><input type="button" class="button" value="닫기" onClick="window.close()"></td>
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
	    <td align="center" width="200">코너명(코인)</td>
	    <td align="center" width="120">등록횟수</td>
	    <td align="center" width="120">지급코인</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
	<% for i=0 to cMomoMngCoinList.FResultCount - 1 %>
	<tr bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">	
	    <td style="padding:0 5 0 5"><%= cMomoMngCoinList.FItemList(i).fgubuntitle %></td>
	    <td align="center">
	    <%
	    	If cMomoMngCoinList.FItemList(i).fgubuncd = "13" Then
	    		Response.Write cMomoMngCoinList.FItemList(i).fcdcount
	    	Else
	    		Response.Write MomoTotalCount(cMomoMngCoinList.FItemList(i).fgubuncd,cMomoMngCoinList.FItemList(i).fcoin)
	    	End If
	    %>
	    </td>
	    <td align="center"><%= FormatNumber(cMomoMngCoinList.FItemList(i).fcoin,0) %></td>
	</tr>
	<%
		vTotalCount = vTotalCount + cMomoMngCoinList.FItemList(i).fcdcount
		vTotalSum	= vTotalSum + cMomoMngCoinList.FItemList(i).fcoin
		next %>
    </tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
    <td align="center" width="200">합계</td>
    <td align="center" width="120"><%=FormatNumber(vTotalCount,0)%></td>
    <td align="center" width="120"><%=FormatNumber(vTotalSum,0)%></td>
</tr>
</table>

<br>

<%
cMomoMngCoinList.FGubun2 = "prodcoupon"
cMomoMngCoinList.FUserCoinLogList

if cMomoMngCoinList.FResultCount > 0 then %>
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td align="center" width="200">구분</td>
    <td align="center" width="120">교환횟수</td>
    <td align="center" width="120">사용코인</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
<% for i=0 to cMomoMngCoinList.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">	
    <td style="padding:0 5 0 5"><%= cMomoMngCoinList.FItemList(i).fgubuntitle %></td>
    <td align="center"><%= FormatNumber(cMomoMngCoinList.FItemList(i).fcdcount,0) %></td>
    <td align="center"><%= FormatNumber(cMomoMngCoinList.FItemList(i).fcoin,0) %></td>
</tr>
<%
	vTotalCount = vTotalCount + cMomoMngCoinList.FItemList(i).fcdcount
	vTotalSum	= vTotalSum + cMomoMngCoinList.FItemList(i).fcoin
	next %>
</tr>
</table>
<% end if %>

<%
	set cMomoMngCoinList = nothing	
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
