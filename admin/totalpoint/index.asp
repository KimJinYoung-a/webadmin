<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  통합 회원 카드
' History : 2009.07.08 강준구 생성
'			2011.01.18 한용민 수정(페이징 클래스 방식으로 변경. 기존펑션내 페이징 잘못되어 있음)
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual = "/lib/util/htmllib.asp" -->
<!-- #include virtual = "/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual = "/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #Include virtual = "/lib/classes/totalpoint/totalpointCls.asp" -->

<%
Dim vUserName, vUserID, vJumin1, vCardNo, vUseYN, vCardGubun , ix,iPerCnt, vParam
Dim opoint ,i ,shopid, fromDate,toDate ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 , page, memberYn, userhp
	shopid 	= requestCheckVar(Request("shopid"),32)
	vUserName		= NullFillWith(requestCheckVar(Request("username"),20),"")
	vUserID			= NullFillWith(requestCheckVar(Request("userid"),32),"")
	vCardGubun		= NullFillWith(requestCheckVar(Request("cardgubun"),4),"")
	vCardNo			= NullFillWith(requestCheckVar(Request("cardno"),20),"")
	vUseYN			= NullFillWith(requestCheckVar(Request("useyn"),20),"")
	memberYn		= requestCheckVar(Request("memberYn"),1)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	page = requestCheckVar(request("page"),10)
	userhp		= requestCheckVar(Request("userhp"),16)

if page="" then page=1

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

if C_ADMIN_USER then
'/매장
elseif (C_IS_SHOP) then

	'//직영점일때
	if C_IS_OWN_SHOP then

		'/어드민권한 점장 미만
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/업체
	if (C_IS_Maker_Upche) then

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if

vParam = "&username="&vUserName&"&cardno="&vCardNo&"&userid="&vUserID&"&useyn="&vUseYN&"&shopid="&shopid&"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2

set opoint = new TotalPoint
	opoint.FPageSize=20
	opoint.FCurrPage=page
 	opoint.FUserName = vUserName
 	opoint.FUserID = vUserID
	opoint.FUseYN = vUseYN
 	opoint.FCardNo = vCardNo
 	opoint.FCardGubun = vCardGubun
 	opoint.frectshopid = shopid
	opoint.FRectStartDay = fromDate
	opoint.FRectEndDay = toDate
	opoint.frectmemberYn = memberYn
	opoint.frectuserhp = userhp
	opoint.GetTotalPointList
%>

<script language="javascript">

function goRead(userseq){
	if (userseq=="0" || userseq==""){
		alert('비회원이거나 정보가 없는 고객은 조회 하실수 없습니다.');
		return;
	}
	
	var popwin = window.open('point_detail.asp?userseq='+userseq+'<%=vParam%>','point_detail','width=650,height=527,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function jsGoPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}

</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 날짜 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;&nbsp;
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
				* 등록매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* 등록매장 : <% drawSelectBoxOffShopdiv_off "shopid",shopid,"1,3,7,11","","" %>
			<% end if %>
		<% else %>
			* 등록매장 : <% drawSelectBoxOffShopdiv_off "shopid",shopid,"1,3,7,11","","" %>
		<% end if %>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색" onClick="jsGoPage('');">
		<!--<input type="button" class="button_s" value="초기리스트" onClick="location.href='/admin/totalpoint/?menupos=<%=g_MenuPos%>'">-->
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 회원여부
		<% drawSelectBoxisusingYN "memberYn",memberYn,"" %>
		* 카드사용YN
		<% drawSelectBoxisusingYN "useyn",vUseYN,"" %>
		&nbsp;&nbsp;
		* 카드구분
		<select name="cardgubun" class="select">
			<option value="">전체</option>
			<option value="1010" <% If vCardGubun = "1010" Then %>selected<% End If %>>POINT1010</option>
			<option value="T" <% If vCardGubun <> "" AND vCardGubun <> "1010" AND vCardGubun <> "3253" Then %>selected<% End If %>>(구)오프라인</option>
			<option value="3253" <% If vCardGubun = "3253" Then %>selected<% End If %>>(구)아이띵소</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 고객명: <input type="text" class="text" name="username" value="<%=vUserName%>" size="8">
		&nbsp;&nbsp;
		* 아이디: <input type="text" class="text" name="userid" value="<%=vUserID%>" size="12">
		&nbsp;&nbsp;
		* 휴대폰번호: <input type="text" class="text" name="userhp" value="<%= userhp %>" size="16" maxlength=16>
		&nbsp;&nbsp;
		* 카드번호: <input type="text" class="text" name="cardno" value="<%=vCardNo%>">
	</td>
</tr>
</table>
</form>

<Br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= opoint.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= opoint.FTotalPage %></b>
	</td>
</tr>
<tr align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td>회원번호</td>
	<td>고객명</td>
	<td>아이디</td>
	<td>카드구분</td>
	<td>카드번호</td>
	<td>카드사용<br>YN</td>
	<td>통합포인트</td>
	<td>가입가맹점</td>
	<td>등록일</td>
	<td>비고</td>
</tr>
<%
if opoint.FresultCount > 0 then

for i=0 to opoint.FresultCount-1
%>
<tr align="center" height="25" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
	<td>
		<% if opoint.FItemList(i).fUserSeq<>"0" then %>
			<%= opoint.FItemList(i).fUserSeq %>
		<% else %>
			비회원
		<% end if %>
	</td>
	<td>
		<%
			If opoint.FItemList(i).fUserName <> "" Then
				If opoint.FItemList(i).fGrade <> "0" Then
					Response.Write "[특별]" & opoint.FItemList(i).fUserName
				Else
					Response.Write opoint.FItemList(i).fUserName
				End If
			Else
				Response.Write "&nbsp;"
			End If
		%>
	</td>
	<td><%= printUserId(opoint.FItemList(i).fOnlineUserID, 2, "*") %></td>
	<td>
		<% If Left(opoint.FItemList(i).fCardNo,4) = "1010" Then %>
			POINT1010
		<% ElseIf Left(opoint.FItemList(i).fCardNo,4) = "3253" Then %>
			아이띵소
		<% Else %>
			오프라인
		<% End If %>
	</td>
	<td><%= opoint.FItemList(i).fCardNo %></td>
	<td><%= opoint.FItemList(i).fUseYN %></td>
	<td align="right"><%=FormatNumber(opoint.FItemList(i).fPoint,0)%></td>
	<td>
		<% If opoint.FItemList(i).fshopname = "" Then %>
			온라인가입
		<% Else %>
			오프라인가입
			<br><%= opoint.FItemList(i).fshopname %>
		<% End If %>
	</td>
	<td><%=opoint.FItemList(i).fRegdate%></td>
	<td><input type="button" class="button" value="적립상세보기" onClick="goRead('<%=opoint.FItemList(i).fUserSeq%>')"></td>
</tr>
<% Next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if opoint.HasPreScroll then %>
			<span class="list_link"><a href="javascript:jsGoPage(<%= opoint.StartScrollPage-1 %>);">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + opoint.StartScrollPage to opoint.StartScrollPage + opoint.FScrollCount - 1 %>
			<% if (i > opoint.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(opoint.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:jsGoPage(<%= i %>);" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if opoint.HasNextScroll then %>
			<span class="list_link"><a href="javascript:jsGoPage(<%= i %>);">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>

<% Else %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
</tr>
<% End If %>

</table>

<%
set opoint = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!--#Include Virtual = "/lib/db/dbclose.asp" -->
