<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 마일리지내역
' History : 2009.04.07 서동석 생성
'			2010.03.26 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopmileagecls.asp"-->
<%
dim page,shopid ,fromDate,toDate ,yyyymmdd1,yyymmdd2 ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,ooffmilde ,i
dim makerid ,CurrencyUnit, CurrencyChar, ExchangeRate ,FmNum, inc3pl, logDesc
dim userid
	shopid = requestCheckVar(request("shopid"),32)
	page = requestCheckVar(request("page"),10)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	makerid = requestCheckVar(request("makerid"),32)
    inc3pl = requestCheckVar(request("inc3pl"),32)
	logDesc = requestCheckVar(request("logDesc"),128)
	userid = requestCheckVar(request("userid"),128)

if page="" then page=1

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-3)
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

'/매장
if (C_IS_SHOP) then

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
		makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
		else
		end if
	end if
end if

set ooffmilde = new COffShopMileage
	ooffmilde.FPageSize=100
	ooffmilde.FCurrpage=page
	ooffmilde.FRectStartDay = fromDate
	ooffmilde.FRectEndDay = toDate
	ooffmilde.FRectShopid=shopid
	ooffmilde.FRectInc3pl = inc3pl
	ooffmilde.FRectLogDesc = logDesc
	ooffmilde.FRectOnlineUserID = userid
	ooffmilde.COffShopMileageList

Call fnGetOffCurrencyUnit(shopid,CurrencyUnit, CurrencyChar, ExchangeRate)
FmNum = CHKIIF(CurrencyUnit="WON" or CurrencyUnit="KRW",0,2)
%>

<script language="javascript">

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* 기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				&nbsp;&nbsp;
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>
					<% if not C_IS_OWN_SHOP and shopid <> "" then %>
						* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
					<% end if %>
				<% else %>
					* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
				<% end if %>
	            &nbsp;&nbsp;
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
				&nbsp;&nbsp;
				* 적요 :
				<input type="text" class="text" size="15" name="logDesc" value="<%= logDesc %>">
				&nbsp;&nbsp;
				* 온라인ID(전체기간) :
				<input type="text" class="text" size="15" name="userid" value="<%= userid %>">
			</td>
		</tr>
		</table>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
	</td>
</tr>
</form>
</table>
<!-- 표 상단바 끝-->
<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
	<tr valign="bottom">
    <td align="left">
    </td>
    <td align="right">
    </td>
	</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<% if ooffmilde.FresultCount>0 then %>
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="20">
		검색결과 : <b><%= ooffmilde.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= ooffmilde.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>회원번호</td>
	<td>온라인ID</td>
	<td>회원명</td>
	<td>샾구분</td>
	<td>마일리지</td>
	<td>적요</td>
	<td>저장일</td>
</tr>

<%
For i = 0 To ooffmilde.FResultCount - 1
%>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= ooffmilde.FItemList(i).Fpointuserno %></td>
	<td ><%= ooffmilde.FItemList(i).Fonlineuserid %></td>
	<td ><%= ooffmilde.FItemList(i).Fpointusername %></td>
	<td><%= ooffmilde.FItemList(i).Fshopid %></td>
	<td align="right"><%= FormatNumber(ooffmilde.FItemList(i).Fpoint,FmNum) %></td>
	<td align="let"><%= ooffmilde.FItemList(i).Fjukyo %></td>
	<td><%= ooffmilde.FItemList(i).Fregdate %></td>
</tr>
<% Next %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan=20>
		<% if ooffmilde.HasPreScroll then %>
			<a href="javascript:frmsubmit('<%= ooffmilde.StarScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ooffmilde.StarScrollPage to ooffmilde.FScrollCount + ooffmilde.StarScrollPage - 1 %>
			<% if i>ooffmilde.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:frmsubmit('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ooffmilde.HasNextScroll then %>
			<a href="javascript:frmsubmit('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% ELSE %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="20">검색 결과가 없습니다.</td>
</tr>
<%END IF%>
</table>

<%
set ooffmilde = Nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
