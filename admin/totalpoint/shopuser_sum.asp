<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 매장 회원 카드 정보
' Hieditor : 2011.01.18 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #Include virtual = "/lib/classes/totalpoint/totalpointCls.asp" -->

<%
dim opoint , i , shopid , memberyn, fromDate,toDate ,yyyy1,mm1,dd1,yyyy2,mm2,dd2, inc3pl
	shopid = request("shopid")
	memberyn = request("memberyn")
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
    inc3pl = request("inc3pl")

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

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if

set opoint = new TotalPoint
 	opoint.frectshopid = shopid
 	opoint.frectmemberyn = memberyn
	opoint.FRectStartDay = fromDate
	opoint.FRectEndDay = toDate
	opoint.FRectInc3pl = inc3pl
	opoint.fshopuser_sum()

%>

<script language="javascript">

function popshopuser(shopid,yyyy1,mm1,dd1,yyyy2,mm2,dd2){
	var popshopuser = window.open('/admin/totalpoint/index.asp?shopid='+shopid+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&menupos=<%=menupos%>','popshopuser','width=1024,height=768,scrollbars=yes,resizable=yes');
	popshopuser.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;&nbsp;
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
				* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
			<% end if %>
		<% else %>
			* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		<% end if %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
        <b>* 매출처구분</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
		<!--회원여부:
		<select name="memberyn">
			<option value="" <% if memberyn = "" then response.write " selected" %>>선택</option>
			<option value="Y" <% if memberyn = "Y" then response.write " selected" %>>회원</option>
			<option value="N" <% if memberyn = "N" then response.write " selected" %>>비회원</option>
		</select>-->
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<Br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		※ 통합이전 (구)카드 내역으로 인하여, 총카드등록 수량과 회원등록&비회원등록 합산 수량은 오차가 있습니다
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= opoint.FresultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">매장</td>
	<td align="center">총카드<Br>등록</td>
	<td align="center">회원<Br>등록</td>
	<td align="center">비회원<Br>등록</td>
</tr>
<% if opoint.FresultCount > 0 then %>
<% for i=0 to opoint.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
	<td align="center">
		<%= opoint.FItemList(i).fshopname %> (<%= opoint.FItemList(i).fregshopid %>)
	</td>
	<td align="center">
		<%= opoint.FItemList(i).fusercount %>
	</td>
	<td align="center">
		<% if opoint.FItemList(i).fmemberY <> 0 then %>
			<a href="javascript:popshopuser('<%= opoint.FItemList(i).fregshopid %>','<%= yyyy1 %>','<%= mm1 %>','<%= dd1 %>','<%= yyyy2 %>','<%= mm2 %>','<%= dd2 %>')" onfocus="this.blur()">
			<%= opoint.FItemList(i).fmemberY %></a>
		<% else %>
			<%= opoint.FItemList(i).fmemberY %>
		<% end if %>
	</td>
	<td align="center">
		<%= opoint.FItemList(i).fmemberN %>
	</td>
</tr>
</form>
<% next %>

<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
</table>

<%
set opoint = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
