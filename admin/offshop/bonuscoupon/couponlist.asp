<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  보너스 쿠폰
' History : 2011.05.12 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/bonuscoupon/bonuscoupon_cls.asp" -->
<%
dim ocoupon, page , i , validsitename ,limityn ,couponname
	page = requestCheckVar(request("page"),10)
	validsitename = requestCheckVar(request("validsitename"),32)
	couponname = requestCheckVar(request("couponname"),128)
	limityn = requestCheckVar(request("limityn"),1)

if page="" then page=1

set ocoupon = new CCouponlist
	ocoupon.FPageSize=50
	ocoupon.FCurrPage = page
	ocoupon.frectvalidsitename = validsitename
	ocoupon.frectcouponname = couponname
	ocoupon.frectlimityn = limityn
	ocoupon.GetCouponMasterList()
%>

<script type='text/javascript'>

function reg() {
	frm_search.submit();
}

// 페이지 이동
function goPage(pg)
{
	document.frm_search.page.value=pg;
	document.frm_search.submit();
}

function couponedit(idx){
	location.href='couponreg.asp?idx='+idx+'&menupos=<%=menupos%>';
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_search" method="GET" action="" onSubmit="return false">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		구분 : <% Drawvalidsitename "validsitename" , validsitename, " onchange='reg();'","Y" %>
		<!--
		쿠폰사용횟수 : <% Drawlimityn "limityn",limityn," onchange='reg();'","Y" %>
		-->
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="reg();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		쿠폰명 : <input type="text" name="couponname" value="<%=couponname%>" size=30>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="center">
		<font color="red">테스트 중입니다.</font>
	</td>
	<td align="right">
		<input type="button" class="button" value="신규등록" onClick="couponedit('');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ocoupon.FResultCount %></b>
	</td>
</tr>
<tr height="30" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>IDX</td>
	<td>구분</td>
	<td>쿠폰명</td>
	<td>사용혜택</td>
	<td>최소<br>구매 금액</td>
	<td>유효기간</td>
	<td>등록일</td>
	<td>발급마감일</td>
	<td>비고</td>
</tr>
<% if ocoupon.FResultCount > 0 then %>
<% for i=0 to ocoupon.FResultCount - 1 %>
<% if ocoupon.FItemList(i).fisusing = "Y" then %>
<tr height="30" align="center" bgcolor="#FFFFFF">
<% else %>
<tr height="30" align="center" bgcolor="silver">
<% end if %>
	<td><%= ocoupon.FItemList(i).FIdx %></td>
	<td>
		<%= validsitenameview(ocoupon.FItemList(i).fvalidsitename) %>
	</td>
	<td><a href="javascript:couponedit('<%= ocoupon.FItemList(i).FIdx %>')"><%= ocoupon.FItemList(i).Fcouponname %></a></td>
	<td><%= ocoupon.FItemList(i).getCouponTypeStr %></td>
	<td align="center"><%= FormatNumber(ocoupon.FItemList(i).Fminbuyprice, 0) %></td>
	<td>
		<acronym title="<%= ocoupon.FItemList(i).Fstartdate %>"><%= Left(ocoupon.FItemList(i).Fstartdate,10) %></acronym>
		~
		<acronym title="<%= ocoupon.FItemList(i).Fexpiredate %>"><%= Left(ocoupon.FItemList(i).Fexpiredate,10) %></acronym>
	</td>
	<td>
		<acronym title="<%= ocoupon.FItemList(i).FRegDate %>"><%= Left(ocoupon.FItemList(i).FRegDate,10) %></acronym>
	</td>
	<td>
		<acronym title="<%= ocoupon.FItemList(i).FOpenFinishDate %>"><%= Left(ocoupon.FItemList(i).FOpenFinishDate,10) %></acronym>
	</td>
	<td align="center">
		<font color="<%= ocoupon.FItemList(i).getCouponStateColor %>"><%= ocoupon.FItemList(i).getCouponStateStr %></font>
	</td>
</tr>
<% next %>
<tr height="30" bgcolor="#FFFFFF">
	<td colspan="10" align="center">
		<!-- 페이지 시작 -->
		<%
		if ocoupon.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & ocoupon.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for i=0 + ocoupon.StartScrollPage to ocoupon.FScrollCount + ocoupon.StartScrollPage - 1

			if i>ocoupon.FTotalpage then Exit for

			if CStr(page)=CStr(i) then
				Response.Write " <font color='red'>" & i & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & i & ")'>" & i & "</a> "
			end if

		next

		if ocoupon.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
		%>
		<!-- 페이지 끝 -->
	</td>
</tr>
<% else %>
	<tr height="30" bgcolor="#FFFFFF">
		<td colspan="10" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>

<%
	set ocoupon = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->