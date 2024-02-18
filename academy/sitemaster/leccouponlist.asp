<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'#######################################################
' Description : 핑거스 보너스쿠폰 관리
' History	:  최초생성자 모름
'              2017.07.07 한용민 수정
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/lecCouponCls.asp" -->
<%
dim ocoupon, page

page = RequestCheckvar(request("page"),10)
if page="" then page=1

set ocoupon = new CCouponMaster
ocoupon.FPageSize=60
ocoupon.FCurrPage = page
ocoupon.GetLecCouponList


dim i
%>
<table width="100%" cellspacing="1" class="a" >
<tr><td align="right"><a href="LecCoupon_edit.asp">[신규등록]</a></td></tr>
</table>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor=#B2B2B2 class=a>
<tr bgcolor="#E6E6E6">
	<td width="50" align="center">IDx</td>
	<td align="center">아이디</td>
	<td align="center">보너스내용</td>
	<td width="50" align="center">사용가능상품</td>
	<td width="150" align="center">사용 혜택</td>
	<td width="50" align="center">최소구매 금액</td>
	<td width="150" align="center">유효기간</td>
	<td width="80" align="center">등록일</td>
	<td width="30" align="center">사용 여부</td>
	<td width="100" align="center">발급자</td>
</tr>
<% for i=0 to ocoupon.FResultCount - 1 %>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= ocoupon.FItemList(i).FIdx %></td>
	<td align="center">
		<%= printUserId(ocoupon.FItemList(i).Fuserid, 2, "*") %>
	</td>
	<td><%= ocoupon.FItemList(i).Fcouponname %></td>
	<td align="center"><%= ocoupon.FItemList(i).Ftargetitemlist %></td>
	<td align="center"><%= ocoupon.FItemList(i).getCouponTypeStr %></td>
	<td align="center"><%= ocoupon.FItemList(i).Fminbuyprice %></td>
	<td align="center"><%= ocoupon.FItemList(i).getAvailDateStr %></td>
	<td align="center"><%= Formatdatetime(ocoupon.FItemList(i).FRegDate,2) %></td>
	<td align="center"><%= ocoupon.FItemList(i).FIsUsing %></td>
	<td align="center"><%= ocoupon.FItemList(i).Freguserid %></td>
</tr>
<% next %>
</table>
<%
set ocoupon = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->