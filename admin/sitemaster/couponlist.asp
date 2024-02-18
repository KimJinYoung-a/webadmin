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
' History : 서동석 생성
'			2022.07.04 한용민 수정(isms취약점수정)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/newcouponcls.asp" -->
<%
dim ocoupon, page , i , validsitename, couponidx, couponname, targetCpnTp
	page = requestCheckvar(request("page"),10)
	validsitename = LEFT(request("validsitename"),20)
	couponidx = getNumeric(requestCheckvar(request("couponidx"),8))
	couponname = requestCheckvar(request("couponname"),20)
	targetCpnTp = requestCheckvar(request("targetCpnTp"),1)
	if page="" then page=1
	
	'//[Fingers]사이트관리>>보너스쿠폰프로모션 페이지에서는 핑거스 쿠폰만 보이도록 박아 넣는다
	if menupos = "1224" or menupos = "1216" then validsitename = "'academy','diyitem'"		
	
set ocoupon = new CCouponMaster
	ocoupon.FPageSize=50
	ocoupon.FCurrPage = page
	ocoupon.frectvalidsitename = validsitename
	ocoupon.FRectIdx = couponidx
	ocoupon.FrectCouponname = couponname
	ocoupon.FrectTargetCpnType = targetCpnTp
	ocoupon.GetCouponMasterList()
%>

<script type="text/javascript">
// 페이지 이동
function goPage(pg)
{
	document.frm_search.page.value=pg;
	document.frm_search.submit();
}
</script>

<!-- 검색 시작 -->
<form name="frm_search" method="GET" action="" onSubmit="return false">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
	<td width="50" rowspan="2" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<label>쿠폰코드 : <input type="text" name="couponidx" class="text" value="<%=couponidx%>" size="8"></label> /
		<label>쿠폰명 : <input type="text" name="couponname" class="text" value="<%=couponname%>" size="20"></label> /
		<label>쿠폰타입 :
			<select name="targetCpnTp" class="select">
			<option value="" <%=chkIIF(targetCpnTp="","selected","")%>>전체</option>
			<option value="C" <%=chkIIF(targetCpnTp="C","selected","")%>>카테고리</option>
			<option value="B" <%=chkIIF(targetCpnTp="B","selected","")%>>브랜드</option>
			</select>
		</label>
	</td>	
	<td width="50" rowspan="2" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="goPage(1);">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>
		<label><input type="radio" name="validsitename" value="" <%=chkIIF(validsitename="","checked","")%>> 전체</label>
		<label><input type="radio" name="validsitename" value="'app'" <%=chkIIF(validsitename="'app'","checked","")%>> 앱쿠폰</label>
		<label><input type="radio" name="validsitename" value="'academy','diyitem'" <%=chkIIF(validsitename="'academy','diyitem'","checked","")%>> 핑거스쿠폰</label>
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">			
	</td>
	<td align="right">			
		<a href="newcouponreg.asp?menupos=<%=menupos%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if ocoupon.FResultCount > 0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ocoupon.FResultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>ID</td>
	<td>쿠폰구분</td>
	<td>보너스내용</td>
	<td>사용 혜택</td>
	<td>최소<br>구매 금액</td>
	<td>최대<br>할인 금액</td>
	<td>유효기간</td>
	<td>등록일</td>
	<td>발급 마감일</td>
	<td>쿠폰타입</td>
	<td>비고</td>
</tr>
<% for i=0 to ocoupon.FResultCount - 1 %>
<% if ocoupon.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" >
<% else %>    
<tr align="center" bgcolor="silver" >
<% end if %>
	<td><%= ocoupon.FItemList(i).FIdx %></td>
	<td>
		<%
			if ocoupon.FItemList(i).fvalidsitename = "" then
				response.write "텐바이텐"
			elseif ocoupon.FItemList(i).fvalidsitename = "academy" then
				response.write "핑거스강좌쿠폰"
			elseif ocoupon.FItemList(i).fvalidsitename = "diyitem" then				
				response.write "핑거스상품쿠폰"
			elseif ocoupon.FItemList(i).fvalidsitename = "mobile" then				
				response.write "모바일"
			elseif ocoupon.FItemList(i).fvalidsitename = "app" then				
				response.write "APP"
			end if
			if ocoupon.FItemList(i).IsFreedeliverCoupon then
				response.write "<font color='red'><Br>무료배송</font>"
			end if
			if ocoupon.FItemList(i).IsWeekendCoupon then
				response.write "<Br>주말쿠폰"
			end if
		%>	
	</td>	
	<td>
		<a href="newcouponreg.asp?idx=<%= ocoupon.FItemList(i).FIdx %>&menupos=<%=menupos%>">
		<%= ReplaceBracket(ocoupon.FItemList(i).Fcouponname) %></a>
	</td>
	<td><%= ocoupon.FItemList(i).getCouponTypeStr %></td>
	<td align="right"><%= formatNumber(ocoupon.FItemList(i).Fminbuyprice,0) %></td>
	<td align="right">
	<% if (ocoupon.FItemList(i).FmxCpnDiscount<>0) then %>
		<%= formatNumber(ocoupon.FItemList(i).FmxCpnDiscount,0) %>
	<% end if%>
	</td>
	<td><%= ocoupon.FItemList(i).getAvailDateStr %></td>
	<td><%= Left(ocoupon.FItemList(i).FRegDate,10) %></td>
	<td><%= ocoupon.FItemList(i).FOpenFinishDate %></td>
	<td ><%= ocoupon.FItemList(i).getCouponTypeNameStr%></td>
	<% if (session("ssAdminPsn")=7) or (session("ssAdminPsn")=14) or (session("ssAdminPsn")=22) or (session("ssAdminPsn")=23) or (session("ssAdminPsn")=30) or (session("ssAdminPsn")=11) or (session("ssAdminPsn")=21) then '//마케팅,컨텐츠,전략기획,입점제휴,개발운영팀,개발팀,온라인MD수입,온라인MD운영 만 사용율보기 (2011.06.21; 허진원) %>
	<td align="center">
		<a href="/admin/datamart/mkt/bonuscouponsummary.asp?page=1&menupos=1021&couponidx=<%= ocoupon.FItemList(i).FIdx %>" target="_blank">사용율(Old)</a><br />
		<a href="/admin/datamart/mkt/bonuscouponsummaryV2.asp?menupos=1021&couponidx=<%= ocoupon.FItemList(i).FIdx %>" target="_blank">사용율(New)</a>
	</td>
	<% else %>
	<td>&nbsp;</td>
	<% end if %>
</tr>   
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="10" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td colspan="11" align="center">
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
</table>

<%
	set ocoupon = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->