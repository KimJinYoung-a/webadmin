<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  상품 쿠폰
' History : 2010.09.30 한용민 생성
'####################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/diyshopitem/itemcouponcls.asp" -->

<%
dim oitemcoupon, page, research ,onlyvalid ,selDate, sSdate, sEdate ,iSerachType, sSearchTxt ,i
	research    = requestCheckVar(request("research"),9)
	page        = requestCheckVar(request("page"),9)
	iSerachType = requestCheckVar(request("iSerachType"),9)
	sSearchTxt  = requestCheckVar(request("sSearchTxt"),32)
	onlyvalid   = requestCheckVar(request("onlyvalid"),9)
	selDate     = requestCheckVar(request("selDate"),10)
	sSdate      = requestCheckVar(request("sSdate"),10)
	sEdate      = requestCheckVar(request("sEdate"),10)	
	if page="" then page=1
	if research="" then onlyvalid="on"

set oitemcoupon = new CItemCouponMaster
	oitemcoupon.FPageSize=30
	oitemcoupon.FCurrPage = page
	oitemcoupon.FRectOnlyValid = onlyvalid
	oitemcoupon.FRectSearchType = iSerachType
	oitemcoupon.FRectSearchTxt = sSearchTxt
	oitemcoupon.FRectSearchDate = selDate
	oitemcoupon.FRectStartDate = sSdate
	oitemcoupon.FRectEndDate   = sEdate
	oitemcoupon.GetItemCouponMasterList()
%>

<script language='javascript'>

function NextPage(page){
    var frm = document.frmSearch;
    frm.page.value = page;
    frm.submit();
}

function RegItemCoupon(){
	var popwin = window.open('itemcouponmasterreg.asp','RegItemCoupon','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function EditItemCoupon(itemcouponidx){
	var popwin = window.open('itemcouponmasterreg.asp?itemcouponidx=' + itemcouponidx,'EditItemCoupon','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function EditCouponItemList(itemcouponidx){
	var popwin = window.open('itemcouponitemlistedit.asp?itemcouponidx=' + itemcouponidx,'EditCouponItemList','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

</script>

<!---- 검색 ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSearch" method="get"  >
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	<select name="iSerachType">
	<option value="1" <%IF Cstr(iSerachType) = "1" THEN%>selected<%END IF%>>쿠폰코드</option>
	<option value="3" <%IF Cstr(iSerachType) = "3" THEN%>selected<%END IF%>>쿠폰명</option>
	<option value="2" <%IF Cstr(iSerachType) = "2" THEN%>selected<%END IF%>>이벤트코드</option>
	</select>
	<input type="text" name="sSearchTxt" value="<%=sSearchTxt%>" size="10" maxlength="10">		
	&nbsp;
	<select name="selDate">
	<option value="S" <%if Cstr(selDate) = "S" THEN %>selected<%END IF%>>시작일 기준</option>
	<option value="E" <%if Cstr(selDate) = "E" THEN %>selected<%END IF%>>종료일 기준</option>
	</select>		
	<input type="text" size="10" name="sSdate" value="<%=sSdate%>" onClick="jsPopCal('sSdate');" style="cursor:hand;">
	~ <input type="text" size="10" name="sEdate" value="<%=sEdate%>" onClick="jsPopCal('sEdate');"  style="cursor:hand;">
	
	<input type="checkbox" name="onlyvalid" <% if onlyvalid="on" then response.write "checked" %> >진행중인쿠폰 만 보기
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frmSearch.submit();">
	</td>
</tr>	
</table>
<!---- /검색 ---->

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">			
		<td align="right">
			<input type="button" class="button" value="신규 상품 쿠폰등록" onclick="RegItemCoupon();">
		</td>		
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oitemcoupon.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oitemcoupon.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oitemcoupon.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">쿠폰번호</td>
	<td align="center">쿠폰구분</td>
	<td align="center">이벤트코드<br>(그룹코드)</td>
	<td >쿠폰명</td>
	<td align="center">할인금액</td>
	<td align="center">시작일</td>
	<td align="center">종료일</td>
	<td align="center">상태</td>
	<td align="center">기본<br>마진구분</td>
	<td align="center">등록일</td>
	<td align="center">비고</td>	
</tr>
<% for i=0 to oitemcoupon.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
	<td><%= oitemcoupon.FItemList(i).Fitemcouponidx %></td>
	<td><font color="<%= oitemcoupon.FItemList(i).getCouponGubunColor %>"><%= oitemcoupon.FItemList(i).getCouponGubunName %></font></td>
	<td>
	    <%= oitemcoupon.FItemList(i).Fevt_code %>
	    <% if Not IsNULL(oitemcoupon.FItemList(i).Fevtgroup_code) then %>
	    (<%= oitemcoupon.FItemList(i).Fevtgroup_code %>)
	    <% end if %>
	</td>
	<td><%= oitemcoupon.FItemList(i).Fitemcouponname %></td>
	<td><%= oitemcoupon.FItemList(i).GetDiscountStr %></td>	
	<td><%= ChkIIF(Right(oitemcoupon.FItemList(i).Fitemcouponstartdate,8)="00:00:00",Left(oitemcoupon.FItemList(i).Fitemcouponstartdate,10),oitemcoupon.FItemList(i).Fitemcouponstartdate) %></td>
	<td><%= ChkIIF(Right(oitemcoupon.FItemList(i).Fitemcouponexpiredate,8)="23:59:59",Left(oitemcoupon.FItemList(i).Fitemcouponexpiredate,10),oitemcoupon.FItemList(i).Fitemcouponexpiredate) %></td>
	<td><font color="<%= oitemcoupon.FItemList(i).GetOpenStateColor %>"><%= oitemcoupon.FItemList(i).GetOpenStateName %></font></td>
	<td><%= oitemcoupon.FItemList(i).GetMargintypeName %></td>
	<td><%= Left(oitemcoupon.FItemList(i).FRegDate,10) %></td>
	<td>
		<input type="button" value="수정" onclick="EditItemCoupon('<%= oitemcoupon.FItemList(i).Fitemcouponidx %>')" class="button">
		<input type="button" value="상품(<%= oitemcoupon.FItemList(i).Fapplyitemcount %>)" onclick="EditCouponItemList('<%= oitemcoupon.FItemList(i).Fitemcouponidx %>');" class="button">
	</td>
</tr>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if oitemcoupon.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oitemcoupon.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oitemcoupon.StarScrollPage to oitemcoupon.FScrollCount + oitemcoupon.StarScrollPage - 1 %>
			<% if i>oitemcoupon.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oitemcoupon.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<%
	set oitemcoupon = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->