<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  할인권 사용 통계 V2
' History : 2019.06.05 허진원 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/BonusCouponSummaryClass.asp"-->
<!-- #include virtual="/lib/classes/items/newcouponcls.asp" -->
<%
dim issuedcount, usingcount, spendcoupon, subtotalprice   , spendmileage, i

dim page, couponidx, userlevel, chkTerm, viewType, isIncULv, sDt, eDt
	page        = RequestCheckVar(request("page"),8)
	chkTerm    = RequestCheckVar(request("chkTerm"),1)
	sDt			= RequestCheckVar(request("sDt"),10)
	eDt			= RequestCheckVar(request("eDt"),10)
	couponidx   = RequestCheckVar(request("couponidx"),8)
	userlevel   = RequestCheckVar(request("userlevel"),1)
	viewType	= RequestCheckVar(request("viewType"),1)
	isIncULv	= RequestCheckVar(request("isIncULv"),1)

if (page="") then page=1
if viewType="" then viewType="D"	'D:일별, H:시간별

'// 보너스쿠폰 기본 정보 접수
Dim ocoupon, couponname, couponStartDt, couponExpireDt
set ocoupon = new CCouponMaster
ocoupon.FRectIdx = couponidx
ocoupon.GetOneCouponMaster
	couponname = ocoupon.FOneItem.Fcouponname
	couponStartDt = ocoupon.FOneItem.Fstartdate
	couponExpireDt = ocoupon.FOneItem.Fexpiredate
set ocoupon = Nothing

if sDt="" then sDt=left(couponStartDt,10)
if eDt="" then eDt=left(couponExpireDt,10)

'// 통계 내용 접수
dim oCouponSummary
set oCouponSummary = new CBonusCouponSummary
oCouponSummary.FPageSize = 100
oCouponSummary.FCurrpage = page
if (chkTerm<>"") then
    oCouponSummary.FRectStartDate = sDt
	oCouponSummary.FRectEndDate = eDt
end if
oCouponSummary.FRectCouponidx  = couponidx
oCouponSummary.FRectUserLevel  = userlevel
oCouponSummary.FRectViewType   = viewType
oCouponSummary.FRectIncULv     = isIncULv
oCouponSummary.getCouponResultSummaryHour
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css">
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript' src="/js/jsCal/js/jscal2.js"></script>
<script type='text/javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style type="text/css">
html {overflow-y:auto;}
#searchFilter span {white-space:nowrap;}
.dimmed {text-align: center; padding-top: 200px;}
</style>
<script type="text/javascript">
function goPage(ipage){
    frm.page.value=ipage;
    frm.submit();
}

$(function(){
	$("#chkTerm").click(function(){
		if($(this).prop("checked")) {
			$("#sDt,#eDt").attr("disabled",false);
		} else {
			$("#sDt,#eDt").attr("disabled",true);
		}
	});

	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
});
</script>
<div class="pad20" style="background-color:#FFF;">
	<h3 class="bMar05"><%=couponname & " (" & left(couponStartDt,10) & "~" & left(couponExpireDt,10) & ")" %></h3>
	<!-- 검색 필터 -->
	<form name="frm" method="get" action="" style="margin:0px;">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
		<table id="searchFilter" class="tbType1 listTb tMar10">
		<colgroup>
			<col width="80" />
			<col width="*" />
			<col width="80" />
		</colgroup>
		<tr>
			<th rowspan="2">검색<br>조건</th>
			<td class="lt">
				<span>
					<input type="checkbox" name="chkTerm" id="chkTerm" <%=ChkIIF(chkTerm<>"","checked","")%> />
					조회기간 : 
					<input id="sDt" name="sDt" value="<%=sDt%>" size="10" maxlength="10" style="cursor:pointer" <%=ChkIIF(chkTerm<>"","","disabled")%> autocomplete="off" /> ~
					<input id="eDt" name="eDt" value="<%=eDt%>" size="10" maxlength="10" style="cursor:pointer" <%=ChkIIF(chkTerm<>"","","disabled")%> autocomplete="off" />
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField:"sDt", trigger:"sDt", max:"<%=edt%>",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_End = new Calendar({
							inputField:"eDt", trigger:"eDt",  min:"<%=sdt%>",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</span>
			</td>	
			<th rowspan="2">
				<input type="button" value="검색" onClick="frm.submit()" class="ui-button" style="font-size:11px;">
			</th>
		</tr>
		<tr>
			<td class="lt">
				<span>쿠폰 번호 : <input type="text" name="couponidx" value="<%= couponidx %>" size="5" maxlength="9"></span> &nbsp;
				<span>사용자레벨 : <% DrawselectboxUserLevel "userlevel",  userlevel, "" %></span> &nbsp;
				<span class="rdoUsing">출력기준 :
					<input type="radio" name="viewType" id="viewTypeD" value="D" <%=chkIIF(viewType="D","checked","")%>/><label for="viewTypeD">일별</label>
					<input type="radio" name="viewType" id="viewTypeH" value="H" <%=chkIIF(viewType="H","checked","")%>/><label for="viewTypeH">시간별</label>
				</span> &nbsp;
				<span><label><input type="checkbox" name="isIncULv" value="Y" <%=chkIIF(isIncULv="Y","checked","")%> /> 등급포함</label></span>
			</td>
		</tr>
		</table>
	</form>
<br />

<!-- 리스트 시작 -->
<%
	dim colCnt: colCnt = chkIIF(isIncULv="Y",8,7)
%>
<table align="center" cellpadding="3" cellspacing="1" class="tbType1 listTb tMar10">
<tr height="25">
	<td colspan="<%=colCnt%>" class="lt">
		검색결과 : <b><%= oCouponSummary.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oCouponSummary.FTotalPage %></b>
	</td>
</tr>
<tr align="center">
	<th><%=chkIIF(viewType="H","발행일시","발행일")%></th>
  	<% if isIncULv="Y" then %><th>레벨</th><% end if %>
  	<th>발행건</th>
  	<th>사용건</th>
  	<th>사용율</th>
  	<th>사용액</th>
  	<th>사용건매출<br>(마일리지제외)</th>
  	<th>마일리지매출</th>
</tr>

<% if oCouponSummary.FresultCount>0 then %>
	<%
	for i=0 to oCouponSummary.FResultCount -1

	issuedcount     = issuedcount + oCouponSummary.FItemList(i).Fissuedcount
	usingcount      = usingcount + oCouponSummary.FItemList(i).Fusingcount
	spendcoupon     = spendcoupon + oCouponSummary.FItemList(i).Fspendcoupon
	subtotalprice   = subtotalprice + oCouponSummary.FItemList(i).Fsubtotalprice
	spendmileage    = spendmileage + oCouponSummary.FItemList(i).Fspendmileage
	%>
	<tr>
		<td align="center"><%= oCouponSummary.FItemList(i).FbaseDate %></td>
		<% if isIncULv="Y" then %><td align="center"><font color="<%= getUserLevelColor(oCouponSummary.FItemList(i).Fuserlevel) %>"><%= getUserLevelStr(oCouponSummary.FItemList(i).Fuserlevel) %></font></td><% end if %>
		<td align="center"><%= FormatNumber(oCouponSummary.FItemList(i).Fissuedcount,0) %></td>
		<td align="center"><%= FormatNumber(oCouponSummary.FItemList(i).Fusingcount,0) %></td>
		<td align="center"><%= oCouponSummary.FItemList(i).getUsingPro() %>%</td>
		<td align="right"><%= FormatNumber(oCouponSummary.FItemList(i).Fspendcoupon,0) %></td>
		<td align="right"><%= FormatNumber(oCouponSummary.FItemList(i).Fsubtotalprice,0) %></td>
		<td align="right"><%= FormatNumber(oCouponSummary.FItemList(i).Fspendmileage,0) %></td>
	</tr>
	<% next %>

	<tr>
	    <td align="center">합계</td>
	    <% if isIncULv="Y" then %><td align="center"></td><% end if %>
	    <td align="center"><%= FormatNumber(issuedcount,0) %></td>
	    <td align="center"><%= FormatNumber(usingcount,0) %></td>
	    <td align="center">
		    <% if issuedcount<>0 then %>
		        <%= CLng(usingcount/issuedcount*100*100)/100 %>%
		    <% end if %>
	    </td>
	    <td align="right"><%= FormatNumber(spendcoupon,0) %></td>
	    <td align="right"><%= FormatNumber(subtotalprice,0) %></td>
	    <td align="right"><%= FormatNumber(spendmileage,0) %></td>
	</tr>
    <tr height="25">
		<td colspan="<%=colCnt%>" align="center">
			<%
			if oCouponSummary.HasPreScroll then
				Response.Write "<a href='javascript:goPage(" & oCouponSummary.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
			else
				Response.Write "[pre] &nbsp;"
			end if
	
			for i=0 + oCouponSummary.StartScrollPage to oCouponSummary.FScrollCount + oCouponSummary.StartScrollPage - 1
	
				if i>oCouponSummary.FTotalpage then Exit for
	
				if CStr(page)=CStr(i) then
					Response.Write " <font color='red'>[" & i & "]</font> "
				else
					Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
				end if
	
			next
	
			if oCouponSummary.HasNextScroll then
				Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
			else
				Response.Write "&nbsp; [next]"
			end if
			%>
		</td>
	</tr>
<% else %>
	<tr>
	    <td colspan="<%=colCnt%>">검색 결과가 없습니다.</td>
	</tr>
<% end if %>

</table>
<%
set oCouponSummary = Nothing
%>
</div>
<div class="dimmed"><img src="/images/loading.gif" width="150px" /></div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
