<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 회원 구매 히스토리
' Hieditor : 2010.06.04 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #Include virtual = "/lib/classes/totalpoint/totalpointCls.asp" -->

<%
Dim ohistory, i, page , vCardNo ,vUserName ,vUserID , parameter, fromDate, toDate, datefg, orderno, userhp
dim posuid , pssnkey , dummikey ,shopid ,disp, inc3pl, yyyy1,mm1,dd1,yyyy2,mm2,dd2, yyyymmdd1,yyymmdd2 ,oldlist
dim beasongyn
	vCardNo			= requestCheckVar(Request("cardno"),20)
	vUserName		= requestCheckVar(Request("username"),20)
	vUserID			= requestCheckVar(Request("userid"),32)
	posuid			= requestCheckVar(Request("posuid"),32)
	pssnkey			= Request("pssnkey")
	dummikey		= Request("dummikey")
	shopid = requestCheckVar(request("shopid"),32)
	page = requestCheckVar(request("page"),10)
    inc3pl = requestCheckVar(request("inc3pl"),1)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	oldlist = requestCheckVar(request("oldlist"),2)
	datefg = requestCheckVar(request("datefg"),16)
	orderno = requestCheckVar(request("orderno"),16)
	userhp = requestCheckVar(request("userhp"),16)
	beasongyn = requestCheckVar(request("beasongyn"),1)

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

if datefg = "" then datefg = "maechul"
if page = "" then page = 1	

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

set ohistory = new TotalPoint
	ohistory.FPageSize = 20
	ohistory.FCurrPage = page
 	ohistory.FUserName = vUserName
 	ohistory.FUserID = vUserID	
 	ohistory.FCardNo = vCardNo
 	ohistory.frectshopid = shopid	
	ohistory.FRectInc3pl = inc3pl
	ohistory.FRectStartDay = fromDate
	ohistory.FRectEndDay = toDate
	ohistory.frectdatefg = datefg
	ohistory.FRectOldData = oldlist
	ohistory.FRectorderno = orderno
	ohistory.FRectuserhp = userhp
	ohistory.FRectbeasongyn = beasongyn
	ohistory.fsell_history_master()	
%>

<script type="text/javascript">

//포인트 적립 상세 보기.
function goRead(userseq){    
	var popwin = window.open('/admin/totalpoint/customer_sell_history_point.asp?userseq='+userseq+'&menupos=<%= menupos %>','addregpoint','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function getOnload(){
    frm.cardno.select();
    frm.cardno.focus();
}

window.onload = getOnload;

//주문내역 상세 보기. 이문재 이사님요청 팝업으로 변경
function gobeasong(orderno){
	var popwin = window.open('/common/offshop/beasong/shopbeasong_input.asp?orderno='+orderno+'&menupos=<%= menupos %>','addregbeasong','width=1280,height=960,scrollbars=yes,resizable=yes');
	//var popwin = window.open('/admin/totalpoint/customer_sell_history_detail.asp?orderno='+orderno+'&menupos=<%= menupos %>','addregdetail','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//폼 전송
function gosubmit(page){
    frm.page.value=page;
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="posuid" value="<%=posuid%>">
<input type="hidden" name="pssnkey" value="<%=pssnkey%>">
<input type="hidden" name="dummikey" value="<%=dummikey%>">
<input type="hidden" name="userseq">
<input type="hidden" name="page">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 기간 : <% drawmaechuldatefg "datefg" ,datefg ,""%>
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >5년이전
		&nbsp;&nbsp;
		<% if C_ADMIN_USER then %>
			* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
		<%
		'직영/가맹점
		elseif (C_IS_SHOP) then
		%>	
			<% if not C_IS_OWN_SHOP and shopid <> "" then %>
				* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<% else %>
			* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
		<% end if %>
		&nbsp;&nbsp;
		* 주문번호 : <input type="text" class="text" name="orderno" value="<%= orderno %>" onKeyPress="if(window.event.keyCode==13) frm.submit();">
	</td>	
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="gosubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">	
		* 카드번호 : <input type="text" class="text" name="cardno" value="<%=vCardNo%>" onKeyPress="if(window.event.keyCode==13) frm.submit();">
		&nbsp;&nbsp;
		* 고객명 : <input type="text" class="text" name="username" value="<%=vUserName%>" size="8" onKeyPress="if(window.event.keyCode==13) frm.submit();">
		&nbsp;&nbsp;
		* 온라인아이디 : <input type="text" class="text" name="userid" value="<%=vUserID%>" size="12" onKeyPress="if(window.event.keyCode==13) frm.submit();">
        &nbsp;&nbsp;
        <b>* 매출처구분</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">	
		* 배송여부 : <% drawSelectBoxisusingYN "beasongyn",beasongyn,"" %>
        &nbsp;&nbsp;
		* 휴대폰번호(주문시입력) : <input type="text" name="userhp" value="<%= userhp %>" size="16" onKeyPress="if(window.event.keyCode==13) frm.submit();" class="text">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ohistory.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= ohistory.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=110>주문번호</td>
	<td>매장명</td>
	<td width=100>매장ID</td>
	<td>성함</td>
	<td>온라인ID</td>
	<td width=110>카드번호</td>
	<td width=80>판매금액</td>
	<td width=80>실결제액</td>
	<td width=140>판매일</td>출고상태
	<td width=50>배송여부</td>
	<td width=90>출고상태</td>
	<td width=90>휴대폰번호<br>(주문시입력)</td>
	<td>비고</td>
</tr>
<%
if ohistory.FresultCount > 0 then

for i=0 to ohistory.FresultCount-1 
%>
<% if ohistory.FItemList(i).fcancelyn = "N" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFaa';>
<% end if %>
	<td>
		<%= ohistory.FItemList(i).forderno %>
		
		<% if ohistory.FItemList(i).fcancelyn = "Y" then %>
			<br>(취소)
		<% end if %>
	</td>
	<td><%= ohistory.FItemList(i).fshopname %></td>
	<td><%= ohistory.FItemList(i).fshopid %></td>
	<td><%= ohistory.FItemList(i).fUserName %></td>
	<td><%= ohistory.FItemList(i).fOnLineUSerID %></td>
	<td><%= ohistory.FItemList(i).fpointuserno %></td>	
	<td><%= FormatNumber(ohistory.FItemList(i).ftotalsum,0) %></td>
	<td><%= FormatNumber(ohistory.FItemList(i).frealsum,0) %></td>
	<td><%= ohistory.FItemList(i).fregdate %></td>
	<td>
		<% if ohistory.FItemList(i).fipkumdiv <> "" and not isnull(ohistory.FItemList(i).fipkumdiv) then %>
			Y
		<% else %>
			N
		<% end if %>
	</td>
	<td><%= ohistory.FItemList(i).shopIpkumDivName %></td>
	<td><%= ohistory.FItemList(i).fuserhp %></td>
	<td>
		<input type="button" value="주문상세" onclick="gobeasong('<%= ohistory.FItemList(i).forderno %>');" class="button">
		<!--<input type="button" value="주문상세" onclick="gobeasong('<%= ohistory.FItemList(i).forderno %>');" class="button">-->

		<% if ohistory.FItemList(i).fUserSeq <> "" then %>
			<input type="button" class="button" value="포인트적립상세" onClick="goRead('<%= ohistory.FItemList(i).fUserSeq %>')">
		<% end if %>
	</td>
</tr>  
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if ohistory.HasPreScroll then %>			
			<span class="list_link"><a href="javascript:gosubmit('<%= ohistory.StartScrollPage-1 %>');">[pre]</a></span>
			
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ohistory.StartScrollPage to ohistory.StartScrollPage + ohistory.FScrollCount - 1 %>
			<% if (i > ohistory.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ohistory.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ohistory.HasNextScroll then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set ohistory = nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->