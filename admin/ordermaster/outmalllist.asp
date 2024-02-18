<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 주문처리
' Hieditor : 2011.04.22 이상구 생성
'			 2012.08.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
dim orderserial
dim searchtype
dim searchrect
Dim ipkumdiv
dim nowdate,searchnextdate

dim outmall
Dim sDt, eDt, research
sDt			= request("sDt")
eDt			= request("eDt")
research	= request("research")
nowdate = Left(CStr(now()),10)

outmall = request("outmall")
searchtype = request("searchtype")
searchrect = request("searchrect")
ipkumdiv = request("ipkumdiv")

If (research = "") Then
	sDt = Date()
	eDt = DateAdd("d",1,sDt)
End If

dim cknodate,ckdelsearch,ckipkumdiv4
cknodate    = requestCheckVar(request("cknodate"),32)
ckdelsearch = requestCheckVar(request("ckdelsearch"),32)
ckipkumdiv4 = requestCheckVar(request("ckipkumdiv4"),32)
orderserial = requestCheckVar(request("orderserial"),32)


dim page
dim ojumun

page = request("page")
if (page="") then page=1

set ojumun = new CJumunMaster
	ojumun.FRectRegStart = sDt
	ojumun.FRectRegEnd = eDt
	ojumun.FRectDelNoSearch=ckdelsearch


if searchtype="01" then
	ojumun.FRectBuyname = searchrect
elseif searchtype="02" then
	ojumun.FRectReqName = searchrect
elseif searchtype="03" then
	ojumun.FRectUserID = searchrect
elseif searchtype="04" then
	ojumun.FRectIpkumName = searchrect
elseif searchtype="06" then
	ojumun.FRectSubTotalPrice = searchrect
elseif searchtype="07" then
	ojumun.FRectAuthCode = searchrect
end if

ojumun.FPageSize = 50
''ojumun.FRectIpkumDiv4 = "on"
ojumun.FRectOrderSerial = orderserial
ojumun.FRectOnlyOutMall = "on"
ojumun.FRectSiteName = outmall
ojumun.FRectIpkumdiv = ipkumdiv
ojumun.FCurrPage = page
ojumun.FRectQryDLVsum ="on" '2013/09/24
ojumun.SearchJumunList

dim ix,iy
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language='javascript'>
function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = 'orderdetail';
    frm.action="viewordermaster.asp"
	frm.submit();

}

function ViewUserInfo(frm){
	//var popwin;
    //popwin = window.open('','userinfo');
    frm.target = 'userinfo';
    frm.action="viewuserinfo.asp"
	frm.submit();

}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function popMisendInfo(ios){
    var popwin = window.open('http://webadmin.10x10.co.kr/admin/ordermaster/misendmaster_main.asp?orderserial='+ios,'popMisendInfo','width=1000,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr bgcolor="#F4F4F4" >
	    <td align="center" width="50" bgcolor="#EEEEEE">검색<br>조건</td>
		<td >
			* 주문번호 :
			<input type="text" name="orderserial" value="<%= orderserial %>" size="11" maxlength="16">
			&nbsp;&nbsp;
			* 검색기간 :
			<input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" />
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
			<input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" />
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<br>
			* 검색조건 :
			<select name="searchtype">
				<option value="">선택</option>
				<option value="01" <% if searchtype="01" then response.write "selected" %> >구매자</option>
				<option value="02" <% if searchtype="02" then response.write "selected" %> >수령인</option>
				<option value="03" <% if searchtype="03" then response.write "selected" %> >아이디</option>
				<option value="04" <% if searchtype="04" then response.write "selected" %> >입금자</option>
				<option value="06" <% if searchtype="06" then response.write "selected" %> >결제금액</option>
				<option value="07" <% if searchtype="07" then response.write "selected" %> >카드승인번호</option>
			</select>
			<input type="text" name="searchrect" value="<%= searchrect %>" size="20" maxlength="25">
			&nbsp;&nbsp;
			* 사이트 : <% drawSelectBoxXSiteOrderInputPartner "outmall",outmall %>
			<!--<select name="outmall">
				<option value="">선택</option>
				<option value="interpark" <% if outmall="interpark" then response.write "selected" %> >interpark</option>
		        <option value="lotteCom" <% if outmall="lotteCom" then response.write "selected" %> >lotteCom</option>
		        <option value="lotteimall" <% if outmall="lotteimall" then response.write "selected" %> >lotteimall</option>
				<option value="bandinlunis" <% if outmall="bandinlunis" then response.write "selected" %> >bandinlunis</option>
				<option value="hanatour" <% if outmall="hanatour" then response.write "selected" %> >hanatour</option>
				<option value="fashionplus" <% if outmall="fashionplus" then response.write "selected" %> >fashionplus</option>
				<option value="wizwid" <% if outmall="wizwid" then response.write "selected" %> >wizwid</option>
				<option value="wconcept" <% if outmall="wconcept" then response.write "selected" %> >wconcept</option>
				<option value="gseshop" <% if outmall="gseshop" then response.write "selected" %> >gseshop</option>
		        <option value="29cm" <% if outmall="29cm" then response.write "selected" %> >29cm</option>
		        <option value="hottracks" <% if outmall="hottracks" then response.write "selected" %> >hottracks</option>
		        <option value="hiphoper" <% if outmall="hottracks" then response.write "selected" %> >hiphoper</option>
		        <option value="gmarket" <% if outmall="gmarket" then response.write "selected" %> >gmarket</option>
		        <option value="cjmallITS" <% if outmall="cjmallITS" then response.write "selected" %> >cjmall아이띵소</option>
			</select>-->
			&nbsp;&nbsp;
			거래상태 :
			<select name="ipkumdiv">
				<option value="">선택</option>
				<option value="4" <% if ipkumdiv="4" then response.write "selected" %> >결제완료</option>
				<option value="5" <% if ipkumdiv="5" then response.write "selected" %> >배송통보</option>
				<option value="6" <% if ipkumdiv="6" then response.write "selected" %> >배송준비</option>
				<option value="8" <% if ipkumdiv="8" then response.write "selected" %> >출고완료</option>
			</select>

			<input type="checkbox" name="ckdelsearch" value="on" <%=CHKIIF(ckdelsearch="on","checked","") %> >취소제외
	    </td>
	    <td align="center" width="50" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<p>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr bgcolor="#FFFFFF">
	<td colspan="16">
		총 건수 : <Font color="#3333FF"><%= FormatNumber(ojumun.FTotalCount,0) %></font>
		&nbsp;총 금액 : <Font color="#3333FF"><%= FormatNumber(ojumun.FSubTotal,0) %></font>
		&nbsp;평균객단가 : <Font color="#3333FF"><%= FormatNumber(ojumun.FAvgTotal,0) %></font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="right">page : <%= ojumun.FCurrPage %>/<%=ojumun.FTotalPage %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" align="center">주문번호</td>
	<td width="100" align="center">제휴몰<br>주문번호</td>
	<td width="80" align="center">Site</td>
	<!--<td width="80" align="center">UserID</td>-->
	<td width="80" align="center">텐배송<br>송장번호</td>
	<td width="65" align="center">구매자</td>
	<td width="65" align="center">수령인</td>
	<td width="60" align="center">할인율</td>
	<td width="72" align="center">결제금액</td>
	<td width="72" align="center">구매총액</td>
	<td width="60" align="center">배송비</td>
	<td width="74" align="center">결제방법</td>
	<td width="74" align="center">거래상태</td>
	<td width="40" align="center">취소<br>삭제</td>
	<td width="120" align="center">주문일</td>
	<% If outmall = "cnglob10x10" Then %>
	<td width="150" align="center">배송일</td>
	<% End If %>
</tr>
<% if ojumun.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<form name="frmOnerder_<%= ojumun.FMasterItemList(ix).FOrderSerial %>" method="post" >
	<input type="hidden" name="orderserial" value="<%= ojumun.FMasterItemList(ix).FOrderSerial %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="sitename" value="<%= ojumun.FMasterItemList(ix).FSiteName %>">
	<input type="hidden" name="userid" value="<%= ojumun.FMasterItemList(ix).UserIDName %>">
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr class="a" bgcolor="#FFFFFF">
	<% else %>
	<tr class="gray" bgcolor="#FFFFFF">
	<% end if %>
		<td align="center"><a href="#" onclick="ViewOrderDetail(frmOnerder_<%= ojumun.FMasterItemList(ix).FOrderSerial %>)" class="zzz"><%= ojumun.FMasterItemList(ix).FOrderSerial %></a></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FAuthcode %></td>
		<td align="center"><font color="<%= ojumun.FMasterItemList(ix).SiteNameColor %>"><%= ojumun.FMasterItemList(ix).FSitename %></font></td>

		<% 'if ojumun.FMasterItemList(ix).UserIDName<>"&nbsp;" then %>
			<!--<td align="center"><a href="#" onclick="ViewUserInfo(frmOnerder_<%'= ojumun.FMasterItemList(ix).FOrderSerial %>)" class="zzz"><%'= ojumun.FMasterItemList(ix).UserIDName %></a></td>-->
		<% 'else %>
			<!--<td align="center"><%'= ojumun.FMasterItemList(ix).UserIDName %></td>-->
		<% 'end if %>

		<td align="center"><%= ojumun.FMasterItemList(ix).Fdeliverno %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FBuyName %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FReqName %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FDisCountrate %></td>
		<td align="right"><font color="<%= ojumun.FMasterItemList(ix).SubTotalColor%>"><%= FormatNumber(ojumun.FMasterItemList(ix).FSubTotalPrice,0) %></font></td>
		<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).FTotalSum,0) %></td>
		<td align="right">
		<% if not IsNULL(ojumun.FMasterItemList(ix).FdlvPaySum) then %>
		<%= FormatNumber(ojumun.FMasterItemList(ix).FdlvPaySum,0) %>
		<% end if %>
		</td>
		<td align="center"><%= ojumun.FMasterItemList(ix).JumunMethodName %></td>
		<td align="center"><a href="javascript:popMisendInfo('<%= ojumun.FMasterItemList(ix).FOrderSerial %>');"><font color="<%= ojumun.FMasterItemList(ix).IpkumDivColor %>"><%= ojumun.FMasterItemList(ix).IpkumDivName %></font></a></td>
		<td align="center"><font color="<%= ojumun.FMasterItemList(ix).CancelYnColor %>"><%= ojumun.FMasterItemList(ix).CancelYnName %></font></td>
		<td align="center"><%= Left(ojumun.FMasterItemList(ix).GetRegDate,16) %></td>
		<% If outmall = "cnglob10x10" Then %>
		<td align="center"><%= Left(ojumun.FMasterItemList(ix).FCvbeadaldate,10) %></td>
		<% End If %>
	</tr>
	</form>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" height="30" align="center">
		<% if ojumun.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
			<% if ix>ojumun.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if ojumun.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
<% end if %>
</table>
<script language="javascript">
	var CAL_Start = new Calendar({
		inputField : "sDt", trigger    : "sDt_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_End.args.min = date;
			CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
	var CAL_End = new Calendar({
		inputField : "eDt", trigger    : "eDt_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_Start.args.max = date;
			CAL_Start.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
</script>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->