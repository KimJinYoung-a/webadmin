<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<%
dim searchtype
dim searchrect

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate
dim orderserial

nowdate = Left(CStr(now()),10)

orderserial = request("orderserial")
searchtype = request("searchtype")
searchrect = request("searchrect")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

dim cknodate,ckdelsearch,ckipkumdiv4
cknodate = request("cknodate")
ckdelsearch = request("ckdelsearch")
ckipkumdiv4 = request("ckipkumdiv4")

dim page
dim ojumun

page = request("page")
if (page="") then page=1

set ojumun = new CJumunMaster

if cknodate="" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if

if ckdelsearch<>"on" then
	ojumun.FRectDelNoSearch="on"
end if


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
end if


if session("ssBctDiv")="999" then
	ojumun.FRectRdSite = session("ssBctID")
else
	ojumun.FRectSiteName = session("ssBctID")
end if

ojumun.FPageSize = 30
ojumun.FRectIpkumDiv4 = ckipkumdiv4
ojumun.FRectOrderSerial = orderserial
ojumun.FCurrPage = page
ojumun.SearchJumunList

dim ix,iy
%>
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

function ReSendMax(iorderserial){
	var popwin;
	popwin = window.open('http://maxshop.maxmovie.com/kspay/maxmovieposter.asp?orderserial=' + iorderserial,'orderdetail','width=400,height=400,scrollbars=yes');
}

function ReSendiKissyou(iorderserial){
	var popwin;
	popwin = window.open('http://designshop.ikissyou.com/ext/ikissyou/ikissyouposter.asp?orderserial=' + iorderserial,'orderdetail','width=400,height=400,scrollbars=yes');
}

</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		주문번호 :
		<input type="text" name="orderserial" value="<%= orderserial %>" size="11" maxlength="16">
		&nbsp;
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>

		<br>
		검색조건 :
		<select name="searchtype">
		<option value="">선택</option>
		<option value="01" <% if searchtype="01" then response.write "selected" %> >구매자</option>
		<option value="02" <% if searchtype="02" then response.write "selected" %> >수령인</option>
		<option value="03" <% if searchtype="03" then response.write "selected" %> >아이디</option>
		<option value="04" <% if searchtype="04" then response.write "selected" %> >입금자</option>
		<option value="06" <% if searchtype="06" then response.write "selected" %> >결제금액</option>
		</select>
		<input type="text" name="searchrect" value="<%= searchrect %>" size="11" maxlength="16">
		&nbsp;
		<input type="checkbox" name="ckipkumdiv4" <% if ckipkumdiv4="on" then response.write "checked" %> >결제완료이상검색
		&nbsp;
		<input type="checkbox" name="ckdelsearch" <% if ckdelsearch="on" then response.write "checked" %> >취소삭제검색
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="1" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="13">
		총 건수 : <Font color="#3333FF"><%= FormatNumber(ojumun.FTotalCount,0) %></font>
		&nbsp;총 금액 : <Font color="#3333FF"><%= FormatNumber(ojumun.FSubTotal,0) %></font>
		&nbsp;평균객단가 : <Font color="#3333FF"><%= FormatNumber(ojumun.FAvgTotal,0) %></font>
	</td>
</tr>
<tr>
	<td colspan="13" align="right">page : <%= ojumun.FCurrPage %>/<%=ojumun.FTotalPage %></td>
</tr>
<tr >
	<td width="100" align="center">주문번호</td>
	<td width="80" align="center">UserID</td>
	<td width="65" align="center">구매자</td>
	<td width="65" align="center">수령인</td>
	<td width="60" align="center">할인율</td>
	<td width="72" align="center">결제금액</td>
	<td width="72" align="center">구매총액</td>
	<td width="74" align="center">결제방법</td>
	<td width="74" align="center">거래상태</td>
	<td width="120" align="center">주문일</td>
	<td width="100" align="center">송장번호</td>
	<td width="80" align="center">취소삭제</td>
	<% if session("ssBctID")="ikissyou" then %>
	<td width="40" align="center">내역<br>전송</td>
	<% end if %>
</tr>
<% if ojumun.FresultCount<1 then %>
<tr>
	<td colspan="13" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<form name="frmOnerder_<%= ojumun.FMasterItemList(ix).FOrderSerial %>" method="post" >
	<input type="hidden" name="orderserial" value="<%= ojumun.FMasterItemList(ix).FOrderSerial %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="sitename" value="<%= ojumun.FMasterItemList(ix).FSiteName %>">
	<input type="hidden" name="userid" value="<%= ojumun.FMasterItemList(ix).UserIDName %>">
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr class="a">
	<% else %>
	<tr class="gray">
	<% end if %>
		<td align="center"><a href="#" onclick="ViewOrderDetail(frmOnerder_<%= ojumun.FMasterItemList(ix).FOrderSerial %>)" class="zzz"><%= ojumun.FMasterItemList(ix).FOrderSerial %></a></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).UserIDName %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FBuyName %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FReqName %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FDisCountrate %></td>
		<td align="right"><font color="<%= ojumun.FMasterItemList(ix).SubTotalColor%>"><%= FormatNumber(ojumun.FMasterItemList(ix).FSubTotalPrice,0) %></font></td>
		<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).FTotalSum,0) %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).JumunMethodName %></td>
		<td align="center"><font color="<%= ojumun.FMasterItemList(ix).IpkumDivColor %>"><%= ojumun.FMasterItemList(ix).IpkumDivName %></font></td>
		<td align="center"><%= Left(ojumun.FMasterItemList(ix).GetRegDate,16) %></td>
		<% if IsNull(ojumun.FMasterItemList(ix).Fdeliverno) or (ojumun.FMasterItemList(ix).Fdeliverno="") then %>
		<td align="center">&nbsp;</td>
		<% else %>
		<td align="center"><a href="http://www.cjgls.co.kr/contents/gls/gls004/gls004_06_01.asp?slipno=<%= ojumun.FMasterItemList(ix).Fdeliverno %>" target="_blank"><%= ojumun.FMasterItemList(ix).Fdeliverno %></a></td>
		<% end if %>
		<% if ojumun.FMasterItemList(ix).FCancelyn<>"N" then %>
			<% if ojumun.FMasterItemList(ix).FCancelyn="Y" then %>
			<td align="center"><font color="red">취소</font></td>
			<% elseif ojumun.FMasterItemList(ix).FCancelyn="D" then %>
			<td align="center"><font color="red">삭제</font></td>
			<% end if %>
		<% else %>
		<td align="center">&nbsp;</td>
		<% end if %>

		<% if session("ssBctID")="ikissyou" then %>
			<% if (ojumun.FMasterItemList(ix).FCancelyn="N") and (Clng(ojumun.FMasterItemList(ix).Fipkumdiv)>1) then %>
			<td align="center"><a href="javascript:ReSendiKissyou('<%= ojumun.FMasterItemList(ix).FOrderSerial %>');">전송</a></td>
			<% end if %>
		<% end if %>
	</tr>
	</form>
	<% next %>
	<tr>
		<td colspan="13" height="30" align="center">
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

<%
set ojumun = Nothing
%>
<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->