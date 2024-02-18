<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터 무통장입금관리
' History : 이상구생성
'			2019.11.20 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/academy/academy_jumuncls.asp"-->
<!-- #include virtual="/lib/classes/payment/ipkumlistcls.asp"-->
<%

dim orderserial
dim searchtype01, searchtype02
dim searchname, searchprice

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2,fromDate,toDate
dim nowdate,searchnextdate,research, bankOnly

nowdate = Left(CStr(now()),10)


searchtype01 = request("searchtype01")
searchtype02 = request("searchtype02")
searchname = request("searchname")
searchprice = request("searchprice")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
bankOnly = request("bankOnly")

dim yyyymmdd1
if (yyyy1="") then
    yyyymmdd1 = dateAdd("m",-1,now())
    yyyy1 = Cstr(Year(yyyymmdd1))
    mm1 = Cstr(Month(yyyymmdd1))
    dd1 = Cstr(day(yyyymmdd1))
end if


if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

searchnextdate = Left(CStr(DateAdd("m",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),-1)),10)

dim ckdate,ckdelsearch,ckipkumdiv4,ckipkumdiv2
ckdate = request("ckdate")
ckdelsearch = request("ckdelsearch")
ckipkumdiv4 = request("ckipkumdiv4")
orderserial = request("orderserial")
ckipkumdiv2 = request("ckipkumdiv2")
research = request("research")

if research="" then ckipkumdiv2="on"
if research="" then ckdate="on"
if research="" then bankOnly="Y"

dim page

page = request("page")
if (page="") then page=1

dim ojumun

'// ============================================================================
dim ipkum,idx, ipkumgubun
idx = request("idx")
ipkumgubun = request("ipkumgubun")

set ipkum = new IpkumChecklist

ipkum.GetipkumlistByIdx

if research="" then
	searchname = ipkum.Fipkumoneitem.Fipkumuser
	searchprice = ipkum.Fipkumoneitem.Fipkumsum

	searchtype01 = "on"
	searchtype02 = "on"
end if

'// ============================================================================
set ojumun = new CJumunMaster

ojumun.FPageSize = 30
ojumun.FRectckdate = ckdate
ojumun.FRectIpkumDiv4 = ckipkumdiv4
ojumun.FRectIpkumDiv2 = ckipkumdiv2
ojumun.FRectOrderSerial = orderserial
ojumun.FRectSearchtype01 = searchtype01
ojumun.FRectSearchtype02 = searchtype02
ojumun.FRectIpkumName = html2db(searchname)
ojumun.FRectSubTotalPrice = searchprice
ojumun.FRectBankOnly = bankOnly

ojumun.FRectRegStart = fromDate
ojumun.FRectRegEnd = toDate

ojumun.FCurrPage = page
if (ipkumgubun = "10x10") then
	ojumun.SearchMatchJumunList
end if


'// ============================================================================
dim oacalecjumun
set oacalecjumun = new CAcademyLecOrderMaster

oacalecjumun.FPageSize = 30
oacalecjumun.FRectckdate = ckdate
oacalecjumun.FRectIpkumDiv4 = ckipkumdiv4
oacalecjumun.FRectIpkumDiv2 = ckipkumdiv2
oacalecjumun.FRectOrderSerial = orderserial
oacalecjumun.FRectSearchtype01 = searchtype01
oacalecjumun.FRectSearchtype02 = searchtype02
oacalecjumun.FRectIpkumName = html2db(searchname)
oacalecjumun.FRectSubTotalPrice = searchprice

oacalecjumun.FRectRegStart = fromDate
oacalecjumun.FRectRegEnd = toDate

oacalecjumun.FCurrPage = page
if (ipkumgubun = "fingers") then
	oacalecjumun.QuickSearchOrderList
end if


dim ix,iy

%>

<script language="javascript">
function jsPopMatchInput(bankdate) {
	var v = "pop_matchorderinput.asp?idx=<%= request("idx") %>&bankdate=" + bankdate;
	var popwin = window.open(v,"jsPopMatchInput","width=350,height=180,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function jsMatchIpkum(frm) {
	if (confirm("매칭하시겠습니까?") == true) {
		frm.submit();
	}
}

</script>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td width="50">Idx</td>
    	<td width="50">은행</td>
    	<td width="65">입금일</td>
    	<td width="70">구분</td>
    	<td width="100">입금자</td>
    	<td width="50">입금액</td>
    	<td width="80">적요</td>
    	<td width="70">확인유무</td>
    	<td width="150">업로드일시</td>
		<td>비고</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" height="30">
    	<td><%= ipkum.Fipkumoneitem.Fidx %></td>
    	<td><%=	ipkum.Fipkumoneitem.Ftenbank %></td>
    	<td><%=	ipkum.Fipkumoneitem.FBankdate %></td>
    	<td><%=	ipkum.Fipkumoneitem.Fgubun %></td>
    	<td><%=	ipkum.Fipkumoneitem.Fipkumuser %></td>
    	<td align="right"><%= FormatNumber(ipkum.Fipkumoneitem.Fipkumsum,0) %></td>
    	<td><%= ipkum.Fipkumoneitem.Fbankname %></td>
    	<td>
    		<% if ipkum.Fipkumoneitem.Fipkumstate=0 then %>미처리
    		<% elseif ipkum.Fipkumoneitem.Fipkumstate=1 then %><font color="red">매칭실패</font>
    		<% elseif ipkum.Fipkumoneitem.Fipkumstate=7 then %><font color="blue">매칭완료</font>
    		<% end if %>
    	</td>
    	<td><%=	ipkum.Fipkumoneitem.Fregdate %></td>
		<td>
			<% if (ipkum.Fipkumoneitem.Fipkumstate <> 7) or (Left(ipkum.Fipkumoneitem.FBankdate, 10) = Left(Now(), 10)) or (DateDiff("d", left(ipkum.Fipkumoneitem.FBankdate,10), Left(Now(), 10)) <= 30) then %>
			<input type="button" class="button" value="수기매칭" onClick="jsPopMatchInput('<%=	ipkum.Fipkumoneitem.FBankdate %>')" <% if not(C_CSUser or C_ADMIN_AUTH) then %>disabled<% end if %> >
			<% end if %>
		</td>
    </tr>
</table>

<p>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="ipkumgubun" value="<%= ipkumgubun %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	주문번호 :
			<input type="text" name="orderserial" value="<%= orderserial %>" size="11" maxlength="16">
			&nbsp;

			<br>
			<input type="checkbox" name="searchtype01" <% if searchtype01<>"" then response.write "checked" %> >구매자/수령인/입금자
			<input type="text" name="searchname" value="<%= searchname %>" size="11" maxlength="16">
			<input type="checkbox" name="searchtype02" <% if searchtype02<>"" then response.write "checked" %> >구매금액
			<input type="text" name="searchprice" value="<%= searchprice %>" size="11" maxlength="16">
			<input type=checkbox name="ckdate" <% if ckdate="on" then  response.write "checked" %> onclick="EnDisabledDateBox(this)">검색기간
			<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;
			<input type="checkbox" name="ckipkumdiv2" <% if ckipkumdiv2="on" then response.write "checked" %> >주문접수만
			&nbsp;
			<input type="checkbox" name="bankOnly" value="Y" <% if bankOnly="Y" then response.write "checked" %> >무통장 주문만
        </td>
        <td align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
			<% if (ipkumgubun = "10x10") then %>
			총 건수 : <Font color="#3333FF"><%= FormatNumber(ojumun.FTotalCount,0) %></font>
			<% end if %>

			<% if (ipkumgubun = "fingers") then %>
			총 건수 : <Font color="#3333FF"><%= FormatNumber(oacalecjumun.FTotalCount,0) %></font>
			<% end if %>

        </td>
        <td align="right">
			<% if (ipkumgubun = "10x10") then %>
			page : <%= ojumun.FCurrPage %>/<%=ojumun.FTotalPage %>
			<% end if %>

			<% if (ipkumgubun = "fingers") then %>
			page : <%= oacalecjumun.FCurrPage %>/<%=oacalecjumun.FTotalPage %>
			<% end if %>

    	</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
	<tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td colspan="2" align="left">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="45">
		<td align="center" width="100">주문번호</td>
		<td width="60">Site</td>
		<td width="70">UserID</td>
		<td width="60">구매자</td>
		<td width="65">수령인</td>
		<td width="65">입금자명</td>
		<td width="60" align="right">구매총액</td>
		<td width="60" align="right">무통장 <br>실 결제금액</td>
		<td width="74">결제방법</td>
		<td width="74">주문상태</td>
		<td width="40">삭제<br>여부</td>
		<td width="70">주문일</td>
		<td>비고</td>
	</tr>
<% if (ojumun.FresultCount > 0) then %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<form name="frmOnerder_<%= ojumun.FMasterItemList(ix).FOrderSerial %>" method="post" action="pop_matchorderlist_Process.asp">
	<input type="hidden" name="mode" value="matchWithOrder">
	<input type="hidden" name="ipkumidx" value="<%= idx %>">
	<input type="hidden" name="orderserial" value="<%= ojumun.FMasterItemList(ix).FOrderSerial %>">
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr align="center" bgcolor="#FFFFFF" height="30">
	<% else %>
	<tr align="center" bgcolor="<%= adminColor("gray") %>" height="30">
	<% end if %>
		<td align="center"><%= ojumun.FMasterItemList(ix).FOrderSerial %></td>
		<td align="center"><font color="<%= ojumun.FMasterItemList(ix).SiteNameColor %>"><%= ojumun.FMasterItemList(ix).FSitename %></font></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).UserIDName %></td>
		<td align="center">
			<% if ojumun.FMasterItemList(ix).FBuyName = searchname then %>
			<font color="blue"><%= ojumun.FMasterItemList(ix).FBuyName %></font>
			<% else %>
			<%= ojumun.FMasterItemList(ix).FBuyName %>
			<% end if %>
		</td>
		<td align="center">
			<% if ojumun.FMasterItemList(ix).FReqName = searchname then %>
			<font color="blue"><%= ojumun.FMasterItemList(ix).FBuyName %></font>
			<% else %>
			<%= ojumun.FMasterItemList(ix).FReqName %>
			<% end if %>
		</td>
		<td align="center">
			<% if ojumun.FMasterItemList(ix).Faccountname = searchname then %>
			<font color="blue"><%= ojumun.FMasterItemList(ix).Faccountname %></font>
			<% else %>
			<%= ojumun.FMasterItemList(ix).Faccountname %>
			<% end if %>
		</td>
		<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).FTotalSum,0) %></td>
		<td align="right">
			<% if (ojumun.FMasterItemList(ix).TotalMajorPaymentPrice = searchprice) then %>
			<font color="blue"><%= FormatNumber(ojumun.FMasterItemList(ix).TotalMajorPaymentPrice,0) %></font>
			<% else %>
			<%= FormatNumber(ojumun.FMasterItemList(ix).FSubTotalPrice,0) %>
			<% end if %>
		</td>
		<td align="center"><%= ojumun.FMasterItemList(ix).JumunMethodName %></td>
		<td align="center"><font color="<%= ojumun.FMasterItemList(ix).IpkumDivColor %>"><%= ojumun.FMasterItemList(ix).IpkumDivName %></font></td>
		<td align="center"><font color="<%= ojumun.FMasterItemList(ix).CancelYnColor %>"><%= ojumun.FMasterItemList(ix).CancelYnName %></font></td>
		<td align="center"><acronym title="<%= Left(ojumun.FMasterItemList(ix).GetRegDate,16) %>"><%= Left(ojumun.FMasterItemList(ix).GetRegDate,10) %></acronym></td>
		<td>
			<% if ojumun.FMasterItemList(ix).IsAvailJumun and ipkum.Fipkumoneitem.Fipkumsum = ojumun.FMasterItemList(ix).FSubTotalPrice and (ipkum.Fipkumoneitem.Fipkumstate <> 7) and (ojumun.FMasterItemList(ix).FIpkumDiv = "2") and (ojumun.FMasterItemList(ix).Faccountname = ipkum.Fipkumoneitem.Fipkumuser) then %>
				<input type="button" class="button" value="매칭(결제완료전환)" onClick="jsMatchIpkum(frmOnerder_<%= ojumun.FMasterItemList(ix).FOrderSerial %>)" <% if not(C_CSUser or C_ADMIN_AUTH) then %>disabled<% end if %> >
			<% else %>
				<%= ojumun.FMasterItemList(ix).Faccountno %>
			<% end if %>
		</td>
	</tr>
	</form>
	<% next %>
<% elseif (oacalecjumun.FresultCount > 0) then %>
	<% for ix=0 to oacalecjumun.FresultCount-1 %>
	<form name="frmOnerder_<%= oacalecjumun.FMasterItemList(ix).FOrderSerial %>" method="post" action="pop_matchorderlist_Process.asp">
	<input type="hidden" name="mode" value="matchWithOrder">
	<input type="hidden" name="ipkumidx" value="<%= idx %>">
	<input type="hidden" name="orderserial" value="<%= oacalecjumun.FMasterItemList(ix).FOrderSerial %>">
	<% if oacalecjumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr align="center" bgcolor="#FFFFFF" height="30">
	<% else %>
	<tr align="center" bgcolor="<%= adminColor("gray") %>" height="30">
	<% end if %>
		<td align="center"><%= oacalecjumun.FMasterItemList(ix).FOrderSerial %></td>
		<td align="center"><font color="<%= oacalecjumun.FMasterItemList(ix).SiteNameColor %>"><%= oacalecjumun.FMasterItemList(ix).FSitename %></font></td>
		<td align="center"><%= oacalecjumun.FMasterItemList(ix).Fuserid %></td>
		<td align="center">
			<% if oacalecjumun.FMasterItemList(ix).FBuyName = searchname then %>
			<font color="blue"><%= oacalecjumun.FMasterItemList(ix).FBuyName %></font>
			<% else %>
			<%= oacalecjumun.FMasterItemList(ix).FBuyName %>
			<% end if %>
		</td>
		<td align="center">
			<% if oacalecjumun.FMasterItemList(ix).FReqName = searchname then %>
			<font color="blue"><%= oacalecjumun.FMasterItemList(ix).FBuyName %></font>
			<% else %>
			<%= oacalecjumun.FMasterItemList(ix).FReqName %>
			<% end if %>
		</td>
		<td align="center">
			<% if oacalecjumun.FMasterItemList(ix).Faccountname = searchname then %>
			<font color="blue"><%= oacalecjumun.FMasterItemList(ix).Faccountname %></font>
			<% else %>
			<%= oacalecjumun.FMasterItemList(ix).Faccountname %>
			<% end if %>
		</td>
		<td align="right"><%= FormatNumber(oacalecjumun.FMasterItemList(ix).FTotalSum,0) %></td>
		<td align="right">
			<% if (oacalecjumun.FMasterItemList(ix).FSubTotalPrice = searchprice) then %>
			<font color="blue"><%= FormatNumber(oacalecjumun.FMasterItemList(ix).FSubTotalPrice,0) %></font>
			<% else %>
			<%= FormatNumber(oacalecjumun.FMasterItemList(ix).FSubTotalPrice,0) %>
			<% end if %>
		</td>
		<td align="center"><%= oacalecjumun.FMasterItemList(ix).JumunMethodName %></td>
		<td align="center"><font color="<%= oacalecjumun.FMasterItemList(ix).IpkumDivColor %>"><%= oacalecjumun.FMasterItemList(ix).IpkumDivName %></font></td>
		<td align="center"><font color="<%= oacalecjumun.FMasterItemList(ix).CancelYnColor %>"><%= oacalecjumun.FMasterItemList(ix).CancelYnName %></font></td>
		<td align="center"><acronym title="<%= Left(oacalecjumun.FMasterItemList(ix).GetRegDate,16) %>"><%= Left(oacalecjumun.FMasterItemList(ix).GetRegDate,10) %></acronym></td>
		<td>
			<% if oacalecjumun.FMasterItemList(ix).IsAvailJumun and ipkum.Fipkumoneitem.Fipkumsum = oacalecjumun.FMasterItemList(ix).FSubTotalPrice and (ipkum.Fipkumoneitem.Fipkumstate <> 7) and (oacalecjumun.FMasterItemList(ix).FIpkumDiv = "2") and (oacalecjumun.FMasterItemList(ix).Faccountname = ipkum.Fipkumoneitem.Fipkumuser) then %>
				<!--
				<input type="button" class="button" value="매칭(결제완료전환)" onClick="jsMatchIpkum(frmOnerder_<%= oacalecjumun.FMasterItemList(ix).FOrderSerial %>)" <% if not(C_CSUser or C_ADMIN_AUTH) then %>disabled<% end if %> >
				-->
			<% else %>
				<%= oacalecjumun.FMasterItemList(ix).Faccountno %>
			<% end if %>
		</td>
	</tr>
	</form>
	<% next %>
<% else %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="13" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
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
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->



<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
