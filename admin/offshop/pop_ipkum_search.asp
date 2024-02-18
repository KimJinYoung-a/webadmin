<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 입금내역
' History : 서동석 생성
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/ipkumlistcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/etcmeachulcls.asp"-->
<%

dim jungsanidx
dim yyyy1, mm1, yyyy2, mm2
dim txammount, jeokyo
dim start_tx_day, end_tx_day
dim excluudematchfinish
dim acctno
dim research
dim inoutgubun, excustomer, ex10x10
dim showdismatch
dim orderby
dim serchjeokyoyn, serchtxammountyn, serchdateyn

jungsanidx 		= requestCheckVar(Request("jungsanidx"),10)
yyyy1 			= requestCheckVar(Request("yyyy1"),4)
mm1 			= requestCheckVar(Request("mm1"),2)
yyyy2 			= requestCheckVar(Request("yyyy2"),4)
mm2 			= requestCheckVar(Request("mm2"),2)

txammount 		= requestCheckVar(Trim(Request("txammount")),20)
jeokyo 			= requestCheckVar(Trim(Request("jeokyo")),100)

excluudematchfinish 	= requestCheckVar(Request("excluudematchfinish"),1)

acctno 			= requestCheckVar(Request("acctno"),20)
research 		= requestCheckVar(Request("research"),2)
inoutgubun 		= requestCheckVar(Request("inoutgubun"),1)
excustomer 		= requestCheckVar(Request("excustomer"),2)
ex10x10 		= requestCheckVar(Request("ex10x10"),20)
showdismatch	= requestCheckVar(Request("showdismatch"),1)
orderby			= requestCheckVar(Request("orderby"),1)

serchjeokyoyn		= requestCheckVar(Request("serchjeokyoyn"),1)
serchtxammountyn	= requestCheckVar(Request("serchtxammountyn"),1)
serchdateyn			= requestCheckVar(Request("serchdateyn"),1)

if (research = "") then
	'excluudematchfinish = "Y"
	inoutgubun = "2"
	excustomer = "Y"
	ex10x10 = "Y"
	showdismatch = ""
	orderby = "Y"
	serchtxammountyn = "Y"
	serchdateyn = "Y"
end if

if (yyyy1 = "") then
	yyyy1 = Year(now)
	mm1 = Month(now)

	yyyy2 = Year(now)
	mm2 = Month(now)
end if

start_tx_day = CStr(DateSerial(yyyy1, mm1, 1))
end_tx_day = CStr(DateSerial(yyyy2, (mm2 + 1), 1))


'// ===========================================================================
dim ofranchulgojungsan
dim jungsan_acctname

set ofranchulgojungsan = new CEtcMeachul
ofranchulgojungsan.FRectidx = jungsanidx

if (jungsanidx <> "") then
	ofranchulgojungsan.getOneEtcMeachul

	jungsan_acctname = ofranchulgojungsan.FOneItem.Fjungsan_acctname

	if (jeokyo = "") and (research = "") and (jungsan_acctname <> "") then
		jeokyo = jungsan_acctname
		serchjeokyoyn = "Y"
	end if
end if


'// ===========================================================================
dim matchexcludecnt
dim oipkum
set oipkum = new IpkumChecklist
	oipkum.FCurrpage=1
	oipkum.FPagesize=100
	oipkum.FScrollCount = 10

	if (serchtxammountyn = "Y") then
		oipkum.FRectTXAmmount = txammount
	end if

	if (serchjeokyoyn = "Y") then
		oipkum.FRectJeokyo = jeokyo
	end if

	if (serchdateyn = "Y") then
		oipkum.FRectTXDayStart = start_tx_day
		oipkum.FRectTXDayEnd = end_tx_day
	end if

	oipkum.FOrderby = orderby

	oipkum.FRectInOutGubun = inoutgubun
	oipkum.FRectExcluudeCustomer = excustomer
	oipkum.FRectExcluude10X10 = ex10x10
	oipkum.FRectExcluudeMatchFinish = excluudematchfinish
	oipkum.FRectAcctNo = acctno

	oipkum.GetipkumlistAccountsNew

	matchexcludecnt = 0
	for i=0 to oipkum.FResultCount-1
		if (oipkum.Fipkumitem(i).Fmatchstate = "X") then
			matchexcludecnt = matchexcludecnt + 1
		end if
	next


'// ===========================================================================
dim i

%>

<script language='javascript'>

function SubmitSearch(frm) {

	if (frm.serchjeokyoyn.checked == true) {
		if (frm.jeokyo.value == "") {
			alert("적요를 입력하세요");
			frm.jeokyo.focus();
			return;
		}
	}

	if (frm.serchtxammountyn.checked == true) {
		if (frm.txammount.value == "") {
			alert("입금액을 입력하세요");
			frm.txammount.focus();
			return;
		}

		if (frm.txammount.value*0 != 0) {
			alert("금액은 숫자만 가능합니다.");
			frm.txammount.focus();
			return;
		}
	}

	document.frm.submit();
}

function SubmitMatch(frm) {
	if (frm.matchprice.value == "") {
		alert("매칭금액을 입력하세요");
		frm.matchprice.focus();
		return;
	}

	if (frm.matchprice.value*0 != 0) {
		alert("매칭금액은 숫자만 가능합니다.");
		frm.matchprice.focus();
		return;
	}

	if (confirm("매칭하시겠습니까?") == true) {
		frm.submit();
	}
}

function SubmitDisMatch(frm) {
	if (confirm("매칭에서 제외하시겠습니까?") == true) {
		frm.mode.value = "dismatch";
		frm.submit();
	}
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="jungsanidx" value="<%= jungsanidx %>">

	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" height="60" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			구분 :
			<select class="select" name="inoutgubun">
				<option value="">전체</option>
				<option value="1" <%if (inoutgubun = "1") then %>selected<% end if %> >출금</option>
				<option value="2" <%if (inoutgubun = "2") then %>selected<% end if %> >입금</option>
			</select>
			&nbsp;
			은행 :
			<% Call drawSelectBoxBankList("acctno", acctno) %>
		</td>
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="SubmitSearch(frm)">
		</td>
	</tr>

	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type="checkbox" name="serchjeokyoyn" value="Y" <% if (serchjeokyoyn = "Y") then %>checked<% end if %> > 적요
			<input type="text" class="text" name="jeokyo" size="10" value="<%= jeokyo %>">
			&nbsp;
			<input type="checkbox" name="serchtxammountyn" value="Y" <% if (serchtxammountyn = "Y") then %>checked<% end if %> > 입금액
			<input type="text" class="text" name="txammount" size="10" value="<%= txammount %>">
			&nbsp;
			<input type="checkbox" name="serchdateyn" value="Y" <% if (serchdateyn = "Y") then %>checked<% end if %> > 검색기간 :
			<% Call DrawYMYMBox(yyyy1, mm1, yyyy2, mm2) %>
		</td>
	</tr>

	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type="checkbox" name="orderby" value="Y" <% if (orderby = "Y") then %>checked<% end if %> > 최근일순
			&nbsp;
			<input type="checkbox" name="excluudematchfinish" value="Y" <% if (excluudematchfinish = "Y") then %>checked<% end if %> > 매칭완료 제외
			&nbsp;
			<input type="checkbox" name="excustomer" value="Y" <% if excustomer<>"" then response.write "checked" %> > 고객입금 제외
			&nbsp;
			<input type="checkbox" name="ex10x10" value="Y" <% if ex10x10<>"" then response.write "checked" %> > 텐바이텐입금 제외
			&nbsp;
			<input type="checkbox" name="showdismatch" value="Y" <% if showdismatch<>"" then response.write "checked" %> > 매칭제외 포함
		</td>
	</tr>

	</form>
</table>
<!-- 검색 끝 -->

<% if (jungsanidx <> "") then %>

	<p>

	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25">
			<td width="120" bgcolor="<%= adminColor("tabletop") %>">IDX</td>
			<td bgcolor="#FFFFFF" colspan="3"><%= ofranchulgojungsan.FOneItem.Fidx %></td>
		</tr>
		<tr height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">매출처</td>
			<td bgcolor="#FFFFFF" ><%= ofranchulgojungsan.FOneItem.Fshopid %></td>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>">정산대상월</td>
			<td bgcolor="#FFFFFF" ><%= ofranchulgojungsan.FOneItem.FYYYYMM %></td>
		</tr>
		<tr height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">구분</td>
			<td bgcolor="#FFFFFF" >
				<%= ofranchulgojungsan.FOneItem.getShopDivName %>
				/
				<font color="<%= ofranchulgojungsan.FOneItem.GetDivCodeColor %>"><%= ofranchulgojungsan.FOneItem.GetDivCodeName %></font>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>">차수</td>
		    <td bgcolor="#FFFFFF" >
				<%= ofranchulgojungsan.FOneItem.FdiffKey %>
		    </td>
		</tr>
		<tr height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">Title</td>
			<td bgcolor="#FFFFFF" colspan="3">
				<%= ofranchulgojungsan.FOneItem.Ftitle %>
			</td>
		</tr>
		<tr height="25">
			<td bgcolor="<%= adminColor("tabletop") %>"><b>총출고가</b></td>
			<td bgcolor="#FFFFFF">
				<%= FormatNumber(ofranchulgojungsan.FOneItem.Ftotalsuplycash,0) %>
				<font color="#AAAAAA">(매출처로 공급한 상품가격)</font>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"><b>총발행금액</b></td>
			<td bgcolor="#FFFFFF">
				<%= FormatNumber(ofranchulgojungsan.FOneItem.Ftotalsum,0) %>
				<font color="#AAAAAA">(계산서 발행 금액)</font>
			</td>
		</tr>
	</table>

<% end if %>
<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11">
		검색결과 : <b><%= oipkum.FTotalCount - matchexcludecnt %></b> (매칭제외 : <%= matchexcludecnt %>)
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td>IDX</td>
	<td width="70">은행명</td>
	<td width="100">계좌번호</td>
	<td width="70">입출금일</td>
	<td>적요</td>
  	<td width="80">입금금액</td>
  	<td width="80">출금금액</td>
  	<td width="80">매칭완료</td>
  	<td width="60">상태</td>
  	<td>매칭금액</td>
  	<td>비고</td>
</tr>
<% if oipkum.FResultCount > 0 then %>
	<% for i=0 to oipkum.FResultCount-1 %>
		<%
		if IsNull(oipkum.Fipkumitem(i).Fmatchstate) then
			oipkum.Fipkumitem(i).Fmatchstate = "N"
		end if
		%>
		<% if (oipkum.Fipkumitem(i).Fmatchstate <> "X") or showdismatch <> "" then %>
		<form name="frmmatch<%= i %>" method="post" action="pop_ipkum_search_process.asp">
		<input type="hidden" name="mode" value="addmatch">
		<input type="hidden" name="jungsanidx" value="<%= jungsanidx %>">
		<input type="hidden" name="inoutidx" value="<%= oipkum.Fipkumitem(i).Finoutidx %>">
		<tr align="center" bgcolor="#FFFFFF" height="25">
			<td><%= oipkum.Fipkumitem(i).Finoutidx %></td>
			<td>
				<%= oipkum.Fipkumitem(i).Fbkname %>
			</td>
			<td>
				<%= oipkum.Fipkumitem(i).Fbkacctno %>
			</td>
			<td>
				<%= mid(oipkum.Fipkumitem(i).Fbkdate,1,4) %>-<%= mid(oipkum.Fipkumitem(i).Fbkdate,5,2) %>-<%= mid(oipkum.Fipkumitem(i).Fbkdate,7,2) %>
			</td>
			<td>
				<%= oipkum.Fipkumitem(i).Fbkjukyo %>
			</td>
		  	<td>
				<% if oipkum.Fipkumitem(i).finout_gubun = "2" then %>
					<%= FormatNumber(oipkum.Fipkumitem(i).Fbkinput,0) %>
				<% end if %>
		  	</td>
		  	<td>
				<% if oipkum.Fipkumitem(i).finout_gubun = "1" then %>
					<%= FormatNumber(oipkum.Fipkumitem(i).Fbkinput,0) %>
				<% end if %>
		  	</td>
		  	<td>
		  		<%= FormatNumber(oipkum.Fipkumitem(i).Ftotmatchedprice,0) %>
		  	</td>
			<td>
				<font color="<%= oipkum.Fipkumitem(i).GetMatchStateColor %>"><%= oipkum.Fipkumitem(i).GetMatchStateName %></font>
			</td>
			<td>
				<input type="text" class="text" name="matchprice" size="10" value="<%= oipkum.Fipkumitem(i).Fbkinput - oipkum.Fipkumitem(i).Ftotmatchedprice %>">
			</td>
			<td>
				<input type="button" class="button_s" value="매칭하기" onClick="SubmitMatch(frmmatch<%= i %>)" <% if (oipkum.Fipkumitem(i).Fmatchstate = "Y") then %>disabled<% end if %>>
				<input type="button" class="button_s" value="매칭제외" onClick="SubmitDisMatch(frmmatch<%= i %>)" <% if (oipkum.Fipkumitem(i).Fmatchstate <> "N") then %>disabled<% end if %>>
			</td>
		</tr>
		</form>
		<% end if %>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>




<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
