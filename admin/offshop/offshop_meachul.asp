<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 가맹점 정산관리(매출)
' History : 2009.04.07 서동석 생성
'			2010.05.13 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->
<%
dim page, shopid , yyyy1 , mm1 , dd1 , yyyy2 , mm2 , dd2 , designer, statecd , divcode
dim i,totalsum, totalsuply, totalerr, totalbuy , fromDate , toDate ,shopdiv
dim bankinoutidx
	yyyy1 = RequestCheckvar(request("yyyy1"),10)
	mm1 = RequestCheckvar(request("mm1"),10)
	dd1 = RequestCheckvar(request("dd1"),10)
	yyyy2 = RequestCheckvar(request("yyyy2"),10)
	mm2 = RequestCheckvar(request("mm2"),10)
	dd2 = RequestCheckvar(request("dd2"),10)
	designer = RequestCheckvar(request("designer"),32)
	statecd  = RequestCheckvar(request("statecd"),10)
	shopid = RequestCheckvar(request("shopid"),32)
	divcode = RequestCheckvar(request("divcode"),10)
    shopdiv = RequestCheckvar(request("shopdiv"),10)
    bankinoutidx = RequestCheckvar(request("bankinoutidx"),32)

if (yyyy1="") then yyyy1 = Cstr(Year(Dateadd("d",now(),-30)))
if (mm1="") then mm1 = Cstr(Month(Dateadd("d",now(),-30)))
''if (dd1="") then dd1 = Cstr(day(Dateadd("d",now(),-30)))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
''if (dd2="") then dd2 = Cstr(day(now()))

'''2010 정산대상 년월로 변경
fromDate = yyyy1+"-"+mm1
toDate = yyyy2+"-"+mm2
''fromDate =DateSerial(yyyy1, mm1, dd1)
''toDate = DateAdd("d",1,DateSerial(yyyy2, mm2, dd2))

page = request("page")
if page="" then page=1

dim ofranchulgojungsan
	set ofranchulgojungsan = new CFranjungsan
	ofranchulgojungsan.FPageSize=50
	ofranchulgojungsan.FCurrpage = page
	ofranchulgojungsan.FRectshopid = shopid
	ofranchulgojungsan.FRectdivcode = divcode
	ofranchulgojungsan.FRectStateCd = statecd
''rw 	fromDate
''rw 	toDate

	if (bankinoutidx = "") then
		'// 입출금IDX 검색시 날짜 제외
		ofranchulgojungsan.FRectStartDate = fromDate
		ofranchulgojungsan.FRectendDate = toDate
	else
		ofranchulgojungsan.FRectBankInOutIdx = bankinoutidx
	end if

	ofranchulgojungsan.FRectShopDiv = shopdiv

	ofranchulgojungsan.getFranJungsanList()
%>

<script language='javascript'>

function popAddFranMeachul(){
	var popwin = window.open('popmeaipchulgojungsanmaker.asp?shopid=' + document.frm.shopid.value,'franmeaip','width=950, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function popMasterEdit(iid){
	var popwin = window.open('popmeaipchulgoedit.asp?idx=' + iid,'franmeaipedit','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function popMasterAdd(){
	var popwin = window.open('popmeaipchulgoedit.asp','franmeaipedit','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function popaddb2c(){
	var popaddb2c = window.open('popb2cmaechul.asp?shopid=' + document.frm.shopid.value,'popaddb2c','width=1024, height=768, scrollbars=yes, resizable=yes');
	popaddb2c.focus();
}

function DelThis(iid){
	if (!confirm('정말로 삭제 하시겠습니까?')){
		return;
	}

	var popwin = window.open('meaipchulgojungsan_process.asp?mode=delmaster&idx=' + iid,'delfrm','width=400, height=400, scrollbars=yes, resizable=yes');

}

function popSubmasterEdit(iid){
	var popwin = window.open('franmeaippopsubmaster.asp?idx=' + iid,'popsubmaster','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function popIpkumSearch(jungsanidx, serchtype, searchstring, yyyy1, mm1, yyyy2, mm2) {
	var popwin;
	if (serchtype == "txammount") {
		popwin = window.open('pop_ipkum_search.asp?jungsanidx=' + jungsanidx + '&serchtype=' + serchtype + '&txammount=' + searchstring + '&yyyy1=' + yyyy1 + '&mm1=' + mm1 + '&yyyy2=' + yyyy2 + '&mm2=' + mm2,'popIpkumSearch','width=900, height=500, scrollbars=yes, resizable=yes');
	} else {
		popwin = window.open('pop_ipkum_search.asp?jungsanidx=' + jungsanidx + '&serchtype=' + serchtype + '&jeokyo=' + searchstring + '&yyyy1=' + yyyy1 + '&mm1=' + mm1 + '&yyyy2=' + yyyy2 + '&mm2=' + mm2,'popIpkumSearch','width=900, height=500, scrollbars=yes, resizable=yes');
	}
	popwin.focus();
}

function popIpkumList(jungsanidx) {
	var popwin = window.open('pop_ipkum_list.asp?jungsanidx=' + jungsanidx,'popIpkumList','width=800, height=500, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function NextPage(page){
	document.frm.page.value=page;
	document.frm.submit();
}

function regOffTax(idx){
	var popwin = window.open("pop_offshop_TaxReg.asp?idx=" + idx,"popOffTaxReg","width=640 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function registerOffShopTax(idx){
	// var popwin = window.open("/cscenter/taxsheet/tax_register_offshop.asp?idx=" + idx,"registerOffShopTax","width=1000 height=800 scrollbars=yes resizable=yes");
	var popwin = window.open("/cscenter/taxsheet/tax_register_new.asp?issuetype=etcmeachul&idx=" + idx,"registerOffShopTax","width=850 height=650 scrollbars=yes resizable=yes");
	popwin.focus();
}

function modifyInvoice(shopid, idx, workidx, invoiceidx){
	if (workidx == "") {
		alert("먼저 작업을 지정하세요");
		return;
	}

	var popwin = window.open("/admin/fran/offinvoice_modify.asp?shopid=" + shopid + "&jungsanidx=" + idx + "&workidx=" + workidx + "&invoiceidx=" + invoiceidx,"modifyInvoice","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

// 공급자용 세금계산서
function popTaxPrint(taxNo, bizNo){
	var s_biz_no = "2118700620";	// 텐바이텐 사업자번호

	//	리얼서버	http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp
	//	테스트		http://ifs.neoport.net/jsp/dti/tx/dti_get_pin.jsp
	var popwinsub = window.open("http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+taxNo+"&cur_biz_no="+s_biz_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+bizNo,"taxview","width=670,height=620,status=no, scrollbars=auto, menubar=no, resizable=yes");

	popwinsub.focus();
}

function goView_Bill36524(tax_no, b_biz_no)
{
		window.open("http://www.bill36524.com/popupBillTax.jsp?NO_TAX=" + tax_no + "&NO_BIZ_NO="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
}

function PopExportSheet(v){
	var popwin;
	popwin = window.open('/admin/fran/cartoonbox_modify.asp?menupos=1357&idx=' + v ,'PopExportSheet','width=1000,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		가맹점 : <% drawSelectBoxOffShop "shopid",shopid %>
		&nbsp;
		구분 :
		<% Call DrawShopDivBox(shopdiv) %>
		&nbsp;
		<select class="select" name="divcode">
			<option value="">전체
			<option value="GC" <% if divcode="GC" then response.write "selected" %> >가맹비
			<option value="MC" <% if divcode="MC" then response.write "selected" %> >매입출고
			<option value="WS" <% if divcode="WS" then response.write "selected" %> >위탁판매
		</select>
		&nbsp;
		상태 :
		<select class="select" name="statecd">
			<option value="">전체
			<option value="0" <% if statecd="0" then response.write "selected" %> >수정중
			<option value="1" <% if statecd="1" then response.write "selected" %> >업체확인중
			<option value="4" <% if statecd="4" then response.write "selected" %> >계산서발행
			<option value="7" <% if statecd="7" then response.write "selected" %> >입금완료
		</select>
		<br>
		검색기간 :
		<% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %>	(정산년월)
		&nbsp;
		입출금IDX :
		<input type="text" class="text" name="bankinoutidx" value="<%= bankinoutidx %>">
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<p>

<br><font size=5>폐기예정 메뉴입니다.</font><br><br>
<font color=red>[경영]매출관리>>기타매출관리</font> 를 이용하세요<br><br>

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" border="0" cellpadding="1" cellspacing="1" class="a" >
<tr>
	<td align="left">
		<input type="button" class="button" value="상품대금작성" onClick="javascript:popAddFranMeachul();" disabled>
		<input type="button" class="button" value="기타비용작성(가맹비등)" onClick="javascript:popMasterAdd();" disabled>
		<input type="button" class="button" value="B2C매출작성" onClick="javascript:popaddb2c();" disabled>
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<h4>7월 가맹점 정산건 작성시 서팀 문의 요망 (제주,진주,일산)</h4>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=35>IDX</td>
	<td width=60>정산년월</td>
	<td width=35>발행<br>차수</td>
	<td width=80>오프샵ID</td>
	<td width=40>구분</td>
	<td width=55>구분</td>
	<td>제목</td>
	<td width=60>발행금액</td>
	<td width=60>실공급액</td>
	<td width=30>오차<br>금액</td>
	<td width=65>매입가액</td>
	<td width=40>수익율</td>
	<td width=65>세금발행일</td>
	<td width=65>입금일</td>
	<td width=65>입금확인액</td>
	<td width=65>현재상태</td>
	<td width=90>계산서발행</td>
	<td width=30>삭제</td>
</tr>
<% if ofranchulgojungsan.FResultCount >0 then %>
<% for i=0 to ofranchulgojungsan.FResultCount-1 %>
<%
totalsum = totalsum + ofranchulgojungsan.FItemList(i).Ftotalsum
totalsuply  = totalsuply + ofranchulgojungsan.FItemList(i).Ftotalsuplycash
totalerr = totalerr  + ofranchulgojungsan.FItemList(i).Ftotalsum -  ofranchulgojungsan.FItemList(i).Ftotalsuplycash
totalbuy = totalbuy + ofranchulgojungsan.FItemList(i).Ftotalbuycash

if IsNull(ofranchulgojungsan.FItemList(i).Ftotmatchedipkumsum) then
	ofranchulgojungsan.FItemList(i).Ftotmatchedipkumsum = 0
end if

%>
<tr bgcolor="#FFFFFF">
	<td align=center><%= ofranchulgojungsan.FItemList(i).Fidx %></td>
	<td align=center><%= ofranchulgojungsan.FItemList(i).FYYYYMM %></td>
	<td align=center><%= ofranchulgojungsan.FItemList(i).FDiffKey %></td>
	<td align=center><a href="javascript:popMasterEdit('<%= ofranchulgojungsan.FItemList(i).Fidx %>');"><%= ofranchulgojungsan.FItemList(i).Fshopid %></a></td>
	<td align=center><%= ofranchulgojungsan.FItemList(i).getShopDivName() %></td>
	<td align=center><font color="<%= ofranchulgojungsan.FItemList(i).GetDivCodeColor %>"><%= ofranchulgojungsan.FItemList(i).GetDivCodeName %></font></td>
	<td><a href="javascript:popSubmasterEdit('<%= ofranchulgojungsan.FItemList(i).Fidx %>');"><%= ofranchulgojungsan.FItemList(i).Ftitle %></a></td>
	<td align=right><%= formatNumber(ofranchulgojungsan.FItemList(i).Ftotalsum,0) %></td>
	<td align=right><%= formatNumber(ofranchulgojungsan.FItemList(i).Ftotalsuplycash,0) %></td>
	<td align=right><%= formatNumber(ofranchulgojungsan.FItemList(i).Ftotalsum-ofranchulgojungsan.FItemList(i).Ftotalsuplycash,0) %></td>
	<td align=right><%= formatNumber(ofranchulgojungsan.FItemList(i).Ftotalbuycash,0) %></td>
	<td align=right>
		<% if ofranchulgojungsan.FItemList(i).Ftotalsum<>0 then %>
		<%= CLng(10000-(ofranchulgojungsan.FItemList(i).Ftotalbuycash/ofranchulgojungsan.FItemList(i).Ftotalsum*100*100))/100 %>%
		<% end if %>
	</td>
	<td align=center><%= Left(ofranchulgojungsan.FItemList(i).Ftaxdate,10) %></td>
	<td align=center>
		<% if (ofranchulgojungsan.FItemList(i).FStateCd >= "1") then %>
			<% if (ofranchulgojungsan.FItemList(i).Fipkumdate = "") or IsNull(ofranchulgojungsan.FItemList(i).Fipkumdate) then %>
				<input type="button" class="button" value="찾기" onClick="popIpkumSearch(<%= ofranchulgojungsan.FItemList(i).Fidx %>, 'txammount', <%= ofranchulgojungsan.FItemList(i).Ftotalsum - ofranchulgojungsan.FItemList(i).Ftotmatchedipkumsum %>, '<%= yyyy1 %>', '<%= mm1 %>', '<%= yyyy2 %>', '<%= mm2 %>')">
			<% else %>
				<a href="javascript:popIpkumSearch(<%= ofranchulgojungsan.FItemList(i).Fidx %>, 'txammount', <%= ofranchulgojungsan.FItemList(i).Ftotalsum - ofranchulgojungsan.FItemList(i).Ftotmatchedipkumsum %>, '<%= yyyy1 %>', '<%= mm1 %>', '<%= yyyy2 %>', '<%= mm2 %>')"><%= ofranchulgojungsan.FItemList(i).Fipkumdate %></a>
			<% end if %>
		<% else %>
			<%= ofranchulgojungsan.FItemList(i).Fipkumdate %>
		<% end if %>
	</td>
	<td align=center>
		<% if (IsNull(ofranchulgojungsan.FItemList(i).Ftotmatchedipkumsum) or (ofranchulgojungsan.FItemList(i).Ftotmatchedipkumsum = 0)) then %>
			<% if (Not IsNull(ofranchulgojungsan.FItemList(i).Fmaymatchedipkumsum)) then %>
				<font color=gray><%= FormatNumber(ofranchulgojungsan.FItemList(i).Ftotalsum,0) %></font>
			<% end if %>
		<% else %>
			<a href="javascript:popIpkumList(<%= ofranchulgojungsan.FItemList(i).Fidx %>)">
				<% if (ofranchulgojungsan.FItemList(i).Ftotalsum = ofranchulgojungsan.FItemList(i).Ftotmatchedipkumsum) then %>
					<%= formatNumber(ofranchulgojungsan.FItemList(i).Ftotmatchedipkumsum,0) %>
				<% else %>
					<font color=red><%= formatNumber(ofranchulgojungsan.FItemList(i).Ftotmatchedipkumsum,0) %></font>
				<% end if %>
			</a>
		<% end if %>
	</td>
	<td align=center><font color="<%= ofranchulgojungsan.FItemList(i).GetStateColor %>"><%= ofranchulgojungsan.FItemList(i).GetStateName %></font></td>
	<td align=center>
		<% if (ofranchulgojungsan.FItemList(i).FStateCd>"0") and (ofranchulgojungsan.FItemList(i).FStateCd<"4") then %>
			<!--
			<input type="button" class="button" value="발행" onclick="regOffTax('<%= ofranchulgojungsan.FItemList(i).Fidx %>');">
			-->
			<% if (ofranchulgojungsan.FItemList(i).Fshopdiv = "7") then %>
				<!-- 수출 -->
				<% if (ofranchulgojungsan.FItemList(i).Finvoiceidx <> "") and (Not IsNull(ofranchulgojungsan.FItemList(i).Finvoiceidx)) then %>
					<a href="javascript:modifyInvoice('<%= ofranchulgojungsan.FItemList(i).Fshopid %>', '<%= ofranchulgojungsan.FItemList(i).Fidx %>', '<%= ofranchulgojungsan.FItemList(i).Fworkidx %>', '<%= ofranchulgojungsan.FItemList(i).Finvoiceidx %>');">IDX : <%= ofranchulgojungsan.FItemList(i).Finvoiceidx %></a>
				<% else %>
					<input type="button" class="button" value="인보이스" onclick="modifyInvoice('<%= ofranchulgojungsan.FItemList(i).Fshopid %>', '<%= ofranchulgojungsan.FItemList(i).Fidx %>', '<%= ofranchulgojungsan.FItemList(i).Fworkidx %>', '<%= ofranchulgojungsan.FItemList(i).Finvoiceidx %>');">
				<% end if %>
				<!--
				<% if (ofranchulgojungsan.FItemList(i).Fworkidx <> "") and (Not IsNull(ofranchulgojungsan.FItemList(i).Fworkidx)) then %>
					<br>
					<input type="button" class="button" value="출고내역" onclick="PopExportSheet('<%= ofranchulgojungsan.FItemList(i).Fworkidx %>');">
				<% end if %>
				-->
			<% else %>
				<input type="button" class="button" value="발행요청" onclick="registerOffShopTax('<%= ofranchulgojungsan.FItemList(i).Fidx %>');">
			<% end if %>
		<% elseif ofranchulgojungsan.FItemList(i).FStateCd>="4" then %>
			<%If ofranchulgojungsan.FItemList(i).FtaxNo <> "" Then %>
				<% if (Left(ofranchulgojungsan.FItemList(i).FtaxNo,2)="TX") then %>
				<a href="javascript:goView_Bill36524('<%=ofranchulgojungsan.FItemList(i).FtaxNo%>','2118700620');"><img src="/images/icon_print02.gif" border="0"></a>
				<% else %>
				<a href="javascript:popTaxPrint('<%=ofranchulgojungsan.FItemList(i).FtaxNo%>','<%=ofranchulgojungsan.FItemList(i).FbizNo%>');"><img src="/images/icon_print02.gif" border="0"></a>
	    		<% end if %>
			<%end if %>
			<% if (ofranchulgojungsan.FItemList(i).Fshopdiv = "7") then %>
				<!-- 수출 -->
				<% if (ofranchulgojungsan.FItemList(i).Finvoiceidx <> "") and (Not IsNull(ofranchulgojungsan.FItemList(i).Finvoiceidx)) then %>
					<a href="javascript:modifyInvoice('<%= ofranchulgojungsan.FItemList(i).Fshopid %>', '<%= ofranchulgojungsan.FItemList(i).Fidx %>', '<%= ofranchulgojungsan.FItemList(i).Fworkidx %>', '<%= ofranchulgojungsan.FItemList(i).Finvoiceidx %>');">IDX : <%= ofranchulgojungsan.FItemList(i).Finvoiceidx %></a>
				<% end if %>
			<% end if %>
		<%end if %>
	</td>
	<td align=center>
		<% if ofranchulgojungsan.FItemList(i).FStateCd="0" then %>
		<a href="javascript:DelThis('<%= ofranchulgojungsan.FItemList(i).Fidx %>');">X</a>
		<% end if %>
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td>총계</td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td align=right><%= formatNumber(totalsum,0) %></td>
	<td align=right><%= formatNumber(totalsuply,0) %></td>
	<td align=right><%= formatNumber(totalerr,0) %></td>
	<td align=right><%= formatNumber(totalbuy,0) %></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" height=20>
	<td colspan=21 align=center>
	<% if ofranchulgojungsan.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ofranchulgojungsan.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ofranchulgojungsan.StarScrollPage to ofranchulgojungsan.FScrollCount + ofranchulgojungsan.StarScrollPage - 1 %>
		<% if i>ofranchulgojungsan.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ofranchulgojungsan.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" >
	<td colspan="21" align="center">[검색 결과가 없습니다.]</td>
</tr>
</table>
<% end if %>

<%
set ofranchulgojungsan = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->