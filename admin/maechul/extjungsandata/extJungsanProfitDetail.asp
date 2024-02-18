<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsanProfitCls.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim fromDate, toDate
dim research : research = requestCheckvar(request("research"),10)
dim sellsite : sellsite = requestCheckvar(request("sellsite"),32)
dim makerid  : makerid = requestCheckvar(request("makerid"),32)
dim rtnExcept : rtnExcept = requestCheckvar(request("rtnExcept"),10)
dim minusgain : minusgain = requestCheckvar(request("minusgain"),10)
dim itemid : itemid = requestCheckvar(request("itemid"),10)
dim page : page = requestCheckvar(request("page"),10)

yyyy1   = request("yyyy1")
mm1     = request("mm1")
dd1     = request("dd1")
yyyy2   = request("yyyy2")
mm2     = request("mm2")
dd2     = request("dd2")

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)
	toDate = LEFT(DateAdd("d", -1,now()),10)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	yyyy2 = Cstr(Year(toDate))
	mm2 = Cstr(Month(toDate))
	dd2 = Cstr(day(toDate))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2)
end if

if (page="") then page=1

Dim oExtJungsanProfit, i
set oExtJungsanProfit = new CExtJungsanProfit
	oExtJungsanProfit.FPageSize = 1000
	oExtJungsanProfit.FCurrPage = 1

	oExtJungsanProfit.FRectStartdate = fromDate
	oExtJungsanProfit.FRectEndDate = toDate

	''oExtJungsanProfit.FRectGroupGubun = grpgubun
	oExtJungsanProfit.FRectSellSite = sellsite
    oExtJungsanProfit.FRectMakerid = makerid
    oExtJungsanProfit.FRectReturnExcept = CHKIIF(rtnExcept="on","1","")
    oExtJungsanProfit.FRectMinusGainOnly = CHKIIF(minusgain="on","1","")
    oExtJungsanProfit.FRectItemid = itemid
    oExtJungsanProfit.GetExtJungsanProfitDetail

Dim oCExtJungsan
if (itemid<>"") then
    SET oCExtJungsan = new CExtJungsan

    oCExtJungsan.FPageSize = 100
	oCExtJungsan.FCurrPage = page

    oCExtJungsan.FRectStartdate = fromDate
	oCExtJungsan.FRectEndDate = toDate
    oCExtJungsan.FRectSellSite = sellsite
    oCExtJungsan.FRectMakerid = makerid
    oCExtJungsan.FRectReturnExcept = CHKIIF(rtnExcept="on","1","")
    oCExtJungsan.FRectMinusGainOnly = CHKIIF(minusgain="on","1","")
    oCExtJungsan.FRectItemid = itemid

    oCExtJungsan.GetExtJungsanByItemDW
end if

Dim TTLextitemno
Dim TTLextTenMeachulPrice, TTLextTenJungsanPrice, TTLtenbuycash, TTLjungsangain
Dim TTLU_extTenJungsanPrice, TTLU_buycash, TTLU_jungsangain
Dim TTLW_extTenJungsanPrice, TTLW_buycash, TTLW_jungsangain
Dim TTLM_extTenJungsanPrice, TTLM_buycash, TTLM_jungsangain
Dim TTLN_extTenMeachulPrice, TTLN_buycash, TTLN_jungsangain


%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function rePage(isellsite,iitemid){
    document.frm.sellsite.value = isellsite;
    document.frm.itemid.value = iitemid;
    document.frm.submit();
}
</script>

<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 제휴몰:
		<select class="select" name="sellsite" onChange="jsGrpAuto(this.value);">
			<option value="">- 전체 -</option>
			<option value="interpark" <% if (sellsite = "interpark") then %>selected<% end if %> >인터파크</option>
			<option value="lotteimall" <% if (sellsite = "lotteimall") then %>selected<% end if %> >롯데아이몰</option>
			<option value="lotteCom" <% if (sellsite = "lotteCom") then %>selected<% end if %> >롯데닷컴</option>
			<option value="11st1010" <% if (sellsite = "11st1010") then %>selected<% end if %> >11번가</option>
			<option value="auction1010" <% if (sellsite = "auction1010") then %>selected<% end if %> >옥션</option>
			<option value="gmarket1010" <% if (sellsite = "gmarket1010") then %>selected<% end if %> >지마켓(NEW)</option>
			<!-- option value="lotteComM" <% if (sellsite = "lotteComM") then %>selected<% end if %> >롯데닷컴(직매출)</option -->
			<option value="gseshop" <% if (sellsite = "gseshop") then %>selected<% end if %> >GS샵</option>
			<!-- option value="dnshop" <% if (sellsite = "dnshop") then %>selected<% end if %> >디앤샵</option -->
			<option value="cjmall" <% if (sellsite = "cjmall") then %>selected<% end if %> >CJ몰</option>
			<!-- option value="wizwid" <% if (sellsite = "wizwid") then %>selected<% end if %> >위즈위드</option -->
			<!-- option value="gabangpop" <% if (sellsite = "gabangpop") then %>selected<% end if %> >패션팝(가방팝)</option -->
			<!-- option value="wconcept" <% if (sellsite = "wconcept") then %>selected<% end if %> >더블유컨셉</option -->
			<!-- option value="privia" <% if (sellsite = "privia") then %>selected<% end if %> >현대프리비아</option -->
			<!-- option value="player" <% if (sellsite = "player") then %>selected<% end if %> >플레이어</option -->
			<option value="homeplus" <% if (sellsite = "homeplus") then %>selected<% end if %> >홈플러스</option>
			<option value="ssg" <% if (sellsite = "ssg") then %>selected<% end if %> >SSG</option>
			<option value="ssg6006" <% if (sellsite = "ssg6006") then %>selected<% end if %> >SSG-이마트</option>
			<option value="ssg6007" <% if (sellsite = "ssg6007") then %>selected<% end if %> >SSG-ssg</option>
			<option value="nvstorefarm" <% if (sellsite = "nvstorefarm") then %>selected<% end if %> >스토어팜</option>
			<option value="ezwel" <% if (sellsite = "ezwel") then %>selected<% end if %> >이지웰페어</option>
			<option value="kakaogift" <% if (sellsite = "kakaogift") then %>selected<% end if %> >카카오기프트</option>
			<option value="coupang" <% if (sellsite = "coupang") then %>selected<% end if %> >쿠팡</option>
			<option value="halfclub" <% if (sellsite = "halfclub") then %>selected<% end if %> >하프클럽</option>
			<option value="hmall" <% if (sellsite = "hmall") then %>selected<% end if %> >Hmall</option>
			
		</select>
		
		&nbsp;
		* 매출일자:
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" style="width:70px;height:50px;" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr>
    <td bgcolor="#FFFFFF">
    브랜드ID : <input type="text" name="makerid" value="<%=makerid%>" size=20 maxlength=32>
    &nbsp;&nbsp;상품코드 : <input type="text" name="itemid" value="<%=itemid%>" size=10 maxlength=10>
    &nbsp;&nbsp;<input type="checkbox" name="rtnExcept" <%=CHKIIF(rtnExcept="on","checked","")%> >반품정산제외
    &nbsp;&nbsp;<input type="checkbox" name="minusgain" <%=CHKIIF(minusgain="on","checked","")%> >마이너스수익만보기
    </td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<p>
<!-- 액션 시작 -->
<form name="frmAct">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	
	</td>
	<td align="right">
	   
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">제휴몰</td>
    <td width="90">브랜드</td>
    <td width="70">상품코드</td>
    <td width="50">이미지</td>
    <td width="60">수량</td>
    <td width="70">매출액</td>
    <td width="70">정산액</td>
    <td width="70">매입액</td>
    <td width="70">수익</td>
    <td>상품명</td>
    <td width="60">판매가</td>
    <td width="60">매입가</td>
    <td width="60">판매상태</td>
    <td width="60">상세</td>
</tr>
<% for i=0 to oExtJungsanProfit.FresultCount -1 %>
<%
    TTLextitemno = TTLextitemno + oExtJungsanProfit.FItemList(i).Fextitemno
    TTLextTenMeachulPrice = TTLextTenMeachulPrice + oExtJungsanProfit.FItemList(i).FextTenMeachulPrice
    TTLextTenJungsanPrice = TTLextTenJungsanPrice + oExtJungsanProfit.FItemList(i).FextTenJungsanPrice
    TTLtenbuycash 		  = TTLtenbuycash + oExtJungsanProfit.FItemList(i).Ftenbuycash
    TTLjungsangain 		  = TTLjungsangain + oExtJungsanProfit.FItemList(i).Fjungsangain
    
%>
<tr bgcolor="FFFFFF" align="right" >
	<td align="left" ><%=oExtJungsanProfit.FItemList(i).FSellsite %></td>
    <td align="left" ><%=oExtJungsanProfit.FItemList(i).FMakerid %></td>
    <td align="center" ><%=oExtJungsanProfit.FItemList(i).FItemID %></td>
    <td align="center" ><img src="<%=oExtJungsanProfit.FItemList(i).FSmallimage %>"></td>
    <td align="center"><%=FormatNumber(oExtJungsanProfit.FItemList(i).Fextitemno,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FextTenMeachulPrice,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FextTenJungsanPrice,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).Ftenbuycash,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).Fjungsangain,0) %></td>
    <td align="left"><%=oExtJungsanProfit.FItemList(i).FItemName %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).Fsellcash,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).Fbuycash,0) %></td>
    <td align="center"><%=oExtJungsanProfit.FItemList(i).getItemStatHtml%></td> 
    <td align="center"><a href="javascript:rePage('<%=oExtJungsanProfit.FItemList(i).FSellsite %>','<%=oExtJungsanProfit.FItemList(i).FItemID %>');">상세</a></td>
</tr>
<% next %>
<tr bgcolor="FFFFFF" align="right">
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td align="center"><%=FormatNumber(TTLextitemno,0)%></td>
    <td><%=FormatNumber(TTLextTenMeachulPrice,0)%></td>
    <td><%=FormatNumber(TTLextTenJungsanPrice,0)%></td>
    <td><%=FormatNumber(TTLtenbuycash,0)%></td>
    <td><%=FormatNumber(TTLjungsangain,0)%></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
</tr>
</table>


<% if (itemid<>"") then %>
<br><p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">제휴몰</td>
	<td width="80">매출일자</td>
	<td width="150">제휴<br>주문번호</td>
	<td width="60">제휴<br>주문순번</td>
	<td width="150">제휴<br>원주문번호</td>
	<td width="40">수량</td>

	<td width="60">판매가</td>
	<td width="60">제휴부담<br>쿠폰</td>
	<td width="60">텐텐부담<br>쿠폰</td>
	<td width="60">쿠폰가</td>
	<td width="70"><b>매출금액</b></td>
	<td width="60">수수료</td>
	<td width="70">정산금액</td>
    <td width="70">매입금액</td>
    <td width="70">수익</td>
	<td width="70">수수료율</td>
    <td width="70">매입구분</td>
	<td width="80">원주문번호</td>
	<td width="100">상품코드</td>
	<td width="60">옵션코드</td>
	<td>비고</td>
</tr>
<% for i=0 to oCExtJungsan.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCExtJungsan.FItemList(i).GetSellSiteName %></td>
	<td><%= oCExtJungsan.FItemList(i).FextMeachulDate %></td>
	<td><%= oCExtJungsan.FItemList(i).FextOrderserial %></td>
	<td><%= oCExtJungsan.FItemList(i).FextOrderserSeq %></td>
	<td><%= oCExtJungsan.FItemList(i).FextOrgOrderserial %></td>
	<td><%= oCExtJungsan.FItemList(i).FextItemNo %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextItemCost, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextOwnCouponPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextTenCouponPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextReducedPrice, 0) %></td>
	<td align="right"><b><%= FormatNumber(oCExtJungsan.FItemList(i).FextTenMeachulPrice, 0) %></b>
	<% if (oCExtJungsan.FItemList(i).GetDiffMeachulPrice<>0) then %>
		<br>(<font color="red"><%=formatNumber(oCExtJungsan.FItemList(i).GetDiffMeachulPrice,0)%></font>)
	<% end if %>
	</td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextCommPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextTenJungsanPrice, 0) %>
	<% if (oCExtJungsan.FItemList(i).GetDiffJungsanPrice<>0) then %>
		<br>(<font color="red"><%=formatNumber(oCExtJungsan.FItemList(i).GetDiffJungsanPrice,0)%></font>)
	<% end if %>
	</td>
    <td align="right"><%=formatNumber(oCExtJungsan.FItemList(i).Ftenbuycash,0)%></td>
    <td align="right"><%=formatNumber(oCExtJungsan.FItemList(i).Fjungsangain,0)%></td>
	<td>
		<%=oCExtJungsan.FItemList(i).GetSusumargin%>
	</td>
    <td ><%=oCExtJungsan.FItemList(i).Fmwdiv%></td>
	<td><%= oCExtJungsan.FItemList(i).FOrgOrderserial %></td>
	<td><%= oCExtJungsan.FItemList(i).Fitemid %></td>
	<td><%= oCExtJungsan.FItemList(i).Fitemoption %></td>
	<td>
		<% if NOT isNULL(oCExtJungsan.FItemList(i).FMinusOrderserial) then %>
			<%= oCExtJungsan.FItemList(i).FMinusOrderserial %>
		<% end if %>

		<% if (oCExtJungsan.FItemList(i).GetDiffReducedPrice <> 0) then %>
			<% if NOT isNULL(oCExtJungsan.FItemList(i).FMinusOrderserial) then %><br><% end if %>
		<%= oCExtJungsan.FItemList(i).GetDiffReducedPrice %>
		<% end if %>
	</td>
    
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="21" align="center">
		<% if oCExtJungsan.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCExtJungsan.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCExtJungsan.StartScrollPage to oCExtJungsan.FScrollCount + oCExtJungsan.StartScrollPage - 1 %>
			<% if i>oCExtJungsan.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCExtJungsan.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>
<% end if %>

<%
if (itemid<>"") then
    set oCExtJungsan = Nothing
end if
set oExtJungsanProfit = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
