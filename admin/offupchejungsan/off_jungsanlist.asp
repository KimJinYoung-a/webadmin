<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인정산
' History : 서동석 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->

<%

dim makerid, yyyy1, mm1, finishflag, page, groupid, vPurchaseType, jgubun, jacctcd, differencekey
dim searchType, searchText, jungsanGubun
makerid 		= requestCheckVar(request("makerid"),32)
yyyy1   		= requestCheckVar(request("yyyy1"),10)
mm1     		= requestCheckVar(request("mm1"),10)
finishflag      = requestCheckVar(request("finishflag"),10)
page            = requestCheckVar(request("page"),10)
vPurchaseType   = requestCheckVar(request("purchasetype"),10)
jgubun          = requestCheckVar(request("jgubun"),10)
jacctcd 		= requestCheckVar(request("jacctcd"),10)
differencekey 	= requestCheckVar(request("differencekey"),10)
searchType 		= requestCheckVar(request("searchType"), 32)
searchText 		= requestCheckVar(request("searchText"), 32)
jungsanGubun    = requestCheckVar(request("jungsanGubun"), 12)

dim comm_cd : comm_cd     = RequestCheckVar(request("comm_cd"),9)

if page="" then page=1


dim taxtype, autojungsan, jungsan_date
taxtype      = requestCheckVar(request("taxtype"),32)
autojungsan  = requestCheckVar(request("autojungsan"),32)
jungsan_date = requestCheckVar(request("jungsan_date"),32)
groupid      = requestCheckVar(request("groupid"),32)

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

''임시
''yyyy1 = "2006"
''mm1="12"


dim ooffjungsan
set ooffjungsan = new COffJungsan
ooffjungsan.FPageSize   = 50
ooffjungsan.FCurrPage = page
ooffjungsan.FRectYYYYMM = yyyy1 + "-" + mm1
ooffjungsan.FRectfinishflag = finishflag
ooffjungsan.FRectMakerid = makerid
ooffjungsan.FRectTaxtype = taxtype
ooffjungsan.FRectAutojungsan = autojungsan
ooffjungsan.FRectJungsanDate = jungsan_date
ooffjungsan.FRectGroupID = groupid
ooffjungsan.FRectPurchaseType = vPurchaseType
ooffjungsan.FRectJungsanGubunCD = comm_cd
ooffjungsan.FRectJGubun = jgubun
ooffjungsan.FRectjacctcd = jacctcd
ooffjungsan.FRectdifferencekey = differencekey
ooffjungsan.FRectSearchType = searchType
ooffjungsan.FRectSearchText = searchText
ooffjungsan.FRectJungsanGubun = jungsanGubun
ooffjungsan.GetOffJungsanMasterList



dim i
dim orgsellmargin, realsellmargin
orgsellmargin   = 0
realsellmargin  = 0
%>
<script language='javascript'>
function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function MakeBrandBatchJungsan(frm){
    if (frm.jgubun.value.length<1){
        alert('정산 방식 구분을 선택 하세요.');
        frm.jgubun.focus();
        return;
    }

    if (frm.differencekey.value.length<1){
        alert('차수 구분을 선택 하세요.');
        frm.differencekey.focus();
        return;
    }

    if (frm.itemvatYN.value.length<1){
        alert('상품 과세 구분을 선택 하세요.');
        frm.itemvatYN.focus();
        return;
    }

    if (confirm('정산내역을 작성 하시겠습니까?')){
        var queryurl = 'off_jungsan_process.asp?mode=brandbatchprocess&jgubun='+frm.jgubun.value+'&makerid=' + frm.makerid.value + '&yyyy=' + frm.yyyy.value + '&mm=' + frm.mm.value + '&differencekey=' + frm.differencekey.value + '&itemvatYN=' + frm.itemvatYN.value+'&ipchulArr='+frm.ipchulArr.value;

        var popwin = window.open(queryurl ,'off_jungsan_process','width=200, height=200, scrollbars=yes, resizable=yes');
    }
}

function PopDetail(idx){
    var popwin = window.open('off_jungsandetailsum.asp?idx=' + idx ,'off_jungsandetailsum','width=960, height=540, scrollbars=yes, resizable=yes');
    popwin.focus();
}

function PopStateChange(idx){
    var popwin = window.open('off_jungsanstateedit.asp?idx=' + idx ,'off_jungsanstateedit','width=960, height=540, scrollbars=yes, resizable=yes');
    popwin.focus();
}

function DelMaster(idx){
    if (confirm('삭제 하시겠습니까?')){
        var popwin = window.open('off_jungsan_process.asp?mode=delmaster&masteridx=' + idx ,'off_jungsan_process','width=200, height=200, scrollbars=yes, resizable=yes');
    }
}

function PopTaxPrintReDirect(itax_no, makerid){
	var popwinsub = window.open("/admin/upchejungsan/red_taxprint.asp?tax_no=" + itax_no + "&makerid=" + makerid,"taxview","width=670,height=620,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popwinsub.focus();
}

function research(t){

}
</script>
<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
    <tr align="center" bgcolor="#FFFFFF" >
        <td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
        <td align="left">
            정산년월 : <% DrawYMBox yyyy1,mm1 %>&nbsp;&nbsp;
            계산서과세구분 :
            <select name="taxtype" class="select">
                <option value="" <% if taxtype="" then response.write "selected" %> >선택
                <option value="01" <% if taxtype="01" then response.write "selected" %> >과세
                <option value="02" <% if taxtype="02" then response.write "selected" %> >면세
                <option value="03" <% if taxtype="03" then response.write "selected" %> >간이
            </select>&nbsp;&nbsp;
            <!--
            수기구분 :
            <select name="autojungsan">
            <option value=""  <% if autojungsan="" then response.write "selected" %> >선택
            <option value="Y" <% if autojungsan="Y" then response.write "selected" %> >자동
            <option value="N" <% if autojungsan="N" then response.write "selected" %> >수기
            </select>&nbsp;&nbsp;
            -->
            정산일 :
            <select name="jungsan_date" class="select">
                <option value="" <% if jungsan_date="" then response.write "selected" %> >선택
                <option value="15일" <% if jungsan_date="15일" then response.write "selected" %> >15일
                <option value="말일" <% if jungsan_date="말일" then response.write "selected" %> >말일
                <option value="수시" <% if jungsan_date="수시" then response.write "selected" %> >수시
                <option value="NULL" <% if jungsan_date="NULL" then response.write "selected" %> >미지정
            </select>
            &nbsp;&nbsp;
            진행상태 : <% DrawOffJungsanStateCombo "finishflag", finishflag %>
            &nbsp;&nbsp;
            업체과세구분 : 
            <select name="jungsanGubun" class="select">
                <option value="" <% if jungsanGubun="" then response.write "selected" %>>전체</option>
                <option value="일반과세" <% if jungsanGubun="일반과세" then response.write "selected" %>>일반과세</option>
                <option value="간이과세" <% if jungsanGubun="간이과세" then response.write "selected" %>>간이과세</option>
                <option value="원천징수" <% if jungsanGubun="원천징수" then response.write "selected" %>>원천징수</option>
                <option value="면세" <% if jungsanGubun="면세" then response.write "selected" %>>면세</option>
                <option value="영세(해외)" <% if jungsanGubun="영세(해외)" then response.write "selected" %>>영세(해외)</option>
            </select>
        </td>
        <td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
    		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
    	</td>
    </tr>
    <tr>
        <td bgcolor="#FFFFFF" >
        구매유형 :
        <% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
        &nbsp;&nbsp;
		브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
		&nbsp;&nbsp;
		업체(그룹코드) : <input type="text" class="text" name="groupid" value="<%= groupid %>" size="12" >&nbsp;&nbsp;
        계정과목코드 : <input type="text" class="text" name="jacctcd" value="<%= jacctcd %>" size="7" >

        </td>
    </tr>
    <tr>
        <td bgcolor="#FFFFFF" >
			정산방식구분 :
			<% drawSelectBoxJGubun "jgubun",jgubun %>
			정산구분 :
			<% drawSelectBoxOFFJungsanCommCDQuery "comm_cd",comm_cd %>
			&nbsp;&nbsp;
			차수
			<input type="text" class="text" name="differencekey" value="<%= differencekey %>" size="2" >
			&nbsp;&nbsp;
			검색조건:
			<select class="select" name="searchType">
				<option></option>
				<option value="socname" <% if (searchType = "socname") then %>selected<% end if %> >업체명</option>
				<option value="socno" <% if (searchType = "socno") then %>selected<% end if %> >사업자번호</option>
			</select>
			&nbsp;
			<input type="text" class="text" name=searchText value="<%= searchText %>" size="15" maxlength="20">
        </td>
    </tr>
    </form>
</table>
<p>
<!-- 표 상단바 끝-->
<% if (makerid<>"") and (yyyy1<>"") and (mm1<>"") then %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="brandbatch" >
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="yyyy" value="<%= yyyy1 %>">
<input type="hidden" name="mm" value="<%= mm1 %>">
<tr bgcolor="#FFFFFF">
    <td>
        <select name="jgubun" class="select">
            <option value="">정산 방식 선택</option>
            <option value="MM">매입</option>
            <option value="CC">수수료</option>
            <option value="CE">기타매출</option>
        </select>
        <select name="differencekey" class="select">
            <option value="">차수 선택
            <option value="0">0차
            <option value="1">1차
            <option value="2">2차
            <option value="3">3차
            <option value="4">4차
            <option value="5">5차
            <option value="6">6차
            <option value="7">7차
            <option value="8">8차
            <option value="9">9차
        </select>
        <select name="itemvatYN" class="select">
            <option value="">상품 과세 구분 선택
            <option value="Y">과세
            <option value="N">면세
        </select>
        <!--
        &nbsp;입출코드<input type="text" name="ipchulArr" value="" size="20">
        -->
        <input type="hidden" name="ipchulArr" value="">
        <input type="button" value=" <%= makerid %> &nbsp;&nbsp;<%= yyyy1 %>년 <%= mm1 %>월 정산 작성 " onClick="MakeBrandBatchJungsan(brandbatch);">
    </td>
</form>
</tr>
</table>
<% end if %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="<%= adminColor("topbar") %>">
        <td colspan="25" align="right" >
            총건수: <%= FormatNumber(ooffjungsan.FTotalCount,0) %> &nbsp;&nbsp;
            총금액: <%= FormatNumber(ooffjungsan.FTotalSum,0) %> &nbsp;&nbsp;
            Page: <%= page %>/<%= ooffjungsan.FTotalPage %>
        </td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="50">정산월</td>
      <td width="50">정산<br>방식</td>
      <td width="60">계정<br>과목</td>
      <td width="30">차수</td>
      <td width="30">과세<br>(계산서)</td>
      <td width="30">과세<br>(상품)</td>
      <td width="90"><a href="javascript:research(frm,'makerid')">브랜드ID</a></td>
      <td width="70">그룹ID</td>
      <td width="56">위탁<br>판매</td>
      <td width="56">업체<br>위탁</td>
      <td width="56">오프<br>매입</td>
      <td width="56">매장<br>매입</td>
      <td width="56">출고<br>매입</td>
      <td width="56">기타<br>내역</td>

      <td width="70">총판매가</td>
      <td width="70">총매출액</td>
      <td width="70">총정산액</td>
      <td width="70">총수수료</td>
      <td width="60">소비<br>마진</td>
      <td width="60">매출<br>마진</td>
      <td width="80"><a href="javascript:research(frm,'state')">상태</a></td>
      <td width="70"><a href="javascript:research(frm,'segum')">세금<br>발행일</a></td>
      <td width="70">입금일</td>
      <td width="60">과세구분</td>
      <td width="30">비고</td>
    </tr>
    <% if ooffjungsan.FResultCount>0 then %>
    <% for i=0 to ooffjungsan.FResultCount-1 %>
    <%
        if (ooffjungsan.FItemList(i).Ftot_orgsellprice<>0) then
            orgsellmargin = CLng((ooffjungsan.FItemList(i).Ftot_orgsellprice-ooffjungsan.FItemList(i).Ftot_jungsanprice)/ooffjungsan.FItemList(i).Ftot_orgsellprice*100*100)/100
        else
            orgsellmargin = 0
        end if

        if (ooffjungsan.FItemList(i).Ftot_realsellprice<>0) then
            realsellmargin = CLng((ooffjungsan.FItemList(i).Ftot_realsellprice-ooffjungsan.FItemList(i).Ftot_jungsanprice)/ooffjungsan.FItemList(i).Ftot_realsellprice*100*100)/100
        else
            realsellmargin = 0
        end if
    %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td ><a href="javascript:PopDetail('<%= ooffjungsan.FItemList(i).Fidx %>');"><%= ooffjungsan.FItemList(i).FYYYYMM %></a></td>
      	<td ><%= ooffjungsan.FItemList(i).getJGubunName %></td>
      	<td ><%= ooffjungsan.FItemList(i).Fjacc_nm %></td>
      	<td ><%= ooffjungsan.FItemList(i).Fdifferencekey %></td>
      	<td ><font color="<%= ooffjungsan.FItemList(i).GetTaxtypeNameColor %>"><%= ooffjungsan.FItemList(i).GetSimpleTaxtypeName %></font></td>
      	<td ><%= ooffjungsan.FItemList(i).GetItemVatTypeName %></td>
      	<td align="left"><a href="javascript:PopUpcheBrandInfoEdit('<%= ooffjungsan.FItemList(i).Fmakerid %>');"><%= ooffjungsan.FItemList(i).Fmakerid %></a></td>
      	<td align="center"><%= ooffjungsan.FItemList(i).FGroupid %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FTW_price,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FUW_price,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FOM_price,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FSM_price,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FCM_price,0) %></td>
        <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FET_price,0) %></td>

        <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftot_orgsellprice,0) %></td>
        <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftot_realsellprice,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftot_jungsanprice,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftotalcommission,0) %></td>
      	<td align="center">
      	    <%= orgsellmargin %> %
      	</td>
      	<td align="center">
      	    <% if orgsellmargin<>realsellmargin then %>
      	    <font color="blue"><%= realsellmargin %></font> %
      	    <% else %>
      	    <%= realsellmargin %> %
      	    <% end if %>
      	</td>
      	<td ><a href="javascript:PopStateChange('<%= ooffjungsan.FItemList(i).Fidx %>');"><font color="<%= ooffjungsan.FItemList(i).GetStateColor %>"><%= ooffjungsan.FItemList(i).GetStateName %></font></a></td>
      	<td ><acronym title="<%= ooffjungsan.FItemList(i).Ftaxinputdate %>"><%= ooffjungsan.FItemList(i).Ftaxregdate %></acronym></td>
      	<td ><%= ooffjungsan.FItemList(i).Fipkumdate %></td>
        <td ><%= ooffjungsan.FItemList(i).Fjungsan_gubun %></td>
      	<td >
      	<% if ooffjungsan.FItemList(i).IsEditenable then %>
      	    <a href="javascript:DelMaster('<%= ooffjungsan.FItemList(i).Fidx %>');"><img src="/images/icon_delete2.gif" border="0" width="20"></a>
      	<% else %>
      	    <% if Not IsNULL(ooffjungsan.FItemList(i).Fneotaxno) then %>
      	        <% if (ooffjungsan.FItemList(i).Fbillsitecode="B") then %>
      	        <img src="/images/icon_print02.gif" width="14" height="14" border=0 onclick="window.open('http://www.bill36524.com/popupBillTax.jsp?NO_TAX=<%= ooffjungsan.FItemList(i).Fneotaxno %>&NO_BIZ_NO=2118700620')" style="cursor:hand">
      	        <% else %>
      	        <%= ooffjungsan.FItemList(i).Fbillsitecode %>
      	        <% end if %>
      	    <% end if %>
      	<% end if %>
      	<a href="/admin/upchejungsan/monthjungsanAdm.asp?makerid=<%= ooffjungsan.FItemList(i).Fmakerid %>&yyyy1=<%= LEFT(ooffjungsan.FItemList(i).Fyyyymm,4) %>&mm1=<%= right(ooffjungsan.FItemList(i).Fyyyymm,2) %>" target="_blank">POP</a>
     	</td>
    </tr>
    <% next %>
    <% else %>
    <tr bgcolor="#FFFFFF">
		<td colspan="25" align="center">[ 검색결과가 없습니다. ]</td>
	</tr>
	<% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center" bgcolor="F4F4F4">
			<% if ooffjungsan.HasPreScroll then %>
				<a href="javascript:NextPage('<%= ooffjungsan.StarScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + ooffjungsan.StarScrollPage to ooffjungsan.FScrollCount + ooffjungsan.StarScrollPage - 1 %>
				<% if i>ooffjungsan.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if ooffjungsan.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set ooffjungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
