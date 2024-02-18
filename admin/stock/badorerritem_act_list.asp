<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 불량오차상품관리
' History : 이상구 생성
'           2021.04.06 한용민 수정(상품구분 일부코드 누락 오류 수정)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%
dim makerid,mode, searchtype, purchasetype, mwdiv, sellyn, onlyisusing, makeruseyn, itemgubun
dim datetype, centermwdiv, monthlymwdiv, yyyy, mm, osummarystock, i, BadOrErrText
	makerid 		= requestcheckvar(request("makerid"),32)
	mode 			= requestcheckvar(request("mode"),32)
	searchtype 		= requestcheckvar(request("searchtype"),3)
	purchasetype 	= requestcheckvar(request("purchasetype"),1)
	mwdiv 			= requestcheckvar(request("mwdiv"),1)
	sellyn 			= requestcheckvar(request("sellyn"),1)
	onlyisusing 	= requestcheckvar(request("onlyisusing"),1)
	makeruseyn	 	= requestcheckvar(request("makeruseyn"),1)
	itemgubun 		= requestcheckvar(request("itemgubun"),3)
	datetype 		= requestcheckvar(request("datetype"),8)
	yyyy 			= requestcheckvar(request("yyyy1"),4)
	mm 				= requestcheckvar(request("mm1"),2)
	centermwdiv		= requestcheckvar(request("centermwdiv"),1)
	monthlymwdiv	= requestcheckvar(request("monthlymwdiv"),1)

if (searchtype = "") then
	searchtype = "bad"
	'datetype = "curr"
	yyyy = Left(now(),4)
	mm   = mid(now(),6,2)
end if

'if (itemgubun = "") then
'	itemgubun = "10"
'end if
datetype = "yyyymm"
' 현재월일경우
if yyyy = Left(now(),4) and mm = mid(now(),6,2) then
	datetype = "curr"
end if

set osummarystock = new CSummaryItemStock
	osummarystock.FRectmakerid = makerid
	osummarystock.FRectSearchType = searchtype
	osummarystock.FRectDatetype   = datetype
	osummarystock.FRectYYYYMM = yyyy+"-"+mm

	'if (datetype = "yyyymm") then
	'	osummarystock.FRectMWDiv = monthlymwdiv
	'else
	'	osummarystock.FRectMWDiv = mwdiv
	'end if
	'osummarystock.FRectlastmwdiv = mwdiv
	osummarystock.FRectMWDiv = mwdiv
	osummarystock.FRectlastmwdiv = monthlymwdiv
	osummarystock.FRectCenterMWDiv = centermwdiv
	osummarystock.FRectSellYN = sellyn
	osummarystock.FRectOnlyIsUsing = onlyisusing
	osummarystock.FRectItemGubun = itemgubun
	osummarystock.FRectPurchaseType = purchasetype
	osummarystock.FRectMakerUseYN = makeruseyn

	if (makerid<>"") then
		osummarystock.FPageSize=500                 ''추가 2016/08/04 class에 막혀있는듯.
		osummarystock.GetBadOrErrItemListByBrand
	else
		osummarystock.GetBadOrErrItemListByBrandGroup
	end if

if (searchtype="bad") then
    BadOrErrText = "불량"
else
    BadOrErrText = "오차등록"
end if

%>
<script type='text/javascript'>

function PopBadOrErrItemReInput(makerid, acttype) {
	var popwin = window.open('/common/pop_badorerritem_re_input.asp?datetype=<%= datetype %>&yyyy1=<%= yyyy %>&mm1=<%= mm %>&makerid=' + makerid + '&searchtype=<%= searchtype %>&acttype=' + acttype + '&mwdiv=<%= mwdiv %>&itemgubun=<%= itemgubun %>&sellyn=<%= sellyn %>&onlyisusing=<%= onlyisusing %>&makeruseyn=<%= makeruseyn %>&purchasetype=<%=purchasetype%>&centermwdiv=<%= centermwdiv %>&monthlymwdiv=<%= monthlymwdiv %>','PopBadOrErrItemReInput','width=1280,height=800,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function SubmitSearchByBrandNew(makerid, mwdiv, itemgubun) {
	var searchitemgubun = "<%= itemgubun %>";

	if ((itemgubun == "") && (searchitemgubun != "")) {
		itemgubun = searchitemgubun;
	}

	var popwin = window.open('?datetype=<%= datetype %>&yyyy1=<%= yyyy %>&mm1=<%= mm %>&searchtype=<%= searchtype %>&purchasetype=<%= purchasetype %>&onlyisusing=<%= onlyisusing %>&sellyn=<%= sellyn %>&mwdiv=' + mwdiv + '&makerid=' + makerid + '&itemgubun=' + itemgubun + '&centermwdiv=<%= centermwdiv %>' + '&monthlymwdiv=<%= monthlymwdiv %>','SubmitSearchByBrandNew','width=1100,height=600,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'PopItemSellEdit','width=500,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popXL(searchtype, purchasetype, mwdiv, sellyn, onlyisusing, makeruseyn, itemgubun, datetype, yyyy, mm, centermwdiv, monthlymwdiv) {
	var popwin = window.open("/admin/stock/badorerritem_xl_download.asp?searchtype=" + searchtype + "&purchasetype=" + purchasetype + "&mwdiv=" + mwdiv + "&sellyn=" + sellyn + "&onlyisusing=" + onlyisusing + "&makeruseyn=" + makeruseyn + "&itemgubun=" + itemgubun + "&datetype=" + datetype + "&yyyy1=" + yyyy + "&mm1=" + mm + "&centermwdiv=" + centermwdiv + "&monthlymwdiv=" + monthlymwdiv,"popXL","width=300,height=200 scrollbars=yes resizable=yes");
	popwin.focus();
}

function ChangePage(v) {
	var frm = document.frm;

	frm.submit();
}

function jsSetBrandAll() {
    var frm = document.frm;
    frm.makerid.value = 'all';
    document.frm.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			기준월 :
			<% ' 이문재 이사님 요청으로 제외시킴. 특정월말 마지막월 기준이 현재 년월 기준과 같은거 아님?	' 2021.04.15 한용민 %>
			<!--<input type="radio" name="datetype" value="curr" <% if (datetype = "curr") then %>checked<% end if %>> 현재기준
			<input type="radio" name="datetype" value="yyyymm" <% if (datetype = "yyyymm") then %>checked<% end if %>> 특정월말기준 -->
			<% Call DrawYMBox(yyyy, mm) %>
			&nbsp;
			<input type="radio" name="searchtype" value="bad" <% if (searchtype = "bad") then %>checked<% end if %>> 불량상품
			<input type="radio" name="searchtype" value="err" <% if (searchtype = "err") then %>checked<% end if %>> 오차등록상품
			&nbsp;
			브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
            <input type="button" class="button" value="전체브랜드" onClick="jsSetBrandAll()">
		</td>
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<b>브랜드 정보</b>
			&nbsp;
			&nbsp;
			사용여부 :
			<select class="select" name="makeruseyn">
				<option value="">-선택-</option>n
				<option value="Y" <% if (makeruseyn = "Y") then %>selected<% end if %> >사용함</option>
				<option value="N" <% if (makeruseyn = "N") then %>selected<% end if %> >사용않함</option>
			</select>
			&nbsp;
			구매유형 : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<b>상품 정보</b>
			&nbsp;
			&nbsp;
			상품구분 :
			<select class="select" name="itemgubun">
				<option value="">-선택-</option><% '이문재 이사님 요청으로 전체 추가 %>
				<option value="10" <% if (itemgubun = "10") then %>selected<% end if %> >온상품(10)</option>
				<option value="OFF" <% if (itemgubun = "OFF") then %>selected<% end if %> >오프전체</option>
				<option value="55" <% if (itemgubun = "55") then %>selected<% end if %> >오프(55)</option>
				<option value="70" <% if (itemgubun = "70") then %>selected<% end if %> >오프(70)</option>
				<option value="75" <% if (itemgubun = "75") then %>selected<% end if %> >오프(75)</option>
				<option value="80" <% if (itemgubun = "80") then %>selected<% end if %> >오프(80)</option>
				<option value="85" <% if (itemgubun = "85") then %>selected<% end if %> >오프(85)</option>
				<option value="90" <% if (itemgubun = "90") then %>selected<% end if %> >오프(90)</option>
			</select>
			&nbsp;
            <%'= CHKIIF(datetype<>"yyyymm", "ON매입구분(현재)", "<del>ON매입구분(현재)</del>") %>ON매입구분(현재) :
			<select class="select" name="mwdiv">
				<option value="">-선택-</option>
				<option value="M" <% if (mwdiv = "M") then %>selected<% end if %> >매입</option>
				<option value="W" <% if (mwdiv = "W") then %>selected<% end if %> >특정</option>
				<option value="U" <% if (mwdiv = "U") then %>selected<% end if %> >업체</option>
				<option value="Z" <% if (mwdiv = "Z") then %>selected<% end if %> >미지정</option>
			</select>
			&nbsp;
            센터매입구분(현재) :
     		<select class="select" name="centermwdiv">
				<option value="">선택</option>
				<option value="M" <%= CHKIIF(centermwdiv="M","selected","")%> >매입</option>
				<option value="W" <%= CHKIIF(centermwdiv="W","selected","")%> >위탁</option>
				<option value="X" <%= CHKIIF(centermwdiv="X","selected","")%> >미지정</option>
			</select>
     		&nbsp;
            <%'= CHKIIF(datetype="yyyymm", "매입구분(월별)", "<del>매입구분(월별)</del>") %>매입구분(재고) :
     		<select class="select" name="monthlymwdiv">
				<option value="">선택</option>
				<option value="M" <%= CHKIIF(monthlymwdiv="M","selected","")%> >매입</option>
				<option value="W" <%= CHKIIF(monthlymwdiv="W","selected","")%> >위탁</option>
				<option value="X" <%= CHKIIF(monthlymwdiv="X","selected","")%> >미지정</option>
			</select>
			&nbsp;
			판매여부(현재) :
			<select class="select" name="sellyn">
				<option value="">-선택-</option>
				<option value="Y" <% if (sellyn = "Y") then %>selected<% end if %> >판매함</option>
				<option value="N" <% if (sellyn = "N") then %>selected<% end if %> >판매않함</option>
			</select>
            &nbsp;
			사용여부(현재) :
			<select class="select" name="onlyisusing">
				<option value="">-선택-</option>
				<option value="Y" <% if (onlyisusing = "Y") then %>selected<% end if %> >사용함</option>
				<option value="N" <% if (onlyisusing = "N") then %>selected<% end if %> >사용않함</option>
			</select>
		</td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>

* 브랜드 및 상품정보는 <font color="red">현재정보를 기준</font>으로 합니다.(특정월 브랜드정보 및 상품정보 고려안함)

<% if makerid<>"" then %>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<% if (searchtype = "bad") then %>
				<input type="button" class="button" value="반품" onclick="PopBadOrErrItemReInput('<%= makerid %>', 'actreturn')" border="0">
				&nbsp;
			<% end if %>
        	<input type="button" class="button" value="로스출고" onclick="PopBadOrErrItemReInput('<%= makerid %>', 'actloss')" border="0">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			검색결과 : <b><%= osummarystock.FTotalCount %></b>
			<% if (osummarystock.FResultCount>=osummarystock.FPageSize) then %>최대 <%=osummarystock.FPageSize%> 건 표시<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>브랜드ID</td>
		<td width="50">이미지</td>
		<td width="40">재고<br>매입<br>구분</td>
		<td width="40">ON<br>매입<br>구분</td>
		<td width="40">센터<br>매입<br>구분</td>
		<td width="30">상품<br>구분</td>
		<td width="50">상품코드</td>
		<td width="40">옵션</td>
		<td>상품명<br><font color="blue">[옵션명]</font></td>

		<td width="50">소비자가</td>
		<td width="50">매입가</td>
		<td width="30">판매<br>여부</td>
		<td width="30">사용<br>여부</td>
		<td width="60"><%= BadOrErrText %><br>수량</td>
		<td width="70">매입가합</td>
		<td width="60">실사<br>유효재고</td>
    </tr>
	<% if osummarystock.FResultCount>0 then %>
	<% for i=0 to osummarystock.FResultCount - 1 %>
	<% if (osummarystock.FItemList(i).Fisusing = "Y") then %>
		<tr bgcolor="#FFFFFF" height="30">
	<% else %>
		<tr bgcolor="#BBBBBB" height="30">
	<% end if %>
    	<td><%= osummarystock.FItemList(i).Fmakerid %></td>
    	<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50" onError="this.src='http://webimage.10x10.co.kr/images/no_image.gif'" ></td>
		<td align="center" style="color:<%=GetMwDivColorCd(osummarystock.FItemList(i).flastmwdiv)%>;"><%= osummarystock.FItemList(i).flastmwdiv %></td>
    	<td align="center" style="color:<%=osummarystock.FItemList(i).GetMwDivColor%>;"><%= osummarystock.FItemList(i).Fmwdiv %></td>
		<td align="center" style="color:<%=GetMwDivColorCd(osummarystock.FItemList(i).Fcentermwdiv)%>;"><%= osummarystock.FItemList(i).Fcentermwdiv %></td>
    	<td align="center"><%= osummarystock.FItemList(i).FItemgubun %></td>
		<td align="center"><a href="javascript:PopItemSellEdit('<%= osummarystock.FItemList(i).FItemid %>');"><%= osummarystock.FItemList(i).FItemid %></a></td>
		<td align="center"><%= osummarystock.FItemList(i).FItemoption %></td>
		<td align="left"><a href="/admin/stock/itemcurrentstock.asp?itemgubun=<%= osummarystock.FItemList(i).FItemgubun %>&itemid=<%= osummarystock.FItemList(i).FItemid %>&itemoption=<%= osummarystock.FItemList(i).FItemoption %>" target=_blank ><%= osummarystock.FItemList(i).FItemname %></a><br><font color="blue">[<%= osummarystock.FItemList(i).FItemOptionName %>]</font></td>

		<td align="right"><%= FormatNumber(osummarystock.FItemList(i).Fsellcash,0) %></td>
		<td align="right"><%= formatnumber(osummarystock.FItemList(i).Fbuycash,0) %></td>
		<td align="center"><%= osummarystock.FItemList(i).Fsellyn %></td>
		<td align="center"><%= osummarystock.FItemList(i).Fisusing %></td>
		<td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fregitemno, 0) %></td>
		<td align="right"><%= formatnumber((osummarystock.FItemList(i).Fbuycash * osummarystock.FItemList(i).Fregitemno),0) %></td>
		<td align="center"><%= FormatNumber(osummarystock.FItemList(i).Frealstock, 0) %></td>
    </tr>
    <% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
</table>

<% else %>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="right">
			<% if (searchtype = "bad") then %>
			<input type="button" class="button" value="엑셀다운로드(불량)" onclick="popXL('bad', '<%= purchasetype %>', '<%= mwdiv %>', '<%= sellyn %>', '<%= onlyisusing %>', '<%= makeruseyn %>', '<%= itemgubun %>', '<%= datetype %>', '<%= yyyy %>', '<%= mm %>', '<%= centermwdiv %>', '<%= monthlymwdiv %>')">
			<% end if %>
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			검색결과 : <b><%= osummarystock.FResultCount %></b>
		</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td rowspan="3">브랜드</td>
		<td rowspan="3">브랜드명</td>
		<td rowspan="3">업체명</td>
		<td width="40" rowspan="3">브랜드<br>사용<br>여부</td>
		<td colspan="11"><%= BadOrErrText %>상품수량</td>
		<td rowspan="3">비고</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td colspan="4">10</td>
		<td colspan="3">90</td>
		<td colspan="3">기타</td>
		<td rowspan="2" width="80">소계</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="55">매입</td>
		<td width="55">위탁</td>
		<td width="55">업배</td>
		<td width="55">미지정</td>
		<td width="55">매입</td>
		<td width="55">위탁</td>
		<td width="55">미지정</td>
		<td width="55">매입</td>
		<td width="55">위탁</td>
		<td width="55">미지정</td>
	</tr>
	<% if osummarystock.FResultCount>0 then %>
	<% for i=0 to osummarystock.FResultCount-1 %>
	<% if (osummarystock.FItemList(i).Fuseyn = "Y") then %>
		<tr bgcolor="#FFFFFF" height="30">
	<% else %>
		<tr bgcolor="#BBBBBB" height="30">
	<% end if %>
	    <td><a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', '', '');"><%= osummarystock.FItemList(i).FMakerid %></a></td>
	    <td><%= osummarystock.FItemList(i).Fmakername %></td>
		<td align="left"><%= osummarystock.FItemList(i).Fcompany_name %></td>
		<td align="center"><%= osummarystock.FItemList(i).Fuseyn %></td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'M', '10');">
	    	<% if (osummarystock.FItemList(i).Fitem10M <> 0) then %><font color="red"><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).Fitem10M, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'W', '10');">
	    	<% if (osummarystock.FItemList(i).Fitem10W <> 0) then %><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).Fitem10W, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'U', '10');">
	    	<% if (osummarystock.FItemList(i).Fitem10U <> 0) then %><font color="green"><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).Fitem10U, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'Z', '10');">
	    	<% if (osummarystock.FItemList(i).Fitem10U <> 0) then %><font color="black"><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).Fitem10Z, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'M', '90');">
	    	<% if (osummarystock.FItemList(i).Fitem90M <> 0) then %><font color="red"><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).Fitem90M, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'W', '90');">
	    	<% if (osummarystock.FItemList(i).Fitem90W <> 0) then %><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).Fitem90W, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'Z', '90');">
	    	<% if (osummarystock.FItemList(i).Fitem90W <> 0) then %><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).Fitem90Z, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'M', 'OFF');">
	    	<% if (osummarystock.FItemList(i).FitemetcM <> 0) then %><font color="red"><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).FitemetcM, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'W', 'OFF');">
	    	<% if (osummarystock.FItemList(i).FitemetcW <> 0) then %><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).FitemetcW, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'Z', 'OFF');">
	    	<% if (osummarystock.FItemList(i).FitemetcW <> 0) then %><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).FitemetcZ, 0) %></a>
	    </td>
	    <td align="center">
	    	<% if ((osummarystock.FItemList(i).FOnCnt + osummarystock.FItemList(i).FOffCnt) <> 0) then %><b><% end if %>
	    	<%= FormatNumber((osummarystock.FItemList(i).FOnCnt + osummarystock.FItemList(i).FOffCnt), 0) %>
	    </td>
	    <td align="left">
	    	<% if (searchtype = "bad") then %>
				<input type="button" class="button" value="반품" onclick="PopBadOrErrItemReInput('<%= osummarystock.FItemList(i).FMakerid %>', 'actreturn')" border="0" <% if (osummarystock.FItemList(i).Fcompany_no = "211-87-00620" and Left(Now(), 7) <> "2022-07") then %>disabled<% end if %> >
				&nbsp;
				<input type="button" class="button" value="매장출고" onclick="PopBadOrErrItemReInput('<%= osummarystock.FItemList(i).FMakerid %>', 'actshopchulgo')" border="0">
				&nbsp;
			<% end if %>
			<% if (searchtype = "bad") then %>
        	<input type="button" class="button" value="폐기처리" onclick="PopBadOrErrItemReInput('<%= osummarystock.FItemList(i).FMakerid %>', 'actloss')" border="0">
			<% else %>
			<input type="button" class="button" value="로스출고" onclick="PopBadOrErrItemReInput('<%= osummarystock.FItemList(i).FMakerid %>', 'actloss')" border="0">
			<% end if %>
	    </td>
	</tr>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="16" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
</table>
<% end if %>

<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
