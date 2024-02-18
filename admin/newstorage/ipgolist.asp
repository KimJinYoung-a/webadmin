<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  물류 입고리스트
' History : 2007.01.01 이상구 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%
dim code, blinkcode, minusjumun, page,designer, statecd, onoffgubun, divcode, rackipgoyn, ipgocheck, yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim fromDate, toDate, vPurchaseType, searchType, searchText, totalsellcash,totalsuply, totalitemno, i
dim alinkcode, linkparam
	page = request("page")
	designer = request("designer")
	statecd = request("statecd")
	code = request("code")				' 입고 코드
	alinkcode = request("alinkcode")
	blinkcode = request("blinkcode")
	onoffgubun = request("onoffgubun")	' 온/오프 구분
	divcode = request("divcode")		' 매입 구분
	rackipgoyn = request("rackipgoyn")	'
	vPurchaseType = requestCheckVar(request("purchasetype"),3)
	searchType = request("searchType")
	searchText = request("searchText")
	minusjumun = request("minusjumun")

	'// 입고일 검색에 필요한 변수 대입
	ipgocheck = request("ipgocheck")
	yyyy1 = request("yyyy1")
	yyyy2 = request("yyyy2")
	mm1	  = request("mm1")
	mm2	  = request("mm2")
	dd1	  = request("dd1")
	dd2	  = request("dd2")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

if onoffgubun="" then onoffgubun="all"

if page="" then page=1

code = Trim(code)
blinkcode = Trim(blinkcode)

linkparam="&designer="&designer&"&searchType="&searchType&"&searchText="&searchText&"&onoffgubun="&onoffgubun&"&divcode="&divcode&"&rackipgoyn="&rackipgoyn
linkparam=linkparam & "&purchasetype="&vpurchasetype&"&minusjumun="&minusjumun&"&code="&"&alinkcode="&alinkcode&"&blinkcode="&"&ipgocheck="&ipgocheck

dim oipchul
set oipchul = new CIpChulStorage
	oipchul.FCurrPage = page
	oipchul.Fpagesize=50
	oipchul.FRectCode = code
	oipchul.FRectBLinkCode = blinkcode
	oipchul.FRectALinkCode = alinkcode
	oipchul.FRectDivcode = divcode
	oipchul.FRectRackipgoyn = rackipgoyn
	oipchul.FRectMinusOnly = minusjumun

	if ipgocheck="on" then
		oipchul.FRectExecuteDtStart = fromDate
		oipchul.FRectExecuteDtEnd   = toDate
	end if

	if code="" then
	oipchul.FRectCodeGubun = "ST"  ''입고
	oipchul.FRectSocID = designer
	oipchul.FRectOnOffGubun = onoffgubun
	end if

	oipchul.FRectSearchType = searchType
	oipchul.FRectSearchText = searchText

	oipchul.FRectBrandPurchaseType = vPurchaseType
	oipchul.GetIpChulgoList

totalsellcash = 0
totalsuply	  = 0
totalitemno = 0
%>

<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script src="/js/jquery.placeholder.min.js"></script>
<script type="text/javascript">

function PopUpcheBrandInfoEdit(v){
	window.open("/admin/member/popupchebrandinfo.asp?designer=" + v,"PopUpcheBrandInfoEdit","width=640,height=580,scrollbars=yes,resizabled=yes");
}

function IpgoInput(){
	location.href="/admin/newstorage/ipgoinput.asp?menupos=<%= menupos %>";
}

function popipgocheck(iidx){
	var popwin = window.open("poplimitcheckipgoNew.asp?idx=" + iidx ,"popipgoproc","width=900,height=550,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function PopIpgoSheet(v,itype){
	var popwin;
	popwin = window.open('popipgosheet.asp?idx=' + v + '&itype=' + itype,'popipgosheet','width=760,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function ExcelSheet(v,itype){
	window.open('popipgosheet.asp?idx=' + v + '&itype=' + itype + '&xl=on');
}

function EnDisabledDateBox(comp){
	document.frm.yyyy1.disabled = !comp.checked;
	document.frm.yyyy2.disabled = !comp.checked;
	document.frm.mm1.disabled = !comp.checked;
	document.frm.mm2.disabled = !comp.checked;
	document.frm.dd1.disabled = !comp.checked;
	document.frm.dd2.disabled = !comp.checked;
}

function NextPage(page){
	ClearPlaceHolder();
	document.frm.page.value = page;
	document.frm.submit();
}

function popXL(fromDate, toDate) {
	<% if ipgocheck<>"on" then %>
	alert("먼저 입고일을 지정하세요.");
	return;
	<% end if %>

	var popwin = window.open("/admin/newstorage/pop_ipgolist_xl_download.asp?fromDate=" + fromDate + "&toDate=" + toDate + "<%=linkparam%>","popXL","width=100,height=100 scrollbars=yes resizable=yes");
	popwin.focus();
}

function SubmitFrm(frm) {
	ClearPlaceHolder();
	if (frm.code.value.length > 0) {
		if (frm.code.value.substring(0,2).toUpperCase() != "ST") {
			alert("입고코드가 아닙니다.");
			return;
		}
	}

	frm.submit();
}

function ClearPlaceHolder() {
	var frm = document.frm;
	frm.code.value = $('#code').val();
	frm.blinkcode.value = $('#blinkcode').val();
}

function popOpenPPMaster(idx) {
	var popwin;

	popwin = window.open('/admin/newstorage/PurchasedProductModify.asp?menupos=9175&idx=' + idx ,'popOpenPPMaster','width=1400,height=768,scrollbars=yes,resizable=yes');
	popwin.focus();
}

$( document ).ready(function() {
    $('textarea').placeholder();
});

</script>

<style>
textarea:-webkit-input-placeholder {color:#acacac;}
textarea:-moz-placeholder {color:#acacac;}
textarea:-ms-input-placeholder {color:#acacac;}
.placeholder { color: #acacac; }
</style>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 브랜드 : <% drawSelectBoxDesignerwithName "designer", designer %>
		&nbsp;
		* 입고코드 :
		<textarea class="textarea" id="code" name="code" cols="12" rows="1" placeholder="최대50개"><%= code %></textarea>
		&nbsp;
		* 주문코드 :
		<textarea class="textarea" id="blinkcode" name="blinkcode" cols="12" rows="1" placeholder="최대50개"><%= blinkcode %></textarea>
		&nbsp;
		* 주문코드(개별) : <input type="text" class="text" name="alinkcode" value="<%= alinkcode %>" size="8" maxlength="8">
		&nbsp;
		<input type="checkbox" name="ipgocheck" <% if ipgocheck="on" then  response.write "checked" %> onclick="EnDisabledDateBox(this)">입고일
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="SubmitFrm(frm)">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 온오프구분:
		<input type="radio" name="onoffgubun" value="all" <% if onoffgubun="all" then response.write "checked" %> >전체
		<input type="radio" name="onoffgubun" value="on" <% if onoffgubun="on" then response.write "checked" %> >온라인
		<input type="radio" name="onoffgubun" value="off" <% if onoffgubun="off" then response.write "checked" %> >오프라인
		&nbsp;
		* 매입구분:
		<input type="radio" name="divcode" value="" <% if divcode="" then response.write "checked" %> >전체
		<input type="radio" name="divcode" value="001" <% if divcode="001" then response.write "checked" %> >매입
		<input type="radio" name="divcode" value="002" <% if divcode="002" then response.write "checked" %> >위탁
		<input type="radio" name="divcode" value="801" <% if divcode="801" then response.write "checked" %> >Off매입
		<input type="radio" name="divcode" value="802" <% if divcode="802" then response.write "checked" %> >Off위탁
		&nbsp;
		* 구매유형 : 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 검색조건 :
		<select class="select" name="searchType">
			<option value="" >전체</option>
			<option value="socname" <% if (searchType = "socname") then %>selected<% end if %> >업체명</option>
			<option value="socno" <% if (searchType = "socno") then %>selected<% end if %> >사업자번호</option>
		</select>
		<input type="text" class="text" name=searchText value="<%= searchText %>" size="15" maxlength="20">
		&nbsp;
		<input type="checkbox" name="minusjumun" <% if minusjumun="on" then response.write "checked" %> >마이너스주문만
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value=" 입고입력 " onclick="IpgoInput();">
	</td>
	<td align="right">
		<% if oipchul.FTotalCount > 0 then %>
			<input type="button" class="button" value=" 엑셀받기 " onclick="popXL('<%= fromDate %>', '<%= toDate %>');">
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="left">
		검색결과 : <b><%= oipchul.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oipchul.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=60>입고코드</td>
	<td width=60>주문코드</td>
	<td width="60">원가IDX</td>
	<td width=60>구매유형</td>
	<td>공급처ID</td>
	<td>공급처</td>
	<td width=50>처리자</td>
	<td width=70>예정일</td>
	<td width=70>입고일</td>
	<td width=60>소비자가</td>
	<td width=60>매입가</td>
	<td width=40>수량</td>
	<td width=50>구분</td>
	<td width=50>마진</td>
	<td width=60>입고처리</td>
	<td width=50>내역서</td>
</tr>
<% if oipchul.FResultCount >0 then %>
<% for i=0 to oipchul.FResultcount-1 %>
<%
totalsellcash = totalsellcash + oipchul.FItemList(i).Ftotalsellcash
totalsuply	  = totalsuply + oipchul.FItemList(i).Ftotalsuplycash
totalitemno	  = totalitemno + oipchul.FItemList(i).ftotalitemno
%>
<tr bgcolor="#FFFFFF" height=24>
	<td align=center><a href="ipgodetail.asp?idx=<%= oipchul.FItemList(i).Fid %>&opage=<%= page %>&menupos=<%=menupos%>"><%= oipchul.FItemList(i).Fcode %></a></td>
	<td align=center>
		<% if Not IsNull(oipchul.FItemList(i).Fblinkcode) then %>
		<a href="/admin/newstorage/orderlist.asp?menupos=537&baljucode=<%= oipchul.FItemList(i).Fblinkcode %>" target="_blank"><%= oipchul.FItemList(i).Fblinkcode %></a>
		<% elseif Not IsNull(oipchul.FItemList(i).Falinkcode) then %>
		<a href="/admin/fran/upchejumunlist.asp?menupos=530&baljucode=<%= oipchul.FItemList(i).Falinkcode %>" target="_blank"><%= oipchul.FItemList(i).Falinkcode %></a>
		<% end if %>
	</td>
	<td align="center">
		<% if (oipchul.FItemList(i).FppMasterIdx <> "" and not(isnull(oipchul.FItemList(i).FppMasterIdx))) then %>
			<a href="#" onclick="popOpenPPMaster(<%= oipchul.FItemList(i).FppMasterIdx %>); return false;"><%= oipchul.FItemList(i).FppMasterIdx %></a>
		<% end if %>
	</td>
	<td align=left><%= oipchul.FItemList(i).fpurchasetypename %></td>
	<td align=left><b><a href="javascript:PopUpcheBrandInfoEdit('<%= oipchul.FItemList(i).Fsocid %>');"><%= oipchul.FItemList(i).Fsocid %></a></b></td>
	<td align=left><%= oipchul.FItemList(i).Fsocname %></td>
	<td align=center><%= oipchul.FItemList(i).Fchargename %></td>
	<td align=center><font color="#777777"><%= Left(oipchul.FItemList(i).Fscheduledt,10) %></font></td>
	<td align=center><%= Left(oipchul.FItemList(i).Fexecutedt,10) %></td>
	<td align=right><font color="<%= oipchul.FItemList(i).GetMinusColor(oipchul.FItemList(i).Ftotalsellcash) %>"><%= FormatNumber(oipchul.FItemList(i).Ftotalsellcash,0) %></font></td>
	<td align=right><font color="<%= oipchul.FItemList(i).GetMinusColor(oipchul.FItemList(i).Ftotalsuplycash) %>"><%= FormatNumber(oipchul.FItemList(i).Ftotalsuplycash,0) %></font></td>
	<td align="right">
		<font color="<%= oipchul.FItemList(i).GetMinusColor(oipchul.FItemList(i).ftotalitemno) %>"><%= FormatNumber(oipchul.FItemList(i).ftotalitemno,0) %></font>
	</td>
	<td align=center><font color="<%= oipchul.FItemList(i).GetDivCodeColor %>"><%= oipchul.FItemList(i).GetDivCodeName %></font></td>
	<td align=right>
	<% if oipchul.FItemList(i).Ftotalsellcash<>0 then %>
	  <%= 100-CLng(oipchul.FItemList(i).Ftotalsuplycash/oipchul.FItemList(i).Ftotalsellcash*100*100)/100 %>%
	<% end if %>
	</td>
	<td align=center>
		<input type="button" class="button" value="입고처리" onClick="popipgocheck('<%= oipchul.FItemList(i).Fid %>')">
	</td>
	<td>
          <a href="javascript:PopIpgoSheet('<%= oipchul.FItemList(i).Fid %>','');"><img src="/images/iexplorer.gif" width=21 border=0></a> <a href="javascript:ExcelSheet('<%= oipchul.FItemList(i).Fid %>','');"><img src="/images/iexcel.gif" width=21 border=0></a>
    </td>
</tr>
<% next %>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align=center>총계</td>
	<td colspan=8></td>
	<td align=right><%= formatNumber(totalsellcash,0) %></td>
	<td align=right><%= formatNumber(totalsuply,0) %></td>
	<td align="right">
		<%= formatNumber(totalitemno,0) %>
	</td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>

<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=20 align=center>[ 검색결과가 없습니다. ]</td>
</tr>
<% end if %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
		<% if oipchul.HasPreScroll then %>
    		<a href="javascript:NextPage('<%= oipchul.StartScrollPage-1 %>')">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oipchul.StartScrollPage to oipchul.FScrollCount + oipchul.StartScrollPage - 1 %>
    		<% if i>oipchul.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oipchul.HasNextScroll then %>
    		<a href="javascript:NextPage('<%= i %>')">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
	</td>
</tr>
</table>


<%
set oipchul = Nothing
%>

<script type="text/javascript">
	EnDisabledDateBox(document.frm.ipgocheck);
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
