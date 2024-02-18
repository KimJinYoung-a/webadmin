<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  상품 쿠폰 관리
' History : 2007.04.08 서동석 생성
'			2022.02.17 한용민 수정(검색조건추가. 디자인 신규버전으로 리뉴얼)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/newitemcouponcls.asp" -->

<%
dim oitemcoupon, page, research, iSerachType, sSearchTxt, selDate, sSdate, sEdate, onlyvalid, couponGubun, itemcoupontype
dim cpnvalue, i
	research    = requestCheckVar(request("research"),9)
	page        = requestCheckVar(request("page"),9)
	iSerachType = requestCheckVar(request("iSerachType"),9)
	sSearchTxt  = requestCheckVar(request("sSearchTxt"),32)
	onlyvalid   = requestCheckVar(request("onlyvalid"),9)
	selDate     = requestCheckVar(request("selDate"),10)
	sSdate      = requestCheckVar(request("sSdate"),10)
	sEdate      = requestCheckVar(request("sEdate"),10)
	couponGubun = requestCheckVar(request("couponGubun"),10)
	itemcoupontype = requestCheckVar(request("itemcoupontype"),10)
	cpnvalue	= requestCheckVar(request("cpnvalue"),10)

cpnvalue = replace(cpnvalue,"원","")
cpnvalue = replace(cpnvalue,"%","")
cpnvalue = Trim(replace(cpnvalue,",",""))

if Not(IsNumeric(cpnvalue)) then cpnvalue=""
''if (itemcoupontype="") then cpnvalue=""

if page="" then page=1
if research="" then onlyvalid="on"
if research="" and couponGubun="" then couponGubun="C"
    
set oitemcoupon = new CItemCouponMaster
	oitemcoupon.FPageSize=30
	oitemcoupon.FCurrPage = page
	oitemcoupon.FRectOnlyValid = onlyvalid
	oitemcoupon.FRectSearchType = iSerachType
	oitemcoupon.FRectSearchTxt = sSearchTxt
	oitemcoupon.FRectSearchDate = selDate
	oitemcoupon.FRectStartDate = sSdate
	oitemcoupon.FRectEndDate   = sEdate
	oitemcoupon.FRectCouponGubun = couponGubun
	oitemcoupon.FRectitemcoupontype = itemcoupontype
	oitemcoupon.FRectItemCouponValue = cpnvalue
	oitemcoupon.GetItemCouponMasterList

%>
<script type='text/javascript'>

function NextPage(page){
    var frm = document.frmSearch;
    frm.page.value = page;
    frm.submit();
}

function RegItemCoupon(){
	var popwin = window.open('itemcouponmasterreg.asp','RegItemCoupon','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function EditItemCoupon(itemcouponidx){
	var popwin = window.open('itemcouponmasterreg.asp?itemcouponidx=' + itemcouponidx,'EditItemCoupon','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function EditItemCouponItemMulti(){
    var popwin = window.open('itemcouponitemlisteidtMulti.asp'  ,'EditItemCouponItemMulti','width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function NvItemCouponExcept(){
	var popwin = window.open('/admin/etc/naverEp/exceptNvCpn.asp'  ,'NvItemCouponExcept','width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function EditCouponItemList(itemcouponidx){
	var popwin = window.open('itemcouponitemlisteidt.asp?itemcouponidx=' + itemcouponidx,'EditCouponItemList','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=350, height=350');
	winCal.focus();
}

function trim(value) {
	return value.replace(/^\s+|\s+$/g,"");
}

function isUInt(val) {
	var re = /^[0-9]+$/;
	return re.test(val);
}

function SubmitFrm(frm) {
	if ((frm.iSerachType.value == "1") || (frm.iSerachType.value == "2")) {
		if (frm.sSearchTxt.value*0 != 0) {
			alert('쿠폰코드/이벤트코드는 숫자만 가능합니다.');
			return;
		}
	}

	frm.sSearchTxt.value = trim(frm.sSearchTxt.value);

	if (frm.iSerachType.value == "1") {
		if (isUInt(frm.sSearchTxt.value) != true) {
			//?????
			//alert("쿠폰코드는 숫자만 가능합니다.");
			//return;
		}
	}

	frm.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frmSearch" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* <select name="iSerachType">
		<option value="1" <%IF Cstr(iSerachType) = "1" THEN%>selected<%END IF%>>쿠폰코드</option>
		<option value="3" <%IF Cstr(iSerachType) = "3" THEN%>selected<%END IF%>>쿠폰명</option>
		<option value="4" <%IF Cstr(iSerachType) = "4" THEN%>selected<%END IF%>>쿠폰설명</option>
		<option value="2" <%IF Cstr(iSerachType) = "2" THEN%>selected<%END IF%>>이벤트코드</option>
		</select>
		<input type="text" name="sSearchTxt" value="<%=sSearchTxt%>" size="10" maxlength="10">
		&nbsp;
		* <select name="selDate">
		<option value="S" <%if Cstr(selDate) = "S" THEN %>selected<%END IF%>>시작일 기준</option>
		<option value="E" <%if Cstr(selDate) = "E" THEN %>selected<%END IF%>>종료일 기준</option>
		<option value="R" <%if Cstr(selDate) = "R" THEN %>selected<%END IF%>>수정일 기준</option>
		</select>
		<input type="text" size="10" name="sSdate" value="<%=sSdate%>" onClick="jsPopCal('sSdate');" style="cursor:hand;">
		~ <input type="text" size="10" name="sEdate" value="<%=sEdate%>" onClick="jsPopCal('sEdate');"  style="cursor:hand;">
		&nbsp;
		<input type="checkbox" name="onlyvalid" <% if onlyvalid="on" then response.write "checked" %> >진행중인쿠폰 만 보기
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="SubmitFrm(document.frmSearch);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 쿠폰구분 
		<select name="couponGubun">
			<option value="" <%=CHKIIF(couponGubun="","selected","") %> >전체
			<option value="C" <%=CHKIIF(couponGubun="C","selected","") %> >일반
			<option value="V" <%=CHKIIF(couponGubun="V","selected","") %> >네이버전용쿠폰
			<option value="P" <%=CHKIIF(couponGubun="P","selected","") %> >지정인발급
			<option value="T" <%=CHKIIF(couponGubun="T","selected","") %> >타겟(E-mail특가)
		</select>
		&nbsp;
		* 할인구분 
		<select name="itemcoupontype">
			<option value="" <%=CHKIIF(itemcoupontype="","selected","") %> >전체
			<option value="1" <%=CHKIIF(itemcoupontype="1","selected","") %> >%
			<option value="2" <%=CHKIIF(itemcoupontype="2","selected","") %> >금액
			<option value="3" <%=CHKIIF(itemcoupontype="3","selected","") %> >배송료
		</select>
		&nbsp;
		* 할인값(%, 원)
		<input type="text" name="cpnvalue" value="<%=cpnvalue%>" size="7" maxlength="10" style="text-align:right">
	</td>
</tr>
</table>
<!-- 검색 끝 -->

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="신규 상품 쿠폰등록" onclick="RegItemCoupon();">
	</td>
	<td align="right">
		<input type="button" class="button" value="네이버전용쿠폰제외관리" onclick="NvItemCouponExcept();">
		&nbsp;
		<input type="button" class="button" value="등록 상품관리" onclick="EditItemCouponItemMulti();">
	</td>
</tr>
</table>
</form>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="14" align="left">
		검색건수 : <%= oitemcoupon.FTotalCount %> 건 Page : <%= page %>/<%= oitemcoupon.FTotalPage %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" width="50">쿠폰번호</td>
	<td align="center" width="80">쿠폰구분</td>
	<td align="center" width="70">이벤트코드<br>(그룹코드)</td>
	<td >쿠폰명</td>
	<td >쿠폰설명</td>
	<td align="center" width="100">할인금액</td>
	<td align="center" width="60">대상상품</td>
	<td align="center" width="100">시작일</td>
	<td align="center" width="100">종료일</td>
	<td align="center" width="70">상태</td>
	<td align="center" width="120">기본<br>마진구분</td>
	<td align="center" width="80">등록일</td>
	<td align="center" width="100">수정일</td>
	<td align="center" width="40">통계</td>
</tr>
<% if oitemcoupon.FResultCount>0 then %>
<% for i=0 to oitemcoupon.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td><%= oitemcoupon.FItemList(i).Fitemcouponidx %></td>
	<td><font color="<%= oitemcoupon.FItemList(i).getCouponGubunColor %>"><%= oitemcoupon.FItemList(i).getCouponGubunName %></font></td>
	<td>
		<%= oitemcoupon.FItemList(i).Fevt_code %>
		<% if Not IsNULL(oitemcoupon.FItemList(i).Fevtgroup_code) then %>
		(<%= oitemcoupon.FItemList(i).Fevtgroup_code %>)
		<% end if %>
	</td>
	<td><a href="javascript:EditItemCoupon('<%= oitemcoupon.FItemList(i).Fitemcouponidx %>')"><%= replace(oitemcoupon.FItemList(i).Fitemcouponname,"4월 정기세일","<strong>4월 정기세일</strong>") %></a></td>
	<td><%= oitemcoupon.FItemList(i).Fitemcouponexplain %></td>
	<td><%= oitemcoupon.FItemList(i).GetDiscountStr %></td>
	<td><a href="javascript:EditCouponItemList('<%= oitemcoupon.FItemList(i).Fitemcouponidx %>');"><%= oitemcoupon.FItemList(i).Fapplyitemcount %> 건</a></td>
	<td><%= ChkIIF(Right(oitemcoupon.FItemList(i).Fitemcouponstartdate,8)="00:00:00",Left(oitemcoupon.FItemList(i).Fitemcouponstartdate,10),oitemcoupon.FItemList(i).Fitemcouponstartdate) %></td>
	<td><%= ChkIIF(Right(oitemcoupon.FItemList(i).Fitemcouponexpiredate,8)="23:59:59",Left(oitemcoupon.FItemList(i).Fitemcouponexpiredate,10),oitemcoupon.FItemList(i).Fitemcouponexpiredate) %></td>
	<td><font color="<%= oitemcoupon.FItemList(i).GetOpenStateColor %>"><%= oitemcoupon.FItemList(i).GetOpenStateName %></font></td>
	<td><%= oitemcoupon.FItemList(i).GetMargintypeName %></td>
	<td><%= Left(oitemcoupon.FItemList(i).FRegDate,10) %></td>
	<td><%=oitemcoupon.FItemList(i).FlastupDt%></td>
	<td>
		<% if oitemcoupon.FItemList(i).Fopenstate>="7" then %>
		<a href="<%=stsAdmURL%>/admin/dataanalysis/report/simpleQry.asp?menupos=4116&qryidx=221&itmCpnIdx=<%=oitemcoupon.FItemList(i).Fitemcouponidx%>"><img src="/images/documents_icon.png" /></a>
		<% end if %>
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="14" align="center">
	<% if oitemcoupon.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oitemcoupon.StarScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oitemcoupon.StarScrollPage to oitemcoupon.FScrollCount + oitemcoupon.StarScrollPage - 1 %>
		<% if i>oitemcoupon.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oitemcoupon.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>')">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set oitemcoupon = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
