<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'#######################################################
' Description : 보너스쿠폰 
' History	:  최초 생성자 모름
'              2017.07.05 한용민 수정
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/newcouponcls.asp" -->
<%
dim ocoupon, page, lp
dim cusUserid, regUserid, masteridx, couponname, coupontype, usingyn, orderserial, chkOld, valSiteType, targetCpnType, i
	cusUserid = request("cusUserid")
	regUserid = request("regUserid")
	masteridx = request("masteridx")
	couponname = request("couponname")
	coupontype = request("coupontype")
	usingyn = request("usingyn")
	orderserial = request("orderserial")
	chkOld = request("chkOld")
	valSiteType = request("valSiteType")
	targetCpnType = request("targetCpnType")
	page = request("page")

if page="" then page=1
if valSiteType="" then valSiteType="T"

set ocoupon = new CCouponMaster
	ocoupon.FPageSize=60
	ocoupon.FCurrPage = page
	ocoupon.FrectCusUserid = cusUserid
	ocoupon.FrectRegUserid = regUserid
	ocoupon.FrectCouponname = couponname
	ocoupon.FRectIdx = masteridx
	ocoupon.FrectCoupontype = coupontype
	ocoupon.FrectUsingyn = usingyn
	ocoupon.FrectOrderserial = orderserial
	ocoupon.FrectChkOld = chkOld
	ocoupon.FrectSiteType = valSiteType
	ocoupon.FrectTargetCpnType = targetCpnType
	ocoupon.GetEventCouponList

%>

<script type="text/javascript">

<!--
	function goPage(pg) {
		frm.page.value=pg;
		frm.submit();
	}

	function newCoupon() {
		location.href="event_coupon_edit.asp";
	}

	function msgOldDB(chk) {
		if(chk.checked) {
			alert("3개월 이전 쿠폰 검색은 DB에 많은 부하를 줄 수 있고 검색시간이 오래걸립니다.\n꼭 필요한 경우에만 체크해주십시오.");
		}
	}

	function chgUsing(fm) {
		if(fm.value=='N') {
			frm.orderserial.disabled=true;
			frm.orderserial.className="text_ro";
		} else {
			frm.orderserial.disabled=false;
			frm.orderserial.className="text";
		}
	}
//-->

</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="1">
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="#EEEEEE">검색<br>조건</td>
	<td align="left">
		고객ID : <input type="text" class="text" name="cusUserid" value="<%=cusUserid%>" size="12" maxlength="32"> &nbsp;
		발급자ID : <input type="text" class="text" name="regUserid" value="<%=regUserid%>" size="12" maxlength="32"> &nbsp;
		쿠폰번호 : <input type="text" class="text" name="masteridx" value="<%=masteridx%>" size="10" maxlength="10"> &nbsp;
		쿠폰명 : <input type="text" class="text" name="couponname" value="<%=couponname%>" size="20" maxlength="20"> &nbsp;
		/ <label><input type="checkbox" name="chkOld" value="Y" <%=chkIIF(chkOld="Y","checked","")%> onclick="msgOldDB(this)"> 3개월 이전 검색</label>
		<br>
     	사용처 :
		<select class="select" name="valSiteType">
		<option value="T">텐바이텐</option>
		<option value="F">더핑거스</option>
		</select> &nbsp; &nbsp;
     	쿠폰종류 :&nbsp;
		<select class="select" name="coupontype">
		<option value="">전체</option>
		<option value="1">%할인</option>
		<option value="2">원할인</option>
		<option value="3">무료배송</option>
		</select> &nbsp; &nbsp;
     	쿠폰구분 :&nbsp;
		<select class="select" name="targetCpnType">
		<option value="">전체</option>
		<option value="B">브랜드</option>
		<option value="C">카테고리</option>
		</select> &nbsp; &nbsp;

     	쿠폰사용여부 :
		<select class="select" name="usingyn" onchange="chgUsing(this)">
		<option value="">전체</option>
		<option value="Y">사용함</option>
		<option value="N">사용안함</option>
		</select>&nbsp; &nbsp;
		주문번호 : <input type="text" class="<%=chkIIF(usingyn="N","text_ro","text")%>" name="orderserial" value="<%=orderserial%>" size="18" maxlength="16"> &nbsp;
		<script language="javascript">
		document.frm.valSiteType.value="<%=valSiteType%>";
		document.frm.coupontype.value="<%=coupontype%>";
		document.frm.targetCpnType.value="<%=targetCpnType%>";
		document.frm.usingyn.value="<%=usingyn%>";
		</script>
	</td>
	<td width="50" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</table>
<!-- 검색 끝 -->
</form>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 0 0;">
<tr>
	<td align="right"><input type="button" class="button" value="신규등록" onClick="newCoupon()"></td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#B2B2B2" class="a">
<% if ocoupon.FResultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="13">
			검색결과 : <b><%= formatNumber(ocoupon.FTotalCount,0) %></b>
			&nbsp;
			페이지 : <b><%= formatNumber(page,0) %>/ <%= formatNumber(ocoupon.FTotalPage,0) %></b>
		</td>
	</tr>
	<tr bgcolor="#E6E6E6">
		<td width="50" align="center">IDx</td>
		<td align="center">아이디</td>
		<td align="center">쿠폰코드</td>
		<td align="center">보너스내용</td>
		<td width="80" align="center">사용주문번호</td>
		<td width="150" align="center">사용 혜택</td>
		<td width="50" align="center">최소구매 금액</td>
		<td width="50" align="center">최대할인 금액</td>
		<td align="center">쿠폰구분</td>
		<td width="150" align="center">유효기간</td>
		<td width="80" align="center">등록일</td>
		<td width="30" align="center">사용 여부</td>
		<td width="100" align="center">발급자</td>
	</tr>
	<% for i=0 to ocoupon.FResultCount - 1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%= ocoupon.FItemList(i).FIdx %></td>
		<td align="center"><%= printUserId(ocoupon.FItemList(i).Fuserid, 2, "*") %></td>
		<td align="center"><%= ocoupon.FItemList(i).FmasterIdx %></td>
		<td><%= ocoupon.FItemList(i).Fcouponname %></td>
		<td align="center"><%= ocoupon.FItemList(i).forderserial %></td>
		<td align="center"><%= ocoupon.FItemList(i).getCouponTypeStr %></td>
		<td align="center"><%= FormatNumber(ocoupon.FItemList(i).Fminbuyprice,0) %></td>
		<td align="center"><%= FormatNumber(chkIIF(ocoupon.FItemList(i).FmxCpnDiscount="" or isNull(ocoupon.FItemList(i).FmxCpnDiscount),"0",ocoupon.FItemList(i).FmxCpnDiscount),0) %></td>
		<td align="center"><%= ocoupon.FItemList(i).getCouponTypeNameStr%></td>
		<td align="center"><%= ocoupon.FItemList(i).getAvailDateStr %></td>
		<td align="center"><%= Formatdatetime(ocoupon.FItemList(i).FRegDate,2) %></td>
		<td align="center"><%= ocoupon.FItemList(i).FIsUsing %></td>
		<td align="center"><%= ocoupon.FItemList(i).Freguserid %></td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="13" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="13" align="center">
		<% if ocoupon.HasPreScroll then %>
			<a href="javascript:goPage(<%= ocoupon.StartScrollPage-1 %>)">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
	
		<% for lp=0 + ocoupon.StartScrollPage to ocoupon.FScrollCount + ocoupon.StartScrollPage - 1 %>
			<% if lp>ocoupon.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(lp) then %>
			<font color="red">[<%= lp %>]</font>
			<% else %>
			<a href="javascript:goPage(<%= lp %>)">[<%= lp %>]</a>
			<% end if %>
		<% next %>
	
		<% if ocoupon.HasNextScroll then %>
			<a href="javascript:goPage(<%= lp %>)">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<%
set ocoupon = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->