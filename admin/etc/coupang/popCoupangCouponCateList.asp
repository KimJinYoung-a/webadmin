<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/coupang/coupangcls.asp"-->
<%
Dim page, oCoupang, i, idx, oCoupangMaster, mallid
page		= request("page")
idx			= request("idx")
mallid		= request("mallid")

If page = "" Then page = 1

Dim startDate, endDate, disCount, couponName, maxDiscountPrice, couponType
couponType = "RATE"
If idx <> "" Then
	SET oCoupangMaster = new CCoupang
		oCoupangMaster.FRectIdx = idx
		oCoupangMaster.getCouponCateOneItem

		startDate = oCoupangMaster.FOneItem.FStartDate
		endDate = oCoupangMaster.FOneItem.FEndDate
		maxDiscountPrice = oCoupangMaster.FOneItem.FMaxDiscountPrice
		disCount = oCoupangMaster.FOneItem.FDisCount
		couponName = oCoupangMaster.FOneItem.FCouponName
	SET oCoupangMaster = nothing
End If

Set oCoupang = new CCoupang
	oCoupang.FCurrPage					= page
	oCoupang.FPageSize					= 50
	oCoupang.getCouponCateList
%>
<link rel="stylesheet" href="/bct.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function popMarginDetail(imidx, icouponid){
	var popdetail=window.open('/admin/etc/coupang/popCoupangCouponDetail.asp?midx='+imidx+'&couponId='+icouponid,'popMarginDetail','width=700,height=700,scrollbars=yes,resizable=yes');
	popdetail.focus();
}
function fnSaveCoupon(){
    if ($("#termSdt").val() == "") {
        alert('시작일을 입력하세요');
        return false;
    }
    if ($("#termEdt").val() == "") {
        alert('종료일을 입력하세요');
        return false;
    }
    if ($("#maxDiscountPrice").val() == "") {
        alert('최대할인금액을 입력하세요');
        $("#maxDiscountPrice").focus();
        return false;
    }
    if ($("#disCount").val() == "") {
        alert('할인율을 입력하세요');
        $("#disCount").focus();
        return false;
    }
    if ($("#couponName").val() == "") {
        alert('프로모션명을 입력하세요');
        $("#couponName").focus();
        return false;
    }
    if (confirm('저장 하시겠습니까?')){
        document.frmSave.target = "xLink";
        document.frmSave.submit();
    }
}

// 선택된 상품 등록
function CoupangCouponRegProcess() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 데이터가 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("데이터가 없습니다.");
		return;
	}

    if (confirm('Coupang에 선택하신 ' + chkSel + '개 쿠폰을 등록 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "COUPONREG";
        document.frmSvArr.action = "<%=apiURL%>/outmall/coupang/actCoupangReq.asp"
        document.frmSvArr.submit();
    }
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSave" method="post" action="procCoupangCoupon.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="cateMaster">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2" bgcolor="<%= adminColor("tabletop") %>">+ 기간별 쿠폰 등록 및 수정</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">기간</td>
	<td align="LEFT">
        <input type="text" id="termSdt" name="startDate" readonly size="11" maxlength="10" value="<%= LEFT(startDate, 10) %>" style="cursor:pointer; text-align:center;" /> ~
        <input type="text" id="termEdt" name="endDate" readonly size="11" maxlength="10" value="<%= LEFT(endDate, 10) %>" style="cursor:pointer; text-align:center;" />
        <script type="text/javascript">
            var CAL_Start = new Calendar({
                inputField : "termSdt", trigger    : "termSdt",
                onSelect: function() {
                    var date = Calendar.intToDate(this.selection.get());
                    CAL_End.args.min = date;
                    CAL_End.redraw();
                    this.hide();

                    if(frm.endDate.value==""||getDayInterval(frm.startDate.value, frm.endDate.value) < 0) frm.endDate.value=frm.startDate.value;
                    doInsertDayInterval();	// 날짜 자동계산
                }, bottomBar: true, dateFormat: "%Y-%m-%d"
            });
            var CAL_End = new Calendar({
                inputField : "termEdt", trigger    : "termEdt",
                onSelect: function() {
                    var date = Calendar.intToDate(this.selection.get());
                    CAL_Start.args.max = date;
                    CAL_Start.redraw();
                    this.hide();

                    if(frm.startDate.value==""||getDayInterval(frm.startDate.value, frm.endDate.value) < 0) frm.startDate.value=frm.endDate.value;
                    doInsertDayInterval();	// 날짜 자동계산
                }, bottomBar: true, dateFormat: "%Y-%m-%d"
            });
        </script>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">최대할인금액</td>
	<td align="LEFT">
		<input type="text" id="maxDiscountPrice" size="10" name="maxDiscountPrice" value="<%= maxDiscountPrice %>">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">할인율</td>
	<td align="LEFT">
		<input type="text" id="disCount" size="3" name="disCount" value="<%= disCount %>">%
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">프로모션명</td>
	<td align="LEFT">
		<input type="text" id="couponName" size="50" name="couponName" value="<%= couponName %>">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">할인방식</td>
	<td align="LEFT">
		<label><input type="radio" name="couponType" value="RATE" <%= Chkiif(couponType="RATE", "checked", "") %>>정률할인</label>
		<label><input type="radio" name="couponType" value="FIXED_WITH_QUANTITY" <%= Chkiif(couponType="FIXED_WITH_QUANTITY", "checked", "") %> >수량별 정액할인</label>
		<label><input type="radio" name="couponType" value="PRICE" <%= Chkiif(couponType="PRICE", "checked", "") %> >정액할인</label>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2">
		<input type="button" class="button" value="저장" onclick="fnSaveCoupon();">
	</td>
</tr>
</form>
</table>

<br /><br />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="3" bgcolor="<%= adminColor("tabletop") %>">기간별 쿠폰 리스트</td>
</tr>
</form>
</table>

<br />
<tr bgcolor="#FFFFFF">
    <td style="padding:5 0 5 0">
	    <table width="100%" class="a">
	    <tr>
	    	<td valign="top">
				<input class="button" type="button" id="btnRegSel" value="할인쿠폰생성" onClick="CoupangCouponRegProcess();">&nbsp;&nbsp;
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
<br />
<!-- 리스트 시작 -->
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		검색결과 : <b><%= FormatNumber(oCoupang.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCoupang.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="20">Idx</td>
	<td width="200">기간</td>
	<td width="100">최대할인금액</td>
	<td width="70">할인율</td>
	<td>프로모션명</td>
	<td width="100">쿠폰ID</td>
	<td width="100">등록일</td>
	<td width="50">관리</td>
</tr>
<% For i=0 to oCoupang.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oCoupang.FItemList(i).FIdx %>"></td>
	<td><%= oCoupang.FItemList(i).FIdx %></td>
	<td style="cursor:pointer;" onclick="popMarginDetail('<%= oCoupang.FItemList(i).FIdx %>', '<%= oCoupang.FItemList(i).FCouponId %>');"><%= LEFT(oCoupang.FItemList(i).FStartDate, 10) %> ~ <%= LEFT(oCoupang.FItemList(i).FEndDate, 10) %></td>
	<td><%= oCoupang.FItemList(i).FMaxDiscountPrice %></td>
	<td><%= oCoupang.FItemList(i).FDiscount %>%</td>
	<td><%= oCoupang.FItemList(i).FCouponName %></td>
	<td>
    <%
    	If Not(IsNULL(oCoupang.FItemList(i).FCouponId)) Then
        	Response.Write oCoupang.FItemList(i).FCouponId
		Else
			If Not(IsNULL(oCoupang.FItemList(i).FRequestedId)) Then
				Response.Write "승인대기" & "<br>(" & oCoupang.FItemList(i).FRequestedId & ")"
			End If
		End If
	%>
	</td>
	<td><%= LEFT(oCoupang.FItemList(i).FRegDate, 10) %></td>
	<td><input type="button" class="button" value="수정" onclick="javascript:location.href='/admin/etc/coupang/popCoupangCouponCateList.asp?idx=<%= oCoupang.FItemList(i).FIdx %>&mallid=<%= mallid %>';"></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="9" align="center" bgcolor="#FFFFFF">
        <% if oCoupang.HasPreScroll then %>
		<a href="javascript:goPage('<%= oCoupang.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oCoupang.StartScrollPage to oCoupang.FScrollCount + oCoupang.StartScrollPage - 1 %>
    		<% if i>oCoupang.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oCoupang.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% Set oCoupang = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
