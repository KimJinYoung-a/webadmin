<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 랜덤 쿠폰 설정 페이지
' Hieditor : 2023.09.25 정태훈 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/randomCouponCls.asp"-->
<%
dim evt_code, oCoupon, i, oIssue, totalRate
evt_code = request("evt_code")

'// 쿠폰 설정 리스트
set oCoupon = new RandomCouponCls
	oCoupon.FRectEvtCode = evt_code
set oIssue = new RandomCouponCls
	oIssue.FRectEvtCode = evt_code
	if evt_code <> "" then
	oCoupon.getRandomCouponList()
	oIssue.getRandomCouponIssueList()
	end if
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script>
function frmsubmit(page){
	frm.submit();
}
function fnRateEdit(idx){
	document.efrm.mode.value="edit";
	document.efrm.idx.value=idx;
	document.efrm.rate.value=$("#rate"+idx).val();
	document.efrm.coupon.value=$("#coupon"+idx).val();
	document.efrm.submit();
}
function fnRateDelete(idx){
	document.efrm.mode.value="delete";
	document.efrm.idx.value=idx;
	document.efrm.submit();
}
function fnRegCouponinfo(){
	$("#regbox").toggle('disable');
}
function frmsubmitCode(){
	var frm = document.wfrm;
	if(frm.evt_code.value==""){
		alert("이벤트 코드를 입력해주세요.");
		frm.evt_code.focus();
	}else if(frm.coupon.value==""){
		alert("쿠폰 번호를 입력해주세요.");
		frm.coupon.focus();
	}else if(frm.rate.value==""){
		alert("당첨 확율을 입력해주세요.");
		frm.rate.focus();
	}else{
		frm.submit();
	}
}
</script>
<style>
.disable {
  display: none;
}
</style>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="reload" value="ON">
<tr align="center" bgcolor="#FFFFFF">
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 이벤트 번호 : <input type="text" name="evt_code" value="<%= evt_code %>" size=10 maxlength=10>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('1');">
	</td>
</tr>
</form>
</table>
<br>
<a href="javascript:fnRegCouponinfo();">신규등록</a><br>
<table width="600" cellpadding="3" cellspacing="1" class="disable" bgcolor="<%= adminColor("tablebg") %>" id="regbox">
<form name="wfrm" method="post" action="/admin/event/coupon/dorateedit.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="add">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		이벤트 번호 : <input type="text" name="evt_code" value="<%= evt_code %>" size=10 maxlength=10>
		쿠폰 번호 : <input type="text" name="coupon" size=10 maxlength=10>
		당첨 확율 : <input type="text" name="rate" size=10 maxlength=10>&nbsp;&nbsp;
		<input type="button" class="button_s" value="등록" onClick="frmsubmitCode();">
	</td>
</tr>
</form>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>IDX</td>
	<td>당첨확율</td>
	<td>쿠폰번호</td>
	<td>쿠폰명</td>
	<td>쿠폰액</td>
	<td>최소구매금액</td>
	<td>수정/삭제</td>
</tr>
<% if oCoupon.FresultCount>0 then %>
	<% for i=0 to oCoupon.FresultCount-1 %>
	<% if oCoupon.FItemList(i).fdeleteYN = "N" then %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFFF';>
	<% else %>
		<tr align="center" bgcolor="#c1c1c1" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#c1c1c1';>
	<% end if %>
		<td>
			<%= oCoupon.FItemList(i).fidx %>
		</td>
		<td>
			<input type="text" name="rate" id="rate<%= oCoupon.FItemList(i).fidx %>" value="<%= oCoupon.FItemList(i).frate %>" size=10 maxlength=10>%
		</td>
		<td>
			<input type="text" name="coupon" id="coupon<%= oCoupon.FItemList(i).fidx %>" value="<%= oCoupon.FItemList(i).fcoupon %>" size=10 maxlength=10>
		</td>
		<td>
			<%= oCoupon.FItemList(i).fcouponname %>
		</td>
		<td>
			<%= oCoupon.FItemList(i).fcouponvalue %>
		</td>
		<td>
			<%= oCoupon.FItemList(i).fminbuyprice %>
		</td>
		<td>
			<input type="button" class="button_s" value="수정" onClick="fnRateEdit(<%= oCoupon.FItemList(i).fidx %>);">
			<input type="button" class="button_s" value="삭제" onClick="fnRateDelete(<%= oCoupon.FItemList(i).fidx %>);">
		</td>
		<% totalRate = totalRate + oCoupon.FItemList(i).frate %>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="7" align="center" class="page_link">당첨확율 <font color="red"><%=totalRate%></font>% (당첨확율 합산이 100% 미만,초과일 경우 오류가 발생합니다.</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="7" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
<br>
<table width="200" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>쿠폰번호</td>
	<td>지급 총 수량</td>
	<td>지급일</td>
</tr>
<% if oIssue.FresultCount>0 then %>
	<% for i=0 to oIssue.FresultCount-1 %>
		<tr align="center" bgcolor="#FFFFFF">
		<td>
			<%= oIssue.FItemList(i).fcoupon %>
		</td>
		<td>
			<%= oIssue.FItemList(i).ftCount %>
		</td>
		<td>
			<%= oIssue.FItemList(i).fregdate %>
		</td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="6" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
<form name="efrm" method="post" action="/admin/event/coupon/dorateedit.asp">
<input type="hidden" name="evt_code" value="<%= evt_code %>">
<input type="hidden" name="idx">
<input type="hidden" name="rate">
<input type="hidden" name="coupon">
<input type="hidden" name="mode" value="edit">
</form>
<%
set oCoupon = nothing
set oIssue = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->