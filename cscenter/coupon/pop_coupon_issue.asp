<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터 쿠폰관리
' History : 이상구생성
'			2023.05.23 한용민 수정(보안 응답패킷 변조 체크 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%
' 관리자, cs팀 만 사용가능
if not(C_ADMIN_AUTH or C_CSUser) then
	response.write "해당매뉴는 관리자 이거나 cs팀만 사용가능합니다."
	dbget.close() : response.end
end if

dim userid, orderserial, jukyo, i, buf, page, ojumun, ix,iy
	userid = requestCheckvar(request("userid"),32)
	orderserial = requestCheckvar(request("orderserial"),32)
	jukyo = requestCheckvar(request("jukyo"),32)

if (userid = "") then
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if

page = 1

set ojumun = new COrderMaster
	ojumun.FPageSize = 5
	ojumun.FCurrPage = page
	ojumun.FRectUserID = userid
	ojumun.FRectOrderSerial = orderserial
	ojumun.QuickSearchOrderList

'' 과거 6개월 이전 내역 검색
if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderList

    if (ojumun.FResultCount>0) then
        response.write "<script>alert('6개월 이전 주문입니다.');</script>"
    end if
end if

'' 검색결과가 1개
dim ResultOneOrderserial
ResultOneOrderserial = ""
if (ojumun.FResultCount=1) then
    ResultOneOrderserial = ojumun.FItemList(0).FOrderSerial
end if

if ((orderserial <> "") and (ojumun.FResultCount <> 1)) then
	response.write "<script>alert('잘못된 주문번호입니다.');</script>"

	orderserial = ""
end if

dim Coupon3000IssueAllow, Coupon5perIssueAllow, CouponDeliverIssueAllow, CouponBirthday
Coupon3000IssueAllow = False
Coupon5perIssueAllow = False
CouponDeliverIssueAllow = True
CouponBirthday = False

' 관리자이거나 cs팀 정규직(어시 이상) 이경우 발행가능
if C_ADMIN_AUTH or C_CSpermanentUser then
	Coupon3000IssueAllow = True
	Coupon5perIssueAllow = True
end if

if C_ADMIN_AUTH or C_CSUser then
	CouponBirthday = True
end if

%>
<script type="text/javascript">

// 생일쿠폰
function IssueCouponBirthday(frm){
	<% if (CouponBirthday <> True) then %>
		alert("생일쿠폰 발행권한이 없습니다.");
		return;
	<% end if %>

	//if (CheckForm(frm) != true) {
	//	return;
	//}

	if (confirm("생일쿠폰을 발행하시겠습니까?") == true) {
		frm.submode.value = "IssueCouponBirthday"
		frm.submit();
	}
}

function IssueCoupon3000(frm)
{
	<% if (Coupon3000IssueAllow <> True) then %>
		alert("발행권한이 없습니다.");
		return;
	<% end if %>

	if (CheckForm(frm) != true) {
		return;
	}

	if (confirm("3000원 할인쿠폰을 발행하시겠습니까?") == true) {
		frm.submode.value = "issuecoupon3000"
		frm.submit();
	}
}

function IssueCoupon5per(frm)
{
	<% if (Coupon5perIssueAllow <> True) then %>
		alert("발행권한이 없습니다.");
		return;
	<% end if %>

	if (CheckForm(frm) != true) {
		return;
	}

	if (confirm("5% 할인쿠폰을 발행하시겠습니까?") == true) {
		frm.submode.value = "issuecoupon5per"
		frm.submit();
	}
}

function IssueCouponDeliver(frm)
{
	<% if (CouponDeliverIssueAllow <> True) then %>
		alert("발행권한이 없습니다.");
		return;
	<% end if %>

	if (CheckForm(frm) != true) {
		return;
	}

	if (confirm("무료배송비 쿠폰(<%=Cstr(getDefaultBeasongPayByDate(now()))%>원)을 발행하시겠습니까?") == true) {
		frm.submode.value = "issuecoupondeliver"
		frm.submit();
	}
}

function CheckForm(frm)
{
	if (frm.orderserial.value == "") {
		alert("관련 주문번호를 입력하세요.");
		return false;
	}

	if (frm.jukyo.value == "") {
		alert("발급사유를 선택하세요.");
		return false;
	}

	return true;
}

function SearchOrderSerial()
{
	document.frmsearch.orderserial.value = document.frm.orderserial.value;
	document.frmsearch.jukyo.value = document.frm.jukyo.value;
	document.frmsearch.submit();
}

function SetOrderSerial(orderserial)
{
	document.frm.orderserial.value = orderserial;
}

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>쿠폰발행</b> &nbsp; 쿠폰발행을 하면 CS메모에 등록되며, 즉시 발행됩니다.
		<br>*
		<% if (Coupon3000IssueAllow = True) then %>
		3000원 할인쿠폰 발행가능,
		<% end if %>
		<% if (Coupon5perIssueAllow = True) then %>
		5% 할인쿠폰 발행가능,
		<% end if %>
		<% if (CouponDeliverIssueAllow = True) then %>
		무료배송비 쿠폰(<%=Cstr(getDefaultBeasongPayByDate(now()))%>원) 발행가능.
		<% end if %>
	</td>
</tr>
</table>

<form name="frmsearch" method="get" action="" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="userid" value="<%= userid %>">
<input type="hidden" name="orderserial" value="<%= orderserial %>">
<input type="hidden" name="jukyo" value="<%= jukyo %>">
</form>
<form name="frm" method="post" action="/cscenter/coupon/domodifycoupon.asp" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="issuecoupon">
<input type="hidden" name="submode" value="">
<input type="hidden" name="userid" value="<%= userid %>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="20%" bgcolor="#f1f1f1">아이디 :</td>
  	<td bgcolor="#FFFFFF" width="25%" >
  	  <b><%= userid %></b>
  	</td>
  	<td height="30" width="20%" bgcolor="#f1f1f1">관련주문번호 :</td>
  	<td bgcolor="#FFFFFF"  >
  	  <input type=text name=orderserial value="<%= ResultOneOrderserial %>">
  	  <input type=button class="button" value="검색" onclick="SearchOrderSerial()">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" bgcolor="#f1f1f1">발급사유 :</td>
  	<td bgcolor="#FFFFFF"  colspan=3>
		<select class="select" name="jukyo">
			<option value=''></option>
     		<option value='배송지연' <% if (jukyo = "배송지연") then %>selected<% end if %>>배송지연</option>
     		<option value='CS서비스' <% if (jukyo = "CS서비스") then %>selected<% end if %>>CS서비스</option>
			<option value='품절' <% if (jukyo = "품절") then %>selected<% end if %>>품절</option>
			<option value='가격오류' <% if (jukyo = "가격오류") then %>selected<% end if %>>가격오류</option>
     		<option value='기타' <% if (jukyo = "기타") then %>selected<% end if %>>기타</option>
     	</select>
  	</td>
  </tr>
</table>
</form>

<p><br><p>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td height=25>구분</td>
    	<td>주문번호</td>
      	<td>구매자</td>
      	<td>구매총액</td>
      	<td>결제방법</td>
      	<td>거래상태</td>
      	<td>주문일</td>
      	<td>비고</td>
    </tr>
	<% if (ojumun.FresultCount > 0) then %>
        <% for i=0 to ojumun.FResultCount - 1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25><%= ojumun.FItemList(ix).CancelYnName %></td>
    	<td><%= ojumun.FItemList(i).FOrderSerial %></td>
    	<td><%= ojumun.FItemList(i).FBuyName %></td>
    	<td><%= FormatNumber(ojumun.FItemList(i).FTotalSum,0) %></td>
    	<td><%= ojumun.FItemList(i).JumunMethodName %></td>
    	<td><%= ojumun.FItemList(i).IpkumDivName %></td>
    	<td><acronym title="<%= ojumun.FItemList(i).FRegDate %>"><%= Left(ojumun.FItemList(i).FRegDate,10) %></acronym></td>
    	<td><input type=button class="button" value="선택" onclick="SetOrderSerial('<%= ojumun.FItemList(i).FOrderSerial %>')"></td>
    </tr>
        <% next %>
	<% else %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td colspan="8"> 검색된 결과가 없습니다.</td>
    </tr>
	<% end if %>
</table>
<br>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="center">
		<input type="button" class="button" value="3000원 할인 쿠폰발행" onClick="IssueCoupon3000(document.frm)">
		<input type="button" class="button" value="5% 할인 쿠폰발행" onClick="IssueCoupon5per(document.frm)">
		<input type="button" class="button" value="무료배송(<%=Cstr(getDefaultBeasongPayByDate(now()))%>원) 쿠폰발행" onClick="IssueCouponDeliver(document.frm)">
		<input type="button" class="button" value="생일 쿠폰발행" onClick="IssueCouponBirthday(document.frm)">
	</td>
</tr>
<tr>
	<td align="center">
		<br>
		<input type="button" class="button" value=" 창 닫 기 " onClick="self.close()">
	</td>
</tr>
</table>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
