<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 핑거스 고객센터 마일리지
' Hieditor : 2015.05.27 이상구 생성
'			 2017.07.07 한용민 수정
'###########################################################
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenterv2/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/order/ordercls.asp"-->
<%
dim userid, orderserial, mileage, jukyo
dim i, buf
	userid = requestCheckvar(request("userid"),32)
	orderserial = requestCheckvar(request("orderserial"),32)
	mileage = requestCheckvar(request("mileage"),32)
	jukyo = requestCheckvar(request("jukyo"),32)

if (userid = "") then
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if

dim page
dim ojumun

page = 1

set ojumun = new COrderMaster
ojumun.FPageSize = 10
ojumun.FCurrPage = page

ojumun.FRectUserID = userid
ojumun.FRectOrderSerial = orderserial

if (Left(orderserial,1) = "B") then
	EXCLUDE_SITENAME = "diyitem"
end if

ojumun.QuickSearchOrderList

dim ix,iy


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

%>
<script language="javascript">

function SubmitForm()
{
	if (document.frm.orderserial.value == "") {
		alert("관련 주문번호를 입력하세요.");
		return;
	}

	/*
	if (document.frmsearch.orderserial.value != document.frm.orderserial.value) {
		alert("먼저 검색을 하세요.");
		return;
	}
	*/

	if (document.frm.mileage.value == "") {
		alert("적립액을 정확히 입력하세요.");
		return;
	}

	if (document.frm.mileage.value*0 != 0) {
		alert("적립액을 정확히 입력하세요.");
		return;
	}

	if (document.frm.mileage.value == 0) {
		alert("적립액은 0 이 될 수 없습니다.");
		return;
	}

	if (document.frm.jukyo.value == "") {
		alert("적립내용이 없습니다.");
		return;
	}

	if (confirm("적립요청 하시겠습니까?") == true) {
		document.frm.submit();
	}
}

function SearchOrderSerial(orderserial)
{
	if (document.frm.orderserial.value == "") {
		alert("주문번호를 입력하세요.");
		return;
	}

	document.frmsearch.orderserial.value = document.frm.orderserial.value;
	document.frmsearch.submit();
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>

</table>
<!-- 표 상단바 끝-->


<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br><b>마일리지 적립요청</b> 적립요청을 하시면, CS처리리스트에 등록되며, 관리자 승인과 함께 적립됩니다.
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <form name="frmsearch" method="get" action="" onsubmit="return false;">
  	<input type="hidden" name="userid" value="<%= userid %>">
  	<input type="hidden" name="orderserial" value="<%= orderserial %>">
  </form>
  <form name="frm" method="post" action="domodifymileage.asp" onsubmit="return false;">
  <input type="hidden" name="mode" value="request">
  <input type="hidden" name="userid" value="<%= userid %>">
  <tr align="left">
  	<td height="30" width="20%" bgcolor="#DDDDFF">아이디 :</td>
  	<td bgcolor="#FFFFFF" width="25%" >
  	  <b><%= userid %></b>
  	</td>
  	<td height="30" width="20%" bgcolor="#DDDDFF">관련주문번호 :</td>
  	<td bgcolor="#FFFFFF"  >
  	  <input type=text name=orderserial value="<%= ResultOneOrderserial %>">
  	  <input type=button value="검색" onclick="SearchOrderSerial()">
  	</td>

  </tr>
  <tr align="left">
  	<td height="30" bgcolor="#DDDDFF">적립액 :</td>
  	<td bgcolor="#FFFFFF" >
	  <input type=text name=mileage value="<%= mileage %>">
  	</td>
  	<td height="30" bgcolor="#DDDDFF">적립내용 :</td>
  	<td bgcolor="#FFFFFF" >
		<select class="select" name="jukyo">
			<option value='' selected>등록안함</option>
     		<option value='입금차액' <% if (jukyo = "입금차액") then %>selected<% end if %>>입금차액</option>
     		<option value='상품차액' <% if (jukyo = "상품차액") then %>selected<% end if %>>상품차액</option>
     		<option value='배송지연' <% if (jukyo = "배송지연") then %>selected<% end if %>>배송지연</option>
     		<option value='CS서비스' <% if (jukyo = "CS서비스") then %>selected<% end if %>>CS서비스</option>
     		<option value='상품대금환불' <% if (jukyo = "상품대금환불") then %>selected<% end if %>>상품대금환불</option>
     		<option value='기타' <% if (jukyo = "기타") then %>selected<% end if %>>기타</option>
     	</select>
  	</td>
  </tr>
</form>
</table>

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
    </tr>
<% if (orderserial <> "") then %>
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
    </tr>
        <% next %>
	<% else %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td colspan="6"> 검색된 결과가 없습니다.</td>
    </tr>
	<% end if %>
<% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="적립요청" onClick="SubmitForm();">
          <input type="button" value=" 창 닫 기 " onClick="self.close()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<!-- 표 하단바 끝-->

<p>
<%
'set OUserInfo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
