<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터 마일리지 적립
' Hieditor : 이상구 생성
'			 2023.09.05 한용민 수정(소스 표준코딩으로 변경. 적립내용 클릭 지정 추가.)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<%
dim userid, orderserial, mileage, jukyo, i, buf, gubun01, gubun02, gubun01name, gubun02name, page, ojumun, defaultCSRefundLimit
dim ix,iy, omakeridList
	userid = requestCheckvar(request("userid"),32)
	orderserial = requestCheckvar(request("orderserial"),32)
	mileage = requestCheckvar(request("mileage"),32)
	jukyo = requestCheckvar(request("jukyo"),32)
	gubun01 = requestCheckvar(request("gubun01"),32)
	gubun02 = requestCheckvar(request("gubun02"),32)
	gubun01name = requestCheckvar(request("gubun01name"),32)
	gubun02name = requestCheckvar(request("gubun02name"),32)

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

if (orderserial = "") and (ojumun.FResultCount=1) then
	orderserial = ojumun.FItemList(0).FOrderSerial
end if

set omakeridList = new COrderMaster
if (orderserial <> "") then
	omakeridList.FRectOrderSerial = orderserial
	omakeridList.getUpcheBeasongMakerList
end if

defaultCSRefundLimit = GetUserRefundAuthLimit(session("ssBctId"))

%>
<script type="text/javascript">

function SubmitForm(){
	var jukyo;

	if (document.frm.orderserial.value == "") {
		alert("관련 주문번호를 입력하세요.");
		return;
	}

	if (document.frm.gubun01.value == "") {
		alert("사유구분을 선택하세요.");
		return;
	}

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

	if (document.frm.mileage.value*1 > <%= defaultCSRefundLimit %>) {
		alert("<%= FormatNumber(defaultCSRefundLimit,0) %> 마일리지를 초과하여 적립할 수 없습니다.\n부득이 더 많은 적립금을 부여해야 될 경우 파트장에게 문의해주세요.");
		return;
	}

	<% if omakeridList.FresultCount > 0 then %>
	if (((document.frm.gubun01.value != "C004") || (document.frm.gubun02.value != "CD13")) && (document.frm.requiremakerid.value == "")) {
		alert("관련 브랜드를 선택하세요.");
		return;
	}
	<% end if %>

	if (document.frm.contents_jupsu.value == "") {
		alert("적립내용이 없습니다.");
		return;
	}

	if (confirm("적립요청 하시겠습니까?") == true) {
		document.frm.submit();
	}
}

function SubmitFormForce() {
	if (document.frm.gubun01.value == "") {
		alert("사유구분을 선택하세요.");
		return;
	}

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

	if (confirm("적립 하시겠습니까?") == true) {
		document.frm.mode.value = "requestForce";
		document.frm.submit();
	}
}

function SetOrderSerial(orderserial){
	var frm = document.frm;
	var DisplayMakerID = false;
	frm.orderserial.value = orderserial;

	if (frm.contents_jupsu.value != "") {
		if (confirm("관련브랜드를 표시하시겠습니까?\n(입력된 접수내용은 사라집니다.)")) {
			DisplayMakerID = true;
		}
	} else {
		DisplayMakerID = true;
	}

	if (DisplayMakerID == true) {
		frm.method = "get";
		frm.action = "";
		frm.submit();
	}
}

function selectGubun(value_gubun01, value_gubun02, value_gubun01name, value_gubun02name, name_gubun01, name_gubun02, name_gubun01name, name_gubun02name ,name_frm, targetDiv) {
    var frm = eval(name_frm);

    eval("document." + name_frm + "." + name_gubun01).value = value_gubun01;
    eval("document." + name_frm + "." + name_gubun02).value = value_gubun02;
    eval("document." + name_frm + "." + name_gubun01name).value = value_gubun01name;
    eval("document." + name_frm + "." + name_gubun02name).value = value_gubun02name;
}

function regcontents_jupsu(contents_jupsu) {
	frm.contents_jupsu.value=contents_jupsu;
}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<b>마일리지 적립요청</b> 적립요청을 하시면, CS처리리스트에 등록되며, 관리자 승인과 함께 적립됩니다.
		<br>* <font color=red><%= FormatNumber(defaultCSRefundLimit,0) %> 마일리지</font> 초과 적립불가
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frm" method="post" action="/cscenter/mileage/domodifymileage.asp" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="request">
<input type="hidden" name="userid" value="<%= userid %>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="left">
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">아이디 :</td>
  		<td bgcolor="#FFFFFF" width="40%" >
  			<b><%= userid %></b>
  		</td>
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">관련주문번호 :</td>
  		<td bgcolor="#FFFFFF" width="40%"  >
  			<input type=text name=orderserial value="<%= orderserial %>">
  		</td>
	</tr>
	<tr align="left">
		<td height="30" bgcolor="<%= adminColor("tabletop") %>">사유구분 :</td>
  		<td bgcolor="#FFFFFF" colspan="3">
			<input type="hidden" name="gubun01" value="<%= gubun01 %>">
			<input type="hidden" name="gubun02" value="<%= gubun02 %>">
			<input class="text_ro" type="text" name="gubun01name" value="<%= gubun01name %>" size="16" Readonly >
			&gt;
			<input class="text_ro" type="text" name="gubun02name" value="<%= gubun02name %>" size="16" Readonly >
			&nbsp;
			[<a href="javascript:selectGubun('C006','CF06','물류관련','출고지연','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">출고지연</a>]
			[<a href="javascript:selectGubun('C004','CD05','공통','품절','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">품절</a>]
			[<a href="javascript:selectGubun('C006','CF01','물류관련','오발송','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">오배송</a>]
			[<a href="javascript:selectGubun('C005','CE03','상품관련','상품등록오류','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">상품등록오류</a>]
			[<a href="javascript:selectGubun('C004','CD12','공통','업체응대불량','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">업체응대불량</a>]
			[<a href="javascript:selectGubun('C004','CD14','공통','기타업체과실','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">기타업체과실</a>]
			&nbsp;
			[<a href="javascript:selectGubun('C004','CD13','공통','CS서비스','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">CS서비스</a>]
			<!--
			[<a href="javascript:selectGubun('C004','CD99','공통','기타','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">기타</a>]
			-->
  		</td>
	</tr>
	<tr align="left">
  		<td height="30" bgcolor="<%= adminColor("tabletop") %>">적립액 :</td>
  		<td bgcolor="#FFFFFF" >
			<input type=text name=mileage value="<%= mileage %>">
  		</td>
  		<td height="30" bgcolor="<%= adminColor("tabletop") %>">관련브랜드 :</td>
  		<td bgcolor="#FFFFFF" >
			<select class="select" name="requiremakerid">
				<option></option>
				<% if orderserial <> "" and omakeridList.FResultCount > 0 then %>
				<% for i=0 to omakeridList.FResultCount - 1 %>
				<option value="<%= omakeridList.FItemList(i).Fmakerid %>"><%= CHKIIF(omakeridList.FItemList(i).Fmakerid="10x10logistics", "텐바이텐배송", omakeridList.FItemList(i).Fmakerid) %></option>
				<% next %>
				<% end if %>
			</select>
		</td>
	</tr>
	<tr align="left">
  		<td height="30" bgcolor="<%= adminColor("tabletop") %>">적립내용 :</td>
  		<td bgcolor="#FFFFFF" colspan="3">
			<textarea class='textarea' id="contents_jupsu" name="contents_jupsu" cols="80" rows="6"></textarea>
			<% if C_ADMIN_AUTH then %>
				<br>
				[<a href="#" onclick="regcontents_jupsu('이벤트 마일리지 보정'); return;">이벤트마일리지보정</a>]
				[<a href="#" onclick="regcontents_jupsu('주문 마일리지 보정'); return;">주문마일리지보정</a>]
			<% end if %>
  		</td>
	</tr>
</table>
</form>

<% if orderserial = "" then %>
	<br />
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
<% end if %>

<!-- 액션 시작 -->
<br />
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="center">
		<input type="button" class="button" value="적립요청" onClick="SubmitForm();">

		<% if C_CSPowerUser or C_ADMIN_AUTH then %>
			&nbsp;
			<input type="button" class="button" value="적립(관리자)" onClick="SubmitFormForce();"> 고객센터파트장이상, 관리자권한
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
