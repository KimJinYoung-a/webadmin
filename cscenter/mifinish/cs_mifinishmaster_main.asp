<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : cs센터
' History	:  2007.06.01 이상구 생성
'              2023.11.15 한용민 수정(6개월이전 데이터도 처리가능하게 로직 변경)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_mifinishcls.asp"-->
<%
dim asid, orderserial, csdetailidx, ocsmifinishmaster, ocsmifinishDetailList, isChulgoState, i
	asid = requestCheckVar(request("asid"),10)
	csdetailidx = request("csdetailidx")

set ocsmifinishmaster = new CCSMifinishMaster
	ocsmifinishmaster.FRectAsid = asid
	ocsmifinishmaster.GetOneCSMaster

	if ocsmifinishmaster.FtotalCount < 1 then
		ocsmifinishmaster.FRectAsid = asid
		ocsmifinishmaster.FRectorder6MonthBefore = "Y"
		ocsmifinishmaster.GetOneCSMaster
	end if

orderserial = ocsmifinishmaster.FOneItem.FOrderSerial

set ocsmifinishDetailList = new CCSMifinishMaster
	ocsmifinishDetailList.FRectAsid = asid
	ocsmifinishDetailList.getMiFinishCSDetailList

	if ocsmifinishDetailList.FTotalCount < 1 then
		ocsmifinishDetailList.FRectAsid = asid
		ocsmifinishDetailList.FRectorder6MonthBefore = "Y"
		ocsmifinishDetailList.getMiFinishCSDetailList
	end if

isChulgoState = (ocsmifinishmaster.FOneItem.Fdivcd = "A000") or (ocsmifinishmaster.FOneItem.Fdivcd = "A100")

%>
<script type="text/javascript">

function confirmSubmit(){
    if (confirm('저장 하시겠습니까?')) {
    	var arrfinishstr = document.getElementsByName("finishstr");

    	for (var i = 0; i < arrfinishstr.length; i++) {
    		// 쉼표 바꾸기
    		arrfinishstr[i].value = arrfinishstr[i].value.replace(/,/g, "_XX_");
    	}

        document.frmmisend.submit();
    }
}

function popMifinishInput(csdetailidx) {
    var popwin = window.open('/cscenter/mifinish/popMifinishInput.asp?csdetailidx=' + csdetailidx,'popMifinishInput','width=1400,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popSendCallChange(iidx){
    if (confirm('고객님께 안내전화를 드렸습니까?')){
        frmmisendOne.csdetailidx.value=iidx;
        frmmisendOne.submit();
    }
}

function SearchThis(){
	location.href="/cscenter/mifinish/cs_mifinishmaster_main.asp?asid=" + frmsearch.asid.value;
}

</script>
<style type="text/css">
<!--
td { font-size:9pt; font-family:Verdana;}

.button {
	font-family: "굴림", "돋움";
	font-size: 10px;
	background-color: #E4E4E4;
	border: 1px solid #000000;
	color: #000000;
	height: 20px;
}
//-->
</style>

<!-- 검색 시작 -->
<form name="frmsearch" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" height="30">
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			ASID : <input type="text" class="text" name="asid" value="<%= asid %>" size=13 >
        	<% if ocsmifinishmaster.FOneItem.Fdeleteyn<>"N" then %>
			<b><font color="#CC3333">[취소CS]</font></b>
			<script language='javascript'>alert('취소된 CS 입니다.');</script>
			<% else %>
			[정상CS]
			<% end if %>
			&nbsp;
			&nbsp;
			구분 : <font color="<%= ocsmifinishmaster.FOneItem.getDivcdColor %>"><%= ocsmifinishmaster.FOneItem.getDivcdStr %></font>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="SearchThis();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" height="30">
		<td align="left">
			주문번호 : <%= orderserial %>
			&nbsp;
			고객명 : <%= ocsmifinishmaster.FOneItem.FBuyName %>
			&nbsp;
			핸드폰번호 : <%= ocsmifinishmaster.FOneItem.FBuyHp %>
			&nbsp;
			이메일 : <%= ocsmifinishmaster.FOneItem.FBuyEmail %>
		</td>
	</tr>
</table>
</form>
<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="csbutton" value="처리내용저장" onclick="confirmSubmit();">
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<form name="frmmisend" method="post" action="cs_mifinishmaster_main_process.asp" style="margin:0px;">
<input type="hidden" name="asid" value="<%= asid %>">
<input type="hidden" name="mode" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">상품코드</td>
		<td width="50">이미지</td>
		<td>상품명<br><font color="blue">[옵션]</font></td>
		<td width="40">접수<br>수량</td>
		<td width="40">부족<br>수량</td>
		<td width="80">처리기준일</td>
		<td width="60">소요<br>일수</td>
		<td width="80">미처리사유</td>
		<td width="80">처리예정일</td>
		<td width="80">물류/업체<br>작성메모</td>
		<td width="35">SMS</td>
		<td width="35">MAIL</td>
		<td width="35">CALL</td>
		<td width="120">CS처리구분</td>
		<td width="100">CS처리메모</td>
	</tr>
	<% for i=0 to ocsmifinishDetailList.FResultCount -1 %>

	<% if ocsmifinishDetailList.FItemList(i).FItemLackNo<>0 then %>
	<tr align="center" bgcolor="<%= adminColor("pink") %>">
	<% else %>
	<tr align="center" bgcolor="FFFFFF">
	<% end if %>
		<td>
			<% if ocsmifinishDetailList.FItemList(i).FIsUpcheBeasong="Y" then %>
			<font color="red"><%= ocsmifinishDetailList.FItemList(i).FItemID %></font>
			<% else %>
			<%= ocsmifinishDetailList.FItemList(i).FItemID %>
			<% end if %>
		</td>
		<td><img src="<%= ocsmifinishDetailList.FItemList(i).FSmallImage %>" width="50" height="50"></td>
		<td align="left">
			<%= ocsmifinishDetailList.FItemList(i).FItemName %>
			<% if ocsmifinishDetailList.FItemList(i).FItemOptionName<>"" then %>
			<br>
			<font color="blue">[<%= ocsmifinishDetailList.FItemList(i).FItemOptionName %>]</font>
			<% end if %>
		</td>
		<td><%= ocsmifinishDetailList.FItemList(i).FRegItemNo %></td>
		<td><font color="red"><b><% if ocsmifinishDetailList.FItemList(i).FItemLackNo=0 then response.write "-" else  response.write  ocsmifinishDetailList.FItemList(i).FItemLackNo end if%></b></font></td>
		<td>
		    <% if IsNULL(ocsmifinishDetailList.FItemList(i).FRegdate) then %>

		    <% else %>
		    <%= Left(ocsmifinishDetailList.FItemList(i).FRegdate,10) %>
		    <% end if %>
		</td>
		<td>
		    <!-- D+2 이상일경우, 빨간색으로 표시 -->
		    <% if (Not IsNULL(ocsmifinishDetailList.FItemList(i).getDPlusDate)) and (ocsmifinishDetailList.FItemList(i).getDPlusDate<>"")  then %>
    		    <% if (ocsmifinishDetailList.FItemList(i).getDPlusDate>=2) then %>
    		    <strong><font color="Red"><%= ocsmifinishDetailList.FItemList(i).getDPlusDateStr %></font></strong>
    		    <% else %>
    		    <%= ocsmifinishDetailList.FItemList(i).getDPlusDateStr %>
    		    <% end if %>
		    <% else %>
		    	<%= ocsmifinishDetailList.FItemList(i).getDPlusDateStr %>
		    <% end if %>
		</td>
		<td>
			<% if Not IsNull(ocsmifinishDetailList.FItemList(i).FMifinishReason) and (CStr(ocsmifinishDetailList.FItemList(i).Fcsdetailidx)=Cstr(csdetailidx)) then %>
				<font color="red">입력중</font>
			<% else %>
				<font color="<%= ocsmifinishDetailList.FItemList(i).getMiFinishCodeColor %>"><%= ocsmifinishDetailList.FItemList(i).getMiFinishCodeName %></font>
			<% end if %>
		</td>
		<td>
			<% if (ocsmifinishDetailList.FItemList(i).FMifinishReason<>"") and (ocsmifinishDetailList.FItemList(i).FMifinishReason<>"00") and (ocsmifinishDetailList.FItemList(i).FMifinishReason<>"05") then %>
				<%= ocsmifinishDetailList.FItemList(i).FMifinishipgodate %>
			<% end if %>
		</td>
		<td><%= ocsmifinishDetailList.FItemList(i).FrequestString %></td>
		<td>
		    <% if ocsmifinishDetailList.FItemList(i).FMifinishReason<>"" then %>
		        <a href="javascript:popMifinishInput('<%= ocsmifinishDetailList.FItemList(i).Fcsdetailidx %>')"><%= ocsmifinishDetailList.FItemList(i).FisSendSMS %></a>
		    <% elseif (ocsmifinishDetailList.FItemList(i).FIsUpcheBeasong="Y") and (ocsmifinishDetailList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMifinishInput('<%= ocsmifinishDetailList.FItemList(i).Fcsdetailidx %>')">N</a>
    	    <% end if %>
		</td>
		<td>
		    <% if ocsmifinishDetailList.FItemList(i).FMifinishReason<>"" then %>
		        <a href="javascript:popMifinishInput('<%= ocsmifinishDetailList.FItemList(i).Fcsdetailidx %>')"><%= ocsmifinishDetailList.FItemList(i).FisSendEmail %></a>
		    <% elseif (ocsmifinishDetailList.FItemList(i).FIsUpcheBeasong="Y") and (ocsmifinishDetailList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMifinishInput('<%= ocsmifinishDetailList.FItemList(i).Fcsdetailidx %>')">N</a>
    	    <% end if %>
		</td>
		<td>
		    <% if ocsmifinishDetailList.FItemList(i).FMifinishReason<>"" and isChulgoState then %>
		        <% if (ocsmifinishDetailList.FItemList(i).FisSendCall<>"Y") then %>
		            <a href="javascript:popSendCallChange('<%= ocsmifinishDetailList.FItemList(i).Fcsdetailidx %>')"><%= ocsmifinishDetailList.FItemList(i).FisSendCall %></a>
		        <% else %>
		            <%= ocsmifinishDetailList.FItemList(i).FisSendCall %>
		        <% end if %>

    	    <% end if %>
		</td>

		<% if (ocsmifinishDetailList.FItemList(i).FMifinishReason <> "") then %>
		<input type="hidden" name="arrcsdetailidx" value="<%= ocsmifinishDetailList.FItemList(i).Fcsdetailidx %>">
		<% end if %>

		<td>
		  <% if (ocsmifinishDetailList.FItemList(i).FMifinishReason <> "") then %>
		      <% if (ocsmifinishDetailList.FItemList(i).FMiFinishState = "7") then %>
		      완료
		      <input type=hidden name=state value="7">
		      <% else %>
		  	<select class="select" name="state">
				<option value="0" <% if (ocsmifinishDetailList.FItemList(i).FMiFinishState = "0") then response.write "selected" end if %>>미처리</option>
				<option value="4" <% if (ocsmifinishDetailList.FItemList(i).FMiFinishState = "4") then response.write "selected" end if %>>고객안내</option><!-- 신규(SMS/mail/통화시) -->
				<option value="6" <% if (ocsmifinishDetailList.FItemList(i).FMiFinishState = "6") then response.write "selected" end if %>>CS처리완료</option>
		  	</select>
		      <% end if %>
		  <% end if %>
		</td>
		<td>
		  <% if (ocsmifinishDetailList.FItemList(i).FMifinishReason <> "") then %>
		  <input type="text" class="text" name="finishstr" value="<%= ocsmifinishDetailList.FItemList(i).FfinishString %>" size="10">
		  <% end if %>
		</td>
	</tr>
	<% next %>
</table>
</form>

<form name="frmmisendOne" method="post" action="cs_mifinishmaster_main_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="SendCallChange">
<input type="hidden" name="asid" value="<%= asid %>">
<input type="hidden" name="csdetailidx" value="">
</form>

<%
set ocsmifinishmaster = Nothing
set ocsmifinishDetailList = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->