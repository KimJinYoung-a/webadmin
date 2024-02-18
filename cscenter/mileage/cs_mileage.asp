<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 고객센터 마일리지관리
' History : 이상구 생성
'			2023.07.21 한용민 수정(마일리지소멸 년별->월별 로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/mileage/sp_pointcls.asp" -->
<!-- #include virtual="/lib/classes/mileage/sp_mileage_logcls.asp" -->
<%
dim i, userid, showdelete, showtype, currpage, showdetail, research, myMileage, myOffMileage, myMileageLog, oExpireMile
dim oBeforeSixMonth, beforeSixMonthSUM, currentdate, expireDate
	userid      = requestCheckVar(trim(request("userid")),32)
	showdelete  = requestCheckVar(trim(request("showdelete")),1)	'삭제내역 표시여부
	showtype    = requestCheckVar(trim(request("showtype")),1)		'보너스(B)구매(O)사용(S) 마일리지
	currpage    = requestCheckVar(getNumeric(trim(request("currpage"))),10)
	showdetail  = requestCheckVar(trim(request("showdetail")),2)
	research  = requestCheckVar(trim(request("research")),2)

if (research = "") then
	''showdelete = "Y"
end if

if (currpage = "") then currpage = 1
if ((showtype <> "S") and (showtype <> "O") and (showtype <> "B") and (showtype <> "X")) then showtype = "A"
if (showdelete = "") then showdelete = "N"
if (showdetail="") then showtype=""

currentdate=date()

' 이번달말일
expireDate = dateadd("d",-1,dateserial(year(dateadd("m",+1,currentdate)),month(dateadd("m",+1,currentdate)),"01"))

set myMileage = new TenPoint
myMileage.FRectUserID = userid
if (userid <> "") then
    myMileage.getTotalMileage
end if

set myOffMileage = new TenPoint
myOffMileage.FGubun = "my10x10"
myOffMileage.FRectUserID = userid
if (userid <> "") then
    myOffMileage.getOffShopMileagePop
end if

set myMileageLog = New CMileageLog
myMileageLog.FPageSize = 100
myMileageLog.FCurrPage = Cint(currpage)
myMileageLog.FRectUserid = userid
myMileageLog.FRectMileageLogType = showtype
myMileageLog.FRectShowDelete = showdelete

if ((userid <> "") and (showtype <> "") and (showdetail<>"")) then
	if (showtype = "A") then
		myMileageLog.getMileageLogAll
		'myMileageLog.getMileageLog
	else
		myMileageLog.getMileageLog
	end if
end if

' 만료예정  마일리지
set oExpireMile = new CMileageLog
	oExpireMile.FRectUserid = userid
	oExpireMile.FRectExpireDate = expireDate
	if (userid<>"") then
		'oExpireMile.getNextExpireMileageSum
		oExpireMile.getNextExpireMileageMonthlySum
	end if

set oBeforeSixMonth = new CMileageLog
oBeforeSixMonth.FRectUserid = userid

beforeSixMonthSUM = 0
if (userid<>"") then
    oBeforeSixMonth.GetRealSumBuyMileageBeforeSixMonth()
	beforeSixMonthSUM = oBeforeSixMonth.FOneItem.FbeforesixmonthSUM
end if

%>
<script type='text/javascript'>

function gotoPage(page){
	document.frmpage.currpage.value = page;
	document.frmpage.submit();
}

function changeType(showtype){
    document.frm.showdetail.value = "on";
	document.frm.showtype.value = showtype;
	document.frm.submit();
}

function popMileageRequest(userid, orderserial, mileage, jukyo) {
	// 필수 : 아이디
	// 옵션 : 주문번호, 마일리지, 적요내용

	if (userid == "") {
		alert("아이디가 없습니다.");
		return;
	}

    var popwin = window.open('/cscenter/mileage/pop_mileage_request.asp?userid=' + userid + '&orderserial=' + orderserial + '&mileage=' + mileage + '&jukyo=' + jukyo,'popMileageRequest','width=1400,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popOffMileageList(userid) {
	if (userid == "") {
		alert("아이디가 없습니다.");
		return;
	}

    var popwin = window.open('/admin/offshop/offmileagelist.asp?menupos=651&userid=' + userid,'popOffMileageList','width=1500,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popYearExpireMileList(yyyymmdd,userid){
    var popwin = window.open('/cscenter/mileage/popAdminExpireMileSummary.asp?yyyymmdd=' + yyyymmdd + '&userid=' + userid,'popAdminExpireMileSummary','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}
function popmonthlyExpireMileList(yyyymmdd,userid){
    var popwin = window.open('/cscenter/mileage/popAdminExpireMileMonthlySummary.asp?yyyymmdd=' + yyyymmdd + '&userid=' + userid+'&menupos=<%=menupos%>','popAdminExpireMileSummary','width=1400,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

<% if C_ADMIN_AUTH then %>
	function SubmitFormDelForce(idx) {
		var frm = document.frmAct;

		if (frm.userid.value == "") {
			alert("에러.");
			return;
		}

		if (idx == "") {
			alert("에러.");
			return;
		}

		if (confirm("[관리자]주의!!!!\n\n지난달에 부여된 마일리지는 삭제하면 안됩니다.\n\n삭제 하시겠습니까?") == true) {
			frm.mode.value = "delForce";
			frm.idx.value = idx;
			frm.submit();
		}
	}

	function jsReCalcSum() {
		var frm = document.frmAct;

		if (frm.userid.value == "") {
			alert("아이디가 없습니다.");
			return;
		}

		if (confirm("[관리자]마일리지를 재계산합니다.\n\n진행하시겠습니까?") == true) {
			frm.mode.value = "recalcmile";
			frm.submit();
		}
	}
<% end if %>

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="showtype" value="<%= showtype %>">
<input type="hidden" name="showdetail" value="<%= showdetail %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		아이디 : <input type="text" class="text" name="userid" value="<%= userid %>" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		&nbsp;
		<input type="checkbox" name="showdelete" <%= chkIIF(showdelete="Y","checked","") %> value="Y">삭제(구매내역의 경우 취소) 표시
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button" value="검색" onclick="document.frm.submit()">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		<% if C_ADMIN_AUTH then %>
			<input type="button" class="button" value="마일리지 재계산[관리자]" onClick="jsReCalcSum()">
		<% end if %>
	</td>
</tr>
</table>
</form>
<br>

<form name="frmWrite" action="userMileage_Process.asp" onsubmit="return checkForm();" style="margin:0px;">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td>
		<input type="button" class="button" value="적립요청" onclick="popMileageRequest('<%= userid %>', '', 0, '');">
		적립요청을 하시면, CS처리리스트에 등록되며, 관리자 승인과 함께 적립됩니다.
	</td>
</tr>
<tr>
	<td>
		<span id="divWrite" style="float:left; display:none">
			<input type="hidden" name="mode" value="INS">
			<input type="hidden" name="userID" value="<%=userID%>">
			주문번호 :
			<input type="text" name="orderSerial" size="11" maxlength="11" onkeydown="onlyNumber(this,event);" class="text">
			&nbsp;
			적립액 :
			<input type="text" name="savePoint" size="5" maxlength="5" style="text-align:right;" onkeydown="onlyNumber(this,event);" class="text">
			&nbsp;
			적립내용 :
			<select class="select" name="etcTitle">
				<option value='' selected>등록안함</option>
				<option value='입금차액'>입금차액</option>
				<option value='상품차액'>상품차액</option>
				<option value='배송지연'>배송지연</option>
				<option value='CS서비스'>CS서비스</option>
				<option value='상품대금환불'>상품대금환불</option>
				<option value='기타'>기타</option>
			</select>
			<input type="submit" class="button" value="등록">
		</span>
	</td>
</tr>
</table>
</form>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<img src="/images/icon_arrow_down.gif" align="absbottom">
		<strong>요약정보</strong>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="35">
	<td height=25>구분</td>
	<td>현재마일리지</td>
	<td>보너스 마일리지</td>
	<td>구매적립 마일리지<br>(온라인+아카데미)</td>
	<td>구매적립예정 마일리지<br>(온라인+아카데미)</td>
	<td>사용한 마일리지</td>
	<td>소멸된 마일리지</td>
</tr>
<% if (userid <> "") then %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25>온라인</td>
    	<td><strong><%=FormatNumber(myMileage.FTotalMileage,0) %></strong></td>
    	<td><%=FormatNumber(myMileage.FBonusMileage,0) %></td>
    	<td><%=FormatNumber(myMileage.FTotJumunmileage + myMileage.FAcademymileage,0) %></td>
      	<td><%=FormatNumber(myMileage.Fmichulmile + myMileage.FmichulmileACA,0) %></td>
      	<td><%=FormatNumber(myMileage.FSpendMileage*-1,0) %></font></td>
      	<td><%=FormatNumber(myMileage.FrealExpiredMileage*-1,0) %></font></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25>오프라인</td>
    	<td><a href="javascript:popOffMileageList('<%= userID %>')"><strong><%=FormatNumber(myOffMileage.FOffShopMileage,0) %></strong></a></td>
    	<td colspan=5></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25>소멸 대상 마일리지</td>
    	<td>
			<!--<a href="javascript:popYearExpireMileList('<%'= oExpireMile.FRectExpireDate %>','<%'= userid %>');">-->
			<a href="#" onclick="popmonthlyExpireMileList('<%= oExpireMile.FRectExpireDate %>','<%= userid %>'); return false;">
			<%= FormatNumber(oExpireMile.FOneItem.getMayExpireTotal,0) %></a>
		</td>
    	<td colspan=5 align=left>
			&nbsp;&nbsp;
			<!--<a href="javascript:popYearExpireMileList('<%'= oExpireMile.FRectExpireDate %>','<%'= userid %>');">-->
			<a href="#" onclick="popmonthlyExpireMileList('<%= oExpireMile.FRectExpireDate %>','<%= userid %>'); return false;">
			* 소멸일자 : <%= expireDate %></a>
		</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25>최근(6개월이내) 구매적립합계</td>
    	<td><%= FormatNumber(myMileage.FRecentJumunMileage,0) %></td>
    	<td colspan=5 align=left></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25>6개월이전 구매적립합계</td>
    	<td>
			<%= FormatNumber(myMileage.FOldJumunmileage,0) %>
			<% if (beforeSixMonthSUM > 0) and (beforeSixMonthSUM <> myMileage.FOldJumunmileage) then %>
			<br /><font color="red">(<%= FormatNumber(beforeSixMonthSUM,0) %>)</font>
			<% end if %>
		</td>
    	<td colspan=5 align=left> &nbsp;&nbsp;* 구매 마일리지에는 6개월전 이전 내역이 표시되지 않습니다.</td>
    </tr>
    	<% if (myMileage.FAcademyMileage>0) then %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25>아카데미 주문적립</td>
    	<td><%= FormatNumber(myMileage.FAcademyMileage,0) %></td>
    	<td colspan=5 align=left> &nbsp;&nbsp;* 구매적립예정 마일리지는 <font color="red">상품출고시</font> 적립됩니다.</td>
    </tr>
		<% else %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25>아카데미 주문적립</td>
    	<td>없음</td>
    	<td colspan=5 align=left> &nbsp;&nbsp;* 구매적립예정 마일리지는 <font color="red">상품출고시</font> 적립됩니다.</td>
    </tr>
    	<% end if %>
<% else %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>온라인</td>
    	<td>-</td>
    	<td>-</td>
      	<td>-</td>
		<td>-</td>
      	<td>-</td>
      	<td>-</td>
    </tr>
<% end if %>
</table>

<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		<img src="/images/icon_arrow_down.gif" align="absbottom">
		<strong>상세내역 : </strong>
		<%=chkIIF(showtype="A","<strong>","")%><a href="javascript:changeType('A')">전체보기</a><%=chkIIF(showtype="A","</strong>","")%>
		|
		<%=chkIIF(showtype="B","<strong>","")%><a href="javascript:changeType('B')">보너스 마일리지</a><%=chkIIF(showtype="B","</strong>","")%>
		|
		<%=chkIIF(showtype="O","<strong>","")%><a href="javascript:changeType('O')">구매 마일리지</a><%=chkIIF(showtype="O","</strong>","")%>
		|
		<%=chkIIF(showtype="S","<strong>","")%><a href="javascript:changeType('S')">사용 마일리지</a><%=chkIIF(showtype="S","</strong>","")%>
		|
		<%=chkIIF(showtype="X","<strong>","")%><a href="javascript:changeType('X')">소멸 마일리지</a><%=chkIIF(showtype="X","</strong>","")%>
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		<% if (showdetail="on") then %>
			검색결과 : <b>총 <%= myMileageLog.FTotalCount %> 건</b> 페이지 : <b><%= currpage %> / <%= myMileageLog.FTotalPage %></b>
		<% end if %>
	</td>
</tr>
<% if (showdetail="on") then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height=25>아이디</td>
		<td>적립일자</td>
		<td>구분</td>
		<td>적요내용</td>
		<td>마일리지</td>
		<td>잔액</td>
		<td>관련주문번호</td>
		<td>삭제여부</td>
		<td>비고</td>
	</tr>
	<% if (myMileageLog.FresultCount > 0) then %>
		<% for i=0 to myMileageLog.FResultCount - 1 %>
		<tr align="center" <% if (myMileageLog.FItemList(i).Fdeleteyn = "Y") then %>bgcolor="#EEEEEE" class="gray"<% else %>bgcolor="#FFFFFF"<% end if %>>
			<td height=25><%= userid %></td>
			<td><%= Left(myMileageLog.FItemList(i).FRegdate,10) %></td>
			<td><% if myMileageLog.FItemList(i).Fmileage >= 0 then %><font color="blue"><% else %><font color="red"><% end if %><%= myMileageLog.FItemList(i).Fstatusflagstring %></font></td>
			<td><%= myMileageLog.FItemList(i).Fjukyo %></td>
			<td align="right"><% if myMileageLog.FItemList(i).Fmileage >= 0 then %><font color="blue"><% else %><font color="red"><% end if %><%= FormatNumber(myMileageLog.FItemList(i).Fmileage, 0) %></font>&nbsp;&nbsp;</td>
			<td align="right">
				<%
				if (showtype = "A") then
					response.write FormatNumber(myMileageLog.FItemList(i).Fremain, 0)
				else
					response.write "--"
				end if
				%>
				&nbsp;&nbsp;
			</td>
			<td><%= myMileageLog.FItemList(i).Forderserial %></td>
			<td><%= myMileageLog.FItemList(i).Fdeleteyn %></td>
			<td>
				<% if C_ADMIN_AUTH and (myMileageLog.FItemList(i).Fstatusflag = "B") and (myMileageLog.FItemList(i).Fid <> "") and (myMileageLog.FItemList(i).Fid > "0") and (myMileageLog.FItemList(i).Fdeleteyn <> "Y") then %>
					<input type="button" class="button" value="삭제[관리자]" onClick="SubmitFormDelForce(<%= myMileageLog.FItemList(i).Fid %>)">
				<% end if %>
			</td>
		</tr>
		<% next %>
		<tr align="center" bgcolor="#FFFFFF">
			<form name="frmpage" method="get" action="" style="margin:0px;">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="userid" value="<%= userid %>">
			<input type="hidden" name="showtype" value="<%= showtype %>">
			<input type="hidden" name="showdelete" value="<%= showdelete %>">
			<input type="hidden" name="currpage" value="<%= currpage %>">
			<input type="hidden" name="showdetail" value="on">
			</form>
			<td colspan="15">
			<% if myMileageLog.HasPreScroll then %>
				<span class="list_link"><a href="javascript:gotoPage(<%= myMileageLog.StartScrollPage-1 %>)">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + myMileageLog.StartScrollPage to myMileageLog.StartScrollPage + myMileageLog.FScrollCount - 1 %>
				<% if (i > myMileageLog.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(myMileageLog.FCurrPage) then %>
				<span class="page_link"><font color="red"><b>[<%= i %>]</b></font></span>
				<% else %>
				<a href="javascript:gotoPage(<%= i %>)" class="list_link"><font color="#000000">[<%= i %>]</font></a>
				<% end if %>
			<% next %>
			<% if myMileageLog.HasNextScroll then %>
				<span class="list_link"><a href="javascript:gotoPage(<%= i %>)">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
			</td>
		</tr>

	<% else %>
		<tr align="center" bgcolor="#FFFFFF">
			<td colspan="15"> 검색된 내용이 없습니다.</td>
		</tr>
	<% end if %>
<% end if %>
</table>

<form name="frmAct" method="post" action="domodifymileage.asp" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="userid" value="<%= userid %>">
<input type="hidden" name="idx" value="">
</form>

<%
set myMileageLog = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
