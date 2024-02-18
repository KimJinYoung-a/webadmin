<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 정산 매출로그 데이터 비교
' Hieditor : 2018.04.23 정윤정 생성 
'###########################################################
Server.ScriptTimeOut = 180
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/outmall_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<!-- #include virtual="/lib/classes/extjungsan/extjungsandiffcls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim sellsite,rdost
dim searchfield,searchtext 
dim chulgoDate_yyyy1,chulgoDate_yyyy2,chulgoDate_mm1,chulgoDate_mm2,chulgoDate_dd1,chulgoDate_dd2
dim confirmDate_yyyy1,confirmDate_yyyy2,confirmDate_mm1,confirmDate_mm2,confirmDate_dd1,confirmDate_dd2
dim chulgoDate_fromDate, chulgoDate_toDate,confirmDate_fromDate, confirmDate_toDate, tmpDate
dim iCurrpage, iTotCnt, iTotPage,iPageSize,iPerCnt
dim arrList, intLoop
dim chkErr
dim extMeachul, logMeachul
  iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
  sellsite = requestCheckVar(Request("sellsite"),10)  
  rdost = requestCheckVar(Request("rdost"),1)
  chkErr = requestCheckVar(Request("chkErr"),1)
  searchfield = requestCheckVar(Request("searchfield"),30)
  searchtext = requestCheckVar(Request("searchtext"),120)
  iPageSize = requestCheckVar(Request("ips"),10)
	chulgoDate_yyyy1   = 	requestCheckvar(request("chulgoDate_yyyy1"),4)
	chulgoDate_mm1     = requestCheckvar(request("chulgoDate_mm1"),2)
	chulgoDate_dd1     = requestCheckvar(request("chulgoDate_dd1"),2)
	chulgoDate_yyyy2   = requestCheckvar(request("chulgoDate_yyyy2"),4)
	chulgoDate_mm2     = requestCheckvar(request("chulgoDate_mm2"),2)
	chulgoDate_dd2     = requestCheckvar(request("chulgoDate_dd2"),2)
	confirmDate_yyyy1   = 	requestCheckvar(request("confirmDate_yyyy1"),4)
	confirmDate_mm1     = requestCheckvar(request("confirmDate_mm1"),2)
	confirmDate_dd1     = requestCheckvar(request("confirmDate_dd1"),2)
	confirmDate_yyyy2   = requestCheckvar(request("confirmDate_yyyy2"),4)
	confirmDate_mm2     = requestCheckvar(request("confirmDate_mm2"),2)
	confirmDate_dd2     = requestCheckvar(request("confirmDate_dd2"),2)
if sellsite ="" then sellsite ="ssg"
if rdost	="" then rdost="1"
IF iCurrpage = "" THEN		iCurrpage = 1
if iPageSize ="" then iPageSize = 20
		iPerCnt = 10		'보여지는 페이지 간격
	
if (chulgoDate_yyyy1="") then
		if rdost="1" then
		chulgoDate_fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 2), 1)
		else
		chulgoDate_fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
		end if
	chulgoDate_toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) ), 1)
 
	chulgoDate_yyyy1 = Cstr(Year(chulgoDate_fromDate))
	chulgoDate_mm1 = Cstr(Month(chulgoDate_fromDate))
	chulgoDate_dd1 = Cstr(day(chulgoDate_fromDate))

	tmpDate = DateAdd("d", -1, chulgoDate_toDate)
	chulgoDate_yyyy2 = Cstr(Year(tmpDate))
	chulgoDate_mm2 = Cstr(Month(tmpDate))
	chulgoDate_dd2 = Cstr(day(tmpDate))
else
	chulgoDate_fromDate = DateSerial(chulgoDate_yyyy1, chulgoDate_mm1, chulgoDate_dd1)
	chulgoDate_toDate = DateSerial(chulgoDate_yyyy2, chulgoDate_mm2, chulgoDate_dd2+1)
end if


if (confirmDate_yyyy1="") then 
	confirmDate_fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1) 
	confirmDate_toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) ), 1)
 
	confirmDate_yyyy1 = Cstr(Year(confirmDate_fromDate))
	confirmDate_mm1 = Cstr(Month(confirmDate_fromDate))
	confirmDate_dd1 = Cstr(day(confirmDate_fromDate))

	tmpDate = DateAdd("d", -1, confirmDate_toDate)
	confirmDate_yyyy2 = Cstr(Year(tmpDate))
	confirmDate_mm2 = Cstr(Month(tmpDate))
	confirmDate_dd2 = Cstr(day(tmpDate))
else
	confirmDate_fromDate = DateSerial(confirmDate_yyyy1, confirmDate_mm1, confirmDate_dd1)
	confirmDate_toDate = DateSerial(confirmDate_yyyy2, confirmDate_mm2, confirmDate_dd2+1)
end if

dim cEJDiff
 set cEJDiff = new CextJungsanDiff
 cEJDiff.FCPage = iCurrpage
 cEJDiff.FPSize = iPageSize
 cEJDiff.FCGFDate = chulgoDate_fromDate
 cEJDiff.FCGTDate = chulgoDate_toDate
 cEJDiff.FCFFDate = confirmDate_fromDate
 cEJDiff.FCFTDate = confirmDate_toDate
 cEJDiff.FSellsite = sellsite
 cEJDiff.FRectST = rdost
 cEJDiff.FRectErr = chkErr
 if rdost ="1" then
 arrList =cEJDiff.fnGetextJsDiffList
else
 arrList =cEJDiff.fnGetlogJsDiffList
end if
 iTotCnt = cEJDiff.FTotCnt
 extMeachul = cEJDiff.FextMeachul
 logMeachul = cEJDiff.FlogMeachul
 set cEJDiff = nothing
 
 if extMeachul ="" or isNull(extMeachul) then extMeachul =0
 if logMeachul ="" or isNull(logMeachul) then logMeachul =0
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
	function jsSearh(){
		$("#btnSubmit").prop("disabled", true); 
		document.frm.submit(); 
		}
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="page" value="">
		<input type="hidden" name="research" value="on">
		<tr  bgcolor="#FFFFFF" >
			<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
			<td align="left"> 
				매출처 : <% fnGetOptOutMall sellsite %>
				 
				&nbsp;&nbsp;
				* 검색조건 :
				<select class="select" name="searchfield">
					<option value=""></option>
					<option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %> >주문번호</option>
				</select>
				<input type="text" class="text" name="searchtext" value="<%= searchtext %>">				
			</td>
			<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
				<input type="button" id="btnSubmit" class="button_s" value="검색" onClick="jsSearh();">
			</td>
		</tr>
		<tr   bgcolor="#FFFFFF" >
			<td align="left"> 
				<input type="radio" name="rdost" class="radio" value="1" <%if rdost ="1" then%>checked<%end if%> > 제휴몰정산 기준 
				<input type="radio" name="rdost" class="radio" value="2" <%if rdost ="2" then%>checked<%end if%>> 매출로그기준
				&nbsp;&nbsp;
 				제휴구매확정일:
			  <% DrawDateBoxdynamic confirmDate_yyyy1, "confirmDate_yyyy1", confirmDate_yyyy2, "confirmDate_yyyy2", confirmDate_mm1, "confirmDate_mm1", confirmDate_mm2, "confirmDate_mm2", confirmDate_dd1, "confirmDate_dd1", confirmDate_dd2, "confirmDate_dd2" %>
			  &nbsp;
				출고일자 :
				<% DrawDateBoxdynamic chulgoDate_yyyy1, "chulgoDate_yyyy1", chulgoDate_yyyy2, "chulgoDate_yyyy2", chulgoDate_mm1, "chulgoDate_mm1", chulgoDate_mm2, "chulgoDate_mm2", chulgoDate_dd1, "chulgoDate_dd1", chulgoDate_dd2, "chulgoDate_dd2" %>
				&nbsp;&nbsp;
			<input type="checkbox" name="chkErr" value="Y" <%if chkErr ="Y" then%> checked<%end if%>>미매칭
			</td>
		</tr>
		<tr  bgcolor="#FFFFFF" >
			<td> 
			
				행표시 :
				<select class="select" name="ips">
					<option value="20" <% if (iPageSize = "20") then %>selected<% end if %> >20</option>
					<option value="100" <% if (iPageSize = "100") then %>selected<% end if %> >100</option>
					<option value="1000" <% if (iPageSize = "1000") then %>selected<% end if %> >1000</option>
					<option value="3000" <% if (iPageSize = "3000") then %>selected<% end if %> >3000</option>
				</select> 
			</td>
		</tr>
	</form>
</table>
<!-- 검색 끝 -->
<p style="padding-top:10px;"></p>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" >
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30">
		검색결과 : <b><%=iTotCnt%></b>
		&nbsp;
		페이지 : <b> /  </b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td rowspan="2">제휴몰</td>
	<td colspan="5">제휴주문</td>
	<td  colspan="5">매출로그</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">	
	<td>주문번호</td>
	<td>상품코드</td>
	<td>옵션코드</td>
	<td>매출금액</td>
	<td></td>
	
	<td>주문번호</td>
	<td>상품코드</td>
	<td>옵션코드</td>
	<td>매출금액</td>
</tr>
<%if isArray(arrList) then%>
	<tr bgcolor="<%=adminColor("pink")%>">
		<td></td>
		<td colspan="3"></td>
		<td align="right"><%=formatnumber(extMeachul,0)%></td>
		<td></td>
		<td colspan="3"></td>
		<td align="right"><%=formatnumber(logMeachul,0)%></td>
		<td></td>
	</tr>
	<%
		for intLoop = 0 To uBound(arrList,2)
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td><%=arrList(0,intLoop)%></td>
		
		<td><%=arrList(1,intLoop)%></td>
		<td><%=arrList(2,intLoop)%></td>
		<td><%=arrList(3,intLoop)%></td>
		<td align="right"><%=arrList(7,intLoop)%></td>
		<td><%=arrList(4,intLoop)%></td>
		
		<td><%=arrList(9,intLoop)%></td>
		<td><%=arrList(10,intLoop)%></td>
		<td><%=arrList(11,intLoop)%></td>
		<td align="right"><%=arrList(13,intLoop)%></td>
		<td></td>
	</tr>
<%next %>
<%end if%>
</table>
<!-- 페이징처리 --> 
<table width="100%" cellpadding="10">
	<tr>
		<td align="center">  
 			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->