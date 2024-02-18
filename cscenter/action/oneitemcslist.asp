<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<%

function CurrStateName(byval v)
	if v < "B006" then
		CurrStateName="접수"
	elseif v = "B006" then
		CurrStateName="업체처리완료"
	elseif v = "B007" then
		CurrStateName="처리완료"
	else
		CurrStateName = v
	end if
end function

function CurrStateColor(byval v)
	if v < "B006" then
		CurrStateColor="#000000"
	elseif v = "B006" then
		CurrStateColor="#000000"
	elseif v = "B007" then
		CurrStateColor="green"
	else
		CurrStateColor = "gray"
	end if
end function

Const MaxRowSize = 1000
dim itemid, itemoption, itemgubun, sitename, gubun01
dim currstate

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate
dim premonthdate
dim datetype, divcd, badOnly

nowdate         = Left(CStr(now()),10)
premonthdate    = DateAdd("d",-14,nowdate)

itemid = requestCheckvar(request("itemid"),10)
itemoption = requestCheckvar(request("itemoption"),4)
currstate = requestCheckvar(request("currstate"),2)

yyyy1   = requestCheckvar(request("yyyy1"),4)
mm1     = requestCheckvar(request("mm1"),2)
dd1     = requestCheckvar(request("dd1"),2)
yyyy2   = requestCheckvar(request("yyyy2"),4)
mm2     = requestCheckvar(request("mm2"),2)
dd2     = requestCheckvar(request("dd2"),2)
datetype = requestCheckvar(request("datetype"),6)
divcd 	= requestCheckvar(request("divcd"),4)
badOnly	= requestCheckvar(request("badOnly"),1)
sitename 	= requestCheckvar(request("sitename"),32)
gubun01 	= requestCheckvar(request("gubun01"),4)

if (yyyy1="") then
	yyyy1 = Left(premonthdate,4)
	mm1   = Mid(premonthdate,6,2)
	dd1   = Mid(premonthdate,9,2)

	nowdate = Left(CStr(now()),10)
	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
else
	nowdate = Left(CStr(DateSerial(yyyy1 , mm1 , dd1)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

if (datetype="") then datetype="reg"
if (itemgubun="") then itemgubun="10"

'상품코드 유효성 검사(2008.08.05;허진원)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

dim sqlStr, RowArr


sqlStr = " select top " & CStr(MaxRowSize) & " "
sqlStr = sqlStr + " 	m.orderserial, m.ipkumdiv "
sqlStr = sqlStr + " 	, (case when a.divcd in ('A000', 'A100', 'A002') then d.confirmitemno * -1 else d.confirmitemno end) as sm "
sqlStr = sqlStr + " 	, m.buyname, m.buyemail, m.buyhp, m.buyphone, m.reqname, m.reqhp, m.reqphone "
sqlStr = sqlStr + " 	, d.itemoptionname, IsNULL(a.currstate, 'B001') as currstate, m.sitename "
sqlStr = sqlStr + " 	, a.finishdate as beasongdate, m.userid, a.divcd "
sqlStr = sqlStr + " 	, C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name "
sqlStr = sqlStr + " from "
sqlStr = sqlStr + " 	db_cs.dbo.tbl_new_as_list a "
sqlStr = sqlStr + " 	join db_cs.dbo.tbl_new_as_detail d "
sqlStr = sqlStr + " 	on "
sqlStr = sqlStr + " 		a.id = d.masterid "
sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_master m "
sqlStr = sqlStr + " 	on "
sqlStr = sqlStr + " 		m.orderserial=a.orderserial "
sqlStr = sqlStr + "		Left Join [db_cs].[dbo].tbl_cs_comm_code C1"
sqlStr = sqlStr + "			on A.divcd=C1.comm_cd"
sqlStr = sqlStr + "		Left Join [db_cs].[dbo].tbl_cs_comm_code C2"
sqlStr = sqlStr + "			on A.gubun01=C2.comm_cd"
sqlStr = sqlStr + "		Left Join [db_cs].[dbo].tbl_cs_comm_code C3"
sqlStr = sqlStr + "			on A.gubun02=C3.comm_cd"
sqlStr = sqlStr + " where "
sqlStr = sqlStr + " 	1 = 1 "
sqlStr = sqlStr + " 	and a.deleteyn <> 'Y' "
sqlStr = sqlStr + " 	and a.divcd not in ('A008', 'A006') "
sqlStr = sqlStr + " 	and d.itemid=" + CStr(itemid)

if itemoption<>"" then
    sqlStr = sqlStr + " and d.itemoption='" + CStr(itemoption) + "'"
end if

if divcd<>"" then
    sqlStr = sqlStr + " and a.divcd = '" + CStr(divcd) + "'"
end if

if gubun01<>"" then
	sqlStr = sqlStr + " and a.gubun01 = '" + CStr(gubun01) + "'"
end if

if badOnly<>"" then
    '// 상품불량
    sqlStr = sqlStr + " and (  "
    sqlStr = sqlStr + "    ((a.gubun01 = 'C004') and (a.gubun02 = 'CD01')) "
    sqlStr = sqlStr + " 	or "
    sqlStr = sqlStr + "    ((a.gubun01 = 'C005') and (a.gubun02 = 'CE01'))"
    sqlStr = sqlStr + " ) "
end if

if (datetype="reg") then
    sqlStr = sqlStr + " and a.regdate >= '" + yyyy1 + "-" + mm1 + "-" + dd1 + "'"
    sqlStr = sqlStr + " and a.regdate < '" + searchnextdate + "'"
elseif (datetype="finish") then
    sqlStr = sqlStr + " and a.currstate = 'B007' and a.finishdate >= '" + yyyy1 + "-" + mm1 + "-" + dd1 + "'"
    sqlStr = sqlStr + " and a.currstate = 'B007' and a.finishdate < '" + searchnextdate + "'"
end if

if sitename<>"" then
	sqlStr = sqlStr + " and m.sitename = '" + sitename + "'"
end if

if currstate="availall" then   '전체
	'//
elseif currstate="reg" then	'접수
	sqlStr = sqlStr + " and a.currstate < 'B007' "
elseif currstate="finish" then	'처리완료
	sqlStr = sqlStr + " and a.currstate = 'B007' "
end if
sqlStr = sqlStr + " order by a.currstate , a.id "


'response.write sqlStr
'response.end



if (itemid<>"") then
    rsget.Open sqlStr,dbget,1
    if not rsget.Eof then
        RowArr = rsget.getRows
    end if
    rsget.Close
end if

dim RowCount, jumuncnt
RowCount = 0
jumuncnt = 0
if IsArray(RowArr) then
    RowCount = Ubound(RowArr,2)
    jumuncnt = RowCount + 1
end if

dim totno, i
totno = 0


dim oitem
set oitem = new CItemInfo
oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItemInfo
end if


dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if
%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>
/*
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}
*/

function editCsErr(comp){
    var frm = comp.form;
    if ((frm.mayDay.value.length!=10)&&(frm.mayDay.value.length!=7)){
        alert('YYYY-MM 또는 YYYY-MM-DD를 입력하세요.');
        return;
    }

    if (frm.errcsno.value.length<1){
        alert('오차 수량을 입력하세요.');
        return;
    }

    if (confirm('CS 오차를 수정하시겠습니까?')){
        frm.submit();
    }
}

$(function(){
	//$("select[name=sitename]").children('option:first').remove();
	$("select[name=sitename]")
		.prepend('<option value="10x10">10x10</option>')
		.prepend('<option value="">전체</option>')
		.val("<%=sitename%>").prop("selected", true);
});
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">

	<tr height="25" bgcolor="#FFFFFF">
	    <td align="center" width="50" bgcolor="#EEEEEE" rowspan="2">검색<br>조건</td>
        <td>
        	* 아이템ID : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="11" maxlength="16">&nbsp;&nbsp;

        	<% if oitemoption.FResultCount>0 then %>
			&nbsp;

			* 옵션선택 :
			<select class="select" name="itemoption">
				<option  value="">----
				<% for i=0 to oitemoption.FResultCount-1 %>
				<option value="<%= oitemoption.FITemList(i).FItemOption %>" <% if itemoption=oitemoption.FITemList(i).FItemOption then response.write "selected" %> ><%= oitemoption.FITemList(i).FOptionName %>
				<% next %>
				</select>
			<% end if %>

			&nbsp;
			* 검색기간 :
			<select class="select" name="datetype">
			    <option value="reg" <%= chkIIF(datetype="reg","selected","") %> >접수일
			    <option value="finish" <%= chkIIF(datetype="finish","selected","") %> >처리일
			</select>
			<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>

            &nbsp;
            * 구분:
            <select class="select" name="divcd">
            	<option value="">전체</option>
            	<option value="">-------------------------</option>
				<option value="A000" <% if (divcd = "A000") then response.write "selected" end if %>>교환출고</option>
				<option value="A100" <% if (divcd = "A100") then response.write "selected" end if %>>교환출고(상품변경)</option>
				<option value="">-------------------------</option>
				<option value="A004" <% if (divcd = "A004") then response.write "selected" end if %>>반품접수(업배)</option>
				<option value="">-------------------------</option>
				<option value="A001" <% if (divcd = "A001") then response.write "selected" end if %>>누락재발송</option>
				<option value="A002" <% if (divcd = "A002") then response.write "selected" end if %>>서비스발송</option>
				<option value="A200" <% if (divcd = "A200") then response.write "selected" end if %>>기타회수</option>
				<option value="">-------------------------</option>
				<option value="A003" <% if (divcd = "A003") then response.write "selected" end if %>>환불요청</option>
				<option value="A005" <% if (divcd = "A005") then response.write "selected" end if %>>외부몰환불요청</option>
				<option value="A007" <% if (divcd = "A007") then response.write "selected" end if %>>신용카드/이체취소요청</option>
				<option value="A700" <% if (divcd = "A700") then response.write "selected" end if %>>업체기타정산</option>
				<option value="A999" <% if (divcd = "A999") then response.write "selected" end if %>>고객추가결제</option>
				<option value="">-------------------------</option>
				<option value="A060" <% if (divcd = "A060") then response.write "selected" end if %>>업체긴급문의</option>
				<option value="A006" <% if (divcd = "A006") then response.write "selected" end if %>>출고시유의사항</option>
				<option value="A008" <% if (divcd = "A008") then response.write "selected" end if %>>주문취소</option>
				<option value="A009" <% if (divcd = "A009") then response.write "selected" end if %>>기타내역(메모)</option>
				<option value="A900" <% if (divcd = "A900") then response.write "selected" end if %>>주문내역변경</option>
				<option value="">-------------------------</option>
				<option value="A010" <% if (divcd = "A010") then response.write "selected" end if %>>회수신청(텐배)</option>
				<option value="">-------------------------</option>
				<option value="A011" <% if (divcd = "A011") then response.write "selected" end if %>>교환회수(텐배)</option>
				<option value="A012" <% if (divcd = "A012") then response.write "selected" end if %>>교환회수(업배)</option>
				<option value="A111" <% if (divcd = "A111") then response.write "selected" end if %>>교환회수(상품변경,텐배)</option>
				<option value="A112" <% if (divcd = "A112") then response.write "selected" end if %>>교환회수(상품변경,업배)</option>
            </select>
        </td>
        <td align="center" width="50" bgcolor="#EEEEEE" rowspan="2">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
        <td>
			* CS상태 :
			<select class="select" name="currstate">
				<option value="availall" <% if currstate="availall" then response.write "selected" %>>전체
				<option value="reg" <% if currstate="reg" then response.write "selected" %>>접수
				<option value="finish" <% if currstate="finish" then response.write "selected" %>>처리완료
			</select>
			&nbsp;

			* 사이트 : <% DrawSelectExtSiteName "sitename", sitename %>
            &nbsp;

			* 사유 : <% drawCSCommCodeBox 1,"Z020","gubun01",gubun01,"" %>
            &nbsp;

            <label><input type="checkbox" name="badOnly" value="Y" <%= CHKIIF(badOnly="Y", "checked", "") %> >
            상품불량만</label>
        </td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p />

* 최대 <%= MaxRowSize %>건 까지만 검색됩니다.

<p />

<% if oitem.FResultCount>0 then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=<%= 6 + oitemoption.FResultCount -1 %> width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
      	<td width="60" bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
      	<td width="300">
      		10 <b><%= CHKIIF(oitem.FOneItem.FItemID>=1000000,Format00(8,oitem.FOneItem.FItemID),Format00(6,oitem.FOneItem.FItemID)) %></b> <%= itemoption %>
      		&nbsp;
      		<!--
      		<input type="button" value="수정" onclick="PopItemSellEdit('<%= itemid %>');">
      		-->
      	</td>
      	<td width="60" bgcolor="<%= adminColor("tabletop") %>">전시여부</td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FDispyn) %>"><%= oitem.FOneItem.FDispyn %></font></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">브랜드ID</td>
      	<td><%= oitem.FOneItem.FMakerid %></td>
      	<td bgcolor="<%= adminColor("tabletop") %>">판매여부</td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FSellyn) %>"><%= oitem.FOneItem.FSellyn %></font></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">상품명</td>
      	<td><%= oitem.FOneItem.FItemName %></td>
      	<td bgcolor="<%= adminColor("tabletop") %>">사용여부</td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FIsUsing) %>"><%= oitem.FOneItem.FIsUsing %></font></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">판매가</td>
      	<td>
      		<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
      		&nbsp;&nbsp;
      		<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
      	    <% if oitem.FOneItem.FSellcash<>0 then %>
			<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
			<% end if %>
			&nbsp;&nbsp;
			<!-- 할인여부/쿠폰적용여부 -->
			<% if (oitem.FOneItem.FSailYn="Y") then %>
			    <font color=red>
			    <% if (oitem.FOneItem.Forgprice<>0) then %>
			        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %> %
			    <% end if %>
			     할인
			    </font>
			<% end if %>

			<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>

			    <font color=green><%= oitem.FOneItem.GetCouponDiscountStr %> 쿠폰
			    (<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>)</font>
			<% end if %>

      	</td>
      	<td bgcolor="<%= adminColor("tabletop") %>">단종여부</td>
      	<td colspan=2>
      		<% if oitem.FOneItem.Fdanjongyn="Y" then %>
			<font color="#33CC33">단종</font>
			<% elseif oitem.FOneItem.Fdanjongyn="S" then %>
			<font color="#33CC33">일시품절</font>
			<% else %>
			생산중
			<% end if %>
		</td>
    </tr>

    <% if oitemoption.FResultCount>1 then %>
	    <!-- 옵션이 있는경우 -->
	    <% for i=0 to oitemoption.FResultCount -1 %>
		    <% if oitemoption.FITemList(i).FOptIsUsing<>"Y" then %>
		    <tr bgcolor="#FFFFFF">
		      	<td bgcolor="<%= adminColor("tabletop") %>"><font color="#AAAAAA">옵션명 :</font></td>
		      	<td><font color="#AAAAAA"><%= oitemoption.FITemList(i).FOptionName %></font></td>
		      	<td bgcolor="<%= adminColor("tabletop") %>"><font color="#AAAAAA">한정여부 : </font></td>
		      	<td><font color="#AAAAAA"><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</font></td>
		      	<td>한정 비교재고 (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
		    </tr>
		    <% else %>

		    <% if oitemoption.FITemList(i).Fitemoption=itemoption then %>
		    <tr bgcolor="#EEEEEE">
		    <% else %>
		    <tr bgcolor="#FFFFFF">
		    <% end if %>
		      	<td>옵션명</td>
		      	<td><%= oitemoption.FITemList(i).FOptionName %></td>
		      	<td>한정여부</td>
		      	<td><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
		      	<td>한정 비교재고 (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
		    </tr>
		    <% end if %>
	    <% next %>
    <% else %>
    	<tr bgcolor="#FFFFFF">
	      	<td bgcolor="<%= adminColor("tabletop") %>">옵션명</td>
	      	<td>-</td>
	      	<td bgcolor="<%= adminColor("tabletop") %>">한정여부</td>
	      	<td><font color="<%= ynColor(oitem.FOneItem.Flimityn) %>"><%= oitem.FOneItem.Flimityn %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
	      	<td>한정 비교재고 (<b><%= oitem.FOneItem.GetLimitStockNo %></b>)</td>
	    </tr>
    <% end if %>
</table>
<% end if %>
<p>


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80">주문번호</td>
		<td >Site</td>
		<td >구분</td>
		<td >사유</td>
		<td width="60">상태</td>
		<td width="40">수량</td>
		<td>옵션명</td>
		<td>회원ID</td>
		<td >구매자</td>
		<td>수령인</td>
		<td width="140">출고일</td>
	</tr>
<%
if IsArray(RowArr) then
	for i=0 to RowCount
%>

	<tr align="center" bgcolor="#FFFFFF">
		<td><%= RowArr(0,i) %></td>
		<td><%= RowArr(12,i) %></td>
		<td><%= RowArr(16,i) %></td>
		<td><%= RowArr(17,i) & ">" & RowArr(18,i) %></td>
		<td><font color="<%= CurrStateColor(RowArr(11,i)) %>"><%= CurrStateName(RowArr(11,i)) %></font></td>
		<td><%= RowArr(2,i) %></td>
		<td><%= DdotFormat(RowArr(10,i),10) %></td>
		<td><%= printUserId(RowArr(14,i),2,"**") %></td>
		<td><%= RowArr(3,i) %></td>
		<td><%= RowArr(7,i) %></td>
		<td><%= RowArr(13,i) %></td>
	</tr>
<%
			totno = totno + RowArr(2,i)
    next
end if

%>
    <tr height="25" bgcolor="#FFFFFF">
        <td align="right" colspan="11">총상품수 <%= totno %> 개 / 총주문건수 <%= jumuncnt %> 건</td>
    </tr>
</table>
<% if (C_ADMIN_AUTH) then %>
<%
Dim StockBaseDate : StockBaseDate = Left(CStr(dateadd("m",-1,now())),7)+"-01"
Dim mayDay : mayDay=yyyy1+"-"+mm1+"-"+dd1
Dim isDailyLog : isDailyLog = CDate(mayDay)>=Cdate(StockBaseDate)
Dim errcsno

IF (isDailyLog) then
    sqlStr = "select yyyymmdd,itemgubun,itemid,itemoption,errcsno"
    sqlStr = sqlStr & " from db_summary.dbo.tbl_daily_logisstock_summary s"
    sqlStr = sqlStr & " where yyyymmdd='"&mayDay&"'"
    sqlStr = sqlStr & " and itemgubun='"&itemgubun&"'"
    sqlStr = sqlStr & " and itemid='"&itemid&"'"
    sqlStr = sqlStr & " and itemoption='"&itemoption&"'"

ELSE
    sqlStr = "select yyyymm,itemgubun,itemid,itemoption,errcsno"
    sqlStr = sqlStr & " from db_summary.dbo.tbl_monthly_logisstock_summary s"
    sqlStr = sqlStr & " where yyyymm='"&Left(mayDay,7)&"'"
    sqlStr = sqlStr & " and itemgubun='"&itemgubun&"'"
    sqlStr = sqlStr & " and itemid='"&itemid&"'"
    sqlStr = sqlStr & " and itemoption='"&itemoption&"'"
END IF

    rsget.Open sqlStr,dbget,1
    if not rsget.Eof then
        errcsno = rsget("errcsno")
    end if
    rsget.Close

%>
<p>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmAct" method="post" action="/admin/stock/stockrefresh_process.asp">
<input type="hidden" name="mode" value="editCsErr">
<input type="hidden" name="itemgubun" value="<%=itemgubun%>">
<input type="hidden" name="itemid" value="<%=itemid%>">
<input type="hidden" name="itemoption" value="<%=itemoption%>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>날짜</td>
    <td>CS오차</td>
    <td>수량</td>
    <td></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
    <td>
        <% if (isDailyLog) then %>
        <input type="text" name="mayDay" value="<%= mayDay %>" size="10" maxlength="10">
        <% else %>
        <input type="text" name="mayDay" value="<%= Left(mayDay,7) %>" size="7" maxlength="7">
        <% end if %>
    </td>
    <td><%= errcsno %></td>
    <td>
    <input type="text" name="errcsno" value="<%= errcsno %>" size="4" maxlength="9">
    </td>
    <td><input type="button" value="수정" onclick="editCsErr(this)"></td>
</tr>
</form>
</table>
<% end if %>
<%
set oitem = Nothing
set oitemoption = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
