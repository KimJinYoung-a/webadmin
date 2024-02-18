<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/resending_reportcls.asp"-->

<%
dim page, research

page	= req("page",1)


Dim ck_date1, ck_date2, finishDate1, finishDate2, regDate1, regDate2, gubun01, gubun02
dim isupchebeasong, divcd
dim finishuser, reguserid

research	= req("research","")
ck_date1	= req("ck_date1","")
ck_date2	= req("ck_date2","")

finishDate1	= req("finishDate1",Date())
finishDate2	= req("finishDate2",Date())
regDate1	= req("regDate1",Date())
regDate2	= req("regDate2",Date())

divCd	= req("divCd","")
gubun01	= req("gubun01","")
gubun02	= req("gubun02","")
isUpcheBeasong	= req("isUpcheBeasong","")
finishuser	= req("finishuser","")
reguserid	= req("reguserid","")

dim obj
set obj = new CReportMaster
obj.FPageSize = 50
obj.FCurrPage = page

obj.FRectFinishUser = finishuser
if (ck_date2="on") then
    obj.FRectRegStart = regDate1
    obj.FRectRegEnd = regDate2
end if
obj.FRectRegUserID = reguserid

obj.getCsListView2 finishDate1, finishDate2, divCd, gubun01, gubun02

dim i
%>
<script language='javascript'>
function goPage(page){
    frm.page.value = page;
    frm.submit();
}

function jsPopCal(fName,sName)
{
	var winCal;
	winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}


window.onload = function()
{
	document.title = "CS처리완료목록조회";
	getGubun02Options('<%=gubun02%>');
}

function getGubun02Options(gubun02)
{
	var f = document.frm;

	switch (f.gubun01.value)
	{
	case "C004":
		var arr =
		[
			["CD01", "단순변심"],
			["CD03", "재주문"],
			["CD04", "사이즈"],
			["CD05", "품절"],
			["CD99", "기타"]
		];
		break;
	case "C005":
		var arr =
		[
			["CE01", "상품불량"],
			["CE02", "상품불만족"],
			["CE03", "상품등록오류"],
			["CE04", "상품설명불량"],
			["CE05", "이벤트오등록"],
			["CE99", "기타"]
		];
		break;
	case "C006":
		var arr =
		[
			["CF01", "오발송"],
			["CF02", "상품파손"],
			["CF03", "구매상품누락"],
			["CF04", "사은품누락"],
			["CF05", "상품품절"],
			["CF06", "출고지연"],
			["CF99", "기타"]
		];
		break;
	case "C007":
		var arr =
		[
			["CG01", "배송지연"],
			["CG02", "택배사파손"],
			["CG03", "택배사분실"]
		];
		break;
	default:
		var arr = [];
		break;
	}

	f.gubun02.length = 1;
	for (i=0;i<arr.length ;i++ )
	{
		var newOpt = document.createElement("OPTION");
		newOpt.value = arr[i][0];
		newOpt.text  = arr[i][1];
		f.gubun02.options.add(newOpt);
	}

	if (gubun02)
		f.gubun02.value = gubun02;

}

</script>
<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get"   >
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">검색<br>조건</td>
		<td align="left">
			<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
			<tr>
				<td>
    				<input type="checkbox" name="ck_date1" checked disabled >
					처리완료일 : <input type="text" size="10" name="finishDate1" value="<%= finishDate1 %>" onClick="jsPopCal('frm','finishDate1');" class="text" style="cursor:hand;">
					~<input type="text" size="10" name="finishDate2" value="<%= finishDate2 %>" onClick="jsPopCal('frm','finishDate2');" class="text" style="cursor:hand;">
    				&nbsp;&nbsp;
					처리자 : <input type="text" class="text" size="15" name="finishuser" value="<%= finishuser %>">
                    &nbsp;&nbsp;
                    <input type="checkbox" name="ck_date2" <%= CHKIIF(ck_date2="on", "checked", "") %>>
					접수일 : <input type="text" size="10" name="regDate1" value="<%= regDate1 %>" onClick="jsPopCal('frm','regDate1');" class="text" style="cursor:hand;">
					~<input type="text" size="10" name="regDate2" value="<%= regDate2 %>" onClick="jsPopCal('frm','regDate2');" class="text" style="cursor:hand;">
    				&nbsp;&nbsp;
					접수자 : <input type="text" class="text" size="15" name="reguserid" value="<%= reguserid %>">
				</td>
			</tr>
			</table>
        </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">
		    <input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td align="left">
	        <table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
			<tr>
				<td>

    				구분:
    				<select name="divCd">
    				<option value="">전체
						<option value="A000" <% if (divcd = "A000") then response.write "selected" end if %>>맞교환출고</option>
						<option value="A001" <% if (divcd = "A001") then response.write "selected" end if %>>누락재발송</option>
						<option value="A002" <% if (divcd = "A002") then response.write "selected" end if %>>서비스발송</option>
						<option value="A003" <% if (divcd = "A003") then response.write "selected" end if %>>환불요청</option>
						<option value="A004" <% if (divcd = "A004") then response.write "selected" end if %>>반품접수(업체배송)</option>
						<option value="A005" <% if (divcd = "A005") then response.write "selected" end if %>>외부몰환불요청</option>
						<option value="A006" <% if (divcd = "A006") then response.write "selected" end if %>>출고시유의사항</option>
						<option value="A007" <% if (divcd = "A007") then response.write "selected" end if %>>신용카드/이체취소요청</option>
						<option value="A008" <% if (divcd = "A008") then response.write "selected" end if %>>주문취소</option>
						<option value="A009" <% if (divcd = "A009") then response.write "selected" end if %>>기타내역(메모)</option>
						<option value="A010" <% if (divcd = "A010") then response.write "selected" end if %>>회수신청(텐바이텐배송)</option>
						<option value="A011" <% if (divcd = "A011") then response.write "selected" end if %>>맞교환회수(텐바이텐배송)</option>
						<option value="A700" <% if (divcd = "A700") then response.write "selected" end if %>>업체기타정산</option>
						<option value="A900" <% if (divcd = "A900") then response.write "selected" end if %>>주문내역변경</option>
    				</select>
    				&nbsp;&nbsp;

    				사유구분 :
    				<select name="gubun01" onchange="getGubun02Options();">
	    				<option value="">전체</option>
	    				<option value="C004" <%=chkIIF(gubun01="C004","selected","")%>>공통</option>
	    				<option value="C005" <%=chkIIF(gubun01="C005","selected","")%>>상품관련</option>
	    				<option value="C006" <%=chkIIF(gubun01="C006","selected","")%>>물류관련</option>
	    				<option value="C007" <%=chkIIF(gubun01="C007","selected","")%>>택배사관련</option>
    				</select>
    				<select name="gubun02">
	    				<option value="">상세</option>
    				</select>


    		    </td>
    		</tr>
    		</table>
	    </td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->
<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td colspan="11" align="right">총 <%= obj.FTotalCount %> 건 <%= page %>/<%= obj.FTotalPage %> &nbsp;</td>
</tr>
<tr align="center" bgcolor="#E6E6E6">
	<!-- td width="60">ID</td -->
	<td width="150">구분</td>
	<td width="150">사유</td>
	<td width="100">주문번호</td>
	<td width="">접수제목</td>
	<td width="80">접수자</td>
	<td width="80">접수일</td>
    <td width="80">처리자</td>
    <td width="80">처리일</td>
</tr>
<% if obj.FResultCount<1 then %>
<tr bgcolor="#FFFFFF">
    <td align="center" colspan="11">
       [검색 결과가 없습니다.]
    </td>
</tr>
<% else %>
<% for i=0 To obj.FREsultCount -1 %>
<tr bgcolor="#FFFFFF">
    <!--td ><%= obj.FItemList(i).FID %></td -->
    <td align="center"><%= obj.FItemList(i).Fdivcd_Name %></td>
    <td align="center"><%= obj.FItemList(i).Fgubun01_Name %> / <%= obj.FItemList(i).Fgubun02_Name %></td>
    <td align="center"><a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= obj.FItemList(i).Forderserial %>');"><%= obj.FItemList(i).Forderserial %></a></td>
    <td align="left">&nbsp; <%= obj.FItemList(i).Ftitle %></td>
    <td align="center"><%= obj.FItemList(i).Freguserid %></td>
    <td ><%= Left(obj.FItemList(i).Fregdate,10) %></td>
	<td align="center"><%= obj.FItemList(i).Ffinishuser %></td>
    <td ><%= Left(obj.FItemList(i).Ffinishdate,10) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="11" align="center">
        <!-- 페이지 시작 -->
			<% sbDisplayPaging "page="&page, obj.FTotalCount, obj.FPageSize, 10%>
    	<!-- 페이지 끝 -->
    </td>
</tr>
<% end if %>
</table>


<%
set obj = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
