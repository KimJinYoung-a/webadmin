<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  온라인 3pl 주문
' History : 2017.03.02 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/db_TPLopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/chulgoclass/chulgoclass.asp" -->

<%
dim shopid , yyyy1,mm1,dd1,yyyy2,mm2,dd2, yyyymmdd1,yyymmdd2,i, reloading
dim fromDate,toDate, page, tplcompanyid, oldlist
dim totsuplyprice , totprofit , totprofit2 , custa ,makerid
	yyyy1 = requestcheckvar(getNumeric(request("yyyy1")),4)
	mm1 = requestcheckvar(getNumeric(request("mm1")),2)
	dd1 = requestcheckvar(getNumeric(request("dd1")),2)
	yyyy2 = requestcheckvar(getNumeric(request("yyyy2")),4)
	mm2 = requestcheckvar(getNumeric(request("mm2")),2)
	dd2 = requestcheckvar(getNumeric(request("dd2")),2)
	page = requestcheckvar(getNumeric(request("page")),10)
	tplcompanyid = requestcheckvar(request("tplcompanyid"),32)
	reloading = requestcheckvar(request("reloading"),2)
	oldlist = requestcheckvar(request("oldlist"),2)

if page = "" then page = 1
if reloading="" and tplcompanyid = "" then tplcompanyid="tplithinkso"

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

if datediff("d",fromDate,toDate)>31 then
	response.write "<script>alert('검색기간은 한달 이내로 검색해주세요.');history.back();</script>"
	response.end
end if

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

dim oorder
set oorder = new Cchulgoitemlist
	oorder.FPageSize = 3000
	oorder.FCurrPage = page
	oorder.FRectStartdate = fromDate
	oorder.FRectEnddate = toDate
	oorder.FRecttplcompanyid = tplcompanyid
	oorder.FRectOldJumun = oldlist
if tplcompanyid="tpliconic" or tplcompanyid="tplmmmg" or tplcompanyid="tplparagon" or tplcompanyid="tplclass101" then
	oorder.fETC3plculgolist
else
	oorder.fonline3plculgolist
end if

%>

<script type="text/javascript">

function frmsubmit(page){
	frm.page.value=page;
	frm.target='';
	frm.action='';
	frm.submit();
}

function exceldown3pl(){
	frm.target='view';
	frm.action='/admin/chulgo/3pl_chulgo_on_excel.asp';
	frm.submit();
	frm.target='';
	frm.action='';
}

</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="reloading" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* 기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역
				&nbsp;
				* 매출처 : <% drawSelectBox3plcompany "tplcompanyid", tplcompanyid, "" %>
			</td>
		</tr>
		</table>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
	</td>
</tr>
</table>
<!-- 표 상단바 끝-->

</form>

<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    </td>
    <td align="right">
    	<input type="button" value="엑셀다운" onclick="exceldown3pl();" class="button">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oorder.FTotalCount %></b>
		<!--&nbsp;페이지 : <b><%= page %>/ <%= oorder.FTotalPage %></b>-->
		&nbsp;&nbsp;※최대 3천건까지 노출 됩니다.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>날짜</td>
	<td>사이트</td>
	<td>총주문건수</td>
	<td>주문건수(+)</td>
	<td>주문건수(-)</td>
	<td>상품수량</td>
	<td>합포건수</td>
	<td>비고</td>
</tr>
<% if oorder.FresultCount > 0 then %>
	<% for i=0 to oorder.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='#FFFFFF';>
		<td>
			<%= oorder.FItemList(i).fyyyymmdd %>
		</td>
		<td>
			<%= oorder.FItemList(i).fsitename %>
		</td>
		<td align="right">
			<%= oorder.FItemList(i).fordercnt %>
		</td>
		<td align="right">
			<%= oorder.FItemList(i).forderpluscnt %>
		</td>
		<td align="right">
			<%= oorder.FItemList(i).forderminuscnt %>
		</td>
		<td align="right">
			<%= oorder.FItemList(i).fitemcnt %>
		</td>
		<td align="right">
			<%= oorder.FItemList(i).fitemcnt2 %>
		</td>
		<td></td>
	</tr>
	<% next %>

	<!--<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oorder.HasPreScroll then %>
				<span class="list_link"><a href="#" onclick="frmsubmit('<%= oorder.StartScrollPage-1 %>'); return false;">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oorder.StartScrollPage to oorder.StartScrollPage + oorder.FScrollCount - 1 %>
				<% if (i > oorder.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oorder.FCurrPage) then %>
					<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
					<a href="#" onclick="frmsubmit('<%= i %>'); return false;" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oorder.HasNextScroll then %>
				<span class="list_link"><a href="#" onclick="frmsubmit('<%= i %>'); return false;">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>-->
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>

<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>

<%
set oorder = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/db_TPLclose.asp" -->
