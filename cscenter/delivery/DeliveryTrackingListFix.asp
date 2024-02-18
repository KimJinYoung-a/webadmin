<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 가출고 리스트
' Hieditor : 2018.03.07 이상구 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_deliverycls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim page, i, j, k
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, basedate, fromdate, todate
dim songjangdiv, makerid, orderserial, checkCnt
dim delayDelivOnly
dim research

page     = requestCheckVar(request("page"),10)
yyyy1   = requestCheckVar(request("yyyy1"),4)
mm1		= requestCheckVar(request("mm1"),2)
dd1		= requestCheckVar(request("dd1"),2)
yyyy2	= requestCheckVar(request("yyyy2"),4)
mm2		= requestCheckVar(request("mm2"),2)
dd2		= requestCheckVar(request("dd2"),2)
songjangdiv		= requestCheckVar(request("songjangdiv"),3)
delayDelivOnly	= requestCheckVar(request("delayDelivOnly"),3)
research		= requestCheckVar(request("research"),3)
makerid			= requestCheckVar(request("makerid"),32)
orderserial		= requestCheckVar(request("orderserial"),32)
checkCnt		= requestCheckVar(request("checkCnt"),32)

If page = "" Then page = 1
If research = "" Then
	delayDelivOnly = "Y"
	''checkCnt = "5"
end if

if (yyyy1="") then
	basedate = Left(CStr(DateAdd("d", -14, now())),10)
	yyyy1 = Left(basedate,4)
	mm1   = Mid(basedate,6,2)
	dd1   = Mid(basedate,9,2)

	basedate = Left(CStr(DateAdd("d", -3, now())),10)
	yyyy2 = Left(basedate,4)
	mm2   = Mid(basedate,6,2)
	dd2   = Mid(basedate,9,2)
end if

fromdate = Left(CStr(DateSerial(yyyy1,mm1 ,dd1)),10)
todate = Left(CStr(DateSerial(yyyy2,mm2 ,dd2+1)),10)

dim oCCSDelivery
set oCCSDelivery = New CCSDelivery
oCCSDelivery.FCurrPage				= page
oCCSDelivery.FPageSize				= 50
oCCSDelivery.FRectStartDate			= fromdate
oCCSDelivery.FRectEndDate			= todate
oCCSDelivery.FRectSongjangDiv		= songjangdiv
oCCSDelivery.FRectDelayDelivOnly	= delayDelivOnly
oCCSDelivery.FRectMakerid			= makerid
oCCSDelivery.FRectOrderserial		= orderserial
oCCSDelivery.FRectCheckCnt			= checkCnt

if (makerid <> "") then
	oCCSDelivery.GetCSMemoDeliveryFixByMakerid()
else
	oCCSDelivery.GetCSMemoDeliveryFixSUM()
end if



%>
<script>

function jsSubmit(frm) {
	frm.submit();
}

function jsSetSongjangDiv(songjangdiv) {
	var frm = document.frm;
	frm.songjangdiv.value = songjangdiv;
	if (frm.songjangdiv.value != songjangdiv) {
		alert('검색불가 택배사입니다.');
		return;
	}
	jsSubmit(frm)
}

function jsSetMakerid(makerid) {
	var frm = document.frm;
	frm.makerid.value = makerid;
	jsSubmit(frm)
}

function goPage(page) {
	var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

function jsReceiveData() {
	var frm = document.frmAct;

	if (confirm("가출고리스트를 가져옵니다.\n\n진행하시겠습니까?")) {
		frm.mode.value = "receivedata";
		frm.submit();
	}
}

</script>
<!-- 검색 시작 -->
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" height="60" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		송장입력일 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		브랜드 : <input type="text" class="text" name="makerid" value="<%= makerid %>">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSubmit(frm);">
	</td>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="25">

	</td>
</tr>
</tr>
</table>
</form>

<p />

<input type="button" class="button" value="가져오기" onClick="jsReceiveData()" disabled>

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(oCCSDelivery.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCCSDelivery.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">idx</td>
	<td width="100">주문번호</td>
	<td width="180">택배사</td>
	<td width="200">송장번호</td>
	<td width="200">브랜드</td>
	<td width="80">결제일</td>
	<td width="80">송장입력일</td>
	<td width="80">실배송일</td>
	<td width="40">조회<br />CNT</td>
	<td width="100">배송조회<br />시작일</td>
	<td width="100">최근<br />배송조회</td>
    <td>비고</td>
</tr>
<% if (oCCSDelivery.FResultCount > 0) then %>
	<% for i = 0 to (oCCSDelivery.FResultCount - 1) %>
	<tr align="center" bgcolor="#FFFFFF" height="25">
		<td><%= oCCSDelivery.FItemList(i).Fidx %></td>
		<td><%= oCCSDelivery.FItemList(i).Forderserial %></td>
		<td><%= oCCSDelivery.FItemList(i).FsongjangName %></td>
		<td>
			<% if (oCCSDelivery.FItemList(i).FsongjangDiv="24") then %>
            <a href="javascript:popDeliveryTrace('<%= oCCSDelivery.FItemList(i).Ffindurl %>','<%= oCCSDelivery.FItemList(i).Fsongjangno %>');"><%= oCCSDelivery.FItemList(i).Fsongjangno %></a>
            <% else %>
            <a target="_blank" href="<%= oCCSDelivery.FItemList(i).Ffindurl + Replace(oCCSDelivery.FItemList(i).Fsongjangno, "-", "") %>"><%= oCCSDelivery.FItemList(i).Fsongjangno %></a>
            <% end if %>
		</td>
		<td><a href="javascript:jsSetMakerid('<%= oCCSDelivery.FItemList(i).Fmakerid %>')"><%= oCCSDelivery.FItemList(i).Fmakerid %></a></td>
		<td>
			<%
			if Not IsNull(oCCSDelivery.FItemList(i).FrealDeliveryDate) then
				if (oCCSDelivery.FItemList(i).Fipkumdate > oCCSDelivery.FItemList(i).FrealDeliveryDate) then
					response.write "<font color='red'><b>" & oCCSDelivery.FItemList(i).Fipkumdate & "</b></font>"
				end if
			end if
			%>
		</td>
		<td><%= oCCSDelivery.FItemList(i).Fbeasongdate %></td>
		<td><%= oCCSDelivery.FItemList(i).FrealDeliveryDate %></td>
		<td><%= oCCSDelivery.FItemList(i).FcheckCnt %></td>
		<td><%= Left(oCCSDelivery.FItemList(i).Fregdate,10) %></td>
		<td><%= Left(oCCSDelivery.FItemList(i).Flastupdate,10) %></td>
    	<td></td>
	</tr>
	<% next %>
	<tr height="20">
	    <td colspan="19" align="center" bgcolor="#FFFFFF">
	        <% if oCCSDelivery.HasPreScroll then %>
			<a href="javascript:goPage('<%= oCCSDelivery.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + oCCSDelivery.StartScrollPage to oCCSDelivery.FScrollCount + oCCSDelivery.StartScrollPage - 1 %>
	    		<% if i>oCCSDelivery.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if oCCSDelivery.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
<% else %>
    <tr height="25" bgcolor="#FFFFFF" align="center">
        <td colspan="12">검색결과가 없습니다.</td>
    </tr>
<% end if %>
</table>

<form name="frmAct" action="DeliveryTrackingList_process.asp">
	<input type="hidden" name="mode">
	<input type="hidden" name="songjangdiv">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
