<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 가출고의심 리스트
' Hieditor : 2018.02.21 이상구 생성
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
	basedate = Left(CStr(DateAdd("d", -7, now())),10)
	yyyy1 = Left(basedate,4)
	mm1   = Mid(basedate,6,2)
	dd1   = Mid(basedate,9,2)

	basedate = Left(CStr(DateAdd("d", -2, now())),10)
	yyyy2 = Left(basedate,4)
	mm2   = Mid(basedate,6,2)
	dd2   = Mid(basedate,9,2)
end if

fromdate = Left(CStr(DateSerial(yyyy1,mm1 ,dd1)),10)
todate = Left(CStr(DateSerial(yyyy2,mm2 ,dd2+1)),10)

dim oCCSDelivery
set oCCSDelivery = New CCSDelivery
oCCSDelivery.FCurrPage				= page
oCCSDelivery.FPageSize				= 100
oCCSDelivery.FRectStartDate			= fromdate
oCCSDelivery.FRectEndDate			= todate
oCCSDelivery.FRectSongjangDiv		= songjangdiv
oCCSDelivery.FRectDelayDelivOnly	= delayDelivOnly
oCCSDelivery.FRectMakerid			= makerid
oCCSDelivery.FRectOrderserial		= orderserial
oCCSDelivery.FRectCheckCnt			= checkCnt

oCCSDelivery.GetCSMemoDeliveryList()

dim oCCSDeliverySUM
set oCCSDeliverySUM = New CCSDelivery
oCCSDeliverySUM.FCurrPage				= 1
oCCSDeliverySUM.FPageSize				= 100
oCCSDeliverySUM.FRectStartDate			= fromdate
oCCSDeliverySUM.FRectEndDate			= todate

oCCSDeliverySUM.GetCSMemoDeliverySUM()

dim songjangName
if (songjangdiv <> "") and oCCSDelivery.FResultCount > 0 then
	if (songjangdiv = CStr(oCCSDelivery.FItemList(0).FsongjangDiv)) then
		songjangName = oCCSDelivery.FItemList(0).FsongjangName
	end if
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

function goPage(page) {
	var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

function jsReTryTracking(songjangdiv) {
	var frm = document.frmAct;
	if (songjangdiv == undefined) {
		alert("택배사 선택 및 검색 후 사용가능합니다.");
		return;
	}

	if (confirm("배송조회 다시하기\n\n조회CNT 가 3~5 회이고 실배송일이 없는 내역에 대해\n배송조회를 다시하도록 합니다.(최근 14일 내역만)\n\n진행하시겠습니까?")) {
		frm.mode.value = "retry";
		frm.songjangdiv.value = songjangdiv;
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
		택배사 :
		<select class="select" name="songjangdiv">
			<option></option>
			<option value="1" <%= CHKIIF(songjangdiv="1", "selected", "") %> >한진택배</option>
			<option value="2" <%= CHKIIF(songjangdiv="2", "selected", "") %> >롯데택배</option>
			<option value="3" <%= CHKIIF(songjangdiv="3", "selected", "") %> >(구)대한통운</option>
			<option value="4" <%= CHKIIF(songjangdiv="4", "selected", "") %> >CJ대한통운</option>
			<option value="8" <%= CHKIIF(songjangdiv="8", "selected", "") %> >우체국택배</option>
			<option value="18" <%= CHKIIF(songjangdiv="18", "selected", "") %> >로젠택배</option>
			<option value="39" <%= CHKIIF(songjangdiv="39", "selected", "") %> >KG로지스</option>
			<option value="41" <%= CHKIIF(songjangdiv="41", "selected", "") %> >드림택배</option>
			<option value="21" <%= CHKIIF(songjangdiv="21", "selected", "") %> >경동택배</option>
			<option value="26" <%= CHKIIF(songjangdiv="26", "selected", "") %> >일양택배</option>
			<option value="28" <%= CHKIIF(songjangdiv="28", "selected", "") %> >동부택배</option>
			<option value="29" <%= CHKIIF(songjangdiv="29", "selected", "") %> >건영택배</option>
			<option value="31" <%= CHKIIF(songjangdiv="31", "selected", "") %> >천일택배</option>
			<option value="33" <%= CHKIIF(songjangdiv="33", "selected", "") %> >호남택배</option>
			<option value="34" <%= CHKIIF(songjangdiv="34", "selected", "") %> >대신화물택배</option>
			<option value="35" <%= CHKIIF(songjangdiv="35", "selected", "") %> >CVSnet택배</option>
			<option value="37" <%= CHKIIF(songjangdiv="37", "selected", "") %> >합동택배</option>
		</select>
		&nbsp;
		브랜드 : <input type="text" class="text" name="makerid" value="<%= makerid %>">
		&nbsp;
		주문번호 : <input type="text" class="text" name="orderserial" value="<%= orderserial %>">
		조회CNT :
		<select class="select" name="checkCnt">
			<option></option>
			<option value="1" <%= CHKIIF(checkCnt="1", "selected", "") %> >1회이상</option>
			<option value="2" <%= CHKIIF(checkCnt="2", "selected", "") %> >2회이상</option>
			<option value="3" <%= CHKIIF(checkCnt="3", "selected", "") %> >3회이상</option>
			<option value="4" <%= CHKIIF(checkCnt="4", "selected", "") %> >4회이상</option>
			<option value="5" <%= CHKIIF(checkCnt="5", "selected", "") %> >5회</option>
		</select>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSubmit(frm);">
	</td>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
        <input type="checkbox" name="delayDelivOnly" value="Y" <%= CHKIIF(delayDelivOnly="Y", "checked", "") %> > 실배송일 오류건만(입금일 이전 배송시작 또는 송장입력 이후 2일이상 배송조회 안됨)
	</td>
</tr>
</tr>
</table>
</form>

<p />

<input type="button" class="button" value="배송조회 다시하기(<%= songjangName %>)" onClick="jsReTryTracking(<%= songjangdiv %>)">

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if (oCCSDeliverySUM.FResultCount > 0) then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% for i = 0 to (oCCSDeliverySUM.FResultCount - 1) %>
		<td width="200"><a href="javascript:jsSetSongjangDiv(<%= oCCSDeliverySUM.FItemList(i).Fsongjangdiv %>)"><%= oCCSDeliverySUM.FItemList(i).FsongjangName %></a></td>
		<td bgcolor="#FFFFFF"><%= oCCSDeliverySUM.FItemList(i).FcheckCnt %></td>
	<% if ((i+1) mod 4) = 0 then %>
	</tr><tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% end if %>
	<% next %>
	<% if (i mod 4) >= 0 then %>
		<% for j = 0 to 4 - (i mod 4) - 1 %>
		<td height="25"></td>
		<td bgcolor="#FFFFFF"></td>
		<% next %>
	<% end if %>
	</tr>
	<% else %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>검색결과 없음</td>
	</tr>
	<% end if %>
</table>

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
	<td width="180">최근조회</td>
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
		<td><%= oCCSDelivery.FItemList(i).Fmakerid %></td>
		<td>
			<%
			if Not IsNull(oCCSDelivery.FItemList(i).FrealDeliveryDate) then
				if (oCCSDelivery.FItemList(i).Fipkumdate > oCCSDelivery.FItemList(i).FrealDeliveryDate) then
					response.write oCCSDelivery.FItemList(i).Fipkumdate
				end if
			end if
			%>
		</td>
		<td><%= oCCSDelivery.FItemList(i).Fbeasongdate %></td>
		<td><%= oCCSDelivery.FItemList(i).FrealDeliveryDate %></td>
		<td><%= oCCSDelivery.FItemList(i).FcheckCnt %></td>
		<td><%= oCCSDelivery.FItemList(i).Flastupdate %></td>
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
