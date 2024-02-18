<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 정산 검토
' Hieditor : 2020/03/30 eastone
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/jungsan/jungsanCheckCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
Dim i
dim research : research = requestCheckvar(request("research"),10)
dim page : page = requestCheckvar(request("page"),10)
dim difftp : difftp = requestCheckvar(request("difftp"),10)
dim chkbysum : chkbysum = requestCheckvar(request("chkbysum"),10)

dim yyyy1, mm1, fromDate, toDate
yyyy1 = requestCheckvar(request("yyyy1"),4)
mm1 = requestCheckvar(request("mm1"),2)

if (yyyy1="") then yyyy1=LEFT(dateadd("d",-4,NOW()),4)
if (mm1="") then mm1=MID(dateadd("d",-4,NOW()),6,2)

if (page="") then page=1
''if (difftp="") then difftp="0"  

fromDate = yyyy1+"-"+mm1+"-01"
toDate = dateADD("m",1,fromDate)

dim yyyymm : yyyymm = yyyy1+"-"+mm1

dim oJungsanCheck
SET oJungsanCheck = new CJungsanCheck
oJungsanCheck.FPageSize = 500
oJungsanCheck.FCurrPage = page
oJungsanCheck.FRectYYYYMM = yyyymm
oJungsanCheck.FRectDiffType = difftp
' oJungsanCheck.FRectCheckBySum = chkbysum

oJungsanCheck.getLogDiffList

dim FormatDotNo : FormatDotNo=0
%>
<script language='javascript'>

/*
function popByExtorderserial(iextorderserial){
	var iUrl = "/admin/maechul/extjungsandata/extJungsanMapEdit.asp?menupos=<%=menupos%>&page=1&research=on";
	
	iUrl += "&searchfield=extOrderserial&searchtext="+iextorderserial;
	var popwin = window.open(iUrl,"extJungsanMapEdit","width=1400,height=800,scrollbars=yes,resizable=yes,status=yes");

	popwin.focus();

}

function actCpnMapByorderserial(sellsite,extOrderserial,Orderserial){
	var popwin = window.open("","extJungsanEditProc","width=600,height=300");
	
	popwin.location.href="/admin/maechul/extjungsandata/extJungsan_process.asp?sellsite="+sellsite+"&extOrderserial="+extOrderserial+"&Orderserial="+Orderserial+"&mode=mapcpnbyorderserial";

	popwin.focus();
}
*/
function popJungsanOrderCheckByOrderserial(iorderserial){
	var popwin = window.open("","popJungsanOrderCheckByOrderserial","width=1200,height=800");
	
	popwin.location.href="/admin/jungsan/popJungsanCheckByOrder.asp?orderserial="+iorderserial;

	popwin.focus();
}

function popReOrderLogByOrderserial(iorderserial){
    var popwin = window.open("","popmaechul_log_process","width=300,height=300");
	
	popwin.location.href="/admin/maechul/maechul_log_process.asp?orderserial="+iorderserial+"&mode=relogorderserialwithque";

	popwin.focus();
}
 
function popDeliveryTrackingSummaryOneOrderserial(iorderserial,imakerid){
    var popwin = window.open("","DeliveryTrackingSummaryOne","width=1200,height=800");
	
	popwin.location.href="/cscenter/delivery/DeliveryTrackingSummaryOne.asp?orderserial="+iorderserial+"&makerid="+imakerid;

	popwin.focus();
}


</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	
		
		* 매출월:
		<% DrawYMBox yyyy1,mm1 %>
        &nbsp;

        * 오차타입
        <input type="radio" name="difftp" <%=CHKIIF(difftp="","checked","")%> value="" >배치로그
        |
        <input type="radio" name="difftp" <%=CHKIIF(difftp="900","checked","")%> value="900" >정산확정검토
        |
        <input type="radio" name="difftp" <%=CHKIIF(difftp="100","checked","")%> value="100" >ON 주문로그
        |
        <input type="radio" name="difftp" <%=CHKIIF(difftp="200","checked","")%> value="200" >OFF 주문로그
        |
        <input type="radio" name="difftp" <%=CHKIIF(difftp="300","checked","")%> value="300" >ON 정산/로그
        |
        <input type="radio" name="difftp" <%=CHKIIF(difftp="400","checked","")%> value="400" >OFF 정산/로그

        <% if (FALSE) then %>
		&nbsp;
        * <input type="checkbox" name="chkbysum" <%=CHKIIF(chkbysum<>"","checked","")%> >합계로보기
        <% end if %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" style="width:70px;height:50px;" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="left" bgcolor="#FFFFFF" >
	<td>
        배치로그는 야간에 전일기준으로 생성됨, 주간에 배치처리 
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<p  >
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="22">
		검색결과 : <b><%= oJungsanCheck.FTotalcount %></b>
		&nbsp;
		<% if oJungsanCheck.FTotalcount>=oJungsanCheck.FPageSize then %>
        (최대 <%=FormatNumber(oJungsanCheck.FPageSize,0)%> 건)
        <% end if %>
	</td>
</tr>
<form name="frm1" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="90">주문번호</td>
	<td width="80">구분</td>
	<td width="80">Que등록일</td>
	<td width="90">사이트</td>
    <td width="60">주문구분</td>
	<td width="60">취소여부(M)</td>
    <td width="60">주문상태</td>

    <td width="60">상품코드</td>
    <td width="60">옵션코드</td>
    <td width="60">수량</td>
	<td width="60">브랜드ID</td>
    <td width="60">구매총액(단가)</td>
    <td width="60">매출총액(단가)</td>
    <td width="60">매입가</td>
    <td width="60">취소여부(D)</td>
    <td width="60">출고일</td>
    <td width="60">배송일</td>
    <td width="60">정산일</td>
	<td>비고</td>

   
</tr>

<% if oJungsanCheck.FresultCount<1 then %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
    <td colspan="22">
       
        [검색결과가 없습니다.]
    </td>
</tr>
<% else %>
<% for i=0 to oJungsanCheck.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td>
        <% if (difftp="900") then %>
        <a href="#" onClick="popDeliveryTrackingSummaryOneOrderserial('<%= oJungsanCheck.FItemList(i).Forderserial %>','<%= NULL2Blank(oJungsanCheck.FItemList(i).Fmakerid) %>'); return false;"><%= oJungsanCheck.FItemList(i).Forderserial %></a>
        <% else %>
        <a href="#" onClick="popJungsanOrderCheckByOrderserial('<%= oJungsanCheck.FItemList(i).Forderserial %>'); return false;"><%= oJungsanCheck.FItemList(i).Forderserial %></a>
        <% end if %>
    </td>
	<td><%= oJungsanCheck.FItemList(i).getLogCheckTypeName %></td>
	<td><%= oJungsanCheck.FItemList(i).FQueregdt %></td>
    <td><%= oJungsanCheck.FItemList(i).Fsitename %></td>
    <td><%= oJungsanCheck.FItemList(i).getJumundivName %></td>
    <td><%= oJungsanCheck.FItemList(i).getCancelynName %></td>
    <td><%= oJungsanCheck.FItemList(i).getIpkumdivname %></td>

    <td><%= oJungsanCheck.FItemList(i).Fitemid %></td>
    <td><%= oJungsanCheck.FItemList(i).Fitemoption %></td>
    <td><%= oJungsanCheck.FItemList(i).Fitemno %></td>
    <td><%= oJungsanCheck.FItemList(i).Fmakerid %></td>
	<td align="right">
        <% if NOT isNULL(oJungsanCheck.FItemList(i).Fitemcost) then %>
        <%= FormatNumber(oJungsanCheck.FItemList(i).Fitemcost, 0) %>
        <% end if %>
    </td>
	<td align="right">
        <% if NOT isNULL(oJungsanCheck.FItemList(i).Freducedprice) then %>
        <%= FormatNumber(oJungsanCheck.FItemList(i).Freducedprice, 0) %>
        <% end if %>
    </td>
   
    <td>
        <% if NOT isNULL(oJungsanCheck.FItemList(i).Fbuycash) then %>
        <%= FormatNumber(oJungsanCheck.FItemList(i).Fbuycash,0) %>
        <% end if %>
    </td>
    <td><%= oJungsanCheck.FItemList(i).Fdcancelyn %></td>
    <td><%= oJungsanCheck.FItemList(i).Fbeasongdate %></td>
    <td><%= oJungsanCheck.FItemList(i).Fdlvfinishdt %></td>
    <td><%= oJungsanCheck.FItemList(i).Fjungsanfixdate %></td>
    
    <td>
    <% if NOT isNULL(oJungsanCheck.FItemList(i).Forderserial) then %>
    <% if LEN(oJungsanCheck.FItemList(i).Forderserial)=11 or LEN(oJungsanCheck.FItemList(i).Forderserial)=16 then %>
    <input type="button" value="로그재작성" onClick="popReOrderLogByOrderserial('<%= oJungsanCheck.FItemList(i).Forderserial %>');">
    <% end if %>
    <% end if %>
    </td>
</tr>
<% next %>
<% end if %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="22" align="center">
		<% if oJungsanCheck.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oJungsanCheck.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oJungsanCheck.StartScrollPage to oJungsanCheck.FScrollCount + oJungsanCheck.StartScrollPage - 1 %>
			<% if i>oJungsanCheck.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oJungsanCheck.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</form>
</table>

<%
set oJungsanCheck = Nothing
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
