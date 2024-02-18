<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 가출고 리스트
' Hieditor : 2019.06.19 
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenter/delivery/deliveryTrackCls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim page, i, j, k
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, basedate, fromdate, todate
dim songjangdiv, makerid '', orderserial  , etcdivinc, bylist
dim research

page     = requestCheckVar(request("page"),10)
yyyy1   = requestCheckVar(request("yyyy1"),4)
mm1		= requestCheckVar(request("mm1"),2)
dd1		= requestCheckVar(request("dd1"),2)
yyyy2	= requestCheckVar(request("yyyy2"),4)
mm2		= requestCheckVar(request("mm2"),2)
dd2		= requestCheckVar(request("dd2"),2)
songjangdiv		= requestCheckVar(request("songjangdiv"),10)
research		= requestCheckVar(request("research"),3)
makerid			= requestCheckVar(request("makerid"),32)
'orderserial		= requestCheckVar(request("orderserial"),32)

If page = "" Then page = 1
If research = "" Then
	
end if

if (yyyy1="") then
	basedate = Left(CStr(DateAdd("d", -7, now())),7)+"-01"
	yyyy1 = Left(basedate,4)
	mm1   = Mid(basedate,6,2)
	dd1   = Mid(basedate,9,2)

	basedate = Left(CStr(DateAdd("d", -0, now())),10)
	yyyy2 = Left(basedate,4)
	mm2   = Mid(basedate,6,2)
	dd2   = Mid(basedate,9,2)
end if

fromdate = Left(CStr(DateSerial(yyyy1,mm1 ,dd1)),10)
todate = Left(CStr(DateSerial(yyyy2,mm2 ,dd2)),10)

dim oDeliveryTrackBrand
set oDeliveryTrackBrand = New CDeliveryTrack
oDeliveryTrackBrand.FCurrPage			= page
oDeliveryTrackBrand.FPageSize			= 100
oDeliveryTrackBrand.FRectStartDate		= fromdate
oDeliveryTrackBrand.FRectEndDate		= todate
oDeliveryTrackBrand.FRectSongjangDiv	= songjangdiv
oDeliveryTrackBrand.FRectMakerid		= makerid
'oDeliveryTrackBrand.FRectOrderserial	= orderserial

oDeliveryTrackBrand.getDeliveryStatusBrandListAdm()

dim iBrandDefaultDlv, iBrandDefaultDlvName
dim iArrBrandDlv
if (makerid<>"") then
    iArrBrandDlv = getBrandAvgDeliverInfo(fromdate,todate,makerid,"0")

    iBrandDefaultDlv = getBrandDefaultDlv(makerid)
    if (isNULL(iBrandDefaultDlv) or iBrandDefaultDlv="") then
        iBrandDefaultDlvName = "미지정"
        iBrandDefaultDlv = ""
    else
        iBrandDefaultDlvName = getSongjangDiv2Val(iBrandDefaultDlv,1)
    end if

end if

%>
<script>

function jsSubmit(frm) {
	frm.submit();
}


function goPage(page) {
	var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

function popDeliveryTrackingSummaryOne(iorderserial,isongjangno,isongjangdiv){
    var iurl = "/cscenter/delivery/DeliveryTrackingSummaryOne.asp?songjangno="+isongjangno+"&orderserial="+iorderserial+"&songjangdiv="+isongjangdiv;
    var popwin = window.open(iurl,'DeliveryTrackingSummaryOne','width=1200 height=800 scrollbars=yes resizable=yes');
    popwin.focus();

}


var ptblrow;
function chgrowcolor(obj){
	obj.parentElement.parentElement.style.background = "#FCE6E0";
    if ((ptblrow)&&(ptblrow.parentElement.parentElement)){
        ptblrow.parentElement.parentElement.style.background = "#FFFFFF";
    }
    ptblrow=obj;
}

var ptbcol;
function chgcolcolor(obj){
	obj.parentElement.style.background = "#FCE6E0";
    if ((ptbcol)&&(ptbcol.parentElement)){
        ptbcol.parentElement.style.background = "#FFFFFF";
    }
    ptbcol=obj;
}


function chgDefaultSongjangDiv(pval,imakerid){
    var comp = document.getElementById("defaultsongjangdlv");
    var selVal = comp.value
    var selTxt = comp.options[comp.selectedIndex].text;
    
    if (selVal.length<1) return;

    if (pval!=selVal){
        if (confirm(imakerid+" 의 기본택배사를 '"+selTxt + "' 로 변경하시겠습니까?")){
            var iurl = "DeliveryTrackingSummary_Process.asp?makerid="+imakerid+"&mode=chgdftsongjangdiv&chgdiv="+selVal;
            var popwin=window.open(iurl,'chgdefaultDiv','width=200 height=200 scrollbars=yes resizable=yes');
            popwin.focus();
        }
    }else{

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
	<td width="50" height="60" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		주문일 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>

		&nbsp;
		택배사 :
		<% Call drawTrackDeliverBox("songjangdiv",songjangdiv,"Y") %>

		&nbsp;
		브랜드 : <input type="text" class="text" name="makerid" value="<%= makerid %>">

        <% if (FALSE) then %>
		&nbsp;
		주문번호 : <input type="text" class="text" name="orderserial" value="<%= orderserial %>">
        <% end if %>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSubmit(frm);">
	</td>

</tr>
</table>
</form>
<p />


<% if (makerid<>"")  then %>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td width="300">
        브랜드 기본택배사 : <%= iBrandDefaultDlvName %><br>
        <%= getSongjangDlvBoxHtml(iBrandDefaultDlv,"defaultsongjangdlv","") %><input type="button" value="기본택배사 변경" onClick="chgDefaultSongjangDiv('<%=iBrandDefaultDlv%>','<%=makerid%>')">
    </td>
    <td align="right">
        <table width="70%" align="right" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
            <td width="120">택배사</td>
            <td width="100">총건수</td>
            <td width="100">배송완료건수</td>
            <td width="100">완료율</td>
            <td width="100">미집하</td>
            <td width="100">미배달</td>
            <td width="100">평균일수<br>(출고후)</td>
            <td width="100">평균일수<br>(발주후)</td>
            <td width="100">평균일수<br>(결제후)</td>
        </tr>
        <%
        if isArray(iArrBrandDlv) then
            For i=0 To UBound(iArrBrandDlv,2)
        %>
        <tr bgcolor="#FFFFFF" align="right">
                <td align="center"><%=iArrBrandDlv(1,i)%></td>
                <td><%=FormatNumber(iArrBrandDlv(2,i),0)%></td>
                <td><%=FormatNumber(iArrBrandDlv(3,i),0)%></td>
                <td align="center">
                <% if (iArrBrandDlv(2,i)<>0) then %>
                    <%= CLNG(iArrBrandDlv(3,i)/iArrBrandDlv(2,i)*100*100)/100 %> %
                <% end if %>
                </td>
                <td><%=FormatNumber(iArrBrandDlv(4,i),0)%></td>
                <td><%=FormatNumber(iArrBrandDlv(5,i),0)%></td>
                <td align="center"><%=iArrBrandDlv(7,i)%> 일</td>
                <td align="center"><%=iArrBrandDlv(8,i)%> 일</td>
                <td align="center"><%=iArrBrandDlv(9,i)%> 일</td>
        </tr>
        <%
            Next	
        end if
        %>
        </table>
    </td>
</tr>
</table>
<% end if %>

<p />



<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="17">
		검색결과 : <b><%= FormatNumber(oDeliveryTrackBrand.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oDeliveryTrackBrand.FTotalPage,0) %></b>
	</td>
    
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">주문번호</td>
    <td width="80">주문상태</td>
    <td width="70">수령인</td>
    <td width="110">주소1</td>
    <td width="160">상품</td>
	<td width="90">택배사</td>
	<td width="90">송장번호</td>
    <td width="80">송장번호검증</td>
	<!--td width="120">브랜드</td-->

    <td width="100">결제일</td>
    <td width="100">업체확인일</td>
	<td width="100">출고일<br>(송장입력일)</td>
    <td width="100">집하일</td>
    
    <td width="100">배송완료일</td>
    <td width="100">정산확정일</td>

    <td width="130">최근상태</td>
    <td width="70">추적</td>
    <td width="40">비고</td>
</tr>
<% if (oDeliveryTrackBrand.FResultCount > 0) then %>
	<% for i = 0 to (oDeliveryTrackBrand.FResultCount - 1) %>
	<tr align="center" bgcolor="#FFFFFF" height="25">
		<td><%= oDeliveryTrackBrand.FItemList(i).Forderserial %>
        <% if oDeliveryTrackBrand.FItemList(i).FSitename<>"10x10" then %>
            <br><%=oDeliveryTrackBrand.FItemList(i).FSitename%>
        <% end if %>
        </td>
        <td><%= oDeliveryTrackBrand.FItemList(i).getOrderDtlStatusName %></td>
        <td><%= GetUsernameWithAsterisk(oDeliveryTrackBrand.FItemList(i).Freqname,true) %></td>
        <td><%= oDeliveryTrackBrand.FItemList(i).Freqzipaddr %></td>
        <td align="left"><%= oDeliveryTrackBrand.FItemList(i).FItemname %>
            <% if (oDeliveryTrackBrand.FItemList(i).FItemoptionName<>"") then %>
            <br><font color="blue">[<%= oDeliveryTrackBrand.FItemList(i).FItemoptionName %>]</font>
            <% end if %>
        </td>
		<td><%= oDeliveryTrackBrand.FItemList(i).Fdivname %></td>
		<td><%= oDeliveryTrackBrand.FItemList(i).Fsongjangno %></td>
        <td><%= oDeliveryTrackBrand.FItemList(i).getDigitChkStr %></td>
		<!-- td><%= oDeliveryTrackBrand.FItemList(i).Fmakerid %></td -->

        <td><%= oDeliveryTrackBrand.FItemList(i).Fipkumdate %></td>
        <td><%= oDeliveryTrackBrand.FItemList(i).Fupcheconfirmdate %></td>
		<td><%= oDeliveryTrackBrand.FItemList(i).Fbeasongdate %></td>
		<td><%= oDeliveryTrackBrand.FItemList(i).Ftrdeparturedt %></td>
        <td><%= oDeliveryTrackBrand.FItemList(i).Fdlvfinishdt %></td>
        <td><%= oDeliveryTrackBrand.FItemList(i).Fjungsanfixdate %></td>

        <td>
            <% if NOT isNULL(oDeliveryTrackBrand.FItemList(i).Fbeasongdate) then %>
            <%= oDeliveryTrackBrand.FItemList(i).getTrackStateUpcheView %>
            <% end if %>
        </td>
        
        <td>
        <% if (oDeliveryTrackBrand.FItemList(i).isValidPopTraceSongjangDiv) then %>
        <a target="_dlv1" onClick="chgcolcolor(this);" href="<%= oDeliveryTrackBrand.FItemList(i).getTrackURI %>">[택배사]</a>
        <% end if %>

        <% if (oDeliveryTrackBrand.FItemList(i).isValidPopTraceSongjangDiv) then %>
        <br><a target="_dlv2" onClick="chgcolcolor(this);" href="<%= oDeliveryTrackBrand.FItemList(i).getTrackNaverURI %>">[네이버]</a>
        <% end if %>
        </td>
    	<td>
        <a href="#" onClick="chgrowcolor(this);popDeliveryTrackingSummaryOne('<%=oDeliveryTrackBrand.FItemList(i).FOrderserial %>','<%=oDeliveryTrackBrand.FItemList(i).Fsongjangno %>','<%=oDeliveryTrackBrand.FItemList(i).Fsongjangdiv %>');return false;">[검토]</a>
        </td>
	</tr>
	<% next %>
	<tr height="20">
	    <td colspan="17" align="center" bgcolor="#FFFFFF">
	        <% if oDeliveryTrackBrand.HasPreScroll then %>
			<a href="javascript:goPage('<%= oDeliveryTrackBrand.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + oDeliveryTrackBrand.StartScrollPage to oDeliveryTrackBrand.FScrollCount + oDeliveryTrackBrand.StartScrollPage - 1 %>
	    		<% if i>oDeliveryTrackBrand.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if oDeliveryTrackBrand.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
<% else %>
    <tr height="25" bgcolor="#FFFFFF" align="center">
        <% if (makerid<>"") then %>
        <td colspan="17">검색결과가 없습니다.</td>
        <% else %>
        <td colspan="17">브랜드를 선택하세요.</td>
        <% end if %>
    </tr>
<% end if %>
</table>
</form>


<%
SET oDeliveryTrackBrand = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
