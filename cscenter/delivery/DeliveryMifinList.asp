<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 미배송 관리 (배송완료일 / 정산확정일)
' 출고후 14일간 배송완료 처리가 안된 주문건.
' Hieditor : 2019.10.23 
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
dim songjangdiv, makerid, orderserial, stp
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
orderserial		= requestCheckVar(request("orderserial"),32)

stp             = requestCheckVar(request("stp"),10)

If page = "" Then page = 1
If research = "" Then
	
end if

if (stp="") then stp="1"

if (yyyy1="") then
	basedate = Left(CStr(DateAdd("d", -7, now())),7)+"-01"
	yyyy1 = Left(basedate,4)
	mm1   = Mid(basedate,6,2)
	dd1   = Mid(basedate,9,2)

	basedate = Left(CStr(DateAdd("d", -1, now())),10)
	yyyy2 = Left(basedate,4)
	mm2   = Mid(basedate,6,2)
	dd2   = Mid(basedate,9,2)
end if

fromdate = Left(CStr(DateSerial(yyyy1,mm1 ,dd1)),10)
todate = Left(CStr(DateSerial(yyyy2,mm2 ,dd2)),10)

dim oDeliveryTrackMifin
set oDeliveryTrackMifin = New CDeliveryTrack
oDeliveryTrackMifin.FCurrPage			= page
oDeliveryTrackMifin.FPageSize			= 100
oDeliveryTrackMifin.FRectStartDate		= fromdate
oDeliveryTrackMifin.FRectEndDate			= todate
oDeliveryTrackMifin.FRectSongjangDiv		= songjangdiv
oDeliveryTrackMifin.FRectMakerid			= makerid
oDeliveryTrackMifin.FRectOrderserial		= orderserial
oDeliveryTrackMifin.FRectSearchType       = stp

oDeliveryTrackMifin.getDeliveryTrackMifinListAdm()


dim iBrandDefaultDlv, iBrandDefaultDlvName
dim iArrBrandDlv
if (makerid<>"") then
    iArrBrandDlv = getBrandAvgDeliverInfo(fromdate,todate,makerid,etcdivinc)

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

function popDeliveryTrackingSummaryOne(iorderserial,isongjangno,isongjangdiv,imakerid){
    var iurl = "/cscenter/delivery/DeliveryTrackingSummaryOne.asp?songjangno="+isongjangno+"&orderserial="+iorderserial+"&songjangdiv="+isongjangdiv+"&makerid="+imakerid;
    var popwin = window.open(iurl,'DeliveryTrackingSummaryOne','width=1200 height=800 scrollbars=yes resizable=yes');
    popwin.focus();

}


function popBrandChulgolistWithDlv(imakerid,isongjangdiv){
    var iurl = "/cscenter/delivery/DeliveryTrackingListBrand.asp?yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&dd1=<%=dd1%>"
    iurl += "&yyyy2=<%=yyyy2%>&mm2=<%=mm2%>&dd2=<%=dd2%>"
    iurl += "&songjangdiv="+isongjangdiv+"&research=<%=research%>"
    iurl += "&makerid="+imakerid;

    var popwin = window.open(iurl,'DeliveryTrackingListBrand','width=1400 height=800 scrollbars=yes resizable=yes');
    popwin.focus();
}

function popBrandChulgolist(imakerid){
    var iurl = "/cscenter/delivery/DeliveryTrackingListBrand.asp?yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&dd1=<%=dd1%>"
    iurl += "&yyyy2=<%=yyyy2%>&mm2=<%=mm2%>&dd2=<%=dd2%>"
    iurl += "&songjangdiv=<%=songjangdiv%>&research=<%=research%>&orderserial=<%=orderserial%>"
    iurl += "&makerid="+imakerid;

    var popwin = window.open(iurl,'DeliveryTrackingListBrand','width=1400 height=800 scrollbars=yes resizable=yes');
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

function switchCheckBox(comp){
    var frm = comp.form;

    if(frm.chkix.length>1){
        for(i=0;i<frm.chkix.length;i++){
            if (!frm.chkix[i].disabled){
                frm.chkix[i].checked = comp.checked;
                AnCheckClick(frm.chkix[i]);
            }
        }
    }else{
        if (!frm.chkix.disabled){
            frm.chkix.checked = comp.checked;
            AnCheckClick(frm.chkix);
        }
    }
}

function AssignDeliverSelect(comp){
    var frm = comp.form;
    var selecidx = frm.basesongjangdlv.selectedIndex;
    var selval   = frm.basesongjangdlv[selecidx].value;

    if (frm.chkix.length>1){
        for (var i=0;i<frm.chgsongjangdiv.length;i++){
            if (frm.chkix[i].checked){
                frm.chgsongjangdiv[i].value=selval;
            }
        }
    }else{
        if (frm.chkix.checked){
            frm.chgsongjangdiv.value=selval;
        }
    }
}

function chgSongjangDivComp(comp,ix){
    var frm = comp.form;
    
    if (comp.value*1>=1){
        if (frm.chkix.length>1){
            frm.chkix[ix].checked=true;
            AnCheckClick(frm.chkix[ix]);
        }else{
            frm.chkix.checked=true;
            AnCheckClick(frm.chkix);
        }
    }
}

function chgSongjangComp(comp,ix){
    var frm = comp.form;

    if (comp.value.length>9){
        if (frm.chkix.length>1){
            frm.chkix[ix].checked=true;
            AnCheckClick(frm.chkix[ix]);
        }else{
            frm.chkix.checked=true;
            AnCheckClick(frm.chkix);
        }
    }

}


function visibleCom(comp){
    return;

    var ibylist = document.getElementById("idbylist");
    if (comp.name=="etcdivinc"){
        if (comp.value=="2"){
            ibylist.style.display="";
        }else if (comp.value=="3"){
            ibylist.style.display="";
            comp.form.bylist.checked=true;
        }else{
            ibylist.style.display="none";
        }
    }
}

function popExceptSongjangBrand(comp){
    var popwin = window.open('DeliveryTrackingEtcFinBrandList.asp','DeliveryTrackingEtcFinDlvList','width=1000 height=800 scrollbars=yes resizable=yes');
    popwin.focus();

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

function popEtcFin(){
    var popwin = window.open('/cscenter/delivery/DeliveryTrackingFake.asp?menupos=4111&etcdivinc=2','popDeliveryTrackingFake','width=1200 height=840 scrollbars=yes resizable=yes');
    popwin.focus();

}

function popErrSongjang(){
    var chulgodtrng = '<%=Left(CStr(DateAdd("d", -7, now())),7)+"-01"%>~<%=Left(CStr(DateAdd("d", -1, now())),10)%>';
    var iurl = '/cscenter/delivery/DeliveryTrackingSummaryDetail.asp?chulgodtrng='+chulgodtrng+'&songjangdiv=&makerid=&isupbea=&mibeatype=999&etcdivinc=0&errchksub=';
    var popwin = window.open(iurl,'DeliveryTrackingSummaryDetail','width=1400 height=800 scrollbars=yes resizable=yes');
    popwin.focus();
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
		송장입력일(출고일) : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>

        <% if (FALSE) then %>
            &nbsp;
            택배사 :
            <% Call drawTrackDeliverBox("songjangdiv",songjangdiv,"Y") %>

            &nbsp;
            브랜드 : <input type="text" class="text" name="makerid" value="<%= makerid %>" onKeyPress="if (event.keyCode == 13) jsSubmit(document.frm);">
            &nbsp;
            주문번호 : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" onKeyPress="if (event.keyCode == 13) jsSubmit(document.frm);">
        
        <% end if %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSubmit(frm);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
    검색조건
    <input type="radio" name="stp" value="0" <%=CHKIIF(stp="0","checked","")%> onClick="visibleCom(this);" >검색기간 미배송
    <input type="radio" name="stp" value="2" <%=CHKIIF(stp="2","checked","")%> onClick="visibleCom(this);" >출고후 D+7 미배송(특정택배사 검토)
    <input type="radio" name="stp" value="1" <%=CHKIIF(stp="1","checked","")%> onClick="visibleCom(this);" >출고후 D+14 미배송
    
    <!-- <input type="radio" name="stp" value="8" <%=CHKIIF(stp="8","checked","")%> onClick="visibleCom(this);" >미정산 -->
    <input type="radio" name="stp" value="9" <%=CHKIIF(stp="9","checked","")%> onClick="visibleCom(this);" >반품존재
    &nbsp;&nbsp;|&nbsp;
    <input type="button" value="기타송장검토" onClick="popEtcFin();">
    &nbsp;
    <input type="button" value="오류송장검토" onClick="popErrSongjang();">

    </td>
</tr>
</table>
</form>
<p />
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td width="100%">
    
    </td>
</tr>
</table>

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
                <td align="center"><a href="#" onClick="popBrandChulgolistWithDlv('<%=makerid%>','<%=iArrBrandDlv(0,i)%>');return false;"><%=iArrBrandDlv(1,i)%></a></td>
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


<form name="frmBChg" method="post" action="DeliveryTrackingSummary_Process.asp">
<input type="hidden" name="mode" value="chgsongjang">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="12">
		검색결과 : <b><%= FormatNumber(oDeliveryTrackMifin.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oDeliveryTrackMifin.FTotalPage,0) %></b>
	</td>
    <td colspan="4" align="left">
        <!-- <input type="button" value="배송완료처리" onClick="CheckNFinishETC(this)";> -->
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <!-- <td width="20"><input type="checkbox" name="chkALL" onClick="switchCheckBox(this);"></td> -->
	<td width="90">주문번호</td>
    <td width="90">수령인</td>
    <td width="110">주소1</td>
    
	<td width="100">택배사</td>
	<td width="110">송장번호</td>
    <td width="100">송장번호검증</td>
	<td width="120">브랜드</td>

	<td width="100">출고일<br>(송장입력일)</td>
    <td width="100">집하일</td>
    <% if (FALSE) then %>
    <td width="100">배송완료일</td>
    <td width="100">정산확정일</td>
    <% end if %>
    <td width="100">배송완료일<br>(추적결과)</td>
    
    <td width="140">최근상태</td>
    <td width="70">추적</td>
    <td width="40">비고</td>
</tr>
<% if (oDeliveryTrackMifin.FResultCount > 0) then %>
	<% for i = 0 to (oDeliveryTrackMifin.FResultCount - 1) %>
    <input type="hidden" name="odetailidx" value="<%= oDeliveryTrackMifin.FItemList(i).Fodetailidx %>">
    <input type="hidden" name="orderserial" value="<%= oDeliveryTrackMifin.FItemList(i).Forderserial %>">
    <input type="hidden" name="songjangno" value="<%= oDeliveryTrackMifin.FItemList(i).Fsongjangno %>">
    <input type="hidden" name="songjangdiv" value="<%= oDeliveryTrackMifin.FItemList(i).FsongjangDiv %>">
	<tr align="center" bgcolor="#FFFFFF" height="25">
        <!-- <td><input type="checkbox" name="chkix" value="<%=i%>" onClick="AnCheckClick(this);" <%=CHKIIF(isNULL(oDeliveryTrackMifin.FItemList(i).Ftrarrivedt),"","disabled") %>></td> -->
		<td><%= oDeliveryTrackMifin.FItemList(i).Forderserial %>
        <% if oDeliveryTrackMifin.FItemList(i).FSitename<>"10x10" then %>
            <br><%=oDeliveryTrackMifin.FItemList(i).FSitename%>
        <% end if %>
        </td>
        <td><%= GetUsernameWithAsterisk(oDeliveryTrackMifin.FItemList(i).Freqname,true) %></td>
        <td><%= oDeliveryTrackMifin.FItemList(i).Freqzipaddr %></td>
        
		<td><%= oDeliveryTrackMifin.FItemList(i).Fdivname %></td>
		<td><%= oDeliveryTrackMifin.FItemList(i).Fsongjangno %></td>
        
        <td><%= oDeliveryTrackMifin.FItemList(i).getDigitChkStr %></td>
		<td><%= oDeliveryTrackMifin.FItemList(i).Fmakerid %></td>
		<td><%= oDeliveryTrackMifin.FItemList(i).Fbeasongdate %></td>
        <td><%= oDeliveryTrackMifin.FItemList(i).Ftrdeparturedt %></td>
		<% if (FALSE) then %>
        <td><%= oDeliveryTrackMifin.FItemList(i).Fdlvfinishdt %></td>
        <td><%= oDeliveryTrackMifin.FItemList(i).Fjungsanfixdate %></td>
        <% end if %>
        
        <td><%= oDeliveryTrackMifin.FItemList(i).Ftrarrivedt %></td>

        <td><%= oDeliveryTrackMifin.FItemList(i).getTrackStateUpcheView %></td>
        <td>
            <% if (oDeliveryTrackMifin.FItemList(i).isValidPopTraceSongjangDiv) then %>
            <a target="_dlv1" onClick="chgcolcolor(this);" href="<%= oDeliveryTrackMifin.FItemList(i).getTrackURI %>">[택배사]</a>
            <% end if %>

            <% if (oDeliveryTrackMifin.FItemList(i).isValidPopTraceSongjangDiv) then %>
            <br><a target="_dlv2" onClick="chgcolcolor(this);" href="<%= oDeliveryTrackMifin.FItemList(i).getTrackNaverURI %>">[네이버]</a>
            <% end if %>
        </td>
    	<td>
        <a href="#" onClick="chgrowcolor(this);popDeliveryTrackingSummaryOne('<%=oDeliveryTrackMifin.FItemList(i).FOrderserial %>','<%=oDeliveryTrackMifin.FItemList(i).Fsongjangno %>','<%=oDeliveryTrackMifin.FItemList(i).Fsongjangdiv %>','<%= oDeliveryTrackMifin.FItemList(i).Fmakerid %>');return false;">[검토]</a>
        </td>
	</tr>
	<% next %>
	<tr height="20">
	    <td colspan="16" align="center" bgcolor="#FFFFFF">
	        <% if oDeliveryTrackMifin.HasPreScroll then %>
			<a href="javascript:goPage('<%= oDeliveryTrackMifin.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + oDeliveryTrackMifin.StartScrollPage to oDeliveryTrackMifin.FScrollCount + oDeliveryTrackMifin.StartScrollPage - 1 %>
	    		<% if i>oDeliveryTrackMifin.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if oDeliveryTrackMifin.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
<% else %>
    <tr height="25" bgcolor="#FFFFFF" align="center">
        <td colspan="19">검색결과가 없습니다.</td>
    </tr>
<% end if %>
</table>
</form>

<%
SET oDeliveryTrackMifin = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
