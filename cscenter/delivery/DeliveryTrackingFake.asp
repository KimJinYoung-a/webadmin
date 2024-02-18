<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 가출고 리스트
' Hieditor : 2019.06.19 서동석 생성
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
dim songjangdiv, makerid, orderserial, etcdivinc, bylist
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
etcdivinc       = requestCheckVar(request("etcdivinc"),10)
bylist          = requestCheckVar(request("bylist"),10)

If page = "" Then page = 1
If research = "" Then
	
end if

if (etcdivinc="") then etcdivinc="0"
if (etcdivinc="0") then bylist="0"
if (bylist="") then bylist="0"

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

dim oDeliveryTrackFake
set oDeliveryTrackFake = New CDeliveryTrack
oDeliveryTrackFake.FCurrPage			= page
oDeliveryTrackFake.FPageSize			= 100
oDeliveryTrackFake.FRectStartDate		= fromdate
oDeliveryTrackFake.FRectEndDate			= todate
oDeliveryTrackFake.FRectSongjangDiv		= songjangdiv
oDeliveryTrackFake.FRectMakerid			= makerid
oDeliveryTrackFake.FRectOrderserial		= orderserial
oDeliveryTrackFake.FRectEtcdivinc       = etcdivinc
oDeliveryTrackFake.FRectByList          = bylist

if (oDeliveryTrackFake.FRectEtcdivinc<>"3") then
    oDeliveryTrackFake.getFakeSongjangGrpBrandListAdm()
else
    oDeliveryTrackFake.getFakeSongjangErrDlvListAdm()
end if

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

/*
function jsSetSongjangDiv(songjangdiv) {
	var frm = document.frm;
	frm.songjangdiv.value = songjangdiv;
	if (frm.songjangdiv.value != songjangdiv) {
		alert('검색불가 택배사입니다.');
		return;
	}
	jsSubmit(frm)
}
*/

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

function popThisByBrand(imakerid){
    var iurl = "/cscenter/delivery/DeliveryTrackingFake.asp?yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&dd1=<%=dd1%>"
    iurl += "&yyyy2=<%=yyyy2%>&mm2=<%=mm2%>&dd2=<%=dd2%>"
    iurl += "&songjangdiv=<%=songjangdiv%>&research=<%=research%>&orderserial=<%=orderserial%>&etcdivinc=<%=etcdivinc%>"
    iurl += "&makerid="+imakerid;

    var popwin = window.open(iurl,'DeliveryTrackingFakepop','width=1400 height=800 scrollbars=yes resizable=yes');
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

function CheckNFinishETC(comp){
    var frm = comp.form;
    var pass = false;

    if (!frm.chkix){
        alert("선택 내역이 없습니다.");
        return;
    }

    if(frm.chkix.length>1){
        for (var i=0;i<frm.chkix.length;i++){
            pass = (pass||frm.chkix[i].checked);
        }
    }else{
        pass = frm.chkix.checked;
    }

    if (!pass) {
        alert("선택 내역이 없습니다.");
        return;
    }

    if(frm.chkix.length>1){
        for (var i=0;i<frm.chkix.length;i++){
            if (frm.chkix[i].checked){
                if (!((frm.chgsongjangdiv[i].value=="99")||(frm.chgsongjangdiv[i].value=="98")||(frm.chgsongjangdiv[i].value=="100"))){
                    alert("기타출고 처리는 기타 또는 퀵,한우리물류 만 가능합니다.");
                    frm.chgsongjangdiv[i].focus();
                    return;
                }else if ((frm.chgsongjangno[i].value).length<1){
                    alert("송장번호를 입력하시기 바랍니다.");
                    frm.chgsongjangno[i].focus();
                    return;
                }
            }
        }
    }else{
        if (frm.chkix.checked){
            if (!((frm.chgsongjangdiv.value=="99")||(frm.chgsongjangdiv.value=="98")||(frm.chgsongjangdiv.value=="100"))){
                alert("기타출고 처리는 기타 또는 퀵,한우리물류 만 가능합니다.");
                return;
            }else if ((frm.chgsongjangno.value).length<1){
                alert("송장번호를 입력하시기 바랍니다.");
                frm.chgsongjangno.focus();
                return;
            }
        }
    }


    if (confirm("선택 내역을 기타출고 완료 처리(배송완료일입력) 하시겠습니까?")){
        frm.mode.value="finetc";
        frm.submit();
    }
}

function CheckNChangeSongjang(comp){
    var frm = comp.form;
    var pass = false;

    if (!frm.chkix){
        alert("선택 내역이 없습니다.");
        return;
    }

    if(frm.chkix.length>1){
        for (var i=0;i<frm.chkix.length;i++){
            pass = (pass||frm.chkix[i].checked);
        }
    }else{
        pass = frm.chkix.checked;
    }

    if (!pass) {
        alert("선택 내역이 없습니다.");
        return;
    }

    if(frm.chkix.length>1){
        for (var i=0;i<frm.chkix.length;i++){
            if (frm.chkix[i].checked){
                if (frm.chgsongjangdiv[i].value.length<1){
                    alert("택배사를 선택하시기 바랍니다.");
                    frm.chgsongjangdiv[i].focus();
                    return;
                }else if ((frm.chgsongjangno[i].value).length<1){
                    alert("송장번호를 입력하시기 바랍니다.");
                    frm.chgsongjangno[i].focus();
                    return;
                }
            }
        }
    }else{
        if (frm.chkix.checked){
            if (frm.chgsongjangdiv.value.length<1){
                alert("택배사를 선택하시기 바랍니다.");
                return;
            }else if ((frm.chgsongjangno.value).length<1){
                alert("송장번호를 입력하시기 바랍니다.");
                frm.chgsongjangno.focus();
                return;
            }
        }
    }


    if (confirm("선택 내역을 수정 하시겠습니까?")){
        frm.mode.value="chgsongjang";
        frm.submit();
    }
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

function visibleCom(comp){
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

function refreshSummary(){
    if (confirm('목록을 재작성 하시겠습니까??')){
        var iurl = "DeliveryTrackingSummary_Process.asp?mode=refreshfakesummary";
        var popwin=window.open(iurl,'etcdlvfinauto','width=200 height=200 scrollbars=yes resizable=yes');
        popwin.focus();
    }
}

function autoFinEtcDlv(){
    if (confirm('기타 택배사 일괄처리 진행 하시겠습니까?')){
        var iurl = "DeliveryTrackingSummary_Process.asp?mode=etcdlvfinauto";
        var popwin=window.open(iurl,'etcdlvfinauto','width=200 height=200 scrollbars=yes resizable=yes');
        popwin.focus();
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
		송장입력일(출고일) : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>

		&nbsp;
		택배사 :
		<% Call drawTrackDeliverBox("songjangdiv",songjangdiv,"Y") %>

		&nbsp;
		브랜드 : <input type="text" class="text" name="makerid" value="<%= makerid %>" onKeyPress="if (event.keyCode == 13) jsSubmit(document.frm);">
		&nbsp;
		주문번호 : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" onKeyPress="if (event.keyCode == 13) jsSubmit(document.frm);">

        <% if (FALSE) then %>
            조회CNT :
            <select class="select" name="checkCnt">
                <option></option>
                <option value="1" <%= CHKIIF(checkCnt="1", "selected", "") %> >1회이상</option>
                <option value="2" <%= CHKIIF(checkCnt="2", "selected", "") %> >2회이상</option>
                <option value="3" <%= CHKIIF(checkCnt="3", "selected", "") %> >3회이상</option>
                <option value="4" <%= CHKIIF(checkCnt="4", "selected", "") %> >4회이상</option>
                <option value="5" <%= CHKIIF(checkCnt="5", "selected", "") %> >5회</option>
            </select>
        <% end if %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSubmit(frm);">
	</td>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
    기타택배조건
    <input type="radio" name="etcdivinc" value="0" <%=CHKIIF(etcdivinc="0","checked","")%> onClick="visibleCom(this);" >전체
    <input type="radio" name="etcdivinc" value="1" <%=CHKIIF(etcdivinc="1","checked","")%> onClick="visibleCom(this);" >기타/퀵 제외
    <input type="radio" name="etcdivinc" value="2" <%=CHKIIF(etcdivinc="2","checked","")%> onClick="visibleCom(this);" >기타/퀵 만 검색

    &nbsp;|&nbsp;
    <input type="radio" name="etcdivinc" value="3" <%=CHKIIF(etcdivinc="3","checked","")%> onClick="visibleCom(this);" >택배사오지정 예상건

    &nbsp;&nbsp;
    <span id="idbylist" style="display:<%=CHKIIF(etcdivinc="2" or etcdivinc="3","","none")%>"><input type="checkbox" name="bylist" value="1" <%=CHKIIF(bylist="1","checked","") %> >리스트로 보기</span>


    </td>
</tr>
</tr>
</table>
</form>
<p />
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td width="100%">
    * 기타 (99) / 퀵(98) 처리규칙 <br>
    1. 자체배송은 출고일 익일로 잡는다. (전일 출고 완료건 기준)<br>
    2. 클래스 상품, 플라워 배송, 화물 배송으로 등록된 상품은 익일로 잡는다 (odeliverfixday in (L,C,X)) <br>
    3. 특정 관리 카테고리는 익일로 잡는다. (가구, 감성채널:플라워, 홈/데코:거울, 홈/데코:조명 <!--, 디지털:PC/노트북 -->) <br>
    4. 특정브랜드 (등록된 브랜드) 는 익일로 잡는다 (방문수령업체(케이크), 가구배송하는업체, 화분.. 액자) <a href="#" onClick="popExceptSongjangBrand(this);return false;"><font color="blue">[기타택배사 자동처리 브랜드 관리]</font></a><br>
    5. 기타 (99) 이면서 송장번호 (직접배송,직배송,직납,직배,직접배달,직접수령,방문수령,가구화물배송,업체직송) 인경우.
    <br>
    경동택배,일양택배,건영택배,천일택배,대신택배,호남택배:추가로 조회해볼 필요가 있음.<br>
    한우리물류 - 추적불가
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

<%
if (makerid="") then    ' and (bylist="0")  프로시저에서는 뺀듯?
%>
    <table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
        <td colspan="12">
            검색결과 : <b><%= FormatNumber(oDeliveryTrackFake.FTotalCount,0) %></b>
            &nbsp;
            페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oDeliveryTrackFake.FTotalPage,0) %></b>

            &nbsp;(30분 단위 서머리 자료임)

            <% if (session("ssBctId")="icommang") then %>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value="기타택배사 일괄처리" onClick="autoFinEtcDlv()">
            &nbsp;&nbsp;
            <input type="button" value="서머리내역 재작성" onClick="refreshSummary();">
            <% end if %>
        </td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="120">브랜드ID</td>
        <td width="100">기간총출고</td>
        <td width="100">기간평균배송소요</td>
        
        <td width="120">총매배주문수<br>(<%=oDeliveryTrackFake.FSumdelayTTLOrderGrp%>)</td>
        <td width="120">순매배주문수<br>(기타,퀵제외)<br>(<%=oDeliveryTrackFake.FSumMibeaTTLOrderGrp%>)</td>
        <td width="100" bgcolor="#AAAA77">미집하주문수<br>(기타,퀵제외)<br>(<%=oDeliveryTrackFake.FSummijiphaTTLOrderGrp%>)</td>
        <td width="100">집하후미이동<br>(기타,퀵제외)<br>(<%=oDeliveryTrackFake.FSumjiphaNoMoveTTLOrderGrp%>)</td>
        
        <td width="120">총매배건수<br>(<%=oDeliveryTrackFake.FSumdelayTTL%>)</td>
        <td width="120">순매배건수<br>(기타,퀵제외)<br>(<%=oDeliveryTrackFake.FSumMibeaTTL%>)</td>
        <td width="100">미집하건수<br>(기타,퀵제외)<br>(<%=oDeliveryTrackFake.FSummijiphaTTL%>)</td>
        <td width="100">집하후미이동<br>(기타,퀵제외)<br>(<%=oDeliveryTrackFake.FSumjiphaNoMoveTTL%>)</td>
        <td>비고</td>
    </tr>
    <% if (oDeliveryTrackFake.FResultCount > 0) then %>
        <% for i = 0 to (oDeliveryTrackFake.FResultCount - 1) %>
        <tr align="center" bgcolor="#FFFFFF" height="25">
            <td><%= oDeliveryTrackFake.FItemList(i).Fmakerid %></td>
            <td></td>
            <td></td>
            <td><%= oDeliveryTrackFake.FItemList(i).FdelayTTLOrderGrp %></td>
            <td><%= oDeliveryTrackFake.FItemList(i).FmibeaTTLOrderGrp %></td>
            <td><%= oDeliveryTrackFake.FItemList(i).FmijiphaTTLOrderGrp %></td>
            <td><%= oDeliveryTrackFake.FItemList(i).FjiphaNoMoveTTLOrderGrp %></td>

            <td><%= oDeliveryTrackFake.FItemList(i).FdelayTTL %></td>
            <td><%= oDeliveryTrackFake.FItemList(i).FmibeaTTL %></td>
            <td><%= oDeliveryTrackFake.FItemList(i).FmijiphaTTL %></td>
            <td><%= oDeliveryTrackFake.FItemList(i).FjiphaNoMoveTTL %></td>
            <td>
                <a href="#" onClick="chgrowcolor(this);popThisByBrand('<%=oDeliveryTrackFake.FItemList(i).Fmakerid %>');return false;">[브랜드별_미출고검토]</a>
                &nbsp;
                <a href="#" onClick="chgrowcolor(this);popBrandChulgolist('<%=oDeliveryTrackFake.FItemList(i).Fmakerid %>');return false;">[브랜드별_전체]</a>
            </td>
        </tr>
        <% next %>
        <tr height="20">
            <td colspan="12" align="center" bgcolor="#FFFFFF">
                <% if oDeliveryTrackFake.HasPreScroll then %>
                <a href="javascript:goPage('<%= oDeliveryTrackFake.StartScrollPage-1 %>');">[pre]</a>
                <% else %>
                    [pre]
                <% end if %>

                <% for i=0 + oDeliveryTrackFake.StartScrollPage to oDeliveryTrackFake.FScrollCount + oDeliveryTrackFake.StartScrollPage - 1 %>
                    <% if i>oDeliveryTrackFake.FTotalpage then Exit for %>
                    <% if CStr(page)=CStr(i) then %>
                    <font color="red">[<%= i %>]</font>
                    <% else %>
                    <a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
                    <% end if %>
                <% next %>

                <% if oDeliveryTrackFake.HasNextScroll then %>
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
<% else %>
<form name="frmBChg" method="post" action="DeliveryTrackingSummary_Process.asp">
<input type="hidden" name="mode" value="chgsongjang">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11">
		검색결과 : <b><%= FormatNumber(oDeliveryTrackFake.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oDeliveryTrackFake.FTotalPage,0) %></b>
	</td>
    <td colspan="2" align="left">
        <input type="button" value="기타내역 배송완료처리" onClick="CheckNFinishETC(this)";>
    </td>
    <td colspan="3" align="right">
        <input type="button" value="선택내역 송장 일괄수정" onClick="CheckNChangeSongjang(this)";>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"><input type="checkbox" name="chkALL" onClick="switchCheckBox(this);"></td>
	<td width="90">주문번호</td>
    <td width="90">수령인</td>
    <td width="110">주소1</td>
    <td width="160">상품</td>
	<td width="100">택배사</td>
    <td width="140">변경할 택배사<br>
    <%= getSongjangDlvBoxHtml(iBrandDefaultDlv,"basesongjangdlv","") %><input type="button" value="v" onClick="AssignDeliverSelect(this)">
    </td>
	<td width="110">송장번호</td>
    <td width="110">변경할 송장번호</td>
    <td width="100">송장번호검증</td>
	<td width="120">브랜드</td>

	<td width="100">출고일<br>(송장입력일)</td>
    <td width="100">집하일</td>
    <!--
	<td width="100">최근추적 Que</td>
    <td width="40">추적<br>회수</td>

    <td width="100">배송완료일<br>(추적결과)</td>
    <td width="100">최근결과일<br>(추적결과)</td>
    -->
    <td width="140">최근상태</td>
    <td width="70">추적</td>
    <td width="40">비고</td>
</tr>
<% if (oDeliveryTrackFake.FResultCount > 0) then %>
	<% for i = 0 to (oDeliveryTrackFake.FResultCount - 1) %>
    <input type="hidden" name="odetailidx" value="<%= oDeliveryTrackFake.FItemList(i).Fodetailidx %>">
    <input type="hidden" name="orderserial" value="<%= oDeliveryTrackFake.FItemList(i).Forderserial %>">
    <input type="hidden" name="songjangno" value="<%= oDeliveryTrackFake.FItemList(i).Fsongjangno %>">
    <input type="hidden" name="songjangdiv" value="<%= oDeliveryTrackFake.FItemList(i).FsongjangDiv %>">
	<tr align="center" bgcolor="#FFFFFF" height="25">
        <td><input type="checkbox" name="chkix" value="<%=i%>" onClick="AnCheckClick(this);" <%=CHKIIF(isNULL(oDeliveryTrackFake.FItemList(i).Ftrarrivedt),"","disabled") %>></td>
		<td><%= oDeliveryTrackFake.FItemList(i).Forderserial %>
        <% if oDeliveryTrackFake.FItemList(i).FSitename<>"10x10" then %>
            <br><%=oDeliveryTrackFake.FItemList(i).FSitename%>
        <% end if %>
        </td>
        <td><%= GetUsernameWithAsterisk(oDeliveryTrackFake.FItemList(i).Freqname,true) %></td>
        <td><%= oDeliveryTrackFake.FItemList(i).Freqzipaddr %></td>
        <td align="left"><%= oDeliveryTrackFake.FItemList(i).FItemname %>
            <% if (oDeliveryTrackFake.FItemList(i).FItemoptionName<>"") then %>
            <br><font color="blue">[<%= oDeliveryTrackFake.FItemList(i).FItemoptionName %>]</font>
            <% end if %>
        </td>
		<td><%= oDeliveryTrackFake.FItemList(i).Fdivname %></td>
        <td>
            <%= getSongjangDlvBoxHtml(oDeliveryTrackFake.FItemList(i).FsongjangDiv,"chgsongjangdiv","onChange='chgSongjangDivComp(this,"&i&")'") %>
        </td>
		<td><%= oDeliveryTrackFake.FItemList(i).Fsongjangno %></td>
        <td>
            <input type="text" name="chgsongjangno" size="14" maxlength="20" value="<%= oDeliveryTrackFake.FItemList(i).Fsongjangno %>" onKeyup="chgSongjangComp(this,<%= i %>);">
        </td>
        <td><%= oDeliveryTrackFake.FItemList(i).getDigitChkStr %></td>
		<td><%= oDeliveryTrackFake.FItemList(i).Fmakerid %></td>
		<td><%= oDeliveryTrackFake.FItemList(i).Fbeasongdate %></td>
		<td><%= oDeliveryTrackFake.FItemList(i).Ftrdeparturedt %></td>
        <td><%= oDeliveryTrackFake.FItemList(i).getTrackStateUpcheView %></td>
        
        <% if (FALSE) then %>
		<td><%= oDeliveryTrackFake.FItemList(i).Fquelastupddt %></td>
        <td><%= oDeliveryTrackFake.FItemList(i).Fquelastupdno %></td>
        
        <td><%= oDeliveryTrackFake.FItemList(i).Ftrarrivedt %></td>
        <td><%= oDeliveryTrackFake.FItemList(i).Ftrupddt %></td>
        <% end if %>
        <td>
        <% if (oDeliveryTrackFake.FItemList(i).isValidPopTraceSongjangDiv) then %>
        <a target="_dlv1" onClick="chgcolcolor(this);" href="<%= oDeliveryTrackFake.FItemList(i).getTrackURI %>">[택배사]</a>
        <% end if %>

        <% if (oDeliveryTrackFake.FItemList(i).isValidPopTraceSongjangDiv) then %>
        <br><a target="_dlv2" onClick="chgcolcolor(this);" href="<%= oDeliveryTrackFake.FItemList(i).getTrackNaverURI %>">[네이버]</a>
        <% end if %>
        </td>
    	<td>
        <a href="#" onClick="chgrowcolor(this);popDeliveryTrackingSummaryOne('<%=oDeliveryTrackFake.FItemList(i).FOrderserial %>','<%=oDeliveryTrackFake.FItemList(i).Fsongjangno %>','<%=oDeliveryTrackFake.FItemList(i).Fsongjangdiv %>');return false;">[검토]</a>
        </td>
	</tr>
	<% next %>
	<tr height="20">
	    <td colspan="16" align="center" bgcolor="#FFFFFF">
	        <% if oDeliveryTrackFake.HasPreScroll then %>
			<a href="javascript:goPage('<%= oDeliveryTrackFake.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + oDeliveryTrackFake.StartScrollPage to oDeliveryTrackFake.FScrollCount + oDeliveryTrackFake.StartScrollPage - 1 %>
	    		<% if i>oDeliveryTrackFake.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if oDeliveryTrackFake.HasNextScroll then %>
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
<% end if %>

<%
SET oDeliveryTrackFake = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
