<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 배송추적 서머리 단건
' Hieditor : 2019.05.23 eastone 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->
<!-- #include virtual="/cscenter/delivery/deliveryTrackCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp" -->
<%
dim i
Dim songjangdiv : songjangdiv	  = requestCheckVar(request("songjangdiv"),10)
Dim songjangno  : songjangno      = requestCheckVar(request("songjangno"),32)
Dim orderserial : orderserial     = requestCheckVar(request("orderserial"),11)
Dim makerid     : makerid         = requestCheckVar(request("makerid"),32)


dim oDeliveryTrackOne
SET oDeliveryTrackOne = New CDeliveryTrack
oDeliveryTrackOne.FRectsongjangno = songjangno

oDeliveryTrackOne.getDeliveryTrackOneInfo()

dim oDeliveryTrackOrder, ordArr
SET oDeliveryTrackOrder = New CDeliveryTrack
oDeliveryTrackOrder.FRectOrderserial = orderserial
oDeliveryTrackOrder.FRectMakerid     = makerid
ordArr = oDeliveryTrackOrder.getDeliveryTrackOrderInfo()


dim iBrandDefaultDlv, iBrandDefaultDlvName
dim iArrBrandDlv
dim ifromdate : ifromdate = LEFT(dateadd("d",-31,now()),10)
dim itodate : itodate = LEFT(now(),10)
if (makerid<>"") then
    iArrBrandDlv = getBrandAvgDeliverInfo(ifromdate,itodate,makerid,"0")

    iBrandDefaultDlv = getBrandDefaultDlv(makerid)
    if (isNULL(iBrandDefaultDlv) or iBrandDefaultDlv="") then
        iBrandDefaultDlvName = "미지정"
        iBrandDefaultDlv = ""
    else
        iBrandDefaultDlvName = getSongjangDiv2Val(iBrandDefaultDlv,1)
    end if

end if

Dim trUri : trUri = getSongjangDiv2Val(songjangdiv,2)
Dim trName  : trName = getSongjangDiv2Val(songjangdiv,1)
%>
<script language="javascript">
function jsSubmit(frm) {
	frm.submit();
}


function addSongjangQue(comp){
    var frm = comp.form;
    var isongjangdiv = frm.addsongjangdiv.value;
    var isongjangno = frm.addsongjangno.value;


    if (isongjangdiv.length<1){
        alert('택배사를 선택하세요.');
        frm.addsongjangdiv.focus();
        return;
    }

    if (isongjangno.length<1){
        alert('송장번호를 입력하세요.');
        frm.isongjangno.focus();
        return;
    }

    frm.mode.value = "retry";
	frm.submit();
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
            if (frm.chkix[ix].disabled==false){
                frm.chkix[ix].checked=true;
                AnCheckClick(frm.chkix[ix]);
            }
        }else{
            if (frm.chkix.disabled==false){
                frm.chkix.checked=true;
                AnCheckClick(frm.chkix);
            }
        }
    }
}

function chgSongjangComp(comp,ix){
    var frm = comp.form;

    if (comp.value.length>9){
        if (frm.chkix.length>1){
            if (frm.chkix[ix].disabled==false){
                frm.chkix[ix].checked=true;
                AnCheckClick(frm.chkix[ix]);
            }

        }else{
            if (frm.chkix.disabled==false){
                frm.chkix.checked=true;
                AnCheckClick(frm.chkix);
            }
        }
    }

}

function chgdlvfinval(comp,ix, jungsandate) {
    var frm = comp.form;

    if (jungsandate == '') {
        jungsandate = '<%=LEFT(now(),10)%>';
    }

    if (frm.chkix.length>1){
        frm.chgdlvfinishdt[ix].value=jungsandate;
        frm.chkix[ix].checked=true;
        AnCheckClick(frm.chkix[ix]);
    }else{
        frm.chgdlvfinishdt.value=jungsandate;
        frm.chkix.checked=true;
        AnCheckClick(frm.chkix);
    }

}


function chkNChangeVal(comp){
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

                if (frm.chgdlvfinishdt[i].value.length<1){
                    /*
                    if (!confirm("배송완료일이 빈값입니다 계속하시겠습니까?")){
                        frm.chgdlvfinishdt[i].focus();
                        return;
                    }
                    */
                }else if (frm.chgdlvfinishdt[i].value.length<10){
                    alert("날짜 형식이 올바르지 않습니다.(YYYY-MM-DD)");
                    frm.chgdlvfinishdt[i].focus();
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
        frm.mode.value="chgdtl";
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

function popByExtorderserial(iextorderserial){
	var iUrl = "/admin/maechul/extjungsandata/extJungsanMapEdit.asp?menupos=1652&page=1&research=on";
	iUrl += "&sellsite="
	iUrl += "&searchfield=extOrderserial&searchtext="+iextorderserial;
	var popwin = window.open(iUrl,"extJungsanMapEdit","width=1400,height=800,crollbars=yes,resizable=yes,status=yes");

	popwin.focus();

}

function popcenter_Action_List(orderserial) {
    var window_width = 1280;
    var window_height = 960;
	var popwin = window.open("<%=replace(manageUrl,"http://","https://")%>/cscenter/action/cs_action.asp?orderserial=" + orderserial ,"cs_action_pop","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}
</script>

<!-- 검색 시작 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td  width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
        &nbsp; 택배사 : <% Call drawTrackDeliverBox("songjangdiv",songjangdiv, "") %>
		&nbsp; 송장번호 : <input type="text" class="text" name="songjangno" value="<%= songjangno %>" size="16" onKeyPress="if (event.keyCode == 13) jsSubmit(document.frm);">
        &nbsp;|&nbsp;
        주문번호 : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="16" onKeyPress="if (event.keyCode == 13) jsSubmit(document.frm);">
        &nbsp;|&nbsp;
        브랜드ID : <input type="text" class="text" name="makerid" value="<%= makerid %>" size="16" onKeyPress="if (event.keyCode == 13) jsSubmit(document.frm);">
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSubmit(document.frm);">
	</td>
</tr>
</form>
</table>


<% if (makerid<>"")  then %>
<p>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td width="300">
        브랜드 기본택배사(<%=makerid%>) : <%= iBrandDefaultDlvName %><br>

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
<p>
<% end if %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="FFFFFF">
<tr>
    <td align="right">
    <% if (trUri<>"") then %>

    <a target="_dlv1" href="<%= trUri + TRIM(replace(songjangno,"-","")) %>">[택배사 추적]</a>

    &nbsp; &nbsp;
    <a target="_dlv2" href="https://search.naver.com/search.naver?query=<%=fnreplaceNvTrName(trName)%>+<%=TRIM(replace(songjangno,"-","")) %>">[네이버 추적]</a>
    &nbsp; &nbsp;
    <% end if %>
    </td>
</tr>

<br>

<p />
<form name="frmQue" method="post" action="DeliveryTrackingSummary_Process.asp">
<input type="hidden" name="mode" value="retry">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11">
        추적 Que 결과
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="100">테이블</td>
    <td width="120">송장번호</td>
    <td width="120">택배사</td>
    <td width="120">Digit Chk</td>
    <td width="120">등록일</td>
    <td width="120">집하일</td>
    <td width="120">배송완료일</td>

    <td width="120">최종업데이트</td>
    <td width="80">추적횟수</td>

    <td width="50"></td>
    <td width="50"></td>

</tr>
<% for i = 0 to (oDeliveryTrackOne.FResultCount - 1) %>
<tr align="center" bgcolor="#FFFFFF">

    <td><%=oDeliveryTrackOne.FItemList(i).getTraceTBLTypeName %></td>
    <td><%=oDeliveryTrackOne.FItemList(i).Fsongjangno %></td>
    <td><%=oDeliveryTrackOne.FItemList(i).getDlvDivName2 %></td>
    <td><%=oDeliveryTrackOne.FItemList(i).getDigitChkStr %></td>
    <td><%=oDeliveryTrackOne.FItemList(i).Fregdt %></td>

    <td ><%=oDeliveryTrackOne.FItemList(i).Fdeparturedt %></td>
    <td ><%=oDeliveryTrackOne.FItemList(i).FdlvfinishDT %></td>
    <td ><%=oDeliveryTrackOne.FItemList(i).Ftraceupddt %></td>
    <td ><%=oDeliveryTrackOne.FItemList(i).FtraceAcctCnt %></td>

    <td align="center">

    </td>
    <td align="center">

    </td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="11" align="right">
    송장추적Que 추가 :
    <% Call drawTrackDeliverBox("addsongjangdiv",songjangdiv, "") %>
    <input type="text" class="text" name="addsongjangno" id="addsongjangno" value="<%= songjangno %>" size="16">
    <input type="button" value="추적Que추가" onClick="addSongjangQue(this);">
    &nbsp;&nbsp;
    </td>
</tr>
</table>
</form>

<br><br>
<p />

<% if isArray(ordArr) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11" align="right">
	</td>
</tr>
<tr align="center" bgcolor="#FFDDDD">
    <td width="100">주문번호</td>
    <td width="70">구매자</td>
    <td width="70">수령인</td>
    <td width="100">주소1</td>
    <td width="120">주문일</td>
    <td width="120">결제일</td>
    <td width="120">발주일</td>

    <td width="50">취소여부</td>
    <td width="80">사이트</td>
    <td width="120">제휴주문번호</td>
    <td width="50"></td>
</tr>
<% if (UBound(ordArr,2)>-1) then %>
<tr align="center" bgcolor="#FFFFFF">

    <td><a href="#" onClick="PopOrderMasterWithCallRingOrderserial('<%=ordArr(0,0) %>');return false;"><%=ordArr(0,0) %></a></td>
    <td><%=GetUsernameWithAsterisk(ordArr(1,0),true) %></td>
    <td><%=GetUsernameWithAsterisk(ordArr(2,0),true) %></td>
    <td><%=ordArr(3,0) %></td>
    <td><%=ordArr(7,0) %></td>
    <td><%=ordArr(8,0) %></td>
    <td><%=ordArr(9,0) %></td>


    <td><%=ordArr(5,0) %></td>
    <td><%=ordArr(11,0) %></td>
    <td>
        <% if (ordArr(11,0)<>"10x10") then %>
        <% if NOT(isNULL(ordArr(29,0))) then %>
        <a href="#" onClick="popByExtorderserial('<%=ordArr(29,0) %>');return false;"><%=ordArr(29,0) %></a>
        <% end if %>
        <% end if %>
    </td>
    <td></td>
</tr>
<% end if %>
</table>

<p>

<%
'' CS내역
dim oJungsanCheckCS
SET oJungsanCheckCS = New CExtJungsan
oJungsanCheckCS.FRectOrderserial = orderserial
if (orderserial<>"") then
    oJungsanCheckCS.getOutJungsanCheckCSInfo()
end if

%>
<% if (oJungsanCheckCS.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="13">
		CS내역 주문번호 : <%= orderserial %>

        &nbsp;<input type="button" class="button" value="관련CS <%=oJungsanCheckCS.FResultCount%>건" class="csbutton" style="width:90px;" onclick="popcenter_Action_List('<%= orderserial %>','','');">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
    <td width="60">csID</td>
    <td width="60">구분</td>
    <td width="80">브랜드ID</td>
    <td width="30">D</td>
    <td width="140">TITLE</td>
    <td width="40">상태</td>
    <td width="70">접수일</td>
    <td width="70">완료일</td>
    <td width="70">확인일</td>
    <td width="70">취소(삭제)일</td>

    <td width="90">연관CsID</td>
    <td width="90">연관주문번호</td>
    <td width="100">비고</td>
</tr>
<% for i=0 to oJungsanCheckCS.FResultCount-1 %>
<%
' if NOT isNULL(oJungsanCheckCS.FItemList(i).getRefOrderSerial) and (oJungsanCheckCS.FItemList(i).getRefOrderSerial<>"") then
'     mapRtnTenOrderserial = oJungsanCheckCS.FItemList(i).getRefOrderSerial
' end if

' if Application("Svr_Info")="Dev" then
'     if (mapRtnTenOrderserial="") then mapRtnTenOrderserial="19040190697"
' end if
%>
<tr align="center" bgcolor="<%=CHKIIF(oJungsanCheckCS.FItemList(i).Fdeleteyn="Y","#DDDDDD","#FFFFFF")%>">
    <td><%=oJungsanCheckCS.FItemList(i).FCsID %></td>
    <td><%=oJungsanCheckCS.FItemList(i).FdivName %></td>
    <td>
        <%=oJungsanCheckCS.FItemList(i).Fmakerid %>
        <% if ((oJungsanCheckCS.FItemList(i).Fmakerid<>"") and (oJungsanCheckCS.FItemList(i).Frequireupche<>"Y")) or ((oJungsanCheckCS.FItemList(i).Fmakerid="") and (oJungsanCheckCS.FItemList(i).Frequireupche="Y")) then %>
        <br>(<%=oJungsanCheckCS.FItemList(i).Frequireupche%>)
        <% end if %>
    </td>
    <td>
        <% if oJungsanCheckCS.FItemList(i).Fdeleteyn<>"N" then response.write "<strong>"&oJungsanCheckCS.FItemList(i).Fdeleteyn&"</strong>" %>
    </td>
    <td align="left"><%=oJungsanCheckCS.FItemList(i).Ftitle %></td>
    <td><%=oJungsanCheckCS.FItemList(i).getCsStateName %> (<%=oJungsanCheckCS.FItemList(i).Fcurrstate%>)</td>
    <td><%=oJungsanCheckCS.FItemList(i).Fregdate %></td>
    <td><%=oJungsanCheckCS.FItemList(i).Ffinishdate %></td>
    <td><%=oJungsanCheckCS.FItemList(i).Fconfirmdate %></td>
    <td><%=oJungsanCheckCS.FItemList(i).Fdeletedate %></td>
    <td><%=oJungsanCheckCS.FItemList(i).Frefasid %></td>
    <td><%=oJungsanCheckCS.FItemList(i).getRefOrderSerial %></td>
    <td></td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="13" align="center">

	</td>
</tr>
</table>
<% end if %>
<% SET oJungsanCheckCS = Nothing %>

<p>

<form name="frmBChg" method="post" action="DeliveryTrackingSummary_Process.asp">
<input type="hidden" name="mode" value="chgdtl">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="right">
        <input type="button" value="선택내역 수정" onClick="chkNChangeVal(this);">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="40"><input type="checkbox" name="chkALL" onClick="switchCheckBox(this);"></td>
    <td width="60">상품코드</td>
    <td width="60">옵션코드</td>
    <td width="80">브랜드ID</td>
    <td width="30">D</td>
    <td width="140">상품명[옵션]</td>
    <td width="40">수량</td>
    <td width="70">구매총액</td>
    <td width="110">확인일</td>
    <td width="110">출고일</td>
    <td width="110">배송일</td>
    <td width="90">정산일</td>
    <td width="110">택배사
    <br><%= getSongjangDlvBoxHtml(songjangdiv,"basesongjangdlv","") %><input type="button" value="v" onClick="AssignDeliverSelect(this)">
    </td>
    <td width="110">송장번호</td>
    <td width="100">비고</td>
</tr>
<% for i=0 to UBound(ordArr,2) %>
<input type="hidden" name="odetailidx" value="<%= ordArr(12,i) %>">
<input type="hidden" name="orderserial" value="<%= ordArr(0,i) %>">
<input type="hidden" name="songjangno" value="<%= ordArr(25,i) %>">
<input type="hidden" name="songjangdiv" value="<%= ordArr(24,i) %>">
<tr align="center" bgcolor="<%=CHKIIF(ordArr(6,i)="Y","#DDDDDD","#FFFFFF")%>">
    <td>
    <% if (ordArr(13,i)=0 or ordArr(13,i)=100) then %>
    <input type="checkbox" name="chkix" value="<%=i%>" disabled >
    <% else %>
    <input type="checkbox" name="chkix" value="<%=i%>" onClick="AnCheckClick(this);" <%=CHKIIF(ordArr(6,i)<>"Y","","disabled") %>><% '정산일 있어도 배송일 입력 허용, 2022-01-26, skyer9 %>
    <% end if %>
    </td>
    <td><%=ordArr(13,i) %></td>
    <td><%=ordArr(14,i) %></td>
    <td><%=ordArr(17,i) %></td>
    <td>
        <%=ordArr(23,i) %>
        /
        <% if ordArr(6,i)<>"N" then response.write "<strong>"&ordArr(6,i)&"</strong>" %>
    </td>
    <td align="left">
        <%=DDotFormat(ordArr(15,i),10) %>
        <%
        if (ordArr(16,i)<>"") then
            response.write "<br><font color=blue>["&ordArr(16,i)&"]</font>"
        end if
        %>
    </td>
    <td><%=ordArr(22,i) %></td>
    <td><%=ordArr(20,i) %></td>

    <td><%=ordArr(18,i) %></td>
    <td><%=ordArr(26,i) %></td>
    <td>
        <% if (ordArr(13,i)=0 or ordArr(13,i)=100) then %>
        <input type="hidden" name="chgdlvfinishdt">
        <%=ordArr(27,i) %>
        <% else %>
        <input type="text" name="chgdlvfinishdt" size="12" maxlength="19" value="<%=ordArr(27,i) %>" onKeyup="chgSongjangComp(this,<%= i %>);" <%=CHKIIF(isNULL(ordArr(28,i)) and Not isNULL(ordArr(26,i)),"","readonly") %>>
        <% if  isNULL(ordArr(27,i)) then %><input type="button" value="T" onclick="chgdlvfinval(this,<%= i %>, '<%=ordArr(28,i) %>')" style="cursor:pointer"><% end if %>
        <% end if %>
    </td>
    <td><%=ordArr(28,i) %></td>
    <td>
        <% if (ordArr(13,i)=0 or ordArr(13,i)=100) then %>
        <input type="hidden" name="chgsongjangdiv">
        <% else %>
        <%= getSongjangDlvBoxHtml(ordArr(24,i),"chgsongjangdiv","onChange='chgSongjangDivComp(this,"&i&")'") %>
        <% end if %>
    </td>
    <td>
    <% if (ordArr(13,i)=0 or ordArr(13,i)=100) then %>
    <input type="hidden" name="chgsongjangno">
    <% else %>
    <input type="text" name="chgsongjangno" size="12" maxlength="20" value="<%=ordArr(25,i) %>" onKeyup="chgSongjangComp(this,<%= i %>);">
    <% end if %>
    </td>
    <td><%=ordArr(30,i) %></td>
</tr>
<% next %>
</table>
</form>
<% end if %>

<br>
<p />

<%
'' 송장 변경로그 by 주문번호
dim oSongjangChgLog
SET oSongjangChgLog = new CDeliveryTrack
oSongjangChgLog.FRectOrderserial = orderserial
if (orderserial<>"") then
    oSongjangChgLog.getSongjangChangeLogList()
end if
%>
<p  >
<% if (oSongjangChgLog.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
        송장변경로그 주문번호 : <%= orderserial %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="60">LogIdx</td>
    <td width="60">상품코드</td>
    <td width="60">옵션코드</td>
    <td width="110">이전택배사</td>
    <td width="110">이전송장번호</td>
    <td width="110">변경택배사</td>
    <td width="110">변경송장번호</td>

    <td width="80">변경자</td>
    <td width="70">등록일</td>
    <td width="70">변경구분</td>

    <td width="70">현재택배사</td>
    <td width="50">현재송장번호</td>
    <td width="90">출고일</td>
    <td width="90">배송일</td>
    <td width="90">정산일</td>

    <td width="100">비고</td>
</tr>
<% for i=0 to oSongjangChgLog.FResultCount-1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><%=oSongjangChgLog.FItemList(i).Fsongjangchgidx %></td>
    <td><%=oSongjangChgLog.FItemList(i).FItemid %></td>
    <td><%=oSongjangChgLog.FItemList(i).FItemOption %></td>
    <td><%=getSongjangDiv2Val(oSongjangChgLog.FItemList(i).Fpsongjangdiv,1) %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fpsongjangno %></td>
    <td><%=getSongjangDiv2Val(oSongjangChgLog.FItemList(i).Fchgsongjangdiv,1) %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fchgsongjangno %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fchguserid %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fregdt %></td>
    <td><%=oSongjangChgLog.FItemList(i).FactionType %></td>
    <td><%=getSongjangDiv2Val(oSongjangChgLog.FItemList(i).Fsongjangdiv,1) %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fsongjangno %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fbeasongdate %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fdlvfinishdt %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fjungsanfixdate %></td>
    <td></td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16" align="center">

	</td>
</tr>
</table>
<% end if %>

<% SET oSongjangChgLog = Nothing %>

<br>
<p />
<%
SET oDeliveryTrackOne = Nothing
SET oDeliveryTrackOrder = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
