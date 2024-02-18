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
dim orderserial : orderserial = requestCheckvar(request("orderserial"),16)

dim oJungsanCheck
SET oJungsanCheck = new CJungsanCheck
oJungsanCheck.FRectOrderserial = orderserial
oJungsanCheck.getLogDiffByOrderserial


Dim pitemgubun, pitemid, pitemoption
pitemid = -1

%>
<script language='javascript'>
function chgCancelOrderJFixdtNULL(iorderserial){
    var frm = document.frmXsiteOrderVal;
    frm.mode.value="chgCancelOrderJFixdtNULL";
    frm.orderserial.value=iorderserial;

    if (confirm("주문 내역 배송완료일 정산을을 NULL 로 변경하시겠습니까?")){
        frm.submit();
    }
}

function chgCancelOrderDetailRealsellprice(iorderserial,iitemid,iitemoption,orgrealsellprice){
    var chgrealsellprice = "";
    chgrealsellprice = prompt("변경할금액", "");
    if (chgrealsellprice == null) return;

    if (chgrealsellprice.length<1){
        alert("실판매가를 입력하세요.");
        return;
    }

    if (!IsDigit(chgrealsellprice)){
        alert('숫자를 입력하세요.');
        return;
    }

    var frm = document.frmXsiteOrderVal;
    frm.mode.value="chgRealOrderRealsellprice";
    frm.orderserial.value=iorderserial;
    frm.itemid.value=iitemid;
    frm.itemoption.value=iitemoption;
    frm.chgval.value=chgrealsellprice;

    if (confirm("주문내역 실판매가 값을 "+orgrealsellprice+" => "+chgrealsellprice+" 로 변경하시겠습니까?")){
        frm.submit();
    }
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
	
		
		* 주문번호 : <input type="text" name="orderserial" value="<%=orderserial%>" size="11" maxlength="16">
        &nbsp;

	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" style="width:70px;height:50px;" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="left" bgcolor="#FFFFFF" >
	<td>

	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<p  >
<% if oJungsanCheck.FresultCount>0 then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="100">주문번호</td>
    <td width="120">사이트</td>
    <td width="70">취소여부</td>
    <td width="70">CHK</td>
    <td>비고</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
    <td><%= oJungsanCheck.FItemList(0).Forderserial %></td>
    <td><%= oJungsanCheck.FItemList(0).Fsitename %></td>
    <td><%= oJungsanCheck.FItemList(0).Fcancelyn %></td>
    <td><%= oJungsanCheck.FItemList(0).getLogCheckTypeName %></td>
    <td align="left">
    <% if  (NOT isNULL(oJungsanCheck.FItemList(0).Fchktype)) then %>
    <% if  oJungsanCheck.FItemList(0).Fchktype=8 then %>
        <input type="button" value="출고/정산일 NULL처리" onClick="chgCancelOrderJFixdtNULL('<%= oJungsanCheck.FItemList(0).Forderserial %>'); return false;">
    <% end if %>
    <% end if %>
    </td>
</tr>
</table>
<% end if %>
<p>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="27">
		
		
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="30">상품<br>구분</td>
	<td width="60">상품코드</td>
	<td width="60">옵션코드</td>
	<td width="90">브랜드ID</td>
    <td width="50">상품<br>수량</td>
	<td width="50">취소<br>여부</td>

    <td width="60">판매가</td>
    <td width="60">구매단가</td>
    <td width="60">매출단가</td>
    <td width="60">매입가</td>
    <td width="60">매입<br>구분</td>
	<td width="60">출고일</td>
    <td width="60">배송일</td>
    <td width="60">정산일</td>
    <td width="40">과세<br>구분</td>
    <td width="10"></td>

    <td width="30">로그<br>Sub</td>
    <td width="60">로그<br>수량</td>
    <td width="60">로그<br>판매가</td>
    <td width="60">로그<br>구매단가</td>
    <td width="60">로그<br>매출단가</td>
    <td width="60">로그<br>매입가</td>
    <td width="60">로그<br>매입구분</td>
	<td width="60">로그<br>출고일</td>
    <td width="60">로그<br>정산일</td>
    <td width="40">로그<br>과세</td>
	<td>비고</td>

   
</tr>

<% if oJungsanCheck.FresultCount<1 then %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
    <td colspan="27">
       
        [검색결과가 없습니다.]
    </td>
</tr>
<% else %>
<% for i=0 to oJungsanCheck.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
    <% if (oJungsanCheck.FItemList(i).Fitemgubun=pitemgubun and oJungsanCheck.FItemList(i).Fitemid=pitemid and oJungsanCheck.FItemList(i).Fitemoption=pitemoption) then %>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
    <% else %>
        <td><%= oJungsanCheck.FItemList(i).Fitemgubun %></td>
        <td><%= oJungsanCheck.FItemList(i).Fitemid %></td>
        <td><%= oJungsanCheck.FItemList(i).Fitemoption %></td>
        <td><%= oJungsanCheck.FItemList(i).Fmakerid %></td>
        <td><%= oJungsanCheck.FItemList(i).Fitemno %></td>
        <td><%= oJungsanCheck.FItemList(i).getDCancelynName %></td>
        
        <td align="right"><%= FormatNumber(oJungsanCheck.FItemList(i).FitemcostCouponNotApplied, 0) %></td>
        <td align="right"><%= FormatNumber(oJungsanCheck.FItemList(i).Fitemcost, 0) %></td>
        <td align="right">
            <% if (oJungsanCheck.FItemList(i).Fdcancelyn="Y") and (oJungsanCheck.FItemList(i).Fsitename<>"10x10") and (LEN(oJungsanCheck.FItemList(i).Forderserial)=11) and (oJungsanCheck.FItemList(i).Fitemid<>0)  then %>
            <a href="#" onClick="chgCancelOrderDetailRealsellprice('<%= oJungsanCheck.FItemList(i).Forderserial %>','<%= oJungsanCheck.FItemList(i).Fitemid %>','<%= oJungsanCheck.FItemList(i).Fitemoption %>',<%= oJungsanCheck.FItemList(i).Freducedprice %>); return false;"><%= FormatNumber(oJungsanCheck.FItemList(i).Freducedprice, 0) %></a>
            <% else %>
            <%= FormatNumber(oJungsanCheck.FItemList(i).Freducedprice, 0) %>
            <% end if %>
        </td>
        <td align="right"><%= FormatNumber(oJungsanCheck.FItemList(i).Fbuycash, 0) %></td>
        <td><%= oJungsanCheck.FItemList(i).Fomwdiv %></td>
        <td><%= oJungsanCheck.FItemList(i).Fbeasongdate %></td>
        <td><%= oJungsanCheck.FItemList(i).Fdlvfinishdt %></td>
        <td><%= oJungsanCheck.FItemList(i).Fjungsanfixdate %></td>
        <td><%= oJungsanCheck.FItemList(i).Fvatinclude %></td>
    <% end if %>
    <%
    pitemgubun  = oJungsanCheck.FItemList(i).Fitemgubun
    pitemid     = oJungsanCheck.FItemList(i).Fitemid
    pitemoption = oJungsanCheck.FItemList(i).Fitemoption
    %>
    <td width="10"></td>
    
    <td><%= oJungsanCheck.FItemList(i).Fsuborderserial %></td>
    <td><%= oJungsanCheck.FItemList(i).Flgitemno %></td>
    <td align="right">
        <% if NOT isNULL(oJungsanCheck.FItemList(i).FlgitemcostCouponNotApplied) then %>
        <%= FormatNumber(oJungsanCheck.FItemList(i).FlgitemcostCouponNotApplied, 0) %>
        <% end if %>
    </td>
    <td align="right">
        <% if NOT isNULL(oJungsanCheck.FItemList(i).Flgitemcost) then %>
        <%= FormatNumber(oJungsanCheck.FItemList(i).Flgitemcost, 0) %>
        <% end if %>
    </td>
    <td align="right">
        <% if NOT isNULL(oJungsanCheck.FItemList(i).FlgreducedPrice) then %>
        <%= FormatNumber(oJungsanCheck.FItemList(i).FlgreducedPrice, 0) %>
        <% end if %>
    </td>
    <td align="right">
        <% if NOT isNULL(oJungsanCheck.FItemList(i).Flgbuycash) then %>
        <%= FormatNumber(oJungsanCheck.FItemList(i).Flgbuycash, 0) %>
        <% end if %>
    </td>
    <td><%= oJungsanCheck.FItemList(i).Flgomwdiv %></td>
	<td><%= oJungsanCheck.FItemList(i).Flgbeasongdate %></td>
    <td><%= oJungsanCheck.FItemList(i).FDTLjFixedDt %></td>
    <td><%= oJungsanCheck.FItemList(i).Flgvatinclude %></td>
    <td></td>
</tr>
<% next %>
<% end if %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="27" align="center">
		
	</td>
</tr>
</table>

<%
set oJungsanCheck = Nothing
%>
<form name="frmXsiteOrderVal" method="post" action="/admin/maechul/extjungsandata/extJungsan_process.asp">
<input type="hidden" name="mode" value="chgRealOrderRealsellprice">
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="chgval" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
