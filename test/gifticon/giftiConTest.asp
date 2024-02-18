<% option Explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/test/gifticon/giftiConCls.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

Dim cpnNum: cpnNum = requestCheckvar(request("cpnNum"),20)
Dim mode  : mode   = requestCheckvar(request("mode"),20)

dim oGicon
dim ret, bufStr

IF (mode="P100") THEN  ''조회
    set oGicon = new CGiftiCon
    ret = oGicon.reqCouponState(cpnNum,"100100")  ''쿠폰번호, 추적번호
    
    if (ret) then
        bufStr =          "SERVICE_CODE:" & oGicon.FConResult.FSERVICE_CODE & VbCRLF
        bufStr = bufStr & "COUPON_NUMBER:" & oGicon.FConResult.FCOUPON_NUMBER & VbCRLF
        bufStr = bufStr & "ERROR_CODE:" & oGicon.FConResult.getResultCode & VbCRLF
        bufStr = bufStr & "MESSAGE:" & oGicon.FConResult.FMESSAGE & VbCRLF
        bufStr = bufStr & "EXCHANGE_COUNT:" & oGicon.FConResult.FEXCHANGE_COUNT & VbCRLF
        bufStr = bufStr & "FBODY_LENGTH:" & oGicon.FConResult.FBODY_LENGTH & VbCRLF ''
        
        bufStr = bufStr & "SubItemCode:" & oGicon.FConResult.FSubItemCode & VbCRLF
        bufStr = bufStr & "SubItemBarCode:" & oGicon.FConResult.FSubItemBarCode & VbCRLF
        bufStr = bufStr & "SubItemEa:" & oGicon.FConResult.FSubItemEa & VbCRLF
        bufStr = bufStr & "SubSupplyID:" & oGicon.FConResult.FSubSupplyID    & VbCRLF   
        bufStr = bufStr & "ItemPrice:" & oGicon.FConResult.getItemPrice & VbCRLF 
        bufStr = bufStr & "SubSupplyPrice:" & oGicon.FConResult.FSubSupplyPrice  & VbCRLF  
        bufStr = bufStr & "SubPartnerCharge:" & oGicon.FConResult.FSubPartnerCharge  & VbCRLF
        bufStr = bufStr & "SubSupplyerCharge:" & oGicon.FConResult.FSubSupplyerCharge & VbCRLF
        bufStr = bufStr & "FSubSubItemType:"&oGicon.FConResult.FSubSubItemType    & VbCRLF
        bufStr = bufStr & "LimitPrice:"    &oGicon.FConResult.FSubLimitPrice     & VbCRLF
        bufStr = bufStr & "DiscountPrice:" &oGicon.FConResult.FSubDiscountPrice  & VbCRLF
        bufStr = bufStr & "SubNotice:[" &oGicon.FConResult.FSubNotice&"]"   & VbCRLF
        bufStr = bufStr & "SubFiller:[" &oGicon.FConResult.FSubFiller&"]"   & VbCRLF 
    else
        response.write "ERR::"&oGicon.FLASTERROR
    end if
    set oGicon = Nothing

ELSEIF (mode="P110") THEN ''승인 
    set oGicon = new CGiftiCon
    ret = oGicon.reqCouponApproval(cpnNum,"100100",10000) ''쿠폰번호, 추적번호, 상품 교환가
    
    if (ret) then
        bufStr =          "SERVICE_CODE:" & oGicon.FConResult.FSERVICE_CODE & VbCRLF
        bufStr = bufStr & "COUPON_NUMBER:" & oGicon.FConResult.FCOUPON_NUMBER & VbCRLF
        bufStr = bufStr & "ERROR_CODE:" & oGicon.FConResult.getResultCode & VbCRLF
        bufStr = bufStr & "MESSAGE:" & oGicon.FConResult.FMESSAGE & VbCRLF
        bufStr = bufStr & "EXCHANGE_COUNT:" & oGicon.FConResult.FEXCHANGE_COUNT & VbCRLF
        
        bufStr = bufStr & "ApprovNO:" & oGicon.FConResult.FApprovNO & VbCRLF
        bufStr = bufStr & "ExchangePrice:" & oGicon.FConResult.FExchangePrice & VbCRLF
    
    end if
    set oGicon = Nothing
ELSEIF (mode="P120") THEN ''승인취소
    set oGicon = new CGiftiCon
    ret = oGicon.reqCouponCancel(cpnNum,"100100",10000) ''쿠폰번호, 추적번호, 상품 교환가
    
    if (ret) then
        bufStr =          "SERVICE_CODE:" & oGicon.FConResult.FSERVICE_CODE & VbCRLF
        bufStr = bufStr & "COUPON_NUMBER:" & oGicon.FConResult.FCOUPON_NUMBER & VbCRLF
        bufStr = bufStr & "ERROR_CODE:" & oGicon.FConResult.getResultCode & VbCRLF
        bufStr = bufStr & "MESSAGE:" & oGicon.FConResult.FMESSAGE & VbCRLF
        bufStr = bufStr & "EXCHANGE_COUNT:" & oGicon.FConResult.FEXCHANGE_COUNT & VbCRLF
        
        bufStr = bufStr & "ApprovNO:" & oGicon.FConResult.FApprovNO & VbCRLF
        bufStr = bufStr & "ExchangePrice:" & oGicon.FConResult.FExchangePrice & VbCRLF
    
    end if
    set oGicon = Nothing
ELSE

END IF
%>

<html>
<head>
<script language='javascript'>
function reqState(frm){
    if (frm.cpnNum.value.length<12){
        alert('쿠폰번호를 입력하세요.');
        return;
    }
    
    frm.mode.value="P100";
    frm.submit();
}

function appReq(frm){
    if (frm.cpnNum.value.length<12){
        alert('쿠폰번호를 입력하세요.');
        return;
    }
    
    frm.mode.value="P110";
    frm.submit();
}

function cancelReq(frm){
    if (frm.cpnNum.value.length<12){
        alert('쿠폰번호를 입력하세요.');
        return;
    }
    
    frm.mode.value="P120";
    frm.submit();
}

</script>
</head>
<body >
<p>
<table width="800" border=1 cellpadding=1 cellspacing=1>
<form name="frmGft" method="get" action="">
<input type="hidden" name="mode" value="">
<Tr>    
    <%
    Dim iArr
    iArr=CHRB(0)
    iArr=iArr&CHRB(0)
    iArr=iArr&CHRB(0)
    iArr=iArr&CHRB(255)
    %>
    <td colspan="2">LL:<%= getNByteLng(iArr,0,4) %></td>
</tr>
<Tr>    
    <td colspan="2">[<%= Dec2Hex(255,4) %>][<%= Dec2Hex(289,4) %>][<%= Dec2Hex(201,4) %>]</td>
</tr>
<Tr>    
    <td colspan="2">
    [<% dim buf : buf= (DecTo4ByteChar(255)) : dPByteArrayDEcimal(buf) %>]
    
    [<%' dPByteArrayDEcimalw(Hex2ByteArray(Dec2Hex(254,4))) %>]
    [<%'= dPByteArrayDEcimal(Hex2ByteArray(Dec2Hex(201,4))) %>]</td>
</tr>
<Tr>    
    <td colspan="2">[<%= ASC(CHR(128)) %>][<%= ASCw(CHRw(289)) %>][<%= ASCw(CHRw(201)) %>][<%= ASCB(CHRB(201)) %>][<%= ASC(CHR("&HC9")) %>]</td>
</tr>

<Tr>   
    <td colspan="2">[<%= ChrB("&H00")&ChrB("&H00")&ChrB("&H00")&ChrB("&H00")&ChrB("&H00")&ChrB("&H00")&ChrB("&H00")&ChrB("&HC9") %>]</td>
</tr>
<Tr>    
    <td colspan="2"><%= Len(CHRB(0)&CHRB(201)) %> <%= Len(CHR(49)&CHR(50)) %></td>
</tr>
<Tr>    
    <td colspan="2"><%= CHRW("&H" & HEX(44032)) %><%= HEX(ASCW("가")) %><%= HEX(ASCW("년")) %><%= CHRW("&H" & "B144") %></td>
</tr>
<tr>
    <td>상품</td>
    <td colspan="2">
        999033886637<br>
        999443414267
    </td>
</tr>
<tr>
    <td>할인권 3,000</td>
    <td colspan="2">
        999285852130<br>
        999692393518<br>
        999517507629<br>
        999687233899
    </td>
</tr>
<tr>
    <td>상품권 10,000</td>
    <td colspan="2">
        999003162323<br>
        999127891875<br>
        999039073026<br>
        999690913464
    </td>
</tr>


<tr>
    <td>giftCon번호</td>
    <td><input type="text"name="cpnNum" value="<%= cpnNum %>" maxlength="12" Size="12"></td>
    <td >
        <input type="button" value="조회" onClick="reqState(frmGft)">
        <input type="button" value="승인" onClick="appReq(frmGft)">
        <input type="button" value="승인취소" onClick="cancelReq(frmGft)">
    </td>
</tr>
<tr>
    <td>결과MSG</td>
    <td colspan="2">
    <textarea cols=100 rows=20><%= bufStr %></textarea>
    </td>
</tr>
</form>
</table>
</body>
</html>