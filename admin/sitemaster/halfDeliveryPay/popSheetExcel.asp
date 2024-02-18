<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/halfDeliveryPay/halfdeliverypaycls.asp"-->
<html xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:x="urn:schemas-microsoft-com:office:excel"
    xmlns="http://www.w3.org/TR/REC-html40">
    <head>
        <meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
        <meta name=ProgId content=Excel.Sheet>
        <meta name=Generator content="Microsoft Excel 9">
    </head>
    <body>
<%
    Dim loginUserId, i, currpage, pagesize, keyword, research, itemid, startdate, enddate, isusing, brandid, itemname, regusertype, regusertext
    Dim oHalfDeliveryPayList

    loginUserId = session("ssBctId") '// 로그인한 사용자 아이디
    itemname = requestcheckvar(request("itemname"), 20) '// 상품명 검색어
    research = requestcheckvar(request("research"), 20) '// 재검색여부
    itemid = requestcheckvar(request("itemid"), 2048) '// 상품코드 검색값
    startdate = requestcheckvar(request("startdate"), 20) '// 시작일 검색값
    enddate = requestcheckvar(request("enddate"), 20) '// 종료일 검색값
    isusing = requestcheckvar(request("isusing"), 20) '// 사용여부 검색값
    brandid = requestcheckvar(request("brandid"), 250) '// 브랜드 아이디 검색값
    regusertype = requestcheckvar(request("regusertype"), 250) '// 작성자 검색옵션(id-아이디, name-이름)
    regusertext = requestcheckvar(request("regusertext"), 250) '// 작성자 검색 값

    set oHalfDeliveryPayList = new CgetHalfDeliveryPay
        oHalfDeliveryPayList.FRectcurrpage = 1
        oHalfDeliveryPayList.FRectpagesize = 2000
        If Trim(research)="on" Then
            oHalfDeliveryPayList.FRectItemIds = itemid
            oHalfDeliveryPayList.FRectItemName = itemname
            oHalfDeliveryPayList.FRectStartdate = startdate
            oHalfDeliveryPayList.FRectEnddate = enddate
            oHalfDeliveryPayList.FRectIsUsing = isusing
            oHalfDeliveryPayList.FRectBrandId = brandid
            oHalfDeliveryPayList.FRectRegUserType = regusertype
            oHalfDeliveryPayList.FRectRegUserText = regusertext
        End If
        oHalfDeliveryPayList.GetHalfDeliveryPayList()

    If oHalfDeliveryPayList.FResultcount < 1 Then
        response.write "<script>alert('데이터가 없습니다.');window.close();</script>"
        response.end
    End If

	Response.Expires=0
	response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=배송비부담설정리스트" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
	Response.CacheControl = "public"

    'response.ContentType = "application/vnd.ms-excel"
    'Response.AddHeader "Content-Disposition", "attachment; filename=배송비부담설정리스트_"& left(now(), 10) & ".xls"
    'response.write "<meta http-equiv=Content-Type content='text/html; charset=euc-kr'>"

    Function AddSpace(byval str)
        if ((str = "") or (IsNull(str))) then
            AddSpace = "&nbsp;"
        else
            AddSpace = str
        end if
    End Function

    function ConvertCurrencyUnit(str)
        if (str = "USD") then
            ConvertCurrencyUnit = "$"
        else
            ConvertCurrencyUnit = "￦"
        end if
    End Function
%>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
            <tr valign="top">
                <td class="td_br" width="80" align="center">번호(idx)</td>
                <td class="td_br" width="100" align="center">상품코드</td>
                <td class="td_br" width="100" align="center">브랜드아이디</td>
                <td class="td_br" align="center">상품명</td>
                <td class="td_br" width="90" align="center">시작일</td>
                <td class="td_br" width="90" align="center">종료일</td>
                <td class="td_br" width="120" align="center">배송구분</td>
                <td class="td_br" width="170" align="center">현재해당상품배송구분</td>                
                <td class="td_br" width="170" align="center">무료배송기준금액</td>
                <td class="td_br" width="100" align="center">배송비</td>
                <td class="td_br" width="120" align="center">배송비부담금액</td>
                <td class="td_br" width="80" align="center">사용여부</td>
                <td class="td_br" width="240" align="center">등록일</td>
                <td class="td_br" width="240" align="center">최종수정일</td>
                <td class="td_br" width="120" align="center">작성자</td>
                <td class="td_br" width="120" align="center">최종수정자</td>                
            </tr>
            <% If oHalfDeliveryPayList.FResultcount > 0 Then %>
                <% For i=0 To oHalfDeliveryPayList.Fresultcount-1 %> 
                    <tr align="center" bgcolor="#FFFFFF">
                        <td class="td_br"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fidx%></td>
                        <td class="td_br"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).FItemId%></td>
                        <td class="td_br"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fbrandid%></td>
                        <td class="td_br" align="left"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fitemname%></td>
                        <td class="td_br"><%=Left(oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fstartdate,10)%></td>
                        <td class="td_br"><%=Left(oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fenddate,10)%></td>
                        <td class="td_br"><%=getBeadalDivname(oHalfDeliveryPayList.FHalfDeliveryPayList(i).FdefaultDeliveryType)%></td>
                        <td class="td_br"><%=getBeadalDivname(oHalfDeliveryPayList.FHalfDeliveryPayList(i).FItemDeliveryType)%></td>                        
                        <td class="td_br"><%=Formatnumber(oHalfDeliveryPayList.FHalfDeliveryPayList(i).FdefaultFreeBeasongLimit,0)%>원</td>
                        <td class="td_br"><%=Formatnumber(oHalfDeliveryPayList.FHalfDeliveryPayList(i).FdefaultDeliverPay,0)%>원</td>
                        <td class="td_br"><%=Formatnumber(oHalfDeliveryPayList.FHalfDeliveryPayList(i).FHalfDeliveryPay,0)%>원</td>
                        <td class="td_br">
                            <%
                                If oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fisusing = "Y" Then
                                    Response.write "사용"
                                Else
                                    Response.write "사용안함"
                                End If
                            %>                
                        </td>
                        <td class="td_br"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fregdate%></td>
                        <td class="td_br"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Flastupdate%></td>
                        <td class="td_br"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fadminid%></td>
                        <td class="td_br"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Flastadminid%></td>                        
                    </tr>
                <% next %>
            <% End If %>
        </table>
    </body>
</html>
<%
    set oHalfDeliveryPayList = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
