<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 출고지시주문리스트
' History : 이상구 생성
'           2023.07.11 한용민 수정(ems 국제우편물 사전통관정보 수정)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/baljucls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->

<script type='text/javascript'>

function ViewOrderDetail(iorderserial){
	var popwin;
    popwin = window.open('viewordermaster.asp?orderserial=' + iorderserial,'orderdetail','scrollbars=yes,resizable=yes,width=800,height=600');
    popwin.focus();
}

function saveSongjang(frm){
    if (confirm('송장번호 저장하시겠습니까?')){
        frm.mode.value="svsongjang";
        frm.submit()
    }
}

function saveTotalWeight(frm){
    if (frm.realweight.value.length<1){
        alert('총 무게를 입력해 주세요');
        frm.realweight.focus();
        return;
    }

    if (confirm('총 무게를 저장하시겠습니까?')){
        frm.mode.value="svttlwight";
        frm.submit()
    }
}

function saveBoxSize(frm) {
    if (frm.boxSizeX.value.length<1){
        alert('박스 사이즈를 입력해 주세요');
        frm.boxSizeX.focus();
        return;
    }

    if (frm.boxSizeY.value.length<1){
        alert('박스 사이즈를 입력해 주세요');
        frm.boxSizeY.focus();
        return;
    }

    if (frm.boxSizeZ.value.length<1){
        alert('박스 사이즈를 입력해 주세요');
        frm.boxSizeZ.focus();
        return;
    }

    if (confirm('박스 사이즈를 저장하시겠습니까?')){
        frm.mode.value="saveBoxSize";
        frm.submit()
    }
}

function jsSendReq(idx, emsGubun, gubun) {
	var popwin;
	var url = '/admin/ordermaster/lib/emsApi_process.asp?mode=sendReq&emsGubun=' + emsGubun + '&idx=' + idx + '&gubun=' + gubun;
	popwin = window.open(url,'jsSendReq','scrollbars=yes,resizable=yes,width=600,height=400');
    popwin.focus();
}

function jsDownXL(idx, songjangdiv, gubun) {
	//var popwin;
	//var url = 'popbaljuList.asp?mode=excel&gubun=' + gubun + '&idx=' + idx + '&songjangdiv=' + songjangdiv;
	//popwin = window.open(url,'jsDownXL','scrollbars=yes,resizable=yes,width=600,height=400');
    //popwin.focus();
	frmbalju.mode.value="excel";
	frmbalju.gubun.value=gubun;
	frmbalju.idx.value=idx;
	frmbalju.songjangdiv.value=songjangdiv;
	frmbalju.action="";
	frmbalju.target="view";
	frmbalju.submit();
}

function baljuSearch() {
	frmbalju.mode.value="";
	frmbalju.action="";
	frmbalju.target="";
	frmbalju.submit();
}

</script>
<%
dim obalju, i, j, k, ipkumdiv
dim idx : idx = trim(RequestCheckVar(request("idx"),10))
dim songjangdiv : songjangdiv = RequestCheckVar(request("songjangdiv"),10)
dim gubun : gubun = RequestCheckVar(request("gubun"),32)
dim mode : mode = RequestCheckVar(request("mode"),10)
dim realweight : realweight = RequestCheckVar(request("realweight"),10)
dim baljusongjangno : baljusongjangno = RequestCheckVar(request("baljusongjangno"),20)
dim cancelyn : cancelyn = RequestCheckVar(request("cancelyn"),1)
dim reload : reload = RequestCheckVar(request("reload"),32)
dim oCOrderDetail, totItemNo, totItemUsDollar
	ipkumdiv = RequestCheckVar(getNumeric(request("ipkumdiv")),1)

if songjangdiv = "" then songjangdiv = getSongjangDivFromIdx(idx) end if
if reload="" and cancelyn="" then cancelyn="N"

set obalju = New CBalju
if ((songjangdiv="90") or (songjangdiv="8") or (songjangdiv="92")) then
    if (songjangdiv="90") or (songjangdiv="92") then
    	'EMS택배
		obalju.FRechWeightGubun = gubun
		obalju.FRectrealweight = realweight
		obalju.FRectbaljusongjangno = baljusongjangno
		obalju.FRectcancelyn = cancelyn
		obalju.FRectipkumdiv = ipkumdiv
    	obalju.getBaljuDetailListEMS idx
    else
    	'우체국택배(군부대)
		obalju.FRectcancelyn = cancelyn
		obalju.FRectipkumdiv = ipkumdiv
    	obalju.getBaljuDetailListMilitary idx
    end if

	If mode = "excel" Then
	    if (songjangdiv="90") then
	    	Response.Buffer = True					'2015-11-12 11:47 김진영 추가(EMS 다운시 한글 깨짐)
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", "attachment; filename=EMS_" & idx & ".xls"
	    elseif (songjangdiv="92") then
	    	Response.Buffer = True
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", "attachment; filename=UPS_" & idx & ".xls"
	    else
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", "attachment; filename=EPOST_" & idx & ".xls"
	    end if

		response.clear

		if (songjangdiv="90") then
			response.write "<meta http-equiv=""content-type"" content=""text/html; charset=euc-kr"">"		'2015-11-12 11:47 김진영 추가(EMS 다운시 한글 깨짐)
			response.write "<table border=1>"

			'// 4칸 떨어져 있어야 엑셀 업로드가 성공한다.
			'// 주문자명이 영문이어야 엑셀 업로드가 성공한다.
			For k = 1 To 4
				response.write "			<tr>" & vbCrLf
				response.write "				<td></td>" & vbCrLf
				response.write "			</tr>" & vbCrLf
			Next
%>
			<tr>
				<!--
				<td></td>
				-->
				<td>상품구분</td>
				<td>수취인명</td>
				<td>수취인EMAIL</td>
				<td>수취인전화1</td>
				<td>수취인전화2</td>
				<td>수취인전화3</td>
				<td>수취인전화4</td>
				<td>수취인전화</td>
				<td>수취인국가코드</td>
				<td>수취인국가명</td>
				<td>수취인우편번호</td>
				<td>수취인 상세주소1</td>
				<td>수취인 상세주소2</td>
				<td>수취인(시/군)</td>
				<td>수취인(주/도)</td>
				<td>수취인 건물명</td>
				<td>총중량(g)</td>
				<td>내용품명</td>
				<td>개수</td>
				<td>순중량(g)</td>
				<td>가격(us$)</td>
				<td>USD</td>
                <td>HSCODE</td>
				<td>생산지</td>
				<td>규격</td>
				<td>보험가입여부</td>
				<td>보험가입금액</td>
				<td>우편물구분</td>
				<td>물품종류</td>
				<td>고객주문번호</td>
				<td>주문인우편번호</td>
				<td>주문인주소</td>
				<td>주문인명</td>
				<td>주문인전화1</td>
				<td>주문인전화2</td>
				<td>주문인전화3</td>
				<td>주문인전화4</td>
				<td>주문인전화</td>
				<td>주문인휴대전화1</td>
				<td>주문인휴대전화2</td>
				<td>주문인휴대전화3</td>
				<td>주문인휴대전화</td>
				<td>주문인EMAIL</td>

				<td>전자상거래여부</td>
				<td>사업자번호</td>
				<td>수출화주</td>
				<td>수출화주 주소</td>
                <td>수출이행등록여부</td>
				<td>수출신고번호1</td>
                <td>전량분할발송여부</td>
                <td>선기적포장개수</td>
				<td>수출신고번호2</td>
                <td>전량분할발송여부</td>
                <td>선기적포장개수</td>
				<td>수출신고번호3</td>
                <td>전량분할발송여부</td>
                <td>선기적포장개수</td>
				<td>수출신고번호4</td>
                <td>전량분할발송여부</td>
                <td>선기적포장개수</td>

				<td>추천우체국코드</td>
                <td>수출면장여부</td>
                <td>브라질세금식별번호</td>
                <td>가로(cm)</td>
                <td>세로(cm)</td>
                <td>높이(cm)</td>
			</tr>
<%
			for i=0 to Ubound(obalju.FBaljuDetailList) -1
%>
			<tr>
				<!--
				<td></td>
				-->
				<td><%=obalju.FBaljuDetailList(i).FitemGubunName%></td>
				<td><%=obalju.FBaljuDetailList(i).FReqName%></td>
				<td><%=obalju.FBaljuDetailList(i).FreqEmail%></td>
				<% if (obalju.FBaljuDetailList(i).FSitename="cnglob10x10") then %>
				<td style="mso-number-format:'\@';"><%= Replace(SplitValue(obalju.FBaljuDetailList(i).FReqHp,"-",0), "+", "") %></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FReqHp,"-",1)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FReqHp,"-",2)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FReqHp,"-",3)%></td>
				<td></td>
				<% else %>
				<td style="mso-number-format:'\@';"><%= Replace(SplitValue(obalju.FBaljuDetailList(i).FreqPhone,"-",0), "+", "") %></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FreqPhone,"-",1)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FreqPhone,"-",2)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FreqPhone,"-",3)%></td>
				<td></td>
			    <% end if %>
				<td><%=obalju.FBaljuDetailList(i).Fdlvcountrycode%></td>
				<td><%=obalju.FBaljuDetailList(i).FcountryNameEn%></td>
				<td style="mso-number-format:'\@';"><%=obalju.FBaljuDetailList(i).Femszipcode%></td>
				<td><%= obalju.FBaljuDetailList(i).FreqAddr1 %></td>
				<td><%= obalju.FBaljuDetailList(i).FreqAddr2 %></td>
				<td></td>
				<td></td>
				<td></td>

				<td><%= obalju.FBaljuDetailList(i).FrealWeight %></td><%'=(obalju.FBaljuDetailList(i).FitemWeigth + 200) ''총중량.%>

				<td>Stationery</td>

				<td>1</td>
				<td><%=obalju.FBaljuDetailList(i).FitemWeigth%></td>
				<td><%=obalju.FBaljuDetailList(i).FitemUsDollar%></td>
                <td>USD</td>
				<td>9609909000</td><% '// Stationery 는 문구류이므로 상품종류를 통칭해서 쓸 수 있다. %>
				<td>KR</td>
				<td></td>
				<td><%=obalju.FBaljuDetailList(i).FInsureYn%></td>
				<td>
				    <% if obalju.FBaljuDetailList(i).FInsureYn="Y" then %>
				    <%=obalju.FBaljuDetailList(i).FItemTotalSum%>
				    <% else %>
				    0
				    <% end if %>
				</td>
				<td>E</td>
				<td></td>

				<td><%=obalju.FBaljuDetailList(i).FOrderserial%></td>
				<td>11154</td>
				<td>83, Yongjeonggyeongje-ro 2-gil, Gunnae-myeon, Pocheon-si, Gyeonggi-do, KOREA</td>
				<td><%=obalju.FBaljuDetailList(i).FReqName%></td>
				<td>82</td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FBuyPhone,"-",0)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FBuyPhone,"-",1)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FBuyPhone,"-",2)%></td>
				<td></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FBuyHp,"-",0)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FBuyHp,"-",1)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FBuyHp,"-",2)%></td>
				<td></td>
				<td><%=obalju.FBaljuDetailList(i).FBuyEmail%></td>

				<td>Y</td>
				<td>2118700620</td>
				<td>TENBYTEN</td>
				<td>83, Yongjeonggyeongje-ro 2-gil, Gunnae-myeon, Pocheon-si, Gyeonggi-do, KOREA</td>
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
				<td></td>

                <td><%=obalju.FBaljuDetailList(i).FboxSizeX%></td>
                <td><%=obalju.FBaljuDetailList(i).FboxSizeY%></td>
                <td><%=obalju.FBaljuDetailList(i).FboxSizeZ%></td>
			</tr>
<%
			Next
			response.write "</table>"

        elseif (songjangdiv="92") then
            '======================================================================
%>
<html lang="ko">
<head>
    <meta charset="euc-kr">
</head>
<body>
<%
			for i=0 to Ubound(obalju.FBaljuDetailList) -1
%>
            <table border=1 width=1100>
            <tr>
                <td colspan="7" align="center" height="40"><h2>COMMERCIAL  INVOICE</h2></td>
            </tr>
            <tr>
                <td colspan="3" width="470" height="30"><b>&nbsp; Shipper / Exporter</b></td>
                <td colspan="4"><b>&nbsp; No. & Date of Invoice</b></td>
            </tr>
            <tr>
                <td colspan="3" height="120" align="left">
                    <br />
                    2118700620<br />
                    TENBYTEN<br />
                    83, Yongjeonggyeongje-ro 2-gil, Gunnae-myeon,<br />
                    Pocheon-si, Gyeonggi-do, KOREA
                </td>
                <td colspan="4" rowspan="5">
                    <br />
                    <b>Date:</b> <%= Left(Now(), 10) %><br /><br /><br />
                    <b>Invoice No:</b><br /><br /><br />
                    <b>PO No:</b> <%=obalju.FBaljuDetailList(i).FOrderserial%><br /><br /><br />
                    <b>Terms of Sale (Incoterm):</b>
                </td>
            </tr>
            <tr>
                <td colspan="3" width="470" height="30"><b>&nbsp; SHIP TO</b></td>
            </tr>
            <tr>
                <td colspan="3" height="120" align="left">
                    <%= obalju.FBaljuDetailList(i).FReqName %><br />
                    <%= obalju.FBaljuDetailList(i).FReqHp %> &nbsp; <%= obalju.FBaljuDetailList(i).FreqPhone %><br />
                    <%= obalju.FBaljuDetailList(i).FreqEmail %><br />
                    <%= obalju.FBaljuDetailList(i).FemsZipCode %><br />
                    <%= obalju.FBaljuDetailList(i).FreqAddr1 %><br />
                    <%= obalju.FBaljuDetailList(i).FreqAddr2 %><br />
                    <%= obalju.FBaljuDetailList(i).FcountryNameEn %> &nbsp; <%= obalju.FBaljuDetailList(i).FprovinceCode %>
                </td>
            </tr>
            <tr>
                <td colspan="3" width="470" height="30"><b>&nbsp; SOLD TO</b></td>
            </tr>
            <tr>
                <td colspan="3" height="120" align="left">
                    SAME AS SHIP TO
                </td>
            </tr>
            <tr height="30">
                <td><b>Port of Loading</b></td>
                <td colspan="2"><b>Final Destination</b></td>
                <td colspan="4" rowspan="4" style="vertical-align: top;">
                    <b>Remark</b><br />
                    ONLY FOR CUSTOMS PURPOSE
                </td>
            </tr>
            <tr height="30">
                <td>KOREA</td>
                <td colspan="2"><%= obalju.FBaljuDetailList(i).FcountryNameEn %></td>
            </tr>
            <tr height="30">
                <td><b>Vessel / Flight</b></td>
                <td colspan="2"><b>Sailing on or About</b></td>
            </tr>
            <tr height="30">
                <td>UPS</td>
                <td colspan="2" align="left"><%= Left(Now(), 10) %></td>
            </tr>
            </table>

            <p />


            <%
            set oCOrderDetail = New CBalju
            oCOrderDetail.getOrderDetailListUPS(obalju.FBaljuDetailList(i).FOrderserial)
            totItemNo = 0
            totItemUsDollar = 0
            %>

            <table border=1 width=1100>
                <tr height="40" align="center">
                    <td colspan="2"></td>
                    <td><b>Description of Goods</b></td>
                    <td><b>Type</b></td>
                    <td><b>Unit($)</b></td>
                    <td><b>QTY</b></td>
                    <td><b>Amount (US$)</b></td>
                </tr>
                <% for j = 0 to Ubound(oCOrderDetail.FOrderDetailList) - 1 %>
                <tr height="30" align="center">
                    <td><%= (j + 1) %></td>
                    <td>EA</td>
                    <td colspan="2"><%= oCOrderDetail.FOrderDetailList(j).Fcatename_e %></td>
                    <td><%= oCOrderDetail.FOrderDetailList(j).FitemUsDollar %></td>
                    <td><%= oCOrderDetail.FOrderDetailList(j).Fitemno %></td>
                    <td><%= (oCOrderDetail.FOrderDetailList(j).FitemUsDollar * oCOrderDetail.FOrderDetailList(j).Fitemno) %></td>
                </tr>
                <%
                	totItemNo = totItemNo + oCOrderDetail.FOrderDetailList(j).Fitemno
                	totItemUsDollar = totItemUsDollar + (oCOrderDetail.FOrderDetailList(j).FitemUsDollar * oCOrderDetail.FOrderDetailList(j).Fitemno)
                next
                %>
                <tr height="30" align="center">
                    <td colspan="4"></td>
                    <td>total</td>
                    <td><%= totItemNo %></td>
                    <td><%= totItemUsDollar %></td>
                </tr>
            </table>
            <%
            Next
            %>
</body>
</html>
<%

		else
            '======================================================================
			response.write "<table border=1>"

			For k = 1 To 11
				response.write "			<tr>" & vbCrLf
				response.write "				<td></td>" & vbCrLf
				response.write "			</tr>" & vbCrLf
			Next
%>
			<tr>
				<td>미반영필드 1</td>
				<td>수취인명</td>
				<td>수취인 우편번호</td>
				<td>수취인 주소</td>
				<td>수취인 주소</td>
				<td>상품명</td>
				<td>수량</td>
				<td>미반영필드 2</td>
				<td>수취인 이동통신</td>
				<td>수취인 전화번호</td>
				<td>주문자명</td>
				<td>주문자 우편번호</td>
				<td>주문자 주소</td>
				<td>주문자 주소</td>
				<td>주문자 전화번호</td>
				<td>주문자 이동통신</td>
				<td>주문번호</td>
				<td>비고</td>
				<td>배송메시지</td>
			</tr>
<%
			for i=0 to Ubound(obalju.FBaljuDetailList) -1
%>
			<tr>
				<td>aaaa</td>
				<td><%=obalju.FBaljuDetailList(i).FReqName%></td>
				<td><%=obalju.FBaljuDetailList(i).FreqZipCode%></td>
				<td><%=obalju.FBaljuDetailList(i).FReqAddr1%></td>
				<td><%=obalju.FBaljuDetailList(i).FReqAddr2%></td>
				<td><%=obalju.FBaljuDetailList(i).FgoodNames%></td>
				<td>1</td>
				<td>aaaa</td>
				<td><%=obalju.FBaljuDetailList(i).FreqHp%></td>
				<td><%=obalju.FBaljuDetailList(i).FreqPhone%></td>
				<td><%=obalju.FBaljuDetailList(i).FBuyName%></td>
				<td><%=obalju.FBaljuDetailList(i).FBuyZipCode%></td>
				<td><%=obalju.FBaljuDetailList(i).FBuyAddr1%></td>
				<td><%=obalju.FBaljuDetailList(i).FBuyAddr2%></td>
				<td><%=obalju.FBaljuDetailList(i).FBuyPhone%></td>
				<td><%=obalju.FBaljuDetailList(i).FBuyHp%></td>
				<td><%=obalju.FBaljuDetailList(i).FOrderserial%></td>
				<td>aaaa</td>
				<td><%=obalju.FBaljuDetailList(i).FEtcStr%></td>
			</tr>
<%
			Next

			response.write "</table>"
		end if
		set obalju = Nothing
		dbget.close	: response.End
	End If
else
	obalju.FRectcancelyn = cancelyn
	obalju.FRectipkumdiv = ipkumdiv
    obalju.getBaljuDetailList idx
end if
%>

<!-- 검색 시작 -->
<form name="frmbalju" method="get" action="" style="margin:0px;">
<input type="hidden" name="editor_no">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="songjangdiv" value="<%= songjangdiv %>">
<input type="hidden" name="gubun" value="<%= gubun %>">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="reload" value="ON">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 출고지시ID : <input type="text" size="8" length="10" value="<%= idx %>" name="idx" onKeyPress="if(event.keyCode==13){baljuSearch();}"/>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="baljuSearch();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 취소여부 : <% drawSelectBoxUsingYN "cancelyn", cancelyn %>
		&nbsp;
		* 거래상태 : <% DrawIpkumDivName "ipkumdiv", ipkumdiv, "" %>
		<% if (songjangdiv="90") then %>
			&nbsp;
			* 중량등록여부 : <% drawSelectBoxUsingYN "realweight", realweight %>
			&nbsp;
			* 운송장등록여부 : <% drawSelectBoxUsingYN "baljusongjangno", baljusongjangno %>
		<% end if %>
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left"></td>
	<td align="right">
		<%
		If songjangdiv = "92" Then
			response.write "<input type='button' value='UPS 엑셀다운(전체)' onClick=""jsDownXL('" & idx & "', '" + songjangdiv + "', '')"" class='button'>"
		End If
		%>
		<%
		If songjangdiv = "90" Then
			response.write "<input type='button' value='EMS & KPACK전송용 엑셀다운(전체)' onClick=""jsDownXL('" & idx & "', '" + songjangdiv + "', '')"" class='button'>"
			response.write "&nbsp;&nbsp;"
			response.write "<input type='button' value='EMS전송용 엑셀다운(2kg 초과만)' onClick=""jsDownXL('" & idx & "', '" + songjangdiv + "', '2kgup')"" class='button'>"
			response.write "&nbsp;&nbsp;"
			response.write "&nbsp;&nbsp;<input type='button' value='K-Packet 전송(2kg 이하)' onClick=""jsSendReq('" + idx + "', 'KPT', '2kgdn')"" disabled class='button'>"
		End If
		%>
		<%
		If songjangdiv = "8" Then
			response.write "<input type='button' value='우체국전송용 엑셀다운' onClick=""jsDownXL('" & idx & "', '" + songjangdiv + "', '')"" class='button'>"
			response.write "&nbsp;&nbsp;"
			response.write "<font color=red>* 다운받은 엑셀파일을 엑셀로 열어서 저장한 후에 우체국에 올리세요.</font>"
		End If
		%>
	</td>
</tr>
<tr>
	<td align="left"></td>
</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			총건수 : <b><%= Ubound(obalju.FBaljuDetailList) %></b>
			&nbsp;
			총금액 : <b><%= FormatNumber(obalju.GetTotalSum,0) %></b>
		</td>
	</tr>

  	<form name="frmview" method="get" style="margin:0px;">
  	<input type="hidden" name="iid" value="<%= idx %>">
  	<input type="hidden" name="menupos" value="<%= menupos %>">

  	<!--
  	<tr bgcolor="FFFFFF">
	  	<td colspan="5">
	  		<input type="button" value="전체선택" onClick="AnSelectAllFrame(true)">
	  		&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" value="선택사항출력" onclick="AnCheckNPrint()">
	  	</td>

	  	<td colspan="10" align="right">
	  		<a href="#" onClick="AnViewUpcheList(frmview)"><font color="#0000FF">[일별 배송리스트]</font></a>
	  	</td>
	</tr>
	-->
	</form>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<!-- <td width="30" align="center">선택</td> -->
		<td width="100">주문번호</td>
		<td width="100">사이트명</td>
		<td width="40">국가</td>
		<td width="120">아이디</td>
		<td width="70">구매자</td>
		<td width="70">수령인</td>
		<td width="70">결제금액</td>
		<td width="70">거래상태</td>
		<td>취소</td>
	<%If ((songjangdiv="90") or (songjangdiv="8")) Then%>
	    <% if (songjangdiv="90") then %>
	    <td width="90">순(상품) 중량</td>
	    <td width="150">실(측정)중량</td>
        <td width="220">박스사이즈</td>
	    <!-- td width="70">실금액</td -->
	    <% end if %>
		<td width="70">송장</td>
	<%End If %>
	</tr>

<% if Ubound(obalju.FBaljuDetailList)<1 then %>
	<tr bgcolor="FFFFFF">
		<td colspan="13" align="center">검색결과가 없습니다.</td>
	</tr>
<% else %>

  	<% for i=0 to Ubound(obalju.FBaljuDetailList) -1 %>
	<form name="frmBuyPrc_<%= obalju.FBaljuDetailList(i).FOrderSerial %>" method="post" action="popBaljuSongjangInput.asp" style="margin:0px;" >

	<input type="hidden" name="orderserial" value="<%= obalju.FBaljuDetailList(i).FOrderSerial %>">

	<input type="hidden" name="songjangdiv" value="<%=songjangdiv%>">
	<input type="hidden" name="idx" value="<%=idx%>">
	<input type="hidden" name="mode" value="">


	<tr bgcolor="FFFFFF">
		<!-- <td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td> -->
		<td align="center"><a href="#" onclick="ViewOrderDetail('<%= obalju.FBaljuDetailList(i).FOrderSerial %>')"><%= obalju.FBaljuDetailList(i).FOrderserial %></a></td>
		<td align="center"><%= obalju.FBaljuDetailList(i).FSiteName %></td>
		<td align="center"><%= obalju.FBaljuDetailList(i).Fdlvcountrycode %></td>
		<td align="center"><%= obalju.FBaljuDetailList(i).FUserID %></td>
		<td align="center"><%= obalju.FBaljuDetailList(i).FBuyName %></td>
		<td align="center"><%= obalju.FBaljuDetailList(i).FReqName %></td>
		<td align="right"><%= FormatNumber(obalju.FBaljuDetailList(i).FSubTotalPrice,0) %></td>
		<td align="center"><font color="<%= obalju.FBaljuDetailList(i).IpkumDivColor %>"><%= IpkumDivName(obalju.FBaljuDetailList(i).Fipkumdiv) %></font></td>
		<td align="center"><font color="<%= obalju.FBaljuDetailList(i).CancelYnColor %>"><%= obalju.FBaljuDetailList(i).CancelYnName %></font></td>
	<%If ((songjangdiv="90") or (songjangdiv="8")) Then%>
	    <% if (songjangdiv="90") then %>
	    <td align="right"><%= obalju.FBaljuDetailList(i).FitemWeigth%> g</td>
	    <td >
	        <input type="text" name="realweight" value="<%= obalju.FBaljuDetailList(i).FrealWeight %>" size="6" maxlength="6" style="text-align:right">(g)
	        <input type="button" value="저장" onClick="saveTotalWeight(this.form)" class='button'>
	    </td>
	    <td >
            <input type="text" class="text" name="boxSizeX" value="<%= obalju.FBaljuDetailList(i).FboxSizeX %>" size=2 AUTOCOMPLETE="off" style="text-align:right">
            *
            <input type="text" class="text" name="boxSizeY" value="<%= obalju.FBaljuDetailList(i).FboxSizeY %>" size=2 AUTOCOMPLETE="off" style="text-align:right">
            *
            <input type="text" class="text" name="boxSizeZ" value="<%= obalju.FBaljuDetailList(i).FboxSizeZ %>" size=2 AUTOCOMPLETE="off" style="text-align:right">
            (cm)
	        <input type="button" value="저장" onClick="saveBoxSize(this.form)" class='button'>
	    </td>
	    <!-- td ><%= obalju.FBaljuDetailList(i).FrealDlvPrice %></td -->
	    <% end if %>
		<td align="center">
		<%If obalju.FBaljuDetailList(i).FIpkumdiv >= "7" Then %>
			<%=obalju.FBaljuDetailList(i).FsongjangNo%>
		<%Else %>
			<input type="text" name="songjangNo" value="<%=obalju.FBaljuDetailList(i).FsongjangNo%>">
			<input type="button" value="송장입력" onClick="saveSongjang(this.form)" class='button'>
		<%End If %>
		</td>
	<%End If %>
	</tr>
	</form>
	<% next %>

<% end if %>
</table>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width="100%" frameborder="0" scrolling="no"></iframe>
<% end if %>
<%
set obalju = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
