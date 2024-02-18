<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/outmall_function.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsanDiffcls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
dim sellsite, yyyy1, mm1 ,yyyy2, mm2, yyyy3, mm3, yyyy4, mm4
dim scmjsdate, scmdeliverdate, omjsdate, scmactdate,sItemDiv
dim chksmdt, chksmdeliverdt, chkomdt, chkactdt
dim rdosmdt, rdosmdeliverdt, rdoomdt, rdoactdt
dim  clsJS, arrList, intLoop, arrSum, intS
dim scmitemno,omitemno,scmsellprice,scmmeachul,scmbuycash,omsellprice,ommeachul,ombuycash, extTenCouponPrice,  extOwnCouponPrice, reducedPrice,	allAtDiscount
dim sISMYN
dim iCurrpage,iPageSize,iPerCnt,isortType,sOrderserial,iTotPage,iTotCnt

    sISMYN ="N"
    sellsite = requestCheckVar(Request("sellsite"),32)
    scmactdate = requestCheckVar(Request("actdt"),7)
    omJsDate = requestCheckVar(Request("omdt"),7)
    scmjsdate = requestCheckVar(Request("smdt"),7)
    scmdeliverdate = requestCheckVar(Request("scmdeliverdate"),7)
    sItemDiv = requestCheckVar(Request("itemdv"),1)
    chkactdt = requestCheckVar(Request("chkactdt"),1)
    chksmdt = requestCheckVar(Request("chksmdt"),1)
    chksmdeliverdt = requestCheckVar(Request("chksmdeliverdt"),1)
    chkomdt= requestCheckVar(Request("chkomdt"),1)
    rdoactdt = requestCheckVar(Request("rdoactdt"),32)
    rdosmdeliverdt = requestCheckVar(Request("rdosmdeliverdt"),32)
    rdosmdt = requestCheckVar(Request("rdosmdt"),32)
    rdoomdt= requestCheckVar(Request("rdoomdt"),32)

 	yyyy1 = requestCheckVar(Request("yyyy1"),4)
  	mm1 = requestCheckVar(Request("mm1"),2)
    yyyy2 = requestCheckVar(Request("yyyy2"),4)
   	mm2 = requestCheckVar(Request("mm2"),2)
    yyyy3 = requestCheckVar(Request("yyyy3"),4)
   	mm3 = requestCheckVar(Request("mm3"),2)
    yyyy4 = requestCheckVar(Request("yyyy4"),4)
   	mm4 = requestCheckVar(Request("mm4"),2)


    '// 초기 파라미터
    '// extJungsanMatchingItem.asp?sellsite=11st1010&actdt=2021-06&omdt=&smdt=2021-07-01&itemdv=I

    if (scmactdate <> "") then
        chkActDT = "Y"
        if (scmactdate = "NOMATCH") or (scmactdate = "N") then
            rdoActDT = "NOMATCH"
        else
            scmactdate 	= Left(scmactdate, 7)
            yyyy1 		= year(scmactdate & "-01")
            mm1 		= month(scmactdate & "-01")
        end if
    end if

    if (scmjsdate <> "") then
        chkSMDT = "Y"
        if (scmjsdate = "NOMATCH") or (scmjsdate = "N") then
            rdoSMDT = "NOMATCH"
        else
            scmjsdate 	= Left(scmjsdate, 7)
            yyyy2 		= year(scmjsdate & "-01")
            mm2 		= month(scmjsdate & "-01")
        end if
    end if

    if (scmdeliverdate <> "") then
        chkSMDeliverDT = "Y"
        if (scmdeliverdate = "NOMATCH") or (scmdeliverdate = "N") then
            rdoSMDeliverDT = "NOMATCH"
        else
            scmdeliverdate 	= Left(scmdeliverdate, 7)
            yyyy4 			= year(scmdeliverdate & "-01")
            mm4 			= month(scmdeliverdate & "-01")
        end if
    end if

    if (omJsDate <> "") then
        chkOMDT = "Y"
        if (omJsDate = "NOMATCH") or (omJsDate = "N") then
            rdoOMDT = "NOMATCH"
        else
            omJsDate 	= Left(omJsDate, 7)
            yyyy3 		= year(omJsDate & "-01")
            mm3 		= month(omJsDate & "-01")
        end if
    end if

 	if sellsite ="" then sellsite ="ssg"

    if yyyy1 <> "" and mm1 <> "" then scmactdate 		= yyyy1 & "-" & Format00(2,mm1)
    if yyyy2 <> "" and mm2 <> "" then scmJsDate 		= yyyy2 & "-" & Format00(2,mm2)
    if yyyy4 <> "" and mm4 <> "" then scmdeliverdate 	= yyyy4 & "-" & Format00(2,mm4)
    if yyyy3 <> "" and mm3 <> "" then omJsDate 			= yyyy3 & "-" & Format00(2,mm3)

	iCurrpage = requestCheckVar(Request("iCP"),4)
	isortType = requestCheckVar(Request("iST"),1)
	sOrderserial = requestCheckVar(Request("sorderserial"),32)

	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF


	iPageSize = 300		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	if isortType = "" THEN isortType =1

    set clsJS = new CextJungsanMapping
        clsJS.FRectOutMall 			= sellsite

        if (chkActDT = "Y") then
            if (rdoActDT = "NOMATCH") then
                clsJS.FRectyyyymm 			= "N"
            else
                clsJS.FRectyyyymm 			= scmactdate
            end if
        end if

        if (chkSMDT = "Y") then
            if (rdoSMDT = "NOMATCH") then
                clsJS.FRectscmJsDate 			= "N"
            else
                clsJS.FRectscmJsDate 			= scmJsDate
            end if
        end if

        if (chkSMDeliverDT = "Y") then
            if (rdoSMDeliverDT = "NOMATCH") then
                clsJS.FRectscmDeliverDate 			= "N"
            else
                clsJS.FRectscmDeliverDate 			= scmdeliverdate
            end if
        end if

        if (chkOMDT = "Y") then
            if (rdoOMDT = "NOMATCH") then
                clsJS.FRectomJsDate 			= "N"
            else
                clsJS.FRectomJsDate 			= omJsDate
            end if
        end if

        clsJS.FRectItemDiv 			= sItemDiv
        clsJS.FRectIsMYN 			= sISMYN
        clsJS.FRectOrderserial 		= sOrderserial
        clsJS.FRectSort 			= isortType
        clsJS.FPSize				= iPageSize
        clsJS.FCPage 				= iCurrpage
    arrList = clsJS.fnGetextMatchingItem
	iTotCnt = clsJS.FTotCnt

	clsJS.fnGetextMatchingItemSUM
	    scmitemno 			= clsJS.Fscmitemno
	    omitemno 			= clsJS.Fomitemno
	    scmsellprice 		= clsJS.Fscmsellprice
	    scmmeachul 			= clsJS.Fscmmeachul
	    scmbuycash 			= clsJS.Fscmbuycash
	    omsellprice 		= clsJS.Fomsellprice
	    ommeachul 			= clsJS.Fommeachul
	    ombuycash			= clsJS.Fombuycash
	    extTenCouponPrice	= clsJS.FextTenCouponPrice
	    extOwnCouponPrice	= clsJS.FextOwnCouponPrice
	    reducedPrice		= clsJS.FreducedPrice
	    allAtDiscount		= clsJS.FallAtDiscount

    set clsJS = nothing

    iTotPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<Script type="text/javascript">
function jsChkDate(sType){
	if (sType=="A") {
		if (!frm.chkActDT.checked ){
			$("#yyyy1").val("");
			$("#mm1").val("") ;
			frm.rdoActDT[0].checked = true;
			frm.rdoActDT[1].checked = false;
		}
	} else if (sType=="S") {
		if (!frm.chkSMDT.checked ){
			$("#yyyy2").val("") ;
			$("#mm2").val("") ;
			frm.rdoSMDT[0].checked = true;
			frm.rdoSMDT[1].checked = false;
		}
	} else if (sType=="O") {
		if (!frm.chkOMDT.checked ){
			$("#yyyy3").val("") ;
			$("#mm3").val("") ;
			frm.rdoOMDT[0].checked = true;
			frm.rdoOMDT[1].checked = false;
		}
	} else if (sType=="D") {
		if (!frm.chkSMDeliverDT.checked ){
			$("#yyyy4").val("") ;
			$("#mm4").val("") ;
			frm.rdoSMDeliverDT[0].checked = true;
			frm.rdoSMDeliverDT[1].checked = false;
		}
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		&nbsp;
		제휴몰:	<% fnGetOptOutMall sellsite %>
		&nbsp;
		&nbsp;&nbsp;
	   구분:
	   <select name="Itemdv">
		<option value="" <%if sitemdiv ="" then%>selected<%end if%>>-전체-</option>
		<option value="I" <%if sitemdiv ="I" then%>selected<%end if%>>상품</option>
		<option value="D" <%if sitemdiv ="D" then%>selected<%end if%>>배송비</option>
	   </select>
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		<input type="checkbox"   name="chkActDT"  value="Y" <%= CHKIIF(chkActDT="Y", "checked", "") %>>
		결제일:
		<input type="radio" id="rdoActDT"  name="rdoActDT"  value="Y" <%= CHKIIF(rdoActDT<>"NOMATCH", "checked", "") %>>
		<% DrawYMSelBox "yyyy1","mm1",yyyy1,mm1 %>
		<input type="radio"  id="rdoActDT"   name="rdoActDT"  value="NOMATCH" <%= CHKIIF(rdoActDT="NOMATCH", "checked", "") %>>
		미매칭&nbsp;&nbsp;


		<input type="checkbox"   name="chkSMDT"  value="Y" <%= CHKIIF(chkSMDT="Y", "checked", "") %>>
		10x10출고일:
		<input type="radio" id="rdoSMDT"  name="rdoSMDT"  value="Y" <%= CHKIIF(rdoSMDT<>"NOMATCH", "checked", "") %>>
		<% DrawYMSelBox "yyyy2","mm2",yyyy2,mm2 %>
		<input type="radio"  id="rdoSMDT" name="rdoSMDT"  value="NOMATCH" <%= CHKIIF(rdoSMDT="NOMATCH", "checked", "") %>>
		미출고&nbsp;&nbsp;


        <input type="checkbox"   name="chkSMDeliverDT"  value="Y" <%= CHKIIF(chkSMDeliverDT="Y", "checked", "") %>>
		10x10정산일:
		<input type="radio" id="rdoSMDeliverDT"  name="rdoSMDeliverDT"  value="Y" <%= CHKIIF(rdoSMDeliverDT<>"NOMATCH", "checked", "") %>>
		<% DrawYMSelBox "yyyy4","mm4",yyyy4,mm4 %>
		<input type="radio"  id="rdoSMDT" name="rdoSMDeliverDT"  value="NOMATCH" <%= CHKIIF(rdoSMDeliverDT="NOMATCH", "checked", "") %>>
		정산이전&nbsp;&nbsp;


		<input type="checkbox" name="chkOMDT"  value="Y" <%= CHKIIF(chkOMDT="Y", "checked", "") %>>
		제휴정산일:
		<input type="radio" id="rdoOMDT"  name="rdoOMDT"  value="Y" <%= CHKIIF(rdoOMDT<>"NOMATCH", "checked", "") %>>
		<% DrawYMSelBox "yyyy3","mm3",yyyy3,mm3 %>&nbsp;
		<input type="radio" id="rdoOMDT"  name="rdoOMDT"  value="NOMATCH" <%= CHKIIF(rdoOMDT="NOMATCH", "checked", "") %>>
		미정산
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		(원)주문번호검색: <input type="text" class="text" name="sorderserial" style="width:150px;" value="<%=sorderserial%>">
		&nbsp;정렬:<select name="iST" class="select">
		<option value="1" <%if isortType="1" or isortType="" then%>selected<%end if%>>결제일</option>
		<option value="2" <%if isortType="2" then%>selected<%end if%>>10x10출고일</option>
		<option value="3" <%if isortType="3" then%>selected<%end if%>>제휴정산일</option>
		<option value="4" <%if isortType="4" then%>selected<%end if%>>주문번호</option>
		</select>
	</td>
</tr>
</form>
</table>
<p>총 <%=iTotCnt%>건 /<%=iTotPage%>페이지
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

	<Tr bgcolor="#E6E6E6" align="center">
		<td rowspan="2">제휴몰</td>
		<td rowspan="2">결제일</td>
		<td rowspan="2">10x10출고일</td>
        <td rowspan="2">10x10정산일</td>
		<td rowspan="2">제휴정산일</td>
		<td rowspan="2">주문번호</td>
		<td rowspan="2">브랜드</td>
		<td rowspan="2">상품코드</td>
		<td rowspan="2">옵션코드</td>
		<td colspan="2">수량</td>
		<!--<td colspan="2">소비자가</td>-->
		<td colspan="2">판매가</td>
		<td colspan="2">10x10쿠폰</td>
		<td colspan="2">제휴쿠폰</td>
		<td colspan="2">매출금액</td>
		<td colspan="2">정산금액</td>
		<td rowspan="2">제휴주문번호</td>
		<td rowspan="2">제휴주문순번</td>
		<td rowspan="2">+/-취소매칭</td>
	</tr>
	<tr bgcolor="#E6E6E6" align="center">
		<td>10x10</td>
		<td>제휴몰</td>
		<!--<td>10x10</td>
		<td>제휴몰</td>-->
		<td>10x10</td>
		<td>제휴몰</td>
		<td>10x10</td>
		<td>제휴몰</td>
		<td>10x10</td>
		<td>제휴몰</td>
		<td>10x10</td>
		<td>제휴몰</td>
		<td>10x10</td>
		<td>제휴몰</td>
	</tr>
	<tr bgcolor="#ffffff" align="right">
		<td colspan="9" align="center">합계</td>
		<td><%=formatnumber(scmitemno,0)%></td>
		<td><%=formatnumber(omitemno,0)%></td>
		<td><%=formatnumber(scmsellprice,0)%></td>
		<td><%=formatnumber(omsellprice,0)%></td>

		<td><%=formatnumber(scmsellprice-reducedPrice,0)%></td>
		<td><%=formatnumber(extTenCouponPrice,0)%></td>
		<td></td>
		<td><%=formatnumber(extOwnCouponPrice,0)%></td>
		<td><%=formatnumber(scmmeachul,0)%></td>
		<td><%=formatnumber(ommeachul,0)%></td>
		<td> </td>
		<td></td>
		<td> </td>
		<td></td>
		<td></td>
	</tr>
	<%if isArray(arrList) then
		for intLoop = 0 To UBound(arrList,2)
	%>
	<tr bgcolor="#ffffff" align="center">
		<td><%=arrList(0,intLoop)%></td>
		<td><%=arrList(1,intLoop)%></td>
		<td><%=arrList(2,intLoop)%></td>
		<td><%=arrList(24,intLoop)%></td>
        <td><%=arrList(3,intLoop)%></td>
		<td><%=arrList(5,intLoop)%>
		<%if arrList(4,intLoop) <> arrList(5,intLoop) then%>
			[<%=arrList(4,intLoop)%>]
		<%end if%>
		</td>
		<td><%=arrList(6,intLoop)%></td>
		<td><%=arrList(7,intLoop)%></td>
		<td><%=arrList(8,intLoop)%></td>

		<td align="right"><span <%if arrList(10,intLoop) <> arrList(9,intLoop) then%>style="color:blue"<%end if%>><%=formatnumber(arrList(9,intLoop),0)%></span></td>
		<td align="right"><span <%if arrList(10,intLoop) <> arrList(9,intLoop) then%>style="color:blue"<%end if%>><%=formatnumber(arrList(10,intLoop),0)%></span></td>
		<td align="right"><span <%if arrList(11,intLoop) <> arrList(14,intLoop) then%>style="color:blue"<%end if%>><%=formatnumber(arrList(11,intLoop),0)%></span></td>
		<td align="right"><span <%if arrList(11,intLoop) <> arrList(14,intLoop) then%>style="color:blue"<%end if%>><%=formatnumber(arrList(14,intLoop),0)%></span></td>

		<td align="right"><span <%if arrList(11,intLoop)-arrList(21,intLoop) <> arrList(19,intLoop) then%>style="color:blue"<%end if%>><%=formatnumber(arrList(11,intLoop)-arrList(21,intLoop),0)%></span></td>
		<td align="right"><span <%if arrList(11,intLoop)-arrList(21,intLoop) <> arrList(19,intLoop) then%>style="color:blue"<%end if%>><%=formatnumber(arrList(19,intLoop),0)%></span></td>
		<td align="right"></td>
		<td align="right"><%=formatnumber(arrList(20,intLoop),0)%></td>

		<td align="right"><span <%if arrList(12,intLoop) <> arrList(15,intLoop) then%>style="color:blue"<%end if%>><%=formatnumber(arrList(12,intLoop),0)%></span></td>
		<td align="right"><span <%if arrList(12,intLoop) <> arrList(15,intLoop) then%>style="color:blue"<%end if%>><%=formatnumber(arrList(15,intLoop),0)%></span></td>
		<td align="right"></td>
		<td align="right"></td>
		<td><%=arrList(17,intLoop)%></td>
		<td><%=arrList(18,intLoop)%></td>
		<td><%=arrList(23,intLoop)%></td>
	</tr>
	<%	next
	end if%>
</table>
<!-- 페이징처리 -->
<table width="100%" cellpadding="10">
	<tr>
		<td align="center">
 			<%sbDisplayPaging "iCP", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
