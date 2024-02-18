<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  사은품 관리
' History : 2008.04.01 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/itemgiftcls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
'Call fnSetEventCommonCode '공통코드 어플리케이션 변수에 세팅

 Dim eCode
 Dim clsGift, arrList, intLoop
 Dim iTotCnt
 Dim iPageSize, iCurrpage ,iDelCnt
 Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
 Dim iSerachType,sSearchTxt,sBrand,  sDate,sSdate,sEdate,igStatus,sgDelivery
 Dim strParm

 eCode     		= requestCheckVar(Request("eC"),10)			'이벤트 코드
 iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
 sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'검색어
 sBrand     	= requestCheckVar(Request("ebrand"),32)		'브랜드
 sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
 sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
 sEdate     	= requestCheckVar(Request("iED"),10)		'종료일
 igStatus		= requestCheckVar(Request("giftstatus"),4)	'사은품 상태
 sgDelivery		= requestCheckVar(Request("selDelivery"),1)	'배송정보

 if igStatus="" then igStatus="6" end if
 if sgDelivery="" then sgDelivery="N" end if

 iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호

	IF iCurrpage = "" THEN	iCurrpage = 1
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격

	IF Cstr(eCode) = "0" THEN eCode = ""

	IF (eCode <> "" AND sSearchTxt = "") THEN
		iSerachType = "2"
		sSearchTxt = eCode
	ELSEIF 	(iSerachType="2" AND sSearchTxt <> "") THEN
		eCode = sSearchTxt
	END IF

    strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&ebrand="&sBrand&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&giftstatus="&igStatus
	set clsGift = new CGift
		clsGift.FECode = eCode
		clsGift.FSearchType = iSerachType
 		clsGift.FSearchTxt  = sSearchTxt
 		clsGift.FBrand		= sBrand
 		clsGift.FDateType   = sDate
 		clsGift.FSDate		= sSdate
 		clsGift.FEDate		= sEdate
 		clsGift.FGStatus	= igStatus
 		clsGift.FGDelivery	= sgDelivery

	 	clsGift.FCPage 		= iCurrpage
	 	clsGift.FPSize 		= iPageSize

		arrList = clsGift.fnGetGiftList	'데이터목록 가져오기
 		iTotCnt = clsGift.FTotCnt	'전체 데이터  수
 	set clsGift = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수

	'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
	Dim  arrgiftscope, arrgifttype,arrgiftstatus
	arrgiftscope 	= fnSetCommonCodeArr("giftscope",False)
	arrgifttype 	= fnSetCommonCodeArr("gifttype",False)
	arrgiftstatus 	= fnSetCommonCodeArr("giftstatus",False)

%>
<script language="JavaScript" src="/js/ttpbarcode.js"></script>
<script language="javascript">
<!--
	//달력
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//수정
	function jsMod(gcode){
		location.href = "giftMod.asp?gC="+gcode+"&menupos=<%=menupos%>&<%=strParm%>";
	}

	//페이징처리
		function jsGoPage(iP){
		document.frmSearch.iC.value = iP;
		document.frmSearch.submit();
	}

	//이동
	function jsGoURL(type,ival){
		if(type=="e"){
			location.href = "/admin/eventmanage/event/event_modify.asp?eC="+ival;
		}
	}

	//상품설정별 페이지이동
	function jsItem(giftscope,gCode, eCode){
		//이벤트등록상품, 선택상품일떄 상품 view, 그외 페이지이동
		if(giftscope == 2 || giftscope == 4 ){
			location.href = "/admin/eventmanage/event/eventitem_regist.asp?eC="+eCode+"&menupos=870";
		}else if(giftscope==5){
			location.href = "giftItemReg.asp?gC="+gCode+"&menupos=<%=menupos%>&<%=strParm%>";
		}
	}

//-->

function DrawReceiptPrintobj_TEC(elementid,printname){
        var objstring = "";
        var e;
        objstring = '<OBJECT name="' + elementid + '" ';
        objstring = objstring + ' classid="clsid:E76C9051-A8C4-458E-9F60-3C14DB9EECF9" ';
        objstring = objstring + ' codebase="http://billyman/Tec_dol.cab#version=1,5,0,0" ';
        objstring = objstring + ' width=0 ';
        objstring = objstring + ' height=0 ';
        objstring = objstring + ' align=center ';
        objstring = objstring + ' hspace=0 ';
        objstring = objstring + ' vspace=0 ';
        objstring = objstring + ' > ';
        objstring = objstring + ' <PARAM Name="PrinterName" Value="' + printname + '"> ';
        objstring = objstring + ' </OBJECT>';

        document.write(objstring);
}

/*
function eventindexprint(ievt_code, ievt_name01, ievt_name02, ievt_startdate, ievt_enddate, ievt_gift_code, ievt_gift_kind, ievt_gift01, ievt_gift02){
	var X = 1;
	var Y = 1;
	var F = 1;

	// TEC_DO3 : 452
	if (TEC_DO3.IsDriver == 1){
           X = 1.05;
           Y = 1.05;
           F = 1.2;

			TEC_DO3.SetPaper(900,600);
			TEC_DO3.OffsetX = 20;
			TEC_DO3.OffsetY = 20;
			TEC_DO3.PrinterOpen();



			TEC_DO3.PrintText(500*X, 30*Y, "Arial Bold", 100*F, 0, 0, ievt_code);

			TEC_DO3.PrintText(50*X, 50*Y, "HY견고딕", 30*F, 0, 0, "[시작일]");
			TEC_DO3.PrintText(250*X, 50*Y, "HY견고딕", 30*F, 1, 0, ievt_startdate);

			TEC_DO3.PrintText(50*X, 100*Y, "HY견고딕", 30*F, 0, 0, "[종료일]");
			TEC_DO3.PrintText(250*X, 100*Y, "HY견고딕", 30*F, 1, 0, ievt_enddate);

			TEC_DO3.PrintText(50*X, 150*Y, "HY견고딕", 30*F, 0, 0, "[이벤트명]");
			TEC_DO3.PrintText(250*X, 150*Y, "HY견고딕", 30*F, 1, 0, ievt_name01);
			TEC_DO3.PrintText(250*X, 200*Y, "HY견고딕", 30*F, 1, 0, ievt_name02);

			TEC_DO3.PrintText(50*X, 250*Y, "HY견고딕", 30*F, 0, 0, "----------------------------------------");



			TEC_DO3.PrintText(50*X, 300*Y, "HY견고딕", 30*F, 0, 0, "[사은품]");

			TEC_DO3.PrintText(50*X, 330*Y, "Arial Bold", 100*F, 1, 0, ievt_gift_code);
			TEC_DO3.PrintText(300*X, 350*Y, "HY견고딕", 30*F, 1, 0, ievt_gift_kind);	<!-- 한줄에 한글 24자까지 넘으면 아래로 -->
			TEC_DO3.PrintText(300*X, 400*Y, "HY견고딕", 30*F, 1, 0, ievt_gift01);
			TEC_DO3.PrintText(300*X, 450*Y, "HY견고딕", 30*F, 1, 0, ievt_gift02);

			TEC_DO3.PrinterClose();


    }else window.status = "TEC B-452 드라이버를 설치해 주세요"
}

DrawReceiptPrintobj_TEC("TEC_DO3","TEC B-452");
*/

function eventIndexBarcodePrint(eventCode, eventName01, eventName02, eventStartdate, eventEnddate, eventGiftCode, eventGiftKind, eventGift01, eventGift02) {
	// /js/barcode.js 참조
	if (initTTPprinter("TTP-243_80x50", "T", "N", "                         www.10x10.co.kr                         ", "Y", "￦", "Y", 3, 0) != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[4]');
		return;
	}

	printTTPOneIndexBarcodeForEventItem(eventCode, eventName01, eventName02, eventStartdate, eventEnddate, eventGiftCode, eventGiftKind, eventGift01, eventGift02, 1);
}
</script>
<!---- 검색 ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmSearch" method="get"  action="giftList_Logistics.asp" onSubmit="return jsSearch(this,'E');">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="iC">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<select class="select" name="selType">
				<option value="1" <%IF Cstr(iSerachType) = "1" THEN%>selected<%END IF%>>사은품코드</option>
				<option value="2" <%IF Cstr(iSerachType) = "2" THEN%>selected<%END IF%>>이벤트코드</option>
			</select>
			<input type="text" class="text" name="sTxt" value="<%=sSearchTxt%>" size="10" maxlength="10">
			&nbsp;
			브랜드:<% drawSelectBoxDesignerwithName "ebrand", sBrand %>
			<br>
			기간:
			<select class="select" name="selDate">
				<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>시작일 기준</option>
				<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>종료일 기준</option>
			</select>
			<input type="text" class="text" size="10" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:hand;">
			~ <input type="text" class="text" size="10" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:hand;">
			&nbsp;
			상태:<%sbGetOptCommonCodeArr "giftstatus", igStatus, True,False,"onChange='javascript:document.frmSearch.submit();'"%>
			&nbsp;
			배송:
			<select class="select" name="selDelivery" onChange="javascript:document.frmSearch.submit();">
				<option value="">전체</option>
				<option value="Y" <%IF sgDelivery="Y" THEN%>selected<%END IF%>>업체</option>
				<option value="N" <%IF sgDelivery="N" THEN%>selected<%END IF%>>텐바이텐</option>
			</select>
		</td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frmSearch.submit();">
		</td>
	</tr>
</table>
<!---- /검색 ---->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="17">검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50">사은품코드</td>
    	<td width="50">이벤트<br>코드</br>(그룹)</td>
    	<td>이벤트명</td>
        <td>상품코드</td>
    	<td>브랜드</td>
    	<td>증정대상</td>
    	<td>증정조건</td>
    	<td>이상</td>
    	<td>미만</td>
    	<td>수량</td>
    	<td>종류</td>
    	<td>시작일</td>
    	<td>종료일</td>
    	<td>상태</td>
    	<td>한정</td>
    	<td>배송</td>
    	<td width="40">인덱스<br>출력</td>
    </tr>
    <%IF isArray(arrList) THEN
    	For intLoop = 0 To UBound(arrList,2)
    %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=arrList(0,intLoop)%></a></td>
    	<td nowrap><%IF arrList(3,intLoop) > 0 THEN%><a href="javascript:jsGoURL('e',<%=arrList(3,intLoop)%>)" title="이벤트 정보수정"><%=arrList(3,intLoop)%></a><%IF arrList(4,intLoop) > 0 THEN%><br>(<%=arrList(4,intLoop)%>)<%END IF%><%END IF%></td>
    	<td align="left"><%=db2html(arrList(1,intLoop))%></td>
        <td align="center">
            <% if Not IsNull(arrList(27,intLoop)) and Not IsNull(arrList(28,intLoop)) and Not IsNull(arrList(29,intLoop)) then %>
            <%= BF_MakeTenBarcode(arrList(27,intLoop), arrList(28,intLoop), arrList(29,intLoop)) %>
            <% end if %>
        </td>
    	<td><%=db2html(arrList(5,intLoop))%></td>
    	<td> <%IF (arrList(2,intLoop) = 2 or arrList(2,intLoop) = 4 or arrList(2,intLoop) = 5) then %>
    		<a href="javascript:jsItem(<%=arrList(2,intLoop)%>,<%=arrList(0,intLoop)%>,<%=arrList(3,intLoop)%>)" title="등록상품 수정"><%=fnGetCommCodeArrDesc(arrgiftscope,arrList(2,intLoop))%><br>(<%=arrList(20,intLoop)%>)</a>
    		<%else%>
    		<%=fnGetCommCodeArrDesc(arrgiftscope,arrList(2,intLoop))%>
    		<%end if%>
    		</td>
    	<td><%=fnGetCommCodeArrDesc(arrgifttype,arrList(6,intLoop))%></td>
    	<td nowrap><%=formatnumber(arrList(7,intLoop),0)%></td>
    	<td nowrap><%=formatnumber(arrList(8,intLoop),0)%></td>
    	<td nowrap><%=arrList(11,intLoop)%></td>
    	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%IF arrList(9,intLoop) > 0 THEN%>[<%=arrList(9,intLoop)%>]<%=arrList(19,intLoop)%><%END IF%></a></td>
    	<td nowrap><%=arrList(13,intLoop)%></td>
    	<td nowrap><%=arrList(14,intLoop)%></td>
    	<td nowrap><%=fnGetCommCodeArrDesc(arrgiftstatus,arrList(15,intLoop))%></td>
    	<td nowrap><%IF arrList(12,intLoop) > 0 THEN%><%=arrList(12,intLoop)%><%END IF%></td>
    	<td nowrap><%IF arrList(21,intLoop)="Y" THEN%><font color="#F08050">업체</font><%ELSE%><font color="#5080F0">텐바이텐</font><%END IF%></td>
    	<td>
    	    <!--
    	    <input type="button" value="출력" class="button" onClick="eventindexprint('<%=arrList(3,intLoop)%>', '<%= left(db2html(arrList(1,intLoop)),20) %>','<%= mid(db2html(arrList(1,intLoop)),21) %>','<%=arrList(13,intLoop)%>','<%=arrList(14,intLoop)%>','<%=arrList(0,intLoop)%>','[<%=arrList(9,intLoop)%>]','<%= left(arrList(19,intLoop),20) %>','<%= mid(arrList(19,intLoop),21) %>')">
    	    -->
    	    <!-- eventindexprint('0', '디자인문구_RECYCLE LETTER','ING ver.3 구매 시 노트증정','2014-02-19','2014-03-16','13688','[17466]','CLASS NOTE ver.8 증정(','색상랜덤)') -->
    	    <input type="button" class="button" value="출력" onClick="eventIndexBarcodePrint('<%=arrList(3,intLoop)%>', '<%= left(db2html(arrList(1,intLoop)),23) %>','<%= mid(db2html(arrList(1,intLoop)),24) %>','<%=arrList(13,intLoop)%>','<%=arrList(14,intLoop)%>','<%=arrList(0,intLoop)%>','[<%=arrList(9,intLoop)%>]','<%= left(arrList(19,intLoop),23) %>','<%= mid(arrList(19,intLoop),24) %>')">

    	</td>

    </tr>
	<% Next
	ELSE
	%>
	<tr>
		<td colspan="17" align="center" bgcolor="#FFFFFF">등록된 내용이 없습니다.</td>
	</tr>
	<%END IF%>
</table>
<!-- 페이징처리 -->
<%
iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1

If (iCurrpage mod iPerCnt) = 0 Then
	iEndPage = iCurrpage
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td valign="bottom" align="center">
         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
		<% else %>[pre]<% end if %>
        <%
			for ix = iStartPage  to iEndPage
				if (ix > iTotalPage) then Exit for
				if Cint(ix) = Cint(iCurrpage) then
		%>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
		<%		else %>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
		<%
				end if
			next
		%>
    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
		<% else %>[next]<% end if %>
        </td>
    </tr>
    </form>
</table>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
