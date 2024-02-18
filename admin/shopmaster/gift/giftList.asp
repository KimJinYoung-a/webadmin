<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  사은품 관리
' History : 2008.04.01 정윤정 생성
'			2013.11.11 한용민 수정
'			2015.12.11 허진원 수정- 상품코드 검색 조건 추가
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
<!-- #include virtual="/lib/classes/event/openGiftCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
'Call fnSetEventCommonCode '공통코드 어플리케이션 변수에 세팅
Dim clsGift, arrList, intLoop, iTotCnt, iPageSize, iCurrpage ,iDelCnt, eCode, runoutrate90up
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt, strParm, fcSc, tmpTitle, tmpGift, tmpGiftForTitle
Dim iSerachType,sSearchTxt,sGiftName,sBrand,  sDate,sSdate,sEdate,igStatus,sgDelivery
dim Category, CategoryMid, DispCategory, iItemid
	eCode     		= requestCheckVar(Request("eC"),10)			'이벤트 코드
	iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
	sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'검색어
	sGiftName		= requestCheckVar(Request("sGN"),64)		'검색 사은품명
	sBrand     	= requestCheckVar(Request("ebrand"),32)		'브랜드
	iItemid		= getNumeric(requestCheckVar(Request("iid"),10))		'상품코드
	sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
	sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
	sEdate     	= requestCheckVar(Request("iED"),10)		'종료일
	igStatus		= requestCheckVar(Request("giftstatus"),4)	'사은품 상태
	sgDelivery		= requestCheckVar(Request("selDelivery"),1)	'배송정보
	fcSc           = requestCheckVar(Request("fcSc"),10)       '''전체증정이벤트 강제지정
	iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
	runoutrate90up = requestCheckVar(Request("runoutrate90up"),2)
	Category	= requestCheckVar(Request("selC"),10) 		'카테고리
	CategoryMid	= requestCheckVar(Request("selCM"),10) 		'카테고리(중분류)
	DispCategory	= requestCheckVar(Request("DispCategory"),10)		'전시카테고리
	
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

'코드 유효성 검사(2008.08.04;허진원)
if sSearchTxt<>"" then
	if Not(isNumeric(sSearchTxt)) then
		if iSerachType="1" then
			Response.Write "<script language=javascript>alert('[" & sSearchTxt & "]은(는) 유효한 사은품코드가 아닙니다.');history.back();</script>"
			dbget.close()	:	response.End
		else
			Response.Write "<script language=javascript>alert('[" & sSearchTxt & "]은(는) 유효한 이벤트코드가 아닙니다.');history.back();</script>"
			dbget.close()	:	response.End
		end if
	end if
end if

''전체사은or 다이어리 이벤트 인지 Check
Dim oOpenGift, iopengiftType, iopengiftName, iopengiftfrontOpen
iopengiftType = 0
set oOpenGift=new CopenGift
oOpenGift.FRectEventCode = eCode
if (eCode<>"") then
    oOpenGift.getOneOpenGift

    if (oOpenGift.FResultcount>0) then
        iopengiftType       = oOpenGift.FOneItem.FopengiftType
        iopengiftName       = oOpenGift.FOneItem.getOpengiftTypeName
        iopengiftfrontOpen  = oOpenGift.FOneItem.FfrontOpen
    end if
end if
set oOpenGift=Nothing

strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&ebrand="&sBrand&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&giftstatus="&igStatus
set clsGift = new CGift
	clsGift.FECode = eCode
	clsGift.FSearchType = iSerachType
	clsGift.FSearchTxt  = sSearchTxt
	clsGift.FGiftName	= sGiftName
	clsGift.FBrand		= sBrand
	clsGift.FItemid		= iItemid
	clsGift.FDateType   = sDate
	clsGift.FSDate		= sSdate
	clsGift.FEDate		= sEdate
	clsGift.FGStatus	= igStatus
	clsGift.FGDelivery	= sgDelivery
	clsGift.frectrunoutrate90up	= runoutrate90up
	clsGift.frectCategory = Category
	clsGift.frectCategoryMid = CategoryMid
	clsGift.frectDispCategory = DispCategory
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

<script language="javascript">

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
	document.frmEvt.iC.value = iP;
	document.frmEvt.submit();
}

//이동
function jsGoURL(type,ival){
	if(type=="e"){
		location.href = "/admin/eventmanage/event/v5/event_register.asp?eC="+ival;
	}
}

//상품설정별 페이지이동
function jsItem(giftscope,gCode, eCode){
	//이벤트등록상품, 선택상품일떄 상품 view, 그외 페이지이동
	if(giftscope == 2 || giftscope == 4 ){
		location.href = "/admin/eventmanage/event/v5/popup/eventitem_regist.asp?eC="+eCode+"&menupos=870";
	}else if(giftscope==5){
		location.href = "giftItemReg.asp?gC="+gCode+"&menupos=<%=menupos%>";
	}
}

</script>

<!---- 검색 ---->
<% if eCode<>"" then %>
<font color="red" size="2"><b>★주의★ 사은품 관리에서 이미 사은품 등록을 했다면 이 메뉴에서 등록하지 마세요. 사은품이 중복으로 발송 됩니다.</b></font>
<% end if %>
<form name="frmEvt" method="get"  action="giftList.asp" onSubmit="return jsSearch(this,'E');" style="margin:0;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="iC">
<input type="hidden" name="fcSc" value="<%=fcSc%>">
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<select name="selType" class="select">
			<option value="1" <%IF Cstr(iSerachType) = "1" THEN%>selected<%END IF%>>사은품코드</option>
			<option value="2" <%IF Cstr(iSerachType) = "2" THEN%>selected<%END IF%>>이벤트코드</option>
		</select>
		<input type="text" class="text" name="sTxt" value="<%=sSearchTxt%>" size="10" maxlength="10">
		&nbsp;&nbsp;
		* 브랜드:
		<% drawSelectBoxDesignerwithName "ebrand", sBrand %>
		&nbsp;&nbsp;
		* 상품코드:
		<input type="text" class="text" name="iid" value="<%=iItemid%>" size="10" maxlength="10">
		&nbsp;&nbsp;
		* 사은품명:
		<input type="text" class="text" name="sGN" value="<%=sGiftName%>" maxlength="64" size="40">
	</td>
	<td  rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="document.frmEvt.submit();">
	</td>
</tr>
<tr  bgcolor="#FFFFFF">
	<td>
		* 기간:
		<select name="selDate" class="select">
			<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>시작일 기준</option>
			<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>종료일 기준</option>
		</select>
		<input type="text" class="text" size="10" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:hand;">
		~ <input type="text" class="text" size="10" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:hand;">
		&nbsp;&nbsp;
		* 상태:
		<%sbGetOptCommonCodeArr "giftstatus", igStatus, True,False,"onChange='document.frmEvt.submit();'"%>
		&nbsp;&nbsp;
		* 배송:
		<select class="select" name="selDelivery" onChange="document.frmEvt.submit();">
			<option value="">전체</option>
			<option value="Y" <%IF sgDelivery="Y" THEN%>selected<%END IF%>>업체</option>
			<option value="N" <%IF sgDelivery="N" THEN%>selected<%END IF%>>텐바이텐</option>
		</select>
		<p>
		* 관리
		<!-- #include virtual="/common/module/categoryselectbox_event.asp"-->
		&nbsp;&nbsp;
		* 전시카테고리 : <%=fnDispCateSelectBox(1,"","DispCategory",DispCategory,"") %>
		&nbsp;&nbsp;
		<input type="checkbox" name="runoutrate90up" value="ON" <% if runoutrate90up="ON" then response.write " checked" %>>소진율90%이상
	</td>
</tr>
</table>
</form>

<% if (iopengiftType<>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#AACCCC" height="25">
        <td align="left" width="200">
            전체사은이벤트 타입 : <b><%= iopengiftName %></b>
	    </td>
	    <td align="left">
            프런트 오픈 상태 : <%= iopengiftfrontOpen %>
	    </td>

	</tr>
</table>
<% end if %>

<!---- /검색 ---->
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
<tr height="40" valign="bottom">
    <td align="left">
    	<input type="button" value="새로등록" class="button" onclick="location.href='giftReg.asp?menupos=<%=menupos%>&eC=<%=eCode%>&fcSc=<%= fcSc %>';" >
    </td>
    <td align="right">
    <%
    	if (iopengiftType<>0) then
	    	'// 등록된 사은품이 있으면 사은품 지급 현황 팝업
	    	IF isArray(arrList) THEN
	    		Dim arrGCd
	    		For intLoop = 0 To UBound(arrList,2)
	    			arrGCd = arrGCd & chkIIF(arrGCd<>"",",","") & arrList(0,intLoop)
	    		Next
    %>
    	<input type="button" value="지급현황" class="button" onclick="fnPopGiftSoldSum('<%=arrGCd%>');">
    	<script type="text/javascript">
    	function fnPopGiftSoldSum(agcd) {
    		window.open("popGiftSoldSumary.asp?arr="+agcd,"popGiftSold","width=550,height=300,scrollbars=yes");
    	}
    	</script>
    <%
    		End If
    	End If
    %>
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="25">검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">사은품<br>관리코드</td>
	<td width=100>사은품구분</td>
	<td width="60">이벤트<br>코드</br>(그룹)</td>
	<td>이벤트명</td>
	<td>브랜드</td>
	<td width="130">증정대상</td>
	<td width="90">증정조건</td>
	<td width="60">이상</td>
	<td width="60">미만</td>
	<td width="30">수량</td>
	<td>사은품명</td>
	<td width="65">시작일</td>
	<td width="65">종료일</td>
	<td width="50">상태</td>
	<td width="30">한정</td>
	<td width="30">소진</td>
	<td width="50">소진율</td>
	<td width="50">배송</td>
	<td width="65">등록일</td>
</tr>
<%IF isArray(arrList) THEN
	For intLoop = 0 To UBound(arrList,2)
		tmpTitle = db2html(arrList(1,intLoop))
		if arrList(9,intLoop) > 0 then
			tmpGift = "[" & arrList(9,intLoop) & "] " & arrList(19,intLoop)
			tmpGiftForTitle = "[" & arrList(9,intLoop) & "] " & arrList(19,intLoop)
		else
			tmpGift = ""
			tmpGiftForTitle = ""
		end if

		'if (Len(tmpTitle) > 30) then
		'	tmpTitle = Left(tmpTitle, 30) & "..."
		'end if
		'if (Len(tmpGift) > 30) then
		'	tmpGift = Left(tmpGift, 30) & "..."
		'end if
%>
<% if arrList(17,intLoop) = "Y" then %>
<tr align="center" bgcolor="#FFFFFF">
<% else %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<% end if %>
	<td height="35" nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=arrList(0,intLoop)%></a></td>
	<td align="left">
		<% if arrList(26,intLoop)="I" then %>
			<% if arrList(28,intLoop)<>"" and not(isnull(arrList(28,intLoop))) then %>
				<%= BF_MakeTenBarcode(arrList(27,intLoop),arrList(28,intLoop),arrList(29,intLoop)) %>
			<% end if %>
		<% elseif arrList(26,intLoop)="B" then %>
			쿠폰:<%= arrList(30,intLoop) %>
		<% end if %>
	</td>
	<td nowrap><%IF arrList(3,intLoop) > 0 THEN%><a href="javascript:jsGoURL('e',<%=arrList(3,intLoop)%>)" title="이벤트 정보수정"><%=arrList(3,intLoop)%></a><%IF arrList(4,intLoop) > 0 THEN%><br>(<%=arrList(4,intLoop)%>)<%END IF%><%END IF%></td>
	<td align="left">&nbsp;<a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="<%= db2html(arrList(1,intLoop)) %>"><%= tmpTitle %></a></td>
	<td align="left"><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=db2html(arrList(5,intLoop))%></a></td>
	<td align="left">
		<%IF (arrList(2,intLoop) = 2 or arrList(2,intLoop) = 4 or arrList(2,intLoop) = 5) then %>
			<a href="javascript:jsItem(<%=arrList(2,intLoop)%>,<%=arrList(0,intLoop)%>,<%=arrList(3,intLoop)%>)" title="등록상품 수정">
				<%=fnGetCommCodeArrDesc(arrgiftscope,arrList(2,intLoop))%>
				<% if (arrList(20,intLoop) <> 0) then %>(<%=arrList(20,intLoop)%>)<% else %>(<font color="red">없음</font>)<% end if %>
			</a>
		<%else%>
    		<%=fnGetCommCodeArrDesc(arrgiftscope,arrList(2,intLoop))%>
		<%end if%>
		</td>
	<td align="left"><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=fnGetCommCodeArrDesc(arrgifttype,arrList(6,intLoop))%></a></td>
	<td align="right" nowrap>
		<a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=formatnumber(arrList(7,intLoop),0)%></a>&nbsp;
	</td>
	<td align="right" nowrap>
		<a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=formatnumber(arrList(8,intLoop),0)%></a>&nbsp;
	</td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=arrList(11,intLoop)%></a></td>
	<td align="left">
		<a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="<%= tmpGiftForTitle %>">
		<p style="width:100px; overflow: hidden;text-overflow: ellipsis;white-space: nowrap;"><%= tmpGift %></p>
		</a>
	</td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="<%IF arrList(22,intLoop) <> "" THEN %><%=arrList(22,intLoop)%><%END IF%>"><%=arrList(13,intLoop)%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="<%IF arrList(23,intLoop) <> "" THEN %><%=arrList(23,intLoop)%><%END IF%>"><%=arrList(14,intLoop)%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=fnGetCommCodeArrDesc(arrgiftstatus,arrList(15,intLoop))%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%IF arrList(12,intLoop) > 0 THEN%><%=arrList(12,intLoop)%><%END IF%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%IF arrList(24,intLoop) > 0 THEN%><%=arrList(24,intLoop)%><%END IF%></a></td>
	<td nowrap>
		<% if arrList(25,intLoop) <> 0 then %>
			<%= arrList(25,intLoop) %> %
		<% end if %>
	</td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%IF arrList(21,intLoop)="Y" THEN%><font color="#F08050">업체</font><%ELSEIF arrList(21,intLoop)="C" THEN%><font color="#F050F0">쿠폰</font><%ELSE%><font color="#5080F0">텐바이텐</font><%END IF%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="<%= arrList(16,intLoop) %>"><%=FormatDate(arrList(16,intLoop),"0000.00.00")%></a></td>
</tr>
<% Next
ELSE
%>
<tr>
	<td colspan="25" align="center" bgcolor="#FFFFFF">등록된 내용이 없습니다.</td>
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
		<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
	<%		else %>
		<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
	<%
			end if
		next
	%>
	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
	<% else %>[next]<% end if %>
    </td>
</tr>
</table>
<!-- 표 하단바 끝-->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->