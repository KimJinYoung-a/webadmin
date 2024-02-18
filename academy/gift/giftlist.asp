<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  사은품 관리
' History : 2010.09.27 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/gift/giftcls.asp"-->
<%
Dim eCode ,iSerachType,sSearchTxt,sGiftName,sBrand,  sDate,sSdate,sEdate,igStatus,sgDelivery
Dim clsGift, arrList, intLoop ,iTotCnt
Dim iPageSize, iCurrpage ,iDelCnt ,iStartPage, iEndPage, iTotalPage, ix,iPerCnt ,strParm
	eCode     		= requestCheckVar(Request("eC"),10)			'이벤트 코드
	iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
	sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'검색어
	sGiftName		= requestCheckVar(Request("sGN"),64)		'검색 사은품명
	sBrand     	= requestCheckVar(Request("ebrand"),32)		'브랜드
	sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
	sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
	sEdate     	= requestCheckVar(Request("iED"),10)		'종료일
	igStatus		= requestCheckVar(Request("giftstatus"),4)	'사은품 상태
	sgDelivery		= requestCheckVar(Request("selDelivery"),1)	'배송정보
 
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

strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&ebrand="&sBrand&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&giftstatus="&igStatus

set clsGift = new CGift
	clsGift.FECode = eCode
	clsGift.FSearchType = iSerachType    
	clsGift.FSearchTxt  = sSearchTxt     
	clsGift.FGiftName	= sGiftName  
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
		document.frmSearch.iC.value = iP;
		document.frmSearch.submit();
	}
	
	//이동
	function jsGoURL(type,ival){
		if(type=="e"){		
			location.href = "/academy/event/event_modi.asp?evtId="+ival;
		}
	}
	
	//상품설정별 페이지이동
	function jsItem(giftscope,gCode, eCode){
		//이벤트등록상품, 선택상품일떄 상품 view, 그외 페이지이동
		if(giftscope == 2 || giftscope == 4 ){
			location.href = "/admin/eventmanage/event/eventitem_regist.asp?eC="+eCode+"&menupos=870";
		}else if(giftscope==5){
			location.href = "giftItemReg.asp?gC="+gCode+"&menupos=<%=menupos%>";
		}
	}

	function jsGoUrl(sUrl){
		self.location.href = sUrl;
	}
		
	
</script>

<!---- 검색 ---->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSearch" method="get"  action="giftList.asp" onSubmit="return jsSearch(this,'E');">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="iC">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<select name="selType">
			<option value="1" <%IF Cstr(iSerachType) = "1" THEN%>selected<%END IF%>>사은품코드</option>
			<option value="2" <%IF Cstr(iSerachType) = "2" THEN%>selected<%END IF%>>이벤트코드</option>
		</select>
		<input type="text" name="sTxt" value="<%=sSearchTxt%>" size="10" maxlength="10">
		&nbsp;브랜드:
		<% drawSelectBoxLecturer "ebrand", sBrand %>
		&nbsp;사은품명:
		<input type="text" name="sGN" value="<%=sGiftName%>" maxlength="64" size="40">			
	</td>
	<td  rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frmSearch.submit();">
	</td>		
</tr>
<tr bgcolor="#FFFFFF">
	<td>
		&nbsp;기간:
		<select name="selDate">
			<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>시작일 기준</option>
			<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>종료일 기준</option>
		</select>		
		<input type="text" size="10" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:hand;">
		~ <input type="text" size="10" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:hand;">     
		&nbsp;상태:		
		<%sbGetOptCommonCodeArr "giftstatus", igStatus, True,False,"onChange='javascript:document.frmSearch.submit();'"%>	
		&nbsp;배송:		
		<select name="selDelivery" onChange="javascript:document.frmSearch.submit();">
			<option value="">전체</option>
			<option value="Y" <%IF sgDelivery="Y" THEN%>selected<%END IF%>>업체</option>
		<!--<option value="N" <%IF sgDelivery="N" THEN%>selected<%END IF%>>텐바이텐</option>-->
		</select>
	</td>
</tr>	
</table>
<!---- /검색 ---->

<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >	
    <tr height="40" valign="bottom">       
        <td align="left">
        	<input type="button" value="새로등록" class="button" onclick="javascript:location.href='giftreg.asp?menupos=<%=menupos%>&eC=<%=eCode%>';" >
        	<% if eCode <> "" then %><input type="button" value="이벤트목록으로" onClick="jsGoUrl('/academy/event/event_list.asp?menupos=814');" class="button"><% end if %>
	    </td>
	    <td align="right"></td>        
	</tr>	
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="16">검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>사은품코드</td>
	<td>이벤트코드</br>(그룹)</td>
	<td>사은품명</td>
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
	<td>등록일</td>
</tr>        
<%IF isArray(arrList) THEN
	For intLoop = 0 To UBound(arrList,2)
%> 
<% if arrList(17,intLoop) = "Y" then %>
<tr align="center" bgcolor="#FFFFFF">
<% else %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<% end if %>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=arrList(0,intLoop)%></a></td>
	<td nowrap><%IF arrList(3,intLoop) > 0 THEN%><a href="javascript:jsGoURL('e',<%=arrList(3,intLoop)%>)" title="이벤트 정보수정"><%=arrList(3,intLoop)%></a><%IF arrList(4,intLoop) > 0 THEN%><br>(<%=arrList(4,intLoop)%>)<%END IF%><%END IF%></td>
	<td align="left">&nbsp;<a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=db2html(arrList(1,intLoop))%></a></td>
	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=db2html(arrList(5,intLoop))%></a></td>    	
	<td> <%IF (arrList(2,intLoop) = 2 or arrList(2,intLoop) = 4 or arrList(2,intLoop) = 5) then %>
		<a href="javascript:jsItem(<%=arrList(2,intLoop)%>,<%=arrList(0,intLoop)%>,<%=arrList(3,intLoop)%>)" title="등록상품 수정"><%=fnGetCommCodeArrDesc(arrgiftscope,arrList(2,intLoop))%><br>(<%=arrList(20,intLoop)%>)</a>
		<%else%>
		<%=fnGetCommCodeArrDesc(arrgiftscope,arrList(2,intLoop))%>
		<%end if%>
		</td>
	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=fnGetCommCodeArrDesc(arrgifttype,arrList(6,intLoop))%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=formatnumber(arrList(7,intLoop),0)%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=formatnumber(arrList(8,intLoop),0)%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=arrList(11,intLoop)%></a></td>
	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%IF arrList(9,intLoop) > 0 THEN%>[<%=arrList(9,intLoop)%>]<%=arrList(19,intLoop)%><%END IF%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="<%IF arrList(22,intLoop) <> "" THEN %><%=arrList(22,intLoop)%><%END IF%>"><%=arrList(13,intLoop)%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="<%IF arrList(23,intLoop) <> "" THEN %><%=arrList(23,intLoop)%><%END IF%>"><%=arrList(14,intLoop)%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=fnGetCommCodeArrDesc(arrgiftstatus,arrList(15,intLoop))%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%IF arrList(12,intLoop) > 0 THEN%><%=arrList(12,intLoop)%><%END IF%></a></td>
		<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%IF arrList(21,intLoop)="Y" THEN%><font color="#F08050">업체</font><%ELSE%><font color="#5080F0">텐바이텐</font><%END IF%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="사은품 정보수정"><%=FormatDate(arrList(16,intLoop),"0000.00.00")%></a></td>    	
</tr>    	
<% Next
ELSE
%>
<tr>
	<td colspan="16" align="center" bgcolor="#FFFFFF">등록된 내용이 없습니다.</td>
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
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->