<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  할인 관리
' History : 2010.09.28 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/sale/salecls.asp"-->
<%
Dim eCode ,strParm ,iSerachType,sSearchTxt,sBrand,  sDate,sSdate,sEdate,isStatus
Dim clsSale, arrList, intLoop ,iStartPage, iEndPage, iTotalPage, ix,iPerCnt
Dim iTotCnt ,iPageSize, iCurrpage ,iDelCnt
	eCode     		= requestCheckVar(Request("eC"),10)			'이벤트 코드
	iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
	sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'검색어
	sBrand     	= requestCheckVar(Request("ebrand"),32)		'브랜드
	sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
	sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
	sEdate     	= requestCheckVar(Request("iED"),10)		'종료일
	isStatus		= requestCheckVar(Request("salestatus"),4)	'할인 상태
	arrList = ""
	iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
 
 	if iSerachType="1" or iSerachType="2" then
 		'검색부분이 번호만 받아야된다면 숫자만 접수
 		sSearchTxt = getNumeric(sSearchTxt)
 	end if
 
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	
	IF Cstr(eCode) = "0" THEN eCode = ""	
	IF (eCode <> "" AND sSearchTxt = "") THEN 
		iSerachType = 2
		sSearchTxt = eCode
	END IF
				
strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&ebrand="&sBrand&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&sstatus="&isStatus    
set clsSale = new CSale
	clsSale.FECode = eCode
	clsSale.FSearchType = iSerachType    
	clsSale.FSearchTxt  = sSearchTxt     
	clsSale.FBrand		= sBrand     	
	clsSale.FDateType   = sDate     		
	clsSale.FSDate		= sSdate     	
	clsSale.FEDate		= sEdate     			
	clsSale.FSStatus	= isStatus
 	clsSale.FCPage 		= iCurrpage
 	clsSale.FPSize 		= iPageSize
 	
	arrList = clsSale.fnGetSaleList	'데이터목록 가져오기
	iTotCnt = clsSale.FTotCnt	'전체 데이터  수
set clsSale = nothing

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수	

Dim arrsalemargin, arrsalestatus
'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
arrsalemargin = fnSetCommonCodeArr("salemargin",False)
arrsalestatus= fnSetCommonCodeArr("salestatus",False)	
%>

<script language="javascript">

	//달력
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
	
	//수정
	function jsMod(scode){
		location.href = "saleReg.asp?sC="+scode+"&menupos=<%=menupos%>&<%=strParm%>";
	}
	
	//페이징처리
		function jsGoPage(iP){
		document.frmSearch.iC.value = iP;
		document.frmSearch.submit();
	}
	
	//이동
	function jsGoURL(type,ival){
		if(type=="e"){		
			location.href = "/academy/event/event_modi.asp?evtid="+ival;
		}else if(type=="i"){
			location.href = "saleItemReg.asp?sC="+ival+"&menupos=<%=menupos%>";
		}
	}
	
	//할인 바로 적용처리
 	function jsSetRealSale(sCode, chkState){  
 		if(chkState !=1){
 			alert("할인중이고 현재날짜가 이벤트 기간중일때만 실시간 처리 가능합니다.");
 			return;
 		}
 		
 		if(confirm("등록된 대상상품에 대해 저장된 할인율이 바로 적용됩니다. 처리하시겠습니까?")){
 			document.frmReal.sC.value = sCode;
 			document.frmReal.submit();
 		}
 	}

</script>

<!---- 검색 ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">	
<form name="frmReal" method="post" action="saleItemProc.asp?<%=strParm%>">
<input type="hidden" name="sC">
<input type="hidden" name="mode" value="P">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<form name="frmSearch" method="get"  action="saleList.asp" onSubmit="return jsSearch(this,'E');">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="iC">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	<select name="selType">
	<option value="1" <%IF Cstr(iSerachType) = "1" THEN%>selected<%END IF%>>할인코드</option>
	<option value="2" <%IF Cstr(iSerachType) = "2" THEN%>selected<%END IF%>>이벤트코드</option>
	<option value="3" <%IF Cstr(iSerachType) = "3" THEN%>selected<%END IF%>>할인명</option>
	</select>
	<input type="text" name="sTxt" value="<%=sSearchTxt%>" size="10" maxlength="10">		
	&nbsp;기간:
	<select name="selDate">
	<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>시작일 기준</option>
	<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>종료일 기준</option>
	</select>		
	<input type="text" size="10" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:hand;">
	~ <input type="text" size="10" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:hand;">      
	&nbsp;상태:
	<%sbGetOptCommonCodeArr  "salestatus", isStatus, True, False,"onChange='javascript:document.frmSearch.submit();'"%>		
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frmSearch.submit();">
	</td>
</tr>	
</table>
<!---- /검색 ---->
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >	
<tr height="40" valign="bottom">       
    <td align="left">
    	<input type="button" value="새로등록" class="button" onclick="javascript:location.href='saleReg.asp?menupos=<%=menupos%>&eC=<%=eCode%>';" >
    </td>
    <td align="right"></td>        
</tr>	
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="13">검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>할인코드</td>
	<td>이벤트코드</br>(그룹코드)</td>
	<td>브랜드명</td>
	<td>할인명</td>    	    	
	<td>할인율</td>
	<td>마진구분</td>
	<td>시작일</td>
	<td>종료일</td>
	<td>상태</td>    	
	<td>상품할인<br>적용시간</td>
	<td colspan="2">처리</td>
	<td>등록일</td>
</tr>        
<% Dim chkState  
IF isArray(arrList) THEN 
	For intLoop = 0 To UBound(arrList,2)  
	chkState = 0  	
	'상태: 오픈, 종료요청 )기간: 현재일기준 기간내
	if (arrList(8,intLoop) = 6 or arrList(8,intLoop) = 7 or arrList(8,intLoop) = 9) and datediff("d",arrList(6,intLoop),date()) >=0 and datediff("d",arrList(7,intLoop),date()) <=0 then
		chkState = 1    	
	end if	
%> 
<tr align="center" bgcolor="#FFFFFF">    
	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="할인 정보수정"><%=arrList(0,intLoop)%></a></td>
	<td><%IF arrList(4,intLoop) > 0 THEN%><a href="javascript:jsGoURL('e',<%=arrList(4,intLoop)%>)" title="이벤트 정보수정"><%=arrList(4,intLoop)%></a><%IF arrList(5,intLoop) > 0 THEN%><br>(<%=arrList(5,intLoop)%>)<%END IF%><%END IF%></td>
	<td><%=arrList(17,intLoop)%></td>
	<td align="left">&nbsp;<a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="할인 정보수정"><%=db2html(arrList(1,intLoop))%></a></td>    	
	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="할인 정보수정"><%=arrList(2,intLoop)%>%</a></td>    
	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="할인 정보수정"><%=fnGetCommCodeArrDesc(arrsalemargin,arrList(3,intLoop))%></a></td>    
	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="할인 정보수정"><%=arrList(6,intLoop)%></a></td>
	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="할인 정보수정"><%=arrList(7,intLoop)%></a></td>
	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="할인 정보수정"><%IF arrList(8,intLoop) = 6 THEN%><font color="blue"><%END IF%><%=fnGetCommCodeArrDesc(arrsalestatus,arrList(8,intLoop))%></a></td>
	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="할인 정보수정">
		<%IF arrList(8,intLoop) = 6 THEN %><%=arrList(15,intLoop)%>
		<%ELSEIF arrList(8,intLoop) = 8 THEN%><%=arrList(16,intLoop)%>
		<%END IF%></a>
	</td>
	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="할인 정보수정">
			<%IF chkState = 1 THEN%><input type="button" value="실시간적용" class="button" onClick="jsSetRealSale(<%=arrList(0,intLoop)%>,<%=chkState%>);"></a><%END IF%>
	</td>    			
	<td>
			<input type="button" value="상품(<%=arrList(13,intLoop)%>)" class="button" onClick="javascript:jsGoURL('i',<%=arrList(0,intLoop)%>)">    		
		</td>
	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="할인 정보수정"><%=FormatDate(arrList(10,intLoop),"0000.00.00")%></a></td>
</tr>
	
<% Next
ELSE
%>
<tr>
	<td colspan="12" align="center" bgcolor="#FFFFFF">등록된 내용이 없습니다.</td>
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