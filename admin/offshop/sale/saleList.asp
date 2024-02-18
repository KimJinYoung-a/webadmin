<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  할인 관리
' History : 2010.12.01 한용민 생성
'####################################################
%>
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/sale/sale_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->

<%
Dim iPageSize, iCurrpage ,iDelCnt , iTotCnt ,clsSale, arrList, intLoop , eCode , shopid
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt , strParm
Dim iSerachType,sSearchTxt,sBrand,  sDate,sSdate,sEdate,isStatus
	eCode     		= requestCheckVar(Request("eC"),10)			'이벤트 코드
	iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
	sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'검색어
	sBrand     	= requestCheckVar(Request("ebrand"),32)		'브랜드
	sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
	sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
	sEdate     	= requestCheckVar(Request("iED"),10)		'종료일
	isStatus		= requestCheckVar(Request("salestatus"),4)	'할인 상태
	arrList = ""
	shopid		= requestCheckVar(Request("shopid"),32)		'매장
	iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호

	'검색부분이 번호만 받아야된다면 숫자만 접수
 	if iSerachType="1" or iSerachType="2" then
 		sSearchTxt = getNumeric(sSearchTxt)
 	end if

	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 50
	iPerCnt = 10

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
	 	clsSale.frectshopid = 	shopid
		arrList = clsSale.fnGetSaleList	'데이터목록 가져오기

 		iTotCnt = clsSale.FTotCnt	'전체 데이터  수
 	set clsSale = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수

	Dim arrsalemargin, arrsalestatus , arrsaleshopmargin
	'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
	arrsalemargin = fnSetCommonCodeArr_off("salemargin",False)
	arrsaleshopmargin = fnSetCommonCodeArr_off("shopsalemargin",False)
	arrsalestatus= fnSetCommonCodeArr_off("salestatus",False)
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
			location.href = "/admin/offshop/event_off/event_modify.asp?evt_code="+ival;
		}else if(type=="i"){
			location.href = "saleItemReg.asp?sC="+ival+"&menupos=<%=menupos%>";
		}
	}

	//할인 바로 적용처리
 	function jsSetRealSale(sCode, chkState){
 		if(chkState !=1){
 			alert("할인중이고 현재날짜가 할인 기간중일때만 실시간 처리 가능합니다.");
 			return;
 		}

 		if(confirm("등록된 대상상품에 대해 저장된 할인율이 실서버에 저장되며,\n\n포스에서 목록 재수신을 하실경우 바로 반영 됩니다.\n\n처리하시겠습니까?")){
 			document.frmReal.sC.value = sCode;
 			document.frmReal.submit();
 		}
 	}

	//할인 전체 실시간 적용
 	function jsSetRealSaleall(){
 		if(confirm("[관리자모드] 할인 전체 실시간 적용\n처리하시겠습니까?")){
			var pop_realall = window.open('/admin/offshop/sale/saleproc.asp?menupos=<%=menupos%>&sM=realall','pop_realall','width=600,height=400,scrollbars=yes,resizable=yes');
			pop_realall.focus();
 		}
 	}

	//동일 적용 복사
	function copyshop(upfrm, onlySameMargin) {
		if (upfrm.copyshopid.value == ''){
			alert('동일 적용 매장을 선택해주세요');
			return;
		}

		if(confirm("선택하신 할인내역과 동일하게 동일 적용매장에 대한 할인 내역을 복사 생성 하시겠습니까?") == true) {
			upfrm.sC.value = '';

			if (!CheckSelected()){
					alert('선택아이템이 없습니다.\n복사할 대상 할인을 선택해 주세요.');
					return;
				}
				var frm;
				var tmp = 0;
					for (var i=0;i<document.forms.length;i++){
						frm = document.forms[i];
						if (frm.name.substr(0,9)=="frmBuyPrc") {
							if (frm.cksel.checked){
								upfrm.sC.value = upfrm.sC.value + frm.sale_code.value;
								tmp = tmp + 1;
							}
						}
					}

				if (tmp != '1'){
					alert('할인 내역은 한가지만 선택 하실수 있습니다');
					return;
				}

			if (onlySameMargin != undefined) {
				upfrm.sOnlySameMargin.value = onlySameMargin;
			}
			upfrm.sM.value = 'copyshop';
			upfrm.action='saleProc.asp';
			upfrm.submit();
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
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan=2>검색<br>조건</td>
	<td align="left">
		<select name="selType">
			<option value="1" <%IF Cstr(iSerachType) = "1" THEN%>selected<%END IF%>>할인코드</option>
			<option value="2" <%IF Cstr(iSerachType) = "2" THEN%>selected<%END IF%>>이벤트코드</option>
			<option value="3" <%IF Cstr(iSerachType) = "3" THEN%>selected<%END IF%>>할인명</option>
		</select>
		<input type="text" name="sTxt" value="<%=sSearchTxt%>" size="30" maxlength="30">
		&nbsp;&nbsp;
		* 기간:
		<select name="selDate">
		<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>시작일 기준</option>
		<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>종료일 기준</option>
		</select>
		<input type="text" size="10" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:hand;">
		~ <input type="text" size="10" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:hand;">
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan=2>
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frmSearch.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 상태:
		<% sbGetOptCommonCodeArr_off "salestatus", isStatus, True, False,"onChange='javascript:document.frmSearch.submit();'"%>
		&nbsp;&nbsp;
		* 매장 : <% drawSelectBoxOffShopdiv_off "shopid",shopid , "1,3,11" ,"","" %>
	</td>
</tr>

</form>
</table>
<!---- /검색 ---->

<Br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
<form name="copyfrm" method="post">
<input type="hidden" name="sC">
<input type="hidden" name="sM">
<input type="hidden" name="sOnlySameMargin">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr height="40" valign="bottom">
    <td align="left">
		<font color="red">[필독]</font> 하루에 한번 새벽5시에 상태값 오픈요청의 경우 오픈으로 자동변경되며, 오픈상태인데 날짜가 지난경우 자동 종료 됩니다
		<Br>매장에서 오픈이나, 종료로 <font color="red">즉시반영</font>을 원하시는경우, <font color="red">반드시 실시간적용</font> 버튼을 누르세요.
		<!--<br>&nbsp;&nbsp;현재 포스에서 <font color="red">할인중인 상품</font>이라면, 할인에 상품 등록이 불가능 합니다.-->
    </td>
    <td align="right">
    	* 동일적용매장 : <% drawSelectBoxOffShopdiv_off "copyshopid","" , "1,3" ,"","" %>
    	<input type="button" value="할인코드 할인율 동일적용" class="button" onclick="copyshop(copyfrm);">
		&nbsp;&nbsp;
		<input type="button" value="상품별 할인율 동일적용(동일마진 상품만)" class="button" onclick="copyshop(copyfrm, 'Y');">
    	&nbsp;&nbsp;
    	<input type="button" value="신규등록" class="button" onclick="javascript:location.href='saleReg.asp?menupos=<%=menupos%>&eC=<%=eCode%>';" >
    </td>
</tr>
</form>

</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="20">
		검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>할인<br>코드</td>
	<!--<td>이벤트코드</br>(그룹코드)</td>-->
	<td>할인명</td>
	<td>매장</td>
	<td>매입마진<br>매장공급마진</td>
	<td>시작일<br>종료일</td>
	<td>상품할인적용시간</td>
	<td>상태</td>
	<td>할인율</td>
	<td>적립<br>포인트</td>
	<td>
		비고
    	<% if C_ADMIN_AUTH then %>
    		<input type="button" value="전체실시간적용" class="button" onclick="jsSetRealSaleall();">
    	<% end if %>
	</td>
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
<form action="" name="frmBuyPrc<%=intLoop%>" method="get">
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFFF';>
	<td align="center" width=25>
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
	</td>
	<td width=70>
		<%=arrList(0,intLoop)%><input type="hidden" name="sale_code" value="<%= arrList(0,intLoop) %>">
	</td>
	<!--<td>-->
		<%' IF arrList(4,intLoop) > 0 THEN%>
			<!--<a href="javascript:jsGoURL('e',<%=arrList(4,intLoop)%>)" title="이벤트 정보수정">-->
			<%'= arrList(4,intLoop) %></a>
			<%' IF arrList(5,intLoop) > 0 THEN %>
				<!--<br>(<%'= arrList(5,intLoop) %>)-->
			<%' END IF %>
		<%' END IF %>
	<!--</td>-->
	<td align="left">
		<%=db2html(arrList(1,intLoop))%>
	</td>
	<td width=140>
		<%=arrList(20,intLoop)%><Br><%=arrList(17,intLoop)%>
	</td>
	<td>
		<%=fnGetCommCodeArrDesc_off(arrsalemargin,arrList(3,intLoop))%>
		<br><%=fnGetCommCodeArrDesc_off(arrsaleshopmargin,arrList(18,intLoop))%>
	</td>
	<td width=80>
		<%=arrList(6,intLoop)%><br><%=arrList(7,intLoop)%>
	</td>
	<td width=170>
		<% if arrList(15,intLoop) <> "" or not isnull(arrList(15,intLoop)) then %>
			오픈:<%=arrList(15,intLoop)%>
		<% end if %>
		<% if arrList(16,intLoop) <> "" or not isnull(arrList(16,intLoop)) then %>
			<br>종료:<%=arrList(16,intLoop)%>
		<% end if %>
	</td>
	<td width=60>
		<%
		'/오픈
		IF arrList(8,intLoop) = 6 THEN
		%>
			<font color="blue"><%=fnGetCommCodeArrDesc_off(arrsalestatus,arrList(8,intLoop))%></font>
		<%
		'/종료
		elseIF arrList(8,intLoop) = 8 THEN
		%>
			<font color="gray"><%=fnGetCommCodeArrDesc_off(arrsalestatus,arrList(8,intLoop))%></font>
		<%
		'/오픈요청 , 종료요청
		elseIF arrList(8,intLoop) = 7 or arrList(8,intLoop) = 9 THEN
		%>
			<font color="red"><%=fnGetCommCodeArrDesc_off(arrsalestatus,arrList(8,intLoop))%></font>
		<% else %>
			<%=fnGetCommCodeArrDesc_off(arrsalestatus,arrList(8,intLoop))%>
		<% end if %>
	</td>
	<td width=50>
		<%=arrList(2,intLoop)%> %
	</td>
	<td width=50>
		<%=arrList(19,intLoop)%> %
	</td>
	<td width=280>
		<input type="button" value="수정" onclick="jsMod(<%=arrList(0,intLoop)%>);" class="button">
		<input type="button" value="상품(<%=arrList(13,intLoop)%>)" class="button" onClick="javascript:jsGoURL('i',<%=arrList(0,intLoop)%>)">
		<%IF chkState = 1 THEN%>
			<input type="button" value="실시간적용" class="button" onClick="jsSetRealSale(<%=arrList(0,intLoop)%>,<%=chkState%>);">
		<%END IF%>
	</td>
</tr>
</form>
<% Next %>
<%
iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1

If (iCurrpage mod iPerCnt) = 0 Then
	iEndPage = iCurrpage
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
<tr align="center" bgcolor="#FFFFFF" >
    <td valign="bottom" align="center" colspan=20>
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
<% ELSE %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">등록된 내용이 없습니다.</td>
</tr>
<%END IF%>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
