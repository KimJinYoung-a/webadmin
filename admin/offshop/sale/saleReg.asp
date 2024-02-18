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
Dim sMode ,eCode, cEvent ,sTitle, dSDay, dEDay, sBrand,eState , sale_shopmargin
Dim sCode, clsSale,isRate, isMargin,sale_shopmarginvalue, isStatus, egCode, isUsing, dOpenDay,isMValue,dCloseDay
Dim intgroup , strParm , shopid , point_rate, shopname
Dim iSerachType,sSearchTxt, sDate,sSdate,sEdate,iCurrpage,sStatus
	eCode     = requestCheckVar(Request("eC"),10)
	sCode     = requestCheckVar(Request("sC"),10)
	isRate = 0
	isUsing = true
	sMode  = "I"
	isStatus =0

IF sCode <> "" THEN
	set clsSale = new CSale
	sMode = "U"
	clsSale.FSCode  = sCode
	clsSale.fnGetSaleConts

	sTitle 		= clsSale.FSName
	isRate 		= clsSale.FSRate
	point_rate = clsSale.fpoint_rate
	isMargin 	= clsSale.FSMargin
	eCode 		= clsSale.FECode
	egCode		= clsSale.FEGroupCode
	dSDay 		= clsSale.FSDate
	dEDay 		= clsSale.FEDate
	isStatus 	= clsSale.FSStatus
	isUsing     = clsSale.FSUsing
	dOpenDay	= clsSale.FOpenDate
	isMValue	= clsSale.FSMarginValue
	sale_shopmargin = clsSale.fsale_shopmargin
	sale_shopmarginvalue	= clsSale.fsale_shopmarginvalue
	dCloseDay 	= clsSale.FCloseDate
	shopid = clsSale.Fshopid
	shopname = getoffshopname(shopid)

	'-검색----------------------------------------
	 iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
	 sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'검색어
	 sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
	 sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
	 sEdate     	= requestCheckVar(Request("iED"),10)		'종료일
	 sStatus		= requestCheckVar(Request("salestatus"),4)	' 상태
	 iCurrpage		= requestCheckVar(Request("iC"),10)	'현재 페이지 번호

	 strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&sstatus="&sStatus
	'---------------------------------------------
	set clsSale = nothing
END IF

IF eCode = "0" THEN eCode = ""

IF eCode <> "" THEN		'이벤트 연관 일경우
	IF sCode = "" THEN
		set cEvent = new cevent_list
			cEvent.Frectevt_code = eCode
			cEvent.fnGetEventConts

			sTitle 	= cEvent.foneitem.fevt_name
			dSDay	= cEvent.foneitem.fevt_startdate
			dEDay	= cEvent.foneitem.fevt_enddate
			isStatus  = cEvent.foneitem.fevt_state
			dOpenDay = cEvent.foneitem.FOpenDate
			shopid = cEvent.foneitem.Fshopid
		set cEvent = nothing
	END IF
END IF

IF dSDay ="" THEN dSDay = date()
IF isStatus < 6 THEN isStatus = 0
if point_rate = "" then point_rate = "0"

'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
Dim  arrsalestatus
	arrsalestatus = fnSetCommonCodeArr_off("salestatus",False)
%>

<script language="javascript">

	//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	function jsSubmitSale(){
		var frm = document.frmReg;

		if(!frm.sSN.value){
			alert("제목을 입력해 주세요");
			return false;
		}

		if(!frm.sSD.value ){
		  	alert("시작일을 입력해주세요");
		  	frm.sSD.focus();
		  	return false;
	  	}

		if(!frm.sED.value ){
		  	alert("종료일을 입력해주세요");
		  	frm.sED.focus();
		  	return false;
	  	}

	  	if(frm.sED.value){
		  	if(frm.sSD.value > frm.sED.value){
			  	alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
			  	frm.sED.focus();
			  	return false;
		  	}
		}

		if(!frm.shopid.value){
			alert("매장을 선택해주세요");
			return false;
		}

		if(frm.shopid.value.substring(0,3) == "ith"){
			alert("등록불가 매장입니다.");
			return false;
		}

		if(typeof(frm.chkstatus)=="object"){
			if(frm.chkstatus.checked) {
				frm.salestatus.value = frm.chkstatus.value;
			}
		}

		var nowDate = "<%=date()%>";
	   if(frm.salestatus.value==7){
	 	if(frm.sOD.value !=""){
	 		nowDate = '<%IF dOpenDay <> ""THEN%><%=FormatDate(dOpenDay,"0000-00-00")%><%END IF%>';
		}

		if(frm.sSD.value < nowDate){
			alert("시작일이 오픈일보다  빠르면 안됩니다. 시작일을 다시 선택해주세요");
		  	frm.sSD.focus();
		  	return false;
		 }
	  }

	  	if(!frm.iSR.value){
			alert("할인율을 입력해 주세요");
			frm.iSR.focus();
			return false;
		}
		if (!IsDouble(frm.iSR.value)){
			alert('할인율은 숫자만 가능합니다.');
			frm.iSR.focus();
			return false;
		}

		if(confirm('저장하시겠습니까?')){
			return true;
		}else{
			return false;
		}
	}

	function jsChSetValue(iVal,itype){

		if (itype == 'salemargin'){
			if(iVal ==5){
				document.all.divM.style.display = "";
			}else{
				document.all.divM.style.display = "none";
			}
		}else if (itype == 'shopsalemargin'){
			if(iVal ==5){
				document.all.divsM.style.display = "";
			}else{
				document.all.divsM.style.display = "none";
			}
		}
	}

</script>

※ 마진구분<br>
동일마진: 판매가 대비 동일 마진율 적용<br>
업체부담: 원판매가의 마진금액만큼 할인판매가에서 차감<br>
반반부담: 할인금액의 1/2금액을 원공급가에서 차감<br>
텐바이텐부담: 원공급가를 할인판매공급가로 고정<br>
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1"  >
<form name="frmReg" method="post" action="saleProc.asp?<%=strParm%>" onSubmit="return jsSubmitSale();">
<input type="hidden" name="sM" value="<%=sMode%>">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="sC" value="<%=sCode%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="sSU" value="1">
<tr>
	<td>
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<%IF eCode <> "" THEN%>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">이벤트코드(그룹)</td>
			<td bgcolor="#FFFFFF" colspan="3"><%=eCode%>
			 <input type="hidden" name="selG" value="0">
			</td>
		</tr>
		<%END IF%>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center"> 제목</td>
			<td bgcolor="#FFFFFF"><%IF eCode <> "" THEN %><%=sTitle%><input type="hidden" name="sSN" value="<%=sTitle%>"><%ELSE%><input type="text" name="sSN" size="30" maxlength="64" value="<%=sTitle%>"><%END IF%></td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center"> 기간</td>
			<td bgcolor="#FFFFFF">
				시작일 :
				<%IF eCode <> "" THEN %>
					<input type="hidden" name="sSD" value="<%=dSDay%>">
					<%=dSDay%> ~
				<%ELSE%>
					<input type="text" name="sSD" size="10" onClick="jsPopCal('sSD');" style="cursor:hand;" value="<%=dSDay%>"> ~
				<%END IF%>
				종료일 :
				<%IF eCode <> "" THEN %>
					<input type="hidden" name="sED" value="<%=dEDay%>">
					<%=dEDay%>
				<%ELSE%>
					<input type="text" name="sED"  size="10" onClick="jsPopCal('sED');" style="cursor:hand;" value="<%=dEDay%>">
				<%END IF%>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"> 할인율</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iSR" size="3" maxlength="3" value="<%=isRate%>" style="text-align:right;">%</td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center">매장</td>
			<td bgcolor="#FFFFFF">
				<% if eCode <> "" then %>
					- 이벤트를 통한 할인의 경우 이벤트의 등록대는 해당 매장과 할인에 등록대는 해당대는 매장이 같아야 합니다<br>
				<% end if %>
				<% if sCode <> "" then %>
					<%= shopname %>(<%= shopid %>)
					<input type="hidden" name="shopid" value="<%= shopid %>">
				<% else %>
					<% drawSelectBoxOffShopdiv_off "shopid",shopid , "1,3,11" ,"","" %>
				<% end if %>

			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"> 상태</td>
			<td bgcolor="#FFFFFF" >
				<input type="hidden" name="sOD" value="<%=dOpenDay%>">
				<input type="hidden" name="salestatus" value="<%=isStatus%>">
				<%=fnGetCommCodeArrDesc_off(arrsalestatus,isStatus)%>
				<%if eCode = "" then%>
					<%IF isStatus =0 then '등록대기 %>
						<input type="checkbox" name="chkstatus" value="7">오픈요청
						<Br>※ 할인셋팅이 끝나고 반드시 <font color="red">오픈요청</font>을 체크 하셔야, 새벽에 자동 오픈처리 됩니다.
						<Br><font color="red">바로 적용</font>시에도 요픈요청을 체크 하셔야, 리스트에 <font color="red">실시간적용</font> 버튼이 활성화 됩니다.
					<%elseif isStatus = 6 or isStatus = 7 then '오픈 %>
						<input type="checkbox" name="chkstatus" value="9">종료요청
						<Br>※ 오픈상태인데 <font color="red">날짜가 지난</font>경우 종료요청 체크를 하지 않아도, 새벽에 자동 종료 됩니다.
						<br><font color="red">강제 종료</font>시에는 종료요청을 체크하셔야, 리스트에 <font color="red">실시간적용</font> 버튼이 활성화 됩니다.
					<%elseif isStatus = 8 then %>
						<div style="padding-top:5px;">종료일: <%=dCloseDay%></div>
					<%end if%>
				<%end if%>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>" align="center">매입마진</td>
			<td bgcolor="#FFFFFF" >
				<%sbGetOptCommonCodeArr_off "salemargin", isMargin, False,True," onchange='jsChSetValue(this.value,""salemargin"");'"%>
				<span id="divM" style="display:<%IF isMargin<> 5 THEN %>none<%END IF%>;">할인마진<input type="text" size="4" name="isMV" maxlength="10" value="<%=isMValue%>" style="text-align:right;">%</span>
				<br><br>
				정책상 <font color="red">매입</font>상품은 <font color="red">텐바이텐부담</font> 으로 등록하셔야 합니다.(매입가 변경 불가)
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">샵공급마진</td>
			<td bgcolor="#FFFFFF">
				<%sbGetOptCommonCodeArr_off "shopsalemargin", sale_shopmargin, False,True," onchange='jsChSetValue(this.value,""shopsalemargin"");'"%>
				<span id="divsM" style="display:<%IF isMargin<> 5 THEN %>none<%END IF%>;">할인마진<input type="text" size="4" name="sale_shopmarginvalue" maxlength="10" value="<%=sale_shopmarginvalue%>" style="text-align:right;">%</span>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">포인트적립</td>
			<td bgcolor="#FFFFFF">
				<!--사용중지
				<input type="text" size="3" name="point_rate" maxlength="3" value="<%'=point_rate%>" style="text-align:right;" readonly>%-->
				<input type="text" size="3" name="point_rate" maxlength="3" value="<%=point_rate%>" style="text-align:right;">%
				<Br>정책상 할인은 포인트적립률이 0% 입니다. 포인트 적립을 원하시면 입력하세요.
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<input type="image" src="/images/icon_save.gif">
		<a href="javascript:history.back();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
