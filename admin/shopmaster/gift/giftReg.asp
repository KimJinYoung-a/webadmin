<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  사은품 관리
' History : 2008.04.01 정윤정 생성
' 			2019.01.31 정태훈 수정 태훈 이벤트 등록시 기프트 박스 정보 같이 등록
'			2020.04.08 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/itemgiftcls.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/openGiftCls.asp"-->
<%
Dim clsGift, eCode, cEvent,cEGroup, arrGroup,intgroup,eState, sStateDesc, sTitle, dSDay, dEDay, dOpenDay, dCloseDay, sBrand, blnGroup,igType
Dim iSiteScope,sPartnerID,arrsitescope, igScope, eregdate, eFolder

if (Request("fcSc")="1") THEN igScope=1  ''전체증정 이벤트

eCode     = requestCheckVar(Request("eC"),10)
igType = 2

IF eCode <> "" THEN		'이벤트 연관 일경우
	set cEvent = new ClsEventSummary
		cEvent.FECode = eCode
		cEvent.fnGetEventConts
		sTitle 	= cEvent.FEName
		dSDay	= cEvent.FESDay
		dEDay	= cEvent.FEEDay
		sBrand	= cEvent.FBrand
		eState  = cEvent.FEState
		dOpenDay= cEvent.FEOpenDate
		dCloseDay=cEvent.FECloseDate
		sStateDesc =cEvent.FEStateDesc
		iSiteScope =cEvent.FEScope
		sPartnerID =cEvent.FPartnerID
	set cEvent = nothing
	eregdate = dSDay
	set cEGroup = new ClsEventGroup
	 	cEGroup.FECode = eCode
	  	arrGroup = cEGroup.fnGetEventItemGroup
	set cEGroup = nothing

	 blngroup = False
	 IF isArray(arrGroup) THEN blngroup = True

	 arrsitescope = fnSetCommonCodeArr("eventscope",True)
END IF

if eState < 6 then eState = 0	'이벤트 상태와 사은품 상태 매칭처리(오픈이전 상태는 모두 대기상태)

''전체사은or 다이어리 이벤트 인지 Check -----------------
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

		igScope = iopengiftType
	end if
end if
set oOpenGift=Nothing

eFolder=eCode
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
<!--

//사은품 종류 등록
function jsSetGiftKind(){
	var gift_delivery;
	var sGKN;
	var makerid;

	if (frmReg.sGN.value==""){
		alert("사은품명을 먼저 입력해 주세요.");
		frmReg.sGN.focus();
		return;
	}
	sGKN=frmReg.sGN.value

	if (frmReg.ebrand.value==""){
		alert("브랜드ID를 먼저 입력해 주세요.");
		frmReg.ebrand.focus();
		return;
	}
	makerid=frmReg.ebrand.value

	if (frmReg.selD.value==""){
		alert("배송방법을 먼저 선택해 주세요.");
		frmReg.selD.focus();
		return;
	}
	gift_delivery=frmReg.selD.value

	var winkind;
	winkind = window.open('/admin/shopmaster/gift/popgiftKindReg.asp?gift_delivery='+gift_delivery+'&makerid='+makerid+'&sGKN='+sGKN,'jsSetGiftKind','width=1280px, height=960px,scrollbars=yes,resizable=yes');
	winkind.focus();
}

//-- jsPopCal : 달력 팝업 --//
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

//사은품 등록
function jsSubmitGift(){
	var frm = document.frmReg;
	if(!frm.sGN.value){
	    <% if (igScope=1 or igScope=9) then %>
	    alert("전체 사은 이벤트인 경우 상세내용 추가등록 저장 확인요망 ");
	    <% else %>
		alert("사은품명을 입력해 주세요");
		<% end if %>
		return;
	}

	if(!frm.sSD.value ){
		alert("시작일을 입력해주세요");
		return;
	}

	if(frm.sED.value){
		if(frm.sSD.value > frm.sED.value){
			alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
			return;
		}
	}

	if(frm.giftscope.value==3){
		if(!frm.ebrand.value){
		alert("브랜드명을 선택해주세요.선택브랜드에 대해 사은품이 지급됩니다.\n\n이벤트 사은품일 경우 이벤트 수정화면에서 브랜드 수정 가능합니다.");
		return;
		}
	}

	if(frm.giftscope.value==4){
		if(!frm.selG.value){
		alert("그룹을 선택해주세요");
		return;
		}
	}

	if (frm.giftkind_linkGbn.value=="B"){
		if ((frm.giftscope.value!=1)&&(frm.giftscope.value!=9)){
			alert('현재 전체 증정 이벤트 또는 다이어리이벤트만 쿠폰 타입 사은품이 가능합니다. \n\n 일반으로 등록 후 전체(다이어리) 사은 이벤트로 변경 후 수정 요망.');
			return;
		}

		if (frm.selD.value!="C"){
			alert('사은품 구분이 쿠폰인경우, 배송타입은 쿠폰만 가능합니다.');
			return;
		}
	}else{
		if (frm.selD.value=="C"){
			alert('사은품 구분이 쿠폰이 아닌경우, 배송타입을 쿠폰으로 설정 불가합니다.');
			return;
		}
	}

	var nowDate = "<%=date()%>";

	if(frm.giftstatus.value==7) {
		if(frm.sOD.value !=""){
			nowDate = '<%IF dOpenDay <> ""THEN%><%=FormatDate(dOpenDay,"0000-00-00")%><%END IF%>';
		}

		if(frm.sSD.value < nowDate){
			//alert("시작일이 오픈일보다  빠르면 안됩니다. 시작일을 다시 선택해주세요");
			//frm.sSD.focus();
			//return;
			if(!confirm("시작일이 오픈일보다 빠르면 안됩니다!!!\n\n 정말 이대로 진행하시렵니까?")) {
				return;
			}
		}
	}

	if(!frm.sGKN.value){
		alert("사은품 종류 입력해 주세요");
		return;
	}

	if(!frm.iGK.value){
		alert("사은품 종류를 확인 버튼을 눌러서 확인해 주세요");
		return;
	}

	<% if (igScope=1 or igScope=9) then %>
		if (frm.giftscope.value!=<%=igScope%>){
			alert('전체 사은타입이 전체 또는 다이어리 증정인경우 증정대상도 동일해야 합니다.');
			return;
		}
	<% end if %>

	if (confirm("저장하시겠습니까?") == true) {
		frm.submit();
	}
}

//-- jsChkGiftType : 모든구매자 조건처리 --//
function jsChkGiftType(iVal){
	if(iVal==1){
		document.all.sGR1.readOnly=true;
		document.all.sGR2.readOnly=true;
		document.all.sGR1.style.backgroundColor='#E6E6E6';
		document.all.sGR2.style.backgroundColor='#E6E6E6';

		document.all.sGR1.value=0;
		document.all.sGR2.value=0;
	}else{
		document.all.sGR1.readOnly=false;
		document.all.sGR2.readOnly=false;
		document.all.sGR1.style.backgroundColor='';
		document.all.sGR2.style.backgroundColor='';

	}

	if(iVal == 2){
		document.all.spanKT.style.display = "none";
	}else{
		document.all.spanKT.style.display = "";
	}
	chkKTdisable();
}

function jsChkgiftgroup(iVal){

  if(iVal ==4){
	document.all.dgiftgroup.style.display = "";
  }else{
	document.all.dgiftgroup.style.display = "none";
  }

   if(iVal ==6){
	document.all.divType1.style.display = "none";
	document.all.divType2.style.display = "none";
  }else{
	document.all.divType1.style.display = "";
	document.all.divType2.style.display = "";
  }
  chkKTdisable();
}

	//-- jsImgView : 이미지 확대화면 새창으로 보여주기 --//
function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/lib/showimage.asp?img='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}

function jsChkLimit(){
	if(document.frmReg.chkLimit.checked){
		document.all.iL.readOnly=false;
		document.all.iL.style.backgroundColor='';
	}else{
		document.all.iL.readOnly=true;
		document.all.iL.style.backgroundColor='#E6E6E6';
		document.frmReg.iL.value = "";
	}
}

	//제휴몰 표기
function jsSetPartner(){
	if(document.frmReg.eventscope.options[document.frmReg.eventscope.selectedIndex].value == 3){
		$("#sSDTime").show();
		$("#sEDTime").show();
		if ($("#sSDTime").val() == ""){
			$("#sSDTime").val("00:00:00");
		}
		if ($("#sEDTime").val() == ""){
			$("#sEDTime").val("23:59:00");
		}
		document.all.spanP.style.display ="";
	}else{
		$("#sSDTime").hide();
		$("#sEDTime").hide();
		$("#sSDTime").val("");
		$("#sEDTime").val("");
		document.all.spanP.style.display ="none";
	}
}

// 사은품등록내역 가져오기
function jsImport(ec){
	var pp = window.open('/admin/shopmaster/gift/popGiftList.asp?eC='+ec,'jsImport','scrollbars=yes,resizable=yes,width=1200,height=600');
}

// 사은품등록내역 가져오기(동일 브랜드)
function jsImportSameBrand(ec) {
	var makerid = document.frmReg.ebrand.value;
	var pp = window.open("/admin/shopmaster/gift/popGiftList.asp?eC=" + ec + "&ebrand=" + makerid,'jsImportSameBrand','scrollbars=yes,resizable=yes,width=1200,height=600');
}

function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		eval("document.all."+sName).value = "";
		eval("document.all."+sSpan).style.display = "none";
	}
}

// 1+1 ,1:1 체크
function jsCheckKT(ev,ch){

	var chk 	= document.getElementById(ev);
	var chftf 	= chk.checked;
	var chk2 	= document.getElementById('tmpchkKT2');
	var chk3 	= document.getElementById('tmpchkKT3');

	chk2.checked=false;
	chk3.checked=false;

	chk.checked=chftf;
	if(chftf){
		document.frmReg.chkKT.value= chk.value;
	}else{
		document.frmReg.chkKT.value=0;
	}
}

// 1+1 disabled
function chkKTdisable(){

	if(document.all.giftscope.value==5){
		if(document.all.gifttype.value!=2){
			document.all.tmpchkKT2.disabled=false;
		} else {
			document.all.tmpchkKT2.disabled=true;
		}
	}else{
		document.all.tmpchkKT2.disabled=true;
	}
}

function TnGiftUsingNum(objval){
	if (objval == "1"){
		document.all.gifttxt1.style.display="";
		document.all.giftimg1.style.display="";
		document.all.gifttxt2.style.display="none";
		document.all.giftimg2.style.display="none";
		document.all.gifttxt3.style.display="none";
		document.all.giftimg3.style.display="none";
	}else if (objval == "2"){
		document.all.gifttxt1.style.display="";
		document.all.giftimg1.style.display="";
		document.all.gifttxt2.style.display="";
		document.all.giftimg2.style.display="";
		document.all.gifttxt3.style.display="none";
		document.all.giftimg3.style.display="none";
	}else if (objval == "3"){
		document.all.gifttxt1.style.display="";
		document.all.giftimg1.style.display="";
		document.all.gifttxt2.style.display="";
		document.all.giftimg2.style.display="";
		document.all.gifttxt3.style.display="";
		document.all.giftimg3.style.display="";
	}else{
		document.all.gifttxt1.style.display="none";
		document.all.giftimg1.style.display="none";
		document.all.gifttxt2.style.display="none";
		document.all.giftimg2.style.display="none";
		document.all.gifttxt3.style.display="none";
		document.all.giftimg3.style.display="none";
		}
}
//-->
</script>
<form name="frmReg" method="post" action="/admin/shopmaster/gift/giftProc.asp" onSubmit="return false;" style="margin:0px;">
<input type="hidden" name="sM" value="I">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="chkKT" value="0">
<input type="hidden" name="giftkind_linkGbn" value="">
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center" colspan=2>※ 이벤트정보</td>
</tr>
<tr>
	<td width="100" bgcolor="#FFFFFF" align="left" colspan=2>
		<input type="button" class="button" value="가져오기" onClick="jsImport('<%= eCode %>');">
		<input type="button" class="button" value="동일브랜드" onClick="jsImportSameBrand('<%= eCode %>');">
	</td>
</tr>

<%IF eCode <> "" THEN%>
	<tr>
		<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">이벤트코드(그룹)</td>
		<td bgcolor="#FFFFFF"><%=eCode%></td>
	</tr>
	<% if (iopengiftType<>0) then %>
		<tr>
			<td width="100" bgcolor="#AACCCC" align="center">전체사은타입</td>
			<td  bgcolor="#FFFFFF" ><%= iopengiftName %>
				<%=CHKIIF(iopengiftfrontOpen="Y","&nbsp;&nbsp;(프런트오픈)","&nbsp;&nbsp;(프런트오픈 <b>안함</b>)")%>
			</td>
		</tr>
	<% end if %>
<%END IF%>
<tr>
	<td width="100" height="30" bgcolor="<%= adminColor("tabletop") %>"  align="center"> 범위</td>
	<td bgcolor="#FFFFFF">
		<%IF eCode <> "" THEN%>
			<input type="hidden" name="eventscope" value="<%=iSiteScope%>">
			<input type="hidden" name="selP" value="<%=sPartnerID%>">
			<%=fnGetCommCodeArrDesc(arrsitescope,iSiteScope)%>&nbsp;<%=sPartnerID%>
		<%ELSE%>
			<%sbGetOptCommonCodeArr "eventscope","",False,True, "onChange=javascript:jsSetPartner();"%>
			<span id="spanP" style="display:none;">
			<select class="select" name="selP">
				<option value="">--제휴몰 전체--</option>
				<% sbOptPartner ""%>
			</select>
		<%END IF%>
	</td>
</tr>
<tr>
	<td width="100" height="30" bgcolor="<%= adminColor("tabletop") %>"  align="center"> 사은품명</td>
	<td bgcolor="#FFFFFF"><%IF eCode <> "" THEN %><%=sTitle%><input type="hidden" name="sGN" value="<%=sTitle%>"><%ELSE%><input type="text" class="text" name="sGN" size="40" maxlength="64"><%END IF%></td>
</tr>
<tr>
	<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center"> 기간</td>
	<td bgcolor="#FFFFFF">
		시작일 : <%IF eCode <> "" THEN %><%=dSDay%><input type="hidden" name="sSD" value="<%=dSDay%>"><%ELSE%><input type="text" class="text" name="sSD" size="10"   onClick="jsPopCal('sSD');"  style="cursor:hand;"><%END IF%>
		<input type="text" name="sSDTime" id="sSDTime" size="10" value="" class="text" style="display:none;">
		~ 종료일 : <%IF eCode <> "" THEN %><%=dEDay%><input type="hidden" name="sED" value="<%=dEDay%>"><%ELSE%><input type="text" class="text" name="sED"  size="10" onClick="jsPopCal('sED');" style="cursor:hand;"><%END IF%>
		<input type="text" name="sEDTime" id="sEDTime" size="10" value="" class="text" style="display:none;">
	</td>
</tr>
<tr>
	<td width="100" height="30" bgcolor="<%= adminColor("tabletop") %>"  align="center">브랜드</td>
	<td bgcolor="#FFFFFF">
		<%IF sBrand <> "" THEN %><%=sBrand%>
			<input type="hidden" name="ebrand" value="<%=sBrand%>">
		<%ELSE%>
			<% drawSelectBoxDesignerwithName "ebrand", "" %>
		<%END IF%>
	</td>
</tr>
<!-- ---------------------------------------------------------------- -->
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center" colspan=2>※ 사은품정보</td>
</tr>
<tr>
	<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">배송방법</td>
	<td bgcolor="#FFFFFF">
		<select name="selD" class="select">
			<option value="N" >텐바이텐배송</option>
			<option value="Y" >업체배송</option>
			<% if (igScope=1)or(igScope=9) then %>
				<option value="C" >쿠폰</option>
			<% end if %>
		</select>
		<span id="icpnSpan" name="icpnSpan" style="display=block">
			쿠폰번호 : <input type="text" class="text_ro" READOnly name="bcouponidx" value="0" size="9" maxlegth="9"> <!-- in Gift_kind -->
		</span>
	</td>
</tr>
<tr>
	<td width="100" height="30" bgcolor="<%= adminColor("tabletop") %>" align="center"  width="100">사은품종류</td>
	<td bgcolor="#FFFFFF">
		<input type="hidden" name="iGK" >
		<input type="text" class="text" name="sGKN" size="30" maxlength="60" onkeyup="document.frmReg.iGK.value='';">
		<input type="button" class="button" value="검색" onClick="jsSetGiftKind();">
		<div id="spanImg"></div>
	</td>
</tr>
<tr>
	<td width="100" height="30" bgcolor="<%= adminColor("tabletop") %>"  align="center">대상상품</td>
	<td bgcolor="#FFFFFF">
		<%sbGetOptGiftCodeValue "giftscope",igScope,blngroup,"onchange='jsChkgiftgroup(this.value);'",eCode%>
		<div id="dgiftgroup" style="display:none;">
			<%IF isArray(arrGroup) THEN%>
				그룹선택: 
				<select class="select" name="selG">
					<option value="">-----</option>
					<% For intgroup = 0 To UBound(arrGroup,2) %>
					<option value="<%=arrGroup(0,intgroup)%>"> <%=arrGroup(0,intgroup)%>(<%=db2html(arrGroup(1,intgroup))%>)</option>
					<% Next %>
				</select>
			<%ELSE%>
				<input type="hidden" name="selG" value="0">
			<%END IF%>
		</div>
	</td>
</tr>
<tr id="divType1" style="display:;">
	<td width="100" height="30" bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">증정조건</td>
	<td bgcolor="#FFFFFF">
		<% if (igScope=9) then %>
			<select class="select" name="gifttype" onchange='jsChkGiftType(this.value);'>
				<option value="2" selected>가격(원)</option>
			<select>
		<% else %>
			<%sbGetOptCommonCodeArr "gifttype", igType, False,True,"onchange='jsChkGiftType(this.value);'"%>
		<% end if %>
	</td>
</tr>
<tr id="divType2" style="display:;">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">증정범위</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="sGR1" class="text" size="10" style="text-align:right" value="0"> 이상 ~ <input type="text" class="text" name="sGR2" size="10" style="text-align:right" value="0"> 미만
		(ex. 20개 이상: 20~0)
	</td>
</tr>
<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center"  width="100">사은품수량</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="iGKC" size="4" maxlength="10" value="1" style="text-align:right;"> 개씩
		<% if (igScope=9) then %>
			<span id="spanKT" style="display:;">
				<label title="같은상품증정" ><input type="checkbox" name="tmpchkKT2" id="tmpchkKT2" disabled onclick="jsCheckKT('tmpchkKT2',this.cheked);" value="2">1+1(동일상품)</label>
				<label title="다른상품증정" ><input type="checkbox" name="tmpchkKT3" id="tmpchkKT3" onclick="jsCheckKT('tmpchkKT3',this.cheked);" value="3">1:1(다른상품)</label>
			</span>
		<% else %>
			<span id="spanKT" style="display:none;">
				<label title="같은상품증정" ><input type="checkbox" name="tmpchkKT2" id="tmpchkKT2" onclick="jsCheckKT('tmpchkKT2',this.cheked);" value="2">1+1(동일상품)</label>
				<label title="다른상품증정" ><input type="checkbox" name="tmpchkKT3" id="tmpchkKT3" onclick="jsCheckKT('tmpchkKT3',this.cheked);" value="3">1:1(다른상품)</label>
			</span>
		<% end if %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" height="30" align="center">사은품한정수량</td>
	<td bgcolor="#FFFFFF">
		<input type="checkbox" name="chkLimit" onClick="jsChkLimit();">한정
		<input type="text" class="text" name="iL" size="4"  style="text-align:right;background-color:#E6E6E6;" readonly> 개(한정수량 있을 경우에만 입력)
	</td>
</tr>
<tr>
	<td width="100" height="30" bgcolor="<%= adminColor("tabletop") %>"  align="center">상태</td>
	<td bgcolor="#FFFFFF">
		<%IF eCode <> "" THEN%>
			<input type="hidden" name="giftstatus" value="<%=eState%>"><%=replace(sStateDesc,"오픈예정","오픈")%>
		<%ELSE%>
			<%sbGetOptCommonCodeArr "giftstatus", "", False,True,""%>
		<%END IF%>
		<input type="hidden" name="sOD" value="">
		<input type="hidden" name="sCD" value="">
	</td>
</tr>
<!-- ---------------------------------------------------------------- -->
<% if eCode<>"" then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center" colspan=2>※ 사은품표시정보(프론트)</td>
	</tr>
	<tr>
		<td height="30" width="100" bgcolor="#FFFFFF" align="left" colspan=2><B>사은품 텍스트 박스 정보</B></td>
	</tr>
	<tr>
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">프론트 노출 여부<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<select name="gift_isusing" id="gift_isusing1" onchange="TnGiftUsingNum(this.value);">
				<option value="1">1개 사용</option>
				<option value="2">2개 사용</option>
				<option value="3">3개 사용</option>
				<option value="0">사은품박스 사용안함</option>
			</select>
			<input type="checkbox" name="gift_infotext" value="Y">한정수량 안내문구
		</td>
	</tr>
	<tr style="display:" id="gifttxt1">
		<td width="100"  bgcolor="<%= adminColor("tabletop") %>">사은품1 내용</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="gift_text1" id="gift_text1_1" size="100" maxlength="64">
		</td>
	</tr>
	<tr style="display:" id="giftimg1">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">사은품1 이미지</td>
		<td bgcolor="#FFFFFF">
			<input type="button" name="btnicon" value="이미지 등록" onClick="jsSetImg('<%=eFolder%>','','gift_img1','spangift_img1')" class="button">
			<input type="hidden" name="gift_img1">
			<div id="spangift_img1" style="padding: 5 5 5 5"></div>
		</td>
	</tr>
	<tr style="display:none" id="gifttxt2">
		<td width="100"  bgcolor="<%= adminColor("tabletop") %>">사은품2 내용</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="gift_text2" id="gift_text2_1" size="100" maxlength="64">
		</td>
	</tr>
	<tr style="display:none" id="giftimg2">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">사은품2 이미지</td>
		<td bgcolor="#FFFFFF">
			<input type="button" name="btnicon" value="이미지 등록" onClick="jsSetImg('<%=eFolder%>','','gift_img2','spangift_img2')" class="button">
			<input type="hidden" name="gift_img2">
			<div id="spangift_img2" style="padding: 5 5 5 5"></div>
		</td>
	</tr>
	<tr style="display:none" id="gifttxt3">
		<td width="100"  bgcolor="<%= adminColor("tabletop") %>">사은품3 내용</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="gift_text3" id="gift_text3_1" size="100" maxlength="64">
		</td>
	</tr>
	<tr style="display:none" id="giftimg3">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">사은품3 이미지</td>
		<td bgcolor="#FFFFFF">
			<input type="button" name="btnicon" value="이미지 등록" onClick="jsSetImg('<%=eFolder%>','','gift_img3','spangift_img3')" class="button">
			<input type="hidden" name="gift_img3">
			<div id="spangift_img3" style="padding: 5 5 5 5"></div>
		</td>
	</tr>
<% end if %>

<tr>
	<td height="30" bgcolor="#FFFFFF" align="center" colspan=2>
		<input type="button" class="button" value="저장하기" onClick="jsSubmitGift();">
		&nbsp;
		<input type="button" class="button" value="취소" onClick="history.back();">
	</td>
</tr>
</table>

</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->