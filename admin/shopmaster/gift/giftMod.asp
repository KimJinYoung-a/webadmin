<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  사은품 관리
' History : 2008.04.01 정윤정 생성
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
Dim clsGift, eCode, cEvent,cEGroup, arrGroup,intgroup, sTitle, dSDay, dEDay, sBrand, blnGroup, dOpenDay, dCloseDay, giftkind_givecnt
Dim tmpsSd, tmpsED,  sSDTime, sEDTime
Dim gCode,igScope,ieGroupCode, igType, igR1,igR2, igStatus, dRegdate, sAdminid, igUsing, igkCode, igkType, igkCnt,igkLimit, igkName,sgkImg
Dim iSerachType,sSearchTxt, sDate,sSdate,sEdate,iCurrpage,sgDelivery, strParm, GiftIsusing, GiftInfoText, giftkind_linkGbn, BCouponIdx
Dim sOldName, GiftText1, GiftImage1, GiftText2, GiftImage2, GiftText3, GiftImage3, iSiteScope,sPartnerID,arrsitescope, i , arrlist, eregdate
dim eFolder
	gCode	  =	requestCheckVar(Request("gC"),10)
	iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
	sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'검색어
	sBrand     	= requestCheckVar(Request("ebrand"),32)		'브랜드
	sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
	sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
	sEdate     	= requestCheckVar(Request("iED"),10)		'종료일
	igStatus		= requestCheckVar(Request("giftstatus"),4)	'사은품 상태

	iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호

	strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&ebrand="&sBrand&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&gstatus="&igStatus

IF gCode = "" THEN
	Alert_return("유입경로에 문제가 있습니다.관리자에게 문의해 주세요")
       dbget.close()	:	response.End
END IF

set clsGift = new CGift
	clsGift.FGCode = gCode
	clsGift.fnGetGiftConts

	sTitle		= clsGift.FGName
	igScope 	= clsGift.FGScope
	eCode		= clsGift.FECode
	ieGroupCode	= clsGift.FEGroupCode
	sBrand		= clsGift.FBrand
	igType		= clsGift.FGType
	igR1		= clsGift.FGRange1
	igR2 		= clsGift.FGRange2
	igkCode		= clsGift.FGKindCode
	igkType		= clsGift.FGKindType
	igkCnt		= clsGift.FGKindCnt
	igkLimit	= clsGift.FGKindlimit
	dSDay		= clsGift.FSDate
	dEDay		= clsGift.FEDate
	igStatus	= clsGift.FGStatus
	igUsing     = clsGift.FGUsing
	dRegdate	= clsGift.FRegdate
	sAdminid 	= clsGift.FAdminid
	igkName 	= clsGift.FGKindName
	sgkImg		= clsGift.FGKindImg
	sgDelivery  = clsGift.FGDelivery
	dOpenDay	= clsGift.FOpenDate
	dCloseDay	= clsGift.FCloseDate
	sOldName	= clsGift.FOldKindName
	iSiteScope	= clsGift.FSiteScope
	sPartnerID	= clsGift.FPartnerID
	BCouponIdx  = clsGift.Fbcouponidx
	giftkind_linkGbn = clsGift.Fgiftkind_linkGbn

	giftkind_givecnt = clsGift.Fgiftkind_givecnt

	If giftkind_givecnt > 0 Then ''사은품 한정제공수량
		arrlist = clsGift.fnLimitgiftCount
	End If

	eregdate = dSDay
	clsGift.FECode = eCode
	clsGift.fnGetEventGiftBox
	GiftIsusing = clsGift.FGiftIsusing
	GiftImage1 = clsGift.FGiftImage1
	GiftText1 = clsGift.FGiftText1
	GiftImage2 = clsGift.FGiftImage2
	GiftText2 = clsGift.FGiftText2
	GiftImage3 = clsGift.FGiftImage3
	GiftText3 = clsGift.FGiftText3
	GiftInfoText = clsGift.FGiftInfoText
set clsGift = nothing

IF eCode = 0 THEN eCode = ""
IF igkLimit = 0 THEN igkLimit = ""
IF isNull(igkLimit) THEN igkLimit = ""

IF eCode <> "" THEN	'이벤트와 연관된 사은품일 경우
	arrsitescope = fnSetCommonCodeArr("eventscope",True) '범위 코드값에 따른 명칭 가져오기
	'그룹리스트
	set cEGroup = new ClsEventGroup
	cEGroup.FECode = eCode
	arrGroup = cEGroup.fnGetEventItemGroup
	set cEGroup = nothing
END IF
	blngroup = False
	IF isArray(arrGroup) THEN blngroup = True

	'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
Dim  arrgiftstatus
arrgiftstatus 	= fnSetCommonCodeArr("giftstatus",False)

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

//사은품 종류 등록
function jsSetGiftKind(gift_code){
	var gift_delivery;
	var sGKN;
	var makerid;

	if (gift_code==""){
		alert("사은품관리코드가 없습니다.");
		return;
	}

	if (frmReg.sGN.value==""){
		alert("사은품명을 먼저 입력해 주세요.");
		frmReg.sGN.focus();
		return;
	}
	sGKN=frmReg.sGN.value

	makerid=frmReg.ebrand.value

	if (frmReg.selD.value==""){
		alert("배송방법을 선택해 주세요.");
		return;
	}
	gift_delivery=frmReg.selD.value

	var winkind;
	winkind = window.open('/admin/shopmaster/gift/popgiftKindReg.asp?gift_delivery='+gift_delivery+'&makerid='+makerid+'&sVM=' + document.frmReg.iGK.value + '&gift_code='+gift_code+'&sGKN='+ document.frmReg.sGKN.value,'popkind','width=1280px, height=960px, scrollbars=yes,resizable=yes');
	winkind.focus();
}

function jsGiftKindManage(){
	var winkind;
	winkind = window.open('popgiftKindManage.asp?iGK='+document.frmReg.iGK.value,'popkindMan','width=850px, height=700px, scrollbars=yes,resizable=yes');
	winkind.focus();
}

//-- jsPopCal : 달력 팝업 --//
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
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

//사은품 등록
function jsSubmitGift(){
	var frm = document.frmReg;
	if(!frm.sGN.value){
		alert("사은품명을 입력해 주세요");
		//frm.sGN.focus();
		return;
	}

	if(!frm.sSD.value || !frm.sED.value ){
		alert("기간을 입력해주세요");
	  //	frm.sSD.focus();
		return;
	}

	if(frm.sSD.value > frm.sED.value){
		alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
		//frm.sED.focus();
		return;
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

	if(!frm.sGKN.value){
		alert("사은품 종류 입력해 주세요");
		return;
	}

	if(!frm.iGK.value){
		alert("사은품 종류를 확인 버튼을 눌러서 확인해 주세요");
		return;
	}

	<% if (igScope=1) then %>
	if (frm.chkLimit.checked){
		//alert('전체 증정 조건인 경우 한정을 체크하실 수 없습니다.');
		//return;
	}
	<% end if %>

	if (frm.giftkind_linkGbn.value=="B"){
		if ((frm.giftscope.value!=1)&&(frm.giftscope.value!=9)){
			alert('현재 전체 증정 이벤트만 쿠폰 타입 사은품이 가능합니다.');
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

	if(frm.giftstatus.value==7){
		if(frm.sOD.value !=""){
			nowDate = '<%IF dOpenDay <> ""THEN%><%=FormatDate(dOpenDay,"0000-00-00")%><%END IF%>';
		}

		if(frm.sSD.value < nowDate){
			//alert("시작일이 오픈일보다  빠르면 안됩니다. 시작일을 다시 선택해주세요");
			//frm.sSD.focus();
			//return;
		}
	}

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
			document.getElementById("tmpchkKT2").checked=false;
			document.getElementById("tmpchkKT3").checked=false;
		}else{
			document.all.spanKT.style.display = "";
		}

		chkKTdisable();

}

function jsChkgiftgroup(iVal){
	// 그룹상품 보여주기
  if(iVal ==4){
	document.all.dgiftgroup.style.display = "";
  }else{
	document.all.dgiftgroup.style.display = "none";
  }

  //당첨자 대상일때 증정조건 감추기
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
	<% if (igScope=1) then %>
	//alert('전체 증정 조건인 경우 한정을 체크하실 수 없습니다.');
	//document.frmReg.chkLimit.checked = false;
	<% end if %>

	if(document.frmReg.chkLimit.checked){
		document.all.iL.readOnly=false;
		document.all.iL.style.backgroundColor='';
		document.all.givecnt.readOnly=false;
		document.all.givecnt.style.backgroundColor='';
	}else{
		document.all.iL.readOnly=true;
		document.all.iL.style.backgroundColor='#E6E6E6';
		document.frmReg.iL.value = "";
		document.all.givecnt.readOnly=true;
		document.all.givecnt.style.backgroundColor='#E6E6E6';
		document.frmReg.givecnt.value = "";
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
		document.all.spanP.style.display ="none";
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

function dpCpnSpan(comp){
	if (comp.value=="C"){
		document.getElementById("icpnSpan").style.display = "block";
	}else{
		document.getElementById("icpnSpan").style.display = "none";
	}
}

function nowcnt(){
	<% If giftkind_givecnt > 0 and IsArray(arrlist) Then %>
		document.getElementById("aaaa").style.display = "block";
	<% else %>
		alert("소진 수량 없음");
	<% end if %>
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

function popgiftdetail(gift_code){
	if (gift_code==""){
		alert("사은품관리코드가 없습니다.");
		return;
	}
	var popdisp = window.open('/admin/shopmaster/gift/giftuserdetail.asp?gift_code='+gift_code+'&menupos=<%= menupos %>','giftdetail','width=1280,height=960,scrollbars=yes,resizable=yes');
	popdisp.focus();
}

</script>

<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
<tr>
    <td align="left">
    	* 주문취소시 사은품 한정수량 복구 안함.
    </td>
    <td align="right"></td>
</tr>
</table>
<!-- 표 중간바 끝-->

<form name="frmReg" method="post" action="/admin/shopmaster/gift/giftProc.asp?<%=strParm%>" onSubmit="return false;" style="margin:0px;">
<input type="hidden" name="sM" value="U">
<input type="hidden" name="gC" value="<%=gCode%>">
<input type="hidden" name="sGD" value="<%=sgDelivery%>">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="chkKT" value="<%=igkType%>">
<input type="hidden" name="giftkind_linkGbn" value="<%=giftkind_linkGbn%>">
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center" colspan=2>※ 이벤트정보</td>
</tr>
<%IF eCode <> "" THEN%>
	<tr>
		<td height="30" width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">이벤트코드(그룹)</td>
		<td bgcolor="#FFFFFF">
			<a href="http://www.10x10.co.kr/event/eventmain.asp?eventid=<%=eCode%>" target="_blank">
			<%=eCode%> <%IF ieGroupCode >0 THEN%>(<%=ieGroupCode%>)<%END IF%></a>
		</td>
	</tr>
	<% if (iopengiftType<>0) then %>
		<tr>
			<td height="30" width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">전체사은타입</td>
			<td  bgcolor="#FFFFFF" >
				<%= iopengiftName %><%=CHKIIF(iopengiftfrontOpen="Y","&nbsp;&nbsp;(프런트오픈)","&nbsp;&nbsp;(프런트오픈 <b>안함</b>)")%>
			</td>
		</tr>
	<% end if %>
<%END IF%>
<tr>
	<td height="30" width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center"> 범위</td>
	<td bgcolor="#FFFFFF">
		<%IF eCode <> "" THEN%>
			<input type="hidden" name="eventscope" value="<%=iSiteScope%>">
			<input type="hidden" name="selP" value="<%=sPartnerID%>">
			<%=fnGetCommCodeArrDesc(arrsitescope,iSiteScope)%>&nbsp;<%=sPartnerID%>
		<%ELSE%>
			<%sbGetOptCommonCodeArr "eventscope",iSiteScope,False,True, "onChange=javascript:jsSetPartner();"%>
			<span id="spanP" style="display:<%IF iSiteScope<> 3 THEN %>none<%END IF%>;">
			<select name="selP">
				<option value="">--제휴몰 전체--</option>
				<% sbOptPartner sPartnerID%>
			</select>
		<%END IF%>
	</td>
</tr>
<tr>
	<td height="30" width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center"> 사은품명</td>
	<td bgcolor="#FFFFFF"><%IF eCode <> "" THEN %><%=sTitle%><input type="hidden" name="sGN" value="<%=sTitle%>"><%ELSE%><input type="text" class="text" name="sGN" size="40" maxlength="64" value="<%=sTitle%>"><%END IF%></td>
</tr>
<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center"> 기간</td>
	<td bgcolor="#FFFFFF">
<%
	If iSiteScope = "3" Then
		tmpsSd	= dSDay
		tmpsED	= dEDay

		dSDay = LEFT(dateconvert(dSDay), 10)
		sSDTime = RIGHT(dateconvert(tmpsSd), 8)

		dEDay = LEFT(dateconvert(dEDay), 10)
		sEDTime = RIGHT(dateconvert(tmpsED), 8)
	End If	
%>
		시작일 : <%IF eCode <> "" THEN %><%=dSDay%><input type="hidden" name="sSD" value="<%=dSDay%>"><%ELSE%><input type="text" class="text" name="sSD" size="10"   value="<%=dSDay%>"  onClick="jsPopCal('sSD');"  style="cursor:hand;"><%END IF%>
		<input type="text" name="sSDTime" id="sSDTime" size="10" value="<%= sSDTime %>" class="text" style="display:<%IF iSiteScope<> 3 THEN %>none<%END IF%>;">
		~ 종료일 : <%IF eCode <> "" THEN %><%=dEDay%><input type="hidden" name="sED" value="<%=dEDay%>"><%ELSE%><input type="text" class="text" name="sED"  size="10"  value="<%=dEDay%>" onClick="jsPopCal('sED');" style="cursor:hand;"><%END IF%>
		<input type="text" name="sEDTime" id="sEDTime" size="10" value="<%= sEDTime %>" class="text" style="display:<%IF iSiteScope<> 3 THEN %>none<%END IF%>;">
	</td>
</tr>
<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">브랜드</td>
	<td bgcolor="#FFFFFF"><% drawSelectBoxDesignerwithName "ebrand", sBrand %></td>
</tr>
<!-- ---------------------------------------------------------------- -->
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center" colspan=2>※ 사은품정보</td>
</tr>
<tr>
	<td width="100" height="30" width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">사은품관리코드</td>
	<td bgcolor="#FFFFFF"><%=gCode%></td>
</tr>
<tr>
	<td width="100" height="30" width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">배송</td>
	<td bgcolor="#FFFFFF">
		<select class="select" name="selD" onChange="dpCpnSpan(this)">
			<option value="N" <%IF sgDelivery = "N" THEN%>selected<%END IF%>>텐바이텐배송</option>
			<option value="Y" <%IF sgDelivery = "Y" THEN%>selected<%END IF%>>업체배송</option>
			<% if (igScope=1)or(igScope=9) then %>
				<option value="C" <%IF sgDelivery = "C" THEN%>selected<%END IF%>>쿠폰</option>
			<% end if %>
		</select>
		<span id="icpnSpan" name="icpnSpan" style="display=<%= chkIIF(sgDelivery="C","block","none") %>">
			쿠폰번호 : <input type="text" class="text_ro" READOnly name="bcouponidx" value="<%= BCouponIdx %>" size="9" maxlegth="9"> <!-- in Gift_kind -->
		</span>
	</td>
</tr>
<tr>
	<td width="100" height="30" bgcolor="<%= adminColor("tabletop") %>" align="center">사은품종류</td>
	<td bgcolor="#FFFFFF">
		<input type="hidden" name="orgiGK" value="<%=igkCode%>">
		<input type="hidden" name="iGK" value="<%=igkCode%>">
		<input type="text" class="text" name="sGKN" size="40" maxlength="60" value ="<%=igkName%>" onkeyup="document.frmReg.iGK.value='';"> <input type="button" class="button" value="확인" onClick="jsSetGiftKind('<%= gCode %>');">

		<% if (igScope=1)or(igScope=9) then %>
		<input type="button" class="button" value="관리" onClick="jsGiftKindManage();">
		<% end if %>

		<div id="spanImg">
		<%IF sgkImg <> "" THEN%><a href="javascript:jsImgView('<%=sgkImg%>')"><img src="<%=sgkImg%>" border="0"></a><%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td width="100" height="30" bgcolor="<%= adminColor("tabletop") %>" align="center">대상상품</td>
	<td bgcolor="#FFFFFF">
		<%sbGetOptGiftCodeValue "giftscope",igScope,blngroup,"onchange='jsChkgiftgroup(this.value);'",eCode%>
		<div id="dgiftgroup" style="display:<%IF NOT (blngroup and igScope = "4") THEN%>none<%END IF%>;">
		<%IF isArray(arrGroup) THEN%>
			그룹선택: 
				<select name="selG">
					<option value="">-----</option>
					<% For intgroup = 0 To UBound(arrGroup,2) %>
					<option value="<%=arrGroup(0,intgroup)%>" <%IF Cstr(ieGroupCode) = Cstr(arrGroup(0,intgroup)) THEN %> selected<%END IF%>> <%=arrGroup(0,intgroup)%>(<%=db2html(arrGroup(1,intgroup))%>)</option>
					<%Next %>
				</select>
			<%ELSE%>
				<input type="hidden" name="selG" value="0">
			<%END IF%>
		</div>
	</td>
</tr>
<% '<!--증정조건 이벤트당첨자일 경우 증정조건 숨긴다--> %>
<tr id="divType1" style="display:<%IF igScope=6 THEN%>none<%END IF%>;">
	<td width="100" height="30" bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">증정조건</td>
	<td bgcolor="#FFFFFF">
		<% if (igScope=9) then %>
			<select name="gifttype" onchange='jsChkGiftType(this.value);'>
				<option value="2" selected>가격(원)</option>
			<select>
		<% else %>
			<%sbGetOptCommonCodeArr "gifttype", igType, False,True,"onchange='jsChkGiftType(this.value);'"%>
		<% end if %>
	</td>
</tr>
<% '<!--증정조건 이벤트당첨자일 경우 증정조건 숨긴다--> %>
<tr id="divType2" style="display:<%IF igScope=6 THEN%>none<%END IF%>;">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">증정범위</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="sGR1" size="10" style="text-align:right;<%IF igType= "1" THEN%>background-color:#E6E6E6;" readonly<%ELSE%>"<%END IF%> value="<%=igR1%>"> 이상 ~ <input type="text" class="text" name="sGR2" size="10" style="text-align:right;<%IF igType= "1" THEN%>background-color:#E6E6E6;" readonly<%ELSE%>"<%END IF%> value="<%=igR2%>"> 미만
	</td>
</tr>
<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">사은품수량</td>
	<td bgcolor="#FFFFFF" >
		<input type="text" class="text" name="iGKC" size="4" maxlength="10" value="<%=igkCnt%>" style="text-align:right;"> 개씩
		<% if (igScope=9) then %>
			<span id="spanKT" style="display:;">
				<label title="같은상품증정" ><input type="checkbox" name="tmpchkKT2" id="tmpchkKT2" disabled onclick="jsCheckKT('tmpchkKT2',this.cheked);" value="2" <%IF igkType = 2 THEN%>checked<%END IF%>>1+1(동일상품) </label>
				<label title="다른상품증정" ><input type="checkbox" name="tmpchkKT3" id="tmpchkKT3" onclick="jsCheckKT('tmpchkKT3',this.cheked);" value="3" <%IF igkType = 3 THEN%>checked<%END IF%>>1:1(다른상품) </label>
			</span>
		<% else %>
			<span id="spanKT" style="display:<%IF igType = 2 THEN%>none<%END IF%>;">
				<label title="동일상품증정" ><input type="checkbox" name="tmpchkKT2" id="tmpchkKT2" onclick="jsCheckKT('tmpchkKT2');" <%IF igScope<>5 Then%>disabled<%End IF%> value="2" <%IF igkType = 2 THEN%>checked<%END IF%>>1+1(동일상품) </label>
				<label title="다른상품증정" ><input type="checkbox" name="tmpchkKT3" id="tmpchkKT3" onclick="jsCheckKT('tmpchkKT3');" value="3" <%IF igkType = 3 THEN%>checked<%END IF%>>1:1(다른상품) </label>
			</span>
		<% end if %>
	</td>
</tr>
<tr>
	<td height="30" bgcolor="<%= adminColor("tabletop") %>"  align="center">사은품한정수량</td>
	<td bgcolor="#FFFFFF">
		<input type="checkbox" name="chkLimit" onClick="jsChkLimit();" <%IF igkLimit <> "" THEN%>checked<%END IF%>>한정
		<input type="text" class="text" name="iL" size="5" value="<%=igkLimit%>" style="text-align:right" <%IF igkLimit ="" THEN%>style="background-color:#E6E6E6;" readonly<%END IF%>>
		<strong>-<input type="text" class="text" size="5" name="givecnt" onclick="nowcnt();" value="<%=giftkind_givecnt %>" style="text-align:right" <%IF igkLimit ="" THEN%>style="background-color:#E6E6E6;" readonly<%END IF%>>=<% if igkLimit<>"" and giftkind_givecnt<>""  then: Response.Write igkLimit-giftkind_givecnt: Else Response.Write "0": End If %></strong>
			(한정수량 있을 경우에만 입력)
		<% If giftkind_givecnt > 0 and IsArray(arrlist) Then %>
		<div id="aaaa" style="display:none;position:absolute; top:400px; left:283px;background-color:#FFF;" class="a">
			<table border="1" cellpadding="0" cellspacing="0" height="132" class="a">
				<%	Dim totcnt : totcnt = 0
						For i = 0 To UBound(arrlist,2)
				%>
				<tr align="center">
					<td width="120"><%=arrlist(0,i)%></td>
					<td width="120"><%=arrlist(1,i)%></td>
				</tr>
				<%
						totcnt = totcnt + arrlist(1,i)
					Next
				%>
				<tr align="center">
					<td>합계</td>
					<td><%=totcnt%></td>
				</tr>
				<tr align="center">
					<td colspan="2" onclick="document.getElementById('aaaa').style.display = 'none';">[닫기]</td>
				</tr>
			</table>
		</div>
		<% End If %>
		<br><br><input type="button" value="고객사은품리스트" onclick="popgiftdetail('<%= gCode %>');" class="button" >
	</td>
</tr>
<tr>
	<td height="30" bgcolor="<%= adminColor("tabletop") %>"  align="center">상태</td>
	<td bgcolor="#FFFFFF">
		<%IF eCode <> "" THEN%>
			<input type="hidden" name="giftstatus" value="<%=igStatus%>"><%=replace(fnGetCommCodeArrDesc(arrgiftstatus,igStatus),"오픈예정","오픈")%>
		<%ELSE%>
			<%sbGetOptStatusCodeValue "giftstatus", igStatus, False,""%>
		<%END IF%>
		<input type="hidden" name="sOD" value="<%=dOpenDay%>">
		<input type="hidden" name="sCD" value="<%=dCloseDay%>">
		<%IF dOpenDay <> "" THEN%><span style="padding-left:10px;">오픈처리일: <%=dOpenDay%></span><%END IF%>
		<%IF dCloseDay <> "" THEN%><br><span style="padding-left:42px;">종료처리일: <%=dCloseDay%></span><%END IF%>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>"  align="center">사용유무</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="sGU" value="Y" <%IF igUsing = "Y" THEN%>checked<%END IF%>>사용 <input type="radio" name="sGU" value="N" <%IF igUsing = "N" THEN%>checked<%END IF%>>사용안함
	</td>
</tr>
<!-- ---------------------------------------------------------------- -->
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center" colspan=2>※ 사은품표시정보(프론트)</td>
</tr>
<!--<tr>-->
	<!--<td height="30" width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">출고내역서<br>(영수증)- OLD</td>-->
	<!--<td bgcolor="#FFFFFF"><%'=db2html(sOldName)%></td>-->
<!--</tr>-->
<tr>
	<td height="30" width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">출고내역서<br>(영수증)- New</td>
	<td bgcolor="#FFFFFF">
		<% =fnComGetEventConditionStr(igkType, igScope,igType,igR1, igR2,igkName,igkCnt, igkCnt,0,0,sBrand)%>
	</td>
</tr>
<% if eCode<>"" then %>
	<tr>
		<td height="30" width="100" bgcolor="#FFFFFF" align="left" colspan=2><B>사은품 텍스트 박스 정보</B></td>
	</tr>
	<tr>
		<td width="100"  bgcolor="<%= adminColor("tabletop") %>">사용여부<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<select name="gift_isusing" id="gift_isusing1" onchange="TnGiftUsingNum(this.value);">
				<option value="1"<% If GiftIsusing=1 Then %> selected<% End If %>>1개 사용</option>
				<option value="2"<% If GiftIsusing=2 Then %> selected<% End If %>>2개 사용</option>
				<option value="3"<% If GiftIsusing=3 Then %> selected<% End If %>>3개 사용</option>
				<option value="0"<% If GiftIsusing=0 Then %> selected<% End If %>>사용 안함</option>
			</select>
			<input type="checkbox" name="gift_infotext" value="Y"<% If GiftInfoText="Y" Then %> checked<% End If %>>한정수량 안내문구
		</td>
	</tr>
	<tr style="display:" id="gifttxt1">
		<td width="100"  bgcolor="<%= adminColor("tabletop") %>">사은품1 내용</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="gift_text1" id="gift_text1_1" value="<%=GiftText1%>" size="100" maxlength="64">
		</td>
	</tr>
	<tr style="display:" id="giftimg1">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">사은품1 이미지</td>
		<td bgcolor="#FFFFFF">
			<input type="button" name="btnicon" value="이미지 등록" onClick="jsSetImg('<%=eFolder%>','<%=GiftImage1%>','gift_img1','spangift_img1')" class="button">
			<input type="hidden" name="gift_img1" value="<%=GiftImage1%>">
			<div id="spangift_img1" style="padding: 5 5 5 5">
				<%IF GiftImage1 <> "" THEN %>
				<a href="javascript:jsImgView('<%=GiftImage1%>')"><img  src="<%=GiftImage1%>" border="0"></a>
				<a href="javascript:jsDelImg('gift_img1','spangift_img1');"><img src="/images/icon_delete2.gif" border="0"></a>
				<%END IF%>
			</div>
		</td>
	</tr>
	<tr style="display:<% If GiftIsusing<2 Then %>none<% End If %>" id="gifttxt2">
		<td width="100"  bgcolor="<%= adminColor("tabletop") %>">사은품2 내용</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="gift_text2" id="gift_text2_1" value="<%=GiftText2%>" size="100" maxlength="64">
		</td>
	</tr>
	<tr style="display:<% If GiftIsusing<2 Then %>none<% End If %>" id="giftimg2">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">사은품2 이미지</td>
		<td bgcolor="#FFFFFF">
			<input type="button" name="btnicon" value="이미지 등록" onClick="jsSetImg('<%=eFolder%>','<%=GiftImage2%>','gift_img2','spangift_img2')" class="button">
			<input type="hidden" name="gift_img2" value="<%=GiftImage2%>">
			<div id="spangift_img2" style="padding: 5 5 5 5">
				<%IF GiftImage2 <> "" THEN %>
				<a href="javascript:jsImgView('<%=GiftImage2%>')"><img  src="<%=GiftImage2%>" border="0"></a>
				<a href="javascript:jsDelImg('gift_img2','spangift_img2');"><img src="/images/icon_delete2.gif" border="0"></a>
				<%END IF%>
			</div>
		</td>
	</tr>
	<tr style="display:<% If GiftIsusing<3 Then %>none<% End If %>" id="gifttxt3">
		<td width="100"  bgcolor="<%= adminColor("tabletop") %>">사은품3 내용</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="gift_text3" id="gift_text3_1" value="<%=GiftText3%>" size="100" maxlength="64">
		</td>
	</tr>
	<tr style="display:<% If GiftIsusing<3 Then %>none<% End If %>" id="giftimg3">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">사은품3 이미지</td>
		<td bgcolor="#FFFFFF">
			<input type="button" name="btnicon" value="이미지 등록" onClick="jsSetImg('<%=eFolder%>','<%=GiftImage3%>','gift_img3','spangift_img3')" class="button">
			<input type="hidden" name="gift_img3" value="<%=GiftImage3%>">
			<div id="spangift_img3" style="padding: 5 5 5 5">
				<%IF GiftImage3 <> "" THEN %>
				<a href="javascript:jsImgView('<%=GiftImage3%>')"><img  src="<%=GiftImage3%>" border="0"></a>
				<a href="javascript:jsDelImg('gift_img3','spangift_img3');"><img src="/images/icon_delete2.gif" border="0"></a>
				<%END IF%>
			</div>
		</td>
	</tr>
<%END IF%>

<tr>
	<td height="30" bgcolor="#FFFFFF" align="center" colspan=2>
		<input type="button" class="button" value="저장하기" onClick="jsSubmitGift();">
		&nbsp;
		<input type="button" class="button" value="취소" onClick="history.back();">
	</td>
</tr>
</table>
</form>

<script type='text/javascript'>

function getOnLoad(){
    alert("다이어리 임시 사은품 증정 \n\n총 결제금액 30,000원 이상 구매시 추가됨\n\n변경시 서팀 문의 요망");
}
<% if gCode="5345" then %>
	window.onload=getOnLoad;
<% end if %>

</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
