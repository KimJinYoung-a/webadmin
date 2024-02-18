<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/event_regist.asp
' Description :  이벤트 개요 등록
' History : 2007.02.07 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
Dim eCode
Dim cEvtCont
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate, etag, eonlyten, eisblogurl,ebrand
Dim echkdisp, ecategory,esale,egift,ecoupon,ecomment,ebbs,eitemps,eapply,ebimg,etemp,emimg,ehtml,eisort,eiaddtype,edid,emid,efwd,selPartner, eDispCate
Dim enameEng, subcopyK , subcopyE
Dim eOneplusone , eFreedelivery , eBookingsell, eDiary
Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
Dim strparm , sCateMid

eCode = Request("eC")
ekind = Request("eK")

elevel = 2 '중요도 보통으로 임시 설정


	'## 검색 #############################
	sDate 		= requestCheckVar(Request("selDate"),1)  	'기간
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)

	sEvt 		= requestCheckVar(Request("selEvt"),10)  	'이벤트 코드/명 검색
	strTxt 		= requestCheckVar(Request("sEtxt"),120)

	sCategory	= requestCheckVar(Request("selC"),10) 		'카테고리
	sCateMid	= requestCheckVar(Request("selCM"),10) 		'카테고리(중분류)
	sState		= requestCheckVar(Request("eventstate"),4)	'이벤트 상태
	sKind 		= requestCheckVar(Request("eventkind"),4)	'이벤트종류
	eKind		= sKind
	edid  		= requestCheckVar(Request("selDId"),32)		'담당 디자이너
	emid  		= requestCheckVar(Request("selMId"),32)		'담당 MD

	ebrand		= requestCheckVar(Request("ebrand"),32)		'브랜드
	esale		= requestCheckVar(Request("chSale"),2) 		'세일유무
	egift		= requestCheckVar(Request("chGift"),2)		'사은품유무
	ecoupon	 	= requestCheckVar(Request("chCoupon"),2)	'쿠폰유무
	eonlyten	= requestCheckVar(Request("chOnlyTen"),2)	'Only-TenByTen유무

	eOneplusone	= requestCheckVar(Request("chOneplusone"),2)	'oneplusone
	eFreedelivery= requestCheckVar(Request("chFreedelivery"),2)	'무료배송
	eBookingsell= requestCheckVar(Request("chBookingsell"),2)	'예약판매
	eDiary= requestCheckVar(Request("chDiary"),2)	'다이어리
	edispCate	= requestCheckVar(Request("disp"),10) 		'전시 카테고리

	strparm  = "selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&selCM="&sCateMid&"&eventstate="&sState&"&eventkind="&sKind&"&selDId="&edid&"&selMId="&emid&_
				"&ebrand="&ebrand&"&chSale="&esale&"&chGift="&egift&"&chCoupon="&ecoupon&"&chOneplusone="&eOneplusone&"&chFreedelivery="&eFreedelivery&"&chBookingsell="&eBookingsell&"&disp="&edispCate&"&chOnlyTen="&eonlyten&"&chDiary="&eDiary
	'#######################################
IF eCode <> "" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'이벤트 코드
	'이벤트 내용 가져오기
	cEvtCont.fnGetEventCont
	ekind =	cEvtCont.FEKind
	eman =	cEvtCont.FEManager
	escope =	cEvtCont.FEScope
	ename =	db2html(cEvtCont.FEName)
	enameEng =	db2html(cEvtCont.FENameEng) '이벤트 영문 추가
	subcopyK =	db2html(cEvtCont.FsubcopyK) '이벤트 영문 추가
	subcopyE =	db2html(cEvtCont.FsubcopyE) '이벤트 영문 추가

	elevel =	cEvtCont.FELevel
	'estate =	cEvtCont.FEState
	eregdate =	cEvtCont.FERegdate

	'이벤트 화면설정 내용 가져오기
	cEvtCont.fnGetEventDisplay
	echkdisp 	= 	cEvtCont.FChkDisp
	ecategory 	=	cEvtCont.FECategory
	eDispCate	=	cEvtCont.FEdispCate
	esale 		= 	cEvtCont.FESale
	egift 		=	cEvtCont.FEGift
	ecoupon 	=	cEvtCont.FECoupon
	ecomment 	=	cEvtCont.FECommnet
	ebbs 		=	cEvtCont.FEBbs
	eitemps	 	=	cEvtCont.FEItemps
	eapply 		=	cEvtCont.FEApply
	eisort 		=	cEvtCont.FEISort
	edid 		=	cEvtCont.FEDId
	emid 		=	cEvtCont.FEMId
	efwd 		=	db2html(cEvtCont.FEFwd)
	etag		= db2html(cEvtCont.FETag)
 	eonlyten		= cEvtCont.FSisOnlyTen
 	eDiary		= cEvtCont.FSisDiary
 	eisblogurl		= cEvtCont.FSisGetBlogURL

	eOneplusone	=	cEvtCont.FEOneplusOne
	eFreedelivery		=	cEvtCont.FEFreedelivery
	eBookingsell		=	cEvtCont.FEBookingsell

	set cEvtCont = nothing
END IF

'2014-08-27 김진영 수정 / 상세내용 체크를 디폴트로 MD팀 요청
echkdisp = 1

%>
<script language="javascript">
<!--
//-- jsEvtSubmit : 이벤트 등록 --//
	function jsEvtSubmit(frm){
	  //브랜드할인이면 이벤트명 조합생성
	  if(frm.eventkind.value=='16') {
	  	if(!frm.ebrand.value){
		  	alert("브랜드를 선택해 주세요");
		  	frm.ebrand.focus();
		  	return false;
	  	}
	  	if(!frm.sEDN.value){
		  	alert("이벤트명을 입력해주세요");
		  	frm.sEDN.focus();
		  	return false;
	  	}
	  	if(frm.sMDc.value<=0){
		  	alert("최대 할인율을 입력해주세요");
		  	frm.sMDc.focus();
		  	return false;
	  	} else {
	  		frm.sEN.value = frm.sEDN.value + "|" + frm.sSDc.value + "|" + frm.sMDc.value;
	  		frm.sENEng.value = frm.sEDNEng.value + "|" + frm.sSDc.value + "|" + frm.sMDc.value; // 영문이벤트
	  	}
	  }

	if(!frm.eventscope.value) {
		alert("이벤트 범위를 선택해주세요");
		frm.chkEscope[0].focus();
		return false;
	}

	  if(!frm.sEN.value){
	  	alert("이벤트명을 입력해주세요");
	  	frm.sEN.focus();
	  	return false;
	  }

	  if(frm.sEN.value.length > 80){
		alert("이벤트명은 60자까지만 가능합니다.다시 입력해주세요.");
	 	frm.sEN.focus();
	  	return false;
	  }

	   if(frm.sENEng.value.length > 120){
		alert("영문이벤트명은 120자까지만 가능합니다.다시 입력해주세요.");
	 	frm.sENEng.focus();
	  	return false;
	  }

  	  if(!frm.sSD.value || !frm.sED.value ){
	  	alert("이벤트 기간을 입력해주세요");
	  	frm.sSD.focus();
	  	return false;
	  }


	  if(frm.sSD.value > frm.sED.value){
	  	alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
	  	frm.sED.focus();
	  	return false;
	  }

	   var nowDate = jsNowDate();


	  	if(frm.sSD.value < nowDate){
	  		alert("시작일이 현재일보다  빠르면 안됩니다. 시작일을 다시 선택해주세요");
	  		frm.sSD.focus();
	  		return false;
	  	}

		if(!frm.selMId.value){
			alert('담당자를 지정하세요');
			return false;
		}

	  if(!frm.eCT.value){
	  		if(GetByteLength(frm.eCT.value) > 200){
	  			alert("comment title은 200자 이내로 작성해주세요");
	  			frm.eCT.focus();
	  			return false;
	  		}
	  	}

  		if(GetByteLength(frm.eTag.value) > 250){
  			alert("Tag는 250자 이내로 작성해주세요");
  			frm.eTag.focus();
  			return false;
  		}

	}

	function jsNowDate(){
	var mydate=new Date()
		var year=mydate.getYear()
		    if (year < 1000)
		        year+=1900

		var day=mydate.getDay()
		var month=mydate.getMonth()+1
		    if (month<10)
		        month="0"+month

		var daym=mydate.getDate()
		    if (daym<10)
		        daym="0"+daym

		return year+"-"+month+"-"+ daym
	}

//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

//-- jsChangeKind : 이벤트종류(Kind)에 따른 화면 View 변경 --//
	function jsChkDisp(){
	 if(document.frmEvt.chkDisp.checked){
	  	eDetail.style.display = "";
	  }else{
	  	eDetail.style.display = "none";
	  }
	}

	function jsChkSubj(chk){
		if(chk=='16') {
			//브랜드할인일경우에는 제목 대신 할인율 범위로 표시
			eNameTr_A.style.display = "none";
			eNameTr_C.style.display = "none";
			eNameTr_B.style.display = "";
		} else {
			eNameTr_A.style.display = "";
			eNameTr_C.style.display = "";
			eNameTr_B.style.display = "none";
		}
	}

//-- jsLastEvent : 지난 이벤트 불러오기 --//
	function jsLastEvent(){
	  var winLast,eKind;
	  eKind = document.frmEvt.eventkind.options[document.frmEvt.eventkind.selectedIndex].value;
	  winLast = window.open('pop_event_lastlist.asp?menupos=<%=menupos%>&eventkind='+eKind,'pLast','width=550,height=600, scrollbars=yes')
	  winLast.focus();
	}


	//제휴몰 표기
	function jsSetPartner(){
		if(document.frmEvt.chkEscope[0].checked&&document.frmEvt.chkEscope[1].checked) {
			document.all.spanP.style.display ="";
			document.frmEvt.eventscope.value="1";
		} else if(document.frmEvt.chkEscope[0].checked) {
			document.all.spanP.style.display ="none";
			document.frmEvt.eventscope.value="2";
		} else if(document.frmEvt.chkEscope[1].checked) {
			document.all.spanP.style.display ="";
			document.frmEvt.eventscope.value="3";
		} else {
			document.all.spanP.style.display ="none";
			document.frmEvt.eventscope.value="";
		}
	}

	// 배너 링크설정 Eable
	function jsEvtLink(bln){
		var d = document.getElementById('elUrl');

		if (bln) {
			d.readOnly=true;
			d.className ="text_ro";
		}else{
			d.readOnly=false;
			d.className="text";
		}
	}
	function workerlist()
	{
		var openWorker = null;
		var worker = frmEvt.selMId.value;
		openWorker = window.open('PopWorkerList.asp?worker='+worker+'&department_id=','openWorker','width=700,height=570,scrollbars=yes');
		openWorker.focus();
	}

	function workerDel()
	{
		var frm = document.frmEvt;

		frm.selMId.value = "";
		frm.doc_workername.value = "";
	}

//-->
</script>
<form name="frmEvt" method="post"  action="event_process.asp" onSubmit="return jsEvtSubmit(this);" style="margin:0px;">
<input type="hidden" name="imod" value="I">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="strparm" value="<%=strparm%>">
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
<tr>
	<td> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 이벤트 개요 등록 </td>
</tr>
<tr>
	<td><input type="button" value="지난 이벤트 내용 불러오기" class="button" onClick="jsLastEvent();"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		   <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>종류</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventkind",ekind,False,"onChange=javascript:jsChkSubj(this.value);"%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>주체</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventmanager",eman,False,""%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>범위</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="hidden" name="eventscope" value="2">
		   			<label><input type="checkbox" name="chkEscope" checked onclick="jsSetPartner()"> 10x10</label>
		   			<label><input type="checkbox" name="chkEscope" onclick="jsSetPartner()"> 제휴몰</label>
		   			<span id="spanP" style="display:none;">
		   			<select name="selP">
		   				<option value="">--제휴몰 전체--</option>
		   				<% sbOptPartner selPartner%>
		   			</select>
		   			</span>
		   		</td>
		   	</tr>
		   	<tr id="eNameTr_A">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>이벤트명</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sEN" size="80" maxlength="120" value="<%=ename%>">
		   		</td>
		   	</tr>
			<tr id="eNameTr_C">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>영문 이벤트명</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sENEng" size="60" maxlength="60" value="<%=enameEng%>">
		   		</td>
		   	</tr>
		   	<tr id="eNameTr_B" style="display:none;">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>이벤트명<br>및 할인율</B></td>
		   		<td bgcolor="#FFFFFF">
		   			이벤트명: <input type="text" name="sEDN" size="50" maxlength="50" value=""><br>
		   			영문이벤트명: <input type="text" name="sEDNEng" size="50" maxlength="50" value=""><br>
		   			할인율: 최저 <input type="text" name="sSDc" size="4" value="0" style="text-align:right;">% ~
		   			최고 <input type="text" name="sMDc" size="4" value="0" style="text-align:right;">%<br>
		   			<font color=gray>※브랜드 스트리트에 보여질 할인율입니다. 실제로 상품에는 적용되지 않으니 상품에는 따로 할인을 적용해주세요.<br>이벤트 링크는 브랜드 스트리트로 연결되니 반드시 상세내용에 브랜드를 선택해주세요.</font>
		   		</td>
		   	</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>서브 카피</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<table width="100%" border="0" cellpadding="0" cellspacing="0">
		   			<tr>
		   				<td width="50%" style="padding-right:3px;"><textarea name="subcopyK" style="width:100%; height:80px;" onclick="if(this.value=='한글')this.value='';" onblur="if(this.value=='')this.value='한글';" value="<%=subcopyK%>"><%=chkiif(subcopyK="","한글",subcopyK)%></textarea></td>
		   				<td width="50%"><textarea name="subcopyE" style="width:100%; height:80px;" onclick="if(this.value=='영문')this.value='';" onblur="if(this.value=='')this.value='영문';" value="<%=subcopyE%>"><%=chkiif(subcopyE="","영문",subcopyE)%></textarea></td>
		   			</tr>
		   			</table>
		   		</td>
			</tr>
		   	<tr>
		   		<td rowspan="2" align="center" bgcolor="<%= adminColor("tabletop") %>"><B>기간</B></td>
		   		<td bgcolor="#FFFFFF">
		   			시작일 : <input type="text" name="sSD" size="10" onClick="jsPopCal('sSD');"  style="cursor:hand;">
		   			~ 종료일 : <input type="text" name="sED"   size="10" onClick="jsPopCal('sED');" style="cursor:hand;">
		   		</td>
		   	</tr>
		   	<tr>
		   		<td  bgcolor="#FFFFFF">
		   			당첨 발표일 : <input type="text" name="sPD" size="10" onClick="jsPopCal('sPD');" style="cursor:hand;">
		   			(당첨자가 있는 경우에만 등록)
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>상태</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptStatusCodeValue "eventstate",estate,false,""%>
		   			<%''sbGetOptStatusCodeAuth "eventstate",0,"N",""%>
		   		</td>
		   	</tr>
			<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>중요도</b></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventlevel",elevel,False,""%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>정렬번호</b></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sortNo" value="0" size="6" maxlength="5" style="text-align:right;" />
		   			(※숫자가 클수록 우선표시 됩니다. / Day&:회차)
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>내용</b></td>
		   		<td bgcolor="#FFFFFF">
		   			상세내용 추가등록 <input type="checkbox" name="chkDisp" onClick="jsChkDisp();" <%IF echkdisp= 1 THEN%>checked<%END IF%>>
		   		</td>
		   	</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<div id="eDetail" style="display:<%IF echkdisp<> 1 THEN%>none;<%END IF%>">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		   	<tr>
		   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">관리 카테고리</td>
		   		<td bgcolor="#FFFFFF">
		   		<%'DrawSelectBoxCategoryOnlyLarge "selCategory", ecategory,"" %>
		   		<!-- #include virtual="/common/module/categoryselectbox_event.asp"-->
		   		</td>
		   	</tr>
		   	<tr>
		   		<td width="100"  align="center" bgcolor="<%= adminColor("tabletop") %>">전시 카테고리</td>
		   		<td bgcolor="#FFFFFF">
		   			<%=fnDispCateSelectBox(1,"","dispcate",eDispCate,"") %>
		   		</td>
		   	</tr>
		   <tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">브랜드</td>
		   		<td bgcolor="#FFFFFF">
		   			<% drawSelectBoxDesignerwithName "ebrand", ebrand %>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">이벤트 타입</td>
		   		<td bgcolor="#FFFFFF">
		   		<input type="checkbox" name="chSale" <%IF esale = "1" THEN%>checked<%END IF%> value="1">할인
		   		<input type="checkbox" name="chGift" <%IF egift = "1" THEN%>checked<%END IF%> value="1">사은품
		   		<input type="checkbox" name="chCoupon" <%IF ecoupon = "1" THEN%>checked<%END IF%> value="1">쿠폰
		   		<input type="checkbox" name="chOnlyTen" <%IF eonlyten ="1" THEN%>checked<%END IF%> value="1">Only-TenByTen
		   		<input type="checkbox" name="chOneplusone" <%IF eOneplusone ="1" THEN%>checked<%END IF%> value="1">1+1
				<input type="checkbox" name="chFreedelivery" <%IF eFreedelivery ="1" THEN%>checked<%END IF%> value="1">무료배송
				<input type="checkbox" name="chBookingsell" <%IF eBookingsell="1" THEN%>checked<%END IF%> value="1">예약판매
				<input type="checkbox" name="chDiary" <%IF eDiary="1" THEN%>checked<%END IF%> value="1">DiaryStory
		   		</td>
			</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">이벤트 기능</td>
		   		<td bgcolor="#FFFFFF">
		   		<input type="checkbox" name="chComm" <%IF ecomment = 1 THEN%>checked<%END IF%> value="1" >코멘트
		   		<input type="checkbox" name="chBbs" <%IF ebbs = 1 THEN%>checked<%END IF%> value="1" >게시판
		   		<input type="checkbox" name="chItemps" <%IF eitemps = 1 THEN%>checked<%END IF%> value="1" >상품후기
		   		<input type="checkbox" name="isblogurl" <%IF eisblogurl THEN%>checked<%END IF%> value="1" onClick="jsChkBlogEnable()">Blog URL
		   		<!--<input type="checkbox" name="chApply" <%IF eapply = 1 THEN%>checked<%END IF%> value="1" >응모-->
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트 링크 타입</td>
		   		<td bgcolor="#FFFFFF">
		   			<label><input type="radio" name="elType" value="E" onclick="jsEvtLink(true);" checked >이벤트</label>
		   			<label><input type="radio" name="elType" value="I" onclick="jsEvtLink(false);" >직접입력</label>
		   			&nbsp;<input type="text" id="elUrl" name="elUrl" size="40" maxlength="128" value="" class="text_ro" readOnly >

		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">상품정렬방법</td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "itemsort",eisort,False,""%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">담당디자이너</td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetDesignerid "selDId",edid,""%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">담당자</td>
		   		<td bgcolor="#FFFFFF">
					<% sbGetwork "selMId",emid,"" %>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">작업전달사항</td>
		   		<td bgcolor="#FFFFFF">
		   			작업구분 <input type="text" name="sWorkTag" size="20" maxlength="16" class="text"> <font color="darkgray">(for Designer)</font>
		   			<textarea name="tFwd" rows="15" style="width:100%;font-size:12px;font-family:'Malgun Gothic',dotum;"><%=efwd%></textarea>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">Comment Title</td>
		   		<td bgcolor="#FFFFFF">
		   			(200자 이내)		   			<Br>
		   			<textarea name="eCT" rows="2" style="width:100%;"></textarea>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">Tag</td>
		   		<td bgcolor="#FFFFFF">
		   			(250자 이내)		   			<Br>
		   			<textarea name="eTag" rows="2" style="width:100%;"><%=etag%></textarea>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">연관 이벤트코드</td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="eLC" size="4" maxlength="10">
		   		</td>
		   	</tr>
		</table>
		</div>
	</td>
</tr>
<tr>
	<td width="100%" align="right">
		<input type="image" src="/images/icon_save.gif">
		<a href="index.asp?menupos=<%=menupos%>&<%=strParm%>"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
