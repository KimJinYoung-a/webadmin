<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event_modify.asp
' Description :  �̺�Ʈ ���� ���
' History : 2007.02.13 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
'--------------------------------------------------------
' ��������
'--------------------------------------------------------
Dim eCode
Dim cEvtCont
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate, echkdisp, eusing, etag, eonlyten, eisblogurl
Dim ecategory,esale,egift,ecoupon,ecomment,ebbs,eitemps,eapply,elktype,elkurl,ebimg,etemp,emimg,ehtml,ehtml5, eisort,eiaddtype, edid, emid ,efwd,ebrand,eicon,ebimg2010
Dim selPartner,dopendate,dclosedate, sWorkTag, ebimgMo, eDispCate, eDateView , ebimgToday , ebimgMo2014 , enamesub
Dim intI
Dim arrGift, intg,blngift
Dim eFolder, backUrl
dim gimg : gimg = ""
Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
Dim ecommenttitle, elinkcode
Dim strparm , sCateMid
Dim cEGroup, arrGroup,intgroup,strG, blngroup
Dim blnFull, blnIteminfo ,blnitemprice, evt_sortNo, blnWide
Dim enameEng , subcopyK , subcopyE
Dim eOneplusone , eFreedelivery , eBookingsell, eDiary
Dim eEtcitemid , eEtcitemimg, eItemListType
Dim eitemid
	eCode		= requestCheckVar(Request("eC"),10)	'�̺�Ʈ�ڵ�
	blnFull		= False
	blnWide		= False
	blnIteminfo	= True
	blnitemprice = False
	'## �˻� #############################
	sDate 		= requestCheckVar(Request("selDate"),1)  	'�Ⱓ
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)

	sEvt 		= requestCheckVar(Request("selEvt"),10)  	'�̺�Ʈ �ڵ�/�� �˻�
	strTxt 		= requestCheckVar(Request("sEtxt"),120)

	sCategory	= requestCheckVar(Request("selC"),10) 		'ī�װ�
	sCateMid	= requestCheckVar(Request("selCM"),10) 		'ī�װ�(�ߺз�)
	sState		= requestCheckVar(Request("eventstate"),4)	'�̺�Ʈ ����
	sKind 		= requestCheckVar(Request("eventkind"),4)	'�̺�Ʈ����
	edid  		= requestCheckVar(Request("selDId"),32)		'��� �����̳�
	emid  		= requestCheckVar(Request("selMId"),32)		'��� MD

	ebrand		= requestCheckVar(Request("ebrand"),32)		'�귣��
	esale		= requestCheckVar(Request("chSale"),2) 		'��������
	egift		= requestCheckVar(Request("chGift"),2)		'����ǰ����
	ecoupon	 	= requestCheckVar(Request("chCoupon"),2)	'��������
	eonlyten	= requestCheckVar(Request("chOnlyTen"),2)	'Only-TenByTen����

	eOneplusone	 	= requestCheckVar(Request("chOneplusone"),2)	'oneplusone
	eFreedelivery	= requestCheckVar(Request("chFreedelivery"),2)	'������
	eBookingsell	= requestCheckVar(Request("chBookingsell"),2)	'�����Ǹ�
	eDiary	= requestCheckVar(Request("chDiary"),2)	'���̾
	edispCate	= requestCheckVar(Request("disp"),10) 		'���� ī�װ�

	strparm  = "selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&selCM="&sCateMid&"&eventstate="&sState&"&eventkind="&sKind&"&selDId="&edid&"&selMId="&emid&_
				"&ebrand="&ebrand&"&chSale="&esale&"&chGift="&egift&"&chCoupon="&ecoupon&"&chOneplusone="&eOneplusone&"&chFreedelivery="&eFreedelivery&"&chBookingsell="&eBookingsell&"&disp="&edispCate&"&chOnlyTen="&eonlyten&"&chDiary="&eDiary
	'#######################################

	IF eCode = "" THEN	'�̺�Ʈ �ڵ尪�� ���� ��� back
		call sbAlertMsg("���ް��� ������ �߻��Ͽ����ϴ�. �����ڿ��� �������ֽʽÿ�", "back","")
	END IF

	eFolder = eCode
'--------------------------------------------------------
' �̺�Ʈ ������ ��������
'--------------------------------------------------------
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'�̺�Ʈ �ڵ�
	'�̺�Ʈ ���� ��������
	cEvtCont.fnGetEventCont
	ekind 		=	cEvtCont.FEKind
	eman 		=	cEvtCont.FEManager
	escope 		=	cEvtCont.FEScope
	selPartner	=	cEvtCont.FEPartnerID
	ename 		=	db2html(cEvtCont.FEName)
	enamesub	=	db2html(cEvtCont.FENamesub) '�̺�Ʈ Ÿ��Ʋ ����ī��
	enameEng =	db2html(cEvtCont.FENameEng) '�̺�Ʈ ���� �߰�
	subcopyK =	db2html(cEvtCont.FsubcopyK) '����ī�� �ѱ�
	subcopyE =	db2html(cEvtCont.FsubcopyE) '����ī�� ����
	esday 		=	cEvtCont.FESDay
	eeday 		=	cEvtCont.FEEDay
	epday 		=	cEvtCont.FEPDay
	elevel 		=	cEvtCont.FELevel
	estate 		=	cEvtCont.FEState
	IF datediff("d",now,eeday) <0 THEN estate = 9 '�Ⱓ �ʰ��� ����ǥ��
	eregdate	=	cEvtCont.FERegdate
	eusing		= 	cEvtCont.FEUsing
	evt_sortNo	= 	cEvtCont.FESortNo
	eitemid		=	cEvtCont.FEitemid
	'�̺�Ʈ ȭ�鼳�� ���� ��������
	cEvtCont.fnGetEventDisplay
	echkdisp 		= cEvtCont.FChkDisp
	tmp_cdl 		= cEvtCont.FECategory
	tmp_cdm			= cEvtCont.FECateMid
	eDispCate		= cEvtCont.FEDispCate

	esale 			= cEvtCont.FESale
	egift 			= cEvtCont.FEGift
	ecoupon 		= cEvtCont.FECoupon
	ecomment 		= cEvtCont.FECommnet
	ebbs 			= cEvtCont.FEBbs
	eitemps 		= cEvtCont.FEItemps
	eapply 			= cEvtCont.FEApply
	elktype			= cEvtCont.FELinkType
	IF elktype="" Then elktype="E" '//��ũŸ�� �⺻�� ����
	elkurl			= cEvtCont.FELinkURL
	ebimg 			= cEvtCont.FEBImg
	ebimg2010		= cEvtCont.FEBImg2010
	ebimgMo			= cEvtCont.FEBImgMobile
	ebimgToday		= cEvtCont.FEBImgMoToday
	ebimgMo2014		= cEvtCont.FEBImgMoListBanner '//2014 ����� ����Ʈ ��� �߰�
	gimg			= cEvtCont.FEGImg
	etemp			= cEvtCont.FETemp
	if etemp = 5 or etemp = 6  THEN	'���۾� �̺�Ʈ �� ��� ó��
		ehtml5 		= db2html(cEvtCont.FEHtml)
	else
		emimg 		= cEvtCont.FEMImg
		ehtml 		= db2html(cEvtCont.FEHtml)
	end if
	eisort 			= cEvtCont.FEISort
	edid 			= cEvtCont.FEDId
	emid 			= cEvtCont.FEMId
	efwd 			= db2html(cEvtCont.FEFwd)
	ebrand			= cEvtCont.FEBrand
	eicon   		= cEvtCont.FEIcon
	ecommenttitle   = db2html(cEvtCont.FECommentTitle)
	elinkcode   	= cEvtCont.FELinkCode
	dopendate		= cEvtCont.FEOpenDate
	dclosedate		= cEvtCont.FECloseDate
 	blnFull			= cEvtCont.FEFullYN
 	blnWide			= cEvtCont.FEWideYN
 	blnIteminfo		= cEvtCont.FEIteminfoYN
 	etag			= db2html(cEvtCont.FETag)
 	eonlyten		= cEvtCont.FSisOnlyTen
 	eisblogurl		= cEvtCont.FSisGetBlogURL
 	sWorkTag		= cEvtCont.FWorkTag

	blnitemprice	 = cEvtCont.FEItempriceYN

	eOneplusone	=	cEvtCont.FEOneplusOne
	eFreedelivery		=	cEvtCont.FEFreedelivery
	eBookingsell		=	cEvtCont.FEBookingsell
	eDiary = cEvtCont.FSisDiary

	eEtcitemid			=	cEvtCont.FEtcitemid
	eEtcitemimg		=	cEvtCont.FEtcitemimg
	eDateView		= cEvtCont.FEdateview
	eItemListType = cEvtCont.FEListType

	set cEvtCont = nothing
	IF elinkcode = 0 THEN elinkcode = ""

	 set cEGroup = new ClsEventGroup
	 	cEGroup.FECode = eCode
	  	arrGroup = cEGroup.fnGetEventItemGroup
	 set cEGroup = nothing

	 blngroup = False
	 IF isArray(arrGroup) THEN blngroup = True

	 If eItemListType = "" OR isNull(eItemListType) Then eItemListType = "1" End If
%>
<script type="text/javaScript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javaScript" src="/js/jquery.iframe-auto-height.js"></script>
<script type="text/javascript">
<!--
//-- jsEvtSubmit(form ��) : �̺�Ʈ ����ó�� --//
	function jsEvtSubmit(frm){

	  //�귣�������̸� �̺�Ʈ�� ���ջ���
	  if(frm.eventkind.value=='16') {
	  	if(!frm.ebrand.value){
		  	alert("�귣�带 ������ �ּ���");
		  	frm.ebrand.focus();
		  	return false;
	  	}
	  	if(!frm.sEDN.value){
		  	alert("�̺�Ʈ���� �Է����ּ���");
		  	frm.sEDN.focus();
		  	return false;
	  	}
	  	if(frm.sMDc.value<=0){
		  	alert("�ִ� �������� �Է����ּ���");
		  	frm.sMDc.focus();
		  	return false;
	  	} else {
	  		frm.sEN.value = frm.sEDN.value + "|" + frm.sSDc.value + "|" + frm.sMDc.value;
	  		frm.sENEng.value = frm.sEDNEng.value + "|" + frm.sSDc.value + "|" + frm.sMDc.value; // �����̺�Ʈ
	  	}
	  }

	if(!frm.eventscope.value) {
		alert("�̺�Ʈ ������ �������ּ���");
		frm.chkEscope[0].focus();
		return false;
	}

	  if(!frm.sEN.value){
	  	alert("�̺�Ʈ���� �Է����ּ���");
	  	if(frm.eventkind.options[frm.eventkind.selectedIndex].value == 4){
	  	 frm.selStatic.focus();
	  	}else{
	  	 frm.sEN.focus();
	  	}
	  	return false;
	  }

	  if(frm.sENEng.value.length > 120){
		alert("�����̺�Ʈ���� 120�ڱ����� �����մϴ�.�ٽ� �Է����ּ���.");
	 	frm.sENEng.focus();
	  	return false;
	  }

	if (frm.selC.value == '110'){
		if (frm.selCM.value==''){
			alert('����ä���� ��ī�װ��� �����ؾ߸� �մϴ�');
			frm.selCM.focus();
			return false;
		}

	}

  	  if(!frm.sSD.value || !frm.sED.value ){
	  	alert("�̺�Ʈ �Ⱓ�� �Է����ּ���");
	  	frm.sSD.focus();
	  	return false;
	  }

	  if(frm.sSD.value > frm.sED.value){
	  	alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
	  	frm.sED.focus();
	  	return false;
	  }

	var nowDate = jsNowDate();

	 if(frm.eventstate.value==7){
	 	if(frm.eOD.value !=""){
	 		nowDate = '<%IF dopendate <> ""THEN%><%=FormatDate(dopendate,"0000-00-00")%><%END IF%>';
		}
 	 }


	 if(frm.eventstate.value < 7){
	 	if(frm.sSD.value < nowDate){
			alert("�������� �����Ϻ���  ������ �ȵ˴ϴ�. �������� �ٽ� �������ּ���");
		  	frm.sSD.focus();
		  	return false;
		 }

  	 	if(frm.sED.value < jsNowDate()){
	  		alert("�������� ���糯¥���� ������ �ȵ˴ϴ�. �������� �ٽ� �������ּ��� ");
	  		frm.sED.focus();
	  		return false;
	  	}
	}

	    if(!frm.eCT.value){
	  		if(GetByteLength(frm.eCT.value) > 200){
	  			alert("comment title�� 200�� �̳��� �ۼ����ּ���");
	  			frm.eCT.focus();
	  			return false;
	  		}
	  	}


  		if(GetByteLength(frm.eTag.value) > 250){
  			alert("Tag�� 250�� �̳��� �ۼ����ּ���");
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

//-- jsPopCal : �޷� �˾� --//
	function jsPopCal(sName){
		var winCal;
		var blnSale, blnGift, blnCoupon;
		blnSale= "<%=esale%>";
		blnGift= "<%=egift%>";
		blnCoupon= "<%=ecoupon%>";

		if (sName!="sPD" && (blnSale=="True" || blnGift=="True"|| blnCoupon=="True")){
			if(confirm("�Ⱓ�� ����� ����, ����ǰ, �������� ������ �˴ϴ�. �Ⱓ�� �����Ͻðڽ��ϱ�?")){
				winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
				winCal.focus();
			}
		}else{
				winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
				winCal.focus();
		}
	}

//-- jsImgDel(�̹��� ����) : �̹��� ȭ�鿡�� �Ⱥ��̰� --//
	function jsImgDel(sType){
	 if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	 	if(sType == "B"){
	 		document.frmEvt.oldBimg.value="";
	 		document.all.imgB.style.display="none";
	 	}else{
	 		document.frmEvt.oldMimg.value="";
	 		document.all.imgM.style.display="none";
	 	}
	 }
	}


//-- jsChangeFrm : ���ø��� ���� ȭ�� ���� ����:���� ���۾� --//
	function jsChangeFrm(iVal){
		$("div[id^='divFrm']").hide();

		if(iVal == 3 || iVal == 7){
			iframG.location.href = "iframe_eventitem_group.asp?eC=<%=eCode%>&ekind=<%=ekind%>";
			$("#divFrm3").show();
			$('#iframG').load(function() {
				$(this).height($(this).contents().find('body')[0].scrollHeight+30);
			});
		}else if(iVal == 5 || iVal == 6 ){
			//iframG.location.href = "about;blank";
			$("#divFrm5").show();
		}else{
			//iframG.location.href = "about;blank";
			$("#divFrm1").show();
		}
	}


//-- jsImgView : �̹��� Ȯ��ȭ�� ��â���� �����ֱ� --//
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}


	function jsSetImg(sFolder, sImg, sName, sSpan){
		document.domain ="10x10.co.kr";
		var winImg;
		winImg = window.open('/admin/eventmanage/common/pop_event_uploadimg.asp?yr=<%=Year(eregdate)%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}

	//-- jsChkDisp : ��ȭ�� ���� --//
	function jsChkDisp(){
	 if(document.frmEvt.chkDisp.checked){
	  	eDetail.style.display = "";
	  }else{
	  	eDetail.style.display = "none";
	  }
	}

	function jsChkSubj(chk){
		if(chk=='16') {
			//�귣�������ϰ�쿡�� ���� ��� ������ ������ ǥ��
			eNameTr_A.style.display = "none";
			eNameTr_C.style.display = "none";
			eNameTr_B.style.display = "";
		} else {
			eNameTr_A.style.display = "";
			eNameTr_C.style.display = "";
			eNameTr_B.style.display = "none";
		}
	}

	function jsManageEventImage(evtcode){
	    var popwin = window.open('<%= uploadImgUrl %>/linkweb/event_admin/eventManageDir.asp?evtcode=' + evtcode,'eventManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
	    popwin.focus();
	}

	function jsManageEventImageNew(evtcode){
	    var popwin = window.open('<%= uploadImgUrl %>/linkweb/event_admin/eventManageDir_new.asp?evtcode=' + evtcode,'eventManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
	    popwin.focus();
	}

	function popCommentXLS(ecd) {
		 var wCmtXls = window.open('pop_event_Comment_xls.asp?eC='+ecd,'pXls','width=400,height=150');
		 wCmtXls.focus();
	}

	//2015.05.19 ���¿�(Ǫ�������� �̺�Ʈ������ �ӽ� ����-�̺�Ʈ ������ ��������)
	function popCommentXLS2(ecd) {
		 var wCmtXls = window.open('pop_event_Comment_xls_2.asp?eC='+ecd,'pXls','width=400,height=150');
		 wCmtXls.focus();
	}

	function popBBSXLS(ecd) {
		 var wBBSXls = window.open('pop_event_board_xls.asp?eC='+ecd,'pXls','width=400,height=150');
		 wBBSXls.focus();
	}

	//���޸� ǥ��
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

	//�̺�ƮŸ�� ���� ������ ���� ���� ���ó��
	function jsChkType(sType,frm){
		if(!frm.checked){
			if(confirm(sType +"������ ������ ��� �ش� "+sType+"���� ���µ� ����ó���˴ϴ�. ������ �����Ͻðڽ��ϱ�?")){
				return;
			}else{
				frm.checked = true;
			}
		}
	}
	// ��� ��ũ���� Eable
	function jsEvtLink(bln){
		if (bln) {
			$("#elUrl").attr("readonly",true);
			$("#elUrl").attr("class","");
			$("#elUrl").addClass("text_ro");
		}else{
			$('#elUrl').removeAttr('readonly');
			$("#elUrl").attr("class","");
			$("#elUrl").addClass("text");
		}
	}

	// ��α�URL�±� �˻�(�ڸ�Ʈ�� üũ�� �Ǿ��־�� ����)
	function jsChkBlogEnable() {
		if($('#isblogurl').prop('checked') == true) {
			if($('#chComm').prop('checked') == false) {
				alert("��α�URL����� �ڸ�Ʈ�� �־�߸� ��밡���մϴ�. �ڸ�Ʈ���θ� �������ּ���.");
				$('#isblogurl').prop('checked',false);
			}
		}
	}

	function workerlist()
	{
		var openWorker = null;
		var worker = frmEvt.selMId.value;
		openWorker = window.open('PopWorkerList.asp?worker='+worker+'&department_id=7','openWorker','width=700,height=570,scrollbars=yes');
		openWorker.focus();
	}

	function workerDel()
	{
		var frm = document.frmEvt;

		frm.selMId.value = "";
		frm.doc_workername.value = "";
	}

	// ��ǰ���� ����Ʈ�˾�
	function jsItemcopylist(){
		var winLast,eKind;
		winLast = window.open('pop_event_itemlist.asp?menupos=<%=menupos%>&eC=<%=eCode%>','pLast','width=550,height=600, scrollbars=yes')
		winLast.focus();
	}
	// ��ǰ �ʱ�ȭ
	function jsItemclear(){
		var frm = document.frmitemclear;

		if(confirm("��ǰ �ʱ�ȭ�� �Ͻðڽ��ϱ�?\n\n��ǰ �ʱ�ȭ�� ������ ������ �Ұ��� �մϴ�.")){
			frm.target = "FrameCKP";
			//frm.target = "blank";
			frm.action = "/admin/eventmanage/event/event_process.asp";
			frm.submit();
		}
	}

	function chgex(v){
		for (i=1;i<4;i++)
		{
			if (v == i)
			{
				$("#notice"+i).css("display","block");
			}else{
				$("#notice"+i).css("display","none");
			}
		}
	}
//-->
</script>
<form name="frmitemclear" method="post">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="imod" value="IC">
</form>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<form name="frmEvt" method="post" action="event_process.asp" onSubmit="return jsEvtSubmit(this);" style="margin:0px;">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="imod" value="U">
<input type="hidden" name="ban" value="<%=ebimg%>">
<input type="hidden" name="ban2010" value="<%=ebimg2010%>">
<input type="hidden" name="banMo" value="<%=ebimgMo%>">
<input type="hidden" name="banMoToday" value="<%=ebimgToday%>">
<input type="hidden" name="banMoList" value="<%=ebimgMo2014%>">
<input type="hidden" name="icon" value="<%=eicon%>">
<input type="hidden" name="main" value="<%=emimg%>">
<input type="hidden" name="gift" value="<%=gimg%>">
<input type="hidden" name="strparm" value="<%=strparm%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="etcitemban" value="<%=eEtcitemimg%>">
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
<tr>
	<td>  <img src="/images/icon_arrow_link.gif" align="absmiddle"> �̺�Ʈ ���� ���  </font></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		   <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>�̺�Ʈ�ڵ�</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0" >
		   			<tr>
		   				<td>
							<%=eCode%>
							<input type="button" value="��ǰ ����" onclick="jsItemcopylist();" class="button"/>
							<%
								Select Case ekind
									Case "19","25","26"
										Response.write "<input type='button' value='��ǰ�ʱ�ȭ' onclick='jsItemclear();' class='button' />"
								End select
							%>
						</td>
		   				<td>
						<%
							'�̺�Ʈ ������ ���� ����Ʈ��ũ ������ ����
							Select Case ekind
								Case "7"		'��Ŭ���ڵ�
									Response.Write "<td><a href='" & vwwwUrl & "/guidebook/weekly_coordinator.asp?eventid=" & eCode & "' target='_blank'>�̸�����</a></td>"
								Case "13"		'��ǰ �̺�Ʈ
									Response.Write "<td><a href='" & vwwwUrl & "/shopping/category_prd.asp?itemid=" & eitemid & "' target='_blank'>�̸�����</a></td>"
								Case "14"		'��ǳ���±�
									Response.Write "<td><a href='" & vwwwUrl & "/guidebook/picnic/picnic.asp?eventid=" & eCode & "' target='_blank'>�̸�����</a></td>"
								Case "16"		'�귣�� �������
									Response.Write "<td><a href='" & vwwwUrl & "/street/street_brand_sub06.asp?makerid=" & ebrand & "&shop_event_yn=Y&shop_event_confirm=Y&shopview=3' target='_blank'>�̸�����</a></td>"
								Case "22"		'DAY&(���̾ص�)
									Response.Write "<td><a href='" & vwwwUrl & "/guidebook/dayand.asp?eventid=" & eCode & "' target='_blank'>�̸�����</a></td>"
								Case "26"		'�����
									Response.Write "<td><a href='" & vmobileUrl & "/event/eventmain.asp?eventid=" & eCode & "' target='_blank'>�̸�����</a></td>"
								Case Else		'�������� �� ��Ÿ
									Response.Write "<td><a href='" & vwwwUrl & "/event/eventmain.asp?eventid=" & eCode & "' target='_blank'>�̸�����</a></td>"
							End Select
						%>
		   				</td>
		   				<td align="right">
		   				<% If sKind = "2" Then %>
		   					<input type="button" value="�Ѹ���List" onClick="window.open('/admin/eventmanage/oneline/?eC=<%=eCode%>&esday=<%=esday%>','oneline','width=600,height=500,scrollbars=yes');">
		   					<img src="/images/icon_excel_reply.gif" alt="�ڸ�Ʈ ������ Excel�ٿ�ε�" onClick="location.href='/admin/eventmanage/oneline/oneline_excel.asp?eC=<%=eCode%>&esday=<%=esday%>';" style="cursor:pointer" align="absmiddle">
		   				<% Else %>
		   					<img src="/images/icon_excel_reply.gif" alt="�ڸ�Ʈ ������ Excel�ٿ�ε�" onClick="popCommentXLS(<%=eCode%>);" style="cursor:pointer" align="absmiddle">
		   					<img src="/images/icon_excel_bbs.gif" alt="�Խ��� ������ Excel�ٿ�ε�" onClick="popBBSXLS(<%=eCode%>);" style="cursor:pointer" align="absmiddle">
		   				<% End If %>
		   					<img src="/images/icon_excel_vote.gif" alt="���� ������ Excel�ٿ�ε�" onClick="window.open('pop_event_votelist_xls.asp?eC=<%=eCode%>','voteXls','width=400,height=150');" style="cursor:pointer" align="absmiddle" title ="xls �ٿ�ε� ȸ�����">
							<img src="/images/icon_excel_vote.gif" alt="���� ������ Excel�ٿ�ε� ��ȸ��"  title ="xls �ٿ�ε� ��ȸ��" onClick="window.open('pop_event_votelist_guest_xls.asp?eC=<%=eCode%>','voteXls','width=400,height=150');" style="cursor:pointer" align="absmiddle">

						<% If eCode = "65010" and session("ssBctId") = "stella0117" then %>
							<img src="/images/icon_excel_reply.gif" alt="�õ����������� Excel�ٿ�ε�" onClick="popCommentXLS2(<%=eCode%>);" style="cursor:pointer" align="absmiddle">
						<% End if %>
		   				</td>
		   			</tr>
		   			</table>
		   		</td>
		   	</tr>
		    <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>�������</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="radio" name="using" value="Y" <%IF eusing="Y" THEN%>checked<%END IF%>>Yes <input type="radio" name="using" value="N" <%IF eusing="N" THEN%>checked<%END IF%>>No
		   		</td>
		   	</tr>
		   <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>����</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventkind",ekind,False,"onChange=javascript:jsChkSubj(this.value);"%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>��ü</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventmanager",eman,False,""%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>����</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="hidden" name="eventscope" value="<%=escope%>">
		   			<label><input type="checkbox" name="chkEscope" <%=chkIIF((escope="1" or escope="2"),"checked","")%> onclick="jsSetPartner()"> 10x10</label>
		   			<label><input type="checkbox" name="chkEscope" <%=chkIIF((escope="1" or escope="3"),"checked","")%> onclick="jsSetPartner()"> ���޸�</label>
		   			<span id="spanP" style="display:<%=chkIIF((escope="1" or escope="3"),"","none")%>">
		   			<select name="selP">
		   				<option value="">--���޸� ��ü--</option>
		   				<% sbOptPartner selPartner%>
		   			</select>
		   			</span>
		   		</td>
		   	</tr>
		   	<tr id="eNameTr_A" style="display:<% if ekind="16" then Response.Write "none" %>;">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>�̺�Ʈ��</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sEN" size="80" maxlength="120" value="<%=ename%>">
		   		</td>
		   	</tr>
			<tr style="display:<% if ekind="16" then Response.Write "none" %>;">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>2014�̺�Ʈ<br/>����ī��</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="subsEN" size="60" maxlength="60" value="<%=enamesub%>">
		   		</td>
		   	</tr>
			<tr id="eNameTr_C">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>���� �̺�Ʈ��</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sENEng" size="60" maxlength="60" value="<%=enameEng%>">
		   		</td>
		   	</tr>
		   	<tr id="eNameTr_B" style="display:<% if ekind<>"16" then Response.Write "none" %>;">
		   	<%
		   		'// �귣�������̸� ������ �������� ǥ��
		   		dim arrEname
				arrEname = Split(ename,"|")
				if Ubound(arrEname)<2 then
					arrEname = ename & "|0|0"
					arrEname = Split(arrEname,"|")
				end if

				If enameEng <> "" then
					Dim arrEnameEng
					arrEnameEng = Split(enameEng,"|")
					if Ubound(arrEnameEng)<2 then
						arrEnameEng = enameEng & "|0|0"
						arrEnameEng = Split(arrEnameEng,"|")
					end If
				End If
		   	%>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>�̺�Ʈ��<br>�� ������</B></td>
		   		<td bgcolor="#FFFFFF">
					�̺�Ʈ��: <input type="text" name="sEDN" size="60" maxlength="60" value="<%=arrEname(0)%>"><br>
					<% If enameEng <> "" Then %>
		   			�����̺�Ʈ��: <input type="text" name="sEDNEng" size="60" maxlength="60" value="<%=arrEnameEng(0)%>"><br>
					<% End If %>
		   			������: ���� <input type="text" name="sSDc" size="4" value="<%=arrEname(1)%>" style="text-align:right;">% ~
		   			�ְ� <input type="text" name="sMDc" size="4" value="<%=arrEname(2)%>" style="text-align:right;">%<br>
		   			<font color=gray>�غ귣�� ��Ʈ��Ʈ�� ������ �������Դϴ�. ������ ��ǰ���� ������� ������ ��ǰ���� ���� ������ �������ּ���.<br>�̺�Ʈ ��ũ�� �귣�� ��Ʈ��Ʈ�� ����Ǵ� �ݵ�� �󼼳��뿡 �귣�带 �������ּ���.</font>
		   		</td>
		   	</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>���� ī��</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<table width="100%" border="0" cellpadding="0" cellspacing="0">
		   			<tr>
		   				<td width="50%" style="padding-right:3px;"><textarea name="subcopyK" style="width:100%; height:80px;" onclick="if(this.value=='�ѱ�')this.value='';" onblur="if(this.value=='')this.value='�ѱ�';" value="<%=subcopyK%>"><%=chkiif(subcopyK="","�ѱ�",subcopyK)%></textarea></td>
		   				<td width="50%"><textarea name="subcopyE" style="width:100%; height:80px;" onclick="if(this.value=='����')this.value='';" onblur="if(this.value=='')this.value='����';" value="<%=subcopyE%>"><%=chkiif(subcopyE="","����",subcopyE)%></textarea></td>
		   			</tr>
		   			</table>
		   		</td>
			</tr>
		   	<tr>
		   		<td rowspan="2" align="center" bgcolor="<%= adminColor("tabletop") %>"><B>�Ⱓ</B></td>
		   		<td bgcolor="#FFFFFF">
		   		<%IF estate = 9 THEN%>
		   			������ : <%=esday%><input type="hidden" name="sSD" size="10" value="<%=esday%>">
		   			~ ������ : <%=eeday%> <input type="hidden" name="sED" value="<%=eeday%>" size="10" >
		   		<%ELSE%>
		   			������ : <input type="text" name="sSD" size="10" value="<%=esday%>" onClick="jsPopCal('sSD');"  style="cursor:hand;">
		   			~ ������ : <input type="text" name="sED" value="<%=eeday%>" size="10" onClick="jsPopCal('sED');" style="cursor:hand;">
		   		<%END IF%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td  bgcolor="#FFFFFF">
		   			��÷ ��ǥ�� : <input type="text" name="sPD" value="<%=epday%>" size="10" onClick="jsPopCal('sPD');" style="cursor:hand;">
		   			(��÷�ڰ� �ִ� ��쿡�� ���)
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>����</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%
		   				sbGetOptStatusCodeValue "eventstate",estate,false,""
		   				''if ekind="22" then
		   				''	'//���̾ص�� ��������Ʈ�� ����ؼ� �������
		   				''	sbGetOptStatusCodeValue "eventstate",estate,false,""
		   				''else
		   				''	sbGetOptStatusCodeAuth "eventstate",estate,"M",""
		   				''end if
		   			%>
		   			<input type="hidden" name="eOD" value="<%=dopendate%>">
		   			<input type="hidden" name="eCD" value="<%=dclosedate%>">
		   			<%IF not isnull(dopendate) THEN%><span style="padding-left:10px;">  ����ó���� : <%=dopendate%>  </span><%END IF%>
		   			<%IF not isnull(dclosedate) THEN%>/ <span style="padding-left:10px;">  ����ó���� : <%=dclosedate%>  </span><%END IF%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>�߿䵵</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventlevel",elevel,False,""%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>���Ĺ�ȣ</b></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sortNo" value="<%=evt_sortNo%>" size="6" maxlength="5" style="text-align:right;" />
		   			(�ؼ��ڰ� Ŭ���� �켱ǥ�� �˴ϴ�. / Day&:ȸ��)
		   		</td>
		   	</tr>
		   		<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>����</b></td>
		   		<td bgcolor="#FFFFFF">
		   			�󼼳��� �߰���� <input type="checkbox" name="chkDisp" onClick="jsChkDisp();" <%IF echkdisp= 1 THEN%>checked<%END IF%>>
		   		</td>
		   	</tr>
		</table>
	</td>

</tr>
<tr>
	<td>
	 <div id="eDetail" style="display:<%IF echkdisp<> 1 THEN%>none;<%END IF%>">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
					   	<tr>
					   		<td width="100"  align="center" bgcolor="<%= adminColor("tabletop") %>">���� ī�װ�</td>
					   		<td bgcolor="#FFFFFF">
					   			<%'DrawSelectBoxCategoryOnlyLarge "selCategory", ecategory,"" %>
					   			<!-- #include virtual="/common/module/categoryselectbox_event.asp"-->
					   		</td>
					   	</tr>
					   	<tr>
					   		<td width="100"  align="center" bgcolor="<%= adminColor("tabletop") %>">���� ī�װ�</td>
					   		<td bgcolor="#FFFFFF">
					   			<%=fnDispCateSelectBox(1,"","dispcate",eDispCate,"") %>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�귣��</td>
					   		<td bgcolor="#FFFFFF">
					   			<% drawSelectBoxDesignerwithName "ebrand", ebrand %>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ Ÿ��</td>
					   		<td bgcolor="#FFFFFF">
					   		<input type="checkbox" name="chSale" <%IF esale THEN%>checked onClick="jsChkType('����',this);"<%END IF%> value="1">����
					   		<input type="checkbox" name="chGift" <%IF egift  THEN%>checked  onClick="jsChkType('����ǰ',this);"<%END IF%> value="1">����ǰ
					   		<input type="checkbox" name="chCoupon" <%IF ecoupon THEN%>checked  onClick="jsChkType('����',this);"<%END IF%> value="1">����
					   		<input type="checkbox" name="chOnlyTen" <%IF eonlyten THEN%>checked<%END IF%> value="1">Only-TenByTen
					   		<input type="checkbox" name="chOneplusone" <%IF eOneplusone THEN%>checked<%END IF%> value="1">1+1
					   		<input type="checkbox" name="chFreedelivery" <%IF eFreedelivery THEN%>checked<%END IF%> value="1">������
					   		<input type="checkbox" name="chBookingsell" <%IF eBookingsell THEN%>checked<%END IF%> value="1">�����Ǹ�
					   		<input type="checkbox" name="chDiary" <%IF eDiary THEN%>checked<%END IF%> value="1">DiaryStory
					   		</td>
						</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ ���</td>
					   		<td bgcolor="#FFFFFF">
					   		<input type="checkbox" name="chComm" id="chComm" <%IF ecomment THEN%>checked<%END IF%> value="1" onClick="jsChkBlogEnable()">�ڸ�Ʈ
					   		<input type="checkbox" name="chBbs" <%IF ebbs THEN%>checked<%END IF%> value="1" >�Խ���
					   		<input type="checkbox" name="chItemps" <%IF eitemps THEN%>checked<%END IF%> value="1" >��ǰ�ı�
					   		<input type="checkbox" name="isblogurl" id="isblogurl" <%IF eisblogurl THEN%>checked<%END IF%> value="1" onClick="jsChkBlogEnable()">Blog URL
					   		<!--<input type="checkbox" name="chApply" <%IF eapply THEN%>checked<%END IF%> value="1" >����-->
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ ��ũ Ÿ��</td>
					   		<td bgcolor="#FFFFFF">
					   			<label><input type="radio" name="elType" value="E" onclick="jsEvtLink(true);"  <% IF elktype="E" Then %>checked<% End IF %> >�̺�Ʈ</label>
					   			<label><input type="radio" name="elType" value="I" onclick="jsEvtLink(false);" <% IF elktype="I" Then %>checked<% End IF %>>�����Է�</label>
					   			&nbsp;<input type="text" name="elUrl" id="elUrl" size="40" maxlength="128" value="<%= elkurl %>" <% IF elktype="E" THEN%>class="text_ro" readOnly<%ELSE%>class="text"<%END IF %>>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">��ǰ���Ĺ��</td>
					   		<td bgcolor="#FFFFFF">
					   			<%sbGetOptEventCodeValue "itemsort",eisort,False,""%>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�������̳�</td>
					   		<td bgcolor="#FFFFFF">
					   			<%sbGetDesignerid "selDId",edid,""%>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�����</td>
					   		<td bgcolor="#FFFFFF"><% sbGetwork "selMId",emid,"" %></td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�۾����޻���</td>
					   		<td bgcolor="#FFFFFF">
					   			�۾����� <input type="text" name="sWorkTag" size="20" maxlength="16" value="<%= sWorkTag %>" class="text"> <font color="darkgray">(for Designer)</font><br />
					   			<textarea name="tFwd" rows="15" style="width:100%;font-size:12px;font-family:'Malgun Gothic',dotum;"><%=efwd%></textarea>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">Comment Title</td>
					   		<td bgcolor="#FFFFFF">
					   			(200�� �̳�)		   			<Br>
					   			<textarea name="eCT" rows="2" style="width:100%;"><%=ecommenttitle%></textarea>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">Tag</td>
					   		<td bgcolor="#FFFFFF">
					   			(250�� �̳�)		   			<Br>
					   			<textarea name="eTag" rows="2" style="width:100%;"><%=etag%></textarea>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">���� �̺�Ʈ�ڵ�</td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="text" name="eLC" size="6" maxlength="10" value="<%=elinkcode%>">
					   			<% if sKind="19" then Response.Write " <font color=darkred>�� PC���� �̺�Ʈ ��ȣ �Է½� �ڸ�Ʈ ����</font>" %>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">���̽��� �� ����</td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="button" value="���̽��� �� ���� ���â" class="button" onClick="window.open('pop_event_facebookapp.asp?ecode=<%=eCode%>','facebookpop','width=500,height=400');">

					   		</td>
					   	</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td style="padding: 10 0 5 0"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ȭ���̹��� ���</td></tr>
			<tr>
				<td>
					<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
						<tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">ȭ�鱸��</td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="checkbox" name="chkFull" value="0" <%IF not blnFull THEN%>checked<%END IF%> onclick="if(this.checked) chkWide.checked=false;"> ���� �޴�&nbsp;&nbsp;
					   			<input type="checkbox" name="chkWide" value="1" <%IF blnWide THEN%>checked<%END IF%> onclick="if(this.checked) chkFull.checked=false;"> ���̵� ������
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">��ǰ����Ʈ ��Ÿ��<br/>(Mobile,App ��)</td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="radio" name="itemlisttype"  value="1"  <%IF eItemListType = "1" THEN%>checked<%END IF%>> ������&nbsp;&nbsp;
					   			<input type="radio" name="itemlisttype"  value="2"  <%IF eItemListType = "2" THEN%>checked<%END IF%>> ����Ʈ��&nbsp;&nbsp;
					   			<input type="radio" name="itemlisttype"  value="3"  <%IF eItemListType = "3" THEN%>checked<%END IF%>> BIG��
					   		</td>
					   	</tr>

					   	<tr>
					   		<td align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">��ǰ����</td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="checkbox" name="chkIteminfo"  value="1"  <%IF blnIteminfo THEN%>checked<%END IF%>> �����
					   		</td>
					   	</tr>
						<tr>
					   		<td align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">��ǰ ��������<br/><font color="#BB8866">[������ ���ΰ�<br/>���⿩��]</font></td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="checkbox" name="chkItemprice"  value="1"  <%IF blnitemprice THEN%>checked<%END IF%>> �������
					   		</td>
					   	</tr>
						<tr>
					   		<td align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ �Ⱓ<br/>���⿩��</td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="checkbox" name="dateview"  value="1"  <%IF eDateView THEN%>checked<%END IF%>> �������
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">��ǥ��ǰ����<br/>��<br/>���</td>
					   		<td bgcolor="#FFFFFF">
					   			<font color="red"><b>�� ī�װ����ΰ� �������̺�Ʈ ����Ʈ�� ������ �̹���.<br/>��ǥ��ǰ�̹����� �ȳ����� ��ǥ��ǰ�ڵ带 �ݵ�� �־����.<br/>��ǥ��ǰ�̹����� ������ ��ǥ��ǰ�ڵ��� �⺻ �̹����� ����ϰ� ��.</b></font><br/>
								��ǥ��ǰ�ڵ� : <input type="text" name="etcitemid" value="<%=eEtcitemid%>"><br/>
								��ǥ��ǰ�̹���(420x420) : <input type="button" name="etcitem" value="��ǰ��ǥ���" onClick="jsSetImg('<%=eFolder%>','<%=eEtcitemimg%>','etcitemban','etciitem')" class="button">
					   			<div id="etciitem" style="padding: 5 5 5 5">
					   				<%IF eEtcitemimg <> "" THEN %>
					   				<img  src="<%=eEtcitemimg%>" border="0">
					   				<a href="javascript:jsDelImg('etcitemban','etciitem');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
					   		</td>
					   	</tr>
						<tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">2011 �⺻���</td>
					   		<td bgcolor="#FFFFFF">
					   		<input type="button" name="btnBan" value="2011 ����̹��� ���" onClick="jsSetImg('<%=eFolder%>','<%=ebimg%>','ban','spanban')" class="button">
					   			<div id="spanban" style="padding: 5 5 5 5">
					   				<%IF ebimg <> "" THEN %>
					   				<img  src="<%=ebimg%>" border="0">
					   				<a href="javascript:jsDelImg('ban','spanban');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
					   		</td>
					   	</tr>
						<tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">2010 �⺻���</td>
					   		<td bgcolor="#FFFFFF">
					   		<input type="button" name="btnBan2010" value="2010 ����̹��� ���" onClick="jsSetImg('<%=eFolder%>','<%=ebimg2010%>','ban2010','spanban2010')" class="button">
					   			<div id="spanban2010" style="padding: 5 5 5 5">
					   				<%IF ebimg2010 <> "" THEN %>
					   				<img  src="<%=ebimg2010%>" border="0">
					   				<a href="javascript:jsDelImg('ban2010','spanban2010');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
					   		</td>
					   	</tr>
						<!-- <tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">����� ����Ʈ���(�� ������Ʈ�� ���� ����)</td>
					   		<td bgcolor="#FFFFFF">
					   		<input type="button" name="btnBanMo" value="����� ����Ʈ��� ���" onClick="jsSetImg('<%=eFolder%>','<%=ebimgMo%>','banMo','spanbanMo')" class="button">
					   			<div id="spanbanMo" style="padding: 5 5 5 5">
					   				<%IF ebimgMo <> "" THEN %>
					   				<img  src="<%=ebimgMo%>" border="0">
					   				<a href="javascript:jsDelImg('banMo','spanbanMo');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
					   			<p style="color:#602030;font-size:11px;">�� ���� �̹��� : JPEG, 50%, 560px �� 380px</p>
					   		</td>
					   	</tr> -->
						<!-- <tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">����� Today���(�� ������Ʈ�� ���� ����)</td>
					   		<td bgcolor="#FFFFFF">
					   		<input type="button" name="btnBanToday" value="����� Today-enjoyevent��� ���" onClick="jsSetImg('<%=eFolder%>','<%=ebimgToday%>','banMoToday','spanbanMoToday')" class="button">
					   			<div id="spanbanMoToday" style="padding: 5 5 5 5">
					   				<%IF ebimgToday <> "" THEN %>
					   				<img  src="<%=ebimgToday%>" border="0">
					   				<a href="javascript:jsDelImg('banMoToday','spanbanMoToday');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
								<p style="color:#602030;font-size:11px;">�� ���� �̹��� : JPEG, 50%, 600px �� 270px</p>
					   		</td>
					   	</tr> -->
						<tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">2014 �����<br/>����Ʈ���</td>
					   		<td bgcolor="#FFFFFF">
					   		<input type="button" name="btnMoBan2014" value="2014 ����� ����Ʈ ���" onClick="jsSetImg('<%=eFolder%>','<%=ebimgMo2014%>','banMoList','spanbanMoList')" class="button">
					   			<div id="spanbanMoList" style="padding: 5 5 5 5">
					   				<%IF ebimgMo2014 <> "" THEN %>
					   				<img  src="<%=ebimgMo2014%>" border="0">
					   				<a href="javascript:jsDelImg('banMoList','spanbanMoList');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
								<p style="color:#602030;font-size:11px;">�� ���� �̹��� : JPEG, 50%, 640px �� 340px</p>
					   		</td>
					   	</tr>
						<tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">�⺻������ </td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="button" name="btnicon" value="����Ʈ������ ���" onClick="jsSetImg('<%=eFolder%>','<%=eicon%>','icon','spanicon')" class="button">
					   			<div id="spanicon" style="padding: 5 5 5 5">
					   				<%IF eicon <> "" THEN %>
					   				<img  src="<%=eicon%>">
					   				<a href="javascript:jsDelImg('icon','spanicon');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">����ǰ �̹��� </td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="button" name="btnicon" value="����ǰ�̹��� ���" onClick="jsSetImg('<%=eFolder%>','<%=gimg%>','gift','spangift')" class="button">
					   			<div id="spangift" style="padding: 5 5 5 5">
					   				<%IF gimg <> "" THEN %>
					   				<a href="javascript:jsImgView('<%=gimg%>')"><img  src="<%=gimg%>" width="400" border="0"></a>
					   				<a href="javascript:jsDelImg('gift','spangift');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td width="30" align="center" rowspan="2"  bgcolor="<%= adminColor("tabletop") %>"> ��<br>��<br>Ʈ<br><br>��<br>��<br>��<br>��<br>��<br> </td>
					   		<td width="65" align="center"  bgcolor="<%= adminColor("tabletop") %>">ȭ�����ø�</td>
					   		<td bgcolor="#FFFFFF"><%sbGetOptEventCodeValue "eventview",etemp,false,"onchange=""jsChangeFrm(this.value);"""%></td>
					   	</tr>
					   	<tr>
					   		<td width="65" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̹���<br>&<br>HTML</td>
					   		<td bgcolor="#FFFFFF">
					   			<!-- 1.���� ž-->
					   			<div id="divFrm1" style="display:;">
						   			<input type="button" name="btnMain" value="����TOP�̹��� ���" onClick="jsSetImg('<%=eFolder%>','<%=emimg%>','main','spanmain')" class="button">
						   			<div id="spanmain" style="padding: 5 5 5 5">
						   				<%IF emimg <> "" THEN %>
						   				<a href="javascript:jsImgView('<%=emimg%>')"><img  src="<%=emimg%>" width="400" border="0"></a>
						   				<a href="javascript:jsDelImg('main','spanmain');"><img src="/images/icon_delete2.gif" border="0"></a>
						   				<%END IF%>
						   			</div>
					   				<hr>
									<span style="color:red;font-weight:800;cursor:pointer;" onclick="chgex('1');">PC-WEB ����</span>||<span style="color:red;font-weight:800;cursor:pointer;" onclick="chgex('2');">Mobile-WEB ����</span>||<span style="color:red;font-weight:800;cursor:pointer;" onclick="chgex('3');">APP ����</span>
									<div id="notice1" style="display:block">
						   			&lt;map name="Mainmap"&gt;<br>
									<font color="blue">��ǰ������ ��ũ��</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoProduct('<font color="blue">��ǰ��ȣ</font>');" onfocus="this.blur();"&gt;<br>
									<font color="blue">�̺�Ʈ�������� ��ũ��</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoEventMain('<font color="blue">�̺�Ʈ�ڵ�</font>');" onfocus="this.blur();"&gt;<br>
									<font color="blue">�̺�Ʈ �׷� �������� ��ũ��</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoEventGroupMain('<font color="blue">�̺�Ʈ�ڵ�</font>','<font color="blue">�׷��ڵ�</font>');" onfocus="this.blur();"&gt;<br>
									<font color="blue">�̺�Ʈ ����ǰ �˾� ��ũ��</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:popShowGiftImg('<font color="blue">�̺�Ʈ�ڵ�</font>');" onfocus="this.blur();"&gt;<br>
									<font color="blue">�귣�������� ��ũ��</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:GoToBrandShop('<font color="blue">�귣����̵�</font>');" onfocus="this.blur();"&gt;<br>
									&lt;/map&gt;<br>
									<font color="blue">���帮�� ���� ��ũ��</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoEventMain_New('<font color="blue">�̺�Ʈ�ڵ�</font>');" onfocus="this.blur();"&gt;<br>
									&lt;/map&gt;
									</map> <br>
									<font color="blue">������ Ÿ��Ʋ �̹����� ��ũ��</font><br>
									&lt;area shape="circle" coords="186,250,144" href="#event_namelink1" onfocus="this.blur();"&gt;<br>
									href="#event_namelink2" href="#event_namelink3" ��� href�� ���ڸ� �ٲ���. &lt;area������ ĭ�� ���������� �� ����.<br>
									</div>
									<div id="notice2" style="display:none">
										<font color="blue">��ǰ������ ��ũ��</font><br>
										&lt;a href="/category/category_itemprd.asp?itemid=<span style="color:red">��ǰ�ڵ�</span>"&gt; ��ǰ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">�̺�Ʈ�������� ��ũ��</font><br>
										&lt;a href="/event/eventmain.asp?eventid=<span style="color:red">�̺�Ʈ�ڵ�</span>"&gt; �̺�Ʈ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">�̺�Ʈ �׷� �������� ��ũ��</font><br>
										&lt;a href="/event/eventmain.asp?eventid=<span style="color:red">�̺�Ʈ�ڵ�</span>&eGc=<span style="color:red">�׷��ڵ�</span>"&gt; �̺�Ʈ �׷� ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">�귣�������� ��ũ��</font><br>
										&lt;a href="/street/street_brand.asp?makerid=<span style="color:red">�귣���ڵ�</span>"&gt; �귣�������� ��ũ &lt;/a&gt;<br>
									</div>
									<div id="notice3" style="display:none">
										�������������� �̵��Ҷ���<br/>
										<font color="blue">��ǰ������ ��ũ��</font><br>
										&lt;a href="/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid=<span style="color:red">��ǰ�ڵ�</span>"&gt; ��ǰ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">�̺�Ʈ�������� ��ũ��</font><br>
										&lt;a href="/apps/appCom/wish/web2014/event/eventmain.asp?eventid=<span style="color:red">�̺�Ʈ�ڵ�</span>"&gt; �̺�Ʈ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">�̺�Ʈ �׷� �������� ��ũ��</font><br>
										&lt;a href="/apps/appCom/wish/web2014/event/eventmain.asp?eventid=<span style="color:red">�̺�Ʈ�ڵ�</span>&eGc=<span style="color:red">�׷��ڵ�</span>"&gt; �̺�Ʈ �׷� ������ ��ũ &lt;/a&gt;<br>
										<br>
										���˾����� ������ ������<br/>
										�ؼ��۾� iframe�߰��Ҷ� �϶� <span style="color:blue">parent.</span> �Լ������� �߰���<br/>
										ex) &lt;a href="#" onclick="<span style="color:blue">parent.</span>fnAPPpopupProduct('<span style="color:red">��ǰ�ڵ�</span>'); return false;"&gt; ��ǰ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">��ǰ������ ��ũ��</font><br>
										&lt;a href="#" onclick="fnAPPpopupProduct('<span style="color:red">��ǰ�ڵ�</span>'); return false;"&gt; ��ǰ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">�̺�Ʈ�������� ��ũ��</font><br>
										&lt;a href="#" onclick="fnAPPpopupEvent('<span style="color:red">�̺�Ʈ�ڵ�</span>'); return false;"&gt; �̺�Ʈ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">�귣�������� ��ũ��</font><br>
										&lt;a href="#" onclick="fnAPPpopupBrand('<span style="color:red">�귣���</span>'); return false;"&gt; �귣�� ��ũ &lt;/a&gt;<br>
										<font color="blue">ī�װ� ��ũ��</font><br>
										&lt;a href="#" onclick="fnAPPpopupCategory('<span style="color:red">ī�װ���ȣ</span>'); return false;"&gt; ī�װ� ��ũ &lt;/a&gt;<br><br>
										�ؾ� �� ���� ��ũ��Ʈ��<br/>
										&lt;script&gt;<br/>
										var chkapp = navigator.userAgent.match('tenapp');<br/>
										if ( chkapp ){<br/>
										&nbsp;&nbsp;&nbsp;//�ۿ��� ��ũ��Ʈ<br/>
										}else{<br/>
										&nbsp;&nbsp;&nbsp;//����Ͽ��� ��ũ��Ʈ<br/>
										}<br/>
										&lt;/script&gt;<br/>
									</div>
									<br>
									<b>�̹��� ��� http://<font color="RED">webimage.</font>10x10.co.kr/event/XXX/</b> �� ����Ǿ����ϴ�.<br>
						   			<textarea name="tHtml" rows="20" style="width:100%;font-size:11px;"><%=ehtml%></textarea>
					   			</div>
					   			<!-- 3.�׷���-->
					   			<div id="divFrm3" style="display:none;">
					   				<iframe id="iframG" src="about:blank" frameborder="0" width="100%" class="autoheight"></iframe>
					   			</div>
					   			<!-- /3.�׷���-->
					   				<!-- 5.���۾�-->
					   			<div id="divFrm5" style="display:none;">
					   				<table border="0" cellpadding="1" cellspacing="3" class="a">
					   					<tr>
					   						<td>
					   						    <!-- <input type="button" value="�̹�������"  onclick="TnFtpUpload('D:/home/cube1010/imgstatic/event/<%= eFolder%>/','/event/<%= eFolder%>/');" class="input_b"> -->
					   						    <input type="button" value="�̹�������(��)"  onclick="jsManageEventImageNew('<%= eFolder%>')" class="input_b">
					   						    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					   						    <input type="button" value="�̹�������(��)"  onclick="jsManageEventImage('<%= eFolder%>')" class="input_b">
					   						    <p>
					   						    <b>�̹�������(��)</b> : ������ ����� �̹��� ����<br>
					   						    <b>�̹�������(��)</b> : ������ ����� �̹��� ����Ʈ(�̹����߰�����. ���ο� �̹��� �߰��� �̹�������(��)������.)<br>
					   						    �� �̺�Ʈ �̹��� �ý��� ���� �������� eventIMG ��� ���ο� ������ �̺�Ʈ����Ҵ ������ �߰��Ͽ� �� �ȿ� �̺�Ʈ�ڵ庰 ������ �����ϰ� �˴ϴ�.<br>
					   						    ���� ����� ���� �ڿ� �̹�������(��)�� ����� �����ʰ� �̹�������(��)�� ����ϰ� �˴ϴ�.<br>
					   						    �׶������� ��������� �����ô��� �ý��۰��� ������ ���� ��ġ�̹Ƿ� ���عٶ��ϴ�.
					   						</td>
					   					</tr>
					   					<tr>
					   					    <td><b>�̹��� ��� http://<font color="RED">webimage.</font>10x10.co.kr/eventIMG/�̺�Ʈ����Ҵ/XXX/</b> �� ����Ǿ����ϴ�.</td>
					   					</tr>
					   					<tr>
					   						<td><textarea name="tHtml5" rows="25" style="width:100%;font-size:11px;"><%=ehtml5%></textarea></td>
					   					</tr>
					   				</table>
					   			</div>
					   			<!-- /5.���۾�-->
					   		</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</div>
	</td>

</tr>
<tr>
	<td width="100%" height="40" align="right">
		<input type="image" src="/images/icon_save.gif">
		<a href="index.asp?<%=strparm%>"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</table>
</form>
 <script language="javascript">
<!--
jsChangeFrm('<%=etemp%>');
//-->
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
