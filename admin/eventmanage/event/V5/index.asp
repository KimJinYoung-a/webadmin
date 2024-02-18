<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/index.asp
' Description :  �̺�Ʈ ��� - ȭ�鼳��
' History : 2007.02.07 ������ ����
'           2012.02.13 ������ - �̴ϴ޷� ��ü
'			2014.03.10 ������ - �����׸� ���̷�(fotoark), ���ְ�(arlejk) ���ܻ��� ����
'           2015.03 ������ - �̺�Ʈ ������
'           2017.04.14 ������ - ���� �����̳� �߰�
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
	Call fnSetEventCommonCode '�����ڵ� ���ø����̼� ������ ����

	'��������
	Dim cEvtList
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
	Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory, sCateMid ,sState,sKind,esale,egift,ecoupon,ebrand,eonlyten,etype_pc, etype_mo,isConfirm,eMng
	Dim strparm
	Dim edgid, edgid2,edgstat1,edgstat2, emdid, epsid, edpid, edgnm, edgnm2, emdnm, epsnm, edpnm, eDiary
	dim eopo,efd,ebs,enew
	dim blnWeb, blnMobile, blnApp, elevel
	dim dispCate, maxDepth
	dim blnReqPublish ,sSort, evt_template, evt_template_mo
	dim isResearch, mdtheme, startESP, endESP
	dim chComm, chItemps, chBbs, isblogurl, endlessView

	'�Ķ���Ͱ� �ޱ� & �⺻ ���� �� ����
	iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
	maxDepth = 2
	
	isResearch = requestCheckVar(Request("isResearch"),1)
	if isResearch ="" then isResearch ="0"
	'## �˻� #############################
	sDate 		= requestCheckVar(Request("selDate"),1)  	'�Ⱓ
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)

	sEvt 		= requestCheckVar(Request("selEvt"),10)  	'�̺�Ʈ �ڵ�/�� �˻�
	strTxt 		= requestCheckVar(Request("sEtxt"),256)

	sCategory	= requestCheckVar(Request("selC"),10) 		'ī�װ�
	sCateMid	= requestCheckVar(Request("selCM"),10) 		'ī�װ�(�ߺз�)
	dispCate	= requestCheckVar(Request("disp"),10) 		'���� ī�װ�
	sState		= requestCheckVar(Request("eventstate"),4)	'�̺�Ʈ ����
	 
	sKind 		= requestCheckVar(Request("eventkind"),32)	'�̺�Ʈ����
	edgid  		= requestCheckVar(Request("sDgId"),32)		'��� �����̳�
''	edgid2 		= requestCheckVar(Request("sDg2Id"),32)		'���� �����̳�
	emdid  		= requestCheckVar(Request("sMdId"),32)		'��� MD
	epsid  		= requestCheckVar(Request("sPsId"),32)		'��� �ۺ���
	edpid  		= requestCheckVar(Request("sDpId"),32)		'��� ������
	
	edgnm  		= requestCheckVar(Request("sdgnm"),32)		'��� �����̳�
''	edgnm2 		= requestCheckVar(Request("sdg2nm"),32)		'���� �����̳�
	emdnm  		= requestCheckVar(Request("smdnm"),32)		'��� MD
	epsnm  		= requestCheckVar(Request("spsnm"),32)		'��� �ۺ���
	edpnm  		= requestCheckVar(Request("sdpnm"),32)		'��� ������

	if Request("designerstatus")<>"" AND Request("designerstatus") <> "," then
		edgstat1	= requestCheckVar(Request("designerstatus")(1),2)		'��� �����̳� ����
		edgstat2	= requestCheckVar(Request("designerstatus")(2),2)		'���� �����̳� ����
	end if

	ebrand		= requestCheckVar(Request("ebrand"),32)		'�귣��
	esale		= requestCheckVar(Request("chSale"),2) 		'��������
	egift		= requestCheckVar(Request("chGift"),2)		'����ǰ����
	ecoupon	 	= requestCheckVar(Request("chCoupon"),2)	'��������
	eonlyten	= requestCheckVar(Request("chOnlyTen"),2)	'Only-TenByTen����
	eDiary		= requestCheckVar(Request("chDiary"),2)	'���̾ ����
	eopo		= requestCheckVar(Request("chopo"),1)	'���÷�����
	efd		= requestCheckVar(Request("chfd"),1)	'������
	ebs		= requestCheckVar(Request("chbs"),1)	'�����Ǹ�
	enew		= requestCheckVar(Request("chnew"),1)	'new
	
	blnWeb		= requestCheckVar(Request("isWeb"),1)
	blnMobile	= requestCheckVar(Request("isMobile"),1)
	blnApp		= requestCheckVar(Request("isApp"),1)
	
	dispCate 	= requestCheckvar(request("disp"),16)
	blnReqPublish= requestCheckvar(request("chkPus"),1)
	sSort       = requestCheckvar(request("sSort"),2)

	etype_pc	= requestCheckvar(request("eventtype_pc"),4)
	etype_mo	= requestCheckvar(request("eventtype_mo"),4)
	eMng    = requestCheckvar(request("eventmanager"),4)
	isConfirm	= requestCheckvar(request("blnCnfm"),1)
	mdtheme  	= requestCheckVar(Request("mdtheme"),1)		'MD��� �̺�Ʈ �׸�
	startESP = requestCheckVar(Request("startESP"),16)
	endESP = requestCheckVar(Request("endESP"),16)
	evt_template = requestCheckVar(Request("evt_template"),2)
	evt_template_mo = requestCheckVar(Request("evt_template_mo"),2)

	chComm		= requestCheckVar(Request("chComm"),1)
	chItemps		= requestCheckVar(Request("chItemps"),1)
	chBbs		= requestCheckVar(Request("chBbs"),1)
	isblogurl		= requestCheckVar(Request("isblogurl"),1)
	elevel		= requestCheckVar(Request("elevel"),1)
	endlessView		= requestCheckVar(Request("endlessView"),1)

	if isResearch="0" and sKind="" then
		skind="1,5,12,13,23,27,28,29,31"
	end if

	'�̺�Ʈ ù������ �����׸��� ���̵��� 
	IF (sKind="" and isResearch="0") or sKind="1,12" THEN
		if (session("ssAdminPsn")="11" or session("ssAdminPsn")="21") and (not ( session("ssBctId")="fotoark" or session("ssBctId")="arlejk" or session("ssBctId")="barbie8711")) then
			'MD�μ���� (��������,��ü,��ǰ,�귣��,���̾,�׽���,�űԵ����̳�) - ���̷�(fotoark), ���ְ�(arlejk), ����ȭ(barbie8711) ����
			sKind = "1,5,12,13,16,17,23,24"
		else
			'��Ÿ (��������,��ü,��ǰ,�귣��,���̾,�׽���,�űԵ����̳�,�����,�귣��Week)
			sKind = "1,5,12,13,16,17,23,24,19,25,26,31"
		end if
	end if
	strparm  = "isWeb="&blnWeb&"&isMobile="&blnMobile&"&isApp="&blnApp&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&selCM="&sCateMid&"&eventstate="&sState&"&eventkind="&sKind&"&sDgId="&edgid&"&sMdId="&emdid&"&spsid="&epsid&"&sdpid="&edpid&_
				"&sdgnm="&edgnm&"&smdnm="&emdnm&"&spsnm="&epsnm&"&sdpnm="&edpnm&"&ebrand="&ebrand&"&chSale="&esale&"&chGift="&egift&"&chCoupon="&ecoupon&"&chOnlyTen="&eonlyten&"&disp="&dispCate&"&chDiary="&eDiary&"&sDg2Id="&edgid2&"&sdg2nm="&edgnm2&"&designerstatus="&edgstat1&"&designerstatus="&edgstat2&"&eventtype_pc="&etype_pc&"&eventtype_mo="&etype_mo&"&elevel="&elevel
	'#######################################
 	if sSort = "" then sSort = "CD"
 	if blnReqPublish= "" then blnReqPublish = False     

 	if sEvt="evt_code" then
		if right(strTxt,1) = "," then strTxt = left(strTxt,len(strTxt)-1)
	end if

	'������ ��������
	set cEvtList = new ClsEvent
		cEvtList.FCPage = iCurrpage		'����������
		cEvtList.FPSize = iPageSize		'���������� ���̴� ���ڵ尹��

		cEvtList.FSfDate 	= sDate		'�Ⱓ �˻� ����
		cEvtList.FSsDate 	= sSdate	'�˻� ������
		cEvtList.FSeDate 	= sEdate	'�˻� ������
		cEvtList.FSfEvt 	= sEvt		'�˻� �̺�Ʈ�� or �̺�Ʈ�ڵ�
		cEvtList.FSeTxt 	= strTxt	'�˻���
		cEvtList.FScategory = sCategory	'�˻� ī�װ�
		cEvtList.FScateMid	= sCateMid	'�˻� ī�װ�(�ߺз�)
		cEvtList.FEDispCate	= dispCate	'�˻� ����ī�װ�
		cEvtList.FSstate 	= sState	'�˻� ����
	 
		cEvtList.FSedid   	= edgid
''		cEvtList.FSedid2   	= edgid2
		cEvtList.FSemid   	= emdid
		cEvtList.FSepsid   	= epsid
		cEvtList.FSedpid   	= edpid
		
		cEvtList.FSednm   	= edgnm
		cEvtList.FSemnm   	= emdnm
		cEvtList.FSepsnm   	= epsnm
		cEvtList.FSedpnm   	= edpnm
		
		cEvtList.FEDgStat1	= edgstat1
		cEvtList.FEDgStat2	= edgstat2
		
		
		cEvtList.FSkind 	= sKind
		cEvtList.FEBrand 	= ebrand
		cEvtList.FSisSale 	= esale
		cEvtList.FSisGift 	= egift
		cEvtList.FSisCoupon	= ecoupon
		cEvtList.FSisOnlyTen= eonlyten
		cEvtList.FSisDiary = eDiary
		cEvtList.FSisoneplusone   = eopo
		cEvtList.FSisfreedelivery = efd
		cEvtList.FSisbookingsell  = ebs
		cEvtList.FSisNew          = enew
	
		cEvtList.FIsWeb = blnWeb
		cEvtList.FIsMobile = blnMobile
		cEvtList.FIsApp = blnApp
		
		cEvtList.FRectEvtManager = eMng
		cEvtList.FRectEvtLevel = elevel
		cEvtList.FRectIsConfirm = isConfirm
		cEvtList.FRectendlessView = endlessView
		cEvtList.FIsReqPublish = blnReqPublish
		cEvtList.FSort          = sSort
		cEvtList.FRectMDTheme   = mdtheme
		cEvtList.FRectEventType_PC = etype_pc
		cEvtList.FRectEventType_MO = etype_mo
		cEvtList.FRectStartESP=startESP
		cEvtList.FRectEndESP=endESP
		cEvtList.FETemp=evt_template
		cEvtList.FETemp_mo=evt_template_mo
		cEvtList.FchComm=chComm
		cEvtList.FchItemps=chItemps
		cEvtList.FchBbs=chBbs
		cEvtList.Fisblogurl=isblogurl
 		arrList = cEvtList.fnGetEventList	'�����͸�� ��������
 		iTotCnt = cEvtList.FTotCnt	'��ü ������  ��
 	set cEvtList = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��

	Dim arreventlevel, arreventstate, arreventkind, arreventtype, arrdsnStat,arreventmanager
	'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
	arreventlevel = fnSetCommonCodeArr("eventlevel",False)
	arreventstate= fnSetCommonCodeArr("eventstate",False)
	arreventkind= fnSetCommonCodeArr("eventkind",False)
	arreventtype= fnSetCommonCodeArr("eventtype",False)
	arrdsnStat = fnSetCommonCodeArr("designerstatus",False)
	arreventmanager = fnSetCommonCodeArr("eventmanager",False)
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
<!-- 
	window.document.domain = "10x10.co.kr";
	function jsGoUrl(sUrl){
		self.location.href = sUrl;
	}

	function jsNewEvent(){
		window.open('about:blank').location.href = 'event_register.asp?menupos=<%=menupos%>&<%=strParm%>';
	}

	function jsSearch(sType){
		var frm = document.frmEvt
		if (sType == "A"){
				frm.iSD.value = "";
				frm.iED.value = "";
				frm.eventstate.value = "";
				frm.sEtxt.value = "";
				frm.selC.value = "";
		}
		if (frm.sEtxt.value!=""){
			if (frm.selEvt.value=="evt_code"){
				if (!IsDouble(frm.sEtxt.value)){
					alert('�̺�Ʈ�ڵ�� ���ڸ� �����մϴ�.');
					frm.sEtxt.focus();
					return;
				}
			}
		}
		frm.action="";
		frm.submit();
	}

	function SubmitForm() {
		jsSearch('E');
	}

	function jsSchedule(){
		var winS;
		winS = window.open('/admin/eventmanage/event/pop_event_schedule.asp','popwin','width=1200, height=800, scrollbars=yes');
		winS.focus();
	}




	function jsCodeManage(){
		var winCode;
		winCode = window.open('/admin/eventmanage/code/popManageCode.asp','popCode','width=450,height=600,scrollbars=yes,resizable=yes');
		winCode.focus();
	}

	function jsDivisionCodeManage(){
		var winCode;
		winCode = window.open('/admin/eventmanage/event/v5/popup/pop_division_Manage.asp','popCode','width=800,height=700,scrollbars=yes,resizable=yes');
		winCode.focus();
	}

	function prize(evt_code){

		 var prize = window.open('/admin/eventmanage/event/pop_event_prize.asp?evt_code='+evt_code,'prize','width=800,height=600,scrollbars=yes,resizable=yes');
		 prize.focus();

	}
	
	function jsGetID(sType, iCid, sUserID){
		var openWorker = window.open('/admin/eventmanage/event/v5/popup/PopWorkerList.asp?sType='+sType+'&department_id='+iCid+'&sUserid='+sUserID,'openWorker','width=350,height=570,scrollbars=yes');
		openWorker.focus();
	}
	
	function jsDelID(sType){ 
		eval("document.frmEvt.s"+sType+"Id").value = "";
		eval("document.frmEvt.s"+sType+"Nm").value = ""; 
	}
	
	 //����Ʈ ����
	 function jsSort(sValue,i){  
	  
	 	document.frmEvt.sSort.value= sValue; 
	 	 
		   if (-1 < eval("document.all.img"+i).src.indexOf("_alpha")){
	        document.frmEvt.sSort.value= sValue+"D";  
	    }else if (-1 < eval("document.all.img"+i).src.indexOf("_bot")){
	     		document.frmEvt.sSort.value= sValue+"A";  
	    }else{
	       document.frmEvt.sSort.value= sValue+"D";  
	    } 
	    
	   
		 document.frmEvt.submit();
	}
	
	//��¥ ����
	function jsSetDate(iValue){
	    var currentDate = new Date(); 

        var month = currentDate.getMonth() + 1;
        var day = currentDate.getDate();
        var year = currentDate.getFullYear();
 
        var preDate = new Date(currentDate.setMonth(month-iValue)); 
        var pmonth = preDate.getMonth() ; 
        var pday = preDate.getDate();
        var pyear = preDate.getFullYear();
        
        if (month <10){
            month = "0"+month;
        }
        
         if (pmonth <10){
            pmonth = "0"+pmonth;
        }
        
         if (pday <10){
            pday = "0"+pday;
        }
        
         if (day <10){
            day = "0"+day;
        }
        
 	    document.frmEvt.iSD.value = pyear+"-"+pmonth+"-"+pday; 
	    document.frmEvt.iED.value = year+"-"+month+"-"+day;
	}

	//
	function makeThumbBanTxt(evtcode,oldview){
	    var popwinThumbTxt = window.open('http://110.93.128.113/pSvr/makeCateEvtBanner.asp?evtcode='+evtcode+'&oldview='+oldview,'popwinThumbTxt','width=680,height=570,scrollbars=yes');
		popwinThumbTxt.focus();
	}

	//�̸�����
	function jsOpen(sPURL,sTG){ 
	    if (sTG =="M" ){ 
	        var winView = window.open(sPURL,"popView","width=400, height=600,scrollbars=yes,resizable=yes,location=yes");
	    }
	}

	//20181105 ��Ƽ3�� ������
	function pop_multi3_manage(eCode){	
		var multi3Window = window.open('/admin/eventmanage/event/V4/pop_manage_multi3.asp?evt_code='+eCode,'multi3Window','width=700, height=900,scrollbars=yes,resizable=yes');
		multi3Window.focus();
	}

	function jsExcelDown(iCurrpage){
		document.frmEvt.target="_blank";
		document.frmEvt.iC.value=iCurrpage;
		document.frmEvt.action="/admin/eventmanage/event/V5/new_event_excel.asp";
		document.frmEvt.submit();
	}

	//�귣�� ID �˻� �˾�â
	function jsSearchBrandIDNew(frmName,compName){
		var compVal = "";
		try{
			compVal = eval("document.all." + frmName + "." + compName).value;
		}catch(e){
			compVal = "";
		}

		var popwin = window.open("/admin/member/popBrandSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal + "&isjsdomain=o","popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

		popwin.focus();
	}

	function fnWorkerInfoSet(eventcode){
		var winworkerView;
		winworkerView = window.open('/admin/eventmanage/event/v5/popup/pop_event_workerinfo.asp?eC='+eventcode+'&fromlist=Y','workerinfo','width=1024,height=800,scrollbars=yes,resizable=yes');
		winworkerView.focus();
	}

	function show_subscript(){
        let subscript_popup = window.open('/admin/eventmanage/event/v5/popup/subscript.asp','popwin','width=1200, height=800, scrollbars=yes');
        subscript_popup.focus();
	}
</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style>
	select {font-size:12px; vertical-align:top;}
	input[type=button], input[type=text] {vertical-align:top;}
</style>
<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmEvt" method="get"  action="index.asp" onSubmit="return jsSearch('E');">
	<input type="hidden" name="menupos" value="<%=menupos%>"> 
	<input type="hidden" name="isResearch" value="1"> 
	<input type="hidden" name="sSort" value="<%=sSort%>">
	<input type="hidden" name="iC">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
			<tr>
				<td>�귣��:</td>
				<td><% NewDrawSelectBoxDesignerwithNameEvent "ebrand", ebrand %></td>
				<td colspan="4">
					���� <!-- #include virtual="/common/module/categoryselectbox_event.asp"-->
					/ ���� ī�װ� : <!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
				</td>
			</tr>
			<tr> 
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">�̺�Ʈ����:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;"><%sbGetNewOptCommonCodeArr "eventkind", sKind, True, True, False,"onChange='javascript:document.frmEvt.submit();'"%></td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">�ڵ�/��:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;"><select name="selEvt">
			    	<option value="evt_code" <%if Cstr(sEvt) = "evt_code" THEN %>selected<%END IF%>>�̺�Ʈ�ڵ�</option>
			    	<option value="evt_name" <%if Cstr(sEvt) = "evt_name" THEN %>selected<%END IF%>>�̺�Ʈ��</option>
			    	<option value="evt_tag" <%if Cstr(sEvt) = "evt_tag" THEN %>selected<%END IF%>>TAG</option>
			    	<option value="evt_sub" <%if Cstr(sEvt) = "evt_sub" THEN %>selected<%END IF%>>����ī��</option>
			    	</select>
			        <input type="text" name="sEtxt" value="<%=strTxt%>" maxlength="256" onkeydown="if(event.keyCode==13) document.frmEvt.submit();" /></td>
			    <td style="border-top:1px solid <%= adminColor("tablebg") %>;">�̺�ƮŸ��:</td>
			    <td style="border-top:1px solid <%= adminColor("tablebg") %>;" >
			    	<input type="checkbox" name="chSale" <%IF Cstr(esale)="1" THEN%> checked <%END IF%>  value="1">����
					<input type="checkbox" name="chCoupon" <%IF Cstr(ecoupon)="1" THEN%> checked<%END IF%> value="1">���� 
					<input type="checkbox" name="chOnlyTen" <%IF Cstr(eonlyten)="1" THEN%> checked<%END IF%> value="1">Only-TenByTen 
					<input type="checkbox" name="chopo" <%IF Cstr(eopo)="1" THEN%> checked<%END IF%> value="1">1+1  
			    </td> 
			</tr>
			<tr>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">��ȹ������:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
					<select name="evt_template" class="select">
						<option value="">����</option>
						<option value="10"<% if evt_template="10" then response.write " selected" %>>���ø� ���</option>
						<option value="6"<% if evt_template="6" then response.write " selected" %>>���۾� ���</option>
					</select>
					<select name="evt_template_mo" class="select">
						<option value="">����</option>
						<option value="11"<% if evt_template_mo="11" then response.write " selected" %>>���ø� ���</option>
						<option value="6"<% if evt_template_mo="6" then response.write " selected" %>>���۾� ���</option>
						<option value="10"<% if evt_template_mo="10" then response.write " selected" %>>Multi3��</option>
					</select>
				</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">�߿䵵 : </td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
					<select name="elevel" class="select">
						<option value="">����</option>
						<option value="1"<% if elevel="1" then response.write " selected" %>>�ֻ�</option>
						<option value="2"<% if elevel="2" then response.write " selected" %>>��</option>
						<option value="3"<% if elevel="3" then response.write " selected" %>>��</option>
					</select>
				</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">&nbsp;</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;"> 
					<input type="checkbox" name="chnew" <%IF Cstr(enew)="1" THEN%> checked<%END IF%> value="1">��Ī
					<input type="checkbox" name="chfd" <%IF Cstr(efd)="1" THEN%> checked<%END IF%> value="1">������
					<input type="checkbox" name="chbs" <%IF Cstr(ebs)="1" THEN%> checked<%END IF%> value="1">�����Ǹ� 
					<input type="checkbox" name="chDiary" <%IF Cstr(eDiary)="1" THEN%> checked<%END IF%> value="1">DiaryStory
				</td>
			</tr>
			<tr>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">�����̳� �۾� ����:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
					<select name="eventtype_pc" class="select">
						<option value="">����</option>
						<option value="0"<% if etype_pc="0" then response.write " selected" %>>MD��</option>
						<option value="20"<% if etype_pc="20" then response.write " selected" %>>��������</option>
					</select>
					<select name="eventtype_mo" class="select">
						<option value="">����</option>
						<option value="0"<% if etype_mo="0" then response.write " selected" %>>MD��</option>
						<option value="20"<% if etype_mo="20" then response.write " selected" %>>��������(Ǯ)</option>
						<option value="50"<% if etype_mo="50" then response.write " selected" %>>��������(���̵�)</option>
					</select>
					<input type="checkbox" name="blnCnfm" <%=chkIIF(Cstr(isConfirm)="1","checked","")%> value="1">���οϷ�
					
				  </td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">�Ⱓ:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;"><select name="selDate">
			    	<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>������ ����</option>
			    	<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>������ ����</option>
			    	<option value="O" <%if Cstr(sDate) = "O" THEN %>selected<%END IF%>>������ ����</option>
			        </select>
			         <input id="iSD" name="iSD" value="<%=sSdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
			        <input id="iED" name="iED" value="<%=sEdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
					
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "iSD", trigger    : "iSD_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_End = new Calendar({
							inputField : "iED", trigger    : "iED_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
					
			        <!--<input type="text" id="iSD" name="iSD" value="<%=sSdate%>" class="text" size="10" maxlength="10" style="vertical-align:top" />
			        <!-- <input type="image" src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" style="vertical-align:top" /> ~
			        <input type="text" id="iED" name="iED" value="<%=sEdate%>" class="text" size="10" maxlength="10" style="vertical-align:top" /> <input type="image" src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" style="vertical-align:top" />
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "iSD", trigger    : "iSD_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_End = new Calendar({
							inputField : "iED", trigger    : "iED_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>-->
					&nbsp;
					<input type="button" value="�ֱ� 1��" class="button" onClick="jsSetDate(1)">
					<input type="button" value="�ֱ� 3��" class="button" onClick="jsSetDate(3)">
					<input type="checkbox" name="endlessView" value="Y" <%IF endlessView="Y" THEN%> checked<%END IF%>> ��ó���
			    </td>
				<td colspan="2" style="border-top:1px solid <%= adminColor("tablebg") %>;">&nbsp;</td>
			</tr>
			<tr>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">�������:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
				    <select name="eventstate" class="select" onChange="SubmitForm();">
				        <option value="">����</option>
                        	<option value="0"<% if sState="0" then response.write " selected" %>> ��ϴ��</option>
                        	<option value="11"<% if sState="11" then response.write " selected" %>> ���ø� �۾���</option>
                        	<option value="3"<% if sState="3" then response.write " selected" %>> �̹�����Ͽ�û</option>
                        	<option value="1"<% if sState="1" then response.write " selected" %>> �����̳� �۾���</option>
                        	<option value="4"<% if sState="4" then response.write " selected" %>> �ۺ��� ��û</option>
                        	<option value="12"<% if sState="12" then response.write " selected" %>> �ۺ��� �۾���</option>
                        	<option value="2"<% if sState="2" then response.write " selected" %>> ���� ��û</option>
                        	<option value="13"<% if sState="13" then response.write " selected" %>> ���� �۾���</option>
                        	<option value="10"<% if sState="10" then response.write " selected" %>> �̺�Ʈ���߿�û</option>
                        	<option value="7"<% if sState="7" then response.write " selected" %>> ���¿���</option>
							<option value="5"<% if sState="5" then response.write " selected" %>> ���¿�û</option>
                        	<option value="6"<% if sState="6" then response.write " selected" %>> ����</option>
                        	<option value="9"<% if sState="9" then response.write " selected" %>> ����</option>
				    </select>
				</td>
			    <td style="border-top:1px solid <%= adminColor("tablebg") %>;">�����:</td>
			    <td style="border-top:1px solid <%= adminColor("tablebg") %>;" colspan="3">
			    	<span style="white-space:nowrap;">��ȹ��  <input type="hidden" name="sMdId" value="<%=emdid%>"><input type="name" name="sMdNm" value="<%=eMDnm%>" class="text"  size="10">&nbsp;<input type="button" class="button" value="����" onClick="jsGetID('Md','162','<%=emdid%>');"> <input type="button" value="&times"  class="button" onClick="jsDelID('Md');" title="����� �����" /></span> &nbsp;&nbsp;
			    	<span style="white-space:nowrap;">�����̳� <input type="hidden" name="sDgId" value="<%=edgid%>"><input type="name" name="sDgNm" value="<%=edgnm%>" class="text"  size="10">&nbsp;<input type="button" class="button" value="����" onClick="jsGetID('Dg','152','<%=edgid%>');">&nbsp;<input type="button" value="&times"  class="button" onClick="jsDelID('Dg');" title="����� �����" /></span> &nbsp;
			    	<span style="white-space:nowrap;">�ۺ���  <input type="hidden" name="sPsId" value="<%=epsid%>"><input type="name" name="sPsNm" value="<%=epsnm%>"class="text"  size="10">&nbsp;<input type="button" class="button" value="����"  onClick="jsGetID('Ps','157','<%=epsid%>');"> <input type="button" value="&times"  class="button" onClick="jsDelID('Ps');" title="����� �����" /></span> &nbsp;&nbsp;
			    	<span style="white-space:nowrap;">������  <input type="hidden" name="sDpId" value="<%=edpid%>"><input type="name" name="sDpNm" value="<%=edpnm%>" class="text"  size="10">&nbsp;<input type="button" class="button" value="����" onClick="jsGetID('Dp','130','<%=edpid%>');"> <input type="button" value="&times"  class="button" onClick="jsDelID('Dp');" title="����� �����" /></span>
			    </td>
			</tr> 
			<tr> 
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">ä��:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
					<input type="checkbox" name="isMobile"  value="1" <%if blnMobile="1" then%>checked<%end if%>> Mobile
					<input type="checkbox" name="isApp"  value="1" <%if blnApp="1" then%>checked<%end if%>> App
					<input type="checkbox" name="isWeb" value="1" <%if blnWeb="1" then%>checked<%end if%>> PC-Web
				</td>
			    <tD colspan="4" style="border-top:1px solid <%= adminColor("tablebg") %>;"><input type="checkbox" name="chkPus" value="1" <%if blnReqPublish THEN%>checked<%end if%>> �ۺ��� ��û�۾�</td>
			</tr>
			<tr> 
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">�̺�Ʈ��ü:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;" colspan="3"><%sbGetOptCommonCodeArr "eventmanager", eMng, True,True,"onChange='document.frmEvt.submit();'"%> &nbsp;
				�������� : <input id="startESP" name="startESP" value="<%=startESP%>" class="text" size="10" maxlength="10" />
				~ <input id="endESP" name="endESP" value="<%=endESP%>" class="text" size="10" maxlength="10" />
				&nbsp;&nbsp;������� : 
				<input type="checkbox" name="chComm"  value="1" <%if chComm="1" then%>checked<%end if%>> �ڸ�Ʈ
				<input type="checkbox" name="chItemps"  value="1" <%if chItemps="1" then%>checked<%end if%>> ��ǰ�ı�
				<input type="checkbox" name="chBbs" value="1" <%if chBbs="1" then%>checked<%end if%>> �����ڸ�Ʈ
				<input type="checkbox" name="isblogurl" value="1" <%if isblogurl="1" then%>checked<%end if%>> blog URL
				</td>
			</tr>
			</table>
        </td>
    		<td  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSearch('E');">
		</td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
    <tr height="40" valign="bottom">
        <td align="left">
        	<input type="button" value="���ε��" onclick="jsNewEvent();" class="button">
	    </td>
	    <td align="right">
	        <input type="button" value="������ Ȯ��" onclick="show_subscript();"  class="button">
	       	<% if iTotCnt>2000 then %>
			   <input type="button" value="����Ʈ�����ٿ�" onclick="alert('2000�� ���Ϸ� �˻��� �ּ���.');"  class="button">
			<% else %>
				<input type="button" value="����Ʈ�����ٿ�" onclick="jsExcelDown(<%=iCurrpage%>);"  class="button">
			<% end if %>
	       	<input type="button" value="������" onclick="jsSchedule();"  class="button">
	       <!--	<input type="button" value="���" onclick=" ">  -->
		   <% if C_ADMIN_AUTH then %><input type="button" value="�����ڵ����" onclick="jsDivisionCodeManage();"  class="button"><%END IF%>
	       <% if C_ADMIN_AUTH then %><input type="button" value="�ڵ����" onclick="jsCodeManage();"  class="button"><%END IF%>
        </td>
	</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="22">�˻���� : <b><%=iTotCnt%></b>&nbsp;&nbsp;������ : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td nowrap rowspan="2">ä��</td>
    	<td nowrap rowspan="2">��ü</td>
    	<td nowrap rowspan="2">�̺�Ʈ����</td>
    	<td nowrap rowspan="2">�̺�Ʈ����</td>
    	<td nowrap rowspan="2" onClick="javascript:jsSort('C','1');" style="cursor:hand;"><b>�̺�Ʈ�ڵ�</b><img src="/images/list_lineup<%IF sSort="CD" THEN%>_bot<%ELSEIF sSort="CA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img1"></td>
    	<td nowrap rowspan="2" onClick="javascript:jsSort('S','2');" style="cursor:hand;"><b>�߿䵵</b> <img src="/images/list_lineup<%IF sSort="SD" THEN%>_bot<%ELSEIF sSort="SA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img2"></td>
      	<td nowrap rowspan="2">�������</td>
      	<td nowrap rowspan="2">���</td>
      	<td nowrap rowspan="2">���̵���</td>
      	<td nowrap rowspan="2">�̺�Ʈ��</td>
      	<td nowrap rowspan="2">ī�װ�</td>
      	<td nowrap rowspan="2">�귣��</td>
      	<td width="60" rowspan="2" onClick="javascript:jsSort('D','3');" style="cursor:hand;"><b>������</b> <img src="/images/list_lineup<%IF sSort="DD" THEN%>_bot<%ELSEIF sSort="DA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img3"></td>
      	<td width="60" rowspan="2">������</td>
      	<td rowspan="2"  onClick="javascript:jsSort('I','4');" style="cursor:hand;"><b>�̹�����û��</b> <img src="/images/list_lineup<%IF sSort="ID" THEN%>_bot<%ELSEIF sSort="IA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img4"></td>
      	<td nowrap colspan="4">�����</td>
      	<td nowrap rowspan="2">����</td>
     </tr>
     <tr align="center" bgcolor="<%= adminColor("tabletop") %>">	 
        <td nowrap>��ȹ��</td>
      	<td nowrap>�����̳�</td>
      	<td nowrap>�ۺ���</td>
      	<td nowrap>������<br />/ �˼���</td>
    </tr>

    <%IF isArray(arrList) THEN
		Dim itemSortvalue
		Dim strURL
		Dim isMobile, isApp, isWeb
		 dim tmpename, ename,eSalePer
		
    	For intLoop = 0 To UBound(arrList,2) 
		
		'2014-08-27 ������ / ������ ������ ����
		Select Case arrList(27,intLoop)
			Case "1"	itemSortvalue = "sitemid"
			Case "2"	itemSortvalue = "slsell"
			Case "3"	itemSortvalue = "sevtitem"
			Case "4"	itemSortvalue = "sbest"
			Case "5"	itemSortvalue = "shsell"
		End Select
		
		isWeb = False
		isMobile = False
		isApp = False
		
		IF isNull(arrList(30,intLoop)) and isNull(arrList(31,intLoop)) and isNull(arrList(32,intLoop)) then
			if arrList(1,intLoop) = "19" THEN
				isWeb = False
				isMobile = True
				isApp = True
			ELSEIF arrList(1,intLoop) = "25"  THEN
				isWeb = False
				isMobile = False
				isApp = True
			ELSEIF arrList(1,intLoop) = "26"  THEN	
				isWeb = False
				isMobile = True
				isApp = False
			ELSE
				isWeb = True
				isMobile = False
				isApp = False	
			END IF
		END IF	
		IF 	 not isNull(arrList(30,intLoop))  THEN	
			isWeb = arrList(30,intLoop)
		END IF	
		IF 	 not isNull(arrList(31,intLoop)) THEN
			 isMobile = arrList(31,intLoop)
		END IF	 
		IF 	 not isNull(arrList(32,intLoop)) THEN
			isApp = arrList(32,intLoop)	
		END IF	
		
		 
    %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>
    	    <% 	
			dim sMoblie,sWeb,sApp
			 sApp = ""
			 sMoblie=""
			Select Case arrList(1,intLoop)
				Case "7"		'��Ŭ���ڵ�
					 sWeb = "<a href='" & vwwwUrl & "/guidebook/weekly_coordinator.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>"  
				Case "13"		'��ǰ �̺�Ʈ
					sWeb =  "<a href='" & vwwwUrl & "/shopping/category_prd.asp?itemid=" & arrList(21,intLoop) & "' target='_blank'>"  
					sMoblie =  "<a href= ""javascript:jsOpen('" & vmobileUrl & "/category/category_itemPrd.asp?itemid=" & arrList(21,intLoop) & "','M');"">" 
				Case "14"		'��ǳ���±�
					sWeb =  "<a href='" & vwwwUrl & "/guidebook/picnic/picnic.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>"  
				Case "5"		'���Ľ����̼�
					sWeb =  "<a href='" & vwwwUrl & "/culturestation/culturestation_event.asp?evt_code=" & arrList(0,intLoop) & "' target='_blank'>"  
					sMoblie =  "<a href='" & mobileUrl & "/culturestation/culturestation_event.asp?evt_code=" & arrList(0,intLoop) & "' target='_blank'>"  
				Case "16"		'�귣�� �������
					sWeb =  "<a href='" & vwwwUrl & "/street/street_brand_sub06.asp?makerid=" & arrList(14,intLoop) & "&shop_event_yn=Y&shop_event_confirm=Y&shopview=3' target='_blank'>"  
				Case "22"		'DAY&(���̾ص�)
					sWeb = "<a href='" & vwwwUrl & "/guidebook/dayand.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>"   
				Case "26"		'�����
					sWeb =  "<a href='" & vwwwUrl & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>"  
					sMoblie =  "<a href= ""javascript:jsOpen('" & vmobileUrl & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "','M');"">"
				Case "29"		'���̽��
					sWeb =  "<a href='" & vwwwUrl & "/HSProject/?eventid=" & arrList(0,intLoop) & "' target='_blank'>"  
					sMoblie =  "<a href= ""javascript:jsOpen('" & vmobileUrl & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "','M');"">"
				Case Else		'�������� �� ��Ÿ
					sWeb = "<a href='" & vwwwUrl & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>"  
					sMoblie = "<a href= ""javascript:jsOpen('" & vmobileUrl & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "','M');"">"  
					sApp ="<a href= ""javascript:jsOpen('" & vmobileUrl & "/apps/appCom/wish/web2014/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "','M');"">"
			End Select 
		%>
    	    <%IF isWeb THEN %> <%=sWeb%>Web</a><%END IF%>
    	    <%=chkIIF(isMobile,"<br />" & sMoblie & "<font color=""blue"">Mobile</font></a>","")%>
    	    <%=chkIIF(isApp,"<br />" & sApp & "<font color=""red"">App</font></a>","")%>
    	</td>
    	<td><%=fnGetCommCodeArrDesc(arreventmanager,arrList(2,intLoop))%></td>
    	<td><%=fnGetCommCodeArrDesc(arreventkind,arrList(1,intLoop))%></td>
    	<td>
			<% If arrList(44,intLoop)="50" Then %>
				PC : ��������(���̵�)
			<% elseIf arrList(44,intLoop)="20" Then %>
				PC : ��������(Ǯ)
			<% else %>
				PC : MD��
			<% End If %><br>
			<% If arrList(45,intLoop)="20" Then %>
				MO : ��������
			<% else %>
				MO : MD��
			<% End If %>
		</td>
		<td><a href="event_register.asp?eC=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a></td>
    	<td><%=fnGetCommCodeArrDesc(arreventlevel,arrList(7,intLoop))%></td>
      	<td>
		  	<% if arrList(54,intLoop)="Y" then %>
			  	��ó���
			<% else %>
				<% if arrList(8,intLoop) = "6" or arrList(8,intLoop) = "7" then %>
					<% if arrList(6,intLoop) < now() then %>
						����
					<% else %>
						<%=fnGetCommCodeArrDesc(arreventstate,arrList(8,intLoop))%>
					<% end if %>
				<% else %>
					<%=fnGetCommCodeArrDesc(arreventstate,arrList(8,intLoop))%>
				<% end if %>
			<% end if %>
			
		</td>
      	<td><% if arrList(1,intLoop)="5" then %><img src="<%=arrList(53,intLoop)%>" width="100" border="0"><% else %><%IF arrList(34,intLoop) <> "" THEN%> <img src="<%=arrList(34,intLoop)%>" width="100" border="0"><%END IF%><%END IF%></td>
        <td><%IF arrList(35,intLoop) <> "" THEN%> <img src="<%=arrList(35,intLoop)%>" width="100" border="0"><%END IF%></td>  
      	<td align="left">
      		<%=chkIIF(Not(arrList(25,intLoop)="" or isNull(arrList(25,intLoop))),"["&arrList(25,intLoop)&"] ","")%>
      		<%   ename =  arrList(4,intLoop)  
      		     eSalePer = ""
      		    if  (arrList(15,intLoop) or arrList(17,intLoop)) then 
            	    tmpename = Split(ename,"|")  
            	  	if Ubound(tmpename)>0 then
            		    ename = tmpename(0)
            		    eSalePer = tmpename(1)
            		 end if
            
                end if
             %>   
      		<%=db2html(ename)%>
      		<% if arrList(15,intLoop)  then%>
      		<font color="red"><%=eSalePer%></font>
      		<% elseif arrList(17,intLoop) then%>
      		<font color="green"><%=eSalePer%></font>
      		<% end if%>
      		<%if arrList(15,intLoop)  then%>&nbsp;<img src="http://fiximage.10x10.co.kr/web2008/category/icon_sale.gif" border="0"><%end if%>
      		<%if arrList(16,intLoop) then%>&nbsp;<img src="http://fiximage.10x10.co.kr/web2008/category/icon_gift.gif" border="0"><%end if%>
      		<%if arrList(17,intLoop) then%>&nbsp;<img src="http://fiximage.10x10.co.kr/web2008/category/icon_coupon.gif" border="0"><%end if%>
      	</td>
      	<td>
      		<%=arrList(12,intLoop)%>
      		<%
      		if arrList(22,intLoop) <> "" then
      			response.write "(" & arrList(22,intLoop) &")"
      		end if
      		'����ī�װ�
      		if arrList(26,intLoop)<>"" then
      			response.write chkIIF(arrList(12,intLoop)<>"","<br/>","") & "<font color='#4030A0'>" & arrList(26,intLoop) & "</font>"
      		end if
      		%>
      	</td>
      	<td><%=arrList(14,intLoop)%></td>
      	<td><%=FormatDate(arrList(5,intLoop),"0000-00-00")%></td>
      	<td><%=FormatDate(arrList(6,intLoop),"0000-00-00")%></td> 
      	<td><%=arrList(36,intLoop)%></td>
      	<td><%=arrList(23,intLoop)%></td>
      	<td><a href="javascript:fnWorkerInfoSet(<%=arrList(0,intLoop)%>);"><%=arrList(11,intLoop)%><% if (arrList(11,intLoop)="" or isnull(arrList(11,intLoop))) and arrList(8,intLoop)="3" then %><span style="color:#B88;">�����̳� ����</span><% end if %></a></td>
      	<td><a href="javascript:fnWorkerInfoSet(<%=arrList(0,intLoop)%>);"><%=arrList(28,intLoop)%><% if (arrList(28,intLoop)="" or isnull(arrList(28,intLoop))) and arrList(8,intLoop)="4" then %><span style="color:#B88;">�ۺ��� ����</span><% end if %></a></td>
      	<td><%=arrList(29,intLoop)%><%=chkiif(arrList(38,intLoop)<>"","<br />" & arrList(38,intLoop),"")%></td>
		<% if arrList(39,intLoop) = 90 then '��Ƽ3��%>		  
			<td align="left" nowrap>		  
			<input type="button" value="�̺�Ʈ����" class="button" onClick="javascript:pop_multi3_manage(<%=arrList(0,intLoop)%>);">      		
			</td>				
		<% else %>      	
			<td align="left" nowrap><input type="button" value="��ǰ" class="button" onClick="javascript:jsGoUrl('/admin/eventmanage/event/v5/popup/eventitem_regist.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>&selsort=<%=itemSortvalue%>')">
				<%IF arrList(13,intLoop) > "1900-01-01" THEN%><input type="button" value="��÷" class="button" onClick="jsGoUrl('/admin/eventmanage/event/eventprize_regist.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>')"><%END IF%>
				<%if arrList(15,intLoop)  then%> <input type="button" value="����(<%=arrList(18,intLoop)%>)" class="button" onClick="jsGoUrl('/admin/shopmaster/sale/salelist.asp?eC=<%=arrList(0,intLoop)%>&menupos=290');"><%end if%>
				<%if arrList(16,intLoop) then%> <input type="button" value="����ǰ(<%=arrList(19,intLoop)%>)" class="button" onClick="jsGoUrl('/admin/shopmaster/gift/giftlist.asp?eC=<%=arrList(0,intLoop)%>&menupos=1045');"><%end if%>
				<!--<%if arrList(17,intLoop) then%> <input type="button" value="����" class="button" onClick="jsGoUrl('coupon');"><%end if%>	-->
				<% If arrList(20,intLoop) = "N" Then %>
				<table cellpadding="0" cellspacing="0" border="0"><tr><td style="padding:3 0 0 0;"><input type="button" class="button" style="width:105;" value="��÷�ھ��� ����" onclick="prize(<%= arrList(0,intLoop) %>);"></td></tr></table>
				<% End IF %>
			</td>
		<% end if %>		  
    </tr>
   <%	Next
   	ELSE
   %>
   	<tr  align="center" bgcolor="#FFFFFF">
   		<td colspan="21">��ϵ� ������ �����ϴ�.</td>
   	</tr>
   <%END IF%>
</table>
 </form>
<!-- ����¡ó�� --> 
<table width="100%" cellpadding="10">
	<tr>
		<td align="center">  
 			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->