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
<!-- #include virtual="/lib/classes/event/eventManageCls_V2.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
	response.write "<script type='text/javascript'>"
	response.write "	alert('���Ұ� ������');history.back();"
	response.write "</script>"
	response.End
	Call fnSetEventCommonCode '�����ڵ� ���ø����̼� ������ ����

	'��������
	Dim cEvtList
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
	Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory, sCateMid ,sState,sKind,esale,egift,ecoupon,ebrand,eonlyten,etype,isConfirm
	Dim strparm
	Dim edgid, edgid2,edgstat1,edgstat2, emdid, epsid, edpid, edgnm, edgnm2, emdnm, epsnm, edpnm, eDiary
	dim eopo,efd,ebs,enew
	dim blnWeb, blnMobile, blnApp
	dim dispCate, maxDepth
	dim blnReqPublish ,sSort
	dim isResearch

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
	strTxt 		= requestCheckVar(Request("sEtxt"),60)

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

	etype		= requestCheckvar(request("eventtype"),4)
	isConfirm	= requestCheckvar(request("blnCnfm"),1)
	
	if isResearch="0" and sKind="" then
		skind="1,12,13,23,27,28,29,31"
	end if


	'�̺�Ʈ ù������ �����׸��� ���̵��� 
	IF (sKind="" and isResearch="0") or sKind="1,12" THEN
		if (session("ssAdminPsn")="11" or session("ssAdminPsn")="21") and (not ( session("ssBctId")="fotoark" or session("ssBctId")="arlejk" or session("ssBctId")="barbie8711")) then
			'MD�μ���� (��������,��ü,��ǰ,�귣��,���̾,�׽���,�űԵ����̳�) - ���̷�(fotoark), ���ְ�(arlejk), ����ȭ(barbie8711) ����
			sKind = "1,12,13,16,17,23,24"
		else
			'��Ÿ (��������,��ü,��ǰ,�귣��,���̾,�׽���,�űԵ����̳�,�����,�귣��Week)
			sKind = "1,12,13,16,17,23,24,19,25,26,31"
		end if
	end if
	strparm  = "isWeb="&blnWeb&"&isMobile="&blnMobile&"&isApp="&blnApp&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&selCM="&sCateMid&"&eventstate="&sState&"&eventkind="&sKind&"&sDgId="&edgid&"&sMdId="&emdid&"&spsid="&epsid&"&sdpid="&edpid&_
				"&sdgnm="&edgnm&"&smdnm="&emdnm&"&spsnm="&epsnm&"&sdpnm="&edpnm&"&ebrand="&ebrand&"&chSale="&esale&"&chGift="&egift&"&chCoupon="&ecoupon&"&chOnlyTen="&eonlyten&"&disp="&dispCate&"&chDiary="&eDiary&"&sDg2Id="&edgid2&"&sdg2nm="&edgnm2&"&designerstatus="&edgstat1&"&designerstatus="&edgstat2
	'#######################################
 	if sSort = "" then sSort = "CD"
 	if blnReqPublish= "" then blnReqPublish = False     
 	    
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
'		cEvtList.FSedid2   	= edgid2
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
		
		cEvtList.FRectEvtType = etype
		cEvtList.FRectIsConfirm = isConfirm
		
		cEvtList.FIsReqPublish = blnReqPublish
		cEvtList.FSort          = sSort
 		arrList = cEvtList.fnGetEventList	'�����͸�� ��������
 		iTotCnt = cEvtList.FTotCnt	'��ü ������  ��
 	set cEvtList = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��

	Dim arreventlevel, arreventstate, arreventkind, arreventtype, arrdsnStat
	'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
	arreventlevel = fnSetCommonCodeArr("eventlevel",False)
	arreventstate= fnSetCommonCodeArr("eventstate",False)
	arreventkind= fnSetCommonCodeArr("eventkind",False)
	arreventtype= fnSetCommonCodeArr("eventtype",False)
	arrdsnStat = fnSetCommonCodeArr("designerstatus",False)
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
<!-- 
	function jsGoUrl(sUrl){
		self.location.href = sUrl;
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

		if(frm.selEvt.value== "evt_code"&&frm.sEtxt.value!=""){
			frm.sEtxt.value = frm.sEtxt.value.replace(/\s/g, "");
			if(!IsDigit(frm.sEtxt.value)){
				alert("�̺�Ʈ�ڵ�� ���ڸ� �����մϴ�.");
				frm.sEtxt.focus();
				return;
			}
		}
 
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

	function prize(evt_code){

		 var prize = window.open('/admin/eventmanage/event/pop_event_prize.asp?evt_code='+evt_code,'prize','width=800,height=600,scrollbars=yes,resizable=yes');
		 prize.focus();

	}
	
	function jsGetID(sType, iCid, sUserID){
		var openWorker = window.open('PopWorkerList.asp?sType='+sType+'&department_id='+iCid+'&sUserid='+sUserID,'openWorker','width=350,height=570,scrollbars=yes');
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
//-->
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
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
			<tr>
				<td>�귣��:</td>
				<td><% drawSelectBoxDesignerwithName "ebrand", ebrand %></td>
				<td colspan="4">
					���� <!-- #include virtual="/common/module/categoryselectbox_event.asp"-->
					/ ���� ī�װ� : <!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
				</td>
			</tr>
			<tr> 
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">�̺�Ʈ����:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;"><%sbGetOptCommonCodeArr "eventkind", sKind, True,True,"onChange='javascript:document.frmEvt.submit();'"%></td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">�ڵ�/��:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;"><select name="selEvt">
			    	<option value="evt_code" <%if Cstr(sEvt) = "evt_code" THEN %>selected<%END IF%>>�̺�Ʈ�ڵ�</option>
			    	<option value="evt_name" <%if Cstr(sEvt) = "evt_name" THEN %>selected<%END IF%>>�̺�Ʈ��</option>
			    	<option value="evt_tag" <%if Cstr(sEvt) = "evt_tag" THEN %>selected<%END IF%>>TAG</option>
			    	<option value="evt_sub" <%if Cstr(sEvt) = "evt_sub" THEN %>selected<%END IF%>>����ī��</option>
			    	</select>
			        <input type="text" name="sEtxt" value="<%=strTxt%>" maxlength="60" onkeydown="if(event.keyCode==13) document.frmEvt.submit();" /></td>
			    <td style="border-top:1px solid <%= adminColor("tablebg") %>;">�̺�ƮŸ��:</td>
			    <td style="border-top:1px solid <%= adminColor("tablebg") %>;" >
			    	<input type="checkbox" name="chSale" <%IF Cstr(esale)="1" THEN%> checked <%END IF%>  value="1">����
					<input type="checkbox" name="chGift" <%IF Cstr(egift)="1" THEN%> checked<%END IF%>  value="1">����ǰ
					<input type="checkbox" name="chCoupon" <%IF Cstr(ecoupon)="1" THEN%> checked<%END IF%> value="1">���� 
					<input type="checkbox" name="chOnlyTen" <%IF Cstr(eonlyten)="1" THEN%> checked<%END IF%> value="1">Only-TenByTen 
					<input type="checkbox" name="chopo" <%IF Cstr(eopo)="1" THEN%> checked<%END IF%> value="1">1+1  
			    </td> 
			</tr>
			<tr>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">�̺�Ʈ����:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
						<%sbGetOptCommonCodeArr "eventtype", eType, True,True,"onChange='javascript:document.frmEvt.submit();'"%> &nbsp; 
						<input type="checkbox" name="blnCnfm" <%=chkIIF(Cstr(isConfirm)="1","checked","")%> value="1">���οϷ�
				</td>
				<td colspan="3" style="border-top:1px solid <%= adminColor("tablebg") %>;">&nbsp;</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;"> 
					<input type="checkbox" name="chfd" <%IF Cstr(efd)="1" THEN%> checked<%END IF%> value="1">������
					<input type="checkbox" name="chbs" <%IF Cstr(ebs)="1" THEN%> checked<%END IF%> value="1">�����Ǹ� 
					<input type="checkbox" name="chDiary" <%IF Cstr(eDiary)="1" THEN%> checked<%END IF%> value="1">DiaryStory
					<input type="checkbox" name="chnew" <%IF Cstr(enew)="1" THEN%> checked<%END IF%> value="1">��Ī
				</td>
			</tr>
			<tr>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">�������:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
				    <% Dim arrCode 
                     	arrCode= fnSetCommonCodeArr("eventstate", False)
 	
 	                 %>
				    <select name="eventstate" class="select" onChange="javascript:SubmitForm();">
				        <option value="">����</option>
				        <% 	IF isArray(arrCode) THEN
                         	For intLoop =0 To UBound(arrCode,2)
                         	    if arrCode(0,intLoop) = 1 THEN
                         	 
                        %>
                            <option value="1^3" <%If CStr(sState) = "1^3" THEN%>selected<%END IF%>>�̹�����Ͽ�û+�����̳��۾���</option>
                        <%      elseif arrCode(0,intLoop) = 9 THEN%>    
                            <option value="6^9" <%If CStr(sState) = "6^9" THEN%>selected<%END IF%>>����+����</option>
                        <%      end if%>
                        	<option value="<%=arrCode(0,intLoop)%>" <%If CStr(sState) = CStr(arrCode(0,intLoop)) THEN%>selected<%END IF%> <%if arrCode(2,intLoop) ="N" then%>style="color:gray;"<%end if%>> <%=arrCode(1,intLoop)%></option>
                        <%
                        	Next
                        	End IF
                        	%>
				     
				    </select>  
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
			    </td>
				<td colspan="2" style="border-top:1px solid <%= adminColor("tablebg") %>;">&nbsp;</td>
			</tr>
			<tr>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">�����λ���:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
					<span style="white-space:nowrap; font-family:'Lucida Console';">PC <%sbGetOptEventCodeValue "designerstatus",edgstat1,True,""%></span> &nbsp;
					<span style="white-space:nowrap; font-family:'Lucida Console';">MW <%sbGetOptEventCodeValue "designerstatus",edgstat2,True,""%></span>
				</td>
			    <td style="border-top:1px solid <%= adminColor("tablebg") %>;">�����:</td>
			    <td style="border-top:1px solid <%= adminColor("tablebg") %>;" colspan="3">
			    	<span style="white-space:nowrap;">��ȹ��  <input type="hidden" name="sMdId" value="<%=emdid%>"><input type="name" name="sMdNm" value="<%=eMDnm%>" class="text"  size="10">&nbsp;<input type="button" class="button" value="����" onClick="jsGetID('Md','162','<%=emdid%>');"> <input type="button" value="&times"  class="button" onClick="jsDelID('Md');" title="����� �����" /></span> &nbsp;&nbsp;
			    	<span style="white-space:nowrap;">�����̳� <input type="hidden" name="sDgId" value="<%=edgid%>"><input type="name" name="sDgNm" value="<%=edgnm%>" class="text"  size="10">&nbsp;<input type="button" class="button" value="����" onClick="jsGetID('Dg','152','<%=edgid%>');">&nbsp;<input type="button" value="&times"  class="button" onClick="jsDelID('Dg');" title="����� �����" /></span> &nbsp;
			    	<span style="white-space:nowrap;">�ۺ��� <input type="hidden" name="sPsId" value="<%=epsid%>"><input type="name" name="sPsNm" value="<%=epsnm%>"class="text"  size="10">&nbsp;<input type="button" class="button" value="����"  onClick="jsGetID('Ps','157','<%=epsid%>');"> <input type="button" value="&times"  class="button" onClick="jsDelID('Ps');" title="����� �����" /></span> &nbsp;&nbsp;
			    	<span style="white-space:nowrap;">������  <input type="hidden" name="sDpId" value="<%=edpid%>"><input type="name" name="sDpNm" value="<%=edpnm%>" class="text"  size="10">&nbsp;<input type="button" class="button" value="����" onClick="jsGetID('Dp','130','<%=edpid%>');"> <input type="button" value="&times"  class="button" onClick="jsDelID('Dp');" title="����� �����" /></span>
			    </td>
			</tr>
			<tr>
			<tr> 
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">ä��:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
					<input type="checkbox" name="isWeb" value="1" <%if blnWeb="1" then%>checked<%end if%>> PC-Web
					<input type="checkbox" name="isMobile"  value="1" <%if blnMobile="1" then%>checked<%end if%>> Mobile
					<input type="checkbox" name="isApp"  value="1" <%if blnApp="1" then%>checked<%end if%>> App
				</td>
			    <tD colspan="4" style="border-top:1px solid <%= adminColor("tablebg") %>;"><input type="checkbox" name="chkPus" value="1" <%if blnReqPublish THEN%>checked<%end if%>> �ۺ��� ��û�۾�</td>
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
        	<input type="button" value="���ε��" onclick="jsGoUrl('event_regist.asp?menupos=<%=menupos%>&<%=strParm%>');" class="button">
	    </td>
	    <td align="right">
	       	<input type="button" value="������" onclick="jsSchedule();"  class="button">
	       <!--	<input type="button" value="���" onclick=" ">  -->
	       <% if C_ADMIN_AUTH then %><input type="button" value="�ڵ����" onclick="jsCodeManage();"  class="button"><%END IF%>
        </td>
	</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="20">�˻���� : <b><%=iTotCnt%></b>&nbsp;&nbsp;������ : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td nowrap rowspan="2">ä��</td>
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
      	<td nowrap colspan="5">�����</td>
      	<td nowrap rowspan="2">����</td>
     </tr>
     <tr align="center" bgcolor="<%= adminColor("tabletop") %>">	 
        <td nowrap>��ȹ��</td>
      	<td colspan="2" nowrap>�����̳�</td>
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
'				Case "15"		'�귣�嵥��
'					sWeb =  "<a href='" & vwwwUrl & "/street/street_brandday.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>"  
'					sMoblie =  "<a href='" & mobileUrl & "/street/street_brandday.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>"  
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
    	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetCommCodeArrDesc(arreventkind,arrList(1,intLoop))%></a></a></td>
    	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetCommCodeArrDesc(arreventtype,arrList(39,intLoop))%></a></a></td>
		<% 	
			
			'�̺�Ʈ ������ ���� ����Ʈ��ũ ������ ����
			IF isMobile or isApp  THEN '�����/���϶�..
				strURL = vmobileUrl
			ELSE	'��Ÿ..
				strURL = vwwwUrl
			END IF	
			Select Case arrList(1,intLoop)
				Case "7"		'��Ŭ���ڵ�
					Response.Write "<td><a href='" & strURL & "/guidebook/weekly_coordinator.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case "13"		'��ǰ �̺�Ʈ
					Response.Write "<td><a href='" & strURL & "/shopping/category_prd.asp?itemid=" & arrList(21,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case "14"		'��ǳ���±�
					Response.Write "<td><a href='" & strURL & "/guidebook/picnic/picnic.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
'				Case "15"		'�귣�嵥��
'					Response.Write "<td><a href='" & strURL & "/street/street_brandday.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case "16"		'�귣�� �������
					Response.Write "<td><a href='" & strURL & "/street/street_brand_sub06.asp?makerid=" & arrList(14,intLoop) & "&shop_event_yn=Y&shop_event_confirm=Y&shopview=3' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case "22"		'DAY&(���̾ص�)
					Response.Write "<td><a href='" & strURL & "/guidebook/dayand.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case "26"		'�����
					Response.Write "<td><a href='" & strURL & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case Else		'�������� �� ��Ÿ
					Response.Write "<td><a href='" & strURL & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
			End Select 
		%>
    	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetCommCodeArrDesc(arreventlevel,arrList(7,intLoop))%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetCommCodeArrDesc(arreventstate,arrList(8,intLoop))%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%IF arrList(34,intLoop) <> "" THEN%> <img src="<%=arrList(34,intLoop)%>" width="100" border="0"><%END IF%></a></td>
        <td>
            <a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%IF arrList(35,intLoop) <> "" THEN%> <img src="<%=arrList(35,intLoop)%>" width="100" border="0"><%END IF%></a>
            <!-- // 2017.04.18 ������� �������� ���� (����� �ؽ�Ʈ �̹��� ���)
            <%IF arrList(37,intLoop) <> "" THEN%> <br><img src="<%=arrList(37,intLoop)%>" width="100" border="0" onclick="makeThumbBanTxt('<%=arrList(0,intLoop)%>','on')"><%ELSE%><%IF arrList(35,intLoop) <> "" THEN%><br><input type="button" value="����" onclick="makeThumbBanTxt('<%=arrList(0,intLoop)%>','')"><%END IF%><%END IF%>
            //-->
        </td>  
      	<td align="left">
      		<a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=chkIIF(Not(arrList(25,intLoop)="" or isNull(arrList(25,intLoop))),"["&arrList(25,intLoop)&"] ","")%>
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
      		</a>
      	</td>
      	<td>
      		<a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>">
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
      		</a>
      	</td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(14,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(5,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(6,intLoop)%></a></td> 
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(36,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(23,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(11,intLoop)%><br /><span style="color:#B88;"><%=fnGetCommCodeArrDesc(arrdsnStat,arrList(42,intLoop))%></span></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(41,intLoop)%><br /><span style="color:#B88;"><%=fnGetCommCodeArrDesc(arrdsnStat,arrList(43,intLoop))%></span></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(28,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(29,intLoop)%><%=chkiif(arrList(38,intLoop)<>"","<br />" & arrList(38,intLoop),"")%></a></td>
      	<td align="left" nowrap><input type="button" value="��ǰ" class="button" onClick="javascript:jsGoUrl('eventitem_regist.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>&selsort=<%=itemSortvalue%>')">
      		<%IF arrList(13,intLoop) > "1900-01-01" THEN%><input type="button" value="��÷" class="button" onClick="jsGoUrl('/admin/eventmanage/event/eventprize_regist.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>')"><%END IF%>
      		<%if arrList(15,intLoop)  then%> <input type="button" value="����(<%=arrList(18,intLoop)%>)" class="button" onClick="jsGoUrl('/admin/shopmaster/sale/salelist.asp?eC=<%=arrList(0,intLoop)%>&menupos=290');"><%end if%>
      		<%if arrList(16,intLoop) then%> <input type="button" value="����ǰ(<%=arrList(19,intLoop)%>)" class="button" onClick="jsGoUrl('/admin/shopmaster/gift/giftlist.asp?eC=<%=arrList(0,intLoop)%>&menupos=1045');"><%end if%>
      		<!--<%if arrList(17,intLoop) then%> <input type="button" value="����" class="button" onClick="jsGoUrl('coupon');"><%end if%>	-->
      		<% If arrList(20,intLoop) = "N" Then %>
      		<table cellpadding="0" cellspacing="0" border="0"><tr><td style="padding:3 0 0 0;"><input type="button" class="button" style="width:105;" value="��÷�ھ��� ����" onclick="prize(<%= arrList(0,intLoop) %>);"></td></tr></table>
      		<% End IF %>
      	</td>
    </tr>
   <%	Next
   	ELSE
   %>
   	<tr  align="center" bgcolor="#FFFFFF">
   		<td colspan="20">��ϵ� ������ �����ϴ�.</td>
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
