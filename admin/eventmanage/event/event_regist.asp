<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/event_regist.asp
' Description :  �̺�Ʈ ���� ���
' History : 2007.02.07 ������ ����
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

elevel = 2 '�߿䵵 �������� �ӽ� ����


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
	eKind		= sKind
	edid  		= requestCheckVar(Request("selDId"),32)		'��� �����̳�
	emid  		= requestCheckVar(Request("selMId"),32)		'��� MD

	ebrand		= requestCheckVar(Request("ebrand"),32)		'�귣��
	esale		= requestCheckVar(Request("chSale"),2) 		'��������
	egift		= requestCheckVar(Request("chGift"),2)		'����ǰ����
	ecoupon	 	= requestCheckVar(Request("chCoupon"),2)	'��������
	eonlyten	= requestCheckVar(Request("chOnlyTen"),2)	'Only-TenByTen����

	eOneplusone	= requestCheckVar(Request("chOneplusone"),2)	'oneplusone
	eFreedelivery= requestCheckVar(Request("chFreedelivery"),2)	'������
	eBookingsell= requestCheckVar(Request("chBookingsell"),2)	'�����Ǹ�
	eDiary= requestCheckVar(Request("chDiary"),2)	'���̾
	edispCate	= requestCheckVar(Request("disp"),10) 		'���� ī�װ�

	strparm  = "selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&selCM="&sCateMid&"&eventstate="&sState&"&eventkind="&sKind&"&selDId="&edid&"&selMId="&emid&_
				"&ebrand="&ebrand&"&chSale="&esale&"&chGift="&egift&"&chCoupon="&ecoupon&"&chOneplusone="&eOneplusone&"&chFreedelivery="&eFreedelivery&"&chBookingsell="&eBookingsell&"&disp="&edispCate&"&chOnlyTen="&eonlyten&"&chDiary="&eDiary
	'#######################################
IF eCode <> "" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'�̺�Ʈ �ڵ�
	'�̺�Ʈ ���� ��������
	cEvtCont.fnGetEventCont
	ekind =	cEvtCont.FEKind
	eman =	cEvtCont.FEManager
	escope =	cEvtCont.FEScope
	ename =	db2html(cEvtCont.FEName)
	enameEng =	db2html(cEvtCont.FENameEng) '�̺�Ʈ ���� �߰�
	subcopyK =	db2html(cEvtCont.FsubcopyK) '�̺�Ʈ ���� �߰�
	subcopyE =	db2html(cEvtCont.FsubcopyE) '�̺�Ʈ ���� �߰�

	elevel =	cEvtCont.FELevel
	'estate =	cEvtCont.FEState
	eregdate =	cEvtCont.FERegdate

	'�̺�Ʈ ȭ�鼳�� ���� ��������
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

'2014-08-27 ������ ���� / �󼼳��� üũ�� ����Ʈ�� MD�� ��û
echkdisp = 1

%>
<script language="javascript">
<!--
//-- jsEvtSubmit : �̺�Ʈ ��� --//
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
	  	frm.sEN.focus();
	  	return false;
	  }

	  if(frm.sEN.value.length > 80){
		alert("�̺�Ʈ���� 60�ڱ����� �����մϴ�.�ٽ� �Է����ּ���.");
	 	frm.sEN.focus();
	  	return false;
	  }

	   if(frm.sENEng.value.length > 120){
		alert("�����̺�Ʈ���� 120�ڱ����� �����մϴ�.�ٽ� �Է����ּ���.");
	 	frm.sENEng.focus();
	  	return false;
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


	  	if(frm.sSD.value < nowDate){
	  		alert("�������� �����Ϻ���  ������ �ȵ˴ϴ�. �������� �ٽ� �������ּ���");
	  		frm.sSD.focus();
	  		return false;
	  	}

		if(!frm.selMId.value){
			alert('����ڸ� �����ϼ���');
			return false;
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
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

//-- jsChangeKind : �̺�Ʈ����(Kind)�� ���� ȭ�� View ���� --//
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

//-- jsLastEvent : ���� �̺�Ʈ �ҷ����� --//
	function jsLastEvent(){
	  var winLast,eKind;
	  eKind = document.frmEvt.eventkind.options[document.frmEvt.eventkind.selectedIndex].value;
	  winLast = window.open('pop_event_lastlist.asp?menupos=<%=menupos%>&eventkind='+eKind,'pLast','width=550,height=600, scrollbars=yes')
	  winLast.focus();
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

	// ��� ��ũ���� Eable
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
	<td> <img src="/images/icon_arrow_link.gif" align="absmiddle"> �̺�Ʈ ���� ��� </td>
</tr>
<tr>
	<td><input type="button" value="���� �̺�Ʈ ���� �ҷ�����" class="button" onClick="jsLastEvent();"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
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
		   			<input type="hidden" name="eventscope" value="2">
		   			<label><input type="checkbox" name="chkEscope" checked onclick="jsSetPartner()"> 10x10</label>
		   			<label><input type="checkbox" name="chkEscope" onclick="jsSetPartner()"> ���޸�</label>
		   			<span id="spanP" style="display:none;">
		   			<select name="selP">
		   				<option value="">--���޸� ��ü--</option>
		   				<% sbOptPartner selPartner%>
		   			</select>
		   			</span>
		   		</td>
		   	</tr>
		   	<tr id="eNameTr_A">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>�̺�Ʈ��</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sEN" size="80" maxlength="120" value="<%=ename%>">
		   		</td>
		   	</tr>
			<tr id="eNameTr_C">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>���� �̺�Ʈ��</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sENEng" size="60" maxlength="60" value="<%=enameEng%>">
		   		</td>
		   	</tr>
		   	<tr id="eNameTr_B" style="display:none;">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>�̺�Ʈ��<br>�� ������</B></td>
		   		<td bgcolor="#FFFFFF">
		   			�̺�Ʈ��: <input type="text" name="sEDN" size="50" maxlength="50" value=""><br>
		   			�����̺�Ʈ��: <input type="text" name="sEDNEng" size="50" maxlength="50" value=""><br>
		   			������: ���� <input type="text" name="sSDc" size="4" value="0" style="text-align:right;">% ~
		   			�ְ� <input type="text" name="sMDc" size="4" value="0" style="text-align:right;">%<br>
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
		   			������ : <input type="text" name="sSD" size="10" onClick="jsPopCal('sSD');"  style="cursor:hand;">
		   			~ ������ : <input type="text" name="sED"   size="10" onClick="jsPopCal('sED');" style="cursor:hand;">
		   		</td>
		   	</tr>
		   	<tr>
		   		<td  bgcolor="#FFFFFF">
		   			��÷ ��ǥ�� : <input type="text" name="sPD" size="10" onClick="jsPopCal('sPD');" style="cursor:hand;">
		   			(��÷�ڰ� �ִ� ��쿡�� ���)
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>����</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptStatusCodeValue "eventstate",estate,false,""%>
		   			<%''sbGetOptStatusCodeAuth "eventstate",0,"N",""%>
		   		</td>
		   	</tr>
			<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>�߿䵵</b></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventlevel",elevel,False,""%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>���Ĺ�ȣ</b></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sortNo" value="0" size="6" maxlength="5" style="text-align:right;" />
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
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		   	<tr>
		   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">���� ī�װ�</td>
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
		   		<input type="checkbox" name="chSale" <%IF esale = "1" THEN%>checked<%END IF%> value="1">����
		   		<input type="checkbox" name="chGift" <%IF egift = "1" THEN%>checked<%END IF%> value="1">����ǰ
		   		<input type="checkbox" name="chCoupon" <%IF ecoupon = "1" THEN%>checked<%END IF%> value="1">����
		   		<input type="checkbox" name="chOnlyTen" <%IF eonlyten ="1" THEN%>checked<%END IF%> value="1">Only-TenByTen
		   		<input type="checkbox" name="chOneplusone" <%IF eOneplusone ="1" THEN%>checked<%END IF%> value="1">1+1
				<input type="checkbox" name="chFreedelivery" <%IF eFreedelivery ="1" THEN%>checked<%END IF%> value="1">������
				<input type="checkbox" name="chBookingsell" <%IF eBookingsell="1" THEN%>checked<%END IF%> value="1">�����Ǹ�
				<input type="checkbox" name="chDiary" <%IF eDiary="1" THEN%>checked<%END IF%> value="1">DiaryStory
		   		</td>
			</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ ���</td>
		   		<td bgcolor="#FFFFFF">
		   		<input type="checkbox" name="chComm" <%IF ecomment = 1 THEN%>checked<%END IF%> value="1" >�ڸ�Ʈ
		   		<input type="checkbox" name="chBbs" <%IF ebbs = 1 THEN%>checked<%END IF%> value="1" >�Խ���
		   		<input type="checkbox" name="chItemps" <%IF eitemps = 1 THEN%>checked<%END IF%> value="1" >��ǰ�ı�
		   		<input type="checkbox" name="isblogurl" <%IF eisblogurl THEN%>checked<%END IF%> value="1" onClick="jsChkBlogEnable()">Blog URL
		   		<!--<input type="checkbox" name="chApply" <%IF eapply = 1 THEN%>checked<%END IF%> value="1" >����-->
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ ��ũ Ÿ��</td>
		   		<td bgcolor="#FFFFFF">
		   			<label><input type="radio" name="elType" value="E" onclick="jsEvtLink(true);" checked >�̺�Ʈ</label>
		   			<label><input type="radio" name="elType" value="I" onclick="jsEvtLink(false);" >�����Է�</label>
		   			&nbsp;<input type="text" id="elUrl" name="elUrl" size="40" maxlength="128" value="" class="text_ro" readOnly >

		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ǰ���Ĺ��</td>
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
		   		<td bgcolor="#FFFFFF">
					<% sbGetwork "selMId",emid,"" %>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�۾����޻���</td>
		   		<td bgcolor="#FFFFFF">
		   			�۾����� <input type="text" name="sWorkTag" size="20" maxlength="16" class="text"> <font color="darkgray">(for Designer)</font>
		   			<textarea name="tFwd" rows="15" style="width:100%;font-size:12px;font-family:'Malgun Gothic',dotum;"><%=efwd%></textarea>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">Comment Title</td>
		   		<td bgcolor="#FFFFFF">
		   			(200�� �̳�)		   			<Br>
		   			<textarea name="eCT" rows="2" style="width:100%;"></textarea>
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
