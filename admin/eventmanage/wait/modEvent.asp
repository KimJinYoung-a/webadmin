<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �̺�Ʈ ��� - ȭ�鼳��
' History :  
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventPartnerWaitCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->

<!-- #include virtual="/admin/lib/adminbodyhead_html5.asp"-->
<%
dim evtCode
evtCode =    requestCheckVar(Request("eC"),10)

if evtCode = "" then	
		Call Alert_return ("���԰�ο� ������ �ֽ��ϴ�. 1�ܰ���� �ٽ� ������ּ���. ")
	response.end
end if
dim dispcate, maxDepth
	dispcate	= requestCheckVar(Request("disp"),10) 		'���� ī�װ�
	maxDepth = 2 '����ī�װ� 2depth���� �����ش�

dim makerid, brandNm
dim eRectGCode,mdRectTheme
eRectGCode =    requestCheckVar(Request("eRGC"),10)
mdRectTheme=    requestCheckVar(Request("mdtm"),1)
dim ClsEvt
dim evtkind,evtmanager  ,evtname,evtstartdate,evtenddate,evtstate,evtregdate,evtusing ,evtlastupdate,adminid, evtcategory  ,evtcateMid,isgift ,brand  ,evttag
dim titlepc, titlemo,issale, iscoupon, saleper, salecper
dim etcitemimg,evt_mo_listbanner,subcopyK  ,evtsubname,mdtheme   ,themecolor,themecolormo ,textbgcolor       
dim giftisusing ,gifttext1 ,giftimg1  ,gifttext2 ,giftimg2  ,gifttext3 ,giftimg3          
dim giftimg1Nm, giftimg2Nm, giftimg3Nm
dim catenm, cateMnm
dim arrList, intLoop,isort , ino
dim arrimg, arrimgmo
set ClsEvt = new CEvent
ClsEvt.FevtCode = evtCode
ClsEvt.fnGetEventST4

evtkind       =clsEvt.Fevtkind      
evtmanager   = clsEvt.Fevtmanager   
evtname      = clsEvt.Fevtname      
evtstartdate  =clsEvt.Fevtstartdate 
evtenddate   = clsEvt.Fevtenddate   
evtstate      =clsEvt.Fevtstate     
evtregdate   = clsEvt.Fevtregdate   
evtusing     = clsEvt.Fevtusing     
evtlastupdate= clsEvt.Fevtlastupdate
adminid      = clsEvt.Fadminid     
dispcate		 =  clsEvt.Fevtdispcate  
catenm 			= clsEvt.FevtCateNm
cateMnm 		= clsEvt.FevtCateMNm
issale       = clsEvt.Fissale       
isgift      =  clsEvt.Fisgift       
iscoupon    =  clsEvt.Fiscoupon     
brand       =  clsEvt.Fbrand        
evttag      =  clsEvt.Fevttag    
brandNm = ClsEvt.FBrandNm
titlepc = ClsEvt.FTitlePC
titlemo = ClsEvt.FTitleMO 
saleper =  ClsEvt.Fsaleper
salecper =  ClsEvt.Fsalecper
etcitemimg        =ClsEvt.Fetcitemimg
evt_mo_listbanner =ClsEvt.Fevt_mo_listbanner 
subcopyK          =ClsEvt.FsubcopyK          
evtsubname        =ClsEvt.Fevtsubname        
mdtheme           =ClsEvt.Fmdtheme           
themecolor        =ClsEvt.Fthemecolor        
themecolormo      =ClsEvt.Fthemecolormo      
textbgcolor       =ClsEvt.Ftextbgcolor       
giftisusing       =ClsEvt.Fgiftisusing       
gifttext1         =ClsEvt.Fgifttext1         
giftimg1          =ClsEvt.Fgiftimg1          
gifttext2         =ClsEvt.Fgifttext2         
giftimg2          =ClsEvt.Fgiftimg2          
gifttext3         =ClsEvt.Fgifttext3         
giftimg3          =ClsEvt.Fgiftimg3          
 
 arrList = clsEvt.fnGetEventGroup
  if mdRectTheme <>"" then mdtheme = mdRectTheme
if mdtheme="3" then
 	ClsEvt.Fsdiv ="W"
 	arrimg 		= ClsEvt.fnGetEventItemImg
 	ClsEvt.Fsdiv ="M"
 	arrimgmo 		= ClsEvt.fnGetEventItemImg
elseif mdtheme ="2" then
	 ClsEvt.Fsdiv ="W"
	arrimg = ClsEvt.fnGetEventSlideImg
	 ClsEvt.Fsdiv ="M"
	 arrimgmo = ClsEvt.fnGetEventSlideImg
end if
set ClsEvt = nothing

if themecolor =""  or isNull(themecolor) then themecolor ="11"
if themecolormo =""  or isNull(themecolormo)  then themecolormo ="11"
if giftisusing ="" then giftisusing =0
dim tmpg 
dim embanNm,etcitemimgNm
if etcitemimg <> "" then
	tmpg = split(etcitemimg,"/")
	etcitemimgNm = tmpg(ubound(tmpg))
end if
if evt_mo_listbanner <> "" then
	tmpg = split(evt_mo_listbanner,"/")
	embanNm = tmpg(ubound(tmpg))
end if
if giftimg1 <> "" then
	tmpg = split(giftimg1,"/")
	giftimg1Nm = tmpg(ubound(tmpg))
end if
if giftimg2 <> "" then
	tmpg = split(giftimg2,"/")
	giftimg2Nm = tmpg(ubound(tmpg))
end if
if giftimg3<> "" then
	tmpg = split(giftimg3,"/")
	giftimg3Nm = tmpg(ubound(tmpg))
end if 

	 
%>
<link rel="stylesheet" type="text/css" href="/css/adminPartnerCommon.css" />

<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script type="text/javascript" src="/js/jquery.swiper-3.3.1.min.js"></script>
<script type="text/javascript" src="/js/tag-it.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
	 //���� �ٿ�ε�
    function jsDownload(sDownURL, sRFN, sFN){
    var winFD = window.open(sDownURL+"/linkweb/board/procDownload.asp?sRFN="+sRFN+"&sFN="+sFN,"popFD","");
    winFD.focus();
 }
 
	//����
	function jsRegEvent(mps){	
		
		var frm = document.frmReg;
		if(!frm.evtNm.value){
			alert("��ȹ������ �Է����ּ���");
			frm.evtNm.focus();
			return;
		}
		
	  if(frm.evtNm.length>60){
	  	alert("��ȹ������ �ִ� 60�ڱ����� �����մϴ�.");
	  	frm.evtNm.focus();
	  	return;
	  }

 	 if(!frm.evtSD.value || !frm.evtED.value ){
	  	alert("��ȹ���� �Ⱓ�� �Է����ּ���"); 
	  	return ;
	  }


	  if(frm.evtSD.value > frm.evtED.value){
	  	alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���"); 
	  	return ;
	  }

	   var nowDate = jsNowDate();


	  	if(frm.evtSD.value < nowDate){
	  		alert("�������� �����Ϻ���  ������ �ȵ˴ϴ�. �������� �ٽ� �������ּ���");
	  		frm.evtSD.focus();
	  		return ;
	  	}
	  	
	  	if(frm.disp.value==0){
	  		alert("��ȹ���� ����ī�װ���  �������ּ���");
	  		return;
	  	}
	  	
	  	$("#evtTag").val("");
		if($("#evtTag").val().length < 1 ){
			$("input[name=tags]").each(function(idx){   
				// �ش� üũ�ڽ��� Value ��������
				var value = $("#evtTag").val();
				var eqValue = $("input[name=tags]:eq(" + idx + ")").val() ;
					if($("#evtTag").val().length < 1 ){
					$("#evtTag").val(eqValue);
					console.log(value + "," + eqValue) ;
				}else{
					$("#evtTag").val(value + "," + eqValue);
				}
			});
		}
		  
	  	if(!frm.evtTag.value){
	  		alert("�˻� Tag�� �Է����ּ���");
	  		frm.evtTag.focus();
	  		return;
	  	}
	  	
	  	if(GetByteLength(frm.evtTag.value)>300){
	  		alert("�˻� Tag�� 300byte (�ѱ� 150��, ���� 300��) �̳��� �Է����ּ���");
	  		frm.evtTag.focus();
	  		return;
	  	}
	  	
	  if (frm.gCnt.value==1){
 		alert("�ּ� 1�� �̻��� �׷��� ��ϵǾ�� �մϴ�.�׷��� ������ּ���");
 		return;
 	}
 	
 	if(!frm.hiddf.value){
 		alert("�⺻��ʸ� ���ε����ּ��� ");
 		return;
 	}
 	if(!frm.hidwb.value){
 		alert("���̵��ʸ� ���ε����ּ��� ");
 		return;
 	}
 	if(!frm.evtNmW.value){
 		alert("PC�׸��� ��ȹ������ �Է����ּ���");
 		frm.evtNmW.focus();
 		return;
 	}
 	if(GetByteLength(frm.evtNmW.value.length)>35){
 		alert("PC�׸��� ��ȹ������ �ִ� 35byte���� �Է°����մϴ�.");
 		frm.evtNmW.focus();
 		return;
 	}
 	
 	if(!frm.subcopyK.value){
 		alert("PC�׸��� ����ī�Ǹ� �Է����ּ���");
 		frm.subcopyK.focus();
 		return;
 	}
 	 	
 	if(GetByteLength(frm.subcopyK.value.length)>120){
 		alert("PC�׸��� ����ī�Ǵ� �ִ� 120byte���� �Է°����մϴ�.");
 		frm.subcopyK.focus();
 		return;
 	} 	
 	if(!frm.evtNmM.value){
 		alert("Mobile�׸��� ��ȹ������ �Է����ּ���");
 		frm.evtNmM.focus();
 		return;
 	}
 	if(GetByteLength(frm.evtNmM.value.length)>35){
 		alert("Mobile�׸��� ��ȹ������ �ִ� 35byte���� �Է°����մϴ�.");
 		frm.evtNmM.focus();
 		return;
 	}
 	if(!frm.evtsubname.value){
 		alert("Mobile�׸��� ����ī�Ǹ� �Է����ּ���");
 		frm.evtsubname.focus();
 		return;
 	}
 	if(GetByteLength(frm.evtsubname.value.length)>120){
 		alert("Mobile�׸��� ����ī�Ǵ� �ִ� 120byte���� �Է°����մϴ�.");
 		frm.evtsubname.focus();
 		return;
 	}
 	
 	frm.hidM.value ="U";
 	frm.target="_self";
 	frm.submit(); 	
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
	
	

//�׷��߰�
function jsAddGroup(){
	document.frmReg.hidGNm.value = "";
 if(!document.getElementById("eGD").value){
 	alert("�׷���� �Է����ּ���");
 	document.getElementById("eGD").focus();
 	return;
 }
 
 if(GetByteLength(document.getElementById("eGD").value)>64){
 	alert("�׷���� 64byte(�ѱ� 32��, ���� 64��) �̳��� �Է����ּ���");
 	document.getElementById("eGD").focus();
 	return;
 }
 
 document.frmReg.hidM.value="GA";
 document.frmReg.hidGNm.value = document.getElementById("eGD").value; 
 document.frmReg.hidGS.value =  document.getElementById("eGS").value;  
 document.frmReg.target = "FrameCKP";  
 document.frmReg.submit();  
}


//�׷����
function jsDelGroup(eGC){
	if(confirm("���õ� �׷��� �����Ͻðڽ��ϱ�?")){
		document.frmReg.hidM.value ="GD";
		document.frmReg.eGC.value = eGC;
		document.frmReg.target = "FrameCKP";
		document.frmReg.submit();		
	}
	return;
}

//�׷����
function jsSetGList(eGC,sMode){  
	var str = $.ajax({
		type: "GET",
		url: "/admin/eventmanage/wait/ajaxGroup.asp",
		data: "eC=<%=evtCode%>&eGC="+eGC,
		dataType: "text",
		async: false
		}).responseText;
	if (str != ""){   
		$("#dList").html(str); 
		if (sMode =="A"){			
			gCnt = parseInt($("#gCnt").val());
			$("#gCnt").val(gCnt+1);
		}else if(sMode=="D"){
			gCnt = parseInt($("#gCnt").val());
			$("#gCnt").val(gCnt-1);
		}
	}
}

//�׷��������
function jsModGSubmit(eGC){
	 
if(!document.getElementById("eMGD").value){
 	alert("�׷���� �Է����ּ���");
 	document.getElementById("eMGD").focus();
 	return;
 }
 
 if(GetByteLength(document.getElementById("eMGD").value)>64){
 	alert("�׷���� 64byte(�ѱ� 32��, ���� 64��) �̳��� �Է����ּ���");
 	document.getElementById("eMGD").focus();
 	return;
 }
 
	document.frmReg.hidM.value ="GM";
	document.frmReg.eGC.value = eGC;
	document.frmReg.hidGNm.value = document.getElementById("eMGD").value;
	document.frmReg.target = "FrameCKP";  
	document.frmReg.submit();	
}



//�׷� ��ǰ���
function jsSetItem(eGC){
	var winItem = window.open('/admin/eventmanage/wait/popRegDispItem.asp?eC=<%=evtCode%>&eGC='+eGC,'popItem','width=1600,height=750,scrollbars=yes,resizable=yes');
 	winItem.focus();
}

//���
	function jsCancel(mps){	
		if (confirm("��ȹ�� ������ �������� �ʰ� ����Ͻðڽ��ϱ�?")){
			location.href = "/admin/eventmanage/wait/index.asp?menupos="+mps;
		} 
			return;
		 
	} 
 
  
  
  // �̹������
 function  jsRegImg(sType, iMW,iMH,pvWidth){
 	var winImg = window.open('/admin/eventmanage/wait/popRegImg.asp?eC=<%=evtCode%>&sType='+sType+'&iMH='+iMH+'&iMW='+iMW+'&pvWidth='+pvWidth,'popImg','width=500,height=350,scrollbars=yes,resizable=yes');
 	winImg.focus();
 }
 
 function  jsRegMultiImg(sType, iMW,iMH){
 	var winImg = window.open('/admin/eventmanage/wait/popRegMultiImg.asp?eC=<%=evtCode%>&sType='+sType+'&iMH='+iMH+'&iMW='+iMW,'popImg','width=500,height=600,scrollbars=yes,resizable=yes');
 	winImg.focus();
 }
 
 //�̹�������
 function jsDelimg(sType){
 	$("#"+sType+"Img").empty();
 	$("#"+sType+"Nm").remove();
 	$("#hid"+sType).val("");
 }
 

												
 //��ǰ���
 function jsRegItem(sdiv){
 	var winItem = window.open('/admin/eventmanage/wait/popRegItem.asp?eC=<%=evtCode%>&sdiv='+sdiv,'popItem','width=1600,height=750,scrollbars=yes,resizable=yes');
 	winItem.focus();
 }
</script>
 <div class="content scrl" style="top:25px;">
							<!-- content--->	
							<div class="cont">
								<div class="pad20 exhibit-manage"> 
									<div class="basicInfo  ">
										<h3 class="bltNo">1. �⺻ ����</h3>										
										<form name="frmReg" method="post" action="procEvent.asp">
											<input type="hidden" name="hidM" value="U">
											<input type="hidden" name="menupos" value="<%=menupos%>">
											<input type="hidden" name="eC" value="<%=evtCode%>">
											<input type="hidden" name="eGC" value="">	
											<input type="hidden" name="hidGNm" value="">
											<input type="hidden" name="hidGS" value="">	
											<input type="hidden" name="makerid" value="<%=brand%>">
											<input type="hidden" name="arrGS" value="">
											<input type="hidden" name="arrGC" value="">
												
										<table class="tbType1 writeTb tMar10">
											<colgroup>
												<col width="14%" /><col width="" />
											</colgroup>
											<tbody>
											<tr>
												<th><div>��ȹ�� �� <strong class="cRd1">*</strong></div></th>
												<td>
													<input type="text" class="formTxt" name="evtNm"  value="<%=evtName%>" placeholder="��ȹ�� ���� �Է����ּ���." style="width:100%" maxlength="120"/>
												</td>
											</tr>
											<tr>
												<th><div>�Ⱓ <strong class="cRd1">*</strong></div></th>
												<td>
													<input type="text" name="evtSD" id="evtSD" value="<%=evtstartdate%>" class="formTxt" style="width:100px" placeholder="������"  onKeyup="alert('�޷����� ������ּ���');this.value='';"/>
													<input type="image" name="evtSD_trigger" id="evtSD_trigger" src="/images/admin_calendar.png" alt="�޷����� �˻�" onclick="return false;" />
													~
													<input type="text" name="evtED" id="evtED"  value="<%=evtenddate%>" class="formTxt" style="width:100px" placeholder="������"  onKeyup="alert('�޷����� ������ּ���');this.value='';"//>
													<input type="image" name="evtED_trigger"  id="evtED_trigger"  src="/images/admin_calendar.png" alt="�޷����� �˻�" onclick="return false;"/>
												</td>
											</tr>
											
											<script type="text/javascript"> 
											var CAL_Start = new Calendar({
												inputField : "evtSD", trigger    : "evtSD_trigger",
												onSelect: function() {
													var date = Calendar.intToDate(this.selection.get());
													CAL_End.args.min = date;
													CAL_End.redraw();
													this.hide();
												}, bottomBar: true, dateFormat: "%Y-%m-%d"
											});
											var CAL_End = new Calendar({
												inputField : "evtED", trigger    : "evtED_trigger",
												onSelect: function() {
													var date = Calendar.intToDate(this.selection.get());
													CAL_Start.args.max = date;
													CAL_Start.redraw();
													this.hide();
												}, bottomBar: true, dateFormat: "%Y-%m-%d"
											});
										</script>
											<tr>
												<th><div>��������</div></th>
												<td>
													<span class="rMar10"><input type="checkbox" id="pdtSale" name="evtSale" value="1" <%if issale then%>checked<%end if%>/> <label for="pdtSale">��ǰ����</label></span>
													<span class="rMar10"><input type="checkbox" id="pdtCp" name="evtCoupon" value="1" <%if iscoupon then%>checked<%end if%>/> <label for="pdtCp">����</label></span>
												</td>
											</tr>
											<tr>
												<th><div>���</div></th>
												<td>
													<span class="rMar10"><input type="checkbox" id="gift" name="evtGift" value="1"  <%if isgift then%>checked<%end if%>/> <label for="gift">����ǰ(GIFT)</label></span>
												</td>
											</tr>
											<tr>
												<th><div>���� ī�װ� <strong class="cRd1">*</strong></div></th>
												<td> 
													<!-- #include virtual="/common/module/dispCateSelectBoxDepth_upche.asp"--> 
												</td>
											</tr>
											<tr>
												<th><div>�˻� Tag <strong class="cRd1">*</strong></div></th>
												<td><%dim tmptagtext, tt %>
														<ul id="singleFieldTags">
															
															<% 
																	If Trim(evtTag) <> "" Then 
																		tmptagtext = Split(evtTag, ",")
																		For tt = 0 To UBound(tmptagtext)
																%>
																			<li><%=tmptagtext(tt)%></li>
																<%
																		Next
																	End If 
																%>
														</ul>
													  <input type="hidden" class="formTxt" name="evtTag"  id="evtTag" value="<%=evtTag%>">
														  <p class="tPad05 cBl2 fs11">- ���� ���� Ű����� �������ּ���. �ݷ� ������ �� �� �ֽ��ϴ�.</p>
												</td>
											</tr>
											</tbody>
										</table> 
									</div>
									<div class="displayInfo tMar50">
								<h3 class="bltNo">2. ��ǰ ���� ����</h3>
								
								<div class="overHidden tMar10">
									<div class="ftLt" id="btnGS">
										<input type="button" class="btnOdrChg btn cBl1 fs12" value="�׷���� ����" />
									</div>
									<div class="ftRt">
										<p class="infoTxt">
											<span><img src='/images/ico_odrchg.png' alt='�׷���� ����' /> �� ��� ���� ��, �Ʒ��� �̵� �� ����Ϸ� ��ư�� Ŭ�����ּ���.</span>
										</p>
									</div>
								</div> 
								<div class="tbListWrap tMar05" id="dList">
									<ul class="thDataList">
										<li> 
											<p style="width:90px">����</p>
											<p class="">�׷�� <strong class="cRd1">*</strong></p>
											<p style="width:150px">��ǰ ���� <strong class="cRd1">*</strong></p>
											<p style="width:150px">����</p>
										</li>
									</ul>
									<ul id="sortable" class="tbDataList">									 
										<% isort =0
												iNo =1
										if isArray(arrList) then 
											%> 
										<%	for intLoop = 0 To UBound(arrList,2)	 
										%> 
										<li id="G<%=arrList(0,intLoop)%>">									
											<p style="width:90px"><%=iNo%></p>
											<p class="lt"><%=arrList(1,intLoop)%></p>
											<p style="width:150px"><input type="button" class="btn3 btnIntb" id="btnItem<%=arrList(0,intLoop)%>" value="��ǰ(<%=arrList(3,intLoop)%>)" onclick="jsSetItem('<%=arrList(0,intLoop)%>');" /></p>
											<p style="width:150px">
												<span id="Gbt<%=arrList(0,intLoop)%>"><a href="javascript:jsSetGList('<%=arrList(0,intLoop)%>','');" class="cBl1 tLine">[����]</a></span>
												<span><a href="javascript:jsDelGroup('<%=arrList(0,intLoop)%>');" class="cBl1 tLine">[����]</a></span>
											</p><input type="hidden" name="eMGS" id="eMGS" value="<%=arrList(2,intLoop)%>">
											<input type="hidden" name="eMGC" id="eMGC" value="<%=arrList(0,intLoop)%>">
										</li> 
									<%	 
											iNo = iNo+ 1
										next  
										isort = arrList(2,intLoop-1)+1
										end if%>										  
										<li class="ui-state-disabled" ><!-- for dev msg : ���� �߰��� �׸��� li�� class="ui-state-disabled" �������ּ��� --> 
											<p style="width:90px"  ><%=iNo%></p>
											<p class="lt"><input type="text" class="formTxt" id="eGD" name="eGD" value="" placeholder="�׷���� �Է����ּ���" style="width:100%" maxlength="64"/></p>
											<p style="width:150px"><input type="button" class="btn3 btnIntb" value="��ǰ(0)" onclick="" disabled="true" /></p>
											<p style="width:150px">
												<a href="javascript:jsAddGroup();" class="cRd1 tLine "><strong>[�߰�]</strong></a> 
											</p><input type="hidden" name="eGS" id="eGS" value="<%=isort%>">
										</li> 
									</ul>  
								</div> 
							</div>   
								<input type="hidden" id="gCnt"	 name="gCnt" value="<%=iNo%>">								
							<div class="saleInfo tMar50">
								<h3 class="bltNo">3. ��ȹ�� ���� ����</h3>
								<table class="tbType1 writeTb tMar10">
									<colgroup>
										<col width="14%" /><col width="" />
									</colgroup>
									<tbody>
									<tr>
										<th><div>��ǰ ���� ���� </div></th>
										<td>
											<span class="rMar20"><input type="text" class="formTxt" name="eSP" value="<%=salePer%>" placeholder="0%" style="width:50px" /> (��:~10%)</span>
										</td>
									</tr>
									<tr>
										<th><div>���� ���� ���� </div></th>
										<td>
											<span class="rMar20"><input type="text" class="formTxt" name="eCP" value="<%=saleCper%>" placeholder="0%" style="width:50px" /> (��:~10%)</span>
										</td>
									</tr>
									</tbody>
								</table>
							</div> 
							<div class="themaInfo tMar50">
								 
							<h3 class="bltNo">4. ��� ��� �̹��� ����</h3>
									<div class="tbListWrap tMar10">
										<table class="tbType1 writeTb">
											<colgroup>
												<col width="18%" /><col width="" />
											</colgroup>
											<tbody>
											<tr>
												<th><div>�⺻ ��� <strong class="cRd1">*</strong></div></th>
												<td>
													<div class="inTbSet">
														<div class="formFile">
															<p> 
																<button type="button" onClick="jsRegImg('df','420','420','105');"  class="btn"  >�̹��� ���</button> 
																<input type="hidden" name="hiddf" id="hiddf" value="<%=etcitemimg%>"> 
															</p>
															<p class="tPad05 fs11 cGy1">- �̹��� ������ : <b>420x420</b>px</p>
														</div>
														<div style="width:105px;">
															<p class="registImg" id="dfImg"> 
																<%if etcitemimg <>"" then%>
																<button type="button" onclick="jsDelimg('df')">X</button>
																<img src="<%=etcitemimg%>" alt="" style="width:105px;" /> 
																<%end if%>
															</p>
														</div>
														<div style="width:156px;" class="lPad20">
															<p class="previewImg lMar20">
																<img src="/images/partner/listbnr_preview1.jpg" alt="" style="width:156px;" />
															</p>
														</div>
													</div>
												</td>
											</tr>
											<tr>
												<th><div>���̵� ��� <strong class="cRd1">*</strong></div></th>
												<td>
													<div class="inTbSet">
														<div class="formFile">
															<p> 
																<button type="button" onClick="jsRegImg('wb','750','406','194');" class="btn">�̹��� ���</button> 														
																<input type="hidden" name="hidwb" id="hidwb" value="<%=evt_mo_listbanner%>"> 
																</p>
																<p class="tPad05 fs11 cGy1">- �̹��� ������ : <b>750*406</b>px</p>																
														</div>
														<div style="width:194px">
															<p class="registImg" id="wbImg">	
																<%if evt_mo_listbanner <> "" then%>															
																<button type="button" onclick="jsDelimg('wb')">X</button>
																<img src="<%=evt_mo_listbanner%>" alt="" style="width:194px" />
																<%end if%>
															</p>
														</div>
														<div style="width:156px;" class="lPad20">
															<p class="previewImg lMar20">
																<img src="/images/partner/listbnr_preview2.jpg" alt="" style="width:156px;" />
															</p>
														</div>
													</div>
												</td>
											</tr>
											</tbody>
										</table>
									</div>
								</div>
								<div class="themaInfo tMar50">
									<div class="overHidden bPad03">
										<h3 class="bltNo ftLt">5. ��ȹ�� �׸� ����</h3>
										<button type="button" class="ftRt" onClick="jsDownload('http://upload.10x10.co.kr','�׸����̵�_20171130.pdf','201712/201712041159331.pdf');"">�׸� ���̵�</button>
									</div>
									<table class="tbType1 writeTb">
										<colgroup>
											<col width="18%" /><col width="" />
										</colgroup>
										<tbody> 
										<tr>
											<th><div>�׸� ���� <strong class="cRd1">*</strong></div></th>
											<td>
												<span class="rMar10"><input type="radio" id="mdTm" name="mdTm" value="1" onClick="jsChTm('A');" <%if mdtheme="1" or isnull(mdtheme) then%>checked<%end if%>> <label for="txtTheme">�ؽ�Ʈ �׸�</label></span>
												<span class="rMar10"><input type="radio" id="mdTm" name="mdTm" value="2" onClick="jsChTm('B');" <%if mdtheme="2" then%>checked<%end if%>> <label for="imgTheme">�̹��� �׸�</label></span>
												<span class="rMar10"><input type="radio" id="mdTm"  name="mdTm" value="3" onClick="jsChTm('C');" <%if mdtheme="3" then%>checked<%end if%>> <label for="pdtTheme">��ǰ �׸�</label></span>
											</td>
										</tr>
										<tr>
											<th><div>�׸� ���� <strong class="cRd1">*</strong></div></th>
											<td id="pvpc"> 
												 <p class="tPad05 cBl2 fs11">-  TV ���α׷�, ������ ��� ���۱� ���� �ܾ�� �������ֽð�, �ٹ����� ������ ���� ������ �Է����ּ���.</p>
												<div id="dvTm" class="themaSetWrap type<%if mdtheme="2" then%>B <% If textbgcolor<>"1" Then %> typeBblack<% End If %><%elseif mdtheme="3" then%>C<%else%>A<%end if%>"><!-- for dev msg : �̺�Ʈ ������ ���� typeA(�ؽ�Ʈ �׸�), typeB(�̹��� �׸�-��ü�Ѹ�), typeC(��ǰ�׸�-�κзѸ�) Ŭ���� �־��ּ���. -->
													<div class="chPcWeb tMar30">
														<p><strong>[PC Web]</strong></p>
														<div class="fullTemplatev17" style="background-color:<%=fnEventColorCode(themecolor)%>;">
															<div class="fullContV17">
																<div class="txtCont">
																	<div class="inner">
																		<a href="<%=wwwUrl%>/street/street_brand_sub06.asp?makerid=<%=makerid%>" class="brandName arrow" target="_top"><%=brandNm%><i></i></a>
																		<p class="title"><textarea   id="evtNmW" name="evtNmW" class="formTxtA" style="width:95%; overflow:hidden; resize:none;" rows="2" maxlength="35"/><%=titlepc%></textarea></p>
																		<p class="subcopy"><textarea  name="subcopyK" id="subcopyK" placeholder="����ī�Ǹ� �Է����ּ���." class="formTxtA" style="width:95%; overflow:hidden; resize:none;" rows="2" maxlength="120" /><%=subcopyK%></textarea></p>
																		<%if isSale or isCoupon then%>
																		<div class="discount">
																			<%if isSale then%>
																			<span class="cRd0V15"><%=saleper%><!--<input type="text" value="<%=saleper%>" placeholder="30%" class="formTxt" style="width:22%" />--></span><!-- for dev msg : ��ǰ���� cRd0V15, �������� cGr0V15 Ŭ���� �־��ּ��� / ��ǰ���� �������� ���ÿ� �� ��� ���������� �տ� + �ٿ��ּ��� -->
																			<%end if %>
																			<%if isCoupon then%>
																			<span class="cGr0V15"><%if isSale then%>+<%end if%><%=salecper%><!--<input type="text" value="<%if isSale then%>+<%end if%><%=salecper%>" placeholder="+5%" class="formTxt" style="width:22%" />%--></span>
																			<%end if%>
																		</div>
																		<%end if%>
																		<p class="boxColorSlt tMar50" id="tbg">
																			<input type="hidden" name="tbgc" id="tbgc" value="1"> 
																			<span class="rMar10"><button type="button" class="colorChip" style="background-color:#fff" value="1"></button></span>
																			<span class="rMar10"><button type="button" class="colorChip" style="background-color:#000" value="2"></button></span>
																			<span class="rMar10"><button type="button" class="btn" value="" style="height:18px; padding:0 8px; vertical-align:top;" >�ڽ� ��� ���� ����</button></span>
																		</p>
																	</div>
																	
																	<p class="colorSlt" id="tmbg">
																		<input type="hidden" name="tmc" id="tmc" value="<%=themecolor%>"> 
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#848484"  value="11" ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#ed6c6c"  value="1" ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#f385af"  value="2" ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#f3a056"  value="3" ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#e7b93c"  value="4" ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#8eba4a"  value="5" ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#43a251"  value="6" ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#50bdd1"  value="7" ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#5aa5ea"  value="8" ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#2672bf"  value="9" ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#2c5a85"  value="10" ></button></span>
																		<span class="rMar10"><button type="button" class="btn" value="" style="height:18px; padding:0 8px; vertical-align:top;" onClick="jsSetTmColor('P');">���� ����</button></span>
																	</p>
																	<p class="pdtImgSch">
																		<button type="button" class="btn3 cGy1" onclick="jsRegItem('W');">�� ��ǰ ���</button>																		
																	</p>
																	<div class="imgSlt">
																		<div>
																			<button type="button" class="btn3" onClick="jsRegMultiImg('tms','1140','560');">��� �̹��� ���</button>
																			<span class="tPad05 fs11 txt">(�̹��� ������ : <b>1140x560</b>px)</span>
																		</div>
																		  
																	</div>  
																</div>
																
																<div class="slide">
																	<% dim imgw(3)
																	if isArray(arrimg) then
																			for intLoop =0 To uBound(arrimg,2)
																			 if mdtheme = "3" then
																		%>
																	<div><img src="<%=webImgUrl%>/image/basic/<%=GetImageSubFolderByItemid(arrimg(1,intLoop)) %>/<%=arrImg(0,intLoop)%>"  ></div>
																	<% else
																		imgw(intLoop) = arrImg(0,intLoop)
																		%>
																	<div><img src="<%=arrImg(0,intLoop)%>"  ></div>
																	<% end if
																			next
																		end if
																	%>
																</div>																
															</div>
														</div>
														<input type="hidden" name="hidtms1" id="hidtms1" value="<%=imgw(0)%>">
														<input type="hidden" name="hidtms2" id="hidtms2" value="<%=imgw(1)%>">
														<input type="hidden" name="hidtms3" id="hidtms3" value="<%=imgw(2)%>">
														<!-- ������ -->
														<!-- for dev msg : �� ������ ����÷� ��� -->
														<div class="pdtGroupBarV17" id="groupBar01" name="groupBar01" style="background-color:<%=fnEventBarColorCode(themecolor)%>;">
															<p>�׷�</p>
															<!-- �귣�� ��ũ�� ��������, �������� ����--><a href="" class="arrow btnBrand">�귣�� ��������<i></i></a>
														</div>
														<!--// ������ -->
													</div>
												 
												
													<div class="chMoApp tMar30">
														<p><strong>[Mobile]</strong></p>
														<div class="event-article">
															<section class="section-event hgroup-event"  id="mobg" style="background-color:<%=fnEventColorCode(themecolormo)%>;">
																<h2><textarea value="" id="evtNmM" name="evtNmM" class="formTxtA" style="width:100%; overflow:hidden; resize:none;" rows="2" maxlength="35"/><%=titlemo%></textarea></h2>
																<p class="subcopy"><textarea id="evtsubname" name="evtsubname"  placeholder="����ī�Ǹ� �Է����ּ���." class="formTxtA" style="width:100%; overflow:hidden; resize:none;" rows="2" maxlength="120"/><%=evtsubname%></textarea></p>
																<%if isSale or isCoupon then%>
																<div class="discount">
																	<%if isSale then%>
																	<b class="red rMar05"><span><%=saleper%><!--<input type="text" value="<%=saleper%>" placeholder="30%" class="formTxt" style="width:18%" /></span>--></b>
																	<%end if%>
																	<%if isCoupon then%>
																	<b class="green"><small>����</small><span><%=salecper%><!--<input type="text" value="<%if isSale then%>+<%end if%><%=salecper%>" placeholder="30%" class="formTxt" style="width:18%" />--></span></b>
																	<%end if%>
																</div>
																<%end if%>
																<div class="btnGroup"><a href="" class="btnV16a"><%=brandNm%></a></div>
																<p class="colorSlt tMar30" id="tmbgmo">
																		<input type="hidden" name="tmcmo" id="tmcmo" value="<%=themecolormo%>"> 
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#848484"  value="11" ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#ed6c6c"  value="1"  ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#f385af"  value="2"  ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#f3a056"  value="3"  ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#e7b93c"  value="4"  ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#8eba4a"  value="5"  ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#43a251"  value="6"  ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#50bdd1"  value="7"  ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#5aa5ea"  value="8"  ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#2672bf"  value="9"  ></button></span>
																		<span class="rMar10"><button type="button" class="colorChip" style="background-color:#2c5a85"  value="10" ></button></span>
																		<span class="rMar10"><button type="button" class="btn" value="" style="height:18px; padding:0 8px; vertical-align:top;" onClick="jsSetTmColor('M');">���� ����</button></span>																	
																</p>
																<p class="pdtImgSch tMar10">
																	<button type="button" class="btn3 cGy1" onclick="jsRegItem('M');">�� ��ǰ���</button>
																</p>
															</section>
															<!-- for dev msg : �ִ� 5�� -->
															<div id="mdRolling" class="swiper">
																<div class="swiper-container">
																	<div class="swiper-wrapper" id="tmsmImg">  	 
																			<%dim imgm(3)
																			if isArray(arrimgmo) then 
																					for intLoop = 0 To ubound(arrimgmo,2)
																						if mdtheme="3" then
																				%>	
																			<div class="swiper-slide">																				
																				<div class="thumbnail"><img src="<%=webImgUrl%>/image/basic/<%=GetImageSubFolderByItemid(arrimgmo(1,intLoop)) %>/<%=arrimgmo(0,intLoop)%>" /></div>																			
																			</div>
																			<%		else 
																						imgm(intLoop) = arrimgmo(0,intLoop)
																				%>
																			<div class="swiper-slide" id="tmsm<%=intLoop%>">																				
																				<div class="thumbnail"><img src="<%=arrimgmo(0,intLoop)%>"  /></div>																			
																			</div>
																		<%		end if 
																				next
																			end if
																			%> 
																	</div>
																	<div class="pagination-line"></div>
																	<button type="button" class="btnNav btnPrev">����</button>
																	<button type="button" class="btnNav btnNext">����</button>
																</div>
															</div>
															<div class="imgSlt">
																<div>
																	<button type="button" class="btn3" onClick="jsRegMultiImg('tmsm','750','528')">��� �̹��� ���</button>
																	<span class="tPad05 fs11 txt">(�̹��� ������ : <b>750x528</b>px)</span>
																</div> 
																
															</div> 
															 
														</div>
														<input type="hidden" name="hidtmsm1" id="hidtmsm1" value="<%=imgm(0)%>">
														<input type="hidden" name="hidtmsm2" id="hidtmsm2" value="<%=imgm(1)%>">
														<input type="hidden" name="hidtmsm3" id="hidtmsm3" value="<%=imgm(2)%>">
														<h3 class="groupBar" >
															<span id="groupBar01Mo" style="background-color:<%=fnEventBarColorCode(themecolormo)%>;"></span><b>�׷�</b>
														</h3>
													</div>
												</div>
											</td>
										</tr>
										</tbody>
									</table>

								</div>

								<div class="giftInfo tMar50">
									<h3 class="bltNo">6. GIFT �ȳ� ���� </h3>
									<table class="tbType1 writeTb tMar10">
										<colgroup>
											<col width="14%" /><col width="" />
										</colgroup>
										<tbody>
										<tr>
											<th><div>����ǰ ����</div></th>
											<td>
												<span class="rMar20">
													<select class="formSlt" id="gUsing" name="gUsing" title="����ǰ ���� ����">
														<option value="0" <%if giftisusing ="0" then%>selected<%end if%>>������</option>
														<option value="1" <%if giftisusing ="1" then%>selected<%end if%>>1�� ����ǰ</option>
														<option value="2" <%if giftisusing ="2" then%>selected<%end if%>>2�� ����ǰ</option>
														<option value="3" <%if giftisusing ="3" then%>selected<%end if%>>3�� ����ǰ</option>
													</select>
												</span>
											</td>
										</tr>
										<tr>
											<th rowspan="2"><div>GIFT1</div></th>
											<td>
												<span class="rMar20"><input type="text" class="formTxt" value="<%=gifttext1%>" name="gtext1" placeholder="��������ǰ(��Ȯ�� ��ǰ��) �� �����Ͻô� ���е鿡�� �������� ������ �帳�ϴ�." style="width:90%" /></span>
											</td>
										</tr>
										<tr>
											<td>
												<div class="inTbSet">
													<div class="formFile">
														<p>
															<button type="button" onClick="jsRegImg('g1','420','420','105');" class="btn">�̹��� ���</button><!--<input type="file" id="formFile" style="width:85%;" />-->
																<input type="hidden" name="hidg1" id="hidg1" value="<%=giftimg1%>"> 
														</p>
														<p class="tPad05 fs11 cGy1">- �̹��� ������ : <b>420x420</b>px</p>
													</div>
													<div style="width:105px;">
														<p class="registImg" id="g1Img">
															<%if  giftimg1 <> "" then %>
															<button type="button" onclick="jsDelimg('g1')">X</button>
															<img src="<%=giftimg1%>" alt="" style="width:105px;" />
															<%end if%>
														</p>
													</div>
												</div>
											</td>
										</tr>
										<tr>
											<th rowspan="2"><div>GIFT2</div></th>
											<td>
												<span class="rMar20"><input type="text" class="formTxt" value="<%=gifttext2%>" name="gtext2" placeholder="��������ǰ(��Ȯ�� ��ǰ��) �� �����Ͻô� ���е鿡�� �������� ������ �帳�ϴ�." style="width:90%" /></span>
											</td>
										</tr>
										<tr>
											<td>
												<div class="inTbSet">
													<div class="formFile">
														<p>
															<button type="button" onClick="jsRegImg('g2','420','420','105');" class="btn">�̹��� ���</button><!--<input type="file" id="formFile" style="width:85%;" />-->
																<input type="hidden" name="hidg2" id="hidg2" value="<%=giftImg2%>"> 
														</p>
														<p class="tPad05 fs11 cGy1">- �̹��� ������ : <b>420x420</b>px</p>
													</div>
													<div style="width:105px;">
														<p class="registImg" id="g2Img">
															<%if giftImg2 <> "" then%>
															<button type="button" onclick="jsDelimg('g2')">X</button>
															<img src="<%=giftImg2%>" alt="" style="width:105px;" />
															<%end if%>
														</p>
													</div>
												</div>
											</td>
										</tr>
										<tr>
											<th rowspan="2"><div>GIFT3</div></th>
											<td>
												<span class="rMar20"><input type="text" class="formTxt" value="<%=gifttext3%>"  name="gtext3" placeholder="��������ǰ(��Ȯ�� ��ǰ��) �� �����Ͻô� ���е鿡�� �������� ������ �帳�ϴ�." style="width:90%" /></span>
											</td>
										</tr>
										<tr>
											<td>
												<div class="inTbSet">
													<div class="formFile">
														<p>
															<button type="button" onClick="jsRegImg('g3','420','420','105');" class="btn">�̹��� ���</button><!--<input type="file" id="formFile" style="width:85%;" />-->
																<input type="hidden" name="hidg3" id="hidg3" value="<%=giftImg3%>"> 
														</p>
														<p class="tPad05 fs11 cGy1">- �̹��� ������ : <b>420x420</b>px</p>
													</div>
													<div style="width:105px;">
														<p class="registImg" id="g3Img">
															<%if giftImg3 <> "" then%>
															<button type="button" onclick="jsDelimg('g3')">X</button>
															<img src="<%=giftImg3%>" alt="" style="width:105px;" />
															<%end if%>
														</p>
													</div>
												</div>
											</td>
										</tr>
										</tbody>
									</table>
								</div>
							</form>		
								<div class="tMar30 ct">
									<input type="button" value="���" onClick="jsCancel('<%=menupos%>')" style="width:100px; height:30px;"   /> 
									<input type="button" value="�Ϸ�" onclick="jsRegEvent();" class="cRd1" style="width:100px; height:30px;" />
								</div>
							</div>
						</div>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<iframe name="FrameCKP" id="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe> 
<script>
$(function(){
//	var sampleTags = ['c++', 'java', 'php', 'coldfusion', 'javascript', 'asp', 'ruby', 'python', 'c', 'scala', 'groovy', 'haskell', 'perl', 'erlang', 'apl', 'cobol', 'go', 'lua', 'piece', '�ѱ�'];

	$('#singleFieldTags').tagit({
//		availableTags: sampleTags,
		placeholderText: "#���� �Է�"
//		Usage : https://github.com/aehlke/tag-it ����
//		autocomplete: {delay: 0, minLength: 2},
//		singleField: true,
//		singleFieldNode: $('#mySingleField')
		 
	});
});
</script>
<script type="text/javascript" src="/js/jquery.slides.min2.js"></script>
 <script >
 	//�׸�����
 function jsChTm(sType){		
		var pCNm = $("#dvTm").attr("class");	
		var nCNm = 	"themaSetWrap type"	+sType;	
		$("#dvTm").removeClass(pCNm);
			$(".slide div").remove();
		$("#mdRolling .swiper-container .swiper-slide div").remove();
		$("#dvTm").addClass(nCNm);
		 	
	 jsSetDisp(sType); 

}

		$(function(){ 
			
			$("#tmbg .colorChip").click(function() {																				
				var chipno = $(this).val() ;
				$("#tmc").val(chipno);
				});
			$("#tmbgmo .colorChip").click(function() {																			
				var chipno = $(this).val() ;
				$("#tmcmo").val(chipno);
				});	
			$("#tbg .colorChip").click(function() {																			
				var chipno = $(this).val() ;
				$("#tbgc").val(chipno);
			
				});	 	
			$("#tbg .btn").click(function() {																			
				var chipno = $("#tbgc").val() ;
				if (chipno==2){
					$("#dvTm").addClass("typeBblack")
				}else{
					$("#dvTm").removeClass("typeBblack")
				}		
			});	
				
																			 
		});
																			
																			
		function jsSetTmColor(sType){
			var chipno ;
			var chipcolor ="";
			
			if (sType=="M" ){
				 chipno =$("#tmcmo").val();
			}else{
				 chipno =$("#tmc").val();
			}	 
			
		 
			 if (chipno == 1 ){
				chipcolor = "#ed6c6c";      
				chipbar ="#cb4848"          
			}else if(chipno==2){          
				chipcolor = "#f385af";  
				chipbar= "#d55787"
			}else if(chipno==3){          
				chipcolor = "#f3a056";      
				chipbar="#e37f35"
			}else if(chipno==4){          
				chipcolor = "#e7b93c";      
				chipbar="#ce8d00"
			}else if(chipno==5){          
				chipcolor = "#8eba4a";      
				chipbar=     "#699426"
			}else if(chipno==6){          
				chipcolor = "#43a251";      
				chipbar= "#358240"
			}else if(chipno==7){          
				chipcolor = "#50bdd1";      
				chipbar=   "#2899ae"
			}else if(chipno==8){          
				chipcolor = "#5aa5ea";      
				chipbar= "#2f7cc3"
			}else if(chipno==9){          
				chipcolor = "#2672bf";      
				chipbar=   "#145290"
			}else if(chipno==10){         
				chipcolor = "#2c5a85";			
				chipbar=		 "#1c3e5d"	
			}else{                        
				chipcolor = "#848484";			
				chipbar="#656565"
			}                             
			 
			 if (sType=="M" ){
			 		$("#groupBar01Mo").css("background",chipbar);
			 		$("#mobg").css("background",chipcolor);
			}else{
				$("#groupBar01").css("background",chipbar);
				$(".fullTemplatev17").css("background",chipcolor);
		}
		}
		
		 
		
		</script> 
<script>
	function jsRollingbg(sType){  		
		if ($(".fullTemplatev17 .slide div").length > 1) {			
		$('.fullTemplatev17 .slide').slidesjs({
			pagination:{effect:'fade'},
			navigation:{effect:'fade'},
			play:{interval:3000, effect:'fade', auto:true},
			effect:{fade: {speed:800, crossfade:true}},
			callback: { 
				complete: function(number) {
					var pluginInstance = $('.fullTemplatev17 .slide').data('plugin_slidesjs');
					setTimeout(function() {
						pluginInstance.play(true);
					}, pluginInstance.options.play.interval);
				}
			}
		});
		jsSetDisp(sType);
	}

	}
 
	
	function jsRollingbgM(sType){
		
	/* rolling for md event */
	if ($("#mdRolling .swiper-container .swiper-slide").length > 1) {
		var mdRolling = new Swiper("#mdRolling .swiper-container", {
			pagination:"#mdRolling .pagination-line",
			paginationClickable:true,
			autoplay:1700,
			loop:true,
			speed:800,
			nextButton:"#mdRolling .btnNext",
			prevButton:"#mdRolling .btnPrev"
		});
	} else {
		var mdRolling = new Swiper("#mdRolling .swiper-container", {
			pagination:false,
			noSwipingClass:".noswiping",
			noSwiping:true
		});
	}

	$("#mdRolling .pagination-line").each(function(){
		var checkItem = $(this).children("span").length;
		if (checkItem == 2) {
			$(this).addClass("grid2");
		}
		if (checkItem == 3) {
			$(this).addClass("grid3");
		}
		if (checkItem == 4) {
			$(this).addClass("grid4");
		}
		if (checkItem == 5) {
			$(this).addClass("grid5");
		}
	});
		}
		
	function jsSetDisp(sType){
			if (sType=="B"){
				//var textW = $(".typeB .fullTemplatev17 .title").outerWidth();
				var textH = $(".typeB .fullTemplatev17 .inner").outerHeight()/2;
				var pgnW = $(".fullTemplatev17 .slide .slidesjs-pagination").outerWidth()/2;
				//$(".fullTemplatev17.typeB .inner").css("width",textW +160);
				$(".typeB .fullTemplatev17 .inner").css("margin-top",-textH);
				$(".typeB .fullTemplatev17 .slide .slidesjs-pagination").css("margin-left",-pgnW);
				$(".typeB .fullTemplatev17 .slidesjs-previous").css("margin-left",-pgnW);
				$(".typeB .fullTemplatev17 .slidesjs-next").css("margin-left",pgnW - 20);
		}else if (sType=="A"){
			var textH = 0;
			$(".typeA .fullTemplatev17 .inner").css("margin-top",-textH);
		}
	}
$(function(){
		if ($(".fullTemplatev17 .slide div").length > 1) {
		$('.fullTemplatev17 .slide').slidesjs({
			pagination:{effect:'fade'},
			navigation:{effect:'fade'},
			play:{interval:3000, effect:'fade', auto:true},
			effect:{fade: {speed:800, crossfade:true}},
			callback: {
				complete: function(number) {
					var pluginInstance = $('.fullTemplatev17 .slide').data('plugin_slidesjs');
					setTimeout(function() {
						pluginInstance.play(true);
					}, pluginInstance.options.play.interval);
				}
			}
		}); 
	}
   
	jsSetDisp('B');

	/* rolling for md event */
	if ($("#mdRolling .swiper-container .swiper-slide").length > 1) {
		var mdRolling = new Swiper("#mdRolling .swiper-container", {
			pagination:"#mdRolling .pagination-line",
			paginationClickable:true,
			autoplay:1700,
			loop:true,
			speed:800,
			nextButton:"#mdRolling .btnNext",
			prevButton:"#mdRolling .btnPrev"
		});
	} else {
		var mdRolling = new Swiper("#mdRolling .swiper-container", {
			pagination:false,
			noSwipingClass:".noswiping",
			noSwiping:true
		});
	}

	$("#mdRolling .pagination-line").each(function(){
		var checkItem = $(this).children("span").length;
		if (checkItem == 2) {
			$(this).addClass("grid2");
		}
		if (checkItem == 3) {
			$(this).addClass("grid3");
		}
		if (checkItem == 4) {
			$(this).addClass("grid4");
		}
		if (checkItem == 5) {
			$(this).addClass("grid5");
		}
	});
});

//
//$(".btnOdrChg").attr("onclick","").unbind("click");
//$(".btnOdrChg").on('click',function() {
//		alert("b");
//		 
//		if ($("#sortable").hasClass('sortable')) { 
//			$("#sortable").removeClass('sortable');  
//		 
//			$("#sortable li ").each(function(idx){
//				var i = parseInt(idx)+1;
//				$("#sortable li p:nth-child(1):eq("+idx+")").html(i);		
//				$("input[name^='eMGS']:eq("+idx+")").val(i);		  
//			}); 	
//			$("#sortable li.ui-state-disabled p:nth-child(1)").html($("#gCnt").val());
//			$("#sortable").sortable("destroy");
//			$(".btnOdrChg").attr("value", "�׷���� ����");
//			$(".infoTxt").hide();
//		} else {
//			$("#sortable").addClass('sortable');
//			
//			$("#sortable li ").each(function(idx){ 				
//				$("#sortable li p:nth-child(1):eq("+idx+")").html("<img src='/images/ico_odrchg.png' alt='�׷���� ����' />");					
//				});	 
//			//$("#sortable li p:nth-child(1)").html("<img src='/images/ico_odrchg.png' alt='�׷���� ����' />");
//			$("#sortable li.ui-state-disabled p:nth-child(1)").html("");
//			$("#sortable").sortable({
//				placeholder:"handling",
//				items:"li:not(.ui-state-disabled)"
//			}).disableSelection();
//			
//			$(".btnOdrChg").attr("value", "�׷���� ����Ϸ�"); 
//			$(".infoTxt").show();  
//			$(".btnOdrChg").on("click",function(){
//				alert("a");
//		 			jsProcGS();
//				});
//			
//		}
//	}); 
	
	
	$(".btnOdrChg").on('click',jsSetGS); 
		
	function jsSetGS(){
		$("#sortable").addClass('sortable');
		$("#sortable li ").each(function(idx){ 				                                                                                    
				$("#sortable li p:nth-child(1):eq("+idx+")").html("<img src='/images/ico_odrchg.png' alt='�׷���� ����' />");					        
		});	                                                                                                                            
		$("#sortable li.ui-state-disabled p:nth-child(1)").html("");                                                                      
		$("#sortable").sortable({                                                                                                         
			placeholder:"handling",                                                                                                         
			items:"li:not(.ui-state-disabled)"                                                                                              
		}).disableSelection();                                                                                                            
                                                                                                                                  
		$(".btnOdrChg").attr("value", "�׷���� ����Ϸ�");                                                                               
		$(".infoTxt").show();    
	
	  $(".btnOdrChg").off("click") ;                                                                                                       
		$(".btnOdrChg").on("click",function(){                                                                           
			jsProcGS();                                                                                                                   
		});                                                                                                                             
	}	
	
	function jsViewGS(){			
			$(".btnOdrChg").attr("value", "�׷���� ����");                              
			$(".infoTxt").hide();                                                        			
			$(".btnOdrChg").off("click") ;           
		 	$(".btnOdrChg").on('click',jsSetGS); 
	}
	
	
	function jsProcGS(){
			$("#sortable li ").each(function(idx){                                       
				var i = parseInt(idx)+1;                                                   
				$("#sortable li p:nth-child(1):eq("+idx+")").html(i);		                   
				$("input[name^='eMGS']:eq("+idx+")").val(i);		                           
			}); 	                                                                       
			$("#sortable li.ui-state-disabled p:nth-child(1)").html($("#gCnt").val());   
			document.frmReg.hidM.value="GS";
				var arrGC,arrGS ; 	
			 if (typeof(document.all.eMGC.length) !="undefined")	{ 				  
				for(var i=0;i< document.all.eMGC.length;i++) {
					if(i==0){
					 	arrGC = document.all.eMGC[i].value;
						arrGS	= document.all.eMGS[i].value;
					}else{
						arrGC = arrGC +"," + document.all.eMGC[i].value;
						arrGS	= arrGS +"," +  document.all.eMGS[i].value;
					}
				}
				document.frmReg.arrGC.value = arrGC
				document.frmReg.arrGS.value =arrGS
				document.frmReg.target = "FrameCKP";
				document.frmReg.submit();
			}else{
				jsViewGS();
			}
	}
</script>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
