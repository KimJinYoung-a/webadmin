<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_basicinfo.asp
' Discription : I��(������) �̺�Ʈ �⺻ ���� ��� â
' History : 2019.01.22 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<%
Dim cEvtCont, estimateSalePrice, salePer, saleCPer, ImgRegdate
Dim eCode, ekind, elevel, isWeb, isMobile, isApp, emdid, emdnm, eDateView
dim esale, egift, ecoupon, eonlyten, eOneplusone, eFreedelivery, eBookingsell
dim eDiary, eNew, ecomment, ebbs, eitemps, eisblogurl, ename, eusing, eman
dim eSdate, eEdate, ePdate, bannerTypeDiv, bannerCouponTxt, bannerGubun, marketing_event_kind
dim eEtcitemid, subcopyK, etag, eSalePer, evt_type, evt_kind, endlessView, eisort, eSTime, eETime
dim evt_startdate, evt_enddate
dim estate '// �̺�Ʈ ����

eCode = Request("eC")
ekind = Request("eK")

elevel = 2 '�߿䵵 �������� �ӽ� ����
isWeb = True
isMobile = True
isApp = True

if emdid = "" then 
    emdid = session("ssBctId")
    emdnm = session("ssBctCname")
end if

'// �ű� ��Ͻ� ����
if eCode = "" then 
	eSdate = date()
	'eEdate = date()
end if 
	
esale = False
egift= False
ecoupon= False
eonlyten= False
eOneplusone= False
eFreedelivery= False
eBookingsell= False
eDiary= False
eNew= False
ecomment = False
ebbs 	= False
eitemps	= False
eisblogurl = False



IF eCode <> "" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'�̺�Ʈ �ڵ�
	'�̺�Ʈ ���� ��������
	cEvtCont.fnGetEventCont
	ekind = cEvtCont.FEKind
    eman = cEvtCont.FEManager
	ename = db2html(cEvtCont.FEName) ' �̺�Ʈ��
	eusing = cEvtCont.FEUsing
	elevel = cEvtCont.FELevel
	isWeb = cEvtCont.FIsWeb
	isMobile = cEvtCont.FIsMobile
	isApp = cEvtCont.FIsApp
	evt_startdate = cEvtCont.FESDay
	evt_enddate = cEvtCont.FEEDay
	eSTime = Hour(cEvtCont.FESDay)
	eETime = Hour(cEvtCont.FEEDay)
    eSdate = left(cEvtCont.FESDay,10)
    eEdate = left(cEvtCont.FEEDay,10)
    ePdate = left(cEvtCont.FEPDay,10)
	subcopyK = db2html(cEvtCont.FsubcopyK) '����ī�� �ѱ� PC
	estate = cEvtCont.FEState
	ImgRegdate = cEvtCont.FEImgRegdate
	IF datediff("d",now,eEdate) <0 THEN estate = 9 '�Ⱓ �ʰ��� ����ǥ��
	if eETime="" then eETime=23
	if ekind = 19 then
	    isWeb = False
	    isMobile = True
	    isApp = True
	    ekind = 1
	elseif ekind = 25 then
	    isWeb = False
	    isMobile = False
	    isApp = True
	    ekind = 1
	elseif ekind = 26 then
	    isWeb = False
	    isMobile = True
	    isApp = False
	    ekind = 1
	elseif not (isWeb  or  isMobile  or isApp) or (isNull(isWeb) and isNull(isMobile) and isNull(isApp))  then 
	    isWeb = True
	    isMobile = False
	    isApp = False    
	    ekind = 1
    end if        
	      
	'�̺�Ʈ ȭ�鼳�� ���� ��������
	cEvtCont.fnGetEventDisplay
	esale 		= 	cEvtCont.FESale
	egift 		=	cEvtCont.FEGift
	ecoupon 	=	cEvtCont.FECoupon
	ecomment 	=	cEvtCont.FECommnet
	ebbs 		=	cEvtCont.FEBbs
	eitemps	 	=	cEvtCont.FEItemps
 	eonlyten	= cEvtCont.FSisOnlyTen
 	eDiary		= cEvtCont.FSisDiary
 	eNew			= cEvtCont.FSisNew
 	eisblogurl	= cEvtCont.FSisGetBlogURL
	eOneplusone		= cEvtCont.FEOneplusOne
	eFreedelivery	= cEvtCont.FEFreedelivery
	eBookingsell	= cEvtCont.FEBookingsell
	eDateView		= cEvtCont.FEdateview
	bannerTypeDiv = cEvtCont.FbannerTypeDiv
	bannerCouponTxt = cEvtCont.FbannerCouponTxt
	bannerGubun = cEvtCont.FbannerGubun
	eEtcitemid		=	cEvtCont.FEtcitemid
	etag = db2html(cEvtCont.FETag)
	If bannerGubun="" Then bannerGubun=1
	evt_type = cEvtCont.Feventtype_pc
    evt_kind = cEvtCont.Feventtype_mo
	endlessView = cEvtCont.FendlessView
	estimateSalePrice = cEvtCont.FestimateSalePrice
	eisort = cEvtCont.FEISort
	salePer = cEvtCont.FsalePer
	saleCPer = cEvtCont.FsaleCPer
	marketing_event_kind = cEvtCont.Fmarketing_event_kind
	set cEvtCont = nothing 

	if estate = "6" or estate = "7" then
		if evt_enddate < now() then
			estate = 9
		end if
	end if

	if (ekind = 1 or ekind = 23) and (eSale or ecoupon) then
	    dim tmpename
	    tmpename = Split(ename,"|") 
	  			 
	  	if Ubound(tmpename)>0 then
		    ename = tmpename(0)
		    eSalePer = tmpename(1)
		 end if

    end if
	if ekind=5 then estimateSalePrice=0
else
    eman=1
    eusing="Y"
    ekind=1
	eDateView = true
end if 

dim idepartmentid, sdepartmentname,clsMem
'�μ��� ��������
set clsMem = new CTenByTenMember
clsMem.Fuserid = emdid
clsMem.fnGetDepartmentInfo
idepartmentid		= clsMem.Fdepartment_id
sdepartmentname = clsMem.FdepartmentNameFull 
set clsMem = nothing
	 
%>
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";

function jsEvtSubmit(frm){
    //ä�μ��� ���� Ȯ��
    if (!frm.blnWeb.checked&&!frm.blnMobile.checked&&!frm.blnApp.checked){
        alert("ä���� �������ּ���");
        frm.blnWeb.focus();
        return false;
    }

    if(!frm.sEN.value){
        alert("�̺�Ʈ���� �Է����ּ���");
        frm.sEN.focus();
        return false;
    }

	if(frm.endlessview.value!="Y"){
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

		if(frm.eventkind.value!=5){
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
		}
	}
	


    if((frm.chComm.checked||frm.chBbs.checked||frm.chItemps.checked||frm.isblogurl.checked)&&frm.sPD.value=="") {
        alert("��÷�� ��ǥ���� �������ּ��� ");
        frm.sPD.focus();
        return false;
    }

    if(frm.imod.value=="PI" || frm.imod.value=="PU"){
		if(frm.etcitemid.value=="PI"){
			alert("��ǥ��ǰ �ڵ带 �Է��� �ּ���.");
			frm.etcitemid.focus();
			return false;
		}
		if(frm.subcopyK.value==""){
			alert("����ī�Ǹ� �Է��� �ּ���.");
			frm.subcopyK.focus();
			return false;
		}		
    }
	if(!$.isNumeric($("#esp").val())){
		alert("�ݾ׿� ���ڸ� �Է����ּ���")
		$("#esp").val(0);
		return false;
	}

    frm.action="basicinfo_process.asp";
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

function fnItemEventSet(div){
    if(div=="13"){
        $("#itemtypediv").show();
		$("#bannergubun").show();
		$("#masteritemcode").show();
		$("#tag").show();
        $("#etype").hide();
		$("#efunction").hide();
		$("#eprize").hide();
		$("#elevel").show();
		$("#eitemsort").show();
		$("#EstimateSalePrice").show();
		<% if eCode>"0" then %>
		$("#imod").val("PU");
		<% else %>
		$("#imod").val("PI");
		<% end if %>
		$("#evt_type").hide();
		$("#evt_kind").hide();
		$("#agent").show();
		$("#etcitemiddiv").show();
		$("#pickupdiv").hide();
    }
    else{
        $("#itemtypediv").hide();
		$("#bannergubun").hide();
		$("#masteritemcode").hide();
		$("#tag").hide();
		$("#efunction").show();
		$("#eprize").show();
		$("#etcitemiddiv").hide();
		if(div=="28"){
			$("#pickupdiv").show();
		}
		else{
			$("#pickupdiv").hide();
		}
		if(div=="5"){
			$("#chItemps").hide();
			$("#chBbs").hide();
			$("#isblogurl").hide();
			$("#etype").hide();
			$("#elevel").hide();
			$("#EstimateSalePrice").hide();
			$("#evt_type").show();
			$("#evt_kind").show();
			$("#agent").hide();
			$("#eitemsort").hide();
		}
		else{
			$("#chItemps").show();
			$("#chBbs").show();
			$("#isblogurl").show();
			$("#etype").show();
			$("#elevel").show();
			$("#EstimateSalePrice").show();
			$("#evt_type").hide();
			$("#evt_kind").hide();
			$("#agent").show();
			$("#eitemsort").show();
		}
		<% if eCode>"0" then %>
		$("#imod").val("BU");
		<% else %>
		$("#imod").val("BI");
		<% end if %>
    }
}

function jsAddByte(target,obj){ 
    var realText = obj.value; 
    var textBit = '';
    var textLen = 0;
    for (var i = 0 ; i < realText.length ; i++) {
        textBit = realText.charAt(i); 
        if(escape(textBit).length > 4) {
            textLen = textLen + 2;
        } else {
            textLen = textLen + 1;
        }

        if (textLen >= 140){
            realText = realText.substr(0,i);
            obj.value = realText;
            break;
        }
    }
	if(target=="1"){
		$("#etk").html(textLen);
	}
	else{
		$("#sck").html(textLen);
	}
}

function TnFavSearchTxt(){
    var winpop = window.open("http://61.252.133.17:5601/app/kibana#/dashboard/5c9d9970-ef60-11e6-9fb4-f3d99fd9206d?_g=(refreshInterval:(display:Off,pause:!f,value:0),time:(from:now-5h%2Fh,mode:quick,to:now))&_a=(filters:!(),options:(darkTheme:!f),panels:!((col:1,id:ca566510-ef5f-11e6-9fb4-f3d99fd9206d,panelIndex:1,row:1,size_x:3,size_y:5,type:visualization),(col:1,id:'%EC%9D%B8%EA%B8%B0%EA%B2%80%EC%83%89%EC%96%B4(MOB)',panelIndex:2,row:6,size_x:3,size_y:5,type:visualization),(col:1,id:'%EC%9D%B8%EA%B8%B0%EA%B2%80%EC%83%89%EC%96%B4(APP)',panelIndex:3,row:11,size_x:3,size_y:5,type:visualization),(col:4,id:'%EC%9D%B8%EA%B8%B0%EA%B2%80%EC%83%89%EC%96%B4-%EC%8B%9C%EA%B0%84%EB%8C%80%EB%B3%84(MOB)',panelIndex:4,row:6,size_x:9,size_y:5,type:visualization),(col:4,id:d06ee1e0-ef62-11e6-9fb4-f3d99fd9206d,panelIndex:5,row:1,size_x:9,size_y:5,type:visualization),(col:4,id:c7604a10-1aa2-11e7-b3b2-cb4977e75f0e,panelIndex:6,row:11,size_x:9,size_y:5,type:visualization)),query:(query_string:(analyze_wildcard:!t,query:'*')),title:'0005.%20%EC%9D%B8%EA%B8%B0%EA%B2%80%EC%83%89%EC%96%B4',uiState:(P-1:(vis:(params:(sort:(columnIndex:!n,direction:!n)))),P-2:(vis:(params:(sort:(columnIndex:!n,direction:!n)))),P-3:(vis:(params:(sort:(columnIndex:!n,direction:!n))))))",'winpop2','width=1450,height=800,scrollbars=yes,resizable=yes');
    winpop.focus();
}

function jsAddByte2(obj){ 
    var realText = obj.value; 
    var textBit = '';
    var textLen = 0;
    for (var i = 0 ; i < realText.length ; i++) {
        textBit = realText.charAt(i); 
        if(escape(textBit).length > 4) {
            textLen = textLen + 2;
        } else {
            textLen = textLen + 1;
        }

        if (textLen >= 500){
            realText = realText.substr(0,i);
            obj.value = realText;
            break;
        }
    }
	$("#Tag").html(textLen);
}

function fnDateViewSet(objval){
	if(objval=="Y"){
		$("#dateset").hide();
	}
	else{
		$("#dateset").show();
	}
}

function fnshowSale(div){
    if(div=="S"){
		if($("#chSale").is(":checked")){
			$("#salediv").show();
		}else{
			$("#salediv").hide();
			$("#sSP").val("");
		}
    }else{
		if($("#chCoupon").is(":checked")){
			$("#coupondiv").show();
		}else{
			$("#coupondiv").hide();
			$("#sCSP").val("");
		}
	}
}
// �̺�Ʈ ��ǰ �ִ� ������ ����
function fnGetMaxSalevalue(saildiv) {
	var evtcd = document.frmEvt.evt_code.value;
	$.ajax({
		type: "POST",
		url: "ajaxGetEvtMaxItemSalePer.asp",
		data: "eC="+evtcd+"&saildiv="+saildiv,
		cache: false,
		success: function(message) {
			if(message) {
				if(saildiv=="S"){
					$("#sSP").val(message);
				}else{
					$("#sCSP").val(message);
				}
			} else {
				alert("�̺�Ʈ�� ��ǰ�� ���ų� �������� ��ǰ�� �����ϴ�.");
			}
		},
		error: function(err) {
			alert(err.responseText);
		}
	});
}

// �̺�Ʈ ��ϴ�� ����
function fnSetEventStateEdit() {
	$.ajax({
		type: "POST",
		url: "ajaxSetEventStateEdit.asp",
		data: "eC=<%=eCode%>",
		cache: false,
		dataType: "JSON",
		success: function(data){
			if(data.response == "OK"){
				location.reload();
				alert(data.message);
			}else if(data.response == "err"){
				alert(data.message);
			}
		},
		error: function(data){
			alert('�ý��� �����Դϴ�.');
		}
	});
}
</script>

<form name="frmEvt" method="post" style="margin:0px;">
<% if eCode>"0" then %>
<% if ekind="13" then %>
<input type="hidden" name="imod" id="imod" value="PU">
<% else %>
<input type="hidden" name="imod" id="imod" value="BU">
<% end if %>
<input type="hidden" name="evt_code" value="<%=eCode%>">
<% else %>
<% if ekind="13" then %>
<input type="hidden" name="imod" id="imod" value="PI">
<% else %>
<input type="hidden" name="imod" id="imod" value="BI">
<% end if %>
<% end if %>
<input type="hidden" name="eIRD" id="imod" value="<%=ImgRegdate%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>��ȹ�� �⺻ ����</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
				<tr>
					<th>����</th>
					<td>
						<% getNewEventKindList "eventkind", ekind, False, True %>
          			</td>
				</tr>

				<tr id="evt_type" style="display:<% If ekind<>"5" Then Response.write "none"%>">
					<th>����</th>
					<td>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="evt_type" value="0"<% if evt_type="0" then %> checked<% end if %>>
								������
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="evt_type" value="1"<% if evt_type="1" then %> checked<% end if %>>
								�о��
								<i class="inputHelper"></i>
							</label>
						</div>
					</td>
				</tr>
				<tr id="evt_kind" style="display:<% If ekind<>"5" Then Response.write "none"%>">
					<th>���Ľ����̼� ����</th>
					<td>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="evt_kind" value="0"<% if evt_kind="0" then %> checked<% end if %>>
								��ȭ
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="evt_kind" value="1"<% if evt_kind="1" then %> checked<% end if %>>
								����
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="evt_kind" value="2"<% if evt_kind="2" then %> checked<% end if %>>
								����
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="evt_kind" value="3"<% if evt_kind="3" then %> checked<% end if %>>
								������
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="evt_kind" value="4"<% if evt_kind="4" then %> checked<% end if %>>
								����
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="evt_kind" value="5"<% if evt_kind="5" then %> checked<% end if %>>
								����
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="evt_kind" value="6"<% if evt_kind="6" then %> checked<% end if %>>
								������
								<i class="inputHelper"></i>
							</label>
						</div>
					</td>
				</tr>

				<tr>
					<th>ä��</th>
					<td>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="blnWeb" value="1"<% if isWeb then %> checked<% end if %>>
								PC
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="blnMobile" value="1" <% if isMobile  then %> checked<% end if %>>
								Mobile
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="blnApp" value="1" <% if isApp  then %> checked<% end if %>>
								APP
								<i class="inputHelper"></i>
							</label>
						</div>
					</td>
				</tr>
				<tr>
					<th>����ī��(��ȹ����)</th>
					<td>
						<input type="text" class="formControl formControl550" placeholder="�Է��ϼ���." name="sEN" id="sEN" maxlength="120" value="<%=ename%>" OnKeyUp="jsAddByte('1',this);">
						<span class="lMar05 cGy1 fs12 vBtm"><span class="cPk2 vBtm" id="etk">50</span><span class="cPk2 vBtm">byte</span>&#47;120byte</span>
						<script type="text/javascript">
							jsAddByte('1',frmEvt.sEN);
						</script>
					</td>
				</tr>
				<tr>
					<th>����ī��<% If ekind="5" Then %>(��÷����)<% end if %></th>
					<td>
						<input type="text" class="formControl formControl550" placeholder="�Է��ϼ���." name="subcopyK" id="subcopyK" maxlength="120" value="<%=subcopyK%>" OnKeyUp="jsAddByte('2',this);">
						<span class="lMar05 cGy1 fs12 vBtm"><span class="cPk2 vBtm" id="sck">50</span><span class="cPk2 vBtm">byte</span>&#47;120byte</span>
						<script type="text/javascript">
							jsAddByte('2',frmEvt.subcopyK);
						</script>
					</td>
				</tr>
				<tr>
					<th>�Ⱓ</th>
					<td>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="endlessview" id="radio4a" value="N"<%IF endlessView="N" THEN%> checked<%END IF%> onClick="fnDateViewSet(this.value);">
								�Ⱓ����
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="endlessview" id="radio4b" value="Y"<%IF endlessView="Y" THEN%> checked<%END IF%> onClick="fnDateViewSet(this.value);">
								��ó���
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="dateview" value="1"<%IF eDateView THEN%> checked<%END IF%>>
								�̺�Ʈ�Ⱓ ���� ����
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="tMar15 tPad15 topLine" id="dateset" style="display:<% If endlessView="Y" Then Response.write "none"%>">
							<div class="formInline"><span class="datepicker">������ <span class="mdi mdi-calendar-month"></span> <input type="text" id="datepicker1" name="sSD" placeholder="���ó�¥" value="<%=eSdate%>" readonly="true"> 
							<select name="sST" class="formControl">
								<option value="00" <% if eSTime="0" then response.write "selected" %>>00</option>
								<option value="01" <% if eSTime="1" then response.write "selected" %>>01</option>
								<option value="02" <% if eSTime="2" then response.write "selected" %>>02</option>
								<option value="03" <% if eSTime="3" then response.write "selected" %>>03</option>
								<option value="04" <% if eSTime="4" then response.write "selected" %>>04</option>
								<option value="05" <% if eSTime="5" then response.write "selected" %>>05</option>
								<option value="06" <% if eSTime="6" then response.write "selected" %>>06</option>
								<option value="07" <% if eSTime="7" then response.write "selected" %>>07</option>
								<option value="08" <% if eSTime="8" then response.write "selected" %>>08</option>
								<option value="09" <% if eSTime="9" then response.write "selected" %>>09</option>
								<option value="10" <% if eSTime="10" then response.write "selected" %>>10</option>
								<option value="11" <% if eSTime="11" then response.write "selected" %>>11</option>
								<option value="12" <% if eSTime="12" then response.write "selected" %>>12</option>
								<option value="13" <% if eSTime="13" then response.write "selected" %>>13</option>
								<option value="14" <% if eSTime="14" then response.write "selected" %>>14</option>
								<option value="15" <% if eSTime="15" then response.write "selected" %>>15</option>
								<option value="16" <% if eSTime="16" then response.write "selected" %>>16</option>
								<option value="17" <% if eSTime="17" then response.write "selected" %>>17</option>
								<option value="18" <% if eSTime="18" then response.write "selected" %>>18</option>
								<option value="19" <% if eSTime="19" then response.write "selected" %>>19</option>
								<option value="20" <% if eSTime="20" then response.write "selected" %>>20</option>
								<option value="21" <% if eSTime="21" then response.write "selected" %>>21</option>
								<option value="22" <% if eSTime="22" then response.write "selected" %>>22</option>
								<option value="23" <% if eSTime="23" then response.write "selected" %>>23</option>
							</select>
							</span></div>
							<div class="formInline"><span class="datepicker">������ <span class="mdi mdi-calendar-month"></span> <input type="text" id="datepicker2" name="sED" placeholder="���ó�¥" value="<%=eEdate%>" readonly="true">
							<select name="sET" class="formControl">
								<option value="00" <% if eETime="0" then response.write "selected" %>>00</option>
								<option value="01" <% if eETime="1" then response.write "selected" %>>01</option>
								<option value="02" <% if eETime="2" then response.write "selected" %>>02</option>
								<option value="03" <% if eETime="3" then response.write "selected" %>>03</option>
								<option value="04" <% if eETime="4" then response.write "selected" %>>04</option>
								<option value="05" <% if eETime="5" then response.write "selected" %>>05</option>
								<option value="06" <% if eETime="6" then response.write "selected" %>>06</option>
								<option value="07" <% if eETime="7" then response.write "selected" %>>07</option>
								<option value="08" <% if eETime="8" then response.write "selected" %>>08</option>
								<option value="09" <% if eETime="9" then response.write "selected" %>>09</option>
								<option value="10" <% if eETime="10" then response.write "selected" %>>10</option>
								<option value="11" <% if eETime="11" then response.write "selected" %>>11</option>
								<option value="12" <% if eETime="12" then response.write "selected" %>>12</option>
								<option value="13" <% if eETime="13" then response.write "selected" %>>13</option>
								<option value="14" <% if eETime="14" then response.write "selected" %>>14</option>
								<option value="15" <% if eETime="15" then response.write "selected" %>>15</option>
								<option value="16" <% if eETime="16" then response.write "selected" %>>16</option>
								<option value="17" <% if eETime="17" then response.write "selected" %>>17</option>
								<option value="18" <% if eETime="18" then response.write "selected" %>>18</option>
								<option value="19" <% if eETime="19" then response.write "selected" %>>19</option>
								<option value="20" <% if eETime="20" then response.write "selected" %>>20</option>
								<option value="21" <% if eETime="21" then response.write "selected" %>>21</option>
								<option value="22" <% if eETime="22" then response.write "selected" %>>22</option>
								<option value="23" <% if eETime="23" then response.write "selected" %>>23</option>
							</select>
							</span></div>
							<p class="tMar15 cPk2 fs12">���� �ð� ������ ������ �ð�(59��59��)���� ���� �˴ϴ�.</p>
						</div>
					</td>
				</tr>
				<tr id="etype" style="display:<% If ekind="5" Then Response.write "none"%>">
					<th>Ÿ��</th>
					<td>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chSale" id="chSale" value="1" onclick="fnshowSale('S');"<% if esale then %> checked<% end if %>>
								����
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chCoupon" id="chCoupon" value="1" onclick="fnshowSale('C');"<% if ecoupon then %> checked<% end if %>>
								����
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chOnlyTen" value="1"<% if eonlyten then %> checked<% end if %>>
								�ܵ�
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chGift" value="1"<% if egift then %> checked<% end if %>>
								GIFT
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chOneplusone" value="1"<% if eOneplusone then %> checked<% end if %>>
								1+1
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chNew" value="1"<% if eNew then %> checked<% end if %>>
								��Ī
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chFreedelivery" value="1"<% if eFreedelivery then %> checked<% end if %>>
								������
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chBookingsell" value="1"<% if eBookingsell then %> checked<% end if %>>
								�����Ǹ�
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chDiary" value="1"<% if eDiary then %> checked<% end if %>>
								DiaryStory
								<i class="inputHelper"></i>
							</label>
						</div>
					</td>
				</tr>
				<tr id="salediv" style="display:<% If not(esale) Then Response.write "none"%>">
					<th>��ǰ ������</th>
					<td>
						<div class="formInline">
							<input type="text" name="sSP" id="sSP" class="formControl formControl150" placeholder="��ǰ ������ �Է�" maxlength="2" value="<%=salePer%>">
							<% if eCode <> "" then %>
							<button class="btn4 btnBlue1 lMar05" onclick="fnGetMaxSalevalue('S');return false;">�ִ밪 ��������</button>
							<% end if %>
						</div>
					</td>
				</tr>
				<tr id="coupondiv" style="display:<% If not(ecoupon) Then Response.write "none"%>">
					<th>���� ������</th>
					<td>
						<div class="formInline">
							<input type="text" name="sCSP" id="sCSP" class="formControl formControl150" placeholder="���� ������ �Է�" maxlength="2" value="<%=saleCPer%>">
							<% if eCode <> "" then %>
							<button class="btn4 btnBlue1 lMar05" onclick="fnGetMaxSalevalue('C');return false;">�ִ밪 ��������</button>
							<% end if %>
						</div>
					</td>
				</tr>
				<tr id="efunction" style="display:<% If ekind="13" Then Response.write "none"%>">
					<th>���</th>
					<td>
						<div class="formInline" id="chItemps" style="display:<% If ekind="5" Then Response.write "none"%>">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chItemps" id="function1" value="1"<% if eitemps then %> checked<% end if %>>
								��ǰ�ı�
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chComm" id="function1" value="1"<% if ecomment then %> checked<% end if %>>
								�ڸ�Ʈ
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline" id="chBbs" style="display:<% If ekind="5" Then Response.write "none"%>">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chBbs" id="function1" value="1"<% if ebbs then %> checked<% end if %>>
								�����ڸ�Ʈ
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline" id="isblogurl" style="display:<% If ekind="5" Then Response.write "none"%>">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="isblogurl" id="function1" value="1"<% if eisblogurl then %> checked<% end if %>>
								blog URL
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="tMar15 tPad15 topLine">
							<div class="formInline"><span class="datepicker">��÷��ǥ�� <span class="mdi mdi-calendar-month"></span> <input type="text" id="datepicker3" placeholder="��¥����" readonly="true" name="sPD" value="<%=ePdate%>"></span></div>
							<p class="tMar15 cPk2 fs12">��÷ ��ǥ���� �ѹ��� ������ �� �ֽ��ϴ�.</p>
						</div>
					</td>
				</tr>
				<tr id="itemtypediv" style="display:<% If ekind<>"13" Then Response.write "none"%>">
					<th>��������</th>
					<td>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="bannerTypeDiv" value="1"<%IF bannerTypeDiv="1" THEN%> checked<%END IF%>>
								����
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="bannerTypeDiv" value="2"<%IF bannerTypeDiv="2" THEN%> checked<%END IF%>>
								���� 
								<i class="inputHelper"></i>
							</label><input type="text" name="bannerCouponTxt" value="<%=bannerCouponTxt%>" size="5">
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="bannerTypeDiv" value="3"<%IF bannerTypeDiv="3" THEN%> checked<%END IF%>>
								GIFT
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="bannerTypeDiv" value="4"<%IF bannerTypeDiv="4" THEN%> checked<%END IF%>>
								������
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="bannerTypeDiv" value="5"<%IF bannerTypeDiv="5" THEN%> checked<%END IF%>>
								������
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="bannerTypeDiv" value="6"<%IF bannerTypeDiv="6" THEN%> checked<%END IF%>>
								1:1
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="bannerTypeDiv" value="7"<%IF bannerTypeDiv="7" THEN%> checked<%END IF%>>
								1+1
								<i class="inputHelper"></i>
							</label>
						</div>
					</td>
				</tr>
				<tr id="bannergubun" style="display:<% If ekind<>"13" Then Response.write "none"%>">
					<th>������</th>
					<td>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="bannerGubun" id="radio5a" value="1"<% if bannerGubun="1" then %> checked<% end if %>>
								����ī��
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="bannerGubun" id="radio5b" value="2"<% if bannerGubun="2" then %> checked<% end if %>>
								����ǰ
								<i class="inputHelper"></i>
							</label>
						</div>
					</td>
				</tr>
				<tr id="etcitemiddiv" style="display:<% If ekind<>"13" Then Response.write "none"%>">
					<th>��ǥ��ǰ</th>
					<td>
						<div class="formInline"><input type="text" name="etcitemid" class="formControl formControl550" placeholder="��ǥ��ǰ �ڵ� �Է�" maxlength="16" value="<%=eEtcitemid%>"></div>
					</td>
				</tr>
				<tr id="tag" style="display:<% If ekind<>"13" Then Response.write "none"%>">
					<th>Tag</th>
					<td>
						<textarea name="eTag" rows="5" cols="50" placeholder="�α��±׸� �Է��غ��ƿ� :)" OnKeyUp="jsAddByte2(this);"><%=etag%></textarea>
						<p class="ftLt tMar20 cGy1 fs12"><span class="cPk2 vBtm" id="Tag">50</span><span class="cPk2 vBtm">byte</span>/500byte</p>
						<button class="ftRt btn4 btnBlue1 tMar10" onclick="TnFavSearchTxt()">�ǽð� �α� �˻��� ����</button>
						<script type="text/javascript">
							jsAddByte2(frmEvt.eTag);
						</script>
					</td>
				</tr>
				<tr>
					<th>�۾�����</th>
					<td>
						<% sbGetOptStatusCodeSort "eventstate", estate, false, "" %> <% if estate=9 then %><button class="btn4 btnBlue1 lMar05" onclick="fnSetEventStateEdit();return false;">��ϴ�� ����</button><% end if %>
					</td>
				</tr>
				<tr id="eitemsort" style="display:<% If ekind="5" Then Response.write "none"%>">
					<th>��ǰ���Ĺ��</th>
					<td>
						<%sbGetOptEventCodeValue "itemsort",eisort,False,""%>
					</td>
				</tr>
				<tr id="elevel" style="display:<% If ekind="5" Then Response.write "none"%>">
					<th>�߿䵵</th>
					<td>
						<%sbGetOptEventCodeValue "eventlevel",elevel,False,""%>
					</td>
				</tr>
				<tr id="EstimateSalePrice" style="display:<% If ekind="5" Then Response.write "none"%>">
					<th>��������</th>
					<td>
						<div class="formInline"><input type="text" name="estimateSalePrice" id="esp" class="formControl formControl550" placeholder="�������� �Է�" maxlength="16" value="<%=estimateSalePrice%>"></div>
					</td>
				</tr>
				<tr id="agent" style="display:<% If ekind="5" Then Response.write "none"%>">
					<th>��ü</th>
					<td>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="eventmanager" id="radio5a" value="1"<% if eman="1" then %> checked<% end if %>>
								10X10
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="eventmanager" id="radio5b" value="2"<% if eman="2" then %> checked<% end if %>>
								��ü
								<i class="inputHelper"></i>
							</label>
						</div>
					</td>
				</tr>
				<tr>
					<th>�������</th>
					<td>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="using" id="radio6a" value="Y"<% if eusing="Y" then %> checked<% end if %>>
								Y
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="using" id="radio6b" value="N"<% if eusing="N" then %> checked<% end if %>>
								N
								<i class="inputHelper"></i>
							</label>
						</div>
					</td>
				</tr>
				<tr id="pickupdiv" style="display:<% If ekind<>"28" Then Response.write "none"%>">
					<th>������ �̺�Ʈ ����</th>
					<td>
					    <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="marketing_event_kind" id="radio7a" value="0" <% if marketing_event_kind=0 or isnull(marketing_event_kind) then %> checked<% end if %>>
                                ����
                                <i class="inputHelper"></i>
                            </label>
                        </div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="marketing_event_kind" id="radio7a" value="1" <% if marketing_event_kind=1 then %> checked<% end if %>>
								����
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="marketing_event_kind" id="radio7b" value="2" <% if marketing_event_kind=2  then %> checked<% end if %>>
								�⼮üũ
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="marketing_event_kind" id="radio7b" value="3" <% if marketing_event_kind=3  then %> checked<% end if %>>
								�α��� ���ϸ���
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="marketing_event_kind" id="radio7b" value="4" <% if marketing_event_kind=4  then %> checked<% end if %>>
								��ǰ ����
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="marketing_event_kind" id="radio7b" value="5" <% if marketing_event_kind=5  then %> checked<% end if %>>
								������ ������
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="marketing_event_kind" id="radio7b" value="6" <% if marketing_event_kind=6  then %> checked<% end if %>>
								����Ǽ�
								<i class="inputHelper"></i>
							</label>
						</div>
					</td>
				</tr>
			</tbody>
		</table>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">���</button>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">����</button>
	</div>
</div>
</form>
<script>
<% if eCode ="" then %>
$(function() {
	$("select[name='eventlevel']").val("3").attr("selected","selected");
});
<% end if %>
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->