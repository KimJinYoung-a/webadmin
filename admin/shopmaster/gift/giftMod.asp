<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����ǰ ����
' History : 2008.04.01 ������ ����
'			2020.04.08 �ѿ�� ����
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
	iSerachType    = requestCheckVar(Request("selType"),4)		'�˻�����
	sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'�˻���
	sBrand     	= requestCheckVar(Request("ebrand"),32)		'�귣��
	sDate     		= requestCheckVar(Request("selDate"),1)		'�˻��� ����
	sSdate     	= requestCheckVar(Request("iSD"),10)		'������
	sEdate     	= requestCheckVar(Request("iED"),10)		'������
	igStatus		= requestCheckVar(Request("giftstatus"),4)	'����ǰ ����

	iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ

	strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&ebrand="&sBrand&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&gstatus="&igStatus

IF gCode = "" THEN
	Alert_return("���԰�ο� ������ �ֽ��ϴ�.�����ڿ��� ������ �ּ���")
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

	If giftkind_givecnt > 0 Then ''����ǰ ������������
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

IF eCode <> "" THEN	'�̺�Ʈ�� ������ ����ǰ�� ���
	arrsitescope = fnSetCommonCodeArr("eventscope",True) '���� �ڵ尪�� ���� ��Ī ��������
	'�׷츮��Ʈ
	set cEGroup = new ClsEventGroup
	cEGroup.FECode = eCode
	arrGroup = cEGroup.fnGetEventItemGroup
	set cEGroup = nothing
END IF
	blngroup = False
	IF isArray(arrGroup) THEN blngroup = True

	'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
Dim  arrgiftstatus
arrgiftstatus 	= fnSetCommonCodeArr("giftstatus",False)

''��ü����or ���̾ �̺�Ʈ ���� Check -----------------
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

//����ǰ ���� ���
function jsSetGiftKind(gift_code){
	var gift_delivery;
	var sGKN;
	var makerid;

	if (gift_code==""){
		alert("����ǰ�����ڵ尡 �����ϴ�.");
		return;
	}

	if (frmReg.sGN.value==""){
		alert("����ǰ���� ���� �Է��� �ּ���.");
		frmReg.sGN.focus();
		return;
	}
	sGKN=frmReg.sGN.value

	makerid=frmReg.ebrand.value

	if (frmReg.selD.value==""){
		alert("��۹���� ������ �ּ���.");
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

//-- jsPopCal : �޷� �˾� --//
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
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
		eval("document.all."+sName).value = "";
		eval("document.all."+sSpan).style.display = "none";
	}
}

//����ǰ ���
function jsSubmitGift(){
	var frm = document.frmReg;
	if(!frm.sGN.value){
		alert("����ǰ���� �Է��� �ּ���");
		//frm.sGN.focus();
		return;
	}

	if(!frm.sSD.value || !frm.sED.value ){
		alert("�Ⱓ�� �Է����ּ���");
	  //	frm.sSD.focus();
		return;
	}

	if(frm.sSD.value > frm.sED.value){
		alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
		//frm.sED.focus();
		return;
	}

	if(frm.giftscope.value==3){
		if(!frm.ebrand.value){
		alert("�귣����� �������ּ���.���ú귣�忡 ���� ����ǰ�� ���޵˴ϴ�.\n\n�̺�Ʈ ����ǰ�� ��� �̺�Ʈ ����ȭ�鿡�� �귣�� ���� �����մϴ�.");
		return;
		}
	}

	if(frm.giftscope.value==4){
		if(!frm.selG.value){
		alert("�׷��� �������ּ���");
		return;
		}
	}

	if(!frm.sGKN.value){
		alert("����ǰ ���� �Է��� �ּ���");
		return;
	}

	if(!frm.iGK.value){
		alert("����ǰ ������ Ȯ�� ��ư�� ������ Ȯ���� �ּ���");
		return;
	}

	<% if (igScope=1) then %>
	if (frm.chkLimit.checked){
		//alert('��ü ���� ������ ��� ������ üũ�Ͻ� �� �����ϴ�.');
		//return;
	}
	<% end if %>

	if (frm.giftkind_linkGbn.value=="B"){
		if ((frm.giftscope.value!=1)&&(frm.giftscope.value!=9)){
			alert('���� ��ü ���� �̺�Ʈ�� ���� Ÿ�� ����ǰ�� �����մϴ�.');
			return;
		}

		if (frm.selD.value!="C"){
			alert('����ǰ ������ �����ΰ��, ���Ÿ���� ������ �����մϴ�.');
			return;
		}
	}else{
		if (frm.selD.value=="C"){
			alert('����ǰ ������ ������ �ƴѰ��, ���Ÿ���� �������� ���� �Ұ��մϴ�.');
			return;
		}
	}

	var nowDate = "<%=date()%>";

	if(frm.giftstatus.value==7){
		if(frm.sOD.value !=""){
			nowDate = '<%IF dOpenDay <> ""THEN%><%=FormatDate(dOpenDay,"0000-00-00")%><%END IF%>';
		}

		if(frm.sSD.value < nowDate){
			//alert("�������� �����Ϻ���  ������ �ȵ˴ϴ�. �������� �ٽ� �������ּ���");
			//frm.sSD.focus();
			//return;
		}
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

//-- jsChkGiftType : ��籸���� ����ó�� --//
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
	// �׷��ǰ �����ֱ�
  if(iVal ==4){
	document.all.dgiftgroup.style.display = "";
  }else{
	document.all.dgiftgroup.style.display = "none";
  }

  //��÷�� ����϶� �������� ���߱�
   if(iVal ==6){
	document.all.divType1.style.display = "none";
	document.all.divType2.style.display = "none";
  }else{
	document.all.divType1.style.display = "";
	document.all.divType2.style.display = "";
  }
  chkKTdisable();
}

//-- jsImgView : �̹��� Ȯ��ȭ�� ��â���� �����ֱ� --//
function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/lib/showimage.asp?img='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}

function jsChkLimit(){
	<% if (igScope=1) then %>
	//alert('��ü ���� ������ ��� ������ üũ�Ͻ� �� �����ϴ�.');
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

	//���޸� ǥ��
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

// 1+1 ,1:1 üũ
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
		alert("���� ���� ����");
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
		alert("����ǰ�����ڵ尡 �����ϴ�.");
		return;
	}
	var popdisp = window.open('/admin/shopmaster/gift/giftuserdetail.asp?gift_code='+gift_code+'&menupos=<%= menupos %>','giftdetail','width=1280,height=960,scrollbars=yes,resizable=yes');
	popdisp.focus();
}

</script>

<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
<tr>
    <td align="left">
    	* �ֹ���ҽ� ����ǰ �������� ���� ����.
    </td>
    <td align="right"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->

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
	<td bgcolor="<%= adminColor("tabletop") %>" align="center" colspan=2>�� �̺�Ʈ����</td>
</tr>
<%IF eCode <> "" THEN%>
	<tr>
		<td height="30" width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">�̺�Ʈ�ڵ�(�׷�)</td>
		<td bgcolor="#FFFFFF">
			<a href="http://www.10x10.co.kr/event/eventmain.asp?eventid=<%=eCode%>" target="_blank">
			<%=eCode%> <%IF ieGroupCode >0 THEN%>(<%=ieGroupCode%>)<%END IF%></a>
		</td>
	</tr>
	<% if (iopengiftType<>0) then %>
		<tr>
			<td height="30" width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">��ü����Ÿ��</td>
			<td  bgcolor="#FFFFFF" >
				<%= iopengiftName %><%=CHKIIF(iopengiftfrontOpen="Y","&nbsp;&nbsp;(����Ʈ����)","&nbsp;&nbsp;(����Ʈ���� <b>����</b>)")%>
			</td>
		</tr>
	<% end if %>
<%END IF%>
<tr>
	<td height="30" width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center"> ����</td>
	<td bgcolor="#FFFFFF">
		<%IF eCode <> "" THEN%>
			<input type="hidden" name="eventscope" value="<%=iSiteScope%>">
			<input type="hidden" name="selP" value="<%=sPartnerID%>">
			<%=fnGetCommCodeArrDesc(arrsitescope,iSiteScope)%>&nbsp;<%=sPartnerID%>
		<%ELSE%>
			<%sbGetOptCommonCodeArr "eventscope",iSiteScope,False,True, "onChange=javascript:jsSetPartner();"%>
			<span id="spanP" style="display:<%IF iSiteScope<> 3 THEN %>none<%END IF%>;">
			<select name="selP">
				<option value="">--���޸� ��ü--</option>
				<% sbOptPartner sPartnerID%>
			</select>
		<%END IF%>
	</td>
</tr>
<tr>
	<td height="30" width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center"> ����ǰ��</td>
	<td bgcolor="#FFFFFF"><%IF eCode <> "" THEN %><%=sTitle%><input type="hidden" name="sGN" value="<%=sTitle%>"><%ELSE%><input type="text" class="text" name="sGN" size="40" maxlength="64" value="<%=sTitle%>"><%END IF%></td>
</tr>
<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center"> �Ⱓ</td>
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
		������ : <%IF eCode <> "" THEN %><%=dSDay%><input type="hidden" name="sSD" value="<%=dSDay%>"><%ELSE%><input type="text" class="text" name="sSD" size="10"   value="<%=dSDay%>"  onClick="jsPopCal('sSD');"  style="cursor:hand;"><%END IF%>
		<input type="text" name="sSDTime" id="sSDTime" size="10" value="<%= sSDTime %>" class="text" style="display:<%IF iSiteScope<> 3 THEN %>none<%END IF%>;">
		~ ������ : <%IF eCode <> "" THEN %><%=dEDay%><input type="hidden" name="sED" value="<%=dEDay%>"><%ELSE%><input type="text" class="text" name="sED"  size="10"  value="<%=dEDay%>" onClick="jsPopCal('sED');" style="cursor:hand;"><%END IF%>
		<input type="text" name="sEDTime" id="sEDTime" size="10" value="<%= sEDTime %>" class="text" style="display:<%IF iSiteScope<> 3 THEN %>none<%END IF%>;">
	</td>
</tr>
<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">�귣��</td>
	<td bgcolor="#FFFFFF"><% drawSelectBoxDesignerwithName "ebrand", sBrand %></td>
</tr>
<!-- ---------------------------------------------------------------- -->
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center" colspan=2>�� ����ǰ����</td>
</tr>
<tr>
	<td width="100" height="30" width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">����ǰ�����ڵ�</td>
	<td bgcolor="#FFFFFF"><%=gCode%></td>
</tr>
<tr>
	<td width="100" height="30" width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">���</td>
	<td bgcolor="#FFFFFF">
		<select class="select" name="selD" onChange="dpCpnSpan(this)">
			<option value="N" <%IF sgDelivery = "N" THEN%>selected<%END IF%>>�ٹ����ٹ��</option>
			<option value="Y" <%IF sgDelivery = "Y" THEN%>selected<%END IF%>>��ü���</option>
			<% if (igScope=1)or(igScope=9) then %>
				<option value="C" <%IF sgDelivery = "C" THEN%>selected<%END IF%>>����</option>
			<% end if %>
		</select>
		<span id="icpnSpan" name="icpnSpan" style="display=<%= chkIIF(sgDelivery="C","block","none") %>">
			������ȣ : <input type="text" class="text_ro" READOnly name="bcouponidx" value="<%= BCouponIdx %>" size="9" maxlegth="9"> <!-- in Gift_kind -->
		</span>
	</td>
</tr>
<tr>
	<td width="100" height="30" bgcolor="<%= adminColor("tabletop") %>" align="center">����ǰ����</td>
	<td bgcolor="#FFFFFF">
		<input type="hidden" name="orgiGK" value="<%=igkCode%>">
		<input type="hidden" name="iGK" value="<%=igkCode%>">
		<input type="text" class="text" name="sGKN" size="40" maxlength="60" value ="<%=igkName%>" onkeyup="document.frmReg.iGK.value='';"> <input type="button" class="button" value="Ȯ��" onClick="jsSetGiftKind('<%= gCode %>');">

		<% if (igScope=1)or(igScope=9) then %>
		<input type="button" class="button" value="����" onClick="jsGiftKindManage();">
		<% end if %>

		<div id="spanImg">
		<%IF sgkImg <> "" THEN%><a href="javascript:jsImgView('<%=sgkImg%>')"><img src="<%=sgkImg%>" border="0"></a><%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td width="100" height="30" bgcolor="<%= adminColor("tabletop") %>" align="center">����ǰ</td>
	<td bgcolor="#FFFFFF">
		<%sbGetOptGiftCodeValue "giftscope",igScope,blngroup,"onchange='jsChkgiftgroup(this.value);'",eCode%>
		<div id="dgiftgroup" style="display:<%IF NOT (blngroup and igScope = "4") THEN%>none<%END IF%>;">
		<%IF isArray(arrGroup) THEN%>
			�׷켱��: 
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
<% '<!--�������� �̺�Ʈ��÷���� ��� �������� �����--> %>
<tr id="divType1" style="display:<%IF igScope=6 THEN%>none<%END IF%>;">
	<td width="100" height="30" bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">��������</td>
	<td bgcolor="#FFFFFF">
		<% if (igScope=9) then %>
			<select name="gifttype" onchange='jsChkGiftType(this.value);'>
				<option value="2" selected>����(��)</option>
			<select>
		<% else %>
			<%sbGetOptCommonCodeArr "gifttype", igType, False,True,"onchange='jsChkGiftType(this.value);'"%>
		<% end if %>
	</td>
</tr>
<% '<!--�������� �̺�Ʈ��÷���� ��� �������� �����--> %>
<tr id="divType2" style="display:<%IF igScope=6 THEN%>none<%END IF%>;">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">��������</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="sGR1" size="10" style="text-align:right;<%IF igType= "1" THEN%>background-color:#E6E6E6;" readonly<%ELSE%>"<%END IF%> value="<%=igR1%>"> �̻� ~ <input type="text" class="text" name="sGR2" size="10" style="text-align:right;<%IF igType= "1" THEN%>background-color:#E6E6E6;" readonly<%ELSE%>"<%END IF%> value="<%=igR2%>"> �̸�
	</td>
</tr>
<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">����ǰ����</td>
	<td bgcolor="#FFFFFF" >
		<input type="text" class="text" name="iGKC" size="4" maxlength="10" value="<%=igkCnt%>" style="text-align:right;"> ����
		<% if (igScope=9) then %>
			<span id="spanKT" style="display:;">
				<label title="������ǰ����" ><input type="checkbox" name="tmpchkKT2" id="tmpchkKT2" disabled onclick="jsCheckKT('tmpchkKT2',this.cheked);" value="2" <%IF igkType = 2 THEN%>checked<%END IF%>>1+1(���ϻ�ǰ) </label>
				<label title="�ٸ���ǰ����" ><input type="checkbox" name="tmpchkKT3" id="tmpchkKT3" onclick="jsCheckKT('tmpchkKT3',this.cheked);" value="3" <%IF igkType = 3 THEN%>checked<%END IF%>>1:1(�ٸ���ǰ) </label>
			</span>
		<% else %>
			<span id="spanKT" style="display:<%IF igType = 2 THEN%>none<%END IF%>;">
				<label title="���ϻ�ǰ����" ><input type="checkbox" name="tmpchkKT2" id="tmpchkKT2" onclick="jsCheckKT('tmpchkKT2');" <%IF igScope<>5 Then%>disabled<%End IF%> value="2" <%IF igkType = 2 THEN%>checked<%END IF%>>1+1(���ϻ�ǰ) </label>
				<label title="�ٸ���ǰ����" ><input type="checkbox" name="tmpchkKT3" id="tmpchkKT3" onclick="jsCheckKT('tmpchkKT3');" value="3" <%IF igkType = 3 THEN%>checked<%END IF%>>1:1(�ٸ���ǰ) </label>
			</span>
		<% end if %>
	</td>
</tr>
<tr>
	<td height="30" bgcolor="<%= adminColor("tabletop") %>"  align="center">����ǰ��������</td>
	<td bgcolor="#FFFFFF">
		<input type="checkbox" name="chkLimit" onClick="jsChkLimit();" <%IF igkLimit <> "" THEN%>checked<%END IF%>>����
		<input type="text" class="text" name="iL" size="5" value="<%=igkLimit%>" style="text-align:right" <%IF igkLimit ="" THEN%>style="background-color:#E6E6E6;" readonly<%END IF%>>
		<strong>-<input type="text" class="text" size="5" name="givecnt" onclick="nowcnt();" value="<%=giftkind_givecnt %>" style="text-align:right" <%IF igkLimit ="" THEN%>style="background-color:#E6E6E6;" readonly<%END IF%>>=<% if igkLimit<>"" and giftkind_givecnt<>""  then: Response.Write igkLimit-giftkind_givecnt: Else Response.Write "0": End If %></strong>
			(�������� ���� ��쿡�� �Է�)
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
					<td>�հ�</td>
					<td><%=totcnt%></td>
				</tr>
				<tr align="center">
					<td colspan="2" onclick="document.getElementById('aaaa').style.display = 'none';">[�ݱ�]</td>
				</tr>
			</table>
		</div>
		<% End If %>
		<br><br><input type="button" value="������ǰ����Ʈ" onclick="popgiftdetail('<%= gCode %>');" class="button" >
	</td>
</tr>
<tr>
	<td height="30" bgcolor="<%= adminColor("tabletop") %>"  align="center">����</td>
	<td bgcolor="#FFFFFF">
		<%IF eCode <> "" THEN%>
			<input type="hidden" name="giftstatus" value="<%=igStatus%>"><%=replace(fnGetCommCodeArrDesc(arrgiftstatus,igStatus),"���¿���","����")%>
		<%ELSE%>
			<%sbGetOptStatusCodeValue "giftstatus", igStatus, False,""%>
		<%END IF%>
		<input type="hidden" name="sOD" value="<%=dOpenDay%>">
		<input type="hidden" name="sCD" value="<%=dCloseDay%>">
		<%IF dOpenDay <> "" THEN%><span style="padding-left:10px;">����ó����: <%=dOpenDay%></span><%END IF%>
		<%IF dCloseDay <> "" THEN%><br><span style="padding-left:42px;">����ó����: <%=dCloseDay%></span><%END IF%>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>"  align="center">�������</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="sGU" value="Y" <%IF igUsing = "Y" THEN%>checked<%END IF%>>��� <input type="radio" name="sGU" value="N" <%IF igUsing = "N" THEN%>checked<%END IF%>>������
	</td>
</tr>
<!-- ---------------------------------------------------------------- -->
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center" colspan=2>�� ����ǰǥ������(����Ʈ)</td>
</tr>
<!--<tr>-->
	<!--<td height="30" width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">�������<br>(������)- OLD</td>-->
	<!--<td bgcolor="#FFFFFF"><%'=db2html(sOldName)%></td>-->
<!--</tr>-->
<tr>
	<td height="30" width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">�������<br>(������)- New</td>
	<td bgcolor="#FFFFFF">
		<% =fnComGetEventConditionStr(igkType, igScope,igType,igR1, igR2,igkName,igkCnt, igkCnt,0,0,sBrand)%>
	</td>
</tr>
<% if eCode<>"" then %>
	<tr>
		<td height="30" width="100" bgcolor="#FFFFFF" align="left" colspan=2><B>����ǰ �ؽ�Ʈ �ڽ� ����</B></td>
	</tr>
	<tr>
		<td width="100"  bgcolor="<%= adminColor("tabletop") %>">��뿩��<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<select name="gift_isusing" id="gift_isusing1" onchange="TnGiftUsingNum(this.value);">
				<option value="1"<% If GiftIsusing=1 Then %> selected<% End If %>>1�� ���</option>
				<option value="2"<% If GiftIsusing=2 Then %> selected<% End If %>>2�� ���</option>
				<option value="3"<% If GiftIsusing=3 Then %> selected<% End If %>>3�� ���</option>
				<option value="0"<% If GiftIsusing=0 Then %> selected<% End If %>>��� ����</option>
			</select>
			<input type="checkbox" name="gift_infotext" value="Y"<% If GiftInfoText="Y" Then %> checked<% End If %>>�������� �ȳ�����
		</td>
	</tr>
	<tr style="display:" id="gifttxt1">
		<td width="100"  bgcolor="<%= adminColor("tabletop") %>">����ǰ1 ����</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="gift_text1" id="gift_text1_1" value="<%=GiftText1%>" size="100" maxlength="64">
		</td>
	</tr>
	<tr style="display:" id="giftimg1">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">����ǰ1 �̹���</td>
		<td bgcolor="#FFFFFF">
			<input type="button" name="btnicon" value="�̹��� ���" onClick="jsSetImg('<%=eFolder%>','<%=GiftImage1%>','gift_img1','spangift_img1')" class="button">
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
		<td width="100"  bgcolor="<%= adminColor("tabletop") %>">����ǰ2 ����</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="gift_text2" id="gift_text2_1" value="<%=GiftText2%>" size="100" maxlength="64">
		</td>
	</tr>
	<tr style="display:<% If GiftIsusing<2 Then %>none<% End If %>" id="giftimg2">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">����ǰ2 �̹���</td>
		<td bgcolor="#FFFFFF">
			<input type="button" name="btnicon" value="�̹��� ���" onClick="jsSetImg('<%=eFolder%>','<%=GiftImage2%>','gift_img2','spangift_img2')" class="button">
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
		<td width="100"  bgcolor="<%= adminColor("tabletop") %>">����ǰ3 ����</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="gift_text3" id="gift_text3_1" value="<%=GiftText3%>" size="100" maxlength="64">
		</td>
	</tr>
	<tr style="display:<% If GiftIsusing<3 Then %>none<% End If %>" id="giftimg3">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">����ǰ3 �̹���</td>
		<td bgcolor="#FFFFFF">
			<input type="button" name="btnicon" value="�̹��� ���" onClick="jsSetImg('<%=eFolder%>','<%=GiftImage3%>','gift_img3','spangift_img3')" class="button">
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
		<input type="button" class="button" value="�����ϱ�" onClick="jsSubmitGift();">
		&nbsp;
		<input type="button" class="button" value="���" onClick="history.back();">
	</td>
</tr>
</table>
</form>

<script type='text/javascript'>

function getOnLoad(){
    alert("���̾ �ӽ� ����ǰ ���� \n\n�� �����ݾ� 30,000�� �̻� ���Ž� �߰���\n\n����� ���� ���� ���");
}
<% if gCode="5345" then %>
	window.onload=getOnLoad;
<% end if %>

</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
