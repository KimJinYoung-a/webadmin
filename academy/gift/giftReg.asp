<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����ǰ ���� 
' History : 2010.09.27 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/gift/giftcls.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->
<%
Dim eCode, cEventsimple,eState, sStateDesc ,clsGift ,iSiteScope,sPartnerID,arrsitescope
Dim sTitle, dSDay, dEDay, dOpenDay, dCloseDay, sBrand, blnGroup,igType
	eCode     = requestCheckVar(Request("eC"),10)
	igType = 2

IF eCode <> "" THEN		'�̺�Ʈ ���� �ϰ��
	
set cEventsimple = new ClsEventSummary
	cEventsimple.FECode = eCode
	cEventsimple.fnGetEventConts
	sTitle 	= cEventsimple.FEName
	dSDay	= cEventsimple.FESDay
	dEDay	= cEventsimple.FEEDay
	sBrand	= cEventsimple.FBrand
	eState  = cEventsimple.FEState
	dOpenDay= cEventsimple.FEOpenDate
	dCloseDay=cEventsimple.FECloseDate
	sStateDesc =cEventsimple.FEStateDesc
	iSiteScope =cEventsimple.FEScope
	sPartnerID =cEventsimple.FPartnerID
set cEventsimple = nothing

blngroup = False
arrsitescope = fnSetCommonCodeArr("eventscope",True)

END IF

if eState < 6 then eState = 0	'�̺�Ʈ ���¿� ����ǰ ���� ��Īó��(�������� ���´� ��� ������)
%>

<script language="javascript">

	//����ǰ ���� ���
	function jsSetGiftKind(){
		var winkind;
		winkind = window.open('popgiftKindReg.asp?sGKN='+document.frmReg.sGKN.value,'popkind','width=450px, height=300px;');
		winkind.focus();
	}


	//-- jsPopCal : �޷� �˾� --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//����ǰ ���
	function jsSubmitGift(){
		var frm = document.frmReg;
		if(!frm.sGN.value){
			alert("������ �Է��� �ּ���");
		//	frm.sGN.focus();
			return false;
		}

		if(!frm.sSD.value ){
		  	alert("�������� �Է����ּ���");
		//  	frm.sSD.focus();
		  	return false;
	  	}

	  	if(frm.sED.value){
		  	if(frm.sSD.value > frm.sED.value){
			  	alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
		//	  	frm.sED.focus();
			  	return false;
		  	}
		}
		if(frm.giftscope.value==3){
			if(!frm.ebrand.value){
			alert("�귣����� �������ּ���.���ú귣�忡 ���� ����ǰ�� ���޵˴ϴ�.\n\n�̺�Ʈ ����ǰ�� ��� �̺�Ʈ ����ȭ�鿡�� �귣�� ���� �����մϴ�.");
			return false;
			}
		}

		if(frm.giftscope.value==4){
			if(!frm.selG.value){
			alert("�׷��� �������ּ���");
			return false;
			}
		}
		var nowDate = "<%=date()%>";

	 if(frm.giftstatus.value==7){
	 	if(frm.sOD.value !=""){
	 		nowDate = '<%IF dOpenDay <> ""THEN%><%=FormatDate(dOpenDay,"0000-00-00")%><%END IF%>';
		}

		if(frm.sSD.value < nowDate){
			alert("�������� �����Ϻ���  ������ �ȵ˴ϴ�. �������� �ٽ� �������ּ���");
		  	//frm.sSD.focus();
		  	return false;
		 }
	  }


		if(!frm.sGKN.value){
			alert("����ǰ ���� �Է��� �ּ���");
			return false;
		}

		if(!frm.iGK.value){
			alert("����ǰ ������ Ȯ�� ��ư�� ������ Ȯ���� �ּ���");
			return false;
		}

		return true;
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
			}else{
				document.all.spanKT.style.display = "";
			}
			chkKTdisable();
	}

	function jsChkgiftgroup(iVal){

	   if(iVal ==6){
		document.all.divType.style.display = "none";
	  }else{
	 	document.all.divType.style.display = "";
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
		if(document.frmReg.chkLimit.checked){
			document.all.iL.readOnly=false;
			document.all.iL.style.backgroundColor='';
		}else{
			document.all.iL.readOnly=true;
			document.all.iL.style.backgroundColor='#E6E6E6';
			document.frmReg.iL.value = "";
		}
	}

	//���޸� ǥ��
	function jsSetPartner(){
		if(document.frmReg.eventscope.options[document.frmReg.eventscope.selectedIndex].value == 3){
			document.all.spanP.style.display ="";
		}else{
			document.all.spanP.style.display ="none";
		}
	}

	// ����ǰ��ϳ��� ��������
	function jsImport(ec){
		var pp = window.open('/academy/gift/popGiftList.asp?eC='+ec,'popim','scrollbars=yes,resizable=yes,width=900,height=600');

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

</script>

<table width="100%" border="0" align="left" class="a" cellpadding="0" cellspacing="1"  >
<form name="frmReg" method="post" action="giftProc.asp" onSubmit="return jsSubmitGift();">
<input type="hidden" name="sM" value="I">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="chkKT" value="0">
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<%IF eCode <> "" THEN%>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">�̺�Ʈ�ڵ�</td>
			<td bgcolor="#FFFFFF" colspan="3"><%=eCode%></td>
		</tr>
		<%END IF%>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center"></td>
			<td bgcolor="#FFFFFF" colspan="3"><input type="button" class="button" value="��������" onClick="jsImport('<%= eCode %>');"></td>
		</tr>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center"> ����</td>
			<td bgcolor="#FFFFFF" colspan="3">
				<%IF eCode <> "" THEN%>
				<input type="hidden" name="eventscope" value="<%=iSiteScope%>">
				<input type="hidden" name="selP" value="<%=sPartnerID%>">
				<%=fnGetCommCodeArrDesc(arrsitescope,iSiteScope)%>&nbsp;<%=sPartnerID%>
				<%ELSE%>
				<%sbGetOptCommonCodeArr "eventscope","",False,True, "onChange=javascript:jsSetPartner();"%>
		   		<span id="spanP" style="display:none;">
		   		<select name="selP">
		   			<option value="">--���޸� ��ü--</option>
		   			<% sbOptPartner ""%>
		   		</select>
		   		<%END IF%>
		   	</td>
		</tr>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center"> ����</td>
			<td bgcolor="#FFFFFF" width="400"><%IF eCode <> "" THEN %><%=sTitle%><input type="hidden" name="sGN" value="<%=sTitle%>"><%ELSE%><input type="text" name="sGN" size="30" maxlength="64"><%END IF%></td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center"> �Ⱓ</td>
			<td bgcolor="#FFFFFF">
				������ : <%IF eCode <> "" THEN %><%=dSDay%><input type="hidden" name="sSD" value="<%=dSDay%>"><%ELSE%><input type="text" name="sSD" size="10"   onClick="jsPopCal('sSD');"  style="cursor:hand;"><%END IF%>
				~ ������ : <%IF eCode <> "" THEN %><%=dEDay%><input type="hidden" name="sED" value="<%=dEDay%>"><%ELSE%><input type="text" name="sED"  size="10" onClick="jsPopCal('sED');" style="cursor:hand;"><%END IF%>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">�������</td>
			<td bgcolor="#FFFFFF"><%sbGetOptGiftCodeValue "giftscope","",blngroup,"onchange='jsChkgiftgroup(this.value);'",eCode%>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">�귣��</td>
			<td bgcolor="#FFFFFF"><%IF sBrand <> "" THEN %><%=sBrand%><input type="hidden" name="ebrand" value="<%=sBrand%>"><%ELSE%><% drawSelectBoxLecturer "ebrand", "" %><%END IF%></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<div id="divType" style="display:;">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">��������</td>
			<td width="400" bgcolor="#FFFFFF">
				<%sbGetOptCommonCodeArr "gifttype", igType, False,True,"onchange='jsChkGiftType(this.value);'"%>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">��������</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="sGR1" size="10" style="text-align:right" value="0"> �̻� ~ <input type="text" name="sGR2" size="10" style="text-align:right" value="0"> �̸�
				(ex. 20�� �̻�: 20~0)
			</td>
		</tr>
		</table>
		</div>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">����ǰ����</td>
			<td  width="400" bgcolor="#FFFFFF">
				<input type="hidden" name="iGK" >
				<input type="text" name="sGKN" size="40" maxlength="60" onkeyup="document.frmReg.iGK.value='';"> <input type="button" class="button" value="Ȯ��" onClick="jsSetGiftKind();">
				<div id="spanImg"></div>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">����ǰ����</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="iGKC" size="4" maxlength="10" value="1" style="text-align:right;"> ����
				<span id="spanKT" style="display:none;">
					<label title="������ǰ����" ><input type="checkbox" name="tmpchkKT2" onclick="jsCheckKT('tmpchkKT2',this.cheked);" value="2">1+1(���ϻ�ǰ) </label>
					<label title="�ٸ���ǰ����" ><input type="checkbox" name="tmpchkKT3" onclick="jsCheckKT('tmpchkKT3',this.cheked);" value="3">1:1(�ٸ���ǰ) </label>
				</span>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">����ǰ��������</td>
			<td bgcolor="#FFFFFF">
				<input type="checkbox" name="chkLimit" onClick="jsChkLimit();">����
				<input type="text" name="iL" size="4"  style="text-align:right;background-color:#E6E6E6;" readonly> ��(�������� ���� ��쿡�� �Է�)
			</td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">��۹��</td>
			<td bgcolor="#FFFFFF">
				<select name="selD">
				<!--<option value="N" >�ٹ����ٹ��</option>-->
				<option value="Y" >��ü���</option>
				</select>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">����</td>
			<td bgcolor="#FFFFFF" colspan="3">
				<%IF eCode <> "" THEN%>
					<input type="hidden" name="giftstatus" value="<%=eState%>"><%=replace(sStateDesc,"���¿���","����")%>
				<%ELSE%>
					<%sbGetOptCommonCodeArr "giftstatus", "", False,True,""%>
				<%END IF%>
				<input type="hidden" name="sOD" value="">
				<input type="hidden" name="sCD" value="">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<input type="image" src="/images/icon_save.gif">
		<a href="javascript:history.back();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->