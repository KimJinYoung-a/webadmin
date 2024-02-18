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

<%
Dim clsGift ,strParm ,sOldName ,iSiteScope,sPartnerID,arrsitescope
Dim eCode, cEvent,intgroup ,iSerachType,sSearchTxt, sDate,sSdate,sEdate,iCurrpage,sgDelivery
Dim sTitle, dSDay, dEDay, sBrand, blnGroup, dOpenDay, dCloseDay
Dim gCode,igScope,ieGroupCode, igType, igR1,igR2, igStatus, dRegdate, sAdminid, igUsing
Dim igkCode, igkType, igkCnt,igkLimit, igkName,sgkImg
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

set clsGift = new CGift
	clsGift.FGCode = gCode
	clsGift.fnGetGiftConts
	
	sTitle		= clsGift.FGName
	igScope 	= clsGift.FGScope
	eCode		= clsGift.FECode	
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
set clsGift = nothing

IF eCode = 0 THEN eCode = ""
IF igkLimit = 0 THEN igkLimit = ""
IF isNull(igkLimit) THEN igkLimit = ""

IF eCode <> "" THEN	'�̺�Ʈ�� ������ ����ǰ�� ���
	arrsitescope = fnSetCommonCodeArr("eventscope",True) '���� �ڵ尪�� ���� ��Ī ��������
END IF

 blngroup = False

  '�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
Dim  arrgiftstatus
arrgiftstatus 	= fnSetCommonCodeArr("giftstatus",False)
%>

<script language="javascript">

	//����ǰ ���� ��� 
	function jsSetGiftKind(){
		var winkind;
		winkind = window.open('popgiftKindReg.asp?sGKN='+document.frmReg.sGKN.value,'popkind','width=470px, height=300px, scrollbars=yes,resizable=yes');
		winkind.focus();
	}

    function jsGiftKindManage(){
		var winkind;
		winkind = window.open('popgiftKindManage.asp?iGK='+document.frmReg.iGK.value,'popkindMan','width=470px, height=300px, scrollbars=yes,resizable=yes');
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
			//frm.sGN.focus();
			return false;
		}

		if(!frm.sSD.value || !frm.sED.value ){
		  	alert("�Ⱓ�� �Է����ּ���");
		  //	frm.sSD.focus();
		  	return false;
	  	}

	  	if(frm.sSD.value > frm.sED.value){
		  	alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
		  	//frm.sED.focus();
		  	return false;
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

		if(!frm.sGKN.value){
			alert("����ǰ ���� �Է��� �ּ���");
			return false;
		}

		if(!frm.iGK.value){
			alert("����ǰ ������ Ȯ�� ��ư�� ������ Ȯ���� �ּ���");
			return false;
		}
        
        <% if (igScope=1) then %>
        if (frm.chkLimit.checked){
	        alert('��ü ���� ������ ��� ������ üũ�Ͻ� �� �����ϴ�.');
	        return false;
	    }
	    <% end if %>
        
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
		// �׷��ǰ �����ֱ�
	  if(iVal ==4){
		document.all.dgiftgroup.style.display = "";
	  }else{
	 	document.all.dgiftgroup.style.display = "none";
	  }

	  //��÷�� ����϶� �������� ���߱�
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
	    <% if (igScope=1) then %>
	    alert('��ü ���� ������ ��� ������ üũ�Ͻ� �� �����ϴ�.');
	    document.frmReg.chkLimit.checked = false;
	    <% end if %>
	    
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
<form name="frmReg" method="post" action="giftProc.asp?<%=strParm%>" onSubmit="return jsSubmitGift();">
<input type="hidden" name="sM" value="U">
<input type="hidden" name="gC" value="<%=gCode%>">
<input type="hidden" name="sGD" value="<%=sgDelivery%>">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="chkKT" value="<%=igkType%>">
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">����ǰ�ڵ�</td>
			<td bgcolor="#FFFFFF" colspan="3"><%=gCode%></td>
		</tr>
		<%IF eCode <> "" THEN%>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">�̺�Ʈ�ڵ�</td>
			<td bgcolor="#FFFFFF" colspan="3"><a href="http://www.10x10.co.kr/event/eventmain.asp?eventid=<%=eCode%>" target="_blank"><%=eCode%></a></td>
		</tr>
		<%END IF%>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"> ����</td>
			<td bgcolor="#FFFFFF"  colspan="3">
			<%IF eCode <> "" THEN%>
				<input type="hidden" name="eventscope" value="<%=iSiteScope%>">
				<input type="hidden" name="selP" value="<%=sPartnerID%>">
				<%=fnGetCommCodeArrDesc(arrsitescope,iSiteScope)%>&nbsp;<%=sPartnerID%>
			<%ELSE%>
				<%sbGetOptCommonCodeArr "eventscope",iSiteScope,False,True, "onChange=javascript:jsSetPartner();"%>
		   		<span id="spanP" style="display:<%IF iSiteScope<> 3 THEN %>none<%END IF%>;">
		  	<%END IF%>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"> ����</td>
			<td bgcolor="#FFFFFF"  width="400"><%IF eCode <> "" THEN %><%=sTitle%><input type="hidden" name="sGN" value="<%=sTitle%>"><%ELSE%><input type="text" name="sGN" size="30" maxlength="64" value="<%=sTitle%>"><%END IF%></td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center"> �Ⱓ</td>
			<td bgcolor="#FFFFFF">
				������ : <%IF eCode <> "" THEN %><%=dSDay%><input type="hidden" name="sSD" value="<%=dSDay%>"><%ELSE%><input type="text" name="sSD" size="10"   value="<%=dSDay%>"  onClick="jsPopCal('sSD');"  style="cursor:hand;"><%END IF%>
				~ ������ : <%IF eCode <> "" THEN %><%=dEDay%><input type="hidden" name="sED" value="<%=dEDay%>"><%ELSE%><input type="text" name="sED"  size="10"  value="<%=dEDay%>" onClick="jsPopCal('sED');" style="cursor:hand;"><%END IF%>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">����ǰ</td>
			<td bgcolor="#FFFFFF"><%sbGetOptGiftCodeValue "giftscope",igScope,blngroup,"onchange='jsChkgiftgroup(this.value);'",eCode%>
			<div id="dgiftgroup" style="display:<%IF NOT (blngroup and igScope = "4") THEN%>none<%END IF%>;">
			</div>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">�귣��</td>
			<td bgcolor="#FFFFFF">
				<% drawSelectBoxLecturer "ebrand", sBrand %>				
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<div id="divType" style="display:<%IF igScope=6 THEN%>none<%END IF%>;"><!--�������� �̺�Ʈ��÷���� ��� �������� �����-->
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">��������</td>
			<td bgcolor="#FFFFFF" width="400">
				<%sbGetOptCommonCodeArr "gifttype", igType, False,True,"onchange='jsChkGiftType(this.value);'"%>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">��������</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="sGR1" size="10" style="text-align:right;<%IF igType= "1" THEN%>background-color:#E6E6E6;" readonly<%ELSE%>"<%END IF%> value="<%=igR1%>"> �̻� ~ <input type="text" name="sGR2" size="10" style="text-align:right;<%IF igType= "1" THEN%>background-color:#E6E6E6;" readonly<%ELSE%>"<%END IF%> value="<%=igR2%>"> �̸�
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
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">����ǰ����</td>
			<td bgcolor="#FFFFFF"  width="400">
			    <input type="hidden" name="orgiGK" value="<%=igkCode%>">
				<input type="hidden" name="iGK" value="<%=igkCode%>">
				<input type="text" name="sGKN" size="40" maxlength="60" value ="<%=igkName%>" onkeyup="document.frmReg.iGK.value='';"> <input type="button" class="button" value="Ȯ��" onClick="jsSetGiftKind();">
				
				<% if (igScope=1) then %>
				<input type="button" class="button" value="����" onClick="jsGiftKindManage();">
				<% end if %>
				
				<div id="spanImg">
				<%IF sgkImg <> "" THEN%><a href="javascript:jsImgView('<%=sgkImg%>')"><img src="<%=sgkImg%>" border="0"></a><%END IF%>
				</div>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">����ǰ����</td>
			<td bgcolor="#FFFFFF" >
				<input type="text" name="iGKC" size="4" maxlength="10" value="<%=igkCnt%>" style="text-align:right;"> ����
				<span id="spanKT" style="display:<%IF igType = 2 THEN%>none<%END IF%>;">
					<label title="���ϻ�ǰ����" ><input type="checkbox" name="tmpchkKT2" onclick="jsCheckKT('tmpchkKT2');" <%IF igScope<>5 Then%>disabled<%End IF%> value="2" <%IF igkType = 2 THEN%>checked<%END IF%>>1+1(���ϻ�ǰ) </label>
					<label title="�ٸ���ǰ����" ><input type="checkbox" name="tmpchkKT3" onclick="jsCheckKT('tmpchkKT3');" value="3" <%IF igkType = 3 THEN%>checked<%END IF%>>1:1(�ٸ���ǰ) </label>
				</span>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">����ǰ��������</td>
			<td bgcolor="#FFFFFF">
				<input type="checkbox" name="chkLimit" onClick="jsChkLimit();" <%IF igkLimit <> "" THEN%>checked<%END IF%>>����
				<input type="text" name="iL" size="4" value="<%=igkLimit%>" style="text-align:right" <%IF igkLimit ="" THEN%>style="background-color:#E6E6E6;" readonly<%END IF%>> (�������� ���� ��쿡�� �Է�)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">���</td>
			<td bgcolor="#FFFFFF">
			<!--'%=fnSetDelivery(sgDelivery)%-->
				<select name="selD">
				<!--<option value="N" <%IF sgDelivery = "N" THEN%>selected<%END IF%>>�ٹ����ٹ��</option>-->
				<option value="Y" <%IF sgDelivery = "Y" THEN%>selected<%END IF%>>��ü���</option>
				</select>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">����</td>
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
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">�������</td>
			<td bgcolor="#FFFFFF">
				<input type="radio" name="sGU" value="Y" <%IF igUsing = "Y" THEN%>checked<%END IF%>>��� <input type="radio" name="sGU" value="N" <%IF igUsing = "N" THEN%>checked<%END IF%>>������
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
</table>
</form>

<script language='javascript'>

	function getOnLoad(){
	    alert("���̾ �ӽ� ����ǰ ���� \n\n�� �����ݾ� 30,000�� �̻� ���Ž� �߰���\n\n����� ���� ���� ���");
	}
	
	<% if gCode="5345" then %>
	window.onload=getOnLoad;
	<% end if %>

</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->