<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  사은품 관리 
' History : 2010.09.27 한용민 생성
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

IF eCode <> "" THEN		'이벤트 연관 일경우
	
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

if eState < 6 then eState = 0	'이벤트 상태와 사은품 상태 매칭처리(오픈이전 상태는 모두 대기상태)
%>

<script language="javascript">

	//사은품 종류 등록
	function jsSetGiftKind(){
		var winkind;
		winkind = window.open('popgiftKindReg.asp?sGKN='+document.frmReg.sGKN.value,'popkind','width=450px, height=300px;');
		winkind.focus();
	}


	//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//사은품 등록
	function jsSubmitGift(){
		var frm = document.frmReg;
		if(!frm.sGN.value){
			alert("제목을 입력해 주세요");
		//	frm.sGN.focus();
			return false;
		}

		if(!frm.sSD.value ){
		  	alert("시작일을 입력해주세요");
		//  	frm.sSD.focus();
		  	return false;
	  	}

	  	if(frm.sED.value){
		  	if(frm.sSD.value > frm.sED.value){
			  	alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
		//	  	frm.sED.focus();
			  	return false;
		  	}
		}
		if(frm.giftscope.value==3){
			if(!frm.ebrand.value){
			alert("브랜드명을 선택해주세요.선택브랜드에 대해 사은품이 지급됩니다.\n\n이벤트 사은품일 경우 이벤트 수정화면에서 브랜드 수정 가능합니다.");
			return false;
			}
		}

		if(frm.giftscope.value==4){
			if(!frm.selG.value){
			alert("그룹을 선택해주세요");
			return false;
			}
		}
		var nowDate = "<%=date()%>";

	 if(frm.giftstatus.value==7){
	 	if(frm.sOD.value !=""){
	 		nowDate = '<%IF dOpenDay <> ""THEN%><%=FormatDate(dOpenDay,"0000-00-00")%><%END IF%>';
		}

		if(frm.sSD.value < nowDate){
			alert("시작일이 오픈일보다  빠르면 안됩니다. 시작일을 다시 선택해주세요");
		  	//frm.sSD.focus();
		  	return false;
		 }
	  }


		if(!frm.sGKN.value){
			alert("사은품 종류 입력해 주세요");
			return false;
		}

		if(!frm.iGK.value){
			alert("사은품 종류를 확인 버튼을 눌러서 확인해 주세요");
			return false;
		}

		return true;
	}

	//-- jsChkGiftType : 모든구매자 조건처리 --//
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

		//-- jsImgView : 이미지 확대화면 새창으로 보여주기 --//
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

	//제휴몰 표기
	function jsSetPartner(){
		if(document.frmReg.eventscope.options[document.frmReg.eventscope.selectedIndex].value == 3){
			document.all.spanP.style.display ="";
		}else{
			document.all.spanP.style.display ="none";
		}
	}

	// 사은품등록내역 가져오기
	function jsImport(ec){
		var pp = window.open('/academy/gift/popGiftList.asp?eC='+ec,'popim','scrollbars=yes,resizable=yes,width=900,height=600');

	}
	
	// 1+1 ,1:1 체크
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
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">이벤트코드</td>
			<td bgcolor="#FFFFFF" colspan="3"><%=eCode%></td>
		</tr>
		<%END IF%>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center"></td>
			<td bgcolor="#FFFFFF" colspan="3"><input type="button" class="button" value="가져오기" onClick="jsImport('<%= eCode %>');"></td>
		</tr>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center"> 범위</td>
			<td bgcolor="#FFFFFF" colspan="3">
				<%IF eCode <> "" THEN%>
				<input type="hidden" name="eventscope" value="<%=iSiteScope%>">
				<input type="hidden" name="selP" value="<%=sPartnerID%>">
				<%=fnGetCommCodeArrDesc(arrsitescope,iSiteScope)%>&nbsp;<%=sPartnerID%>
				<%ELSE%>
				<%sbGetOptCommonCodeArr "eventscope","",False,True, "onChange=javascript:jsSetPartner();"%>
		   		<span id="spanP" style="display:none;">
		   		<select name="selP">
		   			<option value="">--제휴몰 전체--</option>
		   			<% sbOptPartner ""%>
		   		</select>
		   		<%END IF%>
		   	</td>
		</tr>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center"> 제목</td>
			<td bgcolor="#FFFFFF" width="400"><%IF eCode <> "" THEN %><%=sTitle%><input type="hidden" name="sGN" value="<%=sTitle%>"><%ELSE%><input type="text" name="sGN" size="30" maxlength="64"><%END IF%></td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center"> 기간</td>
			<td bgcolor="#FFFFFF">
				시작일 : <%IF eCode <> "" THEN %><%=dSDay%><input type="hidden" name="sSD" value="<%=dSDay%>"><%ELSE%><input type="text" name="sSD" size="10"   onClick="jsPopCal('sSD');"  style="cursor:hand;"><%END IF%>
				~ 종료일 : <%IF eCode <> "" THEN %><%=dEDay%><input type="hidden" name="sED" value="<%=dEDay%>"><%ELSE%><input type="text" name="sED"  size="10" onClick="jsPopCal('sED');" style="cursor:hand;"><%END IF%>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">증정대상</td>
			<td bgcolor="#FFFFFF"><%sbGetOptGiftCodeValue "giftscope","",blngroup,"onchange='jsChkgiftgroup(this.value);'",eCode%>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">브랜드</td>
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
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">증정조건</td>
			<td width="400" bgcolor="#FFFFFF">
				<%sbGetOptCommonCodeArr "gifttype", igType, False,True,"onchange='jsChkGiftType(this.value);'"%>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">증정범위</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="sGR1" size="10" style="text-align:right" value="0"> 이상 ~ <input type="text" name="sGR2" size="10" style="text-align:right" value="0"> 미만
				(ex. 20개 이상: 20~0)
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
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">사은품종류</td>
			<td  width="400" bgcolor="#FFFFFF">
				<input type="hidden" name="iGK" >
				<input type="text" name="sGKN" size="40" maxlength="60" onkeyup="document.frmReg.iGK.value='';"> <input type="button" class="button" value="확인" onClick="jsSetGiftKind();">
				<div id="spanImg"></div>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">사은품수량</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="iGKC" size="4" maxlength="10" value="1" style="text-align:right;"> 개씩
				<span id="spanKT" style="display:none;">
					<label title="같은상품증정" ><input type="checkbox" name="tmpchkKT2" onclick="jsCheckKT('tmpchkKT2',this.cheked);" value="2">1+1(동일상품) </label>
					<label title="다른상품증정" ><input type="checkbox" name="tmpchkKT3" onclick="jsCheckKT('tmpchkKT3',this.cheked);" value="3">1:1(다른상품) </label>
				</span>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">사은품한정수량</td>
			<td bgcolor="#FFFFFF">
				<input type="checkbox" name="chkLimit" onClick="jsChkLimit();">한정
				<input type="text" name="iL" size="4"  style="text-align:right;background-color:#E6E6E6;" readonly> 개(한정수량 있을 경우에만 입력)
			</td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">배송방법</td>
			<td bgcolor="#FFFFFF">
				<select name="selD">
				<!--<option value="N" >텐바이텐배송</option>-->
				<option value="Y" >업체배송</option>
				</select>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">상태</td>
			<td bgcolor="#FFFFFF" colspan="3">
				<%IF eCode <> "" THEN%>
					<input type="hidden" name="giftstatus" value="<%=eState%>"><%=replace(sStateDesc,"오픈예정","오픈")%>
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