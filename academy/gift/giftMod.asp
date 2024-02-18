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

<%
Dim clsGift ,strParm ,sOldName ,iSiteScope,sPartnerID,arrsitescope
Dim eCode, cEvent,intgroup ,iSerachType,sSearchTxt, sDate,sSdate,sEdate,iCurrpage,sgDelivery
Dim sTitle, dSDay, dEDay, sBrand, blnGroup, dOpenDay, dCloseDay
Dim gCode,igScope,ieGroupCode, igType, igR1,igR2, igStatus, dRegdate, sAdminid, igUsing
Dim igkCode, igkType, igkCnt,igkLimit, igkName,sgkImg
	gCode	  =	requestCheckVar(Request("gC"),10)
	iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
	sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'검색어
	sBrand     	= requestCheckVar(Request("ebrand"),32)		'브랜드
	sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
	sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
	sEdate     	= requestCheckVar(Request("iED"),10)		'종료일
	igStatus		= requestCheckVar(Request("giftstatus"),4)	'사은품 상태

	iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호

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

IF eCode <> "" THEN	'이벤트와 연관된 사은품일 경우
	arrsitescope = fnSetCommonCodeArr("eventscope",True) '범위 코드값에 따른 명칭 가져오기
END IF

 blngroup = False

  '공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
Dim  arrgiftstatus
arrgiftstatus 	= fnSetCommonCodeArr("giftstatus",False)
%>

<script language="javascript">

	//사은품 종류 등록 
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
			//frm.sGN.focus();
			return false;
		}

		if(!frm.sSD.value || !frm.sED.value ){
		  	alert("기간을 입력해주세요");
		  //	frm.sSD.focus();
		  	return false;
	  	}

	  	if(frm.sSD.value > frm.sED.value){
		  	alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
		  	//frm.sED.focus();
		  	return false;
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

		if(!frm.sGKN.value){
			alert("사은품 종류 입력해 주세요");
			return false;
		}

		if(!frm.iGK.value){
			alert("사은품 종류를 확인 버튼을 눌러서 확인해 주세요");
			return false;
		}
        
        <% if (igScope=1) then %>
        if (frm.chkLimit.checked){
	        alert('전체 증정 조건인 경우 한정을 체크하실 수 없습니다.');
	        return false;
	    }
	    <% end if %>
        
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
		// 그룹상품 보여주기
	  if(iVal ==4){
		document.all.dgiftgroup.style.display = "";
	  }else{
	 	document.all.dgiftgroup.style.display = "none";
	  }

	  //당첨자 대상일때 증정조건 감추기
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
	    <% if (igScope=1) then %>
	    alert('전체 증정 조건인 경우 한정을 체크하실 수 없습니다.');
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

		//제휴몰 표기
	function jsSetPartner(){
		if(document.frmReg.eventscope.options[document.frmReg.eventscope.selectedIndex].value == 3){
			document.all.spanP.style.display ="";
		}else{
			document.all.spanP.style.display ="none";
		}
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
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">사은품코드</td>
			<td bgcolor="#FFFFFF" colspan="3"><%=gCode%></td>
		</tr>
		<%IF eCode <> "" THEN%>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">이벤트코드</td>
			<td bgcolor="#FFFFFF" colspan="3"><a href="http://www.10x10.co.kr/event/eventmain.asp?eventid=<%=eCode%>" target="_blank"><%=eCode%></a></td>
		</tr>
		<%END IF%>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"> 범위</td>
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
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"> 제목</td>
			<td bgcolor="#FFFFFF"  width="400"><%IF eCode <> "" THEN %><%=sTitle%><input type="hidden" name="sGN" value="<%=sTitle%>"><%ELSE%><input type="text" name="sGN" size="30" maxlength="64" value="<%=sTitle%>"><%END IF%></td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center"> 기간</td>
			<td bgcolor="#FFFFFF">
				시작일 : <%IF eCode <> "" THEN %><%=dSDay%><input type="hidden" name="sSD" value="<%=dSDay%>"><%ELSE%><input type="text" name="sSD" size="10"   value="<%=dSDay%>"  onClick="jsPopCal('sSD');"  style="cursor:hand;"><%END IF%>
				~ 종료일 : <%IF eCode <> "" THEN %><%=dEDay%><input type="hidden" name="sED" value="<%=dEDay%>"><%ELSE%><input type="text" name="sED"  size="10"  value="<%=dEDay%>" onClick="jsPopCal('sED');" style="cursor:hand;"><%END IF%>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">대상상품</td>
			<td bgcolor="#FFFFFF"><%sbGetOptGiftCodeValue "giftscope",igScope,blngroup,"onchange='jsChkgiftgroup(this.value);'",eCode%>
			<div id="dgiftgroup" style="display:<%IF NOT (blngroup and igScope = "4") THEN%>none<%END IF%>;">
			</div>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">브랜드</td>
			<td bgcolor="#FFFFFF">
				<% drawSelectBoxLecturer "ebrand", sBrand %>				
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<div id="divType" style="display:<%IF igScope=6 THEN%>none<%END IF%>;"><!--증정조건 이벤트당첨자일 경우 증정조건 숨긴다-->
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">증정조건</td>
			<td bgcolor="#FFFFFF" width="400">
				<%sbGetOptCommonCodeArr "gifttype", igType, False,True,"onchange='jsChkGiftType(this.value);'"%>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">증정범위</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="sGR1" size="10" style="text-align:right;<%IF igType= "1" THEN%>background-color:#E6E6E6;" readonly<%ELSE%>"<%END IF%> value="<%=igR1%>"> 이상 ~ <input type="text" name="sGR2" size="10" style="text-align:right;<%IF igType= "1" THEN%>background-color:#E6E6E6;" readonly<%ELSE%>"<%END IF%> value="<%=igR2%>"> 미만
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
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">사은품종류</td>
			<td bgcolor="#FFFFFF"  width="400">
			    <input type="hidden" name="orgiGK" value="<%=igkCode%>">
				<input type="hidden" name="iGK" value="<%=igkCode%>">
				<input type="text" name="sGKN" size="40" maxlength="60" value ="<%=igkName%>" onkeyup="document.frmReg.iGK.value='';"> <input type="button" class="button" value="확인" onClick="jsSetGiftKind();">
				
				<% if (igScope=1) then %>
				<input type="button" class="button" value="관리" onClick="jsGiftKindManage();">
				<% end if %>
				
				<div id="spanImg">
				<%IF sgkImg <> "" THEN%><a href="javascript:jsImgView('<%=sgkImg%>')"><img src="<%=sgkImg%>" border="0"></a><%END IF%>
				</div>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">사은품수량</td>
			<td bgcolor="#FFFFFF" >
				<input type="text" name="iGKC" size="4" maxlength="10" value="<%=igkCnt%>" style="text-align:right;"> 개씩
				<span id="spanKT" style="display:<%IF igType = 2 THEN%>none<%END IF%>;">
					<label title="동일상품증정" ><input type="checkbox" name="tmpchkKT2" onclick="jsCheckKT('tmpchkKT2');" <%IF igScope<>5 Then%>disabled<%End IF%> value="2" <%IF igkType = 2 THEN%>checked<%END IF%>>1+1(동일상품) </label>
					<label title="다른상품증정" ><input type="checkbox" name="tmpchkKT3" onclick="jsCheckKT('tmpchkKT3');" value="3" <%IF igkType = 3 THEN%>checked<%END IF%>>1:1(다른상품) </label>
				</span>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">사은품한정수량</td>
			<td bgcolor="#FFFFFF">
				<input type="checkbox" name="chkLimit" onClick="jsChkLimit();" <%IF igkLimit <> "" THEN%>checked<%END IF%>>한정
				<input type="text" name="iL" size="4" value="<%=igkLimit%>" style="text-align:right" <%IF igkLimit ="" THEN%>style="background-color:#E6E6E6;" readonly<%END IF%>> (한정수량 있을 경우에만 입력)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">배송</td>
			<td bgcolor="#FFFFFF">
			<!--'%=fnSetDelivery(sgDelivery)%-->
				<select name="selD">
				<!--<option value="N" <%IF sgDelivery = "N" THEN%>selected<%END IF%>>텐바이텐배송</option>-->
				<option value="Y" <%IF sgDelivery = "Y" THEN%>selected<%END IF%>>업체배송</option>
				</select>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">상태</td>
			<td bgcolor="#FFFFFF">
				<%IF eCode <> "" THEN%>
					<input type="hidden" name="giftstatus" value="<%=igStatus%>"><%=replace(fnGetCommCodeArrDesc(arrgiftstatus,igStatus),"오픈예정","오픈")%>
				<%ELSE%>
					<%sbGetOptStatusCodeValue "giftstatus", igStatus, False,""%>
				<%END IF%>
				<input type="hidden" name="sOD" value="<%=dOpenDay%>">
				<input type="hidden" name="sCD" value="<%=dCloseDay%>">
				<%IF dOpenDay <> "" THEN%><span style="padding-left:10px;">오픈처리일: <%=dOpenDay%></span><%END IF%>
				<%IF dCloseDay <> "" THEN%><br><span style="padding-left:42px;">종료처리일: <%=dCloseDay%></span><%END IF%>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">사용유무</td>
			<td bgcolor="#FFFFFF">
				<input type="radio" name="sGU" value="Y" <%IF igUsing = "Y" THEN%>checked<%END IF%>>사용 <input type="radio" name="sGU" value="N" <%IF igUsing = "N" THEN%>checked<%END IF%>>사용안함
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
	    alert("다이어리 임시 사은품 증정 \n\n총 결제금액 30,000원 이상 구매시 추가됨\n\n변경시 서팀 문의 요망");
	}
	
	<% if gCode="5345" then %>
	window.onload=getOnLoad;
	<% end if %>

</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->