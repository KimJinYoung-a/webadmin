<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  할인 관리 등록
' History : 2010.09.28 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->
<!-- #include virtual="/academy/lib/classes/sale/salecls.asp"-->
<%
Dim sMode ,cEGroup,blngroup,arrGroup,intgroup ,strParm ,iSerachType,sSearchTxt, sDate,sSdate,sEdate,iCurrpage,sStatus
Dim sCode, clsSale,isRate, isMargin, isStatus, egCode, isUsing, dOpenDay,isMValue,dCloseDay	
Dim eCode, cEventsale ,sTitle, dSDay, dEDay, sBrand,eState
	eCode     = requestCheckVar(Request("eC"),10)
	sCode     = requestCheckVar(Request("sC"),10)
	isRate = 0
	isUsing = true
	sMode  = "I"
	isStatus =0

IF sCode <> "" THEN
	set clsSale = new CSale
	sMode = "U"
	clsSale.FSCode  = sCode		
	clsSale.fnGetSaleConts
	
	sTitle 		= clsSale.FSName 		
	isRate 		= clsSale.FSRate 		
	isMargin 	= clsSale.FSMargin 	
	eCode 		= clsSale.FECode 		
	egCode		= clsSale.FEGroupCode 
	dSDay 		= clsSale.FSDate 		
	dEDay 		= clsSale.FEDate		
	isStatus 	= clsSale.FSStatus 	
	isUsing     = clsSale.FSUsing 	
	dOpenDay	= clsSale.FOpenDate
	isMValue	= clsSale.FSMarginValue
	dCloseDay 	= clsSale.FCloseDate
	
	'-검색----------------------------------------
	 iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
	 sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'검색어
	 sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
	 sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
	 sEdate     	= requestCheckVar(Request("iED"),10)		'종료일
	 sStatus		= requestCheckVar(Request("salestatus"),4)	' 상태 
	 iCurrpage		= requestCheckVar(Request("iC"),10)	'현재 페이지 번호
	 
	 strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&sstatus="&sStatus
	'---------------------------------------------   	
	set clsSale = nothing
END IF
IF eCode = "0" THEN eCode = ""
IF eCode <> "" THEN		'이벤트 연관 일경우
	IF sCode = "" THEN
	set cEventsale = new ClsEventSummary
		cEventsale.FECode = eCode
		cEventsale.fnGetEventConts
		sTitle 	= cEventsale.FEName
		dSDay	= cEventsale.FESDay
		dEDay	= cEventsale.FEEDay
		sBrand	= cEventsale.FBrand	
		isStatus= cEventsale.FEState
		dOpenDay= cEventsale.FEOpenDate			
	set cEventsale = nothing
   END IF
	set cEGroup = new ClsEventGroup
	 	cEGroup.FECode = eCode  	
	  	arrGroup = cEGroup.fnGetEventItemGroup		
	set cEGroup = nothing
	 
	 blngroup = False
	 IF isArray(arrGroup) THEN blngroup = True
END IF	
	IF dSDay ="" THEN dSDay = date()
	IF isStatus < 6 THEN isStatus = 0	
	'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
	Dim  arrsalestatus	
	arrsalestatus = fnSetCommonCodeArr("salestatus",False)
%>

<script language="javascript">
			
	//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
	
	function jsSubmitSale(){
		var frm = document.frmReg;
		
		if(!frm.sSN.value){
			alert("제목을 입력해 주세요");
			frm.sSN.focus();
			return false;
		}
		
		if(!frm.sSD.value ){
		  	alert("시작일을 입력해주세요");
		  	frm.sSD.focus();
		  	return false;
	  	}
	  	
	  	if(frm.sED.value){
		  	if(frm.sSD.value > frm.sED.value){
			  	alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
			  	frm.sED.focus();
			  	return false;
		  	}	
		}	
		
		
		
		if(typeof(frm.chkstatus)=="object"){
			if(frm.chkstatus.checked) {
				frm.salestatus.value = frm.chkstatus.value;
			}
		}
	
		var nowDate = "<%=date()%>";	   
	   if(frm.salestatus.value==7){
	 	if(frm.sOD.value !=""){		  
	 		nowDate = '<%IF dOpenDay <> ""THEN%><%=FormatDate(dOpenDay,"0000-00-00")%><%END IF%>';
		}
		
		if(frm.sSD.value < nowDate){
			alert("시작일이 오픈일보다  빠르면 안됩니다. 시작일을 다시 선택해주세요");		  	
		  	return false;
		 }
	  }
	  	
	  	if(!frm.iSR.value){
			alert("할인율을 입력해 주세요");
			frm.iSR.focus();
			return false;
		}
		
		
	}
	
	function jsChSetValue(iVal){
		if(iVal ==5){
			document.all.divM.style.display = "";
		}else{
			document.all.divM.style.display = "none";
		}
	}

</script>

<table width="900" border="0" align="left" class="a" cellpadding="3" cellspacing="1"  >
<form name="frmReg" method="post" action="saleProc.asp?<%=strParm%>" onSubmit="return jsSubmitSale();">
<input type="hidden" name="sM" value="<%=sMode%>">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="sC" value="<%=sCode%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr>
	<td>
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<%IF eCode <> "" THEN%>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">이벤트코드(그룹)</td>
			<td bgcolor="#FFFFFF" colspan="3"><%=eCode%>
			<div id="dgiftgroup" style="display:<%IF NOT blngroup THEN%>none<%END IF%>;">
			<%IF isArray(arrGroup) THEN%>
				그룹선택: <select name="selG">
			   	<%	
			   		For intgroup = 0 To UBound(arrGroup,2)
			   	%>
			   		<option value="<%=arrGroup(0,intgroup)%>" <%IF Cstr(egCode) = Cstr(arrGroup(0,intgroup)) THEN %> selected<%END IF%>> <%=arrGroup(0,intgroup)%>(<%=db2html(arrGroup(1,intgroup))%>)</option>
				<%	Next 
				%>	
			   	</select> 
			 <%ELSE%>  	
			 <input type="hidden" name="selG" value="0">  	
			 <%END IF%>  	
			</div>			
			</td>
		</tr>	
		<%END IF%>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center"> 제목</td>
			<td bgcolor="#FFFFFF"><%IF eCode <> "" THEN %><%=sTitle%><input type="hidden" name="sSN" value="<%=sTitle%>"><%ELSE%><input type="text" name="sSN" size="30" maxlength="64" value="<%=sTitle%>"><%END IF%></td>	
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center"> 기간</td>
			<td bgcolor="#FFFFFF">
				시작일 : <%IF eCode <> "" THEN %><%=dSDay%><input type="hidden" name="sSD" value="<%=dSDay%>"><%ELSE%><input type="text" name="sSD" size="10"   onClick="jsPopCal('sSD');"  style="cursor:hand;" value="<%=dSDay%>"><%END IF%> 
				~ 종료일 : <%IF eCode <> "" THEN %><%=dEDay%><input type="hidden" name="sED" value="<%=dEDay%>"><%ELSE%><input type="text" name="sED"  size="10" onClick="jsPopCal('sED');" style="cursor:hand;" value="<%=dEDay%>"><%END IF%>	
			</td>	
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"> 할인율</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iSR" size="4"  value="<%=isRate%>" style="text-align:right;">%</td>	
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center">마진구분</td>
			<td bgcolor="#FFFFFF"><%sbGetOptCommonCodeArr "salemargin", isMargin, False,True,"onchange='jsChSetValue(this.value);'"%>
			<span id="divM" style="display:<%IF isMargin<> 5 THEN %>none<%END IF%>;">할인마진<input type="text" size="4" name="isMV" maxlength="10" value="<%=isMValue%>" style="text-align:right;">%</span>
			</td>	
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"> 상태</td>
			<td bgcolor="#FFFFFF" >				
					<input type="hidden" name="sOD" value="<%=dOpenDay%>">
					<input type="hidden" name="salestatus" value="<%=isStatus%>">						
					<%=fnGetCommCodeArrDesc(arrsalestatus,isStatus)%>								
				<%if eCode = "" then%>	
				<%IF isStatus =0 then '등록대기 %>						
					<input type="checkbox" name="chkstatus" value="7">오픈요청  													
				<%elseif isStatus = 6 or isStatus = 7 then '오픈 %>						
					<input type="checkbox" name="chkstatus" value="9">종료요청										
				<%elseif isStatus = 8 then %>	
					<div style="padding-top:5px;">종료일: <%=dCloseDay%></div> 
				<%end if%>
				<%end if%>
			</td>	
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">사용유무</td>
			<td bgcolor="#FFFFFF">
				<input type="radio" name="sSU" value="1" <%IF isUsing THEN%>checked<%END IF%>>사용 <input type="radio" name="sSU" value="0" <%IF not isUsing  THEN%>checked<%END IF%>>사용안함
			</td>	
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<input type="image" src="/images/icon_save.gif"> 
		<a href="javascript:history.back();"><img src="/images/icon_cancel.gif" border="0"></a>
		<a href="saleList.asp?menupos=<%=menupos%>"><img src="/images/icon_list.gif" border="0"></a>
	</td>
</tr>
</form>
</table>
<iframe name="dispcate_item" id="dispcate_item" src="saleItemReg.asp?sC=<%=sCode%>" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->