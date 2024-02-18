<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/eventitem_regist_mo.asp
' Description :  이벤트 등록 - 모바일 상품등록
' History : 2007.02.21 정윤정 생성
'           2008.10.20 상품이미지 크기 추가(허진원)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V2.asp"-->
<%
'변수선언
Dim eCode
Dim cEvtItem,cEvtCont,cEGroup
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate,estatedesc, ekinddesc
Dim arrGroup,arrGroup_mo
Dim iTotCnt, arrList,intLoop
Dim iPageSize, iCurrpage ,iDelCnt
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt	
Dim iDispYCnt, iDispNCnt

Dim strG, strSort
Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind,eisort
Dim strparm, Brand
dim makerid, itemname, itemid
dim  itemsort,eItemListType, blnWeb, blnMobile, blnApp
dim eChannel
 
  
strG  		= requestCheckvar(Request("selG"),10)
strSort  	= requestCheckvar(Request("selSort"),1)
	
eCode 		= requestCheckvar(request("eC"),10)
itemid      = requestCheckvar(request("itemid"),255)
itemname    = requestCheckvar(request("itemname"),64)
makerid     = requestCheckvar(request("makerid"),32)
itemsort  	= requestCheckvar(request("itemsort"),32)
eChannel    = requestCheckvar(request("eCh"),1)

	IF eCode = "" THEN	'이벤트 코드값이 없을 경우 back
%>
		<script language="javascript">
		<!--
			alert("전달값에 문제가 발생하였습니다. 관리자에게 문의해주십시오");
			history.back();
		//-->
		</script>
	<%	dbget.close()	:	response.End
	END IF	
	if eChannel = "" then eChannel = "M"
	if itemid<>"" then
	dim iA ,arrTemp,arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp) 
		if trim(arrTemp(iA))<>"" then
			'상품코드 유효성 검사(2008.08.05;허진원)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop

	itemid = left(arrItemid,len(arrItemid)-1)
end if

	'파라미터값 받기 & 기본 변수 값 세팅
	iCurrpage = Request("iC")	'현재 페이지 번호
     
	IF iCurrpage = "" THEN
		iCurrpage = 1	
	END IF	  
	
 
	iPageSize = 30		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	
	'## 검색 #############################			
	sDate = Request("selDate")  '기간 
	sSdate = Request("iSD")
	sEdate = Request("iED")	
	
	sEvt = Request("selEvt")  '이벤트 코드/명 검색
	strTxt = Request("sEtxt")
	
	sCategory	= Request("selC") '카테고리
	sState	 = Request("eventstate")'이벤트 상태	
	sKind = Request("eventkind")	'이벤트종류
	 
	strparm = "selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&eventstate="&sState&"&eventkind="&sKind
	'#######################################
	
	'데이터 가져오기
	
	'--이벤트 개요
	set cEvtCont = new ClsEvent
		cEvtCont.FECode = eCode	'이벤트 코드
		
		cEvtCont.fnGetEventCont	 '이벤트 내용 가져오기
		ekind 		=	cEvtCont.FEKind 
		ekinddesc	=	cEvtCont.FEKindDesc
		eman 		=	cEvtCont.FEManager 
		escope 		=	cEvtCont.FEScope 
		ename 		=	db2html(cEvtCont.FEName)
		esday 		=	cEvtCont.FESDay
		eeday 		=	cEvtCont.FEEDay
		epday 		=	cEvtCont.FEPDay
		elevel 		=	cEvtCont.FELevel
		estate 		=	cEvtCont.FEState
		estatedesc 	= 	cEvtCont.FEStateDesc
		eregdate 	=	cEvtCont.FERegdate
		blnWeb      =   cEvtCont.FIsWeb
        blnMobile   =   cEvtCont.FIsMobile
        blnApp      =   cEvtCont.FIsApp
        
		'이벤트 화면설정 내용 가져오기
		cEvtCont.fnGetEventDisplay		
		Brand		= 	cEvtCont.FEBrand	
		eisort 		=   cEvtCont.FEISort	
		eItemListType =   cEvtCont.FEListType
	set cEvtCont = nothing
	
	'--이벤트 상품	
	 set cEGroup = new ClsEventGroup
 		cEGroup.FECode = eCode  
 		cEGroup.FEChannel = eChannel
  		arrGroup = cEGroup.fnGetEventItemGroup	 
 	set cEGroup = nothing
 	
 	if itemsort = "" then itemsort = eisort
	set cEvtItem = new ClsEvent	
		
		cEvtItem.FPSize = iPageSize	
		cEvtItem.FECode = eCode	
		cEvtItem.FRectMakerid = makerid
		cEvtItem.FRectItemid = itemid
		cEvtItem.FRectItemName = itemname  
       
        cEvtItem.FCPage = iCurrpage
        cEvtItem.FESGroup = strG	
		cEvtItem.FESSort = itemsort	
        cEvtItem.FEChannel = eChannel
 		arrList = cEvtItem.fnGetEventItem 		
 		iTotCnt = cEvtItem.FTotCnt	'전체 데이터  수
        iDispYCnt = cEvtItem.FDispYMCnt
        iDispNCnt = cEvtItem.FDispNMCnt 
        
 	    iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수 
 	     
%>
<script language="javascript">
<!--
// 페이지 이동
function jsGoPage(iCurrpage){
    document.fitem.iC.value = iCurrpage;		
    document.fitem.submit();	
}
	
// 새상품 추가 팝업
function addnewItem(eChannel){
		var popwin;
		popwin = window.open("pop_event_additemlist.asp?eC=<%=eCode%>&makerid=<%= Brand %>&egcode=<%=strG%>&eCh="+eChannel, "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
		popwin.focus();
}
		
	
//정렬
function jsChSort(){ 
		document.fitem.submit();	
}

//그룹검색
function jsSearch(){ 
		document.fsearch.submit();	
}
	
//그룹이동	
function addGroup(){
		var frm,sValue,sGroup;		
		
		frm = document.fitem;
		sValue = "";
		sGroup =frm.eG.options[frm.eG.selectedIndex].value ;
				
		if(!frm.chkI) return;
		if(!sGroup){
		 alert("이동할 그룹이 없습니다.");
		 return;
		}
		
		if (frm.chkI.length > 1){
			for (var i=0;i<frm.chkI.length;i++){
				if(frm.chkI[i].checked){
				   if (sValue==""){
					sValue = frm.chkI[i].value;		
					}else{
					sValue =sValue+","+frm.chkI[i].value;		
					}
				}
			}	
		}else{
			sValue = frm.chkI.value;
		}
		
		if (sValue == "") {
			alert('선택 상품이 없습니다.');
			return;
		}
		
		if(confirm("그룹이동시 PC-Web 및 Mobile/App 이 함께 이동됩니다. 이동하시겠습니까?")){
		document.frmG.selGroup.value = frm.eG.options[frm.eG.selectedIndex].value;
		document.frmG.itemidarr.value = sValue;
		document.frmG.submit();
	   }
}
	
	
	
//전체선택
var ichk;
ichk = 1;
	
function jsChkAll(){			
	    var frm, blnChk;
		frm = document.fitem;
		if(!frm.chkI) return;
		if ( ichk == 1 ){
			blnChk = true;
			ichk = 0;
		}else{
			blnChk = false;
			ichk = 1;
		}
		
 		for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA		
		if ((e.type=="checkbox")) {				
			e.checked = blnChk ;
		}
	}
}

// 전체 이미지크기 일괄 변환
function jsSizeChg(selv) {
    var frm, blnChk;
	frm = document.fitem;
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.itemimgsize.length;i++){
			frm.itemimgsize[i].value=selv;
		}
	} else {
		frm.itemimgsize.value=selv;
	}
}

// 전체 전시여부  일괄 변환
function jsDispChg(selv) { 
    if(selv=="") {return;}
    var frm, blnChk;
	frm = document.fitem;
	if (frm.chkI.length > 1){
		 for (var i=0;i<frm.chkI.length;i++){
		    if (selv=="Y"){ 
			  eval("frm.eDisp_"+i)[0].checked = true;
			  eval("frm.eDisp_"+i)[1].checked = false;
		   }else{
		      eval("frm.eDisp_"+i)[0].checked = false;
		      eval("frm.eDisp_"+i)[1].checked = true;
		   }
		}
	} else {
	    if(selv=="Y") {
		  frm.eDisp_0[0].checked=true;
		  frm.eDisp_0[1].checked=false;
		}else{
		  frm.eDisp_0[0].checked=false;
		  frm.eDisp_0[1].checked=true;
		}
	}
} 
	
//삭제
function jsDel(sType, iValue){	
		var frm;		
		var sValue;		
		frm = document.fitem;
		sValue = "";
		
		if (sType ==0) {
			if(!frm.chkI) return;
			
			if (frm.chkI.length > 1){
			for (var i=0;i<frm.chkI.length;i++){
				if(frm.chkI[i].checked){
				   	if (sValue==""){
						sValue = frm.chkI[i].value;		
				   	}else{
						sValue =sValue+","+frm.chkI[i].value;		
				   	}	
				}
			}	
			}else{
				if(frm.chkI.checked){
					sValue = frm.chkI.value;
				}	
			}
		
			if (sValue == "") {
				alert('선택 상품이 없습니다.');
				return;
			}
			document.frmDel.itemidarr.value = sValue;
		}else{
			document.frmDel.itemidarr.value = iValue;
		}	
		 
		if(confirm("선택하신 상품을 삭제하시겠습니까?")){		
			document.frmDel.submit();
		}
}

// 상품 순서/이미지 사이즈 일괄 저장
function jsSortImgSize() {
	var frm;
	var sValue, sSort, sImgSize,sUsing, sSort_mo, sImgSize_mo,sUsing_mo,sDisp;
	frm = document.fitem;
	sValue = "";
	sSort = ""; 
	sDisp = ""
	var itemid;	
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){ 
			if (frm.chkI[i].checked){
			if(!IsDigit(frm.sSort[i].value)){
				alert("순서지정은 숫자만 가능합니다.");
				frm.sSort[i].focus();
				return;
			}
			 
			
		  itemid = frm.chkI[i].value;		
			if (sValue==""){
				sValue = frm.chkI[i].value;		
			}else{
				sValue =sValue+","+frm.chkI[i].value;		
			}	
			
			// 정렬순서
			if (sSort==""){
				sSort = frm.sSort[i].value;		
			}else{
				sSort =sSort+","+frm.sSort[i].value;		
			}
   
			//전시여부
 			if(sDisp ==""){
 			   if (eval("frm.eDisp_"+i)[0].checked ==true){
 			      sDisp =   eval("frm.eDisp_"+i)[0].value;
 			    }else{
 			       sDisp =   eval("frm.eDisp_"+i)[1].value;
 			    }
 			}else{
 			    if (eval("frm.eDisp_"+i)[0].checked ==true){
 			      sDisp =  sDisp+","+ eval("frm.eDisp_"+i)[0].value;
 			    }else{
 			       sDisp =  sDisp+","+ eval("frm.eDisp_"+i)[1].value;
 			    } 
 		    } 
		}
	}
	}else{
		sValue = frm.chkI.value;
		if(!IsDigit(frm.sSort.value)){
			alert("순서지정은 숫자만 가능합니다.");
			frm.sSort.focus();
			return;
		} 
		 
		sSort   = frm.sSort.value ;
       
        if (frm.eDisp_0[0].checked ==true){
 		    sDisp =  frm.eDisp_0[0].value;
 		}else{
 		    sDisp =  frm.eDisp_0[1].value;
        }
	}

		document.frmSortImgSize.itemidarr.value = sValue;
		document.frmSortImgSize.sortarr.value = sSort; 
		document.frmSortImgSize.disparr.value = sDisp;
	 	document.frmSortImgSize.submit();
}

	//그룹추가
	function jsAddGroup(eCode,gCode,eChannel){
		var winIG; 
		winIG = window.open('pop_eventitem_group.asp?eC='+eCode+'&eGC='+gCode+'&eCh='+eChannel+'&sTarget=item','popG','width=600, height=500,scrollbars=yes,resizable=yes'); 
		winIG.focus();
	}
	 
	 
	//모바일/앱 상품리스트스타일 변경
	function jsChangeListType(){
	    var i, eilt; 
	    for(i=0;i<document.fitem.itemlisttype.length;i++){ 
	        if(document.fitem.itemlisttype[i].checked){
	            eilt = document.fitem.itemlisttype[i].value;
	        }
	    }
 	  
	    document.frmLT.eILT.value = eilt;
	    document.frmLT.submit();
	}
	
	 function jsChkTrue(i){ 
	   if ( document.fitem.chkI.length > 1){ 
	     document.fitem.chkI[i].checked = true;
	}else{
	     document.fitem.chkI.checked =true;
	}
	}
//-->
</script>
<style type="text/css">
div.btmLine {background:url(/images/partner/admin_grade.png) left bottom repeat-x; padding-bottom:5px;}    
.tab {position:relative; z-index:50;}
.tab ul {_zoom:1; border-left:1px solid #ccc; border-bottom:1px solid #ccc; list-style:none; margin:0; padding:0;}
.tab ul:after {content:""; display:block; height:0; clear:both; visibility:hidden;}
.tab ul li {float:left; text-align:center; height:23px;padding-top:7px;border:1px solid #ccc; margin:0 0 -1px -1px; cursor:pointer;   background-color:#fff; }
.tab ul li.selected {background-color:#e3f1fb; position:relative; font-weight:bold;}
.col11 {width:15% !important;}
</style>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a"  style="padding-top:10px">
    <tr>
		<td style="padding-bottom:10"> 
			<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트코드</td>
					<td width="30%" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=eCode%>&nbsp;&nbsp; <a href="<%=vmobileUrl%>/event/eventmain.asp?eventid=<%=eCode%>" target="_blank">[미리보기]</a></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트명</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=ename%></td>
				</tr>	
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">종류</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
						<%=ekinddesc%>
						<% if ekind="16" then %>
							(<%= brand %>)
						<% end if %>
					</td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">상태</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=estatedesc%></td>
				</tr>
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">기간</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=esday%>~ <%=eeday%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">당첨 발표일</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=epday%></td>
				</tr>		
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트상품 정렬</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=fnGetEventCodeDesc("itemsort", eisort)%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">채널</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%if blnWeb then%>PC-WEB <%END IF%><%if blnMobile then%>Mobile <%END IF%><%if blnApp then%>App <%END IF%></td>
				</tr>
			</table>
		</td>
	</tr>	
	<tr>
	    <td>
	        <div class="tab btmLine">
				<ul style="margin-left:-1px">
					<li class="col11"  onclick="location.href='eventitem_regist.asp?eC=<%=eCode%>&menupos=<%=menupos%>'">PC_WEB 상품등록</li>
            		<li class="col11 selected">Moblie/ APP 상품등록</li>
				</ul>
			</div>
		</td>
	</tr>	
	<tr><!-- 검색--->
		<td>
		    <table cellspacing="5"  bgcolor="#e3f1fb" width="100%" class="a" cellpadding="0">
		        <tr>
		            <td bgcolor="#FFFFFF">	 
            			<form name="fsearch" method="post" action="eventitem_regist_mo.asp"> 
            				<input type="hidden" name="eC" value="<%=eCode%>">
            				<input type="hidden" name="eCh" value="<%=eChannel%>">
            				<input type="hidden" name="menupos" value="<%=menupos%>">
            				<input type="hidden" name="mode" value="">
            				<input type="hidden" name="selGroup" value="">
            				<input type="hidden" name="itemsort" value="<%=itemsort%>">
            			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
            				<tr align="center" >
            					<td  width="100" bgcolor="<%= adminColor("tabletop") %>">검색 조건</td>
            					<td align="left"  bgcolor="#ffffff">  	 
            						<table border="0" cellpadding="1" cellspacing="1" class="a">
                						<tr>
                							<td style="white-space:nowrap;">그룹: 
                								<select name="selG" onChange="jsSearch();" class="select">
                						        	<option value="">전체</option>			        	
                						       	<%IF isArray(arrGroup) THEN %>
                						       		<option value="0"  <%IF Cstr(strG) = "0" THEN%>selected<%END IF%>>미지정</option>
                						       	<%	dim sumi, i
                						       		For intLoop = 0 To UBound(arrGroup,2)
                						       		 sumi = 0
                						       	%>
                						       		<option value="<%=arrGroup(9,intLoop)%>" <%IF Cstr(strG) = Cstr(arrGroup(9,intLoop)) THEN %> selected<%END IF%>  <%if not arrGroup(8,intLoop) then%>style="color:gray;"<%end if%>>
                						       		    <%IF arrGroup(5,intLoop) <> 0 THEN%>└&nbsp;<%END IF%><%=arrGroup(0,intLoop)%>(<%=arrGroup(1,intLoop)%>)
                						       		    <% if intLoop < UBound(arrGroup,2)  then 
					                                       for i = 1 to (UBound(arrGroup,2)-intLoop) 
                                    					     if arrGroup(9,intLoop) = arrGroup(9,intLoop+i) then
                                    					        sumi = sumi + 1  
                                    					         %>
                                    					    + <%=arrGroup(0,intLoop+i)%>(<%=arrGroup(1,intLoop+i)%>)
                                    					<%   else 
                                    					        exit for
                                    					     end if 
                                    					    next
                                    					   end if 
                                    					     %>
                						       		    <%if not arrGroup(8,intLoop) then%> -[전시안함]<%end if%>
                						       		 </option>
                						    	<%	 intLoop = intLoop+sumi
                						    	    Next 
                						    	END IF%>	
                						       	</select> 	
                			       			</td> 
                							<td style="white-space:nowrap;padding-left:10px;">브랜드: <%	drawSelectBoxDesignerWithName "makerid", makerid %></td>  
                							<td style="white-space:nowrap;padding-left:10px;">상품코드:</td>
                							<td style="white-space:nowrap;" rowspan="2"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea> </td>
                						</tr> 
                						<tr>
                						    <td colspan="4">상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20"></td>
            						</tr>
            			        	</table>        			        	
            			        </td>
            			        <td   width="100" bgcolor="<%= adminColor("gray") %>">
            						<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch();">
            					</td>
            			    </tr> 
            			</table>
            			</form>
            		</td>
            	</tR> <!-- 검색---> 
            	<tr>
            		<td style="padding-top:10px;" valign="top">  
            		     <div id="divMA">
            		       <form name="fitem" method="post" action="eventitem_regist_mo.asp">
                           <input type="hidden" name="menupos" value="<%=menupos%>">
                           <input type="hidden" name="mode" value="">
                           <input type="hidden" name="iC" value="">
                           <input type="hidden" name="eC" value="<%=eCode%>"> 
                           <input type="hidden" name="eCh" value="<%=eChannel%>"> 
                           <input type="hidden" name="selGroup" value="">
                           <input type="hidden" name="selG" value="<%=strG%>">
                           <input type="hidden" name="makerid" value="<%=makerid%>">
                           <input type="hidden" name="itemname" value="<%=itemname%>">
                           <input type="hidden" name="itemid" value="<%=itemid%>">
 
            		        <table width="100%" border="0" align="center" cellpadding="0"  class="a" cellspacing="1">	  
            		            <tr>
            		                <td>
                                 	    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" class="a" >		
                                 	        <tr>
                                 	           <td colspan="2">상품 리스트 스타일:	
                                 		     		<input type="radio" name="itemlisttype"  value="1"  <%IF eItemListType = "1" THEN%>checked<%END IF%>>격자형 
                                 				    <input type="radio" name="itemlisttype"  value="2"  <%IF eItemListType = "2" THEN%>checked<%END IF%>>리스트형 
                                 				    <input type="radio" name="itemlisttype"  value="3"  <%IF eItemListType = "3" THEN%>checked<%END IF%>>BIG형	
                                 				    &nbsp;&nbsp;<input type="button" value="변경" class="button" onClick="jsChangeListType();" style="width:60px;">
                                 			    </td> 
                                 	        </tr>
                                 		    <tr>
                                 		        
                                     		      <td>
                                 	                <input type="button" value="선택삭제" onClick="jsDel(0,'');" class="button">&nbsp;&nbsp;&nbsp;    
                                     		     	<select name="eG" class="select">
                                     		     	<%IF isArray(arrGroup) THEN %>
                          			            	<option value=""> --선택--</option>
                          			       	    <%
                						       		    For intLoop = 0 To UBound(arrGroup,2)
                						       		     
                						       	%>
                						       		<option value="<%=arrGroup(0,intLoop)%>" <%IF Cstr(strG) = Cstr(arrGroup(0,intLoop)) THEN %> selected<%END IF%>  <%if not arrGroup(8,intLoop) then%>style="color:gray;"<%end if%>>
                						       		    <%IF arrGroup(5,intLoop) <> 0 THEN%>└&nbsp;<%END IF%><%=arrGroup(0,intLoop)%>(<%=arrGroup(1,intLoop)%>)
                						       		     
                						       		    <%if not arrGroup(8,intLoop) then%> -[전시안함]<%end if%>
                						       		 </option>
                						    	<%	  
                						    	    Next 
                						     
                                         		  	  ELSE	
                                         		  	%>
                                     		  	        <option value=""> --그룹없음--</option>
                                     		  	    <%END IF%>	
                                 		     	    </select>   
                                 		     		<input type="button" value="선택그룹이동" onClick="addGroup();" class="button">
                                 		     			&nbsp; 	<input type="button" value="그룹추가" onClick="jsAddGroup('<%=eCode%>','','<%=eChannel%>');" class="button">   
                                 		     		 
                                 		  	    </td>	
                                 		  	    <td align="right"> 
                                 		            <input type="button" value="선택상품 수정" onClick="jsSortImgSize();" class="button">&nbsp; 
                                 		     	    <input type="button" value="새상품 추가" onclick="addnewItem('<%=eChannel%>');" class="button"> 
                                                </td>			      
                                 		    </tr>
                                 		</table>
                                    </td>
                                </tr> 
                                <tr>
                      	            <td> 
                      			        <table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>"> 
                                         <tr bgcolor="#FFFFFF">
                                     	    <td  colspan="20" >
                                     	        <table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="0" > 
                                     	            <tr>
                                         		        <td align="left">[검색결과] <font color="blue">전시 Y: <%=iDispYCnt%></font> /  <font color="red">전시 N: <%=iDispNCnt%></font> / <b>총: <%=iTotCnt%></b>&nbsp;&nbsp;&nbsp;&nbsp;페이지: <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
                                         		        <td align="right">정렬 : <%sbGetOptCommonCodeArr  "itemsort",itemsort,False,False,"onchange='jsChSort();'"%></td>
                                         		    </tr>
                                         		</table>
                                         	</td>
                                     	</tr>
                                        <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
                                     		<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>    				    	
                                     		<td>그룹코드</td>
                                     		<td>상품ID</td>
                                     		<td>이미지</td>
                                     		<td>브랜드</td>
                                     		<td>상품명</td>
                                     		<td>판매가</td>
                                     		<td>매입가</td>
                                     		<td>할인율</td>
                                     		<td>배송</td>	
                                     		<td>판매여부</td>	
                                     		<td>상품사용여부</td>	
                                     		<td>한정여부</td> 
                                     		<td>순서</td>  
                                     		<td> <select name="selDisp" class="select" onChange="jsDispChg(this.value);">
                      			    	        <option value="">전시여부</option>
                      			    	        <option value="Y">Y</option>
                      			    	        <option value="N">N</option>
                      			    	    </select></td> 
                                     	</tr> 
                                     		<%IF isArray(arrList) THEN 
                                     			For intLoop = 0 To UBound(arrList,2)
                                     		%>
                                     	<tr align="center" bgcolor="<%if  arrList(29,intLoop) then%>#FFFFFF<%else%>gray<%end if%>">    
                                     		<td><input type="checkbox" name="chkI" value="<%=arrList(0,intLoop)%>"></td>    				    	
                                     		<td><%IF arrList(1,intLoop) <> 0 THEN%><%=arrList(1,intLoop)%><%END IF%></td>		    				    	
                                     		<td>
                                     				<!-- 2007/05/05 김정인 수정 -- 품절 표시 -->			    		
                                     				<A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a>
                                     				<% if cEvtItem.IsSoldOut(arrList(14,intLoop),arrList(16,intLoop),arrList(20,intLoop),arrList(21,intLoop)) then %>
                                     					<br><img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
                                     				<% end if %>
                                     		</td>
                                     		<td><% if (Not IsNull(arrList(12,intLoop)) ) and (arrList(12,intLoop)<>"") then %>
                                     		     <img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(arrList(0,intLoop))%>/<%=arrList(12,intLoop)%>">
                                     		    <%end if%>
                                     		</td>    	
                                     		<td><%=db2html(arrList(3,intLoop))%></td>
                                     		<td align="left">&nbsp;<%=db2html(arrList(4,intLoop))%></td>
                                     		<td><%
                                     			Response.Write FormatNumber(arrList(7,intLoop),0)
                                     			'할인가
                                     			if arrList(18,intLoop)="Y" then
                                     				Response.Write "<br><font color=#F08050>(할)" & FormatNumber(arrList(9,intLoop),0) & "</font>"
                                     			end if
                                     			'쿠폰가
                                     			if arrList(22,intLoop)="Y" then
                                     				Select Case arrList(23,intLoop)
                                     					Case "1"
                                      						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(5,intLoop)*((100-arrList(24,intLoop))/100),0) & "</font>"
                                     					Case "2"
                                     						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(5,intLoop)-arrList(24,intLoop),0) & "</font>"
                                     				end Select
                                     			end if
                                     		%></td>
                                     		<td><%
                                     			Response.Write FormatNumber(arrList(8,intLoop),0)
                                     			'할인가
                                     			if arrList(18,intLoop)="Y" then
                                     				Response.Write "<br><font color=#F08050>" & FormatNumber(arrList(10,intLoop),0) & "</font>"
                                     			end if
                                     			'쿠폰가
                                     			if arrList(22,intLoop)="Y" then
                                     				if arrList(23,intLoop)="1" or arrList(23,intLoop)="2" then
                                     					if arrList(25,intLoop)=0 or isNull(arrList(25,intLoop)) then
                                     						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(8,intLoop),0) & "</font>"
                                     					else
                                     						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(25,intLoop),0) & "</font>"
                                     					end if
                                     				end if
                                     			end if
                                     		%></td>
                                     			<td><%if arrList(18,intLoop)="Y" then%>
                                     						<font color=#F08050><%=formatnumber(((arrList(7,intLoop)-arrList(9,intLoop))/arrList(7,intLoop))*100,0)%>%</font>
                                     						 
                                     						<%end if%>
                                     				<%if arrList(22,intLoop)="Y" then 
                      						if arrList(23,intLoop)="1" or arrList(23,intLoop)="2" then
                      					        if arrList(25,intLoop)=0 or isNull(arrList(25,intLoop)) then
                      						         Response.Write "<br><font color=#5080F0>" & FormatNumber( arrList(8,intLoop),0) & "</font>"
                      					        else
                      						        Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(24,intLoop),0) 
                      						         if arrList(23,intLoop)="1" then 
                      						         Response.Write "%"
                      						        else
                      						         Response.Write "원"
                      						        end if
                      						         Response.Write "</font>"
                      					        end if
                      				        end if
                      						 end if%>		
                                     			</td>
                                     		   	<td><%= fnColor(cEvtItem.IsUpcheBeasong(arrList(15,intLoop)),"delivery")%></td>    	
                                     			<td><%= fnColor(arrList(14,intLoop),"yn") %></td>
                                     			<td><%= fnColor(arrList(19,intLoop),"yn") %></td>
                                     			<td><%= fnColor(arrList(16,intLoop),"yn") %></td>    				    	
                                     			<td><input type="text" name="sSort" value="<%=arrList(2,intLoop)%>" size="4" style="text-align:right;"></td> 
                                     			<td><input type="radio" name="eDisp_<%=intLoop%>" value="1" <%if arrList(29,intLoop) then%>checked<%end if%> onClick="jsChkTrue('<%=intLoop%>');"><%if arrList(29,intLoop) then%><font color="#5080F0"><%end if%>Y </font>
                      			    	    <input type="radio" name="eDisp_<%=intLoop%>" value="0" <%if not arrList(29,intLoop) then%>checked<%end if%> onClick="jsChkTrue('<%=intLoop%>');"><%if not arrList(29,intLoop) then%><font color="red"><%end if%>N</font></td>
                                     			<!--td><input type="button" value="삭제" onClick="jsDel(1,<%=arrList(0,intLoop)%>);" class="button"></td-->	
                                     		</tr>   
                                     			   <%	Next
                                     			   	ELSE
                                     			   %>
                                     		<tr  align="center" bgcolor="#FFFFFF">
                                     			<td colspan="19">등록된 내용이 없습니다.</td>
                                     		</tr>	
                                     			   <%END IF%>
                                     	</table>
                                    </td>
                               </tr>
                               <tr>
                                    <td> <!-- 페이징처리 -->
                                          <%		
                                     	iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1	
                                     	
                                     	If (iCurrpage mod iPerCnt) = 0 Then																
                                     		iEndPage = iCurrpage
                                     	Else								
                                     		iEndPage = iStartPage + (iPerCnt-1)
                                     	End If	
                                     	%>
                                     	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a"  >
                                     	    <tr valign="bottom" height="25">			      
                                     	        <td valign="bottom" align="center">
                                     	         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
                                     			<% else %>[pre]<% end if %>
                                     	        <%
                                     				for ix = iStartPage  to iEndPage
                                     					if (ix > iTotalPage) then Exit for
                                     					if Cint(ix) = Cint(iCurrpage) then
                                     			%>
                                     				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
                                     			<%		else %>
                                     				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
                                     			<%
                                     					end if
                                     				next
                                     			%>
                                     	    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
                                     			<% else %>[next]<% end if %>
                                     	        </td> 
                                     	    </tr>			  
                                     	</form>    
                                     	</table>
                                    </td>
                                </tr>
                            </table> 
                        </div> 
            		</td>
            	</tR> 
	        </table>
	    </td>
	</tr>
	<tr>
	    <td   align="right"><a href="index.asp?menupos=<%=menupos%>&<%=strparm%>"><img src="/images/icon_list.gif" border="0"></a></td>			        
	 </tr>
</table>
<%
	set cEvtItem = nothing
%>	
<!-- 그룹이동--->
<form name="frmG" method="post" action="eventitem_process.asp">
<input type="hidden" name="mode" value="G">
<input type="hidden" name="iC" value="<%=iCurrpage%>">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="selG" value="<%=strG%>">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="selGroup" value="">
<input type="hidden" name="eCh" value="<%=eChannel%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<!-- 선택삭제--->
<form name="frmDel" method="post" action="eventitem_process.asp">
<input type="hidden" name="mode" value="D">
<input type="hidden" name="iC" value="<%=iCurrpage%>">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="selG" value="<%=strG%>">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="eCh" value="<%=eChannel%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<!-- 순서 및 이미지크기 변경--->
<form name="frmSortImgSize" method="post" action="eventitem_process.asp">
<input type="hidden" name="mode" value="S">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="selG" value="<%=strG%>">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="sortarr" value="">
<input type="hidden" name="sizearr" value=""> 
<input type="hidden" name="disparr" value="">
<input type="hidden" name="makerid" value="<%=makerid%>">
<input type="hidden" name="itemname" value="<%=itemname%>">
<input type="hidden" name="itemid" value="<%=itemid%>"> 
<input type="hidden" name="itemsort" value="<%=itemsort%>">   
<input type="hidden" name="eCh" value="<%=eChannel%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
 
<form name="frmLT" method="post" action="eventitem_process.asp">
<input type="hidden" name="mode" value="L">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="eILT" value="">
<input type="hidden" name="selG" value="<%=strG%>">
<input type="hidden" name="makerid" value="<%=makerid%>">
<input type="hidden" name="itemname" value="<%=itemname%>">
<input type="hidden" name="itemid" value="<%=itemid%>"> 
<input type="hidden" name="itemsort" value="<%=itemsort%>">   
<input type="hidden" name="eCh" value="<%=eChannel%>">   
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<!-- 표 하단바 끝-->		
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
