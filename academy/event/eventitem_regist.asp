<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이벤트 등록 - 상품등록
' History : 2010.09.29 한용민 생성
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
<%
Dim cEvtItem,cEvtCont,cEGroup ,eCode ,arrGroup ,iTotCnt, arrList,intLoop ,iPageSize, iCurrpage ,iDelCnt
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate,estatedesc, ekinddesc ,strparm
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt ,strG, strSort	,sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
	eCode = RequestCheckvar(Request("eC"),10)
	strG  = RequestCheckvar(Request("selG"),10)
	strSort  = RequestCheckvar(Request("selSort"),10)
	
	IF eCode = "" THEN	'이벤트 코드값이 없을 경우 back
%>
		<script language="javascript">
			alert("전달값에 문제가 발생하였습니다. 관리자에게 문의해주십시오");
			history.back();
		</script>
<%	
		dbget.close()	:	response.End
	END IF	
	
	'파라미터값 받기 & 기본 변수 값 세팅
	iCurrpage = RequestCheckvar(Request("iC"),10)	'현재 페이지 번호
	IF iCurrpage = "" Then
			iCurrpage = 1	
	END IF	  
		
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	
	'## 검색 #############################			
	sDate = RequestCheckvar(Request("selDate"),32)  '기간 
	sSdate = RequestCheckvar(Request("iSD"),32)
	sEdate = RequestCheckvar(Request("iED"),32)
	
	sEvt = RequestCheckvar(Request("selEvt"),10)  '이벤트 코드/명 검색
	strTxt = Request("sEtxt")
	
	sCategory	= RequestCheckvar(Request("selC"),10) '카테고리
	sState	 = RequestCheckvar(Request("eventstate"),10)'이벤트 상태	
	sKind = RequestCheckvar(Request("eventkind"),10)	'이벤트종류
		
	strparm = "selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&eventstate="&sState&"&eventkind="&sKind
	'#######################################
	
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
	estatedesc 	= cEvtCont.FEStateDesc
	eregdate 	=	cEvtCont.FERegdate 
set cEvtCont = nothing

'--이벤트 상품	
 set cEGroup = new ClsEventGroup
	cEGroup.FECode = eCode  	
	arrGroup = cEGroup.fnGetEventItemGroup		
set cEGroup = nothing

set cEvtItem = new ClsEvent	
	cEvtItem.FCPage = iCurrpage
	cEvtItem.FPSize = iPageSize	
	cEvtItem.FECode = eCode	
	cEvtItem.FESGroup = strG	
	cEvtItem.FESSort = strSort	
			
	arrList = cEvtItem.fnGetEventItem 		
	iTotCnt = cEvtItem.FTotCnt	'전체 데이터  수

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수 
%>

<script language="javascript">

// 페이지 이동
function jsGoPage(iP){
		document.fitem.iC.value = iP;		
		document.fitem.submit();	
}
	
// 새상품 추가 팝업
function addnewItem(){
		var popwin;
		popwin = window.open("/academy/event/common/pop_event_additemlist.asp?eC=<%=eCode%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		popwin.focus();
}
		
	
//정렬
function jsChSort(){
		document.fitem.submit();	
}

//그룹검색
function jsSearchGroup(){
		document.fitem.submit();	
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
		
		document.frmG.selGroup.value = frm.eG.options[frm.eG.selectedIndex].value;
		document.frmG.itemidarr.value = sValue;
		document.frmG.submit();
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
	var sValue, sSort, sImgSize;
	frm = document.fitem;
	sValue = "";
	sSort = "";
	sImgSize = "";
		
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(!IsDigit(frm.sSort[i].value)){
				alert("순서지정은 숫자만 가능합니다.");
				frm.sSort[i].focus();
				return;
			}
			
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

			// 이미지 사이즈
			if (sImgSize==""){
				sImgSize = frm.itemimgsize[i].value;		
			}else{
				sImgSize =sImgSize+","+frm.itemimgsize[i].value;		
			}	
		}
	}else{
		sValue = frm.chkI.value;
		if(!IsDigit(frm.sSort.value)){
			alert("순서지정은 숫자만 가능합니다.");
			frm.sSort.focus();
			return;
		}
		sSort =  frm.sSort.value; 
		sImgSize =  frm.itemimgsize.value; 
	}

		document.frmSortImgSize.itemidarr.value = sValue;
		document.frmSortImgSize.sortarr.value = sSort;
		document.frmSortImgSize.sizearr.value = sImgSize;
		document.frmSortImgSize.submit();
}

	//그룹추가
	function jsAddGroup(){
		var winIG;
		winIG = window.open('iframe_eventitem_group.asp?ec=<%=eCode%>&T=1','popIG','width=700,height=600,scrollbars=yes');
		winIG.focus();
	}

</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding-bottom:10"> 
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트코드</td>
				<td width="30%" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=eCode%></td>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트명</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=ename%></td>
			</tr>	
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">종류</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=ekinddesc%></td>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">상태</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=estatedesc%></td>
			</tr>
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">기간</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=esday%>~ <%=eeday%></td>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">당첨 발표일</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=epday%></td>
			</tr>			
		</table>
	</td>
</tr>	
		
<tr>
	<td >
		<table width="100%" align="center" cellpadding="3" cellspacing="0" class="a">
			<form name="fitem" method="post" action="eventitem_regist.asp">
			<input type="hidden" name="iC" value="">
			<input type="hidden" name="eC" value="<%=eCode%>">
			<input type="hidden" name="menupos" value="<%=menupos%>">
			<input type="hidden" name="mode" value="">
			<input type="hidden" name="selGroup" value="">
			<tr align="center"  >
				<td align="left">  		        
		        	 그룹검색
		        	<select name="selG" onChange="jsSearchGroup();">
		        	<option value="">전체</option>			        	
		       	<%IF isArray(arrGroup) THEN %>
		       		<option value="0"  <%IF Cstr(strG) = "0" THEN%>selected<%END IF%>>미지정</option>
		       	<%	
		       		For intLoop = 0 To UBound(arrGroup,2)
		       	%>
		       		<option value="<%=arrGroup(0,intLoop)%>" <%IF Cstr(strG) = Cstr(arrGroup(0,intLoop)) THEN %> selected<%END IF%>> <%=arrGroup(0,intLoop)%>(<%=arrGroup(1,intLoop)%>)</option>
		    	<%	Next 
		    	END IF%>	
		       	</select> 			           			        	
		        </td>
		        <td align="right">
		         정렬 : <select name="selSort" onchange="jsChSort();">			         			         		       					       		
		       		<option value="sitemid" >신상품순</option>			       					       		
		       		<option value="sevtitem" <%IF Cstr(strSort) = "sevtitem" THEN %>selected<%END IF%>>순서순</option>
		       		<option value="sbest" <%IF Cstr(strSort) = "sbest" THEN %>selected<%END IF%>>베스트셀러순</option>	
		       		<option value="shsell" <%IF Cstr(strSort) = "shsell" THEN %>selected<%END IF%>>높은가격순</option>			       	
		       		<option value="slsell" <%IF Cstr(strSort) = "slsell" THEN %>selected<%END IF%>>낮은가격순</option>	
		       		<option value="sevtgroup" <%IF Cstr(strSort) = "sevtgroup" THEN %>selected<%END IF%>>그룹순</option>
		       		<option value="sbrand" <%IF Cstr(strSort) = "sbrand" THEN %>selected<%END IF%>>브랜드</option>
		       		</select>			       		
		        </td>			       
			</tr>
		</table>
	</td>
</tR>
		 
<tr>
	<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">												
		    <tr height="35">			      
		        <td align="left">       	
		       	<input type="button" value="선택삭제" onClick="jsDel(0,'');" class="button">&nbsp;&nbsp;&nbsp;      
		       	<select name="eG">
		       	<%IF isArray(arrGroup) THEN
		       		For intLoop = 0 To UBound(arrGroup,2)
		       	%>
		       		<option value=" <%=arrGroup(0,intLoop)%>" ><%IF arrGroup(5,intLoop) <> 0 THEN%>└&nbsp;<%END IF%><%=arrGroup(0,intLoop)%>(<%=arrGroup(1,intLoop)%>)</option>
		    	<%	Next 
		    	  ELSE	
		    	%>
		    	<option value=""> --그룹없음--</option>
		    	<%END IF%>	
		       	</select>
		       		<input type="button" value="선택그룹이동" onClick="addGroup();" class="button">
		       			&nbsp; 	<input type="button" value="그룹추가" onClick="jsAddGroup();" class="button">   
		    	</td>
		    	<td align="right">
		    	<input type="button" value="순서/사이즈 저장" onClick="jsSortImgSize();" class="button">&nbsp; 
		       	<input type="button" value="새상품 추가" onclick="addnewItem();" class="button">
		       	
		        </td>			      
		    </tr>
		</table>
	</td>
</tr>
		 
<tr>
	<td> 
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	    <tr bgcolor="#FFFFFF">
	   		<td colspan="15" align="left">검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
	   	</tr>
	    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    	<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>    				    	
	    	<td>그룹코드</td>
	    	<td align="center">상품ID</td>
			<td align="center">이미지</td>
			<td align="center">브랜드</td>
			<td align="center">상품명</td>
			<td align="center">판매가</td>
			<td align="center">매입가</td>
			<td align="center">배송</td>	
			<td align="center">판매여부</td>	
			<td align="center">사용여부</td>	
			<td align="center">한정여부</td>	
	    	<td>순서</td>
	    	<td>이미지크기<br>
				<select name="selimgsize" onchange=jsSizeChg(this.value)>	
				<option value="1">100px</option>
				<option value="2">200px</option>
				</select>
	    	</td>
	    	<td>비고</td>
	    </tr>
	    <%IF isArray(arrList) THEN 
	    	For intLoop = 0 To UBound(arrList,2)
	    %>
	    <tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';> 
	    	<td><input type="checkbox" name="chkI" value="<%=arrList(0,intLoop)%>"></td>    				    	
	    	<td><%IF arrList(1,intLoop) <> 0 THEN%><%=arrList(1,intLoop)%><%END IF%></td>		    				    	
	    	<td>		    				    		
	    		<A href="<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a>
	    		<% if cEvtItem.IsSoldOut(arrList(14,intLoop),arrList(16,intLoop),arrList(20,intLoop),arrList(21,intLoop)) then %>
	    			<br><img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
	    		<% end if %>
	    	</td>
	    	<td>
	    		<% if (Not IsNull(arrList(12,intLoop)) ) and (arrList(12,intLoop)<>"") then %>
					<img src="<%= imgFingers %>/diyItem/webimage/small/<%=GetImageSubFolderByItemid(arrList(0,intLoop))%>/<%=arrList(12,intLoop)%>">
				<%end if%>
	    	</td>    	
	    	<td><%=db2html(arrList(3,intLoop))%></td>
	    	<td align="left"><%=db2html(arrList(4,intLoop))%></td>
	    	<td>
	    		<%
				Response.Write FormatNumber(arrList(7,intLoop),0)
				'할인가
				if arrList(18,intLoop)="Y" then
					Response.Write "<br><font color=#F08050>(할)" & FormatNumber(arrList(9,intLoop),0) & "</font>"
				end if
				'쿠폰가
				if arrList(22,intLoop)="Y" then
					Select Case arrList(23,intLoop)
						Case "1"
							Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(7,intLoop)*((100-arrList(24,intLoop))/100),0) & "</font>"
						Case "2"
							Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(7,intLoop)-arrList(24,intLoop),0) & "</font>"
					end Select
				end if
				%>
			</td>
	    	<td>
	    		<%
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
			%>
			</td>
	    	<td><%= fnColor(cEvtItem.IsUpcheBeasong(arrList(15,intLoop)),"delivery")%></td>    	
	    	<td><%= fnColor(arrList(14,intLoop),"yn") %></td>
	    	<td><%= fnColor(arrList(19,intLoop),"yn") %></td>
	    	<td><%= fnColor(arrList(16,intLoop),"yn") %></td>    				    	
	    	<td><input type="text" name="sSort" value="<%=arrList(2,intLoop)%>" size="4" style="text-align:right;"></td>
	    	<td><% sbGetOptEventCodeValue "itemimgsize", arrList(27,intLoop), false, "" %></td>
	    	<td><input type="button" value="삭제" onClick="jsDel(1,<%=arrList(0,intLoop)%>);" class="button"></td>	    	
	    </tr>   
	   <%	Next
	   	ELSE
	   %>
	   	<tr  align="center" bgcolor="#FFFFFF">
	   		<td colspan="15">등록된 내용이 없습니다.</td>
	   	</tr>	
	   <%END IF%>
		</table>
		<!-- 페이징처리 -->
		<%		
		iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1	
		
		If (iCurrpage mod iPerCnt) = 0 Then																
			iEndPage = iCurrpage
		Else								
			iEndPage = iStartPage + (iPerCnt-1)
		End If	
		%>
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
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
		        <td width=50 align="right"><a href="event_list.asp?menupos=<%=menupos%>&<%=strparm%>"><img src="/images/icon_list.gif" border="0"></a></td>			        
		    </tr>			  
		</form>    
		</table>
	</td>
</tr>
</table>

<!-- 그룹이동--->
<form name="frmG" method="post" action="eventitem_process.asp">
	<input type="hidden" name="mode" value="G">
	<input type="hidden" name="iC" value="<%=iCurrpage%>">
	<input type="hidden" name="eC" value="<%=eCode%>">
	<input type="hidden" name="selG" value="<%=strG%>">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="selGroup" value="">
	<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<!-- 선택삭제--->
<form name="frmDel" method="post" action="eventitem_process.asp">
	<input type="hidden" name="mode" value="D">
	<input type="hidden" name="iC" value="<%=iCurrpage%>">
	<input type="hidden" name="eC" value="<%=eCode%>">
	<input type="hidden" name="selG" value="<%=strG%>">
	<input type="hidden" name="itemidarr" value="">
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
	<input type="hidden" name="menupos" value="<%=menupos%>">
</form>

<%
	set cEvtItem = nothing
%>	

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
