<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : pop_event_select.asp
' Description :  이벤트 선택
' History : 2019.02.27 정태훈 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<%
	'변수선언
	Dim menupos
	Dim cEvtList	
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt	
	Dim sEvt,strTxt,sKind, mode, eC
	dim blnWeb, blnMobile, blnApp, isWeb, isMobile, isApp
	
	menupos = request("menupos")
	mode = request("mode")
	sKind 	= request("eventkind")
	sEvt 	= request.form("selEvt")
	strTxt 	= request.form("sEtxt")
	eC 	= request("eC")
	blnWeb		= requestCheckVar(Request("isWeb"),1)
	blnMobile	= requestCheckVar(Request("isMobile"),1)
	blnApp		= requestCheckVar(Request("isApp"),1)
	
	'파라미터값 받기 & 기본 변수 값 세팅
	iCurrpage = Request("iC")	'현재 페이지 번호
	IF iCurrpage = "" THEN
		iCurrpage = 1	
	END IF	  
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
		
	'데이터 가져오기
	set cEvtList = new ClsEvent	
		cEvtList.FCPage = iCurrpage	'현재페이지
		cEvtList.FPSize = iPageSize '한페이지에 보이는 레코드갯수 
		
		cEvtList.FIsWeb = blnWeb
		cEvtList.FIsMobile = blnMobile
		cEvtList.FIsApp = blnApp
		cEvtList.FSKind = sKind '검색 종류
		cEvtList.FSfEvt = sEvt '검색 이벤트명 or 이벤트코드
		cEvtList.FSeTxt = strTxt '검색어	
		
 		arrList = cEvtList.fnGetEventLastList	'데이터목록 가져오기
 		iTotCnt = cEvtList.FTotCnt	'전체 데이터  수
 	set cEvtList = nothing
 	 	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
	Dim arreventkind, arreventstate
	'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
	arreventkind = fnSetCommonCodeArr("eventkind",False)
	arreventstate= fnSetCommonCodeArr("eventstate",False)	
%>
<script language="javascript">
<!--
	window.document.domain = "10x10.co.kr";
	//검색
	function jsSearch(){	 
	 if(document.frmLast.selEvt.options[document.frmLast.selEvt.selectedIndex].value == "evt_code" &&  document.frmLast.sEtxt.value !="") {
	   if(!IsDigit(document.frmLast.sEtxt.value)){
	    alert("이벤트 코드는 숫자만 입력가능합니다.");
	    document.frmLast.sEtxt.focus();
	    return false;
	   }
	 }		
	}
	window.document.domain = "10x10.co.kr";
	//부모창에 값 넘기기
	function jsSetEvtRelation(ieC){
		opener.document.frmEvt.evt_code.value = ieC;
		self.close();
	}

	function jsSetEvtCont(ieC){
		document.ibfrm.evt_code.value = ieC;
		document.ibfrm.submit();
	}
	
	//페이지이동
	function jsGoPage(iP){
		document.frmLast.iC.value = iP;
		document.frmLast.submit();
	}
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 지난 이벤트 리스트 </div>
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="0" >
<form name="frmLast" method="post" action="pop_event_select.asp" onSubmit="return jsSearch();">
<input type="hidden" name="iC">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="pTarget" value="<%=request("pTarget")%>">
<input type="hidden" name="eC" value="<%=eC%>">
<tr>
	<td>
	     채널:<input type="checkbox" name="isWeb" value="1" <%if blnWeb="1" then%>checked<%end if%>>PC-Web
			<input type="checkbox" name="isMobile"  value="1" <%if blnMobile="1" then%>checked<%end if%>>Mobile
			<input type="checkbox" name="isApp"  value="1" <%if blnApp="1" then%>checked<%end if%>>App&nbsp;&nbsp;&nbsp;
		종류: <%sbGetOptEventCodeValue "eventkind",sKind,True,"onChange=""document.frmLast.submit();"""%>&nbsp;&nbsp;&nbsp;
		코드/명: <select name="selEvt"> 
        			<option value="evt_code" <%if Cstr(sEvt) = "evt_code" THEN %>selected<%END IF%>>이벤트코드</option>
        			<option value="evt_name" <%if Cstr(sEvt) = "evt_name" THEN %>selected<%END IF%>>이벤트명</option>
        			</select>
        			<input type="text" name="sEtxt" size="15" value="<%=strTxt%>">
        			 <input type="image" src="/images/icon_search.jpg">
    </td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>" >
		     <td align="center" width="15%">채널</td>
			<td align="center" width="10%">코드</td>	
			<td align="center" width="15%">종류</td>
			<td align="center">이벤트명</td>	
			<td align="center" width="15%">상태</td>	
		</tr>
		<%IF isArray(arrList) THEN 
			For intLoop = 0 To UBound(arrList,2)
			isWeb = False
		    isMobile = False
		    isApp = False
		
		IF isNull(arrList(9,intLoop)) and isNull(arrList(10,intLoop)) and isNull(arrList(11,intLoop)) then
			if arrList(1,intLoop) = "19" THEN
				isWeb = False
				isMobile = True
				isApp = True
			ELSEIF arrList(1,intLoop) = "25"  THEN
				isWeb = False
				isMobile = False
				isApp = True
			ELSEIF arrList(1,intLoop) = "26"  THEN	
				isWeb = False
				isMobile = True
				isApp = False
			ELSE
				isWeb = True
				isMobile = False
				isApp = False	
			END IF
		END IF	
		IF 	 not isNull(arrList(9,intLoop))  THEN	
			isWeb = arrList(9,intLoop)
		END IF	
		IF 	 not isNull(arrList(10,intLoop)) THEN
			 isMobile = arrList(10,intLoop)
		END IF	 
		IF 	 not isNull(arrList(11,intLoop)) THEN
			isApp = arrList(11,intLoop)	
		END IF	
			%>
		<% if mode="relation" then %>
		<tr bgcolor="#FFFFFF" onClick="jsSetEvtRelation(<%=arrList(0,intLoop)%>);" style="cursor:hand;" onMouseOver="this.style.backgroundColor='#FFFFEC'" onMouseOut="this.style.backgroundColor='#FFFFFF'">
		<% else %>
		<tr bgcolor="#FFFFFF" onClick="jsSetEvtCont(<%=arrList(0,intLoop)%>);" style="cursor:hand;" onMouseOver="this.style.backgroundColor='#FFFFEC'" onMouseOut="this.style.backgroundColor='#FFFFFF'">
		<% end if %>
		    <td> <%IF isWeb THEN %>  Web <%END IF%><%IF isMobile THEN %>&nbsp; <font color="blue">Mobile</font> <%END IF%><%IF isApp THEN %>&nbsp;<font color="red">App</font><%END IF%></td>
			<td  align="center"><%=arrList(0,intLoop)%></td>
			<td  align="center"><%=fnGetCommCodeArrDesc(arreventkind,arrList(1,intLoop))%></td>
			<td><%=db2html(arrList(4,intLoop))%></td>
			<td  align="center"><%=fnGetCommCodeArrDesc(arreventstate,arrList(8,intLoop))%></td>
		</tr>
		<% Next %>
		<%ELSE%>
		<tr><td colspan="4"  bgcolor="#FFFFFF" align="center">등록된 내용이 없습니다.</td></tr>
		<%END IF%>
		</table>	
</tr>
<tr>
	<td>
		<!-- 페이징처리 -->
<%		
iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1	

If (iCurrpage mod iPerCnt) = 0 Then																
	iEndPage = iCurrpage
Else								
	iEndPage = iStartPage + (iPerCnt-1)
End If	
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
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
</table>
	</td>
</tr>
</form>
</table>
<form method="post" name="ibfrm" action="copyevent_process.asp">
    <input type="hidden" name="evt_code">
	<input type="hidden" name="mode" value="<%=mode%>">
	<input type="hidden" name="eC" value="<%=eC%>">
</form>
<!-- #include virtual="/lib/db/dbclose.asp" -->