<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/site/index.asp
' Description :  이벤트 Static 이미지 관리
' History : 2007.03.27 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventSiteCls.asp"-->
<%
	Call fnSetEventCommonCode '공통코드 어플리케이션 변수에 세팅
	
	'변수선언
	Dim cEvtList
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt	
	Dim slocation, stype, limitCnt
	
	'파라미터값 받기 & 기본 변수 값 세팅
	slocation = Request("sitelocation")
	
	iCurrpage = Request("iC")	'현재 페이지 번호
	IF iCurrpage = "" THEN
		iCurrpage = 1	
	END IF	  
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	
	
	'데이터 가져오기
	set cEvtList = new ClsEvtSite	
		cEvtList.FCPage = iCurrpage	'현재페이지
		cEvtList.FPSize = iPageSize '한페이지에 보이는 레코드갯수 
		cEvtList.FSLocation = slocation '검색 : 위치
		
 		arrList = cEvtList.fnGetList	'데이터목록 가져오기
 		iTotCnt = cEvtList.FTotCnt	'전체 데이터  수
 	set cEvtList = nothing
 	
	iTotalPage 	=  Int(iTotCnt/iPageSize)	'전체 페이지 수
	IF (iTotCnt MOD iPageSize) > 0 THEN	iTotalPage = iTotalPage + 1		
	IF isArray(arrList) THEN stype = arrList(2,0) 	
%>
<script language="javascript">
<!--
	function jsGoPage(iP){
		document.frmEvt.iC.value = iP;
		document.frmEvt.action = "index.asp";
		document.frmEvt.submit();	
	}
	
	
	function AssignTest(slocation,stype){	
	 	var popwin = window.open('','refreshFrm_Test','');
		popwin.focus();
		 frmEvt.target = "refreshFrm_Test";
		 frmEvt.action = "<%=staticImgUrl%>/flash/link/make_event_test_JS.asp?sl=" + slocation+"&st="+stype;
		 frmEvt.submit();			 
	}
	
	function AssignReal(slocation,stype){	  
		 var popwin = window.open('','refreshFrm_Main','');
		 popwin.focus();
		 frmEvt.target = "refreshFrm_Main";
		 frmEvt.action = "<%=staticImgUrl%>/flash/link/make_event_JS.asp?sl=" + slocation+"&st="+stype;
		 frmEvt.submit();
	}
	
	function jsChangeFrm(){		
	 var sl;
	 sl= document.frmSearch.sitelocation.options[document.frmSearch.sitelocation.selectedIndex].value ;	 	
	 self.location.href = "index.asp?sitelocation=" + sl;
	 
	}
//-->
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frmSearch" method="post">
	<input type="hidden" name="iC">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td  >
        	위치 : <%sbGetOptEventCodeValue "sitelocation",slocation,True,"onChange='javascript:jsChangeFrm();'"%>
        	&nbsp;
        	<%IF slocation <> "" THEN %>
        	<!--<a href="javascript:AssignTest(<%=slocation%>,'<%=stype%>');"><img src="/images/icon_search.jpg" border="0" align="absmiddle">미리보기</a>
        	/ --><a href="javascript:AssignReal(<%=slocation%>,'<%=stype%>');"><img src="/images/refreshcpage.gif" align="absmiddle" border="0"> 리얼적용</a>        	
            <%END IF%>
        </td>
        <td  align="left" valign="bottom">        	
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>		
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td colspan="2">
        	+ 위치검색을 해서 적용가능한 이미지 확인 후에만 미리보기/리얼적용이 가능합니다.<br>
        	+ 노란부분이 적용가능한 이미지들입니다.        	
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->
<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<form name="frmEvt" method="post">
<input type="hidden" name="iC">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="sitelocation" value="<%=slocation%>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="35">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">       	
       	<a href="evtsite_regist.asp?menupos=<%=menupos%>"><img src="/images/icon_new_registration.gif" border="0"></a>
       	<% If stype = "flash" Then %>* 플래시 경우 가장 최근에 올린것이 1번입니다.<% End If %>
    	</td>
    	<td align="right">
       
       <!--	<input type="button" value="통계" onclick=" ">  -->
       <!--	정렬: <select name="selSort">
       	<option value="1">이벤트코드내림차순</option>
       	
       	</select>-->
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>idx</td>
    	<td>위치</td>
    	<td>종류</td>
    	<td>이미지</td>
    	<td>사이즈</td>
      	<td>link구분</td>
      	<td>순서</td>
      	<td>사용여부</td>      	
      	<td>등록일</td>      	
    </tr>    
    <%IF isArray(arrList) THEN 
    	For intLoop = 0 To UBound(arrList,2)
    	 IF arrList(2,intLoop) = "flash" THEN
    	   limitCnt = 3
    	 ELSE
    	   limitCnt = 1
    	END IF    
    %>
    <tr align="center"  <%IF slocation<> "" and Cint(intLoop) < Cint(limitCnt) THEN%>bgcolor="#FFFFF4"<%ELSE%>bgcolor="#FFFFFF"<%END IF%>>        	
    	<td><%=arrList(0,intLoop)%></td>
    	<td><a href="evtsite_regist.asp?menupos=<%=menupos%>&idx=<%=arrList(0,intLoop)%>"><%=fnGetEventCodeDesc("sitelocation",arrList(1,intLoop))%></a></td>
    	<td><%=arrList(2,intLoop)%></td>
    	<td><a href="evtsite_regist.asp?menupos=<%=menupos%>&idx=<%=arrList(0,intLoop)%>"><img src="<%=arrList(3,intLoop)%>" width="100" border="0"></a></td>
    	<td><%=arrList(6,intLoop)%> X <%=arrList(7,intLoop)%></td>
    	<td><%=arrList(4,intLoop)%></td>
    	<td><%=arrList(8,intLoop)%></td>
    	<td><%=arrList(10,intLoop)%></td>
    	<td><%=arrList(9,intLoop)%></td>    	
    </tr>   
   <%	Next
   	ELSE
   %>
   	<tr  align="center" bgcolor="#FFFFFF">
   		<td colspan="9">등록된 내용이 없습니다.</td>
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
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
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
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->