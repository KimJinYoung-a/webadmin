<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/pop_event_lastlist.asp
' Description :  지난 이벤트 내용 가져오기
' History : 2007.03.20 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
	'변수선언
	Dim menupos
Dim cEvtList, cDisp, vCateCode, i, eventstate, startdate, enddate
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt, vOpenerForm
	Dim sEvt,strTxt,sKind
	
	startdate = Request("startdate")
	enddate = Request("enddate")
	eventstate = Request("eventstate")
	vCateCode = Request("catecode")
	menupos = request("menupos")
	sKind 	= request("eventkind")
	sEvt 	= request("selEvt")
	strTxt 	= request("sEtxt")
	vOpenerForm = request("openerform")
	
	'파라미터값 받기 & 기본 변수 값 세팅
	iCurrpage = Request("iC")	'현재 페이지 번호
	IF iCurrpage = "" THEN
		iCurrpage = 1	
	END IF
	if sEvt = "" then sEvt = "evt_code"
	if sKind = "" then sKind = "1"
	
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	
	'데이터 가져오기
	set cEvtList = new ClsEvent	
		cEvtList.FCPage = iCurrpage	'현재페이지
		cEvtList.FPSize = iPageSize '한페이지에 보이는 레코드갯수 
		
		cEvtList.FSKind = sKind '검색 종류
		cEvtList.FSfEvt = sEvt '검색 이벤트명 or 이벤트코드
		cEvtList.FSeTxt = strTxt '검색어
		cEvtList.FRectState = eventstate '상태값
		cEvtList.FRectSDate = startdate '시작일
		cEvtList.FRectEDate = enddate '종료일
		cEvtList.FRectDisp = vCateCode
		
 		arrList = cEvtList.fnGetEventLastList	'데이터목록 가져오기
 		iTotCnt = cEvtList.FTotCnt	'전체 데이터  수
 	set cEvtList = nothing
 	 	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
	Dim arreventkind, arreventstate
	'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
	arreventkind = fnSetCommonCodeArr("eventkind",False)
	arreventstate= fnSetCommonCodeArr("eventstate",False)	
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
<!--
	window.onload = function(){
		window.resizeTo(1000, 820);
	}

	//검색
	function jsSearch(){
		if(document.frmLast.sEtxt.value != ""){
			if(document.frmLast.selEvt.options[document.frmLast.selEvt.selectedIndex].value == "evt_code") {
				if(!IsDigit(document.frmLast.sEtxt.value)){
					alert("이벤트 코드는 숫자만 입력가능합니다.");
					document.frmLast.sEtxt.focus();
					return false;
				}
			}
		}
	}
	
	//부모창에 값 넘기기
	function jsSetEvtCont(ieC){
	  if(typeof(opener.document) == "object"){		 
	     <% if (request("pTarget")<>"") then %>
	     opener.location.href = "<%= request("pTarget") %>&eC="+ieC;
	     <% else %>
	     	<% If vOpenerForm <> "" Then %>
	     		opener.<%=vOpenerForm%>.value = ieC;
	    	<% Else %>
		 		opener.location.href = "event_regist.asp?menupos=<%=menupos%>&eC="+ieC;
		 	<% End If %>
		 <% end if %>
		 window.close();
	  }	
	}
	
	//페이지이동
	function jsGoPage(iP){
		document.frmLast.iC.value = iP;
		document.frmLast.submit();
	}
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 지난 이벤트 리스트 </div>
<form name="frmLast" method="get" action="pop_event_lastlist.asp" onSubmit="return jsSearch();" style="margin:0px;">
<input type="hidden" name="iC">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="pTarget" value="<%=request("pTarget")%>">
<input type="hidden" name="openerform" value="<%=vOpenerForm%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="5%" height="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td height="30" align="left">전시카테고리 :
		<%
		SET cDisp = New cDispCate
		cDisp.FCurrPage = 1
		cDisp.FPageSize = 2000
		cDisp.FRectDepth = 1
		cDisp.FRectUseYN = "Y"
		cDisp.GetDispCateList()
		
		If cDisp.FResultCount > 0 Then
			Response.Write "<select name=""catecode"" class=""select"" onChange=""document.frmLast.submit();"">" & vbCrLf
			Response.Write "<option value="""">선택</option>" & vbCrLf
			For i=0 To cDisp.FResultCount-1
				Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(CStr(vCateCode)=CStr(cDisp.FItemList(i).FCateCode),"selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
			Next
			Response.Write "</select>&nbsp;&nbsp;&nbsp;"
		End If
		Set cDisp = Nothing
		%>
	</td>
	<td height="30" rowspan="3"><input type="submit" value=" 검  색 " class="button" style="width:70px;height:50px;" onfocus="this.blur();"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td height="30">
		이벤트종류 : <%sbGetOptEventCodeValue "eventkind",sKind,True,"onChange=""document.frmLast.submit();"""%>&nbsp;&nbsp;&nbsp;
		코드/명 : <select name="selEvt">
        			<option value="evt_code" <%if Cstr(sEvt) = "evt_code" THEN %>selected<%END IF%>>이벤트코드</option>
        			<option value="evt_name" <%if Cstr(sEvt) = "evt_name" THEN %>selected<%END IF%>>이벤트명</option>
        			</select>
        			<input type="text" name="sEtxt" size="15" value="<%=strTxt%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td height="30">
		진행상태 : <% sbGetOptStatusCodeValue "eventstate",eventstate,true,"class=""select"" onChange=""document.frmLast.submit();""" %>&nbsp;&nbsp;&nbsp;
		기간 :&nbsp;
        <input id="startdate" type="text" name="startdate" value="<%= startdate %>" maxlength="10" size="10">
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
        <input type="text" name="dummy0" value="00:00:00" size="8" readonly class="text_ro" />
	    <script type="text/javascript">
		var CAL_Start = new Calendar({
			inputField : "startdate",
			trigger    : "startdate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_End.args.min = date;
				CAL_End.redraw();
				this.hide();
			},
			bottomBar: true,
			dateFormat: "%Y-%m-%d"
		});
		</script>
		&nbsp;~&nbsp;
        <input id="enddate" type="text" name="enddate" value="<%= enddate %>" maxlength="10" size="10">
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absbottom" />
        <input type="text" name="dummy1" value="23:59:59" size="8" readonly class="text_ro" />
	    <script type="text/javascript">
		var CAL_End = new Calendar({
			inputField : "enddate",
			trigger    : "enddate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_Start.args.max = date;
				CAL_Start.redraw();
				this.hide();
			},
			bottomBar: true,
			dateFormat: "%Y-%m-%d"
		});
		</script>
	</td>
</tr>
</table>
<br />
<table width="100%" border="0" align="left" class="a" cellpadding="0" cellspacing="0" >
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>" >
			<td align="center" width="15%">종류</td>
			<td align="center" width="7%">코드</td>
			<td align="center" width="15%">진행상태</td>
			<td align="center">이벤트명</td>
			<td align="center" width="15%">카테고리</td>
			<td align="center" width="10%">시작일</td>
			<td align="center" width="10%">종료일</td>
		</tr>
		<%IF isArray(arrList) THEN 
			For intLoop = 0 To UBound(arrList,2)
			%>
		<tr bgcolor="#FFFFFF" onClick="jsSetEvtCont(<%=arrList(0,intLoop)%>);" style="cursor:hand;" onMouseOver="this.style.backgroundColor='#DDDDDD'" onMouseOut="this.style.backgroundColor='#FFFFFF'">
			<td height="25" align="center"><%=fnGetCommCodeArrDesc(arreventkind,arrList(1,intLoop))%></td>
			<td  align="center"><%=arrList(0,intLoop)%></td>
			<td  align="center"><%=fnGetCommCodeArrDesc(arreventstate,arrList(8,intLoop))%></td>
			<td><%=db2html(arrList(4,intLoop))%></td>
			<td  align="center"><%=arrList(9,intLoop)%></td>
			<td  align="center"><%=arrList(5,intLoop)%></td>
			<td  align="center"><%=arrList(6,intLoop)%></td>
		</tr>
		<% Next %>
		<%ELSE%>
		<tr><td colspan="10"  bgcolor="#FFFFFF" align="center">등록된 내용이 없습니다.</td></tr>
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
</table>
</form>

<!-- #include virtual="/lib/db/dbclose.asp" -->