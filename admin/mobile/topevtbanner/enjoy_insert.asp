<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : enjoy_insert.asp
' Discription : 모바일 카테고리 TOP 2 EVENT
' History : 2015-09-16 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/topeventCls.asp" -->
<!-- #include virtual="/lib/classes/mobile/topcatecodeCls.asp" -->
<%
'###############################################
'이벤트 신규 등록시
'###############################################
%>
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
Dim eCode
Dim idx , evtimg , subImage2 , subImage3 , subImage4 , isusing , mode , gcode
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim evtalt, linkurl 
Dim evttitle
Dim issalecoupontxt
Dim prevDate , ordertext
Dim startdate
Dim enddate
Dim issalecoupon , linktype

Dim cEvtCont
Dim ename , stdt , eddt , sortnum , todaybanner , Molistbanner , evttitle2

	eCode = requestCheckvar(Request("eC"),10)
	gcode = requestCheckvar(Request("gcode"),3)
	idx = requestCheckvar(request("idx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")
	linktype = request("linktype") '이벤트링크타입

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

'// 입력시
IF eCode <> "" And mode = "add" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'이벤트 코드
	'이벤트 내용 가져오기
	cEvtCont.fnGetEventCont
	ename	=	db2html(cEvtCont.FEName)
	stdt	=	cEvtCont.FESDay
	eddt	=	cEvtCont.FEEDay
	'이벤트 화면설정 내용 가져오기
	cEvtCont.fnGetEventDisplay
	todaybanner = cEvtCont.FEBImgMoToday
	Molistbanner = cEvtCont.FEBImgMoListBanner
	
	set cEvtCont = nothing
END IF

'// 수정시
If idx <> "" then
	dim oTopevtOne
	set oTopevtOne = new CMainbanner
	oTopevtOne.FRectIdx = idx
	oTopevtOne.GetOneContents()

	linktype			=	oTopevtOne.FOneItem.Flinktype
	evtalt				=	oTopevtOne.FOneItem.Fevtalt
	linkurl				=	oTopevtOne.FOneItem.Flinkurl
	evtimg				=	oTopevtOne.FOneItem.Fevtimg
	evttitle			=	oTopevtOne.FOneItem.Fevttitle
	issalecoupontxt		=	oTopevtOne.FOneItem.Fissalecoupontxt
	startdate			=	oTopevtOne.FOneItem.Fevtstdate
	enddate				=	oTopevtOne.FOneItem.Fevteddate
	issalecoupon		=	oTopevtOne.FOneItem.Fissalecoupon
	mainStartDate		=	oTopevtOne.FOneItem.Fstartdate
	mainEndDate			=	oTopevtOne.FOneItem.Fenddate 
	isusing				=	oTopevtOne.FOneItem.Fisusing
	ordertext			=	oTopevtOne.FOneItem.Fordertext
	sortnum				=	oTopevtOne.FOneItem.Fsortnum
	todaybanner			=	oTopevtOne.FOneItem.Ftodaybanner
	eCode				=	oTopevtOne.FOneItem.Fevt_code
	Molistbanner		=	oTopevtOne.FOneItem.Fevtmolistbanner
	evttitle2			=	oTopevtOne.FOneItem.Fevttitle2
	gcode				=	oTopevtOne.FOneItem.Fgnbcode

	set oTopevtOne = Nothing
End If 

if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
	else
		sDt = Date()
	end if
	sTm = "00:00:00"
end if

if Not(mainEndDate="" or isNull(mainEndDate)) then
	eDt = left(mainEndDate,10)
	eTm = Num2Str(hour(mainEndDate),2,"0","R") &":"& Num2Str(minute(mainEndDate),2,"0","R") &":"& Num2Str(second(mainEndDate),2,"0","R")
else
	if srcEDT<>"" then
		eDt = left(srcEDT,10)
	else
		eDt = Date()
	end if
	eTm = "23:59:59"
end If

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;

		if (!frm.gcode.value)
		{
			alert('노출 GNB 영역을 선택 해주세요.');
			frm.gcode.focus();
			return;
		}

		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/topevtbanner/";
	}
	$(function(){
	//달력대화창 설정
	var arrDayMin = ["일","월","화","수","목","금","토"];
	var arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
    $("#sDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '이전달', nextText: '다음달', yearSuffix: '년',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 1,
    	showCurrentAtPos: 0,
      	showOn: "button",
      	<% if Idx<>"" then %>maxDate: "<%=eDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#eDt" ).datepicker( "option", "minDate", selectedDate );
    	}
    });
    $("#eDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '이전달', nextText: '다음달', yearSuffix: '년',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 2,
      	showOn: "button",
      	<% if Idx<>"" then %>minDate: "<%=sDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
    	}
    });

});

//-- jsPopCal : 달력 팝업 --//
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function putLinkText(key,gubun) {
	var frm = document.frm;
	var urllink
	if (gubun == "1" )
	{
		urllink = frm.linkurl;
	}
	switch(key) {
		case 'event':
			urllink.value='/event/eventmain.asp?eventid=이벤트번호';
			break;
		case 'itemid':
			urllink.value='/category/category_itemprd.asp?itemid=상품코드';
			break;
	}
}
//지난 이벤트 불러오기
function jsLastEvent(){
  var valsdt , valedt , valgcode
	valsdt = document.frm.sDt.value;
	valedt = document.frm.eDt.value;
	valgcode = document.frm.gcode.value;

	if (valgcode == ""){
		valgcode = "<%=gcode%>";
	}else{
		valgcode = document.frm.gcode.value;
	}

  var winLast,eKind;
  winLast = window.open('pop_event_list.asp?menupos=<%=menupos%>&gcode='+valgcode+'&sDt='+valsdt+'&eDt='+valedt,'pLast','width=550,height=600, scrollbars=yes')
  winLast.focus();
}
function chgmu(v){
	if (v == "1")
	{
		$("#sel11").css("display","block");
		$("#sel21").css("display","none");
		$("#sel22").css("display","none");
	}else{
		$("#sel11").css("display","none");
		$("#sel21").css("display","block");
		$("#sel22").css("display","block");
	}
}
</script>
<table width="1000" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/topeventbanner_proc.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">노출기간</td>
    <td colspan="3">
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<!-- 신규이벤트 등록시 -->
<% If mode = "add" Then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">노출 GNB영역</td>
	<td colspan="3"><% Call drawSelectBoxGNB("gcode" , gcode) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="30">이벤트 링크타입</td>
	<td colspan="3"><label for="load">이벤트 불러오기</label><input type="radio" value="1" name="linktype" id="load" onclick="chgmu('1');" <%=chkiif(linktype="1","checked","")%>/> <label for="self">직접입력</label><input type="radio" value="2" name="linktype" id="self" onclick="chgmu('2');"/></td>
</tr>
<tr bgcolor="#FFFFFF" id="sel11" style="display:<%=chkiif(linktype="1","block","none")%>;">
	<td bgcolor="#FFF999" align="center" height="30">이벤트불러오기</td>
	<td colspan="3"><input type="button" value="이벤트 불러오기" onclick="jsLastEvent();"/><br/><img src="<%=Molistbanner%>" width="200"><br/><%=Molistbanner%></td>
</tr>
<tr bgcolor="#FFFFFF" id="sel21" style="display:none;">
	<td bgcolor="#FFF999" align="center" height="30">이벤트 URL</td>
	<td colspan="3">
		<% IF eCode <> "" And mode = "add" THEN %>
			<input type="hidden" name="linkurl" value="/event/eventmain.asp?eventid=<%=eCode%>">
		<% Else %>
			<input type="text" name="linkurl" size="80" value="<%=linkurl%>"/>
		<% End If %>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('event','1')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','1')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="sel22" style="display:none;">
	<td bgcolor="#FFF999" align="center" width="15%">이벤트 이미지</td>
	<td width="45%">
		<input type="file" name="evtimg" class="file" title="이벤트 #1" require="N" style="width:80%;" />
		<% if evtimg<>"" then %>
		<br>
		<img src="<%= evtimg %>" width="100" /><br><%= evtimg %>
		<% end if %>		
	</td>
	<td bgcolor="#FFF999" align="center" width="10%">이벤트<br/>이미지 alt</td>
	<td width="20%"><input type="text" name="evtalt" value="<%=evtalt%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">이벤트 제목</td>
	<td width="45%"><input type="text" name="evttitle" value="<%=ename%>" size="40"/><!--</br><input type="text" name="evttitle2" value="" size="40"/>--></td>
	<td bgcolor="#FFF999" align="center" width="10%">이벤트 할인</td>
	<td width="20%">할인 : <input type="radio" name="issalecoupon" value="1" <%=chkiif(issalecoupon = 1,"checked","")%>/> 쿠폰 : <input type="radio" name="issalecoupon" value="2" <%=chkiif(issalecoupon = 2,"checked","")%>/> <input type="text" name="issalecoupontxt" size="10" value="<%=issalecoupontxt%>" maxlength="10"/><br/> GIFT : <input type="radio" name="issalecoupon" value="3" <%=chkiif(issalecoupon = 3,"checked","")%>/> 참여 : <input type="radio" name="issalecoupon" value="4" <%=chkiif(issalecoupon = 4,"checked","")%>/>&nbsp;</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">이벤트 시작일 - 종료일</td>
	<td colspan="3">
		<input type="text" name="evtstdate" size="10" value="<%=stdt%>" onClick="jsPopCal('evtstdate');"/>
		-
		<input type="text" name="evteddate" size="10" value="<%=eddt%>" onClick="jsPopCal('evteddate');"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">정렬 번호</td>
	<td colspan="3"><input type="text" name="sortnum" size="10" value="99" maxlength="3"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">작업자 지시사항</td>
	<td colspan="3"><textarea name="ordertext" cols="80" rows="8"/><%=ordertext%></textarea></td>
</tr>
<% Else %>
<!-- 이벤트 수정시 -->
<input type="hidden" value="<%=linktype%>" name="linktype"/>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">노출 GNB영역</td>
	<td colspan="3"><% Call drawSelectBoxGNB("gcode" , gcode) %></td>
</tr>
<% If linktype = "1" Then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">이벤트 이미지</td>
	<td colspan="3"><!-- 구버전<img src="<%=todaybanner%>" width="100"><br/><%=todaybanner%><br/>신버전 --><img src="<%=Molistbanner%>" width="200"><br/><%=Molistbanner%></td>
</tr>
<% Else %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">이벤트 이미지</td>
	<td width="45%">
		<input type="file" name="evtimg" class="file" title="이벤트 #1" require="N" style="width:80%;" />
		<% if evtimg<>"" then %>
		<br>
		<img src="<%= evtimg %>" width="100" /><br><%= evtimg %>
		<% end if %>		
	</td>
	<td bgcolor="#FFF999" align="center" width="10%">이벤트<br/>이미지 alt</td>
	<td width="20%"><input type="text" name="evtalt" value="<%=evtalt%>" size="20" maxlength="20"/></td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">이벤트 제목</td>
	<td width="45%"><input type="text" name="evttitle" value="<%=evttitle%>" size="40"/><!--</br><input type="text" name="evttitle2" value="<%=evttitle2%>" size="40"/>--></td>
	<td bgcolor="#FFF999" align="center" width="10%">이벤트 할인</td>
	<td width="20%">할인 : <input type="radio" name="issalecoupon" value="1" <%=chkiif(issalecoupon = 1,"checked","")%>/> 쿠폰 : <input type="radio" name="issalecoupon" value="2" <%=chkiif(issalecoupon = 2,"checked","")%>/> <input type="text" name="issalecoupontxt" size="10" value="<%=issalecoupontxt%>" maxlength="10"/><br/> GIFT : <input type="radio" name="issalecoupon" value="3" <%=chkiif(issalecoupon = 3,"checked","")%>/> 참여 : <input type="radio" name="issalecoupon" value="4" <%=chkiif(issalecoupon = 4,"checked","")%>/>&nbsp;</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">이벤트 시작일 - 종료일</td>
	<td colspan="3">
		<input type="text" name="evtstdate" size="10" value="<%=startdate%>" onClick="jsPopCal('evtstdate');"/>
		-
		<input type="text" name="evteddate" size="10" value="<%=enddate%>" onClick="jsPopCal('evteddate');"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">이벤트 URL</td>
	<td colspan="3"><input type="text" name="linkurl" size="80" value="<%=linkurl%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('event','1')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','1')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">정렬 번호</td>
	<td colspan="3"><input type="text" name="sortnum" size="10" value="<%=sortnum%>" maxlength="3"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">작업자 지시사항</td>
	<td colspan="3"><textarea name="ordertext" cols="80" rows="8"/><%=ordertext%></textarea></td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->