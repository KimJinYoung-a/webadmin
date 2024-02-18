<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : enjoy_insert.asp
' Discription : 모바일 enjoybanner_new
' History : 2014.06.23 이종화
' 		  : 2018.11.28 최종원 메인 상단 기획전 추가
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/todayenjoyCls.asp" -->
<%
'###############################################
'이벤트 신규 등록시
'###############################################
%>
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
Dim eCode
Dim idx , evtimg , subImage2 , subImage3 , subImage4 , isusing , mode
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
Dim ename , stdt , eddt , sortnum , todaybanner , Molistbanner , evttitle2 , etc_opt , subname , modify_Molistbanner
Dim tag_gift , tag_plusone , tag_launching , tag_actively , sale_per , coupon_per , tag_only
Dim itemid1 , itemid2 , itemid3 , addtype , iteminfo
Dim itemname1 ,  itemname2 , itemname3
Dim itemimg1 ,  itemimg2 , itemimg3


	eCode = requestCheckvar(Request("eC"),10)
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
IF eCode <> "" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'이벤트 코드
	'이벤트 내용 가져오기
	cEvtCont.fnGetEventCont
	ename	=	db2html(cEvtCont.FEName)
	subname	=	db2html(cEvtCont.FENamesub)
	stdt	=	left(cEvtCont.FESDay, 10)
	eddt	=	left(cEvtCont.FEEDay, 10)

	'이벤트 화면설정 내용 가져오기
	cEvtCont.fnGetEventDisplay
	todaybanner = cEvtCont.FEBImgMoToday
	If mode = "add" then
		Molistbanner = cEvtCont.FEBImgMoListBanner
	Else 
		modify_Molistbanner = cEvtCont.FEBImgMoListBanner
	End If 


	dim tmpename
		tmpename = Split(ename,"|") 
			 
	if Ubound(tmpename)>0 then
		ename = tmpename(0)
	end if
	
	set cEvtCont = nothing
END IF



'// 수정시
If idx <> "" then
	dim oEnjoyeventOne
	set oEnjoyeventOne = new CMainbanner
	oEnjoyeventOne.FRectIdx = idx
	oEnjoyeventOne.GetOneContents()

	linktype			=	oEnjoyeventOne.FOneItem.Flinktype
	evtalt				=	oEnjoyeventOne.FOneItem.Fevtalt
	linkurl				=	oEnjoyeventOne.FOneItem.Flinkurl
	evtimg				=	oEnjoyeventOne.FOneItem.Fevtimg
	evttitle			=	oEnjoyeventOne.FOneItem.Fevttitle
	issalecoupontxt		=	oEnjoyeventOne.FOneItem.Fissalecoupontxt
	startdate			=	left(oEnjoyeventOne.FOneItem.Fevtstdate, 10)
	enddate				=	left(oEnjoyeventOne.FOneItem.Fevteddate, 10)
	issalecoupon		=	oEnjoyeventOne.FOneItem.Fissalecoupon
	mainStartDate		=	oEnjoyeventOne.FOneItem.Fstartdate
	mainEndDate			=	oEnjoyeventOne.FOneItem.Fenddate 
	isusing				=	oEnjoyeventOne.FOneItem.Fisusing
	ordertext			=	oEnjoyeventOne.FOneItem.Fordertext
	sortnum				=	oEnjoyeventOne.FOneItem.Fsortnum
	todaybanner			=	oEnjoyeventOne.FOneItem.Ftodaybanner
	eCode				=	oEnjoyeventOne.FOneItem.Fevt_code
	Molistbanner		=	oEnjoyeventOne.FOneItem.Fevtmolistbanner
	evttitle2			=	oEnjoyeventOne.FOneItem.Fevttitle2
	etc_opt				=	oEnjoyeventOne.FOneItem.Fetc_opt

	tag_only			=	oEnjoyeventOne.FOneItem.Ftag_only
	tag_gift			=	oEnjoyeventOne.FOneItem.Ftag_gift
	tag_plusone			=	oEnjoyeventOne.FOneItem.Ftag_plusone
	tag_launching		=	oEnjoyeventOne.FOneItem.Ftag_launching
	tag_actively		=	oEnjoyeventOne.FOneItem.Ftag_actively
	sale_per			=	oEnjoyeventOne.FOneItem.Fsale_per
	coupon_per			=	oEnjoyeventOne.FOneItem.Fcoupon_per

	itemid1				=	oEnjoyeventOne.FOneItem.Fitemid1
	itemid2				=	oEnjoyeventOne.FOneItem.Fitemid2
	itemid3				=	oEnjoyeventOne.FOneItem.Fitemid3
	addtype				=	oEnjoyeventOne.FOneItem.Faddtype
	iteminfo			=	oEnjoyeventOne.FOneItem.Fiteminfo

	set oEnjoyeventOne = Nothing

	Dim ii
	If addtype = 2 then
		If ubound(Split(iteminfo,"^^")) > 0 Then ' 이미지 3개 정보
			For ii = 0 To ubound(Split(iteminfo,","))
				If CStr(itemid1) = CStr(Split(Split(iteminfo,",")(ii),"|")(0)) Then
					itemname1 = Split(Split(iteminfo,",")(ii),"|")(1)
					itemimg1 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(itemid1) & "/" & Split(Split(iteminfo,",")(ii),"|")(2)
				End If

				If CStr(itemid2) = CStr(Split(Split(iteminfo,",")(ii),"|")(0)) Then
					itemname2 = Split(Split(iteminfo,",")(ii),"|")(1)
					itemimg2 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(itemid2) & "/" & Split(Split(iteminfo,",")(ii),"|")(2)
				End If

				If CStr(itemid3) = CStr(Split(Split(iteminfo,",")(ii),"|")(0)) Then
					itemname3 = Split(Split(iteminfo,",")(ii),"|")(1)
					itemimg3 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(itemid3) & "/" & Split(Split(iteminfo,",")(ii),"|")(2)
				End If
			Next 
		End If 
	End If 
End If 

dim dateOption
dateOption = request("dateoption")

if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
		if prevDate = "" then 
			prevDate = sDt
		end if 
	elseif dateOption <> "" then
		sDt = dateOption
	else
		sDt = date
	end if
		sTm = "00:00:00"
end if

if Not(mainEndDate="" or isNull(mainEndDate)) then
	eDt = left(mainEndDate,10)
	eTm = Num2Str(hour(mainEndDate),2,"0","R") &":"& Num2Str(minute(mainEndDate),2,"0","R") &":"& Num2Str(second(mainEndDate),2,"0","R")
else
	if srcEDT<>"" then
		eDt = left(srcEDT,10)
		if prevDate = "" then 
			prevDate = eDt
		end if
	elseif dateOption <> "" then
		eDt = dateOption
	else	
		eDt = date
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

		if (frm.linkurl.value.indexOf("이벤트번호") > 0 || frm.linkurl.value.indexOf("상품코드") > 0){
			alert("링크 값을 확인 해주세요.");
			frm.linkurl.focus();
			return;
		}

		if (frm.addtype[1].checked && frm.addtype[1].value == 2){
			if (frm.itemid1.value == ""){
				alert("상품코드1를 넣어주세요.");
				frm.itemid1.focus();
				return;
			}
			if (frm.itemid2.value == ""){
				alert("상품코드2를 넣어주세요.");
				frm.itemid2.focus();
				return;
			}
			if (frm.itemid3.value == ""){
				alert("상품코드3를 넣어주세요.");
				frm.itemid3.focus();
				return;
			}
		}

		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/appmanage/today/enjoy/";
//		self.location.href="/admin/appmanage/today/enjoy/?menupos=1633&tabs=1";
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
    	numberOfMonths: 2,
    	showCurrentAtPos: 1,
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
	valsdt = document.frm.StartDate.value;
	valedt = document.frm.EndDate.value;

  var winLast,eKind;
  winLast = window.open('pop_event_list.asp?menupos=<%=menupos%>&sDt='+valsdt+'&eDt='+valedt,'pLast','width=550,height=600, scrollbars=yes')
  winLast.focus();
}

function chgmu(v){
	if (v == "1")
	{
		$("#sel11").css("display","");
		$("#sel21").css("display","none");
		$("#sel22").css("display","none");
	}else{
		$("#sel11").css("display","none");
		$("#sel21").css("display","");
		$("#sel22").css("display","");
	}
}
function changeForm(){
	var dispOption = document.frm.addtype.value;	
	var link = dispOption == 1 ? "enjoy_insert.asp" : "mainTopExhibition_insert.asp"
	document.location.href= link + "?menupos=<%=menupos%>&prevDate=<%=sDt%>&dispOption="+dispOption;	
}
function chgtype(v){
	if (v == "1"){
		$("#additem1").css("display","none");
		$("#additem2").css("display","none");
		$("#additem3").css("display","none");
		$("#evttitle2").attr("maxlength",60);
	}else if(v == "3"){
		changeForm();
	}else{
		$("#additem1").css("display","");
		$("#additem2").css("display","");
		$("#additem3").css("display","");
		$("#evttitle2").attr("maxlength",30);
	}
}

// 상품정보 접수
function fnGetItemInfo(iid,v) {
	$.ajax({
		type: "GET",
		url: "/admin/sitemaster/wcms/act_iteminfo.asp?itemid="+iid,
		dataType: "xml",
		cache: false,
		async: false,
		timeout: 5000,
		beforeSend: function(x) {
			if(x && x.overrideMimeType) {
				x.overrideMimeType("text/xml;charset=euc-kr");
			}
		},
		success: function(xml) {
			if($(xml).find("itemInfo").find("item").length>0) {
				var rst = "<img src='" + $(xml).find("itemInfo").find("item").find("smallImage").text() + "' height='70' /><br/>"
					rst += $(xml).find("itemInfo").find("item").find("itemname").text();
				$("#lyItemInfo"+v).fadeIn();
				$("#lyItemInfo"+v).html(rst);
			} else {
				$("#lyItemInfo"+v).fadeOut();
			}
		},
		error: function(xhr, status, error) {
			$("#lyItemInfo"+v).fadeOut();
			/*alert(xhr + '\n' + status + '\n' + error);*/
		}
	});
}
</script>
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/todayenjoy_proc.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<%'2017 상품 추가 ver %>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#FFF999" align="center">기획전 타입</td>
  <td colspan="3">
	<input type="radio" name="addtype" id="typeA" value="1" onclick="chgtype('1');" checked/> <label for="typeA">기본형</label>
	<!--<input type="radio" name="addtype" id="typeB" value="2" onclick="chgtype('2');" disabled/> <label for="typeB">기본형 + 상품3개</label>&nbsp;<br/>-->
	<input type="radio" name="addtype" id="typeC" value="3" onclick="chgtype('3');" /> <label for="typeB">메인상단기획전</label>&nbsp;<br/>	
  </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="10%">노출기간</td>
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
	<td bgcolor="#FFF999" align="center" height="30">이벤트 링크타입</td>
	<td colspan="3">
		<label for="load">이벤트 불러오기</label>
		<input type="radio" value="1" name="linktype" id="load" onclick="chgmu('1');" <%=chkiif(linktype="1","checked","")%>/>
		<label for="self">직접입력</label>
		<input type="radio" value="2" name="linktype" id="self" onclick="chgmu('2');"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="sel11" style="display:<%=chkiif(linktype="1","","none")%>;">
	<td bgcolor="#FFF999" align="center" height="30">이벤트불러오기</td>
	<td colspan="3"><input type="button" value="이벤트 불러오기" onclick="jsLastEvent();"/><img src="<%=Molistbanner%>" width="200"><br/><%=Molistbanner%></td>
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
	<td bgcolor="#FFF999" align="center" width="10%">이벤트 이미지</td>
	<td>
		<input type="file" name="evtimg" class="file" title="이벤트 #1" require="N" style="width:50%;" />
		<% if evtimg<>"" then %>
		<br>
		<img src="<%= evtimg %>" width="100" /><br><%= evtimg %>
		<% end if %>		
	</td>
	<td bgcolor="#FFF999" align="center" width="10%">이벤트<br/>이미지 alt</td>
	<td width="40%"><input type="text" name="evtalt" value="<%=evtalt%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">이벤트 제목</td>
	<td colspan="3"><input type="text" name="evttitle" value="<%=ename%>" size="40"/></br><input type="text" name="evttitle2" id="evttitle2" value="<%=subname%>" size="70"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">이벤트 시작일 - 종료일</td>
	<td colspan="3">
		<input type="text" name="evtstdate" size="10" value="<%=stdt%>" onClick="jsPopCal('evtstdate');"/>
		-
		<input type="text" name="evteddate" size="10" value="<%=eddt%>" onClick="jsPopCal('evteddate');"/>
	</td>
</tr>

<tr bgcolor="#FFFFFF" style="display:none;" id="additem1">
    <td bgcolor="#FFF999" align="center">상품코드1</td>
    <td colspan="3">
        <input type="text" name="itemid1" value="" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'1')" title="상품코드" />
        <div id="lyItemInfo1" style="display:none;"></div>
    </td>
</tr>
<tr bgcolor="#FFFFFF" style="display:none;" id="additem2">
    <td bgcolor="#FFF999" align="center">상품코드2</td>
    <td colspan="3">
        <input type="text" name="itemid2" value="" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'2')" title="상품코드" />
        <div id="lyItemInfo2" style="display:none;"></div>
    </td>
</tr>
<tr bgcolor="#FFFFFF" style="display:none;" id="additem3">
    <td bgcolor="#FFF999" align="center">상품코드3</td>
    <td colspan="3">
        <input type="text" name="itemid3" value="" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'3')" title="상품코드" />
        <div id="lyItemInfo3" style="display:none;"></div>
    </td>
</tr>
<%'2017 상품 추가 ver %>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#FFF999" align="center">태그</td>
  <td>
  	<input type="checkbox" name="tag_only" id="tag_only" value="Y"/> <label for="tag_only">단독</label>
	<input type="checkbox" name="tag_gift" id="tag_gift" value="Y"/> <label for="tag_gift">GIFT</label>
	<input type="checkbox" name="tag_plusone" id="tag_plusone" value="Y"/> <label for="tag_plusone">1+1</label>&nbsp;
	<input type="checkbox" name="tag_launching" id="tag_launching" value="Y"/> <label for="tag_launching">런칭</label>&nbsp;
	<input type="checkbox" name="tag_actively" id="tag_actively" value="Y"/> <label for="tag_actively">참여(코멘트, 게시판 , 상품후기)</label>&nbsp;<br/>
	<font color="red"><strong>※ 단독 > GIFT > 1+1 > 런칭 > 참여 순으로 노출 됩니다.※</strong></font>
  </td>
  <td bgcolor="#FFF999" align="center">할인/쿠폰</td>
  <td>
	<input type="text" name="sale_per" value=""> : 할인율 ex)<font color="red"><strong>~45%</strong></font></br>
	<input type="text" name="coupon_per" value=""> : 쿠폰할인율 ex)<font color="green"><strong>10%</strong></font></br>
	<font color="red"><strong>※있는 경우만 입력 하세요.※</strong></font>
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
<% If linktype = "1" Then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">이벤트 이미지</td>
	<td colspan="3"><!-- 구버전<img src="<%=todaybanner%>" width="100"><br/><%=todaybanner%><br/>신버전 --><img src="<%=chkiif(Molistbanner="",modify_Molistbanner,Molistbanner)%>" width="200"><br/><%=chkiif(Molistbanner="",modify_Molistbanner,Molistbanner)%>
	<% If Molistbanner= "" And modify_Molistbanner <> "" then%>
	<br/>※ 해당 이벤트의 이미지가 등록 되었습니다 저장을 하시면 메인페이지에 적용이 됩니다. ※ 
	<% End If %>
	</td>
</tr>
<% Else %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">이벤트 이미지</td>
	<td>
		<input type="file" name="evtimg" class="file" title="이벤트 #1" require="N" style="width:80%;" />
		<% if evtimg<>"" then %>
		<br>
		<img src="<%= evtimg %>" width="100" /><br><%= evtimg %>
		<% end if %>		
	</td>
	<td bgcolor="#FFF999" align="center" width="10%">이벤트<br/>이미지 alt</td>
	<td width="40%"><input type="text" name="evtalt" value="<%=evtalt%>" size="20" maxlength="20"/></td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">이벤트 제목</td>
	<td colspan="3"><input type="text" name="evttitle" value="<%=evttitle%>" size="40"/></br><input type="text" name="evttitle2" id="evttitle2" value="<%=evttitle2%>" size="70"/></td>
	
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">이벤트 시작일 - 종료일</td>
	<td colspan="3">
		<input type="text" name="evtstdate" size="10" value="<%=startdate%>" onClick="jsPopCal('evtstdate');"/>
		-
		<input type="text" name="evteddate" size="10" value="<%=enddate%>" onClick="jsPopCal('evteddate');"/>
	</td>
</tr>

<tr bgcolor="#FFFFFF" style="display:<%=chkiif(addtype="2","","none")%>;" id="additem1">
    <td bgcolor="#FFF999" align="center">상품코드1</td>
    <td colspan="3">
        <input type="text" name="itemid1" value="<%=itemid1%>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'1')" title="상품코드" />
        <div id="lyItemInfo1" style="display:<%=chkIIF(itemid1="","none","")%>;">
		<%
        	if Not(itemName1="" or isNull(itemName1)) then
        		Response.Write "<img src='" & itemimg1 & "' height='70' /><br/>"
        		Response.Write itemName1
        	end if
        %>
		</div>
    </td>
</tr>
<tr bgcolor="#FFFFFF" style="display:<%=chkiif(addtype="2","","none")%>;" id="additem2">
    <td bgcolor="#FFF999" align="center">상품코드2</td>
    <td colspan="3">
        <input type="text" name="itemid2" value="<%=itemid2%>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'2')" title="상품코드" />
        <div id="lyItemInfo2" style="display:<%=chkIIF(itemid2="","none","")%>;">
		<%
        	if Not(itemName2="" or isNull(itemName2)) then
        		Response.Write "<img src='" & itemimg2 & "' height='70' /><br/>"
        		Response.Write itemName2
        	end if
        %>
		</div>
    </td>
</tr>
<tr bgcolor="#FFFFFF" style="display:<%=chkiif(addtype="2","","none")%>;" id="additem3">
    <td bgcolor="#FFF999" align="center">상품코드3</td>
    <td colspan="3">
        <input type="text" name="itemid3" value="<%=itemid3%>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'3')" title="상품코드" />
        <div id="lyItemInfo3" style="display:<%=chkIIF(itemid3="","none","")%>;">
		<%
        	if Not(itemName3="" or isNull(itemName3)) then
        		Response.Write "<img src='" & itemimg3 & "' height='70' /><br/>"
        		Response.Write itemName3
        	end if
        %>
		</div>
    </td>
</tr>
<%'2017 상품 추가 ver %>

<tr bgcolor="#FFFFFF">
  <td bgcolor="#FFF999" align="center">태그</td>
  <td>
  	<input type="checkbox" name="tag_only" id="tag_only" value="Y" <%=chkiif(tag_only = "Y","checked","")%>/> <label for="tag_only">단독</label><br/>
	<input type="checkbox" name="tag_gift" id="tag_gift" value="Y" <%=chkiif(tag_gift = "Y","checked","")%>/> <label for="tag_gift">GIFT</label>
	<input type="checkbox" name="tag_plusone" id="tag_plusone" value="Y" <%=chkiif(tag_plusone = "Y","checked","")%>/> <label for="tag_plusone">1+1</label>&nbsp;
	<input type="checkbox" name="tag_launching" id="tag_launching" value="Y" <%=chkiif(tag_launching = "Y","checked","")%>/> <label for="tag_launching">런칭</label>&nbsp;
	<input type="checkbox" name="tag_actively" id="tag_actively" value="Y" <%=chkiif(tag_actively = "Y","checked","")%>/> <label for="tag_actively">참여(코멘트, 게시판 , 상품후기)</label>&nbsp;<br/>
	<font color="red"><strong>※ GIFT > 1+1 > 런칭 > 참여 순으로 노출 됩니다.※</strong></font>
  </td>
  <td bgcolor="#FFF999" align="center">할인/쿠폰</td>
  <td>
	<input type="text" name="sale_per" value="<%=sale_per%>"> : 할인율 ex)<font color="red"><strong>~45%</strong></font></br>
	<input type="text" name="coupon_per" value="<%=coupon_per%>"> : 쿠폰할인율 ex)<font color="green"><strong>10%</strong></font></br>
	<font color="red"><strong>※있는 경우만 입력 하세요.※</strong></font>
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