<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : mainTopExhibition_insert.asp
' Discription : 모바일 상단기획전
' histroy	  : 2018.11.28 최종원 메인 상단 기획전 추가
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/todayenjoyCls.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<%
'###############################################
'이벤트 신규 등록시
'###############################################
%>
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
dim encUsrId, tmpTx, tmpRn, userid, indexparam
Dim eCode, PeCode
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
dim dispOption
dim contentType
dim itemId

Dim cEvtCont
Dim ename , stdt , eddt , sortnum , todaybanner , Molistbanner , evttitle2 , etc_opt , subname , modify_Molistbanner
Dim tag_gift , tag_plusone , tag_launching , tag_actively , sale_per , coupon_per , tag_only
Dim itemid1 , itemid2 , itemid3 , addtype , iteminfo
Dim itemname1 ,  itemname2 , itemname3
Dim itemimg1 ,  itemimg2 , itemimg3
dim evtSqureImg

	contentType = request("contentType")
	dispOption = request("dispOption")
	PeCode = requestCheckvar(Request("eC"),10)
	idx = requestCheckvar(request("idx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")
	linktype = request("linktype") '이벤트링크타입
	userid = session("ssBctId")
	indexparam = request("indexparam")
Randomize()
tmpTx = split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z",",")
tmpRn = tmpTx(int(Rnd*26))
tmpRn = tmpRn & tmpTx(int(Rnd*26))
encUsrId = tenEnc(tmpRn & userid)	


If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

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
	startdate			=	oEnjoyeventOne.FOneItem.Fevtstdate
	enddate				=	oEnjoyeventOne.FOneItem.Fevteddate
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
	contentType 		=   oEnjoyeventOne.FOneItem.FcontentType

	set oEnjoyeventOne = Nothing

	if evtimg = "" then
		set cEvtCont = new ClsEvent
		cEvtCont.FECode = eCode	'이벤트 코드
		'이벤트 내용 가져오기
		cEvtCont.fnGetEventDisplay

		evtimg = cEvtCont.FEtcitemimg
		set cEvtCont = nothing	
	end if
End If 

'// 입력시
IF PeCode <> "" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = PeCode	'이벤트 코드
	'이벤트 내용 가져오기
	cEvtCont.fnGetEventCont
	ename	=	db2html(cEvtCont.FEName)
	subname	=	db2html(cEvtCont.FENamesub)
	stdt	=	cEvtCont.FESDay
	eddt	=	cEvtCont.FEEDay	
	'이벤트 화면설정 내용 가져오기
	cEvtCont.fnGetEventDisplay
	todaybanner = cEvtCont.FEBImgMoToday
	evtimg = cEvtCont.FEtcitemimg
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
	
	eCode = PeCode

	set cEvtCont = nothing
END IF

	dim dateOption
	dateOption = request("dateoption")

if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
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
	elseif dateOption <> "" then
		eDt = dateOption
	else	
		eDt = date
	end if
	eTm = "23:59:59"
end If
if indexparam = 1 then
	stdt = startdate
	eddt = enddate
	ename = evttitle
	subname = evttitle2
end if
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
$(function(){	
    $('#startdate').datepicker( {
        changeMonth: true,
        changeYear: true,
        showButtonPanel: true,
        dateFormat: 'yy-mm-dd',     
	});		
    $('#enddate').datepicker( {
        changeMonth: true,
        changeYear: true,
        showButtonPanel: true,
        dateFormat: 'yy-mm-dd',     
    });			
});
	function jsSubmit(){
		var frm = document.frm;		

		if(frm.contentType.value == 1 && frm.evt_code.value == ""){
			alert("이벤트를 넣어주세요.");
			return false;
		}else if(frm.contentType.value == 2 && frm.itemid1.value == ""){
			alert("상품을 넣어주세요.");
			return false;
		}else if(frm.evttitle.value == "" ){
			alert("카피를 넣어주세요.");
			frm.evttitle.focus()
			return false;			
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
	valsdt = document.frm.sDt.value;
	valedt = document.frm.eDt.value;

  var winLast,eKind;
  winLast = window.open('pop_event_list.asp?menupos=<%=menupos%>&sDt='+valsdt+'&eDt='+valedt+'&type=3&idx=<%=idx%>','pLast','width=550,height=600, scrollbars=yes')
  winLast.focus();
}
function changeView(vContentType){
	if(vContentType == 2){
		$("#prdBtn").css("display","")
		$("#evtBtn").css("display","none")  
		$("#prdCode").css("display","")  		
		$("#evtdate").css("display","none")
		$("#evtCode").css("display","none")		
	}else{
		$("#prdBtn").css("display","none")
		$("#evtBtn").css("display","")  
		$("#prdCode").css("display","none")  		
		$("#evtdate").css("display","")
		$("#evtCode").css("display","")		
	}
}
function addnewItem(){
	var popwin; 		
	popwin = window.open("item_regist.asp", "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
	popwin.focus();
}
function jsCheckUpload() {
	var gubun = document.frmUpload.imgtype.value;
	var mainfrm = document.frm
	var test = $("input[id="+gubun+"]").val();
	// console.log(gubun);	
	// console.log(test);
	// return false;
	if($("#fileupload").val()!="") {
		$("#fileupmode").val("upload");

		$('#ajaxform').ajaxSubmit({
			//보내기전 validation check가 필요할경우
			beforeSubmit: function (data, frm, opt) {
				if(!(/\.(jpg|jpeg|png)$/i).test(frm[0].upfile.value)) {
					alert("JPG,PNG 이미지파일만 업로드 하실 수 있습니다.");
					$("#fileupload").val("");
					return false;
				}
				$("#lyrPrgs").show();
			},
			//submit이후의 처리
			success: function(responseText, statusText){
				var resultObj = JSON.parse(responseText)

				if(resultObj.response=="fail") {
					alert(resultObj.faildesc);
				} else if(resultObj.response=="ok") {					
					$("#filepre").val(resultObj.fileurl);
					$("img[id="+gubun+"src]").hide().attr("src",$("#filepre").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
					$("input[id="+gubun+"]").val(resultObj.fileurl);															
				} else {
					alert("처리중 오류가 발생했습니다.\n" + responseText);
				}
				$("#fileupload").val("");
				$("#lyrPrgs").hide();
			},
			//ajax error
			error: function(err){
				alert("ERR: " + err.responseText);
				$("#fileupload").val("");
				$("#lyrPrgs").hide();
			}
		});
	}
}
function setImgType(type){	
	document.frmUpload.imgtype.value = type;
	return false;
}
</script>
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmUpload" id="ajaxform" action="<%=uploadImgUrl%>/linkweb/common/simpleCommonImgUploadProc.asp" method="post" enctype="multipart/form-data" style="display:none; height:0px;width:0px;">
<input style="display:none" type="file" name="upfile" id="fileupload" onchange="jsCheckUpload();" accept="image/*" />
<input type="hidden" name="mode" id="fileupmode" value="upload">
<input type="hidden" name="div" value="TQ">
<input type="hidden" name="upPath" value="/appmanage/roundbanner/">
<input type="hidden" name="tuid" value="<%=encUsrId%>">
<input type="hidden" name="prefile" id="filepre" >	
<input type="hidden" name="imgtype">
</form>	
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/todayenjoy_proc.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="addtype" value="4" />
<%'2017 상품 추가 ver %>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#FFF999" align="center">콘텐츠 타입</td>
  <td colspan="3">  
	<input type="radio" name="contentType" <%=chkiif(contentType = "1","checked","")%> checked value="1" onclick="changeView('1');" /> 이벤트
	<input type="radio" name="contentType" <%=chkiif(contentType = "2","checked","")%> value="2" onclick="changeView('2');" /> 상품&nbsp;<br/>	
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
<tr bgcolor="#FFFFFF" >
	<td bgcolor="#FFF999" align="center" height="30">이벤트/상품 불러오기</td>
	<td colspan="3">			
		<input type="button" id="evtBtn" style="display:<%=chkIIF(contentType=1 or contentType="", "","none")%>" value="이벤트 불러오기" onclick="jsLastEvent();"/>
		<input type="button" id="prdBtn" style="display:<%=chkIIF(contentType=2, "", "none")%>" value="상품 불러오기" onclick="addnewItem();"/>		
	</td>	
</tr>
<tr bgcolor="#FFFFFF" id="prdCode" style="display:<%=chkIIF(contentType=2, "", "none")%>">
	<td bgcolor="#FFF999" align="center" height="30">상품코드</td>
	<td colspan="3">	
		<input type="text" style="width:100px" readonly name="itemid1" value="<%=itemid1%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="evtCode" style="display:<%=chkIIF(contentType=1 or contentType="", "", "none")%>">
	<td bgcolor="#FFF999" align="center" height="30">이벤트코드</td>
	<td colspan="3">	
		<input type="text" style="width:100px" readonly name="evt_code" value="<%=eCode%>">		
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="30">이벤트 / 상품이미지</td>
	<td >
		<div class="inTbSet">												
			<div>	
				<p class="registImg">
					<input type="hidden" id="evtimg" name="evtimg" value="<%=evtimg%>" />
					<img id="evtimgsrc" src="<%=chkIIF(evtimg="" or isNull(evtimg),"/images/admin_login_logo2.png",evtimg)%>" style="height:138px; border:1px solid #EEE;"/>																
				</p>
				<button type="button">																		
					<div onclick="setImgType('evtimg')" >					
						<label for="fileupload" style="cursor:pointer;">
							<%=chkIIF(evtimg="","이미지 업로드","이미지 수정")%>
						</label>					
					</div>							
				</button>										
			</div>	
		</div>							
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">카피</td>
	<td colspan="3">
		<input type="text" name="evttitle" value="<%=ename%>" size="40"/></br>				
	</td>
</tr>
<%'2017 상품 추가 ver %>
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
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->