<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : item_insert.asp
' Discription : 메인 상단 기획전 상품추가 페이지 
' History : 2018.11.26 최종원
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/pcmain_multieventCls.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<%
'###############################################
'이벤트 신규 등록시
'###############################################
%>
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
dim encUsrId, tmpTx, tmpRn, userid
Dim eCode
Dim idx , evtimg , isusing , mode
Dim srcSDT , srcEDT
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim linkurl
Dim maincopy
Dim issalecoupontxt
Dim prevDate , ordertext
Dim startdate
Dim enddate
dim tag_only
dim dispOption
dim contentType
dim contentImg
dim itemId

Dim cEvtCont
Dim ename , stdt , eddt , sortnum , todaybanner , Molistbanner , subcopy , etc_opt , subname
Dim sale_per , coupon_per
userid = session("ssBctId")

	contentType = request("contentType")
	dispOption = request("dispOption")
	eCode = requestCheckvar(Request("eC"),10)
	idx = requestCheckvar(request("idx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")

If idx = "" Then
	mode = "add"
Else
	mode = "modify"
End If

if contentType = "" then
	contentType = 1
end if

Randomize()
tmpTx = split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z",",")
tmpRn = tmpTx(int(Rnd*26))
tmpRn = tmpRn & tmpTx(int(Rnd*26))
encUsrId = tenEnc(tmpRn & userid)	


'// 입력시
IF eCode <> "" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'이벤트 코드
	'이벤트 내용 가져오기
	cEvtCont.fnGetEventCont
	ename	=	db2html(cEvtCont.FEName)
	subname	=	db2html(cEvtCont.FENamesub)
	stdt	=	cEvtCont.FESDay
	eddt	=	cEvtCont.FEEDay
	'이벤트 화면설정 내용 가져오기
	cEvtCont.fnGetEventDisplay
	todaybanner = cEvtCont.FEBImgMoToday
	Molistbanner = cEvtCont.FEBImgMoListBanner

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

	linkurl				=	oEnjoyeventOne.FOneItem.Flinkurl
	maincopy			=	oEnjoyeventOne.FOneItem.Fmaincopy
	startdate			=	oEnjoyeventOne.FOneItem.Fevtstdate
	enddate				=	oEnjoyeventOne.FOneItem.Fevteddate
	mainStartDate		=	oEnjoyeventOne.FOneItem.Fstartdate
	mainEndDate			=	oEnjoyeventOne.FOneItem.Fenddate
	isusing				=	oEnjoyeventOne.FOneItem.Fisusing
	ordertext			=	oEnjoyeventOne.FOneItem.Fordertext
	sortnum				=	oEnjoyeventOne.FOneItem.Fsortnum
	eCode				=	oEnjoyeventOne.FOneItem.Feventid
	Molistbanner		=	oEnjoyeventOne.FOneItem.Fevtmolistbanner
	subcopy				=	oEnjoyeventOne.FOneItem.Fsubcopy
	sale_per			=	oEnjoyeventOne.FOneItem.Fsale_per
	coupon_per			=	oEnjoyeventOne.FOneItem.Fcoupon_per
	tag_only			=	oEnjoyeventOne.FOneItem.Ftag_only
	dispOption			=   oEnjoyeventOne.FOneItem.FdispOption
	contentImg			=   oEnjoyeventOne.FOneItem.FcontentImg
	contentType			=   oEnjoyeventOne.FOneItem.FcontentType
	itemId				=	oEnjoyeventOne.FOneItem.FitemId

	set oEnjoyeventOne = Nothing
End If

if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
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
	else
		eDt = date
	end if
	eTm = "23:59:59"
end If
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;

		if (frm.linkurl.value.indexOf("이벤트번호") > 0 || frm.linkurl.value.indexOf("상품코드") > 0){
			alert("링크 값을 확인 해주세요.");
			frm.linkurl.focus();
			return;
		}

		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.action = "event_proc.asp";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/pcmain/multievent/";
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
			urllink.value='/shopping/category_prd.asp?itemid=상품코드';
			break;
	}
}
//지난 이벤트 불러오기
function jsLastEvent(){
  var valsdt , valedt , valgcode, vDispOption
	valsdt = document.frm.sDt.value;
	valedt = document.frm.eDt.value;
	vDispOption = document.frm.dispOption.value;

  var winLast,eKind;
  winLast = window.open('pop_event_list.asp?menupos=<%=menupos%>&sDt='+valsdt+'&eDt='+valedt+'&dispOption='+vDispOption,'pLast','width=550,height=600, scrollbars=yes')
  winLast.focus();
}
function changeForm(){
	var contentType = document.frm.contentType.value;	
	var dispOption = document.frm.dispOption.value;	
	var link = contentType == 1 ? "event_insert.asp" : "item_insert.asp"
	document.location.href= link + "?menupos=<%=menupos%>&prevDate=<%=sDt%>&dispOption="+dispOption+"&contentType="+contentType;	
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
<form name="frmUpload" id="ajaxform" action="<%=uploadImgUrl%>/linkweb/common/simpleCommonImgUploadProc.asp" method="post" enctype="multipart/form-data" style="display:none; height:0px;width:0px;">
<input style="display:none" type="file" name="upfile" id="fileupload" onchange="jsCheckUpload();" accept="image/*" />
<input type="hidden" name="mode" id="fileupmode" value="upload">
<input type="hidden" name="div" value="TQ">
<input type="hidden" name="upPath" value="/appmanage/multi3img/">
<input type="hidden" name="tuid" value="<%=encUsrId%>">
<input type="hidden" name="prefile" id="filepre" >	
<input type="hidden" name="imgtype">
</form>		
<form name="frm" method="post" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="eventid" value="<%=eCode%>">
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="10%">노출 위치</td>
    <td >
		<select name="dispOption" class="select">							
			<option value="2" <%=chkiif(dispOption="2"," selected","")%>>메인상단기획전</option>
		</select>					
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="10%">콘텐츠구분</td>
	<td>
		<div style="float:left;">
			<input type="radio" onclick="changeForm()" name="contentType" value="1" <%=chkiif(contentType = "1","checked","")%> checked />이벤트 &nbsp;&nbsp;&nbsp; 
			<input type="radio" onclick="changeForm()" name="contentType" value="2" <%=chkiif(contentType = "2","checked","")%>/>상품
		</div>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="10%">노출기간</td>
    <td >
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="30">상품불러오기</td>
	<td >
		<% If Molistbanner <> "" Then %>
		<img src="<%=Molistbanner%>" width="200"><br/><%=Molistbanner%>
		<% End If %>
		<input type="button" value="상품 불러오기" onclick="addnewItem();"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="30">상품이미지</td>
	<td >
		<div class="inTbSet">												
			<div>	
				<p class="registImg">
					<input type="hidden" id="contentImg" name="contentImg" value="<%=contentImg%>" />
					<img id="contentImgsrc" src="<%=chkIIF(contentImg="" or isNull(contentImg),"/images/admin_login_logo2.png",contentImg)%>" style="height:138px; border:1px solid #EEE;"/>																
				</p>
				<button type="button">																		
					<div onclick="setImgType('contentImg')" >					
						<label for="fileupload" style="cursor:pointer;">
							<%=chkIIF(contentImg="","이미지 업로드","이미지 수정")%>
						</label>					
					</div>							
				</button>										
			</div>	
		</div>			
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="30">상품코드</td>
	<td>
		<input type="text" name="itemId" value="<%=itemId%>" readonly>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="30">상품 URL</td>
	<td>		
		<input type="text" name="linkurl" size="80" value="<%=linkurl%>"/>		
	<br/><br/>ex)
		<font color="#707070">		
		- <span style="cursor:pointer" onClick="putLinkText('itemid','1')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">카피</td>
	<td >
		1: <input type="text" name="maincopy" value="<%=chkiif(mode="add",ename,maincopy)%>" size="40" maxlength="65"/></br>
		2: <input type="text" name="subcopy" id="subcopy" value="<%=chkiif(mode="add",subname,subcopy)%>" size="70" maxlength="170"/><font color="red"></font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#FFF999" align="center">할인/쿠폰</td>
  <td>
	<input type="text" name="sale_per" value="<%=sale_per%>"> : 할인율 ex)<font color="red"><strong>~45%</strong></font></br>
	<input type="text" name="coupon_per" value="<%=coupon_per%>"> : 쿠폰할인율 ex)<font color="green"><strong>10%</strong></font></br>
	<font color="red"><strong>※있는 경우만 입력 하세요.※</strong></font>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td ><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">정렬 번호</td>
	<td ><input type="text" name="sortnum" size="10" value="<%=chkiif(mode="add","99",sortnum)%>" maxlength="3"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">작업자 지시사항</td>
	<td ><textarea name="ordertext" cols="80" rows="8"/><%=ordertext%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/></td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->