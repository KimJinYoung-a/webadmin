<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/classes/sitemasterclass/TenQuizCls.asp" -->
<%
Dim userid, encUsrId, tmpTx, tmpRn, i
userid = session("ssBctId")

dim idx
dim chasu
dim TopTitle
dim QuizDescription
dim BackGroundImage
dim MWbackgroundImage
dim PCWBackGroundImage
dim QuestionHintNumber
dim TotalMileage
dim QuizStartDate
dim QuizEndDate
dim TotalQuestionCount
dim StartDescription
dim AdminRegister
dim AdminName
dim AdminModifyer
dim AdminModifyerName
dim RegistDate
dim LastUpDate
dim QuizStatus
dim productEvtNum
dim endAlertTxt
dim waitingAlertTxt
dim QuizStatusValue()
redim preserve QuizStatusValue(3)
QuizStatusValue(1) = "등록 대기"
QuizStatusValue(2) = "오픈"
QuizStatusValue(3) = "종료"
dim mode

dim quizStartTime
dim quizEndTime

'문항 정보
dim tenQuizQuestions
dim tempQuestionType

idx = requestCheckvar(request("idx"),16) 

'테스트데이터

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

If idx <> "" then
	dim tenQuizItem
	set tenQuizItem = new TenQuiz
	tenQuizItem.FRectIdx = idx
	tenQuizItem.GetOneContent()

	idx					= tenQuizItem.FoneItem.Fidx											
	chasu				= tenQuizItem.FoneItem.Fchasu								
	TopTitle			= tenQuizItem.FoneItem.FTopTitle				
	QuizDescription		= tenQuizItem.FoneItem.FQuizDescription					
	BackGroundImage		= tenQuizItem.FoneItem.FBackGroundImage					
	MWbackgroundImage	= tenQuizItem.FoneItem.FMWbackgroundImage
	PCWBackGroundImage	= tenQuizItem.FoneItem.FPCWBackGroundImage
	QuestionHintNumber	= tenQuizItem.FoneItem.FQuestionHintNumber					
	TotalMileage		= tenQuizItem.FoneItem.FTotalMileage								
	QuizStartDate		= tenQuizItem.FoneItem.FQuizStartDate					
	QuizEndDate			= tenQuizItem.FoneItem.FQuizEndDate						
	TotalQuestionCount	= tenQuizItem.FoneItem.FTotalQuestionCount			
	StartDescription	= tenQuizItem.FoneItem.FStartDescription
	productEvtNum		= tenQuizItem.FoneItem.FProductEvtNum
	AdminRegister		= tenQuizItem.FoneItem.FAdminRegister	
	AdminName			= tenQuizItem.FoneItem.FAdminName
	AdminModifyer		= tenQuizItem.FoneItem.FAdminModifyer		
	AdminModifyerName	= tenQuizItem.FoneItem.FAdminModifyerName	
	RegistDate			= tenQuizItem.FoneItem.FRegistDate	
	LastUpDate			= tenQuizItem.FoneItem.FLastUpDate	
	QuizStatus			= tenQuizItem.FoneItem.FQuizStatus	
	endAlertTxt			= tenQuizItem.FoneItem.FEndAlertTxt
	waitingAlertTxt		= tenQuizItem.FoneItem.FWaitingAlertTxt

	set tenQuizItem = Nothing

	set tenQuizQuestions = new TenQuiz	
	tenQuizQuestions.FRectChasu		= chasu
	tenQuizQuestions.FPageSize		= TotalQuestionCount
	tenQuizQuestions.FCurrPage		= 1	
	tenQuizQuestions.GetContentsItemList()
else
TotalMileage = 1000000
End If 

Randomize()
tmpTx = split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z",",")
tmpRn = tmpTx(int(Rnd*26))
tmpRn = tmpRn & tmpTx(int(Rnd*26))
	encUsrId = tenEnc(tmpRn & userid)	
%>
<style type="text/css">
html {overflow:auto;}
body {background-color:#fff;}  
.ui-state-highlight { height: 2.5em; line-height: 2.5em;}
.ui-datepicker{z-index: 99 !important};
</style>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />    
<link rel="stylesheet" href="/js/jquery-ui-timepicker-0.3.3/jquery.ui.timepicker.css?v=0.3.4" type="text/css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" href="/js/jquery-ui-timepicker-0.3.3/include/ui-1.10.0/ui-lightness/jquery-ui-1.10.0.custom.min.css" type="text/css" />
<script type="text/javascript" src="/js/jquery-ui-timepicker-0.3.3/include/jquery-1.9.0.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jquery.swiper-3.3.1.min.js"></script>
<script type="text/javascript" src="/js/tag-it.min.js"></script>
<script type="text/javascript" src="/js/jquery.form.min.js"></script>     
<script type="text/javascript" src="/js/jquery-ui-timepicker-0.3.3/jquery.ui.timepicker.js?v=0.3.3"></script>
    <script type="text/javascript" src="https://apis.google.com/js/plusone.js"></script>
<script type="text/javascript">
$(function(){
	$('#startTime').timepicker();		
	$('#endTime').timepicker();		

    $( "#sortable" ).sortable({
		placeholder: "ui-state-highlight",
		beforeStop: function(){
			reOrderQuizList();
		}
	});
    $( "#sortable" ).disableSelection();	
	var mainfrm = document.frm;		
	initiateValues();
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
    	onClose: function( selectedDate ) {
			$( "#eDt" ).datepicker( "option", "minDate", selectedDate );		
			mainfrm.chasu.value = selectedDate.replace(/-/gi,'');	
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
    	onClose: function( selectedDate ) {
    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
    	}
    });
	
function reOrderQuizList(){
	<% if mode = "modify" then 
	dim ii
	%>
		var originalSeq2 = []
		var newSeq = []
		var questionNumArr = []
		<% for ii=0 to tenQuizQuestions.FtotalCount-1 %>
		questionNumArr.push(parseInt(<%=tenQuizQuestions.FItemList(ii).FIquestionNumber%>))
		<% next %>
		questionNumArr.sort(function(a, b){
			return a - b;
		});		
		$("input[name='newSeq']").each(function(idx, item){
			item.value=questionNumArr[idx];
		});					
		// $("p[name='VquizNumber']").each(function(idx, item){
		// 	item.text(questionNumArr[idx]);
		// });		

		$.each($("input[name='originalSeq2']"), function(k, v){
			originalSeq2.push($(v).val());
		})
		$.each($("input[name='newSeq']"), function(k, v){
			newSeq.push($(v).val());
		})		
		console.log(questionNumArr);
		checkOrder(originalSeq2, newSeq);			
	<% end if %>	
}
function checkOrder(arr1, arr2){
	var result = true;
	var obj = $("#seqAlert");
	var isChangedObj = $("#OrderChangedFlag")

	result = arr1.every((v,i)=> v === arr2[i]);
	// console.log("result : ",result);
	// console.log(arr1);
	// console.log(arr2);	
	if(result){ 
		obj.css("display","none"); 
		isChangedObj.val("");
	}
	else{
		obj.css("display",""); 
		isChangedObj.val(1);
	}
	
}
function initiateValues(){	 
		$("li a").click(function(e){
			e.stopPropagation();
		});
		$("#backgroundImage").css("backgroundColor", $("#backgroundImage option:selected").css("backgroundColor"));
		$("#MWbackgroundImage").css("backgroundColor", $("#MWbackgroundImage option:selected").css("backgroundColor"));
		$("#PCWBackGroundImage").css("backgroundColor", $("#PCWBackGroundImage option:selected").css("backgroundColor"));
		<% 
		if QuizStartDate <> "" then 
		quizStartTime = Num2Str(hour(QuizStartDate),2,"0","R") &":"& Num2Str(minute(QuizStartDate),2,"0","R") 
		%>
		mainfrm.quizStartDate.value = '<%=left(QuizStartDate,10)%>'
		mainfrm.quizStartTime.value = '<%=quizStartTime%>'
		<% else %>
		mainfrm.quizStartTime.value = '10:00'
		<% end if %>	

		<% 
		if QuizEndDate <> "" then 
		quizendTime = Num2Str(hour(QuizEndDate),2,"0","R") &":"& Num2Str(minute(QuizEndDate),2,"0","R") 
		%>
		mainfrm.quizEndDate.value = '<%=left(QuizEndDate,10)%>'
		mainfrm.quizEndTime.value = '<%=quizendTime%>'
		<% else %>
		mainfrm.quizEndTime.value = '22:00'
		<% end if %>	
	}
});	
function jsCheckUpload() {
	var gubun = document.frmUpload.imgtype.value;
	var mainfrm = document.frm
	console.log(gubun);	
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
					if(gubun === "title"){
						$("#lyrBnrImg").hide().attr("src",$("#filepre").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
						$("#topTitle").val(resultObj.fileurl);
					}else if(gubun === "backgroundImage"){
						$("#lyrBnrImg2").hide().attr("src",$("#filepre").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
						$("#backgroundImage").val(resultObj.fileurl);
					}else if(gubun === "MWbackgroundImage"){
						$("#lyrBnrImg3").hide().attr("src",$("#filepre").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
						$("#MWbackgroundImage").val(resultObj.fileurl);
					}else if(gubun === "PCWBackGroundImage"){
						$("#lyrBnrImg4").hide().attr("src",$("#filepre").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
						$("#PCWBackGroundImage").val(resultObj.fileurl);
					}					
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
// 물리적인 파일 삭제 처리
function jsgolist(){
	self.location.href="/admin/sitemaster/tenquiz/";
	}
function addContent(){
	var mainfrm = document.frm;		
		if(mainfrm.backgroundImage.value==""){
			alert('배경 이미지를 입력해 주세요.');
			return false;
		}
		if(mainfrm.quizDescription.value==""){			
			alert('퀴즈 설명을 입력해주세요.');
			mainfrm.quizDescription.focus();
			return false;
		}
		if(mainfrm.questionHintNumber.value==""){
			alert('힌트 문항 번호를 입력해 주세요.');
			mainfrm.questionHintNumber.focus();
			return false;
		}
		if(mainfrm.totalMileage.value==""){
			alert('총 마일리지를 입력해 주세요.');
			mainfrm.totalMileage.focus();
			return false;
		}
		if(mainfrm.quizStartDate.value==""){
			alert('텐퀴즈 시작 날짜를 입력해 주세요.');
			mainfrm.quizStartDate.focus();
			return false;
		}
		if(mainfrm.quizEndDate.value==""){
			alert('텐퀴즈 종료 날짜를 입력해 주세요.');
			mainfrm.quizEndDate.focus();
			return false;
		}
		if(mainfrm.totalQuestionCount.value==""){			
			alert('총 문제 수를 입력해 주세요.');
			mainfrm.totalQuestionCount.focus();
			return false;
		}
		// if(mainfrm.startDescription.value==""){			
		// 	alert('하단 설명을 입력해 주세요.');
		// 	mainfrm.startDescription.focus();
		// 	return false;
		// }
		
		mainfrm.action="tenquizaction.asp";
		mainfrm.submit();		
	}
function jsmodify(v){	
	var popwin = window.open("/admin/sitemaster/tenquiz/popQuestionEdit.asp?listidx=<%=idx%>&subidx="+v, "popup_item", "width=800,height=800,scrollbars=yes,resizable=yes");
	popwin.focus();
}	
function jsdelete(v, qn){
	
	if(confirm(`${qn}번 문제를 삭제하시겠습니까?`)){
		var frm = document.delFrm;
		frm.idx.value = v;
		frm.action = "tenquizaction.asp"
		frm.submit();
	}
}	
function popQuestionEdit(){
	<% if idx <> "" then %>
		<% if tenQuizQuestions.FtotalCount >= TotalQuestionCount then %>
			alert('총 문항 수보다 초과 등록할 수 없습니다.');
		<% else %>
			var popwin = window.open("/admin/sitemaster/tenquiz/popQuestionEdit.asp?listidx=<%=idx%>", "popup_item", "width=800,height=800,scrollbars=yes,resizable=yes");
			popwin.focus();
		<% end if %>
	<% else %>
		alert("퀴즈를 먼저 등록해주세요.");
	<% end if %>	
}	
function setImgType(type){
	document.frmUpload.imgtype.value = type;
	return false;
}
function setImagePath(flatform){
	if(flatform == "a"){		
		$("#lyrBnrImg2").hide().attr("src",$("#backgroundImage").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
		$("#backgroundImage").css("backgroundColor", $("#backgroundImage option:selected").css("backgroundColor"));		
	}else if(flatform == "mw"){
		$("#lyrBnrImg3").hide().attr("src",$("#MWbackgroundImage").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
		$("#MWbackgroundImage").css("backgroundColor", $("#MWbackgroundImage option:selected").css("backgroundColor"));
	}else{
		$("#lyrBnrImg4").hide().attr("src",$("#PCWBackGroundImage").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
		$("#PCWBackGroundImage").css("backgroundColor", $("#PCWBackGroundImage option:selected").css("backgroundColor"));
	}
}
function numOfQuizChk(currentQuizCnt ,quizTotalCnt, stateValue){
	console.log(currentQuizCnt+"/"+quizTotalCnt);
	console.log(stateValue.value);
	if(stateValue.value == 2){
		console.log(currentQuizCnt+"/"+quizTotalCnt);
		if(currentQuizCnt < quizTotalCnt){			
			alert("총 문항 수까지 등록하셔야 오픈하실 수 있습니다. "+currentQuizCnt+"/"+quizTotalCnt);
			stateValue.value = 1
			return false;
		}		
	}
}	
// 업로드 파일 확인 및 처리
</script>
<div class="popWinV17">
	<h1>컨텐츠 등록/수정</h1>
	<button type="button" class="btn btn2" style="position:absolute; right:15px; top:7px;">도움말</button>
	<div class="popContainerV17 pad30">
		<p class="cGn1">* 내용 작성 후 반드시 '저장' 버튼을 눌러주세요.</p>
		<% if mode = "modify" then%>
		<h2 class="tMar20 subType">텐퀴즈 수정</h2>
		<% else %>
		<h2 class="tMar20 subType">텐퀴즈 등록</h2>
		<% end if %>		
		<%if mode <> "add" then%>
		<p class="tPad10 fs11" style="border-top:1px dashed #c9c9c9">
			<span class="cGy1"><%=AdminName&"  "%> <%=RegistDate%> 등록</span><br /><span class="cOr1"><%=AdminModifyerName&"  "%> <%=LastUpDate%> 최종수정</span>
		</p>				
		<% end if %>		
		<form name="frm">
			<table class="tbType1 writeTb tMar10">
				<colgroup>
					<col width="18%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>상태<strong class="cRd1">*</strong></div></th>
					<td>
						<select name="quizStatus" 
						<% if mode = "modify" then %>
						onchange="numOfQuizChk(<%=tenQuizQuestions.FtotalCount%>, <%=TotalQuestionCount%>, this)"
						<% end if %>
						>
						<% for i=1 to 3  %>
							<option value="<%=i%>" <%=chkIIF(QuizStatus=i,"selected","")%>><%=QuizStatusValue(i)%></option>
						<% Next %>
						</select>															
					</td>
				</tr>							
				<!--
				<tr>
					<th><div>타이틀 이미지<strong class="cRd1">*</strong></div></th>
					<td>
						<div class="inTbSet">							
							<div>	
								<p class="registImg">
									<input type="hidden" id="topTitle" name="topTitle" value="<%=topTitle%>" />
									<img id="lyrBnrImg" src="<%=chkIIF(topTitle="" or isNull(topTitle),"/images/admin_login_logo2.png",topTitle)%>" style="height:218px; border:1px solid #EEE;"/>
									<div id="lyrImgUpBtn" class="btn lMar05" style="margin-left:65px;" onclick="setImgType('title')"><label for="fileupload"><%=chkIIF(idx="" and topTitle="","이미지 업로드","이미지 수정")%></label></div>
								</p>				
							</div>					
						</div>
					</td>
				</tr>
				-->
				<tr>
					<th>
						<div>배경 이미지<strong class="cRd1">*</strong></div>
						<!--<div>mobileweb 팝업이미지<strong class="cRd1">*</strong></div>-->
						<div>pcweb 배경 이미지<strong class="cRd1">*</strong></div>
					</th>					
					<td>
						<div class="inTbSet">							
							<div><p style="text-align:center">app</p>
								<p class="registImg" style="text-align:center">									
									<img id="lyrBnrImg2" src="<%=chkIIF(backgroundImage="" or isNull(backgroundImage),"/images/admin_login_logo2.png",backgroundImage)%>" style="height:218px; border:1px solid #EEE;"/>									
									<br/>
									<select name="backgroundImage" id="backgroundImage" style="background-color: #FF7EB4" onchange="setImagePath('a');">
										<option value="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_1.jpg" style="background-color: #FF7EB4" <%=chkIIF(backgroundImage="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_1.jpg","selected","")%>>옵션1</option>
										<option value="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_2.jpg" style="background-color: #BF64FF" <%=chkIIF(backgroundImage="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_2.jpg","selected","")%>>옵션2</option>
										<option value="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_3.jpg" style="background-color: #FA5D72" <%=chkIIF(backgroundImage="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_3.jpg","selected","")%>>옵션3</option>
										<option value="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_4.jpg" style="background-color: #23DE9F" <%=chkIIF(backgroundImage="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_4.jpg","selected","")%>>옵션4</option>
									</select>																							
								</p>				
							</div>	
							<!--
							<div><p style="text-align:center">mobileweb</p>	
								<p class="registImg" style="text-align:center">									
									<img id="lyrBnrImg3" src="<%=chkIIF(MWbackgroundImage="" or isNull(MWbackgroundImage),"/images/admin_login_logo2.png",MWbackgroundImage)%>" style="height:218px; border:1px solid #EEE;"/>									
									<br/>
									<select name="MWbackgroundImage" id="MWbackgroundImage" style="background-color: #FF7EB4" onchange="setImagePath('mw');">
										<option value="" style="background-color: #FF7EB4" <%=chkIIF(true,"selected","")%>>옵션1</option>
										<option value="" style="background-color: #BF64FF" <%=chkIIF(true,"selected","")%>>옵션2</option>
										<option value="" style="background-color: #FA5D72" <%=chkIIF(true,"selected","")%>>옵션3</option>
										<option value="" style="background-color: #23DE9F" <%=chkIIF(true,"selected","")%>>옵션4</option>
									</select>										
								</p>				
							</div>	
							-->
							<div><p style="text-align:center">pcweb</p>	
								<p class="registImg" style="text-align:center">									
									<img id="lyrBnrImg4" src="<%=chkIIF(PCWBackGroundImage="" or isNull(PCWBackGroundImage),"/images/admin_login_logo2.png",PCWBackGroundImage)%>" style="height:218px; border:1px solid #EEE;"/>									
									<br/>
									<select name="PCWBackGroundImage" id="PCWBackGroundImage" style="background-color: #FF7EB4" onchange="setImagePath('pc');">
										<option value="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_1.png" style="background-color: #FF7EB4" <%=chkIIF(PCWBackGroundImage="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_1.png","selected","")%>>옵션1</option>
										<option value="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_2.png" style="background-color: #BF64FF" <%=chkIIF(PCWBackGroundImage="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_2.png","selected","")%>>옵션2</option>
										<option value="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_3.png" style="background-color: #FA5D72" <%=chkIIF(PCWBackGroundImage="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_3.png","selected","")%>>옵션3</option>
										<option value="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_4.png" style="background-color: #23DE9F" <%=chkIIF(PCWBackGroundImage="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_4.png","selected","")%>>옵션4</option>
									</select>										
								</p>				
							</div>																			
						</div>					
					<!--
						<div class="inTbSet">							
							<div><p style="text-align:center">app</p>
								<p class="registImg" style="text-align:center">
									<input type="hidden" id="backgroundImage" name="backgroundImage" value="<%=backgroundImage%>" />
									<img id="lyrBnrImg2" src="<%=chkIIF(backgroundImage="" or isNull(backgroundImage),"/images/admin_login_logo2.png",backgroundImage)%>" style="height:218px; border:1px solid #EEE;"/>
									<div id="lyrImgUpBtn2" class="btn lMar05" style="margin-left:175px;" onclick="setImgType('backgroundImage')"><label for="fileupload"><%=chkIIF(idx="" and backgroundImage="","이미지 업로드","이미지 수정")%></label></div>
								</p>				
							</div>	
							<div><p style="text-align:center">mobileweb</p>	
								<p class="registImg" style="text-align:center">
									<input type="hidden" id="MWbackgroundImage" name="MWbackgroundImage" value="<%=MWbackgroundImage%>" />
									<img id="lyrBnrImg3" src="<%=chkIIF(MWbackgroundImage="" or isNull(MWbackgroundImage),"/images/admin_login_logo2.png",MWbackgroundImage)%>" style="height:218px; border:1px solid #EEE;"/>
									<div id="lyrImgUpBtn3" class="btn lMar05" style="margin-left:175px;" onclick="setImgType('MWbackgroundImage')"><label for="fileupload"><%=chkIIF(idx="" and MWbackgroundImage="","이미지 업로드","이미지 수정")%></label></div>
								</p>				
							</div>	
							<div><p style="text-align:center">pcweb</p>	
								<p class="registImg" style="text-align:center">
									<input type="hidden" id="PCWBackGroundImage" name="PCWBackGroundImage" value="<%=PCWBackGroundImage%>" />
									<img id="lyrBnrImg4" src="<%=chkIIF(PCWBackGroundImage="" or isNull(PCWBackGroundImage),"/images/admin_login_logo2.png",PCWBackGroundImage)%>" style="height:218px; border:1px solid #EEE;"/>
									<div id="lyrImgUpBtn4" class="btn lMar05" style="margin-left:175px;" onclick="setImgType('PCWBackGroundImage')"><label for="fileupload"><%=chkIIF(idx="" and PCWBackGroundImage="","이미지 업로드","이미지 수정")%></label></div>
								</p>				
							</div>																			
						</div>
					-->						
					</td>
				</tr>			
				<tr>
					<th><div>텐퀴즈 설명<strong class="cRd1">*</strong></div></th>
					<td>
						<p><input type="text" name="quizDescription" class="formTxt" style="width:100%;" value="<%=chkIIF(QuizDescription<>"",QuizDescription,"하루에 10문제 풀고 상금(마일리지) 받아가세요!")%>" /></p>
						<p class="tPad05 fs11 cGy1">- 한글 기준 최대 40자까지 입력 가능합니다.</p>
					</td>
				</tr>
				<tr>
					<th><div>다음 차수 힌트 문항번호 <strong class="cRd1">*</strong></div></th>
					<td>
						<p><input type="number" name="questionHintNumber" class="formTxt" min=1 style="width:4%;" value="<%=QuestionHintNumber%>" /></p>
					</td>
				</tr>
				<tr>
					<th><div>총 지급 마일리지 금액 <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="number" name="totalMileage" class="formTxt" style="width:8%;" min="100" max="500" step="100" value='<%=chkIIF(TotalMileage="" or isNull(TotalMileage),100, left(TotalMileage, len(TotalMileage) - 4))%>'/>만 마일리지 
						<p class="tPad05 fs11 cGy1">- 단위 : 만원</p>	
					</td>
				</tr>												
				<tr>
					<th><div>시작일 / 종료일<strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="sDt"  name="quizStartDate" class="formTxt" style="width:8%;" value="" readonly/>
						<input type="text" id="startTime" name="quizStartTime" class="formTxt" style="width:8%;" value="" readonly />						
						 ~ 
						 <input type="text" id="eDt" name="quizEndDate" class="formTxt" style="width:8%;" value="" readonly/>			
						 <input type="text" id="endTime" name="quizEndTime" class="formTxt" style="width:8%;" value="" readonly/>		
					</td>
				</tr>					
				<tr>
					<th><div>차수 <strong class="cRd1">*</strong></div></th>
					<td>
						<p><input type="text" name="chasu" style="background-color:#eeeded" class="formTxt" style="width:14%;" value="<%=chasu%>" readonly/></p>
					</td>
				</tr>			
				<tr>
					<th><div>총 문항 수<strong class="cRd1">*</strong></div></th>
					<td>
						<input type="number" name="totalQuestionCount" class="formTxt" min=1 max=100 style="width:10%;" value="<%=chkIIF(TotalQuestionCount="" or isNull(TotalQuestionCount),10,TotalQuestionCount)%>" />
					</td>
				</tr>	
				<tr>
					<th><div>이벤트 번호<strong class="cRd1">*</strong></div></th>
					<td>
						<p><input type="text" name="productEvtNum" class="formTxt" style="width:8%;" value="<%=productEvtNum%>" /></p>
						<p class="tPad05 fs11 cGy1">- 프론트의 "문제에 나온 상품 보러가기" 버튼을 누르면 이동하는 이벤트 입니다.</p>					
					</td>
				</tr>							
				<tr>
					<th><div>하단 설명</div></th>
					<td>
						<input type="text" name="startDescription" class="formTxt" style="width:50%;" value="<%=StartDescription%>" />
						<p class="tPad05 fs11 cGy1">- 하단 버튼 밑에 보여지는 텍스트 입니다.</p>					
					</td>
				</tr>			
				<tr>
					<th><div>종료 경고문구<strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" name="endAlertTxt" class="formTxt" style="width:50%;" value="<%=chkIIF(endAlertTxt<>"", endAlertTxt, "금일 도전이 종료되었습니다.")%>" />
						<p class="tPad05 fs11 cGy1">- 텐퀴즈 종료 후 종료버튼 누르면 당일 자정까지 보여지는 alert창 텍스트 입니다.</p>					
					</td>
				</tr>			
				<tr>
					<th><div>대기 경고문구<strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" name="waitingAlertTxt" class="formTxt" style="width:50%;" value="<%=chkIIF(waitingAlertTxt<>"", waitingAlertTxt, "10시 응모가능")%>" />
						<p class="tPad05 fs11 cGy1">- 텐퀴즈 대기시 대기버튼을 누르면 보여지는 alert창 텍스트 입니다.</p>					
					</td>
				</tr>											
				</tbody>
			</table>
			<input type="hidden" name="mode" value="<%=mode%>" />
			<input type="hidden" name="idx" value="<%=idx%>">
			<input type="hidden" id="OrderChangedFlag" name="OrderChangedFlag" value="">
		
<!--========================================문항===================================================-->		
<% if mode = "modify" then %>
<h2 class="tMar20 subType">문항 정보</h2>	
			<div class="overHidden">
				<div class="ftLt">
					<input type="button" class="btnRegist btn bold fs12" value="문항 등록" onclick="popQuestionEdit()"/><p class="cGn1" id="seqAlert" style="display:none">* 문항 순서에 변경사항이 있습니다.</p>
				</div>
			</div>
			<div class="pieceList">
				<div class="rt bPad10 rPad10">
					<p class="totalNum">총 문항 수 : <strong><%=tenQuizQuestions.FtotalCount%></strong>/<strong><%=TotalQuestionCount%></strong></p>
				</div>				
				<div class="tbListWrap">
					<ul class="thDataList">
						<li>
							<p style="width:10%">문제 번호</p>							
							<p style="width:30%">문제</p>
							<p style="width:10%">차수</p>
							<p style="width:10%">문제 타입</p>
							<p style="width:10%">문제 답안</p>							
							<p style="width:10%">등록일</p>
							<p style="width:10%">수정일</p>
							<p style="width:10%">수정/삭제</p>
						</li>
					</ul>
					<!-- 리스트 -->
					<ul class="tbDataList" id="sortable">	
<% 
	for i=0 to tenQuizQuestions.FResultCount-1 

	if tenQuizQuestions.FItemList(i).FItype = 1 then
		tempQuestionType = "A"
	else
		tempQuestionType = "B"
	end if
%>							
						<li id="list<%=tenQuizQuestions.FItemList(i).FIquestionNumber%>" style="cursor:pointer;" onmouseover="this.style.backgroundColor='#D8D8D8'" onmouseout="this.style.backgroundColor=''" onclick="window.event.cancelBubble = true;	jsmodify(<%=tenQuizQuestions.FItemList(i).FIidx%>)">
							<p style="width:10%" name="VquizNumber">
								<input type="hidden" name="originalSeq" value="<%=tenQuizQuestions.FItemList(i).FIidx%>" /><%=tenQuizQuestions.FItemList(i).FIquestionNumber%>
								<input type="hidden" name="originalSeq2" value="<%=tenQuizQuestions.FItemList(i).FIquestionNumber%>" />
							</p>
							<p style="width:30%"><input type="hidden" name="newSeq" value="" /><%=tenQuizQuestions.FItemList(i).FIquestion%></p>
							<p style="width:10%"><%=tenQuizQuestions.FItemList(i).FIchasu%></p>							
							<p style="width:10%"><%=tempQuestionType%></p>
							<p style="width:10%"><%=tenQuizQuestions.FItemList(i).FIanswer%></p>							
							<p style="width:10%"><%=tenQuizQuestions.FItemList(i).FIregistDate%></p>
							<p style="width:10%"><%=tenQuizQuestions.FItemList(i).FIlastUpDate%></p>
							<p style="width:10%">
								<a href="javascript:jsmodify(<%=tenQuizQuestions.FItemList(i).FIidx%>)" class="cBl1 tLine">[수정]</a>
								<a href="javascript:jsdelete(<%=tenQuizQuestions.FItemList(i).FIidx%>, <%=tenQuizQuestions.FItemList(i).FIquestionNumber%>)" class="cBl1 tLine">[삭제]</a>
							</p>							
						</li>	
<% Next %>												
					</ul>							
				</div>
			</div>
<% end if %>
	</div>			
	    </form>
	<div class="popBtnWrap">
		<!-- input type="button" value="미리보기" onclick="" class="cBl2" style="width:100px; height:30px;" / -->
		<input type="button" value="취소" onclick="jsgolist();" style="width:100px; height:30px;" />
		<input type="button" value="저장" onclick="addContent();" class="cRd1" style="width:100px; height:30px;" />
		<!-- input type="button" value="수정" onclick="" class="cRd1" style="width:100px; height:30px;" / -->
	</div>
		<!-- 타이틀, 배경 이미지 업로드 Form -->
	<form name="frmUpload" id="ajaxform" action="<%=uploadImgUrl%>/linkweb/common/simpleCommonImgUploadProc.asp" method="post" enctype="multipart/form-data" style="display:none; height:0px;width:0px;">
		<input type="file" name="upfile" id="fileupload" onchange="jsCheckUpload();" accept="image/*" />
		<input type="hidden" name="mode" id="fileupmode" value="upload">
		<input type="hidden" name="div" value="TQ">
		<input type="hidden" name="upPath" value="/appmanage/tenquizimg/">
		<input type="hidden" name="tuid" value="<%=encUsrId%>">
		<input type="hidden" name="prefile" id="filepre" value="<%=topTitle%>">
		<input type="hidden" name="imgtype">
	</form>				
	<form name="delFrm" method="post">
		<input type="hidden" name="idx">
		<input type="hidden" name="mode" value="subdelete">
	</form>					
</div>
