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
QuizStatusValue(1) = "��� ���"
QuizStatusValue(2) = "����"
QuizStatusValue(3) = "����"
dim mode

dim quizStartTime
dim quizEndTime

'���� ����
dim tenQuizQuestions
dim tempQuestionType

idx = requestCheckvar(request("idx"),16) 

'�׽�Ʈ������

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
	var arrDayMin = ["��","��","ȭ","��","��","��","��"];
	var arrMonth = ["1��","2��","3��","4��","5��","6��","7��","8��","9��","10��","11��","12��"];
    $("#sDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
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
		prevText: '������', nextText: '������', yearSuffix: '��',
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
			//�������� validation check�� �ʿ��Ұ��
			beforeSubmit: function (data, frm, opt) {
				if(!(/\.(jpg|jpeg|png)$/i).test(frm[0].upfile.value)) {
					alert("JPG,PNG �̹������ϸ� ���ε� �Ͻ� �� �ֽ��ϴ�.");
					$("#fileupload").val("");
					return false;
				}
				$("#lyrPrgs").show();
			},
			//submit������ ó��
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
					alert("ó���� ������ �߻��߽��ϴ�.\n" + responseText);
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
// �������� ���� ���� ó��
function jsgolist(){
	self.location.href="/admin/sitemaster/tenquiz/";
	}
function addContent(){
	var mainfrm = document.frm;		
		if(mainfrm.backgroundImage.value==""){
			alert('��� �̹����� �Է��� �ּ���.');
			return false;
		}
		if(mainfrm.quizDescription.value==""){			
			alert('���� ������ �Է����ּ���.');
			mainfrm.quizDescription.focus();
			return false;
		}
		if(mainfrm.questionHintNumber.value==""){
			alert('��Ʈ ���� ��ȣ�� �Է��� �ּ���.');
			mainfrm.questionHintNumber.focus();
			return false;
		}
		if(mainfrm.totalMileage.value==""){
			alert('�� ���ϸ����� �Է��� �ּ���.');
			mainfrm.totalMileage.focus();
			return false;
		}
		if(mainfrm.quizStartDate.value==""){
			alert('������ ���� ��¥�� �Է��� �ּ���.');
			mainfrm.quizStartDate.focus();
			return false;
		}
		if(mainfrm.quizEndDate.value==""){
			alert('������ ���� ��¥�� �Է��� �ּ���.');
			mainfrm.quizEndDate.focus();
			return false;
		}
		if(mainfrm.totalQuestionCount.value==""){			
			alert('�� ���� ���� �Է��� �ּ���.');
			mainfrm.totalQuestionCount.focus();
			return false;
		}
		// if(mainfrm.startDescription.value==""){			
		// 	alert('�ϴ� ������ �Է��� �ּ���.');
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
	
	if(confirm(`${qn}�� ������ �����Ͻðڽ��ϱ�?`)){
		var frm = document.delFrm;
		frm.idx.value = v;
		frm.action = "tenquizaction.asp"
		frm.submit();
	}
}	
function popQuestionEdit(){
	<% if idx <> "" then %>
		<% if tenQuizQuestions.FtotalCount >= TotalQuestionCount then %>
			alert('�� ���� ������ �ʰ� ����� �� �����ϴ�.');
		<% else %>
			var popwin = window.open("/admin/sitemaster/tenquiz/popQuestionEdit.asp?listidx=<%=idx%>", "popup_item", "width=800,height=800,scrollbars=yes,resizable=yes");
			popwin.focus();
		<% end if %>
	<% else %>
		alert("��� ���� ������ּ���.");
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
			alert("�� ���� ������ ����ϼž� �����Ͻ� �� �ֽ��ϴ�. "+currentQuizCnt+"/"+quizTotalCnt);
			stateValue.value = 1
			return false;
		}		
	}
}	
// ���ε� ���� Ȯ�� �� ó��
</script>
<div class="popWinV17">
	<h1>������ ���/����</h1>
	<button type="button" class="btn btn2" style="position:absolute; right:15px; top:7px;">����</button>
	<div class="popContainerV17 pad30">
		<p class="cGn1">* ���� �ۼ� �� �ݵ�� '����' ��ư�� �����ּ���.</p>
		<% if mode = "modify" then%>
		<h2 class="tMar20 subType">������ ����</h2>
		<% else %>
		<h2 class="tMar20 subType">������ ���</h2>
		<% end if %>		
		<%if mode <> "add" then%>
		<p class="tPad10 fs11" style="border-top:1px dashed #c9c9c9">
			<span class="cGy1"><%=AdminName&"  "%> <%=RegistDate%> ���</span><br /><span class="cOr1"><%=AdminModifyerName&"  "%> <%=LastUpDate%> ��������</span>
		</p>				
		<% end if %>		
		<form name="frm">
			<table class="tbType1 writeTb tMar10">
				<colgroup>
					<col width="18%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>����<strong class="cRd1">*</strong></div></th>
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
					<th><div>Ÿ��Ʋ �̹���<strong class="cRd1">*</strong></div></th>
					<td>
						<div class="inTbSet">							
							<div>	
								<p class="registImg">
									<input type="hidden" id="topTitle" name="topTitle" value="<%=topTitle%>" />
									<img id="lyrBnrImg" src="<%=chkIIF(topTitle="" or isNull(topTitle),"/images/admin_login_logo2.png",topTitle)%>" style="height:218px; border:1px solid #EEE;"/>
									<div id="lyrImgUpBtn" class="btn lMar05" style="margin-left:65px;" onclick="setImgType('title')"><label for="fileupload"><%=chkIIF(idx="" and topTitle="","�̹��� ���ε�","�̹��� ����")%></label></div>
								</p>				
							</div>					
						</div>
					</td>
				</tr>
				-->
				<tr>
					<th>
						<div>��� �̹���<strong class="cRd1">*</strong></div>
						<!--<div>mobileweb �˾��̹���<strong class="cRd1">*</strong></div>-->
						<div>pcweb ��� �̹���<strong class="cRd1">*</strong></div>
					</th>					
					<td>
						<div class="inTbSet">							
							<div><p style="text-align:center">app</p>
								<p class="registImg" style="text-align:center">									
									<img id="lyrBnrImg2" src="<%=chkIIF(backgroundImage="" or isNull(backgroundImage),"/images/admin_login_logo2.png",backgroundImage)%>" style="height:218px; border:1px solid #EEE;"/>									
									<br/>
									<select name="backgroundImage" id="backgroundImage" style="background-color: #FF7EB4" onchange="setImagePath('a');">
										<option value="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_1.jpg" style="background-color: #FF7EB4" <%=chkIIF(backgroundImage="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_1.jpg","selected","")%>>�ɼ�1</option>
										<option value="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_2.jpg" style="background-color: #BF64FF" <%=chkIIF(backgroundImage="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_2.jpg","selected","")%>>�ɼ�2</option>
										<option value="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_3.jpg" style="background-color: #FA5D72" <%=chkIIF(backgroundImage="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_3.jpg","selected","")%>>�ɼ�3</option>
										<option value="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_4.jpg" style="background-color: #23DE9F" <%=chkIIF(backgroundImage="http://fiximage.10x10.co.kr/m/2018/tenquiz/tit_tenquiz_4.jpg","selected","")%>>�ɼ�4</option>
									</select>																							
								</p>				
							</div>	
							<!--
							<div><p style="text-align:center">mobileweb</p>	
								<p class="registImg" style="text-align:center">									
									<img id="lyrBnrImg3" src="<%=chkIIF(MWbackgroundImage="" or isNull(MWbackgroundImage),"/images/admin_login_logo2.png",MWbackgroundImage)%>" style="height:218px; border:1px solid #EEE;"/>									
									<br/>
									<select name="MWbackgroundImage" id="MWbackgroundImage" style="background-color: #FF7EB4" onchange="setImagePath('mw');">
										<option value="" style="background-color: #FF7EB4" <%=chkIIF(true,"selected","")%>>�ɼ�1</option>
										<option value="" style="background-color: #BF64FF" <%=chkIIF(true,"selected","")%>>�ɼ�2</option>
										<option value="" style="background-color: #FA5D72" <%=chkIIF(true,"selected","")%>>�ɼ�3</option>
										<option value="" style="background-color: #23DE9F" <%=chkIIF(true,"selected","")%>>�ɼ�4</option>
									</select>										
								</p>				
							</div>	
							-->
							<div><p style="text-align:center">pcweb</p>	
								<p class="registImg" style="text-align:center">									
									<img id="lyrBnrImg4" src="<%=chkIIF(PCWBackGroundImage="" or isNull(PCWBackGroundImage),"/images/admin_login_logo2.png",PCWBackGroundImage)%>" style="height:218px; border:1px solid #EEE;"/>									
									<br/>
									<select name="PCWBackGroundImage" id="PCWBackGroundImage" style="background-color: #FF7EB4" onchange="setImagePath('pc');">
										<option value="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_1.png" style="background-color: #FF7EB4" <%=chkIIF(PCWBackGroundImage="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_1.png","selected","")%>>�ɼ�1</option>
										<option value="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_2.png" style="background-color: #BF64FF" <%=chkIIF(PCWBackGroundImage="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_2.png","selected","")%>>�ɼ�2</option>
										<option value="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_3.png" style="background-color: #FA5D72" <%=chkIIF(PCWBackGroundImage="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_3.png","selected","")%>>�ɼ�3</option>
										<option value="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_4.png" style="background-color: #23DE9F" <%=chkIIF(PCWBackGroundImage="http://fiximage.10x10.co.kr/web2018/tenquiz/bg_tenquiz_4.png","selected","")%>>�ɼ�4</option>
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
									<div id="lyrImgUpBtn2" class="btn lMar05" style="margin-left:175px;" onclick="setImgType('backgroundImage')"><label for="fileupload"><%=chkIIF(idx="" and backgroundImage="","�̹��� ���ε�","�̹��� ����")%></label></div>
								</p>				
							</div>	
							<div><p style="text-align:center">mobileweb</p>	
								<p class="registImg" style="text-align:center">
									<input type="hidden" id="MWbackgroundImage" name="MWbackgroundImage" value="<%=MWbackgroundImage%>" />
									<img id="lyrBnrImg3" src="<%=chkIIF(MWbackgroundImage="" or isNull(MWbackgroundImage),"/images/admin_login_logo2.png",MWbackgroundImage)%>" style="height:218px; border:1px solid #EEE;"/>
									<div id="lyrImgUpBtn3" class="btn lMar05" style="margin-left:175px;" onclick="setImgType('MWbackgroundImage')"><label for="fileupload"><%=chkIIF(idx="" and MWbackgroundImage="","�̹��� ���ε�","�̹��� ����")%></label></div>
								</p>				
							</div>	
							<div><p style="text-align:center">pcweb</p>	
								<p class="registImg" style="text-align:center">
									<input type="hidden" id="PCWBackGroundImage" name="PCWBackGroundImage" value="<%=PCWBackGroundImage%>" />
									<img id="lyrBnrImg4" src="<%=chkIIF(PCWBackGroundImage="" or isNull(PCWBackGroundImage),"/images/admin_login_logo2.png",PCWBackGroundImage)%>" style="height:218px; border:1px solid #EEE;"/>
									<div id="lyrImgUpBtn4" class="btn lMar05" style="margin-left:175px;" onclick="setImgType('PCWBackGroundImage')"><label for="fileupload"><%=chkIIF(idx="" and PCWBackGroundImage="","�̹��� ���ε�","�̹��� ����")%></label></div>
								</p>				
							</div>																			
						</div>
					-->						
					</td>
				</tr>			
				<tr>
					<th><div>������ ����<strong class="cRd1">*</strong></div></th>
					<td>
						<p><input type="text" name="quizDescription" class="formTxt" style="width:100%;" value="<%=chkIIF(QuizDescription<>"",QuizDescription,"�Ϸ翡 10���� Ǯ�� ���(���ϸ���) �޾ư�����!")%>" /></p>
						<p class="tPad05 fs11 cGy1">- �ѱ� ���� �ִ� 40�ڱ��� �Է� �����մϴ�.</p>
					</td>
				</tr>
				<tr>
					<th><div>���� ���� ��Ʈ ���׹�ȣ <strong class="cRd1">*</strong></div></th>
					<td>
						<p><input type="number" name="questionHintNumber" class="formTxt" min=1 style="width:4%;" value="<%=QuestionHintNumber%>" /></p>
					</td>
				</tr>
				<tr>
					<th><div>�� ���� ���ϸ��� �ݾ� <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="number" name="totalMileage" class="formTxt" style="width:8%;" min="100" max="500" step="100" value='<%=chkIIF(TotalMileage="" or isNull(TotalMileage),100, left(TotalMileage, len(TotalMileage) - 4))%>'/>�� ���ϸ��� 
						<p class="tPad05 fs11 cGy1">- ���� : ����</p>	
					</td>
				</tr>												
				<tr>
					<th><div>������ / ������<strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="sDt"  name="quizStartDate" class="formTxt" style="width:8%;" value="" readonly/>
						<input type="text" id="startTime" name="quizStartTime" class="formTxt" style="width:8%;" value="" readonly />						
						 ~ 
						 <input type="text" id="eDt" name="quizEndDate" class="formTxt" style="width:8%;" value="" readonly/>			
						 <input type="text" id="endTime" name="quizEndTime" class="formTxt" style="width:8%;" value="" readonly/>		
					</td>
				</tr>					
				<tr>
					<th><div>���� <strong class="cRd1">*</strong></div></th>
					<td>
						<p><input type="text" name="chasu" style="background-color:#eeeded" class="formTxt" style="width:14%;" value="<%=chasu%>" readonly/></p>
					</td>
				</tr>			
				<tr>
					<th><div>�� ���� ��<strong class="cRd1">*</strong></div></th>
					<td>
						<input type="number" name="totalQuestionCount" class="formTxt" min=1 max=100 style="width:10%;" value="<%=chkIIF(TotalQuestionCount="" or isNull(TotalQuestionCount),10,TotalQuestionCount)%>" />
					</td>
				</tr>	
				<tr>
					<th><div>�̺�Ʈ ��ȣ<strong class="cRd1">*</strong></div></th>
					<td>
						<p><input type="text" name="productEvtNum" class="formTxt" style="width:8%;" value="<%=productEvtNum%>" /></p>
						<p class="tPad05 fs11 cGy1">- ����Ʈ�� "������ ���� ��ǰ ��������" ��ư�� ������ �̵��ϴ� �̺�Ʈ �Դϴ�.</p>					
					</td>
				</tr>							
				<tr>
					<th><div>�ϴ� ����</div></th>
					<td>
						<input type="text" name="startDescription" class="formTxt" style="width:50%;" value="<%=StartDescription%>" />
						<p class="tPad05 fs11 cGy1">- �ϴ� ��ư �ؿ� �������� �ؽ�Ʈ �Դϴ�.</p>					
					</td>
				</tr>			
				<tr>
					<th><div>���� �����<strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" name="endAlertTxt" class="formTxt" style="width:50%;" value="<%=chkIIF(endAlertTxt<>"", endAlertTxt, "���� ������ ����Ǿ����ϴ�.")%>" />
						<p class="tPad05 fs11 cGy1">- ������ ���� �� �����ư ������ ���� �������� �������� alertâ �ؽ�Ʈ �Դϴ�.</p>					
					</td>
				</tr>			
				<tr>
					<th><div>��� �����<strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" name="waitingAlertTxt" class="formTxt" style="width:50%;" value="<%=chkIIF(waitingAlertTxt<>"", waitingAlertTxt, "10�� ���𰡴�")%>" />
						<p class="tPad05 fs11 cGy1">- ������ ���� ����ư�� ������ �������� alertâ �ؽ�Ʈ �Դϴ�.</p>					
					</td>
				</tr>											
				</tbody>
			</table>
			<input type="hidden" name="mode" value="<%=mode%>" />
			<input type="hidden" name="idx" value="<%=idx%>">
			<input type="hidden" id="OrderChangedFlag" name="OrderChangedFlag" value="">
		
<!--========================================����===================================================-->		
<% if mode = "modify" then %>
<h2 class="tMar20 subType">���� ����</h2>	
			<div class="overHidden">
				<div class="ftLt">
					<input type="button" class="btnRegist btn bold fs12" value="���� ���" onclick="popQuestionEdit()"/><p class="cGn1" id="seqAlert" style="display:none">* ���� ������ ��������� �ֽ��ϴ�.</p>
				</div>
			</div>
			<div class="pieceList">
				<div class="rt bPad10 rPad10">
					<p class="totalNum">�� ���� �� : <strong><%=tenQuizQuestions.FtotalCount%></strong>/<strong><%=TotalQuestionCount%></strong></p>
				</div>				
				<div class="tbListWrap">
					<ul class="thDataList">
						<li>
							<p style="width:10%">���� ��ȣ</p>							
							<p style="width:30%">����</p>
							<p style="width:10%">����</p>
							<p style="width:10%">���� Ÿ��</p>
							<p style="width:10%">���� ���</p>							
							<p style="width:10%">�����</p>
							<p style="width:10%">������</p>
							<p style="width:10%">����/����</p>
						</li>
					</ul>
					<!-- ����Ʈ -->
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
								<a href="javascript:jsmodify(<%=tenQuizQuestions.FItemList(i).FIidx%>)" class="cBl1 tLine">[����]</a>
								<a href="javascript:jsdelete(<%=tenQuizQuestions.FItemList(i).FIidx%>, <%=tenQuizQuestions.FItemList(i).FIquestionNumber%>)" class="cBl1 tLine">[����]</a>
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
		<!-- input type="button" value="�̸�����" onclick="" class="cBl2" style="width:100px; height:30px;" / -->
		<input type="button" value="���" onclick="jsgolist();" style="width:100px; height:30px;" />
		<input type="button" value="����" onclick="addContent();" class="cRd1" style="width:100px; height:30px;" />
		<!-- input type="button" value="����" onclick="" class="cRd1" style="width:100px; height:30px;" / -->
	</div>
		<!-- Ÿ��Ʋ, ��� �̹��� ���ε� Form -->
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
