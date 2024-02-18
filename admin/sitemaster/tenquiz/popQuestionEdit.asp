<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/classes/sitemasterclass/TenQuizCls.asp" -->
<%
Dim userid, encUsrId, tmpTx, tmpRn
userid = session("ssBctId")
'// 변수 선언
dim i
dim j
dim z
Dim listidx , subIdx, sqlStr
Dim mode
dim chasu
dim totalquestioncount
dim isQuestionNumbers()

dim questionType
dim questionNumber
dim question
dim questionType1Image1
dim questionType1Image2
dim questionType1Image3
dim questionType1Image4
dim questionExample1
dim questionExample2
dim questionExample3
dim questionExample4
dim questionExample1img
dim questionExample2img
dim questionExample3img
dim questionExample4img
dim type2TextExample1
dim type2TextExample2
dim type2TextExample3
dim type2TextExample4
dim answer
dim isusing
dim numOfType1Image

Randomize()
tmpTx = split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z",",")
tmpRn = tmpTx(int(Rnd*26))
tmpRn = tmpRn & tmpTx(int(Rnd*26))
	encUsrId = tenEnc(tmpRn & userid)	

public Function HasQuestionNumber(arr, length, num)
	dim i
	dim result
	dim sum
	result = true
	sum = 0

	For i=1 To length
		if arr(i) = num then
			result = false
			exit for
		else
			result = true
		end if
	Next	

	HasQuestionNumber = result
end Function

'// 파라메터 접수
listidx = request("listidx")
subidx = request("subidx")

If subidx = "" Then
	mode = "subadd"
Else
	mode = "submodify"
End If

If listidx <> "" Then
	sqlStr = " select chasu "
	sqlStr = sqlStr & " , totalquestioncount "
	sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_PlayingTenQuizData "
	sqlStr = sqlStr & " where idx=" & listidx 
	rsget.Open sqlStr, dbget, 1
		chasu = rsget("chasu")
		totalquestioncount = rsget("totalquestioncount")
	rsget.close
End If

If listidx <> "" Then
	redim preserve isQuestionNumbers(totalquestioncount)

	sqlStr = " select questionNumber "
	sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_PlayingTenQuizQuestionData "
	sqlStr = sqlStr & " where chasu=" & chasu
	sqlStr = sqlStr & " and isusing = 'Y'" 
	rsget.Open sqlStr, dbget, 1
		
	if  not rsget.EOF  then
		j = 1
		do until rsget.eof				
			isQuestionNumbers(j) = rsget("questionNumber")			
			j=j+1
			rsget.moveNext
		loop
	end if

	rsget.close
End If

If mode = "submodify" then

	dim tenQuizSubItem
	set tenQuizSubItem = new TenQuiz
	tenQuizSubItem.FRectSubIdx = subidx
	tenQuizSubItem.GetOneSubItem()

	questionType		= tenQuizSubItem.FoneItem.FItype
	questionNumber		= tenQuizSubItem.FoneItem.FIquestionNumber
	question			= tenQuizSubItem.FoneItem.FIquestion
	questionType1Image1	= tenQuizSubItem.FoneItem.FIquestionType1Image1
	questionType1Image2	= tenQuizSubItem.FoneItem.FIquestionType1Image2
	questionType1Image3	= tenQuizSubItem.FoneItem.FIquestionType1Image3
	questionType1Image4	= tenQuizSubItem.FoneItem.FIquestionType1Image4
	type2TextExample1	= tenQuizSubItem.FoneItem.FItype2TextExample1
	type2TextExample2	= tenQuizSubItem.FoneItem.FItype2TextExample2
	type2TextExample3	= tenQuizSubItem.FoneItem.FItype2TextExample3
	type2TextExample4	= tenQuizSubItem.FoneItem.FItype2TextExample4
	if questionType = 1 then
		questionExample1	= tenQuizSubItem.FoneItem.FIquestionExample1
		questionExample2	= tenQuizSubItem.FoneItem.FIquestionExample2
		questionExample3	= tenQuizSubItem.FoneItem.FIquestionExample3
		questionExample4	= tenQuizSubItem.FoneItem.FIquestionExample4	
	else
		questionExample1img	= tenQuizSubItem.FoneItem.FIquestionExample1
		questionExample2img	= tenQuizSubItem.FoneItem.FIquestionExample2
		questionExample3img	= tenQuizSubItem.FoneItem.FIquestionExample3
		questionExample4img	= tenQuizSubItem.FoneItem.FIquestionExample4	
	end if	

	answer				= tenQuizSubItem.FoneItem.FIanswer
	isusing				= tenQuizSubItem.FoneItem.FIIsUsing
	numOfType1Image		= tenQuizSubItem.FoneItem.FINumofType1Image
	
End If 
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
body {background-color:#fff;}
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/evol.colorpicker.min.js"></script>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<script type="text/javascript">
$(function(){
	initiateValues();	
})
function jsCheckUpload() {
	var gubun = document.frmUpload.imgtype.value;
	var mainfrm = document.frm
	var test = $("input[id="+gubun+"]").val();
	console.log(gubun);	
	console.log(test);
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
function chgNumOfImgContainer(){
	var numOfImg = document.frm.numOfType1Image.value;
	for(var i = 4; i >= 1 ; i-- ){		
		$("div[id=type1Img"+i+"]").css("display","")
	}	
	for(var i = 4; i > numOfImg ; i-- ){		
		$("div[id=type1Img"+i+"]").css("display","none")
	}
}
function submitContent(){
	var mainfrm = document.frm;
	
	if(mainfrm.question.value === ""){
		alert("문항을 입력해 주세요.");
		mainfrm.question.focus();
		return false;
	}
	if(mainfrm.type.value == 1 && (
	   mainfrm.questionExample1.value === ""
	|| mainfrm.questionExample2.value === ""
	|| mainfrm.questionExample3.value === ""
	|| mainfrm.questionExample4.value === ""
	)){
		alert("보기를 입력해 주세요.");
		return false;
	}
		mainfrm.action="tenquizaction.asp";
		mainfrm.submit();		
	}
function chkType(){
	var frm = document.frm;
	if(frm.type.value == 1){		
		$("#type1").css('display','');
		$("#type1imgs").css('display','');
		$("#type2").css('display','none');		
	}else{
		$("#type1").css('display','none');
		$("#type1imgs").css('display','none');
		$("#type2").css('display','');		
	}		
}
function initiateValues(){
	chkType();
	chgNumOfImgContainer()
} 

//-- jsImgView : 이미지 확대화면 새창으로 보여주기 --//
function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}
function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=450,height=300');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
function setImgType(type){
	document.frmUpload.imgtype.value = type;
	return false;
}	
</script>

<form name="frm" method="post">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="mode" value="<%=mode%>" />
<input type="hidden" name="idx" value="<%=subIdx%>" />
<div class="popWinV17">
<% If mode = "subadd" Then %>
	<h2 class="tMar20 subType" style="margin-left:30px;">문항 추가</h2>
<% Else %>
	<h2 class="tMar20 subType" style="margin-left:30px;">문항 수정</h2>
<% End if %>
	<div class="popContainerV17 pad30">
		<table class="tbType1 writeTb tMar10">
			<colgroup>
				<col width="18%" /><col width="" />
			</colgroup>
			<tbody>				
			<tr>
				<th><div>문제 유형 <strong class="cRd1">*</strong></div></th>
				<td>
					A타입<input type="radio" name="type" value=1 onclick="chkType();" <%=chkIIF(questionType="1" or questionType="" ,"checked","")%>></BR>
					B타입<input type="radio" name="type" value=2 onclick="chkType();" <%=chkIIF(questionType="2" ,"checked","")%>>
				</td>
			</tr>
			<tr>
				<th><div>문제 번호<strong class="cRd1">*</strong></div></th>
				<td>
					<% if mode<>"submodify" then %>						
					<select name="questionNumber"><!-- 문제번호 선택 불가 변경 -->					
						<% for i=1 to totalquestioncount  %>															
							<% if HasQuestionNumber(isQuestionNumbers, totalquestioncount, i) then %>
							<option value=<%=i%>><%=i%>번 문제</option>						
							<% 
							exit for '//문제번호 선택 불가 변경 
							end if 
							%>
						<% Next %>					
					</select>																
					<% else %>
						<input type="hidden" name="questionNumber" value="<%=questionNumber%>">
						<p><b><%=questionNumber%>번 문제</b></p>
					<% end if %>
				</td>
			</tr>				
			<tr>
				<th><div>차수 <strong class="cRd1">*</strong></div></th>
				<td>
					<p><input type="text" name="chasu" style="background-color:#eeeded" class="formTxt" style="width:14%;" value="<%=chasu%>" readonly/></p>
					<p class="tPad05 fs11 cGy1">- 한글 기준 최대 40자까지 입력 가능합니다.</p>
				</td>
			</tr>				
			<tr>
				<th><div>문항<strong class="cRd1">*</strong></div></th>
				<td>
					<textarea name="question" style="width:100%; height:80px;" value=""><%=question%></textarea>			
				</td>
			</tr>							
<!--type 1일 경우-->				
			<tr id="type1imgs">
				<th><div>문항(A타입)<strong class="cRd1">*</strong></div></th>				
				<td>
					이미지 갯수
					<select name="numOfType1Image" onchange="chgNumOfImgContainer();">
						<option value=1 <%=chkIIF(numOfType1Image = 1,"selected","")%>>1</option>
						<option value=2 <%=chkIIF(numOfType1Image = 2,"selected","")%>>2</option>						
						<option value=4 <%=chkIIF(numOfType1Image = 4,"selected","")%>>4</option>
					</select>				
					</br>
					<div class="inTbSet">							
						<div id="type1Img1">	
							<p class="registImg">
								<input type="hidden" id="questionType1Image1" name="questionType1Image1" value="<%=questionType1Image1%>" />
								<img id="questionType1Image1src" src="<%=chkIIF(questionType1Image1="" or isNull(questionType1Image1),"/images/admin_login_logo2.png",questionType1Image1)%>" style="height:118px; border:1px solid #EEE;"/>								
								<div class="btn lMar05" onclick="setImgType('questionType1Image1')" ><label for="fileupload"><%=chkIIF(questionType1Image1="","이미지 업로드","이미지 수정")%></label></div>
							</p>				
						</div>	
						<div id="type1Img2">	
							<p class="registImg">
								<input type="hidden" id="questionType1Image2" name="questionType1Image2" value="<%=questionType1Image2%>" />
								<img id="questionType1Image2src" src="<%=chkIIF(questionType1Image2="" or isNull(questionType1Image2),"/images/admin_login_logo2.png",questionType1Image2)%>" style="height:118px; border:1px solid #EEE;"/>
								<div class="btn lMar05" onclick="setImgType('questionType1Image2')" ><label for="fileupload"><%=chkIIF(questionType1Image2="","이미지 업로드","이미지 수정")%></label></div>
							</p>				
						</div>							
						<div id="type1Img3">	
							<p class="registImg">
								<input type="hidden" id="questionType1Image3" name="questionType1Image3" value="<%=questionType1Image3%>" />
								<img id="questionType1Image3src" src="<%=chkIIF(questionType1Image3="" or isNull(questionType1Image3),"/images/admin_login_logo2.png",questionType1Image3)%>" style="height:118px; border:1px solid #EEE;"/>
								<div class="btn lMar05" onclick="setImgType('questionType1Image3')" ><label for="fileupload"><%=chkIIF(questionType1Image3="","이미지 업로드","이미지 수정")%></label></div>
							</p>				
						</div>	
						<div id="type1Img4">		
							<p class="registImg">
								<input type="hidden" id="questionType1Image4" name="questionType1Image4" value="<%=questionType1Image4%>" />
								<img id="questionType1Image4src" src="<%=chkIIF(questionType1Image4="" or isNull(questionType1Image4),"/images/admin_login_logo2.png",questionType1Image4)%>" style="height:118px; border:1px solid #EEE;"/>
								<div class="btn lMar05" onclick="setImgType('questionType1Image4')" ><label for="fileupload"><%=chkIIF(questionType1Image4="","이미지 업로드","이미지 수정")%></label></div>
							</p>				
						</div>																							
					</div>					
				</td>
			</tr>
			<tr id="type1">
				<th><div>보기(A타입)<strong class="cRd1">*</strong></div></th>
				<td>
					①<input type="text" name="questionExample1" class="formTxt" style="width:70%;" value="<%=questionExample1%>" /><br>
					②<input type="text" name="questionExample2" class="formTxt" style="width:70%;" value="<%=questionExample2%>" /><br>
					③<input type="text" name="questionExample3" class="formTxt" style="width:70%;" value="<%=questionExample3%>" /><br>
					④<input type="text" name="questionExample4" class="formTxt" style="width:70%;" value="<%=questionExample4%>" /><br>
				</td>
			</tr>								
<!--type 1일 경우-->						
<!--type 2일 경우-->			
			<tr id="type2">
				<th><div>보기(B타입)<strong class="cRd1">*</strong></div></th>
				<td>					
					<div class="inTbSet">							
					①
						<div>	
							<p class="registImg">
								<input type="hidden" id="questionExample1img" name="questionExample1img" value="<%=questionExample1img%>" />
								<img id="questionExample1imgsrc" src="<%=chkIIF(questionExample1img="" or isNull(questionExample1img),"/images/admin_login_logo2.png",questionExample1img)%>" style="height:138px; border:1px solid #EEE;"/>																
							</p>											
							<div><input type="text" name="type2TextExample1" length="" value="<%=type2TextExample1%>" style="width:140px"></div>
							<div class="btn lMar05" onclick="setImgType('questionExample1img')" ><label for="fileupload"><%=chkIIF(questionExample1img="","이미지 업로드","이미지 수정")%></label></div>							
						</div>	
					②	
						<div>	
							<p class="registImg">
								<input type="hidden" id="questionExample2img" name="questionExample2img" value="<%=questionExample2img%>" />
								<img id="questionExample2imgsrc" src="<%=chkIIF(questionExample2img="" or isNull(questionExample2img),"/images/admin_login_logo2.png",questionExample2img)%>" style="height:138px; border:1px solid #EEE;"/>
							</p>
							<div><input type="text" name="type2TextExample2" length="" value="<%=type2TextExample2%>" style="width:140px"></div>				
							<div class="btn lMar05" onclick="setImgType('questionExample2img')" ><label for="fileupload"><%=chkIIF(questionExample2img="","이미지 업로드","이미지 수정")%></label></div>
						</div>						
					</div>	
					<div class="inTbSet">							
					③
						<div>	
							<p class="registImg">
								<input type="hidden" id="questionExample3img" name="questionExample3img" value="<%=questionExample3img%>" />
								<img id="questionExample3imgsrc" src="<%=chkIIF(questionExample3img="" or isNull(questionExample3img),"/images/admin_login_logo2.png",questionExample3img)%>" style="height:138px; border:1px solid #EEE;"/>								
							</p>				
							<div><input type="text" name="type2TextExample3" length="" value="<%=type2TextExample3%>" style="width:140px"></div>
							<div class="btn lMar05" onclick="setImgType('questionExample3img')" ><label for="fileupload"><%=chkIIF(questionExample3img="","이미지 업로드","이미지 수정")%></label></div>
						</div>	
					④	
						<div>	
							<p class="registImg">
								<input type="hidden" id="questionExample4img" name="questionExample4img" value="<%=questionExample4img%>" />
								<img id="questionExample4imgsrc" src="<%=chkIIF(questionExample4img="" or isNull(questionExample4img),"/images/admin_login_logo2.png",questionExample4img)%>" style="height:138px; border:1px solid #EEE;"/>
							</p>		
							<div><input type="text" name="type2TextExample4" length="" value="<%=type2TextExample4%>" style="width:140px"></div>		
							<div class="btn lMar05" onclick="setImgType('questionExample4img')" ><label for="fileupload"><%=chkIIF(questionExample4img="","이미지 업로드","이미지 수정")%></label></div>							
						</div>						
					</div>											
				</td>
			</tr>			
<!--type 2일 경우-->						
			<tr>
				<th><div>답<strong class="cRd1">*</strong></div></th>
				<td>
					<select name="answer">
					<% for i=1 to 4  %>
						<option value="<%=i%>" <%=chkIIF(answer=i,"selected","")%>><%=i%></option>
					<% Next %>
					</select>				
				</td>
			</tr>			
			</tbody>
		</table>
	</div>
	<div class="popBtnWrap">
		<!-- input type="button" value="미리보기" onclick="" class="cBl2" style="width:100px; height:30px;" / -->
		<input type="button" value="취소" onclick="window.close();" style="width:100px; height:30px;" />
		<input type="button" value="저장" onclick="submitContent();" class="cRd1" style="width:100px; height:30px;" />
		<!-- input type="button" value="수정" onclick="" class="cRd1" style="width:100px; height:30px;" / -->
	</div>	
</div>	
</form>
<form name="frmUpload" id="ajaxform" action="<%=uploadImgUrl%>/linkweb/common/simpleCommonImgUploadProc.asp" method="post" enctype="multipart/form-data" style="display:none; height:0px;width:0px;">
	<input type="file" name="upfile" id="fileupload" onchange="jsCheckUpload();" accept="image/*" />
	<input type="hidden" name="mode" id="fileupmode" value="upload">
	<input type="hidden" name="div" value="TQ">
	<input type="hidden" name="upPath" value="/appmanage/tenquizimg/">
	<input type="hidden" name="tuid" value="<%=encUsrId%>">
	<input type="hidden" name="prefile" id="filepre" >	
	<input type="hidden" name="imgtype">
</form>					
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->