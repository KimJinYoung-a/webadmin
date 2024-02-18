<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_attendance_event_setting.asp
' Discription : 출석 체크 이벤트 설정 등록 창
' History : 2023.08.01 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/classes/event/attendanceEventCls.asp"-->
<%
dim mode, main_image, oAttendanceEvent, mo_main_image, mo_main_image2
dim evt_code : evt_code = request("evt_code")
dim idx : idx = request("idx")

'// 이미지 업로드용
Dim userid, encUsrId, tmpTx, tmpRn
    userid = session("ssBctId")

    Randomize()
    tmpTx = split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z",",")
    tmpRn = tmpTx(int(Rnd*26))
    tmpRn = tmpRn & tmpTx(int(Rnd*26))
    encUsrId = tenEnc(tmpRn & userid)
'// 이미지 업로드용

set oAttendanceEvent = new AttendanceEventCls
oAttendanceEvent.FrectEvt_Code = evt_code
oAttendanceEvent.getOneContents()
main_image = oAttendanceEvent.Fmain_image
mo_main_image = oAttendanceEvent.Fmo_main_image
mo_main_image2 = oAttendanceEvent.Fmo_main_image2
if Cint(oAttendanceEvent.Fidx) > 0 then 
    mode = "Modify"
else
    mode = "Add"
end if
if main_image = "" then main_image = "http://webimage.10x10.co.kr/appmanage/startupbanner/SB20230803113300.png"
if oAttendanceEvent.Fbutton_before_day_color = "" then oAttendanceEvent.Fbutton_before_day_color = "#FFA800"
if oAttendanceEvent.Fbutton_before_point_color = "" then oAttendanceEvent.Fbutton_before_point_color = "#00BE9C"
if oAttendanceEvent.Fbutton_before_bg_color = "" then oAttendanceEvent.Fbutton_before_bg_color = "#FEFFB5"
if oAttendanceEvent.Fbutton_after_day_color = "" then oAttendanceEvent.Fbutton_after_day_color = "#FEFFB5"
if oAttendanceEvent.Fbutton_after_point_color = "" then oAttendanceEvent.Fbutton_after_point_color = "#FFF"
if oAttendanceEvent.Fbutton_after_bg_color = "" then oAttendanceEvent.Fbutton_after_bg_color = "#00B49E"
if oAttendanceEvent.Fbutton_today_ring_color = "" then oAttendanceEvent.Fbutton_today_ring_color = "#FFC224"
if oAttendanceEvent.Fcheck_area_bg_color = "" then oAttendanceEvent.Fcheck_area_bg_color = "#FFAAB9"
if oAttendanceEvent.Fcheck_title_color = "" then oAttendanceEvent.Fcheck_title_color = "#111"
if oAttendanceEvent.Fcheck_button_bg_color = "" then oAttendanceEvent.Fcheck_button_bg_color = "#00B49E"
if oAttendanceEvent.Fcheck_button_title_color = "" then oAttendanceEvent.Fcheck_button_title_color = "#FFF"
if oAttendanceEvent.Fcheck_etc_contents_color = "" then oAttendanceEvent.Fcheck_etc_contents_color = "#111"
if oAttendanceEvent.Falarm_bg_color = "" then oAttendanceEvent.Falarm_bg_color = "#FF98AA"

if oAttendanceEvent.Fcheck_etc_contents = "" then
oAttendanceEvent.Fcheck_etc_contents = "지급된 마일리지는 스페셜 마일리지로," & vbcrlf
oAttendanceEvent.Fcheck_etc_contents = oAttendanceEvent.Fcheck_etc_contents + "2023년 8월 21일 23:59까지 사용 가능합니다."
oAttendanceEvent.Fcheck_etc_contents = oAttendanceEvent.Fcheck_etc_contents + "미사용 시 자동으로 소멸됩니다."
end if

if oAttendanceEvent.Falarm_etc_contents = "" then
oAttendanceEvent.Falarm_etc_contents = "본 이벤트는 텐바이텐 APP에서 하루에 한 번씩, 기간 내 최대 9일 간 참여 가능합니다." & vbcrlf
oAttendanceEvent.Falarm_etc_contents = oAttendanceEvent.Falarm_etc_contents + "이벤트가 종료되면 더 이상 참여가 불가합니다."
oAttendanceEvent.Falarm_etc_contents = oAttendanceEvent.Falarm_etc_contents + "지급된 마일리지는 스페셜 마일리지로, 2023년 8월 21일 23:59까지 사용 가능합니다. 미사용 시 자동으로 소멸됩니다."
oAttendanceEvent.Falarm_etc_contents = oAttendanceEvent.Falarm_etc_contents + "해당 이벤트는 내부 사정으로 인해 별도 공지 없이 이벤트가 조기 종료될 수 있습니다."
end if
%>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<script>
function jsEvtSubmit(frm){
    // if(frm.main_image.value==""){
    //     alert("메인 이미지를 등록해주세요.");
    //     return false;
    // }
    if(frm.alarm_etc_contents.value==""){
        alert("유의사항을 등록해주세요.");
        return false;
    }
    frm.action="attendance_process.asp";
	frm.submit();
}
// 업로드 파일 확인 및 처리
function jsCheckUpload() {
	if($("#fileupload").val()!="") {
		$("#fileupmode").val("upload");

		$('#ajaxform').ajaxSubmit({
			//보내기전 validation check가 필요할경우
			beforeSubmit: function (data, frm, opt) {
				if(!(/\.(jpg|jpeg|gif|png)$/i).test(frm[0].upfile.value)) {
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
					document.frmEvt.main_image.value=resultObj.fileurl;
					$("#filepre").val(resultObj.fileurl);
					$("#lyrBnrImg").hide().attr("src",$("#filepre").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
					$("#lyrImgUpBtn").hide();
					$("#lyrImgDelBtn").show();
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
function jsDelImg(){
	if(confirm("이미지를 삭제하시겠습니까?\n\n※ 파일이 완전히 삭제되며 복구 할 수 없습니다.")){
		if($("#filepre").val()!="") {
			$("#fileupmode").val("delete");

			$('#ajaxform').ajaxSubmit({
				//보내기전
				beforeSubmit: function (data, frm, opt) {
					$("#lyrPrgs").show();
				},
				//submit이후의 처리
				success: function(responseText, statusText){
					var resultObj = JSON.parse(responseText)

					if(resultObj.response=="fail") {
						alert(resultObj.faildesc);
					} else if(resultObj.response=="ok") {
						document.frmEvt.main_image.value="";
						$("#lyrBnrImg").hide().attr("src","/images/admin_login_logo2.png").fadeIn("fast");
						$("#filepre").val("");
						$("#lyrImgUpBtn").show();
						$("#lyrImgDelBtn").hide();
					} else {
						alert("처리중 오류가 발생했습니다.\n" + responseText);
					}
					$("#lyrPrgs").hide();
				},
				//ajax error
				error: function(err){
					alert("ERR: " + err.responseText);
					$("#lyrPrgs").hide();
				}
			});
		}
	}
}
</script>
<form name="frmEvt" method="post" style="margin:0px;">
<input type="hidden" name="evt_code" value="<%=evt_code%>">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%=oAttendanceEvent.Fidx%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>출석체크 이벤트 설정</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
				<tr>
					<th>이벤트명</th>
					<td>
						<%=oAttendanceEvent.Fevt_name%>
          			</td>
				</tr>
				<tr>
					<th>메인이미지</th>
					<td>
						<input type="hidden" name="main_image" value="<%=main_image%>" />
                        <div style="width:220px; height:220px;">
                            <div id="lyrPrgs" style="display:none; position:absolute;padding:101px; background-color:rgba(0,0,0,0.2);"><img src="http://fiximage.10x10.co.kr/web2015/giftcard/ajax_loader.gif" alt="progress" /></div>
                            <img id="lyrBnrImg" src="<%=chkIIF(main_image="" or isNull(main_image),"/images/admin_login_logo2.png",main_image)%>" style="height:218px; border:1px solid #EEE;"/>
                        </div>
                        <div id="lyrImgDelBtn" class="btn" style="<%=chkIIF(idx = 0 or main_image="" or IsNull(main_image),"display:none;","")%>" onclick="jsDelImg();">이미지 삭제</button></div>
                        <div id="lyrImgUpBtn" class="btn" style="<%=chkIIF(idx = 0 or main_image="" or IsNull(main_image),"","display:none;")%>"><label for="fileupload">이미지 업로드</label></div>
          			</td>
				</tr>
				<tr>
					<th>메인 이미지 링크</th>
					<td>
						<input type="text" class="formControl formControl550" placeholder="입력하세요." name="main_image_link" id="main_image_link" maxlength="128" value="<%=oAttendanceEvent.Fmain_image_link%>">
					</td>
				</tr>
				<tr>
					<th>출석전 버튼 날짜 컬러</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="입력하세요." name="button_before_day_color" id="button_before_day_color" maxlength="10" value="<%=oAttendanceEvent.Fbutton_before_day_color%>">
					</td>
				</tr>
				<tr>
					<th>출석전 버튼 포인트 컬러</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="입력하세요." name="button_before_point_color" id="button_before_point_color" maxlength="10" value="<%=oAttendanceEvent.Fbutton_before_point_color%>">
					</td>
				</tr>
				<tr>
					<th>출석전 버튼 배경 컬러</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="입력하세요." name="button_before_bg_color" id="button_before_bg_color" maxlength="10" value="<%=oAttendanceEvent.Fbutton_before_bg_color%>">
					</td>
				</tr>
				<tr>
					<th>출석후 버튼 날짜 컬러</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="입력하세요." name="button_after_day_color" id="button_after_day_color" maxlength="10" value="<%=oAttendanceEvent.Fbutton_after_day_color%>">
					</td>
				</tr>
                <tr>
					<th>출석후 버튼 포인트 컬러</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="입력하세요." name="button_after_point_color" id="button_after_point_color" maxlength="10" value="<%=oAttendanceEvent.Fbutton_after_point_color%>">
					</td>
				</tr>
                <tr>
					<th>출석후 버튼 배경 컬러</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="입력하세요." name="button_after_bg_color" id="button_after_bg_color" maxlength="10" value="<%=oAttendanceEvent.Fbutton_after_bg_color%>">
					</td>
				</tr>
                <tr>
					<th>출석 버튼 today 링 컬러</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="입력하세요." name="button_today_ring_color" id="button_today_ring_color" maxlength="10" value="<%=oAttendanceEvent.Fbutton_today_ring_color%>">
					</td>
				</tr>
                <tr>
					<th>하단 출석 영역 배경 컬러</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="입력하세요." name="check_area_bg_color" id="check_area_bg_color" maxlength="10" value="<%=oAttendanceEvent.Fcheck_area_bg_color%>">
					</td>
				</tr>
                <tr>
					<th>하단 출석 영역 타이틀 컬러</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="입력하세요." name="check_title_color" id="check_title_color" maxlength="10" value="<%=oAttendanceEvent.Fcheck_title_color%>">
					</td>
				</tr>
                <tr>
					<th>하단 출석 영역 버튼 배경 컬러</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="입력하세요." name="check_button_bg_color" id="check_button_bg_color" maxlength="10" value="<%=oAttendanceEvent.Fcheck_button_bg_color%>">
					</td>
				</tr>
                <tr>
					<th>하단 출석 영역 버튼 타이틀 컬러</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="입력하세요." name="check_button_title_color" id="check_button_title_color" maxlength="10" value="<%=oAttendanceEvent.Fcheck_button_title_color%>">
					</td>
				</tr>
                <tr>
					<th>하단 출석 영역 안내 문구</th>
					<td>
                        <textarea name="check_etc_contents" rows="3" cols="50" placeholder="안내 문구를 입력해주세요."><%=oAttendanceEvent.Fcheck_etc_contents%></textarea>
					</td>
				</tr>
                <tr>
					<th>하단 출석 영역 안내 문구 컬러</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="입력하세요." name="check_etc_contents_color" id="check_etc_contents_color" maxlength="10" value="<%=oAttendanceEvent.Fcheck_etc_contents_color%>">
					</td>
				</tr>
				<tr>
					<th>알림 신청 영역 배경 컬러</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="입력하세요." name="alarm_bg_color" id="alarm_bg_color" maxlength="128" value="<%=oAttendanceEvent.Falarm_bg_color%>">
					</td>
				</tr>
                <tr>
                    <th>알림 신청 영역 유의사항</th>
                    <td>
                        <textarea name="alarm_etc_contents" rows="3" cols="50" placeholder="기타 공지를 입력해주세요."><%=oAttendanceEvent.Falarm_etc_contents%></textarea>
                    </td>
                </tr>
                <tr>
					<th>팝업 말풍선 배경 컬러</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="입력하세요." name="popup_bubble_bg_color" id="popup_bubble_bg_color" maxlength="128" value="<%=oAttendanceEvent.Fpopup_bubble_bg_color%>">
					</td>
				</tr>
                <tr>
					<th>팝업 말풍선 텍스트 컬러</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="입력하세요." name="popup_bubble_text_color" id="popup_bubble_text_color" maxlength="128" value="<%=oAttendanceEvent.Fpopup_bubble_text_color%>">
					</td>
				</tr>
				<tr>
					<th>마일리지 적요 내용</th>
					<td>
						<input type="text" class="formControl formControl550" placeholder="입력하세요." name="mileage_summary" id="mileage_summary" maxlength="128" value="<%=oAttendanceEvent.Fmileage_summary%>">
					</td>
				</tr>
			</tbody>
		</table>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">취소</button>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">저장</button>
	</div>
</div>
</form>
<!-- 이미지 업로드 Form -->
<form name="frmUpload" id="ajaxform" action="<%=uploadImgUrl%>/linkweb/common/simpleCommonImgUploadProc.asp" method="post" enctype="multipart/form-data" style="display:none; height:0px;width:0px;">
<input type="file" name="upfile" id="fileupload" onchange="jsCheckUpload();" accept="image/*" />
<input type="hidden" name="mode" id="fileupmode" value="upload">
<input type="hidden" name="div" value="SB">
<input type="hidden" name="upPath" value="/appmanage/startupbanner/">
<input type="hidden" name="tuid" value="<%=encUsrId%>">
<input type="hidden" name="prefile" id="filepre" value="<%=main_image%>">
</form>
<% set oAttendanceEvent = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->