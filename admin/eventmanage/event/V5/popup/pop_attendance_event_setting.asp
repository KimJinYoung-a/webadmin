<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_attendance_event_setting.asp
' Discription : �⼮ üũ �̺�Ʈ ���� ��� â
' History : 2023.08.01 ������
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

'// �̹��� ���ε��
Dim userid, encUsrId, tmpTx, tmpRn
    userid = session("ssBctId")

    Randomize()
    tmpTx = split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z",",")
    tmpRn = tmpTx(int(Rnd*26))
    tmpRn = tmpRn & tmpTx(int(Rnd*26))
    encUsrId = tenEnc(tmpRn & userid)
'// �̹��� ���ε��

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
oAttendanceEvent.Fcheck_etc_contents = "���޵� ���ϸ����� ����� ���ϸ�����," & vbcrlf
oAttendanceEvent.Fcheck_etc_contents = oAttendanceEvent.Fcheck_etc_contents + "2023�� 8�� 21�� 23:59���� ��� �����մϴ�."
oAttendanceEvent.Fcheck_etc_contents = oAttendanceEvent.Fcheck_etc_contents + "�̻�� �� �ڵ����� �Ҹ�˴ϴ�."
end if

if oAttendanceEvent.Falarm_etc_contents = "" then
oAttendanceEvent.Falarm_etc_contents = "�� �̺�Ʈ�� �ٹ����� APP���� �Ϸ翡 �� ����, �Ⱓ �� �ִ� 9�� �� ���� �����մϴ�." & vbcrlf
oAttendanceEvent.Falarm_etc_contents = oAttendanceEvent.Falarm_etc_contents + "�̺�Ʈ�� ����Ǹ� �� �̻� ������ �Ұ��մϴ�."
oAttendanceEvent.Falarm_etc_contents = oAttendanceEvent.Falarm_etc_contents + "���޵� ���ϸ����� ����� ���ϸ�����, 2023�� 8�� 21�� 23:59���� ��� �����մϴ�. �̻�� �� �ڵ����� �Ҹ�˴ϴ�."
oAttendanceEvent.Falarm_etc_contents = oAttendanceEvent.Falarm_etc_contents + "�ش� �̺�Ʈ�� ���� �������� ���� ���� ���� ���� �̺�Ʈ�� ���� ����� �� �ֽ��ϴ�."
end if
%>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<script>
function jsEvtSubmit(frm){
    // if(frm.main_image.value==""){
    //     alert("���� �̹����� ������ּ���.");
    //     return false;
    // }
    if(frm.alarm_etc_contents.value==""){
        alert("���ǻ����� ������ּ���.");
        return false;
    }
    frm.action="attendance_process.asp";
	frm.submit();
}
// ���ε� ���� Ȯ�� �� ó��
function jsCheckUpload() {
	if($("#fileupload").val()!="") {
		$("#fileupmode").val("upload");

		$('#ajaxform').ajaxSubmit({
			//�������� validation check�� �ʿ��Ұ��
			beforeSubmit: function (data, frm, opt) {
				if(!(/\.(jpg|jpeg|gif|png)$/i).test(frm[0].upfile.value)) {
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
					document.frmEvt.main_image.value=resultObj.fileurl;
					$("#filepre").val(resultObj.fileurl);
					$("#lyrBnrImg").hide().attr("src",$("#filepre").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
					$("#lyrImgUpBtn").hide();
					$("#lyrImgDelBtn").show();
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
function jsDelImg(){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n�� ������ ������ �����Ǹ� ���� �� �� �����ϴ�.")){
		if($("#filepre").val()!="") {
			$("#fileupmode").val("delete");

			$('#ajaxform').ajaxSubmit({
				//��������
				beforeSubmit: function (data, frm, opt) {
					$("#lyrPrgs").show();
				},
				//submit������ ó��
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
						alert("ó���� ������ �߻��߽��ϴ�.\n" + responseText);
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
		<h1>�⼮üũ �̺�Ʈ ����</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
				<tr>
					<th>�̺�Ʈ��</th>
					<td>
						<%=oAttendanceEvent.Fevt_name%>
          			</td>
				</tr>
				<tr>
					<th>�����̹���</th>
					<td>
						<input type="hidden" name="main_image" value="<%=main_image%>" />
                        <div style="width:220px; height:220px;">
                            <div id="lyrPrgs" style="display:none; position:absolute;padding:101px; background-color:rgba(0,0,0,0.2);"><img src="http://fiximage.10x10.co.kr/web2015/giftcard/ajax_loader.gif" alt="progress" /></div>
                            <img id="lyrBnrImg" src="<%=chkIIF(main_image="" or isNull(main_image),"/images/admin_login_logo2.png",main_image)%>" style="height:218px; border:1px solid #EEE;"/>
                        </div>
                        <div id="lyrImgDelBtn" class="btn" style="<%=chkIIF(idx = 0 or main_image="" or IsNull(main_image),"display:none;","")%>" onclick="jsDelImg();">�̹��� ����</button></div>
                        <div id="lyrImgUpBtn" class="btn" style="<%=chkIIF(idx = 0 or main_image="" or IsNull(main_image),"","display:none;")%>"><label for="fileupload">�̹��� ���ε�</label></div>
          			</td>
				</tr>
				<tr>
					<th>���� �̹��� ��ũ</th>
					<td>
						<input type="text" class="formControl formControl550" placeholder="�Է��ϼ���." name="main_image_link" id="main_image_link" maxlength="128" value="<%=oAttendanceEvent.Fmain_image_link%>">
					</td>
				</tr>
				<tr>
					<th>�⼮�� ��ư ��¥ �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="button_before_day_color" id="button_before_day_color" maxlength="10" value="<%=oAttendanceEvent.Fbutton_before_day_color%>">
					</td>
				</tr>
				<tr>
					<th>�⼮�� ��ư ����Ʈ �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="button_before_point_color" id="button_before_point_color" maxlength="10" value="<%=oAttendanceEvent.Fbutton_before_point_color%>">
					</td>
				</tr>
				<tr>
					<th>�⼮�� ��ư ��� �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="button_before_bg_color" id="button_before_bg_color" maxlength="10" value="<%=oAttendanceEvent.Fbutton_before_bg_color%>">
					</td>
				</tr>
				<tr>
					<th>�⼮�� ��ư ��¥ �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="button_after_day_color" id="button_after_day_color" maxlength="10" value="<%=oAttendanceEvent.Fbutton_after_day_color%>">
					</td>
				</tr>
                <tr>
					<th>�⼮�� ��ư ����Ʈ �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="button_after_point_color" id="button_after_point_color" maxlength="10" value="<%=oAttendanceEvent.Fbutton_after_point_color%>">
					</td>
				</tr>
                <tr>
					<th>�⼮�� ��ư ��� �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="button_after_bg_color" id="button_after_bg_color" maxlength="10" value="<%=oAttendanceEvent.Fbutton_after_bg_color%>">
					</td>
				</tr>
                <tr>
					<th>�⼮ ��ư today �� �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="button_today_ring_color" id="button_today_ring_color" maxlength="10" value="<%=oAttendanceEvent.Fbutton_today_ring_color%>">
					</td>
				</tr>
                <tr>
					<th>�ϴ� �⼮ ���� ��� �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="check_area_bg_color" id="check_area_bg_color" maxlength="10" value="<%=oAttendanceEvent.Fcheck_area_bg_color%>">
					</td>
				</tr>
                <tr>
					<th>�ϴ� �⼮ ���� Ÿ��Ʋ �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="check_title_color" id="check_title_color" maxlength="10" value="<%=oAttendanceEvent.Fcheck_title_color%>">
					</td>
				</tr>
                <tr>
					<th>�ϴ� �⼮ ���� ��ư ��� �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="check_button_bg_color" id="check_button_bg_color" maxlength="10" value="<%=oAttendanceEvent.Fcheck_button_bg_color%>">
					</td>
				</tr>
                <tr>
					<th>�ϴ� �⼮ ���� ��ư Ÿ��Ʋ �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="check_button_title_color" id="check_button_title_color" maxlength="10" value="<%=oAttendanceEvent.Fcheck_button_title_color%>">
					</td>
				</tr>
                <tr>
					<th>�ϴ� �⼮ ���� �ȳ� ����</th>
					<td>
                        <textarea name="check_etc_contents" rows="3" cols="50" placeholder="�ȳ� ������ �Է����ּ���."><%=oAttendanceEvent.Fcheck_etc_contents%></textarea>
					</td>
				</tr>
                <tr>
					<th>�ϴ� �⼮ ���� �ȳ� ���� �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="check_etc_contents_color" id="check_etc_contents_color" maxlength="10" value="<%=oAttendanceEvent.Fcheck_etc_contents_color%>">
					</td>
				</tr>
				<tr>
					<th>�˸� ��û ���� ��� �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="alarm_bg_color" id="alarm_bg_color" maxlength="128" value="<%=oAttendanceEvent.Falarm_bg_color%>">
					</td>
				</tr>
                <tr>
                    <th>�˸� ��û ���� ���ǻ���</th>
                    <td>
                        <textarea name="alarm_etc_contents" rows="3" cols="50" placeholder="��Ÿ ������ �Է����ּ���."><%=oAttendanceEvent.Falarm_etc_contents%></textarea>
                    </td>
                </tr>
                <tr>
					<th>�˾� ��ǳ�� ��� �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="popup_bubble_bg_color" id="popup_bubble_bg_color" maxlength="128" value="<%=oAttendanceEvent.Fpopup_bubble_bg_color%>">
					</td>
				</tr>
                <tr>
					<th>�˾� ��ǳ�� �ؽ�Ʈ �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="popup_bubble_text_color" id="popup_bubble_text_color" maxlength="128" value="<%=oAttendanceEvent.Fpopup_bubble_text_color%>">
					</td>
				</tr>
				<tr>
					<th>���ϸ��� ���� ����</th>
					<td>
						<input type="text" class="formControl formControl550" placeholder="�Է��ϼ���." name="mileage_summary" id="mileage_summary" maxlength="128" value="<%=oAttendanceEvent.Fmileage_summary%>">
					</td>
				</tr>
			</tbody>
		</table>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">���</button>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">����</button>
	</div>
</div>
</form>
<!-- �̹��� ���ε� Form -->
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