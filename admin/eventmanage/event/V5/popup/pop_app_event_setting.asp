<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_app_event_setting.asp
' Discription : ������ �̺�Ʈ ���� ��� â
' History : 2022.01.30 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/classes/event/appDedicatedEventCls.asp"-->
<%
dim mode, bannerImg, oAppEvent, noticeText, oAppDedicated, episodeCount, bannerImg2, bannerImg3
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

set oAppEvent = new AppEventCls
    oAppEvent.FrectEvt_Code = evt_code
    oAppEvent.getOneContents()
bannerImg = oAppEvent.FOneItem.FmainImage
bannerImg2 = oAppEvent.FOneItem.FmainImage2
bannerImg3 = oAppEvent.FOneItem.FmoMainImage

if bannerImg = "" then bannerImg = "http://webimage.10x10.co.kr/appmanage/startupbanner/SB20230613143329.jpg"
if bannerImg2 = "" then bannerImg2 = "http://webimage.10x10.co.kr/appmanage/startupbanner/SB20230613143336.jpg"
if bannerImg3 = "" then bannerImg3 = "http://webimage.10x10.co.kr/appmanage/startupbanner/SB20230613161323.jpg"

if oAppEvent.FOneItem.Ftitle_color = "" then oAppEvent.FOneItem.Ftitle_color = "#1a56c9"
if oAppEvent.FOneItem.Fprize_circle_color = "" then oAppEvent.FOneItem.Fprize_circle_color = "#f35d1b"
if oAppEvent.FOneItem.Fprize_circle_color2 = "" then oAppEvent.FOneItem.Fprize_circle_color2 = "#ff9364"
if oAppEvent.FOneItem.Fitemlist_bg_color = "" then oAppEvent.FOneItem.Fitemlist_bg_color = "#1a56c9"
if oAppEvent.FOneItem.Fbutton_color = "" then oAppEvent.FOneItem.Fbutton_color = "#f36628"
if oAppEvent.FOneItem.Fprize_bg_color = "" then oAppEvent.FOneItem.Fprize_bg_color = "#f35d1b"
if oAppEvent.FOneItem.Fsub_title = "" then oAppEvent.FOneItem.Fsub_title = "��÷�ڴ� �̺�Ʈ�� ����� ��������<br/> ȭ/��/�� �� 12�ø��� �����˴ϴ�."
if oAppEvent.FOneItem.FetcNotice = "" then
oAppEvent.FOneItem.FetcNotice = "��÷�ڿ��Դ� ���� ���� �����Դϴ�." & vbcrlf
oAppEvent.FOneItem.FetcNotice = oAppEvent.FOneItem.FetcNotice + "5���� �ʰ� ��ǰ�� ���, �������� Ȯ�� �� �߼۵˴ϴ�."
end if

if oAppEvent.FOneItem.Fidx > 0 then 
    mode = "Modify"
else
    mode = "Add"
end if

if oAppEvent.FOneItem.Fnotice <> "" then
    noticeText=oAppEvent.FOneItem.Fnotice
else
    noticeText = "-�ٹ����� �ۿ��� ȸ���� ���� �����ϸ�, ID�� 1�� 1ȸ ���������մϴ�." & vbcrlf
    noticeText = noticeText & "-��÷�ڴ� ����� ���ڿ� �̺�Ʈ ������ ������ ��ǥ�Ǹ�, ��÷�� �е鲲�� ���� ���� �����Դϴ�." & vbcrlf
    noticeText = noticeText & "-��ǰ�� ��÷�� �в��� �����Ű� �ʿ��� ���������� ��û�� �� ������, ������������ �ٹ����� �δ��Դϴ�." & vbcrlf
    noticeText = noticeText & "-�̺�Ʈ ���� �� �������� Ȯ���� �ȵ� ��� ��÷�� ��ҵǸ�, ��÷�� �ٸ� �п��� �絵�˴ϴ�." & vbcrlf
    noticeText = noticeText & "-��÷�� ���в��� �ı⸦ ��û�� �����Դϴ�."
end if

set oAppDedicated = new AppEventCls
oAppDedicated.FRectEventCode = evt_code
episodeCount = oAppDedicated.fnGetAppDedicatedCount
set oAppDedicated = nothing
%>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<script>
function jsEvtSubmit(frm){
    if(frm.bannerImg.value==""){
        alert("���� �̹����� ������ּ���.");
        return false;
    }
    if(frm.notice.value==""){
        alert("���ǻ����� ������ּ���.");
        return false;
    }
    frm.action="appDedicated_process.asp";
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
					document.frmEvt.bannerImg.value=resultObj.fileurl;
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
						document.frmEvt.bannerImg.value="";
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
// ���ε� ���� Ȯ�� �� ó��
function jsCheckUpload2() {
	if($("#fileupload2").val()!="") {
		$("#fileupmode2").val("upload");

		$('#ajaxform2').ajaxSubmit({
			//�������� validation check�� �ʿ��Ұ��
			beforeSubmit: function (data, frm, opt) {
				if(!(/\.(jpg|jpeg|gif|png)$/i).test(frm[0].upfile.value)) {
					alert("JPG,PNG �̹������ϸ� ���ε� �Ͻ� �� �ֽ��ϴ�.");
					$("#fileupload2").val("");
					return false;
				}
				$("#lyrPrgs2").show();
			},
			//submit������ ó��
			success: function(responseText, statusText){
				var resultObj = JSON.parse(responseText)

				if(resultObj.response=="fail") {
					alert(resultObj.faildesc);
				} else if(resultObj.response=="ok") {
					document.frmEvt.bannerImg2.value=resultObj.fileurl;
					$("#filepre2").val(resultObj.fileurl);
					$("#lyrBnrImg2").hide().attr("src",$("#filepre2").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
					$("#lyrImgUpBtn2").hide();
					$("#lyrImgDelBtn2").show();
				} else {
					alert("ó���� ������ �߻��߽��ϴ�.\n" + responseText);
				}
				$("#fileupload2").val("");
				$("#lyrPrgs2").hide();
			},
			//ajax error
			error: function(err){
				alert("ERR: " + err.responseText);
				$("#fileupload2").val("");
				$("#lyrPrgs2").hide();
			}
		});
	}
}
// �������� ���� ���� ó��
function jsDelImg2(){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n�� ������ ������ �����Ǹ� ���� �� �� �����ϴ�.")){
		if($("#filepre2").val()!="") {
			$("#fileupmode2").val("delete");

			$('#ajaxform2').ajaxSubmit({
				//��������
				beforeSubmit: function (data, frm, opt) {
					$("#lyrPrgs2").show();
				},
				//submit������ ó��
				success: function(responseText, statusText){
					var resultObj = JSON.parse(responseText)

					if(resultObj.response=="fail") {
						alert(resultObj.faildesc);
					} else if(resultObj.response=="ok") {
						document.frmEvt.bannerImg.value="";
						$("#lyrBnrImg2").hide().attr("src","/images/admin_login_logo2.png").fadeIn("fast");
						$("#filepre2").val("");
						$("#lyrImgUpBtn2").show();
						$("#lyrImgDelBtn2").hide();
					} else {
						alert("ó���� ������ �߻��߽��ϴ�.\n" + responseText);
					}
					$("#lyrPrgs2").hide();
				},
				//ajax error
				error: function(err){
					alert("ERR: " + err.responseText);
					$("#lyrPrgs2").hide();
				}
			});
		}
	}
}
// ���ε� ���� Ȯ�� �� ó��
function jsCheckUpload3() {
	if($("#fileupload3").val()!="") {
		$("#fileupmode3").val("upload");

		$('#ajaxform3').ajaxSubmit({
			//�������� validation check�� �ʿ��Ұ��
			beforeSubmit: function (data, frm, opt) {
				if(!(/\.(jpg|jpeg|gif|png)$/i).test(frm[0].upfile.value)) {
					alert("JPG,PNG �̹������ϸ� ���ε� �Ͻ� �� �ֽ��ϴ�.");
					$("#fileupload3").val("");
					return false;
				}
				$("#lyrPrgs3").show();
			},
			//submit������ ó��
			success: function(responseText, statusText){
				var resultObj = JSON.parse(responseText)

				if(resultObj.response=="fail") {
					alert(resultObj.faildesc);
				} else if(resultObj.response=="ok") {
					document.frmEvt.bannerImg3.value=resultObj.fileurl;
					$("#filepre3").val(resultObj.fileurl);
					$("#lyrBnrImg3").hide().attr("src",$("#filepre3").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
					$("#lyrImgUpBtn3").hide();
					$("#lyrImgDelBtn3").show();
				} else {
					alert("ó���� ������ �߻��߽��ϴ�.\n" + responseText);
				}
				$("#fileupload3").val("");
				$("#lyrPrgs3").hide();
			},
			//ajax error
			error: function(err){
				alert("ERR: " + err.responseText);
				$("#fileupload3").val("");
				$("#lyrPrgs3").hide();
			}
		});
	}
}
// �������� ���� ���� ó��
function jsDelImg3(){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n�� ������ ������ �����Ǹ� ���� �� �� �����ϴ�.")){
		if($("#filepre3").val()!="") {
			$("#fileupmode3").val("delete");

			$('#ajaxform3').ajaxSubmit({
				//��������
				beforeSubmit: function (data, frm, opt) {
					$("#lyrPrgs3").show();
				},
				//submit������ ó��
				success: function(responseText, statusText){
					var resultObj = JSON.parse(responseText)

					if(resultObj.response=="fail") {
						alert(resultObj.faildesc);
					} else if(resultObj.response=="ok") {
						document.frmEvt.bannerImg.value="";
						$("#lyrBnrImg3").hide().attr("src","/images/admin_login_logo2.png").fadeIn("fast");
						$("#filepre3").val("");
						$("#lyrImgUpBtn3").show();
						$("#lyrImgDelBtn3").hide();
					} else {
						alert("ó���� ������ �߻��߽��ϴ�.\n" + responseText);
					}
					$("#lyrPrgs3").hide();
				},
				//ajax error
				error: function(err){
					alert("ERR: " + err.responseText);
					$("#lyrPrgs3").hide();
				}
			});
		}
	}
}
function TnSearchObjOpenWin(){
    var winpop = window.open('/admin/eventmanage/event/v5/popup/pop_app_event_item_regist.asp?evt_code=<%=evt_code%>&stype=w','winpop','width=1300,height=768,scrollbars=yes,resizable=yes');
}
</script>
<form name="frmEvt" method="post" style="margin:0px;">
<input type="hidden" name="evt_code" value="<%=evt_code%>">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%=oAppEvent.FOneItem.Fidx%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>������ ������ ����</h1>
		<p class="tMar15 cPk2 fs12">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;������ ������ ��ǰ ��� �� ��ǰ �̹��� ����� ���������� ��� �ؾ��մϴ�.</p>
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
						<%=oAppEvent.FOneItem.Fevt_name%>
          			</td>
				</tr>
				<tr>
					<th>�����̹���</th>
					<td>
						<input type="hidden" name="bannerImg" value="<%=bannerImg%>" />
                        <div style="width:220px; height:220px;">
                            <div id="lyrPrgs" style="display:none; position:absolute;padding:101px; background-color:rgba(0,0,0,0.2);"><img src="http://fiximage.10x10.co.kr/web2015/giftcard/ajax_loader.gif" alt="progress" /></div>
                            <img id="lyrBnrImg" src="<%=chkIIF(bannerImg="" or isNull(bannerImg),"/images/admin_login_logo2.png",bannerImg)%>" style="height:218px; border:1px solid #EEE;"/>
                        </div>
                        <div id="lyrImgDelBtn" class="btn" style="<%=chkIIF(idx = 0 or bannerImg="" or IsNull(bannerImg),"display:none;","")%>" onclick="jsDelImg();">�̹��� ����</button></div>
                        <div id="lyrImgUpBtn" class="btn" style="<%=chkIIF(idx = 0 or bannerImg="" or IsNull(bannerImg),"","display:none;")%>"><label for="fileupload">�̹��� ���ε�</label></div>
          			</td>
				</tr>
				<tr>
					<th>������ �����̹���</th>
					<td>
						<input type="hidden" name="bannerImg2" value="<%=bannerImg2%>" />
                        <div style="width:220px; height:220px;">
                            <div id="lyrPrgs2" style="display:none; position:absolute;padding:101px; background-color:rgba(0,0,0,0.2);"><img src="http://fiximage.10x10.co.kr/web2015/giftcard/ajax_loader.gif" alt="progress" /></div>
                            <img id="lyrBnrImg2" src="<%=chkIIF(bannerImg2="" or isNull(bannerImg2),"/images/admin_login_logo2.png",bannerImg2)%>" style="height:218px; border:1px solid #EEE;"/>
                        </div>
                        <div id="lyrImgDelBtn2" class="btn" style="<%=chkIIF(idx = 0 or bannerImg2="" or IsNull(bannerImg2),"display:none;","")%>" onclick="jsDelImg2();">�̹��� ����</button></div>
                        <div id="lyrImgUpBtn2" class="btn" style="<%=chkIIF(idx = 0 or bannerImg2="" or IsNull(bannerImg2),"","display:none;")%>"><label for="fileupload2">�̹��� ���ε�</label></div>
          			</td>
				</tr>
				<tr>
					<th>���� �����̹���</th>
					<td>
						<input type="hidden" name="bannerImg3" value="<%=bannerImg3%>" />
                        <div style="width:220px; height:220px;">
                            <div id="lyrPrgs3" style="display:none; position:absolute;padding:101px; background-color:rgba(0,0,0,0.2);"><img src="http://fiximage.10x10.co.kr/web2015/giftcard/ajax_loader.gif" alt="progress" /></div>
                            <img id="lyrBnrImg3" src="<%=chkIIF(bannerImg3="" or isNull(bannerImg3),"/images/admin_login_logo2.png",bannerImg3)%>" style="height:218px; border:1px solid #EEE;"/>
                        </div>
                        <div id="lyrImgDelBtn3" class="btn" style="<%=chkIIF(idx = 0 or bannerImg3="" or IsNull(bannerImg3),"display:none;","")%>" onclick="jsDelImg3();">�̹��� ����</button></div>
                        <div id="lyrImgUpBtn3" class="btn" style="<%=chkIIF(idx = 0 or bannerImg3="" or IsNull(bannerImg3),"","display:none;")%>"><label for="fileupload3">�̹��� ���ε�</label></div>
          			</td>
				</tr>
				<tr>
					<th>���� �� �ٷΰ��� ��ư ��ũ</th>
					<td>
						<input type="text" class="formControl formControl550" placeholder="�Է��ϼ���." name="deeplink" id="deeplink" maxlength="128" value="<%=oAppEvent.FOneItem.Fdeeplink%>">
					</td>
				</tr>
                <tr>
                    <th>ȸ��(��ǰ) ����</th>
                    <td>
                        <input type="button" class="button" value="ȸ�� ����(<%=episodeCount%>)" onclick="TnSearchObjOpenWin();">
                    </td>
                </tr>
				<tr>
					<th>Ÿ��Ʋ �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="title_color" id="title_color" maxlength="10" value="<%=oAppEvent.FOneItem.Ftitle_color%>">
					</td>
				</tr>
				<tr>
					<th>��÷�� �� ���� ��� �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="prize_circle_color" id="prize_circle_color" maxlength="10" value="<%=oAppEvent.FOneItem.Fprize_circle_color%>">
					</td>
				</tr>
				<tr>
					<th>��÷�� �� ��Ʈ��ũ ��� �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="prize_circle_color2" id="prize_circle_color2" maxlength="10" value="<%=oAppEvent.FOneItem.Fprize_circle_color2%>">
					</td>
				</tr>
				<tr>
					<th>����Ư�� ����Ʈ ��� �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="itemlist_bg_color" id="itemlist_bg_color" maxlength="10" value="<%=oAppEvent.FOneItem.Fitemlist_bg_color%>">
					</td>
				</tr>
				<tr>
					<th>�˸� ��ư �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="button_color" id="button_color" maxlength="10" value="<%=oAppEvent.FOneItem.Fbutton_color%>">
					</td>
				</tr>
				<tr>
					<th>��÷�� ����Ʈ ��� �÷�</th>
					<td>
						<input type="text" class="formControl formControl150" placeholder="�Է��ϼ���." name="prize_bg_color" id="prize_bg_color" maxlength="10" value="<%=oAppEvent.FOneItem.Fprize_bg_color%>">
					</td>
				</tr>
				<tr>
					<th>��÷�� �ȳ� ���� Ÿ��Ʋ</th>
					<td>
						<input type="text" class="formControl formControl550" placeholder="�Է��ϼ���." name="sub_title" id="sub_title" maxlength="128" value="<%=oAppEvent.FOneItem.Fsub_title%>">
					</td>
				</tr>
                <tr>
                    <th>��÷�� �ȳ� ��Ÿ ����</th>
                    <td>
                        <textarea name="etcNotice" rows="3" cols="50" placeholder="��Ÿ ������ �Է����ּ���."><%=oAppEvent.FOneItem.FetcNotice%></textarea>
                    </td>
                </tr>
                <tr>
                    <th>���ǻ���</th>
                    <td>
                        <textarea name="notice" rows="8" cols="50" placeholder="�귣�� ���丮�� �Է����ּ���."><%=noticeText%></textarea>
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
<input type="hidden" name="prefile" id="filepre" value="<%=bannerImg%>">
</form>
<form name="frmUpload2" id="ajaxform2" action="<%=uploadImgUrl%>/linkweb/common/simpleCommonImgUploadProc.asp" method="post" enctype="multipart/form-data" style="display:none; height:0px;width:0px;">
<input type="file" name="upfile" id="fileupload2" onchange="jsCheckUpload2();" accept="image/*" />
<input type="hidden" name="mode" id="fileupmode2" value="upload">
<input type="hidden" name="div" value="SB">
<input type="hidden" name="upPath" value="/appmanage/startupbanner/">
<input type="hidden" name="tuid" value="<%=encUsrId%>">
<input type="hidden" name="prefile2" id="filepre2" value="<%=bannerImg2%>">
</form>
<form name="frmUpload3" id="ajaxform3" action="<%=uploadImgUrl%>/linkweb/common/simpleCommonImgUploadProc.asp" method="post" enctype="multipart/form-data" style="display:none; height:0px;width:0px;">
<input type="file" name="upfile" id="fileupload3" onchange="jsCheckUpload3();" accept="image/*" />
<input type="hidden" name="mode" id="fileupmode3" value="upload">
<input type="hidden" name="div" value="SB">
<input type="hidden" name="upPath" value="/appmanage/startupbanner/">
<input type="hidden" name="tuid" value="<%=encUsrId%>">
<input type="hidden" name="prefile3" id="filepre3" value="<%=bannerImg3%>">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->