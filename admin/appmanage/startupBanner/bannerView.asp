<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/appmanage/startupBannerCls.asp" -->
<!-- #include virtual="/partner/lib/adminHead.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->

<%
	Dim cSBanner, idx, bannerTitle,startDate,expireDate,closeType,bannerType,bannerImg,linkType,linkTitle,linkURL,targetOS,targetType,importance,isUsing,status
	Dim startDateHour, startDateMinute, startDateSecond, expireDateHour, expireDateMinute, expireDateSecond
	idx	= getNumeric(requestCheckVar(request("idx"),10))

	if idx<>"" then
		SET cSBanner = New CStartupBanner
		cSBanner.FRectIdx = idx
		cSBanner.GetOneStartupBanner

		if cSBanner.FResultCount>0 then
			bannerTitle = cSBanner.FOneItem.FbannerTitle
			startDate   = cSBanner.FOneItem.FstartDate
			expireDate  = cSBanner.FOneItem.FexpireDate
			closeType   = cSBanner.FOneItem.FcloseType
			bannerType  = cSBanner.FOneItem.FbannerType
			bannerImg   = cSBanner.FOneItem.FbannerImg
			linkType    = cSBanner.FOneItem.FlinkType
			linkTitle   = cSBanner.FOneItem.FlinkTitle
			linkURL     = cSBanner.FOneItem.FlinkURL
			targetOS    = cSBanner.FOneItem.FtargetOS
			targetType  = cSBanner.FOneItem.FtargetType
			importance  = cSBanner.FOneItem.Fimportance
			isUsing     = cSBanner.FOneItem.FisUsing
			status      = cSBanner.FOneItem.Fstatus
		end if

		SET cSBanner = Nothing
	end if

	'// ���ϼ��� ó���� ȸ��ID ��ȣȭ
	Dim userid, encUsrId, tmpTx, tmpRn
	userid = session("ssBctId")

	'// ���������� �ð��߰�
	If Trim(startDate) <> "" Then
		If Len(Hour(startDate)) < 2 Then
			startDateHour = "0"&Hour(startDate)
		Else
			startDateHour = Hour(startDate)
		End If
		If Len(Minute(startDate)) < 2 Then
			startDateMinute = "0"&Minute(startDate)
		Else
			startDateMinute = Minute(startDate)
		End If
		If Len(Second(startDate)) < 2 Then
			startDateSecond = "0"&Second(startDate)
		Else
			startDateSecond = Second(startDate)
		End If
	Else
		startDateHour = "00"
		startDateMinute = "00"
		startDateSecond = "00"
	End If

	If Trim(expireDate) <> "" Then
		If Len(Hour(expireDate)) < 2 Then
			expireDateHour = "0"&Hour(expireDate)
		Else
			expireDateHour = Hour(expireDate)
		End If
		If Len(Minute(expireDate)) < 2 Then
			expireDateMinute = "0"&Minute(expireDate)
		Else
			expireDateMinute = Minute(expireDate)
		End If
		If Len(Second(expireDate)) < 2 Then
			expireDateSecond = "0"&Second(expireDate)
		Else
			expireDateSecond = Second(expireDate)
		End If
	Else
		expireDateHour = "23"
		expireDateMinute = "59"
		expireDateSecond = "59"
	End If

	Randomize()
	tmpTx = split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z",",")
	tmpRn = tmpTx(int(Rnd*26))
	tmpRn = tmpRn & tmpTx(int(Rnd*26))
		encUsrId = tenEnc(tmpRn & userid)
%>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
function fnChgLinkType(val) {
	switch(val) {
		case "event":
			document.frm1.linkTitle.value = "�̺�Ʈ";
			document.frm1.linkURL.value = "/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�";
			break;
		case "spevt":
			document.frm1.linkTitle.value = "��ȹ��";
			document.frm1.linkURL.value = "/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�";
			break;
		case "prd":
			document.frm1.linkTitle.value = "��ǰ����";
			document.frm1.linkURL.value = "/category/category_itemprd.asp?itemid=��ǰ�ڵ�";
			break;
		default:
			document.frm1.linkTitle.value = "";
			document.frm1.linkURL.value = "";
	}
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
					document.frm1.bannerImg.value=resultObj.fileurl;
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
						document.frm1.bannerImg.value="";
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

// ������ �ΰ� ���븸 ���� ��
function jsDelImg2(){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
		document.frm1.bannerImg.value="";
		$("#lyrBnrImg").hide().attr("src","/images/admin_login_logo2.png").fadeIn("fast");
		$("#lyrImgUpBtn").show();
		$("#lyrImgDelBtn").hide();
	}
}

// ����� Ȯ�� �� ó��
function fnSubmit(frm) {
	if(frm.startDate.value.length<10) {
		alert("������ �Է����ּ���.");
		frm.startDate.focus();
		return false;
	}

	if(frm.expireDate.value<10) {
		alert("������ �Է����ּ���.");
		frm.expireDate.focus();
		return false;
	}

	if(!frm.closeType.value) {
		alert("�ݱ� �ɼ��� �������ּ���.");
		frm.closeType.focus();
		return false;
	}

	if(!frm.bannerTitle.value) {
		alert("������ �Է����ּ���.");
		frm.bannerTitle.focus();
		return false;
	}

	if(!frm.bannerType[0].checked&&!frm.bannerType[1].checked) {
		alert("��� ���¸� �������ּ���.");
		frm.bannerType[0].focus();
		return false;
	}
/* 
	if(!frm.bannerImg.value) {
		alert("��� �̹����� �������ּ���.");
		return false;
	}
 */
	if(!frm.linkType.value) {
		alert("��ũ ������ �������ּ���.");
		frm.linkType.focus();
		return false;
	}

	if(!frm.linkURL.value.length) {
		alert("��ũ URL�� �Է����ּ���.");
		frm.linkURL.focus();
		return false;
	}

	if(!frm.importance.value.length) {
		alert("����� �߿䵵�� �������ּ���.");
		frm.importance.focus();
		return false;
	}

	if(confirm("�Է��Ͻ� �������� ����Ͻðڽ��ϱ�?")){
		frm.submit();
	}

}
</script>
</head>
<body>
<div class="popupWrap">
	<div class="popHead">
		<h1><img src="/images/partner/pop_admin_logo.gif" alt="10x10" /></h1>
		<p class="btnClose"><input type="image" src="/images/partner/pop_admin_btn_close.gif" alt="â�ݱ�" onclick="window.close();" /></p>
	</div>
	<div class="popContent scrl" style="padding-top:20px;">
		<div class="contTit bgNone">
			<h2>�۱������ ���</h2>
		</div>
		<div class="cont">
			<form name="frm1" action="doBannerReg.asp" method="post" style="margin:0px;">
			<input type="hidden" name="idx" value="<%=idx%>">
			<input type="hidden" name="mode" value="<%=chkiif(idx="" or isNull(idx),"add","modi")%>">
				<table class="tbType1 writeTb" bgcolor="#FFFFFF">
					<tbody>
						<tr>
							<th width="12%">�Ⱓ</th>
							<td height="30" style="padding-left:5px;">
								<input type="text" name="startDate" value="<%=Left(startDate, 10)%>" class="formTxt" id="termSdt" maxlength="10" style="width:100px" placeholder="������" />
								<input type="text" name="startDateSecond" class="formTxt" maxlength="12" style="width:100px" placeholder="�����Ͻú���" value="<%=startDateHour&":"&startDateMinute&":"&startDateSecond%>" />
								<input type="image" src="/images/admin_calendar.png" alt="�޷����� �˻�" id="ChkStart_trigger" onclick="return false;" />
								~
								<input type="text" name="expireDate" value="<%=Left(expireDate, 10)%>" class="formTxt" id="termEdt" maxlength="10" style="width:100px" placeholder="������" />
								<input type="text" name="expireDateSecond" class="formTxt" maxlength="12" style="width:100px" placeholder="�������Ͻú���" value="<%=expireDateHour&":"&expireDateMinute&":"&expireDateSecond%>" />
								<input type="image" src="/images/admin_calendar.png" alt="�޷����� �˻�" id="ChkEnd_trigger" onclick="return false;" />
								<script type="text/javascript">
									var CAL_Start = new Calendar({
										inputField : "termSdt", trigger    : "ChkStart_trigger",
										onSelect: function() {
											var date = Calendar.intToDate(this.selection.get());
											CAL_End.args.min = date;
											CAL_End.redraw();
											this.hide();
										}, bottomBar: true, dateFormat: "%Y-%m-%d"
									});
									var CAL_End = new Calendar({
										inputField : "termEdt", trigger    : "ChkEnd_trigger",
										onSelect: function() {
											var date = Calendar.intToDate(this.selection.get());
											CAL_Start.args.max = date;
											CAL_Start.redraw();
											this.hide();
										}, bottomBar: true, dateFormat: "%Y-%m-%d"
									});
								</script>
							</td>
							<th style="border-left:1px solid #CCC;">�̹���</th>
						</tr>
						<tr>
							<th>�ݱ� �ɼ�</th>
							<td height="30" style="padding-left:5px;">
								<select name="closeType" class="formSlt" >
									<option value="" <%=chkIIF(closeType="" and idx<>"","selected","")%>>:: ���� ::</option>
									<option value="0" <%=chkIIF(closeType="0","selected","")%>>�ܹ߼�</option>
									<option value="1" <%=chkIIF(closeType="1" or idx="","selected","")%>>���� �׸�����</option>
									<option value="2" <%=chkIIF(closeType="2","selected","")%>>7�ϰ� �����ʱ�</option>
									<option value="9" <%=chkIIF(closeType="9","selected","")%>>�ٽ� �����ʱ�</option>
								</select>
							</td>
							<td align="center" style="border-left:1px solid #CCC;" rowspan="7">
								<p>
									<label><input type="radio" name="bannerType" value="S" class="formCheck" <%=chkIIF(bannerType="S","checked","")%> /> ���簢��(560x560px)</label> &nbsp;
									<label><input type="radio" name="bannerType" value="R" class="formCheck" <%=chkIIF(bannerType="R","checked","")%> /> ������(560x800px)</label>
								</p>
								<p class="tMar05">
									<input type="hidden" name="bannerImg" value="<%=bannerImg%>" />
									<div style="width:220px; height:220px;">
										<div id="lyrPrgs" style="display:none; position:absolute;padding:101px; background-color:rgba(0,0,0,0.2);"><img src="http://fiximage.10x10.co.kr/web2015/giftcard/ajax_loader.gif" alt="progress" /></div>
										<img id="lyrBnrImg" src="<%=chkIIF(bannerImg="" or isNull(bannerImg),"/images/admin_login_logo2.png",bannerImg)%>" style="height:218px; border:1px solid #EEE;"/>
									</div>
									<div id="lyrImgDelBtn" class="btn lMar05" style="<%=chkIIF(idx="" and bannerImg="","display:none;","")%>" onclick="jsDelImg();">�̹��� ����</button></div>
									<div id="lyrImgUpBtn" class="btn lMar05" style="<%=chkIIF(idx="" and bannerImg="","","display:none;")%>"><label for="fileupload">�̹��� ���ε�</label></div>
								</p>
							</td>
						</tr>
						<tr>
							<th>����</th>
							<td height="30" style="padding-left:5px;">
								<input type="text" name="bannerTitle" value="<%=bannerTitle%>" class="formTxt" size="50" maxlength="100" />
							</td>
						</tr>
						<tr>
							<th>��ũ</th>
							<td height="30" style="padding-left:5px;">
								<p>
									���� :
									<select name="linkType" class="formSlt" onchange="fnChgLinkType(this.value);">
										<option value="" <%=chkIIF(linkType="","selected","")%>>:: ���� ::</option>
										<option value="event" <%=chkIIF(linkType="event","selected","")%>>�̺�Ʈ</option>
										<option value="spevt" <%=chkIIF(linkType="spevt","selected","")%>>��ȹ��</option>
										<option value="prd" <%=chkIIF(linkType="prd","selected","")%>>��ǰ����</option>
									</select>
									/ ���� : <input type="text" name="linkTitle" value="<%=linkTitle%>" class="formTxt readonly" size="10" maxlength="30" readonly="readonly" />
								</p>
								<p class="tMar05">�ּ� : <input type="text" name="linkURL" value="<%=linkURL%>" class="formTxt" size="60" maxlength="180" /></p>
							</td>
						</tr>
						<tr>
							<th>Ÿ��</th>
							<td height="30" style="padding-left:5px;">
								<p>
									�ü�� :
									<select name="targetOS" class="formSlt">
										<option value="" <%=chkIIF(targetOS="","selected","")%>>��ü</option>
										<option value="ios" <%=chkIIF(targetOS="ios","selected","")%>>iOS</option>
										<option value="android" <%=chkIIF(targetOS="android","selected","")%>>�ȵ���̵�</option>
									</select>
								</p>
								<p class="tMar05">
									��� :
									<select name="targetType" class="formSlt">
										<option value="00" <%=chkIIF(targetType="" or targetType="00","selected","")%>>����</option>
										<option value="30" <%=chkIIF(targetType="30","selected","")%>>��ȸ��</option>
										<option value="15" <%=chkIIF(targetType="15","selected","")%>>Orange</option>
										<option value="10" <%=chkIIF(targetType="10","selected","")%>>Yellow</option>
										<option value="11" <%=chkIIF(targetType="11","selected","")%>>Green</option>
										<option value="12" <%=chkIIF(targetType="12","selected","")%>>Blue</option>
										<option value="13" <%=chkIIF(targetType="13","selected","")%>>VIP Silver</option>
										<option value="14" <%=chkIIF(targetType="14","selected","")%>>VIP Gold</option>
										<option value="16" <%=chkIIF(targetType="16","selected","")%>>VVIP</option>
										<option value="20" <%=chkIIF(targetType="20","selected","")%>>VIP��ü</option>
									</select>
								</p>
							</td>
						</tr>
						<tr>
							<th>�켱����</th>
							<td height="30" style="padding-left:5px;">
								<select name="importance" class="formSlt" >
									<option value="" <%=chkIIF(importance="","selected","")%>>:: ���� ::</option>
									<option value="10" <%=chkIIF(importance="10","selected","")%>>����</option>
									<option value="30" <%=chkIIF(importance="30","selected","")%>>����</option>
									<option value="50" <%=chkIIF(importance="50","selected","")%>>����</option>
								</select>
							</td>
						</tr>
						<tr>
							<th>��뿩��</th>
							<td height="30" style="padding-left:5px;">
								<label><input type="radio" name="isUsing" value="Y" class="formCheck" <%=chkIIF(isUsing="" or isUsing="Y","checked","")%> /> ���</label>
								<label><input type="radio" name="isUsing" value="N" class="formCheck" <%=chkIIF(isUsing="N","checked","")%> /> ������</label>
							</td>
						</tr>
						<tr>
							<th>�������</th>
							<td height="30" style="padding-left:5px;">
								<select name="status" class="formSlt" >
									<option value="0" <%=chkIIF(status="" or status="0","selected","")%>>��ϴ��</option>
									<option value="5" <%=chkIIF(status="5","selected","")%>>����</option>
									<option value="9" <%=chkIIF(status="9","selected","")%>>����</option>
								</select>
							</td>
						</tr>
					</tboby>
				</table>

				<div class="tPad15 ct">
					<input type="button" value="�� ��" onclick="if(confirm('�۾��� ����ϰ� â�� �ݰڽ��ϱ�?')){self.close();}" class="btn3 btnDkGy" style="margin-right:30px;" />
					<input type="button" value="�� ��" onclick="fnSubmit(this.form);" class="btn3 btnRd" />
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
		</div>
	</div>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->