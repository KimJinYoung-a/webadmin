<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��ȹ�� �����̵� ���� �˾�
' History : 2019-02-19 ����ȭ
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/common/lib/pop_slide/classes/slidemanageCls.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<%
dim mastercode , detailcode , mode , idx , itemid , pickitem , menu , device , bannerImg
dim oSlideManage

'// �̹��� ���ε��
Dim userid, encUsrId, tmpTx, tmpRn
	userid = session("ssBctId")

	Randomize()
	tmpTx = split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z",",")
	tmpRn = tmpTx(int(Rnd*26))
	tmpRn = tmpRn & tmpTx(int(Rnd*26))
	encUsrId = tenEnc(tmpRn & userid)
'// �̹��� ���ε��	

menu = requestCheckVar(request("menu"),15)
idx  = request("idx")
mastercode = request("mastercode")
detailcode = request("detailcode")
device = request("device")

if idx = "" then idx = 0
if idx = 0 then 
	mode = "add"
else
	mode = "modi"
end if 

set oSlideManage = new SlideListCls
	oSlideManage.FRectIdx = idx
	oSlideManage.getSlide()

	if oSlideManage.FItem.Fmastercode <> "" then mastercode = oSlideManage.FItem.Fmastercode
	if oSlideManage.FItem.Fdetailcode <> "" then detailcode = oSlideManage.FItem.Fdetailcode
	if oSlideManage.FItem.Fdevice <> "" then device = oSlideManage.FItem.Fdevice
	if oSlideManage.FItem.Fimageurl <> "" then bannerImg = oSlideManage.FItem.Fimageurl
	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<script type="text/javascript">
// ����ǰ �߰� �˾�
function findProd() {
		var popwin;
		popwin = window.open("/admin/Diary2009/pop_additemlist.asp", "popup_item", "width=900,height=600,scrollbars=yes,resizable=yes");
		popwin.focus();
}

function chgselectbox(v) {
	if (v != '' ){
		location.href = "?idx=<%=idx%>&mode=<%=mode%>&mastercode="+v;
	} else {
		location.href = "?idx=<%=idx%>&mode=<%=mode%>";
	}
}

function regitem() {
	var frm = document.frmreg;

	if (!frm.mastercode.value) {
		alert('��ȹ���� ���� ���ּ���.');
		return false;
	}

	if(!frm.device[0].checked&&!frm.device[1].checked) {
		alert("ä�θ� �������ּ���.");
		frm.device[0].focus();
		return false;
	}

	if(!frm.titlename.value) {
		alert("������ �Է����ּ���.");
		frm.titlename.focus();
		return false;
	}
	
	if(!frm.isvideo[0].checked&&!frm.isvideo[1].checked) {
		alert("������ ��� ������ ���� ���ּ���");
		frm.isvideo[0].focus();
		return false;
	}

	if(!frm.linkurl.value.length) {
		alert("��ũ��(URL)�� �Է����ּ���.");
		frm.linkurl.focus();
		return false;
	}

	if(frm.linkurl.value.includes("�귣����̵�")||frm.linkurl.value.includes("ī�װ�")||frm.linkurl.value.includes("��ǰ�ڵ�")||frm.linkurl.value.includes("�̺�Ʈ�ڵ�")) {
		alert("��Ȯ�� ��ũ�� �Է����ּ���.");
		frm.linkurl.focus();
		return false;
	}

	if(!frm.StartDate.value) {
		alert("������ �Է����ּ���.");
		frm.StartDate.focus();
		return false;
	}

	if(!frm.EndDate.value) {
		alert("������ �Է����ּ���.");
		frm.EndDate.focus();
		return false;
	}

	if ($("#evt_startdate").text() != "") {
		if (frm.StartDate.value < $("#evt_startdate").text())
		{
			alert("��� ������ �̺�Ʈ ������ ������ �� �� �����ϴ�.");
			frm.StartDate.value = $("#evt_startdate").text();
			return false;
		}
	}

	if ($("#evt_enddate").text() != "") {
		if (frm.EndDate.value > $("#evt_enddate").text())
		{
			alert("��� ������ �̺�Ʈ ������ ���İ� �� �� �����ϴ�.");
			frm.EndDate.value = $("#evt_enddate").text();
			return false;
		}
	}
	
	if(confirm("�Է��Ͻ� �������� ����Ͻðڽ��ϱ�?")){
		frm.submit();
	}
}

function fnchggroup() {
	var frm = document.frm;

	if (!frm.mastercode.value) {
		frm.mastercode.value = document.frmreg.mastercode.value;
	}

	frm.submit();
}

// ���ε� ���� Ȯ�� �� ó��
function jsCheckUpload() {
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
					document.frmreg.bannerImg.value=resultObj.fileurl;
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
						document.frmreg.bannerImg.value="";
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

function jsLastEvent(num){
	winLast = window.open('/admin/exhibitionitems/pop_event_lastlist.asp?num='+num,'pLast','width=800,height=600, scrollbars=yes')
	winLast.focus();
}

// ��ũ�� ����
function showDrop(){
	$(".selectLink ul").show();
}

//�����Է�
function populateTextBox(t,v){
	var val = v;
	var linktype = t;
	var device = document.frmreg.device.value;
		
	switch (t) {
		case 'event'  : 
			val = "/event/eventmain.asp?" + v;
			break;
		case 'item' : 
			if (device == "" || device == "P") {
				val = "/shopping/category_prd.asp?" + v;
			} else {
				val = "/category/category_itemprd.asp?" + v;
			}
			break;
		case 'category'  : 
			if (device == "" || device == "P") {
				val = "/shopping/category_list.asp?" + v;
			} else {
				val = "/category/category_list.asp?" + v;
			}
			break;
		case 'brand'  : 
			val = "/street/street_brand.asp?" + v;
			break;
		default   : 
			break;
	}

	$("#linkurl").val(val);
	$(".selectLink ul").css("display","none");
}

function linkcopy(){
	var val = $("#linkurl").val();
	$("#linkurl").attr("value",val);
	$(".selectLink ul").css("display","none");
}

function jsCopyEvtDate() {
	var eventStartDate = $("#evt_startdate").text();
	var eventEndDate = $("#evt_enddate").text();

	document.frmreg.StartDate.value = (eventStartDate != "") ? eventStartDate : "";
	document.frmreg.EndDate.value = (eventEndDate != "") ? eventEndDate : "";
}
</script>
</head>
<body>
<div class="contSectFix scrl">
	<div class="pad20">
		<form name="frm" method="get" action="">
			<input type="hidden" name="menu" value="<%=menu%>"/>
			<input type="hidden" name="mastercode" value=""/>
			<input type="hidden" name="idx" value="<%= idx %>"/>
		</form>
		<form name="frmreg" method="post" action="/admin/common/lib/pop_slide/pop_slide_manage_proc.asp">
		<input type="hidden" name="mode" value="<%= Mode %>"/>
		<input type="hidden" name="idx" value="<%= idx %>"/>
		<input type="hidden" name="menu" value="<%= menu%>"/>
		<table class="tbType1 listTb">
			<tr>
				<td>
					<table class="tbType1 listTb">
						<tr bgcolor="#FFFFFF" height="25">
							<td colspan="2" ><b>�����̵� ���</b></td>
						</tr>
						<tr bgcolor="#FFFFFF">
							<th>��ȹ��</th>
							<td style="text-align:left;">
								<%=DrawSelectAllView("mastercode",mastercode,"fnchggroup",menu)%>
								<% if mastercode > 0 then %>
									<%=DrawSelectDetailView("detailcode",mastercode,detailcode,"",menu)%>
								<% end if %>
							</td>
						</tr>
						<tr bgcolor="#FFFFFF">
							<th>ä��</th>
							<td style="text-align:left;">
								<input type="radio" value="P" name="device" <%=chkiif(device="P","checked","") %>> PC
								<input type="radio" value="M" name="device" <%=chkiif(device="M","checked","") %>> M/A
							</td>
						</tr>
						<tr bgcolor="#FFFFFF">
							<th>����<br/>(�������� �ؽ�Ʈ)</th>
							<td style="text-align:left;">
								<input type="text" name="titlename" value="<%= oSlideManage.FItem.Ftitlename%>" size="50"/>
							</td>
						</tr>
						<tr bgcolor="#FFFFFF">
							<th>����2<br/>(�������� �ؽ�Ʈ)</th>
							<td style="text-align:left;">
								<input type="text" name="subtitlename" value="<%= oSlideManage.FItem.Fsubtitlename%>" size="50"/>
							</td>
						</tr>
						<tr bgcolor="#FFFFFF">
							<th>�ؽ�Ʈ ���� (��������)</th>
							<td style="text-align:left;">
								#<input type="text" name="titlecolor" value="<%= oSlideManage.FItem.Ftitlecolor%>" size="6" maxlength="6"/>
							</td>
						</tr>
						<tr bgcolor="#FFFFFF">
							<th>�̹��� ���</th>
							<td style="text-align:left;">
								<p class="tMar05">
									<input type="hidden" name="bannerImg" value="<%=bannerImg%>" />
									<div style="width:220px; height:220px;">
										<div id="lyrPrgs" style="display:none; position:absolute;padding:101px; background-color:rgba(0,0,0,0.2);"><img src="http://fiximage.10x10.co.kr/web2015/giftcard/ajax_loader.gif" alt="progress" /></div>
										<img id="lyrBnrImg" src="<%=chkIIF(bannerImg="" or isNull(bannerImg),"/images/admin_login_logo2.png",bannerImg)%>" style="height:218px; border:1px solid #EEE;"/>
									</div>
									<div id="lyrImgDelBtn" class="btn" style="<%=chkIIF(idx = 0 or bannerImg="","display:none;","")%>" onclick="jsDelImg();">�̹��� ����</button></div>
									<div id="lyrImgUpBtn" class="btn" style="<%=chkIIF(idx = 0 or bannerImg="","","display:none;")%>"><label for="fileupload">�̹��� ���ε�</label></div>
								</p>
							</td>
						</tr>
						<tr bgcolor="#FFFFFF">
							<th>���� ���</th>
							<td style="text-align:left;">
								#<input type="text" name="Lcolor" value="<%=oSlideManage.FItem.FLcolor%>" size="6" maxlength="6"/> ���� 
								#<input type="text" name="Rcolor" value="<%=oSlideManage.FItem.FRcolor%>" size="6" maxlength="6"/> ����
							</td>
						</tr>
						<tr bgcolor="#FFFFFF">
							<th>������ �������</th>
							<td style="text-align:left;">
								<input type="radio" name="isvideo" value="1" <%=chkiif(oSlideManage.FItem.Fisvideo = "1","checked","")%>/> ���
								<input type="radio" name="isvideo" value="0" <%=chkiif(oSlideManage.FItem.Fisvideo = "0" or oSlideManage.FItem.Fisvideo = "","checked","")%>/> ������
							</td>
						</tr>
						<tr bgcolor="#FFFFFF">
							<th>������ ��ũ��<br/>HTML</th>
							<td style="text-align:left;">
								<textarea name="videohtml" rows="8" cols="60"><%=oSlideManage.FItem.Fvideohtml%></textarea>
							</td>
						</tr>
						<tr bgcolor="#FFFFFF">
							<th>URL</th>
							<td style="text-align:left;">
								<div class="selectLink">
									<input type="text" value="<%=oSlideManage.FItem.Flinkurl%>" placeholder="��ũ�� �Է�(����)" onclick="showDrop();" name="linkurl" id="linkurl" onkeyup="linkcopy();" size="20" autocomplete="off"/>
									<ul style="display:none;">
										<li onclick="populateTextBox('');">���þ���</li>
										<!--<li onclick="populateTextBox('#group�׷��ڵ�');">#group�׷��ڵ�</li>-->
										<li onclick="populateTextBox('event','eventid=�̺�Ʈ�ڵ�');">�̺�Ʈ</li>
										<li onclick="populateTextBox('item','itemid=��ǰ�ڵ�');">��ǰ(O)</li>
										<li onclick="populateTextBox('category','disp=ī�װ�');">ī�װ�</li>
										<li onclick="populateTextBox('brand','makerid=�귣����̵�');">�귣��</li>
									</ul>
								</div>
							</td>
						</tr>
						<tr bgcolor="#FFFFFF">
							<th>�̺�Ʈ�ڵ�</th>
							<td style="text-align:left;">
								<input type="text" name="evt_code" value="<%=oSlideManage.FItem.Feventid%>" size="6"/> <input type="button" value="�̺�Ʈ �ҷ�����" onclick="jsLastEvent(1);"/>
								<input type="button" id="copyDateButton" value="�Ⱓ ����" onclick="jsCopyEvtDate();" style="display:none;"/>
							</td>
						</tr>
						<tr bgcolor="#FFFFFF">
							<th>������</th>
							<td style="text-align:left;">
								<input type="text" name="StartDate" id="startdate" value="<%=oSlideManage.FItem.Fstartdate%>">
								<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" style="vertical-align:middle;"/>
								<script type="text/javascript">
								var CAL_Start = new Calendar({
									inputField : "startdate",
									trigger	: "startdate_trigger",
									onSelect: function() {
										var date = Calendar.intToDate(this.selection.get());
										CAL_End.args.min = date;
										CAL_End.redraw();
										this.hide();
									},
									bottomBar: true,
									dateFormat: "%Y-%m-%d"
								});
								</script>
								<span style="color:red">�̺�Ʈ ������ : <span id="evt_startdate"><%=oSlideManage.FItem.Fevt_startdate%></span></span>
							</td>
						</tr>
						<tr bgcolor="#FFFFFF" >
							<th>������</th>
							<td style="text-align:left;">
								<input type="text" name="EndDate" id="enddate" value="<%=oSlideManage.FItem.Fenddate%>">
								<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" style="vertical-align:middle;"/>
								<script type="text/javascript">
								var CAL_End = new Calendar({
									inputField : "enddate",
									trigger	: "enddate_trigger",
									onSelect: function() {
										var date = Calendar.intToDate(this.selection.get());
										CAL_Start.args.max = date;
										CAL_Start.redraw();
										this.hide();
									},
									bottomBar: true,
									dateFormat: "%Y-%m-%d"
								});
								</script>
								<span style="color:red">�̺�Ʈ ������ : <span id="evt_enddate"><%=oSlideManage.FItem.Fevt_enddate%></span></span>
							</td>
						</tr>
						<tr>
							<th>���ļ���</th>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="text" name="sorting" value="<%=chkiif(oSlideManage.FItem.Fsorting = "","99",oSlideManage.FItem.Fsorting)%>">
							</td>
						</tr>
						<tr>
							<th>��뿩��</th>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="radio" name="isusing" value="1" id="usey" <%=chkiif(oSlideManage.FItem.Fisusing = ""  or oSlideManage.FItem.Fisusing = "1" , "checked" , "")%>> <label for="usey">�����</label>
								<input type="radio" name="isusing" value="0" id="usen" <%=chkiif(oSlideManage.FItem.Fisusing = "0" , "checked" , "")%>> <label for="usen">������</label>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td colspan="2">
					<img src="http://webadmin.10x10.co.kr/images/icon_save.gif" border="0" onClick="regitem();" style="cursor:pointer">
					<img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" border="0" onClick="frmreg.reset();" style="cursor:pointer">
				</td>
			</tr>
		</table>
		</form>
		<%'// �̹��� ���ε� %>
		<form name="frmUpload" id="ajaxform" action="<%=uploadImgUrl%>/linkweb/common/simpleCommonImgUploadProc.asp" method="post" enctype="multipart/form-data" style="display:none; height:0px;width:0px;">
			<input type="file" name="upfile" id="fileupload" onchange="jsCheckUpload();" accept="image/*" />
			<input type="hidden" name="mode" id="fileupmode" value="upload">
			<input type="hidden" name="div" value="SB">
			<input type="hidden" name="upPath" value="/event/swipeimage/">
			<input type="hidden" name="tuid" value="<%=encUsrId%>">
			<input type="hidden" name="prefile" id="filepre" value="<%=bannerImg%>">
		</form>
	</div>
</div>
<%
	set oSlideManage = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->