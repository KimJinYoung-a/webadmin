<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/classes/sitemasterclass/GiftCardImageCls.asp" -->					   
<%
Dim userid, encUsrId, tmpTx, tmpRn, i, idx, mode
dim designId
dim giftcardImage
dim giftcardAlt
dim sortNumber
dim adminRegister
dim adminName
dim adminModifyer
dim adminModifyerName
dim registDate
dim lastUpDate
dim isusing

userid = session("ssBctId")

idx = requestCheckvar(request("idx"),16) 

'�׽�Ʈ������

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

If idx <> "" then
	dim giftCardImgObj
	set giftCardImgObj = new GiftCardImageCls
	giftCardImgObj.FRectIdx = idx
	giftCardImgObj.GetOneContent()

	designId			= giftCardImgObj.FOneItem.FdesignId		
	giftcardImage		= giftCardImgObj.FOneItem.FGiftCardImage		
	giftcardAlt			= giftCardImgObj.FOneItem.FGiftCardAlt	
	sortNumber			= giftCardImgObj.FOneItem.FSortNumber	
	adminRegister		= giftCardImgObj.FOneItem.FAdminRegister	
	adminName			= giftCardImgObj.FOneItem.FAdminName	
	adminModifyer		= giftCardImgObj.FOneItem.FAdminModifyer	
	adminModifyerName	= giftCardImgObj.FOneItem.FAdminModifyerName			
	registDate			= giftCardImgObj.FOneItem.FRegistDate	
	lastUpDate			= giftCardImgObj.FOneItem.FLastUpDate
	isusing				= giftCardImgObj.FOneItem.FIsusing

	set giftCardImgObj = Nothing
else

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
					if(gubun === "cardImg"){
						$("#lyrBnrImg").hide().attr("src",$("#filepre").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
						$("#giftcardImage").val(resultObj.fileurl);
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
	self.location.href="/admin/giftcardimage/";
	}
function addContent(){
	var mainfrm = document.frm;		
		if(mainfrm.giftcardImage.value==""){
			alert('�̹����� �Է����ּ���.');
			return false;
		}		
		mainfrm.action="addimage.asp";
		mainfrm.submit();		
	}
function setImgType(type){
	document.frmUpload.imgtype.value = type;
	return false;
}
// ���ε� ���� Ȯ�� �� ó��
</script>
<div class="popWinV17">
	<h1>����Ʈī�� �̹��� ���/����</h1>
	<button type="button" class="btn btn2" style="position:absolute; right:15px; top:7px;">����</button>
	<div class="popContainerV17 pad30">
		<p class="cGn1">* ���� �ۼ� �� �ݵ�� '����' ��ư�� �����ּ���.</p>
		<% if mode = "modify" then%>
		<h2 class="tMar20 subType">�̹��� ����</h2>
		<% else %>
		<h2 class="tMar20 subType">�̹��� ���</h2>
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
						<input type="radio" name="isusing" value="1" <%=chkIIF(isusing="1" or isusing="" ,"checked","")%>>��� ��<br>
						<input type="radio" name="isusing" value="0" <%=chkIIF(isusing="0","checked","")%>>��� ����
					</td>					
				</tr>							
				<tr>
					<th><div>�̹���<strong class="cRd1">*</strong></div></th>
					<td>
						<div class="inTbSet">							
							<div>	
								<p class="registImg">
									<input type="hidden" id="giftcardImage" name="giftcardImage" value="<%=giftcardImage%>" />
									<img id="lyrBnrImg" src="<%=chkIIF(giftcardImage="" or isNull(giftcardImage),"/images/admin_login_logo2.png",giftcardImage)%>" style="height:218px; border:1px solid #EEE;"/>
									<div id="lyrImgUpBtn" class="btn lMar05" style="margin-left:65px;" onclick="setImgType('cardImg')"><label for="fileupload" style="cursor:pointer"><%=chkIIF(idx="" and giftcardImage="","�̹��� ���ε�","�̹��� ����")%></label></div>
								</p>				
							</div>					
						</div>
					</td>
				</tr>
				<tr>
					<th><div>alt��</div></th>
					<td>
						<input type="text" name="giftcardAlt" value="<%=giftcardAlt%>"  class="formTxt" style="width:10%;" />						
					</td>
				</tr>															
				</tbody>
			</table>
			<input type="hidden" name="mode" value="<%=mode%>" />
			<input type="hidden" name="idx" value="<%=idx%>">
			<input type="hidden" id="OrderChangedFlag" name="OrderChangedFlag" value="">		
<!--========================================����===================================================-->		
	</div>			
	    </form>
	<div class="popBtnWrap">		
		<input type="button" value="���" onclick="jsgolist();" style="width:100px; height:30px;" />
		<input type="button" value="����" onclick="addContent();" class="cRd1" style="width:100px; height:30px;" />		
	</div>
		<!-- ī���̹��� -->
	<form name="frmUpload" id="ajaxform" action="<%=uploadImgUrl%>/linkweb/common/simpleCommonImgUploadProc.asp" method="post" enctype="multipart/form-data" style="display:none; height:0px;width:0px;">
		<input type="file" name="upfile" id="fileupload" onchange="jsCheckUpload();" accept="image/*" />
		<input type="hidden" name="mode" id="fileupmode" value="upload">
		<input type="hidden" name="div" value="TQ">
		<input type="hidden" name="upPath" value="/appmanage/giftcard/">
		<input type="hidden" name="tuid" value="<%=encUsrId%>">
		<input type="hidden" name="prefile" id="filepre" value="<%=giftcardImage%>">
		<input type="hidden" name="imgtype">
	</form>				
	<form name="delFrm" method="post">
		<input type="hidden" name="idx">
		<input type="hidden" name="mode" value="subdelete">
	</form>					
</div>
