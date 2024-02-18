<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ��Ƽ3�� �̺�Ʈ ����
' History : 2018.11.05 ������ ����
'####################################################
%>
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/sitemasterclass/Multi3Cls.asp" -->
<%
dim evt_code, unitIdx
dim encUsrId, tmpTx, tmpRn, userid
userid = session("ssBctId")
evt_code = request("evt_code")
unitIdx = request("unitIdx")

Randomize()
tmpTx = split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z",",")
tmpRn = tmpTx(int(Rnd*26))
tmpRn = tmpRn & tmpTx(int(Rnd*26))
encUsrId = tenEnc(tmpRn & userid)	

%>

<script type="text/javascript">
function addItem(){	
	var frm = document.itenFrm;
	if(!chkValidation(frm))return false;	
	var link = "multi3_process.asp"
	frm.action = link;
	frm.submit();
}
function chkValidation(frm){
	if(frm.itemid.value==""){
		alert("��ǰ�� �������ּ���.");
		return false;
	}
	return true;
}
function findItem(){		
	var popwin; 		
	popwin = window.open("multi3_eventitem_regist.asp", "popup_item_search", "width=1024,height=768,scrollbars=yes,resizable=yes");
	popwin.focus();
}
function jsCheckUpload() {
	var gubun = document.frmUpload.imgtype.value;
	var mainfrm = document.itenFrm	
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
					$("img[id="+gubun+"src]").hide().attr("src",$("#filepre").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
					$("input[id="+gubun+"]").val(resultObj.fileurl);															
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
function setImgType(type){		
	document.frmUpload.imgtype.value = type;
	return false;
}
</script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<h3>��Ƽ3�� ���������� ��ǰ �߰�</h3>
�̺�Ʈ�ڵ� : <%=evt_code%>
<div>			
	<form name="itenFrm">
	<input type="hidden" name="mode" value="itemadd">
	<input type="hidden" name="evt_code" value="<%=evt_code%>">
	<input type="hidden" name="unitIdx" value="<%=unitIdx%>">													
		<table style="border:solid 1px black;margin-top:10px;width:550px;" id="itemContainer">						
			<tr>
				<td rowspan=3>
					<div class="inTbSet" align="center">												
						<div>	
							<p class="registImg">
								<input type="hidden" id="item_img" name="item_img" value="" />
								<img name="item_imgsrc" id="item_imgsrc" src="/images/admin_login_logo2.png" style="height:138px; border:1px solid #EEE;"/>																
							</p>																													
							<button type="button">
								<div onclick="setImgType('item_img')" >
									<label for="fileupload" style="cursor:pointer;">�̹��� ���ε�
									</label>
								</div>
							</button>														
						</div>	
					</div>					
				</td>
				<td style="border-bottom: 1px solid">��ǰid</td>
				<td style="border-bottom: 1px solid">										
					<input type="text" name="itemid" readonly value="">										
					<input type="button" onclick="findItem()" value="��ǰã��">
				</td>						
			</tr>							
			<tr>
				<td style="border-bottom: 1px solid">��ǰ��</td>
				<td style="border-bottom: 1px solid">
					<input type="text" name="item_name">
				</td>												
			</tr>
			<tr>
				<td>��ǰ����</td>
				<td>
					<input style="width:50px" type="number" name="item_order">
				</td>						
			</tr>			
		</table>															
	</form>
</div>
<div align="center">
<input type="button" onclick="addItem();" value="����">
<input type="button" onclick="window.close();" value="���">
</div>
<form name="frmUpload" id="ajaxform" action="<%=uploadImgUrl%>/linkweb/common/simpleCommonImgUploadProc.asp" method="post" enctype="multipart/form-data" style="display:none; height:0px;width:0px;">
<input style="display:none" type="file" name="upfile" id="fileupload" onchange="jsCheckUpload();" accept="image/*" />
<input type="hidden" name="mode" id="fileupmode" value="upload">
<input type="hidden" name="div" value="TQ">
<input type="hidden" name="upPath" value="/appmanage/multi3img/">
<input type="hidden" name="tuid" value="<%=encUsrId%>">
<input type="hidden" name="prefile" id="filepre" >	
<input type="hidden" name="imgtype">
</form>		
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
