<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/street/brandmainCls.asp" -->
<%
Dim idx, reload, ix, makerid
	idx = request("idx")
	reload = request("reload")
	if idx="" then idx=0

If reload="on" then
    response.write "<script>opener.location.reload(); window.close();</script>"
    dbget.close()	:	response.End    
End If

Dim mbrand
Set mbrand = New cBrandMain
	mbrand.FRectIdx = idx
	mbrand.sBrandImageGetOne

	makerid = mbrand.FOneItem.fmakerid
%>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="http://m.10x10.co.kr/lib/css/main.css" />
<script src="http://code.jquery.com/jquery-latest.min.js"></script>
<script src="http://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script type='text/javascript'>
	function SaveMainContents(frm){
		if (frm.makerid.value.length<1){
			alert('브랜드를 입력 하세요.');
			frm.makerid.focus();
			return;
		}
		if (confirm('이미지를 한번더 확인 해주세요. 저장 하시겠습니까?')){
			frm.submit();
		}
	}

	function fileInfo(f){
		var file = f.files; // files 를 사용하면 파일의 정보를 알 수 있음

		var reader = new FileReader(); // FileReader 객체 사용
		reader.onload = function(rst){ // 이미지를 선택후 로딩이 완료되면 실행될 부분
			$('#img_box').empty().html('<img src="' + rst.target.result + '">'); // append 메소드를 사용해서 이미지 추가
			// 이미지는 base64 문자열로 추가
			// 이 방법을 응용하면 선택한 이미지를 미리보기 할 수 있음
		}
		reader.readAsDataURL(file[0]); // 파일을 읽는다, 배열이기 때문에 0 으로 접근
	}

	// typing
	function textclone(k,v){
		var frmtext = $("#"+k);
		frmtext.bind("keyup",function(){
			var msg = $(this).val();
			if ($(this).val().length > 0){
				msg = msg.replace(/(?:\r\n|\r|\n)/g, '<br>');
				$("#"+v).html(msg);
			}else{
				$("#"+v).html("");
			}
		});
	}

	function fnDeleteImage(){
		$('#viewbox').empty();
		$("#file1").val("");
		$("#imgurl").html("");
		$('#img_box').empty();
		$("#dfile").val("Y");
	}
</script>
<div class="popWinV17">
	<h1>등록</h1>
	<div class="popContainerV17 pad10">
		<div class="ftLt col2">
			<form name="frm" method="post" action="<%=staticUploadUrl%>/linkweb/street/do_image_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
			<table class="tbType1 writeTb">
			<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
			<input type="hidden" name="idx" value="<%=idx%>">
			<input type="hidden" name="dfile" id="dfile">
				<tr bgcolor="#FFFFFF">
					<td width="100" align="center">브랜드 :</td>
					<td width="300">
					<% If idx = 0 Then %>
						<% drawSelectBoxDesignerwithName "makerid",makerid %>
					<% Else %>
						<%=makerid%>
						<input type="hidden" name="makerid" value="<%=makerid%>"/>
					<% End If %>
					</td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td width="100" align="center">이미지 : </td>
					<td>
					<input type="file" name="file1" id="file1" value="" size="32" maxlength="32" class="formFile" accept="image/*" onchange="fileInfo(this)">
					<% if mbrand.FOneItem.Fbrandimage<>"" then %>
					<br>
					<div id="viewbox"><img src="<%=uploadUrl%>/brandstreet/main/<%= mbrand.FOneItem.Fbrandimage %>" id="img" width="200" alt="" /></div>
					<br><span id="imgurl"> <%=uploadUrl%>/brandstreet/main/<%= mbrand.FOneItem.Fbrandimage %>
					<br><input type="button" value=" 삭제 " onClick="fnDeleteImage();"></span>
					<% end if %>
					<font color="red">*이미지 사이즈 : 640 X 907</font>
					</td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td width="100" align="center">사용여부</td>
					<td>
						<input type="radio" name="isusing" value="1" checked <%=chkIIF(mbrand.FOneItem.fisusing = "Y","checked","")%> >Y
						<input type="radio" name="isusing" value="0" <%=chkIIF(mbrand.FOneItem.fisusing = "N","checked","")%>>N
					</td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td align="center" colspan="2">
						<input type="button" value=" 저 장 " onClick="SaveMainContents(frm);" class="button">
					</td>
				</tr>	
			</table>
			</form>
		</div>
		<div style="position:fixed;left:48%;top:50px;">
			<div class="lPad30 vTop">
				<%'타입별 템플릿 %>
				<%'rolling image%>
				<div class="text-bnr">
				<section style="width:375px;">
					<div class="thumbnail" id="img_box">
						<% If idx > 0 Then %>
						<img src="<%=uploadUrl%>/brandstreet/main/<%= mbrand.FOneItem.Fbrandimage %>" id="viewimg" alt="" width="375"/>
						<% End If %>
					</div>
				</section>
				</div>
			</div>
		</div>
	</div>
</div>

<%
set mbrand = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->