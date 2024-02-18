<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 이미지등록
' History : 
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/partner/lib/adminHead.asp" -->
<%
Dim sFolder, sImg, sName,sSpan, slen, arrImg, sImgName, vYear
Dim sOpt
dim sType, iMH, iMW,pvWidth
dim sFileType
sFolder = Request.Querystring("sF") 
sImg = Request.Querystring("sImg")
iMH = requestCheckVar(Request("iMH"),10)  
iMW = requestCheckVar(Request("iMW"),10)  
sType = requestCheckVar(Request("sType"),10)  
IF sImg <> "" THEN
	arrImg = split(sImg,"/")
	slen = ubound(arrImg)
	sImgName = arrImg(slen)
END IF	

pvWidth = requestCheckVar(Request("pvWidth"),10) 

dim uploadImgUrl
dim C_IS_SSL_ENABLED : C_IS_SSL_ENABLED = (Request.ServerVariables("HTTPS") = "on")
IF application("Svr_Info")="Dev" THEN
 	uploadImgUrl 		= "http://testupload.10x10.co.kr"
ELSE
	if (C_IS_SSL_ENABLED = True) then
 		uploadImgUrl 		= "https://upload.10x10.co.kr"          '' upload.10x10.co.kr 통해서 Nas Server로 업로드
	else
 		uploadImgUrl 		= "http://upload.10x10.co.kr"          '' upload.10x10.co.kr 통해서 Nas Server로 업로드
	end if
 END IF
%>

<script language="javascript">
<!--
	function jsUpload(){
		if(!document.frmImg.sfImg.value){
			alert("찾아보기 버튼을 눌러 업로드할 이미지를 선택해 주세요.");			
			return;
		}
		document.frmImg.submit();
		document.all.dvLoad.style.display = "";
		
	}

$(document).ready(function(){
	var _URL = window.URL || window.webkitURL;
	$('#sfImg').change(function (){
		var file = $(this)[0].files[0];
		img = new Image();
		var imgwidth = 0;
		var imgheight = 0;
		var maxwidth = <%=iMW%>;
		var maxheight = <%=iMH%>;

		img.src = _URL.createObjectURL(file);
		img.onload = function() {
			imgwidth = this.width;
			imgheight = this.height;

			$("#width").text(imgwidth);
			$("#height").text(imgheight);
			if(imgwidth > maxwidth){
				alert("가로 사이즈 " + maxwidth + "px 이하로 올려주세요!"); 
			}
			if(imgheight > maxheight){
				alert("세로 사이즈 " + maxheight + "px 이하로 올려주세요!"); 
			}
		}
	});
});


//-->
</script>
</head>
<body>
<div class="popupWrap">
	<div class="popHead">
		<h1><img src="/images/partner/pop_admin_logo.gif" alt="10x10" /></h1>
		<p class="btnClose"><input type="image" src="/images/partner/pop_admin_btn_close.gif" alt="창닫기" onclick="window.close();" /></p>
	</div>
	<div class="popContent scrl">
		<div class="contTit bgNone"><!-- for dev msg : 타이틀 영역하단에 searchWrap이 올 경우엔 bgNone 클래스 삭제 -->
			<h2>이미지 등록</h2>
		</div>
		<div class="cont">  
			<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/partnerAdmin/JoinUpload.asp" enctype="MULTIPART/FORM-DATA"  >						
			<input type="hidden" name="iMW" value="<%=iMW%>">
			<input type="hidden" name="iMH" value="<%=iMH%>">
			<input type="hidden" name="iML" value="3">
			<input type="hidden" name="sType" value="<%=sType%>">
			<input type="hidden" name="pvWidth" value="<%=pvWidth%>">
			<table class="tbType1 writeTb tMar10">
				<colgroup>
					<col width="15%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>이미지명</div></th>
					<td> 					
						<div class="formFile">
						<p><input type="file"  name="sfImg"  id="formFile" style="width:90%;" placeholder="<%=sImgName%>"/></p>
							<p class="tPad05 fs11 cGy1">- 이미지 사이즈 : <b><%=iMW%>X<%=iMH%></b>px</p>	
							<p class="tPad05 fs11 cGy1">- 파일 타입 : <b>gif,jpg,png</b></p>	
						</div>
					</td>
				</tr>  
			</table>
			<div class="tPad15 ct">
				<input type="button" value="등록" onclick="jsUpload();" class="btn3 btnDkGy" />			 
			</div> 
			</form>	 
		</div>
	</div>
</div>
</body>
</html>			
<!-- #include virtual="/lib/db/dbclose.asp" -->
<div id="dvLoad" style="display:none;top:70px;left:50;position:absolute;background-color:gray;">
	<table class="tbType1 writeTb tMar10">
		<tr>
			<td height="50" ><p style="padding:10px;">업로드 처리중입니다. 잠시만 기다려주세요~~ </p></td>
		</tr>
	</table>
</div>