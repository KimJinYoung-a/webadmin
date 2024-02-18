<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 이미지등록
' History : 
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 

<!-- #include virtual="/admin/lib/popheader.asp"-->
 
<%
Dim sFolder,  sName,sSpan, slen, arrImg,  vYear
dim sImg1,sImg2,sImg3
dim  sImgName1,sImgName2, sImgName3
Dim sOpt
dim sType, iMH, iMW,pvWidth
dim sFileType
sFolder = Request.Querystring("sF") 
sImg1 = Request.Querystring("sImg1")
sImg2 = Request.Querystring("sImg2")
sImg3 = Request.Querystring("sImg3")
iMH = requestCheckVar(Request("iMH"),10)  
iMW = requestCheckVar(Request("iMW"),10)  
sType = requestCheckVar(Request("sType"),10)  
 

pvWidth = requestCheckVar(Request("pvWidth"),10) 
%>
<link rel="stylesheet" type="text/css" href="/css/adminPartnerCommon.css" />

<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script language="javascript">
<!--
	function jsUpload(){
	 
		document.frmImg.submit();
		document.all.dvLoad.style.display = "";
		
	}

$(document).ready(function(){
	
		var sfpath1 = $(opener.document).find("#hid<%=sType%>1").val(); 
		var sfpath2 = $(opener.document).find("#hid<%=sType%>2").val(); 
 		var sfpath3 = $(opener.document).find("#hid<%=sType%>3").val(); 
	
		var sfimg1 = "<button type='button' onclick=jsDelimg('<%=sType%>1')>X</button><img src='"+sfpath1+"' style='width:105px;' />"
		var sfimg2 = "<button type='button' onclick=jsDelimg('<%=sType%>2')>X</button><img src='"+sfpath2+"' style='width:105px;' />"
		var sfimg3 = "<button type='button' onclick=jsDelimg('<%=sType%>3')>X</button><img src='"+sfpath3+"' style='width:105px;' />"							
									 
 	 	
	if (sfpath1!=""){
		 $("#pv<%=sType%>1").html(sfimg1);	
		 $("#hid<%=sType%>1").val(sfpath1); 
		}
	if (sfpath2!=""){ 
		$("#pv<%=sType%>2").html(sfimg2);	
		 $("#hid<%=sType%>2").val(sfpath2); 
		}
	if (sfpath3!=""){ 
		$("#pv<%=sType%>3").html(sfimg3);	
		 $("#hid<%=sType%>3").val(sfpath3); 
		}
		 
	var _URL = window.URL || window.webkitURL;
		for(i=0;i<3;i++){
		$('#sfImg'+i).change(function (){
			var file = $(this)[i].files[i];
			img = new Image();
			var imgwidth = 0;
			var imgheight = 0;
			var maxwidth = <%=iMW%>;
			var maxheight = <%=iMH%>;

			img.src = _URL.createObjectURL(file);
			img.onload = function() {
				imgwidth = this.width;
				imgheight = this.height;

				//$("#width"+i).text(imgwidth);
				//$("#height"+i).text(imgheight);
				if(imgwidth > maxwidth){
					alert("가로 사이즈 " + maxwidth + "px 이하로 올려주세요!");
				}
				if(imgheight > maxheight){
					alert("세로 사이즈 " + maxheight + "px 이하로 올려주세요!");
				}
			}
		});
	}
 
 	
});

 
 //이미지삭제
 function jsDelimg(sTypev){ 
 	$("#pv"+sTypev).empty();
 	$("#hid"+sTypev).val("");
 }
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
			<h2>이미지 관리</h2>
		</div>
		<div class="cont">  
			<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/event_admin/eventWaitMultiUpload.asp" enctype="MULTIPART/FORM-DATA"  >						
			<input type="hidden" name="iMW" value="<%=iMW%>">
			<input type="hidden" name="iMH" value="<%=iMH%>">
			<input type="hidden" name="iML" value="1">
			<input type="hidden" name="sType" value="<%=sType%>">
			<input type="hidden" name="pvWidth" value="<%=pvWidth%>">
			<input type="hidden" name="hid<%=sType%>1" id="hid<%=sType%>1" value="">
			<input type="hidden" name="hid<%=sType%>2" id="hid<%=sType%>2" value="">
			<input type="hidden" name="hid<%=sType%>3" id="hid<%=sType%>3" value="">
		 	<p class="tPad05 fs11 cGy1">-이미지 삭제는 이미지의 [X] 버튼을 누르신 후 아래쪽의 [저장] 버튼을 눌러주세요</p>
			<table class="tbType1 writeTb tMar10">
				<colgroup>
					<col width="15%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>이미지1</div></th>
					<td>
						<div class="inTbSet">
							<div class="formFile">
								<p><input type="file" id="formFile"  name="sfImg1" style="width:85%;" /></p> 
								<p class="tPad05 fs11 cGy1">- 이미지 사이즈 : <b><%=iMW%>X<%=iMH%></b>px</p>	
								<p class="tPad05 fs11 cGy1">- 파일 타입 : <b>gif,jpg,png</b></p>	
							</div>
							<div style="width:105px;">
								<p class="registImg" id="pv<%=sType%>1">
									
								</p>
							</div>
						</div>								 
					</td>
				</tr>
				<tr>
					<th><div>이미지2</div></th>
					<td> <div class="inTbSet">
							<div class="formFile">
								<p><input type="file" id="formFile"  name="sfImg2" style="width:85%;" /></p> 
								<p class="tPad05 fs11 cGy1">- 이미지 사이즈 : <b><%=iMW%>X<%=iMH%></b>px</p>	
								<p class="tPad05 fs11 cGy1">- 파일 타입 : <b>gif,jpg,png</b></p>	
							</div>
							<div style="width:105px;">
								<p class="registImg" id="pv<%=sType%>2">
									
								</p>
							</div>
						</div>		
						</td>
				</tr>
				<tr>
					<th><div>이미지3</div></th>
					<td> <div class="inTbSet">
							<div class="formFile">
								<p><input type="file" id="formFile"  name="sfImg3" style="width:85%;" /></p> 
								<p class="tPad05 fs11 cGy1">- 이미지 사이즈 : <b><%=iMW%>X<%=iMH%></b>px</p>	
								<p class="tPad05 fs11 cGy1">- 파일 타입 : <b>gif,jpg,png</b></p>	
							</div>
							<div style="width:105px;">
								<p class="registImg" id="pv<%=sType%>3">
									
								</p>
							</div>
						</div>	
						</td>
				</tr>					 
			</table>   
			<div class="tPad15 ct">
				<input type="button" value="저장" onclick="jsUpload();" class="btn3 btnDkGy" />			 
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