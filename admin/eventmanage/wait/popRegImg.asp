<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : �̹������
' History : 
'####################################################
%>
<<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 

<!-- #include virtual="/admin/lib/popheader.asp"-->
 
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
%>
<link rel="stylesheet" type="text/css" href="/css/adminPartnerCommon.css" />

<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script language="javascript">
<!--
	function jsUpload(){
		if(!document.frmImg.sfImg.value){
			alert("ã�ƺ��� ��ư�� ���� ���ε��� �̹����� ������ �ּ���.");			
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
				alert("���� ������ " + maxwidth + "px ���Ϸ� �÷��ּ���!"); 
			}
			if(imgheight > maxheight){
				alert("���� ������ " + maxheight + "px ���Ϸ� �÷��ּ���!"); 
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
		<p class="btnClose"><input type="image" src="/images/partner/pop_admin_btn_close.gif" alt="â�ݱ�" onclick="window.close();" /></p>
	</div>
	<div class="popContent scrl">
		<div class="contTit bgNone"><!-- for dev msg : Ÿ��Ʋ �����ϴܿ� searchWrap�� �� ��쿣 bgNone Ŭ���� ���� -->
			<h2>�̹��� ���</h2>
		</div>
		<div class="cont">  
			<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/event_admin/eventWaitUpload.asp" enctype="MULTIPART/FORM-DATA"  >						
			<input type="hidden" name="iMW" value="<%=iMW%>">
			<input type="hidden" name="iMH" value="<%=iMH%>">
			<input type="hidden" name="iML" value="1">
			<input type="hidden" name="sType" value="<%=sType%>">
			<input type="hidden" name="pvWidth" value="<%=pvWidth%>">
			<table class="tbType1 writeTb tMar10">
				<colgroup>
					<col width="15%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>�̹�����</div></th>
					<td> <p><input type="file"  name="sfImg"  id="sfImg" style="width:90%;" placeholder="<%=sImgName%>"/></p>
							<p class="tPad05 fs11 cGy1">- �̹��� ������ : <b><%=iMW%>X<%=iMH%></b>px</p>	
							<p class="tPad05 fs11 cGy1">- ���� Ÿ�� : <b>gif,jpg,png</b></p>	
					</td>
				</tr>  
			</table>
			<div class="tPad15 ct">
				<input type="button" value="���" onclick="jsUpload();" class="btn3 btnDkGy" />			 
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
			<td height="50" ><p style="padding:10px;">���ε� ó�����Դϴ�. ��ø� ��ٷ��ּ���~~ </p></td>
		</tr>
	</table>
</div>