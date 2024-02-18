<%@ codepage="65001" language=vbscript %>
<% option explicit %>
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"

Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 동영상 삽입"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/DIYitemCls.asp"-->
<!-- #include virtual="/lib/util/base64Lib.asp"-->
<script language="jscript" runat="server">
function jsDecodeURIComponent(v) { return decodeURIComponent(v); }
function jsEncodeURIComponent(v) { return encodeURIComponent(v); }
</script>
<script type="text/javascript" src="/apps/academy/lib/waititemreg.js"></script>
<%
Dim param
param = Base64Decode(jsDecodeURIComponent(request("param")),"UTF-8")
''param = URLDecode(request("param"))
%>
<script>
jQuery(document).ready(function(){
	fnAPPShowRightConfirmBtns();
});
function fnAppCallWinConfirm(){
	var vodurl = $("#vodlink").val();
	var iframeCode='';
	var vodId = '';
	if(vodurl==""){
		alert("동영상 링크 정보를 입력해주세요.");
	} else {
		if (vodurl.indexOf("iframe") != -1){
			//iframe 등록
			vodurl = Base64.encode($("#vodlink").val());
			fnAPPopenerJsCallClose("fnVodLinkSet(\""+vodurl+"\")");
		}
		else {
			//URL 등록
			if (vodurl.indexOf("vimeo") > 0){
				vodId = getId(vodurl,'vimeo');
				//alert(vodId);
				iframeCode = '<iframe width="640" height="360" src="https://player.vimeo.com/video/' + vodId + '" frameborder="0" allowfullscreen></iframe>';
			}else{
				vodId = getId(vodurl,'youtube');
				iframeCode = '<iframe width="560" height="315" src="https://www.youtube.com/embed/' + vodId + '" frameborder="0" allowfullscreen></iframe>';
			}
			$("#vodlink").val(iframeCode);
			vodurl = Base64.encode($("#vodlink").val());
			fnAPPopenerJsCallClose("fnVodLinkSet(\""+vodurl+"\")");
		}
	}
}
function fnVodDelete(){
	fnAPPopenerJsCallClose("fnVodDelSet(\"\")");
}

</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<form name="vod" method="post" onsubmit="return false;">
		<div class="content bgGry">
			<h1 class="hidden">동영상 삽입</h1>
			<div class="vodAdd">
				<div class="linkInsert">
					<textarea rows="5" name="vodlink" id="vodlink" placeholder="동영상 링크를 입력해주세요."><%=param%></textarea>
				</div>
				<% If param<>"" Then %>
				<button type="button" class="btnM1 btnGry tMar1r" onclick="fnVodDelete()">동영상 삭제</button>
				<% Else %>
				<div class="linkNotice">
					<p class="fs1-5r">Youtube, Vimeo만 지원합니다.</p>
					<p class="tMar1-5r">동영상 링크를 복사해서 <br />붙여넣기 해주시면 동영상이 연결됩니다.</p>
				</div>
				<% End If %>
			</div>
		</div>
		</form>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->