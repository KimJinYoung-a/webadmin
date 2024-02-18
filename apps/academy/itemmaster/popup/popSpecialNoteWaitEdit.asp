<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/itemmaster/DIYitemCls.asp"-->
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"

Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 제작 특이사항 입력"
%>
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<%
Dim waititemid
waititemid = requestCheckVar(request("waititemid"),10)

Dim oitem, requirecontents
set oitem = new CWaitItemDetail
oitem.FRectDesignerID = request.cookies("partner")("userid")
if (waititemid<>"") then
oitem.WaitProductDetail(waititemid)
requirecontents=oitem.Frequirecontents
End If
%>
<script>
jQuery(document).ready(function(){
	//$("#requirecontents").keyup(function(){
		fnAPPShowRightConfirmBtns();
	//});
});
function fnAppCallWinConfirm(){
	if($("#requirecontents").val()==""){
		alert("제작 특이사항을 입력해주세요.");
	}
	else{
		document.sform.action="/apps/academy/itemmaster/popup/WaitDIYItemPopupDetailinfoEdit_Process.asp";
		document.sform.target="FrameCKP";
		document.sform.submit();		
	}
}
function fnDetailInfoEnd(){
    var irequirecontents = Base64.encode($("#requirecontents").val());
	fnAPPopenerJsCallClose("fnMakeUnusualSet(\""+irequirecontents+"\")");
	
	//원코드
	//fnAPPopenerJsCallClose("fnMakeUnusualSet(\""+$("#requirecontents").val()+"\")");
	//var irequirecontents = encSpecialCharNativeFun($("#requirecontents").val());
    //alert('fnDetailInfoEnd-'+irequirecontents)
}

</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<form method="post" name="sform" autocomplete="off">
		<input type="hidden" name="waititemid" value="<%=waititemid%>">
		<div class="content bgGry">
			<h1 class="hidden">제작 특이사항 입력</h1>
			<div class="spcNote">
				<div class="linkInsert">
					<textarea rows="5" name="requirecontents" id="requirecontents" placeholder="특이사항이 있을 경우 입력해주세요"><%=requirecontents%></textarea>
				</div>
				<div class="linkNotice">
					<p class="fs1-5r">고객이 알아야할 <br />제작 특이사항을 입력해주세요.</p>
				</div>
			</div>
		</div>
		</form>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<%
set oitem = nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->