<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - Q&A FILTER"
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<%
'전달값 처리
dim statediv
statediv = RequestCheckVar(request("statediv"),1)
If (statediv="") Then statediv="0"
%>
<script>
$(function() {
	// button tab
	$(".selectBtn button").click(function(){
		$(this).parent().parent().find("button").removeClass("selected");
		$(this).addClass("selected");
	});
});

function fnSearchFilterSelect(formname,formdata){
	if(formname!=''){
		eval("$('#"+formname+"')").val(formdata);
	}
}

function fnAppCallWinConfirm(){
	var jsontxt;
	jsontxt = JSON.stringify({"statediv":$("#statediv").val(),"filter":"Y"});
	jsontxt = Base64.encode(jsontxt);
	fnAPPopenerJsCallClose("fnSearchFilterSet(\"" + jsontxt + "\")");
}
</script>
</head>
<body>
<div class="wrap">
	<div class="container">
		<!-- content -->
		<div class="content">
			<h1 class="hidden">필터</h1>
			<form name="searchForm" id="searchForm" method="get" style="margin:0px;">
			<input type="hidden" name="statediv" id="statediv" value="<%=statediv%>">
			<div class="filterWrap">
				<dl class="dfCompos">
					<dt>구분</dt>
					<dd class="selectBtn">
						<ul>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(statediv="0","selected","")%>" onClick="fnSearchFilterSelect('statediv','0');">전체</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(statediv="N","selected","")%>" onClick="fnSearchFilterSelect('statediv','N');">답변중</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(statediv="Y","selected","")%>" onClick="fnSearchFilterSelect('statediv','Y');">답변완료</button></li>
						</ul>
					</dd>
				</dl>
			</div>
			</form>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<script type="text/javascript">
<!--
jQuery(document).ready(function(){
fnAPPShowRightConfirmBtns();
});
//-->
</script>