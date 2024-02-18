<%@ codepage="65001" language=vbscript %>
<% option explicit %>
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 필터"

Dim odiv
odiv = requestCheckVar(request("odiv"),1)

'전달값 처리
dim sdiv, sortupdown, statediv, edate, sdate
sortupdown = RequestCheckVar(request("sortupdown"),1)
sdiv = RequestCheckVar(request("sdiv"),10)
sdate = RequestCheckVar(request("sdate"),10)
edate = RequestCheckVar(request("edate"),10)
statediv = RequestCheckVar(request("statediv"),1)
If (sortupdown="") Then sortupdown="u"
If (sdiv="") Then sdiv="Reg"
If (statediv="") Then statediv="0"
%>
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jqueryui/css/jquery-ui.css" />
<script>

function fnSearchFilterSelect(formname,formdata){
	if(formname!=''){
		eval("$('#"+formname+"')").val(formdata);
	}
}

function fnSearchSortFilterSelect(formdata,btnnum){
	if($("#sdiv").val()==formdata){
		if($("#ssort").val()=="u"){
			$("#ssort").val("d");
			eval("$('#btn"+btnnum +"')").removeClass("srtUp");
			eval("$('#btn"+btnnum +"')").addClass("srtDown");
		}else{
			$("#ssort").val("u");
			eval("$('#btn"+btnnum +"')").removeClass("srtDown");
			eval("$('#btn"+btnnum +"')").addClass("srtUp");
		}
	}else{
		$("#sdiv").val(formdata);
		$("#ssort").val("u");
		eval("$('#btn"+btnnum +"')").removeClass("srtDown");
		eval("$('#btn"+btnnum +"')").addClass("srtUp");
	}
}

function fnAppCallWinConfirm(){
	
	var jsontxt;
	//jsontxt = $("#statediv").val() + "," + $("#sdiv").val() + "," + $("#ssort").val() + "," + $("#datepicker").val() + "," + $("#datepicker2").val() + ",Y";
	jsontxt = JSON.stringify({"statediv":$("#statediv").val(),"sdiv":$("#sdiv").val(),"ssort":$("#ssort").val(),"startdate":$("#datepicker").val(),"enddate":$("#datepicker2").val(),"filter":"Y"});
	jsontxt = Base64.encode(jsontxt);
	//alert(jsontxt);
	fnAPPopenerJsCallClose("fnSearchFilterSet(\"" + jsontxt + "\")");
}

jQuery(document).ready(function(){
	fnAPPShowRightConfirmBtns();

	// button tab
	$(".selectBtn button").click(function(){
		$(this).parent().parent().find("button").removeClass("selected");
		$(this).addClass("selected");
	});
	// datepicker
	$("#datepicker").datepicker({
		showOn:"both",
		buttonImage: "http://image.thefingers.co.kr/apps/2016/ico_calendar.png",
		buttonImageOnly:true,
		buttonText:"출고 예정일을 선택해주세요",
		dateFormat:"yy-mm-dd"
	});
	$("#datepicker2").datepicker({
		showOn:"both",
		buttonImage: "http://image.thefingers.co.kr/apps/2016/ico_calendar.png",
		buttonImageOnly:true,
		buttonText:"출고 예정일을 선택해주세요",
		dateFormat:"yy-mm-dd"
	});
});
</script>
</head>
<body>
<div class="wrap">
	<div class="container">
		<!-- content -->
		<div class="content">
			<h1 class="hidden">필터</h1>
			<!-- for dev msg : 레이어 작업시 이 부분만 가져가면 됩니다.-->
			<form name="searchForm" id="searchForm" method="get" style="margin:0px;">
			<input type="hidden" name="sdiv" id="sdiv" value="Reg">
			<input type="hidden" name="ssort" id="ssort" value="<%=sortupdown%>">
			<input type="hidden" name="statediv" id="statediv" value="<%=statediv%>">
			<div class="filterWrap">
			<div class="filterWrap">
				<dl class="dfCompos">
					<dt>기간</dt>
					<dd class="selectBtn list">
						<div>
							<div class="datePickWrap"><input type="text" id="datepicker" value="<%=sdate%>" readonly class="datepicker" placeholder="시작일"></div>
						</div>
						<div style="width:8%" class="ct cGy4 fs1-5r"> ~ </div>
						<div>
							<div class="datePickWrap"><input type="text" id="datepicker2" value="<%=edate%>" readonly class="datepicker2" placeholder="종료일"></div>
						</div>
					</dd>
				</dl>
				<dl class="dfCompos">
					<dt>구분</dt>
					<dd class="selectBtn">
						<% If odiv = "S" Then %>
						<ul>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(statediv="0","selected","")%>" onClick="fnSearchFilterSelect('statediv','0');">전체</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(statediv="1","selected","")%>" onClick="fnSearchFilterSelect('statediv','1');">확인대기</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(statediv="2","selected","")%>" onClick="fnSearchFilterSelect('statediv','2');">주문취소</button></li>
						</ul>
						<% ElseIf odiv = "D" Then %>
						<ul>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(statediv="0","selected","")%>" onClick="fnSearchFilterSelect('statediv','0');">전체</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(statediv="1","selected","")%>" onClick="fnSearchFilterSelect('statediv','1');">배송대기</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(statediv="2","selected","")%>" onClick="fnSearchFilterSelect('statediv','2');">미출고</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(statediv="4","selected","")%>" onClick="fnSearchFilterSelect('statediv','4');">주문취소</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(statediv="5","selected","")%>" onClick="fnSearchFilterSelect('statediv','5');">출고완료</button></li>
						</ul>
						<% Else %>
						<ul>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(statediv="0","selected","")%>" onClick="fnSearchFilterSelect('statediv','0');">전체</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(statediv="1","selected","")%>" onClick="fnSearchFilterSelect('statediv','1');">미처리</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(statediv="2","selected","")%>" onClick="fnSearchFilterSelect('statediv','2');">처리완료</button></li>
						</ul>
						<% End If %>
					</dd>
				</dl>
				<dl class="dfCompos">
					<dt>정렬기준</dt>
					<dd class="selectBtn">
						<% If odiv = "S" Then %>
						<ul>
							<li class="grid1"><button type="button" class="btnM1 btnGry <%=chkIIF(sdiv="Reg","selected","")%>" onClick="fnSearchSortFilterSelect('Reg',1);"><span class="sort <%=chkIIF(sdiv="Reg" and sortupdown="d","srtDown","srtUp")%>" id="btn1">등록순</span></button></li><!-- for dev msg : 한번 클릭시 button 에 selected 붙여주시고 한번더 클릭하면 span 태그의 srtUp/srtDown 이 토글되면 됩니다 -->
						</ul>
						<% ElseIf odiv = "D" Then %>
						<ul>
							<li class="grid1"><button type="button" class="btnM1 btnGry <%=chkIIF(sdiv="Reg","selected","")%>" onClick="fnSearchSortFilterSelect('Reg',1);"><span class="sort <%=chkIIF(sdiv="Reg" and sortupdown="d","srtDown","srtUp")%>" id="btn1">주문일</span></button></li>
						</ul>
						<% Else %>
						<ul>
							<li class="grid1"><button type="button" class="btnM1 btnGry <%=chkIIF(sdiv="Reg","selected","")%>" onClick="fnSearchSortFilterSelect('Reg',1);"><span class="sort <%=chkIIF(sdiv="Reg" and sortupdown="d","srtDown","srtUp")%>" id="btn1">접수일</span></button></li>
						</ul>
						<% End If %>
					</dd>
				</dl>
			</div>
			</form>
			<!-- //for dev msg : 레이어 작업시 이 부분만 가져가면 됩니다.-->
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->