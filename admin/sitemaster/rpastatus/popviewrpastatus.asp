<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : RPA 상태 내용 확인
' Hieditor : 2020.08.12 원승현 추가
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/rpastatus/rpastatuscls.asp"-->
<%
Dim i, mode
Dim idx
dim oRpaStatusView, loginUserId

idx = requestCheckvar(request("idx"), 50)

loginUserId = session("ssBctId")

if Trim(idx) = "" then
	response.write "<script>alert('정상적인 경로로 접근해주세요.');window.close();</script>"
	response.end
end If

'// rpastatus View 데이터를 가져온다.
set oRpaStatusView = new CgetRpaStatus
	oRpaStatusView.FRectIdx = idx
	oRpaStatusView.getRpaStatusview()
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
body {background-color:#fff;}
</style>
</head>
<body>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script type="text/javascript" src="/js/jquery.swiper-3.3.1.min.js"></script>
<script type="text/javascript" src="/js/tag-it.min.js"></script>
<script type='text/javascript'>
document.domain = "10x10.co.kr";
</script>
<%' 팝업 사이즈 : 750*800 %>
	<div class="popWinV17">
		<h1><%=oRpaStatusView.FOneRpaStatus.Ftitle%></h1>
		<div class="popContainerV17 pad30">
			<table class="tbType1 writeTb">
				<colgroup>
					<col width="18%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>번호(idx) <strong class="cRd1"></strong></div></th>
					<td><%=oRpaStatusView.FOneRpaStatus.Fidx%></td>
				</tr>
				<tr>
					<th><div>구분</div></th>
					<td>
                        <%=getRpaTypeName(oRpaStatusView.FOneRpaStatus.Ftype)%>
					</td>
				</tr>
				<tr>
					<th><div>내용</div></th>
					<td>
                        <div style="text-align:left;width:600px;white-space:pre-line;"><%=replace(oRpaStatusView.FOneRpaStatus.Fcontents,chr(13)&chr(10),"<br>")%></div>
					</td>
				</tr>                

				<tr>
					<th><div>성공여부</div></th>
					<td>
                        <%=getRpaIsSuccessName(oRpaStatusView.FOneRpaStatus.FisSuccess)%>
					</td>
				</tr>
				<tr>
					<th><div>등록일</div></th>
					<td>
						<span class="tPad05 col2"><%=oRpaStatusView.FOneRpaStatus.Fregdate%></span>
					</td>
				</tr>
				</tbody>
			</table>
		</div>
	</div>
</body>
</html>
<%
	set oRpaStatusView = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
