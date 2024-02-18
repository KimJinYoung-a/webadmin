<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [업체] HELP
' History : 2014.06.11 정윤정 생성
'###########################################################
%>
<!-- 필수-------------------------------------->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/partner/incSessionDesigner.asp" --> 
<!-- //필수-------------------------------------->
<!-- #include virtual="/partner/lib/adminHead.asp" --><!--html-->   
 
<%
dim menupos : menupos = requestCheckVar(Request("menupos"),10)  '메뉴번호 
Dim ssBctDiv_UsercDiv :ssBctDiv_UsercDiv = session("ssBctDiv")&"_"&session("ssUserCDiv")
Dim conHelp : conHelp			= Application("topHelp"&ssBctDiv_UsercDiv)
dim conMNum, comSMNum
if menupos <> "" then  '-- 메뉴번호 있을때만 서브메뉴 보여준다.
			conMNum = split(menupos,"^")(0)
			comSMNum = split(menupos,"^")(1) 
end if 
%>
</head>
<body>
<div>
   
				 
<div class="popupWrap">
	<div class="popHead">
		<h1><img src="/images/partner/pop_admin_logo.gif" alt="10x10" /></h1>
		<p class="btnClose"><input type="image" src="/images/partner/pop_admin_btn_close.gif" alt="창닫기" onclick="window.close();" /></p>
	</div>
	<div class="popContent scrl">
		<div class="contTit bgNone"> 
			<h2>HELP</h2> 
		</div>
		<div class="cont">
			 <div class="helpCont">
				<%=conHelp(conMNum,comSMNum)%>
			</div>  
		</div>
	</div>
</div>
</body>
</html>

 

<!-- #include virtual="/lib/db/dbclose.asp" -->