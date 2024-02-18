<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 스타일픽 관리
' Hieditor : 2011.04.06 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->
<%
dim cd1 , cd2 ,i
dim mainidx ,omain ,opageing, cd1pre ,cd1next ,mainimage ,mainimagelink
dim mainidxpre , mainidxnext , mainidxmin , mainidxmax ,Num ,numimg
	mainidx = requestcheckvar(request("mainidx"),10)
	cd1 = requestcheckvar(request("cd1"),10)
	cd2 = requestcheckvar(request("cd2"),10)

'/파라메타 아예 없는경우 스타일 카테고리 기본값 지정	
if cd1="" and mainidx="" then	
	cd1 = pageloadevent(cd1)
end if	

'/mainidx가 있는경우 해당 내역 가져옴 '/mainidx가 없는경우 해당 스타일에 오픈이상내역중 최근 내역 순으로 가져옴
set omain = new cstylepick
	omain.frectcd1 = cd1
	omain.frectmainidx = mainidx
	omain.frectview = "ON"
	omain.fnGetmain_item
	
	if omain.ftotalcount < 1 then
		response.write "<script language='javascript'>"
		response.write "	alert('해당 스타일에 등록되어 있는 내역이 없습니다');"
		'response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.end
	else
		cd1pre = omain.foneitem.fcd1pre
		cd1next = omain.foneitem.fcd1next
		mainimage = omain.foneitem.fmainimage
		mainimagelink = omain.foneitem.fmainimagelink
		mainidx = omain.foneitem.fmainidx
		cd1 = omain.foneitem.fcd1
		
		'/페이징 현재 페이지 일렬번호와  첫페이지 , 이전페이지 ,다음페이지 ,맨끝페이지 가져옴
		set opageing = new cstylepick
			opageing.frectcd1 = cd1
			opageing.frectmainidx = mainidx
			opageing.fnGetmain_pageing
		
			if opageing.ftotalcount > 0 then
				Num = opageing.foneitem.fRowNum		'/현재페이지 넘버
				mainidxpre = opageing.foneitem.fmainidxpre	'/이전페이지
				mainidxnext = opageing.foneitem.fmainidxnext	'/다음페이지
				mainidxmin = opageing.foneitem.fmainidxmin	'/첫페이지
				mainidxmax = opageing.foneitem.fmainidxmax	'/맨끝페이지			
				
				'/현재 페이지 이미지 셋팅
				for i = 0 to len(num) -1				
					numimg = numimg & "<img src='http://fiximage.10x10.co.kr/web2011/stylezine/no_num_"& mid(num,i+1,1) &".gif'>"
				next
			end if	
		set opageing = nothing
	end if
set omain = nothing
%>

<link href="<%=wwwUrl%>/lib/css/2011ten.css" rel="stylesheet" type="text/css">

<script language="javascript">
	
	function mainidxpre(mainidxpre){
		if (mainidxpre==''){
			alert('이전 StylePick 이 없습니다');
			return;
		}
		
		location.href="?cd1=<%=cd1%>&mainidx="+mainidxpre;
	}

	function mainidxnext(mainidxnext){
		if (mainidxnext==''){
			alert('다음 StylePick 이 없습니다');
			return;
		}
		
		location.href="?cd1=<%=cd1%>&mainidx="+mainidxnext;
	}
	
</script>

<!----- 스타일픽 스타일 카테고리 ------>
<table width="960" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
	<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td width="140" style="border-right:1px solid #e5e5e5;"><img src="http://fiximage.10x10.co.kr/web2011/header/top_logo.gif" width="140" height="120"></td>
			<td align="right" valign="top" style="padding-top:51px;"><img src="http://fiximage.10x10.co.kr/web2011/header/stylepick_title.gif" width="365" height="53"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td height="33" align="right" style="border-top:3px solid #dadada;border-bottom:3px solid #dadada;padding-right:7px;"> 
		<!----- 스타일픽 상단메뉴 START ----->
		<%
		dim objcd1
		set objcd1 = new cstylepickMenu
			objcd1.frectisusing = "Y"
			objcd1.getstylepick_cate_cd1()
			
		if objcd1.fresultcount > 0 then
		%>		
		<table border=0 cellspacing=0 cellpadding=0>
		<tr>				
			<% for i = 0 to objcd1.fresultcount -1 %>		
			<td>
				<img src='http://fiximage.10x10.co.kr/web2011/header/stylepick_menu<%=objcd1.FItemList(i).fcd1%><%if cd1 = objcd1.FItemList(i).fcd1 then response.write "on" End if%>.gif'>
			</td>
			<% 
			if i+1 <> objcd1.fresultcount then response.write "<td><img src='http://fiximage.10x10.co.kr/web2011/header/stylepick_dot.gif'></td>"
		
			next
			%>
		</tr>
		</table>
		<%
		end if						
		set objcd1 = nothing
		%>
		<!----- 스타일픽 상단메뉴 END ----->
	</td>
</tr>
</table>

<!----- 스타일픽 메인 START ------>
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-bottom:20px;">
<tr height=20><td></td></tr>
<tr>
	<td>
	<!----- 스타일진 넘버 START ----->
	<div id="style_no" style="position:absolute; width:141px; margin-top:-31px">
	<table width="141" border="0" cellpadding="0" cellspacing="0" bgcolor="#777777">
	<tr>
		<td><img src="http://fiximage.10x10.co.kr/web2011/stylezine/no_btn_first.gif" width="26" height="31"></td>
		<td><img src="http://fiximage.10x10.co.kr/web2011/stylezine/no_btn_prev.gif" width="16" height="31"></td>
		<td width="57" align="center" style="padding-left:2px;"><img src="http://fiximage.10x10.co.kr/web2011/stylezine/no_txt.gif" width="24" height="10"><%=numimg%></td>
		<td><img src="http://fiximage.10x10.co.kr/web2011/stylezine/no_btn_next.gif" width="16" height="31"></td>
		<td><img src="http://fiximage.10x10.co.kr/web2011/stylezine/no_btn_end.gif" width="26" height="31"></td>
	</tr>
	</table>
	</div>
	<!----- 스타일진 넘버 END ----->
	<% if cd1pre <> "" then %><div id="m_left" style="position:absolute; width:56px; margin-top:530px; margin-left:5px;">
		<img src="http://fiximage.10x10.co.kr/web2011/stylezine/style_main_prev.png" width="56" height="56"></div><% end if %>
	<% if cd1next <> "" then %><div id="m_right" style="position:absolute; width:56px; margin-top:530px; margin-left:899px;">
		<img src="http://fiximage.10x10.co.kr/web2011/stylezine/style_main_next.png" width="56" height="56"></div><% end if %>
	<% if mainimage <> "" then %><img src="<%= mainimage %>" border="0" usemap="#Mapmainimage"><%end if %><%= mainimagelink %></td>
</tr>
</table>
<!----- 스타일진 메인 END ------>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->