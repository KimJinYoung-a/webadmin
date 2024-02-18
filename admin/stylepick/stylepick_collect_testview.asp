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
dim cd1 , cd2 ,i ,evtidx ,oevent ,banner_img ,oitem , page ,trgubun ,maketr,PSize
	cd1 = requestcheckvar(request("cd1"),10)
	cd2 = requestcheckvar(request("cd2"),10)	
	evtidx = requestcheckvar(request("evtidx"),10)
	page = request("page")
	if page = "" then page = 1
	trgubun = 0
	maketr = 0
	PSize = 22
	
'/파라메타 아예 없는경우 스타일 카테고리 기본값 지정	
if cd1="" and evtidx="" then
	cd1 = pageloadevent(cd1)
end if

'/evtidx가 있는경우 해당 내역 가져옴 '/evtidx가 없는경우 해당 스타일 기획전에  오픈이상내역중 최근 내역 순으로 가져옴
set oevent = new cstylepick	
	oevent.frectcd1 = cd1
	oevent.frectevtidx = evtidx
	oevent.fnGetEvent_item
	
	if oevent.ftotalcount < 1 then
		response.write "<script language='javascript'>"
		response.write "	alert('해당 스타일에 등록되어 있는 기획전이 없습니다');"
		'response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.end
	else
		banner_img = oevent.foneitem.fbanner_img
		evtidx = oevent.foneitem.fevtidx
		cd1 = oevent.foneitem.fcd1

		'/기획전 상품리스트
		set oitem = new cstylepick
			oitem.FPageSize = PSize
			oitem.FCurrPage = page	
			oitem.frectevtidx = evtidx
			oitem.GetevtItemList

	end if
set oevent = nothing
%>

<script language="javascript">

	function jsGoPage(page){
		document.frm.page.value = page;
		document.frm.submit();
	}

</script>

<link href="<%=wwwUrl%>/lib/css/2011ten.css" rel="stylesheet" type="text/css">

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

<!----- 스타일진 리스트 START ------>
<table width="960" border=0 align="center" cellpadding="0" cellspacing="0" style="margin-bottom:20px;border-bottom:1px solid #e5e5e5;">
<form name="frm" method="get">
<input type="hidden" name="cd1" value="<%=cd1%>">
<input type="hidden" name="cd2" value="<%=cd2%>">
<input type="hidden" name="PSize" value="<%=PSize%>">
<input type="hidden" name="page" value="">
<tr height=20><td></td></tr>
<tr>
	<!----- 왼쪽 타이틀 ----->
	<td colspan="2" rowspan="2" valign="top" style="border-bottom:1px solid #e5e5e5;padding:20px 0 0 15px;">
		<table border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td>
				<img src="http://fiximage.10x10.co.kr/web2011/stylezine/left_title_<%=cd1%>.gif" width="285" height="105"></td>
		</tr>
		<tr>
			<td style="padding:35px 0 0 12px;">
				<%
				dim ocatecd2
				
				'/분류 리스트 카운트
				set ocatecd2 = new cstylepickMenu
					ocatecd2.frectcd1 = cd1
					ocatecd2.fstylepick_cd2_count
				%>
				<table border="0" cellspacing="0" cellpadding="0">
				<tr>	
					<% if Request.ServerVariables("SCRIPT_NAME") = "/stylepick/stylepick_collect_testview.asp" then %>
						<td width="18" height="25"><img src="http://fiximage.10x10.co.kr/web2011/stylezine/list_category_off.gif"></td>
						<td>ALL (<%=ocatecd2.fitemallcount%>)</td>
					<% else %>
						<td width="18" height="25"><img src="http://fiximage.10x10.co.kr/web2011/stylezine/list_category<% if cd2="" then %>_on<%else%>_off<%end if%>.gif"></td>
						<td>ALL (<%=ocatecd2.fitemallcount%>)</td>
					<% end if %>
				</tr>
				<% if ocatecd2.fresultcount >0 then %>
				<% for i = 0 to ocatecd2.fresultcount - 1 %>
					<% 
					'/상품수량이 있는경우만 노출
					if ocatecd2.FItemList(i).fitemcount > 0 then
					%>
					<tr>
						<td height="25"><img src="http://fiximage.10x10.co.kr/web2011/stylezine/list_category<% if cd2=ocatecd2.FItemList(i).fcd2 then %>_on<%else%>_off<%end if%>.gif"></td>
						<td><%=ocatecd2.FItemList(i).fcatename%> (<%=ocatecd2.FItemList(i).fitemcount%>)</td>
					</tr>
					<% end if %>
				<% next %>
				<% end if %>
				</table>
				<% set ocatecd2 = nothing %>
			</td>
		</tr>
		</table>
	</td>
	<!----- 상단 기획 타이틀 ----->
	<td height="195" colspan="4" align="right" valign="top" style="border-left:1px solid #e5e5e5;border-bottom:1px solid #e5e5e5;" width=638><img src="<%=banner_img%>"> </td>
</tr>
<% if oitem.fresultcount > 0 then %>
<tr>
	<%
	for i = 0 to oitem.fresultcount -1
	
	maketr = maketr + 1
	%>	
	<td width="159" height="195" align="center" valign="top" class="style_list">
		<table width="120" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td>
				<img src="<%= oitem.FItemList(i).Flistimage120 %>" width="120" height="120"></td>
		</tr>
		<tr>
			<td align="center" valign="top" style="padding-top:7px;">			
				<%= chrbyte(oitem.FItemList(i).fitemname,38,"Y") %></td>
		</tr>
		</table>
	</td>
	<%
	'//첫줄일경우 td 4번째에서 줄내림
	if trgubun = 0 then
		if maketr = 4 then
				response.write "</tr><tr><td width=159 height=195>&nbsp;</td>"
			maketr = 0
			trgubun = trgubun + 1
		end if
		
	'//첫줄이 아닐경우 td 5번째에서 줄내림		
	else	
		if maketr = 5 then
			response.write "</tr><tr><td width=159 height=195>&nbsp;</td>"
			maketr = 0
			trgubun = trgubun + 1
		end if
	end if
	
	next			
		
	'/첫줄에서 끝날경우 4칸으로 페이징 자리처리, 페이징이 colspan=2 이기때문에 페이지 맨끝이 한줄이라면 한줄내리고 공란처리
	if trgubun = 0 then
		if oitem.fresultcount mod 4 = 1 then response.write "<td width=159 height=195 class='style_list'>&nbsp;</td>"		
		if oitem.fresultcount mod 4 = 3 then response.write "<td width=159 height=195 class='style_list'>&nbsp;</td></tr><tr><td width=159 height=195>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td>"
	
	'/첫줄이 아닐경우 첫줄 라인수인 4를 빼고 5칸을 기준으로 페이징 자리처리, 페이징이 colspan=2 이기때문에 페이지 맨끝이 한줄이라면 한줄내리고 공란처리
	else	
		if (oitem.fresultcount-4) mod 5 = 0 then response.write "<td width=159 height=195 class='style_list'>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td>"
		if (oitem.fresultcount-4) mod 5 = 1 then response.write "<td width=159 height=195 class='style_list'>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td>"
		if (oitem.fresultcount-4) mod 5 = 2 then response.write "<td width=159 height=195 class='style_list'>&nbsp;</td>"		
		if (oitem.fresultcount-4) mod 5 = 4 then response.write "<td width=159 height=195 class='style_list'>&nbsp;</td><tr><td width=159 height=195>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td>"
	end if
	%>
	<!---- 페이지 넘버링 ----->
	<td colspan="2" align="center" style="border-left:1px solid #e5e5e5;border-top:1px solid #e5e5e5;"><img src="http://fiximage.10x10.co.kr/web2011/stylezine/list_pagenum.gif" width="220" height="22"></td>
</tr>
<% else %>
<tr>
	<td align='center' class='style_list' valign='top'>
		검색 결과가 없습니다
	</td>
</tr>	
<% end if %>
</form>
</table>
<!----- 스타일진 리스트 END ------>

<% set oitem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->