<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ��Ÿ���� ����
' Hieditor : 2011.04.06 �ѿ�� ����
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
	
'/�Ķ��Ÿ �ƿ� ���°�� ��Ÿ�� ī�װ� �⺻�� ����	
if cd1="" and evtidx="" then
	cd1 = pageloadevent(cd1)
end if

'/evtidx�� �ִ°�� �ش� ���� ������ '/evtidx�� ���°�� �ش� ��Ÿ�� ��ȹ����  �����̻󳻿��� �ֱ� ���� ������ ������
set oevent = new cstylepick	
	oevent.frectcd1 = cd1
	oevent.frectevtidx = evtidx
	oevent.fnGetEvent_item
	
	if oevent.ftotalcount < 1 then
		response.write "<script language='javascript'>"
		response.write "	alert('�ش� ��Ÿ�Ͽ� ��ϵǾ� �ִ� ��ȹ���� �����ϴ�');"
		'response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.end
	else
		banner_img = oevent.foneitem.fbanner_img
		evtidx = oevent.foneitem.fevtidx
		cd1 = oevent.foneitem.fcd1

		'/��ȹ�� ��ǰ����Ʈ
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

<!----- ��Ÿ���� ��Ÿ�� ī�װ� ------>
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
		<!----- ��Ÿ���� ��ܸ޴� START ----->
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
		<!----- ��Ÿ���� ��ܸ޴� END ----->
	</td>
</tr>
</table>

<!----- ��Ÿ���� ����Ʈ START ------>
<table width="960" border=0 align="center" cellpadding="0" cellspacing="0" style="margin-bottom:20px;border-bottom:1px solid #e5e5e5;">
<form name="frm" method="get">
<input type="hidden" name="cd1" value="<%=cd1%>">
<input type="hidden" name="cd2" value="<%=cd2%>">
<input type="hidden" name="PSize" value="<%=PSize%>">
<input type="hidden" name="page" value="">
<tr height=20><td></td></tr>
<tr>
	<!----- ���� Ÿ��Ʋ ----->
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
				
				'/�з� ����Ʈ ī��Ʈ
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
					'/��ǰ������ �ִ°�츸 ����
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
	<!----- ��� ��ȹ Ÿ��Ʋ ----->
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
	'//ù���ϰ�� td 4��°���� �ٳ���
	if trgubun = 0 then
		if maketr = 4 then
				response.write "</tr><tr><td width=159 height=195>&nbsp;</td>"
			maketr = 0
			trgubun = trgubun + 1
		end if
		
	'//ù���� �ƴҰ�� td 5��°���� �ٳ���		
	else	
		if maketr = 5 then
			response.write "</tr><tr><td width=159 height=195>&nbsp;</td>"
			maketr = 0
			trgubun = trgubun + 1
		end if
	end if
	
	next			
		
	'/ù�ٿ��� ������� 4ĭ���� ����¡ �ڸ�ó��, ����¡�� colspan=2 �̱⶧���� ������ �ǳ��� �����̶�� ���ٳ����� ����ó��
	if trgubun = 0 then
		if oitem.fresultcount mod 4 = 1 then response.write "<td width=159 height=195 class='style_list'>&nbsp;</td>"		
		if oitem.fresultcount mod 4 = 3 then response.write "<td width=159 height=195 class='style_list'>&nbsp;</td></tr><tr><td width=159 height=195>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td>"
	
	'/ù���� �ƴҰ�� ù�� ���μ��� 4�� ���� 5ĭ�� �������� ����¡ �ڸ�ó��, ����¡�� colspan=2 �̱⶧���� ������ �ǳ��� �����̶�� ���ٳ����� ����ó��
	else	
		if (oitem.fresultcount-4) mod 5 = 0 then response.write "<td width=159 height=195 class='style_list'>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td>"
		if (oitem.fresultcount-4) mod 5 = 1 then response.write "<td width=159 height=195 class='style_list'>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td>"
		if (oitem.fresultcount-4) mod 5 = 2 then response.write "<td width=159 height=195 class='style_list'>&nbsp;</td>"		
		if (oitem.fresultcount-4) mod 5 = 4 then response.write "<td width=159 height=195 class='style_list'>&nbsp;</td><tr><td width=159 height=195>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td><td width=159 height=195 class='style_list'>&nbsp;</td>"
	end if
	%>
	<!---- ������ �ѹ��� ----->
	<td colspan="2" align="center" style="border-left:1px solid #e5e5e5;border-top:1px solid #e5e5e5;"><img src="http://fiximage.10x10.co.kr/web2011/stylezine/list_pagenum.gif" width="220" height="22"></td>
</tr>
<% else %>
<tr>
	<td align='center' class='style_list' valign='top'>
		�˻� ����� �����ϴ�
	</td>
</tr>	
<% end if %>
</form>
</table>
<!----- ��Ÿ���� ����Ʈ END ------>

<% set oitem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->