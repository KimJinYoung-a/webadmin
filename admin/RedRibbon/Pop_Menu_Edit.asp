<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/RedRibbon/redRibbonManagerCls.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>
<body>

<%

dim Depth,cdL,cdM,cdS,SortMethod,MnL_viewType
Depth = request("depth")
cdL= request("cdL")
cdM= request("cdM")
cdS= request("cdS")

dim objView

set objView = new giftManagerView
objView.getMenuView cdL,cdM,cdS

if SortMethod ="" then SortMethod ="cashHigh"


dim ECodeNm , EOrderNo

SELECT CASE DEPTH
	CASE "L"
		
		ECodeNm 	= objView.LCodeNm
		EOrderNo 	= objView.OrderNo
	CASE "M"
		
		ECodeNm 	= objView.MCodeNm
		EOrderNo 	= objView.OrderNo
	CASE "S"
		
		ECodeNm 	= objView.SCodeNm
		EOrderNo 	= objView.OrderNo
END SELECT 

dim imageName,imageMain


SELECT CASE DEPTH
	CASE "L"
		imageName="Menu" & cdL 
	CASE "M"
		imageName="MidTop" & cdL & "-" & cdM 
	CASE "S"
		imageName="Mania" & cdL & "-" & cdM & "-" & cdS 
		imageMain="ManiaMain" & cdL & "-" & cdM & "-" & cdS 
END SELECT

'response.write imageName

%>
<script language="javascript">
function subchk(){
	if (isNaN(UpdateFRM.viewidx.value)) {
		alert('숫자만 입력가능합니다');
		return false;
	}
}

function popInputImg(ImgNm,Frmvalue){
	var pop = window.open('Pop_Menu_InputImage.asp?imageName=' + ImgNm +'&Frmvalue=' + Frmvalue ,'pp','width=300,height=150,resizable=yes');
}

function searchCashEdit(cdL,cdM,cdS){
	var pop = window.open('Pop_Menu_CashEdit.asp?cdL=' + cdL + '&cdM=' + cdM + '&cdS=' + cdS,'pp','width=500,height=300,status=yes,resizable=yes');
}

document.domain = "10x10.co.kr";

window.resizeTo(430,400);
</script>

<table width="400" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="UpdateFRM" action="Menu_Process.asp" target="" onSubmit="return subchk();">
	<input type="hidden" name="mode" value="edit">
	<input type="hidden" name="Depth" value="<%= Depth %>">
	<input type="hidden" name="LCode" size="4" value="<%= objView.LCode %>" />
	<input type="hidden" name="MCode" size="4" value="<%= objView.MCode %>" />
	<input type="hidden" name="SCode" size="4" value="<%= objView.SCode %>" />
	<tr>
		<td width="130" bgcolor="#FFFFFF"></td>
		<td bgcolor="#FFFFFF"></td>
	</tr>
<% IF objView.LCode <>"" then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">대 카테고리</td>
		<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.LCode %></font>] <%= objView.LCodeNm %>
	</tr>
<% END IF %>
<% IF objView.MCode <>"" then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">중 카테고리</td>
		<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.MCode %></font>] <%= objView.MCodeNm %>
	</tr>
<% END IF %>
<% IF objView.SCode <>"" then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">소 카테고리</td>
		<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.SCode %></font>] <%= objView.SCodeNm %>
	</tr>
<% END IF %>
	<tr>
		<td colspan="2" height="20" bgcolor="#FFFFFF" align="right"><input type="button" class="button" value="검색가격관리" onclick="searchCashEdit('<%= objView.LCode %>','<%= objView.MCode %>','<%= objView.SCode %>')"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">순서</td>
		<td bgcolor="#FFFFFF"><input type="text" size="4" name="OrderNo" value="<%= EOrderNo %>">(0 ~ 99)</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">카테고리명</td>
		<td bgcolor="#FFFFFF"><input type="text" size="16" name="CodeNm" value="<%= ECodeNm %>"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">정렬순서</td>
		<td bgcolor="#FFFFFF">
			<select name="SortMethod">
				<option value="cashHigh" <% if objView.SortMethod="cashHigh" then response.write "selected" %>>가격순(높은순)</option>
				<option value="cashLow" <% if objView.SortMethod="cashLow" then response.write "selected" %>>가격순(낮은순)</option>
				<option value="itemidHigh" <% if objView.SortMethod="itemidHigh" then response.write "selected" %>>상품번호순(높은순)</option>
				<option value="itemidLow" <% if objView.SortMethod="itemidLow" then response.write "selected" %>>상품번호순(낮은순)</option>
				<option value="OrderNo" <% if objView.SortMethod="OrderNo" then response.write "selected" %>>지정번호순</option>
				<option value="ItemScore" <% if objView.SortMethod="ItemScore" then response.write "selected" %>>인기상품순</option>
			</select>
		</td>
	</tr>

<!-- Depth 별 설정 -->
<% SELECT CASE DEPTH %>
<% CASE "L" %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">표시형식</td>
		<td bgcolor="#FFFFFF">
			<select name="ListType">
				<option value="list" <% if objView.ListType="list" then response.write "selected" %>>상품리스트</option>
				<option value="wish" <% if objView.ListType="wish" then response.write "selected" %>>위시리스트</option>
				<option value="mania" <% if objView.ListType="mania" then response.write "selected" %>>매니아 가이드</option>
				<option value="event" <% if objView.ListType="event" then response.write "selected" %>>이벤트</option>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center"  align="center">메뉴 이미지<br>(활성)</td>
		<td bgcolor="#FFFFFF"><input type="text" name="LCodeImgON" id="LCodeImgON" value="<%= objView.LCodeImgON %>" size="25" /><input type="button" class="button" value="이미지넣기" onclick="popInputImg('<%= imageName %>on','LCodeImgON');"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center"  align="center">메뉴 이미지<br>(비활성)</td>
		<td bgcolor="#FFFFFF"><input type="text" name="LCodeImgOFF" id="LCodeImgOFF" value="<%= objView.LCodeImgOFF %>" size="25" /><input type="button" class="button" value="이미지넣기" onclick="popInputImg('<%= imageName %>','LCodeImgOFF');"></td>
	</tr>
<% CASE "M" %>
	<input type="hidden" name="ListType" size="8" value="<%= objView.ListType %>" />
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">상단 이미지</td>
		<td bgcolor="#FFFFFF"><input type="text" name="MCodeTopImg" id="MCodeTopImg" value="<%= objView.MCodeTopImg %>" size="25" /><input type="button" class="button" value="이미지넣기" onclick="popInputImg('<%= imageName %>','MCodeTopImg');"></td>
	</tr>
<% CASE "S" %>
	<input type="hidden" name="ListType" size="8" value="<%= objView.ListType %>" />
	<% IF objView.ListType="mania" THEN%>
	<!--
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">매니아가이드<br>이미지</td>
		<td bgcolor="#FFFFFF"><input type="text" name="GuideListImg" id="GuideListImg" value="<%'= objView.GuideListImg %>" size="25" /><input type="button" class="button" value="이미지넣기" onclick="popInputImg('<%'= imageName %>','GuideListImg');"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">매니아가이드<br>메인이미지</td>
		<td bgcolor="#FFFFFF"><input type="text" name="GuideTopImg" id="GuideTopImg" value="<%'= objView.GuideTopImg %>" size="25" /><input type="button" class="button" value="이미지넣기" onclick="popInputImg('<%'= imageMain %>','GuideTopImg');"></td>
	</tr>
	-->
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">매니아가이드<br>on 이미지</td>
		<td bgcolor="#FFFFFF"><input type="text" name="guideonimg" id="guideonimg" value="<%= objView.GuideListImg %>" size="25" /><input type="button" class="button" value="이미지넣기" onclick="popInputImg('<%= imageName %>on','guideonimg');"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">매니아가이드<br>off 이미지</td>
		<td bgcolor="#FFFFFF"><input type="text" name="guideoffimg" id="guideoffimg" value="<%= objView.GuideListImg %>" size="25" /><input type="button" class="button" value="이미지넣기" onclick="popInputImg('<%= imageName %>','guideoffimg');"></td>
	</tr>
	<% END IF %>
<% END SELECT %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">사용여부</td>
		<td bgcolor="#FFFFFF">
			
			<input type="radio" name="isusing" value="Y" <% IF objView.IsUsing="Y" Then response.write "checked" %>> 사용 
			<input type="radio" name="isusing" value="N" <% IF objView.IsUsing="N" Then response.write "checked" %>>사용안함</td>
	</tr>
	

	
	<tr>
		<td bgcolor="#FFFFFF" colspan="2" align="center"><input type="submit" class="button" value="적용"></td>
	</tr>
	</form>
</table>

<% set objView = nothing %>
</body>
</html>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->