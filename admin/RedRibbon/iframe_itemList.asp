<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
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
<!--
<link rel="stylesheet" href="/bct.css" type="text/css">-->
</head>
<body topmargin="0" onScroll="SetDivScroll();">

<%

function getYNColor(v)
	if v="Y" then
		getYNColor =v
	else
		getYNColor ="<font color='red'>" & v & "</font>"
	end if
end function


dim cdL, cdM, cdS,SortMethod,PageCount,Page,DIV


cdL= request("cdL")
cdM= request("cdM")
cdS= request("cdS")
SortMethod = request("SortMethod")
PageCount= request("PageCount")
Page= request("pg")
DIV = request("DIV")
if PageCount="" then PageCount = 30

if Page="" then Page= 1


dim objT,objView,i

set objView = new giftManagerView
objView.getMenuView cdL,cdM,cdS


if SortMethod="" then SortMethod = objView.SortMethod
IF Div="" then
	IF objView.ListType="wish" then
		Div="wish"
	else 
		Div="nor"
	end if
	
End if

set objT = new giftManagerCls
objT.FRectCDL = cdL
objT.FRectCDM = cdM
objT.FRectCDS = cdS
objT.FPageSize = PageCount
objT.FCurrPage = Page
objT.FRectSort = SortMethod
objT.FRectDiv = Div


if Div ="wish" then
objT.getBestItemList
else 
objT.getGiftItemList
end if


%>
<script language='javascript'>

function popItemWindow(iid,v){
	
	var frm = document.getElementById(v);
	
	//alert(frm.cdL.value);
	
	if (frm.cdL.value!=''&&frm.cdM.value!=''&&frm.cdS.value!=''){
		window.open("pop_itemAddInfo.asp?cd1=&target=" + frm, "addpop", "width=800,height=500,scrollbars=yes,status=no,resizable=yes");
	} else {
		alert('카테고리를 선택해 주세요');
	}
}

function checkedValue(){
	var tgvalue="";
	var chkbx = document.getElementsByName('cksel');
	
	
	//var rdodiv = document.ListFrm.div;
	//for(i=0;i<rdodiv.length;i++){
	//	if(rdodiv[i].checked&&rdodiv[i].value!=''){
	//		alert('베스트상품은 수정또는 삭제가 불가능 합니다')
	//		return '';
	//	}
	//}
	
	for (var i=0;i<chkbx.length;i++) {
		if (chkbx[i].checked){
			tgvalue=tgvalue  + chkbx[i].value + ",";
		}
	}
	
	if (tgvalue.length < 1){
		alert('하나 이상 상품을 선택해 주세요');
		return '';
	}else{
		return tgvalue;
	}
}

// 상품 추가
function AddItems(){
	var frm = document.addfrm;
	window.open("", "popup_item", "width=100,height=100,scrollbars=no,status=no,resizable=no");
	frm.target="popup_item";
	frm.submit();
}
// 상품 삭제
function DelItems(){

	var arritems = checkedValue();
	
	if (arritems.length < 1){
		return;
	} else {
		
		addfrm.arrItemID.value = arritems;
		addfrm.mode.value="del";
		window.open("", "popup_item", "width=100,height=100,scrollbars=no,status=no,resizable=no");
		addfrm.target="popup_item";
		addfrm.submit();
	}
}

// 상품 이동
function MoveItems(){
	var L,M,S
	L = '<%= cdL %>';
	M = '<%= cdM %>';
	S = '<%= cdS %>';
	
	if (L!=''&&M!=''&&S!=''){
		var arritems = checkedValue();

		if (arritems.length < 1){
			return;
		} else {
			window.open("Pop_item_Move.asp?cdL="  + L + "&cdM=" + M + "&cdS=" + S + "&arrItemID=" + arritems , "popup_item", "width=800,height=500,scrollbars=yes,status=yes,resizable=yes");
		}
	} else {
		alert('카테고리를 선택해 주세요');
		return;
	}		
	
	
}
// 순서 수정
function UpdateRank(){
	
	AnSelectAllChk(true);
	
	var arritems = checkedValue();
	
	var tgvalue="";
	
	if (arritems.length < 1){
		return;
	} else {
		var chkbx = document.getElementsByName('OrderNo');
		
		for (var i=0;i<chkbx.length;i++) {
			tgvalue=tgvalue  + chkbx[i].value + ",";
		}
		
		addfrm.arrOrderNo.value = tgvalue;
		addfrm.arrItemID.value = arritems;
		//addfrm.div.value= document.ListFrm.div.value;
		addfrm.mode.value="update";
		window.open("", "popup_item", "width=100,height=100,scrollbars=no,status=no,resizable=no");
		addfrm.target="popup_item";
		addfrm.submit();
	}

}
function AnSelectAllChk(bool){
	var frm = document.getElementsByName('cksel');
	for (var i=0;i<frm.length;i++){
		if (frm[i].disabled!=true){
			frm[i].checked = bool;
		AnCheckClick(frm[i]);
		}
	}
}


function SetDivScroll() {
	var positionTop = parseInt(document.body.scrollTop, 10);
	var objTable    = document.getElementById('header');
    
	if (objTable != null){
		objTable.style.top = positionTop;
	
	}
}
function frmsub(){
	document.ListFrm.pg.value='1';
	document.ListFrm.submit();
}
</script>

<form name="addfrm" method="get" action="Item_Process.asp">
<input type="hidden" name="mode" value="" >
<input type="hidden" name="cdL" value="<%= cdL %>">
<input type="hidden" name="cdM" value="<%= cdM %>">
<input type="hidden" name="cdS" value="<%= cdS %>">
<input type="hidden" name="div" value="<%= div %>">
<input type="hidden" name="listType" value="<%= objView.ListType %>">
<input type="hidden" name="arrItemID" value="">
<input type="hidden" name="arrOrderNo" value="">
</form>


<div id="header" style="position:absolute;top:0px">
<table width="780" border="0" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="ListFrm" method="get" action="">
	<input type="hidden" name="cdL" value="<%= cdL %>">
	<input type="hidden" name="cdM" value="<%= cdM %>">
	<input type="hidden" name="cdS" value="<%= cdS %>">
	<tr>
		<td colspan="11" bgcolor="#FFFFFF">
			<input type="button" class="button" value="상품추가" onclick="popItemWindow('','addfrm');">
			<input type="button" class="button" value="상품삭제" onclick="DelItems();"> 
			<input type="button" class="button" value="상품수정" onclick="UpdateRank();"> 
			<input type="button" class="button" value="상품이동" onclick="MoveItems();">
			<select name="SortMethod" onchange="this.form.submit();">
				<option value="cashHigh" <% if SortMethod="cashHigh" then response.write "selected" %>>가격(높은순)</option>
				<option value="cashLow" <% if SortMethod="cashLow" then response.write "selected" %>>가격(낮은순)</option>
				<option value="itemidHigh" <% if SortMethod="itemidHigh" then response.write "selected" %>>상품번호(높은순)</option>
				<option value="itemidLow" <% if SortMethod="itemidLow" then response.write "selected" %>>상품번호(낮은순)</option>
				<option value="OrderNo" <% if SortMethod="OrderNo" then response.write "selected" %>>지정번호</option>
				<option value="ItemScore" <% if SortMethod="ItemScore" then response.write "selected" %>>인기상품순</option>
				
				
			</select>
			
			<select name="PageCount"  onchange="this.form.submit();">
				<option value="10" <% if PageCount="10" then response.write "selected" %>>10</option>
				<option value="30" <% if PageCount="30" then response.write "selected" %>>30</option>
				<option value="50" <% if PageCount="50" then response.write "selected" %>>50</option>
				<option value="100" <% if PageCount="100" then response.write "selected" %>>100</option>
			</select>
			정렬
			
			총 <b><%= objT.FTotalCount %></b> 개
			<select name="pg" style="width:40;" onchange="this.form.submit();">
			<% for i=1 to objT.FTotalPage %>
				<option value="<%= i %>" <% if Cint(objT.FCurrPage) = i then response.write "selected" %>><%= i %></option>
			<% next %>
			</select>
			/<b><%= objT.FTotalPage %></b> <b>Pages</b><%=objT.FResultCount %>
			
		</td>

	</tr>
	<tr>
		<td colspan="11" bgcolor="#FFFFFF">
			<input type="radio" name="div" value="nor" <% if DIV="nor" then response.write "checked" %> onclick="frmsub();">일반상품
			<input type="radio" name="div" value="sell" <% if DIV="sell" then response.write "checked" %> onclick="frmsub();">베스트 - SELL
			<input type="radio" name="div" value="wish" <% if DIV="wish" then response.write "checked" %> onclick="frmsub();">베스트 - WISH
			<input type="radio" name="div" value="review" <% if DIV="review" then response.write "checked" %> onclick="frmsub();">베스트 - REVIEW
				
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"><input type="checkbox" name="ckselm" onClick="AnSelectAllChk(this.checked);"></td>
		<td width="25" align="center">No.</td>
		<td width="50" align="center">이미지</td>
		<td width="55" align="center">상품코드 </td>
		<td width="230" align="center">상품명</td>
		<td width="60" align="center">가격</td>
		<td width="100" align="center">브랜드</td>
		<td width="60" align="center">매입구분</td>
		<td width="30" align="center">전시</td>
		<td width="30" align="center">판매</td>
		<td width="50" align="center">순서</td>
	</tr>
	</form>
	
</table>
</div>


<table width="780"  border="0" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td colspan="11"  bgcolor="#FFFFFF">
			<input type="button" class="button" value="상품추가" onclick="">
			<input type="button" class="button" value="상품삭제" onclick=""> 
			<input type="button" class="button" value="상품수정" onclick=""> 
		</td>
	</tr>
	<tr>
		<td colspan="11" bgcolor="#FFFFFF">
			<input type="radio" name="div" value="Sell">베스트 - SELL
			<input type="radio" name="div" value="Wish">베스트 - WISH
			<input type="radio" name="div" value="Review">베스트 - REVIEW
			</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="11"><input type="checkbox" name="ckselm" onClick="AnSelectAllChk(this.checked);"></td>
		
	</tr>
	<% if objT.FResultCount >0 then %>
	<% for i=0 to objT.FResultCount-1 %>
	<tr bgcolor="#FFFFFF">
		<td width="20" align="center"><input type="checkbox" name="cksel" value="<%= objT.FItemList(i).FItemid %>" onClick="AnCheckClick(this);"></td>
		<td width="25" align="center"><%= objT.FTotalCount-objT.FPageSize*(objT.FCurrPage-1)-i %></td>
		<td width="50" align="center"><img src="<%= objT.FItemList(i).FImageSmall %>" width="50" height="50" border="0"></td>
		<td width="55" align="center"><%= objT.FItemList(i).FItemid %></td>
		<td width="230" align="center" style="word-wrap:break-word"><%= objT.FItemList(i).FItemName %></td>
		<td width="60" align="center"><%= FormatNumber(objT.FItemList(i).getRealPrice,0) %></td>
		<td width="100" align="center" style="word-wrap:break-word"><%= objT.FItemList(i).FMakerID %></td>
		<td width="60" align="center">
			<font color="<%= MwDivcolor(objT.FItemList(i).FMwdiv) %>"><%= MwDivName(objT.FItemList(i).FMwdiv) %>&nbsp;
				<% if objT.FItemList(i).FSellcash<>0 then %>
					<%= 100-CLng(objT.FItemList(i).FBuycash/objT.FItemList(i).FSellcash*100*100)/100 %>%
				<% end if %>
			</font>    
		</td>
		<td width="30" align="center"><%= getYNColor(objT.FItemList(i).FDispYn) %></td>
		<td width="30" align="center"><%= getYNColor(objT.FItemList(i).FSellYn) %></td>
		<td width="50" align="center"><input type="text" name="OrderNo" value="<%= objT.FItemList(i).FOrderNo %>" size="4"></td>
	</tr>
	<% next %>
	<% end if %>
</table>
<% set objT = nothing %>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->