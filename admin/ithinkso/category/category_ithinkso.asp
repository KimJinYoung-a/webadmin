<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'###########################################################
' Description : 아이띵소 카테고리 관리
' Hieditor : 2013.05.09 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>

<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language='javascript'>

function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<% if session("sslgnMethod")<>"S" then %>
	<!-- USB키 처리 시작 (2008.06.23;허진원) -->
	<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
	<script language="javascript" src="/js/check_USBToken.js"></script>
	<!-- USB키 처리 끝 -->
<% end if %>
</head>
<body bgcolor="#F4F4F4" onload="checkUSBKey()">

<!-- #include virtual="/lib/classes/ithinkso/category/category_cls_ithinkso.asp"-->

<%
dim CateSeq0, CateSeq1, CateSeq2, CateSeq3, i, cCate, Depth, cCategory, menupos
	CateSeq0	=	requestCheckVar(Request("iCateSeq0"),10)
	CateSeq1	=	requestCheckVar(Request("iCateSeq1"),10)
	CateSeq2	=	requestCheckVar(Request("iCateSeq2"),10)
	CateSeq3	=	requestCheckVar(Request("iCateSeq3"),10)
	Depth 		= 	requestCheckVar(Request("Depth"),10)
	menupos = request("menupos")
	
IF Depth = "" THEN Depth = 0
'if CateSeq0 = "" then CateSeq0 = 1

%>	
		
<script language="javascript">

function frmsubmit(){
	frmSearch.submit();
}

//상위 카테고리 변경에 따른 하위 카테고리 데이터 변경 처리
function jsChCategory(intD){				
	var intT = 0;	
	eval("document.frmSearch.iCateSeq"+intD).value =  eval("document.frmSearch.CateSeq"+intD).options[eval("document.frmSearch.CateSeq"+intD).selectedIndex].value;					
	if(eval("document.frmSearch.iCateSeq"+intD).value ==""){
	  if (intD == 0) {
	    document.frmSearch.Depth.value="";
	    frmsubmit();	
	  }else{
		jsChCategory(intD-1);
	  }
	}else{
		intT= eval("document.frmSearch.CateSeq"+intD).options[eval("document.frmSearch.CateSeq"+intD).selectedIndex].thread;		
								
		document.frmSearch.Depth.value = intD;
		
		frmsubmit();		
	}	
}
		   
//신규등록
function categoryreg(intD){
	if (frmSearch.newCateName.value==''){
		alert('카테고리명을 입력하세요.');
		frmSearch.newCateName.focus();
		return;
	}
	if (frmSearch.newCateOrder.value==''){
		alert('정렬순서를 입력하세요.');
		frmSearch.newCateOrder.focus();
		return;
	}
	if (frmSearch.newisusing.value==''){
		alert('사용유무를 입력하세요.');
		frmSearch.newisusing.focus();
		return;
	}
	
	frmSearch.regDepth.value = intD;
	frmSearch.mode.value="categoryreg";	
	frmSearch.target = "hidCategory";
	frmSearch.action = "/admin/ithinkso/category/category_process_ithinkso.asp";	
	frmSearch.submit();
	frmSearch.regDepth.value = "";
	frmSearch.mode.value = "";	
	frmSearch.target = "";
	frmSearch.action = "";
}

//수정
function categoryedit(intD){
	for (var i = 0; i < frmSearch.CateOrder.length; i++){
		if (!IsDouble(frmSearch.CateOrder[i].value)){
			alert('정렬순서를 입력하세요.');
			frmSearch.CateOrder[i].focus();
			return;
		}
	}

	frmSearch.regDepth.value = intD;
	frmSearch.mode.value="categoryedit";	
	frmSearch.target = "hidCategory";
	frmSearch.action = "/admin/ithinkso/category/category_process_ithinkso.asp";	
	frmSearch.submit();
	frmSearch.regDepth.value = "";
	frmSearch.mode.value = "";	
	frmSearch.target = "";
	frmSearch.action = "";
}

//카테고리매뉴리얼적용
function categoryReal(){
	var categoryReal = window.open('<%= wwwithinksoweb %>/chtml/make_category.asp','categoryReal','width=400,height=300,scrollbars=yes,resizable=yes');
	categoryReal.focus();
}

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="#999999">
			<tr>
				<td width="400" style="padding:5; border-top:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999"  background="/images/menubar_1px.gif">
					<font color="#333333"><b>[ON]해외상품관리&gt;&gt;아이띵소카테고리관리</b></font>
				</td>
				
				<td align="right" style="border-bottom:1px solid #999999" bgcolor="#F4F4F4">
					<!-- 마스터이상 메뉴권한 설정 -->
					
					<a href="Javascript:PopMenuEdit('1491');"><img src="/images/icon_chgauth.gif" border="0" valign="bottom"></a>
					
					<!-- Help 설정 -->
					
					<a href="Javascript:PopMenuHelp('1491');"><img src="/images/icon_help.gif" border="0" valign="bottom"></a>
					
				</td>
				
			</tr>
		</table>
	</td>
</tr>
</table>

<form name="frmSearch" method="post" style="margin:0px;">
<input type="hidden" name="iCateSeq0" value="<%=CateSeq0%>">
<input type="hidden" name="iCateSeq1" value="<%=CateSeq1%>">
<input type="hidden" name="iCateSeq2" value="<%=CateSeq2%>">
<input type="hidden" name="iCateSeq3" value="<%=CateSeq3%>">
<input type="hidden" name="Depth" value="<%=Depth%>">
<input type="hidden" name="regDepth" value="">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<%
		set cCate = new ccategory_ithinkso
			cCate.frectisusing = "Y"
			cCate.getCategoryType_notpaging
		%>
		* 타입 :		
		<select name="CateSeq0" onchange="jsChCategory(0);">
			<option value="">--선택--</option>
			<%
			IF cCate.fresultcount > 0 THEN
				
			For i = 0 To cCate.fresultcount - 1
			%>	
			<option value="<%= cCate.FItemList(i).fCateTypeSeq %>" <% if cstr(CateSeq0) = cstr(cCate.FItemList(i).fCateTypeSeq) then response.write " selected" %>>
				<%= cCate.FItemList(i).fCateTypeName %>
			</option>
			<%
			NEXT
			
			END IF	
			%>
		</select>
		<% set cCate = nothing %>
	 	&nbsp;>&nbsp;
		<% 
		set cCate = new ccategory_ithinkso
			cCate.frectCateTypeSeq = CateSeq0
			cCate.frectisusing = "Y"
			
			if CateSeq0 <> "" then
		 		cCate.getCategory_notpaging
		 	end if
		%>
		대카테 : 
		<select name="CateSeq1" onChange="jsChCategory(1);">	
			<option value="">--전체--</option>
			<%
			IF cCate.fresultcount > 0 THEN
				
			For i = 0 To cCate.fresultcount - 1
			%>				
			<option value="<%= cCate.FItemList(i).fCateSeq %>" <% if cstr(CateSeq1) = cstr(cCate.FItemList(i).fCateSeq) then response.write " selected" %>>
				<%= cCate.FItemList(i).fCateName %>
			</option>
			<%
			NEXT
			
			END IF	
			%>				
		</select>
		<% set cCate = nothing %>
		&nbsp;>&nbsp;		
		중카테 :
		<% 
		set cCate = new ccategory_ithinkso
			cCate.frectCateTypeSeq = CateSeq0
			cCate.frectsubCateSeq1 = CateSeq1
			cCate.frectisusing = "Y"
			
			if CateSeq0 <> "" and Depth > 0 then
		 		cCate.getCategory_notpaging
		 	end if
		%>		
		<select name="CateSeq2" onChange="jsChCategory(2);">
			<option value="">--전체--</option>
			<%
			IF cCate.fresultcount > 0 THEN
				
			For i = 0 To cCate.fresultcount - 1
			%>				
			<option value="<%= cCate.FItemList(i).fCateSeq %>" <% if cstr(CateSeq2) = cstr(cCate.FItemList(i).fCateSeq) then response.write " selected" %>>
				<%= cCate.FItemList(i).fCateName %>
			</option>
			<%
			NEXT
			
			END IF	
			%>
		</select>
		<% set cCate = nothing %>		
	</td>
</tr>
</table>
<!-- 검색 끝 -->
<Br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
    </td>
    <td align="right">
    	<a href="javascript:categoryReal();"><img src="/images/refreshcpage.gif" border="0">카테고리매뉴Real생성</a>
    </td>
</tr>	
</table>
<!-- 표 중간바 끝-->

<iframe id="hidCategory" name="hidCategory" src="about:blank" frameborder="0" width=0 height=0></iframe>

<% if CateSeq0 = "" and Depth = 0 then %>
	<div id='tCT' style='padding:5 5 5 5;'><font color="red">- 카테고리 타입</font></div>
	<%
	set cCategory = new ccategory_ithinkso
		'cCategory.frectisusing = "Y"
		cCategory.getCategoryType_notpaging
	%>
	<div id='dCT' style='display:;'>
	<table width='100%' align='center' cellpadding='3' cellspacing='1' class='a' bgcolor='<%= adminColor("tablebg") %>'>
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>
		<td colspan=10 height=20><font color='blue' size=3><b>* 등록리스트 *</b></font></td>
	</tr>
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>					
		<td>카테고리타입번호</td>
		<td>카테고리타입명</td>
		<td>정렬순서</td>
		<td>사용유무</td>
	</tr>	
	<%
	IF cCategory.fresultcount > 0 THEN
		
	For i = 0 To cCategory.fresultcount - 1
	%>	
	<input type="hidden" name="tCateSeq" value="<%= cCategory.FItemList(i).fCateTypeSeq %>">
	<tr align="center" bgcolor="#FFFFFF" onmouseover="this.style.background='#f1f1f1';" onmouseout="this.style.background='#FFFFFF';">
		<td><%= cCategory.FItemList(i).fCateTypeSeq %></td>
		<td><input type="text" name="CateName" value="<%= cCategory.FItemList(i).fCateTypeName %>" size=20></td>
		<td><input type="text" name="CateOrder" value="<%= cCategory.FItemList(i).fCateTypeOrder %>" size=4></td>
		<td>
		   <select name="isusing">
			   <option value="Y" <% if cCategory.FItemList(i).fIsUsing="Y" then response.write "selected" %>>Y</option>
			   <option value="N" <% if cCategory.FItemList(i).fIsUsing="N" then response.write "selected" %>>N</option>
		   </select>
		</td>
	</tr>
			
	<%	Next %>
	<tr align='center' bgcolor='FFFFFF'>
		<td colspan=10 align='right'><input type='button' class='button' value='수정' onClick='categoryedit(0);'></td>
	</tr>
	<%ELSE%>	
	<tr bgcolor='FFFFFF' align='center'>
		<td colspan=10>등록된 내역이 없습니다</td>
	</tr>				
	<%END IF%>

	<tr align='center' bgcolor='ffffff'>
		<td colspan=10 height=30></td>
	</tr>
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>
		<td colspan=10 height=20><font color='blue' size=3><b>* 신규등록 *</b></font></td>
	</tr>
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>
		<td>카테고리타입번호</td>
		<td>카테고리타입명</td>										
		<td>정렬순서</td>
		<td>사용유무</td>
	</tr>
	<tr align='center' bgcolor='FFFFFF'>
		<td></td>
		<td><input type='text' name='newCateName' size=20></td>
		<td><input type='text' name='newCateOrder' size=4></td>
		<td>
		   <select name="newisusing">
			   <option value="Y">Y</option>
			   <option value="N">N</option>
		   </select>
		</td>
	</tr>						
	<tr align='center' bgcolor='FFFFFF'>
		<td colspan=10 align='right'><input type='button' class='button' value='신규저장' onClick='categoryreg(0);'></td>
	</tr>
	</table>
	</div>
	<% set cCategory = nothing %>
<% else %>
	<div id='tCT' style='padding:5 5 5 5;'>+ 카테고리 타입</div>
<% end if %>

<% if CateSeq0 <> "" and Depth = 0 then %>
	<div id='tC1' style='padding:5 5 5 5;'><font color="red">- 대카테고리</font></div>
	<%
	set cCategory = new ccategory_ithinkso
		cCategory.frectCateTypeSeq = CateSeq0
		'cCategory.frectisusing = "Y"
	 	cCategory.getCategory_notpaging
	%>
	<div id='dC1' style='display:;'>								
	<table width='100%' align='center' cellpadding='3' cellspacing='1' class='a' bgcolor='<%= adminColor("tablebg") %>'>
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>
		<td colspan=10 height=20><font color='blue' size=3><b>* 등록리스트 *</b></font></td>
	</tr>																																																
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>
		<td>카테고리번호</td>
		<td>카테고리타입명</td>										
		<td>정렬순서</td>
		<td>사용유무</td>
	</tr>	
	<%
	IF cCategory.fresultcount > 0 THEN
		
	For i = 0 To cCategory.fresultcount - 1
	%>	
	<input type="hidden" name="tCateSeq" value="<%= cCategory.FItemList(i).fCateSeq %>">
	<tr align="center" bgcolor="#FFFFFF" onmouseover="this.style.background='#f1f1f1';" onmouseout="this.style.background='#FFFFFF';">
		<td><%= cCategory.FItemList(i).fCateSeq %></td>
		<td><input type="text" name="CateName" value="<%= cCategory.FItemList(i).fCateName %>" size="20"></td>
		<td><input type="text" name="CateOrder" value="<%= cCategory.FItemList(i).fCateOrder %>" size="4"></td>
		<td>
		   <select name="isusing">
			   <option value="Y" <% if cCategory.FItemList(i).fIsUsing="Y" then response.write "selected" %>>Y</option>
			   <option value="N" <% if cCategory.FItemList(i).fIsUsing="N" then response.write "selected" %>>N</option>
		   </select>
		</td>
	</tr>						
			
	<%	Next %>
	<tr align='center' bgcolor='FFFFFF'>
		<td colspan=10 align='right'><input type='button' class='button' value='수정' onClick='categoryedit(1);'></td>
	</tr>
	<%ELSE%>	
	<tr bgcolor='FFFFFF' align='center'>
		<td colspan=10>등록된 내역이 없습니다</td>
	</tr>				
	<%END IF%>

	<tr align='center' bgcolor='ffffff'>
		<td colspan=10 height=30></td>
	</tr>
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>
		<td colspan=10 height=20><font color='blue' size=3><b>* 신규등록 *</b></font></td>
	</tr>
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>
		<td>카테고리번호</td>
		<td>카테고리타입명</td>										
		<td>정렬순서</td>
		<td>사용유무</td>
	</tr>
	<tr align='center' bgcolor='FFFFFF'>
		<td></td>
		<td><input type='text' name='newCateName' size=20></td>
		<td><input type='text' name='newCateOrder' size=4></td>
		<td>
		   <select name="newisusing">
			   <option value="Y">Y</option>
			   <option value="N">N</option>
		   </select>
		</td>
	</tr>						
	<tr align='center' bgcolor='FFFFFF'>
		<td colspan=10 align='right'><input type='button' class='button' value='신규저장' onClick='categoryreg(1);'></td>
	</tr>
	</table>					
	</div>
	<% set cCategory = nothing %>
<% else %>
	<div id='tC1' style='padding:5 5 5 5;'>+ 대카테고리</div>
<% end if %>

<% if CateSeq0 <> "" and Depth = 1 then %>
	<div id='tC2' style='padding:5 5 5 5;'><font color="red">- 중카테고리</font></div>
	<%
	set cCategory = new ccategory_ithinkso
		cCategory.frectCateTypeSeq = CateSeq0
		cCategory.frectsubCateSeq1 = CateSeq1
		'cCategory.frectisusing = "Y"
	 	cCategory.getCategory_notpaging
	%>
	<div id='dC2' style='display:;'>								
	<table width='100%' align='center' cellpadding='3' cellspacing='1' class='a' bgcolor='<%= adminColor("tablebg") %>'>
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>
		<td colspan=10 height=20><font color='blue' size=3><b>* 등록리스트 *</b></font></td>
	</tr>	
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>
		<td>카테고리번호</td>
		<td>카테고리타입명</td>										
		<td>정렬순서</td>
		<td>사용유무</td>
	</tr>	
	<%
	IF cCategory.fresultcount > 0 THEN
		
	For i = 0 To cCategory.fresultcount - 1
	%>	
	<input type="hidden" name="tCateSeq" value="<%= cCategory.FItemList(i).fCateSeq %>">
	<tr align="center" bgcolor="#FFFFFF" onmouseover="this.style.background='#f1f1f1';" onmouseout="this.style.background='#FFFFFF';">
		<td><%= cCategory.FItemList(i).fCateSeq %></td>
		<td><input type="text" name="CateName" value="<%= cCategory.FItemList(i).fCateName %>" size="20"></td>
		<td><input type="text" name="CateOrder" value="<%= cCategory.FItemList(i).fCateOrder %>" size="4"></td>
		<td>
		   <select name="isusing">
			   <option value="Y" <% if cCategory.FItemList(i).fIsUsing="Y" then response.write "selected" %>>Y</option>
			   <option value="N" <% if cCategory.FItemList(i).fIsUsing="N" then response.write "selected" %>>N</option>
		   </select>
		</td>
	</tr>						
			
	<%	Next %>
	<tr align='center' bgcolor='FFFFFF'>
		<td colspan=10 align='right'><input type='button' class='button' value='수정' onClick='categoryedit(2);'></td>
	</tr>
	<%ELSE%>	
	<tr bgcolor='FFFFFF' align='center'>
		<td colspan=10>등록된 내역이 없습니다</td>
	</tr>				
	<%END IF%>

	<tr align='center' bgcolor='ffffff'>
		<td colspan=10 height=30></td>
	</tr>
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>
		<td colspan=10 height=20><font color='blue' size=3><b>* 신규등록 *</b></font></td>
	</tr>
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>
		<td>카테고리번호</td>
		<td>카테고리타입명</td>										
		<td>정렬순서</td>
		<td>사용유무</td>
	</tr>
	<tr align='center' bgcolor='FFFFFF'>
		<td></td>
		<td><input type='text' name='newCateName' size=20></td>
		<td><input type='text' name='newCateOrder' size=4></td>
		<td>
		   <select name="newisusing">
			   <option value="Y">Y</option>
			   <option value="N">N</option>
		   </select>
		</td>
	</tr>						
	<tr align='center' bgcolor='FFFFFF'>
		<td colspan=10 align='right'><input type='button' class='button' value='신규저장' onClick='categoryreg(2);'></td>
	</tr>
	</table>
	</div>
	<% set cCategory = nothing %>
<% else %>
	<div id='tC2' style='padding:5 5 5 5;'>+ 중카테고리</div>
<% end if %>

<% if CateSeq0 <> "" and Depth = 2 then %>
	<div id='tC2' style='padding:5 5 5 5;'><font color="red">- 소카테고리</font></div>
	<%
	set cCategory = new ccategory_ithinkso
		cCategory.frectCateTypeSeq = CateSeq0
		cCategory.frectsubCateSeq1 = CateSeq1
		cCategory.frectsubCateSeq2 = CateSeq2	
		'cCategory.frectisusing = "Y"
	 	cCategory.getCategory_notpaging
	%>
	<div id='dC3' style='display:;'>								
	<table width='100%' align='center' cellpadding='3' cellspacing='1' class='a' bgcolor='<%= adminColor("tablebg") %>'>
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>
		<td colspan=10 height=20><font color='blue' size=3><b>* 등록리스트 *</b></font></td>
	</tr>	
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>
		<td>카테고리번호</td>
		<td>카테고리타입명</td>										
		<td>정렬순서</td>
		<td>사용유무</td>
	</tr>	
	<%
	IF cCategory.fresultcount > 0 THEN
		
	For i = 0 To cCategory.fresultcount - 1
	%>	
	<input type="hidden" name="tCateSeq" value="<%= cCategory.FItemList(i).fCateSeq %>">
	<tr align="center" bgcolor="#FFFFFF" onmouseover="this.style.background='#f1f1f1';" onmouseout="this.style.background='#FFFFFF';">
		<td><%= cCategory.FItemList(i).fCateSeq %></td>
		<td><input type="text" name="CateName" value="<%= cCategory.FItemList(i).fCateName %>" size="20"></td>
		<td><input type="text" name="CateOrder" value="<%= cCategory.FItemList(i).fCateOrder %>" size="4"></td>
		<td>
		   <select name="isusing">
			   <option value="Y" <% if cCategory.FItemList(i).fIsUsing="Y" then response.write "selected" %>>Y</option>
			   <option value="N" <% if cCategory.FItemList(i).fIsUsing="N" then response.write "selected" %>>N</option>
		   </select>
		</td>
	</tr>						
			
	<%	Next %>
	<tr align='center' bgcolor='FFFFFF'>
		<td colspan=10 align='right'><input type='button' class='button' value='수정' onClick='categoryedit(3);'></td>
	</tr>
	<%ELSE%>	
	<tr bgcolor='FFFFFF' align='center'>
		<td colspan=10>등록된 내역이 없습니다</td>
	</tr>				
	<%END IF%>

	<tr align='center' bgcolor='ffffff'>
		<td colspan=10 height=30></td>
	</tr>
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>
		<td colspan=10 height=20><font color='blue' size=3><b>* 신규등록 *</b></font></td>
	</tr>
	<tr align='center' bgcolor='<%= adminColor("tabletop") %>'>
		<td>카테고리번호</td>
		<td>카테고리타입명</td>										
		<td>정렬순서</td>
		<td>사용유무</td>
	</tr>
	<tr align='center' bgcolor='FFFFFF'>
		<td></td>
		<td><input type='text' name='newCateName' size=20></td>
		<td><input type='text' name='newCateOrder' size=4></td>
		<td>
		   <select name="newisusing">
			   <option value="Y">Y</option>
			   <option value="N">N</option>
		   </select>
		</td>
	</tr>						
	<tr align='center' bgcolor='FFFFFF'>
		<td colspan=10 align='right'><input type='button' class='button' value='신규저장' onClick='categoryreg(3);'></td>
	</tr>
	</table>					
	</div>
	<% set cCategory = nothing %>
<% else %>
	<div id='tC3' style='padding:5 5 5 5;'>+ 소카테고리</div>
<% end if %>

</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->