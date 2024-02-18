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
' Hieditor : 2011.04.05 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->

<%
dim ocate ,i ,menupos ,ocateone , cd1,cd2,cd3 ,mode , itemadd
dim catetype,catecode,catename,isusing,orderno,lastadminid
	menupos = request("menupos")	
	catetype = request("catetype")
	cd1 = request("cd1")
	cd2 = request("cd2")
	cd3 = request("cd3")
	mode = request("mode")
	
	if catetype = "" then catetype = "CD1"
		
	'//카테고리 수정
	set ocateone = new cstylepickMenu
	ocateone.frectcd1 = cd1
	ocateone.frectcd2 = cd2
	ocateone.frectcd3 = cd3
	
	if cd1 <> "" then
		ocateone.getstylepick_cate_cd1_one()
	
		if ocateone.ftotalcount > 0 then						
			catecode = ocateone.foneitem.fcd1
			catename = ocateone.foneitem.fcatename
			isusing = ocateone.foneitem.fisusing
			orderno = ocateone.foneitem.forderno		
			lastadminid = ocateone.foneitem.flastadminid			
		end if
		
	elseif cd2 <> "" then
		ocateone.getstylepick_cate_cd2_one()
		
		if ocateone.ftotalcount > 0 then						
			catecode = ocateone.foneitem.fcd2
			catename = ocateone.foneitem.fcatename
			isusing = ocateone.foneitem.fisusing
			orderno = ocateone.foneitem.forderno		
			lastadminid = ocateone.foneitem.flastadminid			
		end if
		
	elseif cd3 <> "" then
		ocateone.getstylepick_cate_cd3_one()
		
		if ocateone.ftotalcount > 0 then						
			catecode = ocateone.foneitem.fcd3
			catename = ocateone.foneitem.fcatename
			isusing = ocateone.foneitem.fisusing
			orderno = ocateone.foneitem.forderno		
			lastadminid = ocateone.foneitem.flastadminid			
		end if			
	end if

	'//카테고리 리스트
	set ocate = new cstylepickMenu
	
	if catetype = "CD1" then
		ocate.getstylepick_cate_cd1()
	elseif catetype = "CD2" then
		ocate.getstylepick_cate_cd2()
	elseif catetype = "CD3" then
		ocate.getstylepick_cate_cd3()		
	end if
	
	if orderno = "" then orderno = "1"
	if isusing = "" then isusing = "Y"
	if mode = "" then mode = "itemadd"	
%>

<script language='javascript'>

function Savepick(mode){
    
    if (frmedit.catecode.value == ''){
        alert('카테고리 코드를 입력 하세요.');
        frmedit.catecode.focus();
        return;
    }    

    if (frmedit.catename.value == ''){
        alert('카테고리 명을 입력 하세요.');
        frmedit.catename.focus();
        return;
    }

    if (frmedit.orderno.value == ''){
        alert('정렬순서를 입력 하세요.');
        frmedit.orderno.focus();
        return;
    }
    
    if (frmedit.isusing.value.length<1){
        alert('사용여부를 선택하세요.');
        frmedit.isusing.focus();
        return;
    }
            
    if (confirm('저장 하시겠습니까?')){
        frmedit.mode.value=mode;
        frmedit.submit();
    }    
}

function frmsubmit(){
	frm.submit();
}

function Createcategory(catetype){
	//alert("프론트가 오픈전입니다");
	//return;

	var Createcategory = window.open('<%= wwwUrl %>/chtml/make_stylepick_category_menu.asp?catetype='+catetype,'Createcategory','width=800,height=768,scrollbars=yes,resizable=yes');
	Createcategory.focus();
}

function midcate_manage(code)
{
	var midcategory = window.open('/admin/stylepick/stylepick_midcate.asp?code='+code,'midcategory','width=500,height=500,scrollbars=yes,resizable=yes');
	midcategory.focus();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		카테고리타입 : <% Drawcatetype "catetype",catetype," onchange='frmsubmit();'" %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmedit" method="post" action="/admin/stylepick/stylepick_category_process.asp" >
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="catecodeorg" value="<%= catecode %>">
<input type="hidden" name="isusingorg" value="<%= isusing %>">
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">카테고리타입</td>
    <td>
		<%=GETcatetype(catetype) %><input type="hidden" name="catetype" value="<%= catetype %>">
    </td>
</tr>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">카테고리코드</td>
    <td>
        <input type="text" name="catecode" value="<%= catecode %>" maxlength="3" size="3"> ※숫자3자리
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">카테고리명</td>
    <td>
        <input type="text" name="catename" value="<%= catename %>" maxlength="32" size="32">        
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">정렬순서</td>
    <td>
        <input type="text" name="orderno" value="<%= orderno %>" maxlength="2" size="2"> ex) 1 ~ 99 숫자가 낮을수록 우선순위
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">사용여부</td>
    <td>
		<select name="isusing">
			<option value="" <% if isusing="" then response.write " selected" %>>선택</option>
			<option value="Y" <% if isusing="Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing="N" then response.write " selected" %>>N</option>
		</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2">
    	<% if mode = "itemadd" then %>
    		<input type="button" value="신규저장" onClick="Savepick('itemadd');" class="button">
    	<% else %>
    		<input type="button" value="수정" onClick="Savepick('itemedit');" class="button">
    	<% end if %>
    </td>
</tr>
</form>
</table>

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<% if catetype="CD1" then %>
			<input type="button" value="스타일 카테고리 매뉴 실서버 적용" class="button" onclick="Createcategory('<%= catetype %>');">
		<% end if %>
	</td>
	<td align="right">
		<input type="button" onclick="location.href='?menupos=<%=menupos%>&catetype=<%=catetype%>&itemadd=itemadd'" value="신규등록" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= ocate.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">	
	<td>카테고리타입</td>
	<td>카테고리코드</td>
	<td>카테고리명</td>
	<td>정렬순서</td>
	<td>사용여부</td>
	<td>최종수정인</td>
	<td>비고</td>
</tr>
<% if ocate.FresultCount>0 then %>
<% for i=0 to ocate.FresultCount-1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">			

<% if ocate.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='FFFFaa';>
<% end if %>
	<td>
		<%= GETcatetype(catetype) %>
	</td>
	<td>
		<%= ocate.FItemList(i).fcd1 %><%= ocate.FItemList(i).fcd2 %><%= ocate.FItemList(i).fcd3 %>
	</td>
	<td>
		<%= ocate.FItemList(i).fcatename %>
	</td>
	<td>
		<%= ocate.FItemList(i).forderno %>
	</td>
	<td>
		<%= ocate.FItemList(i).fisusing %>
	</td>
		
	<td>
		<%= ocate.FItemList(i).flastadminid %>
	</td>
	<td>
		<input type="button" onclick="location.href='?catetype=<%= catetype %>&menupos=<%=menupos%>&cd1=<%=ocate.FItemList(i).fcd1%>&cd2=<%=ocate.FItemList(i).fcd2%>&cd3=<%=ocate.FItemList(i).fcd3%>&mode=itemedit'" value="수정" class="button">
		<% If catetype = "CD2" Then %>
		&nbsp;<input type="button" onclick="midcate_manage(<%=Mid(ocate.FItemList(i).fcd2,2,1)%>);" value="중분류관리" class="button">
		<% End If %>
	</td>
</tr>   
</form>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<% 
set ocate = nothing
set ocateone = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->