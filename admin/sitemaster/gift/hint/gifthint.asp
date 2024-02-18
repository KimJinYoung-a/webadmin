<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기프트
' History : 2015.01.26 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/gifthint_cls.asp"-->
<%
dim page, i, themeidx, themetype, title, executetime, isusing, orderno, regdate, lastadminid, lastupdate
dim selectisusing, selectthemeidx, selecttitle, research, isnew, selectthemetype
	themeidx = getNumeric(requestcheckvar(request("themeidx"),10))
	page = getNumeric(requestcheckvar(request("page"),10))
	selectisusing = requestcheckvar(request("selectisusing"),10)
	selectthemeidx = getNumeric(requestcheckvar(request("selectthemeidx"),10))
	selectthemetype = getNumeric(requestcheckvar(request("selectthemetype"),10))
	selecttitle = requestcheckvar(request("selecttitle"),128)
	research = requestcheckvar(request("research"),2)
	isnew = requestcheckvar(request("isnew"),1)

if page="" then page=1

dim othemeone
set othemeone = new Cgifthint
	othemeone.frectthemeidx = themeidx
	
	if themeidx<>"" then
		othemeone.getgifthint_one

        themeidx = othemeone.FOneItem.Fthemeidx
        themetype = othemeone.FOneItem.Fthemetype
        title = othemeone.FOneItem.Ftitle
        executetime = othemeone.FOneItem.Fexecutetime
        isusing = othemeone.FOneItem.Fisusing
        orderno = othemeone.FOneItem.Forderno
        regdate = othemeone.FOneItem.Fregdate
        lastadminid = othemeone.FOneItem.Flastadminid
        lastupdate = othemeone.FOneItem.Flastupdate
	end if
set othemeone = Nothing

dim otheme
set otheme = new Cgifthint
	otheme.FPageSize=20
	otheme.FCurrPage= page
	otheme.frectthemeidx = selectthemeidx
	otheme.frectisusing = selectisusing
	otheme.frecttitle = selecttitle
	otheme.frectthemetype = selectthemetype
	otheme.getgifthint_list

if orderno="" then orderno=99
if isusing="" then isusing="Y"
if executetime="" then executetime="00:00:00"
if selectisusing="" and research="" then selectisusing="Y"
%>

<script type='text/javascript'>

function Savetheme(){
    if (frmtheme.themetype.value.length<1){
        alert('테마타입을 선택 하세요.');
        frmtheme.themetype.focus();
        return;
    }
    if (frmtheme.title.value==''){
        alert('테마명을 입력 하세요.');
        frmtheme.title.focus();
        return;
    }
    if (frmtheme.isusing.value.length<1){
        alert('사용여부를 선택 하세요.');
        frmtheme.isusing.focus();
        return;
    }
    if (frmtheme.orderno.value==''){
		alert('정렬순서를 입력 하세요.');
        frmtheme.orderno.focus();
        return;
    }
    
    if (confirm('저장 하시겠습니까?')){
    	frmtheme.mode.value='regtheme';
    	frmtheme.action='/admin/sitemaster/gift/hint/gifthint_process.asp';
        frmtheme.submit();
    }
}

function chselected(themeidx){
	location.href='/admin/sitemaster/gift/hint/gifthint.asp?themeidx='+themeidx+'&menupos=<%= menupos %>';
}

function chsnewtheme(){
	location.href='/admin/sitemaster/gift/hint/gifthint.asp?isnew=Y&menupos=<%= menupos %>';
}

function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}

function jsSetItem(themeidx){
	var jsSetItem;
	jsSetItem = window.open('/admin/sitemaster/gift/hint/gifthint_item.asp?themeidx='+themeidx+'&menupos=<%= menupos %>','jsSetItem','width=1024,height=768,scrollbars=yes,resizable=yes');
	jsSetItem.focus();
}

</script>

<% 
'/신규등록 & 수정시에만 노출
if isnew="Y" or themeidx<>"" then
%>
	<form name="frmtheme" method="post" action="" style="margin:0px;" >
	<input type="hidden" name="mode">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	
	<% if themeidx<>"" then %>
		<tr bgcolor="<%= adminColor("tabletop") %>">
		    <td width="150">테마번호</td>
		    <td bgcolor="FFFFFF">
	        	<%= themeidx %>
				<input type="hidden" name="themeidx" value="<%= themeidx %>" >
		    </td>
		</tr>
	<% end if %>
	
	<tr bgcolor="<%= adminColor("tabletop") %>">
	    <td width="150">테마타입</td>
	    <td bgcolor="FFFFFF">
	    	<% drawthemetype "themetype", themetype, "" %>
	    </td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
	    <td width="150">테마명</td>
	    <td bgcolor="FFFFFF">
	        <input type="text" name="title" value="<%= ReplaceBracket(title) %>" maxlength="64" size="80">
	    </td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
	    <td width="150">적용시간</td>
	    <td bgcolor="FFFFFF">
			<input type="text" name="executetime" size=7 maxlength=8 value="<%= trim(executetime) %>" class="text">
	    </td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
	    <td width="150">사용여부</td>
	    <td bgcolor="FFFFFF">
	    	<% drawSelectBoxisusingYN "isusing", isusing, "" %>
	    </td>
	</tr>
	<input type="hidden" name="orderno" value="<%= orderno %>" maxlength="2" size="2">
	<!--<tr bgcolor="<%= adminColor("tabletop") %>">
	    <td width="150">정렬순서</td>
	    <td bgcolor="FFFFFF">
	        <input type="text" name="orderno" value="<% ' = orderno %>" maxlength="2" size="2"> 기본값 : 99 , 숫자가 높을수록 상위에 노출되며 기본값으로 두실경우 최신등록 순으로 노출 됩니다.
	    </td>
	</tr>-->
	
	<% if themeidx<>"" then %>
		<tr bgcolor="<%= adminColor("tabletop") %>">
		    <td width="150">등록일</td>
		    <td bgcolor="FFFFFF">
		    	<%= regdate %>
		    </td>
		</tr>
		<tr bgcolor="<%= adminColor("tabletop") %>">
		    <td width="150">최종수정</td>
		    <td bgcolor="FFFFFF">
		    	<%= lastadminid %>
		    	<Br><%= lastupdate %>
		    </td>
		</tr>
	<% end if %>
	
	<tr bgcolor="<%= adminColor("tabletop") %>">
	    <td colspan="2" align="center" bgcolor="FFFFFF">
			<% if themeidx<>"" then %>
	    		<input type="button" value="수정하기" onClick="Savetheme();" class="button">
	    	<% else %>	
	    		<input type="button" value="신규저장하기" onClick="Savetheme();" class="button">
	    	<% end if %>
	    </td>
	</tr>
	</table>
	</form>
	<br><br>
<% end if %>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 번호 : <input type="text" name="selectthemeidx" value="<%=selectthemeidx%>" maxlength="10" size="10" class="text">
		&nbsp;&nbsp;
		* 제목 : <input type="text" name="selecttitle" value="<%=selecttitle%>" maxlength="64" size="80" class="text">
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="gosubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 사용유무 :
		<% drawSelectBoxisusingYN "selectisusing", selectisusing, "" %>
		&nbsp;&nbsp;
		* 테마타입 :
		<% drawthemetype "selectthemetype", selectthemetype, "" %>
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" onClick="chsnewtheme('');" value="신규등록" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%=otheme.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= otheme.FTotalPage %></b>		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<!--<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>-->
	<td>테마번호</td>
	<td>테마타입</td>
	<td>테마명</td>
	<td>적용시간</td>
	<td>사용여부</td>
	<!--<td>정렬순서</td>-->
	<td>최종수정</td>
	<td>비고</td>
</tr>
<% if otheme.fresultcount > 0 then %>
<% For i = 0 to otheme.fresultcount -1 %>
<% if otheme.FItemList(i).fisusing="Y" then %>
	<tr bgcolor="#FFFFFF" align="center">
<% else %>
	<tr bgcolor="#f1f1f1" align="center">
<% end if %>	
	<!--<td><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<% '= otheme.FItemList(i).fthemeidx %>"></td>-->
	<td><%= otheme.FItemList(i).fthemeidx %></td>
	<td>
		<img src="http://imgstatic.10x10.co.kr/offshop/temp/2015/201502/ico_<%= otheme.FItemList(i).fthemetype %>.gif" width=40 height=40>
		<br><%= getthemetype(otheme.FItemList(i).fthemetype) %>
	</td>
	<td><%= ReplaceBracket(otheme.FItemList(i).ftitle) %></td>
	<td><%= otheme.FItemList(i).fexecutetime %></td>
	<td><%= otheme.FItemList(i).fisusing %></td>
	<!--<td><%= otheme.FItemList(i).forderno %></td>-->
	<td><%= otheme.FItemList(i).flastadminid %><Br><%= otheme.FItemList(i).flastupdate %></td>
	<td width=150>
		<input type="button" onClick="chselected('<%=otheme.FItemList(i).fthemeidx%>');" value="수정" class="button">
		<input type="button" class="button" value="상품" onclick="jsSetItem('<%= otheme.FItemList(i).fthemeidx %>','0');"/>
	</td>	
</tr>
<% Next %>
<tr bgcolor="FFFFFF" align="center">
	<td colspan="15">
       	<% If otheme.HasPreScroll Then %>
			<span class="otheme_link"><a href="gosubmit('<%= otheme.StartScrollPage-1 %>'); return false;">[pre]</a></span>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + otheme.StartScrollPage to otheme.StartScrollPage + otheme.FScrollCount - 1 %>
			<% If (i > otheme.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(otheme.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="gosubmit('<%= i %>'); return false;" class="otheme_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If otheme.HasNextScroll Then %>
			<span class="otheme_link"><a href="gosubmit('<%= i %>'); return false;">[next]</a></span>
		<% Else %>
			[next]
		<% End If %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set otheme = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->