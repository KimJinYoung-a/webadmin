<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 포스페이지 관리
' History : 2009.09.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/academy/lib/classes/sitemaster/sitemaster_cls.asp"-->

<%
dim linktype, fixtype
dim poscode, page, listIsUsing, gubun, listGubun

poscode = RequestCheckvar(request("poscode"),10)
page = RequestCheckvar(request("page"),10)
listIsUsing = RequestCheckvar(request("lUse"),1)
listGubun = RequestCheckvar(request("lGubun"),24)
gubun = RequestCheckvar(request("gubun"),24)

if poscode="" then poscode=0
if page="" then page=1
if listIsUsing="" then listIsUsing="Y"
if listGubun="" then listGubun="index"
if gubun="" then gubun="index"

dim oposcode,oposcodeList

set oposcode = new cposcode_list
	oposcode.FRectPosCode = poscode
	oposcode.fposcode_oneitem

set oposcodeList = new cposcode_list
	oposcodeList.FPageSize=10
	oposcodeList.FCurrPage= page
	oposcodeList.FListIsUsing= listIsUsing
	oposcodeList.FListGubun= listGubun
	oposcodeList.fposcode_list

dim i
%>

<script language='javascript'>
function SavePosCode(frm){
    if (frm.poscode.value.length<1){
        alert('구분 코드 값을 입력하세요.');
        frm.poscode.focus();
        return;
    }
    
    if (frm.poscode.value*1<1){
        alert('구분 코드 값은 1 이상입니다.');
        frm.poscode.focus();
        return;
    }
    
    if (frm.posname.value.length<1){
        alert('구분명을 입력하세요.');
        frm.posname.focus();
        return;
    }
    
    
    if (frm.imagetype.value.length<1){
        alert('링크타입을 선택하세요.');
        frm.imagetype.focus();
        return;
    }
    
    if (frm.imagewidth.value.length<1){
        alert('이미지 사이즈 가로를 입력하세요.');
        frm.imagewidth.focus();
        return;
    }
    
    if (frm.imageheight.value.length<1){
        alert('이미지 사이즈 세로를 입력하세요.');
        frm.imageheight.focus();
        return;
    }
        
    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
    
}

function ChangeLinktype(){
    // Do nothing
}
</script>

<!--표 헤드시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>코드 입력</strong></font>
			</td>		
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr height="10" valign="top">
		<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_06.gif"></td>
		<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
	</tr>
</table>
<!--표 헤드끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frmposcode" method="post" action="/academy/sitemaster/imagemake_poscode_process.asp" >
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">그룹구분</td>
    <td>
    <%
    	if oposcode.FOneItem.Fgubun<>"" then
    		Response.Write oposcode.FOneItem.Fgubun
    		Response.Write "<input type='hidden' name='gubun' value='" & oposcode.FOneItem.Fgubun & "' >"
     	else
    		call DrawGroupGubunCombo ("gubun", gubun, "")
    	end if
    %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">구분코드</td>
    <td>
        <% if oposcode.FOneItem.Fposcode<>"" then %>
        <%= oposcode.FOneItem.Fposcode %>
        <input type="hidden" name="poscode" value="<%= oposcode.FOneItem.Fposcode %>" >
        <% else %>
        <input type="text" name="poscode" value="<%= oposcode.FOneItem.Fposcode %>" maxlength="7" size="5">
        (숫자)
        <% end if %>         
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">구분명</td>
    <td>
        <input type="text" name="posname" value="<%= oposcode.FOneItem.Fposname %>" maxlength="32" size="64">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">사용할이미지수</td>
    <td>
        <input type="text" name="imagecount" value="<%= oposcode.FOneItem.fimagecount %>" maxlength="2" size="2">
        (숫자만 입력 하세요)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">이미지 너비</td>
    <td>
        <input type="text" name="imagewidth" value="<%= oposcode.FOneItem.Fimagewidth %>" maxlength="16" size="8">
        (이미지 Width Size 숫자)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">이미지 높이</td>
    <td>
        <input type="text" name="imageheight" value="<%= oposcode.FOneItem.Fimageheight %>" maxlength="16" size="8">
        (이미지 Height Size 숫자 : 0 인경우 height 지정 안함)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">링크타입</td>
    <td>
        <select name="imagetype">
	        <option value="" <% if oposcode.FOneItem.fimagetype = "" then response.write " selected" %>>선택</option>
	        <option value="map" <% if oposcode.FOneItem.fimagetype = "map" then response.write " selected" %>>map</option>
	        <option value="link" <% if oposcode.FOneItem.fimagetype = "link" then response.write " selected" %>>link</option>                
	        <option value="flash" <% if oposcode.FOneItem.fimagetype = "flash" then response.write " selected" %>>flash</option>
	        <option value="multi" <% if oposcode.FOneItem.fimagetype = "multi" then response.write " selected" %>>multi</option>
	        <option value="xml" <% if oposcode.FOneItem.fimagetype = "xml" then response.write " selected" %>>xml</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">사용여부</td>
    <td>
        <% if oposcode.FOneItem.Fisusing="N" then %>
        <input type="radio" name="isusing" value="Y">사용함
        <input type="radio" name="isusing" value="N" checked >사용안함
        <% else %>
        <input type="radio" name="isusing" value="Y" checked >사용함
        <input type="radio" name="isusing" value="N">사용안함
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SavePosCode(frmposcode);"></td>
</tr>
</form>
</table>
<%
set oposcode = Nothing
%>
<br>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td colspan="6">
    	<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
    	<tr>
    		<td>
    			사용여부:
    			<select name="lUse" class="select" onchange="self.location='?lUse='+this.value+'&lGubun=<%=listGubun%>';">
    			<option value="A">전체</option>
    			<option value="Y">사용</option>
    			<option value="N">삭제</option>
    			</select>
    			<script language="javascript">
    			document.getElementById("lUse").value='<%=listIsUsing%>';
    			</script>
    			&nbsp;
    			그룹구분:
    			<% call DrawGroupGubunCombo ("lGubun", listGubun, "onchange=""self.location='?lUse=" & listIsUsing & "&lGubun='+this.value;""") %>
    		</td>
    		<td align="right"><a href="?poscode=&lGubun=<%=listgubun%>&gubun=<%=listgubun%>"><img src="/images/icon_new_registration.gif" border="0"></a></td>
    	</tr>
    	</table>
    </td>
</tr>
<% if oposcodeList.FResultCount > 0 then %>
	<tr bgcolor="#DDDDFF" align="center">
	    <td width="70">그룹</td>
	    <td width="50">code</td>
	    <td>구분명</td>
	    <td width="60">링크타입</td>
	    <td width="60">Image수</td>
	    <td width="60">사용여부</td>
	</tr>
	<% for i=0 to oposcodeList.FResultCount-1 %>
		<% if (CStr(oposcodeList.FItemList(i).FposCode)=poscode) then %>
		<tr bgcolor="#9999CC">
		<% else %>
		<tr bgcolor="#FFFFFF">
		<% end if %>
		    <td align="center"><%= oposcodeList.FItemList(i).Fgubun %></td>
		    <td align="center"><%= oposcodeList.FItemList(i).FposCode %></td>
		    <td ><a href="?poscode=<%= oposcodeList.FItemList(i).FposCode %>&page=<%= page %>&lGubun=<%=listgubun%>"><%= oposcodeList.FItemList(i).FposName %></a></td>
		    <td align="center"><%= oposcodeList.FItemList(i).fimagetype %></td>
		    <td align="center"><%= oposcodeList.FItemList(i).fimagecount %></td>
		    <td align="center"><%= oposcodeList.FItemList(i).Fisusing %></td>
		</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td align="center">내용이 없습니다.</td>
	</tr>	
<% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
       <td valign="bottom" align="center">
		    <% if oposcodeList.HasPreScroll then %>
				<a href="?page=<%= oposcodeList.StartScrollPage-1 %>&lUse=<%=listIsUsing%>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
		
			<% for i=0 + oposcodeList.StartScrollPage to oposcodeList.FScrollCount + oposcodeList.StartScrollPage - 1 %>
				<% if i>oposcodeList.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&lUse=<%=listIsUsing%>">[<%= i %>]</a>
				<% end if %>
			<% next %>
		
			<% if oposcodeList.HasNextScroll then %>
				<a href="?page=<%= i %>&lUse=<%=listIsUsing%>">[next]</a>
			<% else %>
				[next]
			<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set oposcodeList = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->