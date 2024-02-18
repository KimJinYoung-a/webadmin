<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이미지 관리
' History : 2016.08.12 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/event/etcImageMngCls.asp"-->

<%
dim reguserid, lastuserid, page, i, folderidx, folderTitle, realPath, sortkey, isusing
	menupos = getNumeric(requestcheckvar(request("menupos"),10))
	page = getNumeric(requestcheckvar(request("page"),10))
	folderidx = getNumeric(requestcheckvar(request("folderidx"),10))

if page="" then page=1

dim oEtcone
SET oEtcone = new CEtcImageManage
	oEtcone.FRectfolderIdx = folderidx

	if (folderidx<>"") then
		oEtcone.getEtcImage_masterone

		if oEtcone.FResultCount > 0 then
	        folderidx = oEtcone.FOneItem.FfolderIdx
            foldertitle = oEtcone.FOneItem.FfolderTitle
            realpath = oEtcone.FOneItem.FrealPath
            sortkey = oEtcone.FOneItem.Fsortkey
            isusing = oEtcone.FOneItem.Fisusing
		end if
	end if

dim oEtcImage
SET oEtcImage = new CEtcImageManage
	oEtcImage.FPageSize = 30
	oEtcImage.FCurrPage = page
	oEtcImage.getEtcImagemasterList

if sortkey="" then sortkey=100
%>

<script type="text/javascript">

function deletcimage(frm){
    if (frm.folderidx.value==''){
        alert('구분자가 없습니다.');
        frm.folderidx.focus();
        return;
    }

    if (confirm('삭제 하시겠습니까?')){
    	frm.mode.value='etcimgdel';
        frm.submit();
    }
}

function Saveetcimage(frm){
    if (frm.foldertitle.value==''){
        alert('구분명을 선택하세요.');
        frm.foldertitle.focus();
        return;
    }

    if (frm.realpath.value==''){
        alert('적용 경로를 입력하세요.');
        frm.realpath.focus();
        return;
    }

    if (confirm('저장 하시겠습니까?')){
    	frm.mode.value='etcimgedit';
        frm.submit();
    }
}

//신규등록
function newetcimage(){
	location.href='/admin/eventmanage/etcimage/manager/manager.asp?menupos=<%=menupos%>'
}

</script>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmetcimg" method="post" action="/admin/eventmanage/etcimage/manager/manager_process.asp">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode">
<tr bgcolor="#FFFFFF" align="center">
    <td width="150">번호</td>
    <td align="left">
    	<%= folderidx %>
		<input type="hidden" name="folderidx" value="<%= folderidx %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="150">구분명</td>
    <td align="left">
    	<input type="text" name="foldertitle" value="<%= foldertitle %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="150">적용 경로</td>
    <td align="left">
		<input type="text" name="realpath" value="<%= realpath %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="150">정렬</td>
    <td align="left">
		<input type="text" name="sortkey" value="<%= sortkey %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td>사용여부</td>
    <td align="left">
        <% drawSelectBoxisusingYN "isusing", isusing, "" %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>
    	<%
    	'//수정모드
    	if folderidx <> "" then
    	%>
			<input type="button" value="삭제" onClick="deletcimage(frmetcimg);" class="button">
		<% end if %>
	</td>
    <td>
    	<input type="button" value="저장" onClick="Saveetcimage(frmetcimg);" class="button">
    </td>
</tr>
</form>
</table>

<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    </td>
    <td align="right">
    	<input type="button" onclick="newetcimage();" value="신규등록" class="button">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= oEtcImage.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oEtcImage.FTotalPage %></b>
	</td>
</tr>
<% if oEtcImage.FResultCount > 0 then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>번호</td>
	<td>구분명</td>
	<td>적용 경로</td>
	<td>정렬</td>
    <td>사용여부</td>
    <td>비고</td>
</tr>
<% for i=0 to oEtcImage.FResultCount-1 %>

<% if oEtcImage.FItemList(i).ffolderidx = folderidx then %>
	<tr bgcolor="orange" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="orange"; align="center">
<% else %>
	<tr bgcolor="#ffffff" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF"; align="center">
<% end if %>
	<td><%= oEtcImage.FItemList(i).Ffolderidx %></td>
	<td><%= oEtcImage.FItemList(i).FfolderTitle %></td>
	<td><%= oEtcImage.FItemList(i).FrealPath %></td>
    <td><%= oEtcImage.FItemList(i).Fsortkey %></td>
    <td><%= oEtcImage.FItemList(i).Fisusing %></td>
    <td width=60>
    	<input type="button" onclick="location.href='?folderidx=<%= oEtcImage.FItemList(i).ffolderidx %>&page=<%= page %>'" value="수정" class="button">
    </td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
	    <% if oEtcImage.HasPreScroll then %>
			<a href="?page=<%= oEtcImage.StartScrollPage-1 %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oEtcImage.StartScrollPage to oEtcImage.FScrollCount + oEtcImage.StartScrollPage - 1 %>
			<% if i>oEtcImage.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oEtcImage.HasNextScroll then %>
			<a href="?page=<%= i %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>

<% else %>
<tr bgcolor="#FFFFFF">
	<td align="center">내용이 없습니다.</td>
</tr>
<% end if %>
</table>

<%
set oEtcone = Nothing
set oEtcImage = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
