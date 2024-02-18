<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이미지 관리
' History : 2016.07.28 서동석 생성
'			2016.08.12 한용민 수정
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
dim i, reload, folderidx, delyn, page
	folderidx = getNumeric(requestCheckVar(request("folderidx"),10))
	delyn = requestCheckVar(request("delyn"),32)
	page = getNumeric(requestCheckVar(request("page"),10))
	reload = requestCheckVar(request("reload"),32)

if (page="") then page=1
if reload="" and delyn="" then delyn="N"

dim oEtcImage
SET oEtcImage = new CEtcImageManage
	oEtcImage.FPageSize = 30
	oEtcImage.FCurrPage = page
	oEtcImage.FRectFolderidx = folderidx
	oEtcImage.FRectDelYN = delyn
	oEtcImage.getEtcImageList

%>
<script type="text/javascript">

document.domain = '10x10.co.kr';

function jsSearch(page){
	var frm = document.frmSearch;

	frm.page.value=page;
	frm.submit();
}

function popregImage(folderidx, etcimgidx){
    var popwin;
    var folderidx = folderidx;
    if ('<%= folderidx %>'==''){
        alert('검색조건- 구분을 먼저 검색후에 사용하세요.');
        return;
    }

    popwin = window.open('/admin/eventmanage/etcimage/popImageReg.asp?folderidx=<%= folderidx %>&etcimgidx='+etcimgidx+'&menupos=<%=menupos%>','popImageReg','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();	
}

function jsImgCodeManage(){
	var jsImgCodeManage;
	jsImgCodeManage = window.open('/admin/eventmanage/etcimage/manager/manager.asp?menupos=<%=menupos%>','jsImgCodeManage','width=1280,height=960');
	jsImgCodeManage.focus();
}

//링크 복사
function copyLink(imagepath) {
	clipboardData.setData("Text", imagepath);
	alert('선택하신 파일의 링크 경로가 복사되었습니다. 사용하실 곳에 Ctrl+V 하시면됩니다.');
} 

</script>

<form name="frmSearch" method="get"  action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="reload" value="ON">
<input type="hidden" name="iC">

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		구분 : <% Call sbDrawEtcImgGbn("folderidx",folderidx, " onchange='jsSearch("""");'") %>
		&nbsp;
		삭제여부 : <% drawSelectBoxisusingYN "delyn", delyn, " onchange='jsSearch("""");'" %>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="jsSearch('');">
	</td>
</tr>
</table>
<!-- 표 상단바 끝-->

</form>

<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
<tr height="40" valign="bottom">
    <td align="left">
    </td>
    <td align="right">
    	<input type="button" value="새로등록" onclick="popregImage('','');" class="button">

		<% if C_ADMIN_AUTH then %>
			<input type="button" onclick="jsImgCodeManage();" value="코드관리" class="button">
		<% END IF %>
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="20">검색결과 : <b><%=oEtcImage.FTotalCount%></b>&nbsp;&nbsp;페이지 : <b><%=page%> / <%=oEtcImage.FTotalPage%></b></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">일렬번호</td>
	<td width="150">구분</td>
	<td width="50">이미지</td>
	<td>경로</td>
  	<td width="140">최종수정</td>
  	<td width="130">비고</td>
</tr>

<% if (oEtcImage.FTotalCount>0) then %>
	<% for i=0 to oEtcImage.FResultCount -1 %>
	<% if isnull(oEtcImage.FItemList(i).Fdeldt) or oEtcImage.FItemList(i).Fdeldt="" then %>
		<tr align="center" bgcolor="#FFFFFF">
	<% else %>
		<tr align="center" bgcolor="#e1e1e1">
	<% end if %>

		<td><%= oEtcImage.FItemList(i).FetcimgIdx %></td>
	    <td><%= oEtcImage.FItemList(i).FfolderTitle %></td>
		<td>
			<% if oEtcImage.FItemList(i).Fimagename <> "" then %>
				<img src="<%= webImgUrl %>\<%= oEtcImage.FItemList(i).FrealPath %>\<%= oEtcImage.FItemList(i).Fsubfolder %>\<%= oEtcImage.FItemList(i).Fimagename %>" width=50 height=50>
			<% end if %>
		</td>
		<td align="left">
			<% if oEtcImage.FItemList(i).Fimagename <> "" then %>
				<%= webImgUrl %>\<%= oEtcImage.FItemList(i).FrealPath %>\<%= oEtcImage.FItemList(i).Fsubfolder %>\<%= oEtcImage.FItemList(i).Fimagename %>
			<% end if %>
	    </td>
	    <td>
			<% if oEtcImage.FItemList(i).Flastuserid<>"" then %>
				<%= oEtcImage.FItemList(i).Flastupdate %><Br>(<%= oEtcImage.FItemList(i).Flastuserid %>)
			<% end if %>
	    </td>
	    <td>
	    	<input type="button" onclick="popregImage('<%=oEtcImage.FItemList(i).Ffolderidx%>','<%=oEtcImage.FItemList(i).FetcimgIdx%>');" value="수정" class="button">

	    	<% if oEtcImage.FItemList(i).Fimagename <> "" then %>
	    		<input type="button"  id="btnLink" class="button" value="경로복사" title="경로복사" onClick="copyLink('<%= replace(webImgUrl,"\","/") %>/<%= replace(oEtcImage.FItemList(i).FrealPath,"\","/") %>/<%= replace(oEtcImage.FItemList(i).Fsubfolder,"\","/") %>/<%= oEtcImage.FItemList(i).Fimagename %>')">
	    	<% end if %>
	    </td>
	</tr>
	<% next %>

	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="11">
		    <% if oEtcImage.HasPreScroll then %>
				<a href="#" onClick="jsSearch('<%= oEtcImage.FStarScrollPage-1 %>');">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
	
			<% for i=0 + oEtcImage.StartScrollPage to oEtcImage.FScrollCount + oEtcImage.StartScrollPage - 1 %>
				<% if i>oEtcImage.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
					<font color="red">[<%= i %>]</font>
				<% else %>
					<a href="#" onClick="jsSearch('<%= i %>');">[<%= i %>]</a>
				<% end if %>
			<% next %>
	
			<% if oEtcImage.HasNextScroll then %>
				<a href="#" onClick="jsSearch('<%= i %>');">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="20">등록된 내용이 없습니다.</td>
	</tr>
<% end if %>

</table>

<%
SET oEtcImage = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->