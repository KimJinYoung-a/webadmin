<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/topbrand/topbrandnewscls.asp" -->
<%

dim i, page, iscurrtopbrand

page = requestCheckVar(request("page"),10)
if (page = "") then
    page = "1"
end if


'==============================================================================
dim otopbrandnewslist
set otopbrandnewslist = New CTopBrandNews

iscurrtopbrand = otopbrandnewslist.IsCurrentTopBrand(session("ssBctId"))

otopbrandnewslist.FRectMakerID = session("ssBctId")
otopbrandnewslist.FCurrPage = page
'otopbrandnewslist.FRectIsCurrentTopBrand = "Y"

otopbrandnewslist.GetTopBrandNewsList

%>
<script>
function SubmitWrite()
{
    if (frm.title.value.length < 1) {
        alert("제목을 입력하세요.");
        return;
    }

    if (frm.contents.value.length < 1) {
        alert("내용을 입력하세요.");
        return;
    }

    if (confirm("저장하시겠습니까?") == true) {
        frm.submit();
    }
}

function FileCheck(comp,maxfilesize,maxwidth,maxheight){
	if(comp.fileSize > maxfilesize){
		alert("파일사이즈는 "+ maxfilesize + "byte를 넘기실 수 없습니다...");
		return false;
	}

	if ((comp.src!="")&&(comp.width <1)){
		alert('이미지만 가능합니다.');
		return false;
	}

	return true;
}
</script>


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" action="brandnews_process.asp" method=post onsubmit="return false">
    <input type=hidden name=menupos value="<%= menupos %>">
    <input type=hidden name=mode value="write">
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<input type="button" class="icon" value="#">
			<font color="red"><b>브랜드뉴스 등록</b></font>
			&nbsp;
			<% if (iscurrtopbrand = false) then %>
	        	<font color=red><b>현재 탑브랜드가 아닙니다.</b></font>
			<% end if %>
		</td>
	</tr>
	<tr align="center">
        <td width="80" bgcolor="<%= adminColor("tabletop") %>">제목</td>
        <td align="left" bgcolor="#FFFFFF">
        	<input type="text" class="text" name="title" size=75>
        </td>
	</tr>
	<tr align="center">
        <td bgcolor="<%= adminColor("tabletop") %>">내용</td>
        <td align="left" bgcolor="#FFFFFF">
        	<textarea class="textarea" name="contents" cols="75" rows="5"></textarea>
        </td>        
	</tr>
	</form>

	<form name="f" action="brandnews_write.asp" method=get onsubmit="return false">
    <input type=hidden name=menupos value="<%= menupos %>">
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
          <input type="button" class="button" value="등록하기" onClick="SubmitWrite();">
          <input type="button" class="button" value="취소하기" onClick="history.back();">
	    </td>
    </tr>
    </form>
</table>
<!-- 표 하단바 끝-->
<%

set otopbrandnewslist = Nothing

%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->