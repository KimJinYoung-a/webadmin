<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/topbrand/topbrandtalkcls.asp" -->
<%

dim i, page, iscurrtopbrand

page = requestCheckVar(request("page"),10)
if (page = "") then
    page = "1"
end if


'==============================================================================
dim otopbrandtalklist
set otopbrandtalklist = New CTopBrandTalk

iscurrtopbrand = otopbrandtalklist.IsCurrentTopBrand(session("ssBctId"))

otopbrandtalklist.FRectMakerID = session("ssBctId")
otopbrandtalklist.FCurrPage = page
'otopbrandtalklist.FRectIsCurrentTopBrand = "Y"

otopbrandtalklist.GetTopBrandTalkList

%>
<script>
function SubmitWrite()
{
    if (frm.imagetalk.value.length < 1) {
        alert("내용을 입력하세요.");
        return;
    }

    if (frm.image1.value.length < 1) {
        alert("이미지를 올리세요.");
        return;
    }

    if (frm.image1.fileSize > 1000000) {
        alert("이미지는 1메가를 넘을수 없습니다.");
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
    <form name="frm" action="<%=uploadUrl%>/linkweb/doTopBrandTalk.asp" method=post onsubmit="return false" enctype="multipart/form-data">
    <input type=hidden name=menupos value="<%= menupos %>">
    <input type=hidden name=mode value="write">
    
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<input type="button" class="icon" value="#">
			<font color="red"><b>브랜드토크 등록</b></font>
			&nbsp;
			<% if (iscurrtopbrand = false) then %>
        	<font color=red><b>현재 탑브랜드가 아닙니다.</b></font>
			<% end if %>
		</td>
	</tr>
    
    <tr align="center">
        <td width="80" bgcolor="<%= adminColor("tabletop") %>">내용입력</td>
        <td align="left" bgcolor="#FFFFFF">
        	<textarea class="textarea" name="imagetalk" cols="75" rows="5"></textarea>
        </td>
	</tr>
	<tr align="center">
        <td bgcolor="<%= adminColor("tabletop") %>">이미지등록</td>
        <td align="left" bgcolor="#FFFFFF">
        	<input type="file" class="file" name=image1 size=40><br>(1메가 이하, 570px × 248px 크기의 이미지만 업로드가 가능합니다.)
        </td>
	</tr>
    </form>
    
	<form name="f" action="brandtalk_write.asp" method=get onsubmit="return false">
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

set otopbrandtalklist = Nothing

%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->