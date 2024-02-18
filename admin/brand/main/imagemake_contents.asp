<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.29 김진영 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/street/brandmainCls.asp" -->
<%
Dim idx, reload, ix, gubun
	idx = request("idx")
	gubun = request("gubun")
	reload = request("reload")
	if idx="" then idx=0

If reload="on" then
    response.write "<script>opener.location.reload(); window.close();</script>"
    dbget.close()	:	response.End    
End If

Dim mbrand
Set mbrand = New cBrandMain
	mbrand.FRectIdx = idx
	mbrand.FRectGubun = gubun
	mbrand.sMainTop3modify
%>
<script language='javascript'>
function SaveMainContents(frm){
    if (frm.image_order.value.length<1){
        alert('이미지 우선순위를 입력 하세요.');
        frm.image_order.focus();
        return;
    }
    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
}
function SaveMainContents2(frm){
    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmcontents" method="post" action="<%=staticUploadUrl%>/linkweb/street/doMainimage_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="gubun" value="<%=gubun%>">
<input type="hidden" name="ckUserId" value="<%=session("ssBctId")%>">
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">Idx :</td>
	    <td><%= mbrand.FOneItem.Fidx %><input type="hidden" name="idx" value="<%= mbrand.FOneItem.Fidx %>"></td>
	</tr>
<% If gubun = 1 Then %>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">이미지정렬우선순위 :</td>
	    <td>
	    	<input type="text" size="2" maxlength="2" name="image_order"  class="text" value=<%= chkiif(mbrand.FOneItem.Fidx = "","99",mbrand.FOneItem.FImage_order) %>>
	    	&nbsp;실서버 적용시 숫자가 작을경우 우선노출
	    </td>
	</tr>
<% End If %>
	<tr bgcolor="#FFFFFF">
		<td width="150" align="center"><%=chkIIF(gubun="1","롤링 배너 이미지 : ","이미지 : ")%></td>
		<td><input type="file" name="file1" value="" size="32" maxlength="32" class="file">
			<br><font color="RED">1140 x 500픽셀의 이미지만 등록하세요</font>
		<% If mbrand.FOneItem.Fidx <> "" Then %>
		<br><img src="<%=uploadUrl%>/brandstreet/main/<%= mbrand.FOneItem.fimagepath %>" border="0"> 
		<br><%=uploadUrl%>/brandstreet/main/<%= mbrand.FOneItem.fimagepath %>
		<% End If %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">링크값 :</td>
	    <td>
			<input type="text" name="linkpath" value="<%= mbrand.FOneItem.flinkpath %>" maxlength="80" size="80"><br>(상대경로로 표시해 주세요  ex: /street/street_brand_sub01.asp?makerid=ithinkso)
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">사용여부</td>
	    <td>
			<input type="radio" name="isusing" value="Y" checked <%=chkIIF(mbrand.FOneItem.fisusing = "Y","checked","")%> >Y
			<input type="radio" name="isusing" value="N" <%=chkIIF(mbrand.FOneItem.fisusing = "N","checked","")%>>N
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td align="center" colspan=2>
	    	<% If gubun = 1 Then %>
	    	<input type="button" value=" 저 장 " onClick="SaveMainContents(frmcontents);" class="button">
	    	<% Else %>
	    	<input type="button" value=" 저 장 " onClick="SaveMainContents2(frmcontents);" class="button">
	    	<% End If%>
	    </td>
	</tr>	
</form>
</table>
<%
set mbrand = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
