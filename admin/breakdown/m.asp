<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/breakdown/breakdownCls.asp"-->
<%
	Dim vIdx, webImgUrl, cBreakview, vReqEquipment, vWorkType, vWorkTarget, vReqComment, vReqCapImage1, vTeam, vReqName, vReqDate, vReqState
	vIdx = requestCheckVar(Request("a"),5)
	
	IF application("Svr_Info") = "Dev" THEN
		webImgUrl = "http://testwebimage.10x10.co.kr"
	Else
		webImgUrl = "http://webimage.10x10.co.kr"
	End If
	
	If vIdx = "" Then
		Response.Write "<script>alert('�߸��� �����Դϴ�.')</script>"
		dbget.close()
		Response.End
	End If
	
	If IsNumeric(vIdx) = false Then
		Response.Write "<script>alert('�߸��� �����Դϴ�.');</script>"
		dbget.close()
		Response.End
	End If

	Set cBreakview = New CBreakdown
	 	cBreakview.FReqIdx = vIdx
		cBreakview.fnGetBreakdownMobileView
		
		vReqEquipment 	= cBreakview.FReqEquipment
		vWorkType		= cBreakview.FWorkType
		vWorkTarget		= Replace(cBreakview.FWorkTarget,"_break","")
		vReqComment		= cBreakview.FReqComment
		vReqCapImage1	= cBreakview.FReqCapImage1
		vTeam			= cBreakview.FTeam
		vReqName		= cBreakview.FReqName
		vReqDate		= cBreakview.FReqDate
		vReqState		= cBreakview.FReqState
	Set cBreakview = Nothing
%>

<script language="javascript">
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','scrollbars=yes,resizable=yes');
	image_view.focus();
}
</script>

�۾���ȣ : <%=vIdx%>(<%=fnWorkState(vReqState)%>)<br>
�̸� : <%=vReqName%>(<%=vTeam%>)<br>
�۾����� : <%=fnWorkType(vWorkType)%>(<%=fnWorkTargetName(vWorkTarget)%>)<br>
����ڵ� : <%=CHKIIF(vWorkType<>"3",CommonCode("v",vWorkTarget,vReqEquipment),"")%><br>
�ڸ�Ʈ : <%=vReqComment%><br>
ĸ���̹��� : <a href="javascript:image_view('<%=webImgUrl%>/breakdown<%=vReqCapImage1%>');" onfocus="this.blur()"><img src="<%=webImgUrl%>/breakdown<%=vReqCapImage1%>" width="100" border="0"></a><br>

<!-- #include virtual="/lib/db/dbclose.asp" -->