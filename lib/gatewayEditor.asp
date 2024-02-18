<%
'###########################################################
' Description :  탭스 에디터 업로드
' History : 2010.03.30 한용민 생성
'###########################################################
%>
<%
dim suid, sdisplayName, isize

suid = request("uid")
sdisplayName = request("displayName")
isize = request("size")

if suid <> "" and sdisplayName <>"" and isize <> "" then
%>
<script language="javascript">
	var uid ="<%=suid%>";
	var displayName ="<%=sdisplayName%>"
	var size = "<%=isize%>"
	
	var data = '{ uid: "' + uid + '", displayName: "' + displayName + '", size: "' + size + '" }';  
	opener.parent.OnCompleteImageUpload(data); 
	self.close();
</script>	
<%
else
%>
<script language="javascript">
	alert("데이터전송에 문제가 발생하였습니다. 다시 시도해 주십시오");
	self.close();
</script>
<%	
end if
%>