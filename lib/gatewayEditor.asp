<%
'###########################################################
' Description :  �ǽ� ������ ���ε�
' History : 2010.03.30 �ѿ�� ����
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
	alert("���������ۿ� ������ �߻��Ͽ����ϴ�. �ٽ� �õ��� �ֽʽÿ�");
	self.close();
</script>
<%	
end if
%>