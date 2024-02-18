<%
dim userimage : userimage = request("userimage")
%>

<script language="javascript">
<!--	
	opener.document.frm_base.userimage.value = "<%=userimage%>";
	opener.SaveUserImage();
	window.close();	
-->
</script>
