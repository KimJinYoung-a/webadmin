<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual ="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/10x10_board_commentcls.asp" -->
<%
 Dim ocomm,NextPage,comment,uname,pass,idx
  
 idx = request("idx")
 
 set ocomm = new CComment
 ocomm.Comment_delete idx

%>

<script language="JavaScript">
<!--
 alert("삭제되었습니다!");
 history.back();
//-->
</script>

<% 
set ocomm = Nothing
%>
<!-- #include virtual ="/lib/db/dbclose.asp" -->