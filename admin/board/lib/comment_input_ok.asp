<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual ="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/10x10_board_commentcls.asp" -->
<%
 Dim ocomm,NextPage,idx,comment,userid,username
 dim jukyocd,jukyo,usemileage
  
 NextPage = request("NextPage")
 userid = request("userid")
 username = request("username")
 comment = request("comment")
 idx = request("idx")
 
 set ocomm = new CComment
 ocomm.Comment_input idx,userid,username,comment

%>

<script language="JavaScript">
<!--

 location.href = "<% = NextPage %>"

//-->
</script>

<% 
set ocomm = Nothing
%>
<!-- #include virtual ="/lib/db/dbclose.asp" -->