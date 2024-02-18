<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

dim userimage

userimage = requestCheckVar(request("userimage"), 128)

%>
<script language="javascript">

opener.focus();
opener.document.frm_base.userimage.value = "<%= userimage %>";
opener.SaveUserImage();
window.close();

</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->