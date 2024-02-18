<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<script language="JavaScript" src="/cscenter/js/cscenter.js?v=1.1"></script>
<script type="text/javascript">
PopMyQna('', '', '<%=CHKIIF(session("ssBctId")="oesesang52","V","N")%>','','','','','','');
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->