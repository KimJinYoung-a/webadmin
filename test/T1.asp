<%



%>
<script language='javascript'>
function ss(){
    if (confirm('ok?')){
        document.frm.submit();
    }
}
</script>
<form name="frm" method="post" action="https://webadmin.10x10.co.kr/test/t2.asp">
<input type="hidden" name="aa" value="aa1">
<input type="hidden" name="bb" value="bb1">
<input type="button" value="TTT" onClick="ss()">
</form>