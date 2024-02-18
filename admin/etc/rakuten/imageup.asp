<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<% 
dim i,cnt

cnt=3
%>
<script language='javascript'>
   function jsUp(ii){
        var frm  = document.frmA;
        var itemid = frm.itemid[ii].value;
        var partneritemname = frm.partneritemname[ii].value;
        var istest = frm.istest.value;
        
        if (itemid.length<1){
            alert('상품코드 입력요망');
            frm.itemid[ii].focus();
            return;
        }
        
        if (partneritemname.length<1){
            alert('제휴사관리코드 입력요망');
            frm.partneritemname[ii].focus();
            return;
        }
      //  alert(istest);return;
        var upfrm = document.frmUp;
        
        if (confirm('저장하시겠습니까?')){
            upfrm.itemid.value=itemid;
            upfrm.partneritemname.value=partneritemname;
            upfrm.istest.value=istest;
            
            upfrm.target="iifrm";
            <% if (application("Svr_Info")="Dev") then %>
            upfrm.action="http://testimgstatic.10x10.co.kr/linkweb/rakuten/uprakuten_proc.asp";
            <% else %>
            upfrm.action="http://imgstatic.10x10.co.kr/linkweb/rakuten/uprakuten_proc.asp";
            <% end if %>
            upfrm.submit();
        }
   }
</script>
    
<form name="frmA">
<table width="100%" border="0" class="a">
<tr>
    <td>상품코드</td>
    <td>제휴사관리코드</td>
    <td><input type="checkbox" name="istest" value="on"> 테스트</td>
</tr>
<% for i=0 to cnt-1 %>
<tr>
    <td><input type="text" name="itemid" value="" size="10"></td>
    <td><input type="text" name="partneritemname" value="" size="30"></td>
    <td><input type="button" value="이미지업로드" onClick="jsUp('<%=i%>')"></td>
</tr>
<% next %>
</table>
</form>
<p>
<form name="frmUp">
<input type="hidden" name="itemid">
<input type="hidden" name="partneritemname">
<input type="hidden" name="istest">
</form>    
<table width="100%" border="0" class="a">
<tr>
    <iframe width="100%" height="100%" id="iifrm" name="iifrm">
</tr>
 </table>
<!-- #include virtual="/admin/lib/poptail.asp"-->