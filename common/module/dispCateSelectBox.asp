<script type="text/javascript">
$(function(){
	chgDispCate('<%=dispCate%>');
});

function chgDispCate(dc) {
    setTimeout( function() {  //ios10 issue
    	$.ajax({
    		url: "/common/module/dispCateSelectBox_response.asp?disp="+dc,
    		cache: false,
    		async: false,
    		success: function(message) {
           		// 내용 넣기 
           		$("#lyrDispCtBox").empty().html(message);
           		$("#oDispCate").val(dc);
    		}
    	});
    }, 50);
}
</script>
<span id="lyrDispCtBox"></span>
<input type="hidden" name="disp" id="oDispCate" value="<%=dispCate%>">
