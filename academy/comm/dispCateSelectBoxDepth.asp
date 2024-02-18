<script type="text/javascript">
$(function(){
	chgDispCate('<%=dispCate%>','<%=maxDepth%>');
});

function chgDispCate(dc, maxDepth) { 
	 setTimeout( function() {  //ios10 issue 
		$.ajax({
			url: "/academy/comm/dispCateSelectBoxDepth_response.asp?disp="+dc+"&maxD="+maxDepth,
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
