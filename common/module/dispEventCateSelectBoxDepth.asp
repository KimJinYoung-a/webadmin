<script type="text/javascript">
$(function(){
	chgDispCate('<%=dispCate%>','<%=maxDepth%>');
});

function chgDispCate(dc, maxDepth) {
	if(dc==""){
		dc="0";
	}
	 setTimeout( function() {  //ios10 issue 
		$.ajax({
			url: "/common/module/dispEventCateSelectBoxDepth_response.asp?disp="+dc+"&maxD="+maxDepth,
			cache: false,
			async: false,
			success: function(message) {
	       		// ���� �ֱ�
	       		$("#lyrDispCtBox").empty().html(message);
	       		$("#oDispCate").val(dc);
			}
		}); 
		 }, 50);
}
</script>
<span id="lyrDispCtBox"></span>
<input type="hidden" name="disp" id="oDispCate" value="<%=dispCate%>">