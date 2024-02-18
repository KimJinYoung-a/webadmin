<script type="text/javascript">

function chgDispCate(dc, maxDepth, partsort) { 
	 setTimeout( function() {  //ios10 issue 
		$.ajax({
			url: "/common/module/dispCateSelectBoxDepth_response2.asp?disp="+dc+"&maxD="+maxDepth+"&partsort="+partsort,
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
<script type="text/javascript">
	chgDispCate('<%=dispCate%>','<%=maxDepth%>','<%=catesort%>');
</script>