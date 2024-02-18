<%
'###########################################################
' Description : µ•¿Ã≈Õ∫–ºÆ
' History : 2016.01.29 «—øÎπŒ ª˝º∫
'###########################################################
%>
<%
dim calyyyy, calmm
	calyyyy = Request("calyyyy")
	calmm = Request("calmm")

if calyyyy="" then calyyyy=Year(now)
if calmm="" then calmm=Month(now)

%>
<style type="text/css">
	.calBtn { height:20px; border-radius:6px; border: 1px solid #c0c0c0; font-size:18px; font-family:tahoma; }
</style>

<script type="text/javascript">

	$(function(){
		gopurpose('<%= calyyyy %>','<%= calmm %>');

		//setTimeout("gomaechul('<%'= startdate %>','<%'= enddate %>')", 4000);
		gomaechul('<%= startdate %>','<%= enddate %>');

		gosales('<%= startdate %>','<%= enddate %>');
	});

	//∏Ò«• ª—∏≤
	function gopurpose(yyyy, mm) {
		document.frm.calyyyy.value=yyyy;
		document.frm.calmm.value=mm;

		$.ajax({
			type: "get",
			url: "/admin/dataanalysis/md_purpose_ajax.asp?calyyyy="+yyyy+"&calmm="+mm,
			cache: false,
			success: function(str){
				$("#divpurpose").empty().html(str);
				$('#divpurpose').show();
			}
			,error: function(err) {
				alert(err.responseText);
			}
		});
	}

	//∏≈√‚ ª—∏≤
	function gomaechul(startdate, enddate) {
		$.ajax({
			type: "get",
			url: "/admin/dataanalysis/md_maechul_ajax.asp?startdate="+startdate+"&enddate="+enddate,
			cache: false,
			success: function(str){
				$("#divmaechul").empty().html(str);
				$('#divmaechul').show();
			}
			,error: function(err) {
				alert(err.responseText);
			}
		});
	}

	//øµæ˜¿ÃΩ¥ ª—∏≤
	function gosales(startdate, enddate) {
		$.ajax({
			type: "get",
			url: "/admin/dataanalysis/md_sales_ajax.asp?startdate="+startdate+"&enddate="+enddate,
			cache: false,
			success: function(str){
				$("#divsales").empty().html(str);
				$('#divsales').show();
			}
			,error: function(err) {
				alert(err.responseText);
			}
		});
	}

</script>

<br>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr bgcolor="#FFFFFF">
	<td width="38%" valign="top">
		<input type='hidden' name='calyyyy' value='<%= calyyyy %>'>
		<input type='hidden' name='calmm' value='<%= calmm %>'>
		<div id="divpurpose"><img src='http://fiximage.10x10.co.kr/icons/loading16.gif' width=20 height=20></div>
	</td>
	<td width="1%" valign="top"></td>
	<td width="38%" valign="top">
		<div id="divmaechul"><img src='http://fiximage.10x10.co.kr/icons/loading16.gif' width=20 height=20></div>
	</td>
	<td width="1%" valign="top"></td>
	<td valign="top">
		<div id="divsales"><img src='http://fiximage.10x10.co.kr/icons/loading16.gif' width=20 height=20></div>
	</td>
</tr>
</table>
