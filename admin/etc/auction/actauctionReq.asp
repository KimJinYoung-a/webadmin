<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),20)
Dim arrItemid : arrItemid = request("cksel")
Dim auto : auto = request("auto")
Dim CommCD : CommCD = request("CommCD")
Dim goodsGrpCd : goodsGrpCd = request("goodsGrpCd")
Dim i, strParam, iErrStr, ret1
Dim sqlStr, strSql, AssignedRow, SubNodes
Dim chgSellYn, actCnt, retErrStr
Dim buf, buf2, CNT10, CNT20, CNT30, iitemid
Dim ArrRows
Dim retFlag
Dim iMessage
dim iItemName, pregitemname
retFlag   = request("retFlag")
chgSellYn = request("chgSellYn")
arrItemid = Trim(arrItemid)
'rw arrItemid  : 1856104, 1856200
%>
<script type="text/javascript">
	var items = "<%=arrItemid%>";
	var itemArr = items.split(", ");
	var rotation;
	var rno = 0;

	function loadRotation() {
		if(itemArr[rno] == undefined){
			<% if (auto <> "Y") then %>
			alert('완료하였습니다');
			<% end if %>
			return;
		}
		rotation = arrSubmit(itemArr[rno]);
		rno++;
		if(rno > itemArr.length-1){
			clearTimeout(rotation);
			//setTimeout("alert('완료하였습니다')", 500);
		}else{
			//setTimeout('loadRotation()', 2000);
		}
	}

	function arrSubmit(ino){
		document.frmSvArr.target = "xLink2";
        document.frmSvArr.act.value = "<%=cmdparam%>";
        document.frmSvArr.itemid.value = ino;
        document.frmSvArr.chgSellYn.value = "<%=chgSellYn%>";
		document.frmSvArr.ccd.value = "<%=CommCD%>";
		document.frmSvArr.goodsGrpCd.value = "<%=goodsGrpCd%>";
        document.frmSvArr.action = '/admin/etc/auction/auctionActProc.asp';
		document.frmSvArr.submit();
	}
	window.onload = new Function('setTimeout("loadRotation()", 200)');
</script>
<form name="frmSvArr">
	<input type="hidden" name="act">
	<input type="hidden" name="itemid">
	<input type="hidden" name="chgSellYn">
	<input type="hidden" name="ccd">
	<input type="hidden" name="goodsGrpCd">
</form>

<div id="actStr"></div>
<iframe name="xLink2" id="xLink2" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->