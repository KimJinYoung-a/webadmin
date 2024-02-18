<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/que/queItemCls.asp"-->
<%
Dim mallid, actCount
Dim queGroup, i, arrRows, apiaction

mallid		= request("mallid")
apiaction	= request("apiaction")
If mallid = "" Then
	response.write "좌측 제휴몰을 클릭하세요"
	response.end
End If

SET queGroup = new COutmall
	queGroup.FRectMallid = mallid
	queGroup.FRectApiAction = apiaction
	arrRows = queGroup.getMallActionList
	actCount = queGroup.FResultCount
SET queGroup = nothing
%>
<script>
//크롬 업데이트로 alert 수정..2021-07-26
function systemAlert(message){
	alert(message);
}
window.addEventListener("message", (event) => {
    var data = event.data;
    if (typeof(window[data.action]) == "function") {
        window[data.action].call(null, data.message);
    } },
false);
//크롬 업데이트로 alert 수정..2021-07-26 끝

function etcmallAction(v){
	var chkSel=0, strSell;
	var act;

	if (v == "SOLDOUT"){
		act = "EditSellYn";
		document.frmSvArr.chgSellYn.value = "N";
	}else if (v == "DELETE"){
		act = "EditSellYn";
		document.frmSvArr.chgSellYn.value = "X";
	}else if ((v == "EDIT2") || (v == "EDITBATCH")) {
		act = "EDIT";
	}else{
		act = v;
	}

	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

	if ("<%= mallid %>" == "auction1010"){
		if(v == "PRICE"){
			act = "EditInfo"
		}else if(v == "EDITBATCH"){
			act = "EditInfo"
		}
		document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp";
	}
	if ("<%= mallid %>" == "ezwel"){
		if(v == "PRICE"){
			act = "EDIT"
		}
		document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp";
	}
	if ("<%= mallid %>" == "gmarket1010"){
		if(v == "EDITBATCH"){
			act = "EDITPOLICY"
		}
		document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp";
	}
    if ("<%= mallid %>" == "halfclub")  	{document.frmSvArr.action = "<%=apiURL%>/outmall/halfclub/acthalfclubReq.asp";}
	if ("<%= mallid %>" == "11st1010")		{document.frmSvArr.action = "<%=apiURL%>/outmall/11st/act11stReq.asp";}
	if ("<%= mallid %>" == "gsshop")		{document.frmSvArr.action = "<%=apiURL%>/outmall/gsshop/actgsshopReq.asp";}
	if ("<%= mallid %>" == "interpark"){
		if(v == "PRICE"){
			act = "EDIT"
		}
		document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp";
	}
	if ("<%= mallid %>" == "nvstorefarm"){
		if(v == "PRICE"){
			act = "EDIT"
		}
		document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarm/actnvstorefarmReq.asp";
	}
	if ("<%= mallid %>" == "nvstoremoonbangu"){
		if(v == "PRICE"){
			act = "EDIT"
		}
		document.frmSvArr.action = "<%=apiURL%>/outmall/nvstoremoonbangu/actnvstoremoonbanguReq.asp";
	}
	if ("<%= mallid %>" == "Mylittlewhoopee"){
		if(v == "PRICE"){
			act = "EDIT"
		}
		document.frmSvArr.action = "<%=apiURL%>/outmall/Mylittlewhoopee/actMylittlewhoopeeReq.asp";
	}
	if ("<%= mallid %>" == "ssg"){
		if(v == "PRICE"){
			act = "EDIT"
		}
		document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp";
	}
	if ("<%= mallid %>" == "lfmall"){
		if(v == "PRICE"){
			act = "EDIT"
		}
		document.frmSvArr.action = "<%=apiURL%>/outmall/lfmall/actlfmallReq.asp";
	}

	if ("<%= mallid %>" == "kakaostore"){
		if(v == "PRICE"){
			act = "EDIT"
		}
		document.frmSvArr.action = "/admin/etc/kakaostore/actKakaostoreReq.asp";
	}

	if ("<%= mallid %>" == "lotteCom")			{document.frmSvArr.action = "<%=apiURL%>/outmall/LotteCom/actLotteComReq.asp";}
	if ("<%= mallid %>" == "lotteimall")		{document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp";}
	if ("<%= mallid %>" == "lotteon")			{document.frmSvArr.action = "<%=apiURL%>/outmall/lotteon/actlotteonReq.asp";}
	if ("<%= mallid %>" == "cjmall")			{document.frmSvArr.action = "<%=apiURL%>/outmall/cjmall/actCjMallReq.asp";}
	if ("<%= mallid %>" == "11stmy")			{document.frmSvArr.action = "<%=apiURL%>/outmall/11stmy/actmy11stReq.asp";}
	if ("<%= mallid %>" == "coupang")			{document.frmSvArr.action = "<%=apiURL%>/outmall/coupang/actcoupangReq.asp";}
	
	if ("<%= mallid %>" == "hmall1010")			{document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp";}
	if ("<%= mallid %>" == "WMP")				{document.frmSvArr.action = "<%=apiURL%>/outmall/wmp/actWmpReq.asp";}
	if ("<%= mallid %>" == "shintvshopping")	{document.frmSvArr.action = "<%=apiURL%>/outmall/shintvshopping/actShintvshoppingReq.asp";}
	if ("<%= mallid %>" == "skstoa")			{document.frmSvArr.action = "<%=apiURL%>/outmall/skstoa/actskstoaReq.asp";}

    if (confirm('<%= mallid %>에 선택하신 ' + chkSel + '개 상품을 수정하시겠습니까?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = act;
		document.frmSvArr.submit();
    }
}
</script>
<%= mallid %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="mallid" value="<%=mallid%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		액션 :
		<select name="apiaction" class="select">
			<option value="">-선택-</option>
			<option value="SOLDOUT" <%= chkiif(apiaction = "SOLDOUT", "selected", "") %> >SOLDOUT</option>
			<option value="PRICE"	<%= chkiif(apiaction = "PRICE", "selected", "") %> >PRICE</option>
			<option value="EDITBATCH"	<%= chkiif(apiaction = "EDITBATCH", "selected", "") %> >EDIT배치</option>
			<option value="EDIT"	<%= chkiif(apiaction = "EDIT", "selected", "") %> >EDIT</option>
			<% If mallid = "gsshop" Then %>
			<option value="EDITINFO"	<%= chkiif(apiaction = "EDITINFO", "selected", "") %> >EDITINFO(기본정보)</option>
			<% End If %>
			<option value="EDIT2"	<%= chkiif(apiaction = "EDIT2", "selected", "") %> >EDIT2(판매전환)</option>
			<option value="DELETE"	<%= chkiif(apiaction = "DELETE", "selected", "") %> >DELETE</option>
			<option value="CHKSTAT"	<%= chkiif(apiaction = "CHKSTAT", "selected", "") %> >CHKSTAT</option>
		</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br />
<% If apiaction <> "" Then %>
검색건수 : <%= actCount %>건
<input class="button" type="button" value="실행" onClick="etcmallAction('<%= apiaction %>');">
<br />
<% End If %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="chgSellYn" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="33%" align="center">상품코드</td>
	<td width="33%" align="center">액션</td>
	<td width="33%" align="center">카운트</td>
</tr>
<%
If isArray(arrRows) Then
	For i =0 To UBound(arrRows,2)
%>
<tr height="25" bgcolor="FFFFFF" >
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= arrRows(0, i) %>"></td>
	<td><%= arrRows(0, i) %></td>
	<td><%= arrRows(1, i) %></td>
	<td><%= arrRows(2, i) %></td>
</tr>
<%
	Next
Else
%>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4" height="50" align="center">데이터가 없습니다.</td>
</tr>
<% End If %>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="500"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->