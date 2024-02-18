<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/que/queItemCls.asp"-->
<%
Dim mallid, oOutmall, page, i
Dim itemid, apiAction, resultCode, lastUserid, sellyn, pagesize, errMsg
mallid		= request("mallid")
itemid		= request("itemid")
apiAction	= request("apiAction")
resultCode	= request("resultCode")
page 		= request("page")
lastUserid	= request("lastUserid")
sellyn		= request("sellyn")
pagesize	= request("pagesize")
errMsg		= request("errMsg")

If page = "" Then page = 1
If pagesize = "" Then pagesize = 100

If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If

Set oOutmall = new COutmall
	oOutmall.FPageSize 			= pagesize
	oOutmall.FCurrPage			= page
	oOutmall.FRectMallid 		= mallid
	oOutmall.FRectItemid 		= itemid
	oOutmall.FRectApiAction 	= apiAction
	oOutmall.FRectResultCode 	= resultCode
	oOutmall.FRectLastUserid 	= lastUserid
	oOutmall.FRectGSShopSellyn 	= sellyn
	oOutmall.FRectErrMsg	 	= errMsg
	oOutmall.getQueLogbackupList
%>
<script>
//ũ�� ������Ʈ�� alert ����..2021-07-26
function systemAlert(message){
	alert(message);
}
window.addEventListener("message", (event) => {
    var data = event.data;
    if (typeof(window[data.action]) == "function") {
        window[data.action].call(null, data.message);
    } },
false);
//ũ�� ������Ʈ�� alert ����..2021-07-26 ��

function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

// ���õ� ��ǰ ����
function etcmallDeleteProcess(imallid) {
	var chkSel=0, strSell;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

	if (imallid == '11st1010'){
		if (confirm('API�� �����ϴ� ����� �ƴմϴ�.\n\n11���� ���ο��� ������ ó���ؾ� �մϴ�.\n\n ' + chkSel + '�� ���� �Ͻðڽ��ϱ�?')){
			if (confirm('���� �����Ͻðڽ��ϱ�? Ȯ�ι�ư Ŭ���� DB���� ��ǰ�� �����˴ϴ�.')){
				document.frmSvArr.target = "xLink";
				document.frmSvArr.cmdparam.value = "DELETE";
				document.frmSvArr.action = "<%=apiURL%>/outmall/11st/act11stReq.asp"
				document.frmSvArr.submit();
			}
		}
	} else if (imallid == 'interpark'){
		if (confirm('API�� �����ϴ� ����� �ƴմϴ�.\n\������ũ ���ο��� ������ ó���ؾ� �մϴ�.\n\n ' + chkSel + '�� ���� �Ͻðڽ��ϱ�?')){
			if (confirm('���� �����Ͻðڽ��ϱ�? Ȯ�ι�ư Ŭ���� DB���� ��ǰ�� �����˴ϴ�.')){
				document.frmSvArr.target = "xLink";
				document.frmSvArr.cmdparam.value = "DELETE";
				document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actinterparkReq.asp"
				document.frmSvArr.submit();
			}
		}
	} else if (imallid == 'gmarket1010'){
		if (confirm('API�� �����ϴ� ����� �ƴմϴ�.\n\������ ���ο��� ������ ó���ؾ� �մϴ�.\n\n ' + chkSel + '�� ���� �Ͻðڽ��ϱ�?')){
			if (confirm('���� �����Ͻðڽ��ϱ�? Ȯ�ι�ư Ŭ���� DB���� ��ǰ�� �����˴ϴ�.')){
				document.frmSvArr.target = "xLink";
				document.frmSvArr.cmdparam.value = "DELETE";
				document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
				document.frmSvArr.submit();
			}
		}
	}
}
// ���õ� ��ǰ �Ǹſ��� ����
function etcmallSellYnProcess(chkYn, imallid) {
	var chkSel=0, strSell;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

	switch(chkYn) {
		case "Y": strSell="�Ǹ���";break;
		case "N": strSell="ǰ��";break;
	}

	if (imallid == 'gsshop'){
	    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��GSShop���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
	        if (chkYn=="X"){
	            if (!confirm(strSell + '�� �����ϸ� GSShop���� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
	        }
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshop/actgsshopReq.asp"
	        document.frmSvArr.submit();
	    }
	}else if (imallid == 'lotteimall'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n�طԵ�iMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '�� �����ϸ� �Ե�iMall���� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'lotteCom'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '�� �����ϸ� n�طԵ����Ŀ��� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/LotteCom/actLotteComReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'cjmall'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��CJMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '�� �����ϸ� n��CJMall���� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/cjmall/actCjMallReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'auction1010'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��Auction���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '�� �����ϸ� n��Auction���� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'gmarket1010'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��Gmarket���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '�� �����ϸ� n��Gmarket���� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'homeplus'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��Ȩ�÷����� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '�� �����ϸ� n��Ȩ�÷������� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'interpark'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��������ũ�� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '�� �����ϸ� n��������ũ���� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actinterparkReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'nvstorefarm'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n�ؽ�����ʰ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '�� �����ϸ� n�ؽ�����ʿ��� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarm/actnvstorefarmReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'nvstoremoonbangu'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n�ؽ������ ���汸���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '�� �����ϸ� n�ؽ������ ���汸���� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/nvstoremoonbangu/actnvstoremoonbanguReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'Mylittlewhoopee'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n�ؽ�����ʰ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '�� �����ϸ� n�ؽ�����ʿ��� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/Mylittlewhoopee/actMylittlewhoopeeReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == '11st1010'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��11st���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/11st/act11stReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'ssg'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��ssg���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'coupang'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��coupang���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/coupang/actCoupangReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'halfclub'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��halfclub���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/halfclub/acthalfclubReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'hmall1010'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��hmall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'ezwel'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��ezwel���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
	        document.frmSvArr.submit();
		}
	}else if (imallid == 'lotteon'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��lotteon���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteon/actlotteonReq.asp"
	        document.frmSvArr.submit();
		}
	}else if (imallid == 'lfmall'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��lfmall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/lfmall/actlfmallReq.asp"
	        document.frmSvArr.submit();
		}
	}else if (imallid == 'WMP'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n������������ ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/wmp/actWmpReq.asp"
	        document.frmSvArr.submit();
		}
	}else if (imallid == 'skstoa'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��skstoa���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/skstoa/actskstoaReq.asp"
	        document.frmSvArr.submit();
		}
	}else if (imallid == 'shintvshopping'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��shintvshopping���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/shintvshopping/actShintvshoppingReq.asp"
	        document.frmSvArr.submit();
		}
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td>������ : <%= mallid %></td>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="mallid" value="<%=mallid%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>

		API�׼� :
	<% If mallid = "auction1010" Then %>
		<select name="apiAction" class="select">
			<option value="">��ü</option>
			<option value="REG"		 <%= Chkiif(apiAction = "REG", "selected", "")%> >��ǰ���</option>
			<option value="REGOnSale"	 <%= Chkiif(apiAction = "REGOnSale", "selected", "")%> >��Ͽ����Ǹ�������</option>
			<option value="EDIT"	 <%= Chkiif(apiAction = "EDIT", "selected", "")%> >��ǰ����</option>
			<option value="EDIT2"	 <%= Chkiif(apiAction = "EDIT2", "selected", "")%> >��ǰ����(�Ǹ���ȯ)</option>
			<option value="EditInfo"	 <%= Chkiif(apiAction = "EditInfo", "selected", "")%> >��ǰ����(��ǰ)</option>
			<option value="SOLDOUT"  <%= Chkiif(apiAction = "SOLDOUT", "selected", "")%> >ǰ��ó��</option>
			<option value="PRICE"	 <%= Chkiif(apiAction = "PRICE", "selected", "")%> >���ݼ���</option>
			<option value="KEEPSELL" <%= Chkiif(apiAction = "KEEPSELL", "selected", "")%> >�Ǹ�����</option>
		</select>
	<% ElseIf mallid = "gmarket1010" Then %>
		<select name="apiAction" class="select">
			<option value="">��ü</option>
			<option value="REG"		 <%= Chkiif(apiAction = "REG", "selected", "")%> >��ǰ���</option>
			<option value="REGOnSale"	 <%= Chkiif(apiAction = "REGOnSale", "selected", "")%> >��Ͽ����Ǹ�������</option>
			<option value="EDIT"	 <%= Chkiif(apiAction = "EDIT", "selected", "")%> >��ǰ����</option>
			<option value="EDIT2"	 <%= Chkiif(apiAction = "EDIT2", "selected", "")%> >��ǰ����(�Ǹ���ȯ)</option>
			<option value="EDITPOLICY"	 <%= Chkiif(apiAction = "EDITPOLICY", "selected", "")%> >��ǰ����(��ǰ)</option>
			<option value="EDITIMG"	 <%= Chkiif(apiAction = "EDITIMG", "selected", "")%> >�̹�������</option>
			<option value="SOLDOUT"  <%= Chkiif(apiAction = "SOLDOUT", "selected", "")%> >ǰ��ó��</option>
			<option value="PRICE"	 <%= Chkiif(apiAction = "PRICE", "selected", "")%> >���ݼ���</option>
			<option value="KEEPSELL" <%= Chkiif(apiAction = "KEEPSELL", "selected", "")%> >�Ǹ�����</option>
		</select>
	<% ElseIf mallid = "ssg" Then %>
		<select name="apiAction" class="select">
			<option value="">��ü</option>
			<option value="REG"		 <%= Chkiif(apiAction = "REG", "selected", "")%> >��ǰ���</option>
			<option value="EDIT"	 <%= Chkiif(apiAction = "EDIT", "selected", "")%> >��ǰ����</option>
			<option value="EDIT2"	 <%= Chkiif(apiAction = "EDIT2", "selected", "")%> >��ǰ����(�Ǹ���ȯ)</option>
			<option value="SOLDOUT"  <%= Chkiif(apiAction = "SOLDOUT", "selected", "")%> >ǰ��ó��</option>
			<option value="PRICE"	 <%= Chkiif(apiAction = "PRICE", "selected", "")%> >���ݼ���</option>
			<option value="CHKSTAT" <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >�űԻ�ǰ��ȸ</option>
		</select>
	<% ElseIf mallid = "hmall1010" Then %>
		<select name="apiAction" class="select">
			<option value="">��ü</option>
			<option value="REG"		 <%= Chkiif(apiAction = "REG", "selected", "")%> >��ǰ���</option>
			<option value="EDIT"	 <%= Chkiif(apiAction = "EDIT", "selected", "")%> >��ǰ����</option>
			<option value="EDIT2"	 <%= Chkiif(apiAction = "EDIT2", "selected", "")%> >��ǰ����(�Ǹ���ȯ)</option>
			<option value="EDITBATCH"	 <%= Chkiif(apiAction = "EDITBATCH", "selected", "")%> >��ǰ����(��ġó��)</option>
			<option value="OPTEDIT" <%= Chkiif(apiAction = "OPTEDIT", "selected", "")%> >�ɼǼ���</option>
			<option value="SOLDOUT"  <%= Chkiif(apiAction = "SOLDOUT", "selected", "")%> >ǰ��ó��</option>
			<option value="PRICE"	 <%= Chkiif(apiAction = "PRICE", "selected", "")%> >���ݼ���</option>
			<option value="OPTSTAT" <%= Chkiif(apiAction = "OPTSTAT", "selected", "")%> >�����ȸ</option>
			<option value="CHKSTAT" <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >����ȸ</option>
		</select>
	<% ElseIf mallid = "lotteon" Then %>
		<select name="apiAction" class="select">
			<option value="">��ü</option>
			<option value="REG"		 <%= Chkiif(apiAction = "REG", "selected", "")%> >��ǰ���</option>
			<option value="EDIT"	 <%= Chkiif(apiAction = "EDIT", "selected", "")%> >��ǰ����</option>
			<option value="EDIT2"	 <%= Chkiif(apiAction = "EDIT2", "selected", "")%> >��ǰ����(�Ǹ���ȯ)</option>
			<option value="EDITBATCH"	 <%= Chkiif(apiAction = "EDITBATCH", "selected", "")%> >��ǰ����(��ġó��)</option>
			<option value="QTY" <%= Chkiif(apiAction = "QTY", "selected", "")%> >������</option>
			<option value="SOLDOUT"  <%= Chkiif(apiAction = "SOLDOUT", "selected", "")%> >ǰ��ó��</option>
			<option value="PRICE"	 <%= Chkiif(apiAction = "PRICE", "selected", "")%> >���ݼ���</option>
			<option value="OPTSTAT" <%= Chkiif(apiAction = "OPTSTAT", "selected", "")%> >�ɼǻ���</option>
			<option value="CHKSTAT" <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >����ȸ</option>
		</select>
	<% Else %>
		<select name="apiAction" class="select">
			<option value="">��ü</option>
			<option value="REG"  <%= Chkiif(apiAction = "REG", "selected", "")%> >��ǰ���</option>
			<% If mallid = "shintvshopping" or mallid = "skstoa" Then %>
			<option value="CONFIRM" <%= Chkiif(apiAction = "CONFIRM", "selected", "")%> >��Ͻ��ο�û</option>
			<% End If %>
			<option value="EDIT"	 <%= Chkiif(apiAction = "EDIT", "selected", "")%> >��ǰ����</option>
			<option value="EDIT2"	 <%= Chkiif(apiAction = "EDIT2", "selected", "")%> >��ǰ����(�Ǹ���ȯ)</option>
			<option value="EDITBATCH"	 <%= Chkiif(apiAction = "EDITBATCH", "selected", "")%> >��ǰ����(��ġó��)</option>
			<option value="SOLDOUT"  <%= Chkiif(apiAction = "SOLDOUT", "selected", "")%> >ǰ��ó��</option>
			<option value="PRICE"	 <%= Chkiif(apiAction = "PRICE", "selected", "")%> >���ݼ���</option>
			<% If (mallid <> "cjmall" and mallid <> "11stmy") Then %>
			<option value="ITEMNAME" <%= Chkiif(apiAction = "ITEMNAME", "selected", "")%> >��ǰ�����</option>
			<% End If %>
			<% If mallid = "gsshop" Then %>
			<option value="IMAGE"	 <%= Chkiif(apiAction = "IMAGE", "selected", "")%> >�̹�������</option>
			<option value="CONTENT"  <%= Chkiif(apiAction = "CONTENT", "selected", "")%> >��ǰ�������</option>
			<option value="INFODIV"	 <%= Chkiif(apiAction = "INFODIV", "selected", "")%> >���ΰ�ü���</option>
			<option value="CHKSTAT"  <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >�űԻ�ǰ��ȸ</option>
			<% ElseIf mallid = "11stmy" Then %>
			<option value="VIEWOPT"  <%= Chkiif(apiAction = "VIEWOPT", "selected", "")%> >�ɼ���ȸ</option>
			<% ElseIf mallid = "cjmall" or mallid="coupang" or mallid="ezwel" or mallid="interpark" or mallid="lfmall" Then %>
			<option value="CHKSTAT"  <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >�űԻ�ǰ��ȸ</option>
			<% ElseIf (mallid = "lotteimall") OR (mallid = "lotteCom") Then %>
			<option value="CHKSTOCK"  <%= Chkiif(apiAction = "CHKSTOCK", "selected", "")%> >�����ȸ</option>
			<option value="CHKSTAT"  <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >�űԻ�ǰ��ȸ</option>
				<% If mallid = "lotteimall" Then %>
					<option value="DISPVIEW"  <%= Chkiif(apiAction = "DISPVIEW", "selected", "")%> >���û�ǰ��ȸ</option>
				<% Else %>
					<option value="INFODIV"	 <%= Chkiif(apiAction = "INFODIV", "selected", "")%> >���ΰ�ü���</option>
				<% End If %>
			<% End If %>
			<% If mallid = "interpark" Then %>
			<option value="DELETE"  <%= Chkiif(apiAction = "DELETE", "selected", "")%> >��ǰ����</option>
			<% End If %>
		</select>
	<% End If %>
		&nbsp;
		�������� :
		<select name="resultCode" class="select">
			<option value="">��ü</option>
			<option value="OK"  	<%= Chkiif(resultCode = "OK", "selected", "")%> >����</option>
			<option value="ERR"		<%= Chkiif(resultCode = "ERR", "selected", "")%> >����</option>
			<option value="QNull"	<%= Chkiif(resultCode = "QNull", "selected", "")%> >����</option>
		</select>
		&nbsp;
		�����Ǹ� :
		<select name="sellyn" class="select">
			<option value="">��ü</option>
			<option value="Y"  	<%= Chkiif(sellyn = "Y", "selected", "")%> >Y</option>
			<option value="N"  	<%= Chkiif(sellyn = "N", "selected", "")%> >N</option>
		</select>
		&nbsp;
		����ID :
		<select name="lastUserid" class="select">
			<option value="">��ü</option>
			<option value="system"  <%= Chkiif(lastUserid = "system", "selected", "")%> >������</option>
			<option value="etc"	<%= Chkiif(lastUserid = "etc", "selected", "")%> >������</option>
		</select>
		&nbsp;
		ǥ�ð��� :
		<select name="pagesize" class="select">
			<option value="20"  <%= Chkiif(pagesize = "20", "selected", "")%> >20</option>
			<option value="100"  <%= Chkiif(pagesize = "100", "selected", "")%> >100</option>
			<option value="165"  <%= Chkiif(pagesize = "165", "selected", "")%> >165</option>
			<option value="200"  <%= Chkiif(pagesize = "200", "selected", "")%> >200</option>
			<option value="500"  <%= Chkiif(pagesize = "500", "selected", "")%> >500</option>
		</select>
		<br /><br />
		Message �˻� : <input type="text" name="errMsg" id="errMsg" value="<%= errMsg %>">
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<p>
<textarea id="itemidArr"></textarea>
<button onclick="copyId();">Copy</button>
<script>
function copyId() {
	var ttt = document.getElementById("itemidArr");
	ttt.select();
	document.execCommand("copy");
}
</script>

<br />
<input type="button" value="���ư���" class="button" onclick="location.replace('/admin/etc/que/popQueLogList.asp?mallid=<%= mallid %>');">
<br />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="subcmd" value="">
<input type="hidden" name="chgSellYn" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="13">
		�˻���� : <b><%= FormatNumber(oOutmall.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oOutmall.FTotalPage,0) %></b>
	</td>
	<td align="right" valign="top">
		���û�ǰ�� ǰ����
		<input class="button" type="button" id="btnSellYn" value="����" onClick="etcmallSellYnProcess('N', '<%= mallid %>');">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td>������</td>
	<td>API�׼�</td>
	<td>��ǰ�ڵ�</td>
	<td>�켱����</td>
	<td>��Ͻð�</td>
	<td>ť�����ð�</td>
	<td>API�Ϸ�ð�</td>
	<td>�����Ǹ�</td>
	<td>�ɼǼ�</td>
	<td>���м�</td>
	<td>��������</td>
	<td>����ID</td>
	<td width="300">Message</td>
</tr>
<%
	Dim itemidArr, outmallGoodnoArr
%>
<% For i = 0 To oOutmall.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oOutmall.FItemList(i).FItemid %>"></td>
	<td><%= oOutmall.FItemlist(i).FMallid %></td>
	<td><%= oOutmall.FItemlist(i).FApiAction %></td>
	<td><%= oOutmall.FItemlist(i).FItemid %></td>
	<td><%= oOutmall.FItemlist(i).FPriority %></td>
	<td><%= oOutmall.FItemlist(i).FRegdate %></td>
	<td><%= oOutmall.FItemlist(i).FReaddate %></td>
	<td><%= oOutmall.FItemlist(i).FFindate %></td>
	<td>
		<%
			If oOutmall.FItemlist(i).FGSShopSellyn = "Y" Then
				response.write "<font color='BLUE'>"&oOutmall.FItemlist(i).FGSShopSellyn&"</font>"
			Else
				response.write "<font color='RED'>"&oOutmall.FItemlist(i).FGSShopSellyn&"</font>"
			End If
		%>
	</td>
	<td><%= oOutmall.FItemlist(i).FOptioncnt %>:<%= oOutmall.FItemlist(i).FRegedOptCnt %></td>
	<td><%= oOutmall.FItemlist(i).FAccFailCnt %></td>
	<td>
	<%
		Select Case oOutmall.FItemlist(i).FResultCode
			Case "OK"		response.write "<font color='BLUE'>"&oOutmall.FItemlist(i).FResultCode&"</font>"
			Case "ERR"		response.write "<font color='RED'>"&oOutmall.FItemlist(i).FResultCode&"</font>"
			Case Else		response.write "<font color='GRAY'>"&oOutmall.FItemlist(i).FResultCode&"</font>"
		End Select
	%>
	</td>
	<td><%= oOutmall.FItemlist(i).FLastUserid %></td>
	<td width="300"><font title='<%= oOutmall.FItemlist(i).FLastErrMsg %>'><%= left(oOutmall.FItemlist(i).FLastErrMsg, 120) %></font></td>
</tr>
<%
		itemidArr = itemidArr & oOutmall.FItemList(i).FItemid & ","
		If oOutmall.FItemList(i).FOutmallGoodno <> "" Then
			outmallGoodnoArr = outmallGoodnoArr & oOutmall.FItemList(i).FOutmallGoodno & VBCRLF
		End If
	Next

	If Right(itemidArr,1) = "," Then
		itemidArr = Left(itemidArr, Len(itemidArr) - 1)
	End If
%>

<% If mallid = "11st1010" or mallid = "interpark" or mallid = "gmarket1010" Then %>
<textarea id="outmallGoodnoArr"><%= outmallGoodnoArr %></textarea>
<button onclick="copyId2();">Copy</button>
<script>
function copyId2() {
	var ttt2 = document.getElementById("outmallGoodnoArr");
	ttt2.select();
	document.execCommand("copy");
}
</script>
<input class="button" type="button" id="DELETE" value="����" onClick="etcmallDeleteProcess('<%= mallid %>');">
<% End If %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="14" align="center">
	<% If oOutmall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oOutmall.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oOutmall.StartScrollPage To oOutmall.FScrollCount + oOutmall.StartScrollPage - 1 %>
		<% If i>oOutmall.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oOutmall.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
</form>
</table>
<% Set oOutmall = Nothing %>
<script>
	document.getElementById("itemidArr").value = "<%= itemidArr %>";
</script>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="400"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
