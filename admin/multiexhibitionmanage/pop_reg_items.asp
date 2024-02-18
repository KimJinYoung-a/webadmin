<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2008.04.04 ������ ����
'           2010.07.05 ������ - �������� ���� �߰�
'						2013.12.24 ������ - ��ǰ�ڵ� �˻� �޸����ῡ�� ���ͷ� ����
' Description : ��ǰ �߰� - ����, ����ǰ ��ǰ��Ͽ� ���
'				input - actionURL(db ó���� �ʿ��� �Ķ���ͱ��� ����) ex.acURL = "/admin/eventmanage/event/eventitem_process.asp?eC=1234"
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/admin/multiexhibitionmanage/lib/classes/itemsCls.asp"-->
<%
dim target, actionURL
dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, deliverytype, limityn, vatyn, sailyn, couponyn, mwdiv,defaultmargin, keyword , sortDiv
dim cdl, cdm, cds , dispCate
dim reAct, ptype
dim page, paraRoad , sCode
dim mastercode , detailcode

actionURL 	= Replace(ReplaceRequestSpecialChar(request("acURL")),"||","&")

If (session("ssBctID")="areum531") Then				'2018-01-04 ���Ƹ� ��û, �˻� �����ۼ� ������û
	itemid      = requestCheckvar(request("itemid"),1255)
Else
	itemid      = requestCheckvar(request("itemid"),255)
End If

'itemid      = requestCheckvar(request("itemid"),255)
itemname    = requestCheckvar(request("itemname"),64)
makerid     = requestCheckvar(request("makerid"),32)
sellyn      = requestCheckvar(request("sellyn"),2)
usingyn     = requestCheckvar(request("usingyn"),1)
danjongyn   = requestCheckvar(request("danjongyn"),2)
limityn     = requestCheckvar(request("limityn"),2)
sailyn      = requestCheckvar(request("sailyn"),1)
deliverytype= requestCheckvar(request("deliverytype"),1)
mwdiv       = requestCheckvar(request("mwdiv"),2)
couponyn		= requestCheckvar(request("couponyn"),1)
defaultmargin = requestCheckvar(request("defaultmargin"),10)
keyword			= requestCheckvar(request("keyword"),512)
sortDiv			= requestCheckvar(request("sortDiv"),10)
paraRoad	= requestCheckvar(request("PR"),1)
sCode		= requestCheckvar(request("sC"),10)
reAct       = requestCheckvar(request("reAct"),1)
cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
dispCate = requestCheckvar(request("disp"),16)
ptype= requestCheckvar(request("ptype"),8)
page = requestCheckvar(request("page"),10)
mastercode = requestCheckvar(request("mastercode"),10)
detailcode = requestCheckvar(request("detailcode"),10)

if mastercode = "" then mastercode = 0
if detailcode = "" then detailcode = 0

if (page="") then page=1
if sailyn="" and instr(actionURL,"saleitem")>0 and reAct = "" then sailyn="N"			'�������������� �˻��ȰŶ�� �⺻��: ���ξ���(������ ����)
if couponyn="" and instr(actionURL,"saleitem")>0 and reAct = ""  then couponyn="N"
'if sellyn = "" then sellyn ="Y"
if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	itemid = replace(itemid,chr(13),"") '��ǰ�ڵ�˻� ���ͷ�(2013.12.24)
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)

		if trim(arrTemp(iA))<>"" then
			'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.04;������)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop

	if arrItemid <> "" then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		itemid = ""
	end if
end if


'==============================================================================
dim oitem

set oitem = new CItem

oitem.FPageSize         = 30
oitem.FCurrPage         = page
oitem.FRectMakerid      = makerid
oitem.FRectItemid       = itemid
oitem.FRectItemName     = itemname
oitem.FRectKeyword		= keyword

oitem.FRectSellYN       = sellyn
oitem.FRectIsUsing      = usingyn
oitem.FRectDanjongyn    = danjongyn
oitem.FRectLimityn      = limityn
oitem.FRectMWDiv        = mwdiv
oitem.FRectDeliveryType = deliverytype
oitem.FRectSailYn       = sailyn
oitem.FRectCouponYn		= couponyn

oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds
oitem.FRectDispCate		= dispCate
oitem.FRectSortDiv = SortDiv
If ptype="just1day" Then
oitem.FRectDealYn="N"
End If
oitem.GetItemListWithOption

dim i
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function jsSerach(){
	var frm;
	frm = document.frm;
	frm.target = "_self";
	frm.action ="pop_reg_items.asp";
	frm.submit();
}

function SelectItems(sType){
var frm;
var itemcount = 0;
frm = document.frmItem;
frm.sType.value = sType;   //��ü���� or ���û�ǰ ���� ����

	if (sType == "sel"){
		 if(typeof(frm.chkitem) !="undefined"){
	   	   	if(!frm.chkitem.length){
	   	   		if(!frm.chkitem.checked){
	   	   			alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
	   	   			return;
	   	   		}
	   	   		frm.itemidarr.value = frm.chkitem.value;
						frm.itemoptarr.value = frm.optioncode.value;
						frm.itemgubunarr.value = frm.gubuncode.value;
				
	   	   		itemcount = 1;
	   	    }else{
	   	    	for(i=0;i<frm.chkitem.length;i++){
	   	    		if(frm.chkitem[i].checked) {
	   	    			if (frm.itemidarr.value==""){
	   	    				frm.itemidarr.value =  frm.chkitem[i].value;
							frm.itemoptarr.value =  frm.optioncode[i].value;
							frm.itemgubunarr.value = frm.gubuncode[i].value;
	   	    			}else{
	   	    				frm.itemidarr.value = frm.itemidarr.value + "," +frm.chkitem[i].value;
							frm.itemoptarr.value = frm.itemoptarr.value + "," +frm.optioncode[i].value;
							frm.itemgubunarr.value = frm.itemgubunarr.value + "," +frm.gubuncode[i].value;

	   	    			}
	   	    		}
	   	    		itemcount = frm.chkitem.length;
	   	    	}

	   	    	if (frm.itemidarr.value == ""){
	   	    		alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
	   	   			return;
	   	    	}
	   	    }
	   	  }else{
	   	  	alert("�߰��� ��ǰ�� �����ϴ�.");
	   	  	return;
	   	  }
	}else{
		if(typeof(frm.chkitem) !="undefined"){
			itemcount = "<%= oitem.FTotalCount%>";
		  if(confirm("<%= oitem.FTotalCount%>���� �˻��� ��� ��ǰ�� �߰��Ͻðڽ��ϱ�?")){
		  	if(itemcount > 1000) {
		  		alert("��ǰ�� �ִ� 1000�Ǳ��� �����մϴ�. ������ �ٽ� �������ּ��� ");
		  		return;
		  	}
			frm.itemidarr.value = document.frm.itemid.value;
		  }else{
		  	return;
		  }
		}else{
		 	alert("�߰��� ��ǰ�� �����ϴ�.");
	   	  	return;
		}
	}

	// ��ȹ�� ����
	if (!document.frm.mastercode.value || document.frm.mastercode.value == 0) {
			alert("���� ������ ���ּ���");
			frm.mastercode.focus;
			return;
	} else {
		frm.mastercode.value = document.frm.mastercode.value;
	}

	if (!document.frm.detailcode.value) {
			alert("�ɼ� ������ ���ּ���");
			return;
	}

	frm.target = "FrameCKP";
	//frm.target = "blank";
	frm.action = "/admin/multiexhibitionmanage/lib/items_proc.asp";
	frm.itemcount.value = itemcount;
	frm.submit();
    frm.itemidarr.value = "";
	frm.itemoptarr.value = "";
	frm.itemcount.value = 0;
	opener.history.go(0);
	//window.close();
}

function SelectAllItems(){
var frm;
var itemcount = 0;
frm = document.frm;
		itemcount = "<%= oitem.FTotalCount%>";
		if (itemcount >0){
		  if(confirm("<%= oitem.FTotalCount%>���� �˻��� ��� ��ǰ�� �߰��Ͻðڽ��ϱ�?")){
		  	if(itemcount > 1000) {
		  		alert("��ǰ�� �ִ� 1000�Ǳ��� �����մϴ�. ������ �ٽ� �������ּ��� ");
		  		return;
		  	}
		  }else{
		  	return;
		  }
		}else{
		 	alert("�߰��� ��ǰ�� �����ϴ�.");
	   	  	return;
		}

	//frm.target = opener.name;
	frm.sType.value = "all";
	frm.target = "FrameCKP";
	frm.action = "<%=actionURL%>";
	frm.itemcount.value = itemcount;
	frm.submit();
	frm.itemidarr.value = "";
	frm.itemcount.value = 0;
	opener.history.go(0);
}

//��ü ����
function jsChkAll() {
	var frm;
	frm = document.frmItem;
	if (frm.chkAll.checked) {
	  if(typeof(frm.chkitem) !="undefined") {
	   	if (!frm.chkitem.length) {
		   	 	frm.chkitem.checked = true;
		  } else {
				for (i=0;i<frm.chkitem.length;i++) {
					frm.chkitem[i].checked = true;
			 	}
		  }
	  }
	} else {
		if(typeof(frm.chkitem) !="undefined") {
			if(!frm.chkitem.length) {
					frm.chkitem.checked = false;
			}else{
				for(i=0;i<frm.chkitem.length;i++) {
					frm.chkitem[i].checked = false;
				}
			}
		}
	}
}

// �����Ȳ �˾�
function PopItemStock(gubuncode,itemid,itemoption){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun="+ gubuncode +"&itemid="+ itemid +"&itemoption="+ itemoption,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// ������ �̵�
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="pop_reg_items.asp";
	document.frm.submit();
}

function mkbutton(mastercode) {
    var filtercode = 3;
    var targetform = "frmItem";
    var targetname = "detailcode";
    $.ajax({
        method : "get",
        url: "/admin/multiexhibitionmanage/lib/ajax_function.asp",
        data : "mastercode="+mastercode+"&filtercode="+filtercode+"&targetform="+targetform+"&targetname="+targetname,
        cache: false,
        async: false,
        success: function(message) {
            $("#submenu").empty().html(message).css("padding-top","10px");
        }
    });
}

$(function(){
    // init select
	<% if mastercode > 0 then %>
    mkbutton(<%=mastercode%>);
	<% end if %>
});
</script>
</head>
<body>
<div class="contSectFix scrl">
	<div class="pad20">
		<form name="frm" method="post">
		<input type="hidden" name="page" >
		<input type="hidden" name="sType" >
		<input type="hidden" name="itemidarr" >
		<input type="hidden" name="itemoptarr" >
		<input type="hidden" name="itemgubunarr" >
		<input type="hidden" name="itemcount" value="0">
		<input type="hidden" name="mode" value="I">
		<input type="hidden" name="acURL" value="<%=actionURL%>">
		<input type="hidden" name="defaultmargin" value="<%=defaultmargin%>">
		<input type="hidden" name="PR" value="<%=paraRoad%>">
		<input type="hidden" name="sC" value="<%=sCode%>">
		<input type="hidden" name="ptype" value="<%=ptype%>">
		<input type="hidden" name="reAct" value="1">
		<table class="tbType1 listTb">
			<tr bgcolor="<%= adminColor("topbar") %>" >
				<td  style="text-align:left;">
					<table class="tbType1 listTb">
						<tr>
							<td style="text-align:left;">�귣��: <%	drawSelectBoxDesignerWithName "makerid", makerid %></td>
							<td style="text-align:left;">��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20"></td>
							<td style="white-space:nowrap;padding-left:5px;">��ǰ�ڵ�:</td>
							<td style="white-space:nowrap;" rowspan="2"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea> </td>
						</tr>
					<tr>
						<td style="text-align:left;"> <!-- #include virtual="/common/module/categoryselectbox.asp"--></td>
						<td colspan="2" style="text-align:left;">����ī�װ� : <!-- #include virtual="/common/module/dispCateSelectBox.asp"--></td>
					</tr>
					<tr>
						<td colspan="4"  style="text-align:left;">�˻�Ű���� : <input type="text" class="text" name="keyword" value="<%=keyword%>" size="40"><font color="gray" size="2">(����:�������ֽ��ϴ�.)</font>
							<div style="float:right;text-align:left;padding:10px;">
								�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>

								���:<% drawSelectBoxUsingYN "usingyn", usingyn %>

								����:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>

								����:<% drawSelectBoxLimitYN "limityn", limityn %>

								���:<% drawSelectBoxMWU "mwdiv", mwdiv %><br><br>

								����: <% drawSelectBoxSailYN "sailyn", sailyn %>

								����: <% drawSelectBoxCouponYN "couponyn", couponyn %>

								���:<% drawBeadalDiv "deliverytype",deliverytype %>

								����Ʈ: <% drawSelectBoxIsBestSorting "sortDiv", sortDiv%>
							</div>
						</td>
					</tr>
				</table>
				</td>
				<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSerach();">
				</td>
			</tr>
			<tr bgcolor="<%= adminColor("topbar") %>" >
				<td>
					<div style="float:left;">
						<table cellpadding="3" cellspacing="1" class="a" border="0" width="100%">
							<tr align="center" bgcolor="<%= adminColor("topbar") %>">
								<td style="color:red;text-align:left;">�� ��ȹ���� ���� ���ּ���! (�ʼ�) ��</td>
							</tr>
							<tr align="center" bgcolor="<%= adminColor("topbar") %>">
								<td style="text-align:left;">���� ���� &nbsp;&nbsp;&nbsp;<%=DrawSelectAllView("mastercode",mastercode,"mkbutton")%>
								</td>
							</tr>
							<tr>
								<td>
									<div id="submenu" style="text-align:left;"></div>
								</td>
							</tr>
						</table>
					</div>
					
				</td>
			</tr>
		</table>
		</form>
		<div class="tPad15">
			<form name="frmItem" method="post">
			<input type="hidden" name="page" >
			<input type="hidden" name="sType" >
			<input type="hidden" name="itemidarr" >
			<input type="hidden" name="itemoptarr" >
			<input type="hidden" name="itemgubunarr" >
			<input type="hidden" name="itemcount" value="0">
			<input type="hidden" name="mode" value="I">
			<input type="hidden" name="acURL" value="<%=actionURL%>">
			<input type="hidden" name="defaultmargin" value="<%=defaultmargin%>">
			<input type="hidden" name="sC" value="<%=sCode%>">
			<input type="hidden" name="ptype" value="<%=ptype%>">
			<input type="hidden" name="mastercode" value="<%=mastercode%>">
			<input type="hidden" name="detailcode" value="">
			<table class="tbType1 listTb">
				<tr>
					<td  style="text-align:left;">
						<input type="button" value="���û�ǰ �߰�" onClick="SelectItems('sel')" class="button">
						<!-- saleItemProc.asp ��ü���� �߰� ���� �ִµ���.-->
						<!-- /admin/shopmaster/sale/saleItemProc_skyer9.asp �߰��۾� �ʿ� -->
						<%IF paraRoad ="S" THEN '���ΰ��������� ��ü���� ��ư Ȱ��ȭó�� 2014-12-02 ������%>
						<input type="button" value="��ü���� �߰�" onClick="SelectAllItems();" class="button" >
						<%END IF%>
						<!-- -->
					</td>
				</tr>
			</table>

			<table class="tbType1 listTb">
			<tr bgcolor="#FFFFFF">
				<td colspan="16" style="text-align:left;">
				�˻���� : <b><%= oitem.FTotalCount%></b>
				&nbsp;
				������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
				</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td align="center"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
				<td align="center">����</td>
				<td align="center">��ǰID</td>
				<td align="center">�ɼ��ڵ�</td>
				<td align="center">[�ɼ�Ÿ��]<br/><br/>�ɼǸ�</td>
				<td align="center">�̹���</td>
				<td align="center">�귣��</td>
				<td align="center">��ǰ��</td>
				<td align="center">�ǸŰ�</td>
				<td align="center">���԰�</td>
				<td align="center" nowrap>���<br>����</td>
				<td align="center" nowrap>���<br>����</td>
				<td align="center" nowrap>�Ǹ�<br>����</td>
				<td align="center" nowrap>���<br>����</td>
				<td align="center" nowrap>����<br>����</td>
				<td align="center" nowrap>���<br>��Ȳ</td>
			</tr>
			<% if oitem.FresultCount<1 then %>
				<tr bgcolor="#FFFFFF" >
					<td colspan="16" align="center">[�˻������ �����ϴ�.]</td>
				</tr>
			<% end if %>
			<% if oitem.FresultCount > 0 then %>
				<% for i=0 to oitem.FresultCount-1 %>
				<tr class="a" height="25" bgcolor="#FFFFFF">
				<td align="center">
					<input type="checkbox" name="chkitem" value="<%= oitem.FItemList(i).FItemId %>">
					<input type="hidden" name="optioncode" value="<%=chkiif(oitem.FItemList(i).Fitemoption="","0000",oitem.FItemList(i).Fitemoption)%>">
					<input type="hidden" name="gubuncode" value="<%= oitem.FItemList(i).Fitemgubun %>">
				</td>
				<td>&nbsp;<% =oitem.FItemList(i).Fitemgubun %></td>
				<td align="center"><A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).FItemId %>" target="_blank"><%= oitem.FItemList(i).FItemId %></a></td>
				<td>&nbsp;<% =oitem.FItemList(i).Fitemoption %></td>
				<td>&nbsp;[<% =oitem.FItemList(i).Foptiontypename %>]<br/><br/><% =oitem.FItemList(i).Fitemoptionname %></td>
				<td align="center"><%IF oitem.FItemList(i).FSmallImage <> "" THEN%><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border=0 alt=""><%END IF%></td>
					<td align="center"><% =oitem.FItemList(i).Fmakerid %></td>
				<td>&nbsp;<% =oitem.FItemList(i).Fitemname %></td>
				<td align="center"><%
						Response.Write FormatNumber(oitem.FItemList(i).Forgprice,0)
						'���ΰ�
						if oitem.FItemList(i).Fsailyn="Y" then
							Response.Write "<br><font color=#F08050>(��)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
						end if
						'������
						if oitem.FItemList(i).FitemCouponYn="Y" then
							Select Case oitem.FItemList(i).FitemCouponType
								Case "1"
									Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),0) & "</font>"
								Case "2"
									Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,0) & "</font>"
							end Select
						end if
					%></td>
				<td align="center"><%
						Response.Write FormatNumber(oitem.FItemList(i).Forgsuplycash,0)
						'���ΰ�
						if oitem.FItemList(i).Fsailyn="Y" then
							Response.Write "<br><font color=#F08050>" & FormatNumber(oitem.FItemList(i).Fsailsuplycash,0) & "</font>"
						end if
						'������
						if oitem.FItemList(i).FitemCouponYn="Y" then
							if oitem.FItemList(i).FitemCouponType="1" or oitem.FItemList(i).FitemCouponType="2" then
								if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
									Response.Write "<br><font color=#5080F0>" & FormatNumber(oitem.FItemList(i).Forgsuplycash,0) & "</font>"
								else
									Response.Write "<br><font color=#5080F0>" & FormatNumber(oitem.FItemList(i).Fcouponbuyprice,0) & "</font>"
								end if
							end if
						end if
					%></td>
				<td align="center"><%=fnColor(oitem.FItemList(i).IsUpcheBeasong(),"delivery")%></td>
				<td align="center"><%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %></td>
				<td align="center">
				<%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %>
				</td>
				<td align="center">
				<%= fnColor(oitem.FItemList(i).Fisusing,"yn") %>
				</td>
				<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
				<td align="center" nowrap>
				<a href="javascript:PopItemStock('<%= oitem.FItemList(i).Fitemgubun%>','<%= oitem.FItemList(i).FItemId %>','<%= oitem.FItemList(i).Fitemoption%>')" title="�����Ȳ �˾�">[����]</a><br>
				<%IF oitem.FItemList(i).IsSoldOut() THEN%>
					<img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
				<%END IF%>
				</td>
			</tr>
			<% next %>
			<tr>
				<td colspan="16" align="center" bgcolor="#FFFFFF">
				<!-- ����¡ó�� -->
					<% if oitem.HasPreScroll then %>
						<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
						<% else %>
							[pre]
						<% end if %>

						<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
							<% if i>oitem.FTotalpage then Exit for %>
							<% if CStr(page)=CStr(i) then %>
							<font color="red">[<%= i %>]</font>
							<% else %>
							<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
							<% end if %>
						<% next %>

						<% if oitem.HasNextScroll then %>
							<a href="javascript:NextPage('<%= i %>')">[next]</a>
						<% else %>
							[next]
						<% end if %>
				</td>
			</tr>
			<% end if %>
			</table>
			</form>
		</div>
		<div style="padding:5px;text-align:right;font-size:8pt">Ver1.0  lastupdate: 2013.12.24 </div>
		<iframe name="FrameCKP" src="about:blank" frameborder="0" width="600" height="200"></iframe>
	</div>
</div>

<%	set oitem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
