<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
dim idx, page, dispCate, itemname, i, stype, maxDepth, searchdiv, catesort, itemidarr
dim ebrand, sellyn, sailyn, limityn, couponyn, sortDiv
dispCate = requestCheckvar(request("disp"),10)
stype = requestCheckvar(request("stype"),2)
searchdiv = requestCheckvar(request("searchdiv"),10)
maxDepth = 2 '����ī�װ� 2depth���� �����ش�
itemidarr      = requestCheckvar(request("itemidarr"),255)

if idx="" Then idx = requestCheckvar(request("idx"),9)
page = requestCheckvar(request("page"),6)
itemname = requestCheckvar(request("itemname"),100)
if (page="") then page=1
ebrand = requestCheckvar(request("ebrand"),32)
sellyn      = requestCheckvar(request("sellyn"),10)
limityn     = requestCheckvar(request("limityn"),2)
sailyn      = requestCheckvar(request("sailyn"),1)
couponyn		= requestCheckvar(request("couponyn"),1)
sortDiv	= requestCheckvar(request("sortDiv"),5)
if sortDiv="" then sortDiv="new"

if itemname <> "" then
	if checkNotValidHTML(itemname) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if

if itemidarr<>"" then
	dim iA ,arrTemp,arrItemid
	itemidarr = replace(itemidarr,chr(13),"")
	arrTemp = Split(itemidarr,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp) 
		if trim(arrTemp(iA))<>"" then
			'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.05;������)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop

	itemidarr = left(arrItemid,len(arrItemid)-1)
end if

dim oitem, oitembanner

set oitem = new CItem

oitem.FPageSize         = 20
oitem.FCurrPage         = page
oitem.FRectDispCate   = dispCate
oitem.FRectItemid       = itemidarr
oitem.FRectMakerid	= ebrand
oitem.FRectSellYN       = sellyn
oitem.FRectLimityn      = limityn
oitem.FRectSailYn       = sailyn
oitem.FRectCouponYn		= couponyn
oitem.FRectSortDiv		= sortDiv
If searchdiv="itemid" Then
oitem.FRectItemid = itemname
Else
oitem.FRectItemName = itemname
End If
oitem.GetItemList

set oitembanner = new CDealSelect
oitembanner.FRectDealCode = idx
oitembanner.GetDealSelectItemList

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<title></title>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<script language="JavaScript" src="/js/common.js"></script>
<script type="text/javascript">
function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function selectCategory(frm){
	if ((frm.cdl.selectedIndex<0)||(frm.cdm.selectedIndex<0)||(frm.cds.selectedIndex<0)){
		alert('ī�װ��� ���ܰ� ��� �������ּ���.');
		return;
	}

	var cd1 = frm.cdl[frm.cdl.selectedIndex].value;
	var cd2 = frm.cdm[frm.cdm.selectedIndex].value;
	var cd3 = frm.cds[frm.cds.selectedIndex].value;

	var cd1name = frm.cdl[frm.cdl.selectedIndex].text;
	var cd2name = frm.cdm[frm.cdm.selectedIndex].text;
	var cd3name = frm.cds[frm.cds.selectedIndex].text;

	if ((cd1=="")||(cd2=="")||(cd3=="")){
		alert('ī�װ��� ���ܰ� ��� �������ּ���.');
		return;
	}

	opener.setCategory(cd1,cd2,cd3,cd1name,cd2name,cd3name);
	window.close();
}

//�˻�
function jsSearch(){   
	//��ǰ�ڵ� ����&���͸� �Է°����ϵ��� üũ-----------------------------
	var itemid = document.frm.itemid.value;  
	 itemid =  itemid.replace(",","\r");    //�޸��� �ٹٲ�ó�� 
		 for(i=0;i<itemid.length;i++){ 
			if ( itemid.charCodeAt(i) != "13" && itemid.charCodeAt(i) != "10" && "0123456789".indexOf(itemid.charAt(i)) < 0){ 
					alert("��ǰ�ڵ�� ���ڸ� �Է°����մϴ�.");
					return;
			}
		}  
	//---------------------------------------------------------------------
		document.frm.submit();
}

function tnCheckAll(bool, comp){
    var frm = comp.form;
    if (!comp.length){
        comp.checked = bool;
    }else{
        for (var i=0;i<comp.length;i++){
            comp[i].checked = bool;
        }
    }
}

function tnCheckOne(itemid, comp){
	if(comp.value==itemid){
		if(comp.checked){
			comp.checked = false;
		}else{
			comp.checked = true;
		}
	}else{
		for (var i=0;i<comp.length;i++){
			if(comp[i].value==itemid){
				if(comp[i].checked){
					comp[i].checked = false;
				}else{
					comp[i].checked = true;
				}

			}
		}
	}
}

function TnSelectItemReg(rfrm){
	//alert("ok");
	if($("#sortable input:checkbox[name='cksel']:checked").length<1){
		alert("���õ� ��ǰ�� �����ϴ�.");
	}else{
		$("#sortable input:checkbox[name='cksel']").attr('checked', true);
		rfrm.target="FrameCKP";
		rfrm.submit();
	}
}

function TnSelectItemDel(rfrm){
	//alert("ok");
	if($("#sortable input:checkbox[name='cksel']:checked").length<1){
		alert("���õ� ��ǰ�� �����ϴ�.");
	}else{
		rfrm.action="dodealitemdel.asp";
		rfrm.target="FrameCKP";
		rfrm.submit();
	}
}

function TnSaveDealItem(){
	if($("#sortable input:checkbox[name='cksel']:checked").length<1){
		alert("���õ� ��ǰ�� �����ϴ�.");
	}
	else{
		location.href="/admin/itemmaster/deal/selectdealitem.asp?idx=<%=idx%>";
	}
}

function TnDelThemeItemBanner(){
	opener.document.frmEvt.target='FrameCKP';
	opener.document.frmEvt.upback.value='Y';
	opener.document.frmEvt.submit();
	$("#sortable input:checkbox[name='cksel']").attr('checked', false);
	location.reload();
}

function TnViewSelectItemSave(rfrm){
	//alert("ok");
	if($("#sortable input:checkbox[name='cksel']:checked").length<1){
		alert("���õ� ��ǰ�� �����ϴ�.");
	}else{
		rfrm.submit();
	}
}

//�귣�� ID �˻� �˾�â
function jsSearchBrandIDNew(frmName,compName){
	var compVal = "";
	try{
		compVal = eval("document.all." + frmName + "." + compName).value;
	}catch(e){
		compVal = "";
	}

	var popwin = window.open("/admin/member/popBrandSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal + "&isjsdomain=o","popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

</script>
<body>
<!-- �˾� ������ : �ּ� 1100*751 -->
<div class="popWinV17">
	<h1>��ǰ �˻�</h1>
	
	<div class="popContainerV17 noScrl">
		<div class="contL">
			<h2>��ǰ ����</h2>
			<form name="frm" method="get">
			<input type="hidden" name="page" >
			<input type="hidden" name="itemid" >
			<input type='hidden' name='idx' id='idx'value="<%=idx%>">
			<input type='hidden' name='stype' id='stype'value="<%=stype%>">
			<div id="unitType01" class="unitPannel">
				<div class="searchWrap" style="border-top:none;">
					<div class="search">
						<ul>
							<li>
								<label class="formTit">ī�װ� :</label>
								<!-- #include virtual="/common/module/dispCateSelectBoxDepth2.asp"-->
								&nbsp;&nbsp;&nbsp;�귣�� : 
								<% 'NewDrawSelectBoxDesignerwithNameEvent "ebrand", ebrand %>
								<%	drawSelectBoxDesignerWithName "ebrand", ebrand %>
							</li>
						</ul>
					</div>
					<dfn class="line"></dfn>
					<div class="search">
						<ul>
							<li>
								<label class="formTit" for="schWord">�˻��� :</label>
								<select name="searchdiv" class="formSlt">
									<option value="itemname" selected>��ǰ��</option>
									<option value="itemid">��ǰ�ڵ�</option>
								</select>
								<input type="text" class="formTxt" name="itemname" id="schWord" style="width:200px" value="<%=itemname%>" placeholder="�˻�� �Է��ϼ���." />
							</li>
						</ul>
					</div>
					��ǰ�ڵ� : <textarea rows="3" cols="10" name="itemidarr" id="itemidarr"><%=replace(itemidarr,",",chr(10))%></textarea>
					<dfn class="line"></dfn>
					<div class="search">
						<ul>
							<li>
								<label class="formTit" for="schWord">�Ǹ� :</label>
								<% drawSelectBoxSellYN "sellyn", sellyn %>
								&nbsp;���� : <% drawSelectBoxLimitYN "limityn", limityn %>
								&nbsp;���� <% drawSelectBoxSailYN "sailyn", sailyn %>
								&nbsp;����: <% drawSelectBoxCouponYN "couponyn", couponyn %>
							</li>
						</ul>
					</div>
					<input type="button" class="schBtn" value="�˻�" onClick="jsSearch();" />
				</div>
				<div class="tbListWrap tMar15">
					<div class="ftLt pad10">
						<span>�˻���� : <strong><%= FormatNumber(oitem.FTotalCount,0) %></strong></span> <span class="lMar10">������ : <strong><%= FormatNumber(page,0) %> / <%= FormatNumber(oitem.FTotalPage,0) %></strong></span>
					</div>
					<div class="ftRt pad10">
						<label for="sortDiv">���ļ��� : </label>
						<select name="sortDiv" id="sortDiv" class="formSlt" title="���ļ���">
							<option value="new" <%=chkIIF(sortDiv="new","selected","")%>>�Ż��</option>
							<option value="best" <%=chkIIF(sortDiv="best","selected","")%>>�α��</option>
							<option value="cashH" <%=chkIIF(sortDiv="cashH","selected","")%>>����</option>
							<option value="cashL" <%=chkIIF(sortDiv="cashL","selected","")%>>������</option>
						</select>
					</div>
					<ul class="thDataList">
						<li>
							<p class="cell05"><input type="checkbox" onClick="tnCheckAll(this.checked,frm.cksel);" /></p>
							<p class="cell10">��ǰ ID</p>
							<p class="cell10">�̹���</p>
							<p>��ǰ��</p>
							<p class="cell10">����</p>
							<p class="cell10">��ü ID</p>
							<p class="cell10">�Ǹſ���</p>
						</li>
					</ul>
					<ul class="tbDataList" id="items">
						<% for i=0 to oitem.FresultCount-1 %>
						<input type="hidden" id="itemname<%=i%>" value="<%= oitem.FItemList(i).FItemName %>">
						<li>
							<p class="cell05"><input type="checkbox" name="cksel" id="cksel" value="<%= oitem.FItemList(i).FItemID %>" /></p>
							<p class="cell10" onClick="tnCheckOne(<%= oitem.FItemList(i).FItemID %>,frm.cksel);"><%= oitem.FItemList(i).FItemID %></p>
							<p class="cell10" onClick="tnCheckOne(<%= oitem.FItemList(i).FItemID %>,frm.cksel);"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0" /></p>
							<p class="lt" onClick="tnCheckOne(<%= oitem.FItemList(i).FItemID %>,frm.cksel);"><%= oitem.FItemList(i).FItemName %></p>
							<p class="cell10" onClick="tnCheckOne(<%= oitem.FItemList(i).FItemID %>,frm.cksel);">
								<% =FormatNumber(oitem.FItemList(i).Forgprice,0) %>
								<% if oitem.FItemList(i).Fsailyn="Y" then %>
									<br><font color=#F08050>(<% =CLng((oitem.FItemList(i).Forgprice-oitem.FItemList(i).Fsailprice)/oitem.FItemList(i).Forgprice*100) %>%��) <% =FormatNumber(oitem.FItemList(i).Fsailprice,0) %></font>
								<% end if %>
								<% if oitem.FItemList(i).FitemCouponYn="Y" then %>
									<br><font color=#5080F0>(��)<% =FormatNumber(oitem.FItemList(i).GetCouponAssignPrice,0) %></font>
								<% end if %>
							<%  %>
							</p>
							<p class="cell10" onClick="tnCheckOne(<%= oitem.FItemList(i).FItemID %>,frm.cksel);"><%= oitem.FItemList(i).FMakerID %></p>
							<p class="cell10" onClick="tnCheckOne(<%= oitem.FItemList(i).FItemID %>,frm.cksel);"><%IF oitem.FItemList(i).IsSoldOut() THEN%><font color="#ffa8a8">ǰ��</font><% Else %>�Ǹ�<% End If %></p>
						</li>
						<% next %>
					</ul>
					<div class="ct tPad20 cBk1">
						<% if oitem.HasPreScroll then %>
						<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>');">[prev]</a>
						<% else %>
						[pre]
						<% end if %>
						<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
							<% if i>oitem.FTotalpage then Exit for %>
							<% if CStr(page)=CStr(i) then %>
							<span class="cRd1">[<%= i %>]</span>
							<% else %>
							<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
							<% end if %>
						<% next %>
						<% if oitem.HasNextScroll then %>
						<a href="javascript:NextPage('<%= i %>');">[next]</a>
						<% else %>
						[next]
						<% end if %>
					</div>
				</div>
			</div>
			</form>
			<!--// ��ǰ Tab -->
		</div>
		<input type="button" class="btnMove" id="additem" title="�����ؼ� ���" />
		<form name="frmArrupdate" method="post" action="dodealitemreg.asp">
		<input type='hidden' name='item_temp'>
		<input type='hidden' name='mode'>
		<input type="hidden" name="check" id="check" value="<% If oitembanner.FresultCount > 0 Then Response.write oitembanner.FresultCount Else Response.write 0 %>">
		<input type='hidden' name='checkcnt' id='checkcnt'>
		<input type='hidden' name='idx' id='idx'value="<%=idx%>">
		<input type='hidden' name='stype' id='stype'value="<%=stype%>">
		<input type='hidden' name='upback' value="N">
		<div class="contR">
			<h2 style="margin-left:-1px;">���� ����</h2>
			<div class="tbListWrap">
				<ul class="thDataList">
					<li>
						<p class="cell10"><input type="checkbox" onClick="tnCheckAll(this.checked,frmArrupdate.cksel);" /></p>
						<p class="cell25">��ǰ�ڵ�</p>
						<p>��ǰ��</p>
						<p class="cell10">�Ǹ�</p>
					</li>
				</ul>
				<div id="sitem">
				<% If oitembanner.FresultCount > 0 Then %>
				<ul id="sortable" class="tbDataList">
				<% For i=0 To oitembanner.FresultCount-1 %>
					<li id='del<%= oitembanner.FItemList(i).Fitemid %>'>
						<p class='cell10'><input type='checkbox' name='cksel' id="cksel<%=i%>" value='<%= oitembanner.FItemList(i).Fitemid %>' /><input type='hidden' name='sitemname' id='sitemname' value='<%= oitembanner.FItemList(i).Fitemname %>'></p>
						<p class='cell25'><%= oitembanner.FItemList(i).Fitemid %></p>
						<p class='lt'><span class='textOverflow'><%= oitembanner.FItemList(i).Fitemname %></span></p>
						<p class='cell10'><span class='textOverflow'><% if oitembanner.FItemList(i).IsSoldOut then %>ǰ��<% else %>�Ǹ�<% end if %></span></p>
					</li>
				<% Next %>
				</ul>
				<% Else %>
				<ul id="sortable" class="tbDataList"></ul>
				<% End If %>
				</div>
				<div class="pad10 rt">
					<input type="button" class="btn" value="��������" onclick="TnViewSelectItemSave(this.form);" />
					<input type="button" class="btn" value="���û���" onclick="TnSelectItemDel(this.form)" />
				</div>
			</div>
		</div>
		</form>
	</div>
	<div class="popBtnWrap">
		<input type="button" value="���ÿϷ�" onclick="TnSaveDealItem();" class="cRd1" style="width:100px; height:30px;" />
		<input type="button" value="�ݱ�" onclick="self.close();" style="width:100px; height:30px;" />
	</div>
	
</div>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script>
$(function(){
	$("#additem").click(function(){
		
		var CheckOverlap=false;
		$("#items input:checkbox[name='cksel']").each(function(i){
			if($("#items input:checkbox[name='cksel']:eq(" + i + ")").is(":checked")){
				$("#sortable input:checkbox[name='cksel']").each(function(j){
					//alert($("#items input:checkbox[name='cksel']:eq(" + i + ")").val() + " / " + $("#sortable input:checkbox[name='cksel']:eq(" + j + ")").val());
					if($("#items input:checkbox[name='cksel']:eq(" + i + ")").val() == $("#sortable input:checkbox[name='cksel']:eq(" + j + ")").val()){
						CheckOverlap=true;
						$("#items input:checkbox[name='cksel']:eq(" + i + ")").attr('checked', false);
						return false;
					}
					else{
						CheckOverlap=false;
						return;
					}
				});
			}
		});
		//alert((Number($("#check").val())+Number($("#items input:checkbox[name='cksel']:checked").length)));
		if($("#items input:checkbox[name='cksel']:checked").length<1){
			alert("���õ� ��ǰ�� �����ϴ�.");
			return false;
		}else if(CheckOverlap){
			alert("���� ��ǰ�� �ֽ��ϴ�.");
			return false;
		//}else if((Number($("#check").val())+Number($("#items input:checkbox[name='cksel']:checked").length))>5){
		//	alert("��ǰ ��� �߰� ������ 5�� �Դϴ�.");
		//	return false;
		}else{
			// ���߰�
			var oRow;
			$("#items input:checkbox[name='cksel']").each(function(i){
				if($("#items input:checkbox[name='cksel']:eq(" + i + ")").is(":checked")){
					oRow = "					<li id='del" + $("#items input:checkbox[name='cksel']:eq(" + i + ")").val() + "'>"
					oRow += "						<p class='cell10'><input type='checkbox' name='cksel' id='cksel'" + i + " value='" + $("#items input:checkbox[name='cksel']:eq(" + i + ")").val() + "' /><input type='hidden' name='sitemname' id='sitemname'  value='" + $("#itemname"+i).val() + "'></p>"
					oRow += "						<p class='cell25'>" + $("#items input:checkbox[name='cksel']:eq(" + i + ")").val() + "</p>"
					oRow += "						<p class='lt'><span class='textOverflow'>" + string_cut($("#itemname"+i).val(), 20, "...") + "</span></p>"
					oRow += "					</li>"
					//alert(oRow);
					$("#sitem ul").append(oRow);
					$("#check").val(Number($("#check").val())+1);//�ɼ� �߰� ���� ī��Ʈ
				}
			});
		}
		tnCheckAll(true,frmArrupdate.cksel);
		document.frmArrupdate.upback.value="Y";
		TnSelectItemReg(frmArrupdate);
	});

	$("#delitem").click(function(){
		if($("#sortable input:checkbox[name='cksel']:checked").length<1){
			alert("���� ��ǰ�� �������ּ���.");
		}else{
			document.frmArrupdate.mode.value="del";
			document.frmArrupdate.upback.value="Y";
			TnSelectItemDel(frmArrupdate);
		}
	});
});
function string_cut(str, len, tail) {
	return str.substr(0, len)+tail;
}
</script>
<script>
$(function() {
	$("#sortable").sortable({
		placeholder:"handling"
	}).disableSelection();

	$(".tab li").click(function() {
		$(".tab li").removeClass('selected');
		$(this).addClass('selected');
		$('.unitPannel').hide();
		var activeTab = $(this).find("a").attr("href");
		$(activeTab).show();
		return false;
	});
});
chgDispCate('<%=dispCate%>','<%=maxDepth%>');
//document.onload = getOnload();
</script>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<%
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->