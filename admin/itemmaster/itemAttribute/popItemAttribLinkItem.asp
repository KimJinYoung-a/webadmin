<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<%
'###############################################
' Discription : ��ǰ �Ӽ� - ��ǰ ����
' History : 2019.04.23 ������ : �ű� ����
'###############################################

'// ���� ����
Dim attribCd, i
Dim oAttrib
Dim attribDiv,attribDivName,attribName,attribNameAdd,attribUsing
Dim makerid, itemname, includeOption, itemid, dispCate
dim iA ,arrTemp, arrItemid

'// �Ķ���� ����
attribCd = requestCheckVar(request("attribCd"),8)
includeOption = requestCheckVar(request("includeOption"),1)

itemid = requestCheckVar(request("itemid"),255)
dispCate = requestCheckVar(request("disp"),18)

if itemid<>"" then
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

'// �Ӽ����� ����
if attribCd<>"" then
	set oAttrib = new CAttrib
		oAttrib.FRectattribCd = attribCd
		oAttrib.GetOneAttrib
		if oAttrib.FResultCount>0 then
			attribDiv		= oAttrib.FOneItem.FattribDiv
			attribDivName	= oAttrib.FOneItem.FattribDivName
			attribName		= oAttrib.FOneItem.FattribName
			attribNameAdd	= oAttrib.FOneItem.FattribNameAdd
			attribUsing		= oAttrib.FOneItem.FattribUsing
		end if
	set oAttrib = Nothing
else
	Call Alert_Close("��ǰ�Ӽ������� �����ϴ�.")
	dbget.close: response.End
end if

if attribUsing="N" then
	Call Alert_Close("������ ��ǰ�Ӽ������Դϴ�.")
	dbget.close: response.End
end if
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css">
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<link href="/js/jqPagination/jqPagination.css" rel="stylesheet">
<script type="text/javascript" src="/js/jqPagination/jquery.jqpagination.min.js"></script>
<style type="text/css">
html {overflow-y:auto;}
#searchFilter label {white-space:nowrap;}
.dimmed {text-align: center; padding-top: 200px;}
li.selected { outline: 1px solid red; }
</style>
<script type="text/javascript">
$(function(){
	$(".pagination").hide();
	fnGetLinkedItemList(1);
});

// ���ͺ� ���� ����
function setSortable() {
	$( "#tblStandByItem, #tblLinkedItem").sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="24" colspan="2" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		receive: function(event,ui) {
			if(ui.sender.attr("id")=="tblStandByItem") {
				fnPostItemLink("addLinkItem",ui.item.attr("val"),"");
			} else {
				fnPostItemLink("clearLinkItem",ui.item.attr("val"),"");
			}
			chgNoResult();
		},
		connectWith: "#tblStandByItem, #tblLinkedItem",
		cancel: ".noResult",
		dropOnEmpty: true
    }).disableSelection();
}

// �⺻�� ���� ��ȯ
function chgNoResult() {
	if($("#tblStandByItem tr").length>1) {
		$("#tblStandByItem .noResult").hide();
	} else {
		$("#tblStandByItem .noResult").show();
	}
	if($("#tblLinkedItem tr").length>1) {
		$("#tblLinkedItem .noResult").hide();
	} else {
		$("#tblLinkedItem .noResult").show();
	}
}

// ������� ���� ��ǰ �˻� ��� ǥ��
function fnGetFindItemList(pg){
	$.ajax({
		url: "act_itemAttribItemList.asp",
		cache: true,
		type: "POST",
		data: $("#frmSearch").serialize() + "&mode=findItem&page="+pg,
		beforeSend: function() {
			$(".dimmed").show();
		},
		complete: function() {
			$(".dimmed").hide();
		},
		success: function(message) {
			if(message.response=="Ok") {
				var tblCont = '';
				if(message.items.length>0) {
					message.items.forEach(itm => {
						tblCont += '<tr val="'+itm.itemid+'">';
						tblCont += '<td>'+itm.itemid+'</td>';
						tblCont += '<td>'+itm.itemname+'</td>';
						tblCont += '</tr>';
					});
				} else {
					tblCont = '<tr class="noResult"><td colspan="3">�˻��� ��ǰ�� �����ϴ�.</td></tr>';
				}

				$("#tblStandByItem").empty().append(tblCont);
				setSortable();

				if(message.totalPage>1) {
					$("#lyrPgn1").show().jqPagination({
						current_page:pg,
						max_page:message.totalPage,
						paged: function(page) {
							fnGetFindItemList(page);
						}
					});
				} else {
					$("#lyrPgn1").hide();
				}
			} else {
				alert(message.faildesc);
				console.log(message);
			}
		}
		,error: function(err) {
			console.error(err.responseText);
		}
	});
}

// ����� ��ǰ ��� ǥ��
function fnGetLinkedItemList(pg){
	$.ajax({
		url: "act_itemAttribItemList.asp",
		cache: true,
		type: "POST",
		data: $("#frmSearch").serialize() + "&mode=linkedItem&page="+pg,
		beforeSend: function() {
			$(".dimmed").show();
		},
		complete: function() {
			$(".dimmed").hide();
		},
		success: function(message) {
			if(message.response=="Ok") {
				var tblCont = '';
				if(message.items.length>0) {
					message.items.forEach(itm => {
						tblCont += '<tr val="'+itm.itemid+'">';
						tblCont += '<td>'+itm.itemid+'</td>';
						tblCont += '<td>'+itm.itemname+'</td>';
						tblCont += '</tr>';
					});
				} else {
					tblCont = '<tr class="noResult"><td colspan="3">����� ��ǰ�� �����ϴ�.</td></tr>';
				}

				$("#tblLinkedItem").empty().append(tblCont);
				setSortable();

				if(message.totalPage>1) {
					$("#lyrPgn2").show().jqPagination({
						current_page:pg,
						max_page:message.totalPage,
						paged: function(page) {
							fnGetLinkedItemList(page);
						}
					});
				} else {
					$("#lyrPgn2").hide();
				}
			} else {
				alert(message.faildesc);
				console.log(message);
			}
		}
		,error: function(err) {
			console.error(err.responseText);
		}
	});
}

// ��ǰ ���� ó��
function fnPostItemLink(mode,iid,opt){
	$.ajax({
		url: "act_itemAttribLinkProc.asp",
		cache: true,
		type: "POST",
		data: "mode="+mode+"&attribCd=<%=attribCd%>&itemid="+iid+"&itemoption="+opt,
		success: function(message) {
			if(message.response!="Ok") {
				alert(message.faildesc);
				console.log(message);
			}
		}
		,error: function(err) {
			console.error(err.responseText);
		}
	});
}
</script>
<div class="pad20">
	<h3 class="bMar05">��ǰ�Ӽ� - ��ǰ ����</h3>

	<!-- ��ǰ�Ӽ� ���� -->
	<table class="tbType1 listTb">
	<colgroup>
		<col width="90" />
		<col width="*" />
		<col width="90" />
		<col width="*" />
	</colgroup>
	<tr>
		<th>�Ӽ��ڵ�</td>
		<td class="lt"><%=attribCd %></td>
		<th>�Ӽ�����</td>
		<td class="lt"><%=attribDivName%></td>
	</tr>
	<tr>
		<th>�Ӽ���</td>
		<td colspan="3" class="lt">
			<%=attribName & chkIIF(attribNameAdd="" or isNull(attribNameAdd),""," / " & attribNameAdd)%>
		</td>
	</tr>
	</table>

	<!-- ��ǰ �˻� ���� -->
	<form name="frmSearch" id="frmSearch" method="post" action="" style="margin:0px;">
	<input type="hidden" name="attribCd" value="<%=attribCd%>" />
	<table class="tbType1 listTb tMar10">
	<colgroup>
		<col width="70" />
		<col width="*" />
	</colgroup>
	<tr>
		<th>�˻�<br />����</th>
		<td id="searchFilter" class="lt">
			<table cellpadding="0" cellspacing="0">
			<tr>
				<td class="lt" style="border:none;">
					<label>�귣�� : <%	drawSelectBoxDesignerWithName "makerid", makerid %></label> &nbsp;/
					<label>��ǰ�� : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32" /></label>
					<!--<label><input type="checkbox" name="includeOption" value="Y" <%=chkIIF(includeOption="Y","checked","")%> /> �ɼ�����</label>-->
				</td>
				<td rowspan="2" class="lt" style="border:none;">
					<label>��ǰ�ڵ� : <textarea rows="3" cols="10" name="itemid" style="vertical-align:top;"><%=replace(itemid,",",chr(10))%></textarea></label>
				</td>
			</tr>
			<tr>
				<td class="lt" style="border:none;">
					<label>����ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox.asp"--></label>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	</table>
	</form>
	<script>
	function popSelectDispCate(trg,trgNm,disp){
		var param = "dispCate=" + disp + "&frmname=frmSearch&targetcompname=" + trg + "&targetcpndtlnm=" + trgNm
		var popwin = window.open('/common/module/popDispCateSelect.asp?'+param,'popSelectDispCategory','width=700,height=200,scrollbars=yes,resizable=yes');
		popwin.focus();
	}
	</script>
	<div class="tMar10">
		<table style="width:100%;">
		<colgroup>
			<col width="50%" />
			<col width="10" />
			<col width="50%" />
		</colgroup>
		<tr>
			<td class="bPad05">
				<b>���� ����</b>
				<input type="button" value="��ǰ�˻�" onClick="fnGetFindItemList(1)" class="ui-button" style="font-size:11px; float:right;">
			</td>
			<td>&nbsp;&nbsp;</td>
			<td class="bPad05">
				<b>����� ��ǰ</b>
				<input type="button" value="��ǰ�˻�" onClick="fnGetLinkedItemList(1)" class="ui-button" style="font-size:11px; float:right;">
			</td>
		</tr>
		<tr>
			<td class="ct">
				<!-- �˻� ��ǰ ��� -->
				<table  class="tbType1 listTb">
				<tr>
					<th>��ǰ�ڵ�</th>
					<th>��ǰ��</th>
				</tr>
				<tbody id="tblStandByItem">
				<tr class="noResult">
					<td colspan="2">��ǰ�� �˻����ּ���.</td>
				</tr>
				</tbody>
				</table>
				<div id="lyrPgn1" class="pagination">
					<a href="#" class="first" data-action="first">&laquo;</a>
					<a href="#" class="previous" data-action="previous">&lsaquo;</a>
					<input type="text" readonly="readonly" />
					<a href="#" class="next" data-action="next">&rsaquo;</a>
					<a href="#" class="last" data-action="last">&raquo;</a>
				</div>
			</td>
			<td></td>
			<td class="ct">
				<!-- ����� ��ǰ ��� -->
				<table class="tbType1 listTb">
				<tr>
					<th>��ǰ�ڵ�</th>
					<th>��ǰ��</th>
				</tr>
				<tbody id="tblLinkedItem">
				<tr>
					<td colspan="2">�ε���...</td>
				</tr>
				</tbody>
				</table>
				<div id="lyrPgn2" class="pagination">
					<a href="#" class="first" data-action="first">&laquo;</a>
					<a href="#" class="previous" data-action="previous">&lsaquo;</a>
					<input type="text" readonly="readonly" />
					<a href="#" class="next" data-action="next">&rsaquo;</a>
					<a href="#" class="last" data-action="last">&raquo;</a>
				</div>
			</td>
		</tr>
		</table>
	</div>
</div>
<div class="dimmed"><img src="/images/loading.gif" width="150px" /></div>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->