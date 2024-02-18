<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��ȹ�� ����
' History : 2019-01-21 ����ȭ
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/multiexhibitionmanage/lib/classes/itemsCls.asp"-->
<%
dim limited , itemid , poscode
dim isusing , mdpick
dim mastercode , detailcode , detailcode2 , flaglist
dim menu : menu = "exhibitionitem"

dim sBrand,arrItemid
isusing = request("isusingbox")
sBrand 	= request("ebrand")
arrItemid = request("aitem")
mdpick 	= request("mdpick")
poscode = request("menupos")
itemid	= requestCheckvar(request("itemid"),255)
flaglist = requestCheckvar(request("flaglist"),1)

mastercode = requestCheckvar(request("mastercode"),10)
detailcode = requestCheckvar(request("detailcode"),100)

if flaglist = "" then flaglist = 1
if mastercode = "" then mastercode = 0
if detailcode = "" then 
	detailcode = ""
else
	detailcode = trim(detailcode)
end if

dim page , i
	page = requestCheckVar(request("page"),5)
	if page = "" then page = 1

if itemid<>"" then
	dim iA ,arrTemp',arrItemid
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

dim oExhibition
set oExhibition = new ExhibitionCls
	oExhibition.FPageSize = 50
	oExhibition.FCurrPage = page
	oExhibition.FrectMasterCode = mastercode
	oExhibition.FrectDetailCode = detailcode
	oExhibition.FrectMakerid = sBrand
	oExhibition.FRectArrItemid = arrItemid
	oExhibition.Frectpick = mdpick
	if flaglist = 1 then 
		oExhibition.getItemsList
	else
		oExhibition.getOptionItemsList
	end if 

%>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/admin/common/lib/js/front.js"></script>
<script type="text/javascript">
// �ű� ��� �˾� ���� ��ǰ
function popRegItems(){
	var popRegItems = window.open('/admin/multiexhibitionmanage/pop_reg_items.asp','popRegItems','width=1400,height=750,scrollbars=yes,resizable=yes')
	popRegItems.focus();
}

// �׷� ���� �˾�
function popGroupManage() {
	var popGroupManage = window.open('/admin/multiexhibitionmanage/pop_exhibition_manage.asp','popRegNew','width=750,height=750,status=yes')
	popGroupManage.focus();
}

// ��ǰ ����
function fnDelItem(idx) {
	if (confirm("��ǰ�� ���� �Ͻðڽ��ϱ�?") == true){ 
		var frm = document.itemdel
		frm.eidx.value = idx;
		frm.submit();
	}
}

// pick ��ü����
var ichk;
ichk = 1;
function jsChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if ((e.name=="chkI")){
			if ((e.type=="checkbox")) {
				e.checked = blnChk ;
			}
		}
	}
}

// pick �ϰ� ����
function jsSortIsusing() {
	if (confirm("pick��뿩�θ� ���� �Ͻðڽ��ϱ�?") == true){    //Ȯ��
		var frm;
		var sValue;
		var mdpick;
		frm = document.fitem;
		sValue = ""; //idx
		sCheck = ""; //mdpick 1,0
		chkSel	= 0;
	
		if (frm.chkI.length > 1){
			for (var i=0;i<frm.chkI.length;i++){
				if(frm.chkI[i].checked) chkSel++;
	
				if (frm.chkI[i].checked){
					if (sValue==""){
						sValue = frm.chkI[i].value;
					}else{
						sValue =sValue+","+frm.chkI[i].value;
					}
					
					frm.mdpickchk[i].value="1";
					if (sCheck==""){
						sCheck = frm.mdpickchk[i].value;
					}else{
						sCheck =sCheck+","+frm.mdpickchk[i].value;
					}
				}else{
					if (sValue==""){
						sValue = frm.chkI[i].value;
					}else{
						sValue =sValue+","+frm.chkI[i].value;
					}
					frm.mdpickchk[i].value="0";
					if (sCheck==""){
						sCheck = frm.mdpickchk[i].value;
					}else{
						sCheck =sCheck+","+frm.mdpickchk[i].value;
					}
				}
			}
		}else{
			if(frm.chkI.checked) chkSel++;
			if(frm.chkI.checked){
				sValue = frm.chkI.value;
				sCheck = frm.mdpickchk.value;
			}
		}
		document.frmreg.mdpick.value = sCheck;
		document.frmreg.eid.value = sValue;
		document.frmreg.submit();
	}else{
	    return;
	}
}

// �˻� ��ư �߰�
function mkbutton(mastercode) {
    var filtercode = 2;
    var targetform = "refreshFrm";
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

// �ɼ� ��ǰ ������
function optmore(itemid){
	if ($("#optlist"+itemid).css("display") != "none" ){
		$("#optlist"+itemid).hide();
		return false;
	}

	$.ajax({
        method : "get",
        url: "/admin/multiexhibitionmanage/ajax_optionitems.asp",
        data : "itemid="+itemid+"&detailcode=<%=detailcode%>",
        cache: false,
        async: false,
        success: function(message) {
			$("#optlist"+itemid).show();
			$("#optlist"+itemid+"_2").empty().html(message).show();
        }
    });
}

function FnIsUsing(idx,isusing) {
	var text;
	var bgcolor;
	var tempisusing;
	$.ajax({
        method : "post",
        url: "/admin/multiexhibitionmanage/lib/manage_proc.asp",
        data : "itemidx="+idx+"&itemisusing="+isusing+"&itemmode=itemisusing",
        cache: false,
        async: false,
        success: function(message) {
			if (isusing == 0) {
				text = "[�����]";
				bgcolor = "#EC3F1A";
				tempisusing = 1;
			} else {
				text = "[������]";
				bgcolor = "#FFFFFF";
				tempisusing = 0;
			}
			$("#idx"+idx).closest("tr").css("background-color",bgcolor);
			$("#idx"+idx).html("<a href=javascript:FnIsUsing('"+ idx +"','"+ tempisusing +"')>"+text+"</a>");
        }
    });
}

// �����Ȳ �˾�
function PopItemStock(gubuncode,itemid,itemoption){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun="+ gubuncode +"&itemid="+ itemid +"&itemoption="+ itemoption,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

$(function(){
    // init select
	<% if mastercode > 0 then %>
    mkbutton(<%=mastercode%>);
	<% end if %>

	// checkboxó��
	var optarr = "<%=trim(detailcode)%>"
	if (optarr) {
		optarr = optarr.replace(/ /g, '');
		var itemidslice = optarr.split(','), i;
		var itemIdResult = '';

		$("input[name='detailcode']").each(function(){
			for (i = 0; i < itemidslice.length; i++) {
				if (itemidslice[i] == $(this).val()) {
					$(this).prop("checked",true);
				}
			}
		});
	}
});
</script>
<div class="content scrl" style="top:40px;">
	<div class="pad20">
		<!-- ��� �˻��� ���� -->
		<div>
			<h1></h1>
			<form name="refreshFrm" method="get">
			<input type="hidden" name="menupos" value="<%= request("menupos") %>">
			<input type="hidden" name="page" value="">
			<table class="tbType1 listTb">
				<tr bgcolor="#FFFFFF">
					<th width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</th>
					<td style="text-align:left;">
						��ȹ�� : <%=DrawSelectAllView("mastercode",mastercode,"mkbutton")%><br/><br/>
						<input type="radio" name="flaglist" value="1" <%=chkiif(flaglist = 1,"checked","")%>/> ��ǰ����Ʈ <input type="radio" name="flaglist" value="2" <%=chkiif(flaglist = 2,"checked","")%>/> �ɼǸ���Ʈ
						<!--<select name="mdpick">
							<option value="">PICK����</option>
							<option value="0" <% if mdpick = "0" then response.write " selected"%>>X</option>
							<option value="1" <% if mdpick = "1" then response.write " selected"%>>O</option>
						</select>-->
						<br><br>�귣��:
						<% drawSelectBoxDesignerwithName "ebrand", sBrand %>
					</td>
					<th>��ǰ �ڵ�</th>
					<td style="text-align:left;">
						<textarea rows="4" cols="80" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
					</td>
					<td width="50" bgcolor="<%= adminColor("gray") %>">
						<input type="button" class="button_s" value="�˻�" onclick="refreshFrm.submit();">
					</td>
				</tr>
				<tr>
					<td colspan="5">
						<div id="submenu" style="text-align:left;"></div>
					</td>
				</tr>
			</table>
			</form>
		</div>
		<!-- �˻� �� -->
		<!-- �׼� ���� -->
		<div class="tPad15">
			<form name="frmarr" method="post" action="">
			<input type="hidden" name="menupos" value="<%= request("menupos") %>">
			<input type="hidden" name="mode" value="">
			<table class="tbType1 listTb">
			<tr>
				<th width="150">�ڵ� �� ��ǰ ����</th>
				<td style="text-align:left;">
					<div style="float:left;">
						<input type="button" value="�ڵ� ����" onclick="popGroupManage();" class="button">
					</div>
				</td>
			</tr>
			<% if mastercode > 0 then %>
			<tr>
				<th>�����̵� ����</th>
				<td style="text-align:left;">
					<input type="button" value="<%=getMasterCodeName(mastercode)%> �����̵� ����" onclick="popSlideManage('<%=mastercode%>','<%=menu%>');" class="button">
					<div class="tPad15">
						<strong>�̸����� : </strong>
						<input type="button" class="button" value="<%=getMasterCodeName(mastercode)%>" onclick="popSlideView('<%=mastercode%>','0','<%=menu%>')">&nbsp;
						<%=DrawDetailButtons(mastercode,"popSlideView",menu)%>
					</div>
				</td>
			</tr>
			<% end if %>
			</table>
			</form>
			<form name="frmreg" method="post" action="/admin/multiexhibitionmanage/lib/manage_proc.asp">
				<input type="hidden" name="mode" value="pickreg">
				<input type="hidden" name="eid" value="">
				<input type="hidden" name="mdpick" value="">
				<input type="hidden" name="poscode" value="<%=poscode%>">
				<input type="hidden" name="page" value="<%=page%>">
			</form>
		</div>
		<div class="tPad15">
			<!-- ����Ʈ ���� -->
			<table class="tbType1 listTb">
			<form name="fitem" method="post" style="margin:0px;">
			<% IF oExhibition.FResultCount>0 Then %>
				<tr height="25" bgcolor="FFFFFF">
					<td colspan="<%=chkiif(flaglist=1,"11","13")%>" style="text-align:left;">
						<div style="float:left;">
							�˻���� : <b><%= oExhibition.FTotalCount %></b>
							&nbsp;
							������ : <b><%= page %>/ <%= oExhibition.FTotalPage %></b>
						</div>
						<div style="float:right;">
							<input type="button" value="��ǰ���" class="button" onclick="popRegItems();">
						</div>
					</td>
					
				</tr>
				<tr bgcolor="<%= adminColor("tabletop") %>">
					<td>��ȣ</td>
					<td>���� </td>
					<td>�̹��� </td>
					<td><%=chkiif(flaglist=1,"��ǰ��ȣ","����-��ǰ��ȣ-��ǰ�ڵ�")%></td>
					<td>��ǰ�� </td>
					<td>��ü���̵� </td>
					<td>�ǸŰ�</td>
					<td>����</td>
					<td>��౸��</td>
					<% if flaglist <> 1 then %>
					<td>�����Ȳ</td>
					<td>��ǰ����</td>
					<% end if %>
				</tr>
					<% if flaglist = 1 then %>
						<% For i =0 To oExhibition.FResultCount -1 %>
						<tr bgcolor="#FFFFFF">
							<td><%= oExhibition.FItemList(i).Fidx %></td>
							<td>
								<%=getMasterCodeName(oExhibition.FItemList(i).Fmastercode)%>
							</td>
							<td><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=oExhibition.FItemList(i).Fitemid%>" target="_blank"><img src="<%= db2html(oExhibition.FItemList(i).FImageList) %>" width="40" height="40" border="0" style="cursor:pointer"></a></td>
							<td><%= oExhibition.FItemList(i).Fitemid %> </td>
							<td><%= oExhibition.FItemList(i).fitemname %> &nbsp;&nbsp;<span style="color:red"><%=chkiif(oExhibition.FItemList(i).Foptcnt>0,"(+"& oExhibition.FItemList(i).Foptcnt &" <span onclick='optmore("& oExhibition.FItemList(i).Fitemid &")' style='cursor:pointer'>more</span>)","")%></span></td>
							<td><%= oExhibition.FItemList(i).fmakerid %> </td>
							<td>
								<%
								Response.Write FormatNumber(oExhibition.FItemList(i).Forgprice,0)
								'���ΰ�
								if oExhibition.FItemList(i).Fsailyn="Y" then
									Response.Write "<br><font color=#F08050>("&CLng((oExhibition.FItemList(i).Forgprice-oExhibition.FItemList(i).Fsailprice)/oExhibition.FItemList(i).Forgprice*100) & "%��)" & FormatNumber(oExhibition.FItemList(i).Fsailprice,0) & "</font>"
								end if
								'������
								if oExhibition.FItemList(i).FitemCouponYn="Y" then
									Select Case oExhibition.FItemList(i).FitemCouponType
										Case "1"
											Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oExhibition.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
										Case "2"
											Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oExhibition.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
									end Select
								end if
							%>
							</td><%'�ǸŰ�%>
							<td>
								<%
								Response.Write fnPercent(oExhibition.FItemList(i).Forgsuplycash,oExhibition.FItemList(i).Forgprice,1)
								'���ΰ�
								if oExhibition.FItemList(i).Fsailyn="Y" then
									Response.Write "<br><font color=#F08050>" & fnPercent(oExhibition.FItemList(i).Fsailsuplycash,oExhibition.FItemList(i).Fsailprice,1) & "</font>"
								end if
								'������
								if oExhibition.FItemList(i).FitemCouponYn="Y" then
									Select Case oExhibition.FItemList(i).FitemCouponType
										Case "1"
											if oExhibition.FItemList(i).Fcouponbuyprice=0 or isNull(oExhibition.FItemList(i).Fcouponbuyprice) then
												Response.Write "<br><font color=#5080F0>" & fnPercent(oExhibition.FItemList(i).Fbuycash,oExhibition.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
											else
												Response.Write "<br><font color=#5080F0>" & fnPercent(oExhibition.FItemList(i).Fcouponbuyprice,oExhibition.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
											end if
										Case "2"
											if oExhibition.FItemList(i).Fcouponbuyprice=0 or isNull(oExhibition.FItemList(i).Fcouponbuyprice) then
												Response.Write "<br><font color=#5080F0>" & fnPercent(oExhibition.FItemList(i).Fbuycash,oExhibition.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
											else
												Response.Write "<br><font color=#5080F0>" & fnPercent(oExhibition.FItemList(i).Fcouponbuyprice,oExhibition.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
											end if
									end Select
								end if
								%>
							</td><%'����%>
							<td><%=fnColor(oExhibition.FItemList(i).Fmwdiv,"mw")%><br/>
								<%
									If oExhibition.FItemList(i).Fdeliverytype = "1" Then
										response.write "�ٹ�"
									ElseIf oExhibition.FItemList(i).Fdeliverytype = "2" Then
										response.write "����"
									ElseIf oExhibition.FItemList(i).Fdeliverytype = "4" Then
										response.write "�ٹ�"
									ElseIf oExhibition.FItemList(i).Fdeliverytype = "9" Then
										response.write "����"
									ElseIf oExhibition.FItemList(i).Fdeliverytype = "7" Then
										response.write "����"
									End If
								%>
							</td>
						</tr>
						<tr id="optlist<%=oExhibition.FItemList(i).Fitemid%>" style="display:none;">
							<td colspan='9'>
								<table width='100%' id="optlist<%=oExhibition.FItemList(i).Fitemid%>_2"></table>
							</td>
						</tr>
						<% Next %>
					<% else %>
						<% For i = 0 To oExhibition.FResultCount -1 %>
						<tr bgcolor="#FFFFFF">
							<td><%= oExhibition.FItemList(i).Fidx %></td>
							<td>
								<%=getMasterCodeName(oExhibition.FItemList(i).Fmastercode)%>
								<br/><br/>
								<%="��"&getDetailCodeName(oExhibition.FItemList(i).Fmastercode,oExhibition.FItemList(i).Fdetailcode)%>					
							</td>
							<td><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=oExhibition.FItemList(i).Fitemid%>" target="_blank"><img src="<%= db2html(oExhibition.FItemList(i).FImageList) %>" width="40" height="40" border="0" style="cursor:pointer"></a></td>
							<td><%= oExhibition.FItemList(i).Fgubuncode %>-<span style="font-weight:800"><%= oExhibition.FItemList(i).Fitemid %></span>-<%= oExhibition.FItemList(i).Foptioncode %> </td>
							<td><%= oExhibition.FItemList(i).fitemname %> </td>
							<td><%= oExhibition.FItemList(i).fmakerid %> </td>
							<td>
								<%
								Response.Write FormatNumber(oExhibition.FItemList(i).Forgprice,0)
								'���ΰ�
								if oExhibition.FItemList(i).Fsailyn="Y" then
									Response.Write "<br><font color=#F08050>("&CLng((oExhibition.FItemList(i).Forgprice-oExhibition.FItemList(i).Fsailprice)/oExhibition.FItemList(i).Forgprice*100) & "%��)" & FormatNumber(oExhibition.FItemList(i).Fsailprice,0) & "</font>"
								end if
								'������
								if oExhibition.FItemList(i).FitemCouponYn="Y" then
									Select Case oExhibition.FItemList(i).FitemCouponType
										Case "1"
											Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oExhibition.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
										Case "2"
											Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oExhibition.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
									end Select
								end if
							%>
							</td><%'�ǸŰ�%>
							<td>
								<%
								Response.Write fnPercent(oExhibition.FItemList(i).Forgsuplycash,oExhibition.FItemList(i).Forgprice,1)
								'���ΰ�
								if oExhibition.FItemList(i).Fsailyn="Y" then
									Response.Write "<br><font color=#F08050>" & fnPercent(oExhibition.FItemList(i).Fsailsuplycash,oExhibition.FItemList(i).Fsailprice,1) & "</font>"
								end if
								'������
								if oExhibition.FItemList(i).FitemCouponYn="Y" then
									Select Case oExhibition.FItemList(i).FitemCouponType
										Case "1"
											if oExhibition.FItemList(i).Fcouponbuyprice=0 or isNull(oExhibition.FItemList(i).Fcouponbuyprice) then
												Response.Write "<br><font color=#5080F0>" & fnPercent(oExhibition.FItemList(i).Fbuycash,oExhibition.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
											else
												Response.Write "<br><font color=#5080F0>" & fnPercent(oExhibition.FItemList(i).Fcouponbuyprice,oExhibition.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
											end if
										Case "2"
											if oExhibition.FItemList(i).Fcouponbuyprice=0 or isNull(oExhibition.FItemList(i).Fcouponbuyprice) then
												Response.Write "<br><font color=#5080F0>" & fnPercent(oExhibition.FItemList(i).Fbuycash,oExhibition.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
											else
												Response.Write "<br><font color=#5080F0>" & fnPercent(oExhibition.FItemList(i).Fcouponbuyprice,oExhibition.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
											end if
									end Select
								end if
								%>
							</td><%'����%>
							<td><%=fnColor(oExhibition.FItemList(i).Fmwdiv,"mw")%><br/>
								<%
									If oExhibition.FItemList(i).Fdeliverytype = "1" Then
										response.write "�ٹ�"
									ElseIf oExhibition.FItemList(i).Fdeliverytype = "2" Then
										response.write "����"
									ElseIf oExhibition.FItemList(i).Fdeliverytype = "4" Then
										response.write "�ٹ�"
									ElseIf oExhibition.FItemList(i).Fdeliverytype = "9" Then
										response.write "����"
									ElseIf oExhibition.FItemList(i).Fdeliverytype = "7" Then
										response.write "����"
									End If
								%>
							</td>
							<td><a href="javascript:PopItemStock('<%=oExhibition.FItemList(i).Fgubuncode%>','<%=oExhibition.FItemList(i).Fitemid%>','<%=oExhibition.FItemList(i).Foptioncode%>')" title="�����Ȳ �˾�">[����]</a></td>
							<td><input type="button" value="����" onclick="fnDelItem('<%= oExhibition.FItemList(i).Fidx%>');"/></td>
							<%'��౸��%>	
						</tr>
						<% Next %>
					<% end if %>
			<% else %>
				<tr bgcolor="#FFFFFF">
					<td colspan="3" class="page_link">[�˻������ �����ϴ�.]</td>
				</tr>
			<% End IF %>
				<tr bgcolor="#FFFFFF">
					<td colspan="11" align="center">
					<!-- ������ ���� -->
						<a href="?page=1&isusingbox=<%=isusing%>&mastercode=<%=mastercode%>&detailcode=<%=detailcode%>&flaglist=<%=flaglist%>&menupos=<%=poscode%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/pprev_btn.gif" width="10" height="10" border="0"></a>
						<% if oExhibition.HasPreScroll then %>
							<span class="list_link"><a href="?page=<%= oExhibition.StartScrollPage-1 %>&isusingbox=<%=isusing%>&mastercode=<%=mastercode%>&detailcode=<%=detailcode%>&flaglist=<%=flaglist%>&menupos=<%=poscode%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
						<% else %>
						&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;
						<% end if %>
						<% for i = 0 + oExhibition.StartScrollPage to oExhibition.StartScrollPage + oExhibition.FScrollCount - 1 %>
							<% if (i > oExhibition.FTotalpage) then Exit for %>
							<% if CStr(i) = CStr(oExhibition.FCurrPage) then %>
							<span class="page_link"><font color="red"><b><%= i %>&nbsp;&nbsp;</b></font></span>
							<% else %>
							<a href="?page=<%= i %>&isusingbox=<%=isusing%>&mastercode=<%=mastercode%>&detailcode=<%=detailcode%>&flaglist=<%=flaglist%>&menupos=<%=poscode%>" class="list_link"><font color="#000000"><%= i %>&nbsp;&nbsp;</font></a>
							<% end if %>
						<% next %>
						<% if oExhibition.HasNextScroll then %>
							<span class="list_link"><a href="?page=<%= i %>&isusingbox=<%=isusing%>&mastercode=<%=mastercode%>&detailcode=<%=detailcode%>&flaglist=<%=flaglist%>&menupos=<%=poscode%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
						<% else %>
						&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;
						<% end if %>
						<a href="?page=<%= oExhibition.FTotalpage %>&isusingbox=<%=isusing%>&mastercode=<%=mastercode%>&detailcode=<%=detailcode%>&flaglist=<%=flaglist%>&menupos=<%=poscode%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/nnext_btn.gif" width="10" height="10" border="0"></a>
					<!-- ������ �� -->
					</td>
				</tr>
			</form>
			</table>
			<form name="itemdel" method="post" action="/admin/multiexhibitionmanage/lib/manage_proc.asp">
			<input type="hidden" name="eidx" value=""/>
			<input type="hidden" name="mode" value="delitem" />
			<input type="hidden" name="poscode" value="<%=poscode%>"/>
			<input type="hidden" name="page" value="<%=page%>"/>
			</form>
			<!-- ����Ʈ �� -->
		</div>
	</div>
</div>
<% Set oExhibition = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->