<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ��ǰ���
' History : 2013.12.11 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim target, cdl, cdm, cds, page, oitem, i, sailyn, couponyn, mwdiv,defaultmargin, keyword
dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, deliverytype, limityn, vatyn, mode
dim dispCate, itemexists, reload, infodivYn, infodiv
	itemid      = request("itemid")
	itemname    = request("itemname")
	makerid     = request("makerid")
	sellyn      = request("sellyn")
	usingyn     = request("usingyn")
	danjongyn   = request("danjongyn") 
	mwdiv       = request("mwdiv")
	limityn     = request("limityn") 
	sailyn      = request("sailyn")
	couponyn	= request("couponyn")
	defaultmargin = request("defaultmargin")
	deliverytype       = request("deliverytype")
	keyword		= request("keyword")
	cdl = request("cdl")
	cdm = request("cdm")
	cds = request("cds")
	page = request("page")
	mode = request("mode")
	dispCate = requestCheckvar(request("disp"),16)
	itemexists = requestCheckvar(request("itemexists"),1)
	reload = request("reload")
	infodiv  = request("infodiv")
	infodivYn  = requestCheckvar(request("infodivYn"),10)

if mode="" then mode="regitem"
If infodiv <> "" Then
	infodivYn = "Y"	
End If
if (page="") then page=1
'if sailyn="" then sailyn="N"			'�������������� �˻��ȰŶ�� �⺻��: ���ξ���(������ ����)
'if couponyn="" then couponyn="N"
'if sellyn = "" then sellyn ="Y"
if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	arrTemp = Split(itemid,",")

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

	itemid = left(arrItemid,len(arrItemid)-1)
end if

'if reload="" and itemexists="" then itemexists="N"
itemexists="Y"

set oitem = new CItem
	oitem.FPageSize         = 50
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
	oitem.FRectitemexists	= itemexists
	oitem.FRectInfodivYn    = infodivYn
	oitem.FRectInfodiv    = infodiv	
	oitem.GetItem_Evaluate_exclude
%>

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">

	function regitem(mode){
		var regitem = window.open('/admin/itemmaster/Item_Evaluate_exclude_pop.asp?mode='+mode,'regitem','width=1024,height=768,scrollbars=yes,resizable=yes');
		regitem.focus();
	}

	// ������ �̵�
	function goPage(pg){
		document.frm.page.value=pg;
		document.frm.submit();
	}

	// ���õ� �׸� ����/����
	function doedit(mode){
		var i, chk=0;
		var frm = document.frm_list;

		if (frm.Eval_excludeitemid.length){
			for(i=0;i<frm.Eval_excludeitemid.length;i++){
				if(frm.Eval_excludeitemid[i].checked){
					chk++;
				}
			}
		} else {
			if(frm.Eval_excludeitemid.checked){
				chk++;
			}
		}

		if(chk==0){
			alert("��ǰ�� ��� �Ѱ��̻� �������ֽʽÿ�.");
			return;
		} else {
			if(confirm("�����Ͻ� " + chk + "����  �׸��� ��� ���� �Ͻðڽ��ϱ�?")){
				frm.mode.value=mode;
				frm.action="/admin/itemmaster/Item_Evaluate_exclude_process.asp";
				frm.submit();
			} else {
				return;
			}
		}
	}

	//��ü ����
	function jsChkAll(){	
	var frm;
	frm = document.frm_list;
		if (frm.chkAll.checked){			      
		   if(typeof(frm.Eval_excludeitemid) !="undefined"){
		   	   if(!frm.Eval_excludeitemid.length){
			   	 	frm.Eval_excludeitemid.checked = true;	   	 
			   }else{
					for(i=0;i<frm.Eval_excludeitemid.length;i++){
						frm.Eval_excludeitemid[i].checked = true;
				 	}		
			   }	
		   }	
		} else {	  
		  if(typeof(frm.Eval_excludeitemid) !="undefined"){
		  	if(!frm.Eval_excludeitemid.length){
		   	 	frm.Eval_excludeitemid.checked = false;	  
		   	}else{
				for(i=0;i<frm.Eval_excludeitemid.length;i++){
					frm.Eval_excludeitemid[i].checked = false;
				}	
			}		
		  }	
		}
	}

</script>

<!-- ��� �˻��� ���� -->
<table width="100%" align="center" celiadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="reload" value="ON">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		* �귣�� :
		<%	drawSelectBoxDesignerWithName "makerid", makerid %>
		&nbsp;&nbsp;
		* ��ǰ�ڵ� :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="40" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(��ǥ�� �����Է°���)
		<p>
		* ��ǰ�� :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20">			
		&nbsp;&nbsp;
		* �˻�Ű���� : <input type="text" class="text" name="keyword" value="<%=keyword%>" size="40"><font color="gray" size="2">(����:�������ֽ��ϴ�.)</font>
		<p>
		<span style="white-space:nowrap;">* ���� <!-- #include virtual="/common/module/categoryselectbox.asp"--></span>
		<p>
		<span style="white-space:nowrap;">* ����ī�װ� : <!-- #include virtual="/common/module/dispCateSelectBox.asp"--></span>
     	<p>
     	<span style="white-space:nowrap;">* ǰ�������Է¿��� :
     	<select class="select" name="infodivYn">
	        <option value="">��ü</option>
	        <option value="N" <%= CHKIIF(infodivYn="N","selected","") %> >�Է�����</option>
	        <option value="Y" <%= CHKIIF(infodivYn="Y","selected","") %> >�Է¿Ϸ�</option>
        </select></span>
        &nbsp;&nbsp;
		<span style="white-space:nowrap;">* ǰ�� : <% drawSelectBoxinfodiv "infodiv", infodiv, "" %></span>		
		<p>
		* �Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>
		&nbsp;&nbsp;
     	* ���:<% drawSelectBoxUsingYN "usingyn", usingyn %>
		&nbsp;&nbsp;
     	* ����:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
		&nbsp;&nbsp;
     	* ����:<% drawSelectBoxLimitYN "limityn", limityn %>
		&nbsp;&nbsp;
     	* ���:<% drawSelectBoxMWU "mwdiv", mwdiv %>
		&nbsp;&nbsp;
     	* ����: <% drawSelectBoxSailYN "sailyn", sailyn %>
		<p>
     	* ����: <% drawSelectBoxCouponYN "couponyn", couponyn %>
		&nbsp;&nbsp;
     	* ���:<% drawBeadalDiv "deliverytype",deliverytype %>
		<!--&nbsp;&nbsp;
		* ��ǰ��Ͽ��� : -->
		<%' drawSelectBoxisusingYN "itemexists", itemexists,"" %>		
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<Br>

<!-- �׼� ���� -->
<table width="100%" align="center" celiadding="0" cellspacing="0" class="a" style="padding:10 0 0 0;">
<form name="frm_list" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="page" value="<%=page%>">
<tr>
	<td align="left">		
		[ON]��ǰ����>>��ǰ���� ���� <font color="red">ǰ��(�Ƿ���,��ǰ(����깰),������ǰ,�ǰ���ɽ�ǰ/ü��������ǰ)</font>�� �ش�Ǵ� ��ǰ��, �Ϸ翡 �ѹ� ������ �̰��� �ڵ� ����˴ϴ�.
	</td>
	<td align="right">
		<% if oitem.FResultCount>0 then %>
			<input type="button" value="���û���" onClick="doedit('delitem')" class="button">
			&nbsp;
		<% end if %>
			
		<input type="button" value="�űԵ��" onClick="regitem('regitem')" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" celiadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%=FormatNumber(oitem.FTotalCount,0)%></b>
		&nbsp;
		������ : <b><%= page %>/<%=FormatNumber(oitem.Ftotalpage,0)%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<td align="center">��ǰID</td>
	<td align="center">�̹���</td>
	<td align="center">�귣��</td>
	<td align="center">��ǰ��</td>
	<td align="center">�ǸŰ�</td>
	<td align="center" nowrap>���<br>����</td>	
	<td align="center" nowrap>���<br>����</td>
	<td align="center" nowrap>�Ǹ�<br>����</td>	
	<td align="center" nowrap>���<br>����</td>	
	<td align="center" nowrap>����<br>����</td>	
</tr>
<% if oitem.FResultCount=0 then %>
	<tr align="center">
		<td colspan="20" height="30" bgcolor="#FFFFFF">���(�˻�)�� ������ �����ϴ�.</td>
	</tr>
<%
else

for i=0 to oitem.FResultCount - 1
%>
<tr align="center" bgcolor="#FFFFFF">
	<td  align="center"><input type="checkbox" name="Eval_excludeitemid" value="<%= oitem.FItemList(i).fEval_excludeitemid %>"></td>
	<td align="center"><A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).FItemId %>" target="_blank"><%= oitem.FItemList(i).FItemId %></a></td>
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
	<td align="center"><%=fnColor(oitem.FItemList(i).IsUpcheBeasong(),"delivery")%></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %></td>
	<td align="center">
	<%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %>
	</td>
	<td align="center">
	<%= fnColor(oitem.FItemList(i).Fisusing,"yn") %>
	</td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center" bgcolor="#FFFFFF">
		<!-- ����¡ó�� -->
		<% if oitem.HasPreScroll then %>
			<a href="javascript:goPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oitem.HasNextScroll then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% end if %>

</form>
</table>

<%
set oitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->