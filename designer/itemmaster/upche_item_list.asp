<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<%

dim itemid, makerid, itemname, waititemid
dim sellyn, isusing, danjongyn, limityn, mwdiv
dim page, cdl, cdm, cds, dispCate
dim infodivYn, itemdiv,overseaYN

itemid  = RequestCheckVar(request("itemid"),10)
makerid = RequestCheckVar(request("makerid"),32)
itemname = RequestCheckVar(request("itemname"),32)

sellyn  = RequestCheckVar(request("sellyn"),10)
isusing = RequestCheckVar(request("isusing"),10)
danjongyn = RequestCheckVar(request("danjongyn"),10)
limityn = RequestCheckVar(request("limityn"),10)
mwdiv = RequestCheckVar(request("mwdiv"),10)

page = RequestCheckVar(request("page"),10)

cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
dispCate = requestCheckvar(request("disp"),16)
infodivYn  = requestCheckvar(request("infodivYn"),10)
waititemid = requestCheckvar(request("waititemid"),10)
itemdiv = requestCheckvar(request("itemdiv"),2)
overseaYN= requestCheckvar(request("overseaYN"),1)
if (sellyn="") then sellyn="A"

if (page="") then page=1

''if (isusing="") then isusing="Y"
''����ϴ� ��ǰ�� ǥ�÷� ����
isusing="Y"

'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.01;������)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

'==============================================================================
dim oitem

set oitem = new CItem

oitem.FRectMakerId = session("ssBctID")
oitem.FRectItemid = itemid
oitem.FRectItemName = itemname
oitem.FRectDanjongyn = danjongyn
oitem.FRectLimityn = limityn
oitem.FRectMWDiv = mwdiv
oitem.FPageSize = 30
oitem.FCurrPage = page
oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds
oitem.FRectDispCate		= dispCate
oitem.FRectInfodivYn    = infodivYn
oitem.FRectSellReserve = "Y"
oitem.FRectwaititemid  = waititemid
oitem.FRectItemDiv  = itemdiv
oitem.FRectdeliverOverseas = overseaYN

if (sellyn <> "A") then
    oitem.FRectSellYN = sellyn
end if

if (isusing <> "A") then
    oitem.FRectIsUsing = isusing
end if


oitem.GetProductList

dim i

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function NextPage(ipage){
	document.frm.page.value= ipage;
	SubmitSearch();
}
function SubmitSearch(){
	document.frm.action = "/designer/itemmaster/upche_item_list.asp";
	document.frm.target = "";

	if ((document.frm.itemid.value != "") && ((document.frm.itemid.value*0) != 0)) {
	    alert("��ǰ�ڵ忡�� ���ڸ� �Է��� �����մϴ�.");
	    document.frm.itemid.focus();
	    return;
    }
	document.frm.submit();
}


// ============================================================================
// �⺻��������
function editItemInfo(itemid) {

	var param = "itemid=" + itemid;
	popwin = window.open('upche_item_infomodify.asp?' + param ,'editItemInfoPop','width=1100,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ============================================================================
// �ɼǼ���
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('upche_item_optionmodify.asp?' + param ,'editItemOption','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function editSimpleItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/common/pop_upche_simpleitemedit.asp?' + param ,'editSimpleItemOption','width=500,height=650,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ============================================================================
// �̹�������
function editItemImage(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('upche_item_imagemodify.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//�����ٿ�
function nowListExcelDown()
{
	if ((document.frm.itemid.value != "") && ((document.frm.itemid.value*0) != 0)) {
	    alert("��ǰ�ڵ忡�� ���ڸ� �Է��� �����մϴ�.");
	    document.frm.itemid.focus();
	    return;
    }

	document.frm.action = "/designer/itemmaster/upche_item_list_XL.asp";
	document.frm.target = "XLdown";
	document.frm.submit();
}

//�����ٿ�_�ɼ�����
function nowListExcelDownOption(){
	if ((document.frm.itemid.value != "") && ((document.frm.itemid.value*0) != 0)) {
	    alert("��ǰ�ڵ忡�� ���ڸ� �Է��� �����մϴ�.");
	    document.frm.itemid.focus();
	    return;
    }

	document.frm.action = "/designer/itemmaster/upche_item_list_option_XL.asp";
	document.frm.target = "XLdown";
	document.frm.submit();
}

//ǰ������ �ϰ����� �˾�
function popUploadXLSItemInfo() {
	popwin = window.open('pop_item_infoUploadFile.asp','popInfoUpload','width=520,height=300,scrollbars=no,resizable=no');
	popwin.focus();
}

//������������ �ϰ����� �˾�
function popUploadXLSSafetyInfo() {
	popwin = window.open('./itemInfoFile/pop_item_safetyinfoUploadFile.asp','popInfoUpload','width=520,height=270,scrollbars=no,resizable=no');
	popwin.focus();
}

//�ؿܹ������ �ϰ����� �˾�
function popUploadXLSOverSeaInfo(){
    popwin =  window.open('./itemInfoFile/pop_item_overseainfoUploadFile.asp','popInfoUpload','width=520,height=270,scrollbars=no,resizable=no');
	popwin.focus();
}
</script>


<!-- ǥ ��ܹ� ����-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			��ǰ�ڵ� :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="11" maxlength="11" onKeyPress="if (event.keyCode == 13) SubmitSearch();">
			&nbsp;
			��ǰ�� :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="20" onKeyPress="if (event.keyCode == 13) SubmitSearch();">
			<br>
			����<!-- #include virtual="/common/module/categoryselectbox.asp"-->
			&nbsp; ����ī�װ� : <!-- #include virtual="/common/module/dispCateSelectBox_upche.asp"-->
			<input type="hidden" name="waititemid" value=""> <!-- for play auto -->
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:SubmitSearch();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>
			&nbsp;
			����:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
	     	&nbsp;
	     	����:<% drawSelectBoxLimitYN "limityn", limityn %>
	     	&nbsp;
	     	�ŷ�����:<% drawSelectBoxMWU "mwdiv", mwdiv %>
	     	&nbsp;
	     	<font color="red">ǰ�������Է¿���</font>
	     	<select class="select" name="infodivYn">
            <option value="">��ü</option>
            <option value="N" <%= CHKIIF(infodivYn="N","selected","") %> >�Է�����</option>
            <option value="Y" <%= CHKIIF(infodivYn="Y","selected","") %> >�Է¿Ϸ�</option>
            </select>
	     	&nbsp;
			��ǰ����:<% drawSelectBoxItemDiv "itemdiv", itemdiv %>
			&nbsp;
			<font color="red">�ؿܹ�ۿ���</font>
			<select class="select" name="overseaYN">
            <option value="">��ü</option>
            <option value="N" <%= CHKIIF(overseaYN="N","selected","") %> >N</option>
            <option value="Y" <%= CHKIIF(overseaYN="Y","selected","") %> >Y</option>
            </select>

		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" class="a" >
<tr>
	<td align="left" style="padding-top:5px;">
		<input type="button" class="button" style="width:240px;background-color:#F8DFF0;" value="[��ǰ������ð���] �߰����� �ϰ����" onclick="popUploadXLSItemInfo()" title="Excel������ ���ε��Ͽ� [��ǰ������ð���] �߰����� �ϰ�����մϴ�." /> &nbsp;
		<input type="button" class="button" style="width:190px;background-color:#DFF8F0;" value="[�����������]���� �ϰ����" onclick="popUploadXLSSafetyInfo()" title="Excel������ ���ε��Ͽ� [�����������]���� �ϰ�����մϴ�." />
		<input type="button" class="button" style="width:190px;" value="[�ؿܹ��]���� �ϰ����" onclick="popUploadXLSOverSeaInfo()" title="Excel������ ���ε��Ͽ� [�ؿܹ��]���� �ϰ�����մϴ�." />
	</td>
	<td align="right" style="padding:5 0 5 0;">
	    <img src="/images/btn_excel.gif" style="cursor:pointer;display:inline;position:relative;top:5px" onClick="nowListExcelDown()" alt="��ǰ��Ͽ���" title="��ǰ��� �ٿ�ε�">(��ǰ���)
	    &nbsp;
	    <img src="/images/btn_excel.gif" style="cursor:pointer;display:inline;position:relative;top:5px" onClick="nowListExcelDownOption()" alt="�ɼ����Ի�ǰ��Ͽ���" title="��ǰ��� �ٿ�ε�(�ɼ�����)">(�ɼ�����)
	</td>
</tr>
</table>

	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	    <tr bgcolor="#FFFFFF">
	        <td colspan="14" align="right">�ѰǼ� : <%= oitem.FTotalCount %> </td>
	    </tr>
	    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td width="60">��ǰ�ڵ�</td>
			<td width="50">�̹���</td>
			<td width="100">�귣��ID</td>
			<td>��ǰ��</td>
			<td width="60">�ǸŰ�</td>
			<td width="60">���ް�</td>
			<td width="40">����</td>
			<td width="30">�ŷ�<br>����</td>
			<td width="30">�Ǹſ���</td>
			<td width="40">����<br>����</td>
			<td width="40">�ؿܹ��<br>����</td>
			<td width="50">�⺻<br>����</td>
			<td width="50">�̹���</td>
			<td width="70">�ɼ�/����<br>�ǸŰ���</td>
	    </tr>
<% if oitem.FresultCount<1 then %>
	    <tr bgcolor="#FFFFFF">
	    	<td colspan="14" align="center">[�˻������ �����ϴ�.]</td>
	    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
    	<% if (oitem.FItemList(i).Fisusing = "N") then %>
    	<tr class="a" height="25" bgcolor="<%= adminColor("gray") %>">
		<% else %>
		<tr class="a" height="25" bgcolor="#FFFFFF">
		<% end if %>
			<td align="center"><a href="http://www.10x10.co.kr/<%= oitem.FItemList(i).Fitemid %>" target="_blank"><%= oitem.FItemList(i).Fitemid %></a></td>
			<td align="center"><img src="<%= oitem.FItemList(i).FImgSmall %>" width="50" height="50" border="0" alt=""></td>
			<td align="center"><%= oitem.FItemList(i).Fmakerid %></td>
			<td align="left"><% =oitem.FItemList(i).Fitemname %>&nbsp;&nbsp;<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank"><font color="blue">(Ȯ���ϱ�)</font></a></td>
			<td align="right">
			    <%= FormatNumber(oitem.FItemList(i).Forgprice,0) %>
			    <%
			    '���ΰ�
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>("&CLng((oitem.FItemList(i).Forgprice-oitem.FItemList(i).Fsailprice)/oitem.FItemList(i).Forgprice*100) & "%��)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
			end if
			'������
			if oitem.FItemList(i).FitemCouponYn="Y" then
				Select Case oitem.FItemList(i).FitemCouponType
					Case "1"
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
					Case "2"
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
				end Select
			end if
			    %>
			</td>
			<td align="right"><%= FormatNumber(oitem.FItemList(i).Forgsuplycash,0) %>
			    <%
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
			    %>
			</td>
			<td align="right">
			<%
			Response.Write fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).Forgprice,1)
			'���ΰ�
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>" & fnPercent(oitem.FItemList(i).Fsailsuplycash,oitem.FItemList(i).Fsailprice,1) & "</font>"
			end if
			'������
			if oitem.FItemList(i).FitemCouponYn="Y" then
				Select Case oitem.FItemList(i).FitemCouponType
					Case "1"
						if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
						else
							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Fcouponbuyprice,oitem.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
						end if
					Case "2"
						if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
						else
							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Fcouponbuyprice,oitem.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
						end if
				end Select
			end if
		%>
	        </td>
			<td align="center">
				<font color="<%= mwdivColor(oitem.FItemList(i).Fmwdiv) %>"><%= mwdivName(oitem.FItemList(i).Fmwdiv) %></font>
			</td>

			<td align="center">
				<%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %>
			<%IF oitem.FItemList(i).Fsellreservedate <>"" THEN%><div>���¿���: <%=oitem.FItemList(i).Fsellreservedate%></div><%END IF%>
			</td>
			<td align="center">
        		<% if (oitem.FItemList(i).Flimityn = "Y") then %>
             		<%= fnColor(oitem.FItemList(i).Flimityn,"yn") %>
             		<br>(<%= (oitem.FItemList(i).Flimitno - oitem.FItemList(i).Flimitsold) %>)
        		<% else %>
              		<%= fnColor(oitem.FItemList(i).Flimityn,"yn") %>
       			<% end if %>
			</td>
			<td align="center"><%=fnColor(oitem.FItemList(i).FdeliverOverseas,"yn")%>
		    </td>
		    <td align="center">
		    	<img src="/images/icon_modify.gif" border="0" align="absbottom" onClick="editItemInfo('<%= oitem.FItemList(i).FItemId %>');" style="cursor:pointer">
		    </td>
		    <td align="center">
		    	<a href="javascript:editItemImage('<%= oitem.FItemList(i).FItemId %>')">
		    	<img src="/images/icon_modify.gif" border="0" align="absbottom">
		    	</a>
		    </td>
		    <td align="center">
        <% if (oitem.FItemList(i).Fmwdiv = "U") then %>
		      	<a href="javascript:editSimpleItemOption('<%= oitem.FItemList(i).FItemId %>')">
		      	<img src="/images/icon_modify.gif" border="0" align="absbottom">
		      	</a>
        <% else %>
		      	<a href="javascript:editSimpleItemOption('<%= oitem.FItemList(i).FItemId %>')">
		      	<b>[</b>������û<b>]</b>
		      	</a>
        <% end if %>

		    </td>
		</tr>
		<% next %>
	</table>
<% end if %>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
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
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<iframe id="XLdown" name="XLdown" src="about:blank" frameborder="0" width="110" height="110"></iframe>

<% set oitem = nothing %>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
