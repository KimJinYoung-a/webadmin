<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �¶��� �ؿ��ǸŴ���ǰ
' History : 2013.05.06 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->

<%
response.write "������� �Ŵ� �Դϴ�."
response.end

dim itemid, itemname, makerid, sellyn, usingyn, mwdiv, limityn, overSeaYn, weightYn, danjongyn
dim cdl, cdm, cds, sortDiv, sortDiv2, sellcash1, sellcash2, vDate1, vDate2, page, i
dim itemrackcode, vRegUserID, vIsReg, reload
dim sitename
	itemid		= request("itemid")
	itemname	= request("itemname")
	makerid		= request("makerid")
	sellyn		= request("sellyn")
	usingyn		= request("usingyn")
	mwdiv		= request("mwdiv")
	limityn		= request("limityn")
	overSeaYn	= request("overSeaYn")
	weightYn	= request("weightYn")
	itemrackcode= request("itemrackcode")
	sortDiv		= request("sortDiv")
	sortDiv2	= request("sortDiv2")
	vRegUserID	= request("reguserid")
	vIsReg		= request("isreg")
	sellcash1	= request("sellcash1")
	sellcash2	= request("sellcash2")
	vDate1		= request("date1")
	vDate2		= request("date2")
	cdl = request("cdl")
	cdm = request("cdm")
	cds = request("cds")
	page = request("page")
	reload = request("reload")
    sitename = request("sitename")
	danjongyn   = requestCheckvar(request("danjongyn"),10)

'�⺻��
if (page="") then page=1
if sitename="" then sitename="WSLWEB"
'if reload<>"ON" and mwdiv="" then mwdiv="MW"
if reload<>"ON" and overSeaYn="" then overSeaYn="Y"
'if reload<>"ON" and weightYn="" then weightYn="Y"
if sortDiv="" then sortDiv="new"
if sortDiv2="" then sortDiv2="weightup"
if reload<>"ON" and sellyn="" then sellyn="YS"
if reload<>"ON" and usingyn="" then usingyn="Y"
if vIsReg="" then vIsReg="x"

if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	arrTemp = Split(itemid,",")

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

dim oitem
set oitem = new COverSeasItem
	oitem.FPageSize         = 30
	oitem.FCurrPage         = page
	oitem.FRectMakerid      = makerid
	oitem.FRectItemid       = itemid
	oitem.FRectItemName     = itemname
	oitem.FRectSellYN       = sellyn
	oitem.FRectIsUsing      = usingyn
	oitem.FRectLimityn      = limityn
	oitem.FRectMWDiv        = mwdiv
	oitem.FRectIsOversea	= overSeaYn
	oitem.FRectIsWeight		= weightYn
	oitem.FRectRackcode		= itemrackcode
	oitem.FRectCate_Large   = cdl
	oitem.FRectCate_Mid     = cdm
	oitem.FRectCate_Small   = cds
	oitem.FRectSortDiv		= sortDiv
	oitem.FRectSortDiv2		= sortDiv2
	oitem.FRectRegUserID	= vRegUserID
	oitem.FRectRegDate1		= vDate1
	oitem.FRectRegDate2		= vDate2
	oitem.FRectIsReg		= vIsReg
	oitem.FRectSellcash1	= sellcash1
	oitem.FRectSellcash2	= sellcash2
	oitem.FRectSitename     = sitename
	oitem.FRectDanjongyn    = danjongyn

    if (sitename<>"") then
        oitem.GetOverSeasTargetItemListCommon
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('����Ʈ�� �����ϼ���');"
		response.write "</script>"
    end if
%>

<script type='text/javascript'>

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function chgSort(srt,gb){
	if(gb == "2"){
		document.frm.sortDiv2.value= srt;
	}else{
		document.frm.sortDiv.value= srt;
	}
	document.frm.submit();
}

function chgReg(reg){
	document.frm.isreg.value= reg;
	document.frm.submit();
}

function PopItemContent(iitemid){
	var popwin = window.open('/admin/itemmaster/overseas/popItemContent.asp?itemid=' + iitemid +'&sitename=<%=sitename%>','itemWeightEdit','width=1280,height=960,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function num_check(gb){
	if(gb == "1"){
		if(isNaN(document.frm.sellcash1.value) == true)
		{
			alert("���ڸ� �Է����ּ���.");
			document.frm.sellcash1.value = "";
			document.frm.sellcash1.focus();
		}
	}else{
		if(isNaN(document.frm.sellcash2.value) == true)
		{
			alert("���ڸ� �Է����ּ���.");
			document.frm.sellcash2.value = "";
			document.frm.sellcash2.focus();
		}
	}
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get>
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" >
<input type="hidden" name="reload" value="ON">
<input type="hidden" name="sortDiv" value="<%=sortDiv%>">
<input type="hidden" name="sortDiv2" value="<%=sortDiv2%>">
<input type="hidden" name="isreg" value="<%=vIsReg%>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �귣�� : <%	drawSelectBoxDesignerWithName "makerid", makerid %>
		&nbsp;&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<p>
		* ���ڵ� :
		<input type="text" class="text" name="itemrackcode" value="<%= itemrackcode %>" size="12" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		&nbsp;&nbsp;		
		* ��ǰ�ڵ� :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(��ǥ�� �����Է°���)
		&nbsp;&nbsp;
		* ��ǰ�� :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick='NextPage("");'>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		* �Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>
		&nbsp;&nbsp;
     	* ���:<% drawSelectBoxUsingYN "usingyn", usingyn %>
		&nbsp;&nbsp;
     	* ����:<% drawSelectBoxLimitYN "limityn", limityn %>
		&nbsp;&nbsp;
		* ����: <% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
		&nbsp;&nbsp;
     	* �ŷ�����:<% drawSelectBoxMWU "mwdiv", mwdiv %>
		&nbsp;&nbsp;
     	* �ؿܹ��
		<select class="select" name="overSeaYn">
		<option value="">��ü</option>
		<option value="Y" <% if overSeaYn="Y" then Response.write "selected"%>>���</option>
		<option value="N" <% if overSeaYn="N" then Response.write "selected"%>>����</option>
		</select>
		&nbsp;&nbsp;
     	* ���Կ���
		<select class="select" name="weightYn">
		<option value="">��ü</option>
		<option value="Y" <% if weightYn="Y" then Response.write "selected"%>>���</option>
		<option value="N" <% if weightYn="N" then Response.write "selected"%>>����</option>
		</select>
     	<br>
     	* �ǸŰ� :
     	<input type="text" class="text" name="sellcash1" value="<%=sellcash1%>" size="10" onkeyUp="num_check('1')">
     	~<input type="text" class="text" name="sellcash2" value="<%=sellcash2%>" size="10" onkeyUp="num_check('2')">
		&nbsp;&nbsp;
		* ����������Ʈ��:
		<input type="text" name="date1" size="10" maxlength=10 readonly value="<%= vDate1 %>">
		<a href="javascript:calendarOpen(frm.date1);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		&nbsp;~&nbsp;
		<input type="text" name="date2" size="10" maxlength=10 readonly value="<%= vDate2 %>">
		<a href="javascript:calendarOpen(frm.date2);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		&nbsp;&nbsp;
     	* �����
     	<select class="select" name="reguserid">
     	<option value="">��ü</option>
     	<option value="gkclzh" <% if vRegUserID="gkclzh" then Response.write "selected"%>>�ڼ���(gkclzh)</option>
     	<option value="grim0307" <% if vRegUserID="grim0307" then Response.write "selected"%>>���׸�(grim0307)</option>
     	<option value="alsdud001919" <% if vRegUserID="alsdud001919" then Response.write "selected"%>>���ο�(alsdud001919)</option>
     	<option value="">------------------</option>
     	<%
     		Dim vQuery

			vQuery = "SELECT" & vbcrlf
			vQuery = vQuery & " userid, part_sn, username, case part_sn when '11' then 'MD' when '14' then 'MKT' end as part" & vbcrlf
			vQuery = vQuery & " FROM [db_partner].[dbo].[tbl_user_tenbyten]" & vbcrlf
			vQuery = vQuery & " WHERE part_sn IN(11,14) AND isusing = 1" & vbcrlf

			' ��翹���� ó��	' 2018.10.16 �ѿ��
			vQuery = vQuery & " and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
			vQuery = vQuery & " AND posit_sn = '12'" & vbcrlf
			vQuery = vQuery & " ORDER BY part_sn ASC, username ASC" & vbcrlf

			'response.write vQuery & "<br>"
     		rsget.Open vQuery,dbget,1
     		Do Until rsget.Eof
				Response.Write "<option value=""" & rsget("userid") & """ "
				If vRegUserID = rsget("userid") Then
					Response.Write " selected"
				End If
				Response.Write ">" & rsget("part") & " - " & rsget("username") & "</option>"
			rsget.MoveNext
			Loop
			rsget.close()
     	%>
     	</select>
		<br>
		<b><font color="blue">
	    * ����Ʈ : <% drawSelectboxMultiSiteSitename "sitename", sitename, " onchange='NextPage("""");'" %>
	    </font></b>	
	</td>
</tr>
</form>
</table>

<br>
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    	�� �̸Ŵ��� �űԵ�ϸ� �����մϴ�. ������ [ON]�ؿܻ�ǰ����>>�ؿ��ǸŻ�ǰ ���� �ϼ���.
    </td>
    <td align="right">
    </td>
</tr>
</table>
<!-- ǥ �߰��� ��-->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>
				�˻���� : <b><%= oitem.FTotalCount%></b>
				&nbsp;
				������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
			</td>
			<td align="right">
				��ǰ��Ͽ��� :
				<select name="reg" class="select" onchange="chgReg(this.value)">
					<option value="all" <%= CHKIIF(vIsReg="all","selected","") %>>��ü����</option>
					<option value="x" <%= CHKIIF(vIsReg="x","selected","") %>>�̵�ϸ�</option>
					<option value="o" <%= CHKIIF(vIsReg="o","selected","") %>>��ϸ�</option>
				</select>
				&nbsp;&nbsp;&nbsp;
				���Ĺ�� :
				1���� <select name="sort" class="select" onchange="chgSort(this.value,'1')">
					<option value="" <% if sortDiv="" then Response.Write "selected" %>>-����-</option>
					<option value="new" <% if sortDiv="new" then Response.Write "selected" %>>�Ż�ǰ��</option>
					<option value="best" <% if sortDiv="best" then Response.Write "selected" %>>�α��ǰ��</option>
					<option value="min" <% if sortDiv="min" then Response.Write "selected" %>>�������ݼ�</option>
					<option value="hi" <% if sortDiv="hi" then Response.Write "selected" %>>�������ݼ�</option>
					<option value="hs" <% if sortDiv="hs" then Response.Write "selected" %>>������������</option>
					<!--<option value="weight" <% if sortDiv="weight" then Response.Write "selected" %>>��ǰ���Լ�</option>//-->
				</select>
				2���� <select name="sort2" class="select" onchange="chgSort(this.value,'2')">
					<option value="weightup" <% if sortDiv2="weightup" then Response.Write "selected" %>>��ǰ���Գ�����</option>
					<option value="weightdown" <% if sortDiv2="weightdown" then Response.Write "selected" %>>��ǰ���Գ�����</option>
				</select>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">No.</td>
	<td width=50> �̹���</td>
	<td width="100">�귣��ID</td>
	<td> ��ǰ��</td>
	<td width="60">�ǸŰ�</td>
	<td width="60">���԰�</td>
	<td width="30">���<br>����</td>
	<td width="30">�Ǹ�<br>����</td>
	<td width="30">���<br>����</td>
	<td width="30">����<br>����</td>
	<td width="50">����<br>����</td>
	<td width="40">�ؿ�<br>����</td>
	<td width="60">��ǰ<br>����</td>
	<td width="100">�����</td>	
	<td width="100">���</td>
</tr>

<% if oitem.FresultCount > 0 then %>
	<% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF" align="center">
		<td>
			<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="�̸�����">
			<%= oitem.FItemList(i).Fitemid %></a>
			</td>
		<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
		<td align="left"><%= oitem.FItemList(i).Fmakerid %></td>
		<td align="left"><% =oitem.FItemList(i).Fitemname %></td>
		<td align="right">
		<%
			Response.Write "" & FormatNumber(oitem.FItemList(i).Forgprice,0) & ""
			'���ΰ�
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>(��)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
			end if
			'������
			if oitem.FItemList(i).FitemCouponYn="Y" then
				Select Case oitem.FItemList(i).FitemCouponType
					Case "1"
						'Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),0) & "</font>"
					Case "2"
						'Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,0) & "</font>"
				end Select
			end if
		%>
		</td>
		<td align="center"><%= FormatNumber(oitem.FItemList(i).Fbuycash,0) %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
		<td align="center">
			<%= fnColor(oitem.FItemList(i).Fdanjongyn,"dj") %>
		</td>
		<td align="center"><%= fnColor(oitem.FItemList(i).FdeliverOverseas,"yn") %></td>
		<td align="center"><%= FormatNumber(oitem.FItemList(i).FitemWeight,0) %>g</td>
	    <td align="center"><%= oitem.FItemList(i).FRegUserID %></td>
	    <td>
	    	<% If oitem.FItemList(i).fsitename<>"" Then %>
	    		<input type="button" onClick="PopItemContent( '<%= oitem.FItemList(i).Fitemid %>');" value="����" class="button">
	    		<br>
	    		<b>��ǰ��ϿϷ�</b>
	    	<% Else %>
	    		<input type="button" onClick="PopItemContent('<%= oitem.FItemList(i).Fitemid %>');" value="��ǰ���" class="button">
	    		<br>
	    		<font color="red">��ǰ�̵��</font>
	    	<% End If %>
	    </td>
	</tr>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="25" align="center">
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
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>

</table>


<% set oitem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->