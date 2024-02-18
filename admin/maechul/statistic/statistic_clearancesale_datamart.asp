<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : Ŭ����� ���� ���
'	History		: 2016.04.27 �ѿ�� ����
'#############################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/lib/classes/maechul/statistic/ClearanceSaleCls_statistic.asp" -->

<%
Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vSorting, vCateL, vCateM
dim vCateS, vIsBanPum, vPurchasetype, v6Ago, sellchnl, inc3pl, mwdiv,chkShowGubun, dispCate,vBrandID ,itemid
dim page
	vSiteName 	= request("sitename")
	vDateGijun	= NullFillWith(request("date_gijun"),"beasongdate")
	vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",0,now())))
	vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",0,now())))
	vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",0,now())))
	vEYear		= NullFillWith(request("eyear"),Year(now))
	vEMonth		= NullFillWith(request("emonth"),Month(now))
	vEDay		= NullFillWith(request("eday"),Day(now))
	vSorting	= NullFillWith(request("sorting"),"stock")
	vBrandID	= NullFillWith(request("ebrand"),"")
	vCateL		= NullFillWith(request("cdl"),"")
	vCateM		= NullFillWith(request("cdm"),"")
	vCateS		= NullFillWith(request("cds"),"")
	dispCate = requestCheckvar(request("disp"),16)
	itemid      = requestCheckvar(request("itemid"),255)
	vIsBanPum	= NullFillWith(request("isBanpum"),"all")
	vPurchasetype = request("purchasetype")
	v6Ago		= NullFillWith(request("is6ago"),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	mwdiv       = NullFillWith(request("mwdiv"),"")
	inc3pl = request("inc3pl")
	page =requestCheckVar(request("page"),4)
	chkShowGubun = request("chkShowGubun")

v6MonthDate	= DateAdd("m",-6,now())
if page = "" then page = 1

if itemid<>"" then
	dim iA ,arrTemp,arrItemid
	itemid = replace(itemid,",",chr(10))
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

Set cStatistic = New cStaticclearancesale
	cStatistic.FRectSort = vSorting
	cStatistic.FRectCateL = vCateL
	cStatistic.FRectCateM = vCateM
	cStatistic.FRectCateS = vCateS
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectPurchasetype = vPurchasetype
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	'cStatistic.FRect6MonthAgo = v6Ago 
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectMakerid = vBrandID
	cStatistic.FRectInc3pl = inc3pl  ''2014/01/15 �߰�
	cStatistic.FRectDispCate = dispCate
	cStatistic.FRectItemid   = itemid 
	cStatistic.FRectChkShowGubun = chkShowGubun	
	cStatistic.FPageSize = 100
	cStatistic.FCurrPage = page
	cStatistic.fclearancesale_Statistic()

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function searchSubmit(page){
	if ((CheckDateValid(frm.syear.value, frm.smonth.value, frm.sday.value) == true) && (CheckDateValid(frm.eyear.value, frm.emonth.value, frm.eday.value) == true)) {
		$("#btnSubmit").prop("disabled", true);
		frm.page.value=page;
		frm.submit();
	}
}

function pop_stock(itemgubun, itemid, itemoption){
	var pop_stock = window.open('/admin/stock/itemcurrentstock.asp??itemgubun='+itemgubun+'&itemid='+itemid+'&itemoption='+itemoption+'&menupos=709','addreg','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_stock.focus();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="<%= page %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" >
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻�</td>
	<td align="left">
		<table class="a" cellpadding="3" border="0">
		<tr>
			<td height="25" colspan="4">
				 �Ⱓ:
				<select name="date_gijun" class="select">
					<option value="regdate" <%=CHKIIF(vDateGijun="regdate","selected","")%>>�ֹ���</option>
					<option value="ipkumdate" <%=CHKIIF(vDateGijun="ipkumdate","selected","")%>>������</option>
					<option value="beasongdate" <%=CHKIIF(vDateGijun="beasongdate","selected","")%>>�����</option>
				</select>
				<% DrawDateBoxdynamic vSYear,"syear",vEYear,"eyear",vSMonth,"smonth",vEMonth,"emonth",vSDay,"sday",vEDay,"eday" %>
				&nbsp;
				����Ʈ:<% Call Drawsitename("sitename", vSiteName)%>
				&nbsp;ä��:<% drawSellChannelComboBox "sellchnl",sellchnl %>
				&nbsp;<b>����ó:</b> <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
				&nbsp;���Ա���:<% Call DrawBrandMWUPCombo("mwdiv",mwdiv) %>
				&nbsp;
				�ֹ�����:
				<select name="isBanpum" class="select">
					<option value="all" <%=CHKIIF(vIsBanPum="all","selected","")%>>��ǰ����</option>
					<option value="<>" <%=CHKIIF(vIsBanPum="<>","selected","")%>>��ǰ����</option>
					<option value="=" <%=CHKIIF(vIsBanPum="=","selected","")%>>��ǰ�Ǹ�</option>
				</select>
			</td>
		</tr>
		<tr>
			<td colspan="4">
				<Br>
				<!-- #include virtual="/common/module/categoryselectbox.asp"-->
				&nbsp;����ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
				<br>
				�귣��:<input type="text" class="text" name="ebrand" value="<%=vBrandID%>" size="20"> <input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'ebrand');">
				&nbsp;
				��ǰ�ڵ�:<textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
				&nbsp;��������:
				<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
				<!--<input type="checkbox" name="chkShowGubun" value="Y" <% if (chkShowGubun = "Y") then %>checked<% end if %> > ä�α���,���Ա��� ǥ��-->
			</td>
		</tr>
	    </table>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" id="btnSubmit" class="button_s" value="�˻�" onClick="searchSubmit('1');">
	</td>
</tr>
</table>
<!-- �˻� �� -->

<br> 
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25" align="left">
		�˻���� : <b><%= cStatistic.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= cStatistic.FTotalPage %></b>
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25" align="right">
		<input type="radio" name="sorting" value="itemno" <%=CHKIIF(vSorting="itemno","checked","")%>>������
		<input type="radio" name="sorting" value="itemcost" <%=CHKIIF(vSorting="itemcost","checked","")%>>�����
		<input type="radio" name="sorting" value="stock" <%=CHKIIF(vSorting="stock","checked","")%>>����
	</td>
</tr>
</form>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan=5></td>
	<td>A</td>
	<td>B</td>
	<td>C</td>
	<td>D</td>
	<td>E</td>
	<td>F</td>
	<td>G</td>
	<td>H</td>
	<td>I</td>
	<td></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=50>�̹���</td>
	<td width=60>��ǰ�ڵ�<br>�ɼ��ڵ�</td>	
	<td>��ǰ��<Br>�ɼǸ�</td>
	<td>�귣��ID</td>
	<td width=70>Ŭ�����<Br>�����</td>
	<td width=70>�ǸŰ�</td>
	<td width=70>���԰�</td>
	<td width=60>�Ǹż���</td>
	<td width=70>�Ѹ���</td>
	<td width=70>�Ѹ���</td>
	<td width=60>���ͷ�</td>
	<td width=60>Ŭ�����<br>���������</td>
	<td width=60>
		������
		<br>G-C
	</td>
	<td width=60>
		������
		<br>C/G
	</td>
	<td width=40>���</td>
</tr>
<% if cStatistic.FresultCount>0 then %>
	<% for i=0 to cStatistic.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td>
			<img src="<%= cStatistic.FItemList(i).FimageSmall %>" width=50 height=50>
		</td>
		<td>
			<%= cStatistic.FItemList(i).fitemid %>

			<% if cStatistic.FItemList(i).fitemoption<>"" then %>
				<br><%= cStatistic.FItemList(i).fitemoption %>
			<% end if %>
		</td>
		<td align="left">
			<%= cStatistic.FItemList(i).fitemname %>

			<% if cStatistic.FItemList(i).foptionname<>"" then %>
				<br><%= cStatistic.FItemList(i).foptionname %>
			<% end if %>
		</td>
		<td align="center">
			<%= cStatistic.FItemList(i).fmakerid %>
		</td>
		<td align="center">
			<%= cStatistic.FItemList(i).fregdate %>
		</td>
		<td align="right">
			<%= CurrFormat(cStatistic.FItemList(i).fsellcash) %>
		</td>
		<td align="right">
			<%= CurrFormat(cStatistic.FItemList(i).fbuycash) %>
		</td>
		<td align="right">
			<%= CurrFormat(cStatistic.FItemList(i).fitemno) %>
		</td>
		<td align="right">
			<%= CurrFormat(cStatistic.FItemList(i).fitemcostsum) %>
		</td>
		<td align="right">
			<%= CurrFormat(cStatistic.FItemList(i).fbuycashsum) %>
		</td>
		<td align="right">
			<% if cStatistic.FItemList(i).fitemcostsum-cStatistic.FItemList(i).fbuycashsum <> 0 then %>
				<%= round((( (cStatistic.FItemList(i).fitemcostsum-cStatistic.FItemList(i).fbuycashsum) /cStatistic.FItemList(i).fitemcostsum) *100),2) %>%
			<% else %>
				0%
			<% end if %>
		</td>
		<td align="right">
			<%= CurrFormat(cStatistic.FItemList(i).favailsysstock) %>
		</td>
		<td align="right">
			<%= CurrFormat(cStatistic.FItemList(i).favailsysstock-cStatistic.FItemList(i).fitemno) %>
		</td>
		<td align="right">
			<% if cStatistic.FItemList(i).fitemno<>0 and cStatistic.FItemList(i).favailsysstock<>0 then %>
				<%= round(((cStatistic.FItemList(i).fitemno/cStatistic.FItemList(i).favailsysstock) *100),2) %>%
			<% else %>
				0%
			<% end if %>
		</td>
		<td>
			<input type="button" onclick="pop_stock('10','<%= cStatistic.FItemList(i).fitemid %>','<%= cStatistic.FItemList(i).fitemoption %>'); return fasle;" value="���" class="button">
		</td>
	</tr>   
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if cStatistic.HasPreScroll then %>
				<span class="list_link"><a href="" onclick="searchSubmit('<%= cStatistic.StartScrollPage-1 %>'); return false;">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + cStatistic.StartScrollPage to cStatistic.StartScrollPage + cStatistic.FScrollCount - 1 %>
				<% if (i > cStatistic.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(cStatistic.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="" onclick="searchSubmit('<%= i %>'); return false;" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if cStatistic.HasNextScroll then %>
				<span class="list_link"><a href="" onclick="searchSubmit('<%= i %>'); return false;">[next]</a></span>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
Set cStatistic = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->