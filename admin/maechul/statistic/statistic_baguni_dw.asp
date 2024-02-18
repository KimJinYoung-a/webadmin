<%@ language=vbscript %>
<% option explicit

	'��ũ��Ʈ Ÿ�Ӿƿ� �ð� ���� (�⺻ 90��)
	Server.ScriptTimeout = 180
%>
<%
'###########################################################
' Description : ��ٱ�����ȯ����
' History : 2016.03.16 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_dw.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->

<%
Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vSorting, vCateL, vCateM, vCateS
dim vIsBanPum, vPurchasetype, v6Ago, sellchnl, inc3pl, mwdiv, dispCate,vBrandID, chkImg ,itemid, iCurrPage,iPageSize,iTotalPage,iTotCnt
dim syyyy, smm, sdd, eyyyy, emm, edd, reloading, date_gijun
	syyyy		= NullFillWith(request("syyyy"),Year(DateAdd("d",0,now())))
	smm		= NullFillWith(request("smm"),Month(DateAdd("d",0,now())))
	sdd		= NullFillWith(request("sdd"),Day(DateAdd("d",0,now())))
	eyyyy		= NullFillWith(request("eyyyy"),Year(now))
	emm		= NullFillWith(request("emm"),Month(now))
	edd		= NullFillWith(request("edd"),Day(now))
	iPageSize = 100
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= request("sitename")
	vDateGijun	= NullFillWith(request("date_gijun"),"regdate")
	vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",0,now())))
	vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",0,now())))
	vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",0,now())))
	vEYear		= NullFillWith(request("eyear"),Year(now))
	vEMonth		= NullFillWith(request("emonth"),Month(now))
	vEDay		= NullFillWith(request("eday"),Day(now))
	vSorting	= NullFillWith(request("sorting"),"itemsellcntD")
	vBrandID	= NullFillWith(request("ebrand"),"")
	vCateL		= NullFillWith(request("cdl"),"")
	vCateM		= NullFillWith(request("cdm"),"")
	vCateS		= NullFillWith(request("cds"),"")
	dispCate = requestCheckvar(request("disp"),16)
	itemid      = requestCheckvar(request("itemid"),255)
	chkImg		= requestCheckvar(request("chkImg"),1)
	vIsBanPum	= NullFillWith(request("isBanpum"),"all")
	vPurchasetype = request("purchasetype")
	v6Ago		= NullFillWith(request("is6ago"),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	mwdiv       = NullFillWith(request("mwdiv"),"")
	inc3pl = request("inc3pl")
	iCurrPage =requestCheckVar(request("iC"),4)
	reloading    = requestCheckVar(request("reloading"),2)

if iCurrPage = "" then iCurrPage = 1
if chkImg ="" then chkImg = 0	
if reloading="" and vSiteName="" then vSiteName="10x10"

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

Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectSort = vSorting
	cStatistic.FRectCateL = vCateL
	cStatistic.FRectCateM = vCateM
	cStatistic.FRectCateS = vCateS
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectPurchasetype = vPurchasetype
	cStatistic.FRectDateGijun = vDateGijun
	'cStatistic.FRectmaechulStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	'cStatistic.FRectmaechulEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectStartdate = syyyy & "-" & TwoNumber(smm) & "-" & TwoNumber(sdd)
	cStatistic.FRectEndDate = eyyyy & "-" & TwoNumber(emm) & "-" & TwoNumber(edd)
	cStatistic.FRectSiteName = vSiteName
	'cStatistic.FRect6MonthAgo = v6Ago 
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectMakerid = vBrandID
	cStatistic.FRectInc3pl = inc3pl  ''2014/01/15 �߰�
	cStatistic.FRectDispCate = dispCate
	cStatistic.FRectItemid   = itemid 
	cStatistic.FPageSize = iPageSize
	cStatistic.FCurrPage = iCurrPage
	cStatistic.fStatistic_baguni()

iTotCnt = cStatistic.FResultCount	
iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">

function searchSubmit(){
    document.frm.target = "_self"; 
    document.frm.action = "statistic_baguni_dw.asp";  
	document.frm.iC.value="";
	frm.submit();
}

function jstrSort(vsorting){
	var tmpSorting = document.getElementById("img"+vsorting)

	if (-1 < tmpSorting.src.indexOf("_alpha")){
		frm.sorting.value= vsorting+"D";
	}else if (-1 < tmpSorting.src.indexOf("_bot")){
		frm.sorting.value= vsorting+"A";
	}else{
		frm.sorting.value= vsorting+"D";
	}
	searchSubmit();
}

 function jsexceldown(){ 
  
    var icurrpage = $('#selDCnt').val(); 
    document.frm.target =  "XLdown"; 
    document.frm.iC.value =icurrpage;
    document.frm.action = "statistic_baguni_dw_xls.asp";  
    //alert("a");
	document.frm.submit(); 
	
}
</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="reloading" value="on">
<input type="hidden" name="sorting" value="<%= vsorting %>">
<input type="hidden" name="iC" value=""> 

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" >
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>" rowspan=3>�˻�</td>
	<td align="left">
		<select name="date_gijun" class="select">
			<option value="regdate" <%=CHKIIF(vDateGijun="regdate","selected","")%>>�ֹ���</option>
			<option value="ipkumdate" <%=CHKIIF(vDateGijun="ipkumdate","selected","")%>>������</option>
			<option value="beasongdate" <%=CHKIIF(vDateGijun="beasongdate","selected","")%>>�����</option>
			<option value="jfixeddt" <%=CHKIIF(vDateGijun="jfixeddt","selected","")%>>����Ȯ����</option>
		</select>
		<% DrawDateBoxdynamic syyyy,"syyyy",eyyyy,"eyyyy",smm,"smm",emm,"emm",sdd,"sdd",edd,"edd" %>
		&nbsp;&nbsp;����ó : <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
		&nbsp;&nbsp;�������� : 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
		&nbsp;&nbsp;���Ա��� : <% Call DrawBrandMWUCombo("mwdiv",mwdiv) %>
		&nbsp;&nbsp;�귣�� : <input type="text" class="text" name="ebrand" value="<%=vBrandID%>" size="20"> <input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'ebrand');">
		&nbsp;&nbsp;<input type="checkbox" name="chkImg" value="1" <%if chkImg = 1 then%>checked<%end if%>>��ǰ�̹��� ����
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan=3>
		<input type="button" id="btnSubmit" class="button_s" value="�˻�" onClick="searchSubmit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		��ǰ�ڵ� : <textarea rows="2" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		&nbsp;&nbsp;����ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<% 'DrawDateBoxdynamic vsyear,"syear",vEYear,"eyear",vsmonth,"smonth",vemonth,"emonth",vsday,"sday",veday,"eday" %>
		����Ʈ : <% Call Drawsitename("sitename", vSiteName)%>
		&nbsp;&nbsp;ä�� : <% drawSellChannelComboBox "sellchnl",sellchnl %>
		&nbsp;&nbsp;�ֹ����� : 
		<select name="isBanpum" class="select">
			<option value="all" <%=CHKIIF(vIsBanPum="all","selected","")%>>��ǰ����</option>
			<option value="<>" <%=CHKIIF(vIsBanPum="<>","selected","")%>>��ǰ����</option>
			<option value="=" <%=CHKIIF(vIsBanPum="=","selected","")%>>��ǰ�Ǹ�</option>
		</select>
	</td>
</tr>
</table>
<!-- �˻� �� -->
<br> 

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		�� <font color="red">����</font>�� ���� �ɸ��� ������ �Դϴ�. ������ ���� ������ ���ð� ��ٷ� �ּ���.
		<Br>
		�Ǹŵ� ��ǰ�� ��ٱ��ϸ� ���ļ� �Ǹ� �Ǿ�����, ��ñ��� �Ǿ������� �˼� �����ϴ�. �帧�� �Ǵ��ϴ� ������ ����ϼ���.
	</td>
	<td align="right">
		<!--����: <input type="radio" name="sorting" value="itemsellcnt" <%=CHKIIF(vSorting="itemsellcnt","checked","")%>>�Ǹ���ȯ����
		<input type="radio" name="sorting" value="itembagunicnt" <%=CHKIIF(vSorting="itembagunicnt","checked","")%>>��ٱ��ϰǼ���
		<input type="radio" name="sorting" value="itemsellconversrate" <%=CHKIIF(vSorting="itemsellconversrate","checked","")%>>�Ǹ���ȯ����-->
		<span style="width:100%;text-align:right;">
�����ٿ�:
	<% dim iDownCnt, imaxDCnt, iminDCnt 
 	%> 
	<select name="selDCnt" id="selDCnt" class="select">
	    <option value="0">--������ ����--</option>
	    <%
	    if iTotCnt >0 then
	        iDownCnt =  Int(iTotCnt/5000)+1 
	        imaxDCnt = 0
	    for i=1 to iDownCnt 
	        iminDCnt = imaxDCnt + 1
	        if iDownCnt = 1 then
	            imaxDCnt = iTotCnt
	        else    
	            imaxDCnt = 5000*i
	        end if    
	    %>
	    <option value="<%=i%>"><%=iminDCnt%>~<%=imaxDCnt%></option>
	    <%next%>
	    <%end if%> 
	</select>
    <a href="javascript:jsexceldown();"><image src="/images/btn_excel.gif" border="0" align="absmiddle"></a> 
</span>
	</td>
</tr>
</table>
<!-- �׼� �� -->
</form>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�˻���� : <b><%=iTotCnt%></b>
		&nbsp;
		������ : <b><%= iCurrPage %> / <%=iTotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td></td>

	<% IF chkImg = 1 then %>
		<td></td>
	<% END IF %>

	<td></td>
    <td></td>
    <td></td>
	<td>A</td>
	<td>B</td>
    <td>C</td>
    <td>D</td>
    <td>E</td>
    <td>F</td>
    <td>G</td>
    <td>H</td>
    <td>I</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ǰ�ڵ�</td>

	<% IF chkImg = 1 then %>
		<td>�̹���</td>
	<% END IF %>

	<td>�귣��</td>
    <td>ī�װ�</td>
    <td>��ǰ��</td>
	<td onClick="jstrSort('sellcash'); return false;" style="cursor:hand;">
		�ǸŰ�
		<img src="/images/list_lineup<%=CHKIIF(vSorting="sellcashD","_bot","_top")%><%=CHKIIF(instr(vSorting,"sellcash")>0,"_on","")%>.png" id="imgsellcash">
	</td>
	<td onClick="jstrSort('buycash'); return false;" style="cursor:hand;">
		���԰�
		<img src="/images/list_lineup<%=CHKIIF(vSorting="buycashD","_bot","_top")%><%=CHKIIF(instr(vSorting,"buycash")>0,"_on","")%>.png" id="imgbuycash">
	</td>
	<td >-</td>
	<!--
    <td onClick="jstrSort('totbagunicnt'); return false;" style="cursor:hand;">
    	�Ѵ�����
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="totbagunicntD","_bot","_top")%><%=CHKIIF(instr(vSorting,"totbagunicnt")>0,"_on","")%>.png" id="imgtotbagunicnt">
    	<Br>D+E
    </td>
	-->
    <td onClick="jstrSort('itemsellcnt'); return false;" style="cursor:hand;">
    	�Ǹ���ȯ��
    	<br>(�ǸŰǼ�)
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itemsellcntD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemsellcnt")>0,"_on","")%>.png" id="imgitemsellcnt">
    </td>
    <td onClick="jstrSort('itembagunicnt'); return false;" >
    	��ٱ���
    	<br>�����Ǽ�
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itembagunicntD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itembagunicnt")>0,"_on","")%>.png" id="imgitembagunicnt">
    </td>
    <td onClick="jstrSort('itemsellconversrate'); return false;" >
    	�Ǹ���ȯ��
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itemsellconversrateD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemsellconversrate")>0,"_on","")%>.png" id="imgitemsellconversrate">
    </td>
    <td onClick="jstrSort('itemsellsum'); return false;" style="cursor:hand;">
    	��ü����
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itemsellsumD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemsellsum")>0,"_on","")%>.png" id="imgitemsellsum">
    </td>
    <td onClick="jstrSort('totfavcount'); return false;" style="cursor:hand;">
    	�����ü�
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="totfavcountD","_bot","_top")%><%=CHKIIF(instr(vSorting,"totfavcount")>0,"_on","")%>.png" id="imgtotfavcount">
    </td>
    <td onClick="jstrSort('recentfavcount'); return false;" style="cursor:hand;">
    	�ֱ����ü�1��
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="recentfavcountD","_bot","_top")%><%=CHKIIF(instr(vSorting,"recentfavcount")>0,"_on","")%>.png" id="imgrecentfavcount">
    </td>
</tr>
<% if cStatistic.FTotalCount>0 then %>
	<%
	dim tot_totbagunicnt, tot_itemsellcnt, tot_itembagunicnt, tot_itemsellconversrate, tot_itemsellsum
	dim tot_favcount, tot_recentfavcount

	For i = 0 To cStatistic.FTotalCount -1

	tot_totbagunicnt = tot_totbagunicnt + cStatistic.FList(i).ftotbagunicnt
	tot_itemsellcnt = tot_itemsellcnt + cStatistic.FList(i).fitemsellcnt
	tot_itembagunicnt = tot_itembagunicnt + cStatistic.FList(i).fitembagunicnt
	tot_itemsellconversrate = tot_itemsellconversrate + cStatistic.FList(i).fitemsellconversrate
	tot_itemsellsum = tot_itemsellsum + cStatistic.FList(i).fitemsellsum
	tot_favcount = tot_favcount + cStatistic.FList(i).ffavcount
	tot_recentfavcount = tot_recentfavcount + cStatistic.FList(i).frecentfavcount
	%>
	<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFFF';>
		<td><%= cStatistic.FList(i).FitemID %></td>

		<% IF chkImg = 1 then %>
			<td><img src="<%= cStatistic.FList(i).FSmallImage %>" width="50" height="50" border="0"></td>
		<% END IF %>

		<td><%= cStatistic.FList(i).FMakerID %></td>
		<td><%= cStatistic.FList(i).fcatename %></td>
		<td align="left"><%= cStatistic.FList(i).fitemname %></td>
		<td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).fsellcash) %></td>
		<td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).fbuycash) %></td>
		<td align="right"></td>
		<!--td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).ftotbagunicnt) %></td-->
		<td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).fitemsellcnt) %></td>
		<td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).fitembagunicnt) %></td>
		<td align="right"><%= round(NullOrCurrFormat(cStatistic.FList(i).fitemsellconversrate),1) %>%</td>
		<td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).fitemsellsum) %></td>
		<td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).ffavcount) %></td>
		<td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).frecentfavcount) %></td>
	</tr>
	<% Next %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="<% IF chkImg = 1 then %>7<% else %>6<% end if %>">�Ѱ�</td>
		<td align="right"></td>
		<!--td align="right"><%= NullOrCurrFormat(tot_totbagunicnt) %></td-->
		<td align="right"><%= NullOrCurrFormat(tot_itemsellcnt) %></td>
		<td align="right"><%= NullOrCurrFormat(tot_itembagunicnt) %></td>
		<td align="right"><%= round(NullOrCurrFormat(tot_itemsellconversrate/cStatistic.FTotalCount),1) %>%</td>
		<td align="right"><%= NullOrCurrFormat(tot_itemsellsum) %></td>
		<td align="right"><%= NullOrCurrFormat(tot_favcount) %></td>
		<td align="right"><%= NullOrCurrFormat(tot_recentfavcount) %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="25">
			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center">�˻������ �����ϴ�.</td>
	</tr>
<% end if %>
</table>

<% Set cStatistic = Nothing %>
<iframe id="XLdown" name="XLdown" src="about:blank" frameborder="0" width="110" height="110"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
