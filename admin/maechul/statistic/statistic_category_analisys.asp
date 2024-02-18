<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ī�װ�������
' History : 2016.01.20 ������ ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_analisys.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp" -->
<%
Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vCateL, vCateM
dim sellchnl, categbn, vCateS, vCateX, vIsBanPum, vBrandID, vCateGubun, vPurchasetype, v6Ago, vParam, vbizsec
Dim mwdiv, inc3pl, vCateMRate,vTot_CateMRate, dispCate, maxDepth, chkChannel, linkcate, linkdispcate, rdsite, showsuply, showordcnt
Dim incStockAvg, vShowDate
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= request("sitename")
	vDateGijun	= NullFillWith(request("date_gijun"),"regdate")
	vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",0,now())))
	vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",0,now())))
	vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",0,now())))
	vEYear		= NullFillWith(request("eyear"),Year(now))
	vEMonth		= NullFillWith(request("emonth"),Month(now))
	vEDay		= NullFillWith(request("eday"),Day(now))
	vCateL		= NullFillWith(request("cdl"),"")
	vCateM		= NullFillWith(request("cdm"),"")
	vCateS		= NullFillWith(request("cds"),"")
	vCateX      = NullFillWith(request("cdx"),"")
	vIsBanPum	= NullFillWith(request("isBanpum"),"all")
	vPurchasetype = request("purchasetype")
	vBrandID	= NullFillWith(request("ebrand"),"")
	v6Ago		= NullFillWith(request("is6ago"),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	vbizsec     = NullFillWith(request("bizsec"),"")
	mwdiv       = NullFillWith(request("mwdiv"),"")
	rdsite		= NullFillWith(request("rdsite"),"")
	categbn     = NullFillWith(request("categbn"),"")
    inc3pl = request("inc3pl")
    dispCate 	= requestCheckvar(request("disp"),16)
    maxDepth    = requestCheckvar(request("selDepth"),1)
    chkChannel  = requestCheckvar(request("chkChl"),1)
	showsuply   = requestCheckvar(request("showsuply"),10)
	showordcnt  = requestCheckvar(request("showordcnt"),10)
	incStockAvg = requestCheckvar(request("incStockAvg"),10)
	vShowDate	= requestCheckvar(request("showdate"),10)

if maxDepth = ""   then maxDepth = 0
if chkChannel = "" then chkChannel = 0
vCateGubun = "L"
If vCateL <> "" and vCateM <> "" and vCateS<>"" Then
	vCateGubun = "X"
ELSEIF vCateL <> "" and vCateM <> "" THEN
    vCateGubun = "S"
ELSEif vCateL <> "" Then
	vCateGubun = "M"
End IF

if (categbn="") then
    categbn="D"
end if
vParam = CurrURL() & "?menupos="&Request("menupos")&"&date_gijun="&vDateGijun&"&syear="&vSYear&"&smonth="&vSMonth&"&sday="&vSDay&"&eyear="&vEYear&"&emonth="&vEMonth&"&eday="&vEDay&"&sitename="&vSiteName&"&isBanpum="&vIsBanPum&"&ebrand="&vBrandID&"&purchasetype="&vPurchasetype&"&is6ago="&v6Ago&"&mwdiv="&mwdiv&"&categbn="&categbn&"&sellchnl="&sellchnl&"&chkChl="&chkChannel&"&rdsite="&rdsite&"&incStockAvg="&incStockAvg

Dim vTot_OrderCnt, vTot_ItemNO, vTot_OrgitemCost, vTot_ItemcostCouponNotApplied, vTot_ItemCost, vTot_BuyCash
Dim vTot_MaechulProfit, vTot_MaechulProfitPer, vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2
dim vTot_upcheJungsan, vTot_avgipgoPrice, vTot_overValueStockPrice

Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectCateL = vCateL
	cStatistic.FRectCateM = vCateM
	cStatistic.FRectCateS = vCateS
	cStatistic.FRectCateGubun = vCateGubun
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectPurchasetype = vPurchasetype
	cStatistic.FRectMakerID = vBrandID
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	'cStatistic.FRect6MonthAgo = v6Ago
	cStatistic.FRectSellChannelDiv = sellchnl
	'cStatistic.FRectChannelDiv = channelDiv
	'cStatistic.FRectBizSectionCd = vbizsec
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectRdsite = rdsite
	cStatistic.FRectInc3pl = inc3pl  ''2014/01/15 �߰�
	cStatistic.FRectCateGbn = categbn
    cStatistic.FRectChkchannel = chkChannel
	cStatistic.FRectIncStockAvgPrc = (incStockAvg<>"") ''true '' ��ո��԰� ���� ��������.
	cStatistic.FRectBySuplyPrice = CHKIIF(showsuply="on",1,0)
    cStatistic.FRectUseOrderCount = CHKIIF(showordcnt="on",1,0)  '' �ֹ��Ǽ� ǥ��
    cStatistic.FRectShowDate = vShowDate

	if (categbn="M") then
	    cStatistic.fStatistic_category()
	else
	    cStatistic.FRectdispCate = dispCate
        cStatistic.FRectmaxDepth = maxdepth
    	cStatistic.fStatistic_DispCategory  ''2013/10/17 �߰�
    end if

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function popCateSellDetail(cdl,cdm,cds,dispcate){
    window.open("/admin/maechul/statistic/statistic_item_analisys.asp?menupos=1726&date_gijun=<%=vDateGijun%>&syear=<%=vSYear%>&smonth=<%=vSMonth%>&sday=<%=vSDay%>&eyear=<%=vEYear%>&emonth=<%=vEMonth%>&eday=<%=vEDay%>&cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&disp="+dispcate,'','');
}

function MonthDiff(d1, d2) {
	d1 = d1.split("-");
	d2 = d2.split("-");

	d1 = new Date(d1[0], d1[1] - 1, d1[2]);
	d2 = new Date(d2[0], d2[1] - 1, d2[2]);

	var d1Y = d1.getFullYear();
	var d2Y = d2.getFullYear();
	var d1M = d1.getMonth();
	var d2M = d2.getMonth();

	return (d2M+12*d2Y)-(d1M+12*d1Y);
}

function searchSubmit()
{
	//if((frm.syear.value == <%=Year(v6MonthDate)%> && frm.smonth.value < <%=Month(v6MonthDate)%>) && (frm.is6ago.checked == false))
	//{
	//	alert("6�������� �����ʹ� 6�������������͸� üũ�ϼž� �����մϴ�.");
	//}
	//else
	//{
		if ((CheckDateValid(frm.syear.value, frm.smonth.value, frm.sday.value) == true) && (CheckDateValid(frm.eyear.value, frm.emonth.value, frm.eday.value) == true)) {
			//2014-09-01 ������ �ּ�ó��
			//if (MonthDiff(frm.syear.value + "-" + frm.smonth.value + "-" + frm.sday.value, frm.eyear.value + "-" + frm.emonth.value + "-" + frm.eday.value) >= 1) {
			//	alert("�ִ� 1���������� �˻��� �����մϴ�.");
			//	return;
			//}
			$("#btnSubmit").prop("disabled", true);
			frm.submit();
		}
	//}
}

function jsChangeDepth(ivalue){
    var dispDepth  = "<%=maxDepth%>";
    var strDisp = 0;

    if(ivalue < dispDepth){
        if (ivalue == 0){
            strDisp = "";
        }else{
         strDisp = "<%=dispCate%>".substring(0,ivalue*3);
        }

        document.all.disp.value =strDisp ;
    }
    searchSubmit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				* �Ⱓ :&nbsp;
				<select name="date_gijun" class="select">
					<option value="regdate" <%=CHKIIF(vDateGijun="regdate","selected","")%>>�ֹ���</option>
					<option value="ipkumdate" <%=CHKIIF(vDateGijun="ipkumdate","selected","")%>>������</option>
					<option value="beasongdate" <%=CHKIIF(vDateGijun="beasongdate","selected","")%>>��ǰ�����</option>
				</select>
				<%
					'### ��
					Response.Write "<select name=""syear"" class=""select"">"
					For i=Year(now) To 2001 Step -1
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### ��
					Response.Write "<select name=""smonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### ��
					Response.Write "<select name=""sday"" class=""select"">"
					For i=1 To 31
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSDay),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;~&nbsp;"

					'#############################

					'### ��
					Response.Write "<select name=""eyear"" class=""select"">"
					For i=Year(now) To 2001 Step -1
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### ��
					Response.Write "<select name=""emonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### ��
					Response.Write "<select name=""eday"" class=""select"">"
					For i=1 To 31
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEDay),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>"


					'### 6��������������check
					'Response.Write "<input type=""checkbox"" name=""is6ago"" value=""o"" "
					'If v6Ago = "o" Then
					'	Response.Write "checked"
					'End If
					'Response.Write ">6��������������"
				%>
				<%
					'### ����Ʈ����
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* ����Ʈ���� : "
					Call Drawsitename("sitename", vSiteName)

					'Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* �⺻ ����μ� : "
                    'Call DrawBizSectionGain("O,T","bizsec", vbizsec,"")
				%>
				&nbsp;
                	* ä�α���
                	<% drawSellChannelComboBox "sellchnl",sellchnl %>
                &nbsp;
                * �ֹ����� :
				<select name="isBanpum" class="select">
					<option value="all" <%=CHKIIF(vIsBanPum="all","selected","")%>>��ǰ����</option>
					<option value="<>" <%=CHKIIF(vIsBanPum="<>","selected","")%>>��ǰ����</option>
					<option value="=" <%=CHKIIF(vIsBanPum="=","selected","")%>>��ǰ�Ǹ�</option>
				</select>
				&nbsp;
				* ���Ա��� :
				<% Call DrawBrandMWUPCombo("mwdiv",mwdiv) %>
				&nbsp;
				�Ǹ�ó��:
				<% Call DrawRdsiteCombo("rdsite",rdsite) %>
				&nbsp;&nbsp;&nbsp;
				<b>* ����ó����</b>
        	    <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
			</td>
		</tr>
		<tr>
		    <td>

				* �귣�� : <input type="text" class="text" name="ebrand" value="<%=vBrandID%>" size="20"> <input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'ebrand');">
				&nbsp;&nbsp;

				&nbsp;&nbsp;
				* �������� : 
				<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
				&nbsp;&nbsp;
				<input type="radio" name="categbn" value="M" <%=CHKIIF(categbn="M","checked","")%> >����ī�װ�
				<input type="radio" name="categbn" value="D" <%=CHKIIF(categbn="D","checked","")%> >����ī�װ�
				<select name="selDepth" class="select"  onChange="jsChangeDepth(this.value);" <%if categbn = "M" then%>disabled<%end if%>>
				    <option value="0" <%if maxDepth ="0" then%>selected<%end if%>>��(1 Depth)</option>
				    <option value="1" <%if maxDepth ="1" then%>selected<%end if%>>��(2 Depth)</option>
				    <option value="2" <%if maxDepth ="2" then%>selected<%end if%>>��(3 Depth)</option>
				    <option value="3" <%if maxDepth ="3" then%>selected<%end if%>>��(4 Depth)</option>
				</select>
				 <%if maxDepth > 0 and categbn = "D" then %>
				<!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
				<% end if%>
				<input type="checkbox" name="chkChl" value="1" <%if chkChannel ="1" then%>checked<%end if%>>ä�λ󼼺���
			    &nbsp;

			    <input type="checkbox" name="showordcnt" value="on" <%= CHKIIF(showordcnt="on","checked","") %> >�ֹ��Ǽ� ǥ��
			    &nbsp;
			    <input type="checkbox" name="showsuply" value="on" <%= CHKIIF(showsuply="on","checked","") %> >���ް��� ǥ��
			    &nbsp;
			    <input type="checkbox" name="incStockAvg" <%=CHKIIF(incStockAvg<>"","checked","")%>>��ո��԰�����
				&nbsp;
			    <input type="checkbox" name="showdate" value="Y" <%=CHKIIF(vShowDate<>"","checked","")%>>����ǥ��
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" id="btnSubmit" class="button_s" value="�˻�" onClick="javascript:searchSubmit();" ></td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<br>
* �ֹ��Ǽ��� ��ǰ/��ȯ���� ���Ե��� ����<br>
* �˻� �Ⱓ�� ������� ����� �������ϴ�. �׷��� �˻� ��ư�� Ŭ���� �� �ƹ� ������ ����δٰ� ���� �˻���ư�� Ŭ������ ������.<br />
* �ִ� 2000�Ǳ����� ��ȸ�˴ϴ�.
<br>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">
	<%if vShowDate ="Y" then%>
	<td align="center">����</td>
	<%end if%>
	<td align="center"><%=CateGubun(vCateGubun)%>ī�װ�</td>
	<%if chkChannel ="1" then%>
	<td align="center">ä��</td>
	<%end if%>
    <%if (showordcnt="on") then%><td>�ֹ��Ǽ�</td><% end if%>
    <td>��ǰ����</td>
    <% if (NOT C_InspectorUser) then %>
    <td>�Һ��ڰ�[��ǰ]</td>
    <td>�ǸŰ�[��ǰ]<br>(��������)</td>
    <td><b>�����Ѿ�[��ǰ]<br>(��ǰ��������)</b></td>
    <%if chkChannel ="1" then%>
    <td>ä��<br>������</td>
    <%end if%>
    <td><b>���ʽ�����<br>����[��ǰ]</b></td>
    <% end if %>
    <td>��޾�</td>
    <td>�����Ѿ�[��ǰ]<% if (NOT C_InspectorUser) then %><br>(��ǰ��������)<% end if %></td>
    <td><b>�������</b></td>
    <td>������</td>
    <td>�������2<br>(��޾ױ���)</td>
    <td>������</td>
    <td>ī�װ���<br>���� ����</td>
	<td align="center">��ü<br>�����</td>
	<td align="center"><b>ȸ�����</b></td>
	<td align="center">���<br>���԰�</td>
	<td align="center">���<br>����</td>
    <td align="center">���</td>
</tr>
<%
For i = 0 To cStatistic.FTotalCount -1
%>
<tr 	bgcolor=<%if chkChannel ="1" then%>"#e3f1fb"<%else%>"#FFFFFF"<%end if%>>
	<%if vShowDate ="Y" then%>
	<td bgcolor="#ffffff" align="center" <%if chkChannel ="1" then%>rowspan="6"<%end if%>><%= cStatistic.FList(i).Fyyyymmdd %></td>
	<%end if%>
	<td  style="padding-left:5px;" bgcolor="#ffffff"<%if chkChannel ="1" then%>rowspan="6"<%end if%>>
		<%= cStatistic.FList(i).FCategoryName %>&nbsp;
		<%  linkcate = ""
			If vCateGubun = "L" Then
				Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FList(i).FCateL&"""><font color=""gray"">[��]</font></a>"
				IF (cStatistic.FList(i).FCateL="999") then
				    '' Response.Write " <a href=""javascript:popCateSellDetail('"&cStatistic.FList(i).FCateL&"','','','')"">(��)</a>"
				end if
				if categbn = "D" then
				    linkcate = "&disp1="&cStatistic.FList(i).FCateL
				else
				    linkcate = "&cdl="&cStatistic.FList(i).FCateL
				end if
			ElseIf vCateGubun = "M" Then
				Response.Write "<a href="""&vParam&"""><font color=""gray"">[��]</font></a>"
				Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FList(i).FCateL&"&cdm="&cStatistic.FList(i).FCateM&"""><font color=""gray"">[��]</font></a>"
				IF (cStatistic.FList(i).FCateM="") then
				    Response.Write " <a href=""javascript:popCateSellDetail('"&cStatistic.FList(i).FCateL&"','','','')"">(��)</a>"
				end if
				if categbn = "D" then
				    linkcate = "&disp1="&cStatistic.FList(i).FCateL&"&disp2="&cStatistic.FList(i).FCateM
			    else
				    linkcate = "&cdl="&cStatistic.FList(i).FCateL&"&cdm="&cStatistic.FList(i).FCateM
			    end if

			ElseIf vCateGubun = "S" Then
				Response.Write "<a href="""&vParam&"""><font color=""gray"">[��]</font></a>"
				Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FList(i).FCateL&"""><font color=""gray"">[��]</font></a>"
				if (categbn="D") then
                Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FList(i).FCateL&"&cdm="&cStatistic.FList(i).FCateM&"&cds="&cStatistic.FList(i).FCateS&"""><font color=""gray"">[��]</font></a>"
                    linkcate = "&disp1="&cStatistic.FList(i).FCateL&"&disp2="&cStatistic.FList(i).FCateM&"&disp3="&cStatistic.FList(i).FCateS
                else
                    linkcate = "&cdl="&cStatistic.FList(i).FCateL&"&cdm="&cStatistic.FList(i).FCateM&"&cds="&cStatistic.FList(i).FCateS
                end if
            ElseIf vCateGubun = "X" Then
				Response.Write "<a href="""&vParam&"""><font color=""gray"">[��]</font></a>"
				Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FList(i).FCateL&"""><font color=""gray"">[��]</font></a>"
                Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FList(i).FCateL&"&cdm="&cStatistic.FList(i).FCateM&"""><font color=""gray"">[��]</font></a>"
                'Response.Write " <a href=""javascript:popCateSellDetail('"&cStatistic.FList(i).FCateL&"','"&cStatistic.FList(i).FCateM&"','"&cStatistic.FList(i).FCateS&"','"&cStatistic.FList(i).FCateX&"')"">(��)</a>"
             End IF
              linkdispcate =  "&disp="&cStatistic.FList(i).FDispCateCode
			if cStatistic.FTotItemCost ="" or cStatistic.FTotItemCost = 0 then
				vCateMRate = 0
			else
			vCateMRate = (cStatistic.FList(i).FItemCost/cStatistic.FTotItemCost)*100
			end if
	' Response.Write " <a href=""javascript:popCateSellDetail('"&cStatistic.FList(i).FCateL&"','"&cStatistic.FList(i).FCateM&"','"&cStatistic.FList(i).FCateS&"','"&cStatistic.FList(i).FCateX&"')"">(��)</a>"
		%>
	</td>
	<%if chkChannel ="1" then%>
	<td align="center">��ü</td>
	<%end if%>
	<%if (showordcnt="on") then%><td align="center"><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).FCountOrder)) %></td><%end if%>
	<td align="center"><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).FItemNO)) %></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FOrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost) %></b></td>
	<%if chkChannel ="1" then%>
	<td></td>
	<%end if%>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost-cStatistic.FList(i).FReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FBuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).FMaechulProfitPer %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice-cStatistic.FList(i).FBuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).FMaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%=formatnumber(vCateMRate,2)%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FupcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice - cStatistic.FList(i).FupcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FavgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FoverValueStockPrice) %></td>
	<td  align="center" <%if chkChannel ="1" then%>rowspan="6"<%end if%>><a href="/admin/maechul/statistic/statistic_item_analisys.asp?menupos=1726&date_gijun=<%=vDateGijun%>&syear=<%=vSYear%>&smonth=<%=vSMonth%>&sday=<%=vSDay%>&eyear=<%=vEYear%>&emonth=<%=vEMonth%>&eday=<%=vEDay%><%=linkcate&linkdispcate%>&incStockAvg=<%=incStockAvg%>" target="_blank">[��ǰ��]</a></td>
</tr>
<%if chkChannel ="1" then%>
<tr bgcolor="#ffffff" align="Center">
    <td>WEB</td>
    <%if (showordcnt="on") then%><td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fwww_countorder))%></td><%end if%>
    <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fwww_ItemNO))%></td>
    <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Fwww_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fwww_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ItemCost-cStatistic.FList(i).Fwww_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_MaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%=cStatistic.FList(i).Fwww_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ReducedPrice-cStatistic.FList(i).Fwww_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Fwww_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"> </td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_ReducedPrice - cStatistic.FList(i).Fwww_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fwww_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#ffffff" align="Center">
    <td >MOB</td>
    <%if (showordcnt="on") then%><td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fm_countorder))%></td><%end if%>
    <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fm_ItemNO)) %></td>
     <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fm_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Fm_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fm_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_ItemCost-cStatistic.FList(i).Fm_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fm_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Fm_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Fm_ReducedPrice-cStatistic.FList(i).Fm_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"> <%= cStatistic.FList(i).Fm_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"> </td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fm_ReducedPrice - cStatistic.FList(i).Fm_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fm_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#ffffff" align="Center">
    <td >APP</td>
    <%if (showordcnt="on") then%><td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fa_countorder))%></td><%end if%>
    <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fa_ItemNO)) %></td>
     <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fa_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Fa_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fa_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_ItemCost-cStatistic.FList(i).Fa_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fa_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Fa_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Fa_ReducedPrice-cStatistic.FList(i).Fa_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"> <%= cStatistic.FList(i).Fa_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"> </td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fa_ReducedPrice - cStatistic.FList(i).Fa_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fa_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#ffffff" align="Center">
    <td >����</td>
    <%if (showordcnt="on") then%><td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fo_countorder))%></td><%end if%>
    <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Fo_ItemNO)) %></td>
     <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fo_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Fo_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Fo_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_ItemCost-cStatistic.FList(i).Fo_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fo_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Fo_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Fo_ReducedPrice-cStatistic.FList(i).Fo_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"> <%= cStatistic.FList(i).Fo_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"> </td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Fo_ReducedPrice - cStatistic.FList(i).Fo_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fo_overValueStockPrice) %></td>
</tr>
<tr bgcolor="#ffffff" align="Center">
    <td >�ؿܸ�</td>
    <%if (showordcnt="on") then%><td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Ff_countorder))%></td><%end if%>
    <td><%= NullOrCurrFormat(CDbl(cStatistic.FList(i).Ff_ItemNO)) %></td>
     <% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_OrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_ItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#98F791"><b><%= NullOrCurrFormat(cStatistic.FList(i).Ff_ItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%if cStatistic.FList(i).Ff_ItemCost > 0 and cStatistic.FList(i).FItemCost > 0 then%><%=formatnumber((cStatistic.FList(i).Ff_ItemCost/cStatistic.FList(i).FItemCost)*100,0)%><%else%>0<%end if%>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_ItemCost-cStatistic.FList(i).Ff_ReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_ReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).Ff_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).Ff_MaechulProfitper%>%</td>
	<td align="right" style="padding-right:5px;"> <%= NullOrCurrFormat(cStatistic.FList(i).Ff_ReducedPrice-cStatistic.FList(i).Ff_BuyCash) %></td>
	<td align="right" style="padding-right:5px;"> <%= cStatistic.FList(i).Ff_MaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"> </td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_upcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).Ff_ReducedPrice - cStatistic.FList(i).Ff_upcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_avgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Ff_overValueStockPrice) %></td>
</tr>
<%end if%>
<%
    vTot_OrderCnt                   = vTot_OrderCnt + CDbl(NullOrCurrFormat(cStatistic.FList(i).FCountOrder))
	vTot_ItemNO						= vTot_ItemNO + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemNO))
	vTot_OrgitemCost				= vTot_OrgitemCost + CDbl(NullOrCurrFormat(cStatistic.FList(i).FOrgitemCost))
	vTot_ItemcostCouponNotApplied	= vTot_ItemcostCouponNotApplied + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemcostCouponNotApplied))
	vTot_ItemCost					= vTot_ItemCost + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemCost))
	vTot_BonusCouponPrice			= vTot_BonusCouponPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemCost-cStatistic.FList(i).FReducedPrice))
	vTot_ReducedPrice				= vTot_ReducedPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FReducedPrice))
	vTot_BuyCash					= vTot_BuyCash + CDbl(NullOrCurrFormat(cStatistic.FList(i).FBuyCash))
	vTot_MaechulProfit				= vTot_MaechulProfit + CDbl(NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit))
	vTot_MaechulProfit2				= vTot_MaechulProfit2 + CDbl(NullOrCurrFormat(cStatistic.FList(i).FReducedPrice-cStatistic.FList(i).FBuyCash))
	vTot_CateMRate					= vTot_CateMRate + vCateMRate
	vTot_upcheJungsan				= vTot_upcheJungsan + CDbl(NullOrCurrFormat(cStatistic.FList(i).FupcheJungsan))
	vTot_avgipgoPrice				= vTot_avgipgoPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FavgipgoPrice))
	vTot_overValueStockPrice		= vTot_overValueStockPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FoverValueStockPrice))
Next

	vTot_MaechulProfitPer = Round(((vTot_ItemCost - vTot_BuyCash)/CHKIIF(vTot_ItemCost=0,1,vTot_ItemCost))*100,2)
	vTot_MaechulProfitPer2 = Round(((vTot_ReducedPrice - vTot_BuyCash)/CHKIIF(vTot_ReducedPrice=0,1,vTot_ReducedPrice))*100,2)
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<%if vShowDate ="Y" then%>
	<td align="center"></td>
	<%end if%>
	<td align="center" <%if chkChannel ="1" then%>colspan="2"<%end if%>>�Ѱ�</td>
	<%if (showordcnt="on") then%><td align="center"><%=NullOrCurrFormat(vTot_OrderCnt)%></td><%end if%>
	<td align="center"><%=NullOrCurrFormat(vTot_ItemNO)%></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_OrgitemCost)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ItemcostCouponNotApplied)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_ItemCost)%></b></td>
	<%if chkChannel ="1" then%><td></td><%end if%>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_BonusCouponPrice)%></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ReducedPrice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_BuyCash)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer%>%</td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_MaechulProfit2)%></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer2%>%</td>
	<td align="right" style="padding-right:5px;"><%=formatnumber(vTot_CateMRate,2)%>%</td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_upcheJungsan)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_ReducedPrice - vTot_upcheJungsan)%></b></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_avgipgoPrice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_overValueStockPrice)%></td>
	<td></td>
</tr>
</table>
<% Set cStatistic = Nothing

Function CateGubun(g)
	If g = "L" Then
		CateGubun = "��"
	ElseIf vCateGubun = "M" Then
		CateGubun = "��"
	ElseIf vCateGubun = "S" Then
		CateGubun = "��"
	End IF
End Function
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
