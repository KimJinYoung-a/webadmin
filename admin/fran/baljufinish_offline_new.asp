<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/order/logisticsbaljuofflinecls.asp"-->

<%

dim baljudate, baljunum, baljuid, searchtype, IsFinished, sqlStr

baljudate = RequestCheckVar(request("baljudate"),32)
baljunum = RequestCheckVar(request("baljunum"),32)
baljuid = RequestCheckVar(request("baljuid"),32)
searchtype = RequestCheckVar(request("searchtype"),32)

if (searchtype = "") then
    searchtype = "M"
end if

if (baljudate = "") then
    baljudate = Left(now, 10)
end if



'==============================================================================
'��������ڵ庰 ó���� �� ���ϴܴ�.
dim baljuitemoff
set baljuitemoff = New COfflineBalju
baljuitemoff.FRectBaljuNum = baljunum
baljuitemoff.FRectBaljuId = baljuid
baljuitemoff.FRectSiteSeq = GetLogicsSiteSeq		'/lib/classes/order/logisticsbaljuofflinecls.asp
baljuitemoff.FRectOnlyNoPackItem = "Y"

baljuitemoff.GetOfflineBaljuItemList

sqlStr = " select IsFinished " + VbCrLf
sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_Logistics_offline_baljumaster " + VbCrLf
sqlStr = sqlStr + " where baljuKey = " & baljunum & " " + VbCrLf
'response.write sqlStr & "<Br>"
rsget_Logistics.Open sqlStr,dbget_Logistics,1

if  not rsget_Logistics.EOF  then
	IsFinished 		= rsget_Logistics("IsFinished")
end if
rsget_Logistics.Close


dim TotalBaljucount, TotalUpchecount, TotalTenBaljucount
dim TotalNoPackCount, TotalPackCount, TotalDeliverCount, TotalEtcCount

dim i, minboxno, maxboxno

minboxno = -1



'==============================================================================
dim baljushopoff
set baljushopoff = New COfflineBalju
baljushopoff.FRectBaljuId = baljuid
baljushopoff.FRectBaljuNum = baljunum
'baljushopoff.FRectBaljuDate = baljudate
baljushopoff.FRectSiteSeq = GetLogicsSiteSeq		'/lib/function.asp
baljushopoff.FRectSelectedOnly = "N"

baljushopoff.GetOfflineBaljuShopList

Sub DrawMiChulgoDiv3(selectedname,selectedId)
	dim varexists
	varexists = false
%>
	<select class='select' name="<%= selectedname %>">
	<option value='' <% if selectedId="" then response.write " selected" %> ></option>
	<option value='5�ϳ����' <% if selectedId="5�ϳ����" then response.write " selected" %> >5�ϳ����</option>
	<option value='������' <% if selectedId="������" then response.write " selected" %> >������</option>
	<option value='����' <% if selectedId="����" then response.write " selected" %> >����</option>
	<% if (selectedId<>"") and (Not varexists) then %>
	<option value="<%= selectedId %>" id=special selected ><%= selectedId %></option>
	<% else %>
	<option value='��Ÿ�Է�' id=special <% if selectedId="��Ÿ�Է�" then response.write " selected" %> >��Ÿ�Է�</option>
	<% end if %>
	</select>
<%
end Sub


%>
<script>
function WriteBarcode(itemgubun, itemid, itemoption) {
    if (1*itemid>=1000000){
        var tmp = "" + (100000000 + (1 * itemid));
    }else{
        var tmp = "" + (1000000 + (1 * itemid));
    }
    document.frm.barcode.value = itemgubun + tmp.substring(1) + itemoption;
    barcodechulgo();
}

function showSpecialInput(objTarget){
	/*
	   if(objTarget[objTarget.selectedIndex].id=='special'){
	   output = window.showModalDialog("/lib/inputpop.html" , null, "dialogwidth:250px;dialogheight:120px;center:yes;scroll:no;resizable:no;status:no;help:no;");

	   if(output!=''){
	   objTarget[objTarget.selectedIndex].text=output;
	   objTarget[objTarget.selectedIndex].value=output;
	   }else{

	   }
	   }
	 */
}

function FinishBalju(isWait) {
    var f = document.frmarr;
    var u = document.uparr;
	var msg = "";

    u.itemgubun.value = "";
    u.itemid.value = "";
    u.itemoption.value = "";
    u.comment.value = "";

	<% if (IsFinished = "N") then %>
    for (var i = 0; i < f.elements.length; i++) {
        if ((f[i].name == "comment") && (f[i].selectedIndex != 0)) {
            u.itemgubun.value = u.itemgubun.value + "|" + f[i-3].value;
            u.itemid.value = u.itemid.value + "|" + f[i-2].value;
            u.itemoption.value = u.itemoption.value + "|" + f[i-1].value;

            u.comment.value = u.comment.value + "|" + f[i][f[i].selectedIndex].value;
        }
    }
	<% end if %>

	if (isWait === true) {
		u.isWait.value = "Y";
		msg = "�����!!\n\n�ش� ������ðǿ� ���� ������ø� ����� ��ȯ�մϴ�\n5������ ǥ���� ��ǰ�� ���������� ����˴ϴ�.\n���� ���Ϸ�ÿ� �ݿ��˴ϴ�.";
	} else {
		u.isWait.value = "N";
		<% if (IsFinished = "N") then %>
		msg = "�ش� ������ðǿ� ���� ������ø� ���Ϸ� ��ȯ�մϴ�\n5������ ǥ���� ��ǰ�� ���������� ����˴ϴ�.\n��� �ݿ��˴ϴ�.";
		<% else %>
		msg = "�ش� ������ðǿ� ���� ������ø� ���Ϸ� ��ȯ�մϴ�.\n��� �ݿ��˴ϴ�.";
		<% end if %>
	}

    if (confirm(msg) === true) {
        u.submit();
    }
}

</script>

<style>

.nomarginimg {
	display: block; margin: 0; padding: 0;
}

.listSep {
	border-top:0px #CCCCCC solid; height:1px; margin:0; padding:0;
}

.listSep2 {
	border-top:0px #555555 solid; height:1px; margin:0; padding:0;
}

.trheight20 {
	height: 20px;
}

</style>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<form name="frm" onsubmit="return false;">
	<input type=hidden name="baljunum" value="<%= baljunum %>">
	<tr height="10">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10" class="nomarginimg"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10" class="nomarginimg"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>��ǰ���</strong></font>
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr height="25">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td bgcolor="#F3F3FF">
			<br>
			&nbsp;
			��������ڵ� : <%= baljunum %>
			<input type="hidden" name="baljudate" value="<%= Left(baljudate,10) %>">
			&nbsp;
			������ :
			<select class="select" name="baljuid">
				<option value=''  selected>����</option>
				<% for i=0 to baljushopoff.FResultCount -1 %>
				<option value='<%= baljushopoff.FItemList(i).FBaljuId %>'  <% if (baljushopoff.FItemList(i).FBaljuId = baljuid) then %>selected<% end if %>><%= baljushopoff.FItemList(i).FBaljuName %></option>
				<% next %>
			</select>
			<input type="button" class="button" value=" �� �� " onclick="document.frm.submit();" style="width:80px; height: 22px;">
			<input type="button" class="button" value=" ������� " onclick="location.href='baljulist_offline_new.asp'" style="width:80px; height: 22px;">
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr height="10" valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10" class="nomarginimg"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10" class="nomarginimg"></td>
	</tr>
</form>
</table>

<p>

<table width="100%" height="200" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
  <form name="researchFrm" method=get>
  <input type=hidden name="baljuid" value="<%= baljuid %>">
  <input type=hidden name="baljunum" value="<%= baljunum %>">
  <input type=hidden name="baljudate" value="<%= baljudate %>">
  <tr height="10" valign="bottom">
      <td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10" class="nomarginimg"></td>
    <td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
    <td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10" class="nomarginimg"></td>
  </tr>
  <tr height="25" valign="top">
    <td background="/images/tbl_blue_round_04.gif"></td>
    <td background="/images/tbl_blue_round_06.gif">
    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    	<tr>
    	  <td>
            <img src="/images/icon_star.gif" align="absbottom">
            <font color="red"><strong>�������</strong></font> &nbsp
          </td>
    	  <td align="right">
			  <input type=button class=button value=" ���[���] " onclick="FinishBalju(true)" style="width:150px; height: 22px;" <%= CHKIIF(IsFinished="N", "", "disabled") %> >
			  &nbsp;
			  &nbsp;
			  &nbsp;
			  &nbsp;
			  &nbsp;
			  &nbsp;
			  <input type=button class=button value=" ���ó�� " onclick="FinishBalju(false)" style="width:150px; height: 22px;" <%= CHKIIF(IsFinished="Y", "disabled", "") %> >
    	  </td>
    	</tr>
    	</table>
    </td>
    <td background="/images/tbl_blue_round_05.gif"></td>
  </tr>
  </form>
  <tr valign="top">
    <td background="/images/tbl_blue_round_04.gif"></td>
    <td bgcolor="#FFFFFF">
		<table width="100%" border="0" cellspacing="2" cellpadding="0" class="a">
		<tr>
			<td>�귣��ID</td>
			<td align=center width=100>�ֹ��ڵ�</td>
			<td align="center" width=30>���<br>���</td>
			<td width=25>����</td>
			<td align="right" width="70">��ǰ<br>�ڵ�</td>
			<td align="center" width=35>�ɼ�</td>

			<td >��ǰ��</td>
			<td >�ɼǸ�</td>
			<td width=20></td>
			<td align=center width=45>����<br />�ֹ�</td>
			<td align=center width=45>�������<br>����</td>
			<td align=center width=50>�����<br>����</td>
			<td align=center width=50>����غ�<br>(Box in)</td>
			<td align=center width=45>��ŷ<br>�Ϸ�</td>
			<td align=center width=45>��<br />�ֹ�</td>
			<td align=center width=100>���</td>
		</tr>
		<form name="frmarr">
		<% for i=0 to baljuitemoff.FResultCount -1 %>
		<%
                                if ((minboxno = -1) or ((minboxno > baljuitemoff.FItemList(i).FRealBoxNo) and (baljuitemoff.FItemList(i).FBoxSongjangNo = "0"))) then
                                        minboxno = baljuitemoff.FItemList(i).FRealBoxNo
                                end if

                                if (maxboxno < baljuitemoff.FItemList(i).FRealBoxNo) then
                                        maxboxno = baljuitemoff.FItemList(i).FRealBoxNo
                                end if

                                TotalBaljucount      = TotalBaljucount + baljuitemoff.FItemList(i).Ftotalbaljuno
                                TotalUpchecount      = TotalUpchecount +  baljuitemoff.FItemList(i).Ftotalupcheno
                                TotalTenBaljucount   = TotalTenBaljucount +  baljuitemoff.FItemList(i).Ftotaltenbaeno

                                TotalNoPackCount     = TotalNoPackCount + baljuitemoff.FItemList(i).Ftotalnopackno
                                TotalPackCount       = TotalPackCount + baljuitemoff.FItemList(i).Ftotalpackno
                                TotalDeliverCount    = TotalDeliverCount + baljuitemoff.FItemList(i).Ftotaldeliverno
                                TotalEtcCount        = TotalEtcCount + baljuitemoff.FItemList(i).Ftotaletcno

		%>
		<tr>
			<td height="1" colspan="16" bgcolor="#CCCCCC"></td>
		</tr>
		<tr>
			<!--
			<td align="center"><%= Format00(4,baljuitemoff.FItemList(i).Fprtidx) %></td>
			-->
			<td ><%= baljuitemoff.FItemList(i).FMakerid %></td>
			<td align=center><%= baljuitemoff.FItemList(i).Fbaljucode %></td>
			<td align="center">
                <% if (baljuitemoff.FItemList(i).Fmaeipdiv = "U") then %>
                ����
                <% elseif (baljuitemoff.FItemList(i).Fmaeipdiv = "9") then %>
                ����
                <% else %>
                <!--�ٹ�-->
                <% end if %>
		    </td>
		    <input type=hidden name=itemgubun value="<%= baljuitemoff.FItemList(i).FItemGubun %>">
		    <input type=hidden name=itemid value="<%= baljuitemoff.FItemList(i).FItemID %>">
		    <input type=hidden name=itemoption value="<%= baljuitemoff.FItemList(i).FItemOption %>">
			<td align="center"><%= baljuitemoff.FItemList(i).FItemGubun %></td>
			<td align="right"><%= baljuitemoff.FItemList(i).FItemID %></td>
			<td align="center"><%= baljuitemoff.FItemList(i).FItemOption %></td>
			<td>
				<a href="#" onclick="TnPopItemStocknew('<%= baljuitemoff.FItemList(i).FItemGubun %>','<%= baljuitemoff.FItemList(i).FItemID %>','<%= baljuitemoff.FItemList(i).FItemOption %>'); return false;">
				<%= baljuitemoff.FItemList(i).FItemName %></a>
			</td>
			<td ><%= baljuitemoff.FItemList(i).FItemOptionName %></td>
			<td>
			<% if ((searchtype <> "P") and (searchtype <> "C")) then %>
			  =&gt;
			<% end if %>
			</td>
			<td align="center">--</td>
			<td align=center><%= baljuitemoff.FItemList(i).Ftotalbaljuno %></td>
			<td align=center>
				<% if (baljuitemoff.FItemList(i).Ftotalnopackno <> 0) then %>
			  		<font color="blue"><%= baljuitemoff.FItemList(i).Ftotalnopackno %></font>
                <% else %>
                	<%= baljuitemoff.FItemList(i).Ftotalnopackno %>
                <% end if %>
		    </td>
			<td align=center>
            	<% if (baljuitemoff.FItemList(i).Ftotalpackno <> baljuitemoff.FItemList(i).Ftotalbaljuno) then %>
			  		<font color="blue"><%= baljuitemoff.FItemList(i).Ftotalpackno %></font>
                <% else %>
                	<%= baljuitemoff.FItemList(i).Ftotalpackno %>
                <% end if %>
			</td>
			<td align=center>
                <% if (baljuitemoff.FItemList(i).Ftotaldeliverno <> baljuitemoff.FItemList(i).Ftotalbaljuno) then %>
			  		<b><font color="red"><%= baljuitemoff.FItemList(i).Ftotaldeliverno %></font></b>
                <% else %>
                	<b><%= baljuitemoff.FItemList(i).Ftotaldeliverno %></b>
				<% end if %>
			</td>
			<td align="center">--</td>
		    <td align=center>
				<% if (IsFinished="N") then %>
		        <% DrawMiChulgoDiv3 "comment", baljuitemoff.FItemList(i).Fcomment %>
				<% end if %>
		    </td>
		</tr>
		<% next %>
		</form>
		<tr>
			<td height="1" colspan="16" bgcolor="#CCCCCC"></td>
		</tr>
		<tr>
			<td align=center>Total</td>
			<td ></td>
			<td ></td>
			<td ></td>
			<td ></td>
			<td ></td>
			<td ></td>
			<td ></td>
			<td ></td>
			<td ></td>
			<td align=center><%= TotalBaljucount %></td>
			<td align=center><%= TotalNoPackCount %></td>
			<td align=center><%= TotalPackCount %></td>
			<td align=center><b><%= TotalDeliverCount %></b></td>
			<td ></td>
			<td ></td>
		</tr>
		</table>
    </td>
    <td background="/images/tbl_blue_round_05.gif"></td>
  </tr>

  <tr height="10" valign="top">
      <td><img src="/images/tbl_blue_round_07.gif" width="10" height="10" class="nomarginimg"></td>
      <td background="/images/tbl_blue_round_08.gif"></td>
      <td><img src="/images/tbl_blue_round_09.gif" width="10" height="10" class="nomarginimg"></td>
  </tr>
</table>
<form name="uparr" method="post" action="baljufinish_offline_new_process<%= CHKIIF(baljunum=54274, "_TEST", "") %>.asp">
	<input type=hidden name=mode value="chulgoproc">
	<input type=hidden name=isWait value="">
	<input type=hidden name=baljunum value="<%= baljunum %>">
	<input type=hidden name=baljuid value="<%= baljuid %>">
	<input type=hidden name=itemgubun value="">
	<input type=hidden name=itemid value="">
	<input type=hidden name=itemoption value="">
	<input type=hidden name=comment value="">
</form>
<%

set baljuitemoff = Nothing

%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
