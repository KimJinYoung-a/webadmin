<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbiTmsOpen.asp" -->
<!-- #include virtual="/lib/db/dbiTMSHelper.asp"-->  
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/BizProfitCls.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<%
Dim dSDate : dSDate=requestCheckvar(request("dSDate"),10)
Dim dEDate : dEDate=requestCheckvar(request("dEDate"),10)
Dim research : research=requestCheckvar(request("research"),10)
IF (dSDate="") then
    dSDate = Left(DateAdd("m",-1,now()),7)+"-01"
ENd IF

IF (dEDate="") then
    dEDate = Left(DAteAdd("d",-1,Left(now(),7)+"-01"),10)
ENd IF

Dim SANSTS : SANSTS=requestCheckvar(request("SANSTS"),10) ''��ǥ����
Dim bizSecCd : bizSecCd=requestCheckvar(request("bizSecCd"),16)
Dim accusecd  : accusecd=requestCheckvar(request("accusecd"),16)
Dim isINTrans  : isINTrans=requestCheckvar(request("isINTrans"),10) ''���ΰŷ�
Dim divAssign  : divAssign=requestCheckvar(request("divAssign"),10) ''�Ⱥ�����
Dim divdpType  : divdpType=requestCheckvar(request("divdpType"),10) ''�Ⱥ�DP Type
Dim sST

IF Len(accusecd)=3 then accusecd=accusecd&"00"
IF (divAssign="Y") and (divdpType="") then divdpType="0"

Dim oBizProfit
set oBizProfit = new CBizProfit
oBizProfit.FRectBizSecCD=bizSecCd
oBizProfit.FRectStdt = Replace(dSDate,"-","")
oBizProfit.FRectEddt = Replace(dEDate,"-","")
oBizProfit.FRectSANSTS = SANSTS
oBizProfit.FRectAccUseCd = accusecd
oBizProfit.FRectINTRANS = isINTrans
oBizProfit.FRectdivAssign = divAssign
oBizProfit.FRectdivdpType = divdpType

oBizProfit.getBizProfitSum


''����ι�
Dim clsBS, arrBizList
Set clsBS = new CBizSection 
	clsBS.FUSE_YN = "Y"  
	clsBS.FOnlySub = "Y"  
	arrBizList = clsBS.fnGetBizSectionList  
Set clsBS = nothing

Dim intLoop, i
Dim debitSum, creditSum, cntSum
debitSum = 0
creditSum = 0

Dim dpOrgBIZSEC
dpOrgBIZSEC = (divAssign="Y") and ((divdpType="0")or(divdpType="1"))
%>

<script language='javascript'>
//�˻�
function jsSearch(){
	document.frm.submit();
}

//�޷º���
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function jsFillCal(comp1,comp2,val1,val2){
    comp1.value = val1;
    comp2.value = val2;
}

function showProfitDetail(bizSecCd,accusecd){
    //var iURI = "popBizProfitDetail.asp?dSDate=<%=dSDate%>&dEDate=<%=dEDate%>&bizSecCd="+bizSecCd+"&acccdup="+acccdup+"&acccd="+acccd+"&SANSTS=<%=SANSTS%>&isINTrans=<%=isINTrans%>";
    var iURI = "popBizProfitDetail.asp?dSDate=<%=dSDate%>&dEDate=<%=dEDate%>&bizSecCd="+bizSecCd+"&accusecd="+accusecd+"&SANSTS=<%=SANSTS%>&isINTrans=<%=isINTrans%>";
    var popwin = window.open(iURI,'showProfitDetail','scrollbars=yes,resizable=yes,width=900,height=600');
    popwin.focus();
}

function checkComp(comp){
    var frm = comp.form;
    if (comp.name=="divAssign"){
        frm.divdpType[0].disabled=!comp.checked;
        frm.divdpType[1].disabled=!comp.checked;
        frm.divdpType[2].disabled=!comp.checked;
        
        if ((comp.checked)&&(!frm.divdpType[0].checked)&&(!frm.divdpType[1].checked)&&(!frm.divdpType[2].checked)){
            frm.divdpType[0].checked=true;
        }
    }
}
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
	<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="page" value="">
			<input type="hidden" name="research" value="on">
			<tr align="center" bgcolor="#FFFFFF" >
				<td  rowspan="3" width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
				<td align="left">
					
					��ǥ��¥:
					<input type="text" name="dSDate" size="10" value="<%=dSDate%>" onClick="jsPopCal('dSDate');"  style="cursor:hand;">
					-
					<input type="text" name="dEDate" size="10" value="<%=dEDate%>" onClick="jsPopCal('dEDate');"  style="cursor:hand;">
					<input type="button" value="������" class="button" onClick="jsFillCal(document.frm.dSDate,document.frm.dEDate,'<%= Left(DateAdd("m",-2,now()),7)+"-01" %>','<%= Left(DateAdd("d",-1,Left(CStr(dateadd("m",-1,now())),7)+"-01" ),10) %>')";>
					<input type="button" value="����" class="button" onClick="jsFillCal(document.frm.dSDate,document.frm.dEDate,'<%= Left(DateAdd("m",-1,now()),7)+"-01" %>','<%= Left(DateAdd("d",-1,Left(CStr(dateadd("m",0,now())),7)+"-01" ),10) %>')";>
					<input type="button" value="�̹���" class="button" onClick="jsFillCal(document.frm.dSDate,document.frm.dEDate,'<%= Left(DateAdd("m",0,now()),7)+"-01" %>','<%= Left(DateAdd("d",-1,Left(CStr(dateadd("m",1,now())),7)+"-01" ),10) %>')";>
					
					&nbsp;&nbsp;
					<input type="checkbox" name="isINTrans" value="Y" <%= ChkIIF(isINTrans="Y","checked","") %> >���ΰŷ���
					
				</td>
				<td  rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="jsSearch();">
				</td>
			</tr>
			<!--
			<tr align="center" bgcolor="#FFFFFF" >
			    <td align="left">
			        �˻���:
					<input type="text" name="sST" value="<%=sST%>" size="30"><font color="Gray">(����ڵ�Ϲ�ȣ,���ι�ȣ,��ȣ,ǰ���)</font>
					&nbsp;&nbsp;
			        
			        
			    </td>
			</tr>
			-->
			<tr>
			    <td  bgcolor="#FFFFFF">
			        ��ǥ���� :
			        <select Name="SANSTS">
			        <option value="">��ü
			        <option value="1" <%= CHKIIF(SANSTS="1","selected","") %> >����
			        <option value="0" <%= CHKIIF(SANSTS="0","selected","") %> >�̽���
			        </select>
			        &nbsp;&nbsp;
			        
			        
			        &nbsp;&nbsp;
					����ι�:
                    <select name="bizSecCd">
                    <option value="">--����--</option>
                    <% For intLoop = 0 To UBound(arrBizList,2)	%>
                		<option value="<%=arrBizList(0,intLoop)%>" <%IF Cstr(bizSecCd) = Cstr(arrBizList(0,intLoop)) THEN%> selected <%END IF%>><%=arrBizList(1,intLoop)%></option>
                	<% Next %>
                    </select>
                    &nbsp;&nbsp;
                    ���������ڵ�:
					<input type="text" name="accusecd" value="<%=accusecd%>" size="15">
					
                    
			    </td>
			</tr>
			<tr>
			    <td bgcolor="#FFFFFF">
					<input type="checkbox" name="divAssign" value="Y" <%= ChkIIF(divAssign="Y","checked","") %> onClick="checkComp(this)">�Ⱥ�����
					&nbsp;
					::
					&nbsp;
					<input type="radio" name="divdpType" value="0" <%= ChkIIF(divdpType="0","checked","") %> <%= ChkIIF(divAssign="Y","","disabled") %>> �Ⱥг��� �и�ǥ��
					<input type="radio" name="divdpType" value="2" <%= ChkIIF(divdpType="2","checked","") %> <%= ChkIIF(divAssign="Y","","disabled") %>> �Ⱥг��� ���ļ�ǥ��
					<input type="radio" name="divdpType" value="1" <%= ChkIIF(divdpType="1","checked","") %> <%= ChkIIF(divAssign="Y","","disabled") %>> �Ⱥг����� ǥ��
			    </td>
			</tr>
			</form>
		</table>
	</td>
</tr>
</table>

<p>
<!-- ��� �� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="2">
			�˻���� : <b><%=oBizProfit.FTotalCount%></b> &nbsp;
		</td>
		<td colspan="12" align="right">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td >����ι�</td>
		<td >����</td>
		<td >�����ڵ�</td>
		<td >�����з�</td>
		<td >��������</td>
		<td >����</td>
		<td >�뺯</td>
		<td >�Ǽ�</td>
		<% if (dpOrgBIZSEC) then %>
		<td >�Ⱥ���<br>����ι�</td>
		<% end if %>
		<td >��</td>
	</tr>
	<% if oBizProfit.FResultCount>0 then %>
	<% For i = 0 To oBizProfit.FResultCount-1 %>
	<%
	debitSum    = debitSum + oBizProfit.FItemList(i).FdebitSum
	creditSum   = creditSum + oBizProfit.FItemList(i).FcreditSum
	cntSum      = cntSum + oBizProfit.FItemList(i).FjpCNT
	%>
	<tr align="center" bgcolor="<%= CHKIIF(oBizProfit.FItemList(i).IsDivAssigned,"#C7EEC7","#FFFFFF") %>">
	    <td><%= oBizProfit.FItemList(i).FBIZSECTION_NM %></td>
	    <td><%= oBizProfit.FItemList(i).FACC_GRP_NM %></td>
        <td><%= oBizProfit.FItemList(i).FACC_USE_CD %></td>    
        <td>
            <%= oBizProfit.FItemList(i).FACC_CD_UPNM %>
        </td>
        <td>
            <%= oBizProfit.FItemList(i).FACC_NM %>
        </td>      
        <td align="right"><%= FormatNumber(oBizProfit.FItemList(i).FdebitSum,0) %></td> 
        <td align="right"><%= FormatNumber(oBizProfit.FItemList(i).FcreditSum,0) %></td> 
        <td align="center"><%= oBizProfit.FItemList(i).FjpCNT %></td> 
        <% if (dpOrgBIZSEC) then %>
		<td >
		    <% if (oBizProfit.FItemList(i).IsDivAssigned) then %>
		    <%= oBizProfit.FItemList(i).FOrgBIZSECTION_NM %>
		    <% end if %>
		</td>
		<% end if %>
        <td>
        <img src="/images/icon_arrow_link.gif" onClick="showProfitDetail('<%= oBizProfit.FItemList(i).FBIZSECTION_CD %>','<%= oBizProfit.FItemList(i).FACC_USE_CD %>');" style="cursor:pointer">    
        </td> 
        
	</tr>
    <%	Next %>
    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="5"></td>
        <td align="right"><%= FormatNumber(debitSum,0) %></td>
        <td align="right"><%= FormatNumber(creditSum,0) %></td>
        <td align="center"><%= cntSum %></td>
        <% if (dpOrgBIZSEC) then %>
		<td ></td>
		<% end if %>
        <td></td>
    </tr>
	<% ELSE%>
	<tr height=30 align="center" bgcolor="#FFFFFF">
		<td colspan="19">��ϵ� ������ �����ϴ�.</td>
	</tr>
	<%END IF%>
</table>
<%
set oBizProfit = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbiTmsClose.asp" -->