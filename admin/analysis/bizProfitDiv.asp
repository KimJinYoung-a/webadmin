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
Dim yyyymm : yyyymm=requestCheckvar(request("yyyymm"),10)
Dim research : research=requestCheckvar(request("research"),10)
Dim dSDate,dEDate

IF (yyyymm="") then
    yyyymm = Left(DateAdd("m",-1,now()),7)
ENd IF

dSDate  = yyyymm+"-01"
dEDate = Left(DAteAdd("d",-1,dateadd("m",1,dSDate)),10)

Dim SANSTS : SANSTS=requestCheckvar(request("SANSTS"),10) ''��ǥ����
Dim bizSecCd : bizSecCd=requestCheckvar(request("bizSecCd"),16)
Dim accusecd  : accusecd=requestCheckvar(request("accusecd"),16)
Dim isINTrans  : isINTrans=requestCheckvar(request("isINTrans"),10) ''���ΰŷ�
Dim sST

''����ι�
Dim clsBS, arrBizList
Set clsBS = new CBizSection 
	clsBS.FUSE_YN = "Y"  
	clsBS.FOnlySub = "Y"  
	arrBizList = clsBS.fnGetBizSectionList  
Set clsBS = nothing

Dim intLoop, i, j, k
Dim debitSum, creditSum, cntSum
debitSum = 0
creditSum = 0


Dim oBizDivCrossTab
set oBizDivCrossTab = new CBizProfit
oBizDivCrossTab.FRectStdt = dSDate
oBizDivCrossTab.FRectEddt = dEDate
oBizDivCrossTab.FRectBizSecCD = bizSecCd
oBizDivCrossTab.FRectAccUseCd = accusecd
oBizDivCrossTab.FRectSANSTS = SANSTS
oBizDivCrossTab.FRectINTRANS = isINTrans
oBizDivCrossTab.getBizProfitDivCrossTabList

Dim oBizDivCrossList
set oBizDivCrossList = new CBizProfit
oBizDivCrossList.FRectStdt = dSDate
oBizDivCrossList.FRectEddt = dEDate
oBizDivCrossList.FRectBizSecCD = bizSecCd
oBizDivCrossList.FRectAccUseCd = accusecd
oBizDivCrossList.FRectSANSTS = SANSTS
oBizDivCrossList.FRectINTRANS = isINTrans
oBizDivCrossList.getBizProfitDivCrossList

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

function popDivSet(frm){
    var iURI = "popBizDivSet.asp?yyyymm=<%=yyyymm%>";
    var popwin = window.open(iURI,'popBizDivSet','scrollbars=yes,resizable=yes,width=900,height=600');
    popwin.focus();
}
</script>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a">
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
					
					��ǥ���:
					<% CaLL DrawYYYYMMsimpleBox("yyyymm",yyyymm,"onChange=''") %>
					
					&nbsp;
					<input type="button" value="�Ⱥб�Ģ����" class="button" onClick="popDivSet(frm);">
					
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
					
                    &nbsp;&nbsp;
					<input type="checkbox" name="isINTrans" value="Y" <%= ChkIIF(isINTrans="Y","checked","") %> >���ΰŷ���
			    </td>
			</tr>
			<tr>
			    <td bgcolor="#FFFFFF" >
					
			    </td>
			</tr>
			</form>
		</table>
	</td>
</tr>
</table>

<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="7">
			<!--�˻���� : <b></b> &nbsp;-->
		</td>
		<td colspan="<%= oBizDivCrossTab.FREsultCount*2 %>" align="right">
		   
		</td>
	</tr>     
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <!-- td rowspan="2">����ι�</td -->
	    <td rowspan="2">����</td>
		<td rowspan="2">�����з�</td>
		<td rowspan="2">��������/�ڵ�</td>
		<td rowspan="2">��ǥ����</td>
		<td rowspan="2">����</td>
		<td rowspan="2">����</td>
		<td rowspan="2">�Ⱥ�<br>����</td>
		<% for i=0 to oBizDivCrossTab.FREsultCount-1 %>
		<td colspan="2"><%= oBizDivCrossTab.FItemList(i).FBIZSECTION_NM %></td>
		<% next %>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <% for i=0 to oBizDivCrossTab.FREsultCount-1 %>
		<td ><%= FormatNumber(oBizDivCrossTab.FItemList(i).FdebitSum,0) %></td>
		<td >(%)</td>
		<% next %>
	</tr>    
	<% Dim pSLTRKEYSEQ, bFound, tdAdd %>
	<% for i=0 to oBizDivCrossList.FResultCount-1 %>
    	<% if pSLTRKEYSEQ<>oBizDivCrossList.FITemList(i).FSLTRKEY&"_"&oBizDivCrossList.FITemList(i).FSLTRKEY_SEQ then %>
    	<% if i<>0 then %>
        	<% if (tdAdd<oBizDivCrossTab.FResultCount) then %>
        	<% for k=tdAdd to oBizDivCrossTab.FResultCount-1 %>
        	<td >&nbsp;</td>
	        <td >&nbsp;</td>
        	<% next %>
        	<% end if %>
        	</tr>
        	<% tdAdd =0 %>
    	<% end if %>
    	<tr align="center" bgcolor="#FFFFFF">
    	    <!-- td><%= oBizDivCrossList.FITemList(i).ForgBIZSECTION_NM %></td -->
    	    <td><%= oBizDivCrossList.FITemList(i).FACC_GRP_NM %></td>
    	    <td><%= oBizDivCrossList.FITemList(i).FACC_CD_UPNM %></td>
    	    <td><%= oBizDivCrossList.FITemList(i).FACC_NM %><br>(<%= oBizDivCrossList.FITemList(i).FACC_USE_CD %>)</td>
    	    <td><%= oBizDivCrossList.FITemList(i).FSLDATE %></td>
    	    <td><%= oBizDivCrossList.FITemList(i).FACC_CD_RMK %></td>
    	    <td><%= FormatNumber(oBizDivCrossList.FITemList(i).FOrgDEBIT,0) %></td>
    	    <td><%= Left(oBizDivCrossList.FITemList(i).getDivTypeName,2) %><br><%= oBizDivCrossList.FITemList(i).FdivKey %></td>
    	<% end if %>
	    <% bFound = false %>
	    <% for j=0 to oBizDivCrossTab.FResultCount-1 %>
    	    <% if (oBizDivCrossTab.FItemList(j).FBIZSECTION_CD=oBizDivCrossList.FItemList(i).FBIZSECTION_CD) then %>
    	    <% bFound = true %>
    	    
    		<td ><%= FormatNumber(oBizDivCrossList.FItemList(i).FdebitSum,0) %></td>
    		<td ><%= CLNG(oBizDivCrossList.FItemList(i).FdivPro*100)/100 %></td>
    		<% tdAdd=tdAdd+1 %>
    		<% exit for %>
    		<% end if %>
    		
    		<% if (Not bFound) and (tdAdd<=j) then %>
	        <td >&nbsp;</td>
	        <td >&nbsp;</td>
	        <% tdAdd=tdAdd+1 %>
	        <% end if %>
		<% next %>
	    
	<% pSLTRKEYSEQ=oBizDivCrossList.FITemList(i).FSLTRKEY&"_"&oBizDivCrossList.FITemList(i).FSLTRKEY_SEQ %>
	<% next %>
</table>
<%
SET oBizDivCrossTab  = nothing
SET oBizDivCrossList = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbiTmsClose.asp" -->