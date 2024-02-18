<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �̼��� ���ڰ�꼭 ����
' History : 2012.02.07 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/tax/EseroTaxCls.asp"-->
<%
Dim clsEsero, arrList, intLoop
Dim iTotCnt,iPageSize, iTotalPage,page
Dim dSDate,dEDate,ssearchText,itaxsellType,itaxModiType,itaxType, iMapTpYn, iMapTp
Dim totSum, tgType
	iPageSize = 50
	page = requestCheckvar(Request("page"),10)
	if page="" then page=1


	dSDate = requestCheckvar(Request("dSD"),10)
	dEDate = requestCheckvar(Request("dED"),10)
	ssearchText = requestCheckvar(Request("sST"),200)
	itaxsellType = requestCheckvar(Request("iTST"),10)
	itaxModiType = requestCheckvar(Request("iTMT"),10)
	itaxType = requestCheckvar(Request("iTT"),10)
    iMapTpYn   = requestCheckvar(Request("iMapTpYn"),10)
    iMapTp     = requestCheckvar(Request("iMapTp"),10)
    totSum     = requestCheckvar(Request("totSum"),10)
    tgType     = requestCheckvar(Request("tgType"),20)

    if (itaxsellType="") then itaxsellType="0"

Set clsEsero = new CEsero
  clsEsero.FSDate      =dSDate
	clsEsero.FEDate      =dEDate
	clsEsero.FsearchText =ssearchText
	clsEsero.FtaxsellType=itaxsellType
	clsEsero.FtaxModiType=itaxModiType
	clsEsero.FtaxType    =itaxType
	clsEsero.FMappingTypeYN = iMapTpYn
	clsEsero.FMappingType   = iMapTp
	clsEsero.FtotSum     =totSum
	clsEsero.FCurrPage 	= page
	clsEsero.FPageSize 	= iPageSize
	arrList = clsEsero.fnGetEseroTaxList
	iTotCnt = clsEsero.FTotCnt
Set clsEsero = nothing
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<script type="text/javascript" src="/admin/approval/eapp/eapp.js"></script>
<script language="javascript">
<!--
// ������ �̵�
function jsGoPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}

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

	function jsSetTax(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP){
		opener.document.all.dView1.style.display = "";
		opener.document.frm.sEK.value= eTax;
		for(i=0;i<opener.document.frm.rdoDK.length;i++){
			opener.document.frm.rdoDK[i].checked = false;
			if(iDK==9){
				if(opener.document.frm.rdoDK[i].value ==2){
					opener.document.frm.rdoDK[i].checked= true;
				}
			}else{
				if(opener.document.frm.rdoDK[i].value ==1){
					opener.document.frm.rdoDK[i].checked= true;
				}
			}
		}


			if(iVK==1){
				opener.document.frm.sVK.value = "����(�ΰ��� 10%)";
				opener.document.frm.rdoVK.value = 0;
			}else if(iVK==2)	{
				opener.document.frm.sVK.value = "����";
				opener.document.frm.rdoVK.value = 3;
			}else{
				opener.document.frm.sVK.value = "�鼼";
				opener.document.frm.rdoVK.value = 2;
			}

		opener.document.frm.dID.value= dID;
		opener.document.frm.sINm.value= sInm;
		opener.document.frm.mTP.value= jsSetComma(mTP);
		opener.document.frm.mSP.value= jsSetComma(mSP);
		opener.document.frm.mVP.value= jsSetComma(mVP);
		
		opener.jsTexSetting();

		self.close();
	}

	function jsSetTaxNormal(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP){
	    opener.fillTaxInfo(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP);
	    self.close();
	}

	function jsSetTaxWithpayreq(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP,prIdx){
	    opener.fillTaxInfoWithPayreqIdx(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP,prIdx);
	    self.close();
	}

function popErpSending(itaxkey){
    var winD = window.open("popRegfileHand.asp?taxkey="+itaxkey,"popErpSending","width=860, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}

//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
	<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="page" value="">
			<input type="hidden" name="tgType" value="<%= tgType %>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td  rowspan="2" width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
				<td align="left">
					<input type="radio" name="iTST" value="0" <%= CHKIIF(itaxsellType="0","checked","") %> >����
					<input type="radio" name="iTST" value="1" <%= CHKIIF(itaxsellType="1","checked","") %> >����&nbsp;&nbsp;
					 �ۼ���:
					<input type="text" name="dSD" size="10" value="<%=dSDate%>" onClick="jsPopCal('dSD');"  style="cursor:hand;">
					-
					<input type="text" name="dED" size="10" value="<%=dEDate%>" onClick="jsPopCal('dED');"  style="cursor:hand;">
					&nbsp;&nbsp;�˻���:
					<input type="text" name="sST" value="<%=ssearchText%>" size="30"><font color="Gray">(����ڵ�Ϲ�ȣ,��ȣ,ǰ���)</font>
				</td>
				<td  rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="jsSearch();">
				</td>
			</tr>
			<tr>
			    <td  bgcolor="#FFFFFF">
			        ��Ī���� :
			        <select Name="iMapTpYn">
			        <option value="">��ü
			        <option value="Y" <%= CHKIIF(iMapTpYn="Y","selected","") %> >��Ī
			        <option value="N" <%= CHKIIF(iMapTpYn="N","selected","") %> >���Ī
			        </select>
			        &nbsp;&nbsp;
			        ��Ī���� :
			        <select Name="iMapTp">
			        <option value="">��ü
			        <option value="1" <%= CHKIIF(iMapTp="1","selected","") %> >�¶��� ����
			        <option value="2" <%= CHKIIF(iMapTp="2","selected","") %> >�������� ����
			        <option value="9" <%= CHKIIF(iMapTp="9","selected","") %> >��Ÿ ����
			        <option value="11" <%= CHKIIF(iMapTp="11","selected","") %> >����
			        </select>
			        &nbsp;&nbsp;
			        ��꼭����:
			        <select Name="iTMT">
			        <option value="">��ü
			        <option value="0" <%= CHKIIF(itaxModiType="0","selected","") %> >����(�Ϲ�)
			        <option value="1" <%= CHKIIF(itaxModiType="1","selected","") %> >����(����)
			        <option value="9" <%= CHKIIF(itaxModiType="9","selected","") %> >��Ÿ(����)
			        </select>
			        &nbsp;&nbsp;
			        ��������:
			        <select Name="iTT">
			        <option value="">��ü
			        <option value="1" <%= CHKIIF(itaxType="1","selected","") %> >����
			        <option value="2" <%= CHKIIF(itaxType="2","selected","") %> >����
			        <option value="3" <%= CHKIIF(itaxType="3","selected","") %> >�鼼
			        </select>
			        &nbsp;&nbsp;
			        �ݾ�:
			        <input type="text" name="totSum" value="<%= totSum %>" maxlength="9" size="10">
			    </td>
			</tr>
			</form>
		</table>
	</td>
</tr>
<tr>
	<td>
		<!-- ��� �� ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="FFFFFF">
				<td colspan="19">
					�˻���� : <b><%=iTotCnt%></b> &nbsp;
					������ : <b><%= page %> / <%=iTotalPage%></b>
				</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td rowspan="2">�ۼ�����</td>
				<td rowspan="2">���ι�ȣ</td>
				<td colspan="2"><%IF itaxsellType="0" THEN%>������<%ELSE%>���޹޴���<%END IF%></td>
				<td rowspan="2">�հ�ݾ�</td>
				<td rowspan="2">���ް���</td>
				<td rowspan="2">����</td>
				<td rowspan="2">�з�</td>
				<td rowspan="2">����</td>
				<td rowspan="2">ǰ���</td>
				<td rowspan="2">����<br>����</td>
				<td rowspan="2">����<br>����</td>
				<td rowspan="2">����ι�</td>
				<td rowspan="2">ERP<br>���ۻ���</td>
				<td rowspan="2">����</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td>����ڵ�Ϲ�ȣ</td>
				<!-- td>��</td -->
				<td>��ȣ</td>
			</tr>
			<%
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
				%>
			<tr align="center" bgcolor="#FFFFFF">
			    <td><%= arrList(1,intLoop) %></td>
			    <td><a href="javascript:popErpSending('<%= arrList(0,intLoop) %>')"><%= arrList(0,intLoop) %></a></td>
			    <% if arrList(15,intLoop)=1 then %>
			    <td><a href="javascript:popHandMapping('<%= arrList(15,intLoop) %>','<%= arrList(1,intLoop) %>','<%= arrList(0,intLoop) %>','<%= arrList(7,intLoop) %>')"><%= arrList(7,intLoop) %></a></td>
			    <td><%= arrList(9,intLoop) %></td>
			    <% else %>
			    <td><a href="javascript:popHandMapping('<%= arrList(15,intLoop) %>','<%= arrList(1,intLoop) %>','<%= arrList(0,intLoop) %>','<%= arrList(2,intLoop) %>')"><%= arrList(2,intLoop) %></a></td>
			    <td><%= arrList(4,intLoop) %></td>
			    <% end if %>
			    <td align="right"><%= FormatNumber(arrList(12,intLoop),0) %></td>
			    <td align="right"><%= FormatNumber(arrList(13,intLoop),0) %></td>
			    <td align="right"><%= FormatNumber(arrList(14,intLoop),0) %></td>
			    <td><%= getSellTypeName(arrList(15,intLoop)) %></td>
			    <td><%= gettaxModiTypeName(arrList(16,intLoop)) %>/<%= gettaxTypeName(arrList(17,intLoop)) %></td>
			    <td><%= arrList(22,intLoop) %></td>
			    <td><%= getMatchStateName(arrList(31,intLoop)) %></td>
			    <td>
			        <% if (tgType="NRM") and Not IsNULL(arrList(31,intLoop)) and (arrList(29,intLoop)="9") then %>
			        <a href="javascript:jsSetTaxWithpayreq('<%=arrList(0,intLoop)%>','<%=arrList(16,intLoop)%>','<%=arrList(17,intLoop)%>','<%=arrList(1,intLoop)%>','<%=sReSearchText%>','<%=arrList(12,intLoop)%>','<%=arrList(13,intLoop)%>','<%=arrList(14,intLoop)%>','<%= arrList(30,intLoop) %>');"><%= getMatchTypeName(arrList(29,intLoop)) %><br><%= arrList(30,intLoop) %></a>
			        <% else %>
			        <%= getMatchTypeName(arrList(29,intLoop)) %><br><%= arrList(30,intLoop) %>
			        <% end if %>
			    </td>
			    <td><%= getbizSecCDName(arrList(32,intLoop)) %>
			    <% if arrList(35,intLoop)>0 then %>
			    �� <%= arrList(35,intLoop) %>
			    <% end if %>
			    </td>
			    <td>
			        <% if Not IsNULL(arrList(33,intLoop)) then %>
    			    [<%= arrList(33,intLoop) %>]
    			    <%= arrList(34,intLoop) %>
			        <% end if %>
			    </td>
			   <td><%Dim sReSearchText
			    	sReSearchText = replace(arrList(22,intLoop),"'","\'")
			    	sReSearchText = replace(sReSearchText,"""","")
			    	%>
			        <input <%= chkIIF(not IsNULL(arrList(31,intLoop)),"disabled","") %> type="button" class="button" value="����" onClick="<%= CHKIIF(tgType="NRM","jsSetTaxNormal","jsSetTax") %>('<%=arrList(0,intLoop)%>','<%=arrList(16,intLoop)%>','<%=arrList(17,intLoop)%>','<%=arrList(1,intLoop)%>','<%=sReSearchText%>','<%=arrList(12,intLoop)%>','<%=arrList(13,intLoop)%>','<%=arrList(14,intLoop)%>')">

					<%
						if (C_ADMIN_AUTH) or (C_MngPart) then
							if (not IsNULL(arrList(31,intLoop))) and (not IsNULL(arrList(29,intLoop))) then
					%>
						<input type="button" class="button_auth" value="����(������)" onClick="<%= CHKIIF(tgType="NRM","jsSetTaxNormal","jsSetTax") %>('<%=arrList(0,intLoop)%>','<%=arrList(16,intLoop)%>','<%=arrList(17,intLoop)%>','<%=arrList(1,intLoop)%>','<%=sReSearchText%>','<%=arrList(12,intLoop)%>','<%=arrList(13,intLoop)%>','<%=arrList(14,intLoop)%>')">
			        <%
							end if
						end if
					%>
			    </td>
			</tr>
		<%	Next
			ELSE%>
			<tr height=30 align="center" bgcolor="#FFFFFF">
				<td colspan="19">��ϵ� ������ �����ϴ�.</td>
			</tr>
			<%END IF%>
		</table>
	</td>
</tr>
<!-- ������ ���� -->
		<%
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10

		iStartPage = (Int((page-1)/iPerCnt)*iPerCnt) + 1

		If (page mod iPerCnt) = 0 Then
			iEndPage = page
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
			<tr height="25" >
				<td colspan="15" align="center">
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
					    <tr valign="bottom" height="25">
					        <td valign="bottom" align="center">
					         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
							<% else %>[pre]<% end if %>
					        <%
								for ix = iStartPage  to iEndPage
									if (ix > iTotalPage) then Exit for
									if Cint(ix) = Cint(page) then
							%>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
							<%		else %>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
							<%
									end if
								next
							%>
					    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
							<% else %>[next]<% end if %>
					        </td>
					    </tr>
					</table>
				</td>
			</tr>
</table>
<!-- ������ �� -->
</body>
</html>




