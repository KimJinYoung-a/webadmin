<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������û�� ����Ʈ
' History : 2011.10.13 ������  ����
'' ToDo ��Ÿ��(�޿�) ���ۺҰ�(DBŸ�Լ��� or ) // ȯ�� ����..
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/eappListCls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<!-- #include virtual="/lib/classes/approval/edmsCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%

dim research
Dim clsEapp, clsedms
Dim ireportstate ,sadminId
Dim ireportidx
Dim iCurrpage,ipagesize,iTotCnt,iTotalPage
Dim arrList,intLoop
Dim iarap_cd,sarap_nm
Dim searchsdate,searchedate, susername, sreportname
Dim sOrderType
Dim icateidx1, sdatetype,sedmscode
dim department_id, inc_subdepartment

	iPageSize = 30
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1

	sadminId =  session("ssBctId")
	icateidx1	= requestCheckvar(Request("icidx1"),10)
	sdatetype	= requestCheckvar(Request("selDT"),10)
	ireportidx	= requestCheckvar(Request("iridx"),10)
	searchsdate= requestCheckvar(Request("selSD"),10)
	searchedate= requestCheckvar(Request("selED"),10)
	iarap_cd		= requestCheckvar(Request("iaidx"),13)
	sarap_nm		= requestCheckvar(Request("selarap"),50)
  ireportstate= requestCheckvar(Request("selPRS"),4)
  sedmscode 	= requestCheckvar(Request("sec"),10)
	sUserName		= requestCheckvar(Request("sUnm"),30)
	sreportname = requestCheckvar(Request("sRnm"),120)
	sOrderType	= requestCheckvar(Request("selOT"),1)
	department_id = requestCheckvar(Request("department_id"),10)
	inc_subdepartment = requestCheckvar(Request("inc_subdepartment"),1)
	research = requestCheckvar(Request("research"),10)

	if (research = "") then
		''searchsdate = Left(DateAdd("m", -6, Now()), 10)
	end if

	'�޴��� ���� �⺻ ī�װ� ����
	if menupos="1402" and Not(icateidx1="5" or icateidx1="12") then
		icateidx1="5"	'����ǰ��
	elseif menupos="1617" and Not(icateidx1="3" or icateidx1="12") then
		icateidx1="3"	'�λ�
	end if

'���� �⺻ �� ���� ��������
set clsEapp = new CEappList
	clsEapp.Fcateidx1				= icateidx1
	clsEapp.FdateType				= sdatetype
	clsEapp.FStartDate				= searchsdate
	clsEapp.FEndDate				= searchedate
	clsEapp.FUsername				= sUserName
	clsEapp.FreportName				= sreportname
	clsEapp.FreportState    		= ireportstate
	clsEapp.Fedmscode				= sedmscode
 	clsEapp.Farap_cd				= iarap_cd
 	clsEapp.Farap_nm				= sarap_nm
 	clsEapp.FOrderType				= sOrderType
	clsEapp.FCurrpage 				= iCurrpage
	clsEapp.FPagesize				= ipagesize
	clsEapp.Fdepartment_id 			= department_id
	clsEapp.Finc_subdepartment 		= inc_subdepartment

	arrList = clsEapp.fnGetEappList
	iTotCnt = clsEapp.FTotCnt
set clsEapp = nothing
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ �� 

%>
<script language="javascript" src="/admin/approval/eapp/eapp.js"></script>
<script language="javascript">
<!--
	function jsView(iridx){
		var winR = window.open("/admin/approval/eapp/vieweapp.asp?iridx="+iridx,"popR","width=1000, height=600, resizable=yes, scrollbars=yes");
		winR.focus();
	}

	function jsSearch(){
	 document.frm.submit();
	}

	// ������ �̵�
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	}

 	//���� �����׸� ��������
 	function jsSetARAP(dAC, sANM,sACC,sACCNM){
 		document.frm.iaidx.value = dAC;
 		document.frm.selarap.value = sANM;
 	}

//-->
</script>
<style>
	FORM {display:inline;}
	</style>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>
		<form name="frm" method="get" action="index.asp">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value="">
			<input type="hidden" name="research" value="on">
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
				<td align="left">
					�μ�NEW:
					<%= drawSelectBoxDepartment("department_id", department_id) %>
					<input type="checkbox" name="inc_subdepartment" value="N" <% if (inc_subdepartment = "N") then %>checked<% end if %> > ���� �μ����� ����
					&nbsp;&nbsp;
					ī�װ�:
					<select name="icidx1" id="icidx1">
					<%
						IF menupos=1402 THEN
							'����ǰ�� �޴��� ���
					%>
						<option value="5" <%=chkIIF(icateidx1="5","selected","")%>>RP-����ǰ��</option>
						<option value="12" <%=chkIIF(icateidx1="12","selected","")%>>DR-���</option>
					<%
						ELSEIF menupos=1617 Then
							'�λ� �޴��� ���
					%>
						<option value="3" <%=chkIIF(icateidx1="3","selected","")%>>HR-�λ�</option>
						<option value="12" <%=chkIIF(icateidx1="12","selected","")%>>DR-���</option>
					<%
						ELSE
							Response.Write "<option value=""0"">--�ֻ���--</option>"
							Set clsedms = new Cedms
							clsedms.sbGetOptedmsCategory 1,0,icateidx1
							Set clsedms = nothing
						END IF
					%>
					</select>&nbsp;&nbsp;
						<select name="selDT">
							<option value="1" <%IF sDateType ="1" THEN%>selected<%END IF%>>�ۼ���</option>
							<option value="2" <%IF sDateType ="2" THEN%>selected<%END IF%>>����������</option>
						</select>:
						<input type="text" name="selSD" size="10" value="<%=searchSDate%>"><img src="/images/calicon.gif" align="absmiddle" border="0" onClick="jsPopCal('selSD');"  style="cursor:hand;">
						~
						<input type="text" name="selED" size="10" value="<%=searchEDate%>"><img src="/images/calicon.gif" align="absmiddle" border="0" onClick="jsPopCal('selED');"  style="cursor:hand;">
				 </td>
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSearch();">
				</td>
			</tr>
			<tr bgcolor="#FFFFFF" >
				<td>
					  �����ڵ�: <input type="text" name="sec" value="<%=sedmscode%>" size="10" maxlength="10">&nbsp;
					  ǰ�Ǽ���: <input type="text" name="sRnm" size="20" value="<%=sreportname%>">&nbsp;
						�����׸�: <input type="text" name="selarap" value="<%=sarap_nm%>" size="13"><input type="hidden" name="iaidx" value="<%=iarap_cd%>" >
						<input type="button" value="����" class="button" onClick="jsGetARAP();" >&nbsp;
					 �ۼ���:
					<input type="text" name="sUnm" size="8" value="<%=sUserName%>">&nbsp;
					�������:
					<select name="selPRS">
						<option value="">----</option>
						 <%sbOptReportState ireportstate%>
					</select>&nbsp;
						����:
					<select name="selOT">
						<option value="1" <%IF sOrderType ="1" THEN%>selected<%END IF%>>����������</option>
						<option value="2" <%IF sOrderType ="2" THEN%>selected<%END IF%>>�ۼ���</option>
					</select>
				</td>
			</tr>
		</table>
		</form>
	</td>
</tr>
<tr>
	<td> �˻����: <b><%=formatnumber(iTotCnt,0)%></b>  &nbsp;&nbsp;������: <b><%=iCurrpage%>/<%=iTotalPage%></b>
		<!-- ��� �� ���� -->
		<Form name="frmAct" method="post" action="erpLink_Process.asp">
		<input type="hidden" name="LTp" value="A">
		<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a"   border="0">
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td>Idx</td>
					<td>�����ڵ�</td>
					<td>ǰ�Ǽ���</td>
					<td>ǰ�Ǳݾ�</td>
					<td>�����׸�</td>
					<td>��������</td>
					<td>�ۼ���</td>
					<td>������</td>
					<td>����������</td>
					<td>�ۼ���</td>
					<td>����������</td>
					<td>�������</td>
					<td>������û����</td>
				</tr>
				<%IF isArray(arrList) THEN
					For intLoop = 0 To UBound(arrList,2)
				%>
				<tr bgcolor="#FFFFFF" align="center">
					<td><a href="javascript:jsView(<%=arrList(0,intLoop)%>);"><%=arrList(0,intLoop)%></a></td>
					<td nowrap><%=arrList(11,intLoop)%></td>
					<td align="left"><%=arrList(1,intLoop)%></td>
					<td align="right"><%=formatnumber(arrList(2,intLoop),0)%></td>
					<td align="left"><%=arrList(7,intLoop)%></td>
					<td align="left"><%=arrList(17,intLoop)%></td>
					<td nowrap><%=arrList(8,intLoop)%></td>
					<td nowrap><%=arrList(15,intLoop)%></td>
					<td nowrap><%=arrList(14,intLoop)%></td>
					<td><%=arrList(5,intLoop)%></td>
					<td><%=arrList(9,intLoop)%></td>
					<td><%=fnGetReportState(arrList(6,intLoop))%></td>
					<td><%=arrList(16,intLoop)%></td>
				</tr>
				<%
					Next
					ELSE
				%>
				<tr bgcolor="#FFFFFF">
					<td colspan="12" align="center">��ϵ� ������ �����ϴ�.</td>
				</tr>
				<%END IF%>
				</table>
				 </form>
			</td>
		</tr>
<!-- ������ ���� -->
		<%
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10

		iStartPage = (Int((iCurrPage-1)/iPerCnt)*iPerCnt) + 1

		If (iCurrPage mod iPerCnt) = 0 Then
			iEndPage = iCurrPage
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
									if Cint(ix) = Cint(iCurrPage) then
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
	</td>
</tr>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
