<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ ��  ����Ʈ
' History : 2011.05.31 ������ ����
'			2018.10.11 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpPartCls.asp"-->
<%
Dim clsPart, arrList, intLoop, arrType
Dim sOpExppartName, ipartTypeIdx, iTotCnt,iPageSize, iTotalPage,iCurrPage, incNo
	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1

 	sOpExppartName 	= requestCheckvar(Request("sOEPN"),60)
 	iPartTypeIdx	= requestCheckvar(Request("selPT"),10)
	incNo			= requestCheckvar(Request("incNo"),1)

Set clsPart = new COpExpPart
	clsPart.FPartTypeidx 	= iPartTypeIdx
	clsPart.FOpExpPartName 	= sOpExppartName
	clsPart.FRectIncNo 	= incNo
	clsPart.FCurrPage 	= iCurrPage
	clsPart.FPageSize 	= iPageSize
	arrList = clsPart.fnGetOpExpPartList
	iTotCnt = clsPart.FTotCnt

	arrType = clsPart.fnGetOpExpPartTypeList
Set clsPart = nothing

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
%>

<script type="text/javascript">
<!--
// ������ �̵�
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	}

//���ε��
function jsNewReg(){
	var winP = window.open("popPart.asp","popP","width=800,height=960,scrollbars=yes,resizable=yes");
	winP.focus();
}

//����
function jsMod(iOEP){
	var winP = window.open("popPart.asp?hidOEP="+iOEP,"popP","width=800,height=960,scrollbars=yes,resizable=yes");
	winP.focus();
}

//Ÿ�Լ���
function jsModType(){
var winP = window.open("popPartType.asp","popP","width=800,height=600,scrollbars=yes,resizable=yes");
	winP.focus();
}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
	<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
				<td align="left">
					 ����:
					 <select name="selPT">
					 <option value="">--����--</option>
					 <% sbOptPartType arrType,ipartTypeIdx%>
					 </select>
					 &nbsp;&nbsp;
					 �����ó :
					 <input type="text" name="sOEPN" size="20" maxlength="60" value="<%=sOpExppartName%>">
					 &nbsp;&nbsp;
					 <input type="checkbox" name="incNo" value="Y" <% if (incNo = "Y") then %>checked<% end if %> >
					 ������ ����

				</td>
				<td width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<tr>
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left"><input type="button" class="button" value="�űԵ��" onClick="jsNewReg();"></td>
			<td align="right"><input type="button" value="���м���" onClick="jsModType()" class="button"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<!-- ��� �� ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="FFFFFF">
				<td colspan="15">
					�˻���� : <b><%=iTotCnt%></b> &nbsp;
					������ : <b><%= iCurrPage %> / <%=iTotalPage%></b>
				</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td width="80">ǥ�ü���</td>
				<td>����</td>
				<td>�����ó</td>
				<td>�����</td>
				<td>�μ���</td>
				<td>�ڱݰ����μ�</td>
				<td>���ްŷ�ó</td> 
				<td>�����׸�</td> 
				<td>ó��</td>
			</tr>
			<%  
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2) 
				%> 
			<tr height=30 align="center" bgcolor="#<% if (arrList(20,intLoop) = True) then %>FFFFFF<% else %>DDDDDD<% end if %>">
				<td><%=arrList(12,intLoop)%></td>
				<td><%=arrList(2,intLoop)%></td>
				<td><%=arrList(3,intLoop)%></td>
				<td><%=arrList(7,intLoop)%></td>
				<td align="left">
					&nbsp;
					<%
					if arrList(21,intLoop) = 1 then
						response.write arrList(22,intLoop)
					elseif arrList(21,intLoop) > 1 then
						response.write arrList(22,intLoop) + " �� " + CStr(arrList(21,intLoop) - 1)
					end if
					%>
				</td>
				<td><%=arrList(14,intLoop)%></td> 
				<td><%=arrList(17,intLoop)%></td> 
				<td><%=arrList(15,intLoop)%></td> 
				<td><input type="button" value="������" class="button" onClick="jsMod(<%=arrList(0,intLoop)%>)"></td>
			</tr>
		<%      Next
			ELSE%>  
			<tr height=5 align="center" bgcolor="#FFFFFF">
				<td colspan="5">��ϵ� ������ �����ϴ�.</td>
			</tr>
			<%END IF%>
		</table>
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
<!-- ������ �� -->
</body>
</html>
