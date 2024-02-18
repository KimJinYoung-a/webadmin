<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/member/partner/partnerCls.asp"-->

<%
	Dim cPartner, iCurrentpage, iPageSize, i, vReqGubun, vReqName, vReqCompany, vReqGCode, vReqGCodegubun, vReqCompanyNo, vReqStatus

	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	vReqGubun		= requestCheckVar(Request("reqgubun"),20)
	vReqName		= requestCheckVar(Request("reqname"),30)
	vReqCompany		= requestCheckVar(Request("reqcompany"),100)
	vReqGCode		= requestCheckVar(Request("reqgcode"),30)
	vReqGCodegubun	= NullFillWith(requestCheckVar(Request("gcodegubun"),1),"1")
	vReqCompanyNo	= requestCheckVar(Request("reqcompanyno"),20)
	vReqStatus		= requestCheckVar(Request("reqstatus"),1)
	iPageSize 		= 15

	set cPartner = new cPartnerInfoReq
 	cPartner.FCurrPage = iCurrentpage
 	cPartner.FPageSize = iPageSize
 	cPartner.Freqgubun = vReqGubun
 	cPartner.Freqname = vReqName
 	cPartner.Freqcompany = vReqCompany
 	cPartner.Freqgcode = vReqGCode
 	cPartner.Freqgcodegubun = vReqGCodegubun
 	cPartner.FreqcompanyNo = vReqCompanyNo
 	cPartner.Freqstatus = vReqStatus
	cPartner.fRequestlist
%>

<script language='javascript'>
function goWrite(gid,tidx,g)
{
	if(g == "newcompreg"){
		var popeditconf = window.open("/admin/member/partner/upcheinfo_new.asp?groupid=" + gid + "&gb=" + g + "&tidx=" + tidx,"popeditconf","width=720,height=900,resizable=yes,scrollbars=yes");
	}else{
		var popeditconf = window.open("/admin/member/partner/upcheinfo_edit_parent.asp?groupid=" + gid + "&gb=" + g + "&tidx=" + tidx,"popeditconf","width=800,height=680,resizable=yes,scrollbars=yes");
	}
	popeditconf.focus();
}
function jsNewRegist()
{
	var newregist = window.open("popUpchelist.asp","newregist","width=800,height=680,resizable=yes,scrollbars=yes");
	newregist.focus();
}
function goUpchelist()
{
	var newregist = window.open("popUpchelist.asp?gb=search","newregist","width=800,height=680,resizable=yes,scrollbars=yes");
	newregist.focus();
}
function jsNewCompReg()
{
	var newregist = window.open("upcheinfo_edit_child_compnosearch.asp?gb=newcomp","newcompreg","width=720,height=850,resizable=yes,scrollbars=yes");
	newregist.focus();
}
function NextPage(iP)
{
	document.frmpage.iC.value = iP;
	document.frmpage.submit();
}
</script>

<form name="frm" action="index.asp" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<table class="a">
		<tr>
			<td></td>
			<td rowspan="2" style="padding:0 0 0 80px;" align="right" valign="top"><input type="submit" value=" ��  �� " class="button" style="width:70px;height:50px;" onfocus="this.blur();"></td>
		</tr>
		<tr>
			<td>
				<table class="a">
				<tr>
					<td>
						<select name="reqgubun" class="select">
							<option value="">-��û����-</option>
							<option value="companyreginfo" <%=CHKIIF(vReqGubun="companyreginfo","selected","")%>>����ڵ������</option>
							<option value="bankinfo" <%=CHKIIF(vReqGubun="bankinfo","selected","")%>>������������</option>
							<option value="jungsandate" <%=CHKIIF(vReqGubun="jungsandate","selected","")%>>����������</option>
						</select>
						&nbsp;&nbsp;
						��û�� : <input type="text" name="reqname" value="<%=vReqName%>" size="10">
						&nbsp;&nbsp;
						��ü�ڵ� :
							<input type="radio" name="gcodegubun" value="1" <%=CHKIIF(vReqGCodegubun="1","checked","")%>><input type="text" name="reqgcode" value="<%=vReqGCode%>" size="7">
							<input type="button" value="��ü" class="button" onClick="goUpchelist()">
							<input type="radio" name="gcodegubun" value="2" <%=CHKIIF(vReqGCodegubun="2","checked","")%>>���ο����ڹ�ȣ
						&nbsp;&nbsp;
					</td>
				</tr>
				<tr>
					<td>
						ȸ���(��ȣ) : <input type="text" name="reqcompany" value="<%=vReqCompany%>" size="20">
						&nbsp;&nbsp;
						����ڹ�ȣ : <input type="text" name="reqcompanyno" value="<%=vReqCompanyNo%>" size="15">
						&nbsp;&nbsp;
						<select name="reqstatus" class="select">
							<option value="">-�������-</option>
							<option value="1" <%=CHKIIF(vReqStatus="1","selected","")%>>��û</option>
							<option value="2" <%=CHKIIF(vReqStatus="2","selected","")%>>�۾���</option>
							<option value="3" <%=CHKIIF(vReqStatus="3","selected","")%>>����Ϸ�</option>
							<option value="5" <%=CHKIIF(vReqStatus="5","selected","")%>>��ϿϷ�</option>
							<option value="0" <%=CHKIIF(vReqStatus="0","selected","")%>>����</option>
						</select>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<br>
<input type="button" class="button" value="����ڵ��(�ű�)" onClick="jsNewCompReg();">
<input type="button" class="button" value="�����û" onClick="jsNewRegist();">
<br><br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="60" align="center">��û��No.</td>
	<td width="100" align="center">��û����</td>
	<td width="80" align="center">��û��</td>
	<td width="300" align="center">������ ��ü����</td>
	<td width="110" align="center">��ü�ڵ�</td>
	<td align="center">ȸ���(��ȣ)</td>
	<td width="90" align="center">����ڹ�ȣ</td>
	<td width="80" align="center">�������</td>
	<td width="150" align="center">��û��</td>
	<td width="150" align="center">������</td>
</tr>
<% If cPartner.FresultCount > 0 Then %>
	<% For i=0 To cPartner.FresultCount-1 %>
	<tr bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
		<td align="center" style="cursor:pointer" onClick="goWrite('<%=cPartner.FItemList(i).Fgroupid%>','<%=cPartner.FItemList(i).Ftidx%>','<%=cPartner.FItemList(i).Fgubun%>')"><%=cPartner.FItemList(i).Ftidx%></td>
		<td align="center"><%=RequestDocumentName(cPartner.FItemList(i).Fgubun)%></td>
		<td align="center"><%=cPartner.FItemList(i).Fusername%></td>
		<td style="padding-left:10px;">
		<%
			If cPartner.FItemList(i).Fgroupid_old <> "" Then
				Response.Write cPartner.FItemList(i).Fgroupid_old & "&nbsp;&nbsp;"
				Response.Write cPartner.FItemList(i).Fcompany_name_old & "&nbsp;&nbsp;"
				Response.Write socialnoReplace(cPartner.FItemList(i).Fcompany_no_old)
			End If
		%>
		</td>
		<td align="center"><%=CHKIIF(cPartner.FItemList(i).Fgroupid="","���ο����ڹ�ȣ",cPartner.FItemList(i).Fgroupid)%></td>
		<td align="center"><%=cPartner.FItemList(i).Fcompany_name%></td>
		<td align="center"><%=socialnoReplace(cPartner.FItemList(i).Fcompany_no)%></td>
		<td align="center">
		<%
			If cPartner.FItemList(i).Fgubun = "newcompreg" AND cPartner.FItemList(i).Fstatus = "3" Then
				Response.Write RequestStateName("5")
			Else
				Response.Write RequestStateName(cPartner.FItemList(i).Fstatus)
			End If
		%>
		</td>
		<td align="center">
			<%=cPartner.FItemList(i).Fregdate%>
		</td>
		<td align="center">
			<% if (cPartner.FItemList(i).Fstatus = 3) then %><%= cPartner.FItemList(i).Flastupdate %><% end if %>
		</td>
	</tr>
	<% next %>
<% Else %>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="center" colspan="15">[�����Ͱ� �����ϴ�.]</td>
</tr>
<% End If %>
</table>

<form name="frmpage" method="get" action="index.asp" style="margin:0px;">
<input type="hidden" name="iC" value="">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<input type="hidden" name="reqgubun" value="<%=vReqGubun%>">
<input type="hidden" name="reqname" value="<%=vReqName%>">
<input type="hidden" name="reqgcode" value="<%=vReqGCode%>">
<input type="hidden" name="reqcompany" value="<%=vReqCompany%>">
<input type="hidden" name="gcodegubun" value="<%=vReqGCodegubun%>">
<input type="hidden" name="reqcompanyno" value="<%=vReqCompanyNo%>">
<input type="hidden" name="reqstatus" value="<%=vReqStatus%>">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr valign="bottom" height="25">
    <td valign="bottom" align="center">
    	<% if cPartner.HasPreScroll then %>
		<a href="javascript:NextPage('<%= cPartner.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + cPartner.StartScrollPage to cPartner.FScrollCount + cPartner.StartScrollPage - 1 %>
			<% if i>cPartner.FTotalpage then Exit for %>
			<% if CStr(iCurrentpage)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if cPartner.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>

    </td>
</tr>
</table>
</form>

<% set cPartner = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
