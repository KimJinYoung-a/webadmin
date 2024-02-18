<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체정보등록/변경
' History : 2015.05.27 강준구 생성
'			2021.12.06 한용민 수정(권한수정)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/member/partner/partnerCls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->


<%
dim ogroup,i
dim groupid, groupid_old, vTIdx

groupid = request("groupid")
vTIdx = request("tidx")

If vTIdx <> "" Then
	set ogroup = new cPartnerInfoReq
	ogroup.Ftidx = vTIdx
	ogroup.fTIdxGroupID_OLD()
	groupid = ogroup.FOneItem.Fgroupid_old
	If isNull(groupid) Then
		groupid = ogroup.FOneItem.Fgroupid
	End IF
	set ogroup = Nothing
End If

set ogroup = new CPartnerGroup
ogroup.FRectGroupid = groupid
ogroup.GetOneGroupInfo

%>

<script language='javascript'>

function PopUpcheReturnAddrOnly(groupid){
	if (groupid == "") {
		alert("그룹코드가 없습니다.");
		return;
	}


	var popwin = window.open("/admin/member/popupchereturnaddronly.asp?groupid=" + groupid,"popupchereturnaddronly","width=900 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopUpEditConfirm(g,ox)
{
	if(ox == "o")
	{
		parent.document.location.href = "/admin/member/partner/upcheinfo_edit_parent.asp?groupid=<%=groupid%>&gb=" + g + "&compnochgox=o";
	}
	else
	{
		parent.document.location.href = "/admin/member/partner/upcheinfo_edit_parent.asp?groupid=<%=groupid%>&gb=" + g + "";
	}
}

function goChild2uid(uid)
{
	var o_uid = parent.child2.document.frmupche.uid;
	var chktempp = parent.child2.document.forms["frmupche"].elements["uid"];

	if(!(fFindText(chktempp.value,uid)))
	{
		if(o_uid.value == "")
		{
			o_uid.value = o_uid.value + "" + uid;
		}
		else
		{
			o_uid.value = o_uid.value + "," + uid;
		}
	}
	else
	{
		o_uid.value = o_uid.value.replace(uid,"");
		o_uid.value = o_uid.value.replace(",,",",");

		if(o_uid.value.substring(0,1) == ",")
		{
			o_uid.value = o_uid.value.substring(1,o_uid.value.length);
		}


		if(o_uid.value.substring(o_uid.value.length-1,o_uid.value.length) == ",")
		{
			o_uid.value = o_uid.value.substring(0,o_uid.value.length-1);
		}
	}
}

function fFindText(strText,writeText)
{
	var arrText = strText.split(",");
	var trueorfalse = false;

	for(var i=0; i<arrText.length; i++)
	{
		if(writeText == arrText[i])
		{
			trueorfalse = true;
		}
	}

	return trueorfalse;
}
</script>


<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			&nbsp;* <b><font size="2">변경전 업체정보</font></b>
			(업체코드 : <%= ogroup.FOneItem.FGroupId %>&nbsp;&nbsp;업체명 : <%= ogroup.FOneItem.FCompany_name %>)
		</td>
	</tr>
	<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan=4><b>1.업체관련정보</b></td>
	</tr>
	<tr height="25">
		<td width="120" bgcolor="<%= adminColor("tabletop") %>">업체코드</td>
		<td bgcolor="#FFFFFF" width="200"><%= ogroup.FOneItem.FGroupId %></td>
		<td width="90" bgcolor="<%= adminColor("tabletop") %>">업체명</td>
		<td bgcolor="#FFFFFF" width="200">
			<%= ogroup.FOneItem.FCompany_name %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">입점브랜드ID</td>
		<td colspan="3" bgcolor="#FFFFFF" style='word-break:break-all;'>
			<% if ogroup.FOneItem.getBrandList = "" then %>
				<font color="red">현재 진행중인 브랜드가 없습니다.</font>
			<% else %>
				<%
					Dim vTmpBrand
					vTmpBrand = Replace(Trim(ogroup.FOneItem.getBrandList),"'","")
					If vTIdx = "" Then
						For i = LBound(Split(vTmpBrand,",")) To UBound(Split(vTmpBrand,","))
							Response.Write "<span onClick=""goChild2uid('" & Replace(Replace(Trim(Split(vTmpBrand,",")(i)),"<font color=#BBBBBB>",""),"</font>","") & "');"" style=""cursor:pointer"">[" & Trim(Split(vTmpBrand,",")(i)) & "]</span>"
							If i <> UBound(Split(vTmpBrand,",")) Then
								Response.Write ", "
							End If
						Next
					Else
						if C_MngPart or C_ADMIN_AUTH then
							For i = LBound(Split(vTmpBrand,",")) To UBound(Split(vTmpBrand,","))
								Response.Write "<span onClick=""goChild2uid('" & Replace(Replace(Trim(Split(vTmpBrand,",")(i)),"<font color=#BBBBBB>",""),"</font>","") & "');"" style=""cursor:pointer"">[" & Trim(Split(vTmpBrand,",")(i)) & "]</span>"
								If i <> UBound(Split(vTmpBrand,",")) Then
									Response.Write ", "
								End If
							Next
						Else
							Response.Write vTmpBrand
						End If
					End If
				%>
			<% end if %>
			<% If vTIdx = "" Then %><br><font color="blue">↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑<br><center>브랜드ID를 선택 하세요.</center></font><% End If %>
		</td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">
			<table width="100%" cellspacing="0" cellpadding="0" border="0" class="a">
			<tr>
				<td>**사업자등록정보**</td>
				<% If vTIdx = "" Then %>
				<td align="right">
					<input type="button" class="button" style="width:200px;" value="사업자등록변경(사업자번호변경O)" onClick="PopUpEditConfirm('companyreginfo','o')">&nbsp;&nbsp;
					<input type="button" class="button" style="width:200px;" value="사업자등록변경(사업자번호변경X)" onClick="PopUpEditConfirm('companyreginfo','x')">
				</td>
				<% End If %>
			</tr>
			</table>
		</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">회사명(상호)</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.FCompany_name %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">대표자</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fceoname %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">사업자번호</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.getDecCompNo %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">과세구분</td>
		<td bgcolor="#FFFFFF"><%=ogroup.FOneItem.Fjungsan_gubun%></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">사업장소재지</td>
		<td colspan="3" bgcolor="#FFFFFF" >[<%= ogroup.FOneItem.Fcompany_zipcode %>]<%= ogroup.FOneItem.Fcompany_address %> <%= ogroup.FOneItem.Fcompany_address2 %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">업태</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_uptae %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">업종</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_upjong %></td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**업체기본정보**</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">대표전화</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_tel %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">팩스</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_fax %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">사무실 주소</td>
		<td colspan="3" bgcolor="#FFFFFF" >[<%= ogroup.FOneItem.Freturn_zipcode %>]<%= ogroup.FOneItem.Freturn_address %> <%= ogroup.FOneItem.Freturn_address2 %></td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">
			<table width="100%" cellspacing="0" cellpadding="0" border="0" class="a">
			<tr>
				<td>**결제계좌정보**</td>
				<% If vTIdx = "" Then %><td align="right"><input type="button" class="button" value="결제계좌변경" onClick="PopUpEditConfirm('bankinfo','x')"></td><% End If %>
			</tr>
			</table>
		</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">거래은행</td>
		<td colspan="3" bgcolor="#FFFFFF" ><%=ogroup.FOneItem.Fjungsan_bank%></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">계좌번호</td>
		<td colspan="3" bgcolor="#FFFFFF" ><%= ogroup.FOneItem.Fjungsan_acctno %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">예금주명</td>
		<td colspan="3" bgcolor="#FFFFFF" ><%= ogroup.FOneItem.Fjungsan_acctname %></td>
	</tr>
	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">
			<table width="100%" cellspacing="0" cellpadding="0" border="0" class="a">
			<tr>
				<td>**정산일정보**</td>
				<% If vTIdx = "" Then %><td align="right"><input type="button" class="button" value="정산일변경" onClick="PopUpEditConfirm('jungsandate','x')"></td><% End If %>
			</tr>
			</table>
		</td>
	</tr>
    <tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">정산일</td>
		<td colspan="3" bgcolor="#FFFFFF" >
		온라인 : <%= ogroup.FOneItem.Fjungsan_date %>
		&nbsp;
		오프라인 : <%= ogroup.FOneItem.Fjungsan_date_off %>
		</td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**담당자정보**</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">담당자명</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_name %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_phone %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_email %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_hp %></td>
	</tr>

	<tr height="25">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">정산담당자명</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_name %></td>
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_phone %></td>
	</tr>
	<tr height="25">
		<td width="60" bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_email %></td>
		<td width="60" bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_hp %></td>
	</tr>
</table>

<%
set ogroup = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->