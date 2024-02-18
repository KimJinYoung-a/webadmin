<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual = "/admin/incSessionAdmin.asp" -->
<!-- #include virtual = "/lib/util/htmllib.asp" -->
<!-- #include virtual = "/lib/db/dbopen.asp" -->
<!-- #include virtual = "/admin/lib/adminbodyhead.asp" -->
<!-- #include virtual = "/lib/function.asp" -->
<!-- #Include virtual = "/lib/classes/totalpoint/totalpointCls.asp" -->

<%
	Dim iTotCnt, arrList, intLoop, arrFileList, i, arrCardList
	Dim iPageSize, iCurrentpage ,iDelCnt, vParam
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
	
	Dim vUserID, vUserSeq, vUserName, vJumin1, vJumin2_Enc, vCardNo, vPoint, vGrade, vSexFlag, vTelNo, vHpNo, vSearchCardNo
	Dim vZipCode, vAddress, vAddressDetail, vEmail, vEmailYN, vSMSYN, vUserStatus, vLastUpdate, vRegdate, vShopName, vUseYN, vTotalPoint
	
	vUserSeq		= NullFillWith(requestCheckVar(Request("userseq"),10),"")
	vUserID			= NullFillWith(requestCheckVar(Request("userid"),32),"")
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	vUserName		= NullFillWith(requestCheckVar(Request("username"),20),"")
	vCardNo			= NullFillWith(requestCheckVar(Request("cardno"),20),"")
	vUseYN			= NullFillWith(requestCheckVar(Request("useyn"),20),"")
	
	vSearchCardNo	= NullFillWith(requestCheckVar(Request("searchcardno"),20),"")
	
	vParam = "&iC="&iCurrentpage&"&username="&vUserName&"&cardno="&vCardNo&"&userid="&vUserID&"&useyn="&vUseYN&"&pagesize="&iPageSize&""
	'<!--  //-->
		
	If vUserSeq = "" Then
		Response.Write "<script>alert('�߸��� ����Դϴ�.');window.close();</script>"
		Response.End
	Else

		Dim totalpointView
		Set totalpointView = New TotalPoint
		totalpointView.FUserSeq = vUserSeq
		totalpointView.GetTotalPointDetail
	
		If totalpointView.FTotCnt = "0" Then
			Response.Write "<script>alert('�߸��� ����Դϴ�.');window.close();</script>"
			dbget.close()
			Response.End
		ElseIf totalpointView.FTotCnt > "1" Then
			'Response.Write "<script>alert('�ߺ��Ǿ� �� ȸ���Դϴ�.');</script>"
		End If
		
		vUserName		= totalpointView.FUserName
		vJumin1			= totalpointView.FJumin1
		vJumin2_Enc		= totalpointView.FJumin2_Enc
		vCardNo			= totalpointView.FCardNo
		vPoint			= totalpointView.FPoint
		vGrade			= totalpointView.FGrade
		vSexFlag		= totalpointView.FSexFlag
		vTelNo			= totalpointView.FTelNo
		vHpNo			= totalpointView.FHpNo
		vZipCode		= totalpointView.FZipCode
		vAddress		= totalpointView.FAddress
		vAddressDetail	= totalpointView.FAddressDetail
		vEmail			= totalpointView.FEmail
		vEmailYN		= totalpointView.FEmailYN
		vSMSYN			= totalpointView.FSMSYN
		vUserStatus		= totalpointView.FUserStatus
		vLastUpdate		= totalpointView.FLastUpdate
		vRegdate		= totalpointView.FRegdate
		vShopName		= totalpointView.FShopName

		totalpointView.FCardNo = vSearchCardNo
		arrList = totalpointView.GetTotalPointLogList()
		vTotalPoint		= totalpointView.FTotalPoint
		
		arrCardList = totalpointView.GetMemberCardList()
		
		set totalpointView = Nothing
	End If
%>



<!-- �˻� ����
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			���̵� : <input type="text" class="text" name="userid" value="" size="12">
			&nbsp;
			ī���ȣ : <input type="text" class="text" name="cardno" value="">
			<br>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>

	</form>
</table>
�˻� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="4">
			<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
			<tr>
				<td width="50%"><img src="/images/icon_arrow_link.gif" valign="absbottom">&nbsp;<b>�⺻����</b></td>
				<td width="50%" align="right"><input type="button" class="button" value="�� ��" onClick="window.close()"></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">ȸ����ȣ</td>
		<td width="300" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vUserSeq%></td>
		<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
		<%
			If vGrade <> "0" Then
				Response.Write "[���̶��Ư��ȸ��] "
			End If
			Response.Write vUserName & "&nbsp;"
			If vSexFlag = "1" Then
				Response.Write "(��)"
			ElseIf vSexFlag = "2" Then
				Response.Write "(��)"
			Else
				Response.Write "(" & vSexFlag & ")"
			End If
		%>
		</td>
	</tr>

	<tr>
		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">��ȭ��ȣ</td>
		<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vTelNo%></td>
		<td align="center"  bgcolor="<%= adminColor("tabletop") %>"></td>
		<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"></td>
	</tr>
	<tr>
		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�ڵ�����ȣ</td>
		<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vHpNo%></td>
		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">SMS���ſ���</td>
		<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vSMSYN%></td>
	</tr>
	<tr>
		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�̸���</td>
		<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vEmail%></td>
		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">MAIL���ſ���</td>
		<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vEmailYN%></td>
	</tr>
	<tr>
		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�� ��</td>
		<td colspan="3" bgcolor="#FFFFFF" style="padding: 0 0 0 5">[<%=vZipCode%>] <%=vAddress%>&nbsp;<%=vAddressDetail%></td>
	</tr>
	<tr>
		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">���ī���ȣ</td>
		<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vCardNo%></td>
		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">����Ʈ</td>
		<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vTotalPoint%> Point</td>
	</tr>
	</tr>
	<!--
	<tr>
		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">ȸ������</td>
		<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=UserStatus(vUserStatus)%>
		</td>
	</tr>
	<tr>
		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vShopName%> (�����:<%=vRegdate%>)</td>
	</tr>
	//-->
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="point_detail.asp">
<input type="hidden" name="userseq" value="<%=vUserSeq%>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<img src="/images/icon_arrow_link.gif" valign="absbottom">&nbsp;<b>������</b>
			&nbsp;
			<select name="searchcardno" class="select" onChange="frm.submit();">
				<option value="">ī����ü</option>
			<%
				IF isArray(arrCardList) THEN
					For intLoop =0 To UBound(arrCardList,2)

						Response.Write "<option value='" & arrCardList(0,intLoop) & "'"
						If vSearchCardNo = arrCardList(0,intLoop) Then
							Response.Write " selected"
						End If
						Response.Write ">" & arrCardList(0,intLoop) & "</option>"

					Next
				End If
			%>
			</select>
		</td>
	</tr>
</form>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="70">ī�屸��</td>
		<td width="140">�����</td>
		<td width="100">ī���ȣ</td>
		<td width="120">�ŷ�����</td>
		<td width="50">����Ʈ</td>
		<td>�����ֹ���ȣ</td>
	</tr>
	<%
		IF isArray(arrList) THEN
			For intLoop =0 To UBound(arrList,2)
	%>
	
	<tr align="center" bgcolor="#FFFFFF">
		<td>
			<% If Left(arrList(0,intLoop),4) = "1010" Then %>
				POINT1010
			<% ElseIf Left(arrList(0,intLoop),4) = "3253" Then %>
				���̶��
			<% Else %>
				��������
			<% End If %>
		</td>
		<td><%=arrList(7,intLoop)%></td>
		<td><%=arrList(0,intLoop)%></td>
		<td>
			<%
				'### ����Ʈ 0�̰� code�� 3(����Ʈ�̰�)�϶� ī�������� ��Ÿ��.
				If arrList(1,intLoop) = "0" AND arrList(8,intLoop) = "3" Then
					Response.Write arrList(4,intLoop)
				Else
					Response.Write arrList(2,intLoop)
				End IF
			%>
		</td>
		<td><%=arrList(1,intLoop)%></td>
		<td><%=arrList(5,intLoop)%></td>
	</tr>
						
	<%
			Next
		End If
	%>
</table>
<!--
<p>

<table border="0" width="100%" cellpadding="0" cellspacing="0" class="a">
<tr><td align="center"><input type="button" class="button" value="�� ��" onClick="window.close()"></td></tr>
</table>
-->

<!-- #include virtual="/lib/db/dbclose.asp" -->
