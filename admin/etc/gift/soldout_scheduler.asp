<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<%
Dim webImgUrl
webImgUrl = "http://webimage.10x10.co.kr"
%>
<!-- #include virtual="/lib/classes/etc/giftCls.asp"-->
<%
	Dim iCurrentpage, Giftlist, i, iTotCnt, vBody
	Set Giftlist = new ClsGift
	Giftlist.FCurrPage = "1"
	Giftlist.FGubun = ""
	Giftlist.FItemID = ""
	Giftlist.FItemName = ""
	Giftlist.FUseYN = "Y"
	Giftlist.FSoldOUT = "Y"
	Giftlist.FGiftList

	iTotCnt = Giftlist.ftotalcount

	If iTotCnt <> 0 Then
		vBody = vBody & "<table cellpadding=""3"" cellspacing=""1"" border=""1"">" & vbCrLf
		vBody = vBody & "<tr bgcolor=""#E6E6E6"">" & vbCrLf
		vBody = vBody & "	<td align=""center"">����</td>" & vbCrLf
		vBody = vBody & "	<td align=""center"">��ǰ</td>" & vbCrLf
		vBody = vBody & "	<td align=""center"">��ǰ�ڵ�</td>" & vbCrLf
		vBody = vBody & "	<td align=""center"">��ǰ��</td>" & vbCrLf
		vBody = vBody & "	<td align=""center"">���ǸŰ�</td>" & vbCrLf
		vBody = vBody & "	<td align=""center"">��ǰ��</td>" & vbCrLf
		vBody = vBody & "	<td align=""center"">��ۺ�</td>" & vbCrLf
		vBody = vBody & "	<td align=""center"">10x10<br>ǰ������</td>" & vbCrLf
		vBody = vBody & "	<td align=""center"">��뿩��</td>" & vbCrLf
		vBody = vBody & "</tr>" & vbCrLf

		For i = 0 To Giftlist.FResultCount -1

			vBody = vBody & "<tr bgcolor=""FFFFFF"">" & vbCrLf
			vBody = vBody & "	<td width=""70"" align=""center"">" & vbCrLf

			If Giftlist.FItemList(i).fgubun = "giftting" Then
				vBody = vBody & "������"
			ElseIf Giftlist.FItemList(i).fgubun = "gifticon" Then
				vBody = vBody & "����Ƽ��"
			ElseIf Giftlist.FItemList(i).fgubun = "celectory" Then
				vBody = vBody & "�����丮"
			ElseIf Giftlist.FItemList(i).fgubun = "gsisuper" Then
				vBody = vBody & "GS���̽���"
			End IF

			vBody = vBody & "	</td>" & vbCrLf
			vBody = vBody & "	<td width=""60"" align=""center""><a href=""http://www.10x10.co.kr/" & Giftlist.FItemList(i).fitemid & """ target=""_blank""><img src=""" & Giftlist.FItemList(i).fsmallimage & """ border=""0""></a></td>" & vbCrLf
			vBody = vBody & "	<td width=""60"" align=""center"">" & Giftlist.FItemList(i).fitemid & "</td>" & vbCrLf
			vBody = vBody & "	<td>" & Giftlist.FItemList(i).fitemname & "</td>" & vbCrLf
			vBody = vBody & "	<td width=""60"" align=""center"">" & FormatNumber(Giftlist.FItemList(i).ftot_sellcash,0) & "</td>" & vbCrLf
			vBody = vBody & "	<td width=""60"" align=""center"">" & FormatNumber(Giftlist.FItemList(i).fsellcash,0) & "</td>" & vbCrLf
			vBody = vBody & "	<td width=""60"" align=""center"">" & FormatNumber(Giftlist.FItemList(i).fdili_itemcost,0) & "</td>" & vbCrLf
			vBody = vBody & "	<td width=""60"" align=""center"">" & CHKIIF(Giftlist.FItemList(i).fsoldout="True","<b><font color=red>ǰ��</font></b>","�Ǹ���") & "</td>" & vbCrLf
			vBody = vBody & "	<td width=""60"" align=""center"">" & Giftlist.FItemList(i).fuseyn & "</td>" & vbCrLf
			vBody = vBody & "</tr>" & vbCrLf

		Next

		vBody = vBody & "</table>" & vbCrLf

		Call SendMail("admin@10x10.co.kr", "smsgbest@10x10.co.kr;babukim89@10x10.co.kr;amarytak@10x10.co.kr;areum531@10x10.co.kr", "������/����Ƽ��/�����丮/GS���̽��� ǰ����ǰ", vBody)
	End If
	set Giftlist = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->