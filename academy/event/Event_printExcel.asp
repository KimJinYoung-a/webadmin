<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->
<%
	'/// �������¸� Ms-Excel�� ���� ///
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition","Attachment;"

	'// ���� ���� //
	dim evtId

	dim oPart, lp

	'// �Ķ���� ���� //
	evtId = RequestCheckvar(request("evtId"),10)

	'// Ŭ���� ����
	set oPart = new CPart
	oPart.FRectevtId = evtId

	oPart.GetPartAllList


	'// �ʵ尪 ��� //
	Response.write ("�̺�Ʈ ��ȣ : " & chr(9))
	Response.write (evtId & chr(9))
	Response.write (chr(9))
	Response.write ("������ �� :" & chr(9))
	Response.write (oPart.FTotalCount & chr(9))
	Response.Write (chr(13) & chr(10))
	Response.Write (chr(13) & chr(10))

	Response.write ("��ȣ" & chr(9))
	Response.write ("���̵�" & chr(9))
	Response.write ("ȸ�����" & chr(9))
	Response.write ("�̸�" & chr(9))
	Response.write ("����1" & chr(9))
	Response.write ("����2" & chr(9))
	Response.write ("������" & chr(9))
	Response.write ("�����Ͻ�" & chr(9))
	Response.write ("���ž�(6M)" & chr(9))
	Response.write ("ȸ��������" & chr(9))
	Response.write ("��÷Ƚ��" & chr(9))
	Response.Write (chr(13) & chr(10))

	if oPart.FTotalCount>0 then

	'@@ ������ ����
	for lp=0 to oPart.FTotalCount - 1

		Response.write (oPart.FPartList(lp).FprtId & chr(9))
		Response.write (oPart.FPartList(lp).FprtUserId & chr(9))
		Response.write (oPart.FPartList(lp).FprtUserLevel & chr(9))
		Response.write (oPart.FPartList(lp).FprtUserNm & chr(9))
		Response.write (Replace(db2html(oPart.FPartList(lp).FprtCont1), chr(13)&chr(10), " ") & chr(9))
		Response.write (Replace(db2html(oPart.FPartList(lp).FprtCont2), chr(13)&chr(10), " ") & chr(9))
		Response.write (oPart.FPartList(lp).FprtCnt & chr(9))
		Response.write (oPart.FPartList(lp).FprtDate & chr(9))
		Response.write (FormatNumber(oPart.FPartList(lp).FsixMonthOrder,0) & chr(9))
		Response.write (oPart.FPartList(lp).FregDate & chr(9))
		Response.write (oPart.FPartList(lp).FprizeCnt & chr(9))
		Response.Write (chr(13) & chr(10))

	next

	end if

set oPart = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->