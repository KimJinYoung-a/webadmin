<% option Explicit %>
<!-- #include virtual="/designer/incSessionDesignerNoCache.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
'#######################################################
'   History : 2012.10.25 ������ ���� - Tabs Upload ���
'	Description : ���� �ٿ�ε� ó��
'#######################################################

	dim dfPath, FileNo, FileName, DestinationFolder, arrFileName

	FileNo = getNumeric(requestCheckVar(Request("fn"),8))

	if FileNo="" then Response.End
	 
	'// ǰ���ȣ�� �ٿ�ε� ���ϸ�
	Select Case FileNo
		Case "01": FileName = "�ٹ�����_��ǰǰ��_01_�Ƿ�.xls"
		Case "02": FileName = "�ٹ�����_��ǰǰ��_02_����_�Ź�.xls"
		Case "03": FileName = "�ٹ�����_��ǰǰ��_03_����.xls"
		Case "04": FileName = "�ٹ�����_��ǰǰ��_04_�м���ȭ(����_��Ʈ_�׼�����).xls"
		Case "05": FileName = "�ٹ�����_��ǰǰ��_05_ħ����_Ŀư.xls"
		Case "06": FileName = "�ٹ�����_��ǰǰ��_06_����(ħ��_����_��ũ��_DIY��ǰ).xls"
		Case "07": FileName = "�ٹ�����_��ǰǰ��_07_������(TV��).xls"
		Case "08": FileName = "�ٹ�����_��ǰǰ��_08_������_������ǰ(�����_��Ź��_�ı⼼ô��_���ڷ�����).xls"
		Case "09": FileName = "�ٹ�����_��ǰǰ��_09_��������(������_��ǳ��).xls"
		Case "10": FileName = "�ٹ�����_��ǰǰ��_10_�繫����(��ǻ��_��Ʈ��_������).xls"
		Case "11": FileName = "�ٹ�����_��ǰǰ��_11_���б��(������ī�޶�_ķ�ڴ�).xls"
		Case "12": FileName = "�ٹ�����_��ǰǰ��_12_��������(MP3_���ڻ���_��).xls"
		Case "13": FileName = "�ٹ�����_��ǰǰ��_13_�޴���.xls"
		Case "14": FileName = "�ٹ�����_��ǰǰ��_14_������̼�.xls"
		Case "15": FileName = "�ٹ�����_��ǰǰ��_15_�ڵ�����ǰ(�ڵ�����ǰ_��Ÿ_�ڵ�����ǰ).xls"
		Case "16": FileName = "�ٹ�����_��ǰǰ��_16_�Ƿ���.xls"
		Case "17": FileName = "�ٹ�����_��ǰǰ��_17_�ֹ��ǰ.xls"
		Case "18": FileName = "�ٹ�����_��ǰǰ��_18_ȭ��ǰ.xls"
		Case "19": FileName = "�ٹ�����_��ǰǰ��_19_�ͱݼ�_����_�ð��.xls"
		Case "20": FileName = "�ٹ�����_��ǰǰ��_20_��ǰ(����깰).xls"
		Case "21": FileName = "�ٹ�����_��ǰǰ��_21_������ǰ.xls"
		Case "22": FileName = "�ٹ�����_��ǰǰ��_22_�ǰ���ɽ�ǰ.xls"
		Case "23": FileName = "�ٹ�����_��ǰǰ��_23_�����ƿ�ǰ.xls"
		Case "24": FileName = "�ٹ�����_��ǰǰ��_24_�Ǳ�.xls"
		Case "25": FileName = "�ٹ�����_��ǰǰ��_25_��������ǰ.xls"
		Case "26": FileName = "�ٹ�����_��ǰǰ��_26_����.xls"
		Case "27": FileName = "�ٹ�����_��ǰǰ��_27_ȣ��_���_����.xls"
		Case "28": FileName = "�ٹ�����_��ǰǰ��_28_������Ű��.xls"
		Case "29": FileName = "�ٹ�����_��ǰǰ��_29_�װ���.xls"
		Case "30": FileName = "�ٹ�����_��ǰǰ��_30_�ڵ���_�뿩_����(����ī).xls"
		Case "31": FileName = "�ٹ�����_��ǰǰ��_31_��ǰ�뿩_����(������,��,����û����_��).xls"
		Case "32": FileName = "�ٹ�����_��ǰǰ��_32_��ǰ�뿩_����(����,���ƿ�ǰ,����ǰ_��).xls"
		Case "33": FileName = "�ٹ�����_��ǰǰ��_33_������_������(����,����,���ͳݰ���_��).xls"
		Case "34": FileName = "�ٹ�����_��ǰǰ��_34_��ǰ��_����.xls"
		Case "35": FileName = "�ٹ�����_��ǰǰ��_35_��Ÿ.xls"
		Case "900": FileName = "�ٹ�����_�ؿܹ��_����.xls"
		Case "990": FileName = "�ٹ�����_��ǰ�����������_����.xls"
	End Select

	On Error Resume Next
	'���� �ٿ�ε�
	dfPath = server.mappath("/designer/itemmaster/itemInfoFile/")
	DestinationFolder = dfPath & "/" & fileName

	'// �ٿ�ε� ���۳�Ʈ ���� �� �ٿ�ε� ����
	Dim oDownload

	IF (application("Svr_Info")	= "Dev") then
	    Set oDownload = Server.CreateObject("TABS.Download")	   '' - TEST
	ELSE
	    Set oDownload = Server.CreateObject("TABSUpload4.Download")	''REAL
	END IF
	 
	oDownload.FilePath = DestinationFolder
	oDownload.FileName = fileName
	oDownload.TransferFile True

	Set oDownload = Nothing

    IF (ERR) then
		response.write "<script>alert('�˼��մϴ�. ������ �غ����Դϴ�.')</script>"
		response.write "<script>self.close();</script>"
    End if
    On Error Goto 0
%>