<%

'// ��, ��, ����Ͽ� ���������� �����Ϸ��� /imgstatic/lib/badgelib.asp ����
'// ���ο� ���� ����(�α� ��� + ��ŰX = ��α��ν� �˸� ����)

'' /lib/util/myalarmlib.asp
'' /lib/util/scm_myalarmlib.asp

''  msgdiv 	����					�Է� ����
'' ===========================================================================
''  000		��ü�˸�
'' 	001		�ű԰�������			ȸ�����Խ�
'' 	002		��������				�� 1ȸ(����)
'' 	003		��ٱ��� ��ǰ �̺�Ʈ	MyAlarm_CheckNewMyAlarm �����
'' 	004		���� ��ǰ �̺�Ʈ		MyAlarm_CheckNewMyAlarm �����
'' 	005		1:1 ���				�� 1ȸ(����,������¥��)
'' 	006		��ǰ QnA				�� 1ȸ(����,������¥��)
'' 	007		�̺�Ʈ ��÷				�� 1ȸ
''  901		���ɻ�ǰ ����
''  902		�����̺�Ʈ ����

Function MyAlarm_InsertMyAlarm_SCM(userid, msgdiv, title, subtitle, contents, wwwTargetURL)
	dim strSql, i

	'// �ߺ��Է� ����(1:!���, ��ǰQNA�� ���)
	strSql = " [db_my10x10].[dbo].[usp_Ten_MyAlarm_ProcInsertLOG] ('" + CStr(userid) + "', '" + CStr(msgdiv) + "', '" + CStr(html2db(title)) + "', '" + CStr(html2db(subtitle)) + "', '" + CStr(html2db(contents)) + "', '" + CStr(wwwTargetURL) + "') "
	dbget.Execute strSql
End Function

%>
