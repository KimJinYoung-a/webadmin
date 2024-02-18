<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="EUC-KR" %>
<%
'###########################################################
' Description :  �ٹ����� ������
' History : 2018.04.27 �̻� ����(���Ϸ� ���� ���� ���Ϸ��� �߼� ���� ����. ���� �������� ����.)
'			2019.06.24 ������ ����(���ø� ��� �ű� �߰�)
'			2020.05.28 �ѿ�� ����(TMS ���Ϸ� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAppNotiopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mailzinenewcls.asp"-->
<!-- #include virtual="/lib/classes/search/searchMobileCls.asp"-->
<!-- #include virtual="/lib/classes/search/itemCls.asp"-->
<%
	dim sqlStr, result, oaward, oaward2, ix
	result=""
	set oaward = new SearchItemCls
		oaward.FListDiv 			= "newlist"
		oaward.FRectSortMethod	    = "be"
		oaward.FRectSearchFlag 	= "newitem"
		oaward.FPageSize 			= 10
		oaward.FCurrPage 			= 1
		oaward.FSellScope			= "Y"
		oaward.FScrollCount 		= 1
		oaward.FRectSearchItemDiv   ="D"
		oaward.FminPrice			= 20000
		oaward.FSalePercentLow = 0.89
		oaward.getSearchList
		If oaward.FResultCount>0 Then
			For ix=0 to oaward.FResultCount-1
				response.write Cstr(oaward.FItemList(ix).FItemID) & "<br>"
			Next
		end if
	set oaward = Nothing

	set oaward2 = new SearchItemCls
		oaward2.FListDiv 			= "bestlist"
		oaward2.FRectSortMethod	    = "be"
		oaward2.FPageSize 			= 10
		oaward2.FCurrPage 			= 1
		oaward2.FSellScope			= "Y"
		oaward2.FScrollCount 		= 1
		oaward2.FRectSearchItemDiv   ="D"
		oaward2.FminPrice			= 20000
		oaward2.FawardType			= "period"
		oaward2.getSearchList
		If oaward2.FResultCount>0 Then
			For ix=0 to oaward2.FResultCount-1
				response.write Cstr(oaward2.FItemList(ix).FItemID) & "<br>"
			Next
		end if
	set oaward2 = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAppNoticlose.asp" -->