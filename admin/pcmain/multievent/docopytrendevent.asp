<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim sqlStr
dim idx, CreateIDX
idx	= Request("idxarr")

    'TrendEvent ����
    sqlStr = "Insert Into db_sitemaster.dbo.tbl_mobile_main_enjoyevent_new" & VbCrlf
    sqlStr = sqlStr & " (evt_code, addtype, linktype, linkurl, evttitle, evttitle2, startdate, enddate, evtstdate" & VbCrlf
    sqlStr = sqlStr & ", evteddate, regdate, lastupdate, adminid, lastadminid, isusing" & VbCrlf
    sqlStr = sqlStr & ", ordertext, sortnum, sale_per, coupon_per, tag_only)" & VbCrlf
    sqlStr = sqlStr & " select eventid, dispOption, 1, linkurl, maincopy, subcopy, startdate, enddate, evtstdate" & VbCrlf
    sqlStr = sqlStr & " , evteddate, regdate, lastupdate, adminid, lastadminid, isusing, ordertext" & VbCrlf
    sqlStr = sqlStr & " , sortnum, sale_per, coupon_per, tag_only" & VbCrlf
    sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_pcmain_enjoyevent where idx in (" & idx & ")" & VbCrlf
    dbget.Execute(sqlStr)

%>
<script>
<!--
	// ������� ����
	alert("�����߽��ϴ�.");
	window.opener.document.location.href = window.opener.document.URL;    // �θ�â ���ΰ�ħ
	 self.close();        // �˾�â �ݱ�
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->