<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	dim qst_sn, strSql, strResult, lp
	qst_sn = Request("qsn")

	strSql = "exec db_board.dbo.sp_Ten_Survey_Answer_Count " & qst_sn
	rsget.CursorLocation = adUseClient
	rsget.CursorType = adOpenStatic
	rsget.LockType = adLockOptimistic
	rsget.Open strSql,dbget

	'파일 시작
	strResult = "<?xml version='1.0' encoding='EUC-KR' ?>" & vbCrLf &_
				"<chart chartBottomMargin='2' formatNumberScale='0' showLimits='0' divLineThickness='1' decimalPrecision='1' chartTopMargin='2' showShadow='1' showDivLineValue='1' canvasBorderColor='CBCBCB' animation='1' baseFontColor='666666' bgColor='FFFFFF' formatNumber='1' legendBorderColor='FFFFFF' canvasPadding='3' legendBgColor='FFFFFF' chartRightMargin='2' legendPadding='2' legendShadow='0' divLineIsDashed='1' showBorder='0' legendBorderThickness='0' placeValuesInside='1' chartLeftMargin='0' canvasBorderThickness='1' baseFontSize='11' divLineDashGap='3' setAdaptiveYMin='1' plotBorderAlpha='20' >" & vbCrLf

	'내용 작성
	if Not(rsget.EOF) then
		'@카테고리 작성
		strResult = strResult & "<categories>" & vbCrLf
		Do Until rsget.EOF
			strResult = strResult & "	<category name='" & rsget("poll_content") & "' showName='1' showLine='1' />" & vbCrLf
		rsget.MoveNext
		loop
		strResult = strResult & "</categories>" & vbCrLf

		rsget.MoveFirst
		'@값 작성
		strResult = strResult & "<dataset seriesName='답변수' color='' showValues='0' >" & vbCrLf
		Do Until rsget.EOF
			strResult = strResult & "	<set value='" & rsget("ansCnt") & "' />" & vbCrLf
		rsget.MoveNext
		loop
		strResult = strResult & "</dataset>" & vbCrLf

	end if

	'파일종료
	strResult = strResult &_
		"<trendLines></trendLines>" & vbCrLf &_
		"<styles>" & vbCrLf &_
		"	<definition>" & vbCrLf &_
		"		<style name='shadow215' type='shadow' angle='215' distance='3'/>" & vbCrLf &_
		"		<style name='shadow45' type='shadow' angle='45' distance='3'/>" & vbCrLf &_
		"	</definition>" & vbCrLf &_
		"	<application>" & vbCrLf &_
		"		<apply toObject='DATAPLOTCOLUMN' styles='shadow215' />" & vbCrLf &_
		"		<apply toObject='DATAPLOTLINE' styles='shadow215' />" & vbCrLf &_
		"		<apply toObject='DATAPLOT' styles='shadow45' />" & vbCrLf &_
		"	</application>" & vbCrLf &_
		"</styles>" & vbCrLf &_
		"</chart>" & vbCrLf
	
	rsget.Close

	Response.Write strResult
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
