<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<%
'###########################################################
' Description : 위임전결규정
' Hieditor : 정윤정 생성
'			 2018.06.11 한용민 수정
'###########################################################
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 12">
<link rel=File-List href="weim.files/filelist.xml">
<style id="weim_25868_Styles">
<!--table {        mso-displayed-decimal-separator:"\."; mso-displayed-thousand-separator:"\,";}
	.font525868 {
		color:windowtext;
		font-size:8.0pt;
		font-weight:400;
		font-style:normal;
		text-decoration:none;
		font-family:"맑은 고딕", monospace;
		mso-font-charset:129;
	}
	.xl6425868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:400;
		font-style:normal;
		text-decoration:none;
		font-family:"맑은 고딕", monospace;
		mso-font-charset:129;
		mso-number-format:"\@"; text-align:general;
		vertical-align:middle;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:nowrap;
	}
	.xl6525868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:none;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl6625868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:.5pt solid gray;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl6725868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:left;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:none;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl6825868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:left;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:.5pt solid gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl6925868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:.5pt dotted gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl7025868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:.5pt dotted gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl7125868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl7225868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl7325868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:left;
		vertical-align:middle;
		border-top:none;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:none;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl7425868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:left;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:.5pt solid gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl7525868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:.5pt dotted gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl7625868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:.5pt dotted gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl7725868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl7825868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl7925868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:general;
		vertical-align:middle;
		border-top:none;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:none;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl8025868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:left;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:.5pt solid gray;
		border-bottom:none;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl8125868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:.5pt dotted gray;
		border-bottom:none;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl8225868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:.5pt dotted gray;
		border-bottom:none;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl8325868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:none;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl8425868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:left;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt solid gray;
		border-bottom:none;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl8525868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt dotted gray;
		border-bottom:none;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl8625868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt dotted gray;
		border-bottom:none;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl8725868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:none;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl8825868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:none;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl8925868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:general;
		vertical-align:middle;
		border-top:none;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl9025868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:left;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt solid gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl9125868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:400;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:general;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt dotted gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl9225868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:400;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:general;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt dotted gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl9325868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt dotted gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl9425868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl9525868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl9625868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:left;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:none;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl9725868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt dotted gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl9825868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:left;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:.5pt solid gray;
		border-bottom:none;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl9925868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:.5pt dotted gray;
		border-bottom:none;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl10025868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:.5pt dotted gray;
		border-bottom:none;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl10125868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:none;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl10225868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:none;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl10325868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:general;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt dotted gray;
		border-bottom:none;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl10425868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:general;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt dotted gray;
		border-bottom:none;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl10525868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt solid gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl10625868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:.5pt solid gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl10725868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:red;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:general;
		vertical-align:middle;
		border-top:none;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:none;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl10825868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:left;
		vertical-align:middle;
		border-top:none;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl10925868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:general;
		vertical-align:middle;
		border-top:none;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:1.0pt solid gray;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl11025868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:red;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:general;
		vertical-align:middle;
		border-top:none;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:1.0pt solid gray;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl11125868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:left;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:.5pt solid gray;
		border-bottom:1.0pt solid gray;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl11225868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:.5pt dotted gray;
		border-bottom:1.0pt solid gray;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl11325868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:.5pt dotted gray;
		border-bottom:1.0pt solid gray;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl11425868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:1.0pt solid gray;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl11525868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:1.0pt solid gray;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl11625868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:left;
		vertical-align:middle;
		border-top:1.0pt solid gray;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:none;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl11725868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:left;
		vertical-align:middle;
		border-top:1.0pt solid gray;
		border-right:.5pt solid gray;
		border-bottom:none;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl11825868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid gray;
		border-right:.5pt dotted gray;
		border-bottom:none;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl11925868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid gray;
		border-right:.5pt dotted gray;
		border-bottom:none;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl12025868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid gray;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:none;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl12125868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid gray;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:none;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl12225868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:general;
		vertical-align:middle;
		border-top:none;
		border-right:none;
		border-bottom:none;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl12325868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:left;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt solid gray;
		border-bottom:none;
		border-left:1.0pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl12425868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:.5pt solid gray;
		border-bottom:none;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl12525868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt solid gray;
		border-bottom:none;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl12625868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:.5pt solid gray;
		border-bottom:none;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl12725868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:none;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl12825868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:.5pt solid gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:none;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl12925868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:none;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:none;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl13025868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:none;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl13125868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:.5pt solid gray;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl13225868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt solid gray;
		border-bottom:.5pt solid gray;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl13325868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:none;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl13425868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:1.0pt solid #7F7F7F;
		border-bottom:.5pt solid gray;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl13525868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt dotted gray;
		border-bottom:.5pt solid gray;
		border-left:.5pt solid gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl13625868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt dotted gray;
		border-bottom:.5pt solid gray;
		border-left:.5pt dotted gray;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl13725868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:.5pt solid gray;
		border-right:.5pt solid gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl13825868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:underline;
		text-underline-style:single;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:.5pt solid gray;
		border-bottom:1.0pt solid #7F7F7F;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl13925868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:underline;
		text-underline-style:single;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt solid gray;
		border-bottom:none;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl14025868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid #7F7F7F;
		border-right:.5pt solid gray;
		border-bottom:1.0pt solid gray;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl14125868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:1.0pt solid gray;
		border-right:.5pt solid gray;
		border-bottom:none;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	.xl14225868 {
		padding-top:1px;
		padding-right:1px;
		padding-left:1px;
		mso-ignore:padding;
		color:black;
		font-size:9.0pt;
		font-weight:700;
		font-style:normal;
		text-decoration:none;
		font-family:굴림, monospace;
		mso-font-charset:129;
		mso-number-format:General;
		text-align:center;
		vertical-align:middle;
		border-top:none;
		border-right:.5pt solid #7F7F7F;
		border-bottom:none;
		border-left:1.0pt solid #7F7F7F;
		mso-background-source:auto;
		mso-pattern:auto;
		white-space:normal;
	}
	ruby {
		ruby-align:left;
	}
	rt {
		color:windowtext;
		font-size:8.0pt;
		font-weight:400;
		font-style:normal;
		text-decoration:none;
		font-family:"맑은 고딕", monospace;
		mso-font-charset:129;
		mso-char-type:none;
	}
-->
</style>
</head>
<body>
<div id="weim_25868" align=center x:publishsource="Excel">
<table border=0 cellpadding=0 cellspacing=0 width=808 class=xl6425868 style='border-collapse:collapse;table-layout:fixed;width:606pt'>
	<tr>
		<td><img src="http://imgstatic.10x10.co.kr/offshop/sample/photo/Regulations_for_Decision_Making_v1.jpg"></td>
	</tr>
<!--<col class=xl6425868 width=108 style='mso-width-source:userset;mso-width-alt:3456;width:81pt'>
<col class=xl6425868 width=267 style='mso-width-source:userset;mso-width-alt:8544;width:200pt'>
<col class=xl6425868 width=51 style='mso-width-source:userset;mso-width-alt:1632;width:38pt'>
<col class=xl6425868 width=43 style='mso-width-source:userset;mso-width-alt:1376;width:32pt'>
<col class=xl6425868 width=51 style='mso-width-source:userset;mso-width-alt:1632;width:38pt'>
<col class=xl6425868 width=38 style='mso-width-source:userset;mso-width-alt:1216;width:29pt'>
<col class=xl6425868 width=108 style='mso-width-source:userset;mso-width-alt:3456;width:81pt'>
<col class=xl6425868 width=142 style='mso-width-source:userset;mso-width-alt:4544;width:107pt'>
	<tr height=17 style='height:12.75pt'>
		<td colspan=2 height=17 class=xl12725868 dir=LTR width=375 style='border-right:.5pt solid gray;height:12.75pt;width:281pt'>구분</td>
		<td colspan=4 class=xl12725868 dir=LTR width=183 style='border-right:1.0pt solid #7F7F7F;border-left:none;width:137pt'>전결권자<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td rowspan=3 class=xl8325868 dir=LTR width=108 style='border-bottom:.5pt solid gray;width:81pt'>합 의</td>
		<td rowspan=3 class=xl12425868 dir=LTR width=142 style='border-bottom:.5pt solid gray;width:107pt'>비고</td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td rowspan=2 height=32 class=xl13325868 dir=LTR width=108 style='border-bottom:.5pt solid gray;height:24.0pt;border-top:none;width:81pt'>업무</td>
		<td rowspan=2 class=xl12625868 dir=LTR width=267 style='border-bottom:.5pt solid gray;border-top:none;width:200pt'>전결사항</td>
		<td rowspan=2 class=xl8125868 dir=LTR width=51 style='border-bottom:.5pt solid gray;border-top:none;width:38pt'>파트장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td rowspan=2 class=xl8225868 dir=LTR width=43 style='border-bottom:.5pt solid gray;border-top:none;width:32pt'>팀 장</td>
		<td rowspan=2 class=xl8225868 dir=LTR width=51 style='border-bottom:.5pt solid gray;border-top:none;width:38pt'>부문장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl6525868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>대표<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl6625868 dir=LTR width=38 style='height:12.0pt; border-left:none;width:29pt'>이사</td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl6725868 dir=LTR width=108 style='height:12.75pt;border-top:none;width:81pt'><span style='mso-spacerun:yes'>&nbsp;</span>1. 전략 기획</td>
		<td class=xl6825868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(1) 사업계획/전략 및 추진과제 수립<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl6925868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7025868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7025868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7125868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>○</td>
		<td class=xl7225868 dir=LTR width=108 style='border-top:none;border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl13725868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'>이사회 승인<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7325868 dir=LTR width=108 style='height:12.75pt; width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(2) 관리회계 기준 수립<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7725868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>○</td>
		<td class=xl7825868 dir=LTR width=108 style='border-top:none;border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl13825868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'><u style='visibility:hidden;mso-ignore:visibility'></u></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(3) 회의체 운영<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 전사회의체<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○</td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl8925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 부문내 회의체<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9125868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9225868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○</td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl9625868 dir=LTR width=108 style='height:12.75pt;border-top:none;width:81pt'><span style='mso-spacerun:yes'>&nbsp;</span>2. 인사 교육<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(1) 인사전략/제도의 수립<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7725868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7325868 dir=LTR width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(2) 조직의 신설,폐지,명칭변경<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7725868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(3) 인력운용<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 정규직 채용<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 계약/파견/프리랜서 채용<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'>기간연장 포함<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>- 월급계약<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○</td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>- 시급계약<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'>인사파트 통보<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 신분 전환<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(4) 인사관리<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl10725868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 조직 책임자 임면<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7325868 dir=LTR width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 인사이동<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>가. 부문간<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○</td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl13925868 dir=LTR width=142 style='border-left:none;width:107pt'><u style='visibility:hidden;mso-ignore:visibility'>　</u></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>나. 부문내<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○</td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 진급<span style='mso-spacerun:yes'>&nbsp;&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○</td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>④ 인사평가</td>
		<td class=xl10325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl10425868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>가. 부문장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>나. 팀<span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span>장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>다. 팀<span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span>원<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>⑤ 휴가/출장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>가. 부문장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'>CFO통보<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>나. 팀<span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span>장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>다. 팀<span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span>원<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(5) 복리후생<span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl10725868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 기본방침, 제도수립/변경<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7325868 dir=LTR width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 복리후생비 신청<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 dir=LTR width=108 style='border-left:none;width:81pt'>인사파트<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(6) 포상 및 징계<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7725868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'>인사위원회<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(7) 급/상여 기준 결정<span style='mso-spacerun:yes'>&nbsp;</span>및 성과급 지급<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7725868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(8) 목표 합의/ 평가<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 부문장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 팀<span style='mso-spacerun:yes'>&nbsp;&nbsp; </span>장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 팀<span style='mso-spacerun:yes'>&nbsp;&nbsp; </span>원<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(9) 4대보험 관리 및 납부<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○</td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7725868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(10) 교 육<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 교육제도의 수립<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 교육실시<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>인사파트<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 교육 참가<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>가. 부문장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl10725868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>나. 팀<span style='mso-spacerun:yes'>&nbsp;&nbsp; </span>장</td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl10825868 dir=LTR width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</span>다. 팀<span style='mso-spacerun:yes'>&nbsp;&nbsp; </span>원<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl9625868 dir=LTR width=108 style='height:12.0pt;border-top:none;width:81pt'>3. 기업 문화<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(1) 행사<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7325868 dir=LTR width=108 style='height:12.0pt;width:81pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;</span></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 전<span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>사<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 부문 단위<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 dir=LTR width=38 style='border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9525868 dir=LTR width=108 style='border-left:none;width:81pt'>기업문화팀<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(2) 신문/도서/잡지등의 구입/구독<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 dir=LTR width=43 style='border-top:none;border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7625868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7725868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl7825868 dir=LTR width=108 style='border-top:none;border-left:none;width:81pt'>기업문화팀<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(3) 휴양시설 관리<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 신규계약/ 해지<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl8925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 관리/운영 기준 수립<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl9625868 dir=LTR width=108 style='height:12.75pt;border-top:none;width:81pt'>4. 재무<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(1) 회계기준 수립/변경<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7725868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(2) 전표승인<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 dir=LTR width=43 style='border-top:none;border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7625868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7725868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(3) 결<span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>산<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 결산 정책 수립<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 결산 보고<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'>분기단위 이사회 승인<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl10725868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 회계감사<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-left:none;width:200pt'>(4) 고정자산 재물조사<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl7625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7725868 dir=LTR width=38 style='border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(5) 재고조사<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7725868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7825868 dir=LTR width=108 style='border-top:none;border-left:none;width:81pt'>영업지원부문장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(6) 내부회계관리 제도<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 내부회계관리규정 제개정<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'>이사회<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl8925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 운용실태 평가/보고<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 dir=LTR width=38 style='border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'>이사회<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl9625868 dir=LTR width=108 style='height:12.0pt;border-top:none;width:81pt'>5. 세무<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(1) 신고 및 납부<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 법인세(중간예납포함)<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 부가세,원천세,지방세등<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 수정신고,경정,보정<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(2) 사업자 등록 관리(신규,정정)<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7725868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(3) 세무조사<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7725868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(4) 조세 불복<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7725868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(5) 가산세,벌과금 등 납부<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl6525868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl11025868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl11125868 dir=LTR width=267 style='border-left:none;width:200pt'>(6) 지점 및 사무소 설치 / 이전<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl11225868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl11325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl11325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl11425868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl11525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl14025868 dir=LTR width=142 style='border-left:none;width:107pt'>이사회<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7325868 dir=LTR width=108 style='height:12.75pt;width:81pt'>6. 자금<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'>(1) 자금수지 계획/실적<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(2) 자금 조달<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7725868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;
		</span>① 증자<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'>이사회<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 회사채<span style='mso-spacerun:yes'>&nbsp; </span>발행
			<span style='mso-spacerun:yes'>&nbsp;</span>
		</td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'>이사회<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 1년이상<span style='mso-spacerun:yes'>&nbsp; </span>장기차입<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'>5억이상<span style='mso-spacerun:yes'>&nbsp; </span>이사회 부의<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>④ 1년미만 단기차입/한도 약정<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 dir=LTR width=38 style='border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'>5억이상<span style='mso-spacerun:yes'>&nbsp; </span>이사회 부의<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(3) 자금 집행<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 자금 집금<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 100만원 미만 소액/고객환불<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 고액 이체 및 예금 청구서<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>④ 대금 지급기준 변경<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(4) 자금 대여,지급보증,담보제공<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7725868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'>이사회<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(5) 채권관리<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 채권 현황 보고(매월)<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 회수불능채권 대손 처리<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>-. 건당 5백만원 초과<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			</span>-. 건당 5백만원 이하
		</td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 소멸시효 경과 채권 제각<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(6) 자금 운용<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>①주식관련 상품<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>②주식관련 상품 제외<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(7) 견질어음 등 담보물 관리<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7725868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(8) 퇴직연금 가입 및 납부<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7725868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(9) 법인카드(체크카드 포함)<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 신규거래 계약<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 발급/ 재발급<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 사용 예산 부여<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-left:none;width:200pt'>(10) 결제 수단<span style='mso-spacerun:yes'>&nbsp;&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 신용카드<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>가. 가맹점 계약<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>나. 수수료 변경<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			</span>다. 업무 제휴
		</td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>영업부문장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</span>라. 청구 및 입금 관리<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 결제수단 도입/변경/폐지<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(11) 인장관리<span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 법인인감 및 통장 인감<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 사용인감<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl8925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 법인인감 증명서 관리<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl9625868 dir=LTR width=108 style='height:12.75pt;border-top:none;width:81pt'>7. 대외 업무</td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(1) 신고 및 인 허가<span style='mso-spacerun:yes'>&nbsp;&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7325868 dir=LTR width=108 style='height:12.75pt;width:81pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td class=xl8025868 dir=LTR width=267 style='border-left:none;width:200pt'>(2) 신고 사건 대응<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8225868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl6525868 dir=LTR width=38 style='border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8325868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl11625868 dir=LTR width=108 style='height:12.0pt;width:81pt'>8. 법무<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl11725868 dir=LTR width=267 style='border-left:none;width:200pt'>(1) 계약서 검토<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl11825868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl11925868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl11925868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl12025868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl12125868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl14125868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 내부 계약 검토 의뢰<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
		<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 외부 법률자문 의뢰<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(2) 약관, 표준계약서의 제정 및 개폐<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7725868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl7825868 dir=LTR width=108 style='border-top:none;border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(3) 회사 규정의 제정 및 개폐<span style='mso-spacerun:yes'>&nbsp;&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7725868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7825868 dir=LTR width=108 style='border-top:none;border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(4) 분쟁 처리<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7725868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7825868 dir=LTR width=108 style='border-top:none;border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl8925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(5) 가압류 접수/추심금 지급<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 dir=LTR width=43 style='border-top:none;border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7625868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7725868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl7825868 dir=LTR width=108 style='border-top:none;border-left:none;width:81pt'>회계팀장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl9625868 dir=LTR width=108 style='height:12.0pt;border-top:none;width:81pt'>9.IT<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(1) 신규 개발 프로젝트<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 3개월 이상 소요<span style='mso-spacerun:yes'>&nbsp;&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>or 외주비용 1천만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 3개월 미만 소요<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</span>and 외주비용 1천만원 미만<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(2) 보안<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 보안 관련 기준 수립<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 시스템 권한 부여 및 해제
			<span style='mso-spacerun:yes'>&nbsp;</span>
		</td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 보안 점검 및 교육<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl8925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>④ S/W 관리<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl9625868 dir=LTR width=108 style='height:12.75pt;border-top:none;width:81pt'>10.CS<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(1) VOC보고<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7725868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl8925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(2) CS 처리기준 결정 및 변경<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7725868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl9625868 dir=LTR width=108 style='height:12.0pt;border-top:none;width:81pt'>11.물류<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(1) 배송업체 선정 및 변경<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl8925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'>(2) 물류관련 전략 수립<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 dir=LTR width=38 style='border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9525868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl9625868 dir=LTR width=108 style='height:12.75pt;border-top:none;width:81pt'>12.투자<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(1) 부동산 취득/처분<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7725868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7825868 dir=LTR width=108 style='border-top:none;border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'>5억이상<span style='mso-spacerun:yes'>&nbsp; </span>모든 투자<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(2) 부동산의 임대차<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'>이사회 부의<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 신규 계약<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 계약조건 변경 및 해지<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 dir=LTR width=38 style='border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9525868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(3) 타법인 출자 및 유가증권 취득<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7725868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7825868 dir=LTR width=108 style='border-top:none;border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl11625868 dir=LTR width=108 style='height:12.0pt;width:81pt'>13. 비용 집행</td>
		<td class=xl9825868 dir=LTR width=267 style='border-left:none;width:200pt'>(1) 자산의 취득 및 처분<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9925868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl10025868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl10025868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl10125868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl10225868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12425868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7325868 dir=LTR width=108 style='height:12.0pt;width:81pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 취득<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>가. 1천만원 이상<span style='mso-spacerun:yes'>&nbsp;</span>
		</td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>나. 2백만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>다. 2백만원 미만<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>재무팀장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 처분 및 폐기<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>가. 취득가 기준 1억이상<span style='mso-spacerun:yes'>&nbsp; </span>or<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>장부가 기준 1천만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>나. 취득가 기준 2천만원이상 or<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>장부가 기준 2백만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>다. 취득가 기준<span style='mso-spacerun:yes'>&nbsp; </span>2천만원 미만 &amp;
			<span style='mso-spacerun:yes'>&nbsp;</span>
		</td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>재무팀장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>장부가 기준 2백만원 미만<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-left:none;width:200pt'>(2) 자산의 임대차<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 임대차료 년 환산<span style='mso-spacerun:yes'>&nbsp; </span>천만원이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 임대차료 년 환산<span style='mso-spacerun:yes'>&nbsp; </span>천만원미만<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(3) 설비등 공사 (인테리어<span style='mso-spacerun:yes'>&nbsp;</span>등)<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 1천만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 1천만원 미만<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(4) 비경상 외부용역 계약 체결<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 1천만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'>5억이상 이사회 승인<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 1천만원 미만<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(5) 경상적인 외부용역 계약 체결<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 5천만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 1천만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 1천만원 미만<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>재무팀장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-left:none;width:200pt'>(6) 부서 운영 예산<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 예산 배정 기준 수립<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 예산 한도내 운영<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 예산 한도 증액 요청<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(7) 일반관리비등 고정지출 비용 집행<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 계약 및 지급기준에 의한 경비
		<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>재무팀장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 통상수준을<span style='mso-spacerun:yes'>&nbsp; </span>벗어나는지출<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl9325868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(8) 광고비<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 건당 2천만원<span style='mso-spacerun:yes'>&nbsp; </span>이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 건당 5백만원<span style='mso-spacerun:yes'>&nbsp; </span>이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 건당 5백만원<span style='mso-spacerun:yes'>&nbsp; </span>미만
			<span style='mso-spacerun:yes'>&nbsp;</span>
		</td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 dir=LTR width=108 style='border-left:none;width:81pt'>재무팀장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(9) 기부금<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 5백만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 5백만원 미만<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-left:none;width:200pt'>(10) 판촉비<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 건당 1천만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 건당 5백만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 건당 5백만원 미만<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(11) 재고처리<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 손망실 처리<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>가. 1천만원<span style='mso-spacerun:yes'>&nbsp; </span>이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>나. 5백만원<span style='mso-spacerun:yes'>&nbsp; </span>이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>다. 5백만원<span style='mso-spacerun:yes'>&nbsp; </span>미만<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>재무팀장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 평가감 대상 상품 선정 및 시행<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>가. 1천만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>나. 5백만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>다. 5백만원 미만<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>재무팀장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>③ 이상재고 추가발주 제한<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>MD팀장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'>영업지원부문장<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>④ 일괄매각,폐기<span style='mso-spacerun:yes'>&nbsp; </span>및 처분<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>가. 1천만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>나. 5백만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl8925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>다. 5백만원 미만<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 dir=LTR width=108 style='border-left:none;width:81pt'>재무팀장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl9625868 dir=LTR width=108 style='height:12.75pt;border-top:none;width:81pt'><span style='mso-spacerun:yes'>&nbsp;</span>14. 상품<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(1) 상품 진행조건<span style='mso-spacerun:yes'>&nbsp; </span>및 신규입점 기준<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7725868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl7425868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(2) 기준에 벗어난 상품진행 및 입점<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7525868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl7625868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl7625868 dir=LTR width=51 style='border-top:none;border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl7725868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl7825868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl10625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(3) 상품의 매입<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>가. 5천만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'>5억이상 이사회 승인<span style='mso-spacerun:yes'>&nbsp;</span></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>나. 1천만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>다. 1천만원 미만<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(4) 대금지급<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-top:none;border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-top:none;border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>가. 정기 지급<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>나. 대금 선지급<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl7925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>다. 조기 지급/지급보류<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>△<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-left:none;width:200pt'>(5) 수입<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl12225868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl12325868 dir=LTR width=267 style='width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>① 통관 진행<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl14225868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>② 수입대금 결제<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>가. 1억 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>나. 1천만원 이상<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl8925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>다. 1천만원 미만<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 dir=LTR width=108 style='border-left:none;width:81pt'>재무팀장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl10525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl9625868 dir=LTR width=108 style='height:12.75pt;border-top:none;width:81pt'>15. 기타<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8025868 dir=LTR width=267 style='border-top:none;border-left:none;width:200pt'>(1) 윤리 위원회 운영<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-top:none;border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-top:none;border-left:none;width:38pt'></td>
		<td class=xl6525868 dir=LTR width=38 style='border-top:none;border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8325868 dir=LTR width=108 style='border-top:none;border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12625868 dir=LTR width=142 style='border-top:none;border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8025868 dir=LTR width=267 style='border-left:none;width:200pt'>(2) 업무 인수 인계<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8125868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8225868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8225868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl6525868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8325868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12625868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>가. 부문장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8725868 dir=LTR width=38 style='border-left:none;width:29pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8825868 dir=LTR width=108 style='border-left:none;width:81pt'>CFO<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=16 style='height:12.0pt'>
		<td height=16 class=xl7925868 width=108 style='height:12.0pt;width:81pt'></td>
		<td class=xl8425868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>나. 팀<span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span>장<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8525868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl8625868 width=43 style='border-left:none;width:32pt'></td>
		<td class=xl8625868 dir=LTR width=51 style='border-left:none;width:38pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl8725868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl8825868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl12525868 dir=LTR width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<tr height=17 style='height:12.75pt'>
		<td height=17 class=xl10925868 width=108 style='height:12.75pt;width:81pt'></td>
		<td class=xl9025868 dir=LTR width=267 style='border-left:none;width:200pt'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span>다. 팀<span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span>원
		<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9725868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9325868 dir=LTR width=43 style='border-left:none;width:32pt'>○<span style='mso-spacerun:yes'>&nbsp;</span></td>
		<td class=xl9325868 width=51 style='border-left:none;width:38pt'></td>
		<td class=xl9425868 width=38 style='border-left:none;width:29pt'></td>
		<td class=xl9525868 width=108 style='border-left:none;width:81pt'></td>
		<td class=xl10525868 width=142 style='border-left:none;width:107pt'></td>
	</tr>
	<![if supportMisalignedColumns]>
	<tr height=0 style='display:none'>
		<td width=108 style='width:81pt'></td>
		<td width=267 style='width:200pt'></td>
		<td width=51 style='width:38pt'></td>
		<td width=43 style='width:32pt'></td>
		<td width=51 style='width:38pt'></td>
		<td width=38 style='width:29pt'></td>
		<td width=108 style='width:81pt'></td>
		<td width=142 style='width:107pt'></td>
	</tr>
<![endif]>-->
</table>
</div>
</body>
</html>