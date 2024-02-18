/**
 * @license Copyright (c) 2003-2017, CKSource - Frederico Knabben. All rights reserved.
 * For licensing, see LICENSE.md or http://ckeditor.com/license
 */

CKEDITOR.editorConfig = function( config ) {
	config.toolbarGroups = [
		{ name: 'document', groups: [ 'mode', 'document', 'doctools' ] },
		{ name: 'clipboard', groups: [ 'clipboard', 'undo' ] },
		{ name: 'editing', groups: [ 'find', 'selection', 'spellchecker', 'editing' ] },
		{ name: 'forms', groups: [ 'forms' ] },
		{ name: 'basicstyles', groups: [ 'basicstyles', 'cleanup' ] },
		{ name: 'paragraph', groups: [ 'list', 'indent', 'blocks', 'align', 'bidi', 'paragraph' ] },
		'/',
		{ name: 'styles', groups: [ 'styles' ] },
		{ name: 'links', groups: [ 'links' ] },
		{ name: 'insert', groups: [ 'insert' ] },
		{ name: 'colors', groups: [ 'colors' ] },
		{ name: 'tools', groups: [ 'tools' ] },
		{ name: 'others', groups: [ 'others' ] },
		{ name: 'about', groups: [ 'about' ] }
	];

	config.removeButtons = 'Save,NewPage,Replace,Scayt,Textarea,TextField,Radio,Form,Select,Button,ImageButton,HiddenField,Checkbox,CreateDiv,BidiRtl,BidiLtr,Language,Anchor,Flash,PageBreak,About,Blockquote';

	// 툴바 글꼴 설정
	config.font_names = 'malgun gothic, \'맑은 고딕\', sans-serif; dotum, \'돋움\', sans-serif; gulim, \'굴림\', sans-serif; Arial, sans-serif; Verdana, sans-serif; Tahoma, sans-serif; gunsuh, \'궁서\', serif; times new roman, serif;';
	
	// 컨텐츠 필터링(자동 스타일 필터링 추가 허용)
	config.extraAllowedContent = 'video[*]{*};source[*]{*};div{border, border-top, border-bottom, padding, background-color, margin, margin-top, margin-bottom, margin-left, margin-right, width} (alignCt); p{width, padding, margin, margin-top, margin-bottom, margin-left, margin-right, font-weight, font-size, font-family, color, line-height} (alignCt); span{display, padding, font-weight, line-height}; strong{line-height, font-size, font-family, color}; td{padding, font-family, font-size, color, line-height}';

	//폰트 사이즈 셀렉트 구간 설정
	config.fontSize_sizes = '14/14px;16/16px;18/18px;20/20px;22/22px;24/24px;26/26px;28/28px;36/36px;48/48px;';
};