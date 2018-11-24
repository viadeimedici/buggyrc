/*
Copyright (c) 2003-2010, CKSource - Frederico Knabben. All rights reserved.
For licensing, see LICENSE.html or http://ckeditor.com/license
*/

CKEDITOR.editorConfig = function( config )
{
	// Define changes to default configuration here. For example:
	// config.language = 'fr';
	// config.uiColor = '#AADC6E';
	
//	config.toolbar = 'MyToolbar';
//
//    config.toolbar_MyToolbar =
//    [
//		
//		
//		['Source','ShowBlocks','-','NewPage','Preview','Print','-','Templates'],
//		['Cut','Copy','Paste','PasteText','PasteFromWord'],
//		['Undo','Redo','-','Find','Replace','-','SelectAll','RemoveFormat'],
//		'/',
//		['Bold','Italic','Underline','Strike','-','Subscript','Superscript'],
//		['NumberedList','BulletedList'],
//		['JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock'],
//		['Link','Unlink','Anchor'],
//		['Image','Table','HorizontalRule','Smiley','SpecialChar','PageBreak'],
//		'/',
//		['Styles','Format','Font','FontSize'],
//		['TextColor','BGColor'],
//		['Outdent','Indent','Blockquote'],
//		['About'],
//	
//	
//    ];
	
	config.toolbar = 'MyToolbar';

    config.toolbar_MyToolbar =
    [
		
		
		['Source'],
		['Cut','Copy','Paste','PasteText','PasteFromWord'],
		['Undo','Redo','-','SelectAll','RemoveFormat'],
		'/',
		['Bold','Italic','Underline','Strike','-','Subscript','Superscript'],
		['NumberedList','BulletedList'],
		['Link','Unlink','Anchor'],
		
		
//		['Source','ShowBlocks','-','NewPage','Preview','Print','-','Templates'],
//		['Cut','Copy','Paste','PasteText','PasteFromWord'],
//		['Undo','Redo','-','Find','Replace','-','SelectAll','RemoveFormat'],
//		'/',
//		['Bold','Italic','Underline','Strike','-','Subscript','Superscript'],
//		['NumberedList','BulletedList'],
//		['JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock'],
//		['Link','Unlink','Anchor'],
//		['Table','HorizontalRule','Smiley','SpecialChar','PageBreak'],
//		'/',
//		['Styles','Format','Font','FontSize'],
//		['TextColor','BGColor'],
//		['Outdent','Indent','Blockquote'],
//		['About'],
	
	
    ];
};
