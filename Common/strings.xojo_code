#tag Module
Protected Module strings
	#tag Constant, Name = kStrAppTitle, Type = String, Dynamic = True, Default = \"VNS XLSX Reader", Scope = Public, Description = 4170702077696E646F77202F2070616765207469746C652E0A
	#tag EndConstant

	#tag Constant, Name = kStrMenuFileOpen, Type = String, Dynamic = True, Default = \"Open\xE2\x80\xA6", Scope = Public, Description = 46696C65206D656E75202D204F70656E2E2E2E206974656D206C6162656C2E0A
	#tag EndConstant

	#tag Constant, Name = kStrErrorTitle, Type = String, Dynamic = True, Default = \"Cannot open file", Scope = Public, Description = 5469746C65206F6620746865206572726F72206469616C6F672073686F776E207768656E206F70656E696E6720612066696C65206661696C732E0A
	#tag EndConstant

	#tag Constant, Name = kStrErrorNotXLSX, Type = String, Dynamic = True, Default = \"This file is not a valid .xlsx workbook.", Scope = Public, Description = 4572726F72206D6573736167653A2066696C65206973206E6F7420612076616C696420584C535820776F726B626F6F6B2E0A
	#tag EndConstant

	#tag Constant, Name = kStrErrorMalformed, Type = String, Dynamic = True, Default = \"This workbook has malformed content.", Scope = Public, Description = 4572726F72206D6573736167653A20776F726B626F6F6B20636F6E74656E74206973206D616C666F726D65642E0A
	#tag EndConstant

	#tag Constant, Name = kStrErrorEncrypted, Type = String, Dynamic = True, Default = \"Encrypted workbooks are not supported.", Scope = Public, Description = 4572726F72206D6573736167653A20656E6372797074656420776F726B626F6F6B7320617265206E6F7420737570706F727465642E0A
	#tag EndConstant

	#tag Constant, Name = kStrErrorGeneric, Type = String, Dynamic = True, Default = \"Could not read this workbook.", Scope = Public, Description = 4572726F72206D6573736167653A2067656E65726963206661696C757265206E6F7420636F766572656420627920746865206D6F72652073706563696669632063617365732E0A
	#tag EndConstant

	#tag Constant, Name = kStrUploadPrompt, Type = String, Dynamic = True, Default = \"Choose an .xlsx file to view\xE2\x80\xA6", Scope = Public, Description = 5765622066696C652075706C6F616465722070726F6D70742061736B696E6720746865207573657220746F2063686F6F736520616E202E786C73782066696C652E0A
	#tag EndConstant

	#tag Constant, Name = kStrInMemory, Type = String, Dynamic = True, Default = \"Read in memory", Scope = Public, Description = 43617074696F6E20666F722074686520225265616420696E206D656D6F72792220636865636B626F7820636F6E74726F6C6C696E6720584C53585A69702773206F70656E206261636B656E642E0A
	#tag EndConstant

	#tag Constant, Name = kStrParseTime, Type = String, Dynamic = True, Default = \"Parsed in ", Scope = Public, Description = 5072656669782073686F776E206E65787420746F2074686520696E2D6D656D6F727920636865636B626F78206265666F7265207468652070617273652D656C6170736564206D696C6C697365636F6E64732E0A
	#tag EndConstant

	#tag Constant, Name = kStrParseTimeUnit, Type = String, Dynamic = True, Default = \" ms", Scope = Public, Description = 5375666669782073686F776E206166746572207468652070617273652D656C61707365642076616C7565202864656661756C742022206D7322292E0A
	#tag EndConstant

	#tag Constant, Name = kStrSaveButton, Type = String, Dynamic = True, Default = \"Save\xE2\x80\xA6", Scope = Public, Description = 43617074696F6E206F662074686520536176652E2E2E20627574746F6E20696E20746865206D61696E2077696E646F772E0A
	#tag EndConstant

	#tag Constant, Name = kStrSaveDialogTitle, Type = String, Dynamic = True, Default = \"Save as .xlsx or .ods", Scope = Public, Description = 5469746C65206F66207468652073617665206469616C6F673B207468652063686F73656E20657874656E73696F6E20282E786C7378202F202E6F647329207069636B732074686520666F726D61742E0A
	#tag EndConstant

	#tag Constant, Name = kStrSavedPrefix, Type = String, Dynamic = True, Default = \"Saved ", Scope = Public, Description = 5072656669782073686F776E20696E2074686520737461747573206C6162656C2061667465722061207375636365737366756C20736176652C206265666F7265207468652066696C656E616D652E0A
	#tag EndConstant

	#tag Constant, Name = kStrErrorSaveTitle, Type = String, Dynamic = True, Default = \"Cannot save file", Scope = Public, Description = 5469746C65206F6620746865206572726F72206469616C6F672073686F776E207768656E20736176696E6720612066696C65206661696C732E0A
	#tag EndConstant

	#tag Constant, Name = kStrErrorSaveGeneric, Type = String, Dynamic = True, Default = \"Could not save this workbook.", Scope = Public, Description = 4572726F72206D6573736167653A2067656E657269632073617665206661696C7572652E0A
	#tag EndConstant

	#tag Constant, Name = kStrFileTypeXlsx, Type = String, Dynamic = True, Default = \"Excel Workbook", Scope = Public, Description = 446973706C6179206E616D65206F662074686520457863656C202E786C73782066696C65207479706520696E207468652073617665206469616C6F672066696C7465722E0A
	#tag EndConstant

	#tag Constant, Name = kStrFileTypeOds, Type = String, Dynamic = True, Default = \"OpenDocument Spreadsheet", Scope = Public, Description = 446973706C6179206E616D65206F6620746865204F70656E446F63756D656E74202E6F64732066696C65207479706520696E207468652073617665206469616C6F672066696C7465722E0A
	#tag EndConstant

	#tag Constant, Name = kStrNewButton, Type = String, Dynamic = True, Default = \"New", Scope = Public, Description = 43617074696F6E206F6620746865204E657720627574746F6E206372656174696E6720616E20656D70747920776F726B626F6F6B2E0A
	#tag EndConstant

	#tag Constant, Name = kStrAddRow, Type = String, Dynamic = True, Default = \"+ Row", Scope = Public, Description = 43617074696F6E206F6620746865206164642D6F6E652D726F7720627574746F6E2E0A
	#tag EndConstant

	#tag Constant, Name = kStrDelRow, Type = String, Dynamic = True, Default = \"- Row", Scope = Public, Description = 43617074696F6E206F66207468652072656D6F76652D6C6173742D726F7720627574746F6E2E0A
	#tag EndConstant

	#tag Constant, Name = kStrAddCol, Type = String, Dynamic = True, Default = \"+ Col", Scope = Public, Description = 43617074696F6E206F6620746865206164642D6F6E652D636F6C756D6E20627574746F6E2E0A
	#tag EndConstant

	#tag Constant, Name = kStrDelCol, Type = String, Dynamic = True, Default = \"- Col", Scope = Public, Description = 43617074696F6E206F66207468652072656D6F76652D6C6173742D636F6C756D6E20627574746F6E2E0A
	#tag EndConstant

	#tag Constant, Name = kStrUntitledName, Type = String, Dynamic = True, Default = \"untitled", Scope = Public, Description = 536F75726365206E616D6520676976656E20746F206120776F726B626F6F6B20637265617465642066726F6D20736372617463682028647269766573207468652073756767657374656420736176652066696C656E616D65292E0A
	#tag EndConstant

	#tag Constant, Name = kStrDefaultSheetName, Type = String, Dynamic = True, Default = \"Sheet1", Scope = Public, Description = 4E616D65206F66207468652073696E676C6520736865657420696E206120776F726B626F6F6B20637265617465642066726F6D20736372617463682E0A
	#tag EndConstant

	#tag Note, Name = About
		Localizable user-visible strings.
		
		All constants are Dynamic so the IDE Language Editor can override them per
		locale at build time. Code must never use hardcoded user-visible strings -
		always reference a kStr... constant from this module.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
