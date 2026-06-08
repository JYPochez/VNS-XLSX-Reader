#tag Module
Protected Module SpreadsheetReader
	#tag Method, Flags = &h0, Description = 4F70656E2061202E786C7378206F72202E6F64732066726F6D206120466F6C6465724974656D2C206469737061746368696E6720627920657874656E73696F6E2E2052657475726E73207468652073686172656420584C5358576F726B626F6F6B206D6F64656C2E0A
		Function Open(file As FolderItem, mode As XLSXEnums.eOpenMode = XLSXEnums.eOpenMode.Auto) As XLSXWorkbook
		  Var hint As String = If(file <> Nil, file.Name, "")
		  Return Open(file, hint, mode)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Open(file As FolderItem, nameHint As String, mode As XLSXEnums.eOpenMode = XLSXEnums.eOpenMode.Auto) As XLSXWorkbook
		  ' Dispatch by file extension (nameHint lets the Web app pass the original
		  ' uploaded filename when the on-disk FolderItem has a temp name).
		  If nameHint.Lowercase.EndsWith(".ods") Then
		    Return ODSReader.Open(file, mode)
		  End If

		  ' Default to XLSX. If the archive turns out to be ODS (wrong/missing
		  ' extension), retry as ODS once.
		  Try
		    Return XLSXReader.Open(file, mode)
		  Catch ex As XLSXException
		    If ex.Code = XLSXEnums.eParseError.MissingPart Then
		      Return ODSReader.Open(file, mode)
		    End If
		    Raise ex
		  End Try
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4F70656E2061202E786C7378206F72202E6F64732066726F6D207261772062797465733B2064697370617463682062792074686520736F757263654E616D6520657874656E73696F6E2C207769746820616E20584C53582D3E4F44532066616C6C6261636B206F6E204D697373696E67506172742E0A
		Function Open(data As MemoryBlock, sourceName As String, mode As XLSXEnums.eOpenMode = XLSXEnums.eOpenMode.Auto) As XLSXWorkbook
		  If sourceName.Lowercase.EndsWith(".ods") Then
		    Return ODSReader.Open(data, sourceName, mode)
		  End If
		  Try
		    Return XLSXReader.Open(data, sourceName, mode)
		  Catch ex As XLSXException
		    If ex.Code = XLSXEnums.eParseError.MissingPart Then
		      Return ODSReader.Open(data, sourceName, mode)
		    End If
		    Raise ex
		  End Try
		End Function
	#tag EndMethod

	#tag Note, Name = About
		Format-dispatching front door for both readers. The UI calls
		SpreadsheetReader.Open and doesn't care whether the file is .xlsx or .ods.

		Dispatch is by file extension (the nameHint overload lets the Web uploader
		pass the original filename). If the extension is missing or wrong, an XLSX
		parse that fails with MissingPart (no xl/workbook.xml) is retried as ODS.

		All overloads return the shared XLSXWorkbook model and thread the
		XLSXEnums.eOpenMode through to XLSXZip.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
