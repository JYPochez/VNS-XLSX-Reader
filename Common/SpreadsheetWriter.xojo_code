#tag Module
Protected Module SpreadsheetWriter
	#tag Method, Flags = &h0, Description = 466F726D61742D61676E6F7374696320736176653A207468652064657374696E6174696F6E20657874656E73696F6E207069636B73207468652073657269616C697A657220282E6F6473202D3E204F44535772697465722C20656C736520584C5358577269746572292E0A
		Sub Save(wb As XLSXWorkbook, file As FolderItem)
		  ' Format-agnostic save: the destination extension picks the serializer.
		  ' .ods -> ODSWriter, anything else (including .xlsx) -> XLSXWriter.
		  If wb Is Nil Then
		    Raise New XLSXException(XLSXEnums.eParseError.Unsupported, "no workbook")
		  End If
		  If file Is Nil Then
		    Raise New XLSXException(XLSXEnums.eParseError.Unsupported, "no destination")
		  End If
		  If file.Name.Lowercase.EndsWith(".ods") Then
		    ODSWriter.Save(wb, file)
		  Else
		    XLSXWriter.Save(wb, file)
		  End If
		End Sub
	#tag EndMethod

	#tag Note, Name = About
		Format-dispatching front door for saving — the write-side mirror of
		SpreadsheetReader. The UI calls SpreadsheetWriter.Save(wb, file) and the
		destination file's extension picks the serializer: .ods -> ODSWriter,
		anything else -> XLSXWriter (so .xlsx is also the fallback for exotic or
		missing extensions; the UI is expected to normalize the extension first).

		Both serializers consume the shared XLSXWorkbook model, so a workbook
		loaded from one format can be saved into the other (format conversion).
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
