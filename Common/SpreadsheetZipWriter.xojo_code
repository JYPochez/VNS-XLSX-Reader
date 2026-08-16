#tag Class
Protected Class SpreadsheetZipWriter
	#tag Method, Flags = &h0, Description = 5374616D7020616C6C20656E74726965732077697468206F6E6520444F532074696D657374616D702074616B656E20617420777269746572206372656174696F6E2E0A
		Sub Constructor()
		  ' Stamp all entries with one DOS timestamp taken at writer creation.
		  Var now As DateTime = DateTime.Now
		  Var yr As Integer = Max(now.Year, 1980)
		  mDosTime = (now.Hour * 2048) + (now.Minute * 32) + (now.Second \ 2)
		  mDosDate = ((yr - 1980) * 512) + (now.Month * 32) + now.Day
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5175657565206F6E65206172636869766520656E74727920287772697474656E20696E2041646450617274206F72646572292E20666F72636553746F72656420777269746573206974207769746820636F6D7072657373696F6E206D6574686F64203020E2809420726571756972656420666F7220746865204F4453206D696D657479706520656E7472792E0A
		Sub AddPart(name As String, content As String, forceStored As Boolean = False)
		  ' Queue one archive entry. Entries are written in AddPart order, which is
		  ' how ODS gets its mimetype first. forceStored writes the entry with
		  ' compression method 0 (required for the ODS mimetype entry).
		  Var mb As MemoryBlock = content
		  mNames.Add name
		  mDatas.Add mb
		  mForceStored.Add forceStored
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 417373656D626C652074686520617263686976652062797465733A206C6F63616C20656E74726965732C2063656E7472616C206469726563746F72792C20656E642D6F662D63656E7472616C2D6469726563746F72792E204E6F207A697036342E0A
		Function ToMemoryBlock() As MemoryBlock
		  ' Assemble the zip: local entries, then central directory, then EOCD.
		  Var locals() As MemoryBlock
		  Var centrals() As MemoryBlock
		  Var offsets() As Integer
		  Var offset As Integer = 0

		  For i As Integer = 0 To mNames.LastIndex
		    Var nm As String = mNames(i)
		    Var nameLen As Integer = nm.Bytes
		    Var data As MemoryBlock = mDatas(i)
		    Var dataSize As Integer = data.Size
		    Var crc As UInt32 = Crc32Of(data)

		    Var method As Integer = 0
		    Var payload As MemoryBlock = data
		    If Not mForceStored(i) And dataSize > 0 Then
		      Var defl As MemoryBlock = RawDeflate(data)
		      If defl <> Nil And defl.Size < dataSize Then
		        method = 8
		        payload = defl
		      End If
		    End If

		    Var lh As New MemoryBlock(30 + nameLen + payload.Size)
		    lh.LittleEndian = True
		    lh.UInt32Value(0) = &h04034B50
		    lh.UInt16Value(4) = 20
		    lh.UInt16Value(6) = 0
		    lh.UInt16Value(8) = method
		    lh.UInt16Value(10) = mDosTime
		    lh.UInt16Value(12) = mDosDate
		    lh.UInt32Value(14) = crc
		    lh.UInt32Value(18) = payload.Size
		    lh.UInt32Value(22) = dataSize
		    lh.UInt16Value(26) = nameLen
		    lh.UInt16Value(28) = 0
		    lh.StringValue(30, nameLen) = nm
		    If payload.Size > 0 Then
		      lh.StringValue(30 + nameLen, payload.Size) = payload.StringValue(0, payload.Size)
		    End If
		    locals.Add lh
		    offsets.Add offset
		    offset = offset + lh.Size

		    Var ch As New MemoryBlock(46 + nameLen)
		    ch.LittleEndian = True
		    ch.UInt32Value(0) = &h02014B50
		    ch.UInt16Value(4) = 20
		    ch.UInt16Value(6) = 20
		    ch.UInt16Value(8) = 0
		    ch.UInt16Value(10) = method
		    ch.UInt16Value(12) = mDosTime
		    ch.UInt16Value(14) = mDosDate
		    ch.UInt32Value(16) = crc
		    ch.UInt32Value(20) = payload.Size
		    ch.UInt32Value(24) = dataSize
		    ch.UInt16Value(28) = nameLen
		    ch.UInt16Value(30) = 0
		    ch.UInt16Value(32) = 0
		    ch.UInt16Value(34) = 0
		    ch.UInt16Value(36) = 0
		    ch.UInt32Value(38) = 0
		    ch.UInt32Value(42) = offsets(i)
		    ch.StringValue(46, nameLen) = nm
		    centrals.Add ch
		  Next

		  Var cdOffset As Integer = offset
		  Var cdSize As Integer = 0
		  For Each ch As MemoryBlock In centrals
		    cdSize = cdSize + ch.Size
		  Next

		  Var eocd As New MemoryBlock(22)
		  eocd.LittleEndian = True
		  eocd.UInt32Value(0) = &h06054B50
		  eocd.UInt16Value(4) = 0
		  eocd.UInt16Value(6) = 0
		  eocd.UInt16Value(8) = locals.Count
		  eocd.UInt16Value(10) = locals.Count
		  eocd.UInt32Value(12) = cdSize
		  eocd.UInt32Value(16) = cdOffset
		  eocd.UInt16Value(20) = 0

		  Var total As Integer = cdOffset + cdSize + eocd.Size
		  Var out As New MemoryBlock(total)
		  Var pos As Integer = 0
		  For Each chunk As MemoryBlock In locals
		    out.StringValue(pos, chunk.Size) = chunk.StringValue(0, chunk.Size)
		    pos = pos + chunk.Size
		  Next
		  For Each chunk As MemoryBlock In centrals
		    out.StringValue(pos, chunk.Size) = chunk.StringValue(0, chunk.Size)
		    pos = pos + chunk.Size
		  Next
		  out.StringValue(pos, eocd.Size) = eocd.StringValue(0, eocd.Size)
		  Return out
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 417373656D626C6520746865206172636869766520616E6420777269746520697420746F2074686520676976656E2066696C65207669612042696E61727953747265616D2E0A
		Sub SaveTo(file As FolderItem)
		  Var mb As MemoryBlock = ToMemoryBlock
		  Var bs As BinaryStream = BinaryStream.Create(file, True)
		  bs.Write(mb.StringValue(0, mb.Size))
		  bs.Close
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 436F6D707265737320766961204D656D6F7279426C6F636B2E436F6D707265737320616E642073747269702074686520677A69702F7A6C6962207772617070657220646F776E20746F2074686520726177206465666C6174652073747265616D2061207A697020656E747279206E656564732E204E696C207768656E20756E7265636F676E697A656420E280942063616C6C65722073746F72657320756E636F6D707265737365642E0A
		Private Function RawDeflate(data As MemoryBlock) As MemoryBlock
		  ' MemoryBlock.Compress wraps its deflate stream (gzip or zlib depending on
		  ' the framework build). A zip entry needs the RAW deflate bytes, so detect
		  ' the wrapper from the magic bytes and strip header + integrity trailer.
		  ' Returns Nil when the wrapper is unrecognized — caller stores uncompressed.
		  Var comp As MemoryBlock
		  Try
		    comp = data.Compress(MemoryBlock.CompressionLevels.Best)
		  Catch
		    Return Nil
		  End Try
		  If comp Is Nil Or comp.Size < 12 Then Return Nil

		  Var b0 As Integer = comp.UInt8Value(0)
		  Var b1 As Integer = comp.UInt8Value(1)

		  If b0 = &h1F And b1 = &h8B Then
		    ' GZIP: 10-byte header + optional fields per FLG, 8-byte trailer.
		    Var flg As Integer = comp.UInt8Value(3)
		    Var pos As Integer = 10
		    If (flg And 4) <> 0 Then
		      If pos + 2 > comp.Size Then Return Nil
		      Var xlen As Integer = comp.UInt8Value(pos) + (comp.UInt8Value(pos + 1) * 256)
		      pos = pos + 2 + xlen
		    End If
		    If (flg And 8) <> 0 Then
		      While pos < comp.Size And comp.UInt8Value(pos) <> 0
		        pos = pos + 1
		      Wend
		      pos = pos + 1
		    End If
		    If (flg And 16) <> 0 Then
		      While pos < comp.Size And comp.UInt8Value(pos) <> 0
		        pos = pos + 1
		      Wend
		      pos = pos + 1
		    End If
		    If (flg And 2) <> 0 Then pos = pos + 2
		    Var deflateLen As Integer = comp.Size - 8 - pos
		    If deflateLen <= 0 Then Return Nil
		    Var outG As New MemoryBlock(deflateLen)
		    outG.StringValue(0, deflateLen) = comp.StringValue(pos, deflateLen)
		    Return outG
		  End If

		  If b0 = &h78 Then
		    ' ZLIB: 2-byte header, 4-byte Adler-32 trailer.
		    Var deflateLen As Integer = comp.Size - 2 - 4
		    If deflateLen <= 0 Then Return Nil
		    Var outZ As New MemoryBlock(deflateLen)
		    outZ.StringValue(0, deflateLen) = comp.StringValue(2, deflateLen)
		    Return outZ
		  End If

		  Return Nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5461626C652D6261736564204352432D3332202849454545203830322E332C20706F6C796E6F6D69616C20454442383833323029206F662074686520656E747279277320756E636F6D707265737365642062797465732C20617320746865207A697020666F726D61742072657175697265732E0A
		Private Function Crc32Of(data As MemoryBlock) As UInt32
		  ' Standard table-based CRC-32 (IEEE 802.3, polynomial EDB88320),
		  ' as required by the zip format for each entry's uncompressed bytes.
		  If mCrcTable.Count = 0 Then
		    Var kPoly As UInt32 = &hEDB88320
		    For i As Integer = 0 To 255
		      Var c As UInt32 = i
		      For k As Integer = 1 To 8
		        If (c And 1) = 1 Then
		          c = (c \ 2) Xor kPoly
		        Else
		          c = c \ 2
		        End If
		      Next
		      mCrcTable.Add c
		    Next
		  End If
		  Var allOnes As UInt32 = &hFFFFFFFF
		  Var crc As UInt32 = allOnes
		  Var n As Integer = data.Size
		  For i As Integer = 0 To n - 1
		    Var b As UInt32 = data.UInt8Value(i)
		    crc = mCrcTable((crc Xor b) And &hFF) Xor (crc \ 256)
		  Next
		  Return crc Xor allOnes
		End Function
	#tag EndMethod

	#tag Property, Flags = &h21, Description = 456E747279206E616D65732C20696E204164645061727420283D206172636869766529206F726465722E0A
		Private mNames() As String
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 556E636F6D7072657373656420656E7472792062797465732C20706172616C6C656C20746F206D4E616D65732E0A
		Private mDatas() As MemoryBlock
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 5472756520666F7263657320636F6D7072657373696F6E206D6574686F642030202873746F7265642920666F72207468617420656E7472792E0A
		Private mForceStored() As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 4C617A696C79206275696C74203235362D656E747279204352432D3332206C6F6F6B7570207461626C652E0A
		Private mCrcTable() As UInt32
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 444F532D656E636F646564206D6F64696669636174696F6E2074696D65207374616D706564206F6E20657665727920656E7472792E0A
		Private mDosTime As Integer
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 444F532D656E636F646564206D6F64696669636174696F6E2064617465207374616D706564206F6E20657665727920656E7472792E0A
		Private mDosDate As Integer
	#tag EndProperty

	#tag Note, Name = About
		Minimal in-memory ZIP archive writer — the write-side counterpart of
		XLSXZip's Memory backend. Used by XLSXWriter and ODSWriter.

		Usage:
		  Var zip As New SpreadsheetZipWriter
		  zip.AddPart("[Content_Types].xml", xmlString)
		  zip.AddPart("mimetype", mimeString, True)   ' forceStored
		  zip.SaveTo(file)            ' or ToMemoryBlock for in-memory use (Web)

		Why manual instead of FolderItem.Zip:
		  - FolderItem.Zip archives a staged folder and gives no control over entry
		    order or per-entry compression. The OpenDocument spec requires the
		    mimetype entry FIRST and STORED (method 0); AddPart order + forceStored
		    give us exactly that.
		  - No disk staging — pure in-memory, sandbox-friendly, mirrors the Memory
		    read backend.

		Compression: each entry is deflated via MemoryBlock.Compress, whose wrapper
		(gzip or zlib) is stripped down to the raw deflate stream a zip entry needs.
		If the wrapper is unrecognized or compression doesn't shrink the data, the
		entry is stored uncompressed (method 0) — always valid, just larger.

		CRC-32 is computed in-class (table-based); the framework has no public CRC-32.
		No zip64: fine for spreadsheet-sized archives (< 4 GB, < 65k entries).
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
