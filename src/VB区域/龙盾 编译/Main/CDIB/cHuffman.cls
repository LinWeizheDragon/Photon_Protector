VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHuffman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************
'人人为我，我为人人
'枕善居汉化收藏整理
'发布日期：2009/07/29
'描    述：支持真彩色的图标列表控件(更新版)
'网    站：http://www.Mndsoft.com/  (VB6源码博客)
'网    站：http://www.VbDnet.com/   (VB.NET源码博客,主要基于.NET2005)
'e-mail  ：Mndsoft@163.com
'e-mail  ：Mndsoft@126.com
'OICQ    ：88382850
'          如果您有新的好的代码别忘记给枕善居哦!
'****************************************************************************


'Huffman Encoding/Decoding Class
'-------------------------------
' Guts borrowed from Fredrik Qvarfort
' posted at http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=11000&lngWId=1

' Modified and tweaked for custom header usage

Option Explicit

Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray        ' used as DMA overlay on a DIB
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SafeArrayBound ' 32 bytes as used. Can be used for 1D and/or 2D arrays
End Type

Private Type HUFFMANTREE
  ParentNode As Integer
  RightNode As Integer
  LeftNode As Integer
  Value As Integer
  Weight As Long
End Type

Private Type ByteArray
  Count As Byte
  Data() As Byte
End Type

Private potLUT(0 To 7) As Byte

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Friend Function EncodeListData(cHost As c32bppDIB, keyTag() As Byte, Result() As Byte) As Boolean
    ' dibBytes will be a DMA overlay of the ImageList DIB bits
    ' byteStart should be zero
    ' byteLen should be bitmap w*h*4
    ' Result will be the compressed bitmap using Huffman compression
    
    ' Header will look like following:
    ' @0, 4 bytes: huffmann compression header
    ' @+4, 4 bytes: size of original data
    '   when compressing image....
    '       @+8-12, 8 bytes: width/height of DIB
    '       @+16, 4 bytes: compressed size (if compressed)
    '       @+20, 2 bytes: nr of Huffman characters (if compressed)
    '       @+22, n*2 bytes: 2 bytes for each character (if compressed)
    '   when compressing key/tag info...
    '       @+8, 4 bytes: compressed size (if compressed)
    '       @+12, 2 bytes: nr of Huffman characters (if compressed)
    '       @+14, n*2 bytes: 2 bytes for each character (if compressed)
    ' nnnn the compressed data
  
  Dim dibBytes() As Byte
  Dim dibSA As SafeArray
  Dim i As Long
  Dim j As Long
  Dim Char As Byte
  Dim BitPos As Byte
  Dim lNode1 As Long
  Dim lNode2 As Long
  Dim lNodes As Long
  Dim lLength As Long
  Dim Count As Integer
  Dim lWeight1 As Long
  Dim lWeight2 As Long
  Dim ByteValue As Byte
  Dim ResultLen As Long
  Dim Bytes As ByteArray
  Dim NodesCount As Integer
  Dim CharCount() As Long
  Dim Nodes() As HUFFMANTREE
  Dim CharValue(0 To 255) As ByteArray
  
  Dim totalChars As Integer
  Dim currPos As Long
  Dim arrLen As Long
  'Count the frequency of each ASCII code
  
  If cHost Is Nothing Then
    arrLen = Abs(UBound(keyTag) - LBound(keyTag)) + 1
    iparseOverlayHost_Byte dibBytes(), VarPtr(dibSA), 1, arrLen, 0, VarPtr(keyTag(LBound(keyTag)))
  Else
    arrLen = cHost.scanWidth * cHost.Height
    iparseOverlayHost_Byte dibBytes(), VarPtr(dibSA), 1, arrLen, 0, cHost.BitsPointer
  End If
  
  ReDim CharCount(0 To 255)
  For i = 0 To arrLen - 1
    CharCount(dibBytes(i)) = CharCount(dibBytes(i)) + 1
  Next
  
  'Create a leaf for each character
  ReDim Nodes(0 To 511)
  For i = 0 To 255
    If (CharCount(i) > 0) Then
      With Nodes(NodesCount)
        .Weight = CharCount(i)
        .Value = i
        .LeftNode = -1
        .RightNode = -1
        .ParentNode = -1
      End With
      NodesCount = NodesCount + 1
    End If
  Next
  
  'Create the Huffman Tree
  For lNodes = NodesCount To 2 Step -1
    'Get the two leafs with the smallest weights
    lNode1 = -1: lNode2 = -1
    For i = 0 To (NodesCount - 1)
      If (Nodes(i).ParentNode = -1) Then
        If (lNode1 = -1) Then
          lWeight1 = Nodes(i).Weight
          lNode1 = i
        ElseIf (lNode2 = -1) Then
          lWeight2 = Nodes(i).Weight
          lNode2 = i
        ElseIf (Nodes(i).Weight < lWeight1) Then
          If (Nodes(i).Weight < lWeight2) Then
            If (lWeight1 < lWeight2) Then
              lWeight2 = Nodes(i).Weight
              lNode2 = i
            Else
              lWeight1 = Nodes(i).Weight
              lNode1 = i
            End If
          Else
            lWeight1 = Nodes(i).Weight
            lNode1 = i
          End If
        ElseIf (Nodes(i).Weight < lWeight2) Then
          lWeight2 = Nodes(i).Weight
          lNode2 = i
        End If
      End If
    Next
    
    'Create a new leaf
    With Nodes(NodesCount)
      .Weight = lWeight1 + lWeight2
      .LeftNode = lNode1
      .RightNode = lNode2
      .ParentNode = -1
      .Value = -1
    End With
    
    'Set the parentnodes of the two leafs
    Nodes(lNode1).ParentNode = NodesCount
    Nodes(lNode2).ParentNode = NodesCount
    
    'Increase the node counter
    NodesCount = NodesCount + 1
  Next

  'Traverse the tree to get the bit sequence
  'for each character, make temporary room in
  'the data array to hold max theoretical size
  ReDim Bytes.Data(0 To 255)
  Call CreateBitSequences(Nodes(), NodesCount - 1, Bytes, CharValue)
  Erase Bytes.Data()
  Erase Nodes
  
  'Calculate the length of the destination string after encoding
  For i = 0 To 255
    If (CharCount(i) > 0) Then
      lLength = lLength + CharValue(i).Count * CharCount(i)
    End If
  Next
  Erase CharCount
  lLength = lLength \ 8 - ((lLength Mod 8) > 0)
  
  'If the destination is larger than the source
  'string we leave it uncompressed and prefix
  'it with a 4 byte header ("HE0" & vbCr)
  If ((lLength = 0) Or (lLength > arrLen)) Then
    ReDim Preserve Result(0 To arrLen + 7)
    CopyMemory Result(0), 1212493837, 4& ' 1212493837 = 72,69,48,13 (uncompressed data)
    CopyMemory Result(4), arrLen, 4&
  
  Else
      'Create a small array to hold the bit values, this is faster than calculating on-fly
      If potLUT(0) = 0 Then
        potLUT(0) = 1
        For i = 1 To 7
          potLUT(i) = potLUT(i - 1) * 2
        Next
      End If
      
      'Store the number of characters used
      Count = 0
      For i = 0 To 255
        If (CharValue(i).Count > 0) Then
          totalChars = totalChars + 1
          Count = Count + 16 + CharValue(i).Count
        End If
      Next
     
      'Make room for the Huffman Tree in the destination byte array
      If cHost Is Nothing Then
        ResultLen = 4 + 4 + 4 + totalChars * 2 + 2
      Else
      '^^ 4=HEader, 4=origSize, 4=compressSize, totalChars*2= charvals used/value, 2=total chars
        ResultLen = 4 + 4 + 4 + 8 + 2 + totalChars * 2
      End If
      ReDim Result(0 To ResultLen + Count \ 8)
      
      'Store the Huffman Tree into the result
      'converting the bit sequences into bytes
      BitPos = 0
      ByteValue = 0
      For i = 0 To 255
        With CharValue(i)
          If (.Count > 0) Then
            For j = 0 To (.Count - 1)
              If (.Data(j)) Then ByteValue = ByteValue + potLUT(BitPos)
              BitPos = BitPos + 1
              If (BitPos = 8) Then
                Result(ResultLen) = ByteValue
                ResultLen = ResultLen + 1
                ByteValue = 0
                BitPos = 0
              End If
            Next
          End If
        End With
      Next
      If (BitPos > 0) Then
        Result(ResultLen) = ByteValue
        ResultLen = ResultLen + 1
      End If
      
      'Resize the destination to be able to
      'contain the encoded string
      ReDim Preserve Result(0 To ResultLen + lLength)
      
      'Now we can encode the data by exchanging each
      'ASCII byte for its appropriate bit string.
      Char = 0
      BitPos = 0
      For i = 0 To arrLen - 1
        With CharValue(dibBytes(i))
          For j = 0 To (.Count - 1)
            If (.Data(j) = 1) Then Char = Char + potLUT(BitPos)
            BitPos = BitPos + 1
            If (BitPos = 8) Then
              Result(ResultLen) = Char
              ResultLen = ResultLen + 1
              BitPos = 0
              Char = 0
            End If
          Next
        End With
      Next
    
      'Add the last byte
      If (BitPos > 0) Then Result(ResultLen) = Char
      
      ' write header & char count used
      CopyMemory Result(0), 1212494605, 4& '1212494605 = 72,69,51,13
      CopyMemory Result(4), arrLen, 4&
      If cHost Is Nothing Then
        CopyMemory Result(8), CLng(UBound(Result) + 1&), 4&
        CopyMemory Result(12), totalChars, 2&
        ResultLen = 14
      Else
        CopyMemory Result(16), CLng(UBound(Result) + 1&), 4&
        CopyMemory Result(20), totalChars, 2&
        ResultLen = 22
      End If
      ' now write which characters & their count
      For i = 0 To 255
        If (CharValue(i).Count > 0) Then
          Result(ResultLen) = i
          Result(ResultLen + 1) = CharValue(i).Count
          ResultLen = ResultLen + 2
        End If
      Next
      
      ' done
  End If
  
  If Not cHost Is Nothing Then
    CopyMemory Result(8), cHost.Width, 4&
    CopyMemory Result(12), cHost.Height, 4&
  End If
  
  iparseOverlayHost_Byte dibBytes, 0, 0, 0, 0, 0
  EncodeListData = True
  
End Function

Friend Function DecodeByte(inArray() As Byte, byteStart As Long, keyTag() As Byte, cHost As c32bppDIB) As Boolean
  
    ' Header will look like following:
    ' Header will look like following:
    ' @0, 4 bytes: huffmann compression header
    ' @+4, 4 bytes: size of original data
    '   when compressing image....
    '       @+8-12, 8 bytes: width/height of DIB
    '       @+16, 4 bytes: compressed size (if compressed)
    '       @+20, 2 bytes: nr of Huffman characters (if compressed)
    '       @+22, n*2 bytes: 2 bytes for each character (if compressed)
    '   when compressing key/tag info...
    '       @+8, 4 bytes: compressed size (if compressed)
    '       @+12, 2 bytes: nr of Huffman characters (if compressed)
    '       @+14, n*2 bytes: 2 bytes for each character (if compressed)
    ' nnnn the compressed data
  
  Dim i As Long
  Dim j As Long
  Dim Pos As Long
  Dim Char As Byte
  Dim currPos As Long
  Dim Count As Integer
  Dim Result() As Byte, resultSA As SafeArray
  Dim BitPos As Integer
  Dim NodeIndex As Long
  Dim ByteValue As Byte
  Dim ResultLen As Long
  Dim NodesCount As Long
  Dim lResultLen As Long
  Dim Nodes() As HUFFMANTREE
  Dim CharValue() As ByteArray
  
  Dim lCompressSize As Long
  
  If byteStart + 16 > UBound(inArray) Then Exit Function ' invalid array size
  CopyMemory lResultLen, inArray(byteStart + 4), 4&
  
  CopyMemory currPos, inArray(byteStart), 4&
  If (currPos And &HFFFF00FF) = 1212481549 Then
    If ((currPos \ &H100) And &HFF) = 48 Then
        'The data is uncompressed, return the raw data
        
        If cHost Is Nothing Then
            If byteStart + lResultLen + 13 > UBound(inArray) Then Exit Function ' invalid array
            ReDim keyTag(1 To lResultLen)
            CopyMemory keyTag(1), inArray(byteStart + 14), lResultLen
        Else
            If byteStart + lResultLen + 21 > UBound(inArray) Then Exit Function ' invalid array
            CopyMemory i, inArray(byteStart + 8), 4&
            CopyMemory j, inArray(byteStart + 12), 4&
            cHost.InitializeDIB i, j
            CopyMemory ByVal cHost.BitsPointer, inArray(byteStart + 16), lResultLen
        End If
        DecodeByte = True
        Exit Function
    ElseIf ((currPos \ &H100) And &HFF) = 51 Then
        If cHost Is Nothing Then
            currPos = byteStart + 8
        Else
            currPos = byteStart + 16
        End If
    Else
        Exit Function
    End If
  Else
    Exit Function
  End If
  
  CopyMemory lCompressSize, inArray(currPos), 4&
  If byteStart + lCompressSize - 1 > UBound(inArray) Then Exit Function ' invalid array size
  
  'Get the number of characters used
  Call CopyMemory(Count, inArray(currPos + 4), 2)
  If Count < 1 Or Count > 256 Then Exit Function ' corrupted data
  
  'Get the used characters and their
  'respective bit sequence lengths
  currPos = currPos + 6
  ReDim CharValue(0 To 255)
  For i = 1 To Count
    With CharValue(inArray(currPos))
      .Count = inArray(currPos + 1)
      ReDim .Data(0 To .Count - 1)
      currPos = currPos + 2
    End With
  Next
  
  'Create a small array to hold the bit values,
  'this is (still) faster than calculating on-fly
  If potLUT(0) = 0 Then
    potLUT(0) = 1
    For i = 1 To 7
      potLUT(i) = potLUT(i - 1) * 2
    Next
  End If
  'Extract the Huffman Tree, converting the
  'byte sequence to bit sequences
  ByteValue = inArray(currPos)
  BitPos = 0
  For i = 0 To 255
    With CharValue(i)
      If (.Count > 0) Then
        For j = 0 To (.Count - 1)
          If (ByteValue And potLUT(BitPos)) Then .Data(j) = 1
          If (BitPos = 7) Then
            currPos = currPos + 1
            ByteValue = inArray(currPos)
            BitPos = 0
          Else
            BitPos = BitPos + 1
          End If
        Next
      End If
    End With
  Next
  If BitPos > 0 Then currPos = currPos + 1
  
  'Create the Huffman Tree
  NodesCount = 1
  ReDim Nodes(0 To 511)
  Nodes(0).LeftNode = -1
  Nodes(0).RightNode = -1
  Nodes(0).ParentNode = -1
  Nodes(0).Value = -1
  For i = 0 To 255
    Call CreateTree(Nodes(), NodesCount, i, CharValue(i))
  Next
  
  'Create the result array
  If cHost Is Nothing Then
    ReDim keyTag(1 To lResultLen)
    iparseOverlayHost_Byte Result(), VarPtr(resultSA), 1, lResultLen, 0, VarPtr(keyTag(1))
  Else
    
    CopyMemory i, inArray(byteStart + 8), 4&
    CopyMemory j, inArray(byteStart + 12), 4&
    cHost.InitializeDIB i, j
    iparseOverlayHost_Byte Result(), VarPtr(resultSA), 1, lResultLen, 0, cHost.BitsPointer
  End If
  
  ResultLen = 0
  'Decode the actual data
  For currPos = currPos To currPos + lCompressSize - 1
    ByteValue = inArray(currPos)
    For BitPos = 0 To 7
      If (ByteValue And potLUT(BitPos)) Then
        NodeIndex = Nodes(NodeIndex).RightNode
      Else
        NodeIndex = Nodes(NodeIndex).LeftNode
      End If
      If (Nodes(NodeIndex).Value > -1) Then
        Result(ResultLen) = Nodes(NodeIndex).Value
        ResultLen = ResultLen + 1
        If (ResultLen = lResultLen) Then
            currPos = currPos + lCompressSize
            Exit For
        End If
        NodeIndex = 0
      End If
    Next
  Next

  iparseOverlayHost_Byte Result(), 0, 0, 0, 0, 0
  DecodeByte = True
End Function

Private Sub CreateBitSequences(Nodes() As HUFFMANTREE, ByVal NodeIndex As Integer, Bytes As ByteArray, CharValue() As ByteArray)
  
  Dim newBytes As ByteArray
  
  'If this is a leaf we set the characters bit
  'sequence in the CharValue array
  If (Nodes(NodeIndex).Value > -1) Then
    CharValue(Nodes(NodeIndex).Value) = Bytes
    Exit Sub
  End If
  
  'Traverse the left child
  If (Nodes(NodeIndex).LeftNode > -1) Then
    newBytes = Bytes
    newBytes.Data(newBytes.Count) = 0
    newBytes.Count = newBytes.Count + 1
    Call CreateBitSequences(Nodes(), Nodes(NodeIndex).LeftNode, newBytes, CharValue)
  End If
  
  'Traverse the right child
  If (Nodes(NodeIndex).RightNode > -1) Then
    newBytes = Bytes
    newBytes.Data(newBytes.Count) = 1
    newBytes.Count = newBytes.Count + 1
    Call CreateBitSequences(Nodes(), Nodes(NodeIndex).RightNode, newBytes, CharValue)
  End If
  
End Sub

Private Sub CreateTree(Nodes() As HUFFMANTREE, NodesCount As Long, Char As Long, Bytes As ByteArray)

  Dim a As Integer
  Dim NodeIndex As Long
  NodeIndex = 0
  For a = 0 To (Bytes.Count - 1)
    If (Bytes.Data(a) = 0) Then
      'Left node
      If (Nodes(NodeIndex).LeftNode = -1) Then
        Nodes(NodeIndex).LeftNode = NodesCount
        Nodes(NodesCount).ParentNode = NodeIndex
        Nodes(NodesCount).LeftNode = -1
        Nodes(NodesCount).RightNode = -1
        Nodes(NodesCount).Value = -1
        NodesCount = NodesCount + 1
      End If
      NodeIndex = Nodes(NodeIndex).LeftNode
    ElseIf (Bytes.Data(a) = 1) Then
      'Right node
      If (Nodes(NodeIndex).RightNode = -1) Then
        Nodes(NodeIndex).RightNode = NodesCount
        Nodes(NodesCount).ParentNode = NodeIndex
        Nodes(NodesCount).LeftNode = -1
        Nodes(NodesCount).RightNode = -1
        Nodes(NodesCount).Value = -1
        NodesCount = NodesCount + 1
      End If
      NodeIndex = Nodes(NodeIndex).RightNode
    Else
      Stop
    End If
  Next
  
  Nodes(NodeIndex).Value = Char

End Sub

