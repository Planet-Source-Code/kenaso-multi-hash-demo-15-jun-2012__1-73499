VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:  clsHash.cls
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-AUG-2004  Kenneth Ives  kenaso@tx.rr.com
'              Module created
' 31-Dec-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added property HashRounds()
' 02-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated references to RipeMD128_256 and RipeMD160_320.
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Add events to track hash progress on the screen.
' 25-Mar-2011  Kenneth Ives  kenaso@tx.rr.com 
'              Added reference to SHA-512/224 and SHA-512/256 as per
'              FIPS 180-4 dtd Mar-2012 (Supercedes FIPS 180-3)
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Combined all Tiger hashes into a single class.
'              - Removed reference to Tiger2 class.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME  As String = "clsHash"
  
' ***************************************************************************
' Enumerations
' ***************************************************************************
  Public Enum enumHASH_ALGORITHM
      eHASH_MD4            ' 0
      eHASH_MD5            ' 1
      eHASH_SHA1           ' 2
      eHASH_SHA224         ' 3
      eHASH_SHA256         ' 4
      eHASH_SHA384         ' 5
      eHASH_SHA512         ' 6
      eHASH_SHA512_224     ' 7
      eHASH_SHA512_256     ' 8
      eHASH_SHA512_320     ' 9
      eHASH_RIPEMD128      ' 10
      eHASH_RIPEMD160      ' 11
      eHASH_RIPEMD256      ' 12
      eHASH_RIPEMD320      ' 13
      eHASH_TIGER128       ' 14
      eHASH_TIGER160       ' 15
      eHASH_TIGER192       ' 16
      eHASH_TIGER224       ' 17
      eHASH_TIGER256       ' 18
      eHASH_TIGER384       ' 19
      eHASH_TIGER512       ' 20
      eHASH_WHIRLPOOL224   ' 21
      eHASH_WHIRLPOOL256   ' 22
      eHASH_WHIRLPOOL384   ' 23
      eHASH_WHIRLPOOL512   ' 24
  End Enum
  
' ***************************************************************************
' Module Variables
'
'                    +-------------- Module level designator
'                    |  +----------- Data type (Long)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m lng HashMethod
' Variable name:     mlngHashMethod
'
' ***************************************************************************
  Private mblnRetLowercase As Boolean
  Private mlngHashRounds   As Long
  Private mlngHashMethod   As enumHASH_ALGORITHM
    
  ' Define class objects with events (See bottom of this module)
  Private WithEvents mobjMD4       As cMD4
Attribute mobjMD4.VB_VarHelpID = -1
  Private WithEvents mobjMD5       As cMD5
Attribute mobjMD5.VB_VarHelpID = -1
  Private WithEvents mobjSHA1      As cSHA1
Attribute mobjSHA1.VB_VarHelpID = -1
  Private WithEvents mobjSHA2      As cSHA2
Attribute mobjSHA2.VB_VarHelpID = -1
  Private WithEvents mobjRIPEMD128 As cRipeMD128
Attribute mobjRIPEMD128.VB_VarHelpID = -1
  Private WithEvents mobjRIPEMD160 As cRipeMD160
Attribute mobjRIPEMD160.VB_VarHelpID = -1
  Private WithEvents mobjTIGER3    As cTiger3
Attribute mobjTIGER3.VB_VarHelpID = -1
  Private WithEvents mobjWHIRLPOOL As cWhirlpool
Attribute mobjWHIRLPOOL.VB_VarHelpID = -1
  

' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************
    
' Update progress bar
Public Event HashProgress(ByVal lngProgress As Long)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

' Comment out Version Property if not used in this application
Public Property Get Version() As String
    ' Version of the DLL this class is associated with.
    Version = DLL_NAME & ".dll (tm) v" & App.Major & "." & App.Minor & "." & App.Revision & _
              " Copyright (c) 2004-2012 - Kenneth Ives  kenaso@tx.rr.com"
End Property

Public Property Let StopProcessing(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.StopProcessing = True
    gblnStopProcessing = blnData
    DoEvents
End Property

Public Property Get StopProcessing() As Boolean
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.StopProcessing
    StopProcessing = gblnStopProcessing
    DoEvents
End Property

Public Property Let ReturnLowercase(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.ReturnLowercase = False
    mblnRetLowercase = blnData
End Property

Public Property Let HashMethod(ByVal lngHashMethod As enumHASH_ALGORITHM)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashMethod = eHASH_SHA256
    mlngHashMethod = lngHashMethod
End Property

Public Property Let HashRounds(ByVal lngData As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashRounds = 1

    Select Case mlngHashMethod
    
           ' MD5, RipeMD family, SHA family, Whirlpool family
           Case 0 To 13, 21 To 24
                Select Case lngData
                       Case 1 To 10: mlngHashRounds = lngData   ' Good choice
                       Case Else:    mlngHashRounds = 1         ' Use default value
                End Select

           Case 14 To 20    ' Tiger family
                Select Case lngData
                       Case 3 To 15: mlngHashRounds = lngData   ' Good choice
                       Case Else:    mlngHashRounds = 3         ' Use default value
                End Select
    End Select
    
End Property


' ***************************************************************************
' ****                         Methods                                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       HashString
'
' Description:   Creates a hash output string based on string data input.
'
' Parameters:    abytData() - Data string in byte array to be hashed
'
'    Example:    abytData() = StrConv("abc", vbFromUnicode)
'
' Returns:       Formatted hashed data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-AUG-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 20-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array.  Added functionality
'              to return data in various formats.
' 04-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a logic bug in output format.
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              Combined all Tiger hashes into a single class.
' ***************************************************************************
Public Function HashString(ByRef abytData() As Byte) As Variant

    Dim abytHash() As Byte
    
    On Error GoTo HashString_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Is there any data to process?
    If Not CBool(IsArrayInitialized(abytData())) Then
        gblnStopProcessing = True   ' Set flag to stop processing
        Exit Function
    End If
    
    Erase abytHash()  ' Always start with empty arrays
    ResetHashObjects  ' Verify hash objects are empty

    Select Case mlngHashMethod

           Case eHASH_MD4
                Set mobjMD4 = New cMD4   ' Instantiate class object
                With mobjMD4
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    abytHash() = .HashString(abytData())
                End With

           Case eHASH_MD5
                Set mobjMD5 = New cMD5   ' Instantiate class object
                With mobjMD5
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    abytHash() = .HashString(abytData())
                End With

           Case eHASH_SHA1
                Set mobjSHA1 = New cSHA1   ' Instantiate class object
                With mobjSHA1
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    abytHash() = .HashString(abytData())
                End With

           Case eHASH_SHA224, eHASH_SHA256, eHASH_SHA384, eHASH_SHA512, _
                eHASH_SHA512_224, eHASH_SHA512_256, eHASH_SHA512_320
                
                Set mobjSHA2 = New cSHA2   ' Instantiate class object
                With mobjSHA2
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_SHA224:     .HashMethod = 0
                           Case eHASH_SHA256:     .HashMethod = 1
                           Case eHASH_SHA384:     .HashMethod = 2
                           Case eHASH_SHA512:     .HashMethod = 3
                           Case eHASH_SHA512_224: .HashMethod = 4
                           Case eHASH_SHA512_256: .HashMethod = 5
                           Case eHASH_SHA512_320: .HashMethod = 6
                    End Select
                    abytHash() = .HashString(abytData())
                End With

           Case eHASH_RIPEMD128, eHASH_RIPEMD256
                Set mobjRIPEMD128 = New cRipeMD128   ' Instantiate class object
                With mobjRIPEMD128
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_RIPEMD128: .HashMethod = 0
                           Case eHASH_RIPEMD256: .HashMethod = 1
                    End Select
                    abytHash() = .HashString(abytData())
                End With

           Case eHASH_RIPEMD160, eHASH_RIPEMD320
                Set mobjRIPEMD160 = New cRipeMD160   ' Instantiate class object
                With mobjRIPEMD160
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_RIPEMD160: .HashMethod = 0
                           Case eHASH_RIPEMD320: .HashMethod = 1
                    End Select
                    abytHash() = .HashString(abytData())
                End With

           Case eHASH_TIGER128, eHASH_TIGER160, eHASH_TIGER192, _
                eHASH_TIGER224, eHASH_TIGER256, eHASH_TIGER384, _
                eHASH_TIGER512
                
                Set mobjTIGER3 = New cTiger3   ' Instantiate class object
                With mobjTIGER3
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_TIGER128: .HashMethod = 0
                           Case eHASH_TIGER160: .HashMethod = 1
                           Case eHASH_TIGER192: .HashMethod = 2
                           Case eHASH_TIGER224: .HashMethod = 3
                           Case eHASH_TIGER256: .HashMethod = 4
                           Case eHASH_TIGER384: .HashMethod = 5
                           Case eHASH_TIGER512: .HashMethod = 6
                    End Select
                    abytHash() = .HashString(abytData())
                End With

           Case eHASH_WHIRLPOOL224, eHASH_WHIRLPOOL256, _
                eHASH_WHIRLPOOL384, eHASH_WHIRLPOOL512
                
                Set mobjWHIRLPOOL = New cWhirlpool   ' Instantiate class object
                With mobjWHIRLPOOL
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_WHIRLPOOL224: .HashMethod = 0
                           Case eHASH_WHIRLPOOL256: .HashMethod = 1
                           Case eHASH_WHIRLPOOL384: .HashMethod = 2
                           Case eHASH_WHIRLPOOL512: .HashMethod = 3
                    End Select
                    abytHash() = .HashString(abytData())
                End With
    End Select

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        ReDim abytHash(0)  ' Resize to smallest number of elements
    End If

    HashString = abytHash()
           
HashString_CleanUp:
    Erase abytHash()  ' Always empty arrays when not needed
    ResetHashObjects  ' Verify hash objects are empty
    
    On Error GoTo 0   ' Nullify error trap in this routine
    Exit Function

HashString_Error:
    ErrorMsg MODULE_NAME, "HashString", Err.Description
    Resume HashString_CleanUp

End Function

' ***************************************************************************
' Routine:       HashFile
'
' Description:   Creates a hashed data string based on file contents.
'
' Parameters:    abytData() - Data file location in byte array format
'
'                Ex:  abytData() = StrConv("C:\Test Folder\Testfile.doc", vbFromUnicode)
'
' Returns:       Formatted hashed data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-AUG-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 20-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array.  Added functionality
'              to return data in various formats.
' 04-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a logic bug in output format.
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              Combined all Tiger hashes into a single class.
' ***************************************************************************
Public Function HashFile(ByRef abytData() As Byte) As Variant

    Dim abytHash() As Byte
    
    On Error GoTo HashFile_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Is there any data to process?
    If Not CBool(IsArrayInitialized(abytData())) Then
        gblnStopProcessing = True   ' Set flag to stop processing
        Exit Function
    End If
    
    Erase abytHash()  ' Always start with empty arrays
    ResetHashObjects  ' Verify hash objects are empty

    Select Case mlngHashMethod

           Case eHASH_MD4
                Set mobjMD4 = New cMD4   ' Instantiate class object
                With mobjMD4
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    abytHash() = .HashFile(abytData())
                End With

           Case eHASH_MD5
                Set mobjMD5 = New cMD5   ' Instantiate class object
                With mobjMD5
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    abytHash() = .HashFile(abytData())
                End With

           Case eHASH_SHA1
                Set mobjSHA1 = New cSHA1   ' Instantiate class object
                With mobjSHA1
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    abytHash() = .HashFile(abytData())
                End With

           Case eHASH_SHA224, eHASH_SHA256, eHASH_SHA384, eHASH_SHA512, _
                eHASH_SHA512_224, eHASH_SHA512_256, eHASH_SHA512_320
                
                Set mobjSHA2 = New cSHA2   ' Instantiate class object
                With mobjSHA2
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_SHA224:     .HashMethod = 0
                           Case eHASH_SHA256:     .HashMethod = 1
                           Case eHASH_SHA384:     .HashMethod = 2
                           Case eHASH_SHA512:     .HashMethod = 3
                           Case eHASH_SHA512_224: .HashMethod = 4
                           Case eHASH_SHA512_256: .HashMethod = 5
                           Case eHASH_SHA512_320: .HashMethod = 6
                    End Select
                    abytHash() = .HashFile(abytData())
                End With

           Case eHASH_RIPEMD128, eHASH_RIPEMD256
                Set mobjRIPEMD128 = New cRipeMD128   ' Instantiate class object
                With mobjRIPEMD128
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_RIPEMD128: .HashMethod = 0
                           Case eHASH_RIPEMD256: .HashMethod = 1
                    End Select
                    abytHash() = .HashFile(abytData())
                End With

           Case eHASH_RIPEMD160, eHASH_RIPEMD320
                Set mobjRIPEMD160 = New cRipeMD160   ' Instantiate class object
                With mobjRIPEMD160
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_RIPEMD160: .HashMethod = 0
                           Case eHASH_RIPEMD320: .HashMethod = 1
                    End Select
                    abytHash() = .HashFile(abytData())
                End With

           Case eHASH_TIGER128, eHASH_TIGER160, eHASH_TIGER192, _
                eHASH_TIGER224, eHASH_TIGER256, eHASH_TIGER384, _
                eHASH_TIGER512
                
                Set mobjTIGER3 = New cTiger3   ' Instantiate class object
                With mobjTIGER3
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_TIGER128: .HashMethod = 0
                           Case eHASH_TIGER160: .HashMethod = 1
                           Case eHASH_TIGER192: .HashMethod = 2
                           Case eHASH_TIGER224: .HashMethod = 3
                           Case eHASH_TIGER256: .HashMethod = 4
                           Case eHASH_TIGER384: .HashMethod = 5
                           Case eHASH_TIGER512: .HashMethod = 6
                    End Select
                    abytHash() = .HashFile(abytData())
                End With

           Case eHASH_WHIRLPOOL224, eHASH_WHIRLPOOL256, _
                eHASH_WHIRLPOOL384, eHASH_WHIRLPOOL512
                
                Set mobjWHIRLPOOL = New cWhirlpool   ' Instantiate class object
                With mobjWHIRLPOOL
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_WHIRLPOOL224: .HashMethod = 0
                           Case eHASH_WHIRLPOOL256: .HashMethod = 1
                           Case eHASH_WHIRLPOOL384: .HashMethod = 2
                           Case eHASH_WHIRLPOOL512: .HashMethod = 3
                    End Select
                    abytHash() = .HashFile(abytData())
                End With
    End Select

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        ReDim abytHash(0)  ' Resize to smallest number of elements
    End If

    HashFile = abytHash()
           
HashFile_CleanUp:
    Erase abytHash()  ' Always empty arrays when not needed
    ResetHashObjects  ' Verify hash objects are empty
    
    On Error GoTo 0   ' Nullify error trap in this routine
    Exit Function

HashFile_Error:
    ErrorMsg MODULE_NAME, "HashFile", Err.Description
    Resume HashFile_CleanUp

End Function

Private Sub ResetHashObjects()

    ' Free class objects from memory
    Set mobjMD4 = Nothing
    Set mobjMD5 = Nothing
    Set mobjSHA1 = Nothing
    Set mobjSHA2 = Nothing
    Set mobjRIPEMD128 = Nothing
    Set mobjRIPEMD160 = Nothing
    Set mobjTIGER3 = Nothing
    Set mobjWHIRLPOOL = Nothing

End Sub

Private Sub Class_Initialize()
    
    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objXXXX = new cXXXX
                
    ResetHashObjects  ' Verify hash objects are empty
    
    ' Preset property values
    StopProcessing = False      ' Continue processing
    HashMethod = eHASH_SHA256   ' Default algorithm
    ReturnLowercase = False     ' Default return of hashed data
    HashRounds = 1              ' Set default number of passes

End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objXXXX = Nothing
            
    ResetHashObjects

End Sub


' ***************************************************************************
' ****                      Events from subclasses                       ****
' ***************************************************************************
    
' Progress event is raised from the appropriate
' class and then raised from here to calling module

Private Sub mobjMD4_HashProgress(ByVal lngProgress As Long)
    RaiseEvent HashProgress(lngProgress)
End Sub

Private Sub mobjMD5_HashProgress(ByVal lngProgress As Long)
    RaiseEvent HashProgress(lngProgress)
End Sub

Private Sub mobjSHA1_HashProgress(ByVal lngProgress As Long)
    RaiseEvent HashProgress(lngProgress)
End Sub

Private Sub mobjSHA2_HashProgress(ByVal lngProgress As Long)
    RaiseEvent HashProgress(lngProgress)
End Sub

Private Sub mobjRIPEMD128_HashProgress(ByVal lngProgress As Long)
    RaiseEvent HashProgress(lngProgress)
End Sub

Private Sub mobjRIPEMD160_HashProgress(ByVal lngProgress As Long)
    RaiseEvent HashProgress(lngProgress)
End Sub

Private Sub mobjTIGER3_HashProgress(ByVal lngProgress As Long)
    RaiseEvent HashProgress(lngProgress)
End Sub

Private Sub mobjWHIRLPOOL_HashProgress(ByVal lngProgress As Long)
    RaiseEvent HashProgress(lngProgress)
End Sub


