VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRipeMD128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Project:       RipeMD128
'
' Module:        clsRipeMD128_256.cls
'
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
' *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING ***
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
'    You acknowledge that this software is subject to the export control
'    laws and regulations of the United States ("U.S.") and agree to abide
'    by those laws and regulations. Under U.S. law, this software may not
'    be downloaded or otherwise exported, reexported, or transferred to
'    restricted countries, restricted end-users, or for restricted
'    end-uses. The U.S. currently has embargo restrictions against Cuba,
'    Iran, Iraq, Libya, North Korea, Sudan, and Syria. The lists of
'    restricted end-users are maintained on the U.S. Commerce Department's
'    Denied Persons List, the Commerce Department's Entity List, the
'    Commerce Department's List of Unverified Persons, and the U.S.
'    Treasury Department's List of Specially Designated Nationals and
'    Blocked Persons. In addition, this software may not be downloaded or
'    otherwise exported, reexported, or transferred to an end-user engaged
'    in activities related to weapons of mass destruction.
'
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
' Description:   RipeMD is an iterative hash function that operates on
'                32-bit words.   The round function takes as input a 10-word
'                chaining variable and a 16-word message block and maps this
'                to a new chaining variable. All operations are defined on
'                32-bit words. Padding is identical to that of MD4 or MD5.
'
' ***************************************************************************
' Copyright Information
'
'      AUTHOR:   Antoon Bosselaers, ESAT-COSIC
'      DATE:     1 March 1996
'      VERSION:  1.0
'
'      Copyright (c) Katholieke Universiteit Leuven
'      1996, All Rights Reserved
'
'  Conditions for use of the RipeMD-128, -160, -256, -320 Software
'  henceforth refered to as "RIPEMD-X".
'
'  The RIPEMD-X software is freely available for use under the terms and
'  conditions described hereunder, which shall be deemed to be accepted by
'  any user of the software and applicable on any use of the software:
'
'  1. K.U.Leuven Department of Electrical Engineering-ESAT/COSIC shall for
'     all purposes be considered the owner of the RIPEMD-X software and of
'     all copyright, trade secret, patent or other intellectual property
'     rights therein.
'  2. The RIPEMD-X software is provided on an "as is" basis without
'     warranty of any sort, express or implied. K.U.Leuven makes no
'     representation that the use of the software will not infringe any
'     patent or proprietary right of third parties. User will indemnify
'     K.U.Leuven and hold K.U.Leuven harmless from any claims or liabilities
'     which may arise as a result of its use of the software. In no
'     circumstances K.U.Leuven R&D will be held liable for any deficiency,
'     fault or other mishappening with regard to the use or performance of
'     the software.
'  3. User agrees to give due credit to K.U.Leuven in scientific publications
'     or communications in relation with the use of the RIPEMD-X software
'     as follows: RIPEMD-X software written by Antoon Bosselaers,
'     available at:
'
'     http://www.esat.kuleuven.ac.be/~cosicart/ps/AB-9601/
'     http://homes.esat.kuleuven.be/~bosselae/ripemd160.html
'
'  RIPEMD-160 is a 160-bit cryptographic hash function, designed by Hans
'  Dobbertin, Antoon Bosselaers, and Bart Preneel. It is intended to be
'  used as a secure replacement for the 128-bit hash functions MD4, MD5,
'  and RIPEMD. MD4 and MD5 were developed by Ron Rivest for RSA Data
'  Security, while RIPEMD was developed in the framework of the EU project
'  RIPE (RACE Integrity Primitives Evaluation, 1988-1992).
'
'  RIPEMD-160 is a strengthened version of RIPEMD with a 160-bit hash
'  result, and is expected to be secure for the next ten years or more. The
'  design philosophy is to build as much as possible on experience gained
'  by evaluating MD4, MD5, and RIPEMD. Like its predecessors, RIPEMD-160 is
'  tuned for 32-bit processors, which we feel will remain important in the
'  coming decade.
'
'  RIPEMD-128 is a plug-in substitute for RIPEMD (or MD4 and MD5, for that
'  matter) with a 128-bit result. In view of the result of Paul van Oorschot
'  and Mike Wiener mentioned earlier, 128-bit hash results do not offer
'  sufficient protection for the next ten years, and applications using
'  128-bit hash functions should consider upgrading to a 160-bit hash
'  function.
'
'  RIPEMD-256 and RIPEMD-320 are optional extensions of, respectively,
'  RIPEMD-128 and RIPEMD-160, and are intended for applications of hash
'  functions that require a longer hash result without needing a larger
'  security level.  If you require a stronger hash, I recommend you use
'  a member of the SHA2 family of algorithms.
'
' ***************************************************************************
'
'        RIPEMD Algorithm has been compromised
'
'        At the rump session of Crypto 2004 it was announced that Xiaoyun
'        Wang, Dengguo Feng, Xuejia Lai and Hongbo Yu found collisions for
'        MD4, MD5, RIPEMD, and the 128-bit version of HAVAL.
'        http://eprint.iacr.org/2004/199.pdf
'
' ***************************************************************************
'
'  Message Hash results using one (1) pass
'
'      vbNullString (empty string)
'          RIPEMD-128   cdf26213a150dc3ecb610f18f6b38b46
'          RIPEMD-256   02ba4c4e5f8ecd1877fc52d64d30e37a2d9774fb1e5d026380ae0168e3c5522d
'
'      "a"
'          RIPEMD-128   86be7afa339d0fc7cfc785e72f578d33
'          RIPEMD-256   f9333e45d857f5d90a91bab70a1eba0cfb1be4b0783c9acfcd883a9134692925
'
'      "abc"
'          RIPEMD-128   c14a12199c66e4ba84636b0f69144c77
'          RIPEMD-256   afbd6e228b9d8cbbcef5ca2d03e6dba10ac0bc7dcbe4680e1e42d2e975459b65
'
'      "message digest"
'          RIPEMD-128   9e327b3d6e523062afc1132d7df9d1b8
'          RIPEMD-256   87e971759a1ce47a514d5c914c392c9018c7c46bc14465554afcdf54a5070c0e
'
'      "abcdefghijklmnopqrstuvwxyz"
'          RIPEMD-128   fd2aa607f71dc8f510714922b371834e
'          RIPEMD-256   649d3034751ea216776bf9a18acc81bc7896118a5197968782dd1fd97d8d5133
'
'      "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"
'          RIPEMD-128   a1aa0689d0fafa2ddc22e88b49133a06
'          RIPEMD-256   3843045583aac6c8c8d9128573e7a9809afb2a0f34ccc36ea9e72f16f6368e3f
'
'      "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
'          RIPEMD-128   d1e959eb179c911faea4624c60c5c702
'          RIPEMD-256   5740a408ac16b720b84424ae931cbb1fe363d1d0bf4017f1a89f7ea6de77a0b8
'
'      8 times "1234567890"
'          RIPEMD-128   3f45ef194732c2dbb2c4a2c769795fa3
'          RIPEMD-256   06fdcc7a409548aaf91368c06a6275b553e3f099bf0ea4edfd6778df89a890dd
'
'      1 million letter "a"
'          RIPEMD-128   4a7f5723f954eba1216c9d8f6320431f
'          RIPEMD-256   ac953744e10e31514c150d4d8d7b677342e33399788296e43ae4850ce4f97978
'
'      1 million binary zeroes
'          RIPEMD-128   c635f3bfcf17fc0f18bf7c467854bde4
'          RIPEMD-256   6b2f39fb5df66cf987ad61dddae56cbfafb8097e6fe905243e8f2df6bf57eedd
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Rewrote this module.
'              - Verified it will handle files larger than 2gb.
'              - Updated algorithms to more closely resemble the RipeMD-160,
'                C code and the RipeMD-320 pseudocode located at:
'                http://homes.esat.kuleuven.be/~bosselae/ripemd160.html
'                Test results for RipeMD-160 are accurate. However,
'                RipeMD-256 test results do not match the results on the
'                web page because all I had to reference was pseudocode.
'              - Added property HashRounds() to set number of rounds for
'                hashing.
'              - Added property HashMethod() to determine hash algorithm.
'              - Thanks to Blazer for letting me know of bad hash outputs.
'                I apparently forgot to insert a line of code that formats
'                the bit size in hex in HashFile() and HashString()
'                routines.
' 15-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Thanks again to Blazer for his assistance and persistance
'                in correcting the results for RipeMD-256 and RipeMD-320.
'                The hashed results are now accurate in accordance with
'                http://homes.esat.kuleuven.be/~bosselae/ripemd160.html
'              ===========================================================
'              Comments about RipeMD-x hash algorithm
'                ** RIPEMD-256 and RIPEMD-320 provide no better security
'                   than RIPEMD-160. Therefore, they are redundant, and can
'                   only serve to mislead users into thinking that they
'                   provide additional security.
'                   http://lists.dlitz.net/pipermail/pycrypto/2008q4/000039.html
'                ** My observations:  The pseudocode on the RipeMD-160
'                   web page for RipeMD-256 and RipeMD-320 does not even
'                   come close to the actual code changes required for the
'                   displayed results. The pseudocode appears to be stronger
'                   but I may be wrong. However, it seems as if the person
'                   who wrote the pseudocode did not communicate with the
'                   person who made the code changes.  Only two minor coding
'                   tweaks were ever made.  See Transform() routine.
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Add events to track hash progress on the screen.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "clsRipeMD128_256"
  Private Const MAX_ROUNDS  As Long = 10
  Private Const MSG_LEN_32  As Long = 128          ' 128 bytes long
  Private Const KB_32       As Long = &H8000&      ' 32768

' ***************************************************************************
' Enumerations
' ***************************************************************************
  Private Enum enumRIPEMD_ALGORITHM
      eRIPEMD128  ' 0
      eRIPEMD256  ' 1
  End Enum

' ***************************************************************************
' Module Variables
'
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Long)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a lng Hash
' Variable name:     malngHash
'
' ***************************************************************************
  Private mobjMath32       As cMath32
  Private mlngHashMethod   As enumRIPEMD_ALGORITHM
  Private mstrBitsInHex    As String
  Private mblnRetLowercase As Boolean
  Private mcurMaxSize      As Currency
  Private mcurAccrued      As Currency
  Private mcurBitCount     As Currency
  Private mcurCurrBitCnt   As Currency
  Private mlngHashRounds   As Long
  Private malngHash()      As Long
  Private mabytS1()        As Byte
  Private mabytS2()        As Byte
  Private mabytW1()        As Byte
  Private mabytW2()        As Byte


' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************

' Update progress bar
Public Event HashProgress(ByVal lngProgress As Long)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let HashMethod(ByVal lngData As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashMethod = eRIPEMD128
    Select Case lngData
           Case 0 To 1: mlngHashMethod = lngData
           Case Else:   mlngHashMethod = eRIPEMD128   ' Use default value
    End Select

End Property

Public Property Let HashRounds(ByVal lngData As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashRounds = 1
    Select Case lngData
           Case 1 To MAX_ROUNDS: mlngHashRounds = lngData   ' Good selection
           Case Else:            mlngHashRounds = 1         ' Default value
    End Select

End Property

Public Property Let ReturnLowercase(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.ReturnLowercase = False
    mblnRetLowercase = blnData
End Property


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       HashFile
'
' Description:   Function to create a unique hex string representation of 
'                the data passed.  Can process files larger than 2 gb.
'                See Readme.txt file for more information. 
'
' Syntax:
'    With mobjRipeMD
'        abytData() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'        mabytHashed() = .HashFile(abytData())                        ' hash data and return as Byte array
'        txtData.Text = StrConv(mabytHashed(), vbUnicode)             ' convert byte array to string data
'    End With
' 
' Parameters:    strSource - Name of file to be hashed
'
' Returns:       Digital signature of the file
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 31-Dec-2009  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Blazer for letting me know of bad hash outputs.
'              I apparently forgot to insert a line of code that formats the
'              bit size in hex.
' ***************************************************************************
Public Function HashFile(ByRef abytFile() As Byte) As Byte()

    Dim hFile        As Long
    Dim lngBlockSize As Long
    Dim curAmtLeft   As Currency
    Dim curFilePos   As Currency
    Dim strSource    As String
    Dim abytData()   As Byte
    Dim objBigFiles  As cBigFiles

    On Error GoTo HashFile_CleanUp

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    Set objBigFiles = New cBigFiles  ' Instantiate class object
    ResetVariables                   ' Reset class variables
    Initialize                       ' load the work and constant arrays
    mcurAccrued = 0@                 ' Init accrued progress
    curFilePos = 0@                  ' set to first position in the file
    Erase abytData()                 ' make sure array is empty

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    With objBigFiles
        strSource = ByteArrayToString(abytFile())            ' Convert byte array to string
        .CalcFileSize strSource, curAmtLeft, mstrBitsInHex   ' Get the size of the file

        mcurBitCount = curAmtLeft * 8                                      ' Get total bit count for comparison
        mstrBitsInHex = Right$(String$(16, "0") & Hex$(mcurBitCount), 16)  ' Convert bit count into hex
        mstrBitsInHex = mobjMath32.SwapEndianHex(mstrBitsInHex)            ' Reverse hex values in string

        ' Calculate amount with number
        ' of hashing rounds requested
        mcurMaxSize = curAmtLeft * mlngHashRounds

        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo HashFile_CleanUp
        End If

        ' Open source file
        If Not .OpenReadOnly(strSource, hFile) Then
            gblnStopProcessing = True
            GoTo HashFile_CleanUp
        End If

        Do
            Erase abytData()                              ' empty data array
            lngBlockSize = GetBlockSize(curAmtLeft)       ' Process the source file in chunks
            ReDim abytData(lngBlockSize - 1)              ' Size receiving array
            curAmtLeft = curAmtLeft - CCur(lngBlockSize)  ' Adjust amount of file left to process

            ' Read source file
            If Not .API_ReadFile(hFile, curFilePos, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If

            ' Adjust pointers accordingly
            curFilePos = curFilePos + CCur(UBound(abytData) + 1)

            ' Process the message data into the appropriate
            ' string length and perform the bit manipulations.
            If Not LoadWordArray(abytData()) Then
                gblnStopProcessing = True
            End If

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do   ' exit Do..Loop
            End If

            ' If nothing else to read then leave
            DoEvents
            If curAmtLeft < 1 Then
                Exit Do  ' exit Do..Loop
            End If

        Loop

        .API_CloseFile hFile    ' Close all files opened by this routine
    End With

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    ' Concatenate the state array into one data
    ' string and return the digital signature.
    HashFile = FormatOutput()

HashFile_CleanUp:
    objBigFiles.API_CloseFile hFile  ' Close all files opened by this routine
    ResetVariables                   ' Reset class variables
    Erase abytData()                 ' Always empty arrays when not needed
    Set objBigFiles = Nothing        ' Always free objects from memory

End Function

' ***************************************************************************
' Routine:       HashString
'
' Description:   Function to digest a text string and output the result as
'                a string of hexadecimal characters.  Got basic idea from
'                David Ireland of DI Management Services Pty Ltd
'                <www.di-mgt.com.au>.
'
' Syntax:
'    With mobjRipeMD
'        abytData() = StrConv("abc", vbFromUnicode)      ' convert string data to byte array 
'        abytHashed() = .HashString(abytData())          ' hash data and return as Byte array
'        txtData.Text = StrConv(abytHashed(), vbUnicode) ' convert byte array to string data
'    End With
' 
' Parameters:    abytData() - Path/filename in byte array to be hashed
'
' Returns:       Hashed data string in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Feb-2006  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array for security purposes
' 31-Dec-2009  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Blazer for letting me know of bad hash outputs.
'              I apparently forgot to insert a line of code that formats the
'              bit size in hex.
' ***************************************************************************
Public Function HashString(ByRef abytData() As Byte) As Byte()

    On Error GoTo HashString_CleanUp

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

    ResetVariables   ' Reset class variables
    Initialize       ' load work and constant arrays

    mcurAccrued = 0@                                                   ' Init accrued progress
    mcurMaxSize = CCur((UBound(abytData) + 1) * mlngHashRounds)        ' Capture data size
    mcurBitCount = CCur((UBound(abytData) + 1) * 8@)                   ' Get bit count for for comparison
    mstrBitsInHex = Right$(String$(16, "0") & Hex$(mcurBitCount), 16)  ' Convert bit count into hex
    mstrBitsInHex = mobjMath32.SwapEndianHex(mstrBitsInHex)            ' Reverse hex values in string

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashString_CleanUp
    End If

    ' Process the message data into the appropriate
    ' string length and perform the bit manipulations.
    If LoadWordArray(abytData()) Then

        ' Concatenate the state array into one data
        ' string and return the digital signature.
        HashString = FormatOutput()
    Else
        gblnStopProcessing = True
    End If

HashString_CleanUp:
    ResetVariables    ' Reset class variables

End Function


' ***************************************************************************
' ****                  Internal functions and Procedures                ****
' ***************************************************************************

' ***************************************************************************
' Routine:       FormatOutput
'
' Description:   Concatenate the hashed values into one string.
'
'                Output length in hex format:
'                    RipeMD-128 = 32 characters
'                    RipeMD-256 = 64 characters
'
' Returns:       Hex data in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Feb-2006  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified variable names and documented.  Converted data to
'              lowercase prior to leaving routine.
' ***************************************************************************
Private Function FormatOutput() As Byte()

    Dim strHex     As String
    Dim intIndex   As Integer
    Dim intCount   As Integer
    Dim intPointer As Integer
    Dim lngTemp    As Long

    strHex = Space$(128)   ' preload output buffer
    intPointer = 1         ' Start position in output buffer

    ' Concatenate the Hash values into one string
    Select Case mlngHashMethod
           Case eRIPEMD128: intCount = 3  ' 4 elements
           Case eRIPEMD256: intCount = 7  ' 8 elements
    End Select

    ' Decode 4 x 32-bit words into 16 bytes with
    ' LSB (Least Significant Bit) first each time
    With mobjMath32
        For intIndex = 0 To intCount

            ' Convert data back to Big_Endian format
            lngTemp = .SwapEndianLong(malngHash(intIndex))
            Mid$(strHex, intPointer, 8) = .LongToHex(lngTemp)
            intPointer = intPointer + 8

        Next intIndex
    End With

    If mblnRetLowercase Then
        strHex = LCase$(Trim$(strHex))   ' Convert string to lowercase
    Else
        strHex = UCase$(Trim$(strHex))   ' Convert string to uppercase
    End If

    FormatOutput = StringToByteArray(strHex)   ' convert string data to byte array

    RaiseEvent HashProgress(100)   ' Update progress bar

End Function

' ***************************************************************************
' Routine:       LoadWordArray
'
' Description:   All data is stored in Big_Endian format with the Most 
'                Significant Bit (MSB) first. Message data will be
'                converted to Little_Endian format, with the Least
'                Significant Bit (MSB) first, in the Transform()
'                routine prior to being manipulated.
' 
'                The message 'M' shall be padded before hash computation
'                begins. The purpose of this padding is to ensure that the
'                padded message is a multiple of 512 bits.
'
'         32-Bit Format
'                Suppose the length of the message 'M', in bits, is 'l'
'                bits. Append the bit '1' to the end of the message,
'                followed by 'k' zero bits, where 'k' is the smallest
'                non-negative solution to the equation l+1+k=448 mod 512.
'                Then append the 64-bit block that is equal to the number
'                'l' expressed using a binary representation. The length of
'                the padded message should now be a multiple of 512 bits.
'
' Parameters:    abytData() - message to be hashed in a byte array
'
' Returns:       Message data formatted in an array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Feb-2006  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Renamed routine, variables, and modified
' ***************************************************************************
Private Function LoadWordArray(ByRef abytData() As Byte) As Boolean

    Dim lngIndex       As Long
    Dim lngPadding     As Long
    Dim lngMaxBytes    As Long
    Dim lngPosition    As Long
    Dim lngSpaceAlloc  As Long
    Dim lngMsgLenInHex As Long
    Dim strMsg         As String
    Dim strHex         As String
    Dim strTemp        As String
    Dim abytMsg()      As Byte

    On Error GoTo LoadWordArray_Error

    LoadWordArray = True  ' Preset flag to TRUE

    With mobjMath32
        lngSpaceAlloc = 512
        lngMaxBytes = 64
        lngMsgLenInHex = Len(mstrBitsInHex)
        strMsg = .ByteArrayToString(abytData())                   ' convert to string data
        mcurCurrBitCnt = mcurCurrBitCnt + CCur(Len(strMsg) * 8)  ' calc the ongoing bit count

        ' Incoming message is in equal multiples
        If Len(strMsg) Mod lngMaxBytes = 0 Then

            ' Process message in equal lengths. There
            ' is no need for any padding at this time.
            Do
                Erase abytMsg()            ' Always start with an empty array
                strTemp = vbNullString     ' Empty string variables
                strHex = vbNullString

                strTemp = Mid$(strMsg, 1, lngMaxBytes)   ' Capture chunk of data message
                strMsg = Mid$(strMsg, lngMaxBytes + 1)   ' Resize data message to what was not used

                ' See if we have some data to work with
                If Len(strTemp) > 0 Then

                    ' the string of data is the exact length needed
                    If Len(strTemp) = lngMaxBytes Then

                        strHex = Space$(lngSpaceAlloc)
                        lngPosition = 1

                        ' convert data to hex string
                        For lngIndex = 1 To Len(strTemp)
                            Mid$(strHex, lngPosition, 2) = Right$("00" & Hex$(Asc(Mid$(strTemp, lngIndex, 1))), 2)
                            lngPosition = lngPosition + 2
                        Next lngIndex

                        abytMsg() = .HexStringToByteArray(strHex)  ' Convert to byte array
                        Transform abytMsg()                        ' Hash data

                    End If
                Else
                    Exit Do  ' exit Do..Loop
                End If

            Loop While Len(strMsg) > 0
        End If

        strTemp = vbNullString     ' Empty string variables
        strHex = vbNullString

        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            GoTo LoadWordArray_CleanUp
        End If

        ' if the length of the message is not
        ' in equal multiples
        If Len(strMsg) > 0 Then

            If Len(strMsg) >= lngMaxBytes Then

                Do
                    Erase abytMsg()            ' Always start with an empty array
                    strTemp = vbNullString     ' Empty string variables
                    strHex = vbNullString

                    strTemp = Mid$(strMsg, 1, lngMaxBytes)   ' Capture chunk of data message
                    strMsg = Mid$(strMsg, lngMaxBytes + 1)   ' Resize data message to what was not used

                    If Len(strTemp) = lngMaxBytes Then

                        strHex = Space$(lngSpaceAlloc)
                        lngPosition = 1

                        ' convert data to hex string
                        For lngIndex = 1 To Len(strTemp)
                            Mid$(strHex, lngPosition, 2) = Right$("00" & Hex$(Asc(Mid$(strTemp, lngIndex, 1))), 2)
                            lngPosition = lngPosition + 2
                        Next lngIndex

                        abytMsg() = .HexStringToByteArray(strHex)  ' Convert to byte array
                        Transform abytMsg()                        ' Hash data

                    Else
                        strHex = vbNullString
                        Exit Do  ' exit Do..Loop
                    End If

                Loop

                ' An error occurred or user opted to STOP processing
                If gblnStopProcessing Then
                    GoTo LoadWordArray_CleanUp
                End If

            Else
                strTemp = Mid$(strMsg, 1, lngMaxBytes)   ' Capture chunk of data message
                strMsg = Mid$(strMsg, lngMaxBytes + 1)   ' Resize data message to what was not used
            End If

            ' the string of data is less than the length needed
            If (Len(strTemp) > 0) And (Len(strMsg) = 0) Then

                strTemp = strTemp & Chr$(128)
                strHex = Space$(lngSpaceAlloc)
                lngPosition = 1

                ' convert data to hex string
                For lngIndex = 1 To Len(strTemp)
                    Mid$(strHex, lngPosition, 2) = Right$("00" & Hex$(Asc(Mid$(strTemp, lngIndex, 1))), 2)
                    lngPosition = lngPosition + 2
                Next lngIndex

                strHex = Trim$(strHex)  ' remove all leading and trailing blanks
                strTemp = vbNullString

                ' if the len of the hex string plus the element length is less
                ' than or equal to the required record length then append the
                ' appropriate number of zeroes and the hex representation of
                ' the message length
                If (Len(strHex) < MSG_LEN_32) Then

                    ' Calculate amount of padding needed
                    lngPadding = MSG_LEN_32 - Len(strHex)

                    Select Case lngPadding

                           Case lngMsgLenInHex
                                ' If the amount of padding needed is the
                                ' exact length of the bit string then
                                ' append, process and leave
                                strHex = strHex & mstrBitsInHex

                           Case Is < lngMsgLenInHex
                                ' *** String 1 ***
                                ' Hex string is exact length needed.
                                ' Will need to create an extra string.
                                strHex = strHex & String$(lngPadding, "0")
                                abytMsg() = .HexStringToByteArray(strHex)  ' Convert to byte array
                                Transform abytMsg()                        ' Hash data

                                ' *** String 2 ***
                                ' Create a string of zeroes and the
                                ' bit count of the message length.
                                strHex = vbNullString
                                lngPadding = MSG_LEN_32 - lngMsgLenInHex
                                strHex = String$(lngPadding, "0") & mstrBitsInHex

                           Case Is > lngMsgLenInHex
                                ' Adjust amount of padding to fill in
                                ' the gap between data and number of bits
                                lngPadding = MSG_LEN_32 - (Len(strHex) + lngMsgLenInHex)
                                strHex = strHex & String$(lngPadding, "0") & mstrBitsInHex
                    End Select

                    abytMsg() = .HexStringToByteArray(strHex)  ' Convert to byte array
                    Transform abytMsg()                        ' Hash data
                    GoTo LoadWordArray_CleanUp

                Else
                    ' *** String 1 ***
                    ' Hex string is exact length needed.
                    ' Will need to create an extra string.
                    abytMsg() = .HexStringToByteArray(strHex)  ' Convert to byte array
                    Transform abytMsg()                        ' Hash data

                    ' *** String 2 ***
                    ' Create a string of zeroes and the
                    ' bit count of the message length.
                    strHex = vbNullString
                    lngPadding = MSG_LEN_32 - lngMsgLenInHex
                    strHex = String$(lngPadding, "0") & mstrBitsInHex

                    abytMsg() = .HexStringToByteArray(strHex)  ' Convert to byte array
                    Transform abytMsg()                        ' Hash data
                    GoTo LoadWordArray_CleanUp

                End If
            End If
        End If

        ' No data found in current work string
        If Len(strMsg) = 0 Then

            ' See if we have reached the end
            ' of the message file or string
            If mcurCurrBitCnt = mcurBitCount Then

                ' End of message indicator is the starting character followed by a
                ' string of zeroes and the hex representation of the message length
                lngPadding = MSG_LEN_32 - (lngMsgLenInHex + 2)
                strHex = Hex$(128) & String$(lngPadding, "0") & mstrBitsInHex
                abytMsg() = .HexStringToByteArray(strHex)  ' Convert to byte array
                Transform abytMsg()                        ' Hash data
            End If
        End If
    End With

LoadWordArray_CleanUp:
    Erase abytMsg()         ' Always empty arrays when not needed
    strMsg = vbNullString   ' Verify string variables are empty
    strTemp = vbNullString
    strHex = vbNullString

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        LoadWordArray = False
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

LoadWordArray_Error:
    ErrorMsg MODULE_NAME, "LoadWordArray", Err.Description
    gblnStopProcessing = True
    Resume LoadWordArray_CleanUp

End Function

' ***************************************************************************
' Routine:       Transform
'
' Description:   Convert the byte array data into long integers to be placed
'                in the four elements of the state() array.
'
' Parameters:    abytMsg() - Message chunk in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated algorithms to more closely resemble the RipeMD-128,
'              C code and the RipeMD-256 pseudocode
'              located at:
'              http://homes.esat.kuleuven.be/~bosselae/ripemd160.html
'              Test results for RipeMD-128 are accurate.  However, RipeMD-256
'              test results do not match the results on the web page because
'              all I had to reference was pseudocode.
' 15-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Thanks again to Blazer for his assistance and persistance
'              in correcting the results for RipeMD-256 and RipeMD-320.  The
'              hashed results are now accurate in accordance with
'              http://homes.esat.kuleuven.be/~bosselae/ripemd160.html
'              - In its current form, RipeMD-256 is a weak hash algorithm.
'              You should use SHA-256 if you want security.
'              http://www.users.zetnet.co.uk/hopwood/crypto/scan/md.html
'              - RIPEMD-256 is considered insecure, and should not be used
'              unless you absolutely need it for compatibility.
'              http://www.cryptopp.com/docs/ref/class_r_i_p_e_m_d256.html
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added RaiseEvents to track hash progress
' ***************************************************************************
Private Sub Transform(ByRef abytMsg() As Byte)

' Updates 4 alngWord 32-bit values in malngHash()
' Input: the next 64 bytes in abytMsg starting at offset lngIndex
' Assumes at least 64 bytes are present after offset lngIndex

    Dim w1           As Byte   ' Word index
    Dim w2           As Byte

    Dim s1           As Byte   ' Shift value
    Dim s2           As Byte

    Dim a1           As Long
    Dim b1           As Long
    Dim c1           As Long
    Dim d1           As Long
    Dim a2           As Long
    Dim b2           As Long
    Dim c2           As Long
    Dim d2           As Long
    Dim lngIdx       As Long
    Dim lngRound     As Long
    Dim lngIndex     As Long
    Dim lngProgress  As Long
    Dim alngWord()   As Long

    ReDim alngWord(15)  ' Size temp work array
    lngIndex = 0

    a1 = malngHash(0)
    b1 = malngHash(1)
    c1 = malngHash(2)
    d1 = malngHash(3)

    Select Case mlngHashMethod

           Case eRIPEMD128
                a2 = malngHash(0)
                b2 = malngHash(1)
                c2 = malngHash(2)
                d2 = malngHash(3)

           Case eRIPEMD256
                a2 = malngHash(4)
                b2 = malngHash(5)
                c2 = malngHash(6)
                d2 = malngHash(7)
    End Select

    ' ******** Message preparation **************************
    ' Convert data to Little-Endian format
    For lngIdx = 0 To 15

        alngWord(lngIdx) = UnsignedJoin(abytMsg(lngIndex), _
                                        abytMsg(lngIndex + 1), _
                                        abytMsg(lngIndex + 2), _
                                        abytMsg(lngIndex + 3))
        lngIndex = lngIndex + 4

    Next lngIdx
    ' ******** End of message preparation ********************

    ' Perform required number of iterations
    For lngRound = 1 To mlngHashRounds

        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

        For lngIndex = 0 To 63

            w1 = mabytW1(lngIndex)
            w2 = mabytW2(lngIndex)
            s1 = mabytS1(lngIndex)
            s2 = mabytS2(lngIndex)

            Select Case lngIndex

                   Case 0 To 15
                        Select Case lngIndex Mod 4
                               Case 0
                                    FF a1, b1, c1, d1, alngWord(w1), s1
                                    III a2, b2, c2, d2, alngWord(w2), s2
                               Case 1
                                    FF d1, a1, b1, c1, alngWord(w1), s1
                                    III d2, a2, b2, c2, alngWord(w2), s2
                               Case 2
                                    FF c1, d1, a1, b1, alngWord(w1), s1
                                    III c2, d2, a2, b2, alngWord(w2), s2
                               Case 3
                                    FF b1, c1, d1, a1, alngWord(w1), s1
                                    III b2, c2, d2, a2, alngWord(w2), s2
                        End Select

                    Case 16 To 31
                        Select Case lngIndex Mod 4
                               Case 0
                                    GG a1, b1, c1, d1, alngWord(w1), s1
                                    HHH a2, b2, c2, d2, alngWord(w2), s2
                               Case 1
                                    GG d1, a1, b1, c1, alngWord(w1), s1
                                    HHH d2, a2, b2, c2, alngWord(w2), s2
                               Case 2
                                    GG c1, d1, a1, b1, alngWord(w1), s1
                                    HHH c2, d2, a2, b2, alngWord(w2), s2
                               Case 3
                                    GG b1, c1, d1, a1, alngWord(w1), s1
                                    HHH b2, c2, d2, a2, alngWord(w2), s2
                        End Select

                    Case 32 To 47
                        Select Case lngIndex Mod 4
                               Case 0
                                    HH a1, b1, c1, d1, alngWord(w1), s1
                                    GGG a2, b2, c2, d2, alngWord(w2), s2
                               Case 1
                                    HH d1, a1, b1, c1, alngWord(w1), s1
                                    GGG d2, a2, b2, c2, alngWord(w2), s2
                               Case 2
                                    HH c1, d1, a1, b1, alngWord(w1), s1
                                    GGG c2, d2, a2, b2, alngWord(w2), s2
                               Case 3
                                    HH b1, c1, d1, a1, alngWord(w1), s1
                                    GGG b2, c2, d2, a2, alngWord(w2), s2
                        End Select

                    Case Else
                        Select Case lngIndex Mod 4
                               Case 0
                                    II a1, b1, c1, d1, alngWord(w1), s1
                                    FFF a2, b2, c2, d2, alngWord(w2), s2
                               Case 1
                                    II d1, a1, b1, c1, alngWord(w1), s1
                                    FFF d2, a2, b2, c2, alngWord(w2), s2
                               Case 2
                                    II c1, d1, a1, b1, alngWord(w1), s1
                                    FFF c2, d2, a2, b2, alngWord(w2), s2
                               Case 3
                                    II b1, c1, d1, a1, alngWord(w1), s1
                                    FFF b2, c2, d2, a2, alngWord(w2), s2
                        End Select
            End Select

            '************************************
            ' Tweak 1 - Swap with counterpart
            '************************************
            If mlngHashMethod = eRIPEMD256 Then
                Select Case lngIndex
                       Case 15: SwapLong a1, a2
                       Case 31: SwapLong b1, b2
                       Case 47: SwapLong c1, c2
                       Case 63: SwapLong d1, d2
                End Select
            End If

        Next lngIndex

        ' 29-Jan-2010 Update progress bar. Separate
        '             steps for easier debugging
        mcurAccrued = mcurAccrued + MSG_LEN_32
        lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
        RaiseEvent HashProgress(lngProgress)

    Next lngRound

    ' Accumulate the values
    With mobjMath32
        Select Case mlngHashMethod

               Case eRIPEMD128
                    d2 = .UnsignedAdd(.UnsignedAdd(malngHash(1), c1), d2)
                    malngHash(1) = .UnsignedAdd(.UnsignedAdd(malngHash(2), d1), a2)
                    malngHash(2) = .UnsignedAdd(.UnsignedAdd(malngHash(3), a1), b2)
                    malngHash(3) = .UnsignedAdd(.UnsignedAdd(malngHash(0), b1), c2)
                    malngHash(0) = d2

               '************************************
               ' Tweak 2 - Sequential add
               '************************************
               Case eRIPEMD256
                    malngHash(0) = .UnsignedAdd(malngHash(0), a1)
                    malngHash(1) = .UnsignedAdd(malngHash(1), b1)
                    malngHash(2) = .UnsignedAdd(malngHash(2), c1)
                    malngHash(3) = .UnsignedAdd(malngHash(3), d1)
                    malngHash(4) = .UnsignedAdd(malngHash(4), a2)
                    malngHash(5) = .UnsignedAdd(malngHash(5), b2)
                    malngHash(6) = .UnsignedAdd(malngHash(6), c2)
                    malngHash(7) = .UnsignedAdd(malngHash(7), d2)
        End Select
    End With

End Sub

Private Function UnsignedJoin(ByVal a As Byte, _
                              ByVal b As Byte, _
                              ByVal c As Byte, _
                              ByVal d As Byte) As Long
    Dim abytData(4) As Byte

    abytData(0) = a
    abytData(1) = b
    abytData(2) = c
    abytData(3) = d

    ' Join 4 x 8-bit bytes into one 32-bit word
    With mobjMath32
        UnsignedJoin = .SwapEndianLong(.ByteArrayToLong(abytData()))
    End With
    
End Function

' ***************************************************************************
' Routine:       SwapLong
'
' Description:   Swap data with each other without using a temp.  Using Math
'                to do this takes a substantial amount more processing than
'                Logic Gates. The logic gates foundation is in all processors.
'                Which you could argue that so is Math, but running a math
'                process with a temp variable in fact uses a considerable
'                amount more processor cycles than logic gates.
'
'                Using 3 Xor's over 3 Basic Equations greatly speeds up the
'                application. (especially when dealing with millions of
'                equations/Logic Gates)
'
'                This routine works with byte, integer and long values.
'                Change the parameter data type accordingly.
'
' Parameters:    AA - data to be swapped with Value2
'                BB - data to be swapped with Value1
'
' Returns:       Swapped data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Mar-2003  Ulli
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=43903&lngWId=1
' ***************************************************************************
Private Sub SwapLong(ByRef AA As Long, _
                     ByRef BB As Long)

    AA = AA Xor BB
    BB = BB Xor AA
    AA = AA Xor BB

End Sub

' ***************************************************************************
' Routines:      Bit manipulating routines for 128 and 256 bit hash strings
'
' Description:   Manipulates the numeric bits
'
' Returns:       Only the value of "a" is changd.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routines
' 15-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Thanks again to Blazer for his assistance and persistance
'              in correcting the results for RipeMD-256 and RipeMD-320.  The
'              hashed results are now accurate in accordance with
'              http://homes.esat.kuleuven.be/~bosselae/ripemd160.html
'              - In its current form, RipeMD-256 is a weak hash algorithm.
'              You should use SHA-256 if you want security.
'              http://www.users.zetnet.co.uk/hopwood/crypto/scan/md.html
'              RIPEMD-256 is considered insecure, and should not be used
'              unless you absolutely need it for compatibility.
'              http://www.cryptopp.com/docs/ref/class_r_i_p_e_m_d256.html
' ***************************************************************************
Private Sub FF(ByRef a As Long, _
               ByVal b As Long, _
               ByVal c As Long, _
               ByVal d As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += F((b) ^ (c) ^ (d)) + (lngWord);\
        ' (a) = ROL((a), (s));\
        lngTmp = (b Xor c Xor d)
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
    End With

End Sub

Private Sub GG(ByRef a As Long, _
               ByVal b As Long, _
               ByVal c As Long, _
               ByVal d As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += G(((b) & (c)) | (~(b) & (d))) + (x) + 0x5a827999UL;\
        ' (a) = ROL((a), (s));\
        lngTmp = ((b And c) Or ((Not b) And d))
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        lngTmp = .UnsignedAdd(lngTmp, &H5A827999)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
    End With

End Sub

Private Sub HH(ByRef a As Long, _
               ByVal b As Long, _
               ByVal c As Long, _
               ByVal d As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += H(((b) | ~(c)) ^ (d)) + (x) + 0x6ed9eba1UL;\
        ' (a) = ROL((a), (s));\
        lngTmp = ((b Or (Not c)) Xor d)
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        lngTmp = .UnsignedAdd(lngTmp, &H6ED9EBA1)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
    End With

End Sub

Private Sub II(ByRef a As Long, _
               ByVal b As Long, _
               ByVal c As Long, _
               ByVal d As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += I(((b) & (d)) | ((c) & ~(d))) + (x) + 0x8f1bbcdcUL;\
        ' (a) = ROL((a), (s));\
        lngTmp = (b And d) Or (c And (Not d))
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        lngTmp = .UnsignedAdd(lngTmp, &H8F1BBCDC)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
    End With

End Sub

Private Sub FFF(ByRef a As Long, _
                ByVal b As Long, _
                ByVal c As Long, _
                ByVal d As Long, _
                ByVal lngWord As Long, _
                ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += F((b) ^ (c) ^ (d)) + (x);\
        ' (a) = ROL((a), (s));\
        lngTmp = (b Xor c Xor d)
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
    End With

End Sub

Private Sub GGG(ByRef a As Long, _
                ByVal b As Long, _
                ByVal c As Long, _
                ByVal d As Long, _
                ByVal lngWord As Long, _
                ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += G(((b) & (c)) | (~(b) & (d))) + (x) + 0x6d703ef3UL;\
        ' (a) = ROL((a), (s));\
        lngTmp = ((b And c) Or ((Not b) And d))
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        lngTmp = .UnsignedAdd(lngTmp, &H6D703EF3)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
    End With

End Sub

Private Sub HHH(ByRef a As Long, _
                ByVal b As Long, _
                ByVal c As Long, _
                ByVal d As Long, _
                ByVal lngWord As Long, _
                ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += H(((b) | ~(c)) ^ (d)) + (x) + 0x5c4dd124UL;\
        ' (a) = ROL((a), (s));\
        lngTmp = ((b Or (Not c)) Xor d)
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        lngTmp = .UnsignedAdd(lngTmp, &H5C4DD124)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
    End With

End Sub

Private Sub III(ByRef a As Long, _
                ByVal b As Long, _
                ByVal c As Long, _
                ByVal d As Long, _
                ByVal lngWord As Long, _
                ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += I(((b) & (d)) | ((c) & ~(d))) + (x) + 0x50a28be6UL;\
        ' (a) = ROL((a), (s));\
        lngTmp = ((b And d) Or (c And (Not d)))
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        lngTmp = .UnsignedAdd(lngTmp, &H50A28BE6)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
    End With

End Sub

' **************************************************************************
' Routine:       GetBlockSize
'
' Description:   Determines the size of the data to be processed.
'
' Parameters:    curAmtLeft - Amount of data left
'
' Returns:       New record size as a long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine 
' ***************************************************************************
Private Function GetBlockSize(ByVal curAmtLeft As Currency) As Long

    ' Determine amount of data to process
    Select Case curAmtLeft
           Case Is >= KB_32: GetBlockSize = KB_32
           Case Else:        GetBlockSize = CLng(curAmtLeft)
    End Select

End Function

' ***************************************************************************
' Routine:       Initialize
'
' Description:   Preloads the arrays and constants with predetermined values
'                to be used in the hashing calculations.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub Initialize()

    Dim strData  As String
    Dim intIdx   As Integer
    Dim avntData As Variant

    Select Case mlngHashMethod

           Case eRIPEMD128
                ReDim malngHash(4)
                malngHash(0) = &H67452301
                malngHash(1) = &HEFCDAB89
                malngHash(2) = &H98BADCFE
                malngHash(3) = &H10325476

           Case eRIPEMD256
                ReDim malngHash(8)
                malngHash(0) = &H67452301
                malngHash(1) = &HEFCDAB89
                malngHash(2) = &H98BADCFE
                malngHash(3) = &H10325476
                malngHash(4) = &H76543210
                malngHash(5) = &HFEDCBA98
                malngHash(6) = &H89ABCDEF
                malngHash(7) = &H1234567
    End Select

    ' Load Bit Shift array values
    ReDim mabytS1(64)
    ReDim mabytS2(64)

    strData = vbNullString
    avntData = Empty   ' Start with an empty variant

    strData = strData & "11 14 15 12 5 8 7 9 11 13 14 15 6 7 9 8 "
    strData = strData & "7 6 8 13 11 9 7 15 7 12 15 9 11 7 13 12 "
    strData = strData & "11 13 6 7 14 9 13 15 14 8 13 6 5 12 7 5 "
    strData = strData & "11 12 14 15 14 15 9 8 9 14 5 6 8 6 5 12 "

    avntData = Split(strData, Chr$(32))

    For intIdx = 0 To 63
        mabytS1(intIdx) = CByte(avntData(intIdx))
    Next intIdx

    strData = vbNullString
    avntData = Empty   ' Start with an empty variant

    strData = strData & "8 9 9 11 13 15 15 5 7 7 8 11 14 14 12 6 "
    strData = strData & "9 13 15 7 12 8 9 11 7 7 12 7 6 15 13 11 "
    strData = strData & "9 7 15 11 8 6 6 14 12 13 5 14 13 13 7 5 "
    strData = strData & "15 5 8 11 14 14 6 14 6 9 12 9 12 5 15 8 "

    avntData = Split(strData, Chr$(32))

    For intIdx = 0 To 63
        mabytS2(intIdx) = CByte(avntData(intIdx))
    Next intIdx

    strData = vbNullString
    avntData = Empty   ' Start with an empty variant

    ' Load Word array indices
    ReDim mabytW1(64)
    ReDim mabytW2(64)

    strData = strData & "0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 "
    strData = strData & "7 4 13 1 10 6 15 3 12 0 9 5 2 14 11 8 "
    strData = strData & "3 10 14 4 9 15 8 1 2 7 0 6 13 11 5 12 "
    strData = strData & "1 9 11 10 0 8 12 4 13 3 7 15 14 5 6 2 "

    avntData = Split(strData, Chr$(32))

    For intIdx = 0 To 63
        mabytW1(intIdx) = CByte(avntData(intIdx))
    Next intIdx

    strData = vbNullString
    avntData = Empty   ' Start with an empty variant

    strData = strData & "5 14 7 0 9 2 11 4 13 6 15 8 1 10 3 12 "
    strData = strData & "6 11 3 7 0 13 5 10 14 15 8 12 4 9 1 2 "
    strData = strData & "15 5 1 3 7 14 6 9 11 8 12 2 10 0 4 13 "
    strData = strData & "8 6 4 1 3 11 15 0 5 12 2 13 9 7 10 14 "

    avntData = Split(strData, Chr$(32))

    For intIdx = 0 To 63
        mabytW2(intIdx) = CByte(avntData(intIdx))
    Next intIdx

    strData = vbNullString
    avntData = Empty   ' Always empty variants when not needed

End Sub

Private Sub ResetVariables()

    Erase malngHash()   ' Start with empty arrays
    Erase mabytS1()
    Erase mabytS2()
    Erase mabytW1()
    Erase mabytW2()

    Erase malngHash()
    mcurAccrued = 0@
    mcurBitCount = 0@
    mcurCurrBitCnt = 0@
    mstrBitsInHex = vbNullString

End Sub

Private Sub Class_Initialize()

    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objRipeMD128 = new cRipeMD128

    ResetVariables
    Set mobjMath32 = Nothing      ' Free class object from memory
    Set mobjMath32 = New cMath32  ' Instantiate class object

    ' Preset property values
    ReturnLowercase = False       ' Default return of hashed data
    HashRounds = 1                ' Set default number of passes
    HashMethod = eRIPEMD128       ' Set default hash algorithm

End Sub

Private Sub Class_Terminate()

    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objRipeMD128 = Nothing

    ResetVariables
    Set mobjMath32 = Nothing

End Sub



