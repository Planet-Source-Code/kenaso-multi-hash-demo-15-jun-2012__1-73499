VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRipeMD160"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Project:       RipeMD160
'
' Module:        clsRipeMD160_320.cls
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
'          RIPEMD-160   9c1185a5c5e9fc54612808977ee8f548b2258d31
'          RIPEMD-320   22d65d5661536cdc75c1fdf5c6de7b41b9f27325ebc61e8557177d705a0ec880151c3a32a00899b8
'
'      "a"
'          RIPEMD-160   0bdc9d2d256b3ee9daae347be6f4dc835a467ffe
'          RIPEMD-320   ce78850638f92658a5a585097579926dda667a5716562cfcf6fbe77f63542f99b04705d6970dff5d
'
'      "abc"
'          RIPEMD-160   8eb208f7e05d987a9b044a8e98c6b087f15a0bfc
'          RIPEMD-320   de4c01b3054f8930a79d09ae738e92301e5a17085beffdc1b8d116713e74f82fa942d64cdbc4682d
'
'      "message digest"
'          RIPEMD-160   5d0689ef49d2fae572b881b123a85ffa21595f36
'          RIPEMD-320   3a8e28502ed45d422f68844f9dd316e7b98533fa3f2a91d29f84d425c88d6b4eff727df66a7c0197
'
'      "abcdefghijklmnopqrstuvwxyz"
'          RIPEMD-160   f71c27109c692c1b56bbdceb5b9d2865b3708dbc
'          RIPEMD-320   cabdb1810b92470a2093aa6bce05952c28348cf43ff60841975166bb40ed234004b8824463e6b009
'
'      "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"
'          RIPEMD-160   12a053384a9c0c88e405a06c27dcf49ada62eb2b
'          RIPEMD-320   d034a7950cf722021ba4b84df769a5de2060e259df4c9bb4a4268c0e935bbc7470a969c9d072a1ac
'
'      "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
'          RIPEMD-160   b0e20b6e3116640286ed3a87a5713079b21f5189
'          RIPEMD-320   ed544940c86d67f250d232c30b7b3e5770e0c60c8cb9a4cafe3b11388af9920e1b99230b843c86a4
'
'      8 times "1234567890"
'          RIPEMD-160   9b752e45573d4b39f4dbd3323cab82bf63326bfb
'          RIPEMD-320   557888af5f6d8ed62ab66945c6d2a0a47ecd5341e915eb8fea1d0524955f825dc717e4a008ab2d42
'
'      1 million letter "a"
'          RIPEMD-160   52783243c1697bdbe16d37f97f68f08325dc1528
'          RIPEMD-320   bdee37f4371e20646b8b0d862dda16292ae36f40965e8c8509e63d1dbddecc503e2b63eb9245bb66
'
'      1 million binary zeroes
'          RIPEMD-160   0b1413246ae36e0877af8d833c3797fade0f91d9
'          RIPEMD-320   32f19606d9028e500faceb062c01326954e30565bf8f01b32dc6e260e86ad948cc57b3f3d5e1ad7a
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
'                RipeMD-320 test results do not match the results on the
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
'                ** In its current form, RipeMD-320 is a weak hash
'                   algorithm. You should use SHA-256, SHA-512 or Whirlpool
'                   if you want security.
'                   http://www.users.zetnet.co.uk/hopwood/crypto/scan/md.html
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
  Private Const MODULE_NAME As String = "clsRipeMD160_320"
  Private Const MAX_ROUNDS  As Long = 10
  Private Const MSG_LEN_32  As Long = 128          ' 128 bytes long
  Private Const KB_32       As Long = &H8000&      ' 32768

' ***************************************************************************
' Enumerations
' ***************************************************************************
  Private Enum enumRIPEMD_ALGORITHM
      eRIPEMD160  ' 0
      eRIPEMD320  ' 1
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
           Case Else:   mlngHashMethod = eRIPEMD160   ' Use default value
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
'                    RipeMD-160 = 40 characters
'                    RipeMD-320 = 80 characters
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
           Case eRIPEMD160: intCount = 4  ' 5 elements
           Case eRIPEMD320: intCount = 9  ' 10 elements
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
'              Updated algorithms to more closely resemble the RipeMD-160,
'              C code and the RipeMD-320 pseudocode
'              located at:
'              http://homes.esat.kuleuven.be/~bosselae/ripemd160.html
'              Test results for RipeMD-160 are accurate.  However, RipeMD-320
'              test results do not match the results on the web page because
'              all I had to reference was pseudocode.
' 08-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Corrected a logic bug in loading of hash values.
' 15-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Thanks again to Blazer for his assistance and persistance
'              in correcting the results for RipeMD-256 and RipeMD-320.  The
'              hashed results are now accurate in accordance with
'              http://homes.esat.kuleuven.be/~bosselae/ripemd160.html
'              - In its current form, RipeMD-320 is a weak hash algorithm.
'              You should use SHA-256, SHA-512 or Whirlpool if you
'              want security.
'              http://www.users.zetnet.co.uk/hopwood/crypto/scan/md.html
'              RIPEMD-256 and RIPEMD-320 provide no better security than
'              RIPEMD-160. Therefore, they are redundant, and can only serve
'              to mislead users into thinking that they provide additional
'              security.
'              http://lists.dlitz.net/pipermail/pycrypto/2008q4/000039.html
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
    Dim e1           As Long
    Dim a2           As Long
    Dim b2           As Long
    Dim c2           As Long
    Dim d2           As Long
    Dim e2           As Long
    Dim lngIdx       As Long
    Dim lngIndex     As Long
    Dim lngRound     As Long
    Dim lngProgress  As Long
    Dim alngWord()   As Long

    ReDim alngWord(15)  ' Size temp work array
    lngIndex = 0

    a1 = malngHash(0)
    b1 = malngHash(1)
    c1 = malngHash(2)
    d1 = malngHash(3)
    e1 = malngHash(4)

    Select Case mlngHashMethod

           Case eRIPEMD160
                a2 = malngHash(0)
                b2 = malngHash(1)
                c2 = malngHash(2)
                d2 = malngHash(3)
                e2 = malngHash(4)

           Case eRIPEMD320  ' 08-Jan-2010 Corrected from 160 -> 320
                a2 = malngHash(5)
                b2 = malngHash(6)
                c2 = malngHash(7)
                d2 = malngHash(8)
                e2 = malngHash(9)
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

         For lngIndex = 0 To 79

            w1 = mabytW1(lngIndex)
            w2 = mabytW2(lngIndex)
            s1 = mabytS1(lngIndex)
            s2 = mabytS2(lngIndex)

            Select Case lngIndex

                   Case 0 To 15
                        Select Case lngIndex Mod 5
                               Case 0
                                    FF a1, b1, c1, d1, e1, alngWord(w1), s1
                                    JJJ a2, b2, c2, d2, e2, alngWord(w2), s2
                               Case 1
                                    FF e1, a1, b1, c1, d1, alngWord(w1), s1
                                    JJJ e2, a2, b2, c2, d2, alngWord(w2), s2
                               Case 2
                                    FF d1, e1, a1, b1, c1, alngWord(w1), s1
                                    JJJ d2, e2, a2, b2, c2, alngWord(w2), s2
                               Case 3
                                    FF c1, d1, e1, a1, b1, alngWord(w1), s1
                                    JJJ c2, d2, e2, a2, b2, alngWord(w2), s2
                               Case 4
                                    FF b1, c1, d1, e1, a1, alngWord(w1), s1
                                    JJJ b2, c2, d2, e2, a2, alngWord(w2), s2
                        End Select

                    Case 16 To 31
                        Select Case lngIndex Mod 5
                               Case 0
                                    GG a1, b1, c1, d1, e1, alngWord(w1), s1
                                    III a2, b2, c2, d2, e2, alngWord(w2), s2
                               Case 1
                                    GG e1, a1, b1, c1, d1, alngWord(w1), s1
                                    III e2, a2, b2, c2, d2, alngWord(w2), s2
                               Case 2
                                    GG d1, e1, a1, b1, c1, alngWord(w1), s1
                                    III d2, e2, a2, b2, c2, alngWord(w2), s2
                               Case 3
                                    GG c1, d1, e1, a1, b1, alngWord(w1), s1
                                    III c2, d2, e2, a2, b2, alngWord(w2), s2
                               Case 4
                                    GG b1, c1, d1, e1, a1, alngWord(w1), s1
                                    III b2, c2, d2, e2, a2, alngWord(w2), s2
                        End Select

                    Case 32 To 47
                        Select Case lngIndex Mod 5
                               Case 0
                                    HH a1, b1, c1, d1, e1, alngWord(w1), s1
                                    HHH a2, b2, c2, d2, e2, alngWord(w2), s2
                               Case 1
                                    HH e1, a1, b1, c1, d1, alngWord(w1), s1
                                    HHH e2, a2, b2, c2, d2, alngWord(w2), s2
                               Case 2
                                    HH d1, e1, a1, b1, c1, alngWord(w1), s1
                                    HHH d2, e2, a2, b2, c2, alngWord(w2), s2
                               Case 3
                                    HH c1, d1, e1, a1, b1, alngWord(w1), s1
                                    HHH c2, d2, e2, a2, b2, alngWord(w2), s2
                               Case 4
                                    HH b1, c1, d1, e1, a1, alngWord(w1), s1
                                    HHH b2, c2, d2, e2, a2, alngWord(w2), s2
                        End Select

                    Case 48 To 63
                        Select Case lngIndex Mod 5
                               Case 0
                                    II a1, b1, c1, d1, e1, alngWord(w1), s1
                                    GGG a2, b2, c2, d2, e2, alngWord(w2), s2
                               Case 1
                                    II e1, a1, b1, c1, d1, alngWord(w1), s1
                                    GGG e2, a2, b2, c2, d2, alngWord(w2), s2
                               Case 2
                                    II d1, e1, a1, b1, c1, alngWord(w1), s1
                                    GGG d2, e2, a2, b2, c2, alngWord(w2), s2
                               Case 3
                                    II c1, d1, e1, a1, b1, alngWord(w1), s1
                                    GGG c2, d2, e2, a2, b2, alngWord(w2), s2
                               Case 4
                                    II b1, c1, d1, e1, a1, alngWord(w1), s1
                                    GGG b2, c2, d2, e2, a2, alngWord(w2), s2
                        End Select

                    Case Else
                        Select Case lngIndex Mod 5
                               Case 0
                                    JJ a1, b1, c1, d1, e1, alngWord(w1), s1
                                    FFF a2, b2, c2, d2, e2, alngWord(w2), s2
                               Case 1
                                    JJ e1, a1, b1, c1, d1, alngWord(w1), s1
                                    FFF e2, a2, b2, c2, d2, alngWord(w2), s2
                               Case 2
                                    JJ d1, e1, a1, b1, c1, alngWord(w1), s1
                                    FFF d2, e2, a2, b2, c2, alngWord(w2), s2
                               Case 3
                                    JJ c1, d1, e1, a1, b1, alngWord(w1), s1
                                    FFF c2, d2, e2, a2, b2, alngWord(w2), s2
                               Case 4
                                    JJ b1, c1, d1, e1, a1, alngWord(w1), s1
                                    FFF b2, c2, d2, e2, a2, alngWord(w2), s2
                        End Select
            End Select

            '**********************************
            ' Tweak 1 - Swap with counterpart
            '**********************************
            If mlngHashMethod = eRIPEMD320 Then
                Select Case lngIndex
                       Case 15: SwapLong a1, a2
                       Case 31: SwapLong b1, b2
                       Case 47: SwapLong c1, c2
                       Case 63: SwapLong d1, d2
                       Case 79: SwapLong e1, e2
                End Select
            End If

        Next lngIndex

        ' 29-Jan-2010 Update progress bar. Separate
        '             steps for easier debugging
        mcurAccrued = mcurAccrued + MSG_LEN_32
        lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
        RaiseEvent HashProgress(lngProgress)

    Next lngRound

    ' Accumulate hashed values
    With mobjMath32
        Select Case mlngHashMethod

               Case eRIPEMD160
                    d2 = .UnsignedAdd(.UnsignedAdd(malngHash(1), c1), d2)
                    malngHash(1) = .UnsignedAdd(.UnsignedAdd(malngHash(2), d1), e2)
                    malngHash(2) = .UnsignedAdd(.UnsignedAdd(malngHash(3), e1), a2)
                    malngHash(3) = .UnsignedAdd(.UnsignedAdd(malngHash(4), a1), b2)
                    malngHash(4) = .UnsignedAdd(.UnsignedAdd(malngHash(0), b1), c2)
                    malngHash(0) = d2

               '************************************
               ' Tweak 2 - Sequential add
               '************************************
               Case eRIPEMD320
                    malngHash(0) = .UnsignedAdd(malngHash(0), a1)
                    malngHash(1) = .UnsignedAdd(malngHash(1), b1)
                    malngHash(2) = .UnsignedAdd(malngHash(2), c1)
                    malngHash(3) = .UnsignedAdd(malngHash(3), d1)
                    malngHash(4) = .UnsignedAdd(malngHash(4), e1)
                    malngHash(5) = .UnsignedAdd(malngHash(5), a2)
                    malngHash(6) = .UnsignedAdd(malngHash(6), b2)
                    malngHash(7) = .UnsignedAdd(malngHash(7), c2)
                    malngHash(8) = .UnsignedAdd(malngHash(8), d2)
                    malngHash(9) = .UnsignedAdd(malngHash(9), e2)
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
' Routines:      Bit manipulating routines for 160 and 320 bit hash strings
'
' Description:   Manipulates the numeric bits
'
' Returns:       The value of "a" and "c" are changed.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routines
' 15-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Thanks again to Blazer for his assistance and persistance
'                in correcting the results for RipeMD-256 and RipeMD-320.
'                The hashed results are now accurate in accordance with
'                http://homes.esat.kuleuven.be/~bosselae/ripemd160.html
'              - In its current form, RipeMD-320 is a weak hash algorithm.
'                You should use SHA-256, SHA-512 or Whirlpool if you
'                want security.
'                http://www.users.zetnet.co.uk/hopwood/crypto/scan/md.html
'              = RIPEMD-256 and RIPEMD-320 provide no better security than
'                RIPEMD-160. Therefore, they are redundant, and can only
'                serve to mislead users into thinking that they provide
'                additional security.
'                http://lists.dlitz.net/pipermail/pycrypto/2008q4/000039.html
' ***************************************************************************
Private Sub FF(ByRef a As Long, _
               ByVal b As Long, _
               ByRef c As Long, _
               ByVal d As Long, _
               ByVal e As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += F((b) ^ (c) ^ (d)) + (x);\
        ' (a) = ROL((a), (s)) + (e);\
        ' (c) = ROL((c), 10);\
        lngTmp = (b Xor c Xor d)
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
        a = .UnsignedAdd(a, e)
        c = .w32Rotate(c, 10)
    End With

End Sub

Private Sub GG(ByRef a As Long, _
               ByVal b As Long, _
               ByRef c As Long, _
               ByVal d As Long, _
               ByVal e As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += F(((b) & (c)) | (~(b) & (d))) + (x) + 0x5A827999;\
        ' (a) = ROL((a), (s)) + (e);\
        ' (c) = ROL((c), 10);\
        lngTmp = ((b And c) Or ((Not b) And d))
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        lngTmp = .UnsignedAdd(lngTmp, &H5A827999)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
        a = .UnsignedAdd(a, e)
        c = .w32Rotate(c, 10)
    End With

End Sub

Private Sub HH(ByRef a As Long, _
               ByVal b As Long, _
               ByRef c As Long, _
               ByVal d As Long, _
               ByVal e As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += H(((b) | (~(c))) ^ (d))) + (x) + 0x6ed9eba1UL;\
        ' (a) = ROL((a), (s)) + (e);\
        ' (c) = ROL((c), 10);\
        lngTmp = ((b Or (Not c)) Xor d)
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        lngTmp = .UnsignedAdd(lngTmp, &H6ED9EBA1)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
        a = .UnsignedAdd(a, e)
        c = .w32Rotate(c, 10)
    End With

End Sub

Private Sub II(ByRef a As Long, _
               ByVal b As Long, _
               ByRef c As Long, _
               ByVal d As Long, _
               ByVal e As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += F((b AND d) OR (c AND NOT(d)) + (x) + 0x8F1BBCDC;\
        ' (a) = ROL((a), (s)) + (e);\
        ' (c) = ROL((c), 10);\
        lngTmp = ((b And d) Or (c And (Not d)))
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        lngTmp = .UnsignedAdd(lngTmp, &H8F1BBCDC)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
        a = .UnsignedAdd(a, e)
        c = .w32Rotate(c, 10)
    End With

End Sub

Private Sub JJ(ByRef a As Long, _
               ByVal b As Long, _
               ByRef c As Long, _
               ByVal d As Long, _
               ByVal e As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += F(((b) & (d)) | ((c) & (~(d))) + (x) + 0xA953FD4E;\
        ' (a) = ROL((a), (s)) + (e);\
        ' (c) = ROL((c), 10);\
        lngTmp = (b Xor (c Or (Not d)))
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        lngTmp = .UnsignedAdd(lngTmp, &HA953FD4E)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
        a = .UnsignedAdd(a, e)
        c = .w32Rotate(c, 10)
    End With

End Sub

Private Sub FFF(ByRef a As Long, _
                ByVal b As Long, _
                ByRef c As Long, _
                ByVal d As Long, _
                ByVal e As Long, _
                ByVal lngWord As Long, _
                ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += F((b) ^ (c) ^ (d)) + (x);\
        ' (a) = ROL((a), (s)) + (e);\
        ' (c) = ROL((c), 10);\
        lngTmp = (b Xor c Xor d)
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
        a = .UnsignedAdd(a, e)
        c = .w32Rotate(c, 10)
    End With

End Sub

Private Sub GGG(ByRef a As Long, _
                ByVal b As Long, _
                ByRef c As Long, _
                ByVal d As Long, _
                ByVal e As Long, _
                ByVal lngWord As Long, _
                ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += F(((b) & (c)) | (~(b) & (d))) + (x) + 0x7A6D76E9;\
        ' (a) = ROL((a), (s)) + (e);\
        ' (c) = ROL((c), 10);\
        lngTmp = (b And c) Or ((Not b) And d)
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        lngTmp = .UnsignedAdd(lngTmp, &H7A6D76E9)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
        a = .UnsignedAdd(a, e)
        c = .w32Rotate(c, 10)
    End With

End Sub

Private Sub HHH(ByRef a As Long, _
                ByVal b As Long, _
                ByRef c As Long, _
                ByVal d As Long, _
                ByVal e As Long, _
                ByVal lngWord As Long, _
                ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += H(((b) | (~(c))) ^ (d))) + (x) + 0x6D703EF3UL;\
        ' (a) = ROL((a), (s)) + (e);\
        ' (c) = ROL((c), 10);\
        lngTmp = ((b Or (Not c)) Xor d)
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        lngTmp = .UnsignedAdd(lngTmp, &H6D703EF3)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
        a = .UnsignedAdd(a, e)
        c = .w32Rotate(c, 10)
    End With

End Sub

Private Sub III(ByRef a As Long, _
                ByVal b As Long, _
                ByRef c As Long, _
                ByVal d As Long, _
                ByVal e As Long, _
                ByVal lngWord As Long, _
                ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += F(((b) & (d)) | ((c) & (~(d))) + (x) + 0x5C4DD124;\
        ' (a) = ROL((a), (s)) + (e);\
        ' (c) = ROL((c), 10);\
        lngTmp = ((b And d) Or (c And (Not d)))
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        lngTmp = .UnsignedAdd(lngTmp, &H5C4DD124)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
        a = .UnsignedAdd(a, e)
        c = .w32Rotate(c, 10)
    End With

End Sub

Private Sub JJJ(ByRef a As Long, _
                ByVal b As Long, _
                ByRef c As Long, _
                ByVal d As Long, _
                ByVal e As Long, _
                ByVal lngWord As Long, _
                ByVal intBitShift As Integer)

    Dim lngTmp As Long

    With mobjMath32
        ' (a) += F(b XOR (c OR NOT(d))) + (x) + 0x50A28BE6;\
        ' (a) = ROL((a), (s)) + (e);\
        ' (c) = ROL((c), 10);\
        lngTmp = (b Xor (c Or (Not d)))
        lngTmp = .UnsignedAdd(lngTmp, lngWord)
        lngTmp = .UnsignedAdd(lngTmp, &H50A28BE6)
        a = .UnsignedAdd(a, lngTmp)
        a = .w32Rotate(a, intBitShift)
        a = .UnsignedAdd(a, e)
        c = .w32Rotate(c, 10)
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

    Dim intIdx   As Integer
    Dim strData  As String
    Dim avntData As Variant

    Select Case mlngHashMethod

           Case eRIPEMD160
                ReDim malngHash(5)
                malngHash(0) = &H67452301
                malngHash(1) = &HEFCDAB89
                malngHash(2) = &H98BADCFE
                malngHash(3) = &H10325476
                malngHash(4) = &HC3D2E1F0

           Case eRIPEMD320
                ReDim malngHash(10)
                malngHash(0) = &H67452301
                malngHash(1) = &HEFCDAB89
                malngHash(2) = &H98BADCFE
                malngHash(3) = &H10325476
                malngHash(4) = &HC3D2E1F0
                malngHash(5) = &H76543210
                malngHash(6) = &HFEDCBA98
                malngHash(7) = &H89ABCDEF
                malngHash(8) = &H1234567
                malngHash(9) = &H3C2D1E0F
    End Select

    ' Load Bit Shift array values
    ReDim mabytS1(80)
    ReDim mabytS2(80)

    strData = vbNullString       ' Clear data string
    avntData = Empty   ' Empty variant

    strData = strData & "11 14 15 12 5 8 7 9 11 13 14 15 6 7 9 8 "
    strData = strData & "7 6 8 13 11 9 7 15 7 12 15 9 11 7 13 12 "
    strData = strData & "11 13 6 7 14 9 13 15 14 8 13 6 5 12 7 5 "
    strData = strData & "11 12 14 15 14 15 9 8 9 14 5 6 8 6 5 12 "
    strData = strData & "9 15 5 11 6 8 13 12 5 12 13 14 11 8 5 6 "

    avntData = Split(strData, Chr$(32))

    For intIdx = 0 To 79
        mabytS1(intIdx) = CByte(avntData(intIdx))
    Next intIdx

    strData = vbNullString      ' Clear data string
    avntData = Empty  ' Empty variant

    strData = strData & "8 9 9 11 13 15 15 5 7 7 8 11 14 14 12 6 "
    strData = strData & "9 13 15 7 12 8 9 11 7 7 12 7 6 15 13 11 "
    strData = strData & "9 7 15 11 8 6 6 14 12 13 5 14 13 13 7 5 "
    strData = strData & "15 5 8 11 14 14 6 14 6 9 12 9 12 5 15 8 "
    strData = strData & "8 5 12 9 12 5 14 6 8 13 6 5 15 13 11 11 "

    avntData = Split(strData, Chr$(32))

    For intIdx = 0 To 79
        mabytS2(intIdx) = CByte(avntData(intIdx))
    Next intIdx

    strData = vbNullString       ' Clear data string
    avntData = Empty   ' Empty variant

    ' Load Word array indices
    ReDim mabytW1(80)
    ReDim mabytW2(80)

    strData = strData & "0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 "
    strData = strData & "7 4 13 1 10 6 15 3 12 0 9 5 2 14 11 8 "
    strData = strData & "3 10 14 4 9 15 8 1 2 7 0 6 13 11 5 12 "
    strData = strData & "1 9 11 10 0 8 12 4 13 3 7 15 14 5 6 2 "
    strData = strData & "4 0 5 9 7 12 2 10 14 1 3 8 11 6 15 13 "

    avntData = Split(strData, Chr$(32))

    For intIdx = 0 To 79
        mabytW1(intIdx) = CByte(avntData(intIdx))
    Next intIdx

    strData = vbNullString      ' Clear data string
    avntData = Empty  ' Empty variant

    strData = strData & "5 14 7 0 9 2 11 4 13 6 15 8 1 10 3 12 "
    strData = strData & "6 11 3 7 0 13 5 10 14 15 8 12 4 9 1 2 "
    strData = strData & "15 5 1 3 7 14 6 9 11 8 12 2 10 0 4 13 "
    strData = strData & "8 6 4 1 3 11 15 0 5 12 2 13 9 7 10 14 "
    strData = strData & "12 15 10 4 1 5 8 7 6 2 13 14 0 3 9 11 "

    avntData = Split(strData, Chr$(32))

    For intIdx = 0 To 79
        mabytW2(intIdx) = CByte(avntData(intIdx))
    Next intIdx

    strData = vbNullString      ' Clear data string
    avntData = Empty  ' Always empty variants when not needed

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
    ' Ex:  Set objRipeMD160 = new cRipeMD160

    ResetVariables
    Set mobjMath32 = Nothing      ' Free class object from memory
    Set mobjMath32 = New cMath32  ' Instantiate class object

    ' Preset property values
    ReturnLowercase = False       ' Default return of hashed data
    HashRounds = 1                ' Set default number of passes
    HashMethod = eRIPEMD160       ' Set default hash algorithm

End Sub

Private Sub Class_Terminate()

    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objRipeMD160 = Nothing

    ResetVariables
    Set mobjMath32 = Nothing

End Sub



