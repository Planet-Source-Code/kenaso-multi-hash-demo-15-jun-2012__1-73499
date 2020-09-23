VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMD4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Project:       MD4 (Message-Digest Algorithm)
'
' Module:        clsMD4
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
'    MD4 Message-Digest Algorithm has been compromised
'
'    At the rump session of Crypto 2004 it was announced that Xiaoyun
'    Wang, Dengguo Feng, Xuejia Lai and Hongbo Yu found collisions for
'    MD4, MD5, RIPEMD, and the 128-bit version of HAVAL.
'    http://eprint.iacr.org/2004/199.pdf
'
' ***************************************************************************
'
' Description:   This software is provided ``AS IS'' and any express or
'                implied warranties, including, but not limited to, the
'                implied warranties of merchantability and fitness for a
'                particular purpose, are disclaimed.  In no event shall the
'                authors be liable for any direct, indirect, incidental,
'                special, exemplary, or consequential damages (including, but
'                not limited to, procurement of substitute goods or services;
'                loss of use, data, or profits; or business interruption)
'                however caused and on any theory of liability, whether in
'                contract, strict liability, or tort (including negligence
'                or otherwise) arising in any way out of the use of this
'                software, even if advised of the possibility of such damage.
'
'                This is a class which encapsulates a set of MD4 Message
'                Digest functions.  MD4 algorithm produces a 128 bit digital
'                fingerprint (signature) from an dataset of arbitrary length.
'                For details see RFC 1321 (summarized below).  This
'                implementation is derived from the RSA Data Security, Inc.
'                MD4 Message-Digest algorithm reference implementation
'                (originally written in C).
'
'  NOTES:
'               Network Working Group                                       R. Rivest
'               Request for Comments: 1320        MIT Laboratory for Computer Science
'               Obsoletes: RFC 1186                       and RSA Data Security, Inc.
'                                                                          April 1992
'
'
'                                   The MD4 Message-Digest Algorithm
'
'               Status of thie Memo
'
'               This memo provides information for the Internet community.  It does
'               not specify an Internet standard.  Distribution of this memo is
'               unlimited.
'
'
'               Summary
'
'               This document describes the MD4 message-digest algorithm [1]. The
'               algorithm takes as input a message of arbitrary length and produces
'               as output a 128-bit "fingerprint" or "message digest" of the input.
'               It is conjectured that it is computationally infeasible to produce
'               two messages having the same message digest, or to produce any
'               message having a given prespecified target message digest. The MD4
'               algorithm is intended for digital signature applications, where a
'               large file must be "compressed" in a secure manner before being
'               encrypted with a private (secret) key under a public-key cryptosystem
'               such as RSA.
'
'               The MD4 algorithm is designed to be quite fast on 32-bit machines. In
'               addition, the MD4 algorithm does not require any large substitution
'               tables; the algorithm can be coded quite compactly.
'
'
'       RFC Author:
'       Ronald L.Rivest
'       Massachusetts Institute of Technology
'       Laboratory for Computer Science
'       NE43 -324545    Technology Square
'       Cambridge, MA  02139-1986
'
' ***************************************************************************
' Note:   Most test vectors obtained from RFC1320.TXT
'
'  MD4 test results using one (1) pass
'      "a"                          - bde52cb31de33e46245e05fbdbd6fb24
'      "abc"                        - a448017aaf21d8525fc10ae87aa6729d
'      "message digest"             - d9130a8164549fe818874806e1c7014b
'      "abcdefghijklmnopqrstuvwxyz" - d79e1c308aa5bbcdeea8ed63df412da9
'      "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
'                                   - 043f8582f241db351ce627e153e7f0e4
'      8 times "1234567890"         - e33b4ddc9c38f2199c3e7b164fcc0536
'      1 million letter "a"         - bbce80cc6bb65e5c6745e30d4eeca9a4
'      1 million binary zeroes      - d0b30f1d5bd243c0880eab13f4c9c643
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 16-Jan-2010  ßlazer  blazerx@gmail.com
'              Original
' 07-Mar-2010  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented.  Added results to flowerbox that
'              include one milion leter "a" and one million zeroes.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "clsMD4"
  Private Const MAX_ROUNDS  As Long = 10
  Private Const MSG_LEN_32  As Long = 128          ' 128 bytes long
  Private Const KB_32       As Long = &H8000&      ' 32768

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
  Private mstrBitsInHex    As String
  Private mblnRetLowercase As Boolean
  Private mcurMaxSize      As Currency
  Private mcurAccrued      As Currency
  Private mcurBitCount     As Currency
  Private mcurCurrBitCnt   As Currency
  Private mlngHashRounds   As Long
  Private malngHash()      As Long
  Private mobjMath32       As cMath32


' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************

' Update progress bar
Public Event HashProgress(ByVal lngProgress As Long)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

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
'    With mobjMD4
'        abytData() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'        mabytHashed() = .HashFile(abytData())                        ' hash data and return as Byte array
'        txtData.Text = StrConv(mabytHashed(), vbUnicode)             ' convert byte array to string data
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
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
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
    curFilePos = 0@                  ' set to first position in the file
    Erase abytData()                 ' make sure array is empty

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    With objBigFiles
        strSource = ByteArrayToString(abytFile())                ' Convert byte array to string
        .CalcFileSize strSource, curAmtLeft, mstrBitsInHex       ' Get the size of the file

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
    Set objBigFiles = Nothing        ' Always free objects from memory

    Erase abytData()                 ' Always empty arrays when not needed
    ResetVariables                   ' Reset class variables

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
'    With mobjMD4
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

    mcurAccrued = 0@                                                    ' Init accrued progress
    mcurMaxSize = CCur((UBound(abytData) + 1) * mlngHashRounds)         ' Capture data size
    mcurBitCount = CCur((UBound(abytData) + 1) * 8@)                    ' Get bit count for for comparison
    mstrBitsInHex = Right$(String$(16, "0") & Hex$(mcurBitCount), 16)   ' Convert bit count into hex
    mstrBitsInHex = mobjMath32.SwapEndianHex(mstrBitsInHex)             ' Reverse hex values in string

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
    ResetVariables   ' Reset class variables

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
'                    MD4 = 32 characters
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
    Dim intPointer As Integer
    Dim lngTemp    As Long

    strHex = Space$(64)   ' preload output buffer
    intPointer = 1        ' Start position in output buffer

    ' Decode 4 x 32-bit words into 16 bytes with
    ' LSB (Least Significant Bit) first each time
    With mobjMath32
        For intIndex = 0 To 3

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

    FormatOutput = StringToByteArray(strHex)  ' convert string data to byte array

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
' Returns:       TRUE - process completed successfully.
'                FALSE - user reqested to stop processing or an error
'                      occurred.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
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
        Erase abytMsg()       ' Always start with an empty array
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
                Erase abytMsg()          ' Always start with an empty array
                strTemp = vbNullString   ' Empty string variables
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
                    Erase abytMsg()          ' Always start with an empty array
                    strTemp = vbNullString   ' Empty string variables
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
                abytMsg() = .HexStringToByteArray(strHex)   ' Convert to byte array
                Transform abytMsg()                         ' Hash data
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
' Description:   Manipulate the data
'
' Parameters:    abytMsg() - Message chunk in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 07-Mar-2010  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Sub Transform(ByRef abytMsg() As Byte)

    Dim H0          As Long
    Dim H1          As Long
    Dim H2          As Long
    Dim H3          As Long
    Dim lngIdx      As Long
    Dim lngIndex    As Long
    Dim lngRound    As Long
    Dim lngProgress As Long
    Dim alngWord()  As Long

    ReDim alngWord(16) ' Size temp work array
    lngIndex = 0

    H0 = malngHash(0)  ' Initialize temp hash variables
    H1 = malngHash(1)
    H2 = malngHash(2)
    H3 = malngHash(3)

    ' ******** Message preparation **************************
    ' convert data to Little-Endian format
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

        ' Round 1
        FF H0, H1, H2, H3, alngWord(0), 3      ' 1
        FF H3, H0, H1, H2, alngWord(1), 7      ' 2
        FF H2, H3, H0, H1, alngWord(2), 11     ' 3
        FF H1, H2, H3, H0, alngWord(3), 19     ' 4
        FF H0, H1, H2, H3, alngWord(4), 3      ' 5
        FF H3, H0, H1, H2, alngWord(5), 7      ' 6
        FF H2, H3, H0, H1, alngWord(6), 11     ' 7
        FF H1, H2, H3, H0, alngWord(7), 19     ' 8
        FF H0, H1, H2, H3, alngWord(8), 3      ' 9
        FF H3, H0, H1, H2, alngWord(9), 7      ' 10
        FF H2, H3, H0, H1, alngWord(10), 11    ' 11
        FF H1, H2, H3, H0, alngWord(11), 19    ' 12
        FF H0, H1, H2, H3, alngWord(12), 3     ' 13
        FF H3, H0, H1, H2, alngWord(13), 7     ' 14
        FF H2, H3, H0, H1, alngWord(14), 11    ' 15
        FF H1, H2, H3, H0, alngWord(15), 19    ' 16

        ' Round 2
        GG H0, H1, H2, H3, alngWord(0), 3      ' 17
        GG H3, H0, H1, H2, alngWord(4), 5      ' 18
        GG H2, H3, H0, H1, alngWord(8), 9      ' 19
        GG H1, H2, H3, H0, alngWord(12), 13    ' 20
        GG H0, H1, H2, H3, alngWord(1), 3      ' 21
        GG H3, H0, H1, H2, alngWord(5), 5      ' 22
        GG H2, H3, H0, H1, alngWord(9), 9      ' 23
        GG H1, H2, H3, H0, alngWord(13), 13    ' 24
        GG H0, H1, H2, H3, alngWord(2), 3      ' 25
        GG H3, H0, H1, H2, alngWord(6), 5      ' 26
        GG H2, H3, H0, H1, alngWord(10), 9     ' 27
        GG H1, H2, H3, H0, alngWord(14), 13    ' 28
        GG H0, H1, H2, H3, alngWord(3), 3      ' 29
        GG H3, H0, H1, H2, alngWord(7), 5      ' 30
        GG H2, H3, H0, H1, alngWord(11), 9     ' 31
        GG H1, H2, H3, H0, alngWord(15), 13    ' 32

        ' Round 3
        HH H0, H1, H2, H3, alngWord(0), 3      ' 33
        HH H3, H0, H1, H2, alngWord(8), 9      ' 34
        HH H2, H3, H0, H1, alngWord(4), 11     ' 35
        HH H1, H2, H3, H0, alngWord(12), 15    ' 36
        HH H0, H1, H2, H3, alngWord(2), 3      ' 37
        HH H3, H0, H1, H2, alngWord(10), 9     ' 38
        HH H2, H3, H0, H1, alngWord(6), 11     ' 39
        HH H1, H2, H3, H0, alngWord(14), 15    ' 40
        HH H0, H1, H2, H3, alngWord(1), 3      ' 41
        HH H3, H0, H1, H2, alngWord(9), 9      ' 42
        HH H2, H3, H0, H1, alngWord(5), 11     ' 43
        HH H1, H2, H3, H0, alngWord(13), 15    ' 44
        HH H0, H1, H2, H3, alngWord(3), 3      ' 45
        HH H3, H0, H1, H2, alngWord(11), 9     ' 46
        HH H2, H3, H0, H1, alngWord(7), 11     ' 47
        HH H1, H2, H3, H0, alngWord(15), 15    ' 48

        ' 29-Jan-2010 Update progress bar. Separate
        '             steps for easier debugging
        mcurAccrued = mcurAccrued + MSG_LEN_32
        lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
        RaiseEvent HashProgress(lngProgress)

    Next lngRound

    ' Add manipulated values to their original values
    With mobjMath32
        malngHash(0) = .UnsignedAdd(malngHash(0), H0)
        malngHash(1) = .UnsignedAdd(malngHash(1), H1)
        malngHash(2) = .UnsignedAdd(malngHash(2), H2)
        malngHash(3) = .UnsignedAdd(malngHash(3), H3)
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
' 07-Mar-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routines
' ***************************************************************************
Private Sub FF(ByRef a As Long, _
               ByVal b As Long, _
               ByVal c As Long, _
               ByVal d As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer)

    Dim lngTemp As Long

    With mobjMath32
        lngTemp = (b And c) Or ((Not b) And d)
        lngTemp = .UnsignedAdd(lngTemp, lngWord)
        a = .UnsignedAdd(a, lngTemp)
        a = .w32Rotate(a, intBitShift)
    End With

End Sub

Private Sub GG(ByRef a As Long, _
               ByVal b As Long, _
               ByVal c As Long, _
               ByVal d As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer)

    Dim lngTemp As Long

    With mobjMath32
        lngTemp = (b And c) Or (b And d) Or (c And d)
        lngTemp = .UnsignedAdd(lngTemp, lngWord)
        lngTemp = .UnsignedAdd(lngTemp, &H5A827999)
        a = .UnsignedAdd(a, lngTemp)
        a = .w32Rotate(a, intBitShift)
    End With

End Sub

Private Sub HH(ByRef a As Long, _
               ByVal b As Long, _
               ByVal c As Long, _
               ByVal d As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer)

    Dim lngTemp As Long

    With mobjMath32
        lngTemp = (b Xor c Xor d)
        lngTemp = .UnsignedAdd(lngTemp, lngWord)
        lngTemp = .UnsignedAdd(lngTemp, &H6ED9EBA1)
        a = .UnsignedAdd(a, lngTemp)
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

Private Sub Initialize()

    ReDim malngHash(4)

    ' Load initialization values
    malngHash(0) = &H67452301
    malngHash(1) = &HEFCDAB89
    malngHash(2) = &H98BADCFE
    malngHash(3) = &H10325476

End Sub

Private Sub ResetVariables()

    Erase malngHash()
    mcurAccrued = 0@
    mcurBitCount = 0@
    mcurCurrBitCnt = 0@
    mstrBitsInHex = vbNullString

End Sub

Private Sub Class_Initialize()

    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objMD4 = new cMD4

    ResetVariables
    Set mobjMath32 = Nothing      ' Free class object from memory
    Set mobjMath32 = New cMath32  ' Instantiate class object

    ' Preset property values
    ReturnLowercase = False       ' Default return of hashed data
    HashRounds = 1                ' Set default number of passes

End Sub

Private Sub Class_Terminate()

    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objMD4 = Nothing

    ResetVariables
    Set mobjMath32 = Nothing

End Sub



