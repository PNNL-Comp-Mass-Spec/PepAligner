Option Strict On

' This class reads a file containing peptides and aligns them to a Fasta file 
'  (or tab-delimited file) of protein sequences (using NAligner.dll, a C# DLL)
'
' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started December 1, 2007
' Copyright 2007, Battelle Memorial Institute.  All Rights Reserved.

' E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
' Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
' -------------------------------------------------------------------------------
' 
' Licensed under the Apache License, Version 2.0; you may not use this file except
' in compliance with the License.  You may obtain a copy of the License at 
' http://www.apache.org/licenses/LICENSE-2.0
'
' Notice: This computer software was prepared by Battelle Memorial Institute, 
' hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the 
' Department of Energy (DOE).  All rights in the computer software are reserved 
' by DOE on behalf of the United States Government and the Contractor as 
' provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY 
' WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS 
' SOFTWARE.  This notice including this sentence must appear on any copies of 
' this computer software.

Public Class clsAlignPeptides

    Public Sub New()
        InitializeLocalVariables()
    End Sub

#Region "Constants and Enums"
    Public Const DEFAULT_MATRIX_NAME As String = "BLOSUM62"
    Public Const DEFAULT_GAP_OPEN_PENALTY As Double = 10
    Public Const DEFAULT_GAP_EXTEND_PENALTY As Double = 4
    Public Const DEFAULT_ALIGNMENT_SCORE_THRESHOLD As Integer = 32

    Protected Const MAX_ALIGNMENT_SCORE_TO_TRACK As Integer = 1000
#End Region

#Region "Structures"
    Protected Structure udtProteinInfoType
        Public ProteinName As String
        Public ProteinSequence As String
    End Structure

    Protected Structure udtAlignmentStatsType
        Public PeptideResidueCountAligned As Integer
        Public ProteinResidueCountAligned As Integer
        Public PeptideIdentityCoverage As Single
        Public PeptideSimilarityCoverage As Single
        Public PeptideResidueCoverage As Single
        Public ProteinResidueCoverage As Single
        Public Score As Integer
        Public Identity As Integer
        Public Similarity As Integer
        Public Gaps As Integer
    End Structure

#End Region

#Region "Classwide Variables"

    Private mMatrixName As String
    Private mGapOpenPenalty As Single
    Private mGapExtendPenalty As Single

    Private mDelimitedFileFormatCode As ProteinFileReader.DelimitedFileReader.eDelimitedFileFormatCode
    Private mDisplayAlignedPeptides As Boolean
    Private mAlignmentScoreThreshold As Integer

    Private mAlignmentScoreHistogram() As Integer

    Private mProteinCount As Integer
    Private mProteinInfo() As udtProteinInfoType

#End Region

#Region "Interface Functions"

    Public Property AlignmentScoreThreshold() As Integer
        Get
            Return mAlignmentScoreThreshold
        End Get
        Set(ByVal Value As Integer)
            mAlignmentScoreThreshold = Value
        End Set
    End Property

    Public Property DelimitedFileFormatCode() As ProteinFileReader.DelimitedFileReader.eDelimitedFileFormatCode
        Get
            Return mDelimitedFileFormatCode
        End Get
        Set(ByVal Value As ProteinFileReader.DelimitedFileReader.eDelimitedFileFormatCode)
            mDelimitedFileFormatCode = Value
        End Set
    End Property

    Public Property DisplayAlignedPeptides() As Boolean
        Get
            Return mDisplayAlignedPeptides
        End Get
        Set(ByVal Value As Boolean)
            mDisplayAlignedPeptides = Value
        End Set
    End Property

    Public Property MatrixName() As String
        Get
            Return mMatrixName
        End Get
        Set(ByVal Value As String)
            If Not Value Is Nothing AndAlso Value.Length > 0 Then
                mMatrixName = Value
            End If
        End Set
    End Property

    Public Property GapExtendPenalty() As Single
        Get
            Return mGapExtendPenalty
        End Get
        Set(ByVal Value As Single)
            If Value >= 0 Then
                mGapExtendPenalty = Value
            Else
                mGapExtendPenalty = 0
            End If
        End Set
    End Property

    Public Property GapOpenPenalty() As Single
        Get
            Return mGapOpenPenalty
        End Get
        Set(ByVal Value As Single)
            If Value >= 0 Then
                mGapOpenPenalty = Value
            Else
                mGapOpenPenalty = 0
            End If
        End Set
    End Property

#End Region

    Public Function AlignPeptidesToFastaFile(ByVal strPeptideInputFilePath As String, ByVal strFastaFilePath As String) As Boolean
        Dim objProteinFileReader As ProteinFileReader.FastaFileReader

        objProteinFileReader = New ProteinFileReader.FastaFileReader

        ReportProgress("Opening delimited protein file: " & strFastaFilePath)
        Return AlignPeptidesToProteinFile(strPeptideInputFilePath, strFastaFilePath, CType(objProteinFileReader, ProteinFileReader.ProteinFileReaderBaseClass))

    End Function

    Public Function AlignPeptidesToDelimitedTextProteinFile(ByVal strPeptideInputFilePath As String, ByVal strProteinFilePath As String) As Boolean
        Dim objProteinFileReader As ProteinFileReader.DelimitedFileReader

        objProteinFileReader = New ProteinFileReader.DelimitedFileReader
        objProteinFileReader.DelimitedFileFormatCode = mDelimitedFileFormatCode

        ReportProgress("Opening delimited protein file: " & strProteinFilePath)
        Return AlignPeptidesToProteinFile(strPeptideInputFilePath, strProteinFilePath, CType(objProteinFileReader, ProteinFileReader.ProteinFileReaderBaseClass))

    End Function

    Protected Function AlignPeptidesToProteinFile(ByVal strPeptideInputFilePath As String, ByVal strProteinFilePath As String, ByRef objProteinFileReader As ProteinFileReader.ProteinFileReaderBaseClass) As Boolean
        Const PRELOAD_THRESHOLD_MB As Integer = 250

        Dim objFileInfo As System.IO.FileInfo

        Dim blnSuccess As Boolean
        Dim blnProteinsPreloaded As Boolean

        If Not System.IO.File.Exists(strPeptideInputFilePath) Then
            Console.WriteLine("File not found: " & strPeptideInputFilePath)
            Return False
        End If

        If Not System.IO.File.Exists(strProteinFilePath) Then
            Console.WriteLine("File not found: " & strProteinFilePath)
            Return False
        End If

        ' Possibly pre-load the protein info
        objFileInfo = New System.IO.FileInfo(strProteinFilePath)

        If objFileInfo.Length / 1024.0 / 1024.0 <= PRELOAD_THRESHOLD_MB Then
            blnSuccess = PreloadProteinInfo(objProteinFileReader, strProteinFilePath)

            If Not blnSuccess Then
                Return False
            End If
            blnProteinsPreloaded = True
        Else
            blnProteinsPreloaded = False
        End If

        blnSuccess = AlignPeptidesToProteinFile(strPeptideInputFilePath, strProteinFilePath, objProteinFileReader, blnProteinsPreloaded)

    End Function

    Public Function AlignPeptidesToProteinFile(ByVal strPeptideInputFilePath As String, ByVal strProteinFilePath As String) As Boolean
        ' Auto-determine the type of strProteinFilePath
        If System.IO.Path.GetExtension(strProteinFilePath).ToLower = ".fasta" Then
            Return AlignPeptidesToFastaFile(strPeptideInputFilePath, strProteinFilePath)
        Else
            Return AlignPeptidesToDelimitedTextProteinFile(strPeptideInputFilePath, strProteinFilePath)
        End If

    End Function

    Protected Function AlignPeptidesToProteinFile(ByVal strPeptideInputFilePath As String, ByVal strProteinFilePath As String, ByRef objProteinFileReader As ProteinFileReader.ProteinFileReaderBaseClass, ByVal blnProteinsPreloaded As Boolean) As Boolean
        ' If blnProteinsPreloaded = False, then objProteinFileReader will be used to read the proteins from strProteinFilePath for each peptide in strPeptideInputFilePath
        ' If blnProteinsPreloaded = True, then the proteins have already been read into memory and threfore mProteinInfo will be used instead

        Const REPORTING_INTERVAL_THRESHOLD_SEC As Integer = 60

        Dim strOutputFileName As String = String.Empty
        Dim strHistogramOutputFileName As String

        Dim strLineIn As String
        Dim strSplitLine() As String
        Dim strPeptide As String
        Dim strRemainder As String

        Dim intIndex As Integer

        Dim intLinesRead As Integer
        Dim dtProcessingStartTime As System.DateTime
        Dim dtLastReportTime As System.DateTime

        Dim srInFile As System.IO.StreamReader
        Dim srOutFile As System.IO.StreamWriter

        Dim objMatrix As NAligner.Matrix
        Dim objPeptideSequence As NAligner.Sequence

        Dim blnSuccess As Boolean

        Try
            blnSuccess = False

            If Not System.IO.File.Exists(strPeptideInputFilePath) Then
                Console.WriteLine("File not found: " & strPeptideInputFilePath)
                Return False
            End If


            ReportProgress("Initializing alignment matrix: " & mMatrixName)
            If Not InitializeMatrix(objMatrix, mMatrixName) Then
                Console.WriteLine("Initialization failed")
                Return False
            End If

            ReportProgress("Reading " & strPeptideInputFilePath)
            Try
                srInFile = New System.IO.StreamReader(New System.IO.FileStream(strPeptideInputFilePath, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read))
            Catch ex As Exception
                Console.WriteLine("Error opening the input file (" & strPeptideInputFilePath & ") in AlignPeptidesToProteinFile: " & ex.Message)
                Return False
            End Try

            Try
                strHistogramOutputFileName = System.IO.Path.GetFileNameWithoutExtension(strPeptideInputFilePath) & "_output_summary.txt"

                strOutputFileName = System.IO.Path.GetFileNameWithoutExtension(strPeptideInputFilePath) & "_output.txt"
                srOutFile = New System.IO.StreamWriter(New System.IO.FileStream(strOutputFileName, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.Read))

                ' Write the header line
                srOutFile.WriteLine( _
                            "Peptide_Number" & ControlChars.Tab & _
                            "Peptide_Seq_Length" & ControlChars.Tab & _
                            "Protein_Seq_Length" & ControlChars.Tab & _
                            "Aligned_Length" & ControlChars.Tab & _
                            "Peptide_Identity_Coverage" & ControlChars.Tab & _
                            "Peptide_Similarity_Coverage" & ControlChars.Tab & _
                            "Peptide_Residue_Coverage" & ControlChars.Tab & _
                            "Protein_Residue_Coverage" & ControlChars.Tab & _
                            "Score" & ControlChars.Tab & _
                            "Identity" & ControlChars.Tab & _
                            "Similarity" & ControlChars.Tab & _
                            "Gaps" & ControlChars.Tab & _
                            "Peptide_Region" & ControlChars.Tab & _
                            "Alignment" & ControlChars.Tab & _
                            "Protein_Region" & ControlChars.Tab & _
                            "Protein_Name" _
                            )

            Catch ex As Exception
                Console.WriteLine("Error opening the output file (" & strOutputFileName & ") in AlignPeptidesToProteinFile: " & ex.Message)
                Return False
            End Try


            intLinesRead = 0
            dtLastReportTime = System.DateTime.UtcNow
            dtProcessingStartTime = dtLastReportTime

            ReDim mAlignmentScoreHistogram(MAX_ALIGNMENT_SCORE_TO_TRACK + 1)

            ' Read each entry in strPeptideInputFilePath and align against the proteins
            Do While srInFile.Peek >= 0
                strLineIn = srInFile.ReadLine
                intLinesRead += 1

                If Not strLineIn Is Nothing AndAlso strLineIn.Length > 0 Then
                    strLineIn = strLineIn.Trim

                    ' Split the line on tabs, spaces, or commas
                    ' Limiting to just two entries so that strSplitLine(0) will contain the peptide and strSplitLine(1) will contain everything else
                    strSplitLine = strLineIn.Split(New Char() {ControlChars.Tab, ","c, " "c}, 2)

                    If strSplitLine(0).Length > 0 Then
                        strPeptide = strSplitLine(0)
                        If strSplitLine.Length = 2 Then
                            strRemainder = strSplitLine(1)
                        ElseIf strSplitLine.Length > 2 Then
                            strRemainder = FlattenArray(strSplitLine, 1, strSplitLine.Length - 1)
                        Else
                            strRemainder = String.Empty
                        End If

                        objPeptideSequence = New NAligner.Sequence
                        objPeptideSequence.AAList = strPeptide
                        objPeptideSequence.Id = "Peptide_" & intLinesRead.ToString

                        If blnProteinsPreloaded Then
                            ' Align strSplitLine(0) against each of the proteins in memory
                            For intIndex = 0 To mProteinCount - 1
                                blnSuccess = AlignPeptideToProtein(srOutFile, intLinesRead, strPeptide, strRemainder, objPeptideSequence, objMatrix, mProteinInfo(intIndex).ProteinName, mProteinInfo(intIndex).ProteinSequence)
                                If Not blnSuccess Then Exit For
                            Next intIndex
                        Else
                            ' Align strSplitLine(0) against each of the proteins in strProteinFilePath
                            If Not objProteinFileReader.OpenFile(strProteinFilePath) Then
                                Console.WriteLine("Error opening the protein input file: " & strProteinFilePath)
                                Return False
                            End If

                            Do While objProteinFileReader.ReadNextProteinEntry
                                blnSuccess = AlignPeptideToProtein(srOutFile, intLinesRead, strPeptide, strRemainder, objPeptideSequence, objMatrix, objProteinFileReader.ProteinName(), objProteinFileReader.ProteinSequence())
                                If Not blnSuccess Then Exit Do
                            Loop

                            objProteinFileReader.CloseFile()
                        End If
                    End If

                End If

                If intLinesRead Mod 1000 = 0 OrElse System.DateTime.UtcNow.Subtract(dtLastReportTime).TotalSeconds >= REPORTING_INTERVAL_THRESHOLD_SEC Then
                    ReportProgress("Working: " & intLinesRead & " peptides processed")
                    dtLastReportTime = System.DateTime.UtcNow

                    srOutFile.Flush()
                    WriteHistogram(strHistogramOutputFileName, dtProcessingStartTime)
                End If

            Loop

            If intLinesRead > 0 Then
                ' Write out the alignment histogram
                WriteHistogram(strHistogramOutputFileName, dtProcessingStartTime)
            End If

            If Not srInFile Is Nothing Then srInFile.Close()
            If Not srOutFile Is Nothing Then srOutFile.Close()

            blnSuccess = True

        Catch ex As Exception
            Console.WriteLine("Error in AlignPeptidesToProteinFile: " & ex.Message)
            blnSuccess = False
        End Try

        ReportProgress("Done")

        Return blnSuccess

    End Function

    Private Function AlignPeptideToProtein(ByRef srOutFile As System.IO.StreamWriter, ByVal intLinesRead As Integer, ByRef strPeptide As String, ByRef strRemainder As String, ByRef objPeptideSequence As NAligner.Sequence, ByRef objMatrix As NAligner.Matrix, ByRef strProteinName As String, ByRef strProteinSequence As String) As Boolean

        Dim objProteinSequence As NAligner.Sequence
        Dim objAlignment As NAligner.Alignment

        Dim blnSuccess As Boolean

        Try
            objProteinSequence = NAligner.Sequence.Parse(">" & strProteinName & " " & strProteinName & ControlChars.NewLine & strProteinSequence)

            objAlignment = NAligner.SmithWatermanGotoh.Align(objPeptideSequence, objProteinSequence, objMatrix, mGapOpenPenalty, mGapExtendPenalty)
            blnSuccess = True

        Catch ex As Exception
            Console.WriteLine("Error aligning " & strPeptide & " to " & strProteinName & ": " & ex.Message)
            blnSuccess = False
        End Try

        If blnSuccess Then
            Try
                WriteAlignmentEntry(srOutFile, objAlignment, intLinesRead, strPeptide.Length, strProteinName, strProteinSequence.Length, strRemainder)
            Catch ex As Exception
                Console.WriteLine("Error writing to output file: " & ex.Message)
                blnSuccess = False
            End Try

        End If

        Return blnSuccess

    End Function

    Private Function CountLetters(ByVal strText As String) As Integer
        Dim intLetterCount As Integer
        Dim intIndex As Integer

        intLetterCount = 0
        For intIndex = 0 To strText.Length - 1
            If Char.IsUpper(strText.Chars(intIndex)) Then
                intLetterCount += 1
            End If
        Next intIndex

        Return intLetterCount

    End Function

    Private Function CreateDefaultMatrixFile() As String

        Dim swMatrix As System.IO.StreamWriter
        Dim strTempMatrixFilePath As String = String.Empty

        Try
            strTempMatrixFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath, DEFAULT_MATRIX_NAME)

            swMatrix = New System.IO.StreamWriter(strTempMatrixFilePath, False)

            swMatrix.WriteLine("#  Matrix made by matblas from blosum62.iij")
            swMatrix.WriteLine("#  * column uses minimum score")
            swMatrix.WriteLine("#  BLOSUM Clustered Scoring Matrix in 1/2 Bit Units")
            swMatrix.WriteLine("#  Blocks Database = /data/blocks_5.0/blocks.dat")
            swMatrix.WriteLine("#  Cluster Percentage: >= 62")
            swMatrix.WriteLine("#  Entropy =   0.6979, Expected =  -0.5209")
            swMatrix.WriteLine("   A  R  N  D  C  Q  E  G  H  I  L  K  M  F  P  S  T  W  Y  V  B  Z  X  *")
            swMatrix.WriteLine("A  4 -1 -2 -2  0 -1 -1  0 -2 -1 -1 -1 -1 -2 -1  1  0 -3 -2  0 -2 -1  0 -4 ")
            swMatrix.WriteLine("R -1  5  0 -2 -3  1  0 -2  0 -3 -2  2 -1 -3 -2 -1 -1 -3 -2 -3 -1  0 -1 -4 ")
            swMatrix.WriteLine("N -2  0  6  1 -3  0  0  0  1 -3 -3  0 -2 -3 -2  1  0 -4 -2 -3  3  0 -1 -4 ")
            swMatrix.WriteLine("D -2 -2  1  6 -3  0  2 -1 -1 -3 -4 -1 -3 -3 -1  0 -1 -4 -3 -3  4  1 -1 -4 ")
            swMatrix.WriteLine("C  0 -3 -3 -3  9 -3 -4 -3 -3 -1 -1 -3 -1 -2 -3 -1 -1 -2 -2 -1 -3 -3 -2 -4 ")
            swMatrix.WriteLine("Q -1  1  0  0 -3  5  2 -2  0 -3 -2  1  0 -3 -1  0 -1 -2 -1 -2  0  3 -1 -4 ")
            swMatrix.WriteLine("E -1  0  0  2 -4  2  5 -2  0 -3 -3  1 -2 -3 -1  0 -1 -3 -2 -2  1  4 -1 -4 ")
            swMatrix.WriteLine("G  0 -2  0 -1 -3 -2 -2  6 -2 -4 -4 -2 -3 -3 -2  0 -2 -2 -3 -3 -1 -2 -1 -4 ")
            swMatrix.WriteLine("H -2  0  1 -1 -3  0  0 -2  8 -3 -3 -1 -2 -1 -2 -1 -2 -2  2 -3  0  0 -1 -4 ")
            swMatrix.WriteLine("I -1 -3 -3 -3 -1 -3 -3 -4 -3  4  2 -3  1  0 -3 -2 -1 -3 -1  3 -3 -3 -1 -4 ")
            swMatrix.WriteLine("L -1 -2 -3 -4 -1 -2 -3 -4 -3  2  4 -2  2  0 -3 -2 -1 -2 -1  1 -4 -3 -1 -4 ")
            swMatrix.WriteLine("K -1  2  0 -1 -3  1  1 -2 -1 -3 -2  5 -1 -3 -1  0 -1 -3 -2 -2  0  1 -1 -4 ")
            swMatrix.WriteLine("M -1 -1 -2 -3 -1  0 -2 -3 -2  1  2 -1  5  0 -2 -1 -1 -1 -1  1 -3 -1 -1 -4 ")
            swMatrix.WriteLine("F -2 -3 -3 -3 -2 -3 -3 -3 -1  0  0 -3  0  6 -4 -2 -2  1  3 -1 -3 -3 -1 -4 ")
            swMatrix.WriteLine("P -1 -2 -2 -1 -3 -1 -1 -2 -2 -3 -3 -1 -2 -4  7 -1 -1 -4 -3 -2 -2 -1 -2 -4 ")
            swMatrix.WriteLine("S  1 -1  1  0 -1  0  0  0 -1 -2 -2  0 -1 -2 -1  4  1 -3 -2 -2  0  0  0 -4 ")
            swMatrix.WriteLine("T  0 -1  0 -1 -1 -1 -1 -2 -2 -1 -1 -1 -1 -2 -1  1  5 -2 -2  0 -1 -1  0 -4 ")
            swMatrix.WriteLine("W -3 -3 -4 -4 -2 -2 -3 -2 -2 -3 -2 -3 -1  1 -4 -3 -2 11  2 -3 -4 -3 -2 -4 ")
            swMatrix.WriteLine("Y -2 -2 -2 -3 -2 -1 -2 -3  2 -1 -1 -2 -1  3 -3 -2 -2  2  7 -1 -3 -2 -1 -4 ")
            swMatrix.WriteLine("V  0 -3 -3 -3 -1 -2 -2 -3 -3  3  1 -2  1 -1 -2 -2  0 -3 -1  4 -3 -2 -1 -4 ")
            swMatrix.WriteLine("B -2 -1  3  4 -3  0  1 -1  0 -3 -4  0 -3 -3 -2  0 -1 -4 -3 -3  4  1 -1 -4 ")
            swMatrix.WriteLine("Z -1  0  0  1 -3  3  4 -2  0 -3 -3  1 -1 -3 -1  0 -1 -3 -2 -2  1  4 -1 -4 ")
            swMatrix.WriteLine("X  0 -1 -1 -1 -2 -1 -1 -1 -1 -1 -1 -1 -1 -1 -2  0  0 -2 -1 -1 -1 -1 -1 -4 ")
            swMatrix.WriteLine("* -4 -4 -4 -4 -4 -4 -4 -4 -4 -4 -4 -4 -4 -4 -4 -4 -4 -4 -4 -4 -4 -4 -4  1 ")

            swMatrix.Close()

        Catch ex As Exception
            Console.WriteLine("Error creating temporary matrix file (" & strTempMatrixFilePath & "): " & ex.Message)
            strTempMatrixFilePath = String.Empty
        End Try

        Return strTempMatrixFilePath

    End Function

    Private Function FlattenArray(ByRef strArray() As String) As String
        Return FlattenArray(strArray, ControlChars.Tab)
    End Function

    Private Function FlattenArray(ByRef strArray() As String, ByVal chSepChar As Char) As String
        If strArray Is Nothing Then
            Return String.Empty
        Else
            Return FlattenArray(strArray, 0, strArray.Length, chSepChar)
        End If
    End Function

    Private Function FlattenArray(ByRef strArray() As String, ByVal intStartIndex As Integer, ByVal intDataCount As Integer) As String
        If strArray Is Nothing Then
            Return String.Empty
        Else
            Return FlattenArray(strArray, intStartIndex, intDataCount, ControlChars.Tab)
        End If
    End Function

    Private Function FlattenArray(ByRef strArray() As String, ByVal intStartIndex As Integer, ByVal intDataCount As Integer, ByVal chSepChar As Char) As String
        Dim intIndex As Integer
        Dim strResult As String

        If strArray Is Nothing Then
            Return String.Empty
        ElseIf strArray.Length = 0 OrElse intDataCount <= 0 Then
            Return String.Empty
        Else
            If intStartIndex < 0 Then
                intStartIndex = 0
            End If

            If intStartIndex >= strArray.Length Then
                Return String.Empty
            End If

            If intStartIndex + intDataCount > strArray.Length Then
                intDataCount = strArray.Length - intStartIndex
            End If

            strResult = strArray(intStartIndex)
            If strResult Is Nothing Then strResult = String.Empty

            For intIndex = intStartIndex + 1 To intStartIndex + intDataCount - 1
                If intIndex >= strArray.Length Then
                    ' Programming error; exit for
                    Exit For
                End If

                If strArray(intIndex) Is Nothing Then
                    strResult &= chSepChar
                Else
                    strResult &= chSepChar & strArray(intIndex)
                End If
            Next intIndex

            Return strResult
        End If
    End Function

    Public Shared Function GetDelimitedFileFormats() As String

        Return "  0: " & ProteinFileReader.DelimitedFileReader.eDelimitedFileFormatCode.SequenceOnly.ToString & ControlChars.NewLine & _
               "  1: " & ProteinFileReader.DelimitedFileReader.eDelimitedFileFormatCode.ProteinName_Sequence.ToString & ControlChars.NewLine & _
               "  2: " & ProteinFileReader.DelimitedFileReader.eDelimitedFileFormatCode.ProteinName_Description_Sequence.ToString & ControlChars.NewLine & _
               "  3: " & ProteinFileReader.DelimitedFileReader.eDelimitedFileFormatCode.UniqueID_Sequence.ToString

    End Function

    Private Function InitializeMatrix(ByRef objMatrix As NAligner.Matrix, ByVal strMatrixName As String) As Boolean

        Dim blnSuccess As Boolean
        Dim blnMatrixFileFound As Boolean

        Try
            If strMatrixName Is Nothing OrElse strMatrixName.Length = 0 Then
                strMatrixName = DEFAULT_MATRIX_NAME
            End If

            Try
                ' See if strMatrixName exists
                blnMatrixFileFound = False
                If System.IO.File.Exists(strMatrixName) Then
                    blnMatrixFileFound = True
                Else
                    ' Matrix file not found; default to use BLOSUM62
                    ' To do this, we must create a file in the temp folder, then instruct NAligner to load it
                    strMatrixName = CreateDefaultMatrixFile()

                    If System.IO.File.Exists(strMatrixName) Then
                        ' Unable to create the temporary file; cannot continue
                        blnMatrixFileFound = True
                    End If
                End If
            Catch ex As Exception
                blnMatrixFileFound = False
            End Try

            If blnMatrixFileFound Then
                ' Load the matrix
                objMatrix = NAligner.Matrix.Load(strMatrixName)
                blnSuccess = True
            Else
                ' Matrix file not found; cannot continue
                blnSuccess = False
            End If

        Catch ex As Exception
            Console.WriteLine("Error intitializing matrix " & strMatrixName & ": " & ex.Message)
            blnSuccess = False
        End Try

        Return blnSuccess

    End Function

    Private Sub InitializeLocalVariables()
        mMatrixName = DEFAULT_MATRIX_NAME
        mGapOpenPenalty = DEFAULT_GAP_OPEN_PENALTY
        mGapExtendPenalty = DEFAULT_GAP_EXTEND_PENALTY

        mDelimitedFileFormatCode = ProteinFileReader.DelimitedFileReader.eDelimitedFileFormatCode.ProteinName_Description_Sequence
        mDisplayAlignedPeptides = False

        mAlignmentScoreThreshold = DEFAULT_ALIGNMENT_SCORE_THRESHOLD

    End Sub

    Protected Function PreloadProteinInfo(ByRef objProteinFileReader As ProteinFileReader.ProteinFileReaderBaseClass, ByVal strProteinFilePath As String) As Boolean

        mProteinCount = 0
        ReDim mProteinInfo(99)

        If Not objProteinFileReader.OpenFile(strProteinFilePath) Then
            Console.WriteLine("Error opening the protein input file: " & strProteinFilePath)
            Return False
        End If

        Do While objProteinFileReader.ReadNextProteinEntry
            If mProteinCount >= mProteinInfo.Length Then
                ReDim Preserve mProteinInfo(mProteinInfo.Length * 2 - 1)

            End If

            With mProteinInfo(mProteinCount)
                .ProteinName = String.Copy(objProteinFileReader.ProteinName)
                .ProteinSequence = String.Copy(objProteinFileReader.ProteinSequence)
            End With

            mProteinCount += 1
        Loop

        objProteinFileReader.CloseFile()
        Return True

    End Function

    Private Sub ReportProgress(ByVal strProgress As String)
        Console.WriteLine(System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss tt") & " " & strProgress)
    End Sub

    Private Sub WriteAlignmentEntry(ByRef srOutFile As System.IO.StreamWriter, ByRef objAlignment As NAligner.Alignment, ByVal intLinesRead As Integer, ByVal intPeptideLength As Integer, ByRef strProteinName As String, ByVal intProteinSequenceLength As Integer, ByRef strRemainder As String)

        Static objAlignedPair As NAligner.formats.Pair

        Dim udtAlignmentStats As udtAlignmentStatsType
        Dim strLineOut As String

        ' Update the AlignmentScore histogram
        udtAlignmentStats.Score = CInt(objAlignment.Score)

        If udtAlignmentStats.Score < 0 Then
            mAlignmentScoreHistogram(0) += 1
        ElseIf udtAlignmentStats.Score < MAX_ALIGNMENT_SCORE_TO_TRACK Then
            mAlignmentScoreHistogram(udtAlignmentStats.Score) += 1
        Else
            mAlignmentScoreHistogram(MAX_ALIGNMENT_SCORE_TO_TRACK) += 1
        End If

        If objAlignment.Score > mAlignmentScoreThreshold Then

            With udtAlignmentStats
                ' Score was already assigned earlier in this function; no need to re-assign
                ' .Score = CInt(objAlignment.Score)
                .Identity = objAlignment.Identity
                .Similarity = objAlignment.Similarity
                .Gaps = objAlignment.Gaps

                .PeptideResidueCountAligned = CountLetters(objAlignment.Sequence1)
                .ProteinResidueCountAligned = CountLetters(objAlignment.Sequence2)

                .PeptideIdentityCoverage = .Identity / CSng(intPeptideLength) * 100
                .PeptideSimilarityCoverage = .Similarity / CSng(intPeptideLength) * 100
                .PeptideResidueCoverage = .PeptideResidueCountAligned / CSng(intPeptideLength) * 100
                .ProteinResidueCoverage = .ProteinResidueCountAligned / CSng(intProteinSequenceLength) * 100

            End With

            strLineOut = intLinesRead.ToString & ControlChars.Tab & _
                    intPeptideLength.ToString() & ControlChars.Tab & _
                    intProteinSequenceLength.ToString() & ControlChars.Tab & _
                    udtAlignmentStats.PeptideResidueCountAligned.ToString() & ControlChars.Tab & _
                    WriteEntryPercentage(udtAlignmentStats.PeptideIdentityCoverage) & ControlChars.Tab & _
                    WriteEntryPercentage(udtAlignmentStats.PeptideSimilarityCoverage) & ControlChars.Tab & _
                    WriteEntryPercentage(udtAlignmentStats.PeptideResidueCoverage) & ControlChars.Tab & _
                    WriteEntryPercentage(udtAlignmentStats.ProteinResidueCoverage) & ControlChars.Tab & _
                    udtAlignmentStats.Score.ToString() & ControlChars.Tab & _
                    udtAlignmentStats.Identity.ToString() & ControlChars.Tab & _
                    udtAlignmentStats.Similarity.ToString() & ControlChars.Tab & _
                    udtAlignmentStats.Gaps.ToString() & ControlChars.Tab & _
                    objAlignment.Sequence1() & ControlChars.Tab & _
                    objAlignment.MarkupLine() & ControlChars.Tab & _
                    objAlignment.Sequence2() & ControlChars.Tab & _
                    strProteinName

            If Not strRemainder Is Nothing AndAlso strRemainder.Length > 0 Then
                strLineOut &= ControlChars.Tab & strRemainder
            End If

            srOutFile.WriteLine(strLineOut)

            If mDisplayAlignedPeptides Then
                If objAlignedPair Is Nothing Then
                    objAlignedPair = New NAligner.formats.Pair
                End If

                Console.WriteLine()

                Console.WriteLine( _
                                ("Score: " & Math.Round(udtAlignmentStats.Score, 1).ToString).PadRight(21) & _
                                ("Identity: " & Math.Round(udtAlignmentStats.Identity, 1).ToString).PadRight(23) & _
                                ("Similarity: " & Math.Round(udtAlignmentStats.Similarity, 1).ToString).PadRight(20) & _
                                ("Gaps: " & Math.Round(udtAlignmentStats.Gaps, 1).ToString).PadRight(15))

                Console.WriteLine( _
                                ("Peptide Length: " & intPeptideLength.ToString).PadRight(21) & _
                                ("Protein Length: " & intProteinSequenceLength.ToString).PadRight(23) & _
                                ("Aligned Length: " & udtAlignmentStats.PeptideResidueCountAligned.ToString).PadRight(20))

                Console.WriteLine( _
                                ("Peptide Identity Coverage: " & WriteEntryPercentage(udtAlignmentStats.PeptideIdentityCoverage)).PadRight(35) & _
                                ("Peptide Similarity Coverage: " & WriteEntryPercentage(udtAlignmentStats.PeptideSimilarityCoverage)).PadRight(30))

                Console.WriteLine( _
                                ("Peptide Coverage: " & WriteEntryPercentage(udtAlignmentStats.PeptideResidueCoverage)).PadRight(35) & _
                                ("Protein Coverage: " & WriteEntryPercentage(udtAlignmentStats.ProteinResidueCoverage)).PadRight(30))

                Console.WriteLine()
                Console.WriteLine(objAlignedPair.Format(objAlignment))

            End If

        End If

    End Sub

    Private Function WriteEntryPercentage(ByVal sngPercentage As Single) As String
        If sngPercentage > 0 Then
            Return Math.Round(sngPercentage, 1).ToString & "%"
        Else
            Return "0%"
        End If
    End Function

    Private Function WriteEntryPercentage(ByVal intCount As Integer, ByVal intTotal As Integer) As String
        If intTotal > 0 Then
            Return Math.Round(intCount / CDbl(intTotal) * 100, 1).ToString & "%"
        Else
            Return "0%"
        End If
    End Function

    Private Function WriteHistogram(ByVal strHistogramOutputFileName As String, ByVal dtProcessingStartTime As DateTime) As Boolean

        Dim srHistogramFile As System.IO.StreamWriter

        Dim intIndex As Integer
        Dim intIndexMax As Integer
        Dim intEntryCount As Integer

        Dim intProcessingTimeSeconds As Integer
        Dim sngProcessingRate As Single

        Try

            srHistogramFile = New System.IO.StreamWriter(New System.IO.FileStream(strHistogramOutputFileName, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.Read))

        Catch ex As Exception
            Console.WriteLine("Error opening the histogram output file (" & strHistogramOutputFileName & "): " & ex.Message)
            Return False
        End Try

        Try
            ' Find the largest non-zero value in mAlignmentScoreHistogram
            For intIndex = mAlignmentScoreHistogram.Length - 1 To 0 Step -1
                If mAlignmentScoreHistogram(intIndex) > 0 Then
                    Exit For
                End If
            Next intIndex
            intIndexMax = intIndex

            ' Compute the sum of all entries in mAlignmentScoreHistogram()
            intEntryCount = 0
            For intIndex = 0 To intIndexMax
                intEntryCount += mAlignmentScoreHistogram(intIndex)
            Next

            intProcessingTimeSeconds = CInt(System.DateTime.UtcNow.Subtract(dtProcessingStartTime).TotalSeconds)

            If intProcessingTimeSeconds > 0 Then
                sngProcessingRate = intEntryCount / CSng(intProcessingTimeSeconds)
            Else
                sngProcessingRate = 0
            End If

            ' Write the processing stats
            srHistogramFile.WriteLine("Total entries" & ControlChars.Tab & intEntryCount.ToString)
            srHistogramFile.WriteLine("Processing time" & ControlChars.Tab & intProcessingTimeSeconds.ToString & ControlChars.Tab & "seconds")
            srHistogramFile.WriteLine("Processing rate" & ControlChars.Tab & Math.Round(sngProcessingRate, 1).ToString & ControlChars.Tab & "entries/sec")

            srHistogramFile.WriteLine()

            ' Write the histogram header line
            srHistogramFile.WriteLine("Score" & ControlChars.Tab & "Score_Count")

            ' Write out the histogram values
            For intIndex = 0 To intIndexMax
                srHistogramFile.WriteLine(intIndex & ControlChars.Tab & mAlignmentScoreHistogram(intIndex).ToString)
            Next

        Catch ex As Exception
            Console.WriteLine("Error writing to the histogram output file (" & strHistogramOutputFileName & "): " & ex.Message)
            Return False
        Finally
            If Not srHistogramFile Is Nothing Then srHistogramFile.Close()
        End Try

        Return True

    End Function
End Class
