Option Strict On

' This program uses clsAlignPeptides to read a file containing peptides and align
' them to a Fasta file or delimited text file with protein sequences (using NAligner.dll, a C# DLL)
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

Module modMain

	Public Const PROGRAM_DATE As String = "April 25, 2014"

    Private mPeptideInputFilePath As String = String.Empty
    Private mProteinFilePath As String = String.Empty
    Private mDelimitedFileFormatCode As ProteinFileReader.DelimitedFileReader.eDelimitedFileFormatCode

    Private mMatrixName As String
    Private mGapOpenPenalty As Single
    Private mGapExtendPenalty As Single
    Private mAlignmentScoreThreshold As Integer
    Private mDisplayAlignedPeptides As Boolean

    Private mRunTestFunctions As Boolean

	Public Function Main() As Integer
		Dim intReturnCode As Integer
		Dim objAlignPeptides As clsAlignPeptides
		Dim objParseCommandLine As New clsParseCommandLine

		Dim blnProceed As Boolean
		Dim blnSuccess As Boolean

		Try
			' Set the default values
			mMatrixName = clsAlignPeptides.DEFAULT_MATRIX_NAME
			mGapOpenPenalty = clsAlignPeptides.DEFAULT_GAP_OPEN_PENALTY
			mGapExtendPenalty = clsAlignPeptides.DEFAULT_GAP_EXTEND_PENALTY
			mAlignmentScoreThreshold = clsAlignPeptides.DEFAULT_ALIGNMENT_SCORE_THRESHOLD
			mDisplayAlignedPeptides = False

			mDelimitedFileFormatCode = ProteinFileReader.DelimitedFileReader.eDelimitedFileFormatCode.ProteinName_Description_Sequence

			mRunTestFunctions = False

			blnProceed = False
			If objParseCommandLine.ParseCommandLine Then
				If SetOptionsUsingCommandLineParameters(objParseCommandLine) Then blnProceed = True
			End If

			If mRunTestFunctions Then
				modTestFunctions.TestAlignment()
				Exit Function
			End If

			If Not blnProceed OrElse _
			   objParseCommandLine.NeedToShowHelp OrElse _
			   mPeptideInputFilePath.Length = 0 OrElse _
			   mProteinFilePath.Length = 0 Then
				ShowProgramHelp()
				intReturnCode = -1
			Else
				objAlignPeptides = New clsAlignPeptides

				With objAlignPeptides
					.MatrixName = mMatrixName
					.GapOpenPenalty = mGapOpenPenalty
					.GapExtendPenalty = mGapExtendPenalty
					.AlignmentScoreThreshold = mAlignmentScoreThreshold
					.DisplayAlignedPeptides = mDisplayAlignedPeptides

					.DelimitedFileFormatCode = mDelimitedFileFormatCode

					''If Not mParameterFilePath Is Nothing AndAlso mParameterFilePath.Length > 0 Then
					''    .LoadParameterFileSettings(mParameterFilePath)
					''End If
				End With

				blnSuccess = objAlignPeptides.AlignPeptidesToProteinFile(mPeptideInputFilePath, mProteinFilePath)

				If blnSuccess Then
					intReturnCode = 0
				Else
					intReturnCode = -1
				End If

			End If

		Catch ex As Exception
			Console.WriteLine("Error occurred in modMain->Main: " & ControlChars.NewLine & ex.Message)
			intReturnCode = -1
		End Try

		Return intReturnCode

	End Function

    Private Function GetAppVersion() As String
		Return Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString & " (" & PROGRAM_DATE & ")"
    End Function

    Private Function SetOptionsUsingCommandLineParameters(ByVal objParseCommandLine As clsParseCommandLine) As Boolean
        ' Returns True if no problems; otherwise, returns false

        Dim strValue As String = String.Empty
		Dim strValidParameters() As String = New String() {"D", "M", "O", "E", "S", "A", "T"}
        Dim intValue As Integer

        Try
            ' Make sure no invalid parameters are present
            If objParseCommandLine.InvalidParametersPresent(strValidParameters) Then
                Return False
            Else

                ' Query objParseCommandLine to see if various parameters are present
                With objParseCommandLine
                    If .RetrieveValueForParameter("T", strValue) Then
                        mRunTestFunctions = True
                        Return True
                    End If

                    ' This program requires that both input files be entered as non-switch parameters
                    If .NonSwitchParameterCount < 2 Then
                        Return False
                    End If

                    mPeptideInputFilePath = .RetrieveNonSwitchParameter(0)
                    mProteinFilePath = .RetrieveNonSwitchParameter(1)

                    If .RetrieveValueForParameter("D", strValue) Then
                        If Integer.TryParse(strValue, intValue) Then
                            mDelimitedFileFormatCode = CType(intValue, ProteinFileReader.DelimitedFileReader.eDelimitedFileFormatCode)
                        End If
                    End If

                    If .RetrieveValueForParameter("M", strValue) Then
                        mMatrixName = String.Copy(strValue)
                    End If

                    If .RetrieveValueForParameter("O", strValue) Then
                        Single.TryParse(strValue, mGapOpenPenalty)
                    End If

                    If .RetrieveValueForParameter("E", strValue) Then
                        Single.TryParse(strValue, mGapExtendPenalty)                       
                    End If

                    If .RetrieveValueForParameter("S", strValue) Then
                        Integer.TryParse(strValue, mAlignmentScoreThreshold)
                    End If

                    If .RetrieveValueForParameter("A", strValue) Then
                        mDisplayAlignedPeptides = True
                    End If
				End With

                Return True
            End If

		Catch ex As Exception
			Console.WriteLine("Error parsing the command line parameters: " & ControlChars.NewLine & ex.Message)
			Return False
        End Try

    End Function

    Private Sub ShowProgramHelp()

        Try

            Console.WriteLine("This program will read a file containing peptides and align them to a file of " & _
                              "protein sequences (.Fasta or delimited text) using Smith-Waterman alignment.  " & _
                              "The program requires that NAligner.dll be present in the same folder as this .Exe")
            Console.WriteLine()

			Console.WriteLine("Program syntax: " & ControlChars.NewLine & IO.Path.GetFileName(Reflection.Assembly.GetExecutingAssembly().Location))
            Console.WriteLine(" PeptideInputFile.txt ProteinFileName [/D:DelimitedFileFormatCode] ")
            Console.WriteLine(" [/M:MatrixName] [/O:GapOpenValue] [/E:GapExtendValue]")
            Console.WriteLine(" [/S:AlignmentScoreThreshold] [/A]")
            Console.WriteLine()

            Console.WriteLine("The two input file names are required. If either filename contains spaces, then surround it with double quotes.")
            Console.WriteLine("You can provide a Fasta file (extension .Fasta) or " & _
                              "a tab-delimited text file of proteins.  If using the tab-delimited file, then optionally use /D to define the delimited file columns.")
            Console.WriteLine()

            Console.WriteLine("Use /D to specify the column format in a delimited protein file.  The default format is type 2, which means 3 columns " & _
                              "(ProteinName, ProteinDescription, and Sequence).  The available formats are:")
            Console.WriteLine(clsAlignPeptides.GetDelimitedFileFormats())
            Console.WriteLine()

            Console.WriteLine("Use /M to define the matrix name (default is " & clsAlignPeptides.DEFAULT_MATRIX_NAME & ").  If not using " & _
                              clsAlignPeptides.DEFAULT_MATRIX_NAME & ", then the matrix file must be present in the program folder.")

            Console.WriteLine("/O defines the gap open penalty (default " & clsAlignPeptides.DEFAULT_GAP_OPEN_PENALTY.ToString & ")")
            Console.WriteLine("/E defines the gap extend penalty (default " & clsAlignPeptides.DEFAULT_GAP_EXTEND_PENALTY.ToString & ")")
            Console.WriteLine("/S defines the alignment score threshold (default " & clsAlignPeptides.DEFAULT_ALIGNMENT_SCORE_THRESHOLD.ToString & "); alignments below this score value will not be saved")
            Console.WriteLine("/A instructs the software to display the aligned peptides for all alignments passing the score threshold")
            Console.WriteLine()


            Console.WriteLine("Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2007")
            Console.WriteLine("Version: " & GetAppVersion())
            Console.WriteLine()

			Console.WriteLine("E-mail: matthew.monroe@pnnl.gov or matt@alchemistmatt.com")
			Console.WriteLine("Website: http://panomics.pnl.gov/ or http://www.sysbio.org/resources/staff/")

            ' Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
			Threading.Thread.Sleep(750)


		Catch ex As Exception
			Console.WriteLine("Error displaying the program syntax: " & ex.Message)
        End Try

    End Sub

End Module
