Option Strict On

' This class can be used to display alignment results
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

Public Class clsAlignmentSummary

    Public Structure udtAlignmentStatsType
        Public AlignmentScore As Double
        Public SequenceLength As Integer
        Public IdentityCount As Integer
        Public SimilarityCount As Integer
        Public GapCount As Integer
    End Structure

    Public Sub DisplaySummary(ByRef udtAlignmentStats As udtAlignmentStatsType, ByVal strAlignedSeq1 As String, ByVal strAlignedSeq2 As String, ByVal strDiffSeq As String, ByVal intAlignedSeq1StartResidue As Integer, ByVal intAlignedSeq2StartResidue As Integer)
        With udtAlignmentStats
            Console.WriteLine()
            Console.WriteLine("Length: " & .SequenceLength.ToString)
            Console.WriteLine("Identity:      " & .IdentityCount.ToString & "/" & .SequenceLength.ToString & DisplayResultsAppendPercent(.IdentityCount, .SequenceLength))
            Console.WriteLine("Similarity:    " & .SimilarityCount.ToString & "/" & .SequenceLength.ToString & DisplayResultsAppendPercent(.SimilarityCount, .SequenceLength))
            Console.WriteLine("Gaps:          " & .GapCount.ToString & "/" & .SequenceLength.ToString & DisplayResultsAppendPercent(.GapCount, .SequenceLength))
            Console.WriteLine("Score: " & Math.Round(.AlignmentScore, 1).ToString)
        End With

        DisplayResultsPrintAlignedStrings( _
                 strAlignedSeq1, _
                 strAlignedSeq2, _
                 strDiffSeq, _
                intAlignedSeq1StartResidue, _
                intAlignedSeq2StartResidue)

    End Sub

    Protected Sub DisplayResultsPrintAlignedStrings(ByVal strAlignedSeq1 As String, ByVal strAlignedSeq2 As String, ByVal strDiffSeq As String, ByVal intAlignedSeq1StartResidue As Integer, ByVal intAlignedSeq2StartResidue As Integer)
        DisplayResultsPrintAlignedStrings(strAlignedSeq1, strAlignedSeq2, strDiffSeq, intAlignedSeq1StartResidue, intAlignedSeq2StartResidue, 60)
    End Sub

    Protected Sub DisplayResultsPrintAlignedStrings(ByVal strAlignedSeq1 As String, ByVal strAlignedSeq2 As String, ByVal strDiffSeq As String, ByVal intAlignedSeq1StartResidue As Integer, ByVal intAlignedSeq2StartResidue As Integer, ByVal intLineLength As Integer)

        Dim intStartIndex As Integer

        intStartIndex = 0
        Do
            Console.WriteLine()

            DisplayResultsPrintStringRemainder(strAlignedSeq1, intStartIndex, intLineLength, intAlignedSeq1StartResidue)
            DisplayResultsPrintStringRemainder(strDiffSeq, intStartIndex, intLineLength, 0)
            DisplayResultsPrintStringRemainder(strAlignedSeq2, intStartIndex, intLineLength, intAlignedSeq2StartResidue)

            intStartIndex += intLineLength
        Loop While intStartIndex < strAlignedSeq1.Length

    End Sub

    Private Shared Sub DisplayResultsPrintStringRemainder(ByVal strText As String, ByVal intStartIndex As Integer, ByVal intLength As Integer, ByVal intStartResidue As Integer)
        Const PREFIX_LENGTH As Integer = 5

        Dim strPrefix As String

        If intStartResidue > 0 Then
            strPrefix = (intStartResidue + intStartIndex + 1).ToString
        Else
            strPrefix = "  "
        End If
        strPrefix = strPrefix.PadRight(PREFIX_LENGTH)

        If intStartIndex + intLength <= strText.Length Then
            Console.WriteLine(strPrefix & strText.Substring(intStartIndex, intLength))
        Else
            Console.WriteLine(strPrefix & strText.Substring(intStartIndex, strText.Length - intStartIndex))
        End If
    End Sub

    Private Function DisplayResultsAppendPercent(ByVal intCount As Integer, ByVal intTotal As Integer) As String
        If intTotal > 0 Then
            Return " (" & Math.Round(intCount / CDbl(intTotal) * 100, 1).ToString & "%)"
        Else
            Return " (0%)"
        End If
    End Function

End Class
