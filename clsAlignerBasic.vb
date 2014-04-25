Option Strict On

' This class is a VB.NET port of the C code available at http://mycplus.com/out.asp?CID=1&SCID=114
'
' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started December 1, 2007
' Copyright 2007, Battelle Memorial Institute.  All Rights Reserved.

' E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
' Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
' -------------------------------------------------------------------------------
' 

'====
' opsa.cpp
' - Optimal Pair-wise Sequence Alignment
' - implements Smith-Waterman with affine gap penalties
' - requires at least one blank line between the two sequences
' - ignores input lines with non-alphabetical characters
' Notes
' - This program is provided as is with no warranty.
' - The author is not responsible for any damage caused
'   either directly or indirectly by using this program.
' - Anybody is free to do whatever he/she wants with this
'   program as long as this header section is preserved.
' Created on 2005-02-01 by
' - Roger Zhang (rogerz@cs.dal.ca)
' Modifications
' - Roger Zhang on 2005-02-18
'   - Hard coded the blosum matrix to remove external dependency
' -
' Last compiled under Linux with gcc-3
'====

Public Class clsAlignerBasic

    Protected Const LARGE_NUMBER As Double = 65536
    Protected Const GAP_OPENING_COST As Double = 10
    Protected Const GAP_EXTENSION_COST As Double = 0.5
    Protected Const NEW_GAP_COST As Double = GAP_OPENING_COST + GAP_EXTENSION_COST

    Protected mBLOSUM As Int16(,)

    Public Sub New()
        InitializeBlosum()
    End Sub

    Public Function AlignSequences(ByRef strSequence1 As String, ByRef strSequence2 As String, ByRef strDiffSeq As String, ByRef udtAlignmentStats As clsAlignmentSummary.udtAlignmentStatsType) As Double

        Dim dblAlignmentScore As Double

        strSequence1 = CleanSeq(strSequence1)
        strSequence2 = CleanSeq(strSequence2)

        dblAlignmentScore = PerformAlignment(strSequence1, strSequence2)

        udtAlignmentStats.AlignmentScore = dblAlignmentScore

        strDiffSeq = ConstructDiffSeq(strSequence1, strSequence2, udtAlignmentStats)


        Return dblAlignmentScore

    End Function

    Protected Function CleanSeq(ByVal strSequence As String) As String

        Dim intIndex As Integer
        Dim sbCleanSequence As New System.Text.StringBuilder

        If strSequence Is Nothing Then
            Return String.Empty
        Else
            strSequence = strSequence.Trim

            For intIndex = 0 To strSequence.Length - 1
                If Char.IsUpper(strSequence.Chars(intIndex)) Then
                    sbCleanSequence.Append(strSequence.Chars(intIndex))
                End If
            Next

            Return sbCleanSequence.ToString
        End If

    End Function

    Protected Function ConstructDiffSeq(ByVal strAlignedSeq1 As String, ByVal strAlignedSeq2 As String, ByRef udtAlignmentStats As clsAlignmentSummary.udtAlignmentStatsType) As String
        Dim intIndex As Integer

        Dim strDiffSeq As New System.text.StringBuilder

        With udtAlignmentStats
            .SequenceLength = strAlignedSeq1.Length
            .IdentityCount = 0
            .SimilarityCount = 0
            .GapCount = 0
        End With

        If strAlignedSeq1.Length <> strAlignedSeq2.Length Then
            ' Can't construct the difference sequence if the aligned sequences are of differing lengths
            Return String.Empty
        End If

        ' Construct the difference string
        ' Vertical bar means identical residue
        ' Colon means similar residue
        ' Space means a gap is present
        ' Period means difference

        For intIndex = 0 To strAlignedSeq1.Length - 1
            If strAlignedSeq1.Chars(intIndex) = strAlignedSeq2.Chars(intIndex) Then
                strDiffSeq.Append("|"c)
                udtAlignmentStats.IdentityCount += 1

            ElseIf strAlignedSeq1.Chars(intIndex) = "-" OrElse strAlignedSeq2.Chars(intIndex) = "-" Then
                strDiffSeq.Append(" "c)
                udtAlignmentStats.GapCount += 1

            ElseIf strAlignedSeq1.Chars(intIndex) = "I" AndAlso strAlignedSeq2.Chars(intIndex) = "L" OrElse _
                   strAlignedSeq1.Chars(intIndex) = "L" AndAlso strAlignedSeq2.Chars(intIndex) = "I" Then
                strDiffSeq.Append(":"c)
                udtAlignmentStats.SimilarityCount += 1


            ElseIf strAlignedSeq1.Chars(intIndex) = "D" AndAlso strAlignedSeq2.Chars(intIndex) = "N" OrElse _
                strAlignedSeq1.Chars(intIndex) = "N" AndAlso strAlignedSeq2.Chars(intIndex) = "D" Then
                strDiffSeq.Append(":"c)
                udtAlignmentStats.SimilarityCount += 1

            ElseIf strAlignedSeq1.Chars(intIndex) = "S" AndAlso strAlignedSeq2.Chars(intIndex) = "T" OrElse _
                strAlignedSeq1.Chars(intIndex) = "T" AndAlso strAlignedSeq2.Chars(intIndex) = "S" Then
                strDiffSeq.Append(":"c)
                udtAlignmentStats.SimilarityCount += 1

            Else
                strDiffSeq.Append("."c)
            End If
        Next

        ' Bump the Similarity count up by the Identity count
        udtAlignmentStats.SimilarityCount += udtAlignmentStats.IdentityCount

        Return strDiffSeq.ToString

    End Function

    Public Sub DisplayResults(ByVal strAlignedSeq1 As String, ByVal strAlignedSeq2 As String, ByVal strDiffSeq As String, ByRef udtAlignmentStats As clsAlignmentSummary.udtAlignmentStatsType, ByVal blnRemoveLeadingDashes As Boolean)

        Dim intStartIndex As Integer
        Dim intEndIndex As Integer
        Dim objAlignmentSummary As New clsAlignmentSummary

        If blnRemoveLeadingDashes Then
            For intStartIndex = 0 To strAlignedSeq1.Length - 2
                If strAlignedSeq1.Chars(intStartIndex) <> "-"c Then
                    Exit For
                End If
            Next

            If intStartIndex = 0 Then
                For intStartIndex = 0 To strAlignedSeq2.Length - 2
                    If strAlignedSeq2.Chars(intStartIndex) <> "-"c Then
                        Exit For
                    End If
                Next
            End If


            For intEndIndex = strAlignedSeq1.Length - 1 To 1 Step -1
                If strAlignedSeq1.Chars(intEndIndex) <> "-"c Then
                    Exit For
                End If
            Next

            If intEndIndex = strAlignedSeq1.Length - 1 Then
                For intEndIndex = strAlignedSeq2.Length - 1 To 1 Step -1
                    If strAlignedSeq2.Chars(intEndIndex) <> "-"c Then
                        Exit For
                    End If
                Next
            End If

            If intEndIndex < intStartIndex Then intEndIndex = intStartIndex

            objAlignmentSummary.DisplaySummary( _
                    udtAlignmentStats, _
                    strAlignedSeq1.Substring(intStartIndex, intEndIndex - intStartIndex + 1), _
                    strAlignedSeq2.Substring(intStartIndex, intEndIndex - intStartIndex + 1), _
                    strDiffSeq.Substring(intStartIndex, intEndIndex - intStartIndex + 1), _
                    1, _
                    1)

        Else
            objAlignmentSummary.DisplaySummary( _
                    udtAlignmentStats, _
                    strAlignedSeq1, _
                    strAlignedSeq2, _
                    strDiffSeq, _
                    1, _
                    1)

        End If


    End Sub

    Protected Sub InitializeBlosum()

        Dim intRowCount As Integer

        intRowCount = 0
        ReDim mBLOSUM(24, 24)

        ' The blosum 62 scoring matrix
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {4, 0, 0, -2, -1, -2, 0, -2, -1, 0, -1, -1, -1, -2, 0, -1, -1, -1, 1, 0, 0, 0, -3, 0, -2}) ' A
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}) ' B
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {0, 0, 9, -3, -4, -2, -3, -3, -1, 0, -3, -1, -1, -3, 0, -3, -3, -3, -1, -1, 0, -1, -2, 0, -2}) ' c
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {-2, 0, -3, 6, 2, -3, -1, -1, -3, 0, -1, -4, -3, 1, 0, -1, 0, -2, 0, -1, 0, -3, -4, 0, -3}) ' D
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {-1, 0, -4, 2, 5, -3, -2, 0, -3, 0, 1, -3, -2, 0, 0, -1, 2, 0, 0, -1, 0, -2, -3, 0, -2}) ' E
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {-2, 0, -2, -3, -3, 6, -3, -1, 0, 0, -3, 0, 0, -3, 0, -4, -3, -3, -2, -2, 0, -1, 1, 0, 3}) ' F
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {0, 0, -3, -1, -2, -3, 6, -2, -4, 0, -2, -4, -3, 0, 0, -2, -2, -2, 0, -2, 0, -3, -2, 0, -3})  ' G
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {-2, 0, -3, -1, 0, -1, -2, 8, -3, 0, -1, -3, -2, 1, 0, -2, 0, 0, -1, -2, 0, -3, -2, 0, 2})  ' H
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {-1, 0, -1, -3, -3, 0, -4, -3, 4, 0, -3, 2, 1, -3, 0, -3, -3, -3, -2, -1, 0, 3, -3, 0, -1}) ' I
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}) ' J
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {-1, 0, -3, -1, 1, -3, -2, -1, -3, 0, 5, -2, -1, 0, 0, -1, 1, 2, 0, -1, 0, -2, -3, 0, -2}) ' K
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {-1, 0, -1, -4, -3, 0, -4, -3, 2, 0, -2, 4, 2, -3, 0, -3, -2, -2, -2, -1, 0, 1, -2, 0, -1}) ' L
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {-1, 0, -1, -3, -2, 0, -3, -2, 1, 0, -1, 2, 5, -2, 0, -2, 0, -1, -1, -1, 0, 1, -1, 0, -1})  ' M
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {-2, 0, -3, 1, 0, -3, 0, 1, -3, 0, 0, -3, -2, 6, 0, -2, 0, 0, 1, 0, 0, -3, -4, 0, -2}) ' N
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}) ' O
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {-1, 0, -3, -1, -1, -4, -2, -2, -3, 0, -1, -3, -2, -2, 0, 7, -1, -2, -1, -1, 0, -2, -4, 0, -3})  ' P
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {-1, 0, -3, 0, 2, -3, -2, 0, -3, 0, 1, -2, 0, 0, 0, -1, 5, 1, 0, -1, 0, -2, -2, 0, -1}) ' Q
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {-1, 0, -3, -2, 0, -3, -2, 0, -3, 0, 2, -2, -1, 0, 0, -2, 1, 5, -1, -1, 0, -3, -3, 0, -2}) ' R
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {1, 0, -1, 0, 0, -2, 0, -1, -2, 0, 0, -2, -1, 1, 0, -1, 0, -1, 4, 1, 0, -2, -3, 0, -2}) ' S
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {0, 0, -1, -1, -1, -2, -2, -2, -1, 0, -1, -1, -1, 0, 0, -1, -1, -1, 1, 5, 0, 0, -2, 0, -2}) ' T
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}) ' U
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {0, 0, -1, -3, -2, -1, -3, -3, 3, 0, -2, 1, 1, -3, 0, -2, -2, -3, -2, 0, 0, 4, -3, 0, -1}) ' V
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {-3, 0, -2, -4, -3, 1, -2, -2, -3, 0, -3, -2, -1, -4, 0, -4, -2, -3, -3, -2, 0, -3, 11, 0, 2})  ' W
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}) ' X
        AddToBlosum(mBLOSUM, intRowCount, New Int16() {-2, 0, -2, -3, -2, 3, -3, 2, -1, 0, -2, -1, -1, -2, 0, -3, -1, -2, -2, -2, 0, -1, 2, 0, 7}) ' Y

    End Sub

    Private Sub AddToBlosum(ByRef intBlosum(,) As Int16, ByRef intRowCount As Integer, ByVal intRowIndexValues() As Int16)
        Dim intIndex As Integer


        For intIndex = 0 To intRowIndexValues.Length - 1
            intBlosum(intRowCount, intIndex) = intRowIndexValues(intIndex)
        Next

        intRowCount += 1
    End Sub

    Protected Function MaxOf3(ByVal x As Double, ByVal y As Double, ByVal z As Double) As Double
        Return Math.Max(Math.Max(x, y), z)
    End Function

    Protected Function PerformAlignment(ByRef strSequence1 As String, ByRef strSequence2 As String) As Double

        Dim n As Integer
        Dim m As Integer
        Dim i As Integer
        Dim j As Integer

        Dim intAsciiA As Integer

        Dim r(,) As Double
        Dim t(,) As Double
        Dim s(,) As Double

        ' Note: System.Convert.ToInt32(chChar) returns the Ascii value of chChar; equivalent to Asc(chChar)
        intAsciiA = System.Convert.ToInt32("A"c)

        n = strSequence1.Length() + 1
        m = strSequence2.Length() + 1

        ReDim r(n - 1, m - 1)
        ReDim s(n - 1, m - 1)
        ReDim t(n - 1, m - 1)

        '====
        ' initialization
        r(0, 0) = 0
        t(0, 0) = 0
        s(0, 0) = 0

        For i = 1 To n - 1
            r(i, 0) = -LARGE_NUMBER
            s(i, 0) = -GAP_OPENING_COST - i * GAP_EXTENSION_COST
            t(i, 0) = s(i, 0)
        Next i

        For j = 1 To m - 1
            t(0, j) = -LARGE_NUMBER
            s(0, j) = -GAP_OPENING_COST - j * GAP_EXTENSION_COST
            r(0, j) = s(0, j)
        Next j

        '====
        ' Smith-Waterman with affine gap costs

        For i = 1 To n - 1
            For j = 1 To m - 1
                r(i, j) = Math.Max(r(i, j - 1) - GAP_EXTENSION_COST, s(i, j - 1) - NEW_GAP_COST)
                t(i, j) = Math.Max(t(i - 1, j) - GAP_EXTENSION_COST, s(i - 1, j) - NEW_GAP_COST)
                s(i, j) = MaxOf3( _
                      s(i - 1, j - 1) + mBLOSUM(System.Convert.ToInt32(strSequence1.Chars(i - 1)) - intAsciiA, System.Convert.ToInt32(strSequence2.Chars(j - 1)) - intAsciiA), _
                      r(i, j), _
                      t(i, j))
            Next j
        Next i

        '====
        ' back tracking

        i = n - 1
        j = m - 1

        Do While (i > 0 Or j > 0)
            If (s(i, j) = r(i, j)) Then
                strSequence1 = strSequence1.Insert(i, "-"c)
                j -= 1
            ElseIf s(i, j) = t(i, j) Then
                strSequence2 = strSequence2.Insert(j, "-"c)
                i -= 1
            Else
                i -= 1
                j -= 1
            End If

        Loop

        '====
        ' final score

        Return s(n - 1, m - 1)

    End Function

End Class
