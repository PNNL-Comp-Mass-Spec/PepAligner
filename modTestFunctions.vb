Option Strict On

Module modTestFunctions

    Public Sub TestAlignment()

        Dim objAligner As New clsAlignerBasic

        Console.WriteLine("Test port of C alignment code to VB.NET: peptide vs. peptide")
        TestPeptides(objAligner)

        PromptUser()
        Console.WriteLine()

        Console.WriteLine("Test port of C alignment code to VB.NET: peptide vs. protein")
        TestProteins(objAligner)

        PromptUser()
        Console.WriteLine()

        TestNAligner()

        PromptUser()

    End Sub

    Private Sub PromptUser()
        Console.WriteLine()
        Console.WriteLine("Press Enter to continue.")
        Console.ReadLine()
    End Sub

    Private Sub TestPeptides(ByRef objAligner As clsAlignerBasic)

        Dim strSeq1 As String
        Dim strSeq2 As String
        Dim strDiffSeq As String = String.Empty
        Dim dblScore As Double

        Dim udtAlignmentStats As clsAlignmentSummary.udtAlignmentStatsType

        strSeq1 = "ASPFTKQCILGGLNS"
        strSeq2 = "ASDFTKQCLIWLNS"

        dblScore = objAligner.AlignSequences(strSeq1, strSeq2, strDiffSeq, udtAlignmentStats)

        objAligner.DisplayResults(strSeq1, strSeq2, strDiffSeq, udtAlignmentStats, True)

    End Sub

    Private Sub TestProteins(ByRef objAligner As clsAlignerBasic)
        Dim strSeq1 As String
        Dim strSeq2 As String
        Dim strDiffSeq As String = String.Empty
        Dim dblScore As Double

        Dim udtAlignmentStats As clsAlignmentSummary.udtAlignmentStatsType

        strSeq1 = GetPeptideVsProtein1()
        strSeq2 = GetPeptideVsProtein2()

        dblScore = objAligner.AlignSequences(strSeq1, strSeq2, strDiffSeq, udtAlignmentStats)

        objAligner.DisplayResults(strSeq1, strSeq2, strDiffSeq, udtAlignmentStats, True)

    End Sub


    Private Sub TestNAligner()
        Const LOOP_COUNT As Integer = 500
        Dim dtStart As DateTime

        Dim objMatrix As NAligner.Matrix
        Dim objAligner As NAligner.Alignment

        Dim objDisplayer As New clsAlignerBasic
        Dim udtAlignmentStats As clsAlignmentSummary.udtAlignmentStatsType


        Dim opengap As Single
        Dim extendgap As Single

        Dim strProtein1 As String
        Dim strProtein2 As String

        Dim objSequence1 As NAligner.Sequence
        Dim objSequence2 As NAligner.Sequence

        ' NAligner defaults
        opengap = 15.0
        extendgap = 3.0

        ' Water defaults
        opengap = 10
        extendgap = 0.5

        objMatrix = NAligner.Matrix.Load("BLOSUM62")

        '
        ' Example using peptide vs. protein
        '
        Console.WriteLine("Now testing NAligner, which is a C# DLL")
        Console.WriteLine("Looping " & LOOP_COUNT.ToString & " times")

        dtStart = System.DateTime.UtcNow
        Console.WriteLine("Start: " & dtStart.ToLocalTime.ToString)

        strProtein2 = ">ProteinB  Protein" & ControlChars.NewLine & GetPeptideVsProtein2()
        objSequence2 = NAligner.Sequence.Parse(strProtein2)

        For intIndex As Integer = 1 To LOOP_COUNT
            strProtein1 = ">PeptideA  Peptide" & ControlChars.NewLine & GetPeptideVsProtein1()
            objSequence1 = NAligner.Sequence.Parse(strProtein1)

            objAligner = NAligner.SmithWatermanGotoh.Align(objSequence1, objSequence2, objMatrix, opengap, extendgap)

            With udtAlignmentStats
                .AlignmentScore = objAligner.Score
                .SequenceLength = objAligner.Sequence1.Length
                .IdentityCount = objAligner.Identity
                .SimilarityCount = objAligner.Similarity
                .GapCount = objAligner.Gaps
            End With
        Next

        Console.WriteLine("Elapsed time: " & System.DateTime.UtcNow.Subtract(dtStart).TotalSeconds.ToString & " seconds")
        Console.WriteLine()

        Console.WriteLine("Results for peptide vs. protein:")
        objDisplayer.DisplayResults(objAligner.Sequence1, objAligner.Sequence2, objAligner.MarkupLine, udtAlignmentStats, False)

        PromptUser()
        Console.WriteLine()


        '
        ' Example using two proteins
        '
        strProtein1 = ">100K_RAT  100 kDa protein (EC 6.3.2.-)." & ControlChars.NewLine & _
            "MMSARGDFLN YALSLMRSHN DEHSDVLPVL DVCSLKHVAY VFQALIYWIK AMNQQTTLDT" & _
            "PQLERKRTRE LLELGIDNED SEHENDDDTS QSATLNDKDD ESLPAETGQN HPFFRRSDSM" & _
            "TFLGCIPPNP FEVPLAEAIP LADQPHLLQP NARKEDLFGR PSQGLYSSSA GSGKCLVEVT" & _
            "MDRNCLEVLP TKMSYAANLK NVMNMQNRQK KAGEDQSMLA EEADSSKPGP SAHDVAAQLK" & _
            "SSLLAEIGLT ESEGPPLTSF RPQCSFMGMV ISHDMLLGRW RLSLELFGRV FMEDVGAEPG" & _
            "SILTELGGFE VKESKFRREM EKLRNQQSRD LSLEVDRDRD LLIQQTMRQL NNHFGRRCAT" & _
            "TPMAVHRVKV TFKDEPGEGS GVARSFYTAI AQAFLSNEKL PNLDCIQNAN KGTHTSLMQR" & _
            "LRNRGERDRE REREREMRRS SGLRAGSRRD RDRDFRRQLS IDTRPFRPAS EGNPSDDPDP" & _
            "LPAHRQALGE RLYPRVQAMQ PAFASKITGM LLELSPAQLL LLLASEDSLR ARVEEAMELI" & _
            "VAHGRENGAD SILDLGLLDS SEKVQENRKR HGSSRSVVDM DLDDTDDGDD NAPLFYQPGK" & _
            "RGFYTPRPGK NTEARLNCFR NIGRILGLCL LQNELCPITL NRHVIKVLLG RKVNWHDFAF" & _
            "FDPVMYESLR QLILASQSSD ADAVFSAMDL AFAVDLCKEE GGGQVELIPN GVNIPVTPQN" & _
            "VYEYVRKYAE HRMLVVAEQP LHAMRKGLLD VLPKNSLEDL TAEDFRLLVN GCGEVNVQML" & _
            "ISFTSFNDES GENAEKLLQF KRWFWSIVER MSMTERQDLV YFWTSSPSLP ASEEGFQPMP" & _
            "SITIRPPDDQ HLPTANTCIS RLYVPLYSSK QILKQKLLLA IKTKNFGFV"

        strProtein2 = ">104K_THEPA  104 kDa microneme-rhoptry antigen." & ControlChars.NewLine & _
         "MKFLILLFNI LCLFPVLAAD NHGVGPQGAS GVDPITFDIN SNQTGPAFLT AVEMAGVKYL" & _
         "QVQHGSNVNI HRLVEGNVVI WENASTPLYT GAIVTNNDGP YMAYVEVLGD PNLQFFIKSG" & _
         "DAWVTLSEHE YLAKLQEIRQ AVHIESVFSL NMAFQLENNK YEVETHAKNG ANMVTFIPRN" & _
         "GHICKMVYHK NVRIYKATGN DTVTSVVGFF RGLRLLLINV FSIDDNGMMS NRYFQHVDDK" & _
         "YVPISQKNYE TGIVKLKDYK HAYHPVDLDI KDIDYTMFHL ADATYHEPCF KIIPNTGFCI" & _
         "TKLFDGDQVL YESFNPLIHC INEVHIYDRN NGSIICLHLN YSPPSYKAYL VLKDTGWEAT" & _
         "THPLLEEKIE ELQDQRACEL DVNFISDKDL YVAALTNADL NYTMVTPRPH RDVIRVSDGS" & _
         "EVLWYYEGLD NFLVCAWIYV SDGVASLVHL RIKDRIPANN DIYVLKGDLY WTRITKIQFT" & _
         "QEIKRLVKKS KKKLAPITEE DSDKHDEPPE GPGASGLPPK APGDKEGSEG HKGPSKGSDS" & _
         "SKEGKKPGSG KKPGPAREHK PSKIPTLSKK PSGPKDPKHP RDPKEPRKSK SPRTASPTRR" & _
         "PSPKLPQLSK LPKSTSPRSP PPPTRPSSPE RPEGTKIIKT SKPPSPKPPF DPSFKEKFYD" & _
         "DYSKAASRSK ETKTTVVLDE SFESILKETL PETPGTPFTT PRPVPPKRPR TPESPFEPPK" & _
         "DPDSPSTSPS EFFTPPESKR TRFHETPADT PLPDVTAELF KEPDVTAETK SPDEAMKRPR" & _
         "SPSEYEDTSP GDYPSLPMKR HRLERLRLTT TEMETDPGRM AKDASGKPVK LKRSKSFDDL" & _
         "TTVELAPEPK ASRIVVDDEG TEADDEETHP PEERQKTEVR RRRPPKKPSK SPRPSKPKKP" & _
         "KKPDSAYIPS ILAILVVSLI VGIL"

        objSequence1 = NAligner.Sequence.Parse(strProtein1)
        objSequence2 = NAligner.Sequence.Parse(strProtein2)

        objAligner = NAligner.SmithWatermanGotoh.Align(objSequence1, objSequence2, objMatrix, opengap, extendgap)

        With udtAlignmentStats
            .AlignmentScore = objAligner.Score
            .SequenceLength = objAligner.Sequence1.Length
            .IdentityCount = objAligner.Identity
            .SimilarityCount = objAligner.Similarity
            .GapCount = objAligner.Gaps
        End With

        Console.WriteLine("NAligner (C#) results for protein vs. protein:")
        objDisplayer.DisplayResults(objAligner.Sequence1, objAligner.Sequence2, objAligner.MarkupLine, udtAlignmentStats, False)

    End Sub


    Private Function GetPeptideVsProtein1() As String
        Return "KGKIFRLDVANKTVRN"
    End Function

    Private Function GetPeptideVsProtein2() As String
        Return "MFRPTLTPLYIALMLSLNPPIAHADEVQKPSWDVNAPTNAPLKKVKIDVNEGTWMNLSVS" & _
                "PDGQHLVFDLLGDIYQIPVTGGEAKPLAQGISWQMQPVYSPNGKHIAFTSDADGGDNIWI" & _
                "MNADGSNPRTVTSETFRLLNSPAWSPDSQYLIGRKHFTASRSLGAGEVWLYHVAGGEGVK" & _
                "LTERPNDEKDLGEPAYSPDGRYIYFSQDDTPGKTFHYSKDSVNGIYKIKRYDTKTGNIEI" & _
                "LIEGTGGAIRPTPSPDGTKLAYIKRDDFQSSLYLLDLKSGETTKLFGDLDRDMQETWAIH" & _
                "GVYPSMAWTQDNKDIFFWAKGKINRLNVANKTVTNVPFSVKTQLDVQPSVRFKQDIDKDV" & _
                "FDVKMLRMAQVSPDGSKVAFEALGKIWLKSLPDGKMSRLTELGNDIGELYPQWSRDGKNI" & _
                "VFTTWNDQDQGAVQVISAKGGKAKQLTTEPGKYVEPTFAPNGELVVYRKTQGGNLTPRTW" & _
                "SQEPGIYKVDLKTKQNTKITAEGYQAQFGASAERIFFMNSGDDDTPQLASINLDGFDKRV" & _
                "HYSSKHATEFRVSPDGEQLAFAERFKVWVTPFAKHGETVEIGPNASNLPVTQLSVRAGES" & _
                "ISWNSKSNQLYWTLGPELYQTEVDSQYLKKDEQAKPSIINLGFTEKADVPRGTVAFVGGK" & _
                "VITMENDQVIDKGVVIVKDNHIVAVGDANTPIPKDAQVIDISSKSIMPGLFDAHAHGAQA" & _
                "DDEIVPQQNWALYSGLSLGVTTIHDPSNDTTEIFAASEQQKAGNIVGPRIFSTGTILYGA" & _
                "NAPGYTSHIDSVDDAKFHLERLKKVGAFSVKSYNQPRRNQRQQVIAAARELEMMVVPEGG" & _
                "SLLQHNLTMVADGHTTVEHSLPVASIYNDIKQFWGQTKVGYTPTLVVAYGGISGENYWYD" & _
                "KTDVWAHPRLSMYVPSDILQARSMRRPHVPESHYNHFNVAKVANEFNKLGIHPNIGAHGQ" & _
                "REGLAAHWEMWMFAQGGMSNMDVLKTATINPATTFGLDHQLGSIKTGKLADLIVIDGDPL" & _
                "ADIRVTDRNGKLFDAESMNQLNGNKQQRKPFFFEKI"

    End Function

End Module
