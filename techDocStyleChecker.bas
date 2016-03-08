Attribute VB_Name = "techDocStyleChecker"
Option Explicit
Public wordsToRemove(30)  As String
Public sentenceStartWordsToRemove(2) As String
Public wordSimplications(309, 2) As String
Sub checkStyle()

'Clean up previous runs
removeCommentsPrefixedWith "Style checker"

'Run all style checks
checkForWordsToRemove
checkForSentenceStartWordsToRemove
checkForSuggestedSimplifications

End Sub
Sub checkForWordsToRemove()

Dim currentDoc As Document, docRange As Range, searchRange As Range, rCounter As Integer, strComment As String

'Set-up the array for words to find
setupWordsToRemove


'Loop through array and execute a find for each one
Set currentDoc = ActiveDocument
Set docRange = currentDoc.Range
rCounter = 0

While rCounter < UBound(wordsToRemove)
    Set searchRange = docRange.Duplicate
    strComment = "Style checker" & Chr(13) + Chr(10) & "  Words that add no meaning" & Chr(13) + Chr(10) & "    '" & wordsToRemove(rCounter) & "'"
    
    Do
        With searchRange.Find
            .ClearFormatting
            .Text = wordsToRemove(rCounter)
            .MatchCase = False
            .MatchWholeWord = True
            .Execute
        End With
    
        If searchRange.Find.Found Then Call currentDoc.Comments.Add(searchRange, strComment)
        
    Loop Until Not searchRange.Find.Found
    
    Set searchRange = Nothing
    rCounter = rCounter + 1
Wend

Set currentDoc = Nothing
Set docRange = Nothing

End Sub

Sub checkForSentenceStartWordsToRemove()

Dim currentDoc As Document, docRange As Range, searchRange As Range, rCounter As Integer, strComment As String

'Set-up the array for words to find
setupSentenceStartWordsToRemove

'Loop through array and execute a find for each one
Set currentDoc = ActiveDocument
Set docRange = currentDoc.Range
rCounter = 0

While rCounter < UBound(sentenceStartWordsToRemove)
    Set searchRange = docRange.Duplicate
    strComment = "Style checker" & Chr(13) + Chr(10) & "  Sentence starters" & Chr(13) + Chr(10) & "    '" & sentenceStartWordsToRemove(rCounter) & "'"
    
    Do
        With searchRange.Find
            .ClearFormatting
            .Text = sentenceStartWordsToRemove(rCounter)
            .MatchCase = True ' Make case matching true here because the start of the sentence will be detected by an uppercase 1st character
            .MatchWholeWord = True
            .Execute
        End With
    
        If searchRange.Find.Found Then Call currentDoc.Comments.Add(searchRange, strComment)
        
    Loop Until Not searchRange.Find.Found
    
    Set searchRange = Nothing
    rCounter = rCounter + 1
Wend

Set currentDoc = Nothing
Set docRange = Nothing

End Sub

Sub checkForSuggestedSimplifications()

Dim currentDoc As Document, docRange As Range, searchRange As Range, rCounter As Integer, strComment As String

'Set-up the array for words to find
setupWordSimplifications

'Loop through array and execute a find for each one
Set currentDoc = ActiveDocument
Set docRange = currentDoc.Range
rCounter = 0

While rCounter < UBound(wordSimplications)
    Set searchRange = docRange.Duplicate
    strComment = "Style checker" & Chr(13) & Chr(10) & "  Simplification suggestion" & Chr(13) + Chr(10) & "    '" & wordSimplications(rCounter, 0) & _
                    Chr(13) & Chr(10) & "' replace with:" & _
                    Chr(13) & Chr(10) & "    " & wordSimplications(rCounter, 1)
    
    Do
        With searchRange.Find
            .ClearFormatting
            .Text = wordSimplications(rCounter, 0)
            .MatchCase = False
            .MatchWholeWord = True
            .Execute
        End With
    
        If searchRange.Find.Found Then Call currentDoc.Comments.Add(searchRange, strComment)
        
    Loop Until Not searchRange.Find.Found
    
    Set searchRange = Nothing
    rCounter = rCounter + 1
Wend

Set currentDoc = Nothing
Set docRange = Nothing


End Sub
Sub removeCommentsPrefixedWith(commentPrefix As String)

Dim docComments As Comments, commentCounter As Integer, docComment As Comment, commentPrefixLength As Integer

Set docComments = ActiveDocument.Comments

commentPrefixLength = Len(commentPrefix)

'If no length, then somehting is wrong
If commentPrefixLength < 1 Then Exit Sub

'Loop through each comment and check the comment prefix - if it matches the supplied prefix, then remove the comment
commentCounter = 1
While commentCounter <= docComments.Count

    Set docComment = docComments.Item(commentCounter)
    'Check that the comment is longer than the prefix, if not, it can't be a comment we want to delete
    If Len(docComment.Range.Text) > commentPrefixLength Then
        'Check if the initial text of the comment is the same as the prefix
        If Left(docComment.Range.Text, commentPrefixLength) = commentPrefix Then
            docComment.Delete
            'Because comment count is recalculated at the point of deleting a comment, need to decrement the counter 1 to prevent comment sbeing skipped
            commentCounter = commentCounter - 1
        End If
    End If
    
    Set docComment = Nothing
    
    commentCounter = commentCounter + 1
Wend

Set docComments = Nothing

End Sub

Sub setupWordsToRemove()

wordsToRemove(0) = "easy"
wordsToRemove(1) = "easily"
wordsToRemove(2) = "simply"
wordsToRemove(3) = "simple"
wordsToRemove(4) = "obviously"
wordsToRemove(5) = "just"
wordsToRemove(6) = "basically"
wordsToRemove(7) = "procure"
wordsToRemove(8) = "sexy"
wordsToRemove(9) = "insane"
wordsToRemove(10) = "clearly"
wordsToRemove(11) = "be advised"
wordsToRemove(12) = "in the process of"
wordsToRemove(13) = "procure"
wordsToRemove(14) = "take action to"
wordsToRemove(15) = "the month of"
wordsToRemove(16) = "the user of"
wordsToRemove(17) = "of course"
wordsToRemove(18) = "everyone knows"
wordsToRemove(19) = "inter alia"
wordsToRemove(20) = "it is"
wordsToRemove(21) = "literally"
wordsToRemove(22) = "overall"
wordsToRemove(23) = "put simply"
wordsToRemove(24) = "take action to"
wordsToRemove(25) = "the month of"
wordsToRemove(26) = "the use of"
wordsToRemove(27) = "there are"
wordsToRemove(28) = "there is"
wordsToRemove(29) = "very"

End Sub

Sub setupSentenceStartWordsToRemove()

sentenceStartWordsToRemove(0) = "So"
sentenceStartWordsToRemove(1) = "However"


End Sub

Sub setupWordSimplifications()

wordSimplications(0, 0) = "a number of"
wordSimplications(0, 1) = "'many' or 'some'"
wordSimplications(1, 0) = "abundance"
wordSimplications(1, 1) = "'enough' or 'plenty'"
wordSimplications(2, 0) = "accede to"
wordSimplications(2, 1) = "'agree to' or 'allow'"
wordSimplications(3, 0) = "accelerate"
wordSimplications(3, 1) = "'speed up'"
wordSimplications(4, 0) = "accentuate"
wordSimplications(4, 1) = "'stress'"
wordSimplications(5, 0) = "accompany"
wordSimplications(5, 1) = "'go with' or 'with'"
wordSimplications(6, 0) = "accomplish"
wordSimplications(6, 1) = "'carry out' or 'do'"
wordSimplications(7, 0) = "accorded"
wordSimplications(7, 1) = "'given'"
wordSimplications(8, 0) = "accordingly"
wordSimplications(8, 1) = "'so'"
wordSimplications(9, 0) = "accrue"
wordSimplications(9, 1) = "'add' or 'gain'"
wordSimplications(10, 0) = "accurate"
wordSimplications(10, 1) = "'correct' or 'exact' or 'right'"
wordSimplications(11, 0) = "acquiesce"
wordSimplications(11, 1) = "'agree'"
wordSimplications(12, 0) = "acquire"
wordSimplications(12, 1) = "'get'"
wordSimplications(13, 0) = "additional"
wordSimplications(13, 1) = "'added' or 'extra' or 'more' or 'other'"
wordSimplications(14, 0) = "addressees"
wordSimplications(14, 1) = "'you'"
wordSimplications(15, 0) = "addressees are requested"
wordSimplications(15, 1) = "'please'"
wordSimplications(16, 0) = "adjacent to"
wordSimplications(16, 1) = "'next to'"
wordSimplications(17, 0) = "adjustment"
wordSimplications(17, 1) = "'change'"
wordSimplications(18, 0) = "admissible"
wordSimplications(18, 1) = "'accepted' or 'allowed'"
wordSimplications(19, 0) = "advantageous"
wordSimplications(19, 1) = "'helpful'"
wordSimplications(20, 0) = "adversely impact"
wordSimplications(20, 1) = "'hurt'"
wordSimplications(21, 0) = "adversely impact on"
wordSimplications(21, 1) = "'hurt' or 'set back'"
wordSimplications(22, 0) = "advise"
wordSimplications(22, 1) = "'recommend' or 'tell'"
wordSimplications(23, 0) = "afford an opportunity"
wordSimplications(23, 1) = "'allow' or 'let'"
wordSimplications(24, 0) = "aforementioned"
wordSimplications(24, 1) = "'remove'"
wordSimplications(25, 0) = "aggregate"
wordSimplications(25, 1) = "'add' or 'total'"
wordSimplications(26, 0) = "aircraft"
wordSimplications(26, 1) = "'plane'"
wordSimplications(27, 0) = "all of"
wordSimplications(27, 1) = "'all'"
wordSimplications(28, 0) = "alleviate"
wordSimplications(28, 1) = "'ease' or 'reduce'"
wordSimplications(29, 0) = "allocate"
wordSimplications(29, 1) = "'divide'"
wordSimplications(30, 0) = "along the lines of"
wordSimplications(30, 1) = "'as in' or 'like'"
wordSimplications(31, 0) = "already existing"
wordSimplications(31, 1) = "'existing'"
wordSimplications(32, 0) = "alternatively"
wordSimplications(32, 1) = "'or'"
wordSimplications(33, 0) = "ameliorate"
wordSimplications(33, 1) = "'help' or 'improve'"
wordSimplications(34, 0) = "and/or"
wordSimplications(34, 1) = "' … or … or both'"
wordSimplications(35, 0) = "anticipate"
wordSimplications(35, 1) = "'expect'"
wordSimplications(36, 0) = "apparent"
wordSimplications(36, 1) = "'clear' or 'plain'"
wordSimplications(37, 0) = "appreciable"
wordSimplications(37, 1) = "'many'"
wordSimplications(38, 0) = "appropriate"
wordSimplications(38, 1) = "'proper' or 'right'"
wordSimplications(39, 0) = "approximate"
wordSimplications(39, 1) = "'about'"
wordSimplications(40, 0) = "arrive onboard"
wordSimplications(40, 1) = "'arrive'"
wordSimplications(41, 0) = "as a means of"
wordSimplications(41, 1) = "'to'"
wordSimplications(42, 0) = "as of yet"
wordSimplications(42, 1) = "'yet'"
wordSimplications(43, 0) = "as prescribed by"
wordSimplications(43, 1) = "'in'"
wordSimplications(44, 0) = "as to"
wordSimplications(44, 1) = "'about' or 'on'"
wordSimplications(45, 0) = "as yet"
wordSimplications(45, 1) = "'yet'"
wordSimplications(46, 0) = "ascertain"
wordSimplications(46, 1) = "'find out' or 'learn'"
wordSimplications(47, 0) = "assist"
wordSimplications(47, 1) = "'aid' or 'help'"
wordSimplications(48, 0) = "assistance"
wordSimplications(48, 1) = "'aid' or 'help'"
wordSimplications(49, 0) = "at the present time"
wordSimplications(49, 1) = "'at present'"
wordSimplications(50, 0) = "at this time"
wordSimplications(50, 1) = "'now'"
wordSimplications(51, 0) = "attain"
wordSimplications(51, 1) = "'meet'"
wordSimplications(52, 0) = "attempt"
wordSimplications(52, 1) = "'try'"
wordSimplications(53, 0) = "attributable to"
wordSimplications(53, 1) = "'because'"
wordSimplications(54, 0) = "authorise"
wordSimplications(54, 1) = "'allow' or 'let'"
wordSimplications(55, 0) = "authorize"
wordSimplications(55, 1) = "'allow' or 'let'"
wordSimplications(56, 0) = "because of the fact that"
wordSimplications(56, 1) = "'because'"
wordSimplications(57, 0) = "belated"
wordSimplications(57, 1) = "'late'"
wordSimplications(58, 0) = "benefit"
wordSimplications(58, 1) = "'help'"
wordSimplications(59, 0) = "benefit from"
wordSimplications(59, 1) = "'enjoy'"
wordSimplications(60, 0) = "bestow"
wordSimplications(60, 1) = "'award' or 'give'"
wordSimplications(61, 0) = "by means of"
wordSimplications(61, 1) = "'by' or 'with'"
wordSimplications(62, 0) = "by virtue of"
wordSimplications(62, 1) = "'by' or 'under'"
wordSimplications(63, 0) = "capability"
wordSimplications(63, 1) = "'ability'"
wordSimplications(64, 0) = "caveat"
wordSimplications(64, 1) = "'warning'"
wordSimplications(65, 0) = "cease"
wordSimplications(65, 1) = "'stop'"
wordSimplications(66, 0) = "close proximity"
wordSimplications(66, 1) = "'near'"
wordSimplications(67, 0) = "combat environment"
wordSimplications(67, 1) = "'combat'"
wordSimplications(68, 0) = "combined"
wordSimplications(68, 1) = "'joint'"
wordSimplications(69, 0) = "commence"
wordSimplications(69, 1) = "'begin' or 'start'"
wordSimplications(70, 0) = "comply with"
wordSimplications(70, 1) = "'follow'"
wordSimplications(71, 0) = "component"
wordSimplications(71, 1) = "'part'"
wordSimplications(72, 0) = "comprise"
wordSimplications(72, 1) = "'form' or 'include' or 'make up'"
wordSimplications(73, 0) = "concerning"
wordSimplications(73, 1) = "'about' or 'on'"
wordSimplications(74, 0) = "consequently"
wordSimplications(74, 1) = "'so'"
wordSimplications(75, 0) = "consolidate"
wordSimplications(75, 1) = "'combine' or 'join' or 'merge'"
wordSimplications(76, 0) = "constitutes"
wordSimplications(76, 1) = "'forms' or 'is' or 'makes up'"
wordSimplications(77, 0) = "contains"
wordSimplications(77, 1) = "'has'"
wordSimplications(78, 0) = "convene"
wordSimplications(78, 1) = "'meet'"
wordSimplications(79, 0) = "currently"
wordSimplications(79, 1) = "'now'"
wordSimplications(80, 0) = "deem"
wordSimplications(80, 1) = "'believe' or 'consider' or 'think'"
wordSimplications(81, 0) = "delete"
wordSimplications(81, 1) = "'cut' or 'drop'"
wordSimplications(82, 0) = "demonstrate"
wordSimplications(82, 1) = "'prove' or 'show'"
wordSimplications(83, 0) = "depart"
wordSimplications(83, 1) = "'go' or 'leave'"
wordSimplications(84, 0) = "designate"
wordSimplications(84, 1) = "'appoint' or 'choose' or 'name'"
wordSimplications(85, 0) = "desire"
wordSimplications(85, 1) = "'want' or 'wish'"
wordSimplications(86, 0) = "determine"
wordSimplications(86, 1) = "'decide' or 'figure' or 'find'"
wordSimplications(87, 0) = "disclose"
wordSimplications(87, 1) = "'show'"
wordSimplications(88, 0) = "discontinue"
wordSimplications(88, 1) = "'drop' or 'stop'"
wordSimplications(89, 0) = "disseminate"
wordSimplications(89, 1) = "'give' or 'issue' or 'pass' or 'send'"
wordSimplications(90, 0) = "due to the fact that"
wordSimplications(90, 1) = "'because' or 'due to' or 'since'"
wordSimplications(91, 0) = "during the period"
wordSimplications(91, 1) = "'during'"
wordSimplications(92, 0) = "e.g."
wordSimplications(92, 1) = "'for example' or 'such as'"
wordSimplications(93, 0) = "each and every"
wordSimplications(93, 1) = "'each'"
wordSimplications(94, 0) = "economical"
wordSimplications(94, 1) = "'cheap'"
wordSimplications(95, 0) = "effect"
wordSimplications(95, 1) = "'choose' or 'pick'"
wordSimplications(96, 0) = "effect modifications"
wordSimplications(96, 1) = "'make changes'"
wordSimplications(97, 0) = "elect"
wordSimplications(97, 1) = "'choose'"
wordSimplications(98, 0) = "eliminate"
wordSimplications(98, 1) = "'cut' or 'drop' or 'end' or 'stop'"
wordSimplications(99, 0) = "elucidate"
wordSimplications(99, 1) = "'explain'"
wordSimplications(100, 0) = "employ"
wordSimplications(100, 1) = "'use'"
wordSimplications(101, 0) = "encounter"
wordSimplications(101, 1) = "'meet'"
wordSimplications(102, 0) = "endeavor"
wordSimplications(102, 1) = "'try'"
wordSimplications(103, 0) = "ensure"
wordSimplications(103, 1) = "'make sure'"
wordSimplications(104, 0) = "enumerate"
wordSimplications(104, 1) = "'count'"
wordSimplications(105, 0) = "equipments"
wordSimplications(105, 1) = "'equipment'"
wordSimplications(106, 0) = "equitable"
wordSimplications(106, 1) = "'fair'"
wordSimplications(107, 0) = "equivalent"
wordSimplications(107, 1) = "'equal'"
wordSimplications(108, 0) = "establish"
wordSimplications(108, 1) = "'set up' or 'prove' or 'show'"
wordSimplications(109, 0) = "evaluate"
wordSimplications(109, 1) = "'check' or 'test'"
wordSimplications(110, 0) = "evidenced"
wordSimplications(110, 1) = "'showed'"
wordSimplications(111, 0) = "evident"
wordSimplications(111, 1) = "'clear'"
wordSimplications(112, 0) = "exclusively"
wordSimplications(112, 1) = "'only'"
wordSimplications(113, 0) = "exhibit"
wordSimplications(113, 1) = "'show'"
wordSimplications(114, 0) = "expedite"
wordSimplications(114, 1) = "'hasten' or 'hurry' or 'speed up'"
wordSimplications(115, 0) = "expeditious"
wordSimplications(115, 1) = "'fast' or 'quick'"
wordSimplications(116, 0) = "expend"
wordSimplications(116, 1) = "'spend'"
wordSimplications(117, 0) = "expertise"
wordSimplications(117, 1) = "'ability'"
wordSimplications(118, 0) = "expiration"
wordSimplications(118, 1) = "'end'"
wordSimplications(119, 0) = "facilitate"
wordSimplications(119, 1) = "'ease' or 'help'"
wordSimplications(120, 0) = "factual evidence"
wordSimplications(120, 1) = "'evidence' or 'facts'"
wordSimplications(121, 0) = "failed to"
wordSimplications(121, 1) = "'didnâ€™t'"
wordSimplications(122, 0) = "feasible"
wordSimplications(122, 1) = "'can be done' or 'workable'"
wordSimplications(123, 0) = "females"
wordSimplications(123, 1) = "'women'"
wordSimplications(124, 0) = "finalise"
wordSimplications(124, 1) = "'complete' or 'finish'"
wordSimplications(125, 0) = "finalize"
wordSimplications(125, 1) = "'complete' or 'finish'"
wordSimplications(126, 0) = "first and foremost"
wordSimplications(126, 1) = "'first'"
wordSimplications(127, 0) = "for a period of"
wordSimplications(127, 1) = "'for'"
wordSimplications(128, 0) = "for the purpose of"
wordSimplications(128, 1) = "'to'"
wordSimplications(129, 0) = "forfeit"
wordSimplications(129, 1) = "'give up' or 'lose'"
wordSimplications(130, 0) = "formulate"
wordSimplications(130, 1) = "'plan'"
wordSimplications(131, 0) = "forward"
wordSimplications(131, 1) = "'send'"
wordSimplications(132, 0) = "frequently"
wordSimplications(132, 1) = "'often'"
wordSimplications(133, 0) = "function"
wordSimplications(133, 1) = "'act' or 'role' or 'work'"
wordSimplications(134, 0) = "furnish"
wordSimplications(134, 1) = "'give' or 'send'"
wordSimplications(135, 0) = "has a requirement for"
wordSimplications(135, 1) = "'needs'"
wordSimplications(136, 0) = "herein"
wordSimplications(136, 1) = "'here'"
wordSimplications(137, 0) = "heretofore"
wordSimplications(137, 1) = "'until now'"
wordSimplications(138, 0) = "herewith"
wordSimplications(138, 1) = "'here' or 'below'"
wordSimplications(139, 0) = "honest truth"
wordSimplications(139, 1) = "'truth'"
wordSimplications(140, 0) = "however"
wordSimplications(140, 1) = "'but' or 'yet'"
wordSimplications(141, 0) = "i.e."
wordSimplications(141, 1) = "'as in'"
wordSimplications(142, 0) = "identical"
wordSimplications(142, 1) = "'same'"
wordSimplications(143, 0) = "identify"
wordSimplications(143, 1) = "'find' or 'name' or 'show'"
wordSimplications(144, 0) = "if and when"
wordSimplications(144, 1) = "'if' or 'when'"
wordSimplications(145, 0) = "immediately"
wordSimplications(145, 1) = "'at once'"
wordSimplications(146, 0) = "impacted"
wordSimplications(146, 1) = "'affected' or 'changed' or 'harmed'"
wordSimplications(147, 0) = "implement"
wordSimplications(147, 1) = "'carry out' or 'install' or 'put in place' or 'tool' or 'start'"
wordSimplications(148, 0) = "in a timely manner"
wordSimplications(148, 1) = "'on time' or 'promptly'"
wordSimplications(149, 0) = "in accordance with"
wordSimplications(149, 1) = "'by' or 'under' or 'following' or 'per'"
wordSimplications(150, 0) = "in addition"
wordSimplications(150, 1) = "'also' or 'besides' or 'too'"
wordSimplications(151, 0) = "in all likelihood"
wordSimplications(151, 1) = "'probably'"
wordSimplications(152, 0) = "in an effort to"
wordSimplications(152, 1) = "'to'"
wordSimplications(153, 0) = "in between"
wordSimplications(153, 1) = "'between'"
wordSimplications(154, 0) = "in excess of"
wordSimplications(154, 1) = "'more than'"
wordSimplications(155, 0) = "in lieu of"
wordSimplications(155, 1) = "'instead'"
wordSimplications(156, 0) = "in light of the fact that"
wordSimplications(156, 1) = "'because'"
wordSimplications(157, 0) = "in many cases"
wordSimplications(157, 1) = "'often'"
wordSimplications(158, 0) = "in order that"
wordSimplications(158, 1) = "'for' or 'so'"
wordSimplications(159, 0) = "in order to"
wordSimplications(159, 1) = "'to'"
wordSimplications(160, 0) = "in regard to"
wordSimplications(160, 1) = "'about' or 'concerning' or 'on'"
wordSimplications(161, 0) = "in relation to"
wordSimplications(161, 1) = "'about' or 'with' or 'to'"
wordSimplications(162, 0) = "in some instances"
wordSimplications(162, 1) = "'sometimes'"
wordSimplications(163, 0) = "in terms of"
wordSimplications(163, 1) = "'as' or 'for' or 'with'"
wordSimplications(164, 0) = "in the amount of"
wordSimplications(164, 1) = "'for'"
wordSimplications(165, 0) = "in the event of"
wordSimplications(165, 1) = "'if'"
wordSimplications(166, 0) = "in the near future"
wordSimplications(166, 1) = "'soon' or 'shortly'"
wordSimplications(167, 0) = "in view of"
wordSimplications(167, 1) = "'since'"
wordSimplications(168, 0) = "in view of the above"
wordSimplications(168, 1) = "'so'"
wordSimplications(169, 0) = "inasmuch as"
wordSimplications(169, 1) = "'since'"
wordSimplications(170, 0) = "inception"
wordSimplications(170, 1) = "'start'"
wordSimplications(171, 0) = "incumbent upon"
wordSimplications(171, 1) = "'must'"
wordSimplications(172, 0) = "indicate"
wordSimplications(172, 1) = "'show' or 'say' or 'state' or 'write down'"
wordSimplications(173, 0) = "indication"
wordSimplications(173, 1) = "'sign'"
wordSimplications(174, 0) = "initial"
wordSimplications(174, 1) = "'first'"
wordSimplications(175, 0) = "initiate"
wordSimplications(175, 1) = "'start'"
wordSimplications(176, 0) = "interface"
wordSimplications(176, 1) = "'meet' or 'work with'"
wordSimplications(177, 0) = "interpose no objection"
wordSimplications(177, 1) = "'donâ€™t object'"
wordSimplications(178, 0) = "is applicable to"
wordSimplications(178, 1) = "'applies to'"
wordSimplications(179, 0) = "is authorised to"
wordSimplications(179, 1) = "'may'"
wordSimplications(180, 0) = "is authorized to"
wordSimplications(180, 1) = "'may'"
wordSimplications(181, 0) = "is in consonance with"
wordSimplications(181, 1) = "'agrees with' or 'follows'"
wordSimplications(182, 0) = "is responsible for"
wordSimplications(182, 1) = "'handles'"
wordSimplications(183, 0) = "it appears"
wordSimplications(183, 1) = "'seems'"
wordSimplications(184, 0) = "it is essential"
wordSimplications(184, 1) = "'must' or 'need to'"
wordSimplications(185, 0) = "it is requested"
wordSimplications(185, 1) = "'please'"
wordSimplications(186, 0) = "liaison"
wordSimplications(186, 1) = "'discussion'"
wordSimplications(187, 0) = "limited number"
wordSimplications(187, 1) = "'limits'"
wordSimplications(188, 0) = "magnitude"
wordSimplications(188, 1) = "'size'"
wordSimplications(189, 0) = "maintain"
wordSimplications(189, 1) = "'support' or 'keep'"
wordSimplications(190, 0) = "maximum"
wordSimplications(190, 1) = "'greatest' or 'largest' or 'most'"
wordSimplications(191, 0) = "methodology"
wordSimplications(191, 1) = "'method'"
wordSimplications(192, 0) = "minimise"
wordSimplications(192, 1) = "'cut' or 'decrease'"
wordSimplications(193, 0) = "minimize"
wordSimplications(193, 1) = "'cut' or 'decrease'"
wordSimplications(194, 0) = "minimum"
wordSimplications(194, 1) = "'least' or 'small' or 'smallest'"
wordSimplications(195, 0) = "modify"
wordSimplications(195, 1) = "'change'"
wordSimplications(196, 0) = "monitor"
wordSimplications(196, 1) = "'check' or 'track' or 'watch'"
wordSimplications(197, 0) = "multiple"
wordSimplications(197, 1) = "'many'"
wordSimplications(198, 0) = "necessitate"
wordSimplications(198, 1) = "'cause' or 'need'"
wordSimplications(199, 0) = "nevertheless"
wordSimplications(199, 1) = "'besides' or 'even so' or 'still'"
wordSimplications(200, 0) = "not certain"
wordSimplications(200, 1) = "'uncertain'"
wordSimplications(201, 0) = "not later than"
wordSimplications(201, 1) = "'by' or 'before'"
wordSimplications(202, 0) = "not many"
wordSimplications(202, 1) = "'few'"
wordSimplications(203, 0) = "not often"
wordSimplications(203, 1) = "'rarely'"
wordSimplications(204, 0) = "not unless"
wordSimplications(204, 1) = "'only if'"
wordSimplications(205, 0) = "not unlike"
wordSimplications(205, 1) = "'alike' or 'similar'"
wordSimplications(206, 0) = "notify"
wordSimplications(206, 1) = "'let know' or 'tell'"
wordSimplications(207, 0) = "notwithstanding"
wordSimplications(207, 1) = "'despite' or 'in spite of' or 'still'"
wordSimplications(208, 0) = "null and void"
wordSimplications(208, 1) = "'null' or 'void'"
wordSimplications(209, 0) = "numerous"
wordSimplications(209, 1) = "'many'"
wordSimplications(210, 0) = "objective"
wordSimplications(210, 1) = "'aim' or 'goal'"
wordSimplications(211, 0) = "obligate"
wordSimplications(211, 1) = "'bind' or 'compel'"
wordSimplications(212, 0) = "observe"
wordSimplications(212, 1) = "'see'"
wordSimplications(213, 0) = "obtain"
wordSimplications(213, 1) = "'get'"
wordSimplications(214, 0) = "on the contrary"
wordSimplications(214, 1) = "'but' or 'so'"
wordSimplications(215, 0) = "on the other hand"
wordSimplications(215, 1) = "'but' or 'so'"
wordSimplications(216, 0) = "one particular"
wordSimplications(216, 1) = "'one'"
wordSimplications(217, 0) = "operate"
wordSimplications(217, 1) = "'run' or 'use' or 'work'"
wordSimplications(218, 0) = "optimum"
wordSimplications(218, 1) = "'best' or 'greatest' or 'most'"
wordSimplications(219, 0) = "option"
wordSimplications(219, 1) = "'choice'"
wordSimplications(220, 0) = "owing to the fact that"
wordSimplications(220, 1) = "'because' or 'since'"
wordSimplications(221, 0) = "parameters"
wordSimplications(221, 1) = "'limits'"
wordSimplications(222, 0) = "participate"
wordSimplications(222, 1) = "'take part'"
wordSimplications(223, 0) = "particulars"
wordSimplications(223, 1) = "'details'"
wordSimplications(224, 0) = "pass away"
wordSimplications(224, 1) = "'die'"
wordSimplications(225, 0) = "perform"
wordSimplications(225, 1) = "'do'"
wordSimplications(226, 0) = "permit"
wordSimplications(226, 1) = "'let'"
wordSimplications(227, 0) = "pertaining to"
wordSimplications(227, 1) = "'about' or 'of' or 'on'"
wordSimplications(228, 0) = "point in time"
wordSimplications(228, 1) = "'moment' or 'now' or 'point' or 'time'"
wordSimplications(229, 0) = "portion"
wordSimplications(229, 1) = "'part'"
wordSimplications(230, 0) = "possess"
wordSimplications(230, 1) = "'have' or 'own'"
wordSimplications(231, 0) = "practicable"
wordSimplications(231, 1) = "'practical'"
wordSimplications(232, 0) = "preclude"
wordSimplications(232, 1) = "'prevent'"
wordSimplications(233, 0) = "previous"
wordSimplications(233, 1) = "'earlier'"
wordSimplications(234, 0) = "previously"
wordSimplications(234, 1) = "'before'"
wordSimplications(235, 0) = "prior to"
wordSimplications(235, 1) = "'before'"
wordSimplications(236, 0) = "prioritise"
wordSimplications(236, 1) = "'focus on' or 'rank'"
wordSimplications(237, 0) = "prioritize"
wordSimplications(237, 1) = "'focus on' or 'rank'"
wordSimplications(238, 0) = "proceed"
wordSimplications(238, 1) = "'do' or 'go ahead' or 'try'"
wordSimplications(239, 0) = "procure"
wordSimplications(239, 1) = "'buy' or 'get'"
wordSimplications(240, 0) = "proficiency"
wordSimplications(240, 1) = "'skill'"
wordSimplications(241, 0) = "promulgate"
wordSimplications(241, 1) = "'issue' or 'publish'"
wordSimplications(242, 0) = "provide"
wordSimplications(242, 1) = "'give' or 'offer' or 'say'"
wordSimplications(243, 0) = "provided that"
wordSimplications(243, 1) = "'if'"
wordSimplications(244, 0) = "provides guidance for"
wordSimplications(244, 1) = "'guides'"
wordSimplications(245, 0) = "purchase"
wordSimplications(245, 1) = "'buy' or 'sale'"
wordSimplications(246, 0) = "pursuant to"
wordSimplications(246, 1) = "'by' or 'following' or 'per' or 'under'"
wordSimplications(247, 0) = "readily apparent"
wordSimplications(247, 1) = "'clear'"
wordSimplications(248, 0) = "refer back"
wordSimplications(248, 1) = "'refer'"
wordSimplications(249, 0) = "reflect"
wordSimplications(249, 1) = "'say' or 'show'"
wordSimplications(250, 0) = "regarding"
wordSimplications(250, 1) = "'about' or 'of' or 'on'"
wordSimplications(251, 0) = "relative to"
wordSimplications(251, 1) = "'about' or 'on'"
wordSimplications(252, 0) = "relocate"
wordSimplications(252, 1) = "'move'"
wordSimplications(253, 0) = "remain"
wordSimplications(253, 1) = "'stay'"
wordSimplications(254, 0) = "remainder"
wordSimplications(254, 1) = "'rest'"
wordSimplications(255, 0) = "remuneration"
wordSimplications(255, 1) = "'pay' or 'payment'"
wordSimplications(256, 0) = "render"
wordSimplications(256, 1) = "'give' or 'make'"
wordSimplications(257, 0) = "represents"
wordSimplications(257, 1) = "'is'"
wordSimplications(258, 0) = "require"
wordSimplications(258, 1) = "'must' or 'need'"
wordSimplications(259, 0) = "requirement"
wordSimplications(259, 1) = "'need' or 'rule'"
wordSimplications(260, 0) = "reside"
wordSimplications(260, 1) = "'live'"
wordSimplications(261, 0) = "residence"
wordSimplications(261, 1) = "'house'"
wordSimplications(262, 0) = "retain"
wordSimplications(262, 1) = "'keep'"
wordSimplications(263, 0) = "satisfy"
wordSimplications(263, 1) = "'meet' or 'please'"
wordSimplications(264, 0) = "selection"
wordSimplications(264, 1) = "'choice'"
wordSimplications(265, 0) = "set forth in"
wordSimplications(265, 1) = "'in'"
wordSimplications(266, 0) = "shall"
wordSimplications(266, 1) = "'must' or 'will'"
wordSimplications(267, 0) = "should you wish"
wordSimplications(267, 1) = "'if you want'"
wordSimplications(268, 0) = "similar to"
wordSimplications(268, 1) = "'like'"
wordSimplications(269, 0) = "solicit"
wordSimplications(269, 1) = "'ask for' or 'request'"
wordSimplications(270, 0) = "span across"
wordSimplications(270, 1) = "'cross' or 'span'"
wordSimplications(271, 0) = "state-of-the-art"
wordSimplications(271, 1) = "'latest'"
wordSimplications(272, 0) = "strategise"
wordSimplications(272, 1) = "'plan'"
wordSimplications(273, 0) = "strategize"
wordSimplications(273, 1) = "'plan'"
wordSimplications(274, 0) = "submit"
wordSimplications(274, 1) = "'give' or 'send'"
wordSimplications(275, 0) = "subsequent"
wordSimplications(275, 1) = "'after' or 'later' or 'next' or 'then'"
wordSimplications(276, 0) = "subsequently"
wordSimplications(276, 1) = "'after' or 'later' or 'then'"
wordSimplications(277, 0) = "substantial"
wordSimplications(277, 1) = "'large' or 'much'"
wordSimplications(278, 0) = "successfully complete"
wordSimplications(278, 1) = "'complete' or 'pass'"
wordSimplications(279, 0) = "sufficient"
wordSimplications(279, 1) = "'enough'"
wordSimplications(280, 0) = "terminate"
wordSimplications(280, 1) = "'end' or 'stop'"
wordSimplications(281, 0) = "the undersigned"
wordSimplications(281, 1) = "'I'"
wordSimplications(282, 0) = "therefore"
wordSimplications(282, 1) = "'so' or 'thus'"
wordSimplications(283, 0) = "therein"
wordSimplications(283, 1) = "'there'"
wordSimplications(284, 0) = "thereof"
wordSimplications(284, 1) = "'its' or 'their'"
wordSimplications(285, 0) = "this day and age"
wordSimplications(285, 1) = "'today'"
wordSimplications(286, 0) = "time period"
wordSimplications(286, 1) = "'period' or 'time'"
wordSimplications(287, 0) = "timely"
wordSimplications(287, 1) = "'prompt'"
wordSimplications(288, 0) = "took advantage of"
wordSimplications(288, 1) = "'preyed on'"
wordSimplications(289, 0) = "transmit"
wordSimplications(289, 1) = "'send'"
wordSimplications(290, 0) = "transpire"
wordSimplications(290, 1) = "'happen'"
wordSimplications(291, 0) = "under the provisions of"
wordSimplications(291, 1) = "'under'"
wordSimplications(292, 0) = "until such time as"
wordSimplications(292, 1) = "'until'"
wordSimplications(293, 0) = "utilisation"
wordSimplications(293, 1) = "'use'"
wordSimplications(294, 0) = "utilise"
wordSimplications(294, 1) = "'use'"
wordSimplications(295, 0) = "utilization"
wordSimplications(295, 1) = "'use'"
wordSimplications(296, 0) = "utilize"
wordSimplications(296, 1) = "'use'"
wordSimplications(297, 0) = "validate"
wordSimplications(297, 1) = "'confirm'"
wordSimplications(298, 0) = "various different"
wordSimplications(298, 1) = "'different' or 'various'"
wordSimplications(299, 0) = "viable"
wordSimplications(299, 1) = "'practical' or 'workable'"
wordSimplications(300, 0) = "vice"
wordSimplications(300, 1) = "'instead of' or 'versus'"
wordSimplications(301, 0) = "warrant"
wordSimplications(301, 1) = "'call for' or 'permit'"
wordSimplications(302, 0) = "whereas"
wordSimplications(302, 1) = "'because' or 'since'"
wordSimplications(303, 0) = "whether or not"
wordSimplications(303, 1) = "'whether'"
wordSimplications(304, 0) = "with reference to"
wordSimplications(304, 1) = "'about'"
wordSimplications(305, 0) = "with respect to"
wordSimplications(305, 1) = "'about' or 'on'"
wordSimplications(306, 0) = "with the exception of"
wordSimplications(306, 1) = "'except for'"
wordSimplications(307, 0) = "witnessed"
wordSimplications(307, 1) = "'saw' or 'seen'"
wordSimplications(308, 0) = "your office"
wordSimplications(308, 1) = "'you'"


End Sub
