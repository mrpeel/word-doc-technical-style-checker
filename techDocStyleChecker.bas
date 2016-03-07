Attribute VB_Name = "techDocStyleChecker"
Option Explicit
Public wordsToRemove(31)  As String
Public sentenceStartWordsToRemove(2) As String
Public wordSimplications(310, 2) As String
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
wordsToRemove(29) = "type"
wordsToRemove(30) = "very"

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
wordSimplications(14, 0) = "address"
wordSimplications(14, 1) = "'discuss'"
wordSimplications(15, 0) = "addressees"
wordSimplications(15, 1) = "'you'"
wordSimplications(16, 0) = "addressees are requested"
wordSimplications(16, 1) = "'please'"
wordSimplications(17, 0) = "adjacent to"
wordSimplications(17, 1) = "'next to'"
wordSimplications(18, 0) = "adjustment"
wordSimplications(18, 1) = "'change'"
wordSimplications(19, 0) = "admissible"
wordSimplications(19, 1) = "'accepted' or 'allowed'"
wordSimplications(20, 0) = "advantageous"
wordSimplications(20, 1) = "'helpful'"
wordSimplications(21, 0) = "adversely impact"
wordSimplications(21, 1) = "'hurt'"
wordSimplications(22, 0) = "adversely impact on"
wordSimplications(22, 1) = "'hurt' or 'set back'"
wordSimplications(23, 0) = "advise"
wordSimplications(23, 1) = "'recommend' or 'tell'"
wordSimplications(24, 0) = "afford an opportunity"
wordSimplications(24, 1) = "'allow' or 'let'"
wordSimplications(25, 0) = "aforementioned"
wordSimplications(25, 1) = "'remove'"
wordSimplications(26, 0) = "aggregate"
wordSimplications(26, 1) = "'add' or 'total'"
wordSimplications(27, 0) = "aircraft"
wordSimplications(27, 1) = "'plane'"
wordSimplications(28, 0) = "all of"
wordSimplications(28, 1) = "'all'"
wordSimplications(29, 0) = "alleviate"
wordSimplications(29, 1) = "'ease' or 'reduce'"
wordSimplications(30, 0) = "allocate"
wordSimplications(30, 1) = "'divide'"
wordSimplications(31, 0) = "along the lines of"
wordSimplications(31, 1) = "'as in' or 'like'"
wordSimplications(32, 0) = "already existing"
wordSimplications(32, 1) = "'existing'"
wordSimplications(33, 0) = "alternatively"
wordSimplications(33, 1) = "'or'"
wordSimplications(34, 0) = "ameliorate"
wordSimplications(34, 1) = "'help' or 'improve'"
wordSimplications(35, 0) = "and/or"
wordSimplications(35, 1) = "' … or … or both'"
wordSimplications(36, 0) = "anticipate"
wordSimplications(36, 1) = "'expect'"
wordSimplications(37, 0) = "apparent"
wordSimplications(37, 1) = "'clear' or 'plain'"
wordSimplications(38, 0) = "appreciable"
wordSimplications(38, 1) = "'many'"
wordSimplications(39, 0) = "appropriate"
wordSimplications(39, 1) = "'proper' or 'right'"
wordSimplications(40, 0) = "approximate"
wordSimplications(40, 1) = "'about'"
wordSimplications(41, 0) = "arrive onboard"
wordSimplications(41, 1) = "'arrive'"
wordSimplications(42, 0) = "as a means of"
wordSimplications(42, 1) = "'to'"
wordSimplications(43, 0) = "as of yet"
wordSimplications(43, 1) = "'yet'"
wordSimplications(44, 0) = "as prescribed by"
wordSimplications(44, 1) = "'in'"
wordSimplications(45, 0) = "as to"
wordSimplications(45, 1) = "'about' or 'on'"
wordSimplications(46, 0) = "as yet"
wordSimplications(46, 1) = "'yet'"
wordSimplications(47, 0) = "ascertain"
wordSimplications(47, 1) = "'find out' or 'learn'"
wordSimplications(48, 0) = "assist"
wordSimplications(48, 1) = "'aid' or 'help'"
wordSimplications(49, 0) = "assistance"
wordSimplications(49, 1) = "'aid' or 'help'"
wordSimplications(50, 0) = "at the present time"
wordSimplications(50, 1) = "'at present'"
wordSimplications(51, 0) = "at this time"
wordSimplications(51, 1) = "'now'"
wordSimplications(52, 0) = "attain"
wordSimplications(52, 1) = "'meet'"
wordSimplications(53, 0) = "attempt"
wordSimplications(53, 1) = "'try'"
wordSimplications(54, 0) = "attributable to"
wordSimplications(54, 1) = "'because'"
wordSimplications(55, 0) = "authorise"
wordSimplications(55, 1) = "'allow' or 'let'"
wordSimplications(56, 0) = "authorize"
wordSimplications(56, 1) = "'allow' or 'let'"
wordSimplications(57, 0) = "because of the fact that"
wordSimplications(57, 1) = "'because'"
wordSimplications(58, 0) = "belated"
wordSimplications(58, 1) = "'late'"
wordSimplications(59, 0) = "benefit"
wordSimplications(59, 1) = "'help'"
wordSimplications(60, 0) = "benefit from"
wordSimplications(60, 1) = "'enjoy'"
wordSimplications(61, 0) = "bestow"
wordSimplications(61, 1) = "'award' or 'give'"
wordSimplications(62, 0) = "by means of"
wordSimplications(62, 1) = "'by' or 'with'"
wordSimplications(63, 0) = "by virtue of"
wordSimplications(63, 1) = "'by' or 'under'"
wordSimplications(64, 0) = "capability"
wordSimplications(64, 1) = "'ability'"
wordSimplications(65, 0) = "caveat"
wordSimplications(65, 1) = "'warning'"
wordSimplications(66, 0) = "cease"
wordSimplications(66, 1) = "'stop'"
wordSimplications(67, 0) = "close proximity"
wordSimplications(67, 1) = "'near'"
wordSimplications(68, 0) = "combat environment"
wordSimplications(68, 1) = "'combat'"
wordSimplications(69, 0) = "combined"
wordSimplications(69, 1) = "'joint'"
wordSimplications(70, 0) = "commence"
wordSimplications(70, 1) = "'begin' or 'start'"
wordSimplications(71, 0) = "comply with"
wordSimplications(71, 1) = "'follow'"
wordSimplications(72, 0) = "component"
wordSimplications(72, 1) = "'part'"
wordSimplications(73, 0) = "comprise"
wordSimplications(73, 1) = "'form' or 'include' or 'make up'"
wordSimplications(74, 0) = "concerning"
wordSimplications(74, 1) = "'about' or 'on'"
wordSimplications(75, 0) = "consequently"
wordSimplications(75, 1) = "'so'"
wordSimplications(76, 0) = "consolidate"
wordSimplications(76, 1) = "'combine' or 'join' or 'merge'"
wordSimplications(77, 0) = "constitutes"
wordSimplications(77, 1) = "'forms' or 'is' or 'makes up'"
wordSimplications(78, 0) = "contains"
wordSimplications(78, 1) = "'has'"
wordSimplications(79, 0) = "convene"
wordSimplications(79, 1) = "'meet'"
wordSimplications(80, 0) = "currently"
wordSimplications(80, 1) = "'now'"
wordSimplications(81, 0) = "deem"
wordSimplications(81, 1) = "'believe' or 'consider' or 'think'"
wordSimplications(82, 0) = "delete"
wordSimplications(82, 1) = "'cut' or 'drop'"
wordSimplications(83, 0) = "demonstrate"
wordSimplications(83, 1) = "'prove' or 'show'"
wordSimplications(84, 0) = "depart"
wordSimplications(84, 1) = "'go' or 'leave'"
wordSimplications(85, 0) = "designate"
wordSimplications(85, 1) = "'appoint' or 'choose' or 'name'"
wordSimplications(86, 0) = "desire"
wordSimplications(86, 1) = "'want' or 'wish'"
wordSimplications(87, 0) = "determine"
wordSimplications(87, 1) = "'decide' or 'figure' or 'find'"
wordSimplications(88, 0) = "disclose"
wordSimplications(88, 1) = "'show'"
wordSimplications(89, 0) = "discontinue"
wordSimplications(89, 1) = "'drop' or 'stop'"
wordSimplications(90, 0) = "disseminate"
wordSimplications(90, 1) = "'give' or 'issue' or 'pass' or 'send'"
wordSimplications(91, 0) = "due to the fact that"
wordSimplications(91, 1) = "'because' or 'due to' or 'since'"
wordSimplications(92, 0) = "during the period"
wordSimplications(92, 1) = "'during'"
wordSimplications(93, 0) = "e.g."
wordSimplications(93, 1) = "'for example' or 'such as'"
wordSimplications(94, 0) = "each and every"
wordSimplications(94, 1) = "'each'"
wordSimplications(95, 0) = "economical"
wordSimplications(95, 1) = "'cheap'"
wordSimplications(96, 0) = "effect"
wordSimplications(96, 1) = "'choose' or 'pick'"
wordSimplications(97, 0) = "effect modifications"
wordSimplications(97, 1) = "'make changes'"
wordSimplications(98, 0) = "elect"
wordSimplications(98, 1) = "'choose'"
wordSimplications(99, 0) = "eliminate"
wordSimplications(99, 1) = "'cut' or 'drop' or 'end' or 'stop'"
wordSimplications(100, 0) = "elucidate"
wordSimplications(100, 1) = "'explain'"
wordSimplications(101, 0) = "employ"
wordSimplications(101, 1) = "'use'"
wordSimplications(102, 0) = "encounter"
wordSimplications(102, 1) = "'meet'"
wordSimplications(103, 0) = "endeavor"
wordSimplications(103, 1) = "'try'"
wordSimplications(104, 0) = "ensure"
wordSimplications(104, 1) = "'make sure'"
wordSimplications(105, 0) = "enumerate"
wordSimplications(105, 1) = "'count'"
wordSimplications(106, 0) = "equipments"
wordSimplications(106, 1) = "'equipment'"
wordSimplications(107, 0) = "equitable"
wordSimplications(107, 1) = "'fair'"
wordSimplications(108, 0) = "equivalent"
wordSimplications(108, 1) = "'equal'"
wordSimplications(109, 0) = "establish"
wordSimplications(109, 1) = "'set up' or 'prove' or 'show'"
wordSimplications(110, 0) = "evaluate"
wordSimplications(110, 1) = "'check' or 'test'"
wordSimplications(111, 0) = "evidenced"
wordSimplications(111, 1) = "'showed'"
wordSimplications(112, 0) = "evident"
wordSimplications(112, 1) = "'clear'"
wordSimplications(113, 0) = "exclusively"
wordSimplications(113, 1) = "'only'"
wordSimplications(114, 0) = "exhibit"
wordSimplications(114, 1) = "'show'"
wordSimplications(115, 0) = "expedite"
wordSimplications(115, 1) = "'hasten' or 'hurry' or 'speed up'"
wordSimplications(116, 0) = "expeditious"
wordSimplications(116, 1) = "'fast' or 'quick'"
wordSimplications(117, 0) = "expend"
wordSimplications(117, 1) = "'spend'"
wordSimplications(118, 0) = "expertise"
wordSimplications(118, 1) = "'ability'"
wordSimplications(119, 0) = "expiration"
wordSimplications(119, 1) = "'end'"
wordSimplications(120, 0) = "facilitate"
wordSimplications(120, 1) = "'ease' or 'help'"
wordSimplications(121, 0) = "factual evidence"
wordSimplications(121, 1) = "'evidence' or 'facts'"
wordSimplications(122, 0) = "failed to"
wordSimplications(122, 1) = "'didnâ€™t'"
wordSimplications(123, 0) = "feasible"
wordSimplications(123, 1) = "'can be done' or 'workable'"
wordSimplications(124, 0) = "females"
wordSimplications(124, 1) = "'women'"
wordSimplications(125, 0) = "finalise"
wordSimplications(125, 1) = "'complete' or 'finish'"
wordSimplications(126, 0) = "finalize"
wordSimplications(126, 1) = "'complete' or 'finish'"
wordSimplications(127, 0) = "first and foremost"
wordSimplications(127, 1) = "'first'"
wordSimplications(128, 0) = "for a period of"
wordSimplications(128, 1) = "'for'"
wordSimplications(129, 0) = "for the purpose of"
wordSimplications(129, 1) = "'to'"
wordSimplications(130, 0) = "forfeit"
wordSimplications(130, 1) = "'give up' or 'lose'"
wordSimplications(131, 0) = "formulate"
wordSimplications(131, 1) = "'plan'"
wordSimplications(132, 0) = "forward"
wordSimplications(132, 1) = "'send'"
wordSimplications(133, 0) = "frequently"
wordSimplications(133, 1) = "'often'"
wordSimplications(134, 0) = "function"
wordSimplications(134, 1) = "'act' or 'role' or 'work'"
wordSimplications(135, 0) = "furnish"
wordSimplications(135, 1) = "'give' or 'send'"
wordSimplications(136, 0) = "has a requirement for"
wordSimplications(136, 1) = "'needs'"
wordSimplications(137, 0) = "herein"
wordSimplications(137, 1) = "'here'"
wordSimplications(138, 0) = "heretofore"
wordSimplications(138, 1) = "'until now'"
wordSimplications(139, 0) = "herewith"
wordSimplications(139, 1) = "'here' or 'below'"
wordSimplications(140, 0) = "honest truth"
wordSimplications(140, 1) = "'truth'"
wordSimplications(141, 0) = "however"
wordSimplications(141, 1) = "'but' or 'yet'"
wordSimplications(142, 0) = "i.e."
wordSimplications(142, 1) = "'as in'"
wordSimplications(143, 0) = "identical"
wordSimplications(143, 1) = "'same'"
wordSimplications(144, 0) = "identify"
wordSimplications(144, 1) = "'find' or 'name' or 'show'"
wordSimplications(145, 0) = "if and when"
wordSimplications(145, 1) = "'if' or 'when'"
wordSimplications(146, 0) = "immediately"
wordSimplications(146, 1) = "'at once'"
wordSimplications(147, 0) = "impacted"
wordSimplications(147, 1) = "'affected' or 'changed' or 'harmed'"
wordSimplications(148, 0) = "implement"
wordSimplications(148, 1) = "'carry out' or 'install' or 'put in place' or 'tool' or 'start'"
wordSimplications(149, 0) = "in a timely manner"
wordSimplications(149, 1) = "'on time' or 'promptly'"
wordSimplications(150, 0) = "in accordance with"
wordSimplications(150, 1) = "'by' or 'under' or 'following' or 'per'"
wordSimplications(151, 0) = "in addition"
wordSimplications(151, 1) = "'also' or 'besides' or 'too'"
wordSimplications(152, 0) = "in all likelihood"
wordSimplications(152, 1) = "'probably'"
wordSimplications(153, 0) = "in an effort to"
wordSimplications(153, 1) = "'to'"
wordSimplications(154, 0) = "in between"
wordSimplications(154, 1) = "'between'"
wordSimplications(155, 0) = "in excess of"
wordSimplications(155, 1) = "'more than'"
wordSimplications(156, 0) = "in lieu of"
wordSimplications(156, 1) = "'instead'"
wordSimplications(157, 0) = "in light of the fact that"
wordSimplications(157, 1) = "'because'"
wordSimplications(158, 0) = "in many cases"
wordSimplications(158, 1) = "'often'"
wordSimplications(159, 0) = "in order that"
wordSimplications(159, 1) = "'for' or 'so'"
wordSimplications(160, 0) = "in order to"
wordSimplications(160, 1) = "'to'"
wordSimplications(161, 0) = "in regard to"
wordSimplications(161, 1) = "'about' or 'concerning' or 'on'"
wordSimplications(162, 0) = "in relation to"
wordSimplications(162, 1) = "'about' or 'with' or 'to'"
wordSimplications(163, 0) = "in some instances"
wordSimplications(163, 1) = "'sometimes'"
wordSimplications(164, 0) = "in terms of"
wordSimplications(164, 1) = "'as' or 'for' or 'with'"
wordSimplications(165, 0) = "in the amount of"
wordSimplications(165, 1) = "'for'"
wordSimplications(166, 0) = "in the event of"
wordSimplications(166, 1) = "'if'"
wordSimplications(167, 0) = "in the near future"
wordSimplications(167, 1) = "'soon' or 'shortly'"
wordSimplications(168, 0) = "in view of"
wordSimplications(168, 1) = "'since'"
wordSimplications(169, 0) = "in view of the above"
wordSimplications(169, 1) = "'so'"
wordSimplications(170, 0) = "inasmuch as"
wordSimplications(170, 1) = "'since'"
wordSimplications(171, 0) = "inception"
wordSimplications(171, 1) = "'start'"
wordSimplications(172, 0) = "incumbent upon"
wordSimplications(172, 1) = "'must'"
wordSimplications(173, 0) = "indicate"
wordSimplications(173, 1) = "'show' or 'say' or 'state' or 'write down'"
wordSimplications(174, 0) = "indication"
wordSimplications(174, 1) = "'sign'"
wordSimplications(175, 0) = "initial"
wordSimplications(175, 1) = "'first'"
wordSimplications(176, 0) = "initiate"
wordSimplications(176, 1) = "'start'"
wordSimplications(177, 0) = "interface"
wordSimplications(177, 1) = "'meet' or 'work with'"
wordSimplications(178, 0) = "interpose no objection"
wordSimplications(178, 1) = "'donâ€™t object'"
wordSimplications(179, 0) = "is applicable to"
wordSimplications(179, 1) = "'applies to'"
wordSimplications(180, 0) = "is authorised to"
wordSimplications(180, 1) = "'may'"
wordSimplications(181, 0) = "is authorized to"
wordSimplications(181, 1) = "'may'"
wordSimplications(182, 0) = "is in consonance with"
wordSimplications(182, 1) = "'agrees with' or 'follows'"
wordSimplications(183, 0) = "is responsible for"
wordSimplications(183, 1) = "'handles'"
wordSimplications(184, 0) = "it appears"
wordSimplications(184, 1) = "'seems'"
wordSimplications(185, 0) = "it is essential"
wordSimplications(185, 1) = "'must' or 'need to'"
wordSimplications(186, 0) = "it is requested"
wordSimplications(186, 1) = "'please'"
wordSimplications(187, 0) = "liaison"
wordSimplications(187, 1) = "'discussion'"
wordSimplications(188, 0) = "limited number"
wordSimplications(188, 1) = "'limits'"
wordSimplications(189, 0) = "magnitude"
wordSimplications(189, 1) = "'size'"
wordSimplications(190, 0) = "maintain"
wordSimplications(190, 1) = "'support' or 'keep'"
wordSimplications(191, 0) = "maximum"
wordSimplications(191, 1) = "'greatest' or 'largest' or 'most'"
wordSimplications(192, 0) = "methodology"
wordSimplications(192, 1) = "'method'"
wordSimplications(193, 0) = "minimise"
wordSimplications(193, 1) = "'cut' or 'decrease'"
wordSimplications(194, 0) = "minimize"
wordSimplications(194, 1) = "'cut' or 'decrease'"
wordSimplications(195, 0) = "minimum"
wordSimplications(195, 1) = "'least' or 'small' or 'smallest'"
wordSimplications(196, 0) = "modify"
wordSimplications(196, 1) = "'change'"
wordSimplications(197, 0) = "monitor"
wordSimplications(197, 1) = "'check' or 'track' or 'watch'"
wordSimplications(198, 0) = "multiple"
wordSimplications(198, 1) = "'many'"
wordSimplications(199, 0) = "necessitate"
wordSimplications(199, 1) = "'cause' or 'need'"
wordSimplications(200, 0) = "nevertheless"
wordSimplications(200, 1) = "'besides' or 'even so' or 'still'"
wordSimplications(201, 0) = "not certain"
wordSimplications(201, 1) = "'uncertain'"
wordSimplications(202, 0) = "not later than"
wordSimplications(202, 1) = "'by' or 'before'"
wordSimplications(203, 0) = "not many"
wordSimplications(203, 1) = "'few'"
wordSimplications(204, 0) = "not often"
wordSimplications(204, 1) = "'rarely'"
wordSimplications(205, 0) = "not unless"
wordSimplications(205, 1) = "'only if'"
wordSimplications(206, 0) = "not unlike"
wordSimplications(206, 1) = "'alike' or 'similar'"
wordSimplications(207, 0) = "notify"
wordSimplications(207, 1) = "'let know' or 'tell'"
wordSimplications(208, 0) = "notwithstanding"
wordSimplications(208, 1) = "'despite' or 'in spite of' or 'still'"
wordSimplications(209, 0) = "null and void"
wordSimplications(209, 1) = "'null' or 'void'"
wordSimplications(210, 0) = "numerous"
wordSimplications(210, 1) = "'many'"
wordSimplications(211, 0) = "objective"
wordSimplications(211, 1) = "'aim' or 'goal'"
wordSimplications(212, 0) = "obligate"
wordSimplications(212, 1) = "'bind' or 'compel'"
wordSimplications(213, 0) = "observe"
wordSimplications(213, 1) = "'see'"
wordSimplications(214, 0) = "obtain"
wordSimplications(214, 1) = "'get'"
wordSimplications(215, 0) = "on the contrary"
wordSimplications(215, 1) = "'but' or 'so'"
wordSimplications(216, 0) = "on the other hand"
wordSimplications(216, 1) = "'but' or 'so'"
wordSimplications(217, 0) = "one particular"
wordSimplications(217, 1) = "'one'"
wordSimplications(218, 0) = "operate"
wordSimplications(218, 1) = "'run' or 'use' or 'work'"
wordSimplications(219, 0) = "optimum"
wordSimplications(219, 1) = "'best' or 'greatest' or 'most'"
wordSimplications(220, 0) = "option"
wordSimplications(220, 1) = "'choice'"
wordSimplications(221, 0) = "owing to the fact that"
wordSimplications(221, 1) = "'because' or 'since'"
wordSimplications(222, 0) = "parameters"
wordSimplications(222, 1) = "'limits'"
wordSimplications(223, 0) = "participate"
wordSimplications(223, 1) = "'take part'"
wordSimplications(224, 0) = "particulars"
wordSimplications(224, 1) = "'details'"
wordSimplications(225, 0) = "pass away"
wordSimplications(225, 1) = "'die'"
wordSimplications(226, 0) = "perform"
wordSimplications(226, 1) = "'do'"
wordSimplications(227, 0) = "permit"
wordSimplications(227, 1) = "'let'"
wordSimplications(228, 0) = "pertaining to"
wordSimplications(228, 1) = "'about' or 'of' or 'on'"
wordSimplications(229, 0) = "point in time"
wordSimplications(229, 1) = "'moment' or 'now' or 'point' or 'time'"
wordSimplications(230, 0) = "portion"
wordSimplications(230, 1) = "'part'"
wordSimplications(231, 0) = "possess"
wordSimplications(231, 1) = "'have' or 'own'"
wordSimplications(232, 0) = "practicable"
wordSimplications(232, 1) = "'practical'"
wordSimplications(233, 0) = "preclude"
wordSimplications(233, 1) = "'prevent'"
wordSimplications(234, 0) = "previous"
wordSimplications(234, 1) = "'earlier'"
wordSimplications(235, 0) = "previously"
wordSimplications(235, 1) = "'before'"
wordSimplications(236, 0) = "prior to"
wordSimplications(236, 1) = "'before'"
wordSimplications(237, 0) = "prioritise"
wordSimplications(237, 1) = "'focus on' or 'rank'"
wordSimplications(238, 0) = "prioritize"
wordSimplications(238, 1) = "'focus on' or 'rank'"
wordSimplications(239, 0) = "proceed"
wordSimplications(239, 1) = "'do' or 'go ahead' or 'try'"
wordSimplications(240, 0) = "procure"
wordSimplications(240, 1) = "'buy' or 'get'"
wordSimplications(241, 0) = "proficiency"
wordSimplications(241, 1) = "'skill'"
wordSimplications(242, 0) = "promulgate"
wordSimplications(242, 1) = "'issue' or 'publish'"
wordSimplications(243, 0) = "provide"
wordSimplications(243, 1) = "'give' or 'offer' or 'say'"
wordSimplications(244, 0) = "provided that"
wordSimplications(244, 1) = "'if'"
wordSimplications(245, 0) = "provides guidance for"
wordSimplications(245, 1) = "'guides'"
wordSimplications(246, 0) = "purchase"
wordSimplications(246, 1) = "'buy' or 'sale'"
wordSimplications(247, 0) = "pursuant to"
wordSimplications(247, 1) = "'by' or 'following' or 'per' or 'under'"
wordSimplications(248, 0) = "readily apparent"
wordSimplications(248, 1) = "'clear'"
wordSimplications(249, 0) = "refer back"
wordSimplications(249, 1) = "'refer'"
wordSimplications(250, 0) = "reflect"
wordSimplications(250, 1) = "'say' or 'show'"
wordSimplications(251, 0) = "regarding"
wordSimplications(251, 1) = "'about' or 'of' or 'on'"
wordSimplications(252, 0) = "relative to"
wordSimplications(252, 1) = "'about' or 'on'"
wordSimplications(253, 0) = "relocate"
wordSimplications(253, 1) = "'move'"
wordSimplications(254, 0) = "remain"
wordSimplications(254, 1) = "'stay'"
wordSimplications(255, 0) = "remainder"
wordSimplications(255, 1) = "'rest'"
wordSimplications(256, 0) = "remuneration"
wordSimplications(256, 1) = "'pay' or 'payment'"
wordSimplications(257, 0) = "render"
wordSimplications(257, 1) = "'give' or 'make'"
wordSimplications(258, 0) = "represents"
wordSimplications(258, 1) = "'is'"
wordSimplications(259, 0) = "require"
wordSimplications(259, 1) = "'must' or 'need'"
wordSimplications(260, 0) = "requirement"
wordSimplications(260, 1) = "'need' or 'rule'"
wordSimplications(261, 0) = "reside"
wordSimplications(261, 1) = "'live'"
wordSimplications(262, 0) = "residence"
wordSimplications(262, 1) = "'house'"
wordSimplications(263, 0) = "retain"
wordSimplications(263, 1) = "'keep'"
wordSimplications(264, 0) = "satisfy"
wordSimplications(264, 1) = "'meet' or 'please'"
wordSimplications(265, 0) = "selection"
wordSimplications(265, 1) = "'choice'"
wordSimplications(266, 0) = "set forth in"
wordSimplications(266, 1) = "'in'"
wordSimplications(267, 0) = "shall"
wordSimplications(267, 1) = "'must' or 'will'"
wordSimplications(268, 0) = "should you wish"
wordSimplications(268, 1) = "'if you want'"
wordSimplications(269, 0) = "similar to"
wordSimplications(269, 1) = "'like'"
wordSimplications(270, 0) = "solicit"
wordSimplications(270, 1) = "'ask for' or 'request'"
wordSimplications(271, 0) = "span across"
wordSimplications(271, 1) = "'cross' or 'span'"
wordSimplications(272, 0) = "state-of-the-art"
wordSimplications(272, 1) = "'latest'"
wordSimplications(273, 0) = "strategise"
wordSimplications(273, 1) = "'plan'"
wordSimplications(274, 0) = "strategize"
wordSimplications(274, 1) = "'plan'"
wordSimplications(275, 0) = "submit"
wordSimplications(275, 1) = "'give' or 'send'"
wordSimplications(276, 0) = "subsequent"
wordSimplications(276, 1) = "'after' or 'later' or 'next' or 'then'"
wordSimplications(277, 0) = "subsequently"
wordSimplications(277, 1) = "'after' or 'later' or 'then'"
wordSimplications(278, 0) = "substantial"
wordSimplications(278, 1) = "'large' or 'much'"
wordSimplications(279, 0) = "successfully complete"
wordSimplications(279, 1) = "'complete' or 'pass'"
wordSimplications(280, 0) = "sufficient"
wordSimplications(280, 1) = "'enough'"
wordSimplications(281, 0) = "terminate"
wordSimplications(281, 1) = "'end' or 'stop'"
wordSimplications(282, 0) = "the undersigned"
wordSimplications(282, 1) = "'I'"
wordSimplications(283, 0) = "therefore"
wordSimplications(283, 1) = "'so' or 'thus'"
wordSimplications(284, 0) = "therein"
wordSimplications(284, 1) = "'there'"
wordSimplications(285, 0) = "thereof"
wordSimplications(285, 1) = "'its' or 'their'"
wordSimplications(286, 0) = "this day and age"
wordSimplications(286, 1) = "'today'"
wordSimplications(287, 0) = "time period"
wordSimplications(287, 1) = "'period' or 'time'"
wordSimplications(288, 0) = "timely"
wordSimplications(288, 1) = "'prompt'"
wordSimplications(289, 0) = "took advantage of"
wordSimplications(289, 1) = "'preyed on'"
wordSimplications(290, 0) = "transmit"
wordSimplications(290, 1) = "'send'"
wordSimplications(291, 0) = "transpire"
wordSimplications(291, 1) = "'happen'"
wordSimplications(292, 0) = "under the provisions of"
wordSimplications(292, 1) = "'under'"
wordSimplications(293, 0) = "until such time as"
wordSimplications(293, 1) = "'until'"
wordSimplications(294, 0) = "utilisation"
wordSimplications(294, 1) = "'use'"
wordSimplications(295, 0) = "utilise"
wordSimplications(295, 1) = "'use'"
wordSimplications(296, 0) = "utilization"
wordSimplications(296, 1) = "'use'"
wordSimplications(297, 0) = "utilize"
wordSimplications(297, 1) = "'use'"
wordSimplications(298, 0) = "validate"
wordSimplications(298, 1) = "'confirm'"
wordSimplications(299, 0) = "various different"
wordSimplications(299, 1) = "'different' or 'various'"
wordSimplications(300, 0) = "viable"
wordSimplications(300, 1) = "'practical' or 'workable'"
wordSimplications(301, 0) = "vice"
wordSimplications(301, 1) = "'instead of' or 'versus'"
wordSimplications(302, 0) = "warrant"
wordSimplications(302, 1) = "'call for' or 'permit'"
wordSimplications(303, 0) = "whereas"
wordSimplications(303, 1) = "'because' or 'since'"
wordSimplications(304, 0) = "whether or not"
wordSimplications(304, 1) = "'whether'"
wordSimplications(305, 0) = "with reference to"
wordSimplications(305, 1) = "'about'"
wordSimplications(306, 0) = "with respect to"
wordSimplications(306, 1) = "'about' or 'on'"
wordSimplications(307, 0) = "with the exception of"
wordSimplications(307, 1) = "'except for'"
wordSimplications(308, 0) = "witnessed"
wordSimplications(308, 1) = "'saw' or 'seen'"
wordSimplications(309, 0) = "your office"
wordSimplications(309, 1) = "'you'"


End Sub
