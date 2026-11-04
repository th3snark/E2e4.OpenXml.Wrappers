<#
.SYNOPSIS
The scenarios New-WordReferences.ps1 runs in Word. Each is a scriptblock of Word object model calls on a new document (Documents.Add()).
.DESCRIPTION
The name is the folder the snapshot goes to under OpenXml.Wrappers.Test/TestFiles/WordReference, and MarkupConformanceTest has a case with
the same name that does the same thing through the library. Word's indexes are 1-based; the library's are 0-based. Keep the two in step.
#>

[ordered]@{
    'blank-document'                  = { param($doc) }

    'insert-after-tab-and-line-break' = {
        param($doc)
        # [char]11 is a manual line break in Word's text (vbVerticalTab).
        $doc.Content.InsertAfter("Before`tafter the tab" + [char]11 + 'after the break')
    }

    'bold-on'                         = {
        param($doc)
        $doc.Content.InsertAfter('Bold text')
        $doc.Range(0, 4).Font.Bold = $true
    }

    'bold-on-then-off'                = {
        param($doc)
        $doc.Content.InsertAfter('Bold text')
        $doc.Range(0, 4).Font.Bold = $true
        $doc.Range(0, 4).Font.Bold = $false
    }

    'font-name'                       = {
        param($doc)
        $doc.Content.InsertAfter('Arial text')
        $doc.Range(0, 5).Font.Name = 'Arial'
    }

    'table-3x3'                       = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 3, 3)
    }

    # The table the Insert Table command makes: Word 2000 behaviour, which gives it the Table Grid style.
    'table-3x3-word9-behavior'        = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 3, 3, 1, 0)  # wdWord9TableBehavior, wdAutoFitFixed
    }

    'cell-merge'                      = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $table.Cell(1, 1).Merge($table.Cell(1, 2))
    }

    'comment-with-reply'              = {
        param($doc)
        $doc.Content.InsertAfter('Commented text')
        $comment = $doc.Comments.Add($doc.Range(0, 9), 'A comment')
        [void]$comment.Replies.Add($comment.Scope, 'A reply')
    }

    'bookmark'                        = {
        param($doc)
        $doc.Content.InsertAfter('Bookmarked text')
        [void]$doc.Bookmarks.Add('Mark', $doc.Range(0, 10))
    }

    'date-field'                      = {
        param($doc)
        [void]$doc.Fields.Add($doc.Range(0, 0), 31)  # wdFieldDate
    }

    'header-footer-text'              = {
        param($doc)
        $section = $doc.Sections.Item(1)
        $section.Headers.Item(1).Range.Text = 'Header text'  # wdHeaderFooterPrimary
        $section.Footers.Item(1).Range.Text = 'Footer text'
    }

    'section-add'                     = {
        param($doc)
        $doc.Content.InsertAfter('First section')
        [void]$doc.Sections.Add()
    }

    'page-borders'                    = {
        param($doc)
        $doc.Sections.Item(1).Borders.Enable = $true
    }

    'protect-read-only'               = {
        param($doc)
        $doc.Protect(3, $false, 'password')  # wdAllowOnlyReading
    }

    'heading-insert-paragraph-after'  = {
        param($doc)
        $doc.Content.InsertAfter('Heading')
        $doc.Paragraphs.Item(1).Style = -2  # wdStyleHeading1
        $doc.Paragraphs.Item(1).Range.InsertParagraphAfter()
    }

    # Each built-in style the library can add (BuiltInStyle, plus Table Grid), applied in turn: Word writes the definition of each, and of
    # the linked character style of a paragraph style. These are also the definitions the library adds (OpenXml.Wrappers/Word/Templates).
    # Keep the list in step with BuiltInStylesScenario in MarkupConformanceTest.
    'built-in-styles'                 = {
        param($doc)
        $styles = @(-2, -3, -4, -5, -6, -7, -8, -9, -10, -20, -21, -22, -23, -24, -25, -26, -27, -28, -30, -31, -32, -33, -35, -39, -40, -41,
            -42, -43, -44, -63, -64, -65, -67, -68, -75, -76, -77, -81, -82, -86, -87, -88, -89, -91, -180, -181, -182, -261, -262, -263, -264, -265,
            'Comment Subject')  # no WdBuiltinStyle constant; Word adds it with a comment
        foreach ($id in $styles) {
            $style = $doc.Styles.Item([object]$id)
            $doc.Content.InsertAfter('x')
            $paragraph = $doc.Paragraphs.Item($doc.Paragraphs.Count).Range
            if ($style.Type -eq 2) {  # wdStyleTypeCharacter
                $doc.Range($paragraph.Start, $paragraph.Start + 1).Style = $style
            }
            else {
                $paragraph.Style = $style
            }
            $doc.Content.InsertParagraphAfter()
        }
        $table = $doc.Tables.Add($doc.Paragraphs.Item($doc.Paragraphs.Count).Range, 1, 1)
        $table.Style = 'Table Grid'
    }

    'join-paragraphs'                 = {
        param($doc)
        $doc.Content.InsertAfter("First paragraph`rSecond paragraph")
        $paragraph = $doc.Paragraphs.Item(1).Range
        [void]$paragraph.Characters.Last.Delete()
    }
}
