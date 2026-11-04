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

    'join-paragraphs'                 = {
        param($doc)
        $doc.Content.InsertAfter("First paragraph`rSecond paragraph")
        $paragraph = $doc.Paragraphs.Item(1).Range
        [void]$paragraph.Characters.Last.Delete()
    }
}
