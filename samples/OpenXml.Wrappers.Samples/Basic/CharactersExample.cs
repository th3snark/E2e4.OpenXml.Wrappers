using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Basic;

/// <summary>
/// Character positions: a range can be any span of characters, as in Word. Format part of a sentence, step through words, bookmark and
/// comment on a few characters, and replace text without disturbing the formatting around it.
/// </summary>
internal static class CharactersExample
{
    public static void Create(string path)
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, true))
        {
            document.Content.Text = "Character ranges\r"
                + "A range does not have to be a whole paragraph. It can be any run of characters, such as this phrase, a single word or one letter.\r"
                + "The first sentence is plain. The second sentence will be formatted. The third sentence gets a comment.\r"
                + "Replacing text in the middle of formatted text keeps the formatting: the price is 100 dollars and the price is 100 dollars.";

            document.Paragraphs[0].Range.Font.Bold = true;
            document.Paragraphs[0].Range.Font.Size = 20;

            //Document.Range(start, end) is a span of characters, counted from 0 at the start of the text. The end of each paragraph counts as one
            //character, so the second paragraph starts at 17.
            RangeWrapper phrase = document.Content;
            phrase.Find.Execute("such as this phrase");

            //A successful Find narrows the range to the match, so the phrase can be formatted without touching the rest of its run.
            phrase.Font.Italic = true;
            phrase.Font.Color = "1F3864";

            //Words, sentences and characters are ranges too. Words include the space after them.
            RangeWrapper paragraph = document.Paragraphs[1].Range;
            RangeWrapper word = paragraph.Words[3];
            word.Font.Bold = true;

            RangeWrapper sentences = document.Paragraphs[2].Range;
            sentences.Sentences[1].Font.Underline = UnderlineStyle.SingleLine;
            sentences.Sentences[1].Font.Color = "C00000";

            //A comment on the third sentence, and a bookmark on one word.
            document.Comments.Add(sentences.Sentences[2], "Should this say what the comment is for?");
            document.Bookmarks.Add("formatted_word", sentences.Words[2]);

            //Moving the ends of a range: find the word 'first', widen the range to the whole word and then two words further.
            RangeWrapper moving = document.Paragraphs[2].Range;
            moving.Find.Execute("first");
            moving.Expand(TextUnit.Word);
            moving.MoveEnd(TextUnit.Word, 2);
            moving.Font.Bold = true;

            //Replacing part of a run keeps its formatting, and each match can be given its own formatting through the range.
            RangeWrapper prices = document.Paragraphs[3].Range;
            prices.Find.Execute("100 dollars", replaceWith: "$100", replace: ReplaceMode.All);
            RangeWrapper search = document.Paragraphs[3].Range;
            while (search.Find.Execute("$100"))
            {
                search.Font.Bold = true;
                search.Font.Color = "375623";
            }

            //An insertion point: a collapsed range. InsertAfter puts text at that point, taking the formatting of the character before it.
            RangeWrapper point = document.Paragraphs[3].Range;
            point.Find.Execute("Replacing");
            point.Collapse(CollapseDirection.End);
            point.InsertAfter(" (or deleting)");
        }
    }
}
