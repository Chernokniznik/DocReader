using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;

namespace docReader
{
   public static class ConvertWord
    {
        public static void FillMergeFields(string filePath, Dictionary<string, string> mergeFields)
        {


            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                // Get all word document paragraphs
                var paragraphs = document.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
                // Regex to find MergeFields
                Regex containsMergeField = new Regex(@"(?<=MERGEFIELD  )(.*?)(?=  \\\* MERGEFORMAT)", RegexOptions.Compiled);

                // Loop through all paragraphs
                foreach (var paragraph in paragraphs)
                {
                    // If pargraph contains MergeField
                    MatchCollection mergeField = containsMergeField.Matches(paragraph.InnerText);
                    if (mergeField.Count > 0)
                    {
                        int i = 0;
                        // Get Merge field name & remove quotes
                        string mergeFieldName = mergeField[0].ToString();
                        if (mergeFieldName[0] == '"')
                        {
                            mergeFieldName = mergeFieldName.Substring(1, mergeFieldName.Length - 2);
                        }

                        // Check that info about merge field exist
                        if (mergeFields.ContainsKey(mergeFieldName))
                        {
                            bool editMergeField = false;
                            var runs = paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>();
                            // Loop through all run properties
                            foreach (DocumentFormat.OpenXml.Wordprocessing.Run run in runs)
                            {
                                // Look for begin/end of merge field and set edit flag appropriatly
                                var fieldChars = run.Descendants<FieldChar>();
                                if (fieldChars.Count() != 0)
                                {
                                    FieldChar fieldChar = fieldChars.First();
                                    if (fieldChar.FieldCharType.ToString() == "begin")
                                        editMergeField = true;
                                    if (fieldChar.FieldCharType.ToString() == "end")
                                        editMergeField = false;
                                    run.RemoveAllChildren();
                                }
                                // For all other tags
                                else
                                {
                                    // If tag is between begin & end of MergeField
                                    if (editMergeField == true)
                                    {
                                        // If it is the beginning of MergeField text
                                        string innerText = run.InnerText;
                                        if (innerText[0] == '«')
                                        {
                                            // Replace it with appropriate value from dictionary
                                            mergeFieldName = mergeField[i].ToString();
                                            if (mergeFieldName[0] == '"')
                                            {
                                                mergeFieldName = mergeFieldName.Substring(1, mergeFieldName.Length - 2);
                                            }
                                            Text t = new Text(mergeFields[mergeFieldName]);
                                            i++;
                                            run.RemoveAllChildren<Text>();
                                            run.AppendChild<Text>(t);
                                        }
                                        else
                                        {
                                            // Remove all other run properties
                                            run.RemoveAllChildren();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                // Save document
                document.MainDocumentPart.Document.Save();
            }
        }
    }
}
